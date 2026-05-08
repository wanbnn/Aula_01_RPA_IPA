from __future__ import annotations

import asyncio
import os
import platform
from collections.abc import AsyncIterator, Sequence
from contextlib import asynccontextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

from mcp.server.fastmcp import Context, FastMCP
from mcp.server.session import ServerSession


EXCEL_FILE_FORMATS: dict[str, int] = {
    ".xlsx": 51,
    ".xlsm": 52,
    ".xlsb": 50,
    ".xls": 56,
    ".csv": 6,
}

CHART_TYPES: dict[str, int] = {
    "column": 51,
    "bar": 57,
    "line": 4,
    "pie": 5,
    "area": 1,
    "scatter": -4169,
}

H_ALIGN: dict[str, int] = {
    "general": 1,
    "left": -4131,
    "center": -4108,
    "right": -4152,
}

V_ALIGN: dict[str, int] = {
    "top": -4160,
    "center": -4108,
    "bottom": -4107,
}

BORDER_LINE_STYLE_CONTINUOUS = 1
BORDER_WEIGHT_THIN = 2
COLOR_INDEX_AUTOMATIC = -4105
EXPORT_FORMAT_PDF = 0
XL_SRC_RANGE = 1
XL_YES = 1
XL_CELL_TYPE_VISIBLE = 12
XL_UP = -4162
XL_TO_LEFT = -4159


class ExcelMcpError(RuntimeError):
    """User-facing Excel MCP error."""


@dataclass(slots=True)
class RuntimeSettings:
    visible: bool
    display_alerts: bool
    max_read_cells: int
    default_output_dir: Path

    @classmethod
    def from_env(cls) -> "RuntimeSettings":
        return cls(
            visible=_env_flag("EXCEL_MCP_VISIBLE", True),
            display_alerts=_env_flag("EXCEL_MCP_DISPLAY_ALERTS", False),
            max_read_cells=int(os.getenv("EXCEL_MCP_MAX_READ_CELLS", "5000")),
            default_output_dir=Path(os.getenv("EXCEL_MCP_OUTPUT_DIR", str(Path.cwd()))).resolve(),
        )


@dataclass(slots=True)
class AppContext:
    controller: "ExcelController"


def _env_flag(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() not in {"0", "false", "no", "off"}


def _require_windows() -> None:
    if platform.system() != "Windows":
        raise ExcelMcpError("O servidor Excel MCP requer Windows com Microsoft Excel instalado.")


def _load_pywin32() -> tuple[Any, Any]:
    _require_windows()
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise ExcelMcpError(
            "Dependencia ausente: instale pywin32 no Windows com `python -m pip install pywin32`."
        ) from exc
    return pythoncom, win32com.client


def _normalize_path(path: str | None, *, default_dir: Path, default_name: str | None = None) -> Path | None:
    if path is None or not path.strip():
        if default_name is None:
            return None
        return (default_dir / default_name).resolve()
    return Path(path).expanduser().resolve()


def _extension_format(path: Path) -> int:
    suffix = path.suffix.lower()
    if suffix not in EXCEL_FILE_FORMATS:
        raise ExcelMcpError(
            f"Extensao nao suportada: {suffix}. Use uma destas: {', '.join(sorted(EXCEL_FILE_FORMATS))}."
        )
    return EXCEL_FILE_FORMATS[suffix]


def _require_non_empty(value: str, field_name: str) -> str:
    if not value or not value.strip():
        raise ExcelMcpError(f"{field_name} nao pode ser vazio.")
    return value.strip()


def _sheet_name(name: str) -> str:
    value = _require_non_empty(name, "sheet_name")
    invalid = set("[]:*?/\\")
    if any(char in invalid for char in value):
        raise ExcelMcpError("Nome de planilha contem caracteres invalidos: []:*?/\\")
    if len(value) > 31:
        raise ExcelMcpError("Nome de planilha deve ter no maximo 31 caracteres.")
    return value


def _rgb_to_excel(color: str) -> int:
    value = color.strip().lstrip("#")
    if len(value) != 6:
        raise ExcelMcpError("Cores devem usar formato hexadecimal #RRGGBB.")
    try:
        red = int(value[0:2], 16)
        green = int(value[2:4], 16)
        blue = int(value[4:6], 16)
    except ValueError as exc:
        raise ExcelMcpError("Cores devem usar formato hexadecimal #RRGGBB.") from exc
    return red + (green * 256) + (blue * 65536)


def _coerce_matrix(values: Sequence[Sequence[Any]] | Sequence[Any]) -> list[list[Any]]:
    if isinstance(values, (str, bytes)):
        matrix = [[values.decode() if isinstance(values, bytes) else values]]
    else:
        try:
            if not values:
                raise ExcelMcpError("values nao pode ser vazio.")
            first = values[0]  # type: ignore[index]
        except TypeError as exc:
            raise ExcelMcpError("values deve ser uma sequencia de celulas ou linhas.") from exc

        if isinstance(first, (list, tuple)):
            matrix = [list(row) for row in values]  # type: ignore[arg-type]
        else:
            matrix = [list(values)]  # type: ignore[list-item]

    if not matrix or not matrix[0]:
        raise ExcelMcpError("values deve conter pelo menos uma celula.")

    width = len(matrix[0])
    for row in matrix:
        if len(row) != width:
            raise ExcelMcpError("Todas as linhas em values devem ter o mesmo numero de colunas.")
    return matrix


def _matrix_is_blank(matrix: Sequence[Sequence[Any]]) -> bool:
    return all(cell is None or str(cell).strip() == "" for row in matrix for cell in row)


def _range_dimensions(range_obj: Any) -> tuple[int, int]:
    return int(range_obj.Rows.Count), int(range_obj.Columns.Count)


def _address(range_obj: Any) -> str:
    return str(range_obj.Address).replace("$", "")


def _range_from_top_left(sheet: Any, start_cell: str, rows: int, cols: int) -> Any:
    if rows < 1 or cols < 1:
        raise ExcelMcpError("Intervalos devem ter pelo menos uma linha e uma coluna.")
    top_left = sheet.Range(_require_non_empty(start_cell, "start_cell"))
    if top_left.Cells.Count != 1:
        raise ExcelMcpError("start_cell deve referenciar uma unica celula.")
    bottom_right = sheet.Cells(top_left.Row + rows - 1, top_left.Column + cols - 1)
    return sheet.Range(top_left, bottom_right)


def _com_value_to_matrix(value: Any, rows: int, cols: int) -> list[list[Any]]:
    if rows == 1 and cols == 1:
        return [[value]]
    if rows == 1:
        return [list(value)]
    if cols == 1:
        return [[item] for item in value]
    return [list(row) for row in value]


class ExcelController:
    def __init__(self, settings: RuntimeSettings) -> None:
        self._settings = settings
        self._lock = asyncio.Lock()
        self._pythoncom: Any | None = None
        self._win32: Any | None = None
        self._excel: Any | None = None
        self._workbook: Any | None = None
        self._com_initialized = False

    async def create_workbook(self, visible: bool | None = None) -> dict[str, Any]:
        async with self._lock:
            excel = self._ensure_excel(visible=visible)
            workbook = excel.Workbooks.Add()
            self._workbook = workbook
            return self._workbook_state()

    async def open_workbook(self, path: str, visible: bool | None = None, read_only: bool = False) -> dict[str, Any]:
        async with self._lock:
            workbook_path = _normalize_path(path, default_dir=self._settings.default_output_dir)
            if workbook_path is None or not workbook_path.exists():
                raise ExcelMcpError(f"Arquivo nao encontrado: {path}")
            excel = self._ensure_excel(visible=visible)
            self._workbook = excel.Workbooks.Open(str(workbook_path), ReadOnly=read_only)
            return self._workbook_state()

    async def save_workbook(self, path: str | None = None) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            output_path = _normalize_path(path, default_dir=self._settings.default_output_dir)
            if output_path is None:
                if not workbook.Path:
                    output_path = _normalize_path("workbook.xlsx", default_dir=self._settings.default_output_dir)
                else:
                    workbook.Save()
                    return self._workbook_state(saved=True)
            assert output_path is not None
            output_path.parent.mkdir(parents=True, exist_ok=True)
            workbook.SaveAs(str(output_path), FileFormat=_extension_format(output_path))
            return self._workbook_state(saved=True)

    async def close_workbook(self, save: bool = True) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            name = workbook.Name
            workbook.Close(SaveChanges=save)
            self._workbook = None
            return {"closed": True, "workbook": name, "saved": save}

    async def quit_excel(self) -> dict[str, Any]:
        async with self._lock:
            if self._workbook is not None:
                self._workbook.Close(SaveChanges=True)
                self._workbook = None
            if self._excel is not None:
                self._excel.Quit()
                self._excel = None
            self._uninitialize_com()
            return {"quit": True}

    async def workbook_info(self) -> dict[str, Any]:
        async with self._lock:
            return self._workbook_state()

    async def list_sheets(self) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            active_sheet = workbook.ActiveSheet.Name
            return {
                "workbook": workbook.Name,
                "active_sheet": active_sheet,
                "sheets": [sheet.Name for sheet in workbook.Worksheets],
            }

    async def add_sheet(self, sheet_name: str, activate: bool = True) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            name = _sheet_name(sheet_name)
            self._assert_sheet_absent(workbook, name)
            sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
            sheet.Name = name
            if activate:
                sheet.Activate()
            return {"created": True, "sheet": sheet.Name, "sheets": [ws.Name for ws in workbook.Worksheets]}

    async def delete_sheet(self, sheet_name: str) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            if workbook.Worksheets.Count <= 1:
                raise ExcelMcpError("Nao e possivel excluir a unica planilha do arquivo.")
            sheet = self._get_sheet(workbook, sheet_name)
            excel = self._ensure_excel()
            previous_alerts = excel.DisplayAlerts
            excel.DisplayAlerts = False
            try:
                sheet.Delete()
            finally:
                excel.DisplayAlerts = previous_alerts
            return {"deleted": True, "sheet": sheet_name, "sheets": [ws.Name for ws in workbook.Worksheets]}

    async def activate_sheet(self, sheet_name: str) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            sheet.Activate()
            return {"active_sheet": sheet.Name}

    async def write_range(
        self,
        sheet_name: str,
        start_cell: str,
        values: Sequence[Sequence[Any]] | Sequence[Any],
        autofit: bool = True,
    ) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            matrix = _coerce_matrix(values)
            rows = len(matrix)
            cols = len(matrix[0])
            target = _range_from_top_left(sheet, start_cell, rows, cols)
            target.Value = matrix[0][0] if rows == 1 and cols == 1 else tuple(tuple(row) for row in matrix)
            written = _com_value_to_matrix(target.Value, rows, cols)
            if _matrix_is_blank(written) and not _matrix_is_blank(matrix):
                raise ExcelMcpError(f"Excel nao confirmou a escrita no intervalo {_address(target)}.")
            if autofit:
                target.EntireColumn.AutoFit()
            return {
                "written": True,
                "sheet": sheet.Name,
                "range": _address(target),
                "rows": rows,
                "columns": cols,
                "values": written,
            }

    async def read_range(self, sheet_name: str, address: str) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            cell_range = sheet.Range(_require_non_empty(address, "address"))
            rows, cols = _range_dimensions(cell_range)
            total_cells = rows * cols
            if total_cells > self._settings.max_read_cells:
                raise ExcelMcpError(
                    f"Range muito grande ({total_cells} celulas). Limite atual: {self._settings.max_read_cells}."
                )
            return {
                "sheet": sheet.Name,
                "range": _address(cell_range),
                "rows": rows,
                "columns": cols,
                "values": _com_value_to_matrix(cell_range.Value, rows, cols),
            }

    async def clear_range(self, sheet_name: str, address: str, clear_formats: bool = False) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            cell_range = sheet.Range(_require_non_empty(address, "address"))
            if clear_formats:
                cell_range.Clear()
            else:
                cell_range.ClearContents()
            return {"cleared": True, "sheet": sheet.Name, "range": _address(cell_range)}

    async def set_formula(self, sheet_name: str, cell: str, formula: str, autofit: bool = True) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            target = sheet.Range(_require_non_empty(cell, "cell"))
            normalized_formula = _require_non_empty(formula, "formula")
            if not normalized_formula.startswith("="):
                normalized_formula = f"={normalized_formula}"
            target.Formula = normalized_formula
            if autofit:
                target.EntireColumn.AutoFit()
            return {"formula_set": True, "sheet": sheet.Name, "cell": _address(target), "formula": normalized_formula}

    async def format_range(
        self,
        sheet_name: str,
        address: str,
        bold: bool | None = None,
        italic: bool | None = None,
        font_size: float | None = None,
        font_color: str | None = None,
        fill_color: str | None = None,
        number_format: str | None = None,
        horizontal_align: Literal["general", "left", "center", "right"] | None = None,
        vertical_align: Literal["top", "center", "bottom"] | None = None,
        borders: bool = False,
        autofit: bool = False,
    ) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            cell_range = sheet.Range(_require_non_empty(address, "address"))
            if bold is not None:
                cell_range.Font.Bold = bold
            if italic is not None:
                cell_range.Font.Italic = italic
            if font_size is not None:
                if font_size <= 0:
                    raise ExcelMcpError("font_size deve ser maior que zero.")
                cell_range.Font.Size = font_size
            if font_color is not None:
                cell_range.Font.Color = _rgb_to_excel(font_color)
            if fill_color is not None:
                cell_range.Interior.Color = _rgb_to_excel(fill_color)
            if number_format is not None:
                cell_range.NumberFormat = number_format
            if horizontal_align is not None:
                cell_range.HorizontalAlignment = H_ALIGN[horizontal_align]
            if vertical_align is not None:
                cell_range.VerticalAlignment = V_ALIGN[vertical_align]
            if borders:
                for border_index in range(7, 13):
                    border = cell_range.Borders(border_index)
                    border.LineStyle = BORDER_LINE_STYLE_CONTINUOUS
                    border.Weight = BORDER_WEIGHT_THIN
                    border.ColorIndex = COLOR_INDEX_AUTOMATIC
            if autofit:
                cell_range.EntireColumn.AutoFit()
                cell_range.EntireRow.AutoFit()
            return {"formatted": True, "sheet": sheet.Name, "range": _address(cell_range)}

    async def create_table(
        self,
        sheet_name: str,
        address: str,
        table_name: str | None = None,
        style: str = "TableStyleMedium2",
    ) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            requested_address = _require_non_empty(address, "address")
            cell_range = sheet.Range(requested_address)
            rows, cols = _range_dimensions(cell_range)
            if rows < 2 or cols < 1:
                raise ExcelMcpError("A tabela deve ter cabecalho e pelo menos uma linha de dados.")

            values = _com_value_to_matrix(cell_range.Value, rows, cols)
            if _matrix_is_blank(values):
                raise ExcelMcpError(
                    f"O intervalo {requested_address} esta vazio. Escreva os dados com excel_write_range antes de criar a tabela."
                )
            if _matrix_is_blank([values[0]]):
                raise ExcelMcpError("A primeira linha da tabela deve conter cabecalhos nao vazios.")
            if _matrix_is_blank(values[1:]):
                raise ExcelMcpError("A tabela deve conter pelo menos uma linha de dados nao vazia.")

            list_object = sheet.ListObjects.Add(XL_SRC_RANGE, cell_range, None, XL_YES)
            if table_name:
                list_object.Name = _require_non_empty(table_name, "table_name")
            list_object.TableStyle = _require_non_empty(style, "style")
            data_body = list_object.DataBodyRange
            data_body_values = [] if data_body is None else _com_value_to_matrix(
                data_body.Value,
                data_body.Rows.Count,
                data_body.Columns.Count,
            )
            if not data_body_values or _matrix_is_blank(data_body_values):
                raise ExcelMcpError(
                    f"A tabela {list_object.Name} foi criada sem dados. Verifique se o intervalo {requested_address} contem linhas abaixo do cabecalho."
                )
            actual_range = list_object.Range
            return {
                "table_created": True,
                "sheet": sheet.Name,
                "table": list_object.Name,
                "range": _address(actual_range),
                "rows": actual_range.Rows.Count,
                "columns": actual_range.Columns.Count,
                "data_rows": data_body.Rows.Count,
                "data_values": data_body_values,
            }

    async def create_chart(
        self,
        sheet_name: str,
        data_range: str,
        chart_type: Literal["column", "bar", "line", "pie", "area", "scatter"] = "column",
        title: str | None = None,
        left: float = 300,
        top: float = 20,
        width: float = 480,
        height: float = 300,
    ) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            source = sheet.Range(_require_non_empty(data_range, "data_range"))
            chart_object = sheet.ChartObjects().Add(left, top, width, height)
            chart = chart_object.Chart
            chart.SetSourceData(source)
            chart.ChartType = CHART_TYPES[chart_type]
            if title:
                chart.HasTitle = True
                chart.ChartTitle.Text = title
            return {
                "chart_created": True,
                "sheet": sheet.Name,
                "chart_name": chart_object.Name,
                "chart_type": chart_type,
                "source_range": _address(source),
            }

    async def autofit(self, sheet_name: str, address: str | None = None) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            sheet = self._get_sheet(workbook, sheet_name)
            target = sheet.UsedRange if address is None else sheet.Range(_require_non_empty(address, "address"))
            target.EntireColumn.AutoFit()
            target.EntireRow.AutoFit()
            return {"autofit": True, "sheet": sheet.Name, "range": _address(target)}

    async def export_pdf(self, output_path: str, sheet_name: str | None = None) -> dict[str, Any]:
        async with self._lock:
            workbook = self._require_workbook()
            pdf_path = _normalize_path(output_path, default_dir=self._settings.default_output_dir)
            if pdf_path is None:
                raise ExcelMcpError("output_path e obrigatorio.")
            if pdf_path.suffix.lower() != ".pdf":
                raise ExcelMcpError("output_path deve terminar com .pdf.")
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            if sheet_name:
                sheet = self._get_sheet(workbook, sheet_name)
                sheet.ExportAsFixedFormat(EXPORT_FORMAT_PDF, str(pdf_path))
            else:
                workbook.ExportAsFixedFormat(EXPORT_FORMAT_PDF, str(pdf_path))
            return {"exported": True, "path": str(pdf_path), "scope": sheet_name or "workbook"}

    def _ensure_com(self) -> None:
        if self._pythoncom is None or self._win32 is None:
            self._pythoncom, self._win32 = _load_pywin32()
        if not self._com_initialized:
            self._pythoncom.CoInitialize()
            self._com_initialized = True

    def _uninitialize_com(self) -> None:
        if self._com_initialized and self._pythoncom is not None:
            self._pythoncom.CoUninitialize()
        self._com_initialized = False

    def _ensure_excel(self, visible: bool | None = None) -> Any:
        self._ensure_com()
        if self._excel is None:
            assert self._win32 is not None
            self._excel = self._win32.DispatchEx("Excel.Application")
            self._excel.DisplayAlerts = self._settings.display_alerts
        self._excel.Visible = self._settings.visible if visible is None else visible
        return self._excel

    def _require_workbook(self) -> Any:
        if self._workbook is None:
            raise ExcelMcpError("Nenhum workbook ativo. Use excel_create_workbook ou excel_open_workbook primeiro.")
        return self._workbook

    def _workbook_state(self, saved: bool = False) -> dict[str, Any]:
        workbook = self._require_workbook()
        return {
            "workbook": workbook.Name,
            "path": str(Path(workbook.FullName).resolve()) if workbook.Path else None,
            "saved": saved,
            "active_sheet": workbook.ActiveSheet.Name,
            "sheets": [sheet.Name for sheet in workbook.Worksheets],
        }

    def _get_sheet(self, workbook: Any, sheet_name: str) -> Any:
        name = _sheet_name(sheet_name)
        try:
            return workbook.Worksheets(name)
        except Exception as exc:  # pywin32 raises com_error; keep import lazy.
            raise ExcelMcpError(f"Planilha nao encontrada: {name}") from exc

    def _assert_sheet_absent(self, workbook: Any, sheet_name: str) -> None:
        try:
            workbook.Worksheets(sheet_name)
        except Exception:
            return
        raise ExcelMcpError(f"Ja existe uma planilha chamada: {sheet_name}")


SETTINGS = RuntimeSettings.from_env()


@asynccontextmanager
async def app_lifespan(_: FastMCP) -> AsyncIterator[AppContext]:
    controller = ExcelController(SETTINGS)
    try:
        yield AppContext(controller=controller)
    finally:
        await controller.quit_excel()


mcp = FastMCP(
    "excel-automation",
    instructions=(
        "Ferramentas para controlar Microsoft Excel via pywin32/COM no Windows. "
        "Use para criar relatorios, planilhas financeiras, tabelas, formulas, graficos e PDFs. "
        "Sempre crie ou abra um workbook antes de manipular planilhas."
    ),
    json_response=True,
    lifespan=app_lifespan,
)


def _controller_from_ctx(ctx: Context[ServerSession, AppContext]) -> ExcelController:
    return ctx.request_context.lifespan_context.controller


@mcp.tool()
async def excel_create_workbook(ctx: Context[ServerSession, AppContext], visible: bool | None = None) -> dict[str, Any]:
    """Create a new Excel workbook and make it the active workbook."""
    await ctx.info("Creating a new Excel workbook")
    return await _controller_from_ctx(ctx).create_workbook(visible=visible)


@mcp.tool()
async def excel_open_workbook(
    path: str,
    ctx: Context[ServerSession, AppContext],
    visible: bool | None = None,
    read_only: bool = False,
) -> dict[str, Any]:
    """Open an existing workbook and make it active."""
    await ctx.info(f"Opening workbook: {path}")
    return await _controller_from_ctx(ctx).open_workbook(path, visible=visible, read_only=read_only)


@mcp.tool()
async def excel_save_workbook(ctx: Context[ServerSession, AppContext], path: str | None = None) -> dict[str, Any]:
    """Save the active workbook. Provide path to SaveAs a specific file."""
    return await _controller_from_ctx(ctx).save_workbook(path=path)


@mcp.tool()
async def excel_close_workbook(ctx: Context[ServerSession, AppContext], save: bool = True) -> dict[str, Any]:
    """Close the active workbook."""
    return await _controller_from_ctx(ctx).close_workbook(save=save)


@mcp.tool()
async def excel_quit(ctx: Context[ServerSession, AppContext]) -> dict[str, Any]:
    """Close Excel controlled by this MCP server."""
    return await _controller_from_ctx(ctx).quit_excel()


@mcp.tool()
async def excel_workbook_info(ctx: Context[ServerSession, AppContext]) -> dict[str, Any]:
    """Return active workbook metadata."""
    return await _controller_from_ctx(ctx).workbook_info()


@mcp.tool()
async def excel_list_sheets(ctx: Context[ServerSession, AppContext]) -> dict[str, Any]:
    """List sheets in the active workbook."""
    return await _controller_from_ctx(ctx).list_sheets()


@mcp.tool()
async def excel_add_sheet(
    sheet_name: str,
    ctx: Context[ServerSession, AppContext],
    activate: bool = True,
) -> dict[str, Any]:
    """Add a worksheet to the active workbook."""
    return await _controller_from_ctx(ctx).add_sheet(sheet_name, activate=activate)


@mcp.tool()
async def excel_delete_sheet(sheet_name: str, ctx: Context[ServerSession, AppContext]) -> dict[str, Any]:
    """Delete a worksheet from the active workbook."""
    return await _controller_from_ctx(ctx).delete_sheet(sheet_name)


@mcp.tool()
async def excel_activate_sheet(sheet_name: str, ctx: Context[ServerSession, AppContext]) -> dict[str, Any]:
    """Activate an existing worksheet."""
    return await _controller_from_ctx(ctx).activate_sheet(sheet_name)


@mcp.tool()
async def excel_write_range(
    sheet_name: str,
    start_cell: str,
    values: list[list[Any]] | list[Any],
    ctx: Context[ServerSession, AppContext],
    autofit: bool = True,
) -> dict[str, Any]:
    """Write a rectangular matrix or one-dimensional row into a sheet starting at start_cell."""
    return await _controller_from_ctx(ctx).write_range(sheet_name, start_cell, values, autofit=autofit)


@mcp.tool()
async def excel_read_range(sheet_name: str, address: str, ctx: Context[ServerSession, AppContext]) -> dict[str, Any]:
    """Read values from an Excel range address, limited by EXCEL_MCP_MAX_READ_CELLS."""
    return await _controller_from_ctx(ctx).read_range(sheet_name, address)


@mcp.tool()
async def excel_clear_range(
    sheet_name: str,
    address: str,
    ctx: Context[ServerSession, AppContext],
    clear_formats: bool = False,
) -> dict[str, Any]:
    """Clear values or values plus formats from a range."""
    return await _controller_from_ctx(ctx).clear_range(sheet_name, address, clear_formats=clear_formats)


@mcp.tool()
async def excel_set_formula(
    sheet_name: str,
    cell: str,
    formula: str,
    ctx: Context[ServerSession, AppContext],
    autofit: bool = True,
) -> dict[str, Any]:
    """Set an Excel formula in a cell. Leading '=' is optional."""
    return await _controller_from_ctx(ctx).set_formula(sheet_name, cell, formula, autofit=autofit)


@mcp.tool()
async def excel_format_range(
    sheet_name: str,
    address: str,
    ctx: Context[ServerSession, AppContext],
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_color: str | None = None,
    fill_color: str | None = None,
    number_format: str | None = None,
    horizontal_align: Literal["general", "left", "center", "right"] | None = None,
    vertical_align: Literal["top", "center", "bottom"] | None = None,
    borders: bool = False,
    autofit: bool = False,
) -> dict[str, Any]:
    """Apply common formatting to a range. Colors use #RRGGBB."""
    return await _controller_from_ctx(ctx).format_range(
        sheet_name,
        address,
        bold=bold,
        italic=italic,
        font_size=font_size,
        font_color=font_color,
        fill_color=fill_color,
        number_format=number_format,
        horizontal_align=horizontal_align,
        vertical_align=vertical_align,
        borders=borders,
        autofit=autofit,
    )


@mcp.tool()
async def excel_create_table(
    sheet_name: str,
    address: str,
    ctx: Context[ServerSession, AppContext],
    table_name: str | None = None,
    style: str = "TableStyleMedium2",
) -> dict[str, Any]:
    """Create an Excel structured table from a range whose first row contains headers."""
    return await _controller_from_ctx(ctx).create_table(sheet_name, address, table_name=table_name, style=style)


@mcp.tool()
async def excel_create_chart(
    sheet_name: str,
    data_range: str,
    ctx: Context[ServerSession, AppContext],
    chart_type: Literal["column", "bar", "line", "pie", "area", "scatter"] = "column",
    title: str | None = None,
    left: float = 300,
    top: float = 20,
    width: float = 480,
    height: float = 300,
) -> dict[str, Any]:
    """Create a chart object from a data range."""
    return await _controller_from_ctx(ctx).create_chart(
        sheet_name,
        data_range,
        chart_type=chart_type,
        title=title,
        left=left,
        top=top,
        width=width,
        height=height,
    )


@mcp.tool()
async def excel_autofit(
    sheet_name: str,
    ctx: Context[ServerSession, AppContext],
    address: str | None = None,
) -> dict[str, Any]:
    """Auto-fit rows and columns for a range or the used range."""
    return await _controller_from_ctx(ctx).autofit(sheet_name, address=address)


@mcp.tool()
async def excel_export_pdf(
    output_path: str,
    ctx: Context[ServerSession, AppContext],
    sheet_name: str | None = None,
) -> dict[str, Any]:
    """Export the active workbook or a single sheet to PDF."""
    return await _controller_from_ctx(ctx).export_pdf(output_path, sheet_name=sheet_name)


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
