from __future__ import annotations

import unittest

import excel_server


class FakeCell:
    def __init__(self, row: int, column: int) -> None:
        self.Row = row
        self.Column = column
        self.Cells = self
        self.Count = 1


class FakeRange:
    def __init__(self, address: str, row: int = 1, column: int = 1) -> None:
        self.Address = address
        self.Row = row
        self.Column = column
        self.Cells = self
        self.Count = 1


class FakeSheet:
    def __init__(self) -> None:
        self.calls: list[tuple[object, object | None]] = []

    def Range(self, first: str | FakeCell, second: FakeCell | None = None) -> FakeRange:
        self.calls.append((first, second))
        if isinstance(first, str):
            return FakeRange(first, 1, 1)
        self.assert_second_cell(second)
        return FakeRange(f"R{first.Row}C{first.Column}:R{second.Row}C{second.Column}", first.Row, first.Column)

    @staticmethod
    def assert_second_cell(second: FakeCell | None) -> None:
        if second is None:
            raise AssertionError("Expected second cell")

    def Cells(self, row: int, column: int) -> FakeCell:
        return FakeCell(row, column)


class ExcelServerTests(unittest.TestCase):
    def test_coerce_matrix_keeps_scalar_string_as_single_cell(self) -> None:
        self.assertEqual(excel_server._coerce_matrix("abc"), [["abc"]])

    def test_range_from_top_left_uses_two_cell_range_not_resize_property(self) -> None:
        sheet = FakeSheet()

        result = excel_server._range_from_top_left(sheet, "A1", 3, 2)

        self.assertEqual(result.Address, "R1C1:R3C2")
        self.assertEqual(len(sheet.calls), 2)
        self.assertEqual(sheet.calls[0], ("A1", None))
        self.assertIsInstance(sheet.calls[1][0], FakeRange)
        self.assertIsInstance(sheet.calls[1][1], FakeCell)

    def test_range_from_top_left_rejects_non_single_start_range(self) -> None:
        class MultiCellSheet(FakeSheet):
            def Range(self, first: str | FakeCell, second: FakeCell | None = None) -> FakeRange:
                range_obj = super().Range(first, second)
                if isinstance(first, str):
                    range_obj.Count = 2
                return range_obj

        with self.assertRaisesRegex(excel_server.ExcelMcpError, "unica celula"):
            excel_server._range_from_top_left(MultiCellSheet(), "A1:B2", 1, 1)


if __name__ == "__main__":
    unittest.main()
