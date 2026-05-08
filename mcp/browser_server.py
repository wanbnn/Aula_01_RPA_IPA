from __future__ import annotations

import asyncio
import os

# Configura o PLAYWRIGHT_BROWSERS_PATH hardcoded apontando pro local real do APPDATA
# já que o PyInstaller isola certos paths.
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(os.environ.get("LOCALAPPDATA", ""), "ms-playwright")

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from dataclasses import dataclass
from typing import Any, Literal
from urllib.parse import urlparse

from mcp.server.fastmcp import Context, FastMCP
from mcp.server.session import ServerSession
from playwright.async_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    TimeoutError as PlaywrightTimeoutError,
    async_playwright,
)


def _env_flag(name: str, default: bool) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() not in {"0", "false", "no", "off"}


def _truncate(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 3] + "..."


def _validate_url(url: str) -> str:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Use uma URL HTTP ou HTTPS valida.")
    return url


@dataclass(slots=True)
class RuntimeSettings:
    headless: bool
    timeout_ms: int
    viewport_width: int
    viewport_height: int
    max_snapshot_chars: int

    @classmethod
    def from_env(cls) -> "RuntimeSettings":
        return cls(
            headless=_env_flag("BROWSER_MCP_HEADLESS", True),
            timeout_ms=int(os.getenv("BROWSER_MCP_TIMEOUT_MS", "15000")),
            viewport_width=int(os.getenv("BROWSER_MCP_VIEWPORT_WIDTH", "1440")),
            viewport_height=int(os.getenv("BROWSER_MCP_VIEWPORT_HEIGHT", "900")),
            max_snapshot_chars=int(os.getenv("BROWSER_MCP_MAX_TEXT", "4000")),
        )


@dataclass(slots=True)
class AppContext:
    controller: "BrowserController"


class BrowserController:
    def __init__(self, settings: RuntimeSettings) -> None:
        self._settings = settings
        self._playwright: Playwright | None = None
        self._browser: Browser | None = None
        self._context: BrowserContext | None = None
        self._page: Page | None = None
        self._lock = asyncio.Lock()

    async def start(self) -> None:
        self._playwright = await async_playwright().start()

        launch_kwargs: dict[str, Any] = {
            "headless": self._settings.headless,
            "args": [
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-features=IsolateOrigins,site-per-process",
                "--disable-dev-shm-usage",
                "--no-sandbox",
            ],
            "ignore_default_args": ["--enable-automation"],
        }

        # Tenta usar o Chrome real instalado (muito menos detectável que Chromium)
        try:
            self._browser = await self._playwright.chromium.launch(
                channel="chrome",
                **launch_kwargs,
            )
        except Exception:
            self._browser = await self._playwright.chromium.launch(**launch_kwargs)

        self._context = await self._browser.new_context(
            viewport={
                "width": self._settings.viewport_width,
                "height": self._settings.viewport_height,
            },
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="pt-BR",
            timezone_id="America/Sao_Paulo",
            permissions=["geolocation"],
            extra_http_headers={
                "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
                "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
                "sec-ch-ua-mobile": "?0",
                "sec-ch-ua-platform": '"Windows"',
            },
        )

        # Stealth: remove rastros de automação antes de qualquer script da página
        await self._context.add_init_script("""
            // Oculta webdriver
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });

            // Plugins realistas
            Object.defineProperty(navigator, 'plugins', {
                get: () => {
                    const arr = [
                        { name: 'Chrome PDF Plugin', filename: 'internal-pdf-viewer', description: 'Portable Document Format' },
                        { name: 'Chrome PDF Viewer', filename: 'mhjfbmdgcfjbbpaeojofohoefgiehjai', description: '' },
                        { name: 'Native Client', filename: 'internal-nacl-plugin', description: '' },
                    ];
                    arr.__proto__ = PluginArray.prototype;
                    return arr;
                },
            });

            // Languages
            Object.defineProperty(navigator, 'languages', { get: () => ['pt-BR', 'pt', 'en-US', 'en'] });

            // Hardware realista
            Object.defineProperty(navigator, 'hardwareConcurrency', { get: () => 8 });
            Object.defineProperty(navigator, 'deviceMemory', { get: () => 8 });

            // Permissões — evita fingerprint via Notification.permission
            const originalQuery = window.navigator.permissions.query.bind(navigator.permissions);
            window.navigator.permissions.query = (parameters) =>
                parameters.name === 'notifications'
                    ? Promise.resolve({ state: Notification.permission })
                    : originalQuery(parameters);

            // Remove propriedade cdc_ injetada pelo chromedriver
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;

            // Chrome runtime simulado
            if (!window.chrome) {
                window.chrome = { runtime: {} };
            }
        """)

        self._context.set_default_timeout(self._settings.timeout_ms)
        self._page = await self._context.new_page()

    async def close(self) -> None:
        if self._context is not None:
            await self._context.close()
        if self._browser is not None:
            await self._browser.close()
        if self._playwright is not None:
            await self._playwright.stop()

    async def snapshot(self, max_chars: int | None = None) -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            return await self._build_snapshot(page, max_chars=max_chars)

    async def navigate(self, url: str, wait_until: str = "networkidle") -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            await page.goto(_validate_url(url), wait_until=wait_until)
            # Pequena pausa humana pós-carregamento
            await asyncio.sleep(1.2)
            return await self._build_snapshot(page)

    async def click(self, selector: str) -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            await page.locator(selector).first.click()
            await page.wait_for_load_state("networkidle")
            return await self._build_snapshot(page)

    async def type_text(
        self,
        selector: str,
        text: str,
        *,
        clear_first: bool,
        press_enter: bool,
    ) -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            field = page.locator(selector).first
            await field.click()
            if clear_first:
                await field.fill("")
            await field.type(text)
            if press_enter:
                await field.press("Enter")
                await page.wait_for_load_state("networkidle")
            return await self._build_snapshot(page)

    async def press(self, key: str) -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            await page.keyboard.press(key)
            await page.wait_for_load_state("networkidle")
            return await self._build_snapshot(page)

    async def wait_for(self, selector: str, state: str) -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            await page.locator(selector).first.wait_for(state=state)
            return await self._build_snapshot(page)

    async def extract_text(self, selector: str, max_chars: int | None = None) -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            text = await page.locator(selector).first.inner_text()
            return {
                "selector": selector,
                "text": _truncate(text.strip(), max_chars or self._settings.max_snapshot_chars),
                "url": page.url,
            }

    async def go_back(self) -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            await page.go_back(wait_until="networkidle")
            return await self._build_snapshot(page)

    async def inspect_elements(self, selector: str, limit: int) -> dict[str, Any]:
        async with self._lock:
            page = self._require_page()
            items = await page.locator(selector).evaluate_all(
                """
                (elements, maxItems) => elements.slice(0, maxItems).map((element) => ({
                    tag: element.tagName.toLowerCase(),
                    text: (element.innerText || element.textContent || '').trim().slice(0, 160),
                    id: element.id || null,
                    name: element.getAttribute('name'),
                    type: element.getAttribute('type'),
                    placeholder: element.getAttribute('placeholder'),
                    ariaLabel: element.getAttribute('aria-label'),
                    href: element.getAttribute('href'),
                }))
                """,
                limit,
            )
            return {
                "url": page.url,
                "selector": selector,
                "count": len(items),
                "elements": items,
            }

    def _require_page(self) -> Page:
        if self._page is None:
            raise RuntimeError("O navegador ainda nao foi inicializado.")
        return self._page

    async def _build_snapshot(self, page: Page, max_chars: int | None = None) -> dict[str, Any]:
        text = await page.locator("body").inner_text()
        title = await page.title()
        return {
            "url": page.url,
            "title": title,
            "text": _truncate(text.strip(), max_chars or self._settings.max_snapshot_chars),
        }


SETTINGS = RuntimeSettings.from_env()


@asynccontextmanager
async def app_lifespan(_: FastMCP) -> AsyncIterator[AppContext]:
    controller = BrowserController(SETTINGS)
    await controller.start()
    try:
        yield AppContext(controller=controller)
    finally:
        await controller.close()


mcp = FastMCP(
    "autonomous-browser",
    instructions=(
        "Ferramentas para navegar na web de forma autonoma com Playwright. "
        "Use seletores CSS validos para clicar, digitar, esperar e extrair conteudo."
    ),
    json_response=True,
    lifespan=app_lifespan,
)


def _controller_from_ctx(ctx: Context[ServerSession, AppContext]) -> BrowserController:
    return ctx.request_context.lifespan_context.controller


@mcp.tool()
async def browser_navigate(
    url: str,
    ctx: Context[ServerSession, AppContext],
    wait_until: Literal["load", "domcontentloaded", "networkidle"] = "networkidle",
) -> dict[str, Any]:
    """Open a page and return a text snapshot of the current state."""
    await ctx.info(f"Navigating to {url}")
    return await _controller_from_ctx(ctx).navigate(url, wait_until=wait_until)


@mcp.tool()
async def browser_get_page_state(
    ctx: Context[ServerSession, AppContext],
    max_chars: int = SETTINGS.max_snapshot_chars,
) -> dict[str, Any]:
    """Read the current page URL, title, and visible text."""
    return await _controller_from_ctx(ctx).snapshot(max_chars=max_chars)


@mcp.tool()
async def browser_click(selector: str, ctx: Context[ServerSession, AppContext]) -> dict[str, Any]:
    """Click the first element matching a CSS selector."""
    await ctx.info(f"Clicking selector: {selector}")
    return await _controller_from_ctx(ctx).click(selector)


@mcp.tool()
async def browser_type(
    selector: str,
    text: str,
    ctx: Context[ServerSession, AppContext],
    clear_first: bool = True,
    press_enter: bool = False,
) -> dict[str, Any]:
    """Type into the first element matching a CSS selector."""
    await ctx.info(f"Typing into selector: {selector}")
    return await _controller_from_ctx(ctx).type_text(
        selector,
        text,
        clear_first=clear_first,
        press_enter=press_enter,
    )


@mcp.tool()
async def browser_press_key(
    key: str,
    ctx: Context[ServerSession, AppContext],
) -> dict[str, Any]:
    """Press a keyboard key in the active page."""
    return await _controller_from_ctx(ctx).press(key)


@mcp.tool()
async def browser_wait_for(
    selector: str,
    ctx: Context[ServerSession, AppContext],
    state: Literal["attached", "detached", "hidden", "visible"] = "visible",
) -> dict[str, Any]:
    """Wait until a selector reaches the requested state."""
    return await _controller_from_ctx(ctx).wait_for(selector, state)


@mcp.tool()
async def browser_extract_text(
    selector: str,
    ctx: Context[ServerSession, AppContext],
    max_chars: int = SETTINGS.max_snapshot_chars,
) -> dict[str, Any]:
    """Extract visible text from the first element matching a CSS selector."""
    return await _controller_from_ctx(ctx).extract_text(selector, max_chars=max_chars)


@mcp.tool()
async def browser_go_back(ctx: Context[ServerSession, AppContext]) -> dict[str, Any]:
    """Navigate back to the previous page in history."""
    return await _controller_from_ctx(ctx).go_back()


@mcp.tool()
async def browser_inspect_elements(
    ctx: Context[ServerSession, AppContext],
    selector: str = "a, button, input, textarea, select",
    limit: int = 25,
) -> dict[str, Any]:
    """List candidate interactive elements to help choose selectors."""
    return await _controller_from_ctx(ctx).inspect_elements(selector, limit=limit)


def main() -> None:
    try:
        mcp.run(transport="stdio")
    except PlaywrightTimeoutError as exc:
        raise RuntimeError(f"Tempo esgotado em uma operacao do navegador: {exc}") from exc


if __name__ == "__main__":
    main()
