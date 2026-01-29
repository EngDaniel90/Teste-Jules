import flet as ft
import asyncio

async def main(page: ft.Page):
    print(f"page.go is async: {asyncio.iscoroutinefunction(page.go)}")
    print(f"page.update is async: {asyncio.iscoroutinefunction(page.update)}")
    await page.window_destroy()

if __name__ == "__main__":
    ft.app(target=main, port=8551)
