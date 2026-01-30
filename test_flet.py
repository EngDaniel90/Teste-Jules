import flet as ft
import asyncio

async def main(page: ft.Page):
    page.add(ft.Text("Hello", color=ft.Colors.RED))
    print("Success")
    await asyncio.sleep(1)
    page.window_close()

if __name__ == "__main__":
    ft.run(main)
