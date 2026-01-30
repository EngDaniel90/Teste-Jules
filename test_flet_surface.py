import flet as ft
print(f"SURFACE: {ft.Colors.SURFACE}")
try:
    print(f"SURFACE_VARIANT: {ft.Colors.SURFACE_VARIANT}")
except AttributeError as e:
    print(f"SURFACE_VARIANT error: {e}")
