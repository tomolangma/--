from cx_Freeze import setup, Executable

setup(
    name="my_program",
    version="1.0",
    description="My Program",
    executables=[Executable("main.pyw", base="Win32GUI")],
)