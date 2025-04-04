from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': []}

base = 'gui'

executables = [
    Executable('GUI.py', base=base, target_name = 'GUI')
]

setup(name='GUI',
      version = '1',
      description = 'Interfaz de configuración del monitor de correo',
      options = {'build_exe': build_options},
      executables = executables)
