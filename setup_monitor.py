from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {'packages': [], 'excludes': []}

base = 'console'

executables = [
    Executable('main.py', base=base, target_name = 'pcountermailmonitorgnsys')
]

setup(name='mail_monitor',
      version = '1',
      description = 'Servicio del monitoreo de correo',
      options = {'build_exe': build_options},
      executables = executables)
