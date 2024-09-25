from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = collect_data_files('pycel')
hiddenimports = collect_submodules('pycel')