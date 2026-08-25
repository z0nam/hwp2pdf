"""Collect the TkDnD runtime files required by tkinterdnd2."""

from PyInstaller.utils.hooks import collect_data_files, copy_metadata


datas = collect_data_files("tkinterdnd2") + copy_metadata("tkinterdnd2")
