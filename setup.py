from distutils.core import setup

import py2exe

data_files = [(".", ["DejaVuSansCondensed-Bold.ttf", "DejaVuSansCondensed.ttf"])]
setup(
    options={
        'py2exe': {
            'bundle_files': 3,
            'compressed': True
        }
    },
    data_files=data_files,
    windows=[{
            "script":"main.py",
            "icon_resources": [(1, "E-Mailer.ico")],
            "dest_base":"E-Mailer"
            }],
    zipfile=None
)
