import sys
from cx_Freeze import setup, Executable
from controlegeral import resource_path

build_exe_options = {
    "zip_include_packages": ['altgraph==0.17.4','asttokens==2.4.1','auto-py-to-exe==2.42.0','bottle==0.12.25','bottle-websocket==0.2.9','certifi==2024.2.2','cffi==1.16.0','charset-normalizer==3.3.2','colorama==0.4.6','comm==0.2.1','cx-Logging==3.1.0','cx_Freeze==6.15.14','debugpy==1.8.0','decorator==5.1.1','defusedxml==0.7.1','docutils==0.20.1','Eel==0.16.0','exceptiongroup==1.2.0','executing==2.0.1','future==0.18.3','gevent==23.9.1','gevent-websocket==0.10.1','greenlet==3.0.3','idna==3.6','ipykernel==6.29.2','ipython==8.21.0','jedi==0.19.1',
                             'jupyter_client==8.6.0','jupyter_core==5.7.1','Kivy==2.3.0','kivy-deps.angle==0.4.0','kivy-deps.glew==0.3.1','kivy-deps.sdl2==0.7.0','Kivy-Garden==0.1.5','kivy-uix==1.0.0','lief==0.14.0','matplotlib-inline==0.1.6','nest-asyncio==1.6.0','numpy==1.26.4','odfpy==1.4.1','packaging==23.2','pandas==2.2.0','parso==0.8.3','pefile==2023.2.7','platformdirs==4.2.0','prompt-toolkit==3.0.43','psutil==5.9.8','pure-eval==0.2.2','pycparser==2.21','Pygments==2.17.2','pyinstaller==6.3.0','pyinstaller-hooks-contrib==2024.0','pyparsing==3.1.1',
                             'pypiwin32==223','python-dateutil==2.8.2','pytz==2024.1','pywin32==306','pywin32-ctypes==0.2.2','pyzmq==25.1.2','requests==2.31.0','six==1.16.0','stack-data==0.6.3','tornado==6.4','traitlets==5.14.1','tzdata==2023.4','urllib3==2.2.0','wcwidth==0.2.13','whichcraft==0.6.1','zope.event==5.0','zope.interface==6.1'],
    'include_files': ['aspirante.py', 'ControleGeral.kv', 'funcoes_aspirante.py', 'logoEN.png', 'logoM.png', 'teste.ods', 'registro.txt']
}


# base="Win32GUI" should be used only for Windows GUI app
base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="guifoo",
    version="0.1",
    description="My GUI application!",
    options={"build_exe": build_exe_options},
    executables=[Executable("controlegeral.py", base=base, icon='logoEN.ico')])