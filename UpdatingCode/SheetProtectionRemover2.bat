@echo ' NOTEBOOK_PATH=C:\Users\USER\SheetProtectionRemover.ipynb
@echo ' NOTEBOOK_PATH Sould be C drive.
@echo ' Selectable Excel file should be in C drive


@echo off

rem Set the path to your Miniconda installation
set CONDA_PATH=C:\Users\USER\miniconda3

rem Set the name of your virtual environment
set ENV_NAME=myenv

rem Set the path to your Jupyter Notebook file

set NOTEBOOK_PATH=C:\Users\USER\pypy\Excel Sheet Protection Remover by Ripon\SheetProtectionRemover.ipynb
rem Activate the virtual environment
call "%CONDA_PATH%\Scripts\activate.bat" %ENV_NAME%

rem Run the Jupyter Notebook
jupyter nbconvert --execute --to notebook --inplace "%NOTEBOOK_PATH%"

rem Deactivate the virtual environment
call "%CONDA_PATH%\Scripts\deactivate.bat"

rem Show a message box indicating completion
msg * "File Sheet Protection Removed!"
