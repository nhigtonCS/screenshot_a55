@echo off


:start
cls

set python_ver=39

python ./get-pip.py

cd \
cd \python%python_ver%\Scripts\
pip install openpyxl
pip install pillow
pip install tk
pip install easygui
pip install selenium
pip install choco
pip install opencv-python
pip install numpy
pip install chromedriver_autoinstaller

pause
exit