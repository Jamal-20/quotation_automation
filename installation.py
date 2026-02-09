echo ================================================
echo    SAMCO Quotation Generator - Installation Start
echo ================================================

## Prerequisites: 
# Python 3.8+, Microsoft Excel (for PDF)

# Packages Installation 
python -m pip install --upgrade pip

pip install PyQt6==6.6.1   # for GUI framework
pip install pandas==2.1.4   # for data processing
pip install openpyxl==3.1.2  # for Excel handling
pip install pywin32==306 # Excel COM automation


# Verify installations
python -c "import PyQt6; print('PyQt6: OK')" 
python -c "import pandas; print('pandas: OK')" 
python -c "import openpyxl; print('openpyxl: OK')" 
python -c "import win32com; print('pywin32: OK')" 


# As combined install and Check
pip install PyQt6 pandas openpyxl pywin32
python --version
python -c "import PyQt6, pandas, openpyxl, win32com"

echo ================================================
echo    INSTALLATION COMPLETED SUCCESSFULLY!

