# 1. Create Virtual Environment 
python -m venv venv
venv\Scripts\activate

# 2. Install packages
pip install -r requirements.txt
python installation.py

# 3. Run
python quotation_generator.py

# 4. To create a standalone .exe:
# Install PyInstaller
pip install pyinstaller

# Build executable
pyinstaller --onefile --windowed --name "JAMAL_Quotation_Generator" --icon=app.ico quotation_generator.py
