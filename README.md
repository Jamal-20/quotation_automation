

# AutoQ Genz📝 

An automated Windows desktop application that generates **Excel and PDF quotations** from predefined templates.  
Designed to **optimize tender and quotation workflows**, reduce processing time, improve accuracy, and ensure consistent, compliant outputs for **NUPCO, government bids, and enterprise procurement**.

[![Python](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/)
[![Excel](https://img.shields.io/badge/excel-2016+-green.svg)](https://www.microsoft.com/excel)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

![App Screenshot](screenshots/app.png)

---

## 🔄 System Workflow

---

## Overview

**AutoQ Genz** is a purpose-built desktop application created to streamline and standardize the preparation of tender and sales quotations.

Instead of manually editing quotations line by line, the application automatically reads structured data from Excel input files and injects it into professionally formatted quotation templates. The system then produces both:

- **Editable Excel quotations**
- **Print-ready, submission-quality PDF files**

---

## Business Value 🎯 

- ⏱️ **Faster turnaround** for tender and quotation preparation  
- 🎯 **Improved accuracy** by eliminating repetitive manual data entry  
- 📄 **Consistent formatting** aligned with corporate branding and tender requirements  
- 📊 **Scalable processing** for large or batch quotations  
- 👥 **Higher productivity** for sales, procurement, and tender teams  

---

## Key Features ✨

- 🖱️ **Drag & Drop Interface** 
  Easily drop Excel input files into the application

- 📊 **Batch Processing**
  Generate dozens or thousand of quotations in a single run

- 📄 **Dual Output Format**
  Automatically creates both `.xlsx` and `.pdf` files

- ⚡**Real-Time Progress Tracking**
  Visual progress bar with detailed processing logs

- 🔒 **Safe Exit Protection**
  Exit confirmation during active processing

---

### Required Python Libraries
Installed automatically via `Installation.py`

---


## How to Run the Application 🚀 

### Step 1: Prepare Your Files

#### Input File (`Input.xlsx`)
Contains the quotation line items that will be injected into the template.

#### Template File (`Template.xlsx`)
Pre-formatted quotation template with branding, formulas, and layout.

---

### Step 2: Run the Application

```bash
python quotation_generator.py
```
---
### Step 3: Output Files

Generated files will be saved automatically to:
```
Excel Output/ – Editable Excel quotations

PDF Output/ – Final, printable PDF quotations
```
---

### Step 4: Building a Standalone Executable

To create a Windows .exe file:
```
pip install pyinstaller
pyinstaller --onefile --windowed --name "SAMCO_Quotation_Generator" --icon=app.ico quotation_generator.py
```
---
### License📄 

This project is licensed under the MIT License.

---
### Support📞

For support or inquiries:

Email: jamaluddin.jk7@gmail.com

---
