# EXTRACTOR: Employee Timesheet and Project Analysis Tool

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

A Python-based ETL solution for automating employee timesheet analysis and project status reporting from Excel data sources.

## 📖 Repository Description
**EXTRACTOR** is a robust data processing tool that transforms raw employee logs and project data into actionable insights. Built with **Python**, **Pandas**, and **Tkinter**, this application features:
- Automated Excel data extraction and transformation
- Intelligent status tracking (Approved/Missing/Rejected)
- GUI-powered workflow for non-technical users
- Executable packaging for Windows environments

## 🚀 Features
- **Data Extraction**: Processes Cisco-generated RDMS exports and employee lists
- **Smart Transformation**:
  - Name standardization and title removal
  - Dynamic week-of-year calculations
  - Multi-dimensional status categorization
- **Visual Reporting**:
  - Color-coded Excel output
  - Auto-adjusted column widths
  - Merged project grouping
- **Desktop GUI**:
  - File picker interface
  - Sheet selection dropdowns
  - Progress feedback

## ⚙️ Technologies Used
- **Core**: `Python 3.9+`
- **Data Processing**: `Pandas`, `XlsxWriter`, `NumPy`
- **GUI**: `Tkinter`, `PIL (Python Imaging Library)`
- **Packaging**: `PyInstaller`

## 📦 Installation
```bash
pip install pandas xlsxwriter numpy pillow
```

## 🖥️ Usage

1. **Select source Excel files**:  
   - Main dataset (GTE export)  
   - Employee list  
2. **Choose output path**  
3. **Click "Démarrer"** to generate the analysis report  

### Sample input/output:

#### **Input**:  
Employee logs (header row 14)  
```
Project Number | Employee Name | Hours | Status
```
#### **Output**:  
- Pivot table with Y/W columns  
- Status colors  
- Merged project groups  

---

## 🔨 Building the Executable

```bash
pyinstaller --onefile --windowed --noconsole \
            --add-data "logo.png:." \
            --add-data "fav.ico:." \
            --icon=fav.ico app.py
```

---

## 📚 Documentation

Find complete technical specifications and workflow diagrams here:  
**[EXTRACTOR Documentation](https://yoshq.notion.site/EXTRACTOR-af2258a285cd4313a1a4d609aa5b6d40?pvs=4)**
