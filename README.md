# Authority-Letter-Excel-to-Word-Using-Python

A Python application that generates authority letters from Excel data using Word templates with a user-friendly interface.

## Overview

This application automates the process of creating personalized authority letters based on Excel data. It reads information from an Excel spreadsheet and merges it into a Word document template, creating individual files for each record. The program organizes the output files in a structured folder hierarchy based on insurance company, policy number, and insured name.

## Features

- **Mail Merge Automation**: Convert Excel data to personalized Word documents
- **Smart Document Organization**: Automatically creates a structured folder hierarchy
- **Dynamic Placeholder Replacement**: Replace placeholders in Word templates with Excel data
- **Date Handling**: Formats dates correctly and inserts current date where needed
- **User-Friendly Interface**: Intuitive GUI with file selection dialog
- **Progress Tracking**: Real-time progress bar and detailed logging
- **Error Handling**: Robust error detection and reporting

## Screenshots

![final_reordered_grid_image](https://github.com/user-attachments/assets/53ea8444-3470-4f4d-980b-ca03aaa32f2c)


## Requirements

- Python 3.x
- pandas
- python-docx
- tkinter (included with Python)
- ttkbootstrap (for MediNIA.py version)

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/authority-letter-excel-to-word.git
   cd authority-letter-excel-to-word
   ```

2. Install required packages:
   ```
   pip install pandas python-docx ttkbootstrap
   ```

## Usage

### Running the Application

1. Launch the application by running either version:
   ```
   python AuthorityLetter.py
   ```
   or
   ```
   python MediNIA.py
   ```

2. Select the required files:
   - Word Template (.docx file) containing your letter with placeholders
   - Excel Data File (.xlsx file) containing your records
   - Output Folder where generated documents will be saved

3. Click "Start Mail Merge" to begin processing

### Creating Templates

In your Word templates, use placeholders enclosed in square brackets that match your Excel column names:

- Example: `[Insured Name]`, `[Policy No.]`, `[Insurance Company]`
- Use `[Current Date]` to automatically insert the current date

### Excel Data Format

Your Excel file should include columns that match the placeholders in your template. The program specifically requires:
- "Insurance Company" - Used for top-level folder organization
- "Policy No." - Used for subfolder naming and document identification
- "Insured Name" - Used for subfolder naming and document identification

Additional columns can be included and used as placeholders in your template.

### Output Structure

The program creates the following folder structure:
```
Output Folder
  └── Insurance Company Name
       └── Policy_Number_Insured_Name
            └── Document_Policy_Number_Insured_Name.docx
```

## Project Versions

The repository includes two versions of the application:

1. **AuthorityLetter.py**: Standard version using basic tkinter
2. **MediNIA.py**: Enhanced version with improved UI using ttkbootstrap

## How It Works

1. The application reads data from the Excel file into a pandas DataFrame
2. For each row in the DataFrame:
   - Creates the necessary folder structure
   - Opens the Word template
   - Replaces all placeholders with corresponding values from the Excel data
   - Saves the resulting document in the appropriate folder
3. Progress is tracked and displayed in real-time
