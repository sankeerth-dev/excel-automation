# PathAxiom Payroll Automation - Setup Guide

To move this automation to a new computer, follow these simple steps:

## Prerequisites
1. **Python**: Ensure Python (version 3.8 or newer) is installed on the new system.
   - You can download it from [python.org](https://www.python.org/downloads/).
   - **Important:** During installation, make sure to check the box that says **"Add Python to PATH"**.
2. **Microsoft Excel**: Since the script exports to PDF using your local Excel application, Microsoft Excel must be installed and activated on the new computer.

## Step 1: Copy the Files
Copy the entire `excel automation` folder to the new system. The essential files you must include are:
- `payroll_automation.py`
- `Attendance_Master.xlsx` (or your latest template)
- `logo.png`
- `Generate_Payroll.bat`
- `requirements.txt`

*(You do not need to copy the `.venv` folder, as it's better to create a fresh one on the new system.)*

## Step 2: Set Up the Environment
On the new system, open the Command Prompt (`cmd`) or PowerShell inside the folder where you copied the files and run these commands one by one to set up the engine:

1. **Create a Virtual Environment:**
   Run: `python -m venv .venv`

2. **Activate the Virtual Environment:**
   Run: `.venv\Scripts\activate`

3. **Install Dependencies:**
   Run: `pip install -r requirements.txt`

## Step 3: Run the Automation!
Once the setup is complete, you can generate your payrolls exactly how you do it currently:

- Just double-click the **`Generate_Payroll.bat`** file. 
- Ensure your `Attendance_Master.xlsx` has the updated numeric data before running it.

*Note: The `.bat` file automatically uses the `.venv` environment, so you only need to right-click or double-click it. You do not need to repeat Step 2 every time.*
