# WordDocToPdf PowerShell Script

This repository contains a simple PowerShell script designed to convert Microsoft Word documents (`.doc` or `.docx`) into PDF format. The script leverages Microsoft Word's COM object to automate the conversion process.

## Features

- Converts Word documents to PDF format.
- Batch conversion of multiple files in a specified directory.
- Minimal dependencies – only requires PowerShell and Microsoft Word installed on your system.

## Prerequisites

Before using this script, ensure you have:

1. **Microsoft Word** installed on your computer. The script uses the Microsoft Word COM object to perform the conversion.
2. **PowerShell** (typically pre-installed on Windows systems).

## How to Use

### Step 1: Clone or Download the Repository
Clone the repository or download the script directly to your local machine:

~~~bash
git clone https://github.com/Saqqe/WordDocToPdf_PowerShellScript.git
~~~

### Step 2: Open PowerShell
Navigate to the folder where the script is located. For example:

~~~powershell
cd C:\path\to\WordDocToPdf_PowerShellScript
~~~

### Step 3: Run the Script

The script will ask for **source** and **destination** folders *(script will create the destination folder if it does not exist)*

Example:
~~~powershell
PS E:\WordDocToPdf_PowerShellScript> .\docToPdf.ps1
Enter the path to the source folder containing Word documents: C:\Documents\TestingWordFiles
Enter the path to the destination folder for PDF files: C:\Documents\TestingWordFiles\pdfs
~~~

### Troubleshooting
[!CAUTION]

Ensure Execution Policy Allows Scripts: You may need to adjust PowerShell's execution policy if scripts are not allowed to run 
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4 

[!CAUTION]
~~~powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
~~~