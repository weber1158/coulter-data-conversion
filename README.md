## What is this repository?

For those who use a Beckman Multisizer4 to collect Coulter Counter data, you may have noticed that the data files are saved with a `.#m4.XLS` extension, and that these files are saved one sample at a time. 

It can be a pain to manually combine the data from each file into a single spreadsheet, and so this repository provides a `PowerShell` script that automatically extracts the data from these files, converts them to CSVs, and also creates a new file with all of the data aggregated in the same CSV file.

## Requirements
- PowerShell (Version 5.1 or later)
- PowerShell ISE

## How to use
1. Download the `XLStoCSV.ps1` file and save it somewhere on your PC.
2. Copy the folder containing the `.#m4.XLS` files onto a flash drive and then save the folder somewhere on your PC.
3. Make a copy of the `XLStoCSV.ps1` file and paste it in the same folder as the Coulter Counter data.
4. Right-click the `XLStoCSV.ps1` file in the folder and click <kbd>Edit</kbd>. This should open the PowerShell ISE.
    - If the PowerShell ISE does not open by this method, open it manually by using the <kbd>&#8862; Win</kbd> + <kbd>R</kbd> keyboard shortcut and typing `powershell_ise` into the Run application. Then, type the following command in the PowerShell console to change the directory to the correct folder:
  
```powershell
Set-Location "<path>"
```
where `<path>` is the path to the folder where the `XLStoCSV.ps1` file and the Coulter Counter files are stored. For example:

```powershell
Set-Location "C:\Users\JoeSmith\Desktop\Data"
```

5. Type the following into the PowerShell console to execute the `XLStoCSV.ps1` script:

```powershell
PowerShell -ExecutionPolicy Bypass -File .\XLStoCSV.ps1
```

All of the `.#m4.XLS` files will be converted into `.csv` files and saved to a new folder called `CSV_Files`. This folder will also contain a file called `all_data.csv` which contains all of the data aggregated into a single spreadsheet.
