<#

    XLStoCSV

    DESCRIPTION
      Converts Beckman Multisizer4 data (.#m4.XLS files) into CSV files

    INSTRUCTIONS
      Save a copy of this PowerShell script in the same folder as the 
      .XLS files, and run the following command in the PowerShell console:

      > powershell -ExecutionPolicy Bypass -File .\XLStoCSV.ps1

      Note that the directory must be set to the correct folder in order for
      this script to work. For example, if the files are stored on the path
      C:\Users\JoeSmith\Data, then you can set the directory by:

      > cd "C:\Users\JoeSmith\Data"

    RESULTS
      The script will print a message stating when the task is complete. The
      final result will be a new folder called "CSV_Files" within the original
      folder. The "CSV_Files" folder will contain a file called "all_data.csv"
      that stores the size distribution data for each .#m4.XLS file within a
      single spreadsheet.


    COPYRIGHT 2025 Austin M. Weber
    
#>


Write-Host "==================================================================="
Write-Host " "

# Get the current folder
$folderPath = Get-Location

# Notify the user the script is running
Write-Host "... converting XLS files to CSV files. Please wait."

Write-Host " "
Write-Host "-------------------------------------------------------------------"


# Main function
#------------------------------------------------------------------------------|

# Create a variable for the CSV_Files path
$csvFolderPath = Join-Path -Path $folderPath -ChildPath "CSV_Files"

# Create the CSV_Files folder if it doesn't exist
if (-not (Test-Path -Path $csvFolderPath)) {
    New-Item -Path $csvFolderPath -ItemType Directory > $null
}

# Get all .XLS files in the original folder
$xlsFiles = Get-ChildItem -Path $folderPath -Filter *.xls

foreach ($file in $xlsFiles) {
    # Open the XLS file
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($file.FullName)
    $worksheet = $workbook.Sheets.Item(1)

    # Extract data starting from line 38
    $data = @()
    $row = 38
    while ($worksheet.Cells.Item($row, 1).Value() -ne $null) {
        $size = $worksheet.Cells.Item($row, 1).Value()
        $occurrences = $worksheet.Cells.Item($row, 2).Value()
        $data += [PSCustomObject]@{ Bins = $size; Counts = $occurrences }
        $row++
    }

    # Remove the string pattern ".#m4" from the file name
    $csvFileName = $file.BaseName -replace "\.\#m4", ""
    $csvPath = Join-Path -Path $csvFolderPath -ChildPath "$csvFileName.csv"

    # Save the extracted data to a CSV file in the CSV_Files folder
    $data | Export-Csv -Path $csvPath -NoTypeInformation

    # Close the workbook and Excel application
    $workbook.Close($false)
    $excel.Quit()
}

# Transpose the data in each CSV file; this step is necessary for
# concatenation purposes later on.
$csvFiles = Get-ChildItem -Path $csvFolderPath -Filter *.csv

foreach ($csvFile in $csvFiles) {
    $csvData = Import-Csv -Path $csvFile.FullName

    # Transpose the data
    $bins = $csvData.Bins
    $counts = $csvData.Counts
    $transposedData = @(
        $bins -join ","
        $counts -join ","
    )

    # Save the transposed data back to the CSV file
    Set-Content -Path $csvFile.FullName -Value $transposedData
}

# Save the names of the CSV files (excluding ".csv" extension) to a variable
# This step is for later when we add the names of the original files to the
# all_data.csv file
$allFileNames = $csvFiles | Where-Object { $_.Name -ne "all_data.csv" } | ForEach-Object { $_.BaseName }

# Create a copy of the first CSV file as "all_data.csv"
$firstCsvFile = $csvFiles[0]
$allDataPath = Join-Path -Path $csvFolderPath -ChildPath "all_data.csv"
Copy-Item -Path $firstCsvFile.FullName -Destination $allDataPath

# Loop through the remaining CSV files and concatenate the second row into "all_data.csv"
# we only need the second row from each file because the first row is (assumed to be)
# the same in each file and the first row has already been preallocateed from the step
# above.
foreach ($csvFile in $csvFiles[1..($csvFiles.Count - 1)]) {
    if ($csvFile.Name -ne "all_data.csv") {
        $csvData = Get-Content -Path $csvFile.FullName
        $secondRow = $csvData[1]
        Add-Content -Path $allDataPath -Value $secondRow
    }
}

# Add the "Sample" column to "all_data.csv"
# This will be the column containing the names of the original .XLS files
$allData = Get-Content -Path $allDataPath
$sampleColumn = "Sample" + "," + ($allFileNames -join ",")
$allDataWithSample = @()

for ($i = 0; $i -lt $allData.Count; $i++) {
    $row = $allData[$i]
    $sample = if ($i -eq 0) { "Sample" } else { $allFileNames[$i - 1] }
    $allDataWithSample += "$row,$sample"
}

# Shift the last column to the first column position
$shiftedData = @()
foreach ($line in $allDataWithSample) {
    $columns = $line -split ","
    $shiftedData += ($columns[-1] + "," + ($columns[0..($columns.Length - 2)] -join ","))
}

Set-Content -Path $allDataPath -Value $shiftedData

#------------------------------------------------------------------------------|
# End main function

# Notify the user that the script has successfully completed.
Write-Host " "
Write-Host "DATA CONVERSION COMPLETE."
Write-Host " "
Write-Host "All CSV files have been saved to the CSV_Files folder."
Write-Host " "
Write-Host "==================================================================="