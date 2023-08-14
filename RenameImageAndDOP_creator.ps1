# Load the required .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# Create an instance of FolderBrowserDialog
$startPath = "P:\01_Photo_Production\_TempRenameInProgress\"
$folderBrowserDialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
$folderBrowserDialog.SelectedPath = $startPath
$folderBrowserDialog.Description = "Select a folder"

# Show the folder selection dialog
$dialogResult = $folderBrowserDialog.ShowDialog()

# Dialog Result Processing
# Check if the user clicked "OK" and get the selected path
if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $rootDirectory = $folderBrowserDialog.SelectedPath
}
else {
    exit
}

# User Input Validation
# If user cancels the folder selection, exit the script
if ([string]::IsNullOrEmpty($rootDirectory)) {
    Write-Output "No folder selected, exiting..."
    exit
}

# Count total files
$totalFiles = (Get-ChildItem -File -Recurse $rootDirectory | Measure-Object).Count
Write-Output "Directory to process: $rootDirectory"
Write-Output "Total files to process: $totalFiles"

# Initialize variables
$fileCount = 0
$imageFileCount = 0
$dopFileCount = 0
$imageFilesRenamed = 0
$dopFilesRenamed = 0
$skippedFileCount = 0
$currentTime = Get-Date
$lastOutputTime = $currentTime
$lastGcTime = $currentTime
$logFile = "RenameLog_$($currentTime.ToString('yyyyMMdd_HHmmss')).txt"

# Initialize log file
Add-Content -Path "$rootDirectory\$logFile" -Value "Log file created at $currentTime"

# Main Processing - Files Iteration
Get-ChildItem -File -Recurse $rootDirectory | ForEach-Object {
    try {
        $currentFile = $_.FullName
        $currentFileName = $_.BaseName
        $currentFileExtension = $_.Extension.ToLower()

        # Pre-processing: File naming adjustments
        if ($_.Extension -ne $currentFileExtension) {
            Rename-Item -Path $currentFile -NewName ($_.BaseName + $currentFileExtension)
        }

        if ($currentFileName -like "IMG_*") {
            $newFileName = $currentFileName -replace "^IMG_", ""
            Rename-Item -Path $currentFile -NewName ($newFileName + $currentFileExtension)
            $currentFileName = $newFileName
        }

        if ($currentFileExtension -ne ".dop") {
            $exiftoolOutput = & exiftool.exe -d "%Y%m%d_%H%M%S" -DateTimeOriginal $currentFile 2>$null
            if (!$exiftoolOutput) {
                $skippedFileCount++
                Add-Content -Path "$rootDirectory\$logFile" -Value "Skipped (ExifTool output not found): $currentFile"
                throw "ExifTool output not found"
            } else {
                $currentFileDateTime = $exiftoolOutput.Split(":")[1].Trim()
                if ($currentFileDateTime.Length -ne 15) {
                    $skippedFileCount++
                    Add-Content -Path "$rootDirectory\$logFile" -Value "Skipped (Invalid DateTime length): $currentFile"
                    throw "Invalid DateTime length"
                }

                $currentFileDateHour = $currentFileDateTime.Substring(0,11)
                if ($currentFileName -notlike "*$currentFileDateHour*") {
                    $newFileNameWithExtension = "$currentFileDateTime" + "_" + $currentFileName + $currentFileExtension
                    Rename-Item -Path $currentFile -NewName $newFileNameWithExtension
                    $imageFilesRenamed++

                    $dopFile = $currentFile + ".dop"
                    if (Test-Path $dopFile) {
                        $newDopFileName = "$currentFileDateTime" + "_" + $currentFileName + $currentFileExtension + ".dop"
                        Rename-Item -Path $dopFile -NewName $newDopFileName
                        $dopFilesRenamed++
                    }
                }
                $imageFileCount++
                Add-Content -Path "$rootDirectory\$logFile" -Value "Handled: $currentFile"
            }
            $currentTime = Get-Date
            if (($currentTime - $lastOutputTime).TotalSeconds -gt 3) {
                $statusMessage = "$fileCount out of $totalFiles files processed..."
                Write-Output $statusMessage
                Add-Content -Path "$rootDirectory\$logFile" -Value $statusMessage
                $lastOutputTime = Get-Date
            }
        }
        else {
            $dopFileCount++
            Add-Content -Path "$rootDirectory\$logFile" -Value "Not acted upon: $currentFile"
        }
        $fileCount++
        if (($currentTime - $lastGcTime).TotalMinutes -gt 5) {
            [GC]::Collect()
            $lastGcTime = Get-Date
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Add-Content -Path "$rootDirectory\$logFile" -Value "Error processing file `"$currentFile`": $errorMessage"
    }
}

# Post-Processing
$endMessages = @(
    "Done.",
    "Total files processed: $fileCount",
    "Image files processed: $imageFileCount",
    "Image files renamed: $imageFilesRenamed",
    ".dop files processed: $dopFileCount",
    ".dop files renamed: $dopFilesRenamed",
    "Total files skipped: $skippedFileCount"
)

# Write the post-processing messages to console and log file
foreach ($message in $endMessages) {
    Write-Output $message
    Add-Content -Path "$rootDirectory\$logFile" -Value $message
}
