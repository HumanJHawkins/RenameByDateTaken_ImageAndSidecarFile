# Load the required .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# A logging function:
function LogOutput ($textToOutput) {
    Add-Content -Path "$rootDirectory\$logFile" -Value $textToOutput
    Write-Output $textToOutput
}

# Get folder to operate on:
$defaultPath = "P:\01_Photo_Production\_TempRenameInProgress\"
$folderBrowserDialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
$folderBrowserDialog.SelectedPath = $defaultPath
$folderBrowserDialog.Description = "Select a folder"

if (!($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)) {
    Write-Output "Folder selection cancelled. Exiting."
    exit
} elseif ([string]::IsNullOrEmpty($folderBrowserDialog.SelectedPath)) {
    Write-Output "Invalid (Null or Empty) folder selection. Exiting."
    exit
}

$rootDirectory = $folderBrowserDialog.SelectedPath

# Initialize log file and variables
LogOutput "Log file created at $currentTime"
LogOutput "Root directory selected: $rootDirectory"


$totalFiles = (Get-ChildItem -Path $rootDirectory -File -Recurse).Count
$statusUpdateIntervalSeconds = 5

$currentTime = Get-Date
$lastStatusUpdateTime = $currentTime
$logFile = "RenameLog_$($currentTime.ToString('yyyyMMdd_HHmmss')).txt"
$deleteScriptFile = "DeleteOrphanedDOPs_$($currentTime.ToString('yyyyMMdd_HHmmss')).txt"


# Preprocess all files. Since we may be dealing with .dop sidecar
#  files that need to be kept in sync, analyze everything first. 
$fileInfoList = @()
Get-ChildItem -Path $rootDirectory -File -Recurse | ForEach-Object {
    # Store file information
    $fileName = $_.BaseName
    $fileExtension = $_.Extension
    $isDatePrefixed = $fileName -match '^\d{8}_\d{2}_'

    # Presume the usual corresponding filename. If not actually present, revert to null.
    $correspondingFileName = if ($fileExtension -eq ".dop") { 
        # Remove only the last occurrence of .dop
        $fileName -replace '\.dop$', ''
    } else {
        $fileName + $fileExtension + ".dop"
    }

    if(!(Test-Path -Path "$($_.DirectoryName)\$correspondingFileName")) {
        $correspondingFileName = $null
    }


    $fileInfo = New-Object -TypeName PSObject -Property @{
        FullPathName = $_.FullName
        FileName = $fileName
        FileExtension = $fileExtension
        IsSidecarFile = $fileExtension -eq ".dop"
        DateTimeTaken = $null
        DateTimeTakenIsValid = $false
        IsDatePrefixed = $isDatePrefixed
        FilenameAlreadyValid = $isDatePrefixed
        CorrespondingFileName = $correspondingFileName
    }

    # If it's an image file, get the DateTimeTaken
    if (!$fileInfo.IsSidecarFile) {
        try {
            $exiftoolOutput = & exiftool.exe -d "%Y%m%d_%H%M%S" -DateTimeOriginal $_.FullName 2>$null
            if ($exiftoolOutput) {
                $dateTime = $exiftoolOutput -replace '.*: '
                $fileInfo.DateTimeTaken = $dateTime
                $dummy = New-Object DateTime
                $fileInfo.DateTimeTakenIsValid = [DateTime]::TryParseExact($dateTime, "yyyyMMdd_HHmmss", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$dummy)

                if (!$fileInfo.DateTimeTakenIsValid -and ![string]::IsNullOrEmpty($dateTime)) {
                    # Log the DateTimeTaken extraction and validation result
                    $logMessage = "Invalid DateTimeTaken `"$dateTime`" extracted from `"$($_.FullName)`"."
                    LogOutput $logMessage
                }
            } else {
                # Log if no exif data found
                $logMessage = "No DateTimeTaken found for `"$($_.FullName)`""
                LogOutput $logMessage
            }
        }
        catch {
            $errorMessage = "Error retrieving DateTimeTaken for `"$($_.FullName)`": $_.Exception.Message"
            LogOutput $errorMessage
        }
    }

    # Add file info to list
    $fileInfoList += $fileInfo

    # Handle status update
    $currentTime = Get-Date
    if (($currentTime - $lastStatusUpdateTime).TotalSeconds -gt $statusUpdateIntervalSeconds) {
        $statusMessage = "$($fileInfoList.Count) out of $totalFiles files processed..."
        LogOutput $statusMessage
        $lastStatusUpdateTime = Get-Date
    }

}

# Preprocessing report
$filesToRename = $fileInfoList.Where({!$_.IsSidecarFile -and !$_.IsDatePrefixed -and $_.DateTimeTakenIsValid}).Count
$validSidecarFiles = $fileInfoList.Where({$_.IsSidecarFile -and $_.CorrespondingFileName}).Count
$invalidSidecarFiles = $fileInfoList.Where({$_.IsSidecarFile -and !$_.CorrespondingFileName}).Count
$preprocessingCompleteMessage = "Preprocessing complete. Total files: $totalFiles. Files to be renamed: $filesToRename. Files not to be renamed: $($totalFiles - $filesToRename). Valid .dop files: $validSidecarFiles. Invalid .dop files: $invalidSidecarFiles."
LogOutput $preprocessingCompleteMessage


# Rename files
$fileCount = 0
$renamedNonSidecarFiles = 0
$renamedSidecarFiles = 0

$fileInfoList | ForEach-Object {
    # Rename non-sidecar files with a valid DateTimeTaken
    if (!$_.IsSidecarFile -and !$_.IsDatePrefixed -and $_.DateTimeTakenIsValid) {
        try {
            $newName = $_.DateTimeTaken + "_" + $_.FileName + $_.FileExtension
            Rename-Item -Path $_.FullPathName -NewName $newName
            $_.IsDatePrefixed = $true
            $renamedNonSidecarFiles++

            # Rename corresponding sidecar file if there is one
            if ($_.CorrespondingFileName) {
                $correspondingFullPathName = $_.FullPathName.Replace($_.FileName + $_.FileExtension, $_.CorrespondingFileName)
                $newCorrespondingName = $newName + ".dop"
                Rename-Item -Path $correspondingFullPathName -NewName $newCorrespondingName
                $fileInfoList.Where({$_.FullPathName -eq $correspondingFullPathName})[0].IsDatePrefixed = $true
                $renamedSidecarFiles++
            }
        }
        catch {
            $errorMessage = "Error renaming file `"$_.FullPathName)`": $_.Exception.Message"
            LogOutput $errorMessage
        }
    }

    $fileCount++

    # Status reporting
    $currentTime = Get-Date
    if (($currentTime - $lastStatusUpdateTime).TotalSeconds -gt $statusUpdateIntervalSeconds) {
        $statusMessage = "$fileCount out of $totalFiles files processed..."
        LogOutput $statusMessage
        $lastStatusUpdateTime = Get-Date
    }
}

# Report the number of files renamed
$renamingCompleteMessage = "Renaming complete. Non-DOP files renamed: $renamedNonSidecarFiles. DOP files renamed: $renamedSidecarFiles."
LogOutput $renamingCompleteMessage

# Create deletion script for orphaned/non-corresponding DOP files
if ($fileInfoList.Where({$_.IsSidecarFile -and !$_.CorrespondingFileName}).Count > 0) {
    $fileInfoList.Where({$_.IsSidecarFile -and !$_.CorrespondingFileName}) | ForEach-Object {
        Add-Content -Path "$rootDirectory\$deleteScriptFile" -Value "del `"$($_.FullPathName)`""
    }
}

$fileInfoList | ConvertTo-Json | Out-File "$rootDirectory\fileInfoList.json"


# Pause to allow user to read output
Read-Host -Prompt "Press Enter to continue"
