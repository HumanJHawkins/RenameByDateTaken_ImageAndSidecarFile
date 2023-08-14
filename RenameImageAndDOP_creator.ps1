# Load the required .NET assembly
Add-Type -AssemblyName System.Windows.Forms

# A logging function:
function LogOutput ($textToOutput) {
    Add-Content -Path "$rootDirectory\$logFile" -Value $textToOutput
    Write-Output $textToOutput
}

function Rename-FileAndSidecar {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePathName,
        
        [Parameter(Mandatory=$true)]
        [string]$RenameTo
    )

    # Ensure file exists
    if (!(Test-Path -Path $FilePathName)) {
        throw "File $FilePathName does not exist."
    }

    # Rename file
    Rename-Item -Path $FilePathName -NewName $RenameTo

    # Rename sidecar file if exists
    if (Test-Path -Path ("$FilePathName.dop")) {
        Rename-Item -Path ("$FilePathName.dop") -NewName ("$RenameTo.dop")
        return 2   # 2 Files renamed
    }

    return 1   # 1 file renamed
}


# Get folder to operate on:
$defaultPath = "P:\PhotoRenameTest\"
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

# Initialize variables and log file
$totalFiles = (Get-ChildItem -Path $rootDirectory -File -Recurse).Count
$statusUpdateInterval = 20

$currentTime = Get-Date
$lastStatusUpdateTime = $currentTime
$logFile = "RenameLog_$($currentTime.ToString('yyyyMMdd_HHmmss')).txt"
$deleteScriptFile = "DeleteOrphanedDOPs_$($currentTime.ToString('yyyyMMdd_HHmmss')).txt"

LogOutput "Log file created at $currentTime"
LogOutput "Root directory selected: $rootDirectory"


# Preprocess all files. Since we may be dealing with .dop sidecar
#  files that need to be kept in sync, analyze everything first. 
$fileInfoList = New-Object System.Collections.Generic.List[object]
$statusCounter = 0
Get-ChildItem -Path $rootDirectory -File -Recurse | Where-Object { $_.FullName -ne "$rootDirectory\$logFile" } | ForEach-Object {
    # Store file information
    $fileBaseName = $_.BaseName
    $fileExtension = $_.Extension
    $filenameAlreadyValid = $false  # Default to False, this will be updated later in the second loop

    # Presume the usual corresponding filename. If not actually present, revert to null.
    $correspondingFileName = if ($fileExtension -eq ".dop") { 
        # Remove only the last occurrence of .dop
        $fileBaseName -replace '\.dop$', ''
    } else {
        $fileBaseName + $fileExtension + ".dop"
    }

    if(!(Test-Path -Path "$($_.DirectoryName)\$correspondingFileName")) {
        $correspondingFileName = $null
    }

    $fileInfo = New-Object -TypeName PSObject -Property @{
        FullPathName = $_.FullName
        FileName = $fileBaseName + $fileExtension
        IsSidecarFile = $fileExtension -eq ".dop"
        DateTimeTaken = $null
        DateTimeTakenIsValid = $false
        Renamed = $false
        FilenameAlreadyValid = $filenameAlreadyValid
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
                # LogOutput "No DateTimeTaken found for `"$($_.FullName)`""
            }
        }
        catch {
            $errorMessage = "Error retrieving DateTimeTaken for `"$($_.FullName)`": $_.Exception.Message"
            LogOutput $errorMessage
        }
    }

    # Add file info to list
    $fileInfoList.Add($fileInfo)

    # Handle status update
    $statusCounter++
    if ($statusCounter -gt $statusUpdateInterval) {
        LogOutput "$($fileInfoList.Count) out of $totalFiles files processed..."
        $statusCounter = 0
    }
}

# Second loop through $fileInfoList to update $filenameAlreadyValid
# TO DO: Clean up some filenames here... Remove "IMG_", move existing datestamp to the front of name?
$statusCounter = 0
$fileInfoList | ForEach-Object {
    if ($_.DateTimeTakenIsValid) {
        # Extract the full DateTimeTaken and the date with 2-digit year from DateTimeTaken
        $dateTime = $_.DateTimeTaken
        $date = $_.DateTimeTaken.substring(2, 6)

        # Check if the FileName contains the date with 2-digit year and update $filenameAlreadyValid
        if ($_.FileName -match $date) {
            $_.FilenameAlreadyValid = $true
        }
    }
    
    # Handle status update
    $statusCounter++
    if ($statusCounter -gt $statusUpdateInterval) {
        LogOutput "Loop 2: $($fileInfoList.Count) out of $totalFiles files processed..."
        $statusCounter = 0
    }
}


# Preprocessing report
LogOutput "Preprocessing complete."
LogOutput "Total files: $totalFiles"
LogOutput "Files to be renamed: $($fileInfoList.Where({!$_.IsSidecarFile -and !$_.FilenameAlreadyValid -and $_.DateTimeTakenIsValid}).Count)"
$filesToRename = $fileInfoList.Where({!$_.IsSidecarFile -and !$_.FilenameAlreadyValid -and $_.DateTimeTakenIsValid}).Count
LogOutput "Files not to be renamed: $($totalFiles - $filesToRename)"
LogOutput "Valid .dop files: $($fileInfoList.Where({$_.IsSidecarFile -and $_.CorrespondingFileName}).Count)"
LogOutput "Invalid .dop files: $($fileInfoList.Where({$_.IsSidecarFile -and !$_.CorrespondingFileName}).Count)"


# Rename files
$fileCount = 0
$renamedNonSidecarFiles = 0
$renamedSidecarFiles = 0

$fileInfoList | ForEach-Object {
    # Rename non-sidecar files with a valid DateTimeTaken
    if (!$_.IsSidecarFile -and !$_.FilenameAlreadyValid -and $_.DateTimeTakenIsValid) {
        try {
            $newName = $_.DateTimeTaken + "_" + $_.FileName
            $renamedFileCount = Rename-FileAndSidecar $_.FullPathName $newName
            $_.Renamed = $true
            $renamedNonSidecarFiles++

            if($renamedFileCount -eq 2) {
                # $fileInfoList.Where({$_.FullPathName -eq $correspondingFullPathName})[0].Renamed = $true
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
LogOutput "Renaming complete."
LogOutput "Non-DOP files renamed: $renamedNonSidecarFiles"
LogOutput "DOP files renamed: $renamedSidecarFiles"


# Create list for already valid but nonconforming files
$validNonconformingFiles = $fileInfoList.Where({$_.FilenameAlreadyValid -and $_.DateTimeTakenIsValid -and $_.FileName -notmatch "^$($_.DateTimeTaken)"})
if ($validNonconformingFiles.Count -gt 0) {
    $listFile = "ValidNonconformingFiles_$($currentTime.ToString('yyyyMMdd_HHmmss')).txt"
    $validNonconformingFiles | ForEach-Object {
        Add-Content -Path "$rootDirectory\$listFile" -Value $_.FileName
    }
    LogOutput "List of valid but nonconforming files created at: $rootDirectory\$listFile"
}


# Create deletion script for orphaned/non-corresponding DOP files
if ($fileInfoList.Where({$_.IsSidecarFile -and !$_.CorrespondingFileName}).Count -gt 0) {
    $fileInfoList.Where({$_.IsSidecarFile -and !$_.CorrespondingFileName}) | ForEach-Object {
        Add-Content -Path "$rootDirectory\$deleteScriptFile" -Value "del `"$($_.FullPathName)`""
    }
    LogOutput "Deletion script for orphaned sidecar files created at: $rootDirectory\$deleteScriptFile"
}


$fileInfoList | ConvertTo-Json | Out-File "$rootDirectory\fileInfoList.json"
LogOutput "File data saved to: $rootDirectory\fileInfoList.json"


# Pause to allow user to read output
Read-Host -Prompt "Press Enter to continue"
