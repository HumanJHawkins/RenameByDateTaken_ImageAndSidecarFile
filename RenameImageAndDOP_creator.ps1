# Load the required .NET assembly
Add-Type -AssemblyName System.Windows.Forms

$verboseLogging = $false  # Set to $true for verbose logging

function logOutput ($textToOutput, $verbose=$false) {
    if ($verbose -and !$verboseLogging) {
        return
    }
    Add-Content -Path "$rootDirectory\$logFile" -Value $textToOutput
    Write-Output $textToOutput
}

function renameFileWithSidecar {
    param (
        [Parameter(Mandatory=$true)]
        [string]$filePathName,
        
        [Parameter(Mandatory=$true)]
        [string]$RenameTo
    )

    # Ensure file exists
    if (!(Test-Path -Path $filePathName)) {
        throw "File $filePathName does not exist."
    }

    # Rename file
    Rename-Item -Path $filePathName -NewName $RenameTo

    # Rename sidecar file if exists
    if (Test-Path -Path ("$filePathName.dop")) {
        Rename-Item -Path ("$filePathName.dop") -NewName ("$RenameTo.dop")
        return 2   # 2 Files renamed
    }

    return 1   # 1 file renamed
}

function getInternalDateTime {
    param(
        [Parameter(Mandatory=$true)]
        [string]$filePathName
    )

    $exiftoolOutput = & exiftool.exe -d "%Y%m%d_%H%M%S" -DateTimeOriginal $filePathName 2>$null
    $outputResult = logOutput "   getInternalDateTime exiftoolOutput: $exiftoolOutput" $true
    if ($exiftoolOutput) {
        $dateTime = $exiftoolOutput -replace '.*: '
        $outputResult = logOutput "   getInternalDateTime dateTime: $dateTime" $true
        return $dateTime
    } else {
        $outputResult = logOutput "   getInternalDateTime No exiftool output for: $filePathName" 
        return $null
    }
}

function validateExifTimestamp($exifTimestamp=$null) {
    if ($null -eq $exifTimestamp) {
        return $false
    }

    $dummy = New-Object DateTime
    return [DateTime]::TryParseExact($exifTimestamp, "yyyyMMdd_HHmmss", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$dummy)
}

function extractFilenameTimestamp {
    param(
        [Parameter(Mandatory=$true)]
        [string]$fileBaseName
    )

    $outputResult = logOutput "   extractFilenameTimestamp received: $fileBaseName" $true

    # Check for timestamps in order of precedence
    $timestampFormats = @(
        '\d{4}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}',
        '\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{4}', 
        '\d{8}[-_ .]\d{6}', 
        '\d{8}_(AM|PM)\d{4}', 
        '\d{8}_\d{4}', 
        '\d{8}', 
        '\d{4}[-_ .]\d{2}[-_ .]\d{2}'
    )

    foreach($format in $timestampFormats) {
        # Create pattern with separators
        $pattern = '(^|\s|-|_|\.)(' + $format + ')($|\s|-|_|\.|\b)'

        if ($fileBaseName -match $pattern) {
            $outputResult = logOutput "   extractFilenameTimestamp format: $format" $true
            $outputResult = logOutput "   extractFilenameTimestamp Matches[0]: $($Matches[0])" $true
            return $Matches[0]
        }
    }

    return $null
}


function reformatFilenameTimestamp {
    param(
        [Parameter(Mandatory=$true)]
        [string]$filenameTimestamp
    )
    
    $outputResult = logOutput "   reformatFilenameTimestamp received: $filenameTimestamp" $true

    # Define timestamp formats
    $timestampFormats = @(
        '\d{4}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}',
        '\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{4}', 
        '\d{8}[-_ .]\d{6}', 
        '\d{8}_(AM|PM)\d{4}', 
        '\d{8}_\d{4}', 
        '\d{8}', 
        '\d{4}[-_ .]\d{2}[-_ .]\d{2}'
    )

    # Initialize a variable to store the matched format
    $matchedFormat = $null

    foreach($format in $timestampFormats) {
        # Create pattern with separators
        $pattern = '(^|\s|-|_|\.)(' + $format + ')($|\s|-|_|\.|\b)'

        if ($filenameTimestamp -match $pattern) {
            $matchedFormat = $format
            $outputResult = logOutput "   reformatFilenameTimestamp format: $matchedFormat" $true
            break
        }
    }

    $reformedFilenameTimestamp = $filenameTimestamp
    # Perform different transformations based on the matched format
    switch ($matchedFormat) {
        '\d{4}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{2}' {
            $reformedFilenameTimestamp = $reformedFilenameTimestamp -replace '[\._ -]', ''
            $reformedFilenameTimestamp = $reformedFilenameTimestamp.Substring(0, 8) + "_" + $reformedFilenameTimestamp.Substring(8)
        }
        '\d{2}[-_ .]\d{2}[-_ .]\d{2}[-_ .]\d{4}' {
            $year = [int]$reformedFilenameTimestamp.Substring(4, 2) + 2000
            $reformedFilenameTimestamp = $reformedFilenameTimestamp -replace '[\._ -]', ''
            $reformedFilenameTimestamp = $year + $reformedFilenameTimestamp.Substring(0, 4) + "_" + $reformedFilenameTimestamp.Substring(8)
        }
        '\d{8}[-_ .]\d{6}' {
            $reformedFilenameTimestamp = $reformedFilenameTimestamp -replace '[\._ -]', ''
            $reformedFilenameTimestamp = $reformedFilenameTimestamp.Substring(0, 8) + "_" + $reformedFilenameTimestamp.Substring(8)
        }
        '\d{8}_(AM|PM)\d{4}' {
            $revisionType = "Nothing"   # Might be letting an odd seperator through
        }
        '\d{8}_\d{4}' {
            $reformedFilenameTimestamp = $reformedFilenameTimestamp -replace '[\._ -]', ''
            $reformedFilenameTimestamp = $reformedFilenameTimestamp.Substring(0, 8) + "_" + $reformedFilenameTimestamp.Substring(8)
        }
            # No change to 8-digit timestamps
            # '\d{8}' {
            # }
        '\d{4}[-_ .]\d{2}[-_ .]\d{2}' {
            $reformedFilenameTimestamp = $reformedFilenameTimestamp -replace '[\._ -]', ''
        }
        default {
            $outputResult = logOutput "Error: Default reached in reformatFilenameTimestamp."
        }
    }
    
    $outputResult = logOutput "   reformatFilenameTimestamp reformedFilenameTimestamp: $reformedFilenameTimestamp" $true
    return $reformedFilenameTimestamp
}


function cleanupFilename {
    param(
        [Parameter(Mandatory=$true)]
        [string]$baseFileName
    )
    
    $outputResult = logOutput "      cleanupFilename baseFileName: $baseFileName" $true

    # Remove 'IMG_' or 'ContactPhoto-IMG_' or 'VID_' if they are at the beginning
    $baseFileName = $baseFileName -replace '^(IMG_|ContactPhoto-IMG_|VID_)', ''
    
    $outputResult = logOutput "      cleanupFilename after remove IMG_+: $baseFileName" $true

    # Remove '_DxO' from the end of the baseFileName
    $baseFileName = $baseFileName -replace '_DxO$', ''
    $outputResult = logOutput "      cleanupFilename after remove _DxO: $baseFileName" $true

    # Reduce certain repeated characters to a single of that character.
    $baseFileName = $baseFileName -replace '_+', '_'
    $baseFileName = $baseFileName -replace '-+', '-'
    $baseFileName = $baseFileName -replace ' +', ' '
    $baseFileName = $baseFileName -replace '\.+', '.'
    $outputResult = logOutput "      cleanupFilename after remove dup chars: $baseFileName" $true

    # Strip certain characters from the ends.
    $baseFileName = $baseFileName -replace '^[-_ .]+|[-_ .]+$', ''
    $baseFileName = $baseFileName.Trim()
    $outputResult = logOutput "      cleanupFilename after trim+: $baseFileName" $true

    return $baseFileName
}

$global:generatedFilenames = @()
function generateUniqueFilename ($directory, $originalFileName, $fileBaseName, $fileExtension) {
    $i = 0

    
    $outputResult = logOutput "      generateUniqueFilename directory: $directory" $true
    $outputResult = logOutput "      generateUniqueFilename originalFileName: $originalFileName" $true
    $outputResult = logOutput "      generateUniqueFilename fileBaseName: $fileBaseName" $true
    $outputResult = logOutput "      generateUniqueFilename fileExtension: $fileExtension" $true

    # Start a loop that continues until a unique filename is found
    do {
        # Create a potential new filename. Numbers only if necessary.
        if($i -gt 0) {
            $potentialFileName = $fileBaseName + '_' + $i.ToString() + $fileExtension
        } else {
            $potentialFileName = $fileBaseName + $fileExtension
        }
        $outputResult = logOutput "      generateUniqueFilename potentialFileName: $potentialFileName" $true

        # Check if the potential new filename exists in the directory
        $testFilePathName = $directory + '\' + $potentialFileName
        $outputResult = logOutput "X     generateUniqueFilename testFilePathName: $testFilePathName" $true
        # Original file always there, so skip if potential == original.
        if ($potentialFileName -ne $originalFileName) {
            if (Test-Path -Path $testFilePathName) {
                $outputResult = logOutput "      generateUniqueFilename NAME FOUND IN PATH" $true
                $i++
                continue
            }
        }

        # Check if the potential new filename exists in the global filenames
        if ($global:generatedFilenames -contains $potentialFileName) {
            $outputResult = logOutput "      generateUniqueFilename NAME FOUND IN LIST" $true
            $i++
            continue
        }

        # If no duplicates found, exit the loop
        break
    } while ($true)

    # Construct the unique filename
    $outputResult = logOutput "      generateUniqueFilename i: $i" $true
    if($i -gt 0) {
        $global:generatedFilenames += $potentialFileName
        return $potentialFileName
    } else {
        return $fileBaseName + $fileExtension
    }
}


function getConformingFilename {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$fileInfo,

        [Parameter(Mandatory=$false)]
        [string]$exifTimestamp
    )
    
    $outputResult = logOutput "   getConformingFilename exifTimestamp: $exifTimestamp" $true
    $outputResult = logOutput "   getConformingFilename fileInfo: $($fileInfo.FullPathName)" $true

    $exifTimestampValid = validateExifTimestamp $exifTimestamp

    $directory = $fileInfo.DirectoryName
    $originalFileName = $fileInfo.FileName
    $baseFileName = cleanupFilename ([IO.Path]::GetFileNameWithoutExtension($originalFileName))
    $fileExtension = [IO.Path]::GetExtension($originalFileName).ToLower()
    
    $outputResult = logOutput "   getConformingFilename directory: $directory" $true
    $outputResult = logOutput "   getConformingFilename fileInfo: $($fileInfo.FullPathName)" $true
    $outputResult = logOutput "   getConformingFilename Cleaned filename: $baseFileName" $true
    $outputResult = logOutput "   getConformingFilename fileExtension: $fileExtension" $true

    # Extract existing timestamp
    $filenameTimestamp = extractFilenameTimestamp $baseFileName
    $outputResult = logOutput "   getConformingFilename filenameTimestamp: $filenameTimestamp" $true

    # Add conforming datestamp if none found
    if ($exifTimestampValid) {
        # Attempt to erase existing timestamp. If none exists, none will be replaced.
        $escapedfilenameTimestamp = [regex]::Escape($filenameTimestamp)
        $baseFileName = $baseFileName -replace $escapedfilenameTimestamp, ''   
        $baseFileName = $exifTimestamp + '_' + $baseFileName
        $outputResult = logOutput "   getConformingFilename WITH VALID TIMESTAMP: $baseFileName" $true
    } else {
        if($filenameTimestamp -ne $null) {
            $outputResult = logOutput "   getConformingFilename: filenameTimestam IS NOT NULL" $true
            # Get improved timestamp format if possible.
            $reformedTimestamp = reformatFilenameTimestamp $filenameTimestamp
            $outputResult = logOutput "   getConformingFilename reformedTimestamp: $reformedTimestamp" $true

            # Remove the existing timestamp from wherever it is.
            $escapedfilenameTimestamp = [regex]::Escape($filenameTimestamp)
            $baseFileName = $baseFileName -replace $escapedfilenameTimestamp, ''
            $outputResult = logOutput "   getConformingFilename NOT valid + erased timestamp: $baseFileName" $true

            # Prepend the improved timestamp.
            $baseFileName = $reformedTimestamp + '_' + $baseFileName
            $outputResult = logOutput "   getConformingFilename WITH FILENAME TIMESTAMP: $baseFileName" $true
        } else {
            $outputResult = logOutput "   getConformingFilename filenameTimestamp: $filenameTimestamp" $true
            $outputResult = logOutput "   getConformingFilename: EXIF TIMESTAMP INVALID AND NO FILENAME TIMESTAMP" $true
        }
    }


    $baseFileName = cleanupFilename $baseFileName
    
    $outputResult = logOutput "   getConformingFilename locally cleaned to: $baseFileName" $true
    $outputResult = logOutput "   getConformingFilename originalFileName: $originalFileName" $true

    # Build full filename with duplicate check
    $fileName = $baseFileName + $fileExtension
    $outputResult = logOutput "   getConformingFilename fileName: $fileName" $true
    if ($fileName -cne $originalFileName) {
        $fileName = generateUniqueFilename $directory $originalFileName $baseFileName $fileExtension 
    }
    $outputResult = logOutput "   getConformingFilename after dup check: $fileName" $true

    # Add logging if the filename is just the extension
    if($fileName -eq $fileExtension) {
        $logMessage = "WARNING: Filename is only an extension: $fileName ********************************************"
        $outputResult = logOutput $logMessage
    }

    # Return the new filename
    return $fileName
}

# Get folder to operate on:
$defaultPath = "P:\zTest\"
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
$statusUpdateInterval = 9

$currentTime = Get-Date
$logFile = "RenameLog_$($currentTime.ToString('yyyyMMdd_HHmmss')).txt"
$deleteScriptFile = "DeleteOrphanedDOPs_$($currentTime.ToString('yyyyMMdd_HHmmss')).txt"

logOutput "Log file created at $currentTime"
logOutput "Root directory selected: $rootDirectory"


# Preprocess all files. Since we may be dealing with .dop sidecar
#  files that need to be kept in sync, analyze everything first. 
$fileInfoList = New-Object System.Collections.Generic.List[object]
$statusCounter = 0

$totalFiles = (Get-ChildItem -Path $rootDirectory -File -Recurse | Where-Object { $_.Extension -notin @('.txt', '.log', '.dop', '.json') } | Measure-Object).Count
logOutput "Total file count: $totalFiles"

Get-ChildItem -Path $rootDirectory -File -Recurse | 
Where-Object { $_.Extension -notin @('.txt', '.log', '.dop', '.json') } | ForEach-Object {
    # Store file information
    $fileBaseName = $_.BaseName
    $fileExtension = $_.Extension

    $fileInfo = New-Object -TypeName PSObject -Property @{
        FullPathName = $_.FullName
        DirectoryName = $_.Directory.FullName
        FileName = $fileBaseName + $fileExtension
        ExifTimestamp = $null
        NewFileName = $null
    }
    
    $outputResult = logOutput "Preprocessing: $($fileInfo.FullPathName)" $true

    try {
        $fileInfo.ExifTimestamp = getInternalDateTime $_.FullName

        # Call getConformingFilename outside the if-else block
        $newFileName = getConformingFilename -fileInfo $fileInfo -exifTimestamp $fileInfo.ExifTimestamp
        $outputResult = logOutput "   Preprocessing newFileName: $newFileName" $true
        if ($newFileName -cne $fileInfo.FileName) {
            $fileInfo.NewFileName = $newFileName
            $outputResult = logOutput "   Preprocessing fileName changed: $newFileName" $true
        } else {
            $outputResult = logOutput "   New filename $newFileName same as old $($fileInfo.FileName). Not renaming." $true
        }
    }
    catch {
        $outputResult = logOutput $_.Exception.Message
    }


    # Add file info to list
    $fileInfoList.Add($fileInfo)

    # Handle status update
    $statusCounter++
    if ($statusCounter -gt $statusUpdateInterval) {
        logOutput "$($fileInfoList.Count) out of $totalFiles files processed..."
        $statusCounter = 0
    }
}



# Preprocessing report
logOutput "Preprocessing complete."
logOutput "Total files: $totalFiles"

$filesToRename = $fileInfoList.Where({$_.NewFileName -ne $null}).Count
logOutput "Files to be renamed: $filesToRename"

$fileInfoList | ConvertTo-Json | Out-File "$rootDirectory\fileInfoList.json"
logOutput "File data saved to: $rootDirectory\fileInfoList.json"


# Pause to allow user to read output
Read-Host -Prompt "Press Enter to continue"

# Initialize counters
$fileCount = 0
$renamedNonSidecarFiles = 0
$renamedSidecarFiles = 0
$statusCounter = 0

$fileInfoList | ForEach-Object {
    # Rename files if NewFileName is not null
    if ($_.NewFileName -ne $null) {
        try {
            $renamedFileCount = renameFileWithSidecar $_.FullPathName $_.NewFileName
            $renamedNonSidecarFiles++

            if($renamedFileCount -eq 2) {
                $renamedSidecarFiles++
            }
        }
        catch {
            $errorMessage = "Error renaming file `"$_.FullPathName)`": $_.Exception.Message"
            logOutput $errorMessage
        }
    }

    $fileCount++
    
    # Report progress
    $statusCounter++
    if ($statusCounter -gt 100) {
        logOutput "$fileCount out of $totalFiles files processed..."
        $statusCounter = 0
    }
}

# Report final stats
logOutput "Renaming complete."
logOutput "Non-DOP files renamed: $renamedNonSidecarFiles"
logOutput "DOP files renamed: $renamedSidecarFiles"


# TO DO: Create deletion script for orphaned/non-corresponding DOP files


# Pause to allow user to read output
Read-Host -Prompt "Press Enter to continue"