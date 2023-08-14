# Define a parameter block
param (
    [string]$rootDirectory = (New-Object -ComObject "Shell.Application").BrowseForFolder(0, "Select a folder", 0).Self.Path
)

# If user cancels the folder selection, exit the script
if([string]::IsNullOrEmpty($rootDirectory)) {
    Write-Output "No folder selected, exiting..."
    exit
}

# Count total files
$totalFiles = (Get-ChildItem -File -Recurse $rootDirectory | Measure-Object).Count
Write-Output "Directory to process: $rootDirectory"
Write-Output "Total files to process: $totalFiles"

# Create a new text file with the current date/time as its name
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$bulkRenameBatchFile = "$currentDateTime.txt"
Add-Content -Path $bulkRenameBatchFile -Value "ECHO Begining List of files to rename."
Write-Output "Generated script: $bulkRenameBatchFile"

# Use ExifTool to generate the rename commands
$fileCount = 0
$lastOutputTime = Get-Date
Get-ChildItem -File -Recurse $rootDirectory | ForEach-Object {
    # Convert file extension to lowercase
    $newFileName = $_.BaseName + $_.Extension.ToLower()
    Rename-Item -Path $_.FullName -NewName $newFileName

    $currentFile = $_.FullName.ToLower()
    $currentFile = $currentFile -replace "'", "''" # Replace single quote with double single quotes
    $currentFile = $currentFile -replace "_", "`_" # Replace underscore with escaped underscore
    $currentFile = $currentFile -replace "\(", "`(" # Replace opening parenthesis with escaped opening parenthesis
    $currentFile = $currentFile -replace "\)", "`)" # Replace closing parenthesis with escaped closing parenthesis
    $currentFileName = $_.BaseName
    $currentFileExtension = $_.Extension

    if($currentFileExtension -ne ".dop"){
        $exiftoolOutput = & exiftool.exe -d "%Y%m%d_%H%M%S" -DateTimeOriginal $currentFile 2>$null
        if(!$exiftoolOutput){
            # Write-Output "Error reading DateTimeOriginal from $currentFile"
        } else {
            $currentFileDateTime = $exiftoolOutput.Split(":")[1].Trim()
            if(!$currentFileDateTime){
                # Write-Output "Error reading DateTimeOriginal from $currentFile"
                continue
            }
            $newFileNameWithExtension = "$currentFileDateTime" + "_" + $currentFileName + $currentFileExtension
            $dopFile = $currentFile + ".dop"
            Add-Content -Path $bulkRenameBatchFile -Value "ren `"$currentFile`" `"$newFileNameWithExtension`""
            if(Test-Path $dopFile){
                $fileCount++  # This is an extra file that is handled this loop.
                Add-Content -Path $bulkRenameBatchFile -Value "ren `"$dopFile`" `"$newFileNameWithExtension.dop`""
            }
        }
        
        $fileCount++
        $currentTime = Get-Date
        if(($currentTime - $lastOutputTime).TotalSeconds -gt 1){
            Write-Output "$fileCount out of $totalFiles files processed..."
            $lastOutputTime = Get-Date
        }
    }
}

Write-Output "Done."
