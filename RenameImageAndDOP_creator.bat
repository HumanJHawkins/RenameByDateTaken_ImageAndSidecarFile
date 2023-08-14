@echo off
setlocal enabledelayedexpansion

REM Check if the user has provided a directory
if "%~1"=="" (
    echo Error: Please provide a directory.
    exit /b
)

REM Set the root directory
set rootDirectory=%~1

REM Convert all file extensions to lowercase using PowerShell
echo Converting all file extensions to lowercase...
for /R "%rootDirectory%" %%F in (*.*) do (
    powershell -command "Get-Item \"%%F\" | Rename-Item -NewName { $_.BaseName + $_.Extension.ToLower() }"
)
for /R "%rootDirectory%" %%F in (*.*.dop) do (
    powershell -command "$item = Get-Item \"%%F\"; $basename = $item.BaseName; $extension = $item.Extension; $dotIndex = $basename.LastIndexOf('.'); if ($dotIndex -ge 0) { $basename = $basename.Substring(0, $dotIndex) + $basename.Substring($dotIndex).ToLower() }; Rename-Item -Path $item.FullName -NewName ( $basename + $extension )"
)
echo Done converting extensions.

REM Create a new text file with the current date/time as its name
for /f "delims=" %%a in ('wmic OS Get localdatetime  ^| find "."') do set currentDateTime=%%a
set bulkRenameBatchFile=%currentDateTime:~0,8%_%currentDateTime:~8,6%.txt
echo REM Created Script >> %bulkRenameBatchFile%
echo Generated script: %bulkRenameBatchFile%

REM Use ExifTool to generate the rename commands
for /R "%rootDirectory%" %%F in (*.*) do (
    set "currentFile=%%F"
    set "currentFileDir=%%~dpF"
    set "currentFileName=%%~nF"
    set "currentFileExtension=%%~xF"
    
    if /I "!currentFileExtension!" NEQ ".dop" (
        for /f "delims=" %%i in ('exiftool -d "%%Y%%m%%d_%%H%%M%%S" -p "$DateTimeOriginal" "!currentFile!" 2^>nul') do (
            set "currentFileDateTime=%%i"
            if "!currentFileDateTime!" NEQ "" (
                set "newFileNameWithExtension=!currentFileDateTime!_!currentFileName!!currentFileExtension!"
                set "dopFile=!currentFile!.dop"
                echo ren "!currentFile!" "!newFileNameWithExtension!" >> %bulkRenameBatchFile%
                if exist "!dopFile!" (
                    echo ren "!dopFile!" "!newFileNameWithExtension!.dop" >> %bulkRenameBatchFile%
                )
            ) else (
                echo Error reading DateTimeOriginal from "!currentFile!" >&2
            )
        )
    )
)
echo Done.
endlocal
