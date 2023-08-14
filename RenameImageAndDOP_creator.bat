@echo off
setlocal enabledelayedexpansion

REM Set the root directory
set rootDirectory=%1
if not defined rootDirectory set rootDirectory=.

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
