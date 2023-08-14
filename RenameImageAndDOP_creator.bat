@echo off

REM Set the root directory
set root=%1
if not defined root set root=.

REM Create a new text file with the current date/time as its name
for /f "delims=" %%a in ('wmic OS Get localdatetime  ^| find "."') do set dt=%%a
set dt=%dt:~0,8%_%dt:~8,6%
echo REM Created Script >> %dt%.txt
echo Generated script: %dt%.txt

REM Use ExifTool to generate the rename commands
setlocal enabledelayedexpansion
set /a counter=1
set "LastComputedFilename="
for /R "%root%" %%F in (*.*) do (
    if /I "%%~xF" NEQ ".dop" (
        for /f "delims=" %%i in ('exiftool -d "%%Y%%m%%d_%%H%%M%%S" -p "$DateTimeOriginal" "%%F" 2^>nul') do (
            if "%%i" NEQ "" (
                if "!LastComputedFilename!" EQU "%%i" (
                    set /a counter+=1
                ) else (
                    set "LastComputedFilename=%%i"
                    set counter=01
                )
                echo ren "%%F" "%%i_!counter!%%~xF" >> %dt%.txt
                echo ren "%%~dpnF.dop" "%%i_!counter!.dop" >> %dt%.txt
            ) else (
                echo Error reading DateTimeOriginal from "%%F" >&2
            )
        )
    )
)

echo Done.
endlocal
