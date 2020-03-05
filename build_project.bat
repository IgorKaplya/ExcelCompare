echo off
set BuildProject=%1
set BuildConfiguration=%2
set BuildMode=%3
rem Verbosity: quiet, minimal, normal (default), detailed
set BuildVerbosity=minimal
set OutPath=%4
set ProjectDir=%~nx1
echo ------------------------Starting build script------------------------
echo [Conf]
echo     Build project: %BuildProject%
echo     Build configuration: %BuildConfiguration%
echo     Build mode: %BuildMode%
echo     BuildVerbosity: %BuildVerbosity%
echo     OutPath: %OutPath%
echo     ProjectDir: %ProjectDir%
echo.
echo [rsvars]
echo     Setting up variables for ms build.
call "C:\Program Files (x86)\Embarcadero\Studio\20.0\bin\rsvars.bat"
echo.
echo [Build]
echo     Building project.
msbuild %BuildProject% /t:%BuildMode% /p:config=%BuildConfiguration% /v:%BuildVerbosity% /p:DCC_ExeOutput=%OutPath%\%ProjectDir%\%BuildConfiguration%
echo.
echo ------------------------Build script complete------------------------
