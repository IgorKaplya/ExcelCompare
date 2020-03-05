echo off
set TestProjectName=%1
set TestProjectPath=%2
set TestEnvironment=%3
set TestBuildConfiguration=%4

set BuildMode=Build
set OutTmp=%CD%\_out_tmp_test\
echo ------------------------Starting test script------------------------
echo [Conf]
echo     Tested project: %TestProjectPath%%TestProjectName%
echo     Build configuration: %TestBuildConfiguration%
echo     Build mode: %BuildMode%
echo     OutTmp: %OutTmp%
echo     Test Environment: %TestEnvironment%
echo [Build]
call build_project %TestProjectPath%%TestProjectName% %TestBuildConfiguration% %BuildMode% %OutTmp%
echo.
echo [Deploy]
xcopy /s /y %OutTmp%\%TestProjectName%\%TestBuildConfiguration% %TestEnvironment%
echo.
echo [Cleanup]
RD /s /q %OutTmp%
echo.
echo [StartTest]
pushd %CD%
cd /d %TestEnvironment%
call %TestProjectName:~0,-5%exe
popd
echo.
echo [ErrorLevel]
echo     ErrorLevel: %errorlevel%
echo.
echo ------------------------Test script complete------------------------