echo off
set TestProjectName=ProjectTest.dproj
set TestProjectPath=D:\Path\TestX\
set TestEnvironment="C:\TestingEnvironment\"
set TestBuildConfiguration=Release
call test_project %TestProjectName% %TestProjectPath% %TestEnvironment% %TestBuildConfiguration%