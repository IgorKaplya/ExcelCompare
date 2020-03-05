echo off
set BuildProject="FullPath\Project.dproj"
set BuildConfiguration=Release
set BuildMode=Build
set OutPath=%CD%\_out\
build_project %BuildProject% %BuildConfiguration% %BuildMode% %OutPath%

pause