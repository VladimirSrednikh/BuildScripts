if exist "%ProgramFiles(x86)%" (
  call "%ProgramFiles(x86)%""\Embarcadero\RAD Studio\12.0\bin\rsvars.bat"
) else (
  call "%ProgramFiles%""\Embarcadero\RAD Studio\12.0\bin\rsvars.bat"
)
cscript Build_Projects.vbs
pause