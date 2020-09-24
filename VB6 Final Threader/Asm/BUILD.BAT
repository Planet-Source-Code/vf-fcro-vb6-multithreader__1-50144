@echo off
\masm32\bin\ml /c /coff /Cp abfr.asm
\masm32\bin\link /DLL /DEF:abfr.def /SUBSYSTEM:WINDOWS /LIBPATH:\masm32\lib abfr.obj abfr.res
pause