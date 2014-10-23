@echo off
rem Compile mebmd5.pas
rem Some day this will be replaced by a decent Makefile (fpmake?)

del *.o
del *.ppu

c:\lazarus\fpc\2.6.4\bin\i386-win32\ppc386.exe -dWIN32 -TWIN32 -O1 -FuC:\lazarus\fpc\2.6.4\units\i386-win32 -FuC:\lazarus\lcl\units\i386-win32 -FuC:\lazarus\lcl\units\i386-win32\win32 -FuC:\lazarus\lcl -FuC:\lazarus\components\lazutils\lib\i386-win32 -FuC:\lazarus\packager\units\i386-win32 mebmd5.pas

