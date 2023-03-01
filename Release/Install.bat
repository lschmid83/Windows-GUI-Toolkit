xcopy /E "%~dp0Windows GUI Toolkit" "%ProgramFiles%\Windows GUI Toolkit\"
regsvr32 /s "%ProgramFiles%\Windows GUI Toolkit\SSubTmr6.dll"
regsvr32 /s "%ProgramFiles%\Windows GUI Toolkit\PopupMenu.dll"
regsvr32 /s "%ProgramFiles%\Windows GUI Toolkit\CommonControls.ocx"
pause