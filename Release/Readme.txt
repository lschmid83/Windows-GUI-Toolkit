Windows GUI Toolkit
===================

This project provides a complete common controls replacement set for Windows applications in Visual Basic 6 using ActiveX UserControl technology which emulate the Windows XP themes.

Install Visual Basic 6
=======================

1. Download Visual Basic 6 ISO from https://winworldpc.com/product/microsoft-visual-bas/60
2. Mount the Visual Basic 6.0 Enterprise Edition.iso file
3. Install the application
4. Search for an OEM CD-Key online

Installation
============

1. Run the Install.bat file as an Administrator
2. Open Visual Basic 6 as an Administrator
3. Create a New Standard EXE project
4. Select Project -> Components... from the main menu
5. Click Browse... and select c:\Program Files\Windows GUI Toolkit\Common Controls.ocx
6. You will see the new components loaded in the Toolbox and you can add them to the form

If you are using an operating system prior to Windows XP you will have to copy the Release\Windows GUI Toolkit folder to your Program Files directory manually and use the following commands to register the DLL's with the system. 

cd Windows\System
regsvr32 "c:\Program Files\Windows GUI Toolkit\SSubTmr6.dll"
regsvr32 "c:\Program Files\Windows GUI Toolkit\PopupMenu.dll"
regsvr32 "c:\Program Files\Windows GUI Toolkit\CommonControls.ocx"

Replace "c:\Program Files" with the path to your installation folder.

This is due to the %programfiles% environment variable used in the batch script not being available prior to Windows XP.

You will also need to the copy the Release\Windows GUI Toolkit\Trebucbd.TTF font to the c:\Windows\Fonts directory.

Copyright 1999-2001 Lawrence Schmid  