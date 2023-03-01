# Windows GUI Toolkit

This project was my first ever attempt at programming a large application using Visual Basic 6. It aims to provide a complete common controls replacement toolkit for Windows applications using ActiveX UserControl technology which emulate the Windows XP themes from the time. It even still works to this day on all versions since Windows 95 and I think it was quite an impressive achievement to build this.

The project includes common control replacements for the following:

* Titlebar		
* Border			
* TextBox		
* Frame 			
* OptionButton		
* Hyperlink		
* ProgressBar		
* ScrollBar		
* Slider			
* SpinEdit		
* CheckBox
* DriveBox			
* FolderList
* FileList
* ComboBox
* ListBox
* StatusBar
* PictureBox
* CommandButton
* ToolBarButton
* MenuBar

# Install Visual Studio 6.0

Unfortunately modern versions of Visual Studio no longer support ActiveX UserControl technology so it necessary to install Visual Basic 6 if you want to use these components in your own projects.

1. Download Visual Basic 6 ISO from https://winworldpc.com/product/microsoft-visual-bas/60
2. Mount the Visual Basic 6.0 Enterprise Edition.iso file
3. Install the application
4. Search for an OEM CD-Key online

# Building the Project

I would recommend before opening the project you run the Install.bat file from the Release folder. Right click on Install.bat and select Run as Administrator. This will copy all of the necessary files to the c:\Program Files\Windows GUI Toolkit folder and register them with the system. You can then open the Windows GUI Toolkit.vbg project group file. Please remember to open Visual Basic as an Administrator or you may encounter errors loading the project.

If you are using an operating system prior to Windows XP you will have to copy the Release\Windows GUI Toolkit folder to your Program Files directory manually and use the following commands to register the DLL's with the system. 

This is due to the %programfiles% environment variable used in the batch script not being available prior to Windows XP.

- cd Windows\System
- regsvr32 "c:\Program Files\Windows GUI Toolkit\SSubTmr6.dll"
- regsvr32 "c:\Program Files\Windows GUI Toolkit\PopupMenu.dll"
- regsvr32 "c:\Program Files\Windows GUI Toolkit\CommonControls.ocx"

Replace "c:\Program Files" with the path to your installation folder.

You will also need to the copy the Release\Windows GUI Toolkit\Trebucbd.TTF font to the c:\Windows\Fonts directory.

If you want to move or rename the Windows GUI Toolkit installation folder you must reregister the components with the system using regsvr32 and the new path so any applications you build can find the references.

If you prefer to build the DLL's yourself and make changes to the code then you will need to do the following:

1. Open Visual Basic 6 as an Administrator (important)
2. Open the Windows GUI Toolkit.vbg solution
3. Click on the SSubTmr6 project in the Project Group window 
   Select File -> Make SSubTmr6.dll and compile the DLL
4. Click on the Menu project in the Project Group Window 
   Select Project -> References... and add a reference to the SSubTmr6.dll created earlier
5. Select File -> Make Menu.dll and compile the DLL
6. Click on the CommonControls project in the Project Group window 
   Select Project -> References and add references to the SSubTmr6.dll and Menu.dll created earlier
7. Select File -> Make CommonControls.ocx to compile the ActiveX controls
8. Click on the LunaBlue project in Project Group window 
   Select File -> Make LunaBlue.dll
9. Repeat step 7 for the LunaGreen, LunaSilver and Win98 theme projects

# Using the Common Controls

1. Open Visual Basic 6 as an Administrator (important)
2. Create a New Standard EXE project
3. Select Project -> Components... -> 
   Browse to the file C:\Program Files\Windows GUI Toolkit\CommonControls.ocx and click Apply
4. You will see the common control icons displayed in the Toolbox and you will be able to add them to the form

# Modifying Theme Resources

If you want to modify the theme resource files you will need to add the VB 6 Resource Editor.

1. From the main menu select Add-Ins -> Add-In Manager... 
2. Select VB 6 Resource Editor and check Loaded / Unloaded and Load on Startup
3. Open a theme project such as LunaBlue in the Solution Explorer
4. Expand the Related Documents folder and double click on the Graphics.RES resource file
5. In the VB Resource Editor expand the Bitmap node and you will see all of the images which make up the theme
6. The raw image files are packaged in the CommonControls project for example: 
   Windows GUI Toolkit\CommonControls\Resources\Control Graphics\Blue for the LunaBlue theme
7. Find the name of the resource you want to edit. For example FORMSKIN\HASFOCUS\TITLEBAR\LEFT
8. Right click the resource and select properties then copy the ID string to the clipboard
9. Select the resource and press the delete key to remove it from the file
10. Click the Add Bitmap... button in the resource editor toolbar
11. Browse to the raw image bitmap and open it. This will add a new resource to the file named 101
12. Right click the new resource and select properties then paste the resource name you copied earlier
13. Save the resource file
14. From the main menu select File -> Make LunaBlue.thm
15. Browse to the c:\Program Files\Windows GUI Toolkit folder and overwrite the LunaBlue.dll file 

# Creating a Menu

The MenuBar component allows you to easily create menus which can have multiple nested sub-menus, checked items, option buttons and icons. It includes a menu editor property page to help create the structure which is stored in a .dat resource file.

To create a new menu:

1. Add the MenuBar component to the form
2. Right click on the component and select Properties
3. Give the menu item a name and then add the sub items
4. Optionally specify a image index for the menu items from the 
   C:\Program Files\Windows GUI Toolkit\ImageStrip.bmp file
5. The default path for the .dat file which stores the menu structure is created at c:\Menu\menu.dat
   I would recommend when you compile the program you copy this file to the executable path and add the following line to Form_Load() to find the file: MenuBar1.MenuPath = App.Path & "\menu.dat"

# Note

Unfortunaltely modern browsers and Windows Defender flag some Visual Basic 6 executables as potential threats. Therefore I have been unable to provide a release distribution for this program. You will have to compile the source code yourself to run it.

# Acknowledgements

I have to mention a special thanks to the website vbaccelerator.com which provided classes to handle subclassing controls for mouse and window focus events and creating popup menus. Without the examples provided by this site this project would have not been possible.













