Attribute VB_Name = "mDeclares"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - mDeclares.bas                                 '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Last type of border added to the form
Public g_LastBorderType As String

' Has the titlebar already set appearance
Public g_ControlsRefreshed As Boolean 'Holds whether the titlebar has already refreshed controls

' The hWnd of the parent form
Public g_hwnd As Long

Public g_Appearance As AppearanceEnum
Public g_BorderStyle As BorderStyleEnum 'Used by statusbar/borders to check borderstyle
Public g_MinFormWidth As Long
Public g_MinFormHeight As Long 'Used by statusbar/borders/titlebar to hold the minimum height the form can be resized to
Public g_SystemFontSize As String 'Holds whether system is using Small fonts
Public g_StatusVisible As Boolean 'Holds whether a status bar is visible
Public g_FormHasFocus As Boolean

'Objects:
Public m_emr As EMsgResponse 'Used to hold the SubClass message response
Public ToolTip As New cTooltip 'Used to set a tooltip for a control

' Used to retrieve control graphics from theme dll
Public g_ResourceLib As New cLibrary

'Functions:

'Draw text/focus rectange
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'Public Const DT_CENTER = &H1
'Public Const DT_LEFT = &H0
'Public Const DT_RIGHT = &H2

Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

'Pause computer
Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

'Set system color
Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long

'Round form corners
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, _
ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

'Draw Small tilebar icon from file
Public Const MAX_PATH = 260
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000 'system icon index
Public Const SHGFI_LARGEICON = &H0 'large icon
Public Const SHGFI_SMALLICON = &H1 'small icon
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
                                 SHGFI_SHELLICONSIZE Or _
                                 SHGFI_SYSICONINDEX Or _
                                 SHGFI_DISPLAYNAME Or _
                                 SHGFI_EXETYPE

Public Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Public Declare Function SHGetFileInfo Lib "shell32.dll" _
   Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" _
   (ByVal hIml As Long, ByVal i As Long, _
    ByVal hDCDest As Long, ByVal x As Long, _
    ByVal y As Long, ByVal flags As Long) As Long

Public shinfo As SHFILEINFO

' Blitting functions
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" _
   (ByVal hdc As Long) As Long


Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Const SRCCOPY = &HCC0020
    Public Const SRCINVERT = &H660046
    Public Const BLACKNESS = &H42
    Public Const WHITENESS = &HFF0062
    Public Const SRCAND = &H8800C6
    Public Const SRCERASE = &H440328
    Public Const SRCPAINT = &HEE0086

Declare Function CreateBitmap Lib "gdi32" _
  (ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal nPlanes As Long, _
   ByVal nBitCount As Long, _
   lpBits As Any) As Long

' Creates a bitmap in memory:
Declare Function CreateCompatibleBitmap Lib "gdi32" _
        (ByVal hdc As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long


' Sets the backcolour of a device context:
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
        ByVal crColor As Long) As Long

'Write text to screen & drawing functions
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal _
  x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor _
  As Long) As Long

Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As _
  Long) As Long
Public Const TEXT_TRANSPARENT = 1
Public Const TEXT_OPAQUE = 2

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINT_TYPE) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Type POINT_TYPE
  x As Long
  y As Long
End Type

Dim mlngOldPen, mlngNewPen As Long

'Get Desktop width/height
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48

'Type RECT
'   Left As Long
'   Top As Long
'    Right As Long
'    Bottom As Long
'End Type

'Check form focus sub class constant
Public Const WM_ACTIVATEAPP = 28

'Check sytem font size declare
Public Declare Function GetDesktopWindow Lib _
"user32" () As Long

Public Declare Function GetDC Lib "user32" (ByVal _
hwnd As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" _
(ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Declare Function ReleaseDC Lib "user32" _
(ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Const LOGPIXELSX = 88

'Internet connection declares

'Check internet connection
Public Const ERROR_SUCCESS = 0&
Public Const APINULL = 0&
Public ReturnCode As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey _
As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, lpData As Any, lpcbData As Long) As Long

'Disconnect from internet
Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412

Public Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Public Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Public Declare Function RasEnumConnections Lib _
"rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As _
Any, lpcb As Long, lpcConnections As Long) As Long

Public Declare Function RasHangUp Lib "rasapi32.dll" Alias _
"RasHangUpA" (ByVal hRasConn As Long) As Long

Public gstrISPName As String

'Get window System path
Declare Function GetSystemDirectory Lib "KERNEL32" Alias _
"GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal _
nSize As Long) As Long
'Public Const MAX_PATH = 260

'Restart/Shutdown Computer
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'Call ExitWindowsEx(2, 0) = Restart
'Call ExitWindowsEx(1, 0) = ShutDown

'Show combo box drop down
Public Declare Function SendMessageLong Lib _
"user32" Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
ByVal lParam As Long) As Long

Public Const CB_SHOWDROPDOWN = &H14F

'Checks the state of a mouse\keyboard button
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Integer

'Check for mouse over
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

'Opens link in IE
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Set windows caption
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

'Keep window on top
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

'Check form focus declarations
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long

'Sets Position/Width/Height of window
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Show in taskbar
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_APPWINDOW = &H40000

'Rounded form corner declarations
Public Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, _
    ByVal nPolyFillMode As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
    (ByVal hwnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Long) As Long
Public XY() As POINTAPI
Public mlWidth             As Long
Public mlHeight            As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

'Move form delclarations
Public Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'Resize form
Public Const SIZE_SE = &HF008&
Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
'Public Const SRCCOPY = &HCC0020

'Minimize form
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

'Close form
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_CLOSE = &H10

'Hide forms titlebar declarations / Remove border
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

'Form border constant
Public Const WS_THICKFRAME = &H40000

'SetCapture directs ALL mouse input to the window that
'has the mouse "captured".
Public Declare Function GetCapture& Lib "user32" ()
Public Declare Function SetCapture& Lib "user32" (ByVal hwnd&)
Public Declare Function ReleaseCapture& Lib "user32" ()

'Make a directory
Declare Function MakeSureDirectoryPathExists Lib _
    "IMAGEHLP.DLL" (ByVal DirPath As String) As Long

'Minimum form size declarations
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Type MINMAXINFO
ptReserved As POINTAPI
ptMaxSize As POINTAPI
ptMaxPosition As POINTAPI
ptMinTrackSize As POINTAPI
ptMaxTrackSize As POINTAPI
End Type

Public Const WM_GETMINMAXINFO = &H24

' Owner draw item draw:
Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemId As Long
    ItemAction As Long
    ItemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    ItemData As Long
End Type

Private Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Returns the version of Windows that the user is running
Public Function GetWindowsVersion() As String
    Dim osv As OSVERSIONINFO
    osv.OSVSize = Len(osv)

    If GetVersionEx(osv) = 1 Then
        Select Case osv.PlatformID
            Case VER_PLATFORM_WIN32s
                GetWindowsVersion = "Win32s on Windows 3.1"
            Case VER_PLATFORM_WIN32_NT
                GetWindowsVersion = "Windows NT"
                
                Select Case osv.dwVerMajor
                    Case 3
                        GetWindowsVersion = "Windows NT 3.5"
                    Case 4
                        GetWindowsVersion = "Windows NT 4.0"
                    Case 5
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows 2000"
                            Case 1
                                GetWindowsVersion = "Windows XP"
                            Case 2
                                GetWindowsVersion = "Windows Server 2003"
                        End Select
                    Case 6
                        Select Case osv.dwVerMinor
                            Case 0
                                GetWindowsVersion = "Windows Vista/Server 2008"
                            Case 1
                                GetWindowsVersion = "Windows 7/Server 2008 R2"
                            Case 2
                                GetWindowsVersion = "Windows 8/Server 2012"
                            Case 3
                                GetWindowsVersion = "Windows 8.1/Server 2012 R2"
                        End Select
                End Select
        
            Case VER_PLATFORM_WIN32_WINDOWS:
                Select Case osv.dwVerMinor
                    Case 0
                        GetWindowsVersion = "Windows 95"
                    Case 90
                        GetWindowsVersion = "Windows Me"
                    Case Else
                        GetWindowsVersion = "Windows 98"
                End Select
        End Select
    Else
        GetWindowsVersion = "Unable to identify your version of Windows."
    End If
End Function

Public Function IsMouseOver(Ctl As Control) As Boolean
    
    Dim typPt As POINTAPI
    Dim mOver As Long
    Dim hwnd As Long
    Dim CtlLeft&, CtlTop&, CtlRight&, CtlBottom&
    
    On Error Resume Next
    
    'Initialize Variables
    hwnd = 0
    err.Number = 0
    
    'Get controls handle
    hwnd = Ctl.hwnd
    
    'if control does not have a handle, an error is raised
    If err.Number > 0 Then
        'Get the handle of the control's parent control or form
        hwnd = Ctl.Container.hwnd
        
        'Get current cursor position
        Call GetCursorPos(typPt)
        
        'Get the handle of the control under these coordinates
        mOver = WindowFromPoint(typPt.x, typPt.y)
        
        'If the returned control handle is equal to the parent
        'control handle then the mouse is over that parent control
        If mOver <> hwnd Then
            IsMouseOver = False
            Exit Function
        End If
        
        'Get the rect of the questioned control
        'If the window's scalemode property is Pixels
        'then remove the TwipsPerPixel calculations
        CtlLeft = Ctl.Left / Screen.TwipsPerPixelX
        CtlTop = Ctl.Top / Screen.TwipsPerPixelY
        CtlRight = (Ctl.Left + Ctl.Width) / Screen.TwipsPerPixelX
        CtlBottom = (Ctl.Top + Ctl.Height) / Screen.TwipsPerPixelY
        
        'Convert the mouses screen position to the
        'mouses parent control position
        Call ScreenToClient(hwnd, typPt)
        
        'If the mouse is within the questioned control's
        'coordinates then the mouse is over the questioned control
        If typPt.y >= CtlTop And typPt.y <= CtlBottom And typPt.x >= CtlLeft And typPt.x <= CtlRight Then
            IsMouseOver = True
        Else
            IsMouseOver = False
        End If
        
        'Reset error number
        err.Number = 0
        
        'Stop here
        Exit Function
    End If
    
    'Questioned control has a handle so check it directly
    
    'Reset Variables
    err.Number = 0
    hwnd = Ctl.hwnd
    
    'Get current cursor position
    Call GetCursorPos(typPt)
    
    'Get the handle of the control under these coordinates
    mOver = WindowFromPoint(typPt.x, typPt.y)
    
    'If the returned control handle is equal to the questioned
    'control handle then the mouse is over that control
    If mOver = hwnd Then
        IsMouseOver = True
    Else
        IsMouseOver = False
    End If
End Function

Public Function GetAccessKeyFromString(New_String As String)

Dim intcount As Integer
Dim strtemp As String
Dim strsingle As String
Dim strFinal As String
Dim blnGotShortCut As Boolean

blnGotShortCut = False

For intcount = 1 To Len(New_String)

    strtemp = Left(New_String, intcount)
    strsingle = Right(strtemp, 1)
    
    If blnGotShortCut = True Then
    GetAccessKeyFromString = strsingle
    blnGotShortCut = False
    End If
    
    If strsingle = "&" Then
    strFinal = strFinal & strsingle
    blnGotShortCut = True
    End If

Next

End Function

Public Function NumberFromString(strText As String) As String

For intcount = 1 To Len(strText)

    strtemp = Left(strText, intcount)
    strsingle = Right(strtemp, 1)
    
    If IsNumeric(strsingle) Then
    strFinal = strFinal & strsingle
    End If

Next

NumberFromString = strFinal

End Function

Public Sub pSetToolTip(hwnd As Long, Caption As String, Title As String, ByVal Style As ToolTipStyleEnum, ByVal Icon As ToolTipIconEnum)

    ToolTip.SetParentHwnd hwnd
    ToolTip.TipText = Caption
    
    If Style = Standard Then
    ToolTip.Style = TTStandard
    Else
    ToolTip.Style = TTBalloon
    End If
    
    If Icon = Error Then
    ToolTip.Icon = TTIconError
    
    ElseIf Icon = Info Then
    ToolTip.Icon = TTIconInfo
    
    ElseIf Icon = Warning Then
    ToolTip.Icon = TTIconWarning
    
    ElseIf Icon = NoIcon Then
    ToolTip.Icon = TTNoIcon
    End If

    ToolTip.ForeColor = 0
    ToolTip.BackColor = 0
    ToolTip.Title = Title
    ToolTip.Create

End Sub

Public Function ProgramFilesPath()
ProgramFilesPath = GetSystemDrive() & "\Program Files"
End Function

Private Function GetSystemDrive() As String
    GetSystemDrive = Space(1000)
    Call GetWindowsDirectory(GetSystemDrive, Len(GetSystemDrive))
    GetSystemDrive = Left$(GetSystemDrive, 2)
End Function


Public Function InIDE() As Boolean
  On Error Resume Next
  Debug.Print 0 / 0
  InIDE = err.Number <> 0
End Function


Public Function Is32Bit() As Boolean

If GetWindowsVersion() = "Windows 95" Or GetWindowsVersion() = "Windows Me" Or GetWindowsVersion() = "Windows 98" Then
    Is32Bit = True
Else
    Is32Bit = False
End If

End Function


Public Sub SetDefaultTheme()

    If g_ResourceLib.hModule = 0 Then
        g_ResourceLib.Filename = ProgramFilesPath() & "\Windows GUI Toolkit\LunaBlue.dll"
    End If

End Sub

Public Sub SetPen(lngDC As Long, intDrawWidth As Integer, lngColour As Long)

    'Create the pen object and apply it
    mlngNewPen = CreatePen(PS_SOLID, intDrawWidth, lngColour)
    mlngOldPen = SelectObject(lngDC, mlngNewPen)

End Sub

Public Sub LineDraw(lngX1 As Long, lngY1 As Long, lngX2 As Long, lngY2 As Long, lngDC As Long)

Dim pt As POINT_TYPE

    'Move current pen x,y
    Call MoveToEx(lngDC, lngX1, lngY1, pt)
    'Draw line from current x,y to given x,y
    Call LineTo(lngDC, lngX2, lngY2)

End Sub

Public Sub DrawBorder(New_Hdc As Long, New_Width As Long, New_Height As Long, New_Appearance As AppearanceEnum, New_Enabled As Boolean)

'Used by draw border holds the color of the border before its drawn
Dim bdrColor As Long

'Convert height/width to pixels
New_Width = New_Width / Screen.TwipsPerPixelX
New_Height = New_Height / Screen.TwipsPerPixelY

Select Case New_Appearance

Case Is = Blue
    
    If New_Enabled = True Then
    bdrColor = &HB99D7F
    
    Else
    bdrColor = &HBAC7C9
    
    End If

Case Is = Green

    If New_Enabled = True Then
    bdrColor = &H7FB9A4
    
    Else
    bdrColor = &HBAC7C9
    
    End If


Case Is = Silver

    If New_Enabled = True Then
    bdrColor = &HB2ACA5
    
    Else
    bdrColor = &HBAC7C9
    
    End If

End Select


    'Top
    SetPen New_Hdc, 1, bdrColor
    LineDraw 0, 0, New_Width, 0, New_Hdc

    SetPen New_Hdc, 2, &HFFFFFF
    LineDraw 0, 2, New_Width, 2, New_Hdc
    
    'Left
    SetPen New_Hdc, 1, bdrColor
    LineDraw 0, 0, 0, New_Height, New_Hdc

    SetPen New_Hdc, 2, &HFFFFFF
    LineDraw 2, 2, 2, New_Height, New_Hdc
    
    'Bottom
    SetPen New_Hdc, 1, bdrColor
    LineDraw 0, New_Height - 1, New_Width, New_Height - 1, New_Hdc

    SetPen New_Hdc, 2, &HFFFFFF
    LineDraw 2, New_Height - 2, New_Width, New_Height - 2, New_Hdc
        
    'Right
    SetPen New_Hdc, 1, bdrColor
    LineDraw New_Width - 1, 0, New_Width - 1, New_Height, New_Hdc

    SetPen New_Hdc, 2, &HFFFFFF
    LineDraw New_Width - 2, 2, New_Width - 2, New_Height - 2, New_Hdc

End Sub


Public Function CreateFolder2(NewPath) As Boolean

    'Add a trailing slash if none
    If Right(NewPath, 1) <> "\" Then
        NewPath = NewPath & "\"
    End If
    
    'Call API
    If MakeSureDirectoryPathExists(NewPath) <> 0 Then
        'No errors, return True
        CreateFolder2 = True
    End If

End Function

Public Function GetFileNameWithExt2(sPath As String) As String
'Get a filename without the path
Dim posn As Integer, i As Integer
Dim fName As String
posn = 0
'find the position of the last "\" character in filename
For i = 1 To Len(sPath)
    If (Mid(sPath, i, 1) = "\") Then posn = i
Next i
'get filename without path
fName = Right(sPath, Len(sPath) - posn)
GetFileNameWithExt2 = fName
End Function

Public Function GetFileName2(sPath As String) As String
Dim posn As Integer, i As Integer
Dim fName As String
posn = 0
'find the position of the last "\" character in filename
For i = 1 To Len(sPath)
    If (Mid(sPath, i, 1) = "\") Then posn = i
Next i

'get filename without path
fName = Right(sPath, Len(sPath) - posn)

'get filename without extension
posn = InStr(fName, ".")
If posn <> 0 Then
    fName = Left(fName, posn - 1)
End If
GetFileName2 = fName
End Function

Public Sub SetGlobalFontSize()


If g_SystemFontSize = "" Then 'Only find font size if it has not already been set

'Returns whether system is using small fonts
Dim hWndDesk As Long
Dim hDCDesk As Long
Dim logPix As Long
Dim R As Long
hWndDesk = GetDesktopWindow()
hDCDesk = GetDC(hWndDesk)
logPix = GetDeviceCaps(hDCDesk, LOGPIXELSX)
R = ReleaseDC(hWndDesk, hDCDesk)
'IsScreenFontSmall = (logPix = 96)

If logPix = 96 Then
g_SystemFontSize = "Small"
Else
g_SystemFontSize = "Large"
End If

End If

End Sub

Public Function WrongFont() As Long
'Returns the system font size
Dim hWndDesk As Long
Dim hDCDesk As Long
Dim logPix As Long
Dim R As Long
hWndDesk = GetDesktopWindow()
hDCDesk = GetDC(hWndDesk)
WrongFont = GetDeviceCaps(hDCDesk, LOGPIXELSX)
'MsgBox logPix
R = ReleaseDC(hWndDesk, hDCDesk)
'WrongFont = (logPix = 96)
End Function

Public Sub Connect2()
'Connects to internet if not already connected
Dim DefaultConnect As String

'Gets the default inet connection
DefaultConnect = GetSettingString(HKEY_CURRENT_USER, "RemoteAccess", "Default")

'If CheckConnection = False Then
lResult = Shell("rundll32.exe rnaui.DLL,RnaDial " & DefaultConnect, 1)
'End If


End Sub


Public Function CheckConnection2() As Boolean
'Returns the internet status
Dim hKey As Long
Dim lpSubKey As String
Dim phkResult As Long
Dim lpValueName As String
Dim lpReserved As Long
Dim lpType As Long
Dim lpData As Long
Dim lpcbData As Long
CheckConnection2 = False
lpSubKey = "System\CurrentControlSet\Services\RemoteAccess"
ReturnCode = RegOpenKey(HKEY_LOCAL_MACHINE, lpSubKey, phkResult)
If ReturnCode = ERROR_SUCCESS Then
hKey = phkResult
lpValueName = "Remote Connection"
lpReserved = APINULL
lpType = APINULL
lpData = APINULL
lpcbData = APINULL
ReturnCode = RegQueryValueEx(hKey, lpValueName, _
lpReserved, lpType, ByVal lpData, lpcbData)
lpcbData = Len(lpData)
ReturnCode = RegQueryValueEx(hKey, lpValueName, _
lpReserved, lpType, lpData, lpcbData)
If ReturnCode = ERROR_SUCCESS Then
  If lpData = 0 Then
   CheckConnection2 = False
  Else
   CheckConnection2 = True
  End If
End If
RegCloseKey (hKey)
End If
End Function

Public Sub Disconnect2()
'Disconnects the internet

'Add code to close a browser windows so no message appears

Dim i As Long
Dim lpRasConn(255) As RasConn
Dim lpcb As Long
Dim lpcConnections As Long
Dim hRasConn As Long
lpRasConn(0).dwSize = RAS_RASCONNSIZE
lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
lpcConnections = 0
ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, _
lpcConnections)

If ReturnCode = ERROR_SUCCESS Then
    For i = 0 To lpcConnections - 1
        If Trim(ByteToString(lpRasConn(i).szEntryName)) _
            = Trim(gstrISPName) Then
            hRasConn = lpRasConn(i).hRasConn
            ReturnCode = RasHangUp(ByVal hRasConn)
        End If
    Next i
End If


End Sub

Public Function ByteToString(bytString() As Byte) As String
'Needed for disocnnect sub
Dim i As Integer
ByteToString = ""
i = 0
While bytString(i) = 0&
ByteToString = ByteToString & Chr(bytString(i))
i = i + 1
Wend
End Function


Public Function ParentTitleBar(FormHwnd As Long, Visible As Boolean)
'Shows/Hides the forms titlebar

Dim nStyle As Long
Dim tR As RECT
GetWindowRect FormHwnd, tR
 
'Retrieve current style bits.
nStyle = GetWindowLong(FormHwnd, GWL_STYLE)


If Visible Then
    nStyle = nStyle Or WS_CAPTION
Else
    nStyle = nStyle And Not WS_CAPTION
End If

'Set property
Call SetWindowLong(FormHwnd, GWL_STYLE, nStyle)

SetWindowPos FormHwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Function

Public Function RemoveBorder(FormHwnd As Long, Visible As Boolean)
'Removes the forms border

Dim tR As RECT
GetWindowRect FormHwnd, tR

Dim nStyle As Long
   
'Retrieve current style bits.
nStyle = GetWindowLong(FormHwnd, GWL_STYLE)


If Visible Then
    nStyle = nStyle Or WS_THICKFRAME
Else
    nStyle = nStyle And Not WS_THICKFRAME
End If

'Set property
Call SetWindowLong(FormHwnd, GWL_STYLE, nStyle)

'Refresh window
SetWindowPos FormHwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Function


Public Function ShowInTaskBar2(FormHwnd As Long, Visible As Boolean)

Dim tR As RECT
GetWindowRect FormHwnd, tR

Call LockWindowUpdate(FormHwnd)
Call ShowWindow(FormHwnd, vbHide)

Dim nStyleEx As Long
' Retrieve current style bits.
nStyleEx = GetWindowLong(FormHwnd, GWL_EXSTYLE)
   
' Attempt to set requested bit On or Off,
' and redraw.
If Visible Then
    nStyleEx = nStyleEx Or WS_EX_APPWINDOW
Else
    nStyleEx = nStyleEx And Not WS_EX_APPWINDOW
End If

'Set property
Call SetWindowLong(FormHwnd, GWL_EXSTYLE, nStyleEx)

'Refresh window
SetWindowPos FormHwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED

Call ShowWindow(FormHwnd, vbNormalFocus)
Call LockWindowUpdate(0&)

End Function


Public Sub RoundCorners(ParentHwnd, ParentWidth, ParentHeight As Long)

Top = CreateRoundRectRgn(ParentWidth + 1, 60, 0, 0, 15, 15)
Bottom = CreateRectRgn(0, 50, ParentWidth, ParentHeight)
Full = CreateRectRgn(0, 0, 0, 0)

Combined = CombineRgn(Full, Top, Bottom, 2)

SetWindowRgn ParentHwnd, Full, True

End Sub

Public Sub SquareCorners(ParentHwnd, ParentWidth, ParentHeight As Long)

Full = CreateRectRgn(0, 0, ParentWidth, ParentHeight)

SetWindowRgn ParentHwnd, Full, True

End Sub


Public Function MakeLong(ByVal WordHi As Integer, ByVal WordLo As Integer) As Long
   ' High word is coerced to Long to allow it to
   ' overflow limits of multiplication which shifts
   ' it left.
   MakeLong = (CLng(WordHi) * &H10000) Or (WordLo And &HFFFF&)
End Function

