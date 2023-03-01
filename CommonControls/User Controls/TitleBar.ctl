VERSION 5.00
Begin VB.UserControl TitleBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000013&
   CanGetFocus     =   0   'False
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   ScaleHeight     =   465
   ScaleWidth      =   6405
   ToolboxBitmap   =   "TitleBar.ctx":0000
   Begin CommonControls.MaskBox3 imgControlBox 
      Height          =   435
      Left            =   4815
      Top             =   0
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   767
      ScaleHeight     =   29
      ScaleWidth      =   106
   End
   Begin CommonControls.MaskBox3 imgCaptionTruncate 
      Height          =   435
      Left            =   4425
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   767
      ScaleHeight     =   435
      ScaleWidth      =   390
      ScaleMode       =   1
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   420
      TabIndex        =   0
      Top             =   75
      Width           =   855
   End
   Begin VB.Line lineFix 
      BorderColor     =   &H00C8D0D4&
      X1              =   -15
      X2              =   2565
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Shape shapeShadowFix 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H006A240A&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   360
      Top             =   60
      Width           =   4530
   End
   Begin VB.Label lblCaptionShadow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   270
      Left            =   375
      TabIndex        =   1
      Top             =   75
      Width           =   855
   End
End
Attribute VB_Name = "TitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - TitleBar.ctl                                  '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Implements ISubclass

' Member variables
Dim WithEvents m_ParentForm As Form
Attribute m_ParentForm.VB_VarHelpID = -1
Dim WithEvents m_MouseTrack As cMouseTrack
Attribute m_MouseTrack.VB_VarHelpID = -1
Dim m_Appearance As AppearanceEnum
Dim m_BorderStyle As BorderStyleEnum
Dim m_Buttons As ButtonsEnum
Dim m_Caption As String
Dim m_Icon As String
Dim m_WindowState As WindowStateEnum
Dim m_WindowStyle As WindowStyleEnum
Dim m_FocusPath As String
Dim m_ButtonsPath As String
Dim m_WindowStatePath As String
Dim m_Height As Long
Dim m_LostFocusTextColor As OLE_COLOR
Dim m_HasFocusTextColor As OLE_COLOR
Dim m_MinMouseOver As Boolean
Dim m_MaxMouseOver As Boolean
Dim m_CloseMouseOver As Boolean
Dim m_ButtonTop As Long
Dim m_ButtonBottom As Long
Dim m_MinLeft As Long
Dim m_MinRight As Long
Dim m_MaxLeft As Long
Dim m_MaxRight As Long
Dim m_WhatsThisLeft As Long
Dim m_WhatsThisRight As Long
Dim m_ToolWindowTop As Long
Dim m_ToolWindowBottom As Long
Dim m_ToolWindowLeft As Long
Dim m_ToolWindowRight As Long
Dim m_CloseLeft(3) As Long
Dim m_CloseRight(3) As Long
                        
' Enums
Public Enum WindowStateEnum
    Normal = 1
    Minimized = 2
    Maximized = 3
End Enum

' Events
Event frmMinimize()
Event frmMaximize()
Event frmRestore()
Event frmClose()
Event frmMove()
Event frmMoveable(Moveable As Boolean)
Event frmSizeable(Sizeable As Boolean)
Event frmLostFocus()
Event frmGotFocus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Initializes the DLL containing the theme and sets the style of the control based on the appearance property.
Private Sub pSetAppearance()

    Select Case m_Appearance
        Case Is = Blue
            g_ResourceLib.Filename = ProgramFilesPath() & "\Windows GUI Toolkit\LunaBlue.dll"
            g_Appearance = Blue
            m_LostFocusTextColor = &HE4CCBD
            m_HasFocusTextColor = &HFFFFFF
            lblCaptionShadow.ForeColor = &H404000
            Parent.BackColor = &HD8E9EC
        Case Is = Green
            g_ResourceLib.Filename = ProgramFilesPath() & "\Windows GUI Toolkit\LunaGreen.dll"
            g_Appearance = Green
            m_LostFocusTextColor = &HFFFFFF
            m_HasFocusTextColor = &HFFFFFF
            lblCaptionShadow.ForeColor = &H404000
            Parent.BackColor = &HD8E9EC
        Case Is = Silver
            g_ResourceLib.Filename = ProgramFilesPath() & "\Windows GUI Toolkit\LunaSilver.dll"
            g_Appearance = Silver
            m_LostFocusTextColor = &H80000013
            m_HasFocusTextColor = &H0&
            lblCaptionShadow.ForeColor = &HC8D0D4
            Parent.BackColor = &HE3DFE0
        Case Is = Win98
            lblCaptionShadow.Visible = False
            g_ResourceLib.Filename = ProgramFilesPath() & "\Windows GUI Toolkit\Win98.dll"
            g_Appearance = Win98
            m_LostFocusTextColor = &HC8D0D4
            m_HasFocusTextColor = &HFFFFFF
            Parent.BackColor = &HC8D0D4
    End Select
    
    If m_Appearance <> Win98 Then
            
        ' Hide line fix
        lineFix.Visible = False
        
        ' Hide shadow fix
        shapeShadowFix.Visible = False
                
        ' Make caption shadow visible
        lblCaptionShadow.Visible = True
    
        ' Set caption shadow position
        lblCaptionShadow.Font = "Trebuchet MS"
        lblCaptionShadow.FontSize = 10
        lblCaptionShadow.FontBold = True
        lblCaptionShadow.Top = 9 * Screen.TwipsPerPixelY
        lblCaptionShadow.Left = 29 * Screen.TwipsPerPixelX
        
        ' Set titlebar caption position
        If m_WindowStyle = StandardWindow Then
            lblCaption.Font = "Trebuchet MS"
            lblCaption.FontSize = 10
            lblCaption.FontBold = True
            lblCaption.Top = 8 * Screen.TwipsPerPixelY
            lblCaption.Left = 28 * Screen.TwipsPerPixelX
            lblCaptionShadow.Visible = True
            Height = 30 * Screen.TwipsPerPixelY
            m_Height = 30 * Screen.TwipsPerPixelY
        Else
            lblCaption.Font = "Tahoma"
            lblCaption.FontSize = 8
            lblCaption.Top = 6 * Screen.TwipsPerPixelY
            lblCaption.Left = 8 * Screen.TwipsPerPixelX
            lblCaptionShadow.Visible = False
            Height = 22 * Screen.TwipsPerPixelY
            m_Height = 22 * Screen.TwipsPerPixelY
        End If
        
        ' Initialize the position of the title bar buttons
        m_ButtonTop = 5
        m_ButtonBottom = 25
        m_MinLeft = 1
        m_MinRight = 21
        m_MaxLeft = 24
        m_MaxRight = 44
        m_WhatsThisLeft = 1
        m_WhatsThisRight = 21
        m_CloseLeft(1) = 46
        m_CloseRight(1) = 67
        m_CloseLeft(2) = 1
        m_CloseRight(2) = 21
        m_CloseLeft(3) = 24
        m_CloseRight(3) = 44
        m_ToolWindowLeft = 1
        m_ToolWindowRight = 14
        m_ToolWindowTop = 6
        m_ToolWindowBottom = 19
    
    Else ' Win98
           
        ' Show line fix
        lineFix.Visible = True
        lineFix.x2 = UserControl.Width
                           
        ' Show shadow fix
        shapeShadowFix.Visible = True
                                   
        ' Hide caption shadow
        lblCaptionShadow.Visible = False
        
        ' Set titlebar caption position
        If m_WindowStyle = StandardWindow Then
            lblCaption.Font = "Tahoma"
            lblCaption.FontSize = 8
            lblCaption.FontBold = True
            lblCaption.Top = 6 * Screen.TwipsPerPixelY
            lblCaption.Left = 25 * Screen.TwipsPerPixelX
            Height = 23 * Screen.TwipsPerPixelY
            m_Height = 23 * Screen.TwipsPerPixelY
        Else
            lblCaption.Font = "Tahoma"
            lblCaption.FontSize = 8
            lblCaption.FontBold = True
            lblCaption.Top = 5 * Screen.TwipsPerPixelY
            lblCaption.Left = 7 * Screen.TwipsPerPixelX
            Height = 20 * Screen.TwipsPerPixelY
            m_Height = 20 * Screen.TwipsPerPixelY
        End If
    
        m_ButtonTop = 6
        m_ButtonBottom = 20
        m_MinLeft = 1
        m_MinRight = 17
        m_MaxLeft = 17
        m_MaxRight = 33
        m_CloseLeft(1) = 35
        m_CloseRight(1) = 51
        m_CloseLeft(2) = 1
        m_CloseRight(2) = 17
        m_ToolWindowLeft = 1
        m_ToolWindowRight = 11
        m_ToolWindowTop = 6
        m_ToolWindowBottom = 16

    End If

End Sub

' Sets the titlebar and border graphics based on the form focus.
Public Sub pSetWindowFocus(bFocus As FocusEnum)

    ' Setup paths and caption color
    If bFocus = HasFocus Then
        g_FormHasFocus = True
        m_FocusPath = "HasFocus\"
        lblCaption.ForeColor = m_HasFocusTextColor
    Else
        g_FormHasFocus = False
        m_FocusPath = "LostFocus\"
        lblCaption.ForeColor = m_LostFocusTextColor
    End If

    ' Draw titlebar
    If m_WindowStyle = StandardWindow Then
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "TitleBar\Left", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "TitleBar\Center", crBitmap), 30 * Screen.TwipsPerPixelX, 0, Width '* 3, 1 * Screen.TwipsPerPixelY
        Set imgCaptionTruncate.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "TitleBar\Truncate", crBitmap)
    ' Toolwindow titlebar
    Else
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Toolwindow\Left", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Toolwindow\Center", crBitmap), 30 * Screen.TwipsPerPixelX, 0, Width ', m_Height
        Set imgCaptionTruncate.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Toolwindow\Truncate", crBitmap)
    
    End If

    ' Draw title bar buttons
    If WindowState = Normal Then
        If Buttons = All Then
            Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "Standard\" & "Default", crBitmap)
        Else
            If m_WindowStyle = StandardWindow Then
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & "ToolWindowClose\" & "DefaultMax", crBitmap)
            End If
        End If
    
    ElseIf WindowState = Minimized Then
        If Buttons = All Then
            Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "Standard\" & "Default", crBitmap)
        Else
            If m_WindowStyle = StandardWindow Then
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & "ToolWindowClose\" & "DefaultMax", crBitmap)
            End If
        End If
    
    ElseIf WindowState = Maximized Then
        If Buttons = All Then
            Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "Maximized\" & "Default", crBitmap)
        Else
            If m_WindowStyle = StandardWindow Then
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & "ToolWindowClose\" & "DefaultMax", crBitmap)
            End If
        End If
        
    End If

    ' Set the title bar icon
    pSetTitleBarIcon
    
    UserControl_Resize

End Sub

' Maximizes the form.
Public Sub Maximize()
        
    ' Set the window state to maximized
    m_ParentForm.WindowState = 2
      
    ' Get the dimensions of desktop
    Dim rDesktopSize As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, rDesktopSize, 0
           
    ' Resize window to fit desktop
    MoveWindow m_ParentForm.hwnd, rDesktopSize.Left - 4, rDesktopSize.Top - 4, (rDesktopSize.Right - rDesktopSize.Left) + 8, (rDesktopSize.Bottom - rDesktopSize.Top) + 8, True

End Sub

' Restores the form.
Private Sub pRestore()
    
    m_ParentForm.WindowState = 0

End Sub

' Minimizes the form.
Public Sub Minimize()
    CloseWindow (m_ParentForm.hwnd)
End Sub

' Closes the form.
Public Sub ExitApp()
    
    ' Post message to close application
    PostMessage m_ParentForm.hwnd, WM_CLOSE, 0&, 0&

End Sub

' Displays the form context menu.
Private Sub pShowContextMenu(ByVal x As Long, ByVal y As Long)

    ' Post message to show context menu.
    Call SendMessage(m_ParentForm.hwnd, &H313, 0, ByVal MakeLong(y, x))

End Sub

' Sets the icon displayed in the title bar.
Private Sub pSetTitleBarIcon()

    ' Titlebar icon variables
    Dim hImgSmall As Long   ' Handle to the system image list
    Dim fName As String     ' File name to get icon from
    Dim fnFilter As String  ' File name filter
    Dim R As Long

    If m_WindowStyle <> ToolWindow Then
    
        ' Get the system icon associated with that file
        hImgSmall& = SHGetFileInfo(m_Icon, 0&, _
                                    shinfo, Len(shinfo), _
                                    BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
On Error GoTo err
        
        UserControl.MouseIcon = LoadPicture(m_Icon)
        
        ' Draw the associated icon onto the usercontrol
        If m_Appearance <> Win98 Then
            Call ImageList_Draw(hImgSmall&, shinfo.iIcon, _
                                UserControl.hdc, 8, 8, ILD_TRANSPARENT)
        Else
            Call ImageList_Draw(hImgSmall&, shinfo.iIcon, _
                                UserControl.hdc, 5, 4, ILD_TRANSPARENT)
        
        End If
                           
    End If
        
err:
    Exit Sub

End Sub

' Sets the global variable containing the border style (used by border and statusbar controls).
Private Sub pSetBorderStyle()

    If m_BorderStyle = Sizable Then
        g_BorderStyle = Sizable
    ElseIf m_BorderStyle = Fixed Then
        g_BorderStyle = Fixed
    End If

End Sub

' Sets the graphics path for the title bar buttons.
Private Sub pSetButtons()

    ' Select the buttons which are displayed
    Select Case m_Buttons
        Case Is = All
            m_ButtonsPath = "All\"
        Case Is = MinClose
            m_ButtonsPath = "MinClose\"
        Case Is = CloseOnly
            m_ButtonsPath = "Close\"
        Case Is = None
            m_ButtonsPath = "None\"
        Case Is = WhatsThis
            m_ButtonsPath = "WhatsThis\"
    End Select
    
    ' Set the dimensions of the title bar buttons box
    If m_WindowStyle = StandardWindow Then
        
        If m_Appearance <> Win98 Then
        
            imgCaptionTruncate.Width = 16 * Screen.TwipsPerPixelX
            imgCaptionTruncate.Height = 30 * Screen.TwipsPerPixelY
        
            Select Case m_Buttons
                Case Is = All
                    imgControlBox.Width = 74 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 30 * Screen.TwipsPerPixelY
                Case Is = CloseOnly
                    imgControlBox.Width = 28 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 30 * Screen.TwipsPerPixelY
                Case Is = MinClose
                    imgControlBox.Width = 74 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 30 * Screen.TwipsPerPixelY
                Case Is = None
                    imgControlBox.Width = 43 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 30 * Screen.TwipsPerPixelY
                Case Is = WhatsThis
                    imgControlBox.Width = 50 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 29 * Screen.TwipsPerPixelY
            End Select
    
        Else
        
            shapeShadowFix.Width = UserControl.Width
                    
            imgCaptionTruncate.Width = 16 * Screen.TwipsPerPixelX
            imgCaptionTruncate.Height = 25 * Screen.TwipsPerPixelY
            
            Select Case m_Buttons
                Case Is = All
                    imgControlBox.Width = 57 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 23 * Screen.TwipsPerPixelY
                Case Is = CloseOnly
                    imgControlBox.Width = 23 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 23 * Screen.TwipsPerPixelY
                Case Is = MinClose
                    imgControlBox.Width = 57 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 23 * Screen.TwipsPerPixelY
                Case Is = None
                    imgControlBox.Width = 23 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 23 * Screen.TwipsPerPixelY
                Case Is = WhatsThis
                    imgControlBox.Width = 40 * Screen.TwipsPerPixelX
                    imgControlBox.Height = 21 * Screen.TwipsPerPixelY
            End Select
       
        End If

    Else ' Tool window
        
        shapeShadowFix.Visible = False
        
        If m_Appearance <> Win98 Then
            imgCaptionTruncate.Width = 12 * Screen.TwipsPerPixelX
            imgCaptionTruncate.Height = 22 * Screen.TwipsPerPixelY
            imgControlBox.Width = 20 * Screen.TwipsPerPixelX
            imgControlBox.Height = 22 * Screen.TwipsPerPixelY
        Else
            imgCaptionTruncate.Width = 12 * Screen.TwipsPerPixelX
            imgCaptionTruncate.Height = 20 * Screen.TwipsPerPixelY
            imgControlBox.Width = 17 * Screen.TwipsPerPixelX
            imgControlBox.Height = 20 * Screen.TwipsPerPixelY
        End If
    
    End If

End Sub

' Sets the forms window state.
Private Sub pSetWindowState()

    If m_WindowState = Normal Then
        
        m_WindowStatePath = "Standard\"
        If Ambient.UserMode = True Then
            pRestore
            RaiseEvent frmRestore
        End If
       
    ElseIf m_WindowState = Maximized Then
        
        m_WindowStatePath = "Maximized\"
        If Ambient.UserMode = True Then
            Maximize
            RaiseEvent frmMaximize
        End If
               
    ElseIf m_WindowState = Minimized Then
    
        m_WindowStatePath = "Standard\"
        If Ambient.UserMode = True Then
            Minimize
            RaiseEvent frmMinimize
        End If
        
    End If
    
    ' Update the titlebar buttons
    Buttons = Buttons

End Sub

' Clears the title bar button graphics.
Private Sub pClearTitleBarButtons()

    If m_WindowStyle = StandardWindow Then
    
        If m_Buttons = All Then
            Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "Default", crBitmap)
            m_MinMouseOver = False
        Else
            Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
        End If
        
    Else
        Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & "ToolWindowClose\" & "DefaultMax", crBitmap)
    
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    'Setup the control
    pSetAppearance
    pSetButtons
    
    'Redraw control with new appearance
    pSetWindowFocus HasFocus
    
    m_ParentForm_Resize

End Sub

' Initializes a tooltip for a control using the handle.
Public Sub CustomToolTip(ControlHwnd As Long, Caption As String, Title As String, Style As ToolTipStyleEnum, Icon As ToolTipIconEnum)
    pSetToolTip ControlHwnd, Caption, Title, Style, Icon
End Sub

' Returns the Windows temp path.
Public Function TempPath()
    TempPath = GetSettingString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\VolumeCaches\Temporary files", "folder", "") '& "\"
End Function

' Returns the users start menu location.
Public Function StartMenuPath()
    StartMenuPath = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Programs", "") & "\"
End Function

' Returns the users desktop location.
Public Function DesktopPath()
    DesktopPath = GetSettingString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop", "") & "\"
End Function

' Returns the system path.
Public Function SystemPath()

    Dim strFolder As String
    Dim lngResult As Long
    strFolder = String(MAX_PATH, 0)
    lngResult = GetSystemDirectory(strFolder, MAX_PATH)
    If lngResult <> 0 Then
        GetSystemPath = Left(strFolder, InStr(strFolder, _
        Chr(0)) - 1)
    Else
        GetSystemPath = ""
    End If
    
End Function

' Shows the color picker dialog.
Public Function ColorDialog(DefaultColor As Long, ShowFull As Boolean)
    ColorDialog = ColorDlg(m_ParentForm.hwnd, DefaultColor, ShowFull)
End Function

' Shows the open dialog.
Public Function OpenDialog(DialogTitle As String, Filter As String, InitialDir As String)

    Dim CommonDialogAPI As New cOpenSave
    Dim cmobj As Object
    Set cmobj = CommonDialogAPI
    With cmobj
        .hwnd = m_ParentForm.hwnd
        .DialogTitle = DialogTitle
        .Filter = Filter
        .FilterIndex = 1
        .InitDir = InitialDir
        .ShowOpen
        If .CancelError = True Then
            OpenDialog = "Cancel"
        Else
            OpenDialog = .Filename
        End If
    End With
End Function

' Shows the save dialog.
Public Function SaveDialog(DialogTitle As String, Filter As String, InitialDir As String)

    Dim CommonDialogAPI As New cOpenSave
    Dim cmobj As Object
    Set cmobj = CommonDialogAPI
    With cmobj
        .hwnd = m_ParentForm.hwnd
        .DialogTitle = DialogTitle
        .Filter = Filter
        .FilterIndex = 1
        .InitDir = InitialDir
        .ShowSave
        If .CancelError = True Then
            SaveDialog = "Cancel"
        Else
            SaveDialog = .Filename
        End If
    End With
End Function

' Shows the browse for folder dialog.
Public Function BrowseFolder(Title As String)
    BrowseFolder = BrowseForFolder(m_ParentForm.hwnd, Title, BIF_DONTGOBELOWDOMAIN)
End Function

' Executes a file.
Public Sub ShellApp(sFile As String)
    Call ShellExecute(0&, vbNullString, sFile, vbNullString, vbNullString, vbNormalFocus)
End Sub

' Creates a directory structure.
Public Sub CreateFolder(sPath As String)
    Dim DirReturn As Boolean
    DirReturn = CreateFolder2(sPath)
End Sub

' Get a filename without the path.
Public Function GetFileNameWithExt(sPath As String) As String
    GetFileNameWithExt = GetFileNameWithExt2(sPath)
End Function

' Get a filename without the path and extension.
Public Function GetFileName(sPath As String) As String
    GetFileName = GetFileName2(sPath)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Subclassing

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    m_emr = emrConsume
    ISubclass_MsgResponse = m_emr
End Property

' Tell the subclasser what to do for this message.
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    
    ' Check if the active window is about to be activated
    If iMsg = WM_ACTIVATEAPP Then
    
      If wParam = 0 Then
        
        pSetWindowFocus LostFocus
        RaiseEvent frmLostFocus
      
        If m_WindowStyle = StandardWindow Then
            lblCaptionShadow.Visible = False
        End If
      
        m_ParentForm.Refresh
      
      Else
      
        pSetWindowFocus HasFocus
        RaiseEvent frmGotFocus
  
        If m_WindowStyle = StandardWindow Then
            lblCaptionShadow.Visible = True
        End If
            
        m_ParentForm.Refresh

      End If
    
    End If

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when an application creates an instance of the control.
Private Sub UserControl_Initialize()

    ' Initializes global font size variable with the system font size
    SetGlobalFontSize

End Sub

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
 
    ' Initialize default properties
    If g_Appearance = Empty Then
        m_Appearance = Blue
    Else
        m_Appearance = g_Appearance
    End If
    
    m_BorderStyle = Sizable
    m_Buttons = All
    m_Caption = "TitleBar"
    g_MinFormHeight = 41
    g_MinFormWidth = 128
    m_Icon = ProgramFilesPath() & "\Windows GUI Toolkit\Icon.ico"
   
    m_WindowState = Normal
    m_WindowStyle = StandardWindow
   
    ' Run procedures to setup control
    pSetAppearance
    pSetButtons
    pSetWindowState
    
    lblCaption.Caption = m_Caption
    lblCaptionShadow.Caption = m_Caption

    ' Redraw Control
    pSetWindowFocus HasFocus
     
    Extender.Align = 1
    Width = 20
 

End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Check if the environment is in design mode or end user mode
    If Ambient.UserMode = True Then
        
        ' Set the global parent window hWnd
        Set m_ParentForm = UserControl.Parent
        g_hwnd = m_ParentForm.hwnd
                
        ' Attach the parent window focus message
        AttachMessage Me, m_ParentForm.hwnd, WM_ACTIVATEAPP
    
    End If
    
    ' Attach mouse tracking message to title bar buttons
    Set m_MouseTrack = New cMouseTrack
    m_MouseTrack.AttachMouseTracking imgControlBox.hwnd
         
    ' Read the properties
    m_Appearance = PropBag.ReadProperty("Appearance", Blue)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", Sizable)
    m_Buttons = PropBag.ReadProperty("Buttons", All)
    m_Caption = PropBag.ReadProperty("Caption", "TitleBar")
    g_MinFormHeight = PropBag.ReadProperty("MinHeight", 41)
    g_MinFormWidth = PropBag.ReadProperty("MinWidth", 128)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_WindowState = PropBag.ReadProperty("WindowState", Normal)
    m_WindowStyle = PropBag.ReadProperty("WindowStyle", StandardWindow)
    m_Icon = PropBag.ReadProperty("Icon", ProgramFilesPath() & "\Windows GUI Toolkit\Icon.ico")

    ' Run procedures to setup control
    pSetAppearance
    pSetBorderStyle
    pSetButtons
    pSetWindowState
    pSetTitleBarIcon
    
    lblCaption.Caption = m_Caption
    lblCaptionShadow.Caption = m_Caption
    SetWindowText g_hwnd, m_Caption
 
    ' Redraw Control
    pSetWindowFocus HasFocus
    
    ' Loop through all controls on parent form
    Dim ctrl As Object
    For Each ctrl In Parent.Controls
    
        ' Move to next control if it doesnt have a refresh property
        On Error Resume Next
        
            ' Update control appearance
            ctrl.Refresh
            
    Next
    
    g_ControlsRefreshed = True
       
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' Write properties to storage
    Call PropBag.WriteProperty("Appearance", m_Appearance, Blue)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, Sizable)
    Call PropBag.WriteProperty("Buttons", m_Buttons, All)
    Call PropBag.WriteProperty("Caption", m_Caption, "TitleBar")
    Call PropBag.WriteProperty("MinHeight", g_MinFormHeight, 41)
    Call PropBag.WriteProperty("MinWidth", g_MinFormWidth, 128)
    Call PropBag.WriteProperty("WindowState", m_WindowState, Normal)
    Call PropBag.WriteProperty("WindowStyle", m_WindowStyle, StandardWindow)
    Call PropBag.WriteProperty("Icon", m_Icon, ProgramFilesPath() & "\Windows GUI Toolkit\Icon.ico")
 
End Sub

Private Sub CaptionTruncate_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub CaptionTruncate_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseDown(Button, Shift, imgCaptionTruncate.Left + x, imgCaptionTruncate.Top + y)
End Sub

Private Sub CaptionTruncate_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseMove(Button, Shift, imgCaptionTruncate.Left + x, imgCaptionTruncate.Top + y)
End Sub

Private Sub imgControlBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If BorderStyle = Sizable Then
    
        ' Resize form
        If WindowState <> Maximized Then
     
            ' Resize right-up
            If x > imgControlBox.ScaleWidth - 15 And y < 4 Then
                imgControlBox.MousePointer = 6
                ReleaseCapture
                SendMessage m_ParentForm.hwnd, 274, 61445, 0
            ElseIf x > imgControlBox.ScaleWidth - 4 And y < 15 Then
                imgControlBox.MousePointer = 6
                ReleaseCapture
                SendMessage m_ParentForm.hwnd, 274, 61445, 0
            ' Resize up-down
            ElseIf x < imgControlBox.ScaleWidth - 15 And y < 4 Then
                imgControlBox.MousePointer = 7
                ReleaseCapture
                SendMessage m_ParentForm.hwnd, 274, 61443, 0
            ' Resize left-right
            ElseIf x > imgControlBox.ScaleWidth - 4 And y > 15 Then
                imgControlBox.MousePointer = 9
                ReleaseCapture
                SendMessage m_ParentForm.hwnd, 274, 61442, 0
            End If
       
        End If
    
    End If

    ' Set title bar buttons mouse down graphic
    If Button = 1 Then
    
        If Buttons = All Then
        
            ' Minimize button
            If m_MaxMouseOver = False And m_CloseMouseOver = False Then
                If x >= m_MinLeft And x <= m_MinRight And y >= m_ButtonTop And y <= m_ButtonBottom Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "MinPressed", crBitmap)
                    m_MinMouseOver = True
                End If
            ' Max/Restore button
            ElseIf m_MinMouseOver = False And m_CloseMouseOver = False Then
                If x >= m_MaxLeft And x <= m_MaxRight And y >= m_ButtonTop And y <= m_ButtonBottom Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "MaxPressed", crBitmap)
                    m_MaxMouseOver = True
                End If
            ' Close button
            ElseIf m_MinMouseOver = False And m_MaxMouseOver = False Then
                If x >= m_CloseLeft(1) And x <= m_CloseRight(1) And y >= m_ButtonTop And y <= m_ButtonBottom Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "ClosePressed", crBitmap)
                    m_CloseMouseOver = True
                End If
            End If
        
        End If
    
        If Buttons = CloseOnly Then
        
            If m_WindowStyle = StandardWindow Then
                
                ' Close button
                If m_MinMouseOver = False And m_MaxMouseOver = False Then
                    If x >= m_CloseLeft(2) And x <= m_CloseRight(2) And y >= m_ButtonTop And y <= m_ButtonBottom Then
                        Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & m_ButtonsPath & "ClosePressed", crBitmap)
                        m_CloseMouseOver = True
                    End If
                End If
        
            Else
        
                ' Close button
                If m_MinMouseOver = False And m_MaxMouseOver = False Then
                    If x >= m_ToolWindowLeft And x <= m_ToolWindowRight And y >= m_ToolWindowTop And y <= m_ToolWindowBottom Then
                        Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & "ToolWindowClose\" & "ClosePressed", crBitmap)
                        m_CloseMouseOver = True
                    End If
                End If
        
            End If
        
        End If
    
        If Buttons = MinClose Then
        
            ' Minimize button
            If m_MaxMouseOver = False And m_CloseMouseOver = False Then
                If x >= m_MinLeft And x <= m_MinRight And y >= m_ButtonTop And y <= m_ButtonBottom Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & m_ButtonsPath & "MinPressed", crBitmap)
                    m_MinMouseOver = True
                End If
            End If
        
            ' Close button
            If m_MinMouseOver = False And m_MaxMouseOver = False Then
                If x >= m_CloseLeft(1) And x <= m_CloseRight(1) And y >= m_ButtonTop And y <= m_ButtonBottom Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & m_ButtonsPath & "ClosePressed", crBitmap)
                    m_CloseMouseOver = True
                End If
            End If
        
        End If
        
        If Buttons = WhatsThis Then
        
            ' Max/Restore button
            If m_MinMouseOver = False And m_CloseMouseOver = False Then
                If x >= m_WhatsThisLeft And x <= m_WhatsThisRight And y >= m_ButtonTop And y <= m_ButtonBottom Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & m_ButtonsPath & "MaxPressed", crBitmap)
                End If
            End If
            
            ' Close button
            If m_MinMouseOver = False And m_MaxMouseOver = False Then
                If x >= m_CloseLeft(3) And x <= m_CloseRight(3) And y >= m_ButtonTop And y <= m_ButtonBottom Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\" & "Buttons\" & m_ButtonsPath & "ClosePressed", crBitmap)
                    m_CloseMouseOver = True
                End If
            End If

        End If
    
    End If

End Sub

Private Sub imgControlBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Tracking is initialised by entering the control:
    If Not (m_MouseTrack.Tracking) Then
        m_MouseTrack.StartMouseTracking
    End If
    
    If BorderStyle = Sizable Then
    
        If WindowState <> Maximized Then
     
            'Resize right-up
            If x > imgControlBox.ScaleWidth - 15 And y < 4 Then
                imgControlBox.MousePointer = 6
            ElseIf x > imgControlBox.ScaleWidth - 4 And y < 15 Then
                imgControlBox.MousePointer = 6
            'Resize up-down
            ElseIf x < imgControlBox.ScaleWidth - 15 And y < 4 Then
                imgControlBox.MousePointer = 7
            'Resize left-right
            ElseIf x > imgControlBox.ScaleWidth - 4 And y > 15 Then
                imgControlBox.MousePointer = 9
            End If
       
        End If
    
    End If
    
    ' Reset mouse pointer
    imgControlBox.MousePointer = 1
     
    If Buttons = All Then
    
        ' Minimize button
        If m_MaxMouseOver = False And m_CloseMouseOver = False Then
    
            If x >= m_MinLeft And x <= m_MinRight And y >= m_ButtonTop And y <= m_ButtonBottom Then
                If m_MinMouseOver = False Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "MinMouseOver", crBitmap)
                    imgControlBox.ToolTipText = "Minimize"
                    m_MinMouseOver = True
                End If
            
            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "Default", crBitmap)
                imgControlBox.ToolTipText = ""
                m_MinMouseOver = False
            End If
        End If
        
        ' Max / restore button
        If m_MinMouseOver = False And m_CloseMouseOver = False Then
       
            If x >= m_MaxLeft And x <= m_MaxRight And y >= m_ButtonTop And y <= m_ButtonBottom Then
                If m_MaxMouseOver = False Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "MaxMouseOver", crBitmap)
                    If WindowState = Normal Then
                        imgControlBox.ToolTipText = "Maximize"
                    Else
                        imgControlBox.ToolTipText = "Restore Down"
                    End If
                    m_MaxMouseOver = True
                End If
            
            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "Default", crBitmap)
                imgControlBox.ToolTipText = ""
                m_MaxMouseOver = False
            End If
            
        End If
                  
        ' Close button
        If m_MinMouseOver = False And m_MaxMouseOver = False Then
                      
            If x >= m_CloseLeft(1) And x <= m_CloseRight(1) And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                If m_CloseMouseOver = False Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "CloseMouseOver", crBitmap)
                    imgControlBox.ToolTipText = "Close"
                    m_CloseMouseOver = True
                End If
            
            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "Default", crBitmap)
                imgControlBox.ToolTipText = ""
                m_CloseMouseOver = False
            End If
            
         End If
        
    End If
        
    If Buttons = MinClose Then
    
        ' Minimize button
        If m_MaxMouseOver = False And m_CloseMouseOver = False Then
        
            If x >= m_MinLeft And x <= m_MinRight And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
        
                If m_MinMouseOver = False Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "MinMouseOver", crBitmap)
                    imgControlBox.ToolTipText = "Minimize"
                    m_MinMouseOver = True
                End If
        
            Else
              
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
                imgControlBox.ToolTipText = ""
                m_MinMouseOver = False
        
            End If
        
        End If

        ' Close button
        If m_MinMouseOver = False And m_MaxMouseOver = False Then
        
            If x >= m_CloseLeft(1) And x <= m_CloseRight(1) And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
        
                If m_CloseMouseOver = False Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "CloseMouseOver", crBitmap)
                    imgControlBox.ToolTipText = "Close"
                    m_CloseMouseOver = True
                End If

            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
                imgControlBox.ToolTipText = ""
                m_CloseMouseOver = False
            End If
        End If

    End If
   
    If Buttons = CloseOnly Then
            
        If m_WindowStyle = StandardWindow Then
                    
          ' Close button
            If m_MinMouseOver = False And m_MaxMouseOver = False Then
            
                If x >= m_CloseLeft(2) And x <= m_CloseRight(2) And y >= m_ButtonTop And y <= m_ButtonBottom Then 'Mouse over
                
                    If m_CloseMouseOver = False Then
                        Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "CloseMouseOver", crBitmap)
                        imgControlBox.ToolTipText = "Close"
                        m_CloseMouseOver = True
                    End If
                
                Else
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
                    imgControlBox.ToolTipText = ""
                    m_CloseMouseOver = False
                End If
                
            End If
          
        Else
        
            ' Close button
            If m_MinMouseOver = False And m_MaxMouseOver = False Then
            
                If x >= m_ToolWindowLeft And x <= m_ToolWindowRight And y >= m_ToolWindowTop And y <= m_ToolWindowBottom Then 'Mouse over
                
                    If m_CloseMouseOver = False Then
                        Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & "ToolWindowClose\" & "CloseMouseOver", crBitmap)
                        imgControlBox.ToolTipText = "Close"
                        m_CloseMouseOver = True
                    End If
                
                Else
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & "ToolWindowClose\" & "DefaultMax", crBitmap)
                    imgControlBox.ToolTipText = ""
                    m_CloseMouseOver = False
                End If
            End If

        End If
    
    End If
    
    If Buttons = WhatsThis Then
    
        'Max / restore button
        If m_MinMouseOver = False And m_CloseMouseOver = False Then
       
            If x >= m_WhatsThisLeft And x <= m_WhatsThisRight And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                
                If m_MaxMouseOver = False Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "MaxMouseOver", crBitmap)
                    imgControlBox.ToolTipText = "Help"
                    m_MaxMouseOver = True
                End If
            
            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
                imgControlBox.ToolTipText = ""
                m_MaxMouseOver = False
            End If
            
        End If
         
         
        ' Close button
        If m_MinMouseOver = False And m_MaxMouseOver = False Then
                      
            If x >= m_CloseLeft(3) And x <= m_CloseRight(3) And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                If m_CloseMouseOver = False Then
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "CloseMouseOver", crBitmap)
                    imgControlBox.ToolTipText = "Close"
                    m_CloseMouseOver = True
                End If
            
            Else
                Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "DefaultMax", crBitmap)
                imgControlBox.ToolTipText = ""
                m_CloseMouseOver = False
            End If
         End If

    End If

End Sub

Private Sub imgControlBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        If Buttons = All Then
            
            ' Minimize button
            If m_MaxMouseOver = False And m_CloseMouseOver = False Then
                If x >= m_MinLeft And x <= m_MinRight And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "m_MinMouseOver", crBitmap)
                    m_MinMouseOver = False
                    Minimize
                End If
            End If
            
            ' Max / restore button
            If m_MinMouseOver = False And m_CloseMouseOver = False Then
                If x >= m_MaxLeft And x <= m_MaxRight And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                    If WindowState = Maximized Then
                        WindowState = Normal
                    Else
                        WindowState = Maximized
                    End If
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "m_MaxMouseOver", crBitmap)
                    m_MaxMouseOver = False
                End If
             End If
            
            ' Close button
            If m_MinMouseOver = False And m_MaxMouseOver = False Then
                If x >= m_CloseLeft(1) And x <= m_CloseRight(1) And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & m_WindowStatePath & "m_CloseMouseOver", crBitmap)
                    m_CloseMouseOver = False
                    RaiseEvent frmClose
                    ExitApp
                End If
            End If
        
        End If
    
        If Buttons = CloseOnly Then
            
            If m_WindowStyle = StandardWindow Then
            
                ' Close button
                If m_MinMouseOver = False And m_MaxMouseOver = False Then
                    If x >= m_CloseLeft(2) And x <= m_CloseRight(2) And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                        Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "m_CloseMouseOver", crBitmap)
                        m_CloseMouseOver = False
                        RaiseEvent frmClose
                        ExitApp
                    End If
                 End If
        
            Else
            
                ' Close button
                If m_MinMouseOver = False And m_MaxMouseOver = False Then
                    If x >= m_ToolWindowLeft And x <= m_ToolWindowRight And y >= m_ToolWindowTop And y <= m_ToolWindowBottom Then  'Mouse over
                        Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & "ToolWindowClose\" & "m_CloseMouseOver", crBitmap)
                        m_CloseMouseOver = False
                        RaiseEvent frmClose
                        ExitApp
                    End If
                 End If
        
            End If
    
        End If
        
        If Buttons = MinClose Then
            
            ' Minimize button
            If m_MaxMouseOver = False And m_CloseMouseOver = False Then
                If x >= m_MinLeft And x <= m_MinRight And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "m_MinMouseOver", crBitmap)
                    m_MinMouseOver = False
                    Minimize
                    RaiseEvent frmMinimize
                End If
            End If
            
            ' Close button
            If m_MinMouseOver = False And m_MaxMouseOver = False Then
                If x >= m_CloseLeft(1) And x <= m_CloseRight(1) And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "m_CloseMouseOver", crBitmap)
                    m_CloseMouseOver = False
                    RaiseEvent frmClose
                    ExitApp
                End If
            End If
        
        End If
        
        If Buttons = WhatsThis Then
        
            ' Max / restore button
            If m_MinMouseOver = False And m_CloseMouseOver = False Then
                If x >= m_WhatsThisLeft And x <= m_WhatsThisRight And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "m_MaxMouseOver", crBitmap)
                    Screen.MouseIcon = imgControlBox.MouseIcon
                    Screen.MousePointer = 99
                    m_MaxMouseOver = False
                End If
             End If
            
            ' Close button
            If m_MinMouseOver = False And m_MaxMouseOver = False Then
               If x >= m_CloseLeft(3) And x <= m_CloseRight(3) And y >= m_ButtonTop And y <= m_ButtonBottom Then  'Mouse over
                    Set imgControlBox.Picture = PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Buttons\" & m_ButtonsPath & "m_CloseMouseOver", crBitmap)
                    m_CloseMouseOver = False
                    RaiseEvent frmClose
                    ExitApp
                End If
            End If
        
        End If
    
    End If

    pSetWindowFocus HasFocus

End Sub

Private Sub m_MouseTrack_MouseLeave()
    pClearTitleBarButtons
End Sub

Private Sub UserControl_DblClick()

    ' Maximize / restore window if title bar is double clicked
    If Buttons = All Then
        If WindowState = Maximized Then
            WindowState = Normal
        Else
            WindowState = Maximized
        End If
    End If

End Sub

Private Sub lblCaption_DblClick()
    Call UserControl_DblClick
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblCaption.MousePointer = 1
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Move form/show titlebar menu
Call UserControl_MouseDown(Button, Shift, x + 430, y + 120)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        If BorderStyle = Sizable Then
        
            ' Send message to resize form
            If WindowState <> Maximized Then
         
                ' Resize left-up
                If x < 15 * Screen.TwipsPerPixelX And y < 4 * Screen.TwipsPerPixelY Then
                    UserControl.MousePointer = 8
                    ReleaseCapture
                    SendMessage m_ParentForm.hwnd, 274, 61444, 0
                ElseIf x < 4 * Screen.TwipsPerPixelX And y < 15 * Screen.TwipsPerPixelY Then
                    UserControl.MousePointer = 8
                    ReleaseCapture
                    SendMessage m_ParentForm.hwnd, 274, 61444, 0
                ' Resize up-down
                ElseIf x > 15 * Screen.TwipsPerPixelX And y < 4 * Screen.TwipsPerPixelY Then
                    UserControl.MousePointer = 7
                    ReleaseCapture
                    SendMessage m_ParentForm.hwnd, 274, 61443, 0
                ' Resize right-left
                ElseIf x < 4 * Screen.TwipsPerPixelX And y > 15 * Screen.TwipsPerPixelY Then
                    UserControl.MousePointer = 9
                    ReleaseCapture
                    SendMessage m_ParentForm.hwnd, 274, 61441, 0
                'Resize Left-Right
                ElseIf x > UserControl.Width - (4 * Screen.TwipsPerPixelX) And y > 15 * Screen.TwipsPerPixelY Then
                    UserControl.MousePointer = 9
                    ReleaseCapture
                    SendMessage m_ParentForm.hwnd, 274, 61442, 0
                End If
           
            End If
        
        End If

    End If

    ' If mouse in titlebar area
    If y < 440 Then
        
        If Button = 1 Then
            
            ' If form isn't maximized raise form move event
            If WindowState <> Maximized Then
                
                RaiseEvent frmMove
                  
                ' Send message to move form
                Dim lngReturnValue As Long
                Call ReleaseCapture
                lngReturnValue = SendMessage(m_ParentForm.hwnd, WM_NCLBUTTONDOWN, _
                HTCAPTION, 0&)
                End If
                
        ElseIf Button = 2 Then
    
            ' Show context menu
            pShowContextMenu (m_ParentForm.Left / Screen.TwipsPerPixelX) + (x / Screen.TwipsPerPixelX), (m_ParentForm.Top / Screen.TwipsPerPixelY) + (y / Screen.TwipsPerPixelY)
        
        End If
    
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If BorderStyle = Sizable Then
    
        If WindowState <> Maximized Then
     
            ' Resize left-up
            If x < 15 * Screen.TwipsPerPixelX And y < (4 * Screen.TwipsPerPixelY) Then
                UserControl.MousePointer = 8
            ElseIf x < (4 * Screen.TwipsPerPixelY) And y < 15 * Screen.TwipsPerPixelY Then
            UserControl.MousePointer = 8
            ' Resize up-down
            ElseIf x > 15 * Screen.TwipsPerPixelX And y < (4 * Screen.TwipsPerPixelY) Then
            UserControl.MousePointer = 7
            ' Resize right-left
            ElseIf x < (4 * Screen.TwipsPerPixelY) And y > 15 * Screen.TwipsPerPixelY Then
                UserControl.MousePointer = 9
            ' Resize left-right
            ElseIf x > UserControl.Width - (4 * Screen.TwipsPerPixelY) And y > 15 * Screen.TwipsPerPixelY Then
                UserControl.MousePointer = 9
            End If
       
        End If
    
    End If
    
    UserControl.MousePointer = 1
    pClearTitleBarButtons

End Sub

Private Sub m_ParentForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    pClearTitleBarButtons
End Sub

                       
Private Sub UserControl_Resize()

On Error GoTo err

    Height = m_Height
    imgControlBox.Left = UserControl.Width - (imgControlBox.ScaleWidth * Screen.TwipsPerPixelX)
    
    If m_WindowStyle = StandardWindow Then
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "TitleBar\Center", crBitmap), 30 * Screen.TwipsPerPixelX, 0, Width ', imgControlBox.Height
    Else
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & m_FocusPath & "Toolwindow\Center", crBitmap), 30 * Screen.TwipsPerPixelX, 0, Width
    End If

    ' Shows the title bar caption truncate graphic
    If Width < (lblCaption.Left + lblCaption.Width) + (imgControlBox.ScaleWidth * Screen.TwipsPerPixelX) Then
        imgCaptionTruncate.Visible = True
        imgCaptionTruncate.Left = UserControl.Width - ((imgControlBox.ScaleWidth * Screen.TwipsPerPixelX) + imgCaptionTruncate.Width)
    Else
        imgCaptionTruncate.Visible = False
    End If

err:
    Exit Sub

End Sub

Private Sub m_ParentForm_Resize()

    ' Round the corners of the form
    If Ambient.UserMode = True Then
            
        If m_WindowStyle = StandardWindow Then

            If m_Appearance <> Win98 Then
                RoundCorners Parent.hwnd, m_ParentForm.ScaleWidth / Screen.TwipsPerPixelX, m_ParentForm.ScaleHeight / Screen.TwipsPerPixelY
            Else
                SquareCorners Parent.hwnd, m_ParentForm.ScaleWidth / Screen.TwipsPerPixelX, m_ParentForm.ScaleHeight / Screen.TwipsPerPixelY
            End If
        
        Else
            SquareCorners Parent.hwnd, m_ParentForm.ScaleWidth / Screen.TwipsPerPixelX, m_ParentForm.ScaleHeight / Screen.TwipsPerPixelY
        End If
    
        ' Allow the form to be maximized from the context menu
        If m_ParentForm.WindowState = 2 Then
            WindowState = Maximized
        ElseIf m_ParentForm.WindowState = 0 Then
            If m_WindowState <> Normal Then
                WindowState = Normal
            End If
        End If
    
    End If
    
End Sub

Private Sub m_ParentForm_Load()

    ' Set parent form icon
    Set m_ParentForm.Icon = UserControl.MouseIcon
    m_ParentForm.ScaleMode = 1
    
    ' Remove border and titlebar
    ParentTitleBar m_ParentForm.hwnd, False
    RemoveBorder m_ParentForm.hwnd, False
    
     'Set the forms window state when form is loaded
    If m_WindowState = Maximized Then
        m_WindowStatePath = "Maximized\"
        RaiseEvent frmMaximize
    ElseIf m_WindowState = Minimized Then
        m_WindowStatePath = "Standard\"
        Minimize
        RaiseEvent frmMinimize
    ElseIf m_WindowState = Normal Then
        m_WindowStatePath = "Standard\"
        RaiseEvent frmRestore
    End If
    
    ' Begin subclassing for window resize (mFormResize module)
    gHW = m_ParentForm.hwnd
    Hook
  
End Sub

Private Sub m_ParentForm_Unload(Cancel As Integer)

    ' Stop subclassing
    Unhook

End Sub

Public Function ProgramFiles()
 ProgramFiles = ProgramFilesPath
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Appearance() As AppearanceEnum
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceEnum)
    
    ' Initialize the appearance
    m_Appearance = New_Appearance
    pSetAppearance
    Buttons = Buttons
    
    ' Redraw control with new appearance
    pSetWindowFocus HasFocus
        
    ' Loop through the controls on parent form
    Dim ctrl As Object
    For Each ctrl In Parent.Controls
       
        ' Resume if the control does not have a refresh method
        On Error Resume Next
        ctrl.Refresh
    Next

    m_ParentForm_Resize

    PropertyChanged "Appearance"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get BorderStyle() As BorderStyleEnum
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleEnum)
  
    If New_BorderStyle = Sizable Then
        g_BorderStyle = Sizable
        m_BorderStyle = New_BorderStyle
    ElseIf New_BorderStyle = Fixed Then
        g_BorderStyle = Fixed
        m_BorderStyle = New_BorderStyle
    
    End If
    
    PropertyChanged "BorderStyle"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Buttons() As ButtonsEnum
    Buttons = m_Buttons
End Property

Public Property Let Buttons(ByVal New_Buttons As ButtonsEnum)

    If m_WindowStyle <> ToolWindow Then
        m_Buttons = New_Buttons
    Else
        m_Buttons = CloseOnly
    End If

    If Ambient.UserMode = False Then
        
        ' Enabled the title bar buttons
        Select Case m_Buttons
            Case Is = All
                Parent.ControlBox = True
                Parent.MinButton = True
                Parent.MaxButton = True
                Parent.WhatsThisButton = False
            Case Is = CloseOnly
                Parent.ControlBox = True
                Parent.MinButton = False
                Parent.MaxButton = False
                Parent.WhatsThisButton = False
            Case Is = MinClose
                Parent.ControlBox = True
                Parent.MinButton = True
                Parent.MaxButton = False
                Parent.WhatsThisButton = False
            Case Is = None
                Parent.ControlBox = False
                Parent.MinButton = False
                Parent.MaxButton = False
                Parent.WhatsThisButton = False
            Case Is = WhatsThis
                Parent.ControlBox = True
                Parent.MinButton = False
                Parent.MaxButton = False
                Parent.WhatsThisButton = True
        End Select
    
    End If

    ' Sets the graphics path of the buttons
    pSetButtons
    
    ' Refresh control graphics
    pSetWindowFocus HasFocus

    PropertyChanged "Buttons"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    lblCaptionShadow.Caption = New_Caption
    lblCaption.Caption = New_Caption
    SetWindowText g_hwnd, m_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Icon() As String
    Icon = m_Icon
End Property

Public Property Let Icon(ByVal New_Icon As String)
    m_Icon = New_Icon
    pSetTitleBarIcon
    PropertyChanged "Icon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinHeight() As Long
    MinHeight = g_MinFormHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Long)
    If Ambient.UserMode = True Then
        g_MinFormHeight = New_MinHeight
    Else
        MsgBox "The MinHeight property can only be set at runtime.", vbOKOnly + vbDefaultButton1 + vbInformation, "MinHeight"
    End If
    PropertyChanged "MinHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinWidth() As Long
    MinWidth = g_MinFormWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
    g_MinFormWidth = New_MinWidth
    PropertyChanged "MinWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get WindowState() As WindowStateEnum
    WindowState = m_WindowState
End Property

Public Property Let WindowState(ByVal New_WindowState As WindowStateEnum)
    m_WindowState = New_WindowState
    pSetWindowState
    PropertyChanged "WindowState"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get WindowStyle() As WindowStyleEnum
    WindowStyle = m_WindowStyle
End Property

Public Property Let WindowStyle(ByVal New_WindowStyle As WindowStyleEnum)
    
    m_WindowStyle = New_WindowStyle
    pSetAppearance

    If m_WindowStyle = StandardWindow Then
        
        Buttons = All
    
        'Set titlebar caption position
        If m_Appearance <> Win98 Then
            lblCaption.Font = "Trebuchet MS"
            lblCaption.FontSize = 10
            lblCaption.FontBold = True
            lblCaption.Top = 8 * Screen.TwipsPerPixelY
            lblCaption.Left = 28 * Screen.TwipsPerPixelX
            lblCaptionShadow.Visible = True
            imgCaptionTruncate.Width = 16 * Screen.TwipsPerPixelX
            imgCaptionTruncate.Height = 30 * Screen.TwipsPerPixelY

        Else
            lblCaption.Font = "Tahoma"
            lblCaption.FontSize = 8
            lblCaption.FontBold = True
            lblCaption.Top = 6 * Screen.TwipsPerPixelY
            lblCaption.Left = 25 * Screen.TwipsPerPixelX
            lblCaptionShadow.Visible = False
            imgCaptionTruncate.Width = 16 * Screen.TwipsPerPixelX
            imgCaptionTruncate.Height = 25 * Screen.TwipsPerPixelY
        End If

    Else ' Tool window
    
        Buttons = CloseOnly
        
        ' Set titlebar caption position
        If m_Appearance <> Win98 Then
            lblCaption.Font = "Tahoma"
            lblCaption.FontSize = 8
            lblCaption.FontBold = True
            lblCaption.Top = 6 * Screen.TwipsPerPixelY
            lblCaption.Left = 3 * Screen.TwipsPerPixelX
            lblCaptionShadow.Visible = False
            imgCaptionTruncate.Width = 12 * Screen.TwipsPerPixelX
            imgCaptionTruncate.Height = 22 * Screen.TwipsPerPixelY
        Else
            lblCaption.Font = "Tahoma"
            lblCaption.FontSize = 8
            lblCaption.FontBold = True
            lblCaption.Top = 5 * Screen.TwipsPerPixelY
            lblCaption.Left = 6 * Screen.TwipsPerPixelX
            lblCaptionShadow.Visible = False
            imgCaptionTruncate.Width = 12 * Screen.TwipsPerPixelX
            imgCaptionTruncate.Height = 20 * Screen.TwipsPerPixelY
        End If
    End If

    PropertyChanged "WindowStyle"
    
End Property
