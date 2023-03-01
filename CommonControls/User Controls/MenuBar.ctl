VERSION 5.00
Begin VB.UserControl MenuBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   PropertyPages   =   "MenuBar.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   4335
   ToolboxBitmap   =   "MenuBar.ctx":0013
   Begin VB.Shape MenuFix 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   2790
      Top             =   0
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label MenuButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   8
      Top             =   120
      Width           =   45
   End
   Begin VB.Label MenuButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   7
      Top             =   120
      Width           =   45
   End
   Begin VB.Label MenuButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   405
      TabIndex        =   6
      Top             =   120
      Width           =   45
   End
   Begin VB.Label MenuButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   510
      TabIndex        =   5
      Top             =   120
      Width           =   45
   End
   Begin VB.Label MenuButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   615
      TabIndex        =   4
      Top             =   120
      Width           =   45
   End
   Begin VB.Label MenuButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   705
      TabIndex        =   3
      Top             =   120
      Width           =   45
   End
   Begin VB.Label MenuButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   810
      TabIndex        =   2
      Top             =   120
      Width           =   45
   End
   Begin VB.Label MenuButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   915
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label MenuButton 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   45
   End
End
Attribute VB_Name = "MenuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - MenuBar.ctl                                   '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Implements ISubclass

' Member variables
Dim m_MenuBorder As Boolean
Dim m_MenuGradientFill As Boolean
Dim m_MenuImageStrip As String
Dim m_MenuPath As String
Dim m_MenuBackColor As OLE_COLOR
Dim m_MenuIndex As Long
Private WithEvents m_MenuObject As cPopupMenu
Attribute m_MenuObject.VB_VarHelpID = -1
Private m_ImageList As New cImageList
Private WithEvents m_MouseTrack As cMouseTrack
Attribute m_MouseTrack.VB_VarHelpID = -1

' Events
Event Click(MenuIndex As Integer, ItemIndex As Integer)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Deserializes menu object from file and creates menu.
Private Sub pReadFromFile()

    ' Hide all menu items
    Dim intcount As Integer
    For intcount = 1 To 9
        MenuButton(intcount).Visible = False
        MenuButton(intcount).AutoSize = True
    Next
    
    ' Check menu exists
    If Dir(MenuPath) = "" Then
        Exit Sub
    End If
    
    ' Deserialize menu data from file
    On Error Resume Next
    m_MenuObject.RestoreFromFile , MenuPath
    
    ' Retrieve array of stored menu structures
    Dim storeMenu() As cStoreMenu
    storeMenu = m_MenuObject.RetrieveStoredMenu
       
    ' Loop through the menu buttons
    Dim Menu As Variant
    intcount = 1
    For Each Menu In storeMenu
     
        If (storeMenu(intcount).MenuName <> "") Then
     
            Dim a As String
            a = storeMenu(intcount).MenuName
     
            MenuButton(intcount).Visible = True
            MenuButton(intcount).Caption = " " & Right(storeMenu(intcount).MenuName, Len(storeMenu(intcount).MenuName) - 1) & " "
                
            ' Set menu item width
            If intcount <> 1 Then
                MenuButton(intcount).Left = MenuButton(intcount - 1).Left + MenuButton(intcount - 1).Width
            End If
            MenuButton(intcount).Height = 250

            
            ' Win98 fix
            If g_Appearance <> Win98 Then
                MenuButton(intcount).Top = 4 * Screen.TwipsPerPixelY
            Else
                MenuButton(intcount).Top = 6 * Screen.TwipsPerPixelY
            End If
        End If
    
        intcount = intcount + 1
    Next

End Sub

' Shows the menu for the menu button clicked.
Private Sub pShowMenu(iMenuIndex As Integer)
    
    If MenuButton(iMenuIndex).Caption <> "" Then   'show menu
    
        Call MenuButton_MouseMove(iMenuIndex, 1, 1, 1, 1)
        
        ' Restore menu
        m_MenuObject.Restore (iMenuIndex & Trim(MenuButton(iMenuIndex).Caption))

        ' Show the menu and store the return value from the menu item selected
        Dim lMenuIndex As Long
        lMenuIndex = m_MenuObject.ShowPopupMenu(MenuButton(iMenuIndex).Left, MenuButton(iMenuIndex).Top + MenuButton(iMenuIndex).Height)
        
        ' If a menu item has been selected
        If lMenuIndex > 0 Then
          
            ' Set checked / toggled state of menu items
            If InStr(m_MenuObject.ItemKey(lMenuIndex), "Option") <> 0 Then
                m_MenuObject.GroupToggle lMenuIndex
            ElseIf InStr(m_MenuObject.ItemKey(lMenuIndex), "Check") <> 0 Then
                m_MenuObject.Checked(lMenuIndex) = Not (m_MenuObject.Checked(lMenuIndex))
            End If
            
            ' Store the menu with changes made
            m_MenuObject.Store iMenuIndex & Trim(MenuButton(iMenuIndex).Caption)
            
            m_MenuObject.StoreToFile , MenuPath
            
            RaiseEvent Click(CInt(iMenuIndex), CInt(lMenuIndex))
          
        End If
            
        MenuButton(iMenuIndex).BackStyle = 0
        MenuButton(iMenuIndex).ForeColor = &H0&
        MenuFix.Visible = False
    
    End If

End Sub

' Paints the control.
Private Sub pPaintComponent(bFocus As Boolean)

    ' Draw background
    If g_Appearance <> Win98 Then
        UserControl.PaintPicture LoadResPicture("WinXP\ToolBarBackGround\Back", vbResBitmap), 0, 0, Width, Height
        UserControl.PaintPicture LoadResPicture("WinXP\ToolBarBackGround\Bottom", vbResBitmap), 0, Height - (2 * Screen.TwipsPerPixelY), Width, 2 * Screen.TwipsPerPixelY
    Else
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Back", vbResBitmap), 0, 0, Width, Height
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Bottom", vbResBitmap), 0, Height - (2 * Screen.TwipsPerPixelY), Width, 2 * Screen.TwipsPerPixelY
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Top", vbResBitmap), 0, 0, Width, 2 * Screen.TwipsPerPixelY
        
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\TopLeft", vbResBitmap), 4 * Screen.TwipsPerPixelX, 0
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Left", vbResBitmap), 4 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, 2 * Screen.TwipsPerPixelY, Height
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\BottomLeft", vbResBitmap), 4 * Screen.TwipsPerPixelX, Height - (2 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Right", vbResBitmap), Width - (6 * Screen.TwipsPerPixelX), 0, 2 * Screen.TwipsPerPixelY, Height
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\TopRight", vbResBitmap), Width - (6 * Screen.TwipsPerPixelX), 0, 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\BottomRight", vbResBitmap), Width - (6 * Screen.TwipsPerPixelX), Height - (2 * Screen.TwipsPerPixelY)
    
    End If

    ' Draw border
    If bFocus = True Then
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\Borders\Left", crBitmap), 0, 0, , Height
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\Borders\Right", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 0, 4 * Screen.TwipsPerPixelX, Height
    
    Else
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\LostFocus\Borders\Left", crBitmap), 0, 0, , Height
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\LostFocus\Borders\Right", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 0, 4 * Screen.TwipsPerPixelX, Height
    
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control
Public Sub Refresh()

    pReadFromFile
    pPaintComponent True
    
    Select Case g_Appearance
        Case Is = Blue
            m_MenuBackColor = &HC56A1D
        Case Is = Green
            m_MenuBackColor = &H70A093
        Case Is = Silver
            m_MenuBackColor = &HBFB4B2
        Case Is = Win98
            m_MenuBackColor = &H7F0000
        End Select
    
    m_MenuObject.Appearance = g_Appearance
    
    UserControl_Resize

End Sub

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
            pPaintComponent False
        Else
            pPaintComponent True
        End If

    End If

End Function

Private Sub Label1_Click()
pShowMenu (1)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when an application creates an instance of the control.
Private Sub UserControl_Initialize()

    ' Initialize image list
    Set m_ImageList = New cImageList
    m_ImageList.Create
    m_ImageList.ColourDepth = ILC_COLOR32
                
    ' Initialize menu object
    Set m_MenuObject = New cPopupMenu
    m_MenuObject.hWndOwner = UserControl.hwnd
    m_MenuObject.ImageList = m_ImageList.hIml

End Sub

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    ' Initialize default theme
    SetDefaultTheme
        
    ' Initialize default properties
    m_MenuBorder = False
    m_MenuGradientFill = False
    m_MenuImageStrip = ProgramFilesPath() & "\Windows GUI Toolkit\ImageStrip.bmp"
    m_MenuPath = "C:\Menu\menu.dat"

    ' Update control
    Refresh

End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Initialize default theme
    SetDefaultTheme

    ' Check if the environment is in design mode or end user mode
    If Ambient.UserMode = True Then
        
        ' Attach the parent window focus message
        AttachMessage Me, Parent.hwnd, WM_ACTIVATEAPP
    
    End If

    ' Initialize the mouse tracking
    Set m_MouseTrack = New cMouseTrack
    m_MouseTrack.AttachMouseTracking UserControl.hwnd

    ' Read the properties
    m_MenuBorder = PropBag.ReadProperty("MenuBorder", False)
    m_MenuGradientFill = PropBag.ReadProperty("MenuGradientFill", False)
    m_MenuImageStrip = PropBag.ReadProperty("MenuImageStrip", "")
    m_MenuPath = PropBag.ReadProperty("MenuPath", "C:\Menu\menu.dat")
    m_ImageList.AddFromFile m_MenuImageStrip, IMAGE_BITMAP
    m_MenuObject.DrawBorder = m_MenuBorder
    m_MenuObject.GradientHighlight = m_MenuGradientFill

    ' Update control
    If g_ControlsRefreshed = True Then
        Refresh
    End If

End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    ' Write properties to storage
    Call PropBag.WriteProperty("MenuBorder", m_MenuBorder, False)
    Call PropBag.WriteProperty("MenuGradientFill", m_MenuGradientFill, False)
    Call PropBag.WriteProperty("MenuImageStrip", m_MenuImageStrip, "")
    Call PropBag.WriteProperty("MenuPath", m_MenuPath, "C:\Menu\menu.dat")

End Sub

' Occurs when all references to the control are removed from memory by setting all the variables that refer to the object to Nothing.
Private Sub UserControl_Terminate()

    ' Clean up
    m_ImageList.Destroy
    Set m_ImageList = Nothing
    
    m_MenuObject.DestroySubClass
    Set m_MenuObject = Nothing

End Sub

Private Sub m_MouseTrack_MouseLeave()

    ' Loop through menu buttons
    For intcount = 1 To 9
        
        MenuButton(intcount).MousePointer = 1
        
        If m_MenuIndex = 0 Then
            MenuFix.Visible = False
        End If
        
        ' For all the button which dont have the mouse over reset backcolor
        If intcount <> m_MenuIndex Then
            MenuButton(intcount).BackStyle = 0
            MenuButton(intcount).ForeColor = &H0&
        End If
    Next

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    m_MenuIndex = 0
    
    ' Tracking is initialised by entering the control
    If Not (m_MouseTrack.Tracking) Then
        m_MouseTrack.StartMouseTracking
    End If

End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    If g_Appearance <> Win98 Then
        Height = 24 * Screen.TwipsPerPixelY
    Else
        Height = 26 * Screen.TwipsPerPixelY
    End If
    
    pPaintComponent True

err:
    Exit Sub

End Sub

Private Sub MenuButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Show menu for the selected index
    m_MenuIndex = Index
    If Button = 1 Then 'left click
        pShowMenu (Index)
    End If
End Sub

Private Sub MenuButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
       
    ' Tracking is initialised by entering the control
    If Not (m_MouseTrack.Tracking) Then
        m_MouseTrack.StartMouseTracking
    End If
    
    ' Show menu fix
    MenuFix.Visible = True
    MenuFix.Top = MenuButton(Index).Top - (3 * Screen.TwipsPerPixelY)
    
    If MenuFix.Left <> MenuButton(Index).Left Then

        MenuFix.Left = MenuButton(Index).Left
        MenuFix.Width = MenuButton(Index).Width
    End If
    
    ' Change buttons backcolor
    MenuButton(Index).BackStyle = 1
    MenuButton(Index).BackColor = m_MenuBackColor
    MenuFix.BorderColor = m_MenuBackColor
    MenuFix.FillColor = m_MenuBackColor
    
    If g_Appearance <> Silver Then
        MenuButton(Index).ForeColor = 16777215
    End If
        
    ' Loop through menu buttons
    Dim intcount As Integer
    For intcount = 1 To 9
    MenuButton(intcount).MousePointer = 1
        
        ' For all the button which dont have the mouse over reset backcolor
        If intcount <> Index Then
            MenuButton(intcount).BackStyle = 0
            MenuButton(intcount).ForeColor = &H0&
        End If
    Next

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get MenuBorder() As Boolean
    MenuBorder = m_MenuBorder
End Property

Public Property Let MenuBorder(ByVal New_MenuBorder As Boolean)
    m_MenuBorder = New_MenuBorder
    m_MenuObject.DrawBorder = m_MenuBorder
    PropertyChanged "MenuBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get MenuGradientFill() As Boolean
    MenuGradientFill = m_MenuGradientFill
End Property

Public Property Let MenuGradientFill(ByVal New_MenuGradientFill As Boolean)
    m_MenuGradientFill = New_MenuGradientFill
    m_MenuObject.GradientHighlight = m_MenuGradientFill
    PropertyChanged "MenuGradientFill"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get MenuImageStrip() As String
    MenuImageStrip = m_MenuImageStrip
End Property

Public Property Let MenuImageStrip(ByVal New_MenuImageStrip As String)
    m_MenuImageStrip = New_MenuImageStrip
    PropertyChanged "MenuImageStrip"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get MenuPath() As String
    MenuPath = m_MenuPath
End Property

Public Property Let MenuPath(ByVal New_MenuPath As String)
    m_MenuPath = New_MenuPath
    pReadFromFile
    PropertyChanged "MenuPath"
End Property
