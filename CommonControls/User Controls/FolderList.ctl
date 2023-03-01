VERSION 5.00
Begin VB.UserControl FolderList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   ScaleHeight     =   2685
   ScaleWidth      =   3735
   ToolboxBitmap   =   "FolderList.ctx":0000
   Begin CommonControls.MaskBox3 Right 
      Align           =   4  'Align Right
      Height          =   2430
      Left            =   3600
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   4286
      ScaleHeight     =   2430
      ScaleWidth      =   135
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Bottom 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      Top             =   2550
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   238
      ScaleHeight     =   135
      ScaleWidth      =   3735
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Left 
      Align           =   3  'Align Left
      Height          =   2430
      Left            =   0
      Top             =   120
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   4286
      ScaleHeight     =   2430
      ScaleWidth      =   150
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Top 
      Align           =   1  'Align Top
      Height          =   120
      Left            =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   212
      ScaleHeight     =   120
      ScaleWidth      =   3735
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin VB.DirListBox lstMain 
      Height          =   2340
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   3270
   End
End
Attribute VB_Name = "FolderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - FolderList.ctl                                '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipTitle As String
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_Path As String
Dim m_Enabled As Boolean

' Events
Event OLECompleteDrag(Effect As Long) 'MappingInfo=lstMain,lstMain,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstMain,lstMain,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=lstMain,lstMain,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=lstMain,lstMain,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=lstMain,lstMain,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=lstMain,lstMain,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event Change() 'MappingInfo=lstMain,lstMain,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Attribute Change.VB_MemberFlags = "200"
Event Click() 'MappingInfo=lstMain,lstMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event Scroll() 'MappingInfo=lstMain,lstMain,-1,Scroll
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Paints the control border.
Private Sub pPaintBorder()
    
    ' Set the enabled graphic path
    Dim strEnabled As String
    If m_Enabled = True Then
        strEnabled = "Enabled\"
    Else
        strEnabled = "Disabled\"
    End If
    
    ' Paint the border
    Top.Cls
    Left.Cls
    Bottom.Cls
    Right.Cls
   
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopLeft", crBitmap), 0, 0
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    Left.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Left", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Left.Height
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomLeft", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width, 3 * Screen.TwipsPerPixelY
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    Right.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Right", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Height

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

Public Sub Refresh()

    lstMain.Top = 1 * Screen.TwipsPerPixelY
    lstMain.Left = 1 * Screen.TwipsPerPixelX
    Right.Width = 3 * Screen.TwipsPerPixelX
    Bottom.Height = 3 * Screen.TwipsPerPixelY
    Left.Width = 3 * Screen.TwipsPerPixelX
    Top.Height = 3 * Screen.TwipsPerPixelY
    
    pPaintBorder
    lstMain.Enabled = m_Enabled

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
        
    m_Enabled = True
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipTitle = ""
    m_ToolTipStyle = Standard
    m_Path = ""
 
    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY
 
    Refresh

End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    SetDefaultTheme
    
    lstMain.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    Set lstMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lstMain.ForeColor = PropBag.ReadProperty("ForeColor", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    lstMain.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    lstMain.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    lstMain.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_Path = PropBag.ReadProperty("Path", "")
    lstMain.Path = m_Path
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
     
    If g_ControlsRefreshed = True Then
        Refresh
    End If
  
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", lstMain.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Font", lstMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lstMain.ForeColor, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", lstMain.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDragMode", lstMain.OLEDragMode, 0)
    Call PropBag.WriteProperty("OLEDropMode", lstMain.OLEDropMode, 0)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("Path", m_Path, "")

End Sub

Private Sub lstMain_Change()
    RaiseEvent Change
End Sub

Private Sub lstMain_Click()
    RaiseEvent Click
End Sub

Private Sub lstMain_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lstMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lstMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lstMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub lstMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub lstMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub lstMain_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    lstMain.Width = Width - (2 * Screen.TwipsPerPixelX)
    lstMain.Height = Height - (2 * Screen.TwipsPerPixelY)
    Height = lstMain.Height + (2 * Screen.TwipsPerPixelY)
    pPaintBorder

err:
    Exit Sub

End Sub

Private Sub lstMain_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub lstMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub lstMain_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub lstMain_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub lstMain_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub lstMain_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
    hwnd = lstMain.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,OLEDrag
Public Sub OLEDrag()
    lstMain.OLEDrag
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = lstMain.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lstMain.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    lstMain.Enabled = New_Enabled
    pPaintBorder
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lstMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lstMain.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = lstMain.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lstMain.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = lstMain.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set lstMain.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = lstMain.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    lstMain.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,OLEDragMode
Public Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = lstMain.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
    lstMain.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
    OLEDropMode = lstMain.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    lstMain.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Path() As String
    Path = lstMain.Path
End Property

Public Property Let Path(ByVal New_Path As String)
    m_Path = New_Path
    On Error GoTo Error_Handler
    lstMain.Path = m_Path
    PropertyChanged "Path"
Error_Handler:

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
Attribute ToolTipCaption.VB_Description = "Returns/sets the text displayed in the tooltip."
    ToolTipCaption = m_ToolTipCaption
End Property

Public Property Let ToolTipCaption(ByVal New_ToolTipCaption As String)
    m_ToolTipCaption = New_ToolTipCaption
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipIcon() As ToolTipIconEnum
Attribute ToolTipIcon.VB_Description = "Returns/sets the icon displayed in the tooltip."
    ToolTipIcon = m_ToolTipIcon
End Property

Public Property Let ToolTipIcon(ByVal New_ToolTipIcon As ToolTipIconEnum)
    m_ToolTipIcon = New_ToolTipIcon
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipTitle() As String
Attribute ToolTipTitle.VB_Description = "Returns/sets the title displayed in the tooltip."
    ToolTipTitle = m_ToolTipTitle
End Property

Public Property Let ToolTipTitle(ByVal New_ToolTipTitle As String)
    m_ToolTipTitle = New_ToolTipTitle
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipStyle() As ToolTipStyleEnum
Attribute ToolTipStyle.VB_Description = "Returns/sets the style of the tooltip i.e Standad or Balloon."
    ToolTipStyle = m_ToolTipStyle
End Property

Public Property Let ToolTipStyle(ByVal New_ToolTipStyle As ToolTipStyleEnum)
    m_ToolTipStyle = New_ToolTipStyle
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipStyle"
End Property
