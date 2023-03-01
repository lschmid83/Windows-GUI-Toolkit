VERSION 5.00
Begin VB.UserControl FileList 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ScaleHeight     =   2970
   ScaleWidth      =   3855
   ToolboxBitmap   =   "FileList.ctx":0000
   Begin CommonControls.MaskBox3 Top 
      Align           =   1  'Align Top
      Height          =   120
      Left            =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   212
      ScaleHeight     =   120
      ScaleWidth      =   3855
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Left 
      Align           =   3  'Align Left
      Height          =   2715
      Left            =   0
      Top             =   120
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   4789
      ScaleHeight     =   2715
      ScaleWidth      =   150
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Bottom 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      Top             =   2835
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   238
      ScaleHeight     =   135
      ScaleWidth      =   3855
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Right 
      Align           =   4  'Align Right
      Height          =   2715
      Left            =   3720
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   4789
      ScaleHeight     =   2715
      ScaleWidth      =   135
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin VB.FileListBox lstMain 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3585
   End
End
Attribute VB_Name = "FileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - FileList.ctl                                  '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
Dim m_BackColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_ForeColor As OLE_COLOR
Dim m_Path As String

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
Event Click() 'MappingInfo=lstMain,lstMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event DblClick() 'MappingInfo=lstMain,lstMain,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
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
Event PathChange() 'MappingInfo=lstMain,lstMain,-1,PathChange
Attribute PathChange.VB_Description = "Occurs when the path is changed by setting the FileName or Path property in code."
Event PatternChange() 'MappingInfo=lstMain,lstMain,-1,PatternChange
Attribute PatternChange.VB_Description = "Occurs when the file listing pattern, such as *.*, is changed using FileName or Pattern in code."
Event Scroll() 'MappingInfo=lstMain,lstMain,-1,Scroll
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Event MouseLeave()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If m_Enabled = True Then
        lstMain.BackColor = m_BackColor
        lstMain.ForeColor = m_ForeColor
        lstMain.Enabled = True
    Else
        lstMain.Enabled = False
    
        If g_Appearance <> Win98 Then
            lstMain.ForeColor = &H92A1A1
            lstMain.BackColor = &H8000000F
        Else
            lstMain.ForeColor = &H808080
            lstMain.BackColor = m_BackColor
        End If
    
    End If

End Sub

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

' Updates the control.
Public Sub Refresh()

    lstMain.Top = 0
    lstMain.Left = 0
    Right.Width = 3 * Screen.TwipsPerPixelX
    Bottom.Height = 3 * Screen.TwipsPerPixelY
    Left.Width = 3 * Screen.TwipsPerPixelX
    Top.Height = 3 * Screen.TwipsPerPixelY
    
    pPaintBorder
    pSetEnabled

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
    
    m_BackColor = &HFFFFFF
    m_Enabled = True
    m_ForeColor = 0
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    m_Path = ""

    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY

    Refresh

End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    lstMain.Archive = PropBag.ReadProperty("Archive", True)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    Set lstMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    lstMain.Hidden = PropBag.ReadProperty("Hidden", False)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    lstMain.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    lstMain.Normal = PropBag.ReadProperty("Normal", True)
    lstMain.Pattern = PropBag.ReadProperty("Pattern", "*.*")
    lstMain.ReadOnly = PropBag.ReadProperty("ReadOnly", True)
    lstMain.System = PropBag.ReadProperty("System", False)
    lstMain.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    lstMain.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_Path = PropBag.ReadProperty("Path", "")
    lstMain.Path = m_Path

    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

    If g_ControlsRefreshed = True Then
        Refresh
    End If

End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Archive", lstMain.Archive, True)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Font", lstMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("Hidden", lstMain.Hidden, False)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", lstMain.MousePointer, 0)
    Call PropBag.WriteProperty("Normal", lstMain.Normal, True)
    Call PropBag.WriteProperty("Pattern", lstMain.Pattern, "*.*")
    Call PropBag.WriteProperty("ReadOnly", lstMain.ReadOnly, True)
    Call PropBag.WriteProperty("System", lstMain.System, False)
    Call PropBag.WriteProperty("OLEDragMode", lstMain.OLEDragMode, 0)
    Call PropBag.WriteProperty("OLEDropMode", lstMain.OLEDropMode, 0)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("Path", m_Path, "")

End Sub

Private Sub lstMain_Click()
    RaiseEvent Click
End Sub

Private Sub lstMain_DblClick()
    RaiseEvent DblClick
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

Private Sub lstMain_PathChange()
    RaiseEvent PathChange
End Sub

Private Sub lstMain_PatternChange()
    RaiseEvent PatternChange
End Sub

Private Sub lstMain_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    lstMain.Width = Width
    lstMain.Height = Height
    Height = lstMain.Height
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

Private Sub lstMain_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub lstMain_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
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
'MappingInfo=lstMain,lstMain,-1,OLEDrag
Public Sub OLEDrag()
    lstMain.OLEDrag
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
    hwnd = lstMain.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Archive
Public Property Get Archive() As Boolean
Attribute Archive.VB_Description = "Determines whether a FileListBox control displays files with Archive attributes."
    Archive = lstMain.Archive
End Property

Public Property Let Archive(ByVal New_Archive As Boolean)
    lstMain.Archive() = New_Archive
    PropertyChanged "Archive"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    lstMain.BackColor = m_BackColor
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
    pSetEnabled
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
    If Enabled = True Then
        Set lstMain.Font = New_Font
    End If
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    lstMain.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Hidden
Public Property Get Hidden() As Boolean
Attribute Hidden.VB_Description = "Determines whether a FileListBox control displays files with Hidden attributes."
    Hidden = lstMain.Hidden
End Property

Public Property Let Hidden(ByVal New_Hidden As Boolean)
    lstMain.Hidden() = New_Hidden
    PropertyChanged "Hidden"
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
'MappingInfo=lstMain,lstMain,-1,MultiSelect
Public Property Get MultiSelect() As Integer
Attribute MultiSelect.VB_Description = "Returns/sets a value that determines whether a user can make multiple selections in a control."
    MultiSelect = lstMain.MultiSelect
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Normal
Public Property Get Normal() As Boolean
Attribute Normal.VB_Description = "Determines whether a FileListBox control displays files with Normal attributes."
    Normal = lstMain.Normal
End Property

Public Property Let Normal(ByVal New_Normal As Boolean)
    lstMain.Normal() = New_Normal
    PropertyChanged "Normal"
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
'MappingInfo=lstMain,lstMain,-1,Pattern
Public Property Get Pattern() As String
Attribute Pattern.VB_Description = "Returns/sets a value indicating the filenames displayed in a control at run time."
    Pattern = lstMain.Pattern
End Property

Public Property Let Pattern(ByVal New_Pattern As String)
    lstMain.Pattern() = New_Pattern
    PropertyChanged "Pattern"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,ReadOnly
Public Property Get ReadOnly() As Boolean
Attribute ReadOnly.VB_Description = "Returns/sets a value that determines whether files with read-only attributes are displayed in the file list or not."
    ReadOnly = lstMain.ReadOnly
End Property

Public Property Let ReadOnly(ByVal New_ReadOnly As Boolean)
    lstMain.ReadOnly() = New_ReadOnly
    PropertyChanged "ReadOnly"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,System
Public Property Get System() As Boolean
Attribute System.VB_Description = "Determines whether a FileListBox control displays files with System attributes."
    System = lstMain.System
End Property

Public Property Let System(ByVal New_System As Boolean)
    lstMain.System() = New_System
    PropertyChanged "System"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
Attribute ToolTipCaption.VB_Description = "Returns/sets the text displayed in the tooltip.\r\n"
    ToolTipCaption = m_ToolTipCaption
End Property

Public Property Let ToolTipCaption(ByVal New_ToolTipCaption As String)
    m_ToolTipCaption = New_ToolTipCaption
    PropertyChanged "ToolTipCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipIcon() As ToolTipIconEnum
Attribute ToolTipIcon.VB_Description = "Returns/sets the icon displayed in the tooltip.\r\n"
    ToolTipIcon = m_ToolTipIcon
End Property

Public Property Let ToolTipIcon(ByVal New_ToolTipIcon As ToolTipIconEnum)
    m_ToolTipIcon = New_ToolTipIcon
    PropertyChanged "ToolTipIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipStyle() As ToolTipStyleEnum
Attribute ToolTipStyle.VB_Description = "Returns/sets the style of the tooltip i.e Standad or Balloon."
    ToolTipStyle = m_ToolTipStyle
End Property

Public Property Let ToolTipStyle(ByVal New_ToolTipStyle As ToolTipStyleEnum)
    m_ToolTipStyle = New_ToolTipStyle
    PropertyChanged "ToolTipStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipTitle() As String
Attribute ToolTipTitle.VB_Description = "Returns/sets the title displayed in the tooltip."
    ToolTipTitle = m_ToolTipTitle
End Property

Public Property Let ToolTipTitle(ByVal New_ToolTipTitle As String)
    m_ToolTipTitle = New_ToolTipTitle
    PropertyChanged "ToolTipTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(ByVal New_Path As String)
    m_Path = New_Path
    On Error GoTo Error_Handler
    lstMain.Path = m_Path
    PropertyChanged "Path"
Error_Handler:
End Property
