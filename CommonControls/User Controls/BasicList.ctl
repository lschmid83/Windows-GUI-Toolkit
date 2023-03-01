VERSION 5.00
Begin VB.UserControl ListBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "BasicList.ctx":0000
   Begin CommonControls.MaskBox3 Right 
      Align           =   4  'Align Right
      Height          =   3345
      Left            =   4665
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   5900
      ScaleHeight     =   3345
      ScaleWidth      =   135
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Bottom 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      Top             =   3465
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   238
      ScaleHeight     =   135
      ScaleWidth      =   4800
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Left 
      Align           =   3  'Align Left
      Height          =   3345
      Left            =   0
      Top             =   120
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   5900
      ScaleHeight     =   3345
      ScaleWidth      =   150
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Top 
      Align           =   1  'Align Top
      Height          =   120
      Left            =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   212
      ScaleHeight     =   120
      ScaleWidth      =   4800
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin VB.ListBox lstMain 
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3870
   End
End
Attribute VB_Name = "ListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - ListBox.ctl                                   '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR

' Events
Event Click() 'MappingInfo=lstMain,lstMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=lstMain,lstMain,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Scroll() 'MappingInfo=lstMain,lstMain,-1,Scroll
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=lstMain,lstMain,-1,OLECompleteDrag
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lstMain,lstMain,-1,OLEDragDrop
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=lstMain,lstMain,-1,OLEDragOver
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=lstMain,lstMain,-1,OLEGiveFeedback
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=lstMain,lstMain,-1,OLESetData
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=lstMain,lstMain,-1,OLEStartDrag

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If UserControl.Enabled = True Then
        lstMain.ForeColor = m_ForeColor
        lstMain.BackColor = m_BackColor
    Else
        If g_Appearance <> Win98 Then
        lstMain.ForeColor = &H92A1A1
        lstMain.BackColor = &H8000000F
        
        Else
        lstMain.ForeColor = &H808080
        lstMain.BackColor = &HFFFFFF
        
        End If
    End If
    
End Sub

' Paints the control border.
Private Sub pPaintBorder()
    
    ' Set the enabled graphic path
    Dim sEnabled As String
    If UserControl.Enabled = True Then
        sEnabled = "Enabled\"
    Else
        sEnabled = "Disabled\"
    End If
    
    ' Paint the border
    Top.Cls
    Left.Cls
    Bottom.Cls
    Right.Cls

    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "TopLeft", crBitmap), 0, 0
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    Left.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Left", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Left.Height
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "BottomLeft", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width, 3 * Screen.TwipsPerPixelY
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    Right.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Right", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Height

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    lstMain.Top = 1 * Screen.TwipsPerPixelY
    lstMain.Left = 1 * Screen.TwipsPerPixelX
    Right.Width = 3 * Screen.TwipsPerPixelX
    Bottom.Height = 3 * Screen.TwipsPerPixelY
    Left.Width = 3 * Screen.TwipsPerPixelX
    Top.Height = 3 * Screen.TwipsPerPixelY
    
    Call UserControl_Resize
    pSetEnabled

End Sub

' Adds an item to the control.
Public Sub AddItem(Text As String, Optional Index As Long)
    lstMain.AddItem Text, Index
End Sub

' Removes an item from the control.
Public Sub RemoveItem(Index As Integer)
    lstMain.RemoveItem (Index)
End Sub

' Clears the control.
Public Sub Clear()
    lstMain.Clear
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
        
    SetDefaultTheme
    
    m_BackColor = &HFFFFFF
    m_ForeColor = 0
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    
    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY

    Refresh

End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lstMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    lstMain.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    lstMain.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    lstMain.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")

    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    
    If g_ControlsRefreshed = True Then
        Refresh
    End If

       
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lstMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", lstMain.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDragMode", lstMain.OLEDragMode, 0)
    Call PropBag.WriteProperty("OLEDropMode", lstMain.OLEDropMode, 0)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")

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

Private Sub lstMain_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub lstMain_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,OLEDrag
Public Sub OLEDrag()
    lstMain.OLEDrag
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

Property Let FontBold(Bold As Boolean)
    lstMain.FontBold = Bold
End Property

Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = lstMain.FontBold
End Property

Property Let FontItalic(Italic As Boolean)
    lstMain.FontItalic = Italic
End Property

Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = lstMain.FontItalic
End Property

Property Let FontName(Name As String)
    lstMain.FontName = Name
End Property

Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = lstMain.FontName
End Property

Property Let FontSize(Size As Long)
    lstMain.FontSize = Size
End Property

Property Get FontSize() As Long
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = lstMain.FontSize
End Property

Property Let FontStrikeThru(StrikeThru As Boolean)
    lstMain.FontStrikeThru = StrikeThru
End Property

Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_MemberFlags = "400"
    FontStrikeThru = lstMain.FontStrikeThru
End Property

Property Let FontUnderline(UnderLine As Boolean)
    lstMain.FontUnderline = UnderLine
End Property

Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = lstMain.FontUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    lstMain.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
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
    Set lstMain.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    lstMain.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,IntegralHeight
Public Property Get IntegralHeight() As Boolean
Attribute IntegralHeight.VB_Description = "Returns/Sets a value indicating whether the control displays partial items."
    IntegralHeight = lstMain.IntegralHeight
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
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control.\r\n"
    OLEDropMode = lstMain.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    lstMain.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = lstMain.Sorted
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Style
Public Property Get Style() As Integer
Attribute Style.VB_Description = "Returns/sets a value that determines whether checkboxes are displayed inside a ListBox control."
    Style = lstMain.Style
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
Attribute ToolTipIcon.VB_Description = "Returns/sets the icon displayed in the tooltip.\r\n"
    ToolTipIcon = m_ToolTipIcon
End Property

Public Property Let ToolTipIcon(ByVal New_ToolTipIcon As ToolTipIconEnum)
    m_ToolTipIcon = New_ToolTipIcon
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
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
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
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
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = lstMain.hwnd
End Property
