VERSION 5.00
Begin VB.UserControl TextBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   ScaleHeight     =   945
   ScaleWidth      =   2265
   ToolboxBitmap   =   "TextBox.ctx":0000
   Begin VB.TextBox txtMain 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - TextBox.ctl                                   '
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
Dim m_MouseIcon As Picture
Dim m_MousePointer As Integer
Dim m_Border As Boolean

' Events
Event Change() 'MappingInfo=txtMain,txtMain,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Attribute Change.VB_MemberFlags = "200"
Event Click() 'MappingInfo=txtMain,txtMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event DblClick() 'MappingInfo=txtMain,txtMain,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtMain,txtMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtMain,txtMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtMain,txtMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtMain,txtMain,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtMain,txtMain,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtMain,txtMain,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=txtMain,txtMain,-1,OLECompleteDrag
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtMain,txtMain,-1,OLEDragDrop
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=txtMain,txtMain,-1,OLEDragOver
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=txtMain,txtMain,-1,OLEGiveFeedback
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=txtMain,txtMain,-1,OLESetData
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=txtMain,txtMain,-1,OLEStartDrag

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If UserControl.Enabled = True Then
        txtMain.BackColor = m_BackColor
        txtMain.ForeColor = m_ForeColor
    Else
        If g_Appearance <> Win98 Then
            txtMain.ForeColor = &H92A1A1
            txtMain.BackColor = &H8000000F
        Else
            txtMain.ForeColor = &H808080
            txtMain.BackColor = &HFFFFFF
        End If
    End If

End Sub

' Sets the position of the textbox based on the border property.
Private Sub pSetBorder()

    If m_Border = True Then
        txtMain.Left = 3 * Screen.TwipsPerPixelX
        txtMain.Top = 3 * Screen.TwipsPerPixelY
    Else
        txtMain.Left = 0
        txtMain.Top = 0
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

    ' Draw the border
    UserControl.Cls
    
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "TopLeft", crBitmap), 0, 0
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Left", crBitmap), 0, 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, Height - (3 * Screen.TwipsPerPixelY), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY


End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Update the control.
Public Sub Refresh()

    pSetBorder
    pSetEnabled
    UserControl_Resize

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
    m_Border = True
    Text = Ambient.DisplayName
    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY

    Refresh
   
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    txtMain.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    txtMain.Locked = PropBag.ReadProperty("Locked", False)
    txtMain.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtMain.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtMain.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    txtMain.Text = PropBag.ReadProperty("Text", "Text1")
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_Border = PropBag.ReadProperty("Border", True)
    txtMain.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    txtMain.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
   
    pSetToolTip txtMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
   
    If g_ControlsRefreshed = True Then
        Refresh
    End If
     
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", txtMain.Alignment, 0)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", txtMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("Locked", txtMain.Locked, False)
    Call PropBag.WriteProperty("MaxLength", txtMain.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtMain.MousePointer, 0)
    Call PropBag.WriteProperty("PasswordChar", txtMain.PasswordChar, "")
    Call PropBag.WriteProperty("Text", txtMain.Text, "Text1")
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("Border", m_Border, True)
    Call PropBag.WriteProperty("OLEDragMode", txtMain.OLEDragMode, 0)
    Call PropBag.WriteProperty("OLEDropMode", txtMain.OLEDropMode, 0)
  
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    If m_Border = True Then
        txtMain.Width = Width - (6 * Screen.TwipsPerPixelX)
        txtMain.Height = Height - (6 * Screen.TwipsPerPixelX)
        pPaintBorder
    Else
        txtMain.Width = Width
        txtMain.Height = Height
        Height = txtMain.Height
    End If

err:
    Exit Sub

End Sub

Private Sub txtMain_Change()
    RaiseEvent Change
End Sub

Private Sub txtMain_Click()
    RaiseEvent Click
End Sub

Private Sub txtMain_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub txtMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub txtMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub txtMain_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub txtMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub txtMain_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub txtMain_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub txtMain_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub txtMain_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,OLEDrag
Public Sub OLEDrag()
    txtMain.OLEDrag
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

Property Let FontBold(Bold As Boolean)
    txtMain.FontBold = Bold
End Property

Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = txtMain.FontBold
End Property

Property Let FontItalic(Italic As Boolean)
    txtMain.FontItalic = Italic
End Property

Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = txtMain.FontItalic
End Property

Property Let FontName(Name As String)
    txtMain.FontName = Name
End Property

Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = txtMain.FontName
End Property

Property Let FontSize(Size As Long)
    txtMain.FontSize = Size
End Property

Property Get FontSize() As Long
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = txtMain.FontSize
End Property

Property Let FontStrikeThru(StrikeThru As Boolean)
    txtMain.FontStrikeThru = StrikeThru
End Property

Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_MemberFlags = "400"
    FontStrikeThru = txtMain.FontStrikeThru
End Property

Property Let FontUnderline(UnderLine As Boolean)
    txtMain.FontUnderline = UnderLine
End Property

Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = txtMain.FontUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = txtMain.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtMain.Alignment() = New_Alignment
    PropertyChanged "Alignment"
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
    txtMain.BackColor() = New_BackColor
    UserControl_Resize
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Refresh
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtMain.Font = New_Font
    Refresh
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
    txtMain.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Specifies whether the selection in a Masked edit control is hidden when the control loses focus."
    HideSelection = txtMain.HideSelection
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = txtMain.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtMain.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtMain.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtMain.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = txtMain.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtMain.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = txtMain.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    txtMain.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = txtMain.MultiLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
    OLEDropMode = txtMain.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    txtMain.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,OLEDragMode
Public Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control."
    OLEDragMode = txtMain.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
    txtMain.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,PasswordChar
Public Property Get PasswordChar() As String
Attribute PasswordChar.VB_Description = "Returns/sets a value that determines whether characters typed by a user or placeholder characters are displayed in a control."
    PasswordChar = txtMain.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtMain.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,ScrollBars
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
    ScrollBars = txtMain.ScrollBars
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
    hwnd = txtMain.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = -517
    Text = txtMain.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtMain.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
Attribute ToolTipCaption.VB_Description = "Returns/sets the text displayed in the tooltip."
    ToolTipCaption = m_ToolTipCaption
End Property

Public Property Let ToolTipCaption(ByVal New_ToolTipCaption As String)
    m_ToolTipCaption = New_ToolTipCaption
    pSetToolTip txtMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
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
    pSetToolTip txtMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
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
    pSetToolTip txtMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
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
    pSetToolTip txtMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Border() As Boolean
Attribute Border.VB_Description = "Return/sets whether the control has a border."
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    pSetBorder
    UserControl_Resize
    PropertyChanged "Border"
End Property
