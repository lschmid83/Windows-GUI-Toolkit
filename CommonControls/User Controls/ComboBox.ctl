VERSION 5.00
Begin VB.UserControl ComboBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   600
   ScaleWidth      =   2400
   ToolboxBitmap   =   "ComboBox.ctx":0000
   Begin CommonControls.MaskBox3 imgButton 
      Height          =   330
      Left            =   1470
      Top             =   0
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   582
      ScaleHeight     =   330
      ScaleWidth      =   270
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Bottom 
      Height          =   45
      Left            =   0
      Top             =   285
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   1455
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Top 
      Height          =   45
      Left            =   45
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   1455
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Left 
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   503
      ScaleHeight     =   285
      ScaleWidth      =   45
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin VB.ComboBox cmbMain 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   1740
   End
End
Attribute VB_Name = "ComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - ComboBox.ctl                                  '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_ToolTipCaption As String
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipTitle As String
Dim m_Enabled As Boolean

' Events
Event Change() 'MappingInfo=cmbMain,cmbMain,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event Click() 'MappingInfo=cmbMain,cmbMain,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=cmbMain,cmbMain,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event DropDown() 'MappingInfo=cmbMain,cmbMain,-1,DropDown
Attribute DropDown.VB_Description = "Occurs when the list portion of a ComboBox control is about to drop down."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event Scroll() 'MappingInfo=cmbMain,cmbMain,-1,Scroll
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."
Event OLECompleteDrag(Effect As Long) 'MappingInfo=cmbMain,cmbMain,-1,OLECompleteDrag
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=cmbMain,cmbMain,-1,OLEDragDrop
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=cmbMain,cmbMain,-1,OLEDragOver
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=cmbMain,cmbMain,-1,OLEGiveFeedback
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=cmbMain,cmbMain,-1,OLESetData
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=cmbMain,cmbMain,-1,OLEStartDrag

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If m_Enabled = True Then
        
        UserControl.Enabled = True
        pPaintDropdown "UnPressed"
        cmbMain.BackColor = m_BackColor
        cmbMain.ForeColor = m_ForeColor
        
        If Locked = True Then
            UserControl.Enabled = False
        End If
        
    Else
    
        UserControl.Enabled = False
        pPaintDropdown "Disabled"
        cmbMain.ForeColor = &H92A1A1
        cmbMain.BackColor = &H8000000F
  
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
    
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopLeft", crBitmap), 0, 0
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    Left.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Left", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Left.Height
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomLeft", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY

End Sub

' Paints the dropdown arrow with mouse over graphics.
Private Sub pPaintDropdown(sState As String)

    imgButton.Cls

    ' Draw the dropdown button
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Back", crBitmap), 5 * Screen.TwipsPerPixelX, 5 * Screen.TwipsPerPixelY, imgButton.Width, imgButton.Height
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Top", crBitmap), 0, 0
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Left", crBitmap), 0, 5 * Screen.TwipsPerPixelY, 5 * Screen.TwipsPerPixelX, imgButton.Height - (10 * Screen.TwipsPerPixelY)
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Right", crBitmap), imgButton.Width - (4 * Screen.TwipsPerPixelX), 5 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelX, imgButton.Height - (10 * Screen.TwipsPerPixelY)
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Bottom", crBitmap), 0, imgButton.Height - (5 * Screen.TwipsPerPixelY)
    
    ' Draw the arrow
    If g_Appearance <> Win98 Then
        imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Arrow", crBitmap), 6 * Screen.TwipsPerPixelX, (imgButton.Height / 2) - (3 * Screen.TwipsPerPixelY)
    Else
    
        If sState <> "Pressed" Then
            imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Arrow", crBitmap), 6 * Screen.TwipsPerPixelX, (imgButton.Height / 2) - (2 * Screen.TwipsPerPixelY)
        Else
            imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Arrow", crBitmap), 7 * Screen.TwipsPerPixelX, (imgButton.Height / 2) - (1 * Screen.TwipsPerPixelY)
        End If
    
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    cmbMain.Top = 0
    cmbMain.Left = 0
    
    Top.Top = 0
    Top.Left = 0
    Top.Height = 3 * Screen.TwipsPerPixelY
    Left.Top = 3 * Screen.TwipsPerPixelY
    Left.Left = 0
    Left.Width = 3 * Screen.TwipsPerPixelX
    
    Bottom.Left = 0
    Bottom.Height = 3 * Screen.TwipsPerPixelY
    imgButton.Top = 0
    imgButton.Width = 20 * Screen.TwipsPerPixelX
    
    pSetEnabled
    pPaintBorder
    UserControl_Resize

End Sub

' Adds an item to the control.
Public Sub AddItem(Text As String, Optional Index As Long)
    cmbMain.AddItem Text, Index
End Sub

' Removes an item from the control.
Public Sub RemoveItem(Index As Long)
    cmbMain.RemoveItem (Index)
End Sub

' Clears the control.
Public Sub Clear()
    cmbMain.Clear
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme

    m_BackColor = &HFFFFFF
    m_ForeColor = 0
    m_ToolTipStyle = Standard
    m_ToolTipIcon = NoIcon
    m_ToolTipTitle = ""
    m_ToolTipCaption = ""
    m_Enabled = True

    Width = 110 * Screen.TwipsPerPixelX

    Refresh
    
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    Set cmbMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    cmbMain.Locked = PropBag.ReadProperty("Locked", False)
    If Ambient.UserMode = True Then
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    MousePointer = PropBag.ReadProperty("MousePointer", 0)
    End If
    cmbMain.Text = PropBag.ReadProperty("Text", "Combo1")
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    cmbMain.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
    cmbMain.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)

    If g_ControlsRefreshed = True Then
        Refresh
    End If
   
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Font", cmbMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("Locked", cmbMain.Locked, False)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", cmbMain.MousePointer, 0)
    Call PropBag.WriteProperty("Text", cmbMain.Text, "Combo1")
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("OLEDragMode", cmbMain.OLEDragMode, 0)
    Call PropBag.WriteProperty("OLEDropMode", cmbMain.OLEDropMode, 0)
  
End Sub

Private Sub UserControl_Resize()

On Error GoTo err
    
    Top.Width = Width
    Left.Height = Height
    Bottom.Top = Height - (3 * Screen.TwipsPerPixelY)
    Bottom.Width = Width
    imgButton.Left = Width - imgButton.Width
    imgButton.Height = cmbMain.Height
    
    cmbMain.Width = Width
    Height = cmbMain.Height
    
    pPaintBorder
   
err:
    Exit Sub

End Sub

Private Sub imgButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Show the dropdown selection
    pPaintDropdown "Pressed"
    SendMessageLong cmbMain.hwnd, CB_SHOWDROPDOWN, 1, 1
    pPaintDropdown "UnPressed"

End Sub

Private Sub imgButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    With imgButton
    
        If GetCapture() = .hwnd Then
         
            ' Mouse has left control
            If x < 0 Or x > .Width Or y < 0 Or y > .Height Then
            
                Call ReleaseCapture
                pPaintDropdown "UnPressed"
        
            End If
    
        Else
        
            ' Mouse has entered control
            Call SetCapture(.hwnd)
            pPaintDropdown "MouseOver"
                     
        End If
    
    End With


End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties
Property Let ListIndex(Index As Integer)
    cmbMain.ListIndex = Index
End Property

Property Get ListIndex() As Integer
    ListIndex = cmbMain.ListIndex
End Property

Property Let FontBold(Bold As Boolean)
    cmbMain.FontBold = Bold
End Property

Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = cmbMain.FontBold
End Property

Property Let FontItalic(Italic As Boolean)
    cmbMain.FontItalic = Italic
End Property

Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = cmbMain.FontItalic
End Property

Property Let FontName(Name As String)
    cmbMain.FontName = Name
End Property

Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = cmbMain.FontName
End Property

Property Let FontSize(Size As Long)
    cmbMain.FontSize = Size
End Property

Property Get FontSize() As Long
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = cmbMain.FontSize
End Property

Property Let FontStrikeThru(StrikeThru As Boolean)
    cmbMain.FontStrikeThru = StrikeThru
End Property

Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_MemberFlags = "400"
    FontStrikeThru = cmbMain.FontStrikeThru
End Property

Property Let FontUnderline(UnderLine As Boolean)
    cmbMain.FontUnderline = UnderLine
End Property

Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = cmbMain.FontUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = cmbMain.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,OLEDrag
Public Sub OLEDrag()
    cmbMain.OLEDrag
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    cmbMain.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    pSetEnabled
    pPaintBorder
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cmbMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cmbMain.Font = New_Font
    Refresh
    pPaintDropdown "UnPressed"
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
    cmbMain.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,IntegralHeight
Public Property Get IntegralHeight() As Boolean
Attribute IntegralHeight.VB_Description = "Returns/Sets a value indicating whether the control displays partial items."
    IntegralHeight = cmbMain.IntegralHeight
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = cmbMain.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    cmbMain.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = cmbMain.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set cmbMain.MouseIcon = New_MouseIcon
    Set imgButton.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = cmbMain.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    cmbMain.MousePointer() = New_MousePointer
    imgButton.MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,OLEDragMode
Public Property Get OLEDragMode() As OLEDragConstants
Attribute OLEDragMode.VB_Description = "Returns/sets whether this object can act as an OLE drag/drop source, and whether this process is started automatically or under programmatic control.\r\n"
    OLEDragMode = cmbMain.OLEDragMode
End Property

Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
    cmbMain.OLEDragMode() = New_OLEDragMode
    PropertyChanged "OLEDragMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
    OLEDropMode = cmbMain.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    cmbMain.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = cmbMain.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    cmbMain.Text() = New_Text
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
    PropertyChanged "ToolTipCaption"
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
'MemberInfo=14,0,0,0
Public Property Get ToolTipIcon() As ToolTipIconEnum
Attribute ToolTipIcon.VB_Description = "Returns/sets the icon displayed in the tooltip."
    ToolTipIcon = m_ToolTipIcon
End Property

Public Property Let ToolTipIcon(ByVal New_ToolTipIcon As ToolTipIconEnum)
    m_ToolTipIcon = New_ToolTipIcon
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
    PropertyChanged "ToolTipTitle"
End Property

Private Sub cmbMain_Change()
    RaiseEvent Change
End Sub

Private Sub cmbMain_Click()
    RaiseEvent Click
End Sub

Private Sub cmbMain_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub cmbMain_DropDown()
    RaiseEvent DropDown
End Sub

Private Sub cmbMain_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cmbMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cmbMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub cmbMain_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub cmbMain_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub cmbMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub cmbMain_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub cmbMain_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub cmbMain_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub cmbMain_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub
