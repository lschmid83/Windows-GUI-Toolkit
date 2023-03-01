VERSION 5.00
Object = "{72D18DD4-0DA7-11D2-8E21-00B404C10000}#2.3#0"; "VBALODCL.OCX"
Begin VB.UserControl CheckedList 
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3555
   ScaleWidth      =   4800
   ToolboxBitmap   =   "CheckedListBox.ctx":0000
   Begin XPGUIControls10.MaskBox3 Left 
      Align           =   3  'Align Left
      Height          =   3285
      Left            =   0
      Top             =   135
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   5794
      ScaleHeight     =   3285
      ScaleWidth      =   135
      ScaleMode       =   1
   End
   Begin XPGUIControls10.MaskBox3 Top 
      Align           =   1  'Align Top
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   238
      ScaleHeight     =   135
      ScaleWidth      =   4800
      ScaleMode       =   1
   End
   Begin XPGUIControls10.MaskBox3 Bottom 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      Top             =   3420
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   238
      ScaleHeight     =   135
      ScaleWidth      =   4800
      ScaleMode       =   1
   End
   Begin XPGUIControls10.MaskBox3 Right 
      Align           =   4  'Align Right
      Height          =   3285
      Left            =   4665
      Top             =   135
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   5794
      ScaleHeight     =   3285
      ScaleWidth      =   135
      ScaleMode       =   1
   End
   Begin ODCboLst.OwnerDrawComboList lstMain 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4180
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      Style           =   7
      MaxLength       =   0
   End
End
Attribute VB_Name = "CheckedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Citex Software                                                       '
'XP GUI Toolkit v1.0 - CheckedList Component v1.0                        '
'Copyright 2001 Lawrence Schmid                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Dim strEnabled As String

'Property Variables:
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As Variant
Dim m_ToolTipStyle As Variant
Dim m_ToolTipTitle As String
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_Locked As Boolean
Dim m_DefaultStyle As ComboBoxDefaultStyleEnum
Dim m_DrawMode As ComboBoxDrawModeEnum


'Event Declarations:
Event MouseLeave()
Event Change()
Event Click() 'MappingInfo=lstMain,lstMain,-1,Click
Event DblClick() 'MappingInfo=lstMain,lstMain,-1,DblClick
Event DrawItem(Index As Long, hDC As Long, bSelected As Boolean, bEnabled As Boolean, LeftPixels As Long, TopPixels As Long, RightPixels As Long, BottomPixels As Long, hFntOld As Long) 'MappingInfo=lstMain,lstMain,-1,DrawItem
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=lstMain,lstMain,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=lstMain,lstMain,-1,MouseUp
Event SelCancel() 'MappingInfo=lstMain,lstMain,-1,SelCancel





'------------------------------------------------------------------------------------------
'Public control subs

Public Sub ChangeTheme()

lstMain.Top = 1 * Screen.TwipsPerPixelY
lstMain.Left = 1 * Screen.TwipsPerPixelX
Right.Width = 3 * Screen.TwipsPerPixelX
Bottom.Height = 3 * Screen.TwipsPerPixelY
Left.Width = 3 * Screen.TwipsPerPixelX
Top.Height = 3 * Screen.TwipsPerPixelY

pSetDefaultStyle
pSetDrawStyle
pSetEnabled
pSetBorder

UserControl_Resize

End Sub

Public Sub RefreshTheme()

lstMain.Top = 1 * Screen.TwipsPerPixelY
lstMain.Left = 1 * Screen.TwipsPerPixelX
Right.Width = 3 * Screen.TwipsPerPixelX
Bottom.Height = 3 * Screen.TwipsPerPixelY
Left.Width = 3 * Screen.TwipsPerPixelX
Top.Height = 3 * Screen.TwipsPerPixelY

pSetDefaultStyle
pSetDrawStyle
pSetEnabled
pSetBorder

UserControl_Resize

End Sub

Public Sub AddItem(Text As String)
lstMain.AddItem Text

End Sub

Public Sub AddItemAndData( _
        ByVal Text As String, _
        Optional ByVal IconIndex As Long = -1, _
        Optional ByVal Indent As Long = 0, _
        Optional ByVal ForeColour As OLE_COLOR = -1, _
        Optional ByVal BackColour As OLE_COLOR = -1, _
        Optional ByVal ItemData As Long = 0, _
        Optional ByVal ExtraData As Long = 0, _
        Optional ByVal Height As Long = -1, _
        Optional ByVal TextXAlign As EODCLItemXAlign = eixLeft, _
        Optional ByVal TextYAlign As EODCLItemYAlign = eixTop, _
        Optional ByRef Font As StdFont = Nothing)

lstMain.AddItemAndData Text, IconIndex, Indent, ForeColor, BackColor, ItemData, ExtraData, Height, TextXAlign, TextYAlign, Font


End Sub


Public Sub InsertItem(Text As String, Index As Long)

lstMain.InsertItem Text, Index

End Sub

Public Sub InsertItemAndData( _
        ByVal Text As String, _
        Index As Long, _
        Optional ByVal IconIndex As Long = -1, _
        Optional ByVal Indent As Long = 0, _
        Optional ByVal ForeColour As OLE_COLOR = -1, _
        Optional ByVal BackColour As OLE_COLOR = -1, _
        Optional ByVal ItemData As Long = 0, _
        Optional ByVal ExtraData As Long = 0, _
        Optional ByVal Height As Long = -1, _
        Optional ByVal TextXAlign As EODCLItemXAlign = eixLeft, _
        Optional ByVal TextYAlign As EODCLItemYAlign = eixTop, _
        Optional ByRef Font As StdFont = Nothing)

lstMain.InsertItemAndData Text, Index, IconIndex, Indent, ForeColor, BackColor, ItemData, ExtraData, Height, TextXAlign, TextYAlign, Font

End Sub

Public Sub Clear()

lstMain.Clear

End Sub

Public Function ListCount() As Long

ListCount = lstMain.ListCount

End Function

Property Get List(ByVal Index As Long) As String
List = lstMain.List(Index)

End Property

Property Let List(ByVal Index As Long, ByVal Text As String)
lstMain.List(Index) = Text

End Property

Public Sub ListIndex(ByVal Index As Long)

lstMain.ListIndex = Index

End Sub

Public Sub RemoveItem(ByVal Index As Long)

lstMain.RemoveItem (Index)

End Sub

Public Function FindItemIndex(TextToFind As String, ExactMatch As Boolean) As Long

FindItemIndex = lstMain.FindItemIndex(TextToFind, ExactMatch)

End Function

Public Sub ShowDropDown(Visible As Boolean)

lstMain.ShowDropDown Visible

End Sub

Public Sub ShowDropDownAtPosition(X As Long, Y As Long, Optional Width As Long = 0, Optional Height As Long = 0)

lstMain.ShowDropDownAtPosition X, Y, Width, Height

End Sub

Public Function IsComboDropped()

IsComboDropped = lstMain.ComboIsDropped

End Function

Property Let ImageList(ImgList As Variant)

lstMain.ImageList = ImgList

End Property


Property Get SelLength() As Long

If Ambient.UserMode = True Then
SelLength = lstMain.SelLength
End If

End Property

Property Get SelStart() As Long

If Ambient.UserMode = True Then
SelStart = lstMain.SelStart
End If

End Property

Property Get SelText() As String

SelText = lstMain.SelText

End Property


Property Let ItemBackColor(Index As Long, BackColor As OLE_COLOR)
lstMain.ItemBackColor(Index) = BackColor
End Property

Property Get ItemBackColor(Index As Long) As OLE_COLOR
BackColor = lstMain.ItemBackColor(Index)
End Property


Property Let ItemData(Index As Long, ItemData As Long)
lstMain.ItemData(Index) = ItemData
End Property

Property Get ItemData(Index As Long) As Long
ItemData = lstMain.ItemData(Index)
End Property


Property Let ItemExtraData(Index As Long, ItemExtraData As Long)
lstMain.ItemExtraData(Index) = ItemExtraData
End Property

Property Get ItemExtraData(Index As Long) As Long
ItemExtraData = lstMain.ItemExtraData(Index)
End Property


Property Let ItemFont(Index As Long, ItemFont As StdFont)
lstMain.ItemFont(Index) = ItemFont
End Property

Property Get ItemFont(Index As Long) As StdFont
ItemFont = lstMain.ItemFont(Index)
End Property


Property Let ItemForeColor(Index As Long, ItemForeColor As OLE_COLOR)
lstMain.ItemForeColor(Index) = ItemForeColor
End Property

Property Get ItemForeColor(Index As Long) As OLE_COLOR
ItemForeColor = lstMain.ItemForeColor(Index)

End Property


Property Let ItemHeight(Index As Long, ItemHeight As Long)
lstMain.ItemHeight(Index) = ItemHeight
End Property

Property Get ItemHeight(Index As Long) As Long
ItemHeight = lstMain.ItemHeight(Index)

End Property


Property Let ItemIcon(Index As Long, ItemIcon As Long)
lstMain.ItemIcon(Index) = ItemIcon
End Property

Property Get ItemIcon(Index As Long) As Long
ItemIcon = lstMain.ItemIcon(Index)
End Property

Property Let ItemIndent(Index As Long, Indent As Long)
lstMain.ItemIndent(Index) = Indent
End Property

Property Get ItemIndent(Index As Long) As Long
ItemIndent = lstMain.ItemIndent(Index)

End Property


Property Let ItemOverLine(Index As Long, OverLine As Boolean)
lstMain.ItemOverLine(Index) = OverLine
End Property

Property Get ItemOverLine(Index As Long) As Boolean
ItemOverLine = lstMain.ItemOverLine(Index)

End Property


Property Let ItemUnderLine(Index As Long, UnderLine As Boolean)
lstMain.ItemUnderLine(Index) = UnderLine
End Property

Property Get ItemUnderLine(Index As Long) As Boolean
ItemUnderLine = lstMain.ItemUnderLine(Index)

End Property

Property Let ItemXAlign(Index As Long, Alignment As ItemXAlignEnum)
lstMain.ItemXAlign(Index) = Alignment
End Property

Property Get ItemXAlign(Index As Long) As ItemXAlignEnum
ItemXAlign = lstMain.ItemXAlign(Index)

End Property

Property Let ItemYAlign(Index As Long, Alignment As ItemYAlignEnum)
lstMain.ItemYAlign(Index) = Alignment
End Property

Property Get ItemYAlign(Index As Long) As ItemYAlignEnum
ItemYAlign = lstMain.ItemYAlign(Index)

End Property



'-----------------------------------------------------------------------------------------
'Private internal subs


Private Sub pSetDrawStyle()

If m_DefaultStyle = NoStyle Then 'Only allow DrawStyles to be set if a default style is not set.
                                 'This is because the if a default style is set it automatically sets the
                                 'draw style of the combobox
    Select Case m_DrawMode
    
    Case Is = dmStandard
    lstMain.ClientDraw = ecdNoClientDraw
    
    Case Is = dmColorPickerNoNames
    lstMain.ClientDraw = ecdColourPickerNoNames
    
    Case Is = dmColorPicker
    lstMain.ClientDraw = ecdColourPickerWithNames
    
    Case Is = dmFontPicker
    lstMain.ClientDraw = ecdFontPicker
    
    Case Is = dmParagraphStyles
    lstMain.ClientDraw = ecdParagraphStyles
    
    End Select
    
End If

End Sub

Private Sub pSetDefaultStyle()

Select Case m_DefaultStyle

Case Is = ColorPicker
lstMain.Clear
Dim i As Long
For i = 1 To lstMain.ListCount
lstMain.RemoveItem (i)
Next
lstMain.ClientDraw = ecdSysColourPicker
m_DrawMode = dmColorPicker
LoadSysColorList lstMain
lstMain.Font.Bold = False
lstMain.Font.Italic = False
lstMain.Font.Name = "MS Sans Serif"
lstMain.Font.Size = 8
lstMain.ClientDraw = ecdSysColourPicker
lstMain.Font = lstMain.Font


Case Is = ColorPickerNoNames
lstMain.Clear
lstMain.ClientDraw = ecdColourPickerNoNames
m_DrawMode = dmColorPickerNoNames
lstMain.Font.Bold = False
lstMain.Font.Italic = False
lstMain.Font.Name = "MS Sans Serif"
lstMain.Font.Size = 8
LoadSysColorList lstMain


Case Is = FontPicker
lstMain.Clear
lstMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
lstMain.Font.Bold = False
lstMain.Font.Italic = False
lstMain.Font.Name = "MS Sans Serif"
lstMain.Font.Size = 8
LoadFontList lstMain, "", -1, -1

Case Is = FontViewer
lstMain.Clear
lstMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
lstMain.Font.Bold = False
lstMain.Font.Italic = False
lstMain.Font.Name = "MS Sans Serif"
lstMain.Font.Size = 8
LoadFontListViewer lstMain, "", -1, -1, True

Case Is = FontViewerNoIcons
lstMain.Clear
lstMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
lstMain.Font.Bold = False
lstMain.Font.Italic = False
lstMain.Font.Name = "MS Sans Serif"
lstMain.Font.Size = 8
LoadFontListViewer lstMain, "", -1, -1, False

Case Is = ParagraphStyles
lstMain.Clear
lstMain.ClientDraw = ecdParagraphStyles
lstMain.Font.Bold = False
lstMain.Font.Italic = False
lstMain.Font.Name = "MS Sans Serif"
lstMain.Font.Size = 8
m_DrawMode = dmParagraphStyles
LoadParagraphStyles lstMain


End Select

End Sub


Private Sub pSetEnabled()

If Ambient.UserMode = False Then
UserControl.Enabled = False

    If m_Enabled = True Then
    lstMain.Enabled = True
    
    Else
    lstMain.Enabled = False
    
    End If


Else

    If m_Enabled = True Then
    UserControl.Enabled = True
    lstMain.Enabled = True
    
        If Locked = True Then
        UserControl.Enabled = False
        End If
    
    Else
    UserControl.Enabled = False
    lstMain.Enabled = False
    
    End If

End If



End Sub


Private Sub pSetBorder()

If m_Enabled = True Then
strEnabled = "Enabled\"
Else
strEnabled = "Disabled\"

End If


Top.Cls
Left.Cls
Bottom.Cls
Right.Cls

Top.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopLeft", crBitmap), 0, 0
Top.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
Top.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY

Left.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "Left", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Left.Height

Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomLeft", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width, 3 * Screen.TwipsPerPixelY
Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY

Right.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "Right", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Height


End Sub





'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Columns
Public Property Get Columns() As Integer
    Columns = lstMain.Columns
End Property

Public Property Let Columns(ByVal New_Columns As Integer)
    lstMain.Columns() = New_Columns
    PropertyChanged "Columns"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
m_BackColor = New_BackColor

If Enabled = True Then
lstMain.BackColor = New_BackColor
End If


PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
m_ForeColor = New_ForeColor

If Enabled = True Then
lstMain.ForeColor = New_ForeColor
End If

PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
m_Enabled = New_Enabled

pSetEnabled
UserControl_Resize

PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Font
Public Property Get Font() As Font
    Set Font = lstMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lstMain.Font = New_Font
    PropertyChanged "Font"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,Sorted
Public Property Get Sorted() As Boolean
    Sorted = lstMain.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    lstMain.Sorted() = New_Sorted
    PropertyChanged "Sorted"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
    FullRowSelect = lstMain.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    lstMain.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    PropertyChanged "Locked"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DefaultStyle() As ComboBoxDefaultStyleEnum
    DefaultStyle = m_DefaultStyle
End Property

Public Property Let DefaultStyle(ByVal New_DefaultStyle As ComboBoxDefaultStyleEnum)
m_DefaultStyle = New_DefaultStyle

pSetDefaultStyle

If m_DefaultStyle = NoStyle Then
m_DrawMode = dmStandard
lstMain.ClientDraw = ecdClientDrawOnly
End If

PropertyChanged "DefaultStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DrawMode() As ComboBoxDrawModeEnum
    DrawMode = m_DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As ComboBoxDrawModeEnum)
    m_DrawMode = New_DrawMode
    PropertyChanged "DrawMode"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
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
ToolTipTitle = m_ToolTipTitle
End Property

Public Property Let ToolTipTitle(ByVal New_ToolTipTitle As String)
m_ToolTipTitle = New_ToolTipTitle
pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

PropertyChanged "ToolTipTitle"
End Property



Private Sub Bottom_Paint()
pSetBorder
End Sub



'-----------------------------------------------------------------------------------------------------
'Read\Write\Init properties


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
    
    m_BackColor = &HFFFFFF
    m_ForeColor = 0
    m_Enabled = True
    m_Locked = False
    m_DefaultStyle = NoStyle
    m_DrawMode = dmStandard
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""

    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY

    ChangeTheme


End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)


    SetDefaultTheme

    lstMain.Columns = PropBag.ReadProperty("Columns", 1)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    Set lstMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lstMain.Sorted = PropBag.ReadProperty("Sorted", False)
    lstMain.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    m_Locked = PropBag.ReadProperty("Locked", False)
    m_DefaultStyle = PropBag.ReadProperty("DefaultStyle", NoStyle)
    m_DrawMode = PropBag.ReadProperty("DrawMode", dmStandard)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
  
    If m_Enabled = True Then
    lstMain.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    lstMain.ForeColor = PropBag.ReadProperty("ForeColor", 0)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    End If

    If glbControlsRefreshed = True Then
    RefreshTheme 'Refresh control because tilebar refresh code has already run before this
                 'This is needed to make sure the correct appearance is displayed
    End If
    
    pSetToolTip lstMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Columns", lstMain.Columns, 1)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Font", lstMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("Sorted", lstMain.Sorted, False)
    Call PropBag.WriteProperty("FullRowSelect", lstMain.FullRowSelect, False)
    Call PropBag.WriteProperty("Locked", m_Locked, False)
    Call PropBag.WriteProperty("DefaultStyle", m_DefaultStyle, NoStyle)
    Call PropBag.WriteProperty("DrawMode", m_DrawMode, dmStandard)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")

End Sub

'-------------------------------------------------------------------------------------------
'Control events


Private Sub UserControl_Resize()

On Error GoTo err

lstMain.Width = Width - (2 * Screen.TwipsPerPixelX)
lstMain.Height = Height - (2 * Screen.TwipsPerPixelX)

err:
Exit Sub


End Sub


Private Sub lstMain_Click()
    RaiseEvent Click
End Sub

Private Sub lstMain_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lstMain_DrawItem(Index As Long, hDC As Long, bSelected As Boolean, bEnabled As Boolean, LeftPixels As Long, TopPixels As Long, RightPixels As Long, BottomPixels As Long, hFntOld As Long)
    RaiseEvent DrawItem(Index, hDC, bSelected, bEnabled, LeftPixels, TopPixels, RightPixels, BottomPixels, hFntOld)
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

Private Sub lstMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lstMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lstMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lstMain_SelCancel()
    RaiseEvent SelCancel
End Sub




