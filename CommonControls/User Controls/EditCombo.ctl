VERSION 5.00
Object = "{72D18DD4-0DA7-11D2-8E21-00B404C10000}#2.3#0"; "VBALODCL.OCX"
Begin VB.UserControl EditCombo 
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   HasDC           =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   2325
   ToolboxBitmap   =   "EditCombo.ctx":0000
   Begin XPGUIControls10.MaskBox3 imgButton 
      Height          =   330
      Left            =   1470
      Top             =   0
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   582
      ScaleHeight     =   330
      ScaleWidth      =   270
      ScaleMode       =   1
   End
   Begin XPGUIControls10.MaskBox3 Bottom 
      Height          =   45
      Left            =   0
      Top             =   285
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   1470
      ScaleMode       =   1
   End
   Begin XPGUIControls10.MaskBox3 Top 
      Height          =   45
      Left            =   0
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   1470
      ScaleMode       =   1
   End
   Begin XPGUIControls10.MaskBox3 Left 
      Height          =   240
      Left            =   0
      Top             =   45
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   423
      ScaleHeight     =   240
      ScaleWidth      =   45
      ScaleMode       =   1
   End
   Begin ODCboLst.OwnerDrawComboList cmbMain 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      Sorted          =   -1  'True
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
      BackColor       =   16777215
      Style           =   0
      FullRowSelect   =   -1  'True
   End
End
Attribute VB_Name = "EditCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Citex Software                                                       '
'XP GUI Toolkit v1.0 - EditCombo Component v1.0                        '
'Copyright 2001 Lawrence Schmid                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'General Variables:
Dim blnToolTipSet As Boolean
Dim strEnabled As String

'Property Variables:
Dim m_ToolTipCaption As String
Dim m_DrawMode As ComboBoxDrawModeEnum
Dim m_DefaultStyle As ComboBoxDefaultStyleEnum
Dim m_Enabled As Boolean
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_Locked As Boolean

'Event Declarations:
Event MouseLeave()
Public Event Click()
Public Event Change()
Public Event DblClick()
Public Event CloseUp()
Public Event DropDown()
Public Event SelCancel()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event AutoCompleteSelection(ByVal Text As String, ByVal Index As Long)

Private Sub Bottom_Paint()
pSetBorder
End Sub

Private Sub imgButton_Paint()

If m_Enabled = True Then
pSetGraphic "UnPressed"
Else
pSetGraphic "Disabled"

End If

End Sub



'-----------------------------------------------------------------------------------------------------
'Public subs/functions


Public Sub RefreshTheme()

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

UserControl_Resize

pSetDefaultStyle
pSetDrawStyle
pSetEnabled
pSetBorder

UserControl_Resize


End Sub

Public Sub ChangeTheme()

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

UserControl_Resize

pSetDefaultStyle
pSetDrawStyle
pSetEnabled
pSetBorder

UserControl_Resize


End Sub


Public Sub AddItem(Text As String)

cmbMain.AddItem Text

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

cmbMain.AddItemAndData Text, IconIndex, Indent, ForeColor, BackColor, ItemData, ExtraData, Height, TextXAlign, TextYAlign, Font


End Sub


Public Sub InsertItem(Text As String, Index As Long)

cmbMain.InsertItem Text, Index

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

cmbMain.InsertItemAndData Text, Index, IconIndex, Indent, ForeColor, BackColor, ItemData, ExtraData, Height, TextXAlign, TextYAlign, Font

End Sub

Public Sub Clear()

cmbMain.Clear

End Sub

Public Function ListCount() As Long

ListCount = cmbMain.ListCount

End Function

Property Get List(ByVal Index As Long) As String
List = cmbMain.List(Index)

End Property

Property Let List(ByVal Index As Long, ByVal Text As String)
cmbMain.List(Index) = Text

End Property

Public Sub ListIndex(ByVal Index As Long)

cmbMain.ListIndex = Index

End Sub

Public Sub RemoveItem(ByVal Index As Long)

cmbMain.RemoveItem (Index)

End Sub

Public Function FindItemIndex(TextToFind As String, ExactMatch As Boolean) As Long

FindItemIndex = cmbMain.FindItemIndex(TextToFind, ExactMatch)

End Function

Public Sub ShowDropDown(Visible As Boolean)

cmbMain.ShowDropDown Visible

End Sub

Public Sub ShowDropDownAtPosition(X As Long, Y As Long, Optional Width As Long = 0, Optional Height As Long = 0)

cmbMain.ShowDropDownAtPosition X, Y, Width, Height

End Sub

Public Function IsComboDropped()

IsComboDropped = cmbMain.ComboIsDropped

End Function




Property Let ImageList(ImgList As Variant)

cmbMain.ImageList = ImgList

End Property


Property Get SelLength() As Long

If Ambient.UserMode = True Then
SelLength = cmbMain.SelLength
End If

End Property

Property Let SelLength(Length As Long)

cmbMain.SelLength = Length

End Property

Property Get SelStart() As Long

If Ambient.UserMode = True Then
SelStart = cmbMain.SelStart
End If

End Property
Property Let SelStart(ByVal Start As Long)

cmbMain.SelStart = Start

End Property

Property Get SelText() As String

SelText = cmbMain.SelText

End Property


Property Let ItemBackColor(Index As Long, BackColor As OLE_COLOR)
cmbMain.ItemBackColor(Index) = BackColor
End Property

Property Get ItemBackColor(Index As Long) As OLE_COLOR
BackColor = cmbMain.ItemBackColor(Index)
End Property


Property Let ItemData(Index As Long, ItemData As Long)
cmbMain.ItemData(Index) = ItemData
End Property

Property Get ItemData(Index As Long) As Long
ItemData = cmbMain.ItemData(Index)
End Property


Property Let ItemExtraData(Index As Long, ItemExtraData As Long)
cmbMain.ItemExtraData(Index) = ItemExtraData
End Property

Property Get ItemExtraData(Index As Long) As Long
ItemExtraData = cmbMain.ItemExtraData(Index)
End Property


Property Let ItemFont(Index As Long, ItemFont As StdFont)
cmbMain.ItemFont(Index) = ItemFont
End Property

Property Get ItemFont(Index As Long) As StdFont
ItemFont = cmbMain.ItemFont(Index)
End Property


Property Let ItemForeColor(Index As Long, ItemForeColor As OLE_COLOR)
cmbMain.ItemForeColor(Index) = ItemForeColor
End Property

Property Get ItemForeColor(Index As Long) As OLE_COLOR
ItemForeColor = cmbMain.ItemForeColor(Index)

End Property


Property Let ItemHeight(Index As Long, ItemHeight As Long)
cmbMain.ItemHeight(Index) = ItemHeight
End Property

Property Get ItemHeight(Index As Long) As Long
ItemHeight = cmbMain.ItemHeight(Index)

End Property


Property Let ItemIcon(Index As Long, ItemIcon As Long)
cmbMain.ItemIcon(Index) = ItemIcon
End Property

Property Get ItemIcon(Index As Long) As Long
ItemIcon = cmbMain.ItemIcon(Index)
End Property

Property Let ItemIndent(Index As Long, Indent As Long)
cmbMain.ItemIndent(Index) = Indent
End Property

Property Get ItemIndent(Index As Long) As Long
ItemIndent = cmbMain.ItemIndent(Index)

End Property


Property Let ItemOverLine(Index As Long, OverLine As Boolean)
cmbMain.ItemOverLine(Index) = OverLine
End Property

Property Get ItemOverLine(Index As Long) As Boolean
ItemOverLine = cmbMain.ItemOverLine(Index)

End Property


Property Let ItemUnderLine(Index As Long, UnderLine As Boolean)
cmbMain.ItemUnderLine(Index) = UnderLine
End Property

Property Get ItemUnderLine(Index As Long) As Boolean
ItemUnderLine = cmbMain.ItemUnderLine(Index)

End Property

Property Let ItemXAlign(Index As Long, Alignment As ItemXAlignEnum)
cmbMain.ItemXAlign(Index) = Alignment
End Property

Property Get ItemXAlign(Index As Long) As ItemXAlignEnum
ItemXAlign = cmbMain.ItemXAlign(Index)

End Property

Property Let ItemYAlign(Index As Long, Alignment As ItemYAlignEnum)
cmbMain.ItemYAlign(Index) = Alignment
End Property

Property Get ItemYAlign(Index As Long) As ItemYAlignEnum
ItemYAlign = cmbMain.ItemYAlign(Index)

End Property




'-----------------------------------------------------------------------------------------------------
'Private internal subs


Private Sub pSetDrawStyle()

If m_DefaultStyle = NoStyle Then 'Only allow DrawStyles to be set if a default style is not set.
                                 'This is because the if a default style is set it automatically sets the
                                 'draw style of the combobox
    Select Case m_DrawMode
    
    Case Is = dmStandard
    cmbMain.ClientDraw = ecdNoClientDraw
    
    Case Is = dmColorPickerNoNames
    cmbMain.ClientDraw = ecdColourPickerNoNames
    
    Case Is = dmColorPicker
    cmbMain.ClientDraw = ecdColourPickerWithNames
    
    Case Is = dmFontPicker
    cmbMain.ClientDraw = ecdFontPicker
    
    Case Is = dmParagraphStyles
    cmbMain.ClientDraw = ecdParagraphStyles
    
    End Select
    
End If

End Sub

Private Sub pSetDefaultStyle()

Select Case m_DefaultStyle

Case Is = ColorPicker
cmbMain.ClientDraw = ecdSysColourPicker
m_DrawMode = dmColorPicker
LoadSysColorList cmbMain

Case Is = ColorPickerNoNames
cmbMain.ClientDraw = ecdColourPickerNoNames
m_DrawMode = dmColorPickerNoNames
LoadSysColorList cmbMain

Case Is = FontPicker
cmbMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
LoadFontList cmbMain, "", -1, -1

Case Is = FontViewer
cmbMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
LoadFontListViewer cmbMain, "", -1, -1, True

Case Is = FontViewerNoIcons
cmbMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
LoadFontListViewer cmbMain, "", -1, -1, False

Case Is = ParagraphStyles
cmbMain.Clear
cmbMain.ClientDraw = ecdParagraphStyles
m_DrawMode = dmParagraphStyles
LoadParagraphStyles cmbMain

End Select

End Sub


Private Sub pSetEnabled()

If Ambient.UserMode = False Then
UserControl.Enabled = False

    If m_Enabled = True Then
    pSetGraphic "UnPressed"
    cmbMain.Visible = True
    
    Else
    pSetGraphic "Disabled"
    cmbMain.Visible = False
    
    End If


Else

    If m_Enabled = True Then
    UserControl.Enabled = True
    pSetGraphic "UnPressed"
    cmbMain.Enabled = True

    
        If Locked = True Then
        UserControl.Enabled = False
        End If
    
    Else
    UserControl.Enabled = False
    pSetGraphic "Disabled"
    cmbMain.Enabled = False
    cmbMain.ForeColor = &H92A1A1
    cmbMain.BackColor = &H8000000F
    
    
    End If

End If


End Sub

Private Sub pSetBorder()

If m_Enabled = True Then
strEnabled = "Enabled\"
Else
strEnabled = "Disabled\"
End If

If ResourceLib.hModule <> 0 Then

Top.Cls
Left.Cls
Bottom.Cls

Top.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopLeft", crBitmap), 0, 0
Top.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY

Left.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "Left", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Left.Height

Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomLeft", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY

Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\" & strEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY

End If

End Sub

Private Sub pSetGraphic(New_Value As String)

imgButton.Cls


imgButton.PaintPicture PictureFromResource(ResourceLib.hModule, "ComboBox\" & New_Value & "\Back", crBitmap), 5 * Screen.TwipsPerPixelX, 5 * Screen.TwipsPerPixelY, imgButton.Width, imgButton.Height
imgButton.PaintPicture PictureFromResource(ResourceLib.hModule, "ComboBox\" & New_Value & "\Top", crBitmap), 0, 0
imgButton.PaintPicture PictureFromResource(ResourceLib.hModule, "ComboBox\" & New_Value & "\Left", crBitmap), 0, 5 * Screen.TwipsPerPixelY, 5 * Screen.TwipsPerPixelX, imgButton.Height - (10 * Screen.TwipsPerPixelY)

imgButton.PaintPicture PictureFromResource(ResourceLib.hModule, "ComboBox\" & New_Value & "\Right", crBitmap), imgButton.Width - (4 * Screen.TwipsPerPixelX), 5 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelX, imgButton.Height - (10 * Screen.TwipsPerPixelY)
imgButton.PaintPicture PictureFromResource(ResourceLib.hModule, "ComboBox\" & New_Value & "\Bottom", crBitmap), 0, imgButton.Height - (5 * Screen.TwipsPerPixelY)

If glbAppearance <> Win98 Then
imgButton.PaintPicture PictureFromResource(ResourceLib.hModule, "ComboBox\" & New_Value & "\Arrow", crBitmap), 6 * Screen.TwipsPerPixelX, (imgButton.Height / 2) - (3 * Screen.TwipsPerPixelY)

Else

    If New_Value <> "Pressed" Then
    imgButton.PaintPicture PictureFromResource(ResourceLib.hModule, "ComboBox\" & New_Value & "\Arrow", crBitmap), 6 * Screen.TwipsPerPixelX, (imgButton.Height / 2) - (2 * Screen.TwipsPerPixelY)
    Else
    imgButton.PaintPicture PictureFromResource(ResourceLib.hModule, "ComboBox\" & New_Value & "\Arrow", crBitmap), 7 * Screen.TwipsPerPixelX, (imgButton.Height / 2) - (1 * Screen.TwipsPerPixelY)
   
    End If

End If

End Sub



'---------------------------------------------------------------------------------------------------
'Control Properties


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,AutoCompleteListItemsOnly
Public Property Get AutoCompleteListItemsOnly() As Boolean
Attribute AutoCompleteListItemsOnly.VB_Description = "Returns/sets whether a combobox with AutoComplete on should only allow selections of items in the list, or should allow any text to be entered."
    AutoCompleteListItemsOnly = cmbMain.AutoCompleteListItemsOnly
End Property

Public Property Let AutoCompleteListItemsOnly(ByVal New_AutoCompleteListItemsOnly As Boolean)
    cmbMain.AutoCompleteListItemsOnly() = New_AutoCompleteListItemsOnly
    PropertyChanged "AutoCompleteListItemsOnly"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,DropDownWidth
Public Property Get DropDownWidth() As Long
Attribute DropDownWidth.VB_Description = "Returns/sets the width of drop down portion of a combo box."
    DropDownWidth = cmbMain.DropDownWidth
End Property

Public Property Let DropDownWidth(ByVal New_DropDownWidth As Long)
    cmbMain.DropDownWidth() = New_DropDownWidth
    PropertyChanged "DropDownWidth"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DrawMode() As ComboBoxDrawModeEnum
Attribute DrawMode.VB_Description = "Returns/sets whether the combobox is automatically setup in a common style."
    DrawMode = m_DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As ComboBoxDrawModeEnum)
    m_DrawMode = New_DrawMode
    PropertyChanged "DrawMode"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
    Set Font = cmbMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Set cmbMain.Font = New_Font

UserControl_Resize
Refresh

PropertyChanged "Font"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Returns/sets whether items in the combobox are fully highlighted or if just the text is highlighted."
    FullRowSelect = cmbMain.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    cmbMain.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets whether the combobox is locked."
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
'cmbMain.Locked() = New_Locked
m_Locked = New_Locked

If New_Locked = True Then
UserControl.Enabled = False
Else
    If m_Enabled = True Then
    UserControl.Enabled = True
    End If


End If


PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum amount of text which can be entered into the edit portion of a drop-down combo box."
    MaxLength = cmbMain.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    cmbMain.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
Set MouseIcon = imgButton.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
Set imgButton.MouseIcon = New_MouseIcon
PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
MousePointer = imgButton.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
imgButton.MousePointer() = New_MousePointer
PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,Sorted
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Returns/sets whether the list items in the control will be sorted."
    Sorted = cmbMain.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    cmbMain.Sorted() = New_Sorted
    PropertyChanged "Sorted"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,DoAutoComplete
Public Property Get AutoComplete() As Boolean
Attribute AutoComplete.VB_Description = "Returns/sets whether the combobox will attempt to automatically complete the user's typing based on the contents of the list."
    AutoComplete = cmbMain.DoAutoComplete
End Property

Public Property Let AutoComplete(ByVal New_AutoComplete As Boolean)
    cmbMain.DoAutoComplete() = New_AutoComplete
    PropertyChanged "AutoComplete"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DefaultStyle() As ComboBoxDefaultStyleEnum
Attribute DefaultStyle.VB_Description = "Returns/sets whether the combobox is automatically setup in a common style."
DefaultStyle = m_DefaultStyle
End Property

Public Property Let DefaultStyle(ByVal New_DefaultStyle As ComboBoxDefaultStyleEnum)
m_DefaultStyle = New_DefaultStyle

pSetDefaultStyle

If m_DefaultStyle = NoStyle Then
m_DrawMode = dmStandard
cmbMain.ClientDraw = ecdClientDrawOnly
End If


PropertyChanged "DefaultStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
m_Enabled = New_Enabled

pSetEnabled
pSetBorder

PropertyChanged "Enabled"
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
Public Property Get ToolTipIcon() As ToolTipIconEnum
Attribute ToolTipIcon.VB_Description = "Returns/sets the icon displayed in the tooltip."
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
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

If m_Enabled = True Then
m_ForeColor = New_ForeColor
cmbMain.ForeColor = New_ForeColor
End If

PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

If m_Enabled = True Then
m_BackColor = New_BackColor
cmbMain.BackColor = New_BackColor
End If


PropertyChanged "BackColor"
End Property



'----------------------------------------------------------------------------------------------
'Read\Write\Init properties

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
    
    m_DefaultStyle = NoStyle
    m_Enabled = True
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    m_ForeColor = 0
    m_BackColor = &HFFFFFF
    m_DrawMode = dmStandard
    m_Locked = False
    
    Width = 110 * Screen.TwipsPerPixelX
    
    ChangeTheme

    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    cmbMain.AutoCompleteListItemsOnly = PropBag.ReadProperty("AutoCompleteListItemsOnly", False)
    cmbMain.DropDownWidth = PropBag.ReadProperty("DropDownWidth", 0)
    Set cmbMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Locked = PropBag.ReadProperty("Locked", False)
    cmbMain.MaxLength = PropBag.ReadProperty("MaxLength", 30000)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    imgButton.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    cmbMain.Sorted = PropBag.ReadProperty("Sorted", False)
    cmbMain.DoAutoComplete = PropBag.ReadProperty("AutoComplete", False)
    m_DefaultStyle = PropBag.ReadProperty("DefaultStyle", NoStyle)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_DrawMode = PropBag.ReadProperty("DrawMode", dmStandard)
    cmbMain.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    
    blnToolTipSet = False
    
    If m_Enabled = True Then
    cmbMain.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    cmbMain.ForeColor = PropBag.ReadProperty("ForeColor", 0)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    End If
      
    If glbControlsRefreshed = True Then
    RefreshTheme 'Refresh control because tilebar refresh code has already run before this
                 'This is needed to make sure the correct appearance is displayed
    End If
      
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoCompleteListItemsOnly", cmbMain.AutoCompleteListItemsOnly, False)
    Call PropBag.WriteProperty("DropDownWidth", cmbMain.DropDownWidth, 0)
    Call PropBag.WriteProperty("Font", cmbMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("Locked", m_Locked, False)
    Call PropBag.WriteProperty("MaxLength", cmbMain.MaxLength, 30000)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", imgButton.MousePointer, 0)
    Call PropBag.WriteProperty("Sorted", cmbMain.Sorted, False)
    Call PropBag.WriteProperty("AutoComplete", cmbMain.DoAutoComplete, False)
    Call PropBag.WriteProperty("DefaultStyle", m_DefaultStyle, NoStyle)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("DrawMode", m_DrawMode, dmStandard)
    Call PropBag.WriteProperty("FullRowSelect", cmbMain.FullRowSelect, False)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")

End Sub


'---------------------------------------------------------------------------------------------------
'Control Events

Private Sub cmbMain_AutoCompleteSelection(ByVal sItem As String, ByVal lIndex As Long)
RaiseEvent AutoCompleteSelection(sItem, lIndex)
End Sub

Private Sub cmbMain_Change()
RaiseEvent Change
End Sub

Private Sub cmbMain_Click()
RaiseEvent Click
End Sub

Private Sub cmbMain_CloseUp()
pSetGraphic "UnPressed"
RaiseEvent CloseUp
End Sub

Private Sub cmbMain_DblClick()
RaiseEvent DblClick
End Sub

Private Sub cmbMain_DropDown()
pSetGraphic "Pressed"
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

Private Sub cmbMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub cmbMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub cmbMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub cmbMain_SelCancel()
RaiseEvent SelCancel
End Sub


Private Sub imgButton_MouseDown(imgButton As Integer, Shift As Integer, X As Single, Y As Single)


pSetGraphic "Pressed"
cmbMain.ShowDropDown True


End Sub

Private Sub imgButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If glbFormHasFocus = True Then

If blnToolTipSet = False Then
pSetToolTip imgButton.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
blnToolTipSet = True
End If

With imgButton
If GetCapture() = .hwnd Then
 
    If ((X < 0) Or (X > .Width)) Or ((Y < 0) Or (Y > .Height)) Then
    'if the mouse is outside the bounds of the control
    ' release the mouse and reset the backcolor
    Call ReleaseCapture

        pSetGraphic "UnPressed" 'Set unpressed value because mouse
                             'has left control
     
        blnToolTipSet = False
     
    End If 'Checking if mouse is inside control
    
    
Else ' otherwise capture the mouse and change the backcolor of the control

Call SetCapture(.hwnd)

    pSetGraphic "MouseOver"

    'RaiseEvent MouseMove(imgButton, Shift, x, y)
             
End If

End With


End If

End Sub

Private Sub imgButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pSetGraphic "UnPressed"
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


cmbMain.DropDownWidth = Width / Screen.TwipsPerPixelX

pSetBorder

err:
Exit Sub

End Sub




