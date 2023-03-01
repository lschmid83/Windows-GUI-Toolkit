VERSION 5.00
Object = "{436403CD-EDD8-11D2-8040-00C04FA4EE99}#12.0#0"; "VBALCBEX.OCX"
Begin VB.UserControl ExplorerCombo 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   ScaleHeight     =   525
   ScaleWidth      =   2220
   ToolboxBitmap   =   "ExplorerCombo].ctx":0000
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
   Begin vbalComboEx.vbalCboEx cmbMain 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ExtendedUI      =   0   'False
      DropDownWidth   =   0
   End
End
Attribute VB_Name = "ExplorerCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Citex Software                                                       '
'XP GUI Toolkit v1.0 - ComboBox Component v1.0                        '
'Copyright 2001 Lawrence Schmid                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


Dim blnToolTipSet As Boolean
Dim strEnabled As String

'Property Variables:
Dim m_ToolTipCaption As String
Dim m_DefaultStyle As ExplorerComboDefaultStyleEnum
Dim m_Enabled As Boolean
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String

'Event Declarations:
Event AutoCompleteSelection(ByVal sItem As String, ByVal lIndex As Long) 'MappingInfo=cmbMain,cmbMain,-1,AutoCompleteSelection
Attribute AutoCompleteSelection.VB_Description = "Raised whenever the Auto Complete mode selects an item."
Event BeginEdit(ByVal iIndex As Long) 'MappingInfo=cmbMain,cmbMain,-1,BeginEdit
Attribute BeginEdit.VB_Description = "Raised when the user begins editing the text box section of the ComboBox."
Event Change() 'MappingInfo=cmbMain,cmbMain,-1,Change
Attribute Change.VB_Description = "Raised when the text in the combo box is changed."
Event Click() 'MappingInfo=cmbMain,cmbMain,-1,Click
Attribute Click.VB_Description = "Raised when the user selects an item by clicking on it and when the ListIndex property is set in code."
Event CloseUp() 'MappingInfo=cmbMain,cmbMain,-1,CloseUp
Attribute CloseUp.VB_Description = "Raised when the ComboBox portion of the control is closed up."
Event DblClick() 'MappingInfo=cmbMain,cmbMain,-1,DblClick
Attribute DblClick.VB_Description = "Raised when the control is double clicked."
Event DrawItem(ByVal ItemIndex As Long, ByVal hDC As Long, ByVal bSelected As Boolean, ByVal bEnabled As Boolean, ByVal LeftPixels As Long, ByVal TopPixels As Long, ByVal RightPixels As Long, ByVal BottomPixels As Long, ByVal hFntOld As Long) 'MappingInfo=cmbMain,cmbMain,-1,DrawItem
Attribute DrawItem.VB_Description = "Raised when an item in the combo box must be drawn."
Event DropDown() 'MappingInfo=cmbMain,cmbMain,-1,DropDown
Attribute DropDown.VB_Description = "Raised whenever the drop-down portion of the combo-box is shown."
Event EndEdit(ByVal iIndex As Long, ByVal bChanged As Boolean, ByVal sText As String, eWHy As ECCXEndEditReason, ByVal iNewIndex As Long) 'MappingInfo=cmbMain,cmbMain,-1,EndEdit
Attribute EndEdit.VB_Description = "Raised when the user finished editing the text portion of the Combo Box."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Raised when a KeyDown occurs in the control."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Raised when a Key is pressed  in the control."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Raised when a KeyUp occurs in the control."
Event RequestDropDownResize(lLeft As Long, lTop As Long, lRight As Long, lBottom As Long, bCancel As Boolean) 'MappingInfo=cmbMain,cmbMain,-1,RequestDropDownResize
Attribute RequestDropDownResize.VB_Description = "Raised when the drop down portion of the control is about to be shown.  You can customise the position at which the drop down appears."
Event MouseLeave()


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


pSetDefaultStyle
pSetEnabled
pSetBorder

UserControl_Resize



End Sub


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


pSetDefaultStyle
pSetEnabled
pSetBorder

UserControl_Resize

End Sub


Public Sub AddItem(Text As String)

cmbMain.AddItem Text

End Sub

Public Sub AddItemAndData( _
        Text As String, _
        Optional IconIndex As Long = -1, _
        Optional IconSelected As Long = -1, _
        Optional ItemData As Long = 0, _
        Optional Indent As Long = 0)

cmbMain.AddItemAndData Text, IconIndex, IconSelected, ItemData, Indent

End Sub


Public Sub InsertItem(Text As String, Optional IndexBefore As Long = -1)

cmbMain.InsertItem Text, IndexBefore

End Sub

Public Sub InsertItemAndData( _
        Text As String, _
        Optional IndexBefore As Long = -1, _
        Optional IconIndex As Long = -1, _
        Optional IconSelected As Long = -1, _
        Optional ItemData As Long = 0, _
        Optional Indent As Long = 0)
        
        
cmbMain.InsertItemAndData Text, IndexBefore, IconIndex, IconSelected, ItemData, Indent


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

Property Get SelStart() As Long

If Ambient.UserMode = True Then
SelStart = cmbMain.SelStart
End If

End Property

Property Get SelText() As String

SelText = cmbMain.SelText

End Property

'--------------------------------------------------------------------------------------
'Private Internal subs

Private Sub pSetToolTip()

    ToolTip.SetParentHwnd imgButton.hwnd
    ToolTip.TipText = m_ToolTipCaption
    
    If m_ToolTipStyle = Standard Then
    ToolTip.Style = TTStandard
    Else
    ToolTip.Style = TTBalloon
    End If
    
    If m_ToolTipIcon = Error Then
    ToolTip.Icon = TTIconError
    
    ElseIf m_ToolTipIcon = Info Then
    ToolTip.Icon = TTIconInfo
    
    ElseIf m_ToolTipIcon = Warning Then
    ToolTip.Icon = TTIconWarning
    
    ElseIf m_ToolTipIcon = NoIcon Then
    ToolTip.Icon = TTNoIcon
    End If

    ToolTip.ForeColor = 0
    ToolTip.BackColor = 0
    ToolTip.Title = m_ToolTipTitle
    ToolTip.Create

End Sub


Private Sub pSetDefaultStyle()


Select Case m_DefaultStyle

Case Is = ExplorerNone
cmbMain.Clear

Case Is = DriveList
LoadDriveList cmbMain, False


'cmbMain.DrawStyle = eccxDriveList

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
    cmbMain.Visible = True
        
    Else
    UserControl.Enabled = False
    pSetGraphic "Disabled"
    cmbMain.Visible = False
    
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
Public Property Get DefaultStyle() As ExplorerComboDefaultStyleEnum
Attribute DefaultStyle.VB_Description = "Returns/sets whether the combobox is automatically setup in a common style."
DefaultStyle = m_DefaultStyle
End Property

Public Property Let DefaultStyle(ByVal New_DefaultStyle As ExplorerComboDefaultStyleEnum)
m_DefaultStyle = New_DefaultStyle

pSetDefaultStyle

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





Private Sub cmbMain_Change()
    RaiseEvent Change


End Sub

Private Sub cmbMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)




End Sub



'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
    
    m_DefaultStyle = NoStyle
    m_Enabled = True
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""

    Width = 110 * Screen.TwipsPerPixelX

    ChangeTheme

    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    cmbMain.AutoCompleteListItemsOnly = PropBag.ReadProperty("AutoCompleteListItemsOnly", False)
    cmbMain.DropDownWidth = PropBag.ReadProperty("DropDownWidth", 0)
    Set cmbMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
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
    
    blnToolTipSet = False

   
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
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", imgButton.MousePointer, 0)
    Call PropBag.WriteProperty("Sorted", cmbMain.Sorted, False)
    Call PropBag.WriteProperty("AutoComplete", cmbMain.DoAutoComplete, False)
    Call PropBag.WriteProperty("DefaultStyle", m_DefaultStyle, NoStyle)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")

End Sub


'---------------------------------------------------------------------------------------------------
'Control Events

Private Sub imgButton_MouseDown(imgButton As Integer, Shift As Integer, X As Single, Y As Single)

pSetGraphic "Pressed"
cmbMain.ShowDropDown True

End Sub

Private Sub imgButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If glbFormHasFocus = True Then


If blnToolTipSet = False Then

pSetToolTip

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


Private Sub UserControl_KeyPress(KeyAscii As Integer)
MsgBox ""
If KeyAscii = 13 Then

cmbMain.AddItem cmbMain.Text

End If


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

pSetBorder

err:
Exit Sub

End Sub

Private Sub cmbMain_AutoCompleteSelection(ByVal sItem As String, ByVal lIndex As Long)
    RaiseEvent AutoCompleteSelection(sItem, lIndex)
End Sub

Private Sub cmbMain_BeginEdit(ByVal iIndex As Long)
    RaiseEvent BeginEdit(iIndex)
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

Private Sub cmbMain_DrawItem(ByVal ItemIndex As Long, ByVal hDC As Long, ByVal bSelected As Boolean, ByVal bEnabled As Boolean, ByVal LeftPixels As Long, ByVal TopPixels As Long, ByVal RightPixels As Long, ByVal BottomPixels As Long, ByVal hFntOld As Long)
    RaiseEvent DrawItem(ItemIndex, hDC, bSelected, bEnabled, LeftPixels, TopPixels, RightPixels, BottomPixels, hFntOld)
End Sub

Private Sub cmbMain_DropDown()
pSetGraphic "Pressed"
RaiseEvent DropDown
End Sub


Private Sub cmbMain_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cmbMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub cmbMain_RequestDropDownResize(lLeft As Long, lTop As Long, lRight As Long, lBottom As Long, bCancel As Boolean)
    RaiseEvent RequestDropDownResize(lLeft, lTop, lRight, lBottom, bCancel)
End Sub



