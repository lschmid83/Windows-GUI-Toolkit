VERSION 5.00
Object = "{F1909D6D-FB9D-11D3-B06C-00500427A693}#1.0#0"; "XUITREEVIEW6.OCX"
Begin VB.UserControl TreeView 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   ScaleHeight     =   2025
   ScaleWidth      =   2715
   ToolboxBitmap   =   "TreeView.ctx":0000
   Begin xuiTreeView6.TreeView tvMain 
      Height          =   2025
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   3572
      RootLines       =   0   'False
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Indent          =   16
      MaxScrollTime   =   -1
   End
End
Attribute VB_Name = "TreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Citex Software                                                       '
'XP GUI Toolkit v1.0 - TreeView Component v1.0                        '
'Copyright 2001 Lawrence Schmid                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


'Property Variables:
Dim m_Border As Boolean
Dim m_Enabled As Boolean

'Event Declarations:
Event AfterLabelEdit(hItem As Long, NewText As String, Cancel As Boolean) 'MappingInfo=tvMain,tvMain,-1,AfterLabelEdit
Event BeforeLabelEdit(hItem As Long, Cancel As Boolean) 'MappingInfo=tvMain,tvMain,-1,BeforeLabelEdit
Event Click(X As Long, Y As Long, RightButton As Boolean) 'MappingInfo=tvMain,tvMain,-1,Click
Attribute Click.VB_Description = "Raised when an item in the TreeView is clicked."
Event CustomSort(hItemParent As Long, hItem1 As Long, hItem2 As Long, ByVal Order As SortOrderConstants) 'MappingInfo=tvMain,tvMain,-1,CustomSort
Event DblClick(X As Long, Y As Long) 'MappingInfo=tvMain,tvMain,-1,DblClick
Attribute DblClick.VB_Description = "Raised when an item is double-clicked in the TreeView."
Event DragBegin(ByVal hItem As Long) 'MappingInfo=tvMain,tvMain,-1,DragBegin
Event DragEnd(MoveItem As Boolean) 'MappingInfo=tvMain,tvMain,-1,DragEnd
Event DragError(ErrCode As ErrCodeConstants) 'MappingInfo=tvMain,tvMain,-1,DragError
Event DragMove(X As Long, Y As Long) 'MappingInfo=tvMain,tvMain,-1,DragMove
Event EnterPress() 'MappingInfo=tvMain,tvMain,-1,EnterPress
Event GetItemToolTipText(hItem As Long, sText As String) 'MappingInfo=tvMain,tvMain,-1,GetItemToolTipText
Event InitLabelEditControl(ByVal hWndEdit As Long) 'MappingInfo=tvMain,tvMain,-1,InitLabelEditControl
Event ItemClick(hItem As Long, RightButton As Boolean) 'MappingInfo=tvMain,tvMain,-1,ItemClick
Event ItemDblClick(hItem As Long) 'MappingInfo=tvMain,tvMain,-1,ItemDblClick
Event ItemDelete(hItem As Long) 'MappingInfo=tvMain,tvMain,-1,ItemDelete
Event ItemExpand(hItem As Long, ByVal ExpandType As ExpandTypeConstants) 'MappingInfo=tvMain,tvMain,-1,ItemExpand
Event ItemExpanding(hItem As Long, ByVal ExpandType As ExpandTypeConstants) 'MappingInfo=tvMain,tvMain,-1,ItemExpanding
Event ItemExpandingCancel(hItem As Long, ByVal ExpandType As ExpandTypeConstants, Cancel As Boolean) 'MappingInfo=tvMain,tvMain,-1,ItemExpandingCancel
Event KeyDown(ByVal KeyCode As KeyCodeConstants) 'MappingInfo=tvMain,tvMain,-1,KeyDown
Event SelChanged() 'MappingInfo=tvMain,tvMain,-1,SelChanged
Event SelChanging() 'MappingInfo=tvMain,tvMain,-1,SelChanging
Event SingleExpanded(hItem As Long, ByVal ExpandType As ExpandTypeConstants) 'MappingInfo=tvMain,tvMain,-1,SingleExpanded
Event TVLostFocus() 'MappingInfo=tvMain,tvMain,-1,TVLostFocus

'----------------------------------------------------------------------------------------
'Public Control subs

Public Sub ChangeTheme()

pSetAppearance
pSetEnabled
UserControl_Resize

End Sub


Public Sub RefreshTheme()

pSetAppearance
pSetEnabled
UserControl_Resize

End Sub

Public Function Add(RelativeItem, _
               Relation As RelationConstants, _
               Key As String, _
               Text As String, _
               Optional IconIndex As Long = -1, _
               Optional SelectedIconIndex As Long = -1, _
               Optional IntegralHeight As Long = -1, _
               Optional Bold As Boolean = False) As Long
                       

Add = tvMain.Add(RelativeItem, Relation, Key, Text, IconIndex, SelectedIconIndex, IntegralHeight, Bold)

End Function


Public Sub Clear(Optional Item As Long = -1)

tvMain.Clear Item


End Sub

Public Sub EnsureVisible(Item As Boolean)

tvMain.EnsureVisible Item

End Sub

Public Sub HitTest(X As Long, Y As Long)

tvMain.HitTest X, Y

End Sub

Public Sub HitTestInfo(X As Long, Y As Long)

tvMain.HitTestInfo X, Y

End Sub

Public Sub IsState(Item, Value As Long, Optional UseAsMask As Boolean = False)

tvMain.IsState Item, Value, UseAsMask

End Sub

Public Sub ItemToggle(Item)

tvMain.ItemToggle Item

End Sub

Public Sub LabelEdit(Item)

tvMain.LabelEdit Item

End Sub

Public Sub LabelEditEnd(SaveChanges As Boolean)

tvMain.LabelEditEnd SaveChanges

End Sub

Public Sub Refresh()

tvMain.Refresh

End Sub

Public Sub Remove(Item)

tvMain.Remove Item

End Sub

Public Sub StockCustomSort(Item1 As Long, Item2 As Long, SortType As StockCustomSortConstants, Optional Backwards As Boolean = False)

tvMain.StockCustomSort Item1, Item2, SortType, Backwards

End Sub



Property Get Count() As Long
Count = tvMain.Count
End Property

Property Let ImageListHwnd(hwnd As Long)
tvMain.hImageList = hwnd
End Property

Property Get ImageListHwnd() As Long
ImageListHwnd = tvMain.hImageList
End Property

Property Let StateImageListHwnd(hwnd As Long)
tvMain.hStateImageList = hwnd
End Property

Property Get StateImageListHwnd() As Long
StateImageListHwnd = tvMain.hStateImageList
End Property



Property Let DisableCustomDraw(CustomDraw As Boolean)
tvMain.DisableCustomDraw = CustomDraw
End Property

Property Get DisableCustomDraw() As Boolean
DisableCustomDraw = tvMain.DisableCustomDraw
End Property


Property Let ItemBackColor(Item, BackColor As OLE_COLOR)
tvMain.ItemBackColor(Item) = BackColor
End Property

Property Get ItemBackColor(Item) As OLE_COLOR
ItemBackColor = tvMain.ItemBackColor(Item)
End Property


Property Let ItemBold(Item, Bold As OLE_COLOR)
tvMain.ItemBold(Item) = Bold
End Property

Property Get ItemBold(Item) As OLE_COLOR
ItemBold = tvMain.ItemBold(Item)
End Property


Property Let ItemChecked(Item, Checked As Boolean)
tvMain.ItemChecked(Item) = Checked
End Property

Property Get ItemChecked(Item) As Boolean
ItemChecked = tvMain.ItemChecked(Item)
End Property

'Property Let ItemChild(Item, Child As Long)
'tvMain.ItemChild(Item) = Child
'End Property

Property Get ItemChild(Item) As Long
ItemChild = tvMain.ItemChild(Item)
End Property

Property Let ItemColor(Item, Color As OLE_COLOR)
tvMain.ItemColor(Item) = Color
End Property

Property Get ItemColor(Item) As OLE_COLOR
ItemColor = tvMain.ItemColor(Item)
End Property

Property Let ItemCut(Item, Cut As Boolean)
tvMain.ItemCut(Item) = Cut
End Property

Property Get ItemCut(Item) As Boolean
ItemCut = tvMain.ItemCut(Item)
End Property

Property Let ItemData(Item, Data As Long)
tvMain.ItemData(Item) = Data
End Property

Property Get ItemData(Item) As Long
ItemData = tvMain.ItemData(Item)
End Property


Property Let ItemDropHighlight(Item, Highlight As Boolean)
tvMain.ItemDropHighlight(Item) = Highlight
End Property

Property Get ItemDropHighlight(Item) As Boolean
ItemDropHighlight = tvMain.ItemDropHighlight(Item)
End Property

Property Let ItemExpanded(Item, Expanded As Boolean)
tvMain.ItemExpanded(Item) = Expanded
End Property

Property Get ItemExpanded(Item) As Boolean
ItemExpanded = tvMain.ItemExpanded(Item)
End Property

'Property Let ItemExpandedOnce(Item, ExpandedOnce As Boolean)
'tvMain.ItemExpandedOnce(Item) = ExpandedOnce
'End Property

Property Get ItemExpandedOnce(Item) As Boolean
ItemExpandedOnce = tvMain.ItemExpandedOnce(Item)
End Property

Property Let ItemExpandedPartial(Item, ExpandedPartial As Boolean)
tvMain.ItemExpandedPartial(Item) = ExpandedPartial
End Property

Property Get ItemExpandedPartial(Item) As Boolean
ItemExpandedPartial = tvMain.ItemExpandedPartial(Item)
End Property


Property Let ItemFont(Item, Font As StdFont)
tvMain.ItemFont(Item) = Font
End Property

Property Get ItemFont(Item) As StdFont
ItemFont = tvMain.ItemFont(Item)
End Property


'Property Let ItemHandle(Item, Handle As Long)
'tvMain.ItemHandle(Item) = Handle
'End Property

Property Get ItemHandle(Item) As Long
ItemHandle = tvMain.ItemHandle(Item)
End Property

'Property Let ItemHasChildren(Item, HasChildren As Boolean)
'tvMain.ItemHasChildren(Item) = HasChildren
'End Property

Property Get ItemHasChildren(Item) As Boolean
ItemHasChildren = tvMain.ItemHasChildren(Item)
End Property


Property Let ItemImage(Item, IconIndex As Long)
tvMain.ItemImage(Item) = IconIndex
End Property

Property Get ItemImage(Item) As Long
ItemImage = tvMain.ItemImage(Item)

End Property

'Property Let ItemIndex(Key, Index As String)
'tvMain.ItemIndex(Key) = Index
'End Property

Property Get ItemIndex(Key) As String
'ItemImage = tvMain.ItemIndex(Key)

End Property

Property Let ItemIntegralHeight(Item, Height As Long)
tvMain.ItemIntegralHeight(Item) = Height
End Property

Property Get ItemIntegralHeight(Item) As Long
ItemIntegralHeight = tvMain.ItemIntegralHeight(Item)

End Property

Property Let ItemKey(Item, Key As String)
tvMain.ItemKey(Item) = Key
End Property

Property Get ItemKey(Item) As String
ItemKey = tvMain.ItemKey(Item)
End Property


Property Let ItemMouseOverColor(Item, Color As OLE_COLOR)
tvMain.ItemMouseOverColor(Item) = Color
End Property

Property Get ItemMouseOverColor(Item) As OLE_COLOR
ItemMouseOverColor = tvMain.ItemMouseOverColor(Item)

End Property


'Property Let ItemNextSibling(Item, NextSibling As Long)
'tvMain.ItemNextSibling(Item) = NextSibling
'End Property

Property Get ItemNextSibling(Item) As Long
ItemNextSibling = tvMain.ItemNextSibling(Item)

End Property

'Property Let ItemNextVisible(Item, NextVisible As Long)
'tvMain.ItemNextVisible(Item) = NextVisible
'End Property

Property Get ItemNextVisible(Item) As Long
ItemNextVisible = tvMain.ItemNextVisible(Item)

End Property

Property Let ItemNumber(Item, Number As Long)
tvMain.ItemNumber(Item) = Number
End Property

Property Get ItemNumber(Item) As Long
ItemNumber = tvMain.ItemNumber(Item)

End Property

'Property Let ItemParent(Item, Parent As Long)
'tvMain.ItemParent(Item) = Parent
'End Property

Property Get ItemParent(Item) As Long
ItemParent = tvMain.ItemParent(Item)

End Property

Property Let ItemPlusMinus(Item, PlusMinus As Boolean)
tvMain.ItemPlusMinus(Item) = PlusMinus
End Property

Property Get ItemPlusMinus(Item) As Boolean
ItemPlusMinus = tvMain.ItemPlusMinus(Item)

End Property

'Property Let ItemPrevious(Item, Previous As Long)
'tvMain.ItemPrevious(Item) = Previous
'End Property

Property Get ItemPrevious(Item) As Long
ItemPrevious = tvMain.ItemPrevious(Item)

End Property

'Property Let ItemPreviousVisible(Item, PreviousVisible As Long)
'tvMain.ItemPreviousVisible(Item) = PreviousVisible
'End Property

Property Get ItemPreviousVisible(Item) As Long
ItemPreviousVisible = tvMain.ItemPreviousVisible(Item)

End Property

Property Let ItemSelected(Item, Selected As Boolean)
tvMain.ItemSelected(Item) = Selected
End Property

Property Get ItemSelected(Item) As Boolean
ItemSelected = tvMain.ItemSelected(Item)

End Property

Property Let ItemSelectedImage(Item, IconIndex As Long)
tvMain.ItemSelectedImage(Item) = IconIndex
End Property

Property Get ItemSelectedImage(Item) As Long
ItemSelectedImage = tvMain.ItemSelectedImage(Item)

End Property

Property Let ItemStateImage(Item, IconIndex As Long)
tvMain.ItemStateImage(Item) = IconIndex
End Property

Property Get ItemStateImage(Item) As Long
ItemStateImage = tvMain.ItemStateImage(Item)

End Property


Property Let ItemText(Item, Text As String)
tvMain.ItemText(Item) = Text
End Property

Property Get ItemText(Item) As String
ItemText = tvMain.ItemText(Item)

End Property



'---------------------------------------------------------------------------------------------------
'Private Internal subs

Private Sub pSetAppearance()
   
If glbAppearance <> Win98 Then
UserControl.BorderStyle = 0
tvMain.Top = 3 * Screen.TwipsPerPixelY
tvMain.left = 3 * Screen.TwipsPerPixelX

Else
UserControl.BorderStyle = 1
tvMain.Top = 0
tvMain.left = 0

End If

End Sub

Private Sub pSetEnabled()
   
If Ambient.UserMode = False Then
UserControl.Enabled = False

    If m_Enabled = True Then
    tvMain.Enabled = True
    
    Else
    tvMain.Enabled = False
    
    End If


Else

    If m_Enabled = True Then
    UserControl.Enabled = True
    tvMain.Enabled = True
    
    
    Else
    UserControl.Enabled = False
    tvMain.Enabled = False
    
    End If

End If


End Sub


Private Sub pSetBorder()

If ResourceLib.hModule <> 0 Then
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\TopLeft", crBitmap), 0, 0
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Left", crBitmap), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, Height - (3 * Screen.TwipsPerPixelX), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
End If

End Sub

  
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the backcolor of the control."
    BackColor = tvMain.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    tvMain.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,CheckBoxes
Public Property Get CheckBoxes() As Boolean
Attribute CheckBoxes.VB_Description = "Gets/sets whether checkboxes are shown for each item in the TreeView."
    CheckBoxes = tvMain.CheckBoxes
End Property

Public Property Let CheckBoxes(ByVal New_CheckBoxes As Boolean)
    tvMain.CheckBoxes() = New_CheckBoxes
    PropertyChanged "CheckBoxes"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,DragExpandTime
Public Property Get DragExpandTime() As Long
    DragExpandTime = tvMain.DragExpandTime
End Property

Public Property Let DragExpandTime(ByVal New_DragExpandTime As Long)
    tvMain.DragExpandTime() = New_DragExpandTime
    PropertyChanged "DragExpandTime"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,DragScrollTime
Public Property Get DragScrollTime() As Long
    DragScrollTime = tvMain.DragScrollTime
End Property

Public Property Let DragScrollTime(ByVal New_DragScrollTime As Long)
    tvMain.DragScrollTime() = New_DragScrollTime
    PropertyChanged "DragScrollTime"
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

PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,ExplorerBar
Public Property Get ExplorerBar() As Boolean
    ExplorerBar = tvMain.ExplorerBar
End Property

Public Property Let ExplorerBar(ByVal New_ExplorerBar As Boolean)
    tvMain.ExplorerBar() = New_ExplorerBar
    PropertyChanged "ExplorerBar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = tvMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set tvMain.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = tvMain.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    tvMain.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,FullRowSelect
Public Property Get FullRowSelect() As Boolean
    FullRowSelect = tvMain.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    tvMain.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,HotTracking
Public Property Get HotTracking() As Boolean
    HotTracking = tvMain.HotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    tvMain.HotTracking() = New_HotTracking
    PropertyChanged "HotTracking"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,Indent
Public Property Get Indent() As Long
    Indent = tvMain.Indent
End Property

Public Property Let Indent(ByVal New_Indent As Long)
    tvMain.Indent() = New_Indent
    PropertyChanged "Indent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,ItemHeight
Public Property Get ItemHeight() As Long
    ItemHeight = tvMain.ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Long)
    tvMain.ItemHeight() = New_ItemHeight
    PropertyChanged "ItemHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,LabelEditing
Public Property Get LabelEditing() As Boolean
    LabelEditing = tvMain.LabelEditing
End Property

Public Property Let LabelEditing(ByVal New_LabelEditing As Boolean)
    tvMain.LabelEditing() = New_LabelEditing
    PropertyChanged "LabelEditing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,LineColor
Public Property Get LineColor() As Long
    LineColor = tvMain.LineColor
End Property

Public Property Let LineColor(ByVal New_LineColor As Long)
    tvMain.LineColor() = New_LineColor
    PropertyChanged "LineColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,Lines
Public Property Get Lines() As Boolean
    Lines = tvMain.Lines
End Property

Public Property Let Lines(ByVal New_Lines As Boolean)
    tvMain.Lines() = New_Lines
    PropertyChanged "Lines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,PlusMinus
Public Property Get PlusMinus() As Boolean
    PlusMinus = tvMain.PlusMinus
End Property

Public Property Let PlusMinus(ByVal New_PlusMinus As Boolean)
    tvMain.PlusMinus() = New_PlusMinus
    PropertyChanged "PlusMinus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,RootLines
Public Property Get RootLines() As Boolean
    RootLines = tvMain.RootLines
End Property

Public Property Let RootLines(ByVal New_RootLines As Boolean)
    tvMain.RootLines() = New_RootLines
    PropertyChanged "RootLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,ScrollBars
Public Property Get ScrollBars() As Boolean
    ScrollBars = tvMain.ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As Boolean)
    tvMain.ScrollBars() = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,ShowNumber
Public Property Get ShowNumber() As Boolean
    ShowNumber = tvMain.ShowNumber
End Property

Public Property Let ShowNumber(ByVal New_ShowNumber As Boolean)
    tvMain.ShowNumber() = New_ShowNumber
    PropertyChanged "ShowNumber"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,ShowSelected
Public Property Get ShowSelected() As Boolean
    ShowSelected = tvMain.ShowSelected
End Property

Public Property Let ShowSelected(ByVal New_ShowSelected As Boolean)
    tvMain.ShowSelected() = New_ShowSelected
    PropertyChanged "ShowSelected"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,SingleExpand
Public Property Get SingleExpand() As Boolean
    SingleExpand = tvMain.SingleExpand
End Property

Public Property Let SingleExpand(ByVal New_SingleExpand As Boolean)
    tvMain.SingleExpand() = New_SingleExpand
    PropertyChanged "SingleExpand"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,ToolTips
Public Property Get ToolTips() As Boolean
    ToolTips = tvMain.ToolTips
End Property

Public Property Let ToolTips(ByVal New_ToolTips As Boolean)
    tvMain.ToolTips() = New_ToolTips
    PropertyChanged "ToolTips"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"
End Property






Private Sub UserControl_Resize()

On Error GoTo err

If glbAppearance <> Win98 Then
tvMain.Width = Width - (6 * Screen.TwipsPerPixelX)
tvMain.Height = Height - (6 * Screen.TwipsPerPixelY)

UserControl.Cls
pSetBorder

Else
tvMain.Width = Width - (4 * Screen.TwipsPerPixelX)
tvMain.Height = Height - (4 * Screen.TwipsPerPixelY)

End If


err:
Exit Sub

End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
    
    m_Border = True
      
    ChangeTheme
   
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)


    SetDefaultTheme
    
    tvMain.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    
    'If Ambient.UserMode = False Then
    'tvMain.CheckBoxes = PropBag.ReadProperty("CheckBoxes", False)
    'tvMain.ItemHeight = PropBag.ReadProperty("ItemHeight", 0)
    
   ' End If
    
    tvMain.DragExpandTime = PropBag.ReadProperty("DragExpandTime", 2000)
    tvMain.DragScrollTime = PropBag.ReadProperty("DragScrollTime", 500)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    tvMain.ExplorerBar = PropBag.ReadProperty("ExplorerBar", False)
    Set tvMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    tvMain.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    tvMain.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    tvMain.HotTracking = PropBag.ReadProperty("HotTracking", False)
    tvMain.Indent = PropBag.ReadProperty("Indent", 0)
    tvMain.LabelEditing = PropBag.ReadProperty("LabelEditing", False)
    tvMain.LineColor = PropBag.ReadProperty("LineColor", -2147483640)
    tvMain.Lines = PropBag.ReadProperty("Lines", False)
    tvMain.PlusMinus = PropBag.ReadProperty("PlusMinus", False)
    tvMain.RootLines = PropBag.ReadProperty("RootLines", False)
    tvMain.ScrollBars = PropBag.ReadProperty("ScrollBars", True)
    tvMain.ShowNumber = PropBag.ReadProperty("ShowNumber", False)
    tvMain.ShowSelected = PropBag.ReadProperty("ShowSelected", False)
    tvMain.SingleExpand = PropBag.ReadProperty("SingleExpand", False)
    tvMain.ToolTips = PropBag.ReadProperty("ToolTips", False)
    m_Border = PropBag.ReadProperty("Border", True)

    If glbControlsRefreshed = True Then
    RefreshTheme 'Refresh control because tilebar refresh code has already run before this
                 'This is needed to make sure the correct appearance is displayed
    End If

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", tvMain.BackColor, &H80000005)
    Call PropBag.WriteProperty("CheckBoxes", tvMain.CheckBoxes, False)
    Call PropBag.WriteProperty("DragExpandTime", tvMain.DragExpandTime, 2000)
    Call PropBag.WriteProperty("DragScrollTime", tvMain.DragScrollTime, 500)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("ExplorerBar", tvMain.ExplorerBar, False)
    Call PropBag.WriteProperty("Font", tvMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", tvMain.ForeColor, &H80000008)
    Call PropBag.WriteProperty("FullRowSelect", tvMain.FullRowSelect, False)
    Call PropBag.WriteProperty("HotTracking", tvMain.HotTracking, False)
    Call PropBag.WriteProperty("Indent", tvMain.Indent, 0)
    Call PropBag.WriteProperty("ItemHeight", tvMain.ItemHeight, 0)
    Call PropBag.WriteProperty("LabelEditing", tvMain.LabelEditing, False)
    Call PropBag.WriteProperty("LineColor", tvMain.LineColor, -2147483640)
    Call PropBag.WriteProperty("Lines", tvMain.Lines, False)
    Call PropBag.WriteProperty("PlusMinus", tvMain.PlusMinus, False)
    Call PropBag.WriteProperty("RootLines", tvMain.RootLines, False)
    Call PropBag.WriteProperty("ScrollBars", tvMain.ScrollBars, True)
    Call PropBag.WriteProperty("ShowNumber", tvMain.ShowNumber, False)
    Call PropBag.WriteProperty("ShowSelected", tvMain.ShowSelected, False)
    Call PropBag.WriteProperty("SingleExpand", tvMain.SingleExpand, False)
    Call PropBag.WriteProperty("ToolTips", tvMain.ToolTips, False)
    Call PropBag.WriteProperty("Border", m_Border, True)
  
End Sub

