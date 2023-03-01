VERSION 5.00
Object = "{017E002E-D7CC-11D2-8E21-44B10AC10000}#15.2#0"; "VBALGRID.OCX"
Begin VB.UserControl GridBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "GridBox.ctx":0000
   Begin vbAcceleratorGrid.vbalGrid grdMain 
      Height          =   3600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "GridBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Citex Software                                                       '
'XP GUI Toolkit v1.0 - GridBox Component v1.0                         '
'Copyright 2001 Lawrence Schmid                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

'General Variables:
Dim m_Enabled As Boolean

'Event Declarations:
Event ColumnClick(ByVal lCol As Long) 'MappingInfo=grdMain,grdMain,-1,ColumnClick
Attribute ColumnClick.VB_Description = "Raised when the user clicks a column."
Event ColumnOrderChanged() 'MappingInfo=grdMain,grdMain,-1,ColumnOrderChanged
Event ColumnWidthChanged(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean) 'MappingInfo=grdMain,grdMain,-1,ColumnWidthChanged
Event ColumnWidthChanging(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean) 'MappingInfo=grdMain,grdMain,-1,ColumnWidthChanging
Attribute ColumnWidthChanging.VB_Description = "Raised whilst a column's width is being changed."
Event ColumnWidthStartChange(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean) 'MappingInfo=grdMain,grdMain,-1,ColumnWidthStartChange
Attribute ColumnWidthStartChange.VB_Description = "Raised when the user is about to start changing the width of a column."
Event DblClick(ByVal lRow As Long, ByVal lCol As Long) 'MappingInfo=grdMain,grdMain,-1,DblClick
Attribute DblClick.VB_Description = "Raised when the user double clicks on the grid."
Event HeaderRightClick(ByVal X As Single, ByVal Y As Single) 'MappingInfo=grdMain,grdMain,-1,HeaderRightClick
Attribute HeaderRightClick.VB_Description = "Raised when the user right clicks on the grid's header."
Event KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean) 'MappingInfo=grdMain,grdMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Raised when a key is pressed in the control."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=grdMain,grdMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Raised after the KeyDown event when the key press has been converted to an ASCII code."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=grdMain,grdMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Raised when a key is released on the grid."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, bDoDefault As Boolean) 'MappingInfo=grdMain,grdMain,-1,MouseDown
Attribute MouseDown.VB_Description = "Raised when the a mouse button is pressed over the control."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=grdMain,grdMain,-1,MouseMove
Attribute MouseMove.VB_Description = "Raised when the mouse moves over the control, or when the mouse moves anywhere and a mouse button has been pressed over the control."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=grdMain,grdMain,-1,MouseUp
Attribute MouseUp.VB_Description = "Raised when a mouse button is released after having been pressed over the control."
Event RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean) 'MappingInfo=grdMain,grdMain,-1,RequestEdit
Attribute RequestEdit.VB_Description = "Raised when the grid has the Editable property set to True and the user's actions request editing of the current cell."
Event RequestRow(ByVal lRow As Long, ByVal sKey As String, ByVal bVisible As Boolean, ByVal lHeight As Long, ByVal bGroupRow As Boolean, bNoMoreRows As Boolean) 'MappingInfo=grdMain,grdMain,-1,RequestRow
Attribute RequestRow.VB_Description = "Raised when the grid is in Virtual mode and the grid has been scrolled to expose a new row.  Set bNoMoreRows to True to indicate all rows have been added."
Event RequestRowData(ByVal lRow As Long) 'MappingInfo=grdMain,grdMain,-1,RequestRowData
Attribute RequestRowData.VB_Description = "Raised in virtual mode when a new row has been added in response to RequestRow. Respond by filling in the cells for that row."
Event SelectionChange(ByVal lRow As Long, ByVal lCol As Long) 'MappingInfo=grdMain,grdMain,-1,SelectionChange
Attribute SelectionChange.VB_Description = "Raised when the user changes the selected cell."


'--------------------------------------------------------------------------------------------
'Public control subs

Public Sub ChangeTheme()

grdMain.Top = 3 * Screen.TwipsPerPixelY
grdMain.left = 3 * Screen.TwipsPerPixelX

pSetAppearance
pSetEnabled

UserControl_Resize

End Sub


Public Sub RefreshTheme()

grdMain.Top = 3 * Screen.TwipsPerPixelY
grdMain.left = 3 * Screen.TwipsPerPixelX

pSetEnabled

UserControl_Resize

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,CancelEdit
Public Sub CancelEdit()
    grdMain.CancelEdit
End Sub

Public Sub AddColumn(Optional Key As String, _
                     Optional Text As String, _
                     Optional Alignment As ECGHdrTextAlignFlags, _
                     Optional IconIndex As Long = -1, _
                     Optional ColumnWidth As Long = -1, _
                     Optional Visible As Boolean = True, _
                     Optional Fixed As Boolean = False, _
                     Optional KeyBefore, _
                     Optional IncludeInSelect As Boolean = True, _
                     Optional FmtString As String, _
                     Optional RowTextColumn As Boolean = False, _
                     Optional SortType As cShellSortTypeConstants)
                     
grdMain.AddColumn Key, Text, Alignment, IconIndex, ColumnWidth, Visible, Fixed, KeyBefore, IncludeInSelect, FmtString, RowTextColumn, SortType


End Sub


Public Sub AddRow(Optional RowBefore As Long = -1, _
                  Optional Key As String, _
                  Optional Visible As Boolean = True, _
                  Optional Height As Long = -1, _
                  Optional GroupRow As Boolean = False, _
                  Optional GroupColStartindex As Long)
                  
                  

grdMain.AddRow RowBefore, Key, Visible, Height, GroupRow, GroupColStartindex

End Sub


Public Sub AutoHeightRow(Row As Long, Optional MinimumHeight As Long = -1)

grdMain.AutoHeightRow Row, MinimumHeight

End Sub

Public Sub AutoWidthColumn(Key)

grdMain.AutoWidthColumn Key

End Sub

Public Sub CellBoundary(Row As Long, Column As Long, left As Long, Top As Long, Width As Long, Height As Long)

grdMain.CellBoundary Row, Column, left, Top, Width, Height

End Sub

Public Sub CellDefaultBackColor(Row As Long, Column As Long)

grdMain.CellDefaultBackColor Row, Column

End Sub

Public Sub CellDefaultFont(Row As Long, Column As Long)

grdMain.CellDefaultFont Row, Column


End Sub

Public Sub CellDefaultForeColor(Row As Long, Column As Long)

grdMain.CellDefaultForeColor Row, Column

End Sub

Public Sub CellDetails(Row As Long, _
                       Column As Long, _
                       Optional Text As String, _
                       Optional Alignment As ECGHdrTextAlignFlags, _
                       Optional IconIndex As Long = -1, _
                       Optional BackColor As OLE_COLOR = -1, _
                       Optional ForeColor As OLE_COLOR = -1, _
                       Optional Font As StdFont, _
                       Optional Indent As Long, _
                       Optional ExtraIconIndex As Long = -1, _
                       Optional ItemData As Long)


grdMain.CellDetails Row, Column, Text, Alignment, IconIndex, BackColor, ForeColor, Font, Indent, ExtraIconIndex, ItemData

End Sub

Public Sub CellFromPoint(xPixels As Long, yPixels As Long, Row As Long, Column As Long)

grdMain.CellFromPoint xPixels, yPixels, Row, Column

End Sub

Public Sub Clear(Optional RemoveColumns As Boolean = False)

grdMain.Clear RemoveColumns


End Sub

Public Sub ClearSelection()

grdMain.ClearSelection

End Sub

Public Sub Draw()

grdMain.Draw

End Sub

Public Sub EnsureVisible(Row As Long, Column As Long)

grdMain.EnsureVisible Row, Column

End Sub

Public Sub FindSearchMatchRow(SearchText As String, Optional LoopThrough As Boolean = True, Optional VisibleRowsOnly As Boolean = True)

grdMain.FindSearchMatchRow SearchText, LoopThrough, VisibleRowsOnly


End Sub

Public Sub RemoveColumn(Key)

grdMain.RemoveColumn Key

End Sub

Public Sub RemoveRow(Row As Long)

grdMain.RemoveRow Row

End Sub

Public Sub SetHeaders()

grdMain.SetHeaders

End Sub

Public Sub Sort()

grdMain.Sort

End Sub


'-----------------------------------------------------------------------------------------
'Private Internal subs

Private Sub pSetAppearance()





End Sub

Private Sub pSetEnabled()

If Ambient.UserMode = False Then
UserControl.Enabled = False

grdMain.Visible = False
UserControl.BackColor = &HFFFFFF

Else
grdMain.Virtual = True

    If m_Enabled = True Then
    grdMain.Enabled = True
    
    Else
    
    grdMain.Enabled = False
    
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

  

'-----------------------------------------------------------------------------------------------------------------
'Control properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the background color of the grid."
    BackColor = grdMain.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    grdMain.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,DefaultRowHeight
Public Property Get DefaultRowHeight() As Long
Attribute DefaultRowHeight.VB_Description = "Gets/sets the height which will be used as a default for rows in the grid."
    DefaultRowHeight = grdMain.DefaultRowHeight
End Property

Public Property Let DefaultRowHeight(ByVal New_DefaultRowHeight As Long)
    grdMain.DefaultRowHeight() = New_DefaultRowHeight
    PropertyChanged "DefaultRowHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,DrawFocusRectangle
Public Property Get DrawFocusRectangle() As Boolean
Attribute DrawFocusRectangle.VB_Description = "Gets/sets whether a focus rectangle (dotted line around the selection) will be shown."
    DrawFocusRectangle = grdMain.DrawFocusRectangle
End Property

Public Property Let DrawFocusRectangle(ByVal New_DrawFocusRectangle As Boolean)
    grdMain.DrawFocusRectangle() = New_DrawFocusRectangle
    PropertyChanged "DrawFocusRectangle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,Editable
Public Property Get Editable() As Boolean
Attribute Editable.VB_Description = "Gets/sets whether the grid will be editable (i.e. raise RequestEdit events)."
    Editable = grdMain.Editable
End Property

Public Property Let Editable(ByVal New_Editable As Boolean)
    grdMain.Editable() = New_Editable
    PropertyChanged "Editable"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gets/sets whether the grid is enabled or not.  Note the grid can still be read when it is disabled, but cannot be selected or edited."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
m_Enabled = New_Enabled
    
pSetEnabled


PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gets/sets the font used by the control."
Attribute Font.VB_UserMemId = -512
    Set Font = grdMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set grdMain.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gets/sets the foreground color used to draw the control."
    ForeColor = grdMain.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    grdMain.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,GridLineColor
Public Property Get GridLineColor() As Long
Attribute GridLineColor.VB_Description = "Gets/sets the colour used to draw grid lines."
    GridLineColor = grdMain.GridLineColor
End Property

Public Property Let GridLineColor(ByVal New_GridLineColor As Long)
    grdMain.GridLineColor() = New_GridLineColor
    PropertyChanged "GridLineColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,GridLines
Public Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "Gets/sets whether grid-lines are drawn or not."
    GridLines = grdMain.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As Boolean)
    grdMain.GridLines() = New_GridLines
    PropertyChanged "GridLines"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,HeaderButtons
Public Property Get HeaderButtons() As Boolean
    HeaderButtons = grdMain.HeaderButtons
End Property

Public Property Let HeaderButtons(ByVal New_HeaderButtons As Boolean)
    grdMain.HeaderButtons() = New_HeaderButtons
    PropertyChanged "HeaderButtons"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,HeaderFlat
Public Property Get HeaderFlat() As Boolean
    HeaderFlat = grdMain.HeaderFlat
End Property

Public Property Let HeaderFlat(ByVal New_HeaderFlat As Boolean)
    grdMain.HeaderFlat() = New_HeaderFlat
    PropertyChanged "HeaderFlat"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,Virtual
Public Property Get Virtual() As Boolean
    Virtual = grdMain.Virtual
End Property

Public Property Let Virtual(ByVal New_Virtual As Boolean)
    grdMain.Virtual() = New_Virtual
    PropertyChanged "Virtual"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,Header
Public Property Get Header() As Boolean
Attribute Header.VB_Description = "Gets/sets whether the grid has a header or not."
    Header = grdMain.Header
End Property

Public Property Let Header(ByVal New_Header As Boolean)
    grdMain.Header() = New_Header
    PropertyChanged "Header"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,HeaderDragReOrderColumns
Public Property Get HeaderDragReOrderColumns() As Boolean
Attribute HeaderDragReOrderColumns.VB_Description = "Gets/sets whether the grid's header columns can be dragged around to reorder them."
    HeaderDragReOrderColumns = grdMain.HeaderDragReOrderColumns
End Property

Public Property Let HeaderDragReOrderColumns(ByVal New_HeaderDragReOrderColumns As Boolean)
    grdMain.HeaderDragReOrderColumns() = New_HeaderDragReOrderColumns
    PropertyChanged "HeaderDragReOrderColumns"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,HeaderHeight
Public Property Get HeaderHeight() As Long
    HeaderHeight = grdMain.HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal New_HeaderHeight As Long)
    grdMain.HeaderHeight() = New_HeaderHeight
    PropertyChanged "HeaderHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,HeaderHotTrack
Public Property Get HeaderHotTrack() As Boolean
Attribute HeaderHotTrack.VB_Description = "Gets/sets whether the grid's header tracks mouse movements and highlights the header column the mouse is over or not."
    HeaderHotTrack = grdMain.HeaderHotTrack
End Property

Public Property Let HeaderHotTrack(ByVal New_HeaderHotTrack As Boolean)
    grdMain.HeaderHotTrack() = New_HeaderHotTrack
    PropertyChanged "HeaderHotTrack"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,HighlightBackColor
Public Property Get HighlightBackColor() As OLE_COLOR
    HighlightBackColor = grdMain.HighlightBackColor
End Property

Public Property Let HighlightBackColor(ByVal New_HighlightBackColor As OLE_COLOR)
    grdMain.HighlightBackColor() = New_HighlightBackColor
    PropertyChanged "HighlightBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,HighlightForeColor
Public Property Get HighlightForeColor() As OLE_COLOR
    HighlightForeColor = grdMain.HighlightForeColor
End Property

Public Property Let HighlightForeColor(ByVal New_HighlightForeColor As OLE_COLOR)
    grdMain.HighlightForeColor() = New_HighlightForeColor
    PropertyChanged "HighlightForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,MultiSelect
Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Gets/sets whether multiple grid cells or rows can be selected or not."
    MultiSelect = grdMain.MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    grdMain.MultiSelect() = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,RowMode
Public Property Get RowMode() As Boolean
Attribute RowMode.VB_Description = "Gets/sets whether cells can be selected in the grid (False) or rows (True)."
    RowMode = grdMain.RowMode
End Property

Public Property Let RowMode(ByVal New_RowMode As Boolean)
    grdMain.RowMode() = New_RowMode
    PropertyChanged "RowMode"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,BackgroundPicture
Public Property Get BackPicture() As Picture
Attribute BackPicture.VB_Description = "Gets/sets a picture to be used as the grid's background."
    Set BackPicture = grdMain.BackgroundPicture
End Property

Public Property Set BackPicture(ByVal New_BackPicture As Picture)
    Set grdMain.BackgroundPicture = New_BackPicture
    PropertyChanged "BackPicture"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=grdMain,grdMain,-1,DisableIcons
Public Property Get DrawDisabledIcons() As Boolean
Attribute DrawDisabledIcons.VB_Description = "Gets/sets whether icons are drawn disabled when the control is disabled."
    DrawDisabledIcons = grdMain.DisableIcons
End Property

Public Property Let DrawDisabledIcons(ByVal New_DrawDisabledIcons As Boolean)
    grdMain.DisableIcons() = New_DrawDisabledIcons
    PropertyChanged "DrawDisabledIcons"
End Property



'-----------------------------------------------------------------------------------------
'Read\Write\Init properties


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme

    m_Enabled = True

    ChangeTheme
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
    
    SetDefaultTheme

    grdMain.BackColor = PropBag.ReadProperty("BackColor", &H80000001)
    grdMain.DefaultRowHeight = PropBag.ReadProperty("DefaultRowHeight", 20)
    grdMain.DrawFocusRectangle = PropBag.ReadProperty("DrawFocusRectangle", True)
    grdMain.Editable = PropBag.ReadProperty("Editable", False)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    Set grdMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    grdMain.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    grdMain.GridLineColor = PropBag.ReadProperty("GridLineColor", -2147483633)
    grdMain.GridLines = PropBag.ReadProperty("GridLines", False)
    grdMain.Header = PropBag.ReadProperty("Header", True)
    grdMain.HeaderDragReOrderColumns = PropBag.ReadProperty("HeaderDragReOrderColumns", True)
    grdMain.HeaderHeight = PropBag.ReadProperty("HeaderHeight", 20)
    grdMain.HeaderHotTrack = PropBag.ReadProperty("HeaderHotTrack", True)
    grdMain.HighlightBackColor = PropBag.ReadProperty("HighlightBackColor", -2147483635)
    grdMain.HighlightForeColor = PropBag.ReadProperty("HighlightForeColor", -2147483634)
    grdMain.MultiSelect = PropBag.ReadProperty("MultiSelect", False)
    grdMain.RowMode = PropBag.ReadProperty("RowMode", False)
    Set grdMain.BackgroundPicture = PropBag.ReadProperty("BackPicture", Nothing)
    grdMain.DisableIcons = PropBag.ReadProperty("DrawDisabledIcons", True)
    grdMain.HeaderButtons = PropBag.ReadProperty("HeaderButtons", True)
    grdMain.HeaderFlat = PropBag.ReadProperty("HeaderFlat", False)
    grdMain.Virtual = PropBag.ReadProperty("Virtual", False)
    

    grdMain.Visible = True

    If glbControlsRefreshed = True Then
    RefreshTheme 'Refresh control because tilebar refresh code has already run before this
                 'This is needed to make sure the correct appearance is displayed
    End If
  

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("BackColor", grdMain.BackColor, &H80000001)
    Call PropBag.WriteProperty("DefaultRowHeight", grdMain.DefaultRowHeight, 20)
    Call PropBag.WriteProperty("DrawFocusRectangle", grdMain.DrawFocusRectangle, True)
    Call PropBag.WriteProperty("Editable", grdMain.Editable, False)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Font", grdMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", grdMain.ForeColor, &H80000008)
    Call PropBag.WriteProperty("GridLineColor", grdMain.GridLineColor, -2147483633)
    Call PropBag.WriteProperty("GridLines", grdMain.GridLines, False)
    Call PropBag.WriteProperty("Header", grdMain.Header, True)
    Call PropBag.WriteProperty("HeaderDragReOrderColumns", grdMain.HeaderDragReOrderColumns, True)
    Call PropBag.WriteProperty("HeaderHeight", grdMain.HeaderHeight, 20)
    Call PropBag.WriteProperty("HeaderHotTrack", grdMain.HeaderHotTrack, True)
    Call PropBag.WriteProperty("HighlightBackColor", grdMain.HighlightBackColor, -2147483635)
    Call PropBag.WriteProperty("HighlightForeColor", grdMain.HighlightForeColor, -2147483634)
    Call PropBag.WriteProperty("MultiSelect", grdMain.MultiSelect, False)
    Call PropBag.WriteProperty("RowMode", grdMain.RowMode, False)
    Call PropBag.WriteProperty("BackPicture", grdMain.BackgroundPicture, Nothing)
    Call PropBag.WriteProperty("DrawDisabledIcons", grdMain.DisableIcons, True)
    Call PropBag.WriteProperty("HeaderButtons", grdMain.HeaderButtons, True)
    Call PropBag.WriteProperty("HeaderFlat", grdMain.HeaderFlat, False)
    Call PropBag.WriteProperty("Virtual", grdMain.Virtual, False)

End Sub



'------------------------------------------------------------------------------
'Control Events:

Private Sub UserControl_Resize()

On Error GoTo err

grdMain.Width = Width - (6 * Screen.TwipsPerPixelX)
grdMain.Height = Height - (6 * Screen.TwipsPerPixelY)

UserControl.Cls
pSetBorder

err:
Exit Sub

End Sub


Private Sub grdMain_ColumnClick(ByVal lCol As Long)
    RaiseEvent ColumnClick(lCol)
End Sub

Private Sub grdMain_ColumnOrderChanged()
    RaiseEvent ColumnOrderChanged
End Sub

Private Sub grdMain_ColumnWidthChanged(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
    RaiseEvent ColumnWidthChanged(lCol, lWidth, bCancel)
End Sub

Private Sub grdMain_ColumnWidthChanging(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
    RaiseEvent ColumnWidthChanging(lCol, lWidth, bCancel)
End Sub

Private Sub grdMain_ColumnWidthStartChange(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
    RaiseEvent ColumnWidthStartChange(lCol, lWidth, bCancel)
End Sub

Private Sub grdMain_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    RaiseEvent DblClick(lRow, lCol)
End Sub

Private Sub grdMain_HeaderRightClick(ByVal X As Single, ByVal Y As Single)
    RaiseEvent HeaderRightClick(X, Y)
End Sub

Private Sub grdMain_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    RaiseEvent KeyDown(KeyCode, Shift, bDoDefault)
End Sub

Private Sub grdMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub grdMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub grdMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, bDoDefault As Boolean)
    RaiseEvent MouseDown(Button, Shift, X, Y, bDoDefault)
End Sub

Private Sub grdMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub grdMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub grdMain_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    RaiseEvent RequestEdit(lRow, lCol, iKeyAscii, bCancel)
End Sub

Private Sub grdMain_RequestRow(ByVal lRow As Long, ByVal sKey As String, ByVal bVisible As Boolean, ByVal lHeight As Long, ByVal bGroupRow As Boolean, bNoMoreRows As Boolean)
    RaiseEvent RequestRow(lRow, sKey, bVisible, lHeight, bGroupRow, bNoMoreRows)
End Sub

Private Sub grdMain_RequestRowData(ByVal lRow As Long)
    RaiseEvent RequestRowData(lRow)
End Sub

Private Sub grdMain_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    RaiseEvent SelectionChange(lRow, lCol)
End Sub



