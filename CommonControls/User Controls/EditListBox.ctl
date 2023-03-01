VERSION 5.00
Object = "{72D18DD4-0DA7-11D2-8E21-00B404C10000}#2.3#0"; "VBALODCL.OCX"
Begin VB.UserControl EditList 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin XPGUIControls10.MaskBox3 Right 
      Align           =   4  'Align Right
      Height          =   3510
      Left            =   4755
      Top             =   45
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   6191
      ScaleHeight     =   3510
      ScaleWidth      =   45
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin XPGUIControls10.MaskBox3 Bottom 
      Align           =   2  'Align Bottom
      Height          =   45
      Left            =   0
      Top             =   3555
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   4800
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin XPGUIControls10.MaskBox3 left 
      Align           =   3  'Align Left
      Height          =   3510
      Left            =   0
      Top             =   45
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   6191
      ScaleHeight     =   3510
      ScaleWidth      =   45
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin XPGUIControls10.MaskBox3 Top 
      Align           =   1  'Align Top
      Height          =   45
      Left            =   0
      Top             =   0
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   4800
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
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
      Style           =   6
      MaxLength       =   0
   End
End
Attribute VB_Name = "EditList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Citex Software                                                       '
'XP GUI Toolkit v1.0 - StandardList Component v1.0                        '
'Copyright 2001 Lawrence Schmid                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


'Property Variables:
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_MouseIcon As Picture
Dim m_MousePointer As MousePointerConstants
Dim m_Locked As Boolean
Dim m_FolderView As Boolean
Dim m_DefaultStyle As ComboBoxDefaultStyleEnum
Dim m_DrawMode As ComboBoxDrawModeEnum
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String

Public Sub Refresh()

lstMain.Top = 0
lstMain.left = 0


pSetDefaultStyle
pSetDrawStyle
pSetEnabled
pSetBorder

UserControl_Resize

End Sub

Private Sub pSetDrawStyle()

If m_DefaultStyle = NoStyle Then 'Only allow DrawStyles to be set if a default style is not set.
                                 'This is because the if a default style is set it automatically sets the
                                 'draw style of the combobox
    Select Case m_DrawMode
    
    Case Is = dmStandard
    lstMain.ClientDraw = ecdClientDrawOnly
    
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
lstMain.ClientDraw = ecdSysColourPicker
m_DrawMode = dmColorPicker
LoadSysColorList lstMain

Case Is = ColorPickerNoNames
lstMain.ClientDraw = ecdColourPickerNoNames
m_DrawMode = dmColorPickerNoNames
LoadSysColorList lstMain

Case Is = FontPicker
lstMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
LoadFontList lstMain, "", -1, -1

Case Is = FontViewer
lstMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
LoadFontListViewer lstMain, "", -1, -1, True

Case Is = FontViewerNoIcons
lstMain.ClientDraw = ecdFontPicker
m_DrawMode = dmFontPicker
LoadFontListViewer lstMain, "", -1, -1, False

Case Is = ParagraphStyles
lstMain.Clear
lstMain.ClientDraw = ecdParagraphStyles
m_DrawMode = dmParagraphStyles
LoadParagraphStyles lstMain


End Select

End Sub


Private Sub pSetEnabled()

If Ambient.UserMode = False Then
UserControl.Enabled = False

    If m_Enabled = True Then
    lstMain.Visible = True
    
    Else
    lstMain.Visible = False
    
    End If


Else

    If m_Enabled = True Then
    UserControl.Enabled = True
    lstMain.Visible = True
    
        If Locked = True Then
        UserControl.Enabled = False
        End If
    
    Else
    UserControl.Enabled = False
    lstMain.Visible = False
    
    End If

End If



End Sub


Private Sub pSetBorder()

If ResourceLib.hModule <> 0 Then
Top.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\TopLeft", crBitmap), 0, 0
Top.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
Top.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY

left.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Left", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, left.Height

Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\BottomLeft", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width, 3 * Screen.TwipsPerPixelY
Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY

Right.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Right", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Height
End If

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
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
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
'MemberInfo=11,0,0,0
Public Property Get MouseIcon() As Picture
    Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set m_MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MousePointer() As Integer
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
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
'MemberInfo=0,0,0,0
Public Property Get FolderView() As Boolean
    FolderView = m_FolderView
End Property

Public Property Let FolderView(ByVal New_FolderView As Boolean)
    m_FolderView = New_FolderView
    PropertyChanged "FolderView"
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
'MemberInfo=14,0,0,0
Public Property Get ToolTipIcon() As Variant
    ToolTipIcon = m_ToolTipIcon
End Property

Public Property Let ToolTipIcon(ByVal New_ToolTipIcon As Variant)
    m_ToolTipIcon = New_ToolTipIcon
    PropertyChanged "ToolTipIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipStyle() As Variant
    ToolTipStyle = m_ToolTipStyle
End Property

Public Property Let ToolTipStyle(ByVal New_ToolTipStyle As Variant)
    m_ToolTipStyle = New_ToolTipStyle
    PropertyChanged "ToolTipStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipTitle() As String
    ToolTipTitle = m_ToolTipTitle
End Property

Public Property Let ToolTipTitle(ByVal New_ToolTipTitle As String)
    m_ToolTipTitle = New_ToolTipTitle
    PropertyChanged "ToolTipTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,NoDimWhenOutOfFocus
Public Property Get DimWhenOutOfFocus() As Boolean
    DimWhenOutOfFocus = lstMain.NoDimWhenOutOfFocus
End Property

Public Property Let DimWhenOutOfFocus(ByVal New_DimWhenOutOfFocus As Boolean)
    lstMain.NoDimWhenOutOfFocus() = New_DimWhenOutOfFocus
    PropertyChanged "DimWhenOutOfFocus"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lstMain,lstMain,-1,NoGrayWhenDisabled
Public Property Get GreyWhenDisabled() As Boolean
    GreyWhenDisabled = lstMain.NoGrayWhenDisabled
End Property

Public Property Let GreyWhenDisabled(ByVal New_GreyWhenDisabled As Boolean)
    lstMain.NoGrayWhenDisabled() = New_GreyWhenDisabled
    PropertyChanged "GreyWhenDisabled"
End Property


Private Sub UserControl_Resize()


lstMain.Width = Width
lstMain.Height = Height
pSetBorder

End Sub




'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
    
    m_BackColor = &HFFFFFF
    m_ForeColor = 0
    m_Enabled = True
    Set m_MouseIcon = LoadPicture("")
    m_MousePointer = 0
    m_Locked = False
    m_FolderView = False
    m_DefaultStyle = NoStyle
    m_DrawMode = dmStandard
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""

    Refresh

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)


    SetDefaultTheme

    lstMain.Columns = PropBag.ReadProperty("Columns", 1)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    Set lstMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set m_MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_MousePointer = PropBag.ReadProperty("MousePointer", 1)
    lstMain.Sorted = PropBag.ReadProperty("Sorted", False)
    lstMain.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    m_Locked = PropBag.ReadProperty("Locked", False)
    m_FolderView = PropBag.ReadProperty("FolderView", False)
    m_DefaultStyle = PropBag.ReadProperty("DefaultStyle", NoStyle)
    m_DrawMode = PropBag.ReadProperty("DrawMode", dmStandard)
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    lstMain.NoDimWhenOutOfFocus = PropBag.ReadProperty("DimWhenOutOfFocus", False)
    lstMain.NoGrayWhenDisabled = PropBag.ReadProperty("GreyWhenDisabled", False)

    Refresh


End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Columns", lstMain.Columns, 1)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Font", lstMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, 1)
    Call PropBag.WriteProperty("Sorted", lstMain.Sorted, False)
    Call PropBag.WriteProperty("FullRowSelect", lstMain.FullRowSelect, False)
    Call PropBag.WriteProperty("Locked", m_Locked, False)
    Call PropBag.WriteProperty("FolderView", m_FolderView, False)
    Call PropBag.WriteProperty("DefaultStyle", m_DefaultStyle, NoStyle)
    Call PropBag.WriteProperty("DrawMode", m_DrawMode, dmStandard)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("DimWhenOutOfFocus", lstMain.NoDimWhenOutOfFocus, False)
    Call PropBag.WriteProperty("GreyWhenDisabled", lstMain.NoGrayWhenDisabled, False)

End Sub

