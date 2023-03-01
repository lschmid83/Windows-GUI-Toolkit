VERSION 5.00
Object = "{72D18DD4-0DA7-11D2-8E21-00B404C10000}#2.3#0"; "VBALODCL.OCX"
Begin VB.UserControl BrowseFolderList 
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   ScaleHeight     =   3270
   ScaleWidth      =   4455
   ToolboxBitmap   =   "BrowseFolderList.ctx":0000
   Begin XPGUIControls10.MaskBox3 Right2 
      Align           =   4  'Align Right
      Height          =   3180
      Left            =   4410
      Top             =   45
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   4736
      ScaleHeight     =   3180
      ScaleWidth      =   45
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin XPGUIControls10.MaskBox3 Bottom 
      Align           =   2  'Align Bottom
      Height          =   45
      Left            =   0
      Top             =   3225
      Width           =   4455
      _ExtentX        =   7355
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   4455
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin XPGUIControls10.MaskBox3 left2 
      Align           =   3  'Align Left
      Height          =   3180
      Left            =   0
      Top             =   45
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   4736
      ScaleHeight     =   3180
      ScaleWidth      =   45
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin XPGUIControls10.MaskBox3 Top 
      Align           =   1  'Align Top
      Height          =   45
      Left            =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   4455
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin ODCboLst.OwnerDrawComboList lstMain 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4180
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
      Style           =   4
      FullRowSelect   =   -1  'True
      MaxLength       =   0
   End
   Begin VB.Label lblPath 
      Caption         =   "Label1"
      Height          =   570
      Left            =   105
      TabIndex        =   1
      Top             =   2565
      Width           =   2460
   End
End
Attribute VB_Name = "BrowseFolderList"
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


' API version of Dir/FileSystemObject.
' Harder to use but quicker...
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName(0 To MAX_PATH - 1) As Byte
   cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Enum EWin32FileAttributes
   FILE_ATTRIBUTE_ARCHIVE = &H20
   FILE_ATTRIBUTE_COMPRESSED = &H800
   FILE_ATTRIBUTE_DIRECTORY = &H10
   FILE_ATTRIBUTE_HIDDEN = &H2
   FILE_ATTRIBUTE_NORMAL = &H80
   FILE_ATTRIBUTE_READONLY = &H1
   FILE_ATTRIBUTE_SYSTEM = &H4
   FILE_ATTRIBUTE_TEMPORARY = &H100
End Enum

Private m_cIL As cSysImageList
Private m_bInterlock As Boolean

'Property Variables:
Dim m_LargeIcons As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_MouseIcon As Picture
Dim m_MousePointer As MousePointerConstants
Dim m_Locked As Boolean
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
'Default Property Values:
Const m_def_LargeIcons = 0



Private Sub LoadPath(ByVal sPath As String)
Dim hFiles As Long
Dim fd As WIN32_FIND_DATA
Dim sName As String
Dim f As Boolean
Dim bSkip As Boolean, bHaveFile As Boolean
Dim iC As Long

   If Not m_bInterlock Then
    '  chkLargeIcons.Enabled = False
      m_bInterlock = True
      Screen.MousePointer = vbArrowHourglass
      lstMain.Clear
   
      ' Get files in directory:
      hFiles = FindFirstFile(sPath & "*.*", fd)
      f = Not (hFiles = INVALID_HANDLE_VALUE)
      
      Do While f
         iC = iC + 1
         ' Add to list:
         sName = ByteZToStr(fd.cFileName)
         bSkip = False
         If (fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            If bHaveFile Then
               lstMain.InsertItemAndData sName, 0, m_cIL.ItemIndex(sPath & sName, True), , , , fd.dwFileAttributes, , m_cIL.IconSizeX, , eixVCentre
               bSkip = True
            End If
         Else
            bHaveFile = True
         End If
         If Not bSkip Then
            lstMain.AddItemAndData sName, m_cIL.ItemIndex(sPath & sName, True), , , , fd.dwFileAttributes, , m_cIL.IconSizeX, , eixVCentre
         End If
         
         ' ...
         If iC Mod 50 = 0 Then
            DoEvents
         End If
         ' Keep looping until no more files
          f = FindNextFile(hFiles, fd)
      Loop
      
      ' Done:
      f = FindClose(hFiles)
      Screen.MousePointer = vbDefault
      m_bInterlock = False
   '   chkLargeIcons.Enabled = True
   End If
   
   lstMain.Sorted = True
  
End Sub
Private Function ByteZToStr(ByRef b() As Byte) As String
Dim iPos As Long
Dim s As String
   s = StrConv(b, vbUnicode)
   iPos = InStr(s, vbNullChar)
   If iPos > 0 Then
      ByteZToStr = left$(s, iPos - 1)
   Else
      ByteZToStr = s
   End If
End Function
Private Sub SetPath(ByVal sPath As String)
   If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
   lblPath.Caption = sPath
   LoadPath sPath
End Sub


Private Sub lstmain_DblClick()
Dim eType As EWin32FileAttributes
Dim sFile As String
Dim iPos As Long
   
   If lstMain.ListIndex > -1 Then
      sFile = lblPath.Caption & lstMain.List(lstMain.ListIndex)
      eType = lstMain.ItemData(lstMain.ListIndex)
      If (eType And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
         If Not m_bInterlock Then
            Select Case lstMain.List(lstMain.ListIndex)
            Case "."
               ' refresh
               SetPath lblPath.Caption
            Case ".."
               ' Find previous directory
               sFile = lblPath.Caption
               ' InstrRev in VB6
               iPos = Len(sFile) - 1
               Do
                  If Mid$(sFile, iPos, 1) = "\" Then
                     sFile = left$(sFile, iPos - 1)
                     Exit Do
                  End If
                  iPos = iPos - 1
               Loop
               SetPath sFile
            Case Else
               ' New directory
               SetPath sFile
            End Select
         End If
      Else
         If vbYes = MsgBox("Start '" & sFile & "?", vbYesNo Or vbQuestion) Then
          '  ShellEx sFile
         End If
      End If
   End If
   
   lstMain.Sorted = True
   
End Sub




Public Sub Refresh()

lstMain.Top = 0
lstMain.left = 0

pSetEnabled
pSetBorder
lstMain.Sorted = True
pSetFolderView
lstMain.Sorted = True

UserControl_Resize

End Sub


Private Sub pSetFolderView()

   Set m_cIL = New cSysImageList
   m_cIL.IconSizeX = 16
   m_cIL.IconSizeY = 16
   m_cIL.Create
      
   lstMain.ImageList = m_cIL.hIml
         
   SetPath "C:\"


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

left2.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Left", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, left2.Height

Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\BottomLeft", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width, 3 * Screen.TwipsPerPixelY
Bottom.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY

Right2.PaintPicture PictureFromResource(ResourceLib.hModule, "ControlBorder\Enabled\Right", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Height
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
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""

    Refresh

    m_LargeIcons = m_def_LargeIcons
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
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    lstMain.NoDimWhenOutOfFocus = PropBag.ReadProperty("DimWhenOutOfFocus", False)
    lstMain.NoGrayWhenDisabled = PropBag.ReadProperty("GreyWhenDisabled", False)

    Refresh


    m_LargeIcons = PropBag.ReadProperty("LargeIcons", m_def_LargeIcons)
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
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("DimWhenOutOfFocus", lstMain.NoDimWhenOutOfFocus, False)
    Call PropBag.WriteProperty("GreyWhenDisabled", lstMain.NoGrayWhenDisabled, False)

    Call PropBag.WriteProperty("LargeIcons", m_LargeIcons, m_def_LargeIcons)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get LargeIcons() As Boolean
    LargeIcons = m_LargeIcons
End Property

Public Property Let LargeIcons(ByVal New_LargeIcons As Boolean)
    m_LargeIcons = New_LargeIcons
    PropertyChanged "LargeIcons"
End Property

