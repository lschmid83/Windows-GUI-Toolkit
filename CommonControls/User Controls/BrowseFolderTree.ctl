VERSION 5.00
Begin VB.UserControl BrowseFolderTree 
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3315
   ScaleHeight     =   3240
   ScaleWidth      =   3315
   ToolboxBitmap   =   "BrowseFolderTree.ctx":0000
   Begin VB.PictureBox Picture1 
      Height          =   2925
      Left            =   120
      ScaleHeight     =   2865
      ScaleWidth      =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   2460
   End
End
Attribute VB_Name = "BrowseFolderTree"
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

Private Enum SHFolders
    CSIDL_PERSONAL = &H5
End Enum


Private Type WIN32_FIND_DATA
   dwFileAttributes  As Long
   ftCreationTime    As Currency   'was FILETIME
   ftLastAccessTime  As Currency   'Currency allows direct
   ftLastWriteTime   As Currency   'storage in Grid
   nFileSizeHigh     As String * 4 'Long
   nFileSizeLow      As String * 4 'Long
   dwReserved0       As Long
   dwReserved1       As Long
   cFileName         As String * 260
   cAlternate        As String * 14
End Type

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Private Const MAX_PATH = 260
Private buffer          As String * MAX_PATH
Private FvFilter     As Variant
Private m_MyDocs     As String
Const Desktop$ = "Desktop"


'Objects:
Private ilsicons As New cImageList

'Property Variables:
Dim m_Border As Boolean
Dim m_DefaultStyle As TreeViewDefaultStylesEnum

Private Sub pCreateExplorerDriveView()


   Const shell32$ = "Shell32.Dll"

    'Use API (MUCH faster than scripting)
'------------------------------
   Dim FirstFixed  As Integer
   Dim MaxPwr      As Integer
   Dim Pwr         As Integer
'------------------------------
   Dim DrvBitMask  As Long
   Dim DriveType   As Long
'------------------------------
   Dim MyDrive     As String
   Dim MyPic       As String
   Dim MyKey       As String
'------------------------------
  ' Dim nod1        As Node
   Dim sI          As SHFILEINFO
   Dim rc          As RECT
'------------------------------

Dim m_MyDocs As String


tvMain.Lines = True
tvMain.RootLines = True
tvMain.PlusMinus = True
'tvMain.ImageList = ImageList1
  

Dim Desktop As Long
Dim MyComputer As Long
Dim MyDrives(10) As Long

'Add Desktop
MyDrive = GetResourceStringFromFile(shell32, 4162) 'Desktop
Desktop = tvMain.Add(0&, AlphabeticalChild, "dt", "Desktop")  ', MyDrive)
     
     
'Add MyDocuments
MyDrive = GetResourceStringFromFile(shell32, 9100) 'My Documents
m_MyDocs = FolderLocation(CSIDL_PERSONAL)
tvMain.Add Desktop, FirstChild, "md", "My Documents"

'Add MyComputer
MyDrive = GetResourceStringFromFile(shell32, 9216) 'My Computer
MyComputer = tvMain.Add(Desktop, FirstChild, "mc", "MyComputer")


DrvBitMask = GetLogicalDrives()
' DrvBitMask is a bitmask representing
' available disk drives. Bit position 0
' is drive A, bit position 2 is drive C, etc.


    ' If function fails, return value is zero.
    If DrvBitMask Then
    ' Get & search each available drive
      MaxPwr = Int(Log(DrvBitMask) / Log(2))   ' a little math...
      For Pwr = 0 To MaxPwr
         If 2 ^ Pwr And DrvBitMask Then
            MyDrive = Chr$(65 + Pwr) & ":\"
            DriveType = GetDriveType(MyDrive)
            Select Case DriveType
               Case 0, 1: MyPic = "dl"
               Case 2:
                  If Pwr < 2 Then 'A or B (Diskette)
                     MyPic = "f35"
                  Else 'other Removable
                     MyPic = "rem"
                  End If
               Case 3: MyPic = "hd"
               Case 4: MyPic = "rte"
               Case 5: MyPic = "cd"
               Case 6: MyPic = "ram"
            End Select
            'Get Drive DisplayName.
            SHGetFileInfo MyDrive, 0&, sI, Len(sI), SHGFI_DISPLAYNAME
            MyDrives(Pwr) = tvMain.Add(MyComputer, AlphabeticalChild, sI.szDisplayName, sI.szDisplayName)
            
                       
            
            If (FirstFixed = 0) And (DriveType = 3) Then
               FirstFixed = tvMain.Count 'TV.Nodes.Count
            End If
            
            tvMain.Add MyDrives(Pwr), FirstChild, "Drive" & Pwr, ""
            
         End If
      Next
   End If
   

End Sub

Private Sub tvMain_ItemExpand(hItem As Long, ExpandType As xuiTreeView6.ExpandTypeConstants)

EnumFilesUnder hItem
MsgBox Right(tvMain.ItemKey(hItem), 2)




End Sub


Private Sub EnumFilesUnder(Item As Long)

   On Error GoTo PROC_ERR
    Dim sPath As String
    Dim sExt As String
    Dim hFind As Long, L4 As Long
    Dim OldPath As String
    Dim W32FD As WIN32_FIND_DATA
    'Dim n2 As Node
    Dim Folders As Long
    Dim FolderPic As String
    
   ' TV.Visible = False
    OldPath = ""
    sPath = BuildFullPath(9) & "*.*"
    'old sPath = ucase$(n.FullPath & "\*.*")
    hFind = FindFirstFile(sPath, W32FD)
    Do
        ' Get the filename, if any.
        sPath = StripNull(W32FD.cFileName)
        If Len(sPath) = 0 Or StrComp(sPath, OldPath) = 0 Then
            ' Nothing found?
            Exit Do
        ElseIf Asc(sPath) <> 46 Then
           'do we have a folder?
           If (W32FD.dwFileAttributes And vbDirectory) Then 'Yes
               FolderPic = "cl"
                             
               Folders = tvMain.Add(Item, AlphabeticalChild, sPath, sPath)
               
               
            '   Set n2 = TV.Nodes.Add(n, tvwChild, , sPath, FolderPic)
             '  n2.ExpandedImage = "op"
               'causes duplicate keys in My Documents
               'n2.Key = BuildFullPath(n2)
               ' Add a dummy item so the + sign is
               ' displayed
               If hasSubDirectory(BuildFullPath(5) & sPath & "\") Then
                  'TV.Nodes.Add n2, tvwChild
                    
               End If
           Else  'do we have a matching file?
              sExt = GetExt(sPath)
              For L4 = 0 To UBound(FvFilter)
                 If sPath Like FvFilter(L4) Then 'Yes
                    Select Case sExt
                       Case "zip", "cab", "ace", "rar"
                          FolderPic = sExt
                       Case Else
                          FolderPic = "new"
                    End Select
                   ' Set n2 = TV.Nodes.Add(n, tvwChild, , sPath, FolderPic)
                    'n2.Key = BuildFullPath(n2)
                    ' TV.Nodes.Item(TV.Nodes.Count).Bold = True
                    '***Node colors don't work if you are using
                    '   background (wallpaper) in Treeview
                '  TV.Nodes.Item(TV.Nodes.Count).BackColor = vbBlue '&H98CCD0   '&HE0E0E0    'grey
                  '  TV.Nodes.Item(TV.Nodes.Count).ForeColor = vbWhite    'RGB(248, 240, 136) 'Tree ylw
                    Exit For
                 End If
              Next
           End If
        End If
        FindNextFile hFind, W32FD
        OldPath = sPath
    Loop
    FindClose hFind
   ' TV.Visible = True
    Exit Sub

PROC_EXIT:
  Exit Sub
PROC_ERR:
'  If ErrMsgBox("EnumFilesUnder") = vbRetry Then Resume Next

End Sub

Private Function BuildFullPath(Item As Long) As String
   On Error GoTo PROC_ERR
   Dim iPos As Integer
   Dim sExt As String
   Dim MyPath As String
   Dim MyDocs2 As String
   MyPath = "C:\" 'Nod.FullPath
   
  ' tvmain.
   
   iPos = InStrRev(MyPath, ":")
   If iPos < 2 Then
     ' Select Case MyPath
   ' If nod.Key = QualifyPath(m_MyDocs) Then
       MyDocs2 = Mid(m_MyDocs, 4)
       BuildFullPath = Replace(MyPath, Desktop & "\" & MyDocs2, m_MyDocs)
       GoTo CheckExt
   ' End If
   End If
   MyPath = Mid$(MyPath, iPos - 1) 'Pick up drive letter

   iPos = InStr(MyPath, "\")
   If iPos > 1 Then
      BuildFullPath = left$(MyPath, 2) & Mid$(MyPath, iPos)
   Else
      BuildFullPath = left$(MyPath, 2)
   End If
CheckExt:
  ' sExt = GetExt(Nod.Text)
   If sExt <> "" Then
      For iPos = 0 To UBound(FvFilter)
         If sExt = GetExt(FvFilter(iPos)) Then 'Match
            Exit Function
         End If
      Next
   End If
   
   BuildFullPath = QualifyPath(BuildFullPath)

PROC_EXIT:
  Exit Function
PROC_ERR:
  'If ErrMsgBox("BuildFullPath") = vbRetry Then Resume Next

End Function

Private Function hasSubDirectory(ByVal sPath As String) As Boolean
    On Error GoTo PROC_ERR
    
    Dim hFind As Long
    Dim OldPath As String
    Dim W32FD As WIN32_FIND_DATA
    Dim L4 As Long
    
    OldPath = ""
    
    hFind = FindFirstFile(sPath & "*.*", W32FD)
    Do
        ' Get the filename, if any.
        sPath = StripNull(W32FD.cFileName)
        If Len(sPath) = 0 Or StrComp(sPath, OldPath) = 0 Then
            ' Nothing found?
            Exit Do
        ElseIf Asc(sPath) <> 46 Then
            ' return true if we have found a directory under this path
            If (W32FD.dwFileAttributes And vbDirectory) Then
                hasSubDirectory = True
                Exit Do
            End If
            For L4 = 0 To UBound(FvFilter)
                If sPath Like FvFilter(L4) Then
                  hasSubDirectory = True
                  Exit Do
                End If
            Next
        End If
        FindNextFile hFind, W32FD
        OldPath = sPath
    Loop
    FindClose hFind

PROC_EXIT:
  Exit Function
PROC_ERR:
 ' If ErrMsgBox("hasSubDirectory") = vbRetry Then Resume Next

End Function


Private Function QualifyPath(ByVal MyString As String) As String
   If Right$(MyString, 1) <> "\" Then
      QualifyPath = MyString & "\"
   Else
      QualifyPath = MyString
   End If
End Function
Private Function FolderLocation(lFolder As SHFolders) As String
Dim buffer As Long
   Dim lp As Long
   'Get the PIDL for this folder
  ' SHGetSpecialFolderLocation MyForm.hwnd, lFolder, lp
   SHGetSpecialFolderLocation 0&, lFolder, lp
   SHGetPathFromIDList lp, buffer
   FolderLocation = StripNull(buffer)
   'Free the PIDL
   CoTaskMemFree lp

End Function

Private Function StripNull(ByVal StrIn As String) As String
   On Error GoTo PROC_ERR
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         StripNull = left$(StrIn, nul - 1)
      Case 1
         StripNull = ""
      Case 0
         StripNull = Trim$(StrIn)
   End Select

PROC_EXIT:
  Exit Function
PROC_ERR:
 '; If ErrMsgBox("StripNull") = vbRetry Then Resume Next

End Function

Private Function GetExt(ByVal Name As String) As String
   On Error GoTo PROC_ERR
   Dim j As Integer
   j = InStrRev(Name, ".")
   If j > 0 And j < Len(Name) Then
      GetExt = LCase$(Mid$(Name, j + 1))
   End If

PROC_EXIT:
  Exit Function
PROC_ERR:
 ' If ErrMsgBox("GetExt") = vbRetry Then Resume Next

End Function

Private Sub pCreateOutlook()
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
   With tvMain
      .Clear
      '.ExplorerBar = tv.e
      '.RootLines = True
      '.Lines = True
      '.PlusMinus = True
      .FullRowSelect = False
      .SingleExpand = False
      .ShowNumber = True
      .InternalBorderX = 0
      .InternalBorderY = 0
      l = .Add(0&, FirstChild, "PERSONAL", "Personal Folders", ilsicons.ItemIndex("PERSONAL FOLDERS"))
      .Sorted(l) = False
      j = .Add(l, FirstChild, "CALENDAR", "Calendar", ilsicons.ItemIndex("CALENDAR"))
      j = .Add(j, NextSibling, "CONTACTS", "Contacts", ilsicons.ItemIndex("CONTACTS"))
      j = .Add(j, NextSibling, "DELETED", "Deleted Items", ilsicons.ItemIndex("DELETED ITEMS"))
      j = .Add(j, NextSibling, "DRAFTS", "Drafts", ilsicons.ItemIndex("DRAFTS"))
      .ItemBold(j) = True
      .ItemNumber(j) = 2
      j = .Add(j, NextSibling, "INBOX", "Inbox", ilsicons.ItemIndex("INBOX"))
      .ItemBold(j) = True
      .ItemNumber(j) = 1732
      For i = 1 To 10
         .Add j, LastChild, "FOLDER" & i, "Folder " & i, ilsicons.ItemIndex("MAILFOLDER")
      Next i
      .ItemExpanded(j) = True
      j = .Add(j, NextSibling, "JOURNAL", "Journal", ilsicons.ItemIndex("JOURNAL"))
      j = .Add(j, NextSibling, "NOTES", "Notes", ilsicons.ItemIndex("NOTES"))
      j = .Add(j, NextSibling, "OUTBOX", "Outbox", ilsicons.ItemIndex("OUTBOX"))
      j = .Add(j, NextSibling, "SENT ITEMS", "Send Items", ilsicons.ItemIndex("SENT ITEMS"))
      j = .Add(j, NextSibling, "TASKS", "Tasks", ilsicons.ItemIndex("TASKS"))
      .ItemExpanded(l) = True
   End With
End Sub

Public Sub Refresh()

pSetAppearance
UserControl_Resize

End Sub

Private Sub pSetDefaultStyle()

If m_DefaultStyle = OutlookFolderList Then

Set ilsicons = New cImageList
ilsicons.Create
ilsicons.ColourDepth = ILC_COLOR32
ilsicons.AddFromFile "C:\OutlookTreeViewIcons.bmp", IMAGE_BITMAP
tvMain.hImageList = ilsicons.hIml


pCreateOutlook


End If


End Sub



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
    BackColor = tvMain.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    tvMain.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=tvMain,tvMain,-1,CheckBoxes
Public Property Get CheckBoxes() As Boolean
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
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get DefaultStyle() As TreeViewDefaultStylesEnum
    DefaultStyle = m_DefaultStyle
End Property

Public Property Let DefaultStyle(ByVal New_DefaultStyle As TreeViewDefaultStylesEnum)
    m_DefaultStyle = New_DefaultStyle
    PropertyChanged "DefaultStyle"
End Property









Private Sub UserControl_Resize()

If glbAppearance <> Win98 Then
tvMain.Width = Width - (6 * Screen.TwipsPerPixelX)
tvMain.Height = Height - (6 * Screen.TwipsPerPixelY)

UserControl.Cls
pSetBorder

Else
tvMain.Width = Width - (4 * Screen.TwipsPerPixelX)
tvMain.Height = Height - (4 * Screen.TwipsPerPixelY)

End If


End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
    
    m_Border = True
    m_DefaultStyle = None
    
    
    Refresh
    'pSetDefaultStyle
pCreateExplorerDriveView
    
    
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
    
    If Ambient.UserMode = False Then
    UserControl.Enabled = False
    Else
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    End If
    
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
    m_DefaultStyle = PropBag.ReadProperty("DefaultStyle", None)


    Refresh
pCreateExplorerDriveView

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", tvMain.BackColor, &H80000005)
    Call PropBag.WriteProperty("CheckBoxes", tvMain.CheckBoxes, False)
    Call PropBag.WriteProperty("DragExpandTime", tvMain.DragExpandTime, 2000)
    Call PropBag.WriteProperty("DragScrollTime", tvMain.DragScrollTime, 500)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
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
    Call PropBag.WriteProperty("DefaultStyle", m_DefaultStyle, None)

End Sub


