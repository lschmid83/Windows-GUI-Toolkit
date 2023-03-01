Attribute VB_Name = "mDefaultStyles"
'Contains Code to automatically setup a combo, list or gridbox
'to a specific style


Option Explicit


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dw As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, ByVal dwAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, lpBuffer As Any) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long



Dim m_cCbo As Object
Private m_sDriveStrings As String
Private m_iType As Long
Private m_iCharSet As Long

Private Const RASTER_FONTTYPE = 1&
Private Const DEVICE_FONTTYPE = 2&
Private Const TRUETYPE_FONTTYPE = 4&

Private Const MAX_PATH = 260
Public Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const ANSI_CHARSET = 0

Private Enum EShellGetFileInfoConstants
   SHGFI_ICON = &H100                       ' // get icon
   SHGFI_DISPLAYNAME = &H200                ' // get display name
   SHGFI_TYPENAME = &H400                   ' // get type name
   SHGFI_ATTRIBUTES = &H800                 ' // get attributes
   SHGFI_ICONLOCATION = &H1000              ' // get icon location
   SHGFI_EXETYPE = &H2000                   ' // return exe type
   SHGFI_SYSICONINDEX = &H4000              ' // get system icon index
   SHGFI_LINKOVERLAY = &H8000               ' // put a link overlay on icon
   SHGFI_SELECTED = &H10000                 ' // show icon in selected state
   SHGFI_ATTR_SPECIFIED = &H20000           ' // get only specified attributes
   SHGFI_LARGEICON = &H0                    ' // get large icon
   SHGFI_SMALLICON = &H1                    ' // get small icon
   SHGFI_OPENICON = &H2                     ' // get open icon
   SHGFI_SHELLICONSIZE = &H4                ' // get shell size icon
   SHGFI_PIDL = &H8                         ' // pszPath is a pidl
   SHGFI_USEFILEATTRIBUTES = &H10           ' // use passed dwFileAttribute
End Enum


Public Enum EDriveType
   DRIVE_REMOVABLE = 2
   DRIVE_FIXED = 3
   DRIVE_REMOTE = 4
   DRIVE_CDROM = 5
   DRIVE_RAMDISK = 6
End Enum

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type ENUMLOGFONTEX
    elfLogFont As LOGFONT
    elfFullName(LF_FULLFACESIZE - 1) As Byte
    elfStyle(LF_FACESIZE - 1) As Byte
    elfScript(LF_FACESIZE - 1) As Byte
End Type


Public Sub LoadParagraphStyles(ByRef cbo As Object)
    Dim sFnt As New StdFont, lHeight As Long
    lHeight = 32
    With cbo
        
        
        sFnt.Name = "Arial"
        sFnt.Size = 14
        sFnt.Bold = True
        sFnt.Italic = False
        
        .AddItemAndData "Heading 1", , 8, , , sFnt.Size, , lHeight, , eixVCentre, sFnt
        sFnt.Name = "Arial"
        sFnt.Size = 12
        sFnt.Bold = False
        sFnt.Italic = True
  
        .AddItemAndData "Heading 2", , 8, , , sFnt.Size, , lHeight, , eixVCentre, sFnt
        sFnt.Name = "Arial"
        sFnt.Size = 10
        sFnt.Bold = True
        sFnt.Italic = False

        .AddItemAndData "Heading 3", , 8, , , sFnt.Size, , lHeight, , eixVCentre, sFnt
        sFnt.Name = "Times New Roman"
        sFnt.Size = 10
        sFnt.Bold = False
        sFnt.Italic = False
   
        
        sFnt.Name = "Courier New"
        sFnt.Size = 8
        sFnt.Bold = False
        sFnt.Italic = False
        
        
        .AddItemAndData "Normal", , 8, , , sFnt.Size, , lHeight, , eixVCentre, sFnt
        
        .AddItemAndData "Centred", , 8, , , sFnt.Size, , lHeight, eixCentre, eixVCentre, sFnt
    
    
        On Error GoTo err
       .ListIndex = 1
    End With

err:
Exit Sub

End Sub


Public Sub LoadSysColorList(ByRef cbo As Object)
      'assign system color names
   With cbo
      .Clear
      .AddItemAndData "3DDKShadow", , , , vb3DDKShadow
      .AddItemAndData "3DFace", , , , vb3DFace
      .AddItemAndData "3DHighlight", , , , vb3DHighlight
      .AddItemAndData "3DLight", , , , vb3DLight
      .AddItemAndData "3DShadow", , , , vb3DShadow
      .AddItemAndData "ActiveBorder", , , , vbActiveBorder
      .AddItemAndData "ActiveTitleBar", , , , vbActiveTitleBar
      .AddItemAndData "ApplicationWorkspace", , , , vbApplicationWorkspace
      .AddItemAndData "ButtonFace", , , , vbButtonFace
      .AddItemAndData "ButtonShadow", , , , vbButtonShadow
      .AddItemAndData "ButtonText", , , , vbButtonText
      .AddItemAndData "Desktop", , , , vbDesktop
      .AddItemAndData "GrayText", , , , vbGrayText
      .AddItemAndData "Highlight", , , , vbHighlight
      .AddItemAndData "HighlightText", , , , vbHighlightText
      .AddItemAndData "InactiveBorder", , , , vbInactiveBorder
      .AddItemAndData "InactiveCaptionText", , , , vbInactiveCaptionText
      .AddItemAndData "InactiveTitleBar", , , , vbInactiveTitleBar
      .AddItemAndData "InfoBackground", , , , vbInfoBackground
      .AddItemAndData "InfoText", , , , vbInfoText
      .AddItemAndData "MenuBar", , , , vbMenuBar
      .AddItemAndData "MenuText", , , , vbMenuText
      .AddItemAndData "ScrollBars", , , , vbScrollBars
      .AddItemAndData "TitleBarText", , , , vbTitleBarText
      .AddItemAndData "WindowBackground", , , , vbWindowBackground
      .AddItemAndData "WindowFrame", , , , vbWindowFrame
      .AddItemAndData "WindowText", , , , vbWindowText
      
      On Error GoTo err
      .ListIndex = 0
   End With


err:
Exit Sub
End Sub



Public Function LoadFontList(ByRef cbo As Object, ByVal sFace As String, ByVal iType As Long, ByVal lCharSet As Long) As Long
Dim tLF As LOGFONT
Dim i As Integer
Dim lHDC As Long

   ' No re-entrancy..
   If m_cCbo Is Nothing Then
      Set m_cCbo = cbo
      cbo.Clear
    '  cbo.Redraw = False
      
      ' Set up to load the fonts:
      m_iType = iType
      m_iCharSet = lCharSet
      ' Convert the face name into a byte array:
      If Len(sFace) > 0 Then
         For i = 1 To Len(sFace)
            tLF.lfFaceName(i - 1) = Asc(Mid$(sFace, i, 1))
         Next i
      End If
      If lCharSet <= 0 Then lCharSet = ANSI_CHARSET
      tLF.lfCharSet = lCharSet
      
    '  InitSystemImageList cbo, False
      cbo.Sorted = True
'      cbo.ExtendedStyle(eccxCaseSensitiveSearch) = False
      ' Start the enumeration:
      lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      LoadFontList = EnumFontFamiliesEx(lHDC, tLF, AddressOf EnumFontFamExProc, (Len(sFace) > 0), 0)
      DeleteDC lHDC
      
   '   cbo.Redraw = True
      
      ' Clear up reference to the caller:
      Set m_cCbo = Nothing
   End If
   
   On Error GoTo err
   cbo.ListIndex = 0
   
err:
Exit Function

   
End Function

Public Function LoadFontListViewer(ByRef cbo As Object, ByVal sFace As String, ByVal iType As Long, ByVal lCharSet As Long, ViewIcons As Boolean) As Long
Dim tLF As LOGFONT
Dim i As Integer
Dim lHDC As Long

   ' No re-entrancy..
   If m_cCbo Is Nothing Then
      Set m_cCbo = cbo
      cbo.Clear
    '  cbo.Redraw = False
      
      ' Set up to load the fonts:
      m_iType = iType
      m_iCharSet = lCharSet
      ' Convert the face name into a byte array:
      If Len(sFace) > 0 Then
         For i = 1 To Len(sFace)
            tLF.lfFaceName(i - 1) = Asc(Mid$(sFace, i, 1))
         Next i
      End If
      If lCharSet <= 0 Then lCharSet = ANSI_CHARSET
      tLF.lfCharSet = lCharSet
      
    '  InitSystemImageList cbo, False
      cbo.Sorted = True
'      cbo.ExtendedStyle(eccxCaseSensitiveSearch) = False
      ' Start the enumeration:
      lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      
      If ViewIcons = True Then
      LoadFontListViewer = EnumFontFamiliesEx(lHDC, tLF, AddressOf EnumFontFamExProc2, (Len(sFace) > 0), 0)
      Else
      LoadFontListViewer = EnumFontFamiliesEx(lHDC, tLF, AddressOf EnumFontFamExProc3, (Len(sFace) > 0), 0)
      End If
      
      DeleteDC lHDC
      
   '   cbo.Redraw = True
      
      ' Clear up reference to the caller:
      Set m_cCbo = Nothing
   End If
   
   On Error GoTo err
   cbo.ListIndex = 0
   
   
err:
Exit Function

End Function


Public Function EnumFontFamExProc(ByVal lpelfe As Long, ByVal lpntme As Long, ByVal iFontType As Long, ByVal lParam As Long) As Long
' The callback function for EnumFontFamiliesEx.

' lpelf points to an ENUMLOGFONTEX structure, lpntm points to either
' a NEWTEXTMETRICEX (if true type) or a TEXTMETRIC (non-true type)
' structure.

Dim tLFEx As ENUMLOGFONTEX
Dim sFace As String, sScript As String
Dim sStyle As String, sFullName As String
Dim lPos As Long
Dim sItem As String
Dim iIcon As Long
Dim FileInfo As SHFILEINFO
    
   CopyMemory tLFEx, ByVal lpelfe, LenB(tLFEx) ' Get the ENUMLOGFONTEX info
   ' Face Name
   sFace = StrConv(tLFEx.elfLogFont.lfFaceName, vbUnicode)
   lPos = InStr(sFace, Chr$(0))
   If (lPos > 0) Then sFace = left$(sFace, (lPos - 1))
    
   ' Script
   sScript = StrConv(tLFEx.elfScript, vbUnicode)
   lPos = InStr(sScript, Chr$(0))
   If (lPos > 0) Then sScript = left$(sScript, (lPos - 1))
    
   ' mbShowStyle
   If lParam = True Then
      ' Style
      sStyle = StrConv(tLFEx.elfStyle, vbUnicode)
      lPos = InStr(sStyle, Chr$(0))
      If (lPos > 0) Then sStyle = left$(sStyle, (lPos - 1))
   Else
      sStyle = ""
   End If
    
   ' Full Name
   sFullName = StrConv(tLFEx.elfFullName, vbUnicode)
   lPos = InStr(sFullName, Chr$(0))
   If (lPos > 0) Then sFullName = left$(sFullName, (lPos - 1))
    
   ' Only display printer and true type fonts:
   If (m_iType > 0) Then
      If (iFontType And m_iType) <> m_iType Then
         EnumFontFamExProc = 1
         Exit Function
      End If
   End If
    
   ' Only display a given font once:
   If m_cCbo.FindItemIndex(sFace, True) = -1 Then
       'm_cSink.AddFont sFace, sStyle, sScript, tLFEx.elfLogFont.lfCharSet, m_bPrinterFont
           If (iFontType And TRUETYPE_FONTTYPE) = TRUETYPE_FONTTYPE Then
        ' m_cCbo.AddItemAndData sFace, 0, 1
        FileInfo.iIcon = 0
        ' SHGetFileInfo "X.TTF", FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICOn
      ElseIf (iFontType And RASTER_FONTTYPE) = RASTER_FONTTYPE Then
          'm_cCbo.AddItemAndData sFace, 1, 1
          FileInfo.iIcon = 1
         'SHGetFileInfo "X.FON", FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICOn
      Else
         'FileInfo.iIcon = -1
      End If
      m_cCbo.AddItemAndData sFace, FileInfo.iIcon, FileInfo.iIcon
   
    
   End If
   ' Ask for more fonts:
   EnumFontFamExProc = 1
    
End Function

Public Function EnumFontFamExProc2(ByVal lpelfe As Long, ByVal lpntme As Long, ByVal iFontType As Long, ByVal lParam As Long) As Long
' The callback function for EnumFontFamiliesEx.

' lpelf points to an ENUMLOGFONTEX structure, lpntm points to either
' a NEWTEXTMETRICEX (if true type) or a TEXTMETRIC (non-true type)
' structure.

Dim FontToAdd As StdFont
Dim tLFEx As ENUMLOGFONTEX
Dim sFace As String, sScript As String
Dim sStyle As String, sFullName As String
Dim lPos As Long
Dim sItem As String
Dim iIcon As Long
Dim FileInfo As SHFILEINFO
    
    Set FontToAdd = New StdFont
    
   CopyMemory tLFEx, ByVal lpelfe, LenB(tLFEx) ' Get the ENUMLOGFONTEX info
   ' Face Name
   sFace = StrConv(tLFEx.elfLogFont.lfFaceName, vbUnicode)
   lPos = InStr(sFace, Chr$(0))
   If (lPos > 0) Then sFace = left$(sFace, (lPos - 1))
    
   ' Script
   sScript = StrConv(tLFEx.elfScript, vbUnicode)
   lPos = InStr(sScript, Chr$(0))
   If (lPos > 0) Then sScript = left$(sScript, (lPos - 1))
    
   ' mbShowStyle
   If lParam = True Then
      ' Style
      sStyle = StrConv(tLFEx.elfStyle, vbUnicode)
      lPos = InStr(sStyle, Chr$(0))
      If (lPos > 0) Then sStyle = left$(sStyle, (lPos - 1))
   Else
      sStyle = ""
   End If
    
   ' Full Name
   sFullName = StrConv(tLFEx.elfFullName, vbUnicode)
   lPos = InStr(sFullName, Chr$(0))
   If (lPos > 0) Then sFullName = left$(sFullName, (lPos - 1))
    
   ' Only display printer and true type fonts:
   If (m_iType > 0) Then
      If (iFontType And m_iType) <> m_iType Then
         EnumFontFamExProc2 = 1
         Exit Function
      End If
   End If
    
   ' Only display a given font once:
   If m_cCbo.FindItemIndex(sFace, True) = -1 Then
       'm_cSink.AddFont sFace, sStyle, sScript, tLFEx.elfLogFont.lfCharSet, m_bPrinterFont
           If (iFontType And TRUETYPE_FONTTYPE) = TRUETYPE_FONTTYPE Then
        ' m_cCbo.AddItemAndData sFace, 0, 1
        FileInfo.iIcon = 0
        ' SHGetFileInfo "X.TTF", FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICOn
      ElseIf (iFontType And RASTER_FONTTYPE) = RASTER_FONTTYPE Then
          'm_cCbo.AddItemAndData sFace, 1, 1
          FileInfo.iIcon = 1
         'SHGetFileInfo "X.FON", FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICOn
      Else
         'FileInfo.iIcon = -1
      End If
         
         FontToAdd.Name = sFace
         m_cCbo.AddItemAndData sFace, FileInfo.iIcon, FileInfo.iIcon, , , , , , , , FontToAdd

    
   End If
   ' Ask for more fonts:
   EnumFontFamExProc2 = 1
    
End Function


Public Function EnumFontFamExProc3(ByVal lpelfe As Long, ByVal lpntme As Long, ByVal iFontType As Long, ByVal lParam As Long) As Long
' The callback function for EnumFontFamiliesEx.

' lpelf points to an ENUMLOGFONTEX structure, lpntm points to either
' a NEWTEXTMETRICEX (if true type) or a TEXTMETRIC (non-true type)
' structure.

Dim FontToAdd As StdFont
Dim tLFEx As ENUMLOGFONTEX
Dim sFace As String, sScript As String
Dim sStyle As String, sFullName As String
Dim lPos As Long
Dim sItem As String
Dim iIcon As Long
Dim FileInfo As SHFILEINFO
    
    Set FontToAdd = New StdFont
    
   CopyMemory tLFEx, ByVal lpelfe, LenB(tLFEx) ' Get the ENUMLOGFONTEX info
   ' Face Name
   sFace = StrConv(tLFEx.elfLogFont.lfFaceName, vbUnicode)
   lPos = InStr(sFace, Chr$(0))
   If (lPos > 0) Then sFace = left$(sFace, (lPos - 1))
    
   ' Script
   sScript = StrConv(tLFEx.elfScript, vbUnicode)
   lPos = InStr(sScript, Chr$(0))
   If (lPos > 0) Then sScript = left$(sScript, (lPos - 1))
    
   ' mbShowStyle
   If lParam = True Then
      ' Style
      sStyle = StrConv(tLFEx.elfStyle, vbUnicode)
      lPos = InStr(sStyle, Chr$(0))
      If (lPos > 0) Then sStyle = left$(sStyle, (lPos - 1))
   Else
      sStyle = ""
   End If
    
   ' Full Name
   sFullName = StrConv(tLFEx.elfFullName, vbUnicode)
   lPos = InStr(sFullName, Chr$(0))
   If (lPos > 0) Then sFullName = left$(sFullName, (lPos - 1))
    
   ' Only display printer and true type fonts:
   If (m_iType > 0) Then
      If (iFontType And m_iType) <> m_iType Then
         EnumFontFamExProc3 = 1
         Exit Function
      End If
   End If
    
   ' Only display a given font once:
   If m_cCbo.FindItemIndex(sFace, True) = -1 Then
       'm_cSink.AddFont sFace, sStyle, sScript, tLFEx.elfLogFont.lfCharSet, m_bPrinterFont
           If (iFontType And TRUETYPE_FONTTYPE) = TRUETYPE_FONTTYPE Then
        ' m_cCbo.AddItemAndData sFace, 0, 1
        FileInfo.iIcon = 0
        ' SHGetFileInfo "X.TTF", FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICOn
      ElseIf (iFontType And RASTER_FONTTYPE) = RASTER_FONTTYPE Then
          'm_cCbo.AddItemAndData sFace, 1, 1
          FileInfo.iIcon = 1
         'SHGetFileInfo "X.FON", FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES Or SHGFI_SMALLICOn
      Else
         'FileInfo.iIcon = -1
      End If
         
         FontToAdd.Name = sFace
         m_cCbo.AddItemAndData sFace, , , , , , , , , , FontToAdd

    
   End If
   ' Ask for more fonts:
   EnumFontFamExProc3 = 1
    
End Function

Public Sub LoadDriveList(ByVal cbo As Object, ByVal bLargeIcons As Boolean)
'// ==========================================================================
'// Load Items - collects all drive information and place it into the listbox,
'// return number of items added to the list: a negative value is an error;
'// ==========================================================================
Dim lAllDriveStrings As Long
Dim sDrive As String
Dim lR As Long
Dim dwIconSize As Long
Dim FileInfo As SHFILEINFO
Dim iPos As Long, iLastPos As Long
Dim iType As EDriveType
Dim hIml As Long
Dim dwFlags As Long
Dim lDefIndex As Long

   cbo.Clear
 
   '// allocate buffer for the drive strings: GetLogicalDriveStrings will tell
   '// me how much is needed (minus the trailing zero-byte)
   lAllDriveStrings = GetLogicalDriveStrings(0, ByVal 0&)

   m_sDriveStrings = String$(lAllDriveStrings + 1, 0) 'new _TCHAR[ lAllDriveStrings + sizeof( _T("")) ]; // + for trailer
   lR = GetLogicalDriveStrings(lAllDriveStrings, ByVal m_sDriveStrings)
   Debug.Assert lR = (lAllDriveStrings - 1)
  
  InitSystemImageList cbo, bLargeIcons
  
   '// now loop over each drive (string)
   If bLargeIcons Then
      dwIconSize = SHGFI_LARGEICON
   Else
      dwIconSize = SHGFI_SMALLICON
   End If
   
   iLastPos = 1
   Do
      iPos = InStr(iLastPos, m_sDriveStrings, vbNullChar)
      
      If iPos <> 0 Then
         sDrive = Mid$(m_sDriveStrings, iLastPos, iPos - iLastPos)
         iLastPos = iPos + 1
      Else
         sDrive = Mid$(m_sDriveStrings, iLastPos)
      End If
      If Not sDrive = vbNullString Then
         lR = SHGetFileInfo(sDrive, FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX Or dwIconSize)
         If (lR = 0) Then  '// failure - which can be ignored
            Debug.Print "SHGetFileInfo failed, no more details available"
         Else
            
            '// insert icon and string into list box
            cbo.AddItemAndData FileInfo.szDisplayName, FileInfo.iIcon, FileInfo.iIcon, 0
            If lDefIndex = 0 Then
               iType = GetDriveType(left$(sDrive, 2))
               If iType = 1 Or iType = DRIVE_FIXED Then
                  lDefIndex = cbo.NewIndex
               End If
            End If
         End If
         cbo.ListIndex = lDefIndex
      Else
         iPos = 0
      End If
   Loop While iPos <> 0
   
End Sub
Public Sub InitSystemImageList(ByRef cbo As Object, ByVal bLargeIcons As Boolean)
Dim dwFlags As Long
Dim hIml As Long
Dim FileInfo As SHFILEINFO

   dwFlags = SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX
   If Not (bLargeIcons) Then
      dwFlags = dwFlags Or SHGFI_SMALLICON
   End If
   '// Load the image list - use an arbitrary file extension for the
   '// call to SHGetFileInfo (we don't want to touch the disk, so use
   '// FILE_ATTRIBUTE_NORMAL && SHGFI_USEFILEATTRIBUTES).
   hIml = SHGetFileInfo(".txt", FILE_ATTRIBUTE_NORMAL, FileInfo, LenB(FileInfo), dwFlags)
       
   ' MFC code sample says to do this, but this looks dubious to me.  Likely
   ' you will disrupt Explorer in Win9x...
   '// Make the background colour transparent, works better for lists etc.
   'ImageList_SetBkColor m_hIml, CLR_NONE
   
   cbo.ImageList = hIml

End Sub









