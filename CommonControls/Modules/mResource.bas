Attribute VB_Name = "mResource"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - mResource.bas                                 '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal uID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibModule As Long) As Long


Private Declare Function EnumResourceLanguages Lib "KERNEL32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceNamesByNum Lib "KERNEL32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceNamesByAny Lib "KERNEL32" Alias "EnumResourceNamesA" (ByVal hModule As Long, lpType As Any, ByVal lpEnumFunc As String, ByVal lParam As Long) As Long
Private Declare Function EnumResourceTypes Lib "KERNEL32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function lstrlen Lib "KERNEL32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_FIXED = &H0
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cX As Long, ByVal cY As Long, ByVal uFlags As Long) As Long
Private Declare Function LoadImageString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cX As Long, ByVal cY As Long, ByVal uFlags As Long) As Long
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2

Private Declare Function LoadStringAsAny Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, wID As Any, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LoadStringWAsAny Lib "user32" Alias "LoadStringW" (ByVal hInstance As Long, wID As Any, lpBuffer As Any, ByVal nBufferMax As Long) As Long

Private Declare Function FindResource Lib "KERNEL32" Alias "FindResourceA" (ByVal hInstance As Long, lpName As Any, lpType As Any) As Long
Private Declare Function LockResource Lib "KERNEL32" (ByVal hResData As Long) As Long
Private Declare Function LoadResource Lib "KERNEL32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function SizeofResource Lib "KERNEL32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "KERNEL32" (ByVal hResData As Long) As Long

Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_MONOCHROME = &H1
Private Const LR_COLOR = &H2
Private Const LR_COPYRETURNORG = &H4
Private Const LR_COPYDELETEORG = &H8
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000&
Private Const LR_CREATEDIBSECTION = &H2000&
Private Const LR_COPYFROMRESOURCE = &H4000&
Private Const LR_SHARED = &H8000&

Private m_cR As cResources


Public Function GetResourceStringFromFile(sModule As String, idString As Long) As String

   Dim hModule As Long
   Dim nChars As Long
Dim buffer As Long
   hModule = LoadLibrary(sModule)
   If hModule Then
      nChars = LoadString(hModule, idString, buffer, MAX_PATH)
      If nChars Then
         GetResourceStringFromFile = Left$(buffer, nChars)
      End If
      FreeLibrary hModule
   End If
End Function

Public Property Get PictureFromResource( _
      ByVal hMod As Long, _
      ByVal sName As String, _
      ByVal eType As CRStandardResourceTypeConstants _
   ) As IPicture
Dim hBmp As Long
Dim hIcon As Long
Dim hCur As Long
Dim lErr As Long
Dim lID As Long

   If eType = crBitmap Then
      If IsNumeric(sName) Then
         lID = CLng(sName)
         hBmp = LoadImageLong(hMod, lID, IMAGE_BITMAP, 0, 0, 0)
      Else
         hBmp = LoadImageString(hMod, sName, IMAGE_BITMAP, 0, 0, 0)
      End If
      If Not hBmp = 0 Then
         Set PictureFromResource = BitmapToPicture(hBmp)
      Else
         'lErr = Err.LastDllError
         'Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cResource", WinError(lErr)
      End If
   ElseIf eType = crGroupIcon Then
      If IsNumeric(sName) Then
         lID = CLng(sName)
         hIcon = LoadImageLong(hMod, lID, IMAGE_ICON, 0, 0, 0)
      Else
         hIcon = LoadImageString(hMod, sName, IMAGE_ICON, 0, 0, 0)
      End If
      If Not hIcon = 0 Then
         Set PictureFromResource = IconToPicture(hIcon)
      Else
        ' lErr = Err.LastDllError
        ' Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cResource", WinError(lErr)
      End If
   ElseIf eType = crGroupCursor Then
      If IsNumeric(sName) Then
         lID = CLng(sName)
         hCur = LoadImageLong(hMod, lID, IMAGE_CURSOR, 0, 0, 0)
      Else
         hCur = LoadImageString(hMod, sName, IMAGE_CURSOR, 0, 0, 0)
      End If
      If Not hCur = 0 Then
         Set PictureFromResource = IconToPicture(hCur)
      Else
        ' lErr = Err.LastDllError
        ' Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cResource", WinError(lErr)
      End If
      
   End If
End Property

Public Sub SaveResource( _
      ByVal hMod As Long, _
      ByVal sName As String, _
      ByVal sType As String, _
      ByVal sFile As String _
   )
Dim sBuf As String
Dim hGbl As Long
Dim hRes As Long
Dim lID As Long
Dim lSize As Long
Dim lPtr As Long
Dim lR As Long
Dim iFile As Integer
Dim lErr As Long
   
On Error GoTo ErrorHandler
   
   If IsNumeric(sName) Then
      lID = CLng(sName)
      sName = "#" & sName
   End If
   If IsNumeric(sType) Then
      hRes = FindResource(hMod, ByVal sName, ByVal CLng(sType))
   Else
      hRes = FindResource(hMod, ByVal sName, ByVal sType)
   End If
   If Not hRes = 0 Then
      hGbl = LoadResource(hMod, hRes)
      If Not hGbl = 0 Then
         lPtr = LockResource(hGbl)
         If Not lPtr = 0 Then
            lSize = SizeofResource(hMod, hRes)
            If lSize > 0 Then
               ReDim b(0 To lSize) As Byte
               CopyMemory b(0), ByVal lPtr, lSize
               
               On Error Resume Next
                  Kill sFile
               On Error GoTo ErrorHandler
               
               iFile = FreeFile
               Open sFile For Binary Access Write Lock Read As #iFile
               Put #iFile, , b
               Close #iFile
               iFile = 0
                                             
            End If
         Else
           ' lErr = Err.LastDllError
            'Err.Raise vbObjectError + 1048 + 4, App.EXEName & ".cResource", WinError(lErr)
         End If
      Else
         'lErr = Err.LastDllError
         'Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cResource", WinError(lErr)
      End If
      FreeResource hRes
   Else
      'Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cResource", "Specified Resource Not Found."
   End If
   Exit Sub
      
ErrorHandler:
   err.Raise err.Number, App.EXEName & ".cResource", err.Description
   If Not (iFile = 0) Then
      Close #iFile
   End If
   Exit Sub
End Sub

Public Function GetResourceNames(cR As cResources, ByVal vType As Variant) As Long
Dim lR As Long
Dim lErr As Long
Dim lType As Long
Dim sType As String
Dim b() As Byte
Dim lpType As Long
Dim hMem As Long
Dim lPtr As Long

   Set m_cR = cR
   If (VarType(vType) = vbLong) Then
      lType = vType
      lR = EnumResourceNamesByNum(cR.hModule, lType, AddressOf EnumResNamesProc, 0)
   Else
      sType = vType
      b = StrConv(sType, vbFromUnicode)
      ReDim Preserve b(0 To UBound(b) + 1) As Byte
      hMem = GlobalAlloc(GPTR, UBound(b) + 1)
      If Not hMem = 0 Then
         lPtr = GlobalLock(hMem)
         If Not lPtr = 0 Then
            CopyMemory ByVal lPtr, b(0), UBound(b) + 1
            lR = EnumResourceNamesByNum(cR.hModule, lPtr, AddressOf EnumResNamesProc, 0)
            GlobalUnlock lPtr
         End If
         GlobalFree hMem
      End If
   End If
   If (lR = 0) Then
      lErr = err.LastDllError
      'Err.Raise vbObjectError + 1048 + 2, App.EXEName & ".cResource", WinError(lErr)
   End If
   Set m_cR = Nothing
   GetResourceNames = lR
      
End Function

Public Function EnumResNamesProc( _
      ByVal hMod As Long, _
      ByVal lpszType As Long, _
      ByVal lpszName As Long, _
      ByVal lParam As Long _
   ) As Long
Dim sName As String
Dim lName As Long
Dim b() As Byte
Dim lLen As Long

   If (lpszName And &HFFFF0000) = 0 Then
      ' resource number:
      lName = lpszName And &HFFFF&
      m_cR.AddResourceName lName, ""
   Else
      ' resource string:
      lLen = lstrlen(lpszName)
      If (lLen > 0) Then
         ReDim b(0 To lLen - 1) As Byte
         CopyMemory b(0), ByVal lpszName, lLen
         sName = StrConv(b, vbUnicode)
      End If
      m_cR.AddResourceName 0, sName
   End If
   EnumResNamesProc = 1
End Function

Public Function GetResourceTypes(cR As cResources) As Long
Dim lR As Long
Dim lErr As Long
   Set m_cR = cR
   lR = EnumResourceTypes(cR.hModule, AddressOf EnumResTypesProc, 0)
   If (lR = 0) Then
      lErr = err.LastDllError
      Set m_cR = Nothing
      'Err.Raise vbObjectError + 1048 + 1, App.EXEName & ".cResource", WinError(lErr)
   End If
   Set m_cR = Nothing
   GetResourceTypes = lR
End Function

Private Function EnumResTypesProc( _
      ByVal hMod As Long, _
      ByVal lpszType As Long, _
      ByVal lParam As Long _
   ) As Long
Dim lType As Long
Dim sType As String
Dim lLen As Long
Dim b() As Byte
   If (lpszType And &HFFFF0000) = 0 Then
      ' standard resource type:
      lType = lpszType And &HFFFF&
      m_cR.AddResourceType lType, ""
   Else
      ' string:
      lLen = lstrlen(lpszType)
      If (lLen > 0) Then
         ReDim b(0 To lLen - 1) As Byte
         CopyMemory b(0), ByVal lpszType, lLen
         sType = StrConv(b, vbUnicode)
      End If
      m_cR.AddResourceType 0, sType
   End If
   
   EnumResTypesProc = 1
   
End Function
