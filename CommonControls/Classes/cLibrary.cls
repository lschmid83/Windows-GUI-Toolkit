VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - cLibrary.cls                                  '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Declare Function LoadLibraryEx Lib "KERNEL32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long

' Missing from VB API declarations:
Private Const DONT_RESOLVE_DLL_REFERENCES = &H1&
Private Const LOAD_LIBRARY_AS_DATAFILE = &H2&
Private Const LOAD_WITH_ALTERED_SEARCH_PATH = &H8&
Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibModule As Long) As Long

Private m_sFileName As String
Private m_hMod As Long

Public Property Get Filename() As String
   Filename = m_sFileName
End Property

Public Property Let Filename(ByVal sFIleName As String)
   ClearUp
   m_sFileName = sFIleName
   If m_sFileName <> "" Then
      m_hMod = LoadLibraryEx(m_sFileName, 0, 0)
      If (m_hMod = 0) Then
       '  Err.Raise vbObjectError + 1048 + 1, App.EXEName & ".cLibrary", WinError(Err.LastDllError)
      End If
   End If
End Property

Public Property Get hModule() As Long
   hModule = m_hMod
End Property

Private Sub ClearUp()
   If (m_hMod <> 0) Then
      FreeLibrary m_hMod
   End If
   m_hMod = 0
   m_sFileName = ""
End Sub


Private Sub Class_Terminate()
   ClearUp
End Sub
