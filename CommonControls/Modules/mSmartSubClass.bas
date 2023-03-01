Attribute VB_Name = "mSmartSubClass"
' ----------------------------------------------------
' Module mSmartSubClass
'
' Version... 1.0
' Date...... 24 April 2001
'
' Copyright (C) 2001 Andrés Pons (andres@vbsmart.com)
' ----------------------------------------------------

'API declarations:
Option Explicit

Public Const SSC_OLDPROC = "SSC_OLDPROC"
Public Const SSC_OBJADDR = "SSC_OBJADDR"

Private Declare Function GetProp Lib "user32" Alias "GetPropA" ( _
    ByVal hWnd As Long, _
    ByVal lpString As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

'
' Function StartSubclassWindowProc()
'
' This is the first windowproc that receives messages
' for all subclassed windows.
' The aim of this function is to just collect the message
' and deliver it to the right SmartSubClass instance.
'
Public Function SmartSubClassWindowProc( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

    Dim lRet As Long
    Dim oSmartSubClass As SmartSubClass

    'Get the memory address of the class instance...
    lRet = GetProp(hWnd, SSC_OBJADDR)
    
    If lRet <> 0 Then
        'oSmartSubClass will point to the class instance
        'without incrementing the class reference counter...
        CopyMemory oSmartSubClass, lRet, 4
        
        'Send the message to the class instance...
        SmartSubClassWindowProc = oSmartSubClass.WindowProc(hWnd, uMsg, wParam, lParam)

        'Remove the address from memory...
        CopyMemory oSmartSubClass, 0&, 4
    End If
    
End Function
