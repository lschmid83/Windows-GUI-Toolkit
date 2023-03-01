Attribute VB_Name = "mColorDialog"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - mColorDialog.bas                              '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type TCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias _
        "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public CustomColors(0 To 15) As Long
Public Function ColorDlg(hWndParent As Long, DefColor As Long, _
       Optional ShowExpDlg As Boolean = 0) As Long
    
   Dim i
   Dim c As Long
   Dim CC As TCHOOSECOLOR
    
   'Initialise Custom Colours
   For i = 0 To 15
      CustomColors(i) = QBColor(15)
   Next i
    
   CustomColors(0) = &HE3FCFD
   'Show Dialog
   With CC
        
       .rgbResult = DefColor
       .hWndOwner = hWndParent
       .lpCustColors = VarPtr(CustomColors(0))
       .flags = &H101
        
       If ShowExpDlg Then .flags = .flags Or &H2
        
       .lStructSize = Len(CC)
       c = ChooseColor(CC)
        
       If c Then
          ColorDlg = .rgbResult
       Else
          ColorDlg = -1
       End If
        
   End With
    
End Function
