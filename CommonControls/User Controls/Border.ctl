VERSION 5.00
Begin VB.UserControl Border 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   HasDC           =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   300
   ToolboxBitmap   =   "Border.ctx":0000
End
Attribute VB_Name = "Border"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - Border.ctl                                    '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Implements ISubclass

' Member variables
Private WithEvents m_ParentForm As Form
Attribute m_ParentForm.VB_VarHelpID = -1
Dim m_BorderType As BorderEnum
Dim m_Focus As Boolean
                        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions
                        
' Sets the dimensions of the control based on the border type.
Private Sub pSetBorderType()
    
    If m_BorderType = bdrBottom Then
        UserControl.Height = 4 * Screen.TwipsPerPixelY
    ElseIf m_BorderType = bdrLeft Then
        UserControl.Width = 4 * Screen.TwipsPerPixelX
    ElseIf m_BorderType = bdrRight Then
        UserControl.Width = 4 * Screen.TwipsPerPixelX
    End If

End Sub

' Paints the control.
Public Sub pPaintComponent(bFocus As Boolean)
    
    m_Focus = bFocus
    
    ' Set the focus path
    Dim sFocusPath As String
    If bFocus = True Then
        sFocusPath = "HasFocus"
    Else
        sFocusPath = "LostFocus"
    End If
        
    ' Draw the graphics
    UserControl.Cls
    
    If BorderType = bdrLeft Then
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sFocusPath & "\Borders\Left", crBitmap), 0, 0, 4 * Screen.TwipsPerPixelX, Height
    ElseIf BorderType = bdrRight Then
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sFocusPath & "\Borders\Right", crBitmap), 0, 0, 4 * Screen.TwipsPerPixelX, Height
    ElseIf BorderType = bdrBottom Then
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sFocusPath & "\Borders\Bottom", crBitmap), 4 * Screen.TwipsPerPixelX, 0, Width, 4 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sFocusPath & "\Borders\BottomLeft", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sFocusPath & "\Borders\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 0
    End If
   
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions
    
' Update the control.
Public Sub Refresh()
    pPaintComponent True
End Sub
                        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Subclassing

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    m_emr = emrConsume
    ISubclass_MsgResponse = m_emr
End Property

' Tell the subclasser what to do for this message.
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    ' Check if the active window is about to be activated
    If iMsg = WM_ACTIVATEAPP Then
    
        If wParam = 0 Then
            pPaintComponent False
        Else
            pPaintComponent True
        End If
    
    End If

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()

    ' Initialize default theme
    SetDefaultTheme
        
    ' Set the border type based on the last border control added to the form
    If g_LastBorderType = "" Then
        m_BorderType = bdrLeft
        Extender.Align = 3
        Height = 20
        g_LastBorderType = "bdrLeft"
    ElseIf g_LastBorderType = "bdrLeft" Then
        m_BorderType = bdrRight
        Extender.Align = 4
        Height = 20
        g_LastBorderType = "bdrRight"
    ElseIf g_LastBorderType = "bdrRight" Then
        m_BorderType = bdrBottom
        Extender.Align = 2
        Width = 20
        g_LastBorderType = ""
    End If
    
    ' Set border type
    pSetBorderType

    ' Update control
    Refresh
    
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Initialize default theme
    SetDefaultTheme

    ' Set the focus
    m_Focus = True
      
    ' Check if the environment is in design mode or end user mode
    If Ambient.UserMode = True Then
    
        ' Set the global parent window hWnd
        If g_hwnd = 0 Then
            Set m_ParentForm = UserControl.Parent
            g_hwnd = m_ParentForm.hwnd
        End If
    
        ' Attach the parent window focus message
        AttachMessage Me, g_hwnd, WM_ACTIVATEAPP
        
    End If
    
    ' Retrieve the border type property
    m_BorderType = PropBag.ReadProperty("BorderType", bdrLeft)

    ' Update control
    If g_ControlsRefreshed = True Then
        Refresh
    End If

End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderType", m_BorderType, bdrLeft)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 1 Or g_BorderStyle <> Sizable Then
        Exit Sub
    End If
        
    ' Resize the parent form based on the border type
    If BorderType = bdrLeft Then
        ReleaseCapture
        SendMessage g_hwnd, 274, 61441, 0
    ElseIf BorderType = bdrRight Then
        ReleaseCapture
        SendMessage g_hwnd, 274, 61442, 0
    ElseIf BorderType = bdrBottom Then
    
        ' Left-down
        If x < 3 * Screen.TwipsPerPixelX Then
            ReleaseCapture
            SendMessage g_hwnd, 274, 61447, 0
        ' Right-Down
        ElseIf x > UserControl.Width - (3 * Screen.TwipsPerPixelX) Then
            ReleaseCapture
            SendMessage g_hwnd, 274, 61448, 0
        ' Down
        Else
            ReleaseCapture
            SendMessage g_hwnd, 274, 61446, 0
        End If
    
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If g_BorderStyle <> Sizable Then
        Exit Sub
    End If

    ' Set the mouse pointer
    If BorderType = bdrLeft Then
        UserControl.MousePointer = 9
    ElseIf BorderType = bdrRight Then
        UserControl.MousePointer = 9
    ElseIf BorderType = bdrBottom Then
    
        If x < 3 * Screen.TwipsPerPixelX Then
            UserControl.MousePointer = 6
        ElseIf x > UserControl.Width - (3 * Screen.TwipsPerPixelX) Then
            UserControl.MousePointer = 8
        Else
            UserControl.MousePointer = 7
        End If
    
    End If

End Sub

Private Sub UserControl_Resize()

On Error GoTo err
    
    If BorderType = bdrLeft Then
        UserControl.Width = 4 * Screen.TwipsPerPixelX
    
    ElseIf BorderType = bdrRight Then
        UserControl.Width = 4 * Screen.TwipsPerPixelX
    
    ElseIf BorderType = bdrBottom Then
        UserControl.Height = 4 * Screen.TwipsPerPixelY
    End If
    
    pPaintComponent m_Focus
    
err:
    Exit Sub

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get BorderType() As BorderEnum
Attribute BorderType.VB_Description = "Returns/sets the border type."
    BorderType = m_BorderType
End Property

Public Property Let BorderType(ByVal New_BorderType As BorderEnum)
    
    m_BorderType = New_BorderType
    
    pSetBorderType
    pPaintComponent True
    
    If m_BorderType = bdrBottom Then
        Extender.Align = 2
        Width = 30
    ElseIf m_BorderType = bdrLeft Then
        Extender.Align = 3
        Height = 30
    ElseIf m_BorderType = bdrRight Then
        Extender.Align = 4
        Height = 30
    
    End If

    PropertyChanged "BorderType"
    
End Property
