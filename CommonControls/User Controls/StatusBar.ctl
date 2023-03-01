VERSION 5.00
Begin VB.UserControl StatusBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   ControlContainer=   -1  'True
   ScaleHeight     =   345
   ScaleWidth      =   4155
   ToolboxBitmap   =   "StatusBar.ctx":0000
End
Attribute VB_Name = "StatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - StatusBar.ctl                                 '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Implements ISubclass

' Member variables
Private WithEvents m_ParentForm As Form
Attribute m_ParentForm.VB_VarHelpID = -1
Dim m_Caption As String
Dim m_ResizeHandle As Boolean
Dim m_StatusVisible As Boolean
Dim m_HasFocus As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Paints the component
Private Sub pPaintComponent(bFocus As Boolean)

    m_HasFocus = bFocus
    
    ' Set the enabled graphics path
    Dim sEnabled As String
    sEnabled = "HasFocus"
    If bFocus = False Then
        sEnabled = "LostFocus"
    End If
        
    ' Draw the statusbar
    UserControl.Cls
    If m_StatusVisible = True Then

        If g_Appearance <> Win98 Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\StatusBar\Back", crBitmap), 0, 0, Width
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\StatusBar\Back", crBitmap), 5 * Screen.TwipsPerPixelX, 0
        End If
        
        ' Draw caption
        Dim tUsrRect As RECT
        Call SetRect(tUsrRect, 8, 5, (Width / Screen.TwipsPerPixelX) - 10, Height)
        Call DrawText(UserControl.hdc, m_Caption, -1, tUsrRect, DT_LEFT)
        Call SetRectEmpty(tUsrRect)
        
        ' Draw borders
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\Borders\Left", crBitmap), 0, 0, , Height
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\Borders\Bottom", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY), Width
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\Borders\Right", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 0, 4 * Screen.TwipsPerPixelX, Height
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\Borders\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\Borders\BottomLeft", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY)

        ' Draw resize handle
        If m_ResizeHandle = True Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\HasFocus\StatusBar\ResizeHandle", crBitmap), Width - (16 * Screen.TwipsPerPixelX), 0
        End If

    Else
        
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\Borders\Bottom", crBitmap), 0, 0, Width
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\Borders\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "FormSkin\" & sEnabled & "\Borders\BottomLeft", crBitmap), 0, 0
    
    End If

End Sub

' Sets the statusbar visibility and initializes min form height.
Private Sub pSetStatusVisible()

    If m_StatusVisible = True Then
        g_StatusVisible = True
        g_MinFormHeight = 57
    Else
        g_StatusVisible = False
        Height = 4 * Screen.TwipsPerPixelY
        g_MinFormHeight = 34
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Update control
Public Sub Refresh()

    Height = 10
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
    
    SetDefaultTheme
   
    m_Caption = Ambient.DisplayName
    m_ResizeHandle = True
    m_StatusVisible = True
    Extender.Align = 2
    Width = 20
        
    Refresh
    
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

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
    
    m_Caption = PropBag.ReadProperty("Caption", "Statusbar1")
    m_ResizeHandle = PropBag.ReadProperty("ResizeHandle", True)
    m_StatusVisible = PropBag.ReadProperty("StatusVisible", True)

    If g_ControlsRefreshed = True Then
        Refresh
    End If

End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("Caption", m_Caption, "Statusbar1")
    Call PropBag.WriteProperty("ResizeHandle", m_ResizeHandle, True)
    Call PropBag.WriteProperty("StatusVisible", m_StatusVisible, True)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If g_BorderStyle = Sizable Then
    
        If m_ResizeHandle = True Then
            
            ' Bottom Border
            If y > UserControl.Height - 60 Then
            
                ' Resize-down
                If x > 200 And x < UserControl.Width - 300 Then
                    UserControl.MousePointer = 7
                ' Resize right-down
                ElseIf x > UserControl.Width - 300 Then
                    UserControl.MousePointer = 8
                ' Resize left-down
                ElseIf x < 200 Then
                    UserControl.MousePointer = 6
                End If
            
            
            Else 'Mouse in upper area
                    
                ' Normal mouse pointer
                If x > 200 And x < UserControl.Width - 300 Then
                    UserControl.MousePointer = 1
                ' Resize left-down
                ElseIf x < 60 Then
                    UserControl.MousePointer = 6
                ' Resize right-down
                ElseIf x > UserControl.Width - 300 Then
                    UserControl.MousePointer = 8
                End If
 
            End If
        
        End If
    
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If g_BorderStyle = Sizable Then
    
        If m_ResizeHandle = True Then
        
            If y > UserControl.Height - 60 Then
                
                ' Resize down
                If x > 200 And x < UserControl.Width - 300 Then
                    ReleaseCapture
                    SendMessage g_hwnd, 274, 61446, 0
                ' Resize right-down
                ElseIf x > UserControl.Width - 300 Then
                    ReleaseCapture
                    SendMessage g_hwnd, 274, 61448, 0
                ' Resize left-down
                ElseIf x < 200 Then
                    ReleaseCapture
                    SendMessage g_hwnd, 274, 61447, 0
                End If
       
            Else 'Mouse in upper area
            
            
                ' Normal mouse pointer
                If x > 200 And x < UserControl.Width - 300 Then
                    UserControl.MousePointer = 1
                ' Resize left-down
                ElseIf x < 60 Then
                    ReleaseCapture
                    SendMessage g_hwnd, 274, 61447, 0
                'Resize right-down
                ElseIf x > UserControl.Width - 300 Then
                    ReleaseCapture
                    SendMessage g_hwnd, 274, 61448, 0
                End If
      
            End If

        End If
    
    End If

End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    If StatusVisible = True Then
    
        If g_Appearance <> Win98 Then
            Height = 27 * Screen.TwipsPerPixelY
        Else
            Height = 24 * Screen.TwipsPerPixelY
        End If
    
    Else
        Height = 4 * Screen.TwipsPerPixelY
    End If
    
    pPaintComponent m_HasFocus

err:
    Exit Sub

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the statusbar."
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    pPaintComponent m_HasFocus
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ResizeHandle() As Boolean
Attribute ResizeHandle.VB_Description = "Returns/sets whether the resize handle is visible."
    ResizeHandle = m_ResizeHandle
End Property

Public Property Let ResizeHandle(ByVal New_ResizeHandle As Boolean)
    m_ResizeHandle = New_ResizeHandle
    UserControl_Resize
    PropertyChanged "ResizeHandle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get StatusVisible() As Boolean
Attribute StatusVisible.VB_Description = "Returns/sets whether the statusbar is visible or just the border."
    StatusVisible = m_StatusVisible
End Property

Public Property Let StatusVisible(ByVal New_StatusVisible As Boolean)
    m_StatusVisible = New_StatusVisible
    pSetStatusVisible
    pPaintComponent m_HasFocus
    UserControl_Resize
    PropertyChanged "StatusVisible"
End Property
