VERSION 5.00
Begin VB.UserControl Hyperlink 
   AutoRedraw      =   -1  'True
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   ScaleHeight     =   570
   ScaleWidth      =   1680
   ToolboxBitmap   =   "Hyperlink.ctx":0000
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - Hyperlink.ctl                                 '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_FocusRectangle As Boolean
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipCaption As String
Dim m_ToolTipTitle As String
Dim m_MouseOverUnderline As Boolean
Dim m_URL As String
Dim m_MouseOverColor As OLE_COLOR
Dim m_MouseLeaveColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_HasFocus As Boolean

' Events
Event MouseLeave()
Event Click() 'MappingInfo=lblMain,lblMain,-1,Click
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblMain,lblMain,-1,MouseDown
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblMain,lblMain,-1,MouseMove
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=lblMain,lblMain,-1,MouseUp
Attribute MouseUp.VB_UserMemId = -607

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the backcolor of the control based on the appearance.
Private Sub pSetAppearance()

    If g_Appearance = Blue Then
        UserControl.BackColor = &HD8E9EC
    ElseIf g_Appearance = Green Then
        UserControl.BackColor = &HD8E9EC
    ElseIf g_Appearance = Silver Then
        UserControl.BackColor = &HE3DFE0
    ElseIf g_Appearance = Win98 Then
        UserControl.BackColor = &HC8D0D4
    End If

End Sub

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If UserControl.Enabled = True Then
        UserControl.ForeColor = m_ForeColor
    Else
    
        If g_Appearance <> Win98 Then
            UserControl.ForeColor = &H92A1A1
        Else
            UserControl.ForeColor = &H808080
        End If
    End If

End Sub

' Paints the control.
Private Sub pPaintComponent()
    
    ' Draw caption
    Dim lTextAlign As Long
    Dim tUsrRect As RECT
    UserControl.Cls
    lTextAlign = DT_LEFT
    Call SetRect(tUsrRect, 1, 1, (Width / Screen.TwipsPerPixelX) - 2, (Height / Screen.TwipsPerPixelY))
    Call DrawText(UserControl.hdc, lblMain.Caption, -1, tUsrRect, lTextAlign)
    Call SetRectEmpty(tUsrRect)
    
    ' Draw focus rectangle
    If m_HasFocus = True And m_FocusRectangle = True Then
        UserControl.ForeColor = 0
        Call SetRect(tUsrRect, 0, 0, (Width / Screen.TwipsPerPixelX), (Height / Screen.TwipsPerPixelY))
        Call DrawFocusRect(UserControl.hdc, tUsrRect)
        Call SetRectEmpty(tUsrRect)
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    pSetAppearance
    Set UserControl.MouseIcon = LoadResPicture("Hand", vbResCursor)
    
    lblMain.Top = 0
    lblMain.Left = 0
    
    pSetEnabled
    Call UserControl_Resize
    pPaintComponent

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme

    m_URL = ""
    m_MouseOverColor = &H40C0&
    m_MouseOverUnderline = True
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipCaption = ""
    m_ToolTipTitle = ""
    m_ForeColor = 0
    m_FocusRectangle = True
    Caption = Ambient.DisplayName
    
    Width = 1020
    Height = 270

    Refresh
    
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    lblMain.Caption = PropBag.ReadProperty("Caption", "Label1")
    Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblMain.Font = PropBag.ReadProperty("Font", UserControl.Font)
    Set UserControl.Font = lblMain.Font
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.ForeColor = m_ForeColor
    m_URL = PropBag.ReadProperty("URL", "")
    m_MouseOverColor = PropBag.ReadProperty("MouseOverColor", &H40C0&)
    m_MouseOverUnderline = PropBag.ReadProperty("MouseOverUnderline", True)
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_FocusRectangle = PropBag.ReadProperty("FocusRectangle", True)

    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

    If g_ControlsRefreshed = True Then
        Refresh
    End If
   
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", lblMain.Caption, "Label1")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("URL", m_URL, "")
    Call PropBag.WriteProperty("MouseOverColor", m_MouseOverColor, &H40C0&)
    Call PropBag.WriteProperty("MouseOverUnderline", m_MouseOverUnderline, True)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("FocusRectangle", m_FocusRectangle, True)

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 1 Then
        
        ' Execute the link in a browser
        Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
    
        UserControl.MousePointer = 1
            
        UserControl.ForeColor = m_MouseLeaveColor
        
        ' Underline the hyperlink
        If MouseOverUnderline = True Then
            UserControl.FontUnderline = False
        End If
   
        ' Repaint the component
        pPaintComponent
                 
    End If
    
    RaiseEvent MouseDown(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    With UserControl
        
        If GetCapture() = .hwnd Then
                    
            ' Mouse has left control
            If x < 0 Or x > .Width Or y < 0 Or y > .Height Then

                Call ReleaseCapture
            
                UserControl.MousePointer = 1
                    
                UserControl.ForeColor = m_MouseLeaveColor
                
                If MouseOverUnderline = True Then
                    UserControl.FontUnderline = False
                End If

                pPaintComponent
      
            End If
        
        Else
            
            ' Mouse has entered control
            Call SetCapture(.hwnd)
        
            RaiseEvent MouseMove(Button, Shift, x, y)
            UserControl.MousePointer = 99
                
            If ForeColor <> MouseOverColor Then
                 
                m_MouseLeaveColor = UserControl.ForeColor
                UserControl.ForeColor = MouseOverColor
            End If

            If MouseOverUnderline = True Then
                
                If UserControl.FontUnderline = False Then
                UserControl.FontUnderline = True
                End If
            
            End If
 
            pPaintComponent
        
        End If
    
    End With

End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    Width = lblMain.Width + (3 * Screen.TwipsPerPixelX)
    Height = lblMain.Height + (2 * Screen.TwipsPerPixelY)
        
err:
    Exit Sub
    
End Sub

Private Sub UserControl_EnterFocus()

    m_HasFocus = True
    pPaintComponent

End Sub

Private Sub UserControl_ExitFocus()

    m_HasFocus = False
    pPaintComponent

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMain,lblMain,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    pPaintComponent
    PropertyChanged "BackColor"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMain,lblMain,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon.\r\n"
Attribute Caption.VB_UserMemId = -518
    Caption = lblMain.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblMain.Caption() = New_Caption
    UserControl_Resize
    pPaintComponent
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMain,lblMain,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events.\r\n"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    pSetEnabled
    pPaintComponent
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMain,lblMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Set lblMain.Font = New_Font
    UserControl_Resize
    pPaintComponent
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMain,lblMain,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    UserControl.ForeColor = m_ForeColor
    pPaintComponent
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FocusRectangle() As Boolean
Attribute FocusRectangle.VB_Description = "Returns/sets whether the tab focus rectangle is drawn when the control has focus."
    FocusRectangle = m_FocusRectangle
End Property

Public Property Let FocusRectangle(ByVal New_FocusRectangle As Boolean)
    m_FocusRectangle = New_FocusRectangle
    PropertyChanged "FocusRectangle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get MouseOverUnderline() As Boolean
Attribute MouseOverUnderline.VB_Description = "Returns/sets whether the text is uderlined when the mouse is over the control."
    MouseOverUnderline = m_MouseOverUnderline
End Property

Public Property Let MouseOverUnderline(ByVal New_MouseOverUnderline As Boolean)
    m_MouseOverUnderline = New_MouseOverUnderline
    PropertyChanged "MouseOverUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipIcon() As ToolTipIconEnum
Attribute ToolTipIcon.VB_Description = "Returns/sets the icon displayed in the tooltip."
    ToolTipIcon = m_ToolTipIcon
End Property

Public Property Let ToolTipIcon(ByVal New_ToolTipIcon As ToolTipIconEnum)
    m_ToolTipIcon = New_ToolTipIcon
    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipStyle() As ToolTipStyleEnum
Attribute ToolTipStyle.VB_Description = "Returns/sets the style of the tooltip i.e Standad or Balloon."
    ToolTipStyle = m_ToolTipStyle
End Property

Public Property Let ToolTipStyle(ByVal New_ToolTipStyle As ToolTipStyleEnum)
    m_ToolTipStyle = New_ToolTipStyle
    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
Attribute ToolTipCaption.VB_Description = "Returns/sets the text displayed in the tooltip.\r\n"
    ToolTipCaption = m_ToolTipCaption
End Property

Public Property Let ToolTipCaption(ByVal New_ToolTipCaption As String)
    m_ToolTipCaption = New_ToolTipCaption
    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipTitle() As String
Attribute ToolTipTitle.VB_Description = "Returns/sets the title displayed in the tooltip.\r\n"
    ToolTipTitle = m_ToolTipTitle
End Property

Public Property Let ToolTipTitle(ByVal New_ToolTipTitle As String)
    m_ToolTipTitle = New_ToolTipTitle
    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get URL() As String
Attribute URL.VB_Description = "Returns/sets the URL which is opened when the control is clicked."
    URL = m_URL
End Property

Public Property Let URL(ByVal New_URL As String)
    m_URL = New_URL
    PropertyChanged "URL"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MouseOverColor() As OLE_COLOR
Attribute MouseOverColor.VB_Description = "Returns/sets the text color when the mouse is over the control."
    MouseOverColor = m_MouseOverColor
End Property

Public Property Let MouseOverColor(ByVal New_MouseOverColor As OLE_COLOR)
    m_MouseOverColor = New_MouseOverColor
    PropertyChanged "MouseOverColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
    hwnd = UserControl.hwnd
End Property
