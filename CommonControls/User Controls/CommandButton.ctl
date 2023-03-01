VERSION 5.00
Begin VB.UserControl CommandButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   ScaleHeight     =   1005
   ScaleWidth      =   2625
   ToolboxBitmap   =   "CommandButton.ctx":0000
   Begin VB.Line lineFix 
      X1              =   0
      X2              =   2580
      Y1              =   0
      Y2              =   0
   End
   Begin CommonControls.MaskBox imgPicture 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   420
      _ExtentX        =   0
      _ExtentY        =   0
   End
End
Attribute VB_Name = "CommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - CommandButton.ctl                             '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_Alignment As AlignmentConstants
Dim m_FocusRectangle As Boolean
Dim m_ForeColor As OLE_COLOR
Dim m_DropDownArrow As Boolean
Dim m_HasFocus As Boolean
Dim m_Default As Boolean
Dim m_Caption As String
Dim m_Picture As Picture
Dim m_MaskColor As OLE_COLOR
Dim m_UseMaskColor As Boolean
Dim m_State As String
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
Dim m_PictureHeight As Long
Dim m_PictureWidth As Long

' Events
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseLeave()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

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
Private Sub pPaintComponent(sState As String)
    
    m_State = sState
     
    ' Draw button graphics
    UserControl.Cls
    
    ' Set line fix color
    lineFix.x2 = UserControl.Width
    If g_Appearance = Blue Then
        lineFix.BorderColor = &HD8E9EC
    ElseIf g_Appearance = Green Then
        lineFix.BorderColor = &HD8E9EC
    ElseIf g_Appearance = Silver Then
        lineFix.BorderColor = &HE3DFE0
    ElseIf g_Appearance = Win98 Then
        lineFix.BorderColor = &HC8D0D4
    End If
    
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\Back", crBitmap), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, Width - (6 * Screen.TwipsPerPixelX), Height - (6 * Screen.TwipsPerPixelY)
    
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\TopLeft", crBitmap), 0, 10
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 10
       
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\Left", crBitmap), 0, 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelY, Height - (6 * Screen.TwipsPerPixelY)
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelY, Height - (6 * Screen.TwipsPerPixelY)
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\Top", crBitmap), 3 * Screen.TwipsPerPixelY, 10, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    
    If InIDE Or Is32Bit = True Then
        lineFix.Visible = False
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelY)
    
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\Bottom", crBitmap), 3 * Screen.TwipsPerPixelY, Height - (3 * Screen.TwipsPerPixelY), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    Else
        lineFix.Visible = True
        lineFix.x2 = UserControl.Width
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
    
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "CommandButton\" & m_State & "\Bottom", crBitmap), 3 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    End If
 
    ' Draw dropdown arror
    If m_DropDownArrow = True Then
    
        If m_State <> "Pressed" Then
        
            If m_State <> "Disabled" Then
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\UnPressed\Arrow", crBitmap), Width - (13 * Screen.TwipsPerPixelX), (Height / 2) - (1 * Screen.TwipsPerPixelY)
            Else
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Disabled\Arrow", crBitmap), Width - (13 * Screen.TwipsPerPixelX), (Height / 2) - (1 * Screen.TwipsPerPixelY)
            End If
        
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\UnPressed\Arrow", crBitmap), Width - (12 * Screen.TwipsPerPixelX), (Height / 2)
        End If
    
    End If
        
    ' Set caption alignment
    Dim lTextAlign As Long
    Select Case m_Alignment
        Case vbCenter
            lTextAlign = DT_CENTER
        Case vbLeftJustify
            lTextAlign = DT_LEFT
        Case vbRightJustify
            lTextAlign = DT_RIGHT
    End Select
    
    ' Draw caption and picture
    Dim tUsrRect As RECT
    If Not m_Picture Is Nothing Then
      
        If m_State = "Pressed" Then
            Call SetRect(tUsrRect, (m_PictureWidth / Screen.TwipsPerPixelX) + 15, (((Height / 2) / Screen.TwipsPerPixelY) - UserControl.Font.Size) + 3, (Width / Screen.TwipsPerPixelX) - (m_PictureWidth / Screen.TwipsPerPixelX), Height)
            Call DrawText(UserControl.hdc, m_Caption, -1, tUsrRect, lTextAlign)
            Call SetRectEmpty(tUsrRect)
            m_HasFocus = True
        Else
            Call SetRect(tUsrRect, (m_PictureWidth / Screen.TwipsPerPixelX) + 14, (((Height / 2) / Screen.TwipsPerPixelY) - UserControl.Font.Size) + 2, (Width / Screen.TwipsPerPixelX) - (m_PictureWidth / Screen.TwipsPerPixelX), Height)
            Call DrawText(UserControl.hdc, m_Caption, -1, tUsrRect, lTextAlign)
            Call SetRectEmpty(tUsrRect)
        End If
    ' Draw caption
    Else
    
        If m_State = "Pressed" Then
            m_HasFocus = True
            Call SetRect(tUsrRect, 11, (((Height / 2) / Screen.TwipsPerPixelY) - UserControl.Font.Size) + 3, (Width / Screen.TwipsPerPixelX) - 3, Height)
            Call DrawText(UserControl.hdc, m_Caption, -1, tUsrRect, lTextAlign)
            Call SetRectEmpty(tUsrRect)
        Else
            Call SetRect(tUsrRect, 10, (((Height / 2) / Screen.TwipsPerPixelY) - UserControl.Font.Size) + 2, (Width / Screen.TwipsPerPixelX) - 3, Height)
            Call DrawText(UserControl.hdc, m_Caption, -1, tUsrRect, lTextAlign)
            Call SetRectEmpty(tUsrRect)
        End If

    End If

    'Draw focus rectangle
    If m_HasFocus = True And m_FocusRectangle = True Then
        Call SetRect(tUsrRect, 4, 4, (Width / Screen.TwipsPerPixelX) - 4, (Height / Screen.TwipsPerPixelY) - 4)
        Call DrawFocusRect(UserControl.hdc, tUsrRect)
        Call SetRectEmpty(tUsrRect)
    End If
    
    If Not m_Picture Is Nothing Then
        If m_State = "Pressed" Then
            imgPicture.Left = 11 * Screen.TwipsPerPixelX
            imgPicture.Top = imgPicture.Top + 20
        Else
            imgPicture.Left = 10 * Screen.TwipsPerPixelX
            imgPicture.Top = (Height / 2) - (m_PictureHeight / 2)
        End If

    End If
    
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()
    pSetEnabled
    pPaintComponent "UnPressed"
    
    
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme

    m_DropDownArrow = False
    Set UserControl.Font = UserControl.Font
    m_MaskColor = &HFF00FF
    m_Caption = Ambient.DisplayName
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    m_Default = False
    m_FocusRectangle = True
    m_Alignment = 0
    
    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY

    Refresh
       
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_Caption = PropBag.ReadProperty("Caption", "")
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    If Not m_Picture Is Nothing Then
        m_PictureHeight = ScaleY(m_Picture.Height)
        m_PictureWidth = ScaleX(m_Picture.Width)
    End If
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_Default = PropBag.ReadProperty("Default", False)
    m_MaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
    m_UseMaskColor = PropBag.ReadProperty("UseMaskColor", True)
    m_DropDownArrow = PropBag.ReadProperty("DropDownArrow", False)
    m_FocusRectangle = PropBag.ReadProperty("FocusRectangle", True)
    m_Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.AccessKeys = GetAccessKeyFromString(m_Caption)

    If m_UseMaskColor = True Then
        imgPicture.SetMaskColor m_MaskColor
    End If
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    
    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

    If g_ControlsRefreshed = True Then
        Refresh
    End If
    
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("Caption", m_Caption, "")
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, &HFF00FF)
    Call PropBag.WriteProperty("UseMaskColor", m_UseMaskColor, True)
    Call PropBag.WriteProperty("Default", m_Default, False)
    Call PropBag.WriteProperty("DropDownArrow", m_DropDownArrow, False)
    Call PropBag.WriteProperty("FocusRectangle", m_FocusRectangle, True)
    Call PropBag.WriteProperty("Alignment", m_Alignment, 0)

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    m_HasFocus = True
    pPaintComponent m_State

End Sub

Private Sub UserControl_EnterFocus()
    
    m_HasFocus = True
    If m_State <> "Pressed" Then
        pPaintComponent "HasFocus"
    End If

End Sub

Private Sub UserControl_ExitFocus()

    m_HasFocus = False
    pPaintComponent "UnPressed"

End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 1 Then
        pPaintComponent "Pressed"
        RaiseEvent MouseDown(Button, Shift, x, y)
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    With UserControl
        If GetCapture() = .hwnd Then
                    
            ' Mouse has left control
            If x < 0 Or x > .Width Or y < 0 Or y > .Height Then
        
                Call ReleaseCapture
                
                ' Set the state
                If m_HasFocus = False Then
                    pPaintComponent "UnPressed"
                Else
                    pPaintComponent "HasFocus"
                End If
                
                RaiseEvent MouseLeave
            
            End If
        
        
        Else
            
            ' Mouse has entered control
            Call SetCapture(.hwnd)
        
            If g_Appearance <> Win98 Then
                pPaintComponent "MouseOver"
            
            Else
                If m_HasFocus = True Then
                    pPaintComponent "HasFocus"
                Else
                    pPaintComponent "UnPressed"
                End If
            
            End If

            RaiseEvent MouseMove(Button, Shift, x, y)
        
        End If
    
    End With
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()

On Error GoTo err
   
    pPaintComponent "UnPressed"

err:
    Exit Sub

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of the controls text.\r\n"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    pPaintComponent m_State
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Default() As Boolean
Attribute Default.VB_Description = "Returns/sets whether this button is the default button for the form."
    Default = m_Default
End Property

Public Property Let Default(ByVal New_Default As Boolean)
    m_Default = New_Default
    PropertyChanged "Default"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DropDownArrow() As Boolean
Attribute DropDownArrow.VB_Description = "Returns/sets whether the drop-down arrow is visible on the button."
    DropDownArrow = m_DropDownArrow
End Property

Public Property Let DropDownArrow(ByVal New_DropDownArrow As Boolean)
    m_DropDownArrow = New_DropDownArrow
    UserControl_Resize
    PropertyChanged "DropDownArrow"
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
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    pSetEnabled
    If New_Enabled = True Then
        pPaintComponent "UnPressed"
    Else
        pPaintComponent "Disabled"
    End If
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    pPaintComponent m_State
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MaskColor() As OLE_COLOR
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    UserControl.ForeColor() = m_ForeColor
    pPaintComponent m_State
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    pPaintComponent m_State
    UserControl.AccessKeys = GetAccessKeyFromString(m_Caption)
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    If Not New_Picture Is Nothing Then
        Set m_Picture = New_Picture
        Set imgPicture.Picture = New_Picture
        m_PictureHeight = ScaleY(m_Picture.Height)
        m_PictureWidth = ScaleX(m_Picture.Width)
        UserControl_Resize
        Refresh
    End If
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
Attribute ToolTipCaption.VB_Description = "Returns/sets the text displayed in the tooltip."
    ToolTipCaption = m_ToolTipCaption
End Property

Public Property Let ToolTipCaption(ByVal New_ToolTipCaption As String)
    m_ToolTipCaption = New_ToolTipCaption
    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipIcon() As ToolTipIconEnum
Attribute ToolTipIcon.VB_Description = "Returns/sets the icon displayed in the tooltip.\r\n"
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
Public Property Get ToolTipTitle() As String
Attribute ToolTipTitle.VB_Description = "Returns/sets the title displayed in the tooltip."
    ToolTipTitle = m_ToolTipTitle
End Property

Public Property Let ToolTipTitle(ByVal New_ToolTipTitle As String)
    m_ToolTipTitle = New_ToolTipTitle
    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipTitle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get UseMaskColor() As Boolean
    UseMaskColor = m_UseMaskColor
End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)
    m_UseMaskColor = New_UseMaskColor
    PropertyChanged "UseMaskColor"
End Property
