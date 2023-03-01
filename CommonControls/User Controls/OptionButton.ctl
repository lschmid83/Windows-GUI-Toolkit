VERSION 5.00
Begin VB.UserControl OptionButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   555
   ScaleWidth      =   2190
   ToolboxBitmap   =   "OptionButton.ctx":0000
   Begin VB.Line lineFix 
      X1              =   0
      X2              =   2580
      Y1              =   0
      Y2              =   0
   End
   Begin CommonControls.MaskBox imgPicture 
      Height          =   285
      Left            =   15
      Top             =   105
      Width           =   420
      _ExtentX        =   0
      _ExtentY        =   0
   End
End
Attribute VB_Name = "OptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - OptionButton.ctl                              '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_Font As Font
Dim m_Alignment As AlignmentConstants
Dim m_FocusRectangle As Boolean
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
Dim m_Appearance As AppearanceEnum
Dim m_Caption As String
Dim m_Value As OLE_OPTEXCLUSIVE
Dim m_ForeColor As OLE_COLOR
Dim m_MouseDown As Boolean
Dim m_State As String
Dim m_HasFocus As Boolean

' Events
Event Click()
Attribute Click.VB_UserMemId = -600
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_UserMemId = -607
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event MouseLeave()

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

' Paints the component.
Private Sub pPaintComponent(sState As String)
    
    m_State = sState
        
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
        
    ' Set the enabled graphic path
    Dim sEnabled As String
    If UserControl.Enabled = True Then
        sEnabled = "Enabled\"
    Else
        sEnabled = "Disabled\"
    End If
    
    ' Draw graphics
    UserControl.Cls
    Set imgPicture.Picture = PictureFromResource(g_ResourceLib.hModule, "OptionButton\" & sEnabled & m_State, crBitmap)
    
    ' Draw focus rectangle
    Dim UsrRect As RECT
    If m_HasFocus = True And m_FocusRectangle = True Then
        Call SetRect(UsrRect, 16, 1, (Width / Screen.TwipsPerPixelX), (Height / Screen.TwipsPerPixelY) - 1)
        Call DrawFocusRect(UserControl.hdc, UsrRect)
        Call SetRectEmpty(UsrRect)
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
    
    ' Draw caption
    Call SetRect(UsrRect, 18, 1, (Width / Screen.TwipsPerPixelX) - 1, Height)
    Call DrawText(UserControl.hdc, m_Caption, -1, UsrRect, lTextAlign)
    Call SetRectEmpty(UsrRect)

End Sub

' Sets the value of the control.
Private Sub pSetValue()

If m_Value = True Then
    pPaintComponent "Pressed"
Else
    pPaintComponent "UnPressed"
End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    imgPicture.Left = 1 * Screen.TwipsPerPixelX
    
    pSetAppearance
    pSetValue
    pSetEnabled
    UserControl_Resize
  
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()

    SetDefaultTheme

    m_ForeColor = 0
    m_Caption = Ambient.DisplayName
    m_Alignment = vbLeftJustify
    m_Value = False
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    Set m_Font = Ambient.Font
    m_FocusRectangle = True
    
    Width = 110 * Screen.TwipsPerPixelX
    Height = 16 * Screen.TwipsPerPixelY

    Refresh
   
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Caption = PropBag.ReadProperty("Caption", "OptionButton1")
    m_Value = PropBag.ReadProperty("Value", False)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Balloon)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Alignment = PropBag.ReadProperty("Alignment", 0)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_FocusRectangle = PropBag.ReadProperty("FocusRectangle", True)
    UserControl.AccessKeys = GetAccessKeyFromString(m_Caption)

    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

    If g_ControlsRefreshed = True Then
        Refresh
    End If
 
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, "OptionButton1")
    Call PropBag.WriteProperty("Value", m_Value, False)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Balloon)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Alignment", m_Alignment, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("FocusRectangle", m_FocusRectangle, True)

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        m_MouseDown = True
        
        If m_Value = True Then
            pPaintComponent "MouseDownPressed"
        Else
            pPaintComponent "MouseDownUnPressed"
        End If
    
        RaiseEvent MouseDown(Button, Shift, x, y)
    
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    With UserControl
        
        If GetCapture() = .hwnd Then
        
            ' Mouse has left control
            If x < 0 Or x > .Width Or y < 0 Or y > .Height Then
                
                Call ReleaseCapture
                        
                If m_State = "MouseOver" Or m_State = "MouseDownUnPressed" Or m_State = "UnPressed" Then
                    pPaintComponent "UnPressed"
                Else
                    pPaintComponent "Pressed"
                End If
                     
                RaiseEvent MouseLeave
            
            End If
 
        Else
        
            ' Mouse has entered control
            Call SetCapture(.hwnd)
        
            If Button = 1 Then
                
                If m_MouseDown = True Then
                
                    If m_Value = True Then
                        pPaintComponent "MouseDownPressed"
                    Else
                        pPaintComponent "MouseDownUnPressed"
                    End If
            
                End If
                
            Else
            
                If m_MouseDown = False Then
           
                    If m_State = "UnPressed" Then
                        pPaintComponent "MouseOver"
                    ElseIf m_State = "Pressed" Then
                        
                        If g_Appearance <> Win98 Then
                            pPaintComponent "MouseOverPressed"
                        Else
                            pPaintComponent "Pressed"
                        End If
                        
                    End If
                
                End If
                
            End If
        
            RaiseEvent MouseMove(Button, Shift, x, y)
        
        End If
    
    End With

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
 
        If m_MouseDown = True Then
            Value = True
            RaiseEvent MouseUp(Button, Shift, x, y)
        End If
    
        m_MouseDown = False
    
    End If

End Sub

Private Sub UserControl_EnterFocus()
    m_HasFocus = True
    pPaintComponent m_State
End Sub

Private Sub UserControl_ExitFocus()
    m_HasFocus = False
    pPaintComponent m_State
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    imgPicture.Top = (Height / 2) - (7 * Screen.TwipsPerPixelY)
    pPaintComponent m_State

err:
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

Property Let FontBold(Bold As Boolean)
    UserControl.FontBold = Bold
End Property

Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = UserControl.FontBold
End Property

Property Let FontItalic(Italic As Boolean)
    UserControl.FontItalic = Italic
End Property

Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = UserControl.FontItalic
End Property

Property Let FontName(Name As String)
    UserControl.FontName = Name
End Property

Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = UserControl.FontName
End Property

Property Let FontSize(Size As Long)
    UserControl.FontSize = Size
End Property

Property Get FontSize() As Long
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = UserControl.FontSize
End Property

Property Let FontStrikeThru(StrikeThru As Boolean)
    UserControl.FontStrikeThru = StrikeThru
End Property

Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_MemberFlags = "400"
    FontStrikeThru = UserControl.FontStrikeThru
End Property

Property Let FontUnderline(UnderLine As Boolean)
    UserControl.FontUnderline = UnderLine
End Property

Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = UserControl.FontUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of the controls text."
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    pPaintComponent m_State
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get FocusRectangle() As Boolean
Attribute FocusRectangle.VB_Description = "Returns/sets whether the tab focus rectangle is drawn when the control has focus."
    FocusRectangle = m_FocusRectangle
End Property

Public Property Let FocusRectangle(ByVal New_FocusRectangle As Boolean)
    m_FocusRectangle = New_FocusRectangle
    pPaintComponent m_State
    PropertyChanged "FocusRectangle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    pPaintComponent m_State
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    UserControl.ForeColor = m_ForeColor
    pPaintComponent m_State
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    pPaintComponent m_State
    PropertyChanged "Font"
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
    pSetValue
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
    hwnd = UserControl.hwnd
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
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    UserControl.OLEDropMode() = New_OLEDropMode
    UserControl.OLEDropMode = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon.\r\n"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    UserControl.AccessKeys = New_Caption
    pPaintComponent m_State
    PropertyChanged "Caption"
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
'MemberInfo=0,0,0,0
Public Property Get Value() As OLE_OPTEXCLUSIVE
Attribute Value.VB_Description = "Returns/sets the value of the optionbutton i.e On or Off."
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal sState As OLE_OPTEXCLUSIVE)
    m_Value = sState
    pSetValue
    PropertyChanged "Value"
End Property
