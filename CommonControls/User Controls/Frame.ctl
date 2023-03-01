VERSION 5.00
Begin VB.UserControl Frame 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   3075
   ToolboxBitmap   =   "Frame.ctx":0000
   Begin VB.Line lineFix 
      X1              =   0
      X2              =   2580
      Y1              =   0
      Y2              =   0
   End
   Begin CommonControls.MaskBox TopRight 
      Height          =   225
      Left            =   2160
      Top             =   0
      Width           =   240
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin CommonControls.MaskBox BottomLeft 
      Height          =   225
      Left            =   0
      Top             =   1935
      Width           =   240
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin CommonControls.MaskBox BottomRight 
      Height          =   225
      Left            =   2160
      Top             =   1950
      Width           =   240
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin CommonControls.MaskBox TopLeft 
      Height          =   225
      Left            =   0
      Top             =   0
      Width           =   240
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - Frame.ctl                                     '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_ForeColor As OLE_COLOR
Dim m_Caption As String

' Events
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the control graphics.
Private Sub pSetGraphics()

    Set TopLeft.Picture = PictureFromResource(g_ResourceLib.hModule, "Frame\TopLeft", crBitmap)
    Set TopRight.Picture = PictureFromResource(g_ResourceLib.hModule, "Frame\TopRight", crBitmap)
    Set BottomLeft.Picture = PictureFromResource(g_ResourceLib.hModule, "Frame\BottomLeft", crBitmap)
    Set BottomRight.Picture = PictureFromResource(g_ResourceLib.hModule, "Frame\BottomRight", crBitmap)

End Sub

' Sets the style of the control based on the appearance.
Private Sub pSetAppearance()

    ' Set line fix color
    If g_Appearance = Blue Then
        lineFix.BorderColor = &HD8E9EC
    ElseIf g_Appearance = Green Then
        lineFix.BorderColor = &HD8E9EC
    ElseIf g_Appearance = Silver Then
        lineFix.BorderColor = &HE3DFE0
    ElseIf g_Appearance = Win98 Then
        lineFix.BorderColor = &HC8D0D4
    End If

    If g_Appearance = Blue Then
        UserControl.BackColor = &HD8E9EC
        lblCaption.BackColor = &HD8E9EC
        m_ForeColor = &HFF0000
    ElseIf g_Appearance = Green Then
        UserControl.BackColor = &HD8E9EC
        lblCaption.BackColor = &HD8E9EC
        m_ForeColor = &HFF0000
    ElseIf g_Appearance = Silver Then
        UserControl.BackColor = &HE3DFE0
        lblCaption.BackColor = &HE3DFE0
        m_ForeColor = &HFF0000
    ElseIf g_Appearance = Win98 Then
        UserControl.BackColor = &HC8D0D4
        lblCaption.BackColor = &HC8D0D4
        m_ForeColor = 0
    End If

End Sub

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If UserControl.Enabled = True Then
        lblCaption.ForeColor = m_ForeColor
    Else
        If g_Appearance <> Win98 Then
            lblCaption.ForeColor = &H92A1A1
        Else
            lblCaption.ForeColor = &H808080
        End If
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

Public Sub Refresh()

    pSetAppearance
    pSetGraphics
    pSetEnabled
    UserControl_Resize

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
        
    SetDefaultTheme

    m_ForeColor = 0
    Caption = Ambient.DisplayName

    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY

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
    Caption = PropBag.ReadProperty("Caption", "")
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)

    If g_ControlsRefreshed = True Then
        Refresh
    End If
 
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, "")
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)

End Sub

Private Sub UserControl_Resize()

On Error GoTo err
    
    TopLeft.Top = (lblCaption.Height / 2) - 1 * Screen.TwipsPerPixelY
    TopRight.Top = (lblCaption.Height / 2) - 1 * Screen.TwipsPerPixelY
    
    TopRight.Left = UserControl.Width - TopRight.Width
    
    BottomLeft.Top = UserControl.Height - BottomLeft.Height
    
    BottomRight.Top = UserControl.Height - BottomRight.Height
    BottomRight.Left = UserControl.Width - BottomRight.Width
    
    UserControl.Cls
    
    If InIDE Or Is32Bit = True Then
    
        lineFix.Visible = False
            
        If g_Appearance <> Win98 Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Left", crBitmap), 0, TopLeft.Top + (4 * Screen.TwipsPerPixelY), 1 * Screen.TwipsPerPixelX, Height - ((TopLeft.Top + TopLeft.Height) + BottomLeft.Height)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Right", crBitmap), Width - (1 * Screen.TwipsPerPixelX), TopRight.Top + (4 * Screen.TwipsPerPixelY), 1 * Screen.TwipsPerPixelX, Height - ((TopLeft.Top + TopLeft.Height) + BottomLeft.Height)
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Top", crBitmap), TopLeft.Width, TopLeft.Top, Width - (TopLeft.Width + TopRight.Width), 1 * Screen.TwipsPerPixelY
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Bottom", crBitmap), TopLeft.Width, Height - (1 * Screen.TwipsPerPixelY), Width - (TopLeft.Width + TopRight.Width), 1 * Screen.TwipsPerPixelY
        
        Else
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Left", crBitmap), 0, TopLeft.Top + (4 * Screen.TwipsPerPixelY), 2 * Screen.TwipsPerPixelX, Height - ((TopLeft.Top + TopLeft.Height) + BottomLeft.Height)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Right", crBitmap), Width - (2 * Screen.TwipsPerPixelX), TopRight.Top + (4 * Screen.TwipsPerPixelY), 2 * Screen.TwipsPerPixelX, Height - ((TopLeft.Top + TopLeft.Height) + BottomLeft.Height)
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Top", crBitmap), TopLeft.Width, TopLeft.Top, Width - (TopLeft.Width + TopRight.Width), 2 * Screen.TwipsPerPixelY
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Bottom", crBitmap), TopLeft.Width, Height - (2 * Screen.TwipsPerPixelY), Width - (TopLeft.Width + TopRight.Width), 2 * Screen.TwipsPerPixelY
        
        End If

    Else
    
        lineFix.Visible = True
        lineFix.x2 = UserControl.Width
            
        If g_Appearance <> Win98 Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Left", crBitmap), 0, TopLeft.Top + (3 * Screen.TwipsPerPixelY), 1 * Screen.TwipsPerPixelX, Height - ((TopLeft.Top + TopLeft.Height) + BottomLeft.Height)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Right", crBitmap), Width - (1 * Screen.TwipsPerPixelX), TopRight.Top + (3 * Screen.TwipsPerPixelY), 1 * Screen.TwipsPerPixelX, Height - ((TopLeft.Top + TopLeft.Height) + BottomLeft.Height)
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Top", crBitmap), TopLeft.Width, TopLeft.Top - 1 * Screen.TwipsPerPixelY, Width - (TopLeft.Width + TopRight.Width), 1 * Screen.TwipsPerPixelY
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Bottom", crBitmap), TopLeft.Width, Height - (2 * Screen.TwipsPerPixelY), Width - (TopLeft.Width + TopRight.Width), 1 * Screen.TwipsPerPixelY
        
        Else
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Left", crBitmap), 0, TopLeft.Top + (3 * Screen.TwipsPerPixelY), 2 * Screen.TwipsPerPixelX, Height - ((TopLeft.Top + TopLeft.Height) + BottomLeft.Height)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Right", crBitmap), Width - (2 * Screen.TwipsPerPixelX), TopRight.Top + (3 * Screen.TwipsPerPixelY), 2 * Screen.TwipsPerPixelX, Height - ((TopLeft.Top + TopLeft.Height) + BottomLeft.Height)
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Top", crBitmap), TopLeft.Width, TopLeft.Top - 1 * Screen.TwipsPerPixelY, Width - (TopLeft.Width + TopRight.Width), 2 * Screen.TwipsPerPixelY
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "Frame\Bottom", crBitmap), TopLeft.Width, Height - (3 * Screen.TwipsPerPixelY), Width - (TopLeft.Width + TopRight.Width), 2 * Screen.TwipsPerPixelY
        
        End If
    
    End If

err:
    Exit Sub

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

Property Let FontBold(Bold As Boolean)
    lblCaption.FontBold = Bold
    UserControl_Resize
End Property

Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = lblCaption.FontBold
End Property

Property Let FontItalic(Italic As Boolean)
    lblCaption.FontItalic = Italic
    UserControl_Resize
End Property

Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = lblCaption.FontItalic
End Property

Property Let FontName(Name As String)
    lblCaption.FontName = Name
    UserControl_Resize
End Property

Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = lblCaption.FontName
End Property

Property Let FontSize(Size As Long)
    lblCaption.FontSize = Size
    UserControl_Resize
End Property

Property Get FontSize() As Long
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = lblCaption.FontSize
End Property

Property Let FontStrikeThru(StrikeThru As Boolean)
    lblCaption.FontStrikeThru = StrikeThru
    UserControl_Resize
End Property

Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_MemberFlags = "400"
    FontStrikeThru = lblCaption.FontStrikeThru
End Property

Property Let FontUnderline(UnderLine As Boolean)
    lblCaption.FontUnderline = UnderLine
    UserControl_Resize
End Property

Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = lblCaption.FontUnderline
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
    lblCaption.BackColor = New_BackColor
    UserControl_Resize
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    lblCaption.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object.\r\n"
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    UserControl_Resize
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
    PropertyChanged "Enabled"
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
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the frame."
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    lblCaption.Caption = " " & m_Caption & " "
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property
