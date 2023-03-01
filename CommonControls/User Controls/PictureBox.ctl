VERSION 5.00
Begin VB.UserControl PictureBox 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   1665
   ScaleWidth      =   2250
   ToolboxBitmap   =   "PictureBox.ctx":0000
End
Attribute VB_Name = "PictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - PictureBox.ctl                                '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_Border As Boolean
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipTitle As String
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_AutoSize As Boolean
Dim m_UseMaskColor As Boolean
Dim m_BackColor As OLE_COLOR

' Events
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event MouseLeave()
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the backcolor of the control based on the appearance.
Private Sub pSetAppearance()

    Select Case g_Appearance
    
        Case Is = Blue
            UserControl.BackColor = &HFFFFFF
        Case Is = Green
            UserControl.BackColor = &HFFFFFF
        Case Is = Silver
            UserControl.BackColor = &HFFFFFF
        Case Is = Win98
            UserControl.BackColor = &HFFFFFF
        
    End Select

End Sub

' Sets the forecolor of the control based on the enabled state.
Private Sub pSetEnabled()
    
    If UserControl.Enabled = True Then
        UserControl.BackColor = m_BackColor
    Else
        If g_Appearance <> Win98 Then
            UserControl.BackColor = &H8000000F
        End If
    End If
End Sub

' Paints the control border.
Private Sub pPaintBorder()

    ' Set the enabled graphic path
    Dim sEnabled As String
    If UserControl.Enabled = True Then
        sEnabled = "Enabled\"
    Else
        sEnabled = "Disabled\"
    End If
    
    ' Paint the border
    UserControl.Cls
    If m_Border = True Then
    
        ' Paint the picture
        If UserControl.Picture <> Empty Then
            UserControl.PaintPicture UserControl.Picture, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        End If
        
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "TopLeft", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Left", crBitmap), 0, 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, Height - (3 * Screen.TwipsPerPixelY), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY

    Else
    
        If UserControl.Picture <> Empty Then
            UserControl.PaintPicture UserControl.Picture, 0, 0
        End If
    
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    pSetAppearance
    
    If m_UseMaskColor = True Then
        UserControl.BackStyle = 0
    Else
        UserControl.BackStyle = 1
    End If
    
    pSetEnabled
    UserControl_Resize

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
        
    SetDefaultTheme

    m_AutoSize = False
    m_Border = True
    m_UseMaskColor = False
    m_BackColor = &HFFFFFF
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipTitle = ""
    m_ToolTipStyle = Standard

    Width = 81 * Screen.TwipsPerPixelX
    Height = 33 * Screen.TwipsPerPixelY

    Extender.TabStop = False
    
    Refresh
        
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    SetDefaultTheme
    
    m_Border = PropBag.ReadProperty("Border", True)
    m_BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_UseMaskColor = PropBag.ReadProperty("UseMaskColor", False)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_AutoSize = PropBag.ReadProperty("AutoSize", False)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)

    If g_ControlsRefreshed = True Then
        Refresh
    End If
    
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H0&)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, &HFF00FF)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, False)
    Call PropBag.WriteProperty("UseMaskColor", m_UseMaskColor, False)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("Border", m_Border, True)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    If m_UseMaskColor = False Then
        pPaintBorder
    End If
    
    RaiseEvent Resize

err:
    Exit Sub

End Sub

Private Sub UserControl_Paint()

    If m_Border = True Then
        pPaintBorder
    End If

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    UserControl.BackColor() = New_BackColor
    UserControl_Resize
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Border() As Boolean
Attribute Border.VB_Description = "Returns/sets whether the control has a border."
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    UserControl_Resize
    PropertyChanged "Border"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Refresh
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
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
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)

    Set UserControl.Picture = New_Picture
    
    ' Set mask picture
    If m_UseMaskColor = True Then
        Set UserControl.MaskPicture = New_Picture
    
        If m_AutoSize = True Then
            Width = UserControl.ScaleX(New_Picture.Width)
            Height = UserControl.ScaleY(New_Picture.Height)
        End If
    
    Else

        UserControl_Resize
        
        ' Set size of picture box based on image
        If m_AutoSize = True Then
        
            If m_Border = True Then
                Width = UserControl.ScaleX(New_Picture.Width) + (6 * Screen.TwipsPerPixelX)
                Height = UserControl.ScaleY(New_Picture.Height) + (6 * Screen.TwipsPerPixelY)
            Else
                Width = UserControl.ScaleX(New_Picture.Width)
                Height = UserControl.ScaleY(New_Picture.Height)
            End If
        
        End If
    
    End If
    
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Returns/sets whether the control is automatically resized to display its entire contrents."
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    Set Picture = Picture
    PropertyChanged "AutoSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets whether the color assigned on the MaskColor property is used as the transparent color in the picture."
    UseMaskColor = m_UseMaskColor
End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)
    m_UseMaskColor = New_UseMaskColor
    Set Picture = Picture
    If m_UseMaskColor = True Then
        UserControl.BackStyle = 0
    Else
        UserControl.BackStyle = 1
    End If
    UserControl_Resize
    PropertyChanged "UseMaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
Attribute ToolTipCaption.VB_Description = "Returns/sets the text displayed in the tooltip."
    ToolTipCaption = m_ToolTipCaption
End Property

Public Property Let ToolTipCaption(ByVal New_ToolTipCaption As String)
    m_ToolTipCaption = New_ToolTipCaption
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
    PropertyChanged "ToolTipTitle"
End Property
