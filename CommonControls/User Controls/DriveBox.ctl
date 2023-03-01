VERSION 5.00
Begin VB.UserControl DriveBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   HasDC           =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   2355
   ToolboxBitmap   =   "DriveBox.ctx":0000
   Begin CommonControls.MaskBox3 imgButton 
      Height          =   330
      Left            =   1470
      Top             =   0
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   582
      ScaleHeight     =   330
      ScaleWidth      =   270
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Bottom 
      Height          =   45
      Left            =   0
      Top             =   285
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   1455
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Top 
      Height          =   45
      Left            =   45
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   79
      ScaleHeight     =   45
      ScaleWidth      =   1455
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin CommonControls.MaskBox3 Left 
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   45
      _ExtentX        =   79
      _ExtentY        =   503
      ScaleHeight     =   285
      ScaleWidth      =   45
      ScaleMode       =   1
      AutoRedraw      =   -1  'True
   End
   Begin VB.DriveListBox cmbMain 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1740
   End
End
Attribute VB_Name = "DriveBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - DriveBox.ctl                                  '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_AutoDisabledColor As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_Enabled As Boolean
Dim m_ForeColor As OLE_COLOR
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String

' Events
Event OLECompleteDrag(Effect As Long) 'MappingInfo=cmbMain,cmbMain,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=cmbMain,cmbMain,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer) 'MappingInfo=cmbMain,cmbMain,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=cmbMain,cmbMain,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=cmbMain,cmbMain,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=cmbMain,cmbMain,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event Change() 'MappingInfo=cmbMain,cmbMain,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Attribute Change.VB_MemberFlags = "200"
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=cmbMain,cmbMain,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event Scroll() 'MappingInfo=cmbMain,cmbMain,-1,Scroll
Attribute Scroll.VB_Description = "Occurs when you reposition the scroll box on a control."

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If m_Enabled = True Then
        UserControl.Enabled = True
        pPaintDropdown "UnPressed"
        cmbMain.Enabled = True
        cmbMain.BackColor = m_BackColor
        cmbMain.ForeColor = m_ForeColor
    Else
        UserControl.Enabled = False
        pPaintDropdown "Disabled"
        cmbMain.Enabled = False
        cmbMain.ForeColor = &H92A1A1
        cmbMain.BackColor = &H8000000F
    End If

End Sub

' Paints the control border.
Private Sub pPaintBorder()
    
    ' Set the enabled graphic path
    Dim sEnabled As String
    If m_Enabled = True Then
        sEnabled = "Enabled\"
    Else
        sEnabled = "Disabled\"
    End If
        
    ' Paint the border
    Top.Cls
    Left.Cls
    Bottom.Cls
    
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "TopLeft", crBitmap), 0, 0
    Top.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
    Left.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Left", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, Left.Height
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "BottomLeft", crBitmap), 0, 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    Bottom.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & sEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY

End Sub

' Paints the dropdown arrow with mouse over graphics.
Private Sub pPaintDropdown(sState As String)
    
    imgButton.Cls
    
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Back", crBitmap), 5 * Screen.TwipsPerPixelX, 5 * Screen.TwipsPerPixelY, imgButton.Width, imgButton.Height
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Top", crBitmap), 0, 0
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Left", crBitmap), 0, 5 * Screen.TwipsPerPixelY, 5 * Screen.TwipsPerPixelX, imgButton.Height - (10 * Screen.TwipsPerPixelY)
    
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Right", crBitmap), imgButton.Width - (4 * Screen.TwipsPerPixelX), 5 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelX, imgButton.Height - (10 * Screen.TwipsPerPixelY)
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Bottom", crBitmap), 0, imgButton.Height - (5 * Screen.TwipsPerPixelY)
    imgButton.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ComboBox\" & sState & "\Arrow", crBitmap), 6 * Screen.TwipsPerPixelX, (imgButton.Height / 2) - (3 * Screen.TwipsPerPixelY)

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    cmbMain.Top = 0
    cmbMain.Left = 0
    
    Top.Top = 0
    Top.Left = 0
    Top.Height = 3 * Screen.TwipsPerPixelY
    Left.Top = 3 * Screen.TwipsPerPixelY
    Left.Left = 0
    Left.Width = 3 * Screen.TwipsPerPixelX
    
    Bottom.Left = 0
    Bottom.Height = 3 * Screen.TwipsPerPixelY
    imgButton.Top = 0
    imgButton.Width = 20 * Screen.TwipsPerPixelX
    
    UserControl_Resize
    
    pSetEnabled
    pPaintBorder
    
    UserControl_Resize

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()

    SetDefaultTheme

    m_BackColor = &HFFFFFF
    m_Enabled = True
    m_ForeColor = 0
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    
    Width = 110 * Screen.TwipsPerPixelX

    Refresh
    
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme
 
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    Set cmbMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", 0)
    If Ambient.UserMode = True Then
        Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
        MousePointer = PropBag.ReadProperty("MousePointer", 0)
    End If
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    cmbMain.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)

    pSetToolTip imgButton.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    pSetToolTip cmbMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

    If g_ControlsRefreshed = True Then
        Refresh
    End If
   
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Font", cmbMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("OLEDropMode", cmbMain.OLEDropMode, 0)
  
End Sub

Private Sub imgButton_MouseDown(imgButton As Integer, Shift As Integer, x As Single, y As Single)
    
    pPaintDropdown "Pressed"
    SendMessageLong cmbMain.hwnd, CB_SHOWDROPDOWN, 1, 1
    pPaintDropdown "UnPressed"

End Sub

Private Sub imgButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    With imgButton
    
        If GetCapture() = .hwnd Then
            
            ' Mouse has left control
            If x < 0 Or x > .Width Or y < 0 Or y > .Height Then

                Call ReleaseCapture
                pPaintDropdown "UnPressed"
             
            End If
 
        Else
        
            ' Mouse has entered control
            Call SetCapture(.hwnd)
            pPaintDropdown "MouseOver"
                     
        End If
        
    End With


End Sub

Private Sub imgButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pPaintDropdown "UnPressed"
End Sub

Private Sub cmbMain_Change()
    RaiseEvent Change
End Sub

Private Sub cmbMain_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cmbMain_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cmbMain_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub cmbMain_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    Top.Width = Width
    Left.Height = Height
    Bottom.Top = Height - (3 * Screen.TwipsPerPixelY)
    Bottom.Width = Width
    imgButton.Left = Width - imgButton.Width
    imgButton.Height = cmbMain.Height
    
    cmbMain.Width = Width - (1 * Screen.TwipsPerPixelX)
    Height = cmbMain.Height
    
    pPaintBorder

err:
    Exit Sub

End Sub

Private Sub cmbMain_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub cmbMain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub cmbMain_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub cmbMain_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub cmbMain_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub cmbMain_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

Property Let Drive(New_Drive As String)
    cmbMain.Drive = New_Drive
End Property

Property Get Drive() As String
    Drive = cmbMain.Drive
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
    hwnd = cmbMain.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,OLEDrag
Public Sub OLEDrag()
    cmbMain.OLEDrag
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    cmbMain.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    pSetEnabled
    pPaintBorder
    UserControl_Resize
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cmbMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    If Enabled = True Then
        Set cmbMain.Font = New_Font
        Refresh
        pPaintDropdown "UnPressed"
    End If
    PropertyChanged "Font"
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
    cmbMain.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = cmbMain.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set cmbMain.MouseIcon = New_MouseIcon
    Set imgButton.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = cmbMain.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    cmbMain.MousePointer() = New_MousePointer
    imgButton.MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmbMain,cmbMain,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
    OLEDropMode = cmbMain.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    cmbMain.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
Attribute ToolTipCaption.VB_Description = "Returns/sets the text displayed in the tooltip."
    ToolTipCaption = m_ToolTipCaption
End Property

Public Property Let ToolTipCaption(ByVal New_ToolTipCaption As String)
    m_ToolTipCaption = New_ToolTipCaption
    pSetToolTip imgButton.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    pSetToolTip cmbMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
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
    pSetToolTip imgButton.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    pSetToolTip cmbMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
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
    pSetToolTip imgButton.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    pSetToolTip cmbMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
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
    pSetToolTip imgButton.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    pSetToolTip cmbMain.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipTitle"
End Property
