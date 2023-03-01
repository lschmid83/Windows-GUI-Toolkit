VERSION 5.00
Begin VB.UserControl ExtraControls 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   585
   ControlContainer=   -1  'True
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   570
   ScaleWidth      =   585
   ToolboxBitmap   =   "ExtraControls.ctx":0000
   Begin VB.Line lineFix2 
      BorderColor     =   &H00C8D0D4&
      X1              =   0
      X2              =   2580
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line lineFix 
      X1              =   0
      X2              =   2580
      Y1              =   0
      Y2              =   0
   End
   Begin CommonControls.MaskBox imgPicture 
      Height          =   165
      Left            =   0
      Top             =   0
      Width           =   195
      _ExtentX        =   0
      _ExtentY        =   0
   End
End
Attribute VB_Name = "ExtraControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - ExtraControls.ctl                             '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
Dim m_ControlType As ExtaControlsEnum
Dim m_ControlPath As String
Dim m_HasFocus As Boolean
Dim m_State As String

' Events
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
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Paints scrollbar down arrow.
Private Sub pPaintDownArrow(New_Value As String)
    
    m_State = New_Value
    UserControl.Picture = PictureFromResource(g_ResourceLib.hModule, "ScrollBar\Vertical\Down\" & m_State, crBitmap)

End Sub

' Paints scrollbar up arrow.
Private Sub pPaintUpArrow(New_Value As String)

    m_State = New_Value
    UserControl.Picture = PictureFromResource(g_ResourceLib.hModule, "ScrollBar\Vertical\Up\" & m_State, crBitmap)

End Sub

' Paints scrollbar left arrow.
Private Sub pPaintLeftArrow(New_Value As String)

    m_State = New_Value
    UserControl.Picture = PictureFromResource(g_ResourceLib.hModule, "ScrollBar\Horizontal\Left\" & m_State, crBitmap)

End Sub

' Paints scrollbar right arrow.
Private Sub pPaintRightArrow(New_Value As String)

    m_State = New_Value
    UserControl.Picture = PictureFromResource(g_ResourceLib.hModule, "ScrollBar\Horizontal\Right\" & m_State, crBitmap)

End Sub

' Paints dropdown selection arrow button.
Private Sub pPaintComboArrow(New_Value As String)

    m_State = New_Value
    
    UserControl.Cls
      
    'Loads the correct background picture
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ArrowButton\" & m_State & "\Back", crBitmap), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, Width - (6 * Screen.TwipsPerPixelX), Height - (6 * Screen.TwipsPerPixelY)
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ArrowButton\" & m_State & "\Top", crBitmap), 0, 0, Width, 3 * Screen.TwipsPerPixelY
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ArrowButton\" & m_State & "\Left", crBitmap), 0, 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ArrowButton\" & m_State & "\Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ArrowButton\" & m_State & "\Bottom", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY), Width, 3 * Screen.TwipsPerPixelY
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ArrowButton\" & m_State & "\Arrow", crBitmap), Width / 2 - (4 * Screen.TwipsPerPixelX), Height / 2 - (2 * Screen.TwipsPerPixelY)

End Sub

' Paints the toolbar truncate button.
Private Sub pPaintTruncateButton(New_Value As String)

    m_State = New_Value
    
    UserControl.Cls
    
    If InIDE Or Is32Bit = True Then
        lineFix.Visible = False
    Else
     
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
    
    End If
    
    ' Draw the button and border
    If m_State <> "UnPressed" And m_State <> "Disabled" Then
    
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, Width - (8 * Screen.TwipsPerPixelX), Height - (8 * Screen.TwipsPerPixelY)
    
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopLeft", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 0
        
        If InIDE Or Is32Bit = True Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (5 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
        End If
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Left", crBitmap), 0, 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Right", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Top", crBitmap), 4 * Screen.TwipsPerPixelY, 0, Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
        If InIDE Or Is32Bit = True Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY), Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY), Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
        End If
    Else
        If InIDE Or Is32Bit = True Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 0, 0, Width, Height
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 0, 0, Width, Height - (1 * Screen.TwipsPerPixelY)
        End If
    End If
    
    ' Draw the arrow
    If g_Appearance <> Win98 Then
        UserControl.PaintPicture LoadResPicture("WinXP\TruncateButton\Arrow", vbResBitmap), Width - (13 * Screen.TwipsPerPixelX), 6 * Screen.TwipsPerPixelY
    Else
    
        If m_State <> "Pressed" Then
            UserControl.PaintPicture LoadResPicture("Win98\TruncateButton\Arrow", vbResBitmap), Width - (12 * Screen.TwipsPerPixelX), 5 * Screen.TwipsPerPixelY
        Else
            UserControl.PaintPicture LoadResPicture("Win98\TruncateButton\Arrow", vbResBitmap), Width - (11 * Screen.TwipsPerPixelX), 6 * Screen.TwipsPerPixelY
        End If
    
    End If

End Sub

' Paints the separator.
Private Sub pPaintSeparator(SepStyle As String)

    If InIDE Or Is32Bit = True Then
        lineFix.Visible = False
    Else

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
    End If

    If g_Appearance <> Win98 Then
    
        If SepStyle = "Vertical" Then
            Width = 1 * Screen.TwipsPerPixelX
            UserControl.PaintPicture LoadResPicture("WinXP\Separator\Back", vbResBitmap), 0, 0, Width, Height
        Else
            Height = 1 * Screen.TwipsPerPixelY
            UserControl.PaintPicture LoadResPicture("WinXP\Separator\Back", vbResBitmap), 0, 0, Width, Height
        End If
    
    Else
    
        If SepStyle = "Vertical" Then
            Width = 2 * Screen.TwipsPerPixelX
            UserControl.PaintPicture LoadResPicture("Win98\Separator\Vertical", vbResBitmap), 0, 0, 2 * Screen.TwipsPerPixelX, Height
        Else
            Height = 2 * Screen.TwipsPerPixelY
            UserControl.PaintPicture LoadResPicture("Win98\Separator\Horizontal", vbResBitmap), 0, 0, Width, 2 * Screen.TwipsPerPixelX
        End If
    
    End If

End Sub

' Paints the toolbar background.
Private Sub pPaintToolbarBackground()

    If InIDE Or Is32Bit = True Then
        lineFix.Visible = False
    Else

        ' Set line fix color
        lineFix.x2 = UserControl.Width
        If g_Appearance = Blue Then
            lineFix.BorderColor = &HD8E9EC
        ElseIf g_Appearance = Green Then
            lineFix.BorderColor = &HD8E9EC
        ElseIf g_Appearance = Silver Then
            lineFix.BorderColor = &HE3DFE0
        ElseIf g_Appearance = Win98 Then
            lineFix.BorderColor = &HFFFFFF
            lineFix2.Visible = True
            lineFix2.x2 = UserControl.Width
        End If
    End If
    
    If g_Appearance <> Win98 Then
        UserControl.PaintPicture LoadResPicture("WinXP\ToolBarBackGround\Back", vbResBitmap), 0, 0, Width, Height
        UserControl.PaintPicture LoadResPicture("WinXP\ToolBarBackGround\Bottom", vbResBitmap), 0, Height - (2 * Screen.TwipsPerPixelY), Width, 2 * Screen.TwipsPerPixelY
    Else
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Back", vbResBitmap), 0, 0, Width, Height
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Bottom", vbResBitmap), 0, Height - (2 * Screen.TwipsPerPixelY), Width, 2 * Screen.TwipsPerPixelY
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Top", vbResBitmap), 0, 0, Width, 2 * Screen.TwipsPerPixelY
        
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\TopLeft", vbResBitmap), 0, 0
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Left", vbResBitmap), 0, 2 * Screen.TwipsPerPixelY, 2 * Screen.TwipsPerPixelY, Height
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\BottomLeft", vbResBitmap), 0, Height - (2 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\Right", vbResBitmap), Width - (2 * Screen.TwipsPerPixelX), 0, 2 * Screen.TwipsPerPixelY, Height
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\TopRight", vbResBitmap), Width - (2 * Screen.TwipsPerPixelX), 0, 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY
        UserControl.PaintPicture LoadResPicture("Win98\ToolBarBackGround\BottomRight", vbResBitmap), Width - (2 * Screen.TwipsPerPixelX), Height - (2 * Screen.TwipsPerPixelY)
    End If

End Sub

' Paints the control border.
Private Sub pPaintControlBorder(BorderType As String)

    ' Set the enabled graphic path
    Dim strEnabled As String
    If UserControl.Enabled = True Then
        strEnabled = "Enabled\"
    Else
        strEnabled = "Disabled\"
    End If
    
    ' Draw the border type
    Select Case BorderType
    
        Case Is = "TopBorder"
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopLeft", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
            
        Case Is = "LeftBorder"
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopLeft", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Left", crBitmap), 0, 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        
        Case Is = "RightBorder"
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        
        Case Is = "BottomBorder"
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, Height - (3 * Screen.TwipsPerPixelY), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\" & strEnabled & "BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
         
    End Select
    
End Sub

' Sets the control type.
Private Sub pSetControlType()

    lineFix.Visible = False
    lineFix2.Visible = False
    
    Select Case m_ControlType
        
        Case Is = PowerButton
            
            m_ControlPath = "PowerButton\"
            Set imgPicture.Picture = LoadResPicture(m_ControlPath & "UnPressed", vbResBitmap)
        
        Case Is = SleepButton
            
            m_ControlPath = "SleepButton\"
            Set imgPicture.Picture = LoadResPicture(m_ControlPath & "UnPressed", vbResBitmap)
 
        Case Is = RestartButton
            
            m_ControlPath = "RestartButton\"
            Set imgPicture.Picture = LoadResPicture(m_ControlPath & "UnPressed", vbResBitmap)
 
        Case Is = GoButton
            
            If g_Appearance <> Win98 Then
                m_ControlPath = "WinXP\GoButton\"
            Else
                m_ControlPath = "Win98\GoButton\"
            End If
            
            UserControl.Picture = LoadResPicture(m_ControlPath & "UnPressed", vbResBitmap)
        
        Case Is = DownArrow
            
            If Enabled = True Then
                pPaintDownArrow "UnPressed"
            Else
                pPaintDownArrow "Disabled"
            End If
        
        Case Is = UpArrow
            
            If Enabled = True Then
                pPaintUpArrow "UnPressed"
            Else
                pPaintUpArrow "Disabled"
            End If
        
        Case Is = LeftArrow
        
            If Enabled = True Then
                pPaintLeftArrow "UnPressed"
            Else
                pPaintLeftArrow "Disabled"
            End If

        Case Is = RightArrow
        
            If Enabled = True Then
                pPaintRightArrow "UnPressed"
            Else
                pPaintRightArrow "Disabled"
            End If
            
        Case Is = ToolBarBackground
        
            If g_Appearance = Win98 Then
                lineFix2.Visible = True
            End If
            
            lineFix.Visible = True
    
        Case Is = SeparatorHorizontal
            lineFix.Visible = True
            
        Case Is = SeparatorVertical
            lineFix.Visible = True
            
        Case Is = TruncateButton
            lineFix.Visible = True
                
    End Select

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    pSetControlType
    UserControl_Resize

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
       
    m_ControlType = PowerButton
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    
    Refresh
    
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
       
    SetDefaultTheme

    m_ControlType = PropBag.ReadProperty("ControlType", PowerButton)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)

    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

    If g_ControlsRefreshed = True Then
        Refresh
    End If
 
End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("ControlType", m_ControlType, PowerButton)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseDown(Button, Shift, x, y)

    If Button = 1 Then
    
        If m_ControlType = PowerButton Or m_ControlType = RestartButton Or m_ControlType = SleepButton Then
            Set imgPicture.Picture = LoadResPicture(m_ControlPath & "Pressed", vbResBitmap)
        ElseIf m_ControlType = GoButton Then
            Set UserControl.Picture = LoadResPicture(m_ControlPath & "Pressed", vbResBitmap)
        ElseIf m_ControlType = TruncateButton Then
            pPaintTruncateButton "Pressed"
        ElseIf m_ControlType = DownArrow Then
            pPaintDownArrow "Pressed"
        ElseIf m_ControlType = UpArrow Then
            pPaintUpArrow "Pressed"
        ElseIf m_ControlType = LeftArrow Then
            pPaintLeftArrow "Pressed"
        ElseIf m_ControlType = RightArrow Then
            pPaintRightArrow "Pressed"
        End If
 
    End If

End Sub

Private Sub UserControl_EnterFocus()

    m_HasFocus = True
    
End Sub

Private Sub UserControl_ExitFocus()

    m_HasFocus = False
    
    If m_ControlType = PowerButton Or m_ControlType = RestartButton Or m_ControlType = SleepButton Then
        UserControl.Picture = LoadResPicture(m_ControlPath & "UnPressed", vbResBitmap)
    End If
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If g_FormHasFocus = True Then
    
        If m_ControlType = GoButton Or m_ControlType = PowerButton Or m_ControlType = RestartButton Or m_ControlType = SleepButton Or m_ControlType = DownArrow Or m_ControlType = UpArrow Or m_ControlType = TruncateButton Or m_ControlType = DownArrow Or m_ControlType = UpArrow Or m_ControlType = LeftArrow Or m_ControlType = RightArrow Then
            
            With UserControl
                
                If GetCapture() = .hwnd Then
                         
                    ' Mouse has left control
                    If x < 0 Or x > .Width Or y < 0 Or y > .Height Then
     
                        Call ReleaseCapture
                                   
                        ' Power \ Restart \ Sleep buttons
                        If m_ControlType = PowerButton Or m_ControlType = RestartButton Or m_ControlType = SleepButton Then
                            
                            If m_HasFocus = False Then
                                Set imgPicture.Picture = LoadResPicture(m_ControlPath & "UnPressed", vbResBitmap)
                            Else
                                Set imgPicture.Picture = LoadResPicture(m_ControlPath & "MouseOver", vbResBitmap)
                            End If
                        
                        ElseIf m_ControlType = GoButton Then
                            UserControl.Picture = LoadResPicture(m_ControlPath & "UnPressed", vbResBitmap)
                        
                        ElseIf m_ControlType = TruncateButton Then
                            pPaintTruncateButton "UnPressed"
                        
                        ElseIf m_ControlType = DownArrow Then
                            pPaintDownArrow "UnPressed"
                        
                        ElseIf m_ControlType = UpArrow Then
                            pPaintUpArrow "UnPressed"
                        
                        ElseIf m_ControlType = LeftArrow Then
                            pPaintLeftArrow "UnPressed"
                        
                        ElseIf m_ControlType = RightArrow Then
                            pPaintRightArrow "UnPressed"
                        End If

                    End If
                
                Else
                        
                    ' Mouse has entered control
                    Call SetCapture(.hwnd)
                    
                    If m_ControlType = PowerButton Or m_ControlType = RestartButton Or m_ControlType = SleepButton Then
                        Set imgPicture.Picture = LoadResPicture(m_ControlPath & "MouseOver", vbResBitmap)
                    
                    ElseIf m_ControlType = GoButton Then
                            UserControl.Picture = LoadResPicture(m_ControlPath & "MouseOver", vbResBitmap)
                                            
                    ElseIf m_ControlType = TruncateButton Then
                        pPaintTruncateButton "MouseOver"
                    
                    ElseIf m_ControlType = DownArrow Then
                        pPaintDownArrow "MouseOver"
                    
                    ElseIf m_ControlType = UpArrow Then
                        pPaintUpArrow "MouseOver"
                    
                    ElseIf m_ControlType = LeftArrow Then
                        pPaintLeftArrow "MouseOver"
                    
                    ElseIf m_ControlType = RightArrow Then
                        pPaintRightArrow "MouseOver"
                    End If
                
                    RaiseEvent MouseMove(Button, Shift, x, y)
    
                End If
            
            End With
        
        Else

            RaiseEvent MouseMove(Button, Shift, x, y)
        
        End If
    
    End If

End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    Select Case m_ControlType
        
        Case Is = PowerButton
            Width = 33 * Screen.TwipsPerPixelX
            Height = 33 * Screen.TwipsPerPixelY
        
        Case Is = SleepButton
            Width = 33 * Screen.TwipsPerPixelX
            Height = 33 * Screen.TwipsPerPixelY
        
        Case Is = SleepButton
            Width = 33 * Screen.TwipsPerPixelX
            Height = 33 * Screen.TwipsPerPixelY
        
        Case Is = GoButton
            Width = 51 * Screen.TwipsPerPixelX
            Height = 26 * Screen.TwipsPerPixelY
        
        Case Is = TopBorder
            pPaintControlBorder "TopBorder"
            Height = 3 * Screen.TwipsPerPixelY
        
        Case Is = BottomBorder
            pPaintControlBorder "BottomBorder"
            Height = 3 * Screen.TwipsPerPixelY
        
        Case Is = LeftBorder
            pPaintControlBorder "LeftBorder"
            Width = 3 * Screen.TwipsPerPixelX
        
        Case Is = RightBorder
            pPaintControlBorder "RightBorder"
            Width = 3 * Screen.TwipsPerPixelX
        
        Case Is = ToolBarBackground
            pPaintToolbarBackground
        
        Case Is = SeparatorHorizontal
            pPaintSeparator "Horizontal"
        
        Case Is = SeparatorVertical
            pPaintSeparator "Vertical"
        
        Case Is = TruncateButton
            pPaintTruncateButton "UnPressed"
        
        Case Is = ComboArrow
            pPaintComboArrow "UnPressed"
        
        Case Is = DownArrow
            Width = 17 * Screen.TwipsPerPixelX
            Height = 17 * Screen.TwipsPerPixelY
        
        Case Is = UpArrow
            Width = 17 * Screen.TwipsPerPixelX
            Height = 17 * Screen.TwipsPerPixelY
        
        Case Is = LeftArrow
            Width = 17 * Screen.TwipsPerPixelX
            Height = 17 * Screen.TwipsPerPixelY
        
        Case Is = RightArrow
            Width = 17 * Screen.TwipsPerPixelX
            Height = 17 * Screen.TwipsPerPixelY
    
    End Select

err:
    Exit Sub

End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ControlType() As ExtaControlsEnum
Attribute ControlType.VB_Description = "Returns/sets the type of control."
    ControlType = m_ControlType
End Property

Public Property Let ControlType(ByVal New_ControlType As ExtaControlsEnum)
    
    m_ControlType = New_ControlType
    
    pSetControlType
    
    Select Case m_ControlType
    
        Case Is = SeparatorHorizontal
            Width = 60 * Screen.TwipsPerPixelX
        
        Case Is = SeparatorVertical
            Height = 60 * Screen.TwipsPerPixelY
        
        Case Is = ToolBarBackground
            Height = 25 * Screen.TwipsPerPixelY
            Width = 180 * Screen.TwipsPerPixelX
        
        Case Is = TruncateButton
            Width = 20 * Screen.TwipsPerPixelX
            Height = 40 * Screen.TwipsPerPixelY
        
        Case Is = LeftBorder
            Height = 60 * Screen.TwipsPerPixelY
        
        Case Is = RightBorder
            Height = 60 * Screen.TwipsPerPixelY
        
        Case Is = TopBorder
            Width = 60 * Screen.TwipsPerPixelY
        
        Case Is = BottomBorder
            Width = 60 * Screen.TwipsPerPixelY
        
    End Select
    
    UserControl_Resize
    
    PropertyChanged "ControlType"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Refresh
    UserControl_Resize
    PropertyChanged "Enabled"
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
Attribute ToolTipStyle.VB_Description = "Returns/sets the style of the tooltip i.e Standad or Balloon.\r\n"
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
Attribute ToolTipTitle.VB_Description = "Returns/sets the title displayed in the tooltip.\r\n"
    ToolTipTitle = m_ToolTipTitle
End Property

Public Property Let ToolTipTitle(ByVal New_ToolTipTitle As String)
    m_ToolTipTitle = New_ToolTipTitle
    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    PropertyChanged "ToolTipTitle"
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
Attribute hwnd.VB_UserMemId = -515
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

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

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub
