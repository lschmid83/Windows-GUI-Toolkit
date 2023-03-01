VERSION 5.00
Begin VB.UserControl ToolbarButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   CanGetFocus     =   0   'False
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3090
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   Enabled         =   0   'False
   FillColor       =   &H00FFFFFF&
   ScaleHeight     =   1560
   ScaleWidth      =   3090
   ToolboxBitmap   =   "ToolbarButton.ctx":0000
   Begin VB.Line lineFix2 
      BorderColor     =   &H00929D9D&
      X1              =   150
      X2              =   2730
      Y1              =   15
      Y2              =   15
   End
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
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   705
      TabIndex        =   0
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "ToolbarButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - ToolbarButton.ctl                             '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_ButtonType As ButtonStyleEnum
Dim m_AutoPressedForeColor As Boolean
Dim m_CheckButton As Boolean
Dim m_Value As Boolean
Dim m_MaskColor As OLE_COLOR
Dim m_Picture As Picture
Dim m_DisabledPicture As Picture
Dim m_DownPicture As Picture
Dim m_MouseOverPicture As Picture
Dim m_PictureAlign As PictureAlignEnum
Dim m_UseMaskColor As Boolean
Dim m_ForeColor As OLE_COLOR
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
Dim m_State As String
Dim m_CaptionX As Integer
Dim m_CaptionY As Integer
Dim m_PictureX As Integer
Dim m_PictureY As Integer

' Events
Event DropDownClick()
Event DropDownMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event DropDownMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event DropDownMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
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
Event MouseLeave()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If UserControl.Enabled = True Then
        
        Set imgPicture.Picture = m_Picture
        lblCaption.ForeColor = m_ForeColor
        
        If m_CheckButton = True Then
            If m_Value = True Then
                pPaintComponent m_ButtonType, "Pressed"
            Else
                pPaintComponent m_ButtonType, "UnPressed"
            End If
        Else
            pPaintComponent m_ButtonType, "UnPressed"
        End If
    
    Else
  
        Set imgPicture.Picture = DisabledPicture
    
        If g_Appearance <> Win98 Then
            lblCaption.ForeColor = &H92A1A1
        Else
            lblCaption.ForeColor = &H808080
        End If
        
        pPaintComponent m_ButtonType, "Disabled"
    End If

End Sub

' Paints the component.
Private Sub pPaintComponent(ByVal New_ButtonStyle As ButtonStyleEnum, New_Value As String)

    m_State = New_Value
    
    UserControl.Cls
    
    ' Set line fix color
    lineFix.X1 = 0
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
    
    If InIDE Or Is32Bit = True Then
        lineFix.Visible = False
        lineFix2.Visible = False
    Else
        lineFix.Visible = True
        lineFix2.Visible = False
    End If

    ' Draw the standard button
    If New_ButtonStyle = StandardButton Then
    
        If m_State <> "UnPressed" And m_State <> "Disabled" Then
        
            ' Draw the background
            If g_Appearance <> Win98 Then
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, Width - (8 * Screen.TwipsPerPixelX), Height - (8 * Screen.TwipsPerPixelY)
            Else
            
                If m_State <> "HasFocus" Then
                    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, Width - (8 * Screen.TwipsPerPixelX), Height - (8 * Screen.TwipsPerPixelY)
                Else
                    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY
                
                End If
            End If
        
            ' Draw the border
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopLeft", crBitmap), 0, 0
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 0
            
            If InIDE Or Is32Bit = True Then
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY), Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
            Else
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY), Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
            End If
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Left", crBitmap), 0, 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Right", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Top", crBitmap), 4 * Screen.TwipsPerPixelY, 0, Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                       
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 0, 0, Width, Height
        End If
        
    ' Draw the dropdown button
    ElseIf New_ButtonStyle = DropDown Then
    
        If m_State <> "UnPressed" And m_State <> "Disabled" Then
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, Width - (8 * Screen.TwipsPerPixelX), Height - (8 * Screen.TwipsPerPixelY)
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopLeft", crBitmap), 0, 0
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 0
            
            If InIDE Or Is32Bit = True Then
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY), Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
            Else
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY), Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
            End If
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Left", crBitmap), 0, 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Right", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Top", crBitmap), 4 * Screen.TwipsPerPixelY, 0, Width - (8 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                   
            If m_State <> "Pressed" Then
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Arrow", crBitmap), Width - (9 * Screen.TwipsPerPixelX), (Height / 2) - (1 * Screen.TwipsPerPixelY)
            Else
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Arrow", crBitmap), Width - (8 * Screen.TwipsPerPixelX), (Height / 2)
            End If
            
        
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 0, 0, Width, Height
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Arrow", crBitmap), Width - (9 * Screen.TwipsPerPixelY), (Height / 2) - (1 * Screen.TwipsPerPixelY)
        End If
    
    ' Draw the dual button
    ElseIf New_ButtonStyle = Dual Then
    
        If m_State <> "UnPressed" And m_State <> "Pressed" And m_State <> "DropDownPressed" And m_State <> "Disabled" Then
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, Width - (8 * Screen.TwipsPerPixelX), Height - (8 * Screen.TwipsPerPixelY)
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopLeft", crBitmap), 0, 0
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownTop", crBitmap), Width - (14 * Screen.TwipsPerPixelX), 0
               
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Left", crBitmap), 0, 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Top", crBitmap), 4 * Screen.TwipsPerPixelY, 0, Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
            
            If InIDE Or Is32Bit = True Then
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY), Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownLeft", crBitmap), Width - 210, (4 * Screen.TwipsPerPixelY), (4 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownBottom", crBitmap), Width - (14 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
            Else
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY), Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownLeft", crBitmap), Width - 210, (4 * Screen.TwipsPerPixelY), (4 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownBottom", crBitmap), Width - (14 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
            End If
            
            
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Arrow", crBitmap), Width - (9 * Screen.TwipsPerPixelX), (Height / 2) - (1 * Screen.TwipsPerPixelY)
            
        ElseIf m_State = "UnPressed" Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 0, 0, Width, Height
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Arrow", crBitmap), Width - (9 * Screen.TwipsPerPixelY), (Height / 2) - (1 * Screen.TwipsPerPixelY)
    
        ElseIf m_State = "DropDownPressed" Then
    
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\MouseOver\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, Width - (8 * Screen.TwipsPerPixelX), Height - (8 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\Back", crBitmap), Width - (10 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, Width, Height
        
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\MouseOver\TopLeft", crBitmap), 0, 0
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\DropDownTop", crBitmap), Width - (14 * Screen.TwipsPerPixelX), 0
            
            If InIDE Or Is32Bit = True Then
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\MouseOver\BottomLeft", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\DropDownRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\MouseOver\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY), Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\DropDownLeft", crBitmap), Width - 210, (4 * Screen.TwipsPerPixelY), (4 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\DropDownBottom", crBitmap), Width - (14 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
        
            Else
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\MouseOver\BottomLeft", crBitmap), 0, Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\DropDownRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\MouseOver\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY), Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\DropDownLeft", crBitmap), Width - 210, (4 * Screen.TwipsPerPixelY), (4 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\DropDownBottom", crBitmap), Width - (14 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
        
            End If
                      
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\MouseOver\Left", crBitmap), 0, 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\MouseOver\Top", crBitmap), 4 * Screen.TwipsPerPixelY, 0, Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Pressed\Arrow", crBitmap), Width - (8 * Screen.TwipsPerPixelX), Height / 2
           
    
        ElseIf m_State = "Disabled" Then

            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Disabled\Back", crBitmap), 0, 0, Width, Height
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\Disabled\Arrow", crBitmap), Width - (10 * Screen.TwipsPerPixelX), (Height / 2) - (1 * Screen.TwipsPerPixelY)
       
        ElseIf m_State = "Pressed" Then
        
            If g_Appearance <> Win98 Then
                lineFix2.X1 = 50
                lineFix2.x2 = UserControl.Width - 100
                lineFix2.BorderColor = &H929D9D
            Else
                lineFix2.X1 = 0
                lineFix2.x2 = UserControl.Width
                lineFix2.BorderColor = &H808080
            End If
        
            If InIDE Or Is32Bit = True Then
                
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, Width - (8 * Screen.TwipsPerPixelX), Height - (8 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopLeft", crBitmap), 0, 0
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownTop", crBitmap), Width - (14 * Screen.TwipsPerPixelX), 0
                            
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (4 * Screen.TwipsPerPixelY), Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownBottom", crBitmap), Width - (14 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownLeft", crBitmap), Width - 210, (4 * Screen.TwipsPerPixelY), (4 * Screen.TwipsPerPixelX), Height - (4 * Screen.TwipsPerPixelY)
            
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Top", crBitmap), 4 * Screen.TwipsPerPixelY, 0, Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
        
            Else
                lineFix2.Visible = True
                
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Back", crBitmap), 4 * Screen.TwipsPerPixelX, 4 * Screen.TwipsPerPixelY, Width - (8 * Screen.TwipsPerPixelX), Height - (8 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\TopLeft", crBitmap), 0, 1 * Screen.TwipsPerPixelY
                
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownTop", crBitmap), Width - (14 * Screen.TwipsPerPixelX), 0

                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\BottomLeft", crBitmap), 0, Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownRight", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Bottom", crBitmap), 4 * Screen.TwipsPerPixelY, Height - (5 * Screen.TwipsPerPixelY), Width - (18 * Screen.TwipsPerPixelX), 4 * Screen.TwipsPerPixelY
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownBottom", crBitmap), Width - (14 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\DropDownLeft", crBitmap), Width - 210, (4 * Screen.TwipsPerPixelY), (4 * Screen.TwipsPerPixelX), Height - (5 * Screen.TwipsPerPixelY)
        
                UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Top", crBitmap), 4 * Screen.TwipsPerPixelY, 0, Width - (18 * Screen.TwipsPerPixelX), 5 * Screen.TwipsPerPixelY
        
            End If
            
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Left", crBitmap), 0, 4 * Screen.TwipsPerPixelY, 4 * Screen.TwipsPerPixelY, Height - (8 * Screen.TwipsPerPixelY)
               
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ToolBarButton\" & m_State & "\Arrow", crBitmap), Width - (8 * Screen.TwipsPerPixelX), (Height / 2) + (1 * Screen.TwipsPerPixelY)
           
        End If
    
    
    End If

    ' Indent caption
    If m_State = "Pressed" Then
        imgPicture.Left = imgPicture.Left + 15
        imgPicture.Top = imgPicture.Top + 15
        lblCaption.Left = lblCaption.Left + 15
        lblCaption.Top = lblCaption.Top + 15
    
    Else
        imgPicture.Left = m_PictureX
        imgPicture.Top = m_PictureY
        lblCaption.Left = m_CaptionX
        lblCaption.Top = m_CaptionY
    
    End If
    
    
    If UserControl.Enabled = True Then
    
        If m_AutoPressedForeColor = True Then
        
            If g_Appearance <> Win98 Then
            
                If m_State = "Pressed" Then
                    lblCaption.ForeColor = &HFFFFFF
                Else
                    lblCaption.ForeColor = m_ForeColor
                End If
            
            End If
        
        End If
    
    End If
   

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

Public Sub Refresh()

    pSetEnabled
    UserControl_Resize

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
  
    SetDefaultTheme

    m_AutoPressedForeColor = True
    m_MaskColor = &HFF00FF
    Set m_DisabledPicture = LoadPicture("")
    Set m_DownPicture = LoadPicture("")
    Set m_MouseOverPicture = LoadPicture("")
    m_ForeColor = 0
    m_ButtonType = StandardButton
    m_CheckButton = False
    m_Value = False
    m_PictureAlign = ButtonCenter
    m_UseMaskColor = True
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    Caption = Ambient.DisplayName
    
    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY
    
    Refresh

End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    lblCaption.Alignment = PropBag.ReadProperty("Alignment", 2)
    lblCaption.Caption = PropBag.ReadProperty("Caption", "Label1")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_MaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
    Set m_DisabledPicture = PropBag.ReadProperty("DisabledPicture", Nothing)
    Set m_DownPicture = PropBag.ReadProperty("DownPicture", Nothing)
    Set m_MouseOverPicture = PropBag.ReadProperty("MouseOverPicture", Nothing)
    m_CheckButton = PropBag.ReadProperty("CheckButton", False)
    m_Value = PropBag.ReadProperty("Value", False)
    m_PictureAlign = PropBag.ReadProperty("PictureAlign", ButtonLeft)
    m_UseMaskColor = PropBag.ReadProperty("UseMaskColor", True)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_AutoPressedForeColor = PropBag.ReadProperty("AutoPressedForeColor", True)
    m_ButtonType = PropBag.ReadProperty("ButtonType", 0)

    If m_UseMaskColor = True Then
        imgPicture.SetMaskColor m_MaskColor
    End If
    Set Picture = PropBag.ReadProperty("Picture", Nothing)

    If g_ControlsRefreshed = True Then
        Refresh
    End If


End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Label1")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, &HFF00FF)
    Call PropBag.WriteProperty("DisabledPicture", m_DisabledPicture, Nothing)
    Call PropBag.WriteProperty("DownPicture", m_DownPicture, Nothing)
    Call PropBag.WriteProperty("MouseOverPicture", m_MouseOverPicture, Nothing)
    Call PropBag.WriteProperty("PictureAlign", m_PictureAlign, ButtonLeft)
    Call PropBag.WriteProperty("UseMaskColor", m_UseMaskColor, True)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("CheckButton", m_CheckButton, False)
    Call PropBag.WriteProperty("Value", m_Value, False)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("Alignment", lblCaption.Alignment, 2)
    Call PropBag.WriteProperty("AutoPressedForeColor", m_AutoPressedForeColor, True)
    Call PropBag.WriteProperty("ButtonType", m_ButtonType, 0)
    
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseDown(Button, Shift, x, y)

    If m_ButtonType <> Dual Then
    
        If Button = 1 Then
        
            If m_CheckButton = False Then
            
                pPaintComponent m_ButtonType, "Pressed"
                Set UserControl.MaskPicture = DownPicture
                
                If UserControl.MaskPicture <> Empty Then
                    Set imgPicture.Picture = DownPicture
                    UserControl_Resize
                End If
  
            Else
            
                pPaintComponent m_ButtonType, "Pressed"
                If m_Value = True Then
                    Value = False
                Else
                    Value = True
                End If

            End If
        
        End If

    Else
    
        If x > Width - (14 * Screen.TwipsPerPixelX) Then
            pPaintComponent m_ButtonType, "DropDownPressed"
        Else
            pPaintComponent m_ButtonType, "Pressed"
        End If
        
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
    With UserControl
        
        If GetCapture() = .hwnd Then
     
            ' Mouse has left control
            If x < 0 Or x > .Width Or y < 0 Or y > .Height Then

                Call ReleaseCapture

                Set UserControl.MaskPicture = m_Picture
                Set imgPicture.Picture = m_Picture
      
                If m_CheckButton = False Then
                    pPaintComponent m_ButtonType, "UnPressed"
                Else
                    If m_Value = True Then
                        pPaintComponent m_ButtonType, "HasFocus"
                    Else
    
                        pPaintComponent m_ButtonType, "UnPressed"
                    End If
                End If
                
                RaiseEvent MouseLeave
   
            End If
        
        Else
    
            ' Mouse has entered control
            Call SetCapture(.hwnd)
            
            Set UserControl.MaskPicture = MouseOverPicture
            
            If UserControl.MaskPicture <> Empty Then
                Set imgPicture.Picture = MouseOverPicture
                UserControl_Resize
            End If
            
            If m_CheckButton = False Then
                If m_State <> "MouseOver" And m_State <> "Pressed" Then
                    pPaintComponent m_ButtonType, "MouseOver"
                End If
            Else
                
                If m_Value = True Then
                    If m_State <> "Pressed" Then
                        pPaintComponent m_ButtonType, "HasFocusMouseOver"
                    End If
                Else
                    pPaintComponent m_ButtonType, "MouseOver"
                End If
                
            End If
            
            RaiseEvent MouseMove(Button, Shift, x, y)
        
        End If
    
    End With
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        If m_CheckButton = False Then
            
            Set UserControl.MaskPicture = Picture
            Set imgPicture.Picture = Picture
              
            RaiseEvent MouseUp(Button, Shift, x, y)
        
        End If
    
    End If
    
    pPaintComponent m_ButtonType, "UnPressed"

End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseUp(Button, Shift, x, y)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call UserControl_MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    ' Calculate image and caption position
    If imgPicture.Picture <> Empty Then

        If m_PictureAlign = ButtonCenter Then
             
             If lblCaption.Caption = "" Then
                imgPicture.Top = ((UserControl.Height / 2) - (imgPicture.Height / 2)) - 20
             Else
                imgPicture.Top = ((UserControl.Height / 2) - (imgPicture.Height / 2)) - 110
             End If
             
             imgPicture.Left = (UserControl.Width / 2) - (imgPicture.Width / 2)
             
             lblCaption.Top = UserControl.Height - (lblCaption.Height + 100)
             lblCaption.Left = 80
             lblCaption.Width = UserControl.Width - 160
        
        Else
        
            lblCaption.Alignment = 0
            
            imgPicture.Top = (UserControl.Height / 2) - (imgPicture.Height / 2)
            imgPicture.Left = 80
            
            lblCaption.Top = (UserControl.Height / 2) - (lblCaption.Height / 2)
            lblCaption.Left = imgPicture.Left + imgPicture.Width + 90
            lblCaption.Width = UserControl.Width - 160
        
        End If
    
    Else 'Button has no picture
    
        lblCaption.Top = (Height / 2) - (lblCaption.Height / 2)
        lblCaption.Left = 6 * Screen.TwipsPerPixelX
        lblCaption.Width = Width - (12 * Screen.TwipsPerPixelX)
        
    End If
    
    m_PictureX = imgPicture.Left
    m_PictureY = imgPicture.Top
    m_CaptionX = lblCaption.Left
    m_CaptionY = lblCaption.Top
        
    If m_State <> "" Then
        pPaintComponent m_ButtonType, m_State
    End If

err:
    Exit Sub

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

Property Let FontBold(Bold As Boolean)
    lblCaption.FontBold = Bold
End Property

Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = lblCaption.FontBold
End Property

Property Let FontItalic(Italic As Boolean)
    lblCaption.FontItalic = Italic
End Property

Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = lblCaption.FontItalic
End Property

Property Let FontName(Name As String)
    lblCaption.FontName = Name
End Property

Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
    FontName = lblCaption.FontName
End Property

Property Let FontSize(Size As Long)
    lblCaption.FontSize = Size
End Property

Property Get FontSize() As Long
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = lblCaption.FontSize
End Property

Property Let FontStrikeThru(StrikeThru As Boolean)
    lblCaption.FontStrikeThru = StrikeThru
End Property

Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_MemberFlags = "400"
    FontStrikeThru = lblCaption.FontStrikeThru
End Property

Property Let FontUnderline(UnderLine As Boolean)
    lblCaption.FontUnderline = UnderLine
End Property

Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = lblCaption.FontUnderline
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoPressedForeColor() As Boolean
Attribute AutoPressedForeColor.VB_Description = "Returns/sets whether the buttons fore color is automatically set when the button is pressed."
    AutoPressedForeColor = m_AutoPressedForeColor
End Property

Public Property Let AutoPressedForeColor(ByVal New_AutoPressedForeColor As Boolean)
    m_AutoPressedForeColor = New_AutoPressedForeColor
    PropertyChanged "AutoPressedForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Retruns/sets the alignment of the controls text."
    Alignment = lblCaption.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    lblCaption.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get CheckButton() As Boolean
Attribute CheckButton.VB_Description = "Returns/sets whether the button only has two states i.e Checked/UnChecked."
    CheckButton = m_CheckButton
End Property

Public Property Let CheckButton(ByVal New_CheckButton As Boolean)

    If m_ButtonType = StandardButton Then
        m_CheckButton = New_CheckButton
    
        If New_CheckButton = True Then
            Value = Value
        End If
    
    Else
        m_CheckButton = False
    End If
    
    PropertyChanged "CheckButton"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    pSetEnabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    If UserControl.Enabled = True Then
        lblCaption.ForeColor() = New_ForeColor
    End If
    PropertyChanged "ForeColor"
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
'MappingInfo=imgPicture,imgPicture,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    Set imgPicture.Picture = New_Picture
    UserControl_Resize
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    m_MaskColor = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get DisabledPicture() As Picture
Attribute DisabledPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is disabled."
    Set DisabledPicture = m_DisabledPicture
End Property

Public Property Set DisabledPicture(ByVal New_DisabledPicture As Picture)
    Set m_DisabledPicture = New_DisabledPicture
    PropertyChanged "DisabledPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get DownPicture() As Picture
Attribute DownPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is pressed."
    Set DownPicture = m_DownPicture
End Property

Public Property Set DownPicture(ByVal New_DownPicture As Picture)
    Set m_DownPicture = New_DownPicture
    PropertyChanged "DownPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get MouseOverPicture() As Picture
Attribute MouseOverPicture.VB_Description = "Returns/sets a graphic to be displayed when the mouse is over the button."
    Set MouseOverPicture = m_MouseOverPicture
End Property

Public Property Set MouseOverPicture(ByVal New_MouseOverPicture As Picture)
    Set m_MouseOverPicture = New_MouseOverPicture
    PropertyChanged "MouseOverPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get PictureAlign() As PictureAlignEnum
Attribute PictureAlign.VB_Description = "Returns/sets the position of the picture on the button."
    PictureAlign = m_PictureAlign
End Property

Public Property Let PictureAlign(ByVal New_PictureAlign As PictureAlignEnum)
    m_PictureAlign = New_PictureAlign
    UserControl_Resize
    PropertyChanged "PictureAlign"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets whether the color assigned in the MaskColor property is used as the transparent color in the buttons picture."
    UseMaskColor = m_UseMaskColor
End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)
    m_UseMaskColor = New_UseMaskColor
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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the state of the button if CheckButton = True."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    
    m_Value = New_Value
    
    If Enabled = True Then
    
        If m_CheckButton = True Then
        m_Value = New_Value
        
            If m_Value = True Then
            
                If m_State <> "Pressed" Then
                pPaintComponent m_ButtonType, "Pressed"
                End If
            
            Else
            
                If m_State <> "UnPressed" Then
                pPaintComponent m_ButtonType, "UnPressed"
                End If
                
            End If
        
        End If
    
    End If

    PropertyChanged "Value"
    
End Property

''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
Public Property Get buttontype() As ButtonStyleEnum
    buttontype = m_ButtonType
End Property

Public Property Let buttontype(ByVal New_ButtonType As ButtonStyleEnum)
    m_ButtonType = New_ButtonType
    pPaintComponent m_ButtonType, "UnPressed"
    PropertyChanged "ButtonType"
End Property
