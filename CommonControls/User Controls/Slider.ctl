VERSION 5.00
Begin VB.UserControl Slider 
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   HasDC           =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   1005
   ToolboxBitmap   =   "Slider.ctx":0000
   Begin CommonControls.MaskBox2 ScrollButton 
      Height          =   330
      Left            =   15
      TabIndex        =   0
      Top             =   1410
      Width           =   570
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin CommonControls.MaskBox Background 
      Height          =   2865
      Left            =   30
      Top             =   15
      Width           =   810
      _ExtentX        =   0
      _ExtentY        =   0
   End
End
Attribute VB_Name = "Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - Slider.ctl                                    '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_FocusRectangle As Boolean
Dim m_LargeChange As Long
Dim m_Max As Long
Dim m_Min As Long
Dim m_Orientation As OrientationEnum
Dim m_Value As Long
Dim m_HasFocus As Boolean
Dim m_Initialized As Boolean
Dim m_State As String
Dim m_MouseX As Long
Dim m_MouseY As Long
Dim m_Sliding As Boolean
Dim m_LastPosition As Long
Dim m_ValueChanged As Boolean
Dim m_OrientationSet As Boolean

' Events
Event MouseLeave()
Event Change()
Attribute Change.VB_MemberFlags = "200"
Event Scroll()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the appearance property.
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

    Select Case m_Orientation
    
    Case Is = Vertical
        Set Background.Picture = PictureFromResource(g_ResourceLib.hModule, "slider\vertical\background", crBitmap)
    Case Horizontal
        Set Background.Picture = PictureFromResource(g_ResourceLib.hModule, "slider\horizontal\background", crBitmap)
        ScrollButton.Top = 0
    End Select

    pSetGraphic "UnPressed"

End Sub

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If UserControl.Enabled = True Then
        pSetGraphic "UnPressed"
    Else
        pSetGraphic "Disabled"
    End If

End Sub

' Sets the graphics for the control based on the orientation.
Private Sub pSetGraphic(sState As String)

    m_State = sState
    
    If Orientation = Vertical Then
        Set ScrollButton.Picture = PictureFromResource(g_ResourceLib.hModule, "Slider\Vertical\" & sState, crBitmap)
    Else
        Set ScrollButton.Picture = PictureFromResource(g_ResourceLib.hModule, "Slider\Horizontal\" & sState, crBitmap)
    End If

End Sub

' Reposition the thumb based on the orientation of the slider and the value.
Private Sub pChangeValue()
    
    Dim MinY As Single
    Dim MaxY As Single
    Dim MinX As Single
    Dim MaxX As Single
    Dim NewPosn As Long

    With UserControl
    
        If Orientation = Vertical Then
        
            MinY = Min
            MaxY = Background.Height - ScrollButton.Height
            
            NewPosn = (Value - Min) / (Max - Min) * (MaxY - MinY) + MinY
            
            If NewPosn >= Background.Top Then
                ScrollButton.Top = NewPosn
            Else
                ScrollButton.Top = Background.Top
            End If
             
            ScrollButton.Left = 6 * Screen.TwipsPerPixelX
        
              
        Else
        
            MinX = Min
            MaxX = Background.Width - ScrollButton.Width
                 
            NewPosn = (Value - Min) / (Max - Min) * (MaxX - MinX) + MinX
                  
            If NewPosn >= Background.Left Then
                ScrollButton.Left = NewPosn
            Else
                ScrollButton.Left = Background.Left
            End If
                  
            'Move the Thumb based on the Alignment
            ScrollButton.Top = 3 * Screen.TwipsPerPixelY
            
        End If
        
    End With

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates the control.
Public Sub Refresh()

    Select Case m_Orientation
        Case Is = Vertical
            Background.Left = 2 * Screen.TwipsPerPixelX
            Background.Top = 2 * Screen.TwipsPerPixelY
            Set Background.Picture = PictureFromResource(g_ResourceLib.hModule, "slider\vertical\background", crBitmap)
        Case Horizontal
            Background.Top = 4 * Screen.TwipsPerPixelY
            Set Background.Picture = PictureFromResource(g_ResourceLib.hModule, "slider\horizontal\background", crBitmap)
            ScrollButton.Top = 0
                 
    End Select
    
    pSetAppearance
    pChangeValue
    pSetEnabled
    UserControl_Resize

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()

    SetDefaultTheme

    m_FocusRectangle = True
    m_LargeChange = 20
    m_Max = 100
    m_Min = 0
    m_Orientation = Vertical
    m_Value = 0
    m_OrientationSet = True

    m_Initialized = True

    Height = 110 * Screen.TwipsPerPixelY

    Refresh
        
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    m_OrientationSet = False
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FocusRectangle = PropBag.ReadProperty("FocusRectangle", True)
    m_LargeChange = PropBag.ReadProperty("LargeChange", 20)
    m_Max = PropBag.ReadProperty("Max", 100)
    m_Min = PropBag.ReadProperty("Min", 1)
    m_Orientation = PropBag.ReadProperty("Orientation", Vertical)
    m_Value = PropBag.ReadProperty("Value", 1)
    m_OrientationSet = True
    m_Initialized = True

    If g_ControlsRefreshed = True Then
        Refresh
    End If

End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("LargeChange", m_LargeChange, 20)
    Call PropBag.WriteProperty("Max", m_Max, 100)
    Call PropBag.WriteProperty("Min", m_Min, 1)
    Call PropBag.WriteProperty("Orientation", m_Orientation, Vertical)
    Call PropBag.WriteProperty("Value", m_Value, 1)
    Call PropBag.WriteProperty("FocusRectangle", m_FocusRectangle, True)

End Sub

Private Sub UserControl_EnterFocus()

    m_HasFocus = True
    
    If m_HasFocus = True And m_FocusRectangle = True Then
        Dim UsrRect As RECT
        Call SetRect(UsrRect, 0, 0, (Width / Screen.TwipsPerPixelX), (Height / Screen.TwipsPerPixelY))
        Call DrawFocusRect(UserControl.hdc, UsrRect)
        Call SetRectEmpty(UsrRect)
    End If

End Sub

Private Sub UserControl_ExitFocus()

    m_HasFocus = False
    UserControl.Cls

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Orientation = Vertical Then
      If y > ScrollButton.Top Then
         Value = m_Value + LargeChange
      Else
         Value = m_Value - LargeChange
      End If
   
   Else
      If x > ScrollButton.Left Then
         Value = m_Value + LargeChange
      Else
         Value = m_Value - LargeChange
      End If
   
   End If

End Sub

Private Sub ScrollButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        If Orientation = Vertical Then
        
            '  Set variables for Thumb Scrolling
            m_MouseY = y
            m_Sliding = True
            m_LastPosition = ScrollButton.Top + y
        
        Else
            m_MouseX = x
            m_Sliding = True
            m_LastPosition = ScrollButton.Left + x
        
        End If

        pSetGraphic "Pressed"
    
    End If 'If button =1

End Sub

Private Sub ScrollButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim NewPosn As Long
    Dim MaxY As Long
    Dim MinY As Long
    Dim MaxX As Long
    Dim MinX As Long

    If m_State <> "Pressed" Then
    
        With ScrollButton
        
            If GetCapture() = .hwnd Then
             
                ' Mouse has left control
                If x < 0 Or x > .Width Or y < 0 Or y > .Height Then
                    Call ReleaseCapture
                    pSetGraphic "UnPressed"
                End If
            
            Else
                ' Mouse has entered control
                Call SetCapture(.hwnd)
                pSetGraphic "MouseOver"
              
            End If
        End With
    
    End If
    
    ' Exit if we are not Sliding
    If Not m_Sliding Then Exit Sub
        
        ' Make scroll button appear pressed
        If m_State <> "Pressed" Then
            pSetGraphic "Pressed"
        End If

        ' Determine the Orientation
        If Orientation = Vertical Then
      
            MinY = 0
            MaxY = Background.Height - (ScrollButton.Height)
        
            'Determine the Position of the Thumb
            NewPosn = ScrollButton.Top + y - m_MouseY
            
            ' Limit NewPosn to Min/Max values
            If NewPosn >= MaxY Then
                NewPosn = MaxY
            ElseIf NewPosn <= MinY Then
                NewPosn = MinY
            End If
             
            If NewPosn >= Background.Top Then
                ScrollButton.Top = NewPosn
            Else
                ScrollButton.Top = Background.Top
            End If

          'Calculate the new Value based on the position of the Thumb between the Up and Down Buttons
          Value = ((ScrollButton.Top - MinY) / (MaxY - MinY)) * (Max - Min) + Min
                   
          If Value <= 2 Then
            Value = 0
          End If
                               
          'Trigger the Event
          RaiseEvent Scroll
          
          'Save position
          m_LastPosition = NewPosn
                   
          'Set the variable so we know if we should trigger
          'the Change Event on MouseUp
          m_ValueChanged = True
         
        
        Else ' Horizontal scrolling
         
            MinX = 0
            MaxX = Background.Width - ScrollButton.Width
                    
            'Determine the Position of the Thumb
            NewPosn = ScrollButton.Left + x - m_MouseX
                 
                ' Limit NewPosn to Min/Max values
                If NewPosn >= MaxX Then
                NewPosn = MaxX
                End If
                 
                If NewPosn <= MinX Then
                NewPosn = MinX
                End If
                 
            If NewPosn >= Background.Left Then
            ScrollButton.Left = NewPosn
            Else
            ScrollButton.Left = Background.Left
      
      End If
      
      ' Calculate the new Value based on the position of the Thumb between the Left and Right Buttons
      Value = ((ScrollButton.Left - MinX) / (MaxX - MinX)) * (Max - Min) + Min
      
      ' Trigger the Event
      RaiseEvent Scroll
               
      ' Save position
      m_LastPosition = NewPosn
               
      ' Set the variable so we know if we should trigger the Change Event on MouseUp
      m_ValueChanged = True
     
    End If


End Sub

Private Sub ScrollButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    pSetGraphic "UnPressed"
    m_Sliding = False
    If m_ValueChanged Then RaiseEvent Change

End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    If m_OrientationSet = True Then
    
        If m_Orientation = Vertical Then
            Width = 34 * Screen.TwipsPerPixelX
            Background.Height = UserControl.Height - (4 * Screen.TwipsPerPixelY)
        Else
            Height = 32 * Screen.TwipsPerPixelY
            Background.Width = UserControl.Width - (5 * Screen.TwipsPerPixelX)
        End If
    
    End If
    
    If m_Initialized = True Then
        pChangeValue
    End If

err:
    Exit Sub

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
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
'MemberInfo=0,0,0,0
Public Property Get FocusRectangle() As Boolean
    FocusRectangle = m_FocusRectangle
End Property

Public Property Let FocusRectangle(ByVal New_FocusRectangle As Boolean)
    m_FocusRectangle = New_FocusRectangle
    PropertyChanged "FocusRectangle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get LargeChange() As Long
Attribute LargeChange.VB_Description = "Returns/sets the amount of change to Value when user clicks the background of the slider."
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Long)
    m_LargeChange = New_LargeChange
    PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum Value of the slider."
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Private Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum Value of the slider."
    Min = m_Min
End Property

Private Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Orientation() As OrientationEnum
Attribute Orientation.VB_Description = "Returns/sets the type of slider i.e Horizontal or Vertical."
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationEnum)
    
    m_Orientation = New_Orientation
    
    If New_Orientation = Vertical Then
        Height = Width
    Else
        Width = Height
    End If
    
    pSetAppearance
    pChangeValue
    UserControl_Resize
    
    PropertyChanged "Orientation"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the Value of the slider."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value

    ' Make sure we are within the given range
    If New_Value >= Min And New_Value <= Max Then
        m_Value = New_Value
    ElseIf New_Value < Min Then
        m_Value = Min
    ElseIf New_Value > Max Then
        m_Value = Max
    End If
      
    ' Set position of scroll button
    If m_Sliding = False Then
        pChangeValue
    End If
    
    RaiseEvent Change
    
    PropertyChanged "Value"
    
End Property
