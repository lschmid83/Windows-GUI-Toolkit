VERSION 5.00
Begin VB.UserControl ScrollBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   HasDC           =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   315
   ToolboxBitmap   =   "ScrollBar.ctx":0000
   Begin VB.PictureBox Down 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2415
      Width           =   315
   End
   Begin VB.PictureBox ScrollButton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1140
      Width           =   315
   End
   Begin VB.PictureBox UP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   315
   End
End
Attribute VB_Name = "ScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - ScrollBar.ctl                                 '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member Variables
Dim m_Min As Long
Dim m_Max As Long
Dim m_Orientation As Variant
Dim m_SmallChange As Long
Dim m_LargeChange As Long
Dim m_Value As Long
Dim m_State As String
Dim m_Initialized As Boolean
Dim m_MouseX As Long
Dim m_MouseY As Long
Dim m_Sliding As Boolean
Dim m_LastPosition As Long
Dim m_ValueChanged As Boolean

' Events
Event Change()
Attribute Change.VB_MemberFlags = "200"
Event Scroll()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the orientation property.
Private Sub pSetAppearance()

    UserControl.Cls

    Select Case m_Orientation
    
        Case Vertical
                     
            Up.Picture = PictureFromResource(g_ResourceLib.hModule, "scrollbar\vertical\up\unpressed", crBitmap)
            Down.Picture = PictureFromResource(g_ResourceLib.hModule, "scrollbar\vertical\down\unpressed", crBitmap)
            ScrollButton.Picture = PictureFromResource(g_ResourceLib.hModule, "scrollbar\vertical\button\unpressed", crBitmap)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "scrollbar\vertical\background", crBitmap), 0, 0, Width, Height
            
            Up.Left = 0
            Up.Top = 0
            ScrollButton.Left = 0
            
        Case Horizontal
            Down.Picture = PictureFromResource(g_ResourceLib.hModule, "scrollbar\horizontal\left\unpressed", crBitmap)
            Up.Picture = PictureFromResource(g_ResourceLib.hModule, "scrollbar\horizontal\right\unpressed", crBitmap)
            ScrollButton.Picture = PictureFromResource(g_ResourceLib.hModule, "scrollbar\horizontal\button\unpressed", crBitmap)
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "scrollbar\horizontal\background", crBitmap), 0, 0, Width, Height
                   
            Down.Left = 0
            Down.Top = 0
            ScrollButton.Top = 0
    
    End Select

End Sub

' Sets the style of the control based on the enabled property.
Private Sub pSetEnabled()

    If UserControl.Enabled = True Then
        ScrollButton.Visible = True
        pSetGraphic "Up", "UnPressed"
        pSetGraphic "Down", "UnPressed"
    Else
        ScrollButton.Visible = False
        pSetGraphic "Up", "Disabled"
        pSetGraphic "Down", "Disabled"
    
    End If

End Sub

' Sets the button graphics based on the orientation and state.
Private Sub pSetGraphic(sButton As String, sState As String)

    m_State = sState
    
    If sButton = "Up" Then
        
        If Orientation = Vertical Then
            Up.Picture = PictureFromResource(g_ResourceLib.hModule, "Scrollbar\Vertical\Up\" & sState, crBitmap)
        Else
            Up.Picture = PictureFromResource(g_ResourceLib.hModule, "Scrollbar\Horizontal\Right\" & sState, crBitmap)
        End If
        
    ElseIf sButton = "Down" Then
        
        If Orientation = Vertical Then
            Down.Picture = PictureFromResource(g_ResourceLib.hModule, "Scrollbar\Vertical\Down\" & sState, crBitmap)
        Else
            Down.Picture = PictureFromResource(g_ResourceLib.hModule, "Scrollbar\Horizontal\Left\" & sState, crBitmap)
        End If
    
    ElseIf sButton = "ScrollButton" Then
        
        If Orientation = Vertical Then
            ScrollButton.Picture = PictureFromResource(g_ResourceLib.hModule, "Scrollbar\Vertical\Button\" & sState, crBitmap)
        Else
            ScrollButton.Picture = PictureFromResource(g_ResourceLib.hModule, "Scrollbar\Horizontal\Button\" & sState, crBitmap)
        End If
        
    End If

End Sub

' Reposition the thumb based on the orientation of the slider and the value.
Private Sub pChangeValue()

    Dim MinY As Single
    Dim MaxY As Single
    Dim MinX As Single
    Dim MaxX As Single
    
    ' Get what Percent Value is based on the Min / Max Values
    ' We then multiple this percent by the distance between the Top / Bottom buttons,
    ' taking into account the width/height of the thumb.
    
    With UserControl
    
        If m_Orientation = Vertical Then
        
            MinY = Up.Height
            MaxY = Down.Top - ScrollButton.Height
        
            ScrollButton.Top = (Value - Min) / (Max - Min) * (MaxY - MinY) + MinY
            ScrollButton.Left = 0
      
        Else
        
            MinX = Down.Width
            MaxX = Up.Left - ScrollButton.Width
                 
            ScrollButton.Left = (Value - Min) / (Max - Min) * (MaxX - MinX) + MinX
            ScrollButton.Top = 0
        
        End If
        
    End With

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

' Updates control.
Public Sub Refresh()

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

    m_LargeChange = 20
    m_Max = 100
    m_Min = 1
    m_Orientation = Vertical
    m_SmallChange = 3
    m_Value = 1
    m_Initialized = True
    Extender.TabStop = False
    
    Height = 110 * Screen.TwipsPerPixelY

    Refresh
    
End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_LargeChange = PropBag.ReadProperty("LargeChange", 20)
    m_Max = PropBag.ReadProperty("Max", 100)
    m_Min = PropBag.ReadProperty("Min", 1)
    m_Orientation = PropBag.ReadProperty("Orientation", Vertical)
    m_SmallChange = PropBag.ReadProperty("SmallChange", 3)
    m_Value = PropBag.ReadProperty("Value", 1)
    m_Initialized = True

    If g_ControlsRefreshed = True Then
        Refresh
    End If

End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("LargeChange", m_LargeChange, 20)
    Call PropBag.WriteProperty("Max", m_Max, 100)
    Call PropBag.WriteProperty("Min", m_Min, 1)
    Call PropBag.WriteProperty("Orientation", m_Orientation, Vertical)
    Call PropBag.WriteProperty("SmallChange", m_SmallChange, 3)
    Call PropBag.WriteProperty("Value", m_Value, 1)

End Sub

Private Sub Up_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        pSetGraphic "Up", "Pressed"
        
        ' Keep adding to value while button is pressed
        Dim currenttime!
        Dim RepeatDelayTime!
        RepeatDelayTime! = 0
            
        Do Until (GetAsyncKeyState(&H1) = 0)
        
            If Orientation = Vertical Then
                Value = m_Value - SmallChange
            Else
                Value = m_Value + SmallChange
            End If
        
            UserControl.Refresh
        
            currenttime = Timer
            
            Do Until Timer > currenttime + RepeatDelayTime!
                If GetAsyncKeyState(&H1) = 0 Then
                    Exit Do
                End If
            Loop
        
        Loop
        
    End If

End Sub


Private Sub Up_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    With Up
        
        If GetCapture() = .hwnd Then
         
            ' Mouse has entered control
            If x < 0 Or x > .Width Or y < 0 Or y > .Height Then

                Call ReleaseCapture
                
                pSetGraphic "Up", "UnPressed"
          
            End If
    
        Else
        
            ' Mouse has left control
            Call SetCapture(.hwnd)
            pSetGraphic "Up", "MouseOver"
         
        End If
       
    End With

End Sub


Private Sub Down_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        pSetGraphic "Down", "Pressed"
        
        ' Keep adding to value while button is pressed
        Dim currenttime!
        Dim RepeatDelayTime!
        RepeatDelayTime! = 0
            
        Do Until (GetAsyncKeyState(&H1) = 0)
        
            If Orientation = Vertical Then
                Value = m_Value + SmallChange
            Else
                Value = m_Value + -SmallChange
            End If
        
            UserControl.Refresh
        
            currenttime = Timer
            Do Until Timer > currenttime + RepeatDelayTime!
                If GetAsyncKeyState(&H1) = 0 Then
                    Exit Do
                End If
            Loop
        
        Loop
    
    End If

End Sub

Private Sub Down_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    With Down
    
        If GetCapture() = .hwnd Then
                     
            ' Mouse has left control
            If x < 0 Or x > .Width Or y < 0 Or y > .Height Then

                Call ReleaseCapture
                pSetGraphic "Down", "UnPressed"
           
            End If
    
        Else
            
            ' Mouse has entered control
            Call SetCapture(.hwnd)
            pSetGraphic "Down", "MouseOver"
         
        End If
    
    End With

End Sub

Private Sub Down_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pSetGraphic "Down", "UnPressed"
End Sub

Private Sub ScrollButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
        
        ' Set variables for thumb scrolling
        If Orientation = Vertical Then
           
            m_MouseY = y
            m_Sliding = True
            m_LastPosition = ScrollButton.Top + y
        
        Else
            
            m_MouseX = x
            m_Sliding = True
            m_LastPosition = ScrollButton.Left + x
        
        End If
        
        pSetGraphic "ScrollButton", "Pressed"
    
    End If

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
                    pSetGraphic "ScrollButton", "UnPressed"
                End If
   
            Else
            
                ' Mouse has entered control
                Call SetCapture(.hwnd)
                pSetGraphic "ScrollButton", "MouseOver"
              
            End If
        End With
    
    End If
    
    ' Exit if we are not Sliding - This value is set to
    ' True in the MouseDown and False in the MouseUp
    If Not m_Sliding Then Exit Sub
    
    ' Determine the Orientation
    If Orientation = Vertical Then
   
       ' We Add the .Top value to the Height to take
       ' in account the ButtonsVisible property
       MinY = Up.Top + Up.Height
       MaxY = Down.Top - (ScrollButton.Height)
    
         ' Determine the Position of the Thumb
         NewPosn = ScrollButton.Top + y - m_MouseY
         If NewPosn >= MaxY Then
             NewPosn = MaxY
         End If
         If NewPosn <= MinY Then
             NewPosn = MinY
         End If
         
         ' Don't need to do anything if we haven't moved
         If NewPosn <> m_LastPosition Then
             
             ' Move the Thumb
             ScrollButton.Move ScrollButton.Left, NewPosn
             
             ' Calculate the new Value based on the position of the Thumb between the Up and Down Buttons
             Value = ((ScrollButton.Top - MinY) / (MaxY - MinY)) * (Max - Min) + Min
             
             ' Trigger the Event
             RaiseEvent Scroll
             
             ' Save position
             m_LastPosition = NewPosn
             
             ' Set the variable so we know if we should trigger the Change Event on MouseUp
             m_ValueChanged = True
             
         End If
    
    Else  ' Horizontal scrolling
       
       ' We Add the .Left value to the Height to take
       ' in account the ButtonsVisible property
       MinX = Down.Left + Down.Width
       MaxX = Up.Left - ScrollButton.Width
            
         ' Determine the Position of the Thumb
         NewPosn = ScrollButton.Left + x - m_MouseX
         If NewPosn >= MaxX Then
             NewPosn = MaxX
         End If
         If NewPosn <= MinX Then
             NewPosn = MinX
         End If
         
         ' Don't need to do anything if we haven't moved
         If NewPosn <> m_LastPosition Then
             
             ' Move the Thumb
             ScrollButton.Move NewPosn, ScrollButton.Top
             
             ' Calculate the new Value based on the position of the Thumb between the Left and Right Buttons
             Value = ((ScrollButton.Left - MinX) / (MaxX - MinX)) * (Max - Min) + Min
             
             ' Trigger the Event
             RaiseEvent Scroll
             
             ' Save position
             m_LastPosition = NewPosn
             
             ' Set the variable so we know if we should trigger the Change Event on MouseUp
             m_ValueChanged = True
         End If
    
    End If

End Sub


Private Sub ScrollButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    pSetGraphic "ScrollButton", "UnPressed"
    m_Sliding = False
    If m_ValueChanged Then RaiseEvent Change
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    UserControl.Cls
    
    If m_Orientation = Vertical Then
        
        Width = Up.Width
            
        ' Draw the background
        If g_Appearance <> Win98 Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "scrollbar\vertical\background", crBitmap), 0, 0, Width, Height
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "scrollbar\vertical\background", crBitmap), 0, 0
        End If
        
        Down.Move 0, (UserControl.ScaleHeight - Down.Height)
        
    Else
        
        If m_Initialized = True Then
            Height = Down.Height
        End If
        
        ' Draw the background
        If g_Appearance <> Win98 Then
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "scrollbar\horizontal\background", crBitmap), 0, 0, Width, Height
        Else
            UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "scrollbar\horizontal\background", crBitmap), 0, 0
        
        End If
        
        Up.Move UserControl.ScaleWidth - (Up.Width), 0
                     
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
'MemberInfo=8,0,0,0
Public Property Get LargeChange() As Long
Attribute LargeChange.VB_Description = "Returns/sets the amount of change to Value when the user clicks the scrollbar area."
    LargeChange = m_LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Long)
    m_LargeChange = New_LargeChange
    PropertyChanged "LargeChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum Value for the scrollbar."
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Private Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum Value of the scrollbar."
    Min = m_Min
End Property

Private Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Orientation() As OrientationEnum
Attribute Orientation.VB_Description = "Returns/sets the type of scrollbar i.e Horizontal or Vertical."
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
    
    PropertyChanged "Orientation"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get SmallChange() As Long
Attribute SmallChange.VB_Description = "Returns/sets the amount of change to Value when one of the scroll arrows are pressed."
    SmallChange = m_SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Long)
    m_SmallChange = New_SmallChange
    PropertyChanged "SmallChange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the value of the scrollbar."
    Value = m_Value
End Property

Public Property Let Value(ByVal sState As Long)
    
    m_Value = sState
    
    ' Make sure we are within the given range
    If sState >= Min And sState <= Max Then
        m_Value = sState
    ElseIf sState < Min Then
        m_Value = Min
    ElseIf sState > Max Then
        m_Value = Max
    End If
    
    ' Set position of scroll button
    If m_Sliding = False Then
        pChangeValue
    End If
    
    RaiseEvent Change
    
    PropertyChanged "Value"
    
End Property
