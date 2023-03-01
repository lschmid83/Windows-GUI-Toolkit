VERSION 5.00
Begin VB.UserControl SpinEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ScaleHeight     =   585
   ScaleWidth      =   2685
   ToolboxBitmap   =   "SpinEdit.ctx":0000
   Begin CommonControls.MaskBox2 Down 
      Height          =   240
      Left            =   1005
      TabIndex        =   1
      Top             =   240
      Width           =   360
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin CommonControls.MaskBox2 Up 
      Height          =   165
      Left            =   990
      TabIndex        =   2
      Top             =   15
      Width           =   390
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin VB.TextBox txtSpinEdit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "SpinEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - SpinEdit.ctl                                  '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_AllowNegative As Boolean
Dim m_Change As Long
Dim m_Max As Long
Dim m_Min As Long
Dim m_Value As Long
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_State As String

' Events
Event Change()
Event Click() 'MappingInfo=txtSpinEdit,txtSpinEdit,-1,Click
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtSpinEdit,txtSpinEdit,-1,KeyDown
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtSpinEdit,txtSpinEdit,-1,KeyPress
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtSpinEdit,txtSpinEdit,-1,KeyUp
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtSpinEdit,txtSpinEdit,-1,MouseDown
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtSpinEdit,txtSpinEdit,-1,MouseMove
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=txtSpinEdit,txtSpinEdit,-1,MouseUp
Attribute MouseUp.VB_UserMemId = -607

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Sets the style of the control based on the appearance property.
Private Sub pSetAppearance()

    txtSpinEdit.Top = 4 * Screen.TwipsPerPixelY
 
    If UserControl.Enabled = True Then
        pSetGraphic "Up", "UnPressed"
        pSetGraphic "Down", "UnPressed"
        txtSpinEdit.ForeColor = ForeColor
        txtSpinEdit.BackColor = BackColor
    Else
        pSetGraphic "Up", "Disabled"
        pSetGraphic "Down", "Disabled"
    
        If g_Appearance <> Win98 Then
            txtSpinEdit.ForeColor = &H92A1A1
            txtSpinEdit.BackColor = &H8000000F
        Else
            txtSpinEdit.ForeColor = &H808080
            txtSpinEdit.BackColor = &HFFFFFF
        End If
    
    End If

End Sub

' Paints the control border.
Private Sub pPaintBorder()

    UserControl.Cls
    
    If UserControl.Enabled = True Then
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Enabled\TopLeft", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Enabled\Left", crBitmap), 0, 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Enabled\Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Enabled\BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Enabled\Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Enabled\TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Enabled\Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, Height - (3 * Screen.TwipsPerPixelY), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Enabled\BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    
    Else
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Disabled\TopLeft", crBitmap), 0, 0
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Disabled\Left", crBitmap), 0, 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Disabled\Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelX, Height - (6 * Screen.TwipsPerPixelY)
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Disabled\BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Disabled\Top", crBitmap), 3 * Screen.TwipsPerPixelX, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Disabled\TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0, 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Disabled\Bottom", crBitmap), 3 * Screen.TwipsPerPixelX, Height - (3 * Screen.TwipsPerPixelY), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
        UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ControlBorder\Disabled\BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelY), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY
    
    End If

End Sub

' Sets the graphics for the control based the button type.
Private Sub pSetGraphic(New_Button As String, New_Value As String)
    
    m_State = New_Value
   
    If New_Button = "Up" Then
        Set Up.Picture = PictureFromResource(g_ResourceLib.hModule, "SpinEdit\Up\" & New_Value, crBitmap)
    Else
        Set Down.Picture = PictureFromResource(g_ResourceLib.hModule, "SpinEdit\Down\" & New_Value, crBitmap)
    End If

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

Public Sub Refresh()

    pSetAppearance
    Value = m_Value
    UserControl_Resize

End Sub






'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtSpinEdit,txtSpinEdit,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = txtSpinEdit.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtSpinEdit.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtSpinEdit,txtSpinEdit,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    txtSpinEdit.BackColor() = New_BackColor
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
    pSetAppearance
    pPaintBorder
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtSpinEdit,txtSpinEdit,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    txtSpinEdit.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtSpinEdit,txtSpinEdit,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = txtSpinEdit.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtSpinEdit.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtSpinEdit,txtSpinEdit,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = txtSpinEdit.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    txtSpinEdit.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AllowNegative() As Boolean
Attribute AllowNegative.VB_Description = "Returns/sets whether negative numbers can be displayed in the control."
    AllowNegative = m_AllowNegative
End Property

Public Property Let AllowNegative(ByVal New_AllowNegative As Boolean)
    m_AllowNegative = New_AllowNegative
    PropertyChanged "AllowNegative"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Change() As Long
Attribute Change.VB_Description = "Returns/sets the change to Value when the up/down arrows are pressed."
Attribute Change.VB_MemberFlags = "200"
    Change = m_Change
End Property

Public Property Let Change(ByVal New_Change As Long)
    m_Change = New_Change
    PropertyChanged "Change"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/sets whether numbers can be entered into the spinedit."
    Locked = txtSpinEdit.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtSpinEdit.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum value."
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets the minimum value."
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the value of the control."
Attribute Value.VB_UserMemId = -518
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)

    If Enabled = True Then
        m_Value = New_Value
        If Value > Max Then
            txtSpinEdit.Text = Max
        ElseIf Value < Min Then
            txtSpinEdit.Text = Min
        Else
            txtSpinEdit.Text = Value
        End If
    End If

    PropertyChanged "Value"
End Property

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme
    
    m_ForeColor = 0
    m_BackColor = &HFFFFFF
    m_AllowNegative = False
    m_Change = 1
    m_Max = 100
    m_Min = 1
    m_Value = 1

    Width = 110 * Screen.TwipsPerPixelX

    Refresh

End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    txtSpinEdit.Alignment = PropBag.ReadProperty("Alignment", 1)
    m_BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    txtSpinEdit.Locked = PropBag.ReadProperty("Locked", False)
    m_ForeColor = PropBag.ReadProperty("ForeColor", &H0&)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtSpinEdit.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_AllowNegative = PropBag.ReadProperty("AllowNegative", False)
    m_Change = PropBag.ReadProperty("Change", 1)
    m_Max = PropBag.ReadProperty("Max", 100)
    m_Min = PropBag.ReadProperty("Min", 1)
    m_Value = PropBag.ReadProperty("Value", 1)

    If g_ControlsRefreshed = True Then
        Refresh
    End If

End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", txtSpinEdit.Alignment, 1)
    Call PropBag.WriteProperty("BackColor", m_BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("Locked", txtSpinEdit.Locked, False)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H0&)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtSpinEdit.MousePointer, 0)
    Call PropBag.WriteProperty("AllowNegative", m_AllowNegative, False)
    Call PropBag.WriteProperty("Change", m_Change, 1)
    Call PropBag.WriteProperty("Max", m_Max, 100)
    Call PropBag.WriteProperty("Min", m_Min, 1)
    Call PropBag.WriteProperty("Value", m_Value, 1)

End Sub

Private Sub Down_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        pSetGraphic "Down", "Pressed"
        
        ' Keep adding to value while button is pressed
        Dim currenttime!
        Dim RepeatDelayTime!
        RepeatDelayTime! = 0.1
            
        If AllowNegative = False Then
            If Value <= 0 Then
                Value = 0
                Exit Sub
            End If
        End If
            
        If Value <= Min Then
            Value = Min
            Exit Sub
        End If
            
        Do Until (GetAsyncKeyState(&H1) = 0)
        
            If AllowNegative = False Then
                If Value <= 0 Then
                    Value = 0
                    Exit Sub
                End If
            End If
        
            If Value <= Min Then
                Value = Min
                Exit Sub
            End If
        
            Value = Value - Change
            txtSpinEdit.Refresh
        
            currenttime = Timer
            
            RaiseEvent Change
            Do Until Timer > currenttime + RepeatDelayTime!
                If GetAsyncKeyState(&H1) = 0 Then
                    Exit Do
                End If
            Loop
        
        Loop
    
    End If

End Sub

Private Sub Down_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If g_FormHasFocus = True Then
    
        With Down
            If GetCapture() = .hwnd Then
             
                ' Mouse has left control
                If x < 0 Or x > .Width Or y < 0 Or y > .Height Then
                    Call ReleaseCapture
                    pSetGraphic "Down", "UnPressed"
                End If
             
            Else
            
                Call SetCapture(.hwnd)
                pSetGraphic "Down", "MouseOver"
                
            End If
        
        End With
    
    End If

End Sub

Private Sub txtSpinEdit_Change()

    If IsNumeric(txtSpinEdit.Text) <> True Then
        txtSpinEdit.Text = Value
    Else
        Value = txtSpinEdit.Text
    End If

End Sub

Private Sub Up_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
    
        pSetGraphic "Up", "Pressed"
        
        ' Keep adding to value while button is pressed
        Dim currenttime!
        Dim RepeatDelayTime!
        RepeatDelayTime! = 0.1
            
        If Value >= Max Then
            Value = Max
            Exit Sub
        End If
            
        Do Until (GetAsyncKeyState(&H1) = 0)
        
            If Value >= Max Then
                Value = Max
                Exit Sub
            End If
            
            Value = Value + Change
            txtSpinEdit.Refresh
            RaiseEvent Change
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

    If g_FormHasFocus = True Then
    
      With Up
      
          If GetCapture() = .hwnd Then
           
              ' Mouse has left control
              If x < 0 Or x > .Width Or y < 0 Or y > .Height Then
                  Call ReleaseCapture
                  pSetGraphic "Up", "UnPressed"
              End If
          
          Else
              
              ' Mouse has entered control
              Call SetCapture(.hwnd)
              pSetGraphic "Up", "MouseOver"
    
          End If
      End With
    
    End If


End Sub

Private Sub txtSpinEdit_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    pPaintBorder
    
    Up.Left = Width - (18 * Screen.TwipsPerPixelX)
    Down.Left = Width - (18 * Screen.TwipsPerPixelX)
    
    Up.Top = 3 * Screen.TwipsPerPixelY
    Down.Top = Up.Top + Up.Height
 
    txtSpinEdit.Left = 3 * Screen.TwipsPerPixelX
    txtSpinEdit.Width = UserControl.Width - (23 * Screen.TwipsPerPixelX)
    
    Height = 22 * Screen.TwipsPerPixelY

err:
    Exit Sub

End Sub
