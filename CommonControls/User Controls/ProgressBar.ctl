VERSION 5.00
Begin VB.UserControl ProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   HasDC           =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   2970
   ToolboxBitmap   =   "ProgressBar.ctx":0000
   Begin VB.Line lineFix 
      X1              =   0
      X2              =   2580
      Y1              =   0
      Y2              =   0
   End
   Begin CommonControls.MaskBox imgProgress 
      Height          =   225
      Left            =   135
      Top             =   60
      Width           =   1365
      _ExtentX        =   0
      _ExtentY        =   0
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - ProgressBar.ctl                               '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Member variables
Dim m_Min As Long
Dim m_Max As Long
Dim m_Size As SizeEnum2
Dim m_Value As Long
Dim m_Height As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private functions

' Paints the component
Private Sub pPaintComponent()

    ' Set the height of the progressbar
    Dim sSize As String
    If m_Size = Large Then
        sSize = "Large\"
        m_Height = 23 * Screen.TwipsPerPixelY
    ElseIf m_Size = Small Then
        sSize = "Small\"
        m_Height = 15 * Screen.TwipsPerPixelY
    End If
        
    ' Draw graphics
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
    
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ProgressBar\" & sSize & "Center", crBitmap), 3 * Screen.TwipsPerPixelX, 1 * Screen.TwipsPerPixelY, Width - 6 * Screen.TwipsPerPixelX, m_Height - (2 * Screen.TwipsPerPixelY)
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ProgressBar\" & sSize & "Left", crBitmap), 0, 1 * Screen.TwipsPerPixelY
    Set imgProgress.Picture = PictureFromResource(g_ResourceLib.hModule, "ProgressBar\" & sSize & "Progress", crBitmap)
    imgProgress.Left = 4 * Screen.TwipsPerPixelX
    If InIDE Or Is32Bit = True Then
        imgProgress.Top = 4 * Screen.TwipsPerPixelY
    Else
        imgProgress.Top = 5 * Screen.TwipsPerPixelY
    End If
    
    UserControl.PaintPicture PictureFromResource(g_ResourceLib.hModule, "ProgressBar\" & sSize & "Right", crBitmap), Width - (4 * Screen.TwipsPerPixelX), 1 * Screen.TwipsPerPixelY
    
    If m_Max > 0 Then
        UserControl.ScaleWidth = m_Max
    End If
    
    imgProgress.Width = m_Value
    
    UserControl.ScaleMode = 1
    If imgProgress.Width - (8 * Screen.TwipsPerPixelX) > 0 Then
        imgProgress.Width = imgProgress.Width - (8 * Screen.TwipsPerPixelX)
        imgProgress.Visible = True
    Else
        imgProgress.Visible = False
    End If
    Height = m_Height

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public functions

Public Sub Refresh()
    pPaintComponent
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Events

' Occurs when a new instance of an object is created.
Private Sub UserControl_InitProperties()
    
    ' Initialize default theme
    SetDefaultTheme
    
    ' Initialize default properties
    m_Size = Small
    m_Value = 0
    m_Min = 0
    m_Max = 100

    Width = 110 * Screen.TwipsPerPixelX
    
    ' Update control
    Refresh

End Sub

' Occurs when loading an old instance of an object that has a saved state.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
       
    ' Initialize default theme
    SetDefaultTheme
    
    ' Read properties
    m_Size = PropBag.ReadProperty("Size", Small)
    m_Value = PropBag.ReadProperty("Value", 1)
    m_Min = PropBag.ReadProperty("Min", 0)
    m_Max = PropBag.ReadProperty("Max", 100)
    
    ' Update control
    If g_ControlsRefreshed = True Then
        Refresh
    End If

End Sub

' Occurs when an instance of an object is to be saved.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' Write properties to storage
    Call PropBag.WriteProperty("Size", m_Size, Small)
    Call PropBag.WriteProperty("Value", m_Value, 1)
    Call PropBag.WriteProperty("Min", m_Min, 0)
    Call PropBag.WriteProperty("Max", m_Max, 100)
    
End Sub

Private Sub UserControl_Resize()

On Error GoTo err

    Height = m_Height
    pPaintComponent

err:
    Exit Sub

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Size() As SizeEnum2
Attribute Size.VB_Description = "Returns/sets the size of the progressbar i.e Small or Large."
    Size = m_Size
End Property

Public Property Let Size(ByVal New_Size As SizeEnum2)
    m_Size = New_Size
    Refresh
    PropertyChanged "Size"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "Returns/sets the value of the progressbar."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    If m_Value >= m_Min And m_Value <= m_Max Then
        pPaintComponent
    Else
        m_Value = m_Min
        pPaintComponent
    End If
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Private Property Get Min() As Long
    Min = m_Min
End Property

Private Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    pPaintComponent
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Max() As Long
Attribute Max.VB_Description = "Returns/sets the maximum value of the progress bar."
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    pPaintComponent
    PropertyChanged "Max"
End Property
