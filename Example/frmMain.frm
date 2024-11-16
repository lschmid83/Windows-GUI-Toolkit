VERSION 5.00
Object = "{05931B16-A732-4810-88BE-423C9A7A76B5}#1.0#0"; "CommonControls.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMain"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin CommonControls.ComboBox ComboBox1 
      Height          =   315
      Left            =   2415
      TabIndex        =   15
      Top             =   1650
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommonControls.PictureBox PictureBox2 
      Height          =   795
      Left            =   7350
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4290
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1402
      BackColor       =   16777215
      AutoSize        =   -1  'True
      UseMaskColor    =   -1  'True
   End
   Begin CommonControls.SpinEdit SpinEdit1 
      Height          =   330
      Left            =   6045
      TabIndex        =   13
      Top             =   1650
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   582
   End
   Begin CommonControls.ExtraControls ExtraControls5 
      Height          =   435
      Left            =   1140
      Top             =   4875
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   767
      ControlType     =   8
   End
   Begin CommonControls.CommandButton CommandButton1 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   615
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Caption         =   "Change the theme"
   End
   Begin CommonControls.ToolbarButton ToolbarButton1 
      Height          =   615
      Left            =   240
      Top             =   1365
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      Caption         =   "ToolbarButton1"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlign    =   1
      Alignment       =   0
      ButtonType      =   3
   End
   Begin CommonControls.ListBox ListBox1 
      Height          =   1065
      Left            =   2280
      TabIndex        =   11
      Top             =   4260
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   1879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommonControls.FileList FileList1 
      Height          =   1650
      Left            =   5400
      TabIndex        =   10
      Top             =   2370
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   2910
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommonControls.FolderList FolderList1 
      Height          =   1695
      Left            =   2250
      TabIndex        =   9
      Top             =   2370
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2990
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      Path            =   "c:\"
   End
   Begin CommonControls.DriveBox DriveBox1 
      Height          =   315
      Left            =   255
      TabIndex        =   8
      Top             =   2355
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommonControls.CheckBox CheckBox2 
      Height          =   225
      Left            =   6045
      TabIndex        =   7
      Top             =   1305
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   397
      BackColor       =   14215660
      Caption         =   "CheckBox2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommonControls.CheckBox CheckBox1 
      Height          =   225
      Left            =   6045
      TabIndex        =   6
      Top             =   930
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   397
      BackColor       =   14215660
      Caption         =   "CheckBox1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin CommonControls.Slider Slider1 
      Height          =   1530
      Left            =   4800
      TabIndex        =   5
      Top             =   585
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   2699
      BackColor       =   14215660
      Min             =   0
      Value           =   0
   End
   Begin CommonControls.ExtraControls ExtraControls4 
      Height          =   390
      Left            =   270
      Top             =   4875
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
      ControlType     =   7
   End
   Begin CommonControls.ExtraControls ExtraControls3 
      Height          =   495
      Left            =   1440
      Top             =   4260
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ControlType     =   6
      ToolTipCaption  =   "Restart"
   End
   Begin CommonControls.ExtraControls ExtraControls2 
      Height          =   495
      Left            =   855
      Top             =   4260
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ControlType     =   5
      ToolTipCaption  =   "Sleep"
   End
   Begin CommonControls.ExtraControls ExtraControls1 
      Height          =   495
      Left            =   270
      Top             =   4260
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ToolTipCaption  =   "Power"
   End
   Begin CommonControls.OptionButton OptionButton2 
      Height          =   240
      Left            =   2400
      TabIndex        =   4
      Top             =   1305
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   423
      BackColor       =   14215660
      ForeColor       =   0
      Caption         =   "OptionButton2"
      ToolTipStyle    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommonControls.OptionButton OptionButton1 
      Height          =   240
      Left            =   2400
      TabIndex        =   3
      Top             =   930
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   423
      BackColor       =   14215660
      ForeColor       =   0
      Value           =   -1  'True
      ToolTipStyle    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommonControls.Frame Frame1 
      Height          =   1545
      Index           =   0
      Left            =   2235
      Top             =   570
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   2725
      BackColor       =   14215660
      ForeColor       =   16711680
      Caption         =   "Frame1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CommonControls.TextBox TextBox1 
      Height          =   1200
      Left            =   225
      TabIndex        =   2
      Top             =   2880
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   2117
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "TextBox1"
   End
   Begin CommonControls.ScrollBar ScrollBar1 
      Height          =   4980
      Left            =   8400
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   570
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8784
   End
   Begin CommonControls.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      Top             =   5565
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   714
      Caption         =   "StatusBar1"
      ResizeHandle    =   0   'False
   End
   Begin CommonControls.Border Border2 
      Align           =   4  'Align Right
      Height          =   5115
      Left            =   8655
      Top             =   450
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   9022
      BorderType      =   2
   End
   Begin CommonControls.TitleBar TitleBar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   794
      BorderStyle     =   1
      Buttons         =   2
      Caption         =   "Windows GUI Toolkit Example v1.0"
   End
   Begin CommonControls.Border Border1 
      Align           =   3  'Align Left
      Height          =   5115
      Left            =   0
      Top             =   450
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   9022
   End
   Begin CommonControls.Frame Frame1 
      Height          =   1545
      Index           =   1
      Left            =   5880
      Top             =   570
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   2725
      BackColor       =   14215660
      ForeColor       =   16711680
      Caption         =   "Frame2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblSliderValue 
      BackColor       =   &H00D8E9EC&
      Caption         =   "0"
      Height          =   300
      Left            =   5460
      TabIndex        =   12
      Top             =   1875
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ToolBarButtons As Integer

Private Sub CommandButton1_Click()

    If TitleBar1.Appearance = Blue Then
        TitleBar1.Appearance = Green
    ElseIf TitleBar1.Appearance = Green Then
        TitleBar1.Appearance = Silver
    ElseIf TitleBar1.Appearance = Silver Then
        TitleBar1.Appearance = Win98
    ElseIf TitleBar1.Appearance = Win98 Then
        TitleBar1.Appearance = Blue
    End If
    
    Call Form_Resize
  
End Sub

Private Sub Form_Load()

    FolderList1.Path = "c:\"
    FileList1.Path = "c:\"
  
    For i = 10 To 1 Step -1
        ComboBox1.AddItem ("Item" & i)
    Next i
    ComboBox1.ListIndex = 0

    For i = 10 To 1 Step -1
        ListBox1.AddItem ("Item" & i)
    Next i

    frmMain.Height = 6000

    ' Restore theme
    SetAppearance (GetSetting(App.EXEName, "appearance", "theme", 1))
    
    Call Form_Resize

End Sub

Private Sub Form_Activate()
    CommandButton1.SetFocus
End Sub

Private Sub Form_Resize()
    
    ' Resize toolbar
    If TitleBar1.Appearance <> Win98 Then
        ScrollBar1.Top = 480
        ScrollBar1.left = frmMain.Width - 330
        ScrollBar1.Height = 5100
    Else
        ScrollBar1.Top = 340
        ScrollBar1.left = frmMain.Width - 320
        ScrollBar1.Height = 5310
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Save theme
    SaveSetting App.EXEName, "appearance", "theme", TitleBar1.Appearance

End Sub

Private Sub MenuBar1_Click(MenuIndex As Integer, ItemIndex As Integer)

    ' Set theme
    If MenuIndex = 4 Then
        SetAppearance (ItemIndex)
    End If
    
    ' Exit
    If MenuIndex = 1 And ItemIndex = 9 Then
        TitleBar1.ExitApp
    End If
    
End Sub

Private Sub SetAppearance(Appearance As Integer)

        ' Initialize toolbar position and size based on theme
        lblSliderValue.BackColor = &HD8E9EC

        If Appearance = 1 Then
            TitleBar1.Appearance = Blue
        ElseIf Appearance = 2 Then
            TitleBar1.Appearance = Green
        ElseIf Appearance = 3 Then
            TitleBar1.Appearance = Silver
            lblSliderValue.BackColor = &HE3DFE0
        ElseIf Appearance = 4 Then
            TitleBar1.Appearance = Win98
            lblSliderValue.BackColor = &HC8D0D4
        End If
        
        Form_Resize

End Sub

Private Sub Slider1_Change()

    lblSliderValue.Caption = Slider1.Value

End Sub

Private Sub DriveBox1_Change()
    FolderList1.Path = DriveBox1.Drive & "\"
End Sub

Private Sub FolderList1_Change()
    FileList1.Path = FolderList1.Path
End Sub



