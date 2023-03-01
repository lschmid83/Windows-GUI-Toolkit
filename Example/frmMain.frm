VERSION 5.00
Object = "{05931B16-A732-4810-88BE-423C9A7A76B5}#1.0#0"; "CommonControls.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmMain"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin CommonControls.ComboBox ComboBox1 
      Height          =   315
      Left            =   2415
      TabIndex        =   18
      Top             =   2370
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5010
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
      TabIndex        =   16
      Top             =   2370
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   582
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4770
      Top             =   5565
   End
   Begin CommonControls.ExtraControls ExtraControls5 
      Height          =   435
      Left            =   1140
      Top             =   5595
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   767
      ControlType     =   8
   End
   Begin CommonControls.CommandButton CommandButton1 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1335
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
      Caption         =   "CommandButton1"
   End
   Begin CommonControls.ToolbarButton ToolbarButton1 
      Height          =   615
      Left            =   240
      Top             =   2085
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
   Begin CommonControls.PictureBox PictureBox1 
      Height          =   1080
      Left            =   5400
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4980
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   1905
      BackColor       =   16777215
   End
   Begin CommonControls.ListBox ListBox1 
      Height          =   1065
      Left            =   2280
      TabIndex        =   13
      Top             =   4980
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
      TabIndex        =   12
      Top             =   3090
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
      TabIndex        =   11
      Top             =   3090
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
      TabIndex        =   10
      Top             =   3075
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
      TabIndex        =   9
      Top             =   2025
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
      TabIndex        =   8
      Top             =   1650
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
      TabIndex        =   7
      Top             =   1305
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   2699
      BackColor       =   14215660
      Min             =   0
      Value           =   0
   End
   Begin CommonControls.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   7500
      TabIndex        =   6
      Top             =   6360
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   397
      Value           =   0
   End
   Begin CommonControls.Hyperlink Hyperlink1 
      Height          =   225
      Left            =   7230
      TabIndex        =   5
      Top             =   5850
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   397
      BackColor       =   14215660
      Caption         =   "Citex Software"
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
      URL             =   "https://www.citexsoftware.co.uk"
   End
   Begin CommonControls.ExtraControls ExtraControls4 
      Height          =   390
      Left            =   270
      Top             =   5595
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
      ControlType     =   7
   End
   Begin CommonControls.ExtraControls ExtraControls3 
      Height          =   495
      Left            =   1440
      Top             =   4980
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ControlType     =   6
      ToolTipCaption  =   "Restart"
   End
   Begin CommonControls.ExtraControls ExtraControls2 
      Height          =   495
      Left            =   855
      Top             =   4980
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ControlType     =   5
      ToolTipCaption  =   "Sleep"
   End
   Begin CommonControls.ExtraControls ExtraControls1 
      Height          =   495
      Left            =   270
      Top             =   4980
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      ToolTipCaption  =   "Power"
   End
   Begin CommonControls.OptionButton OptionButton2 
      Height          =   240
      Left            =   2400
      TabIndex        =   4
      Top             =   2025
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
      Top             =   1650
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
      Top             =   1290
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
      Top             =   3600
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
   Begin CommonControls.ExtraControls ToolBarSeparator 
      Height          =   330
      Index           =   1
      Left            =   1530
      Top             =   810
      Width           =   15
      _ExtentX        =   26
      _ExtentY        =   582
      ControlType     =   2
   End
   Begin CommonControls.ExtraControls ToolBarSeparator 
      Height          =   330
      Index           =   0
      Left            =   1125
      Top             =   810
      Width           =   15
      _ExtentX        =   26
      _ExtentY        =   582
      ControlType     =   2
   End
   Begin CommonControls.ScrollBar ScrollBar1 
      Height          =   5100
      Left            =   8400
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1170
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   8996
   End
   Begin CommonControls.ToolbarButton ToolbarButton 
      Height          =   330
      Index           =   0
      Left            =   105
      Top             =   795
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      Caption         =   ""
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
      ButtonType      =   1
   End
   Begin CommonControls.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      Top             =   6270
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   714
      Caption         =   "StatusBar1"
      ResizeHandle    =   0   'False
   End
   Begin CommonControls.Border Border2 
      Align           =   4  'Align Right
      Height          =   5460
      Left            =   8655
      Top             =   810
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   9631
      BorderType      =   2
   End
   Begin CommonControls.ExtraControls ToolBar1 
      Height          =   375
      Left            =   75
      Top             =   810
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   661
      ControlType     =   3
   End
   Begin CommonControls.MenuBar MenuBar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      Top             =   450
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   635
      MenuImageStrip  =   "C:\Program Files\Windows GUI Toolkit\ImageStrip.bmp"
      MenuPath        =   "C:\Program Files\Windows GUI Toolkit\ExampleMenu.dat"
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
      Height          =   5460
      Left            =   0
      Top             =   810
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   9631
   End
   Begin CommonControls.Frame Frame1 
      Height          =   1545
      Index           =   1
      Left            =   5880
      Top             =   1290
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
      TabIndex        =   15
      Top             =   2595
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ToolBarButtons As Integer

Private Sub Form_Load()

    ' Initialise file paths
    MenuBar1.MenuImageStrip = TitleBar1.ProgramFiles() & "\Windows GUI Toolkit\ImageStrip.bmp"
    MenuBar1.MenuPath = TitleBar1.ProgramFiles() & "\Windows GUI Toolkit\ExampleMenu.dat"
    Set PictureBox1.Picture = LoadPicture(TitleBar1.ProgramFiles() & "\Windows GUI Toolkit\Images\windows-xp-logo.bmp")
    Set PictureBox2.Picture = LoadPicture(TitleBar1.ProgramFiles() & "\Windows GUI Toolkit\Images\citex-logo.bmp")

    FolderList1.Path = "c:\"
    FileList1.Path = "c:\"

    ' Create toolbar buttons
    Dim i As Integer
    Dim left As Integer
    left = 440
    ToolBarButtons = 6
    ToolbarButton(0).Caption = ""
    Set ToolbarButton(0).Picture = LoadPicture("c:\Program Files\Windows GUI Toolkit\Images\new.bmp")
    For i = 1 To ToolBarButtons
        Load ToolbarButton(i)
        With ToolbarButton(i)
            .Top = 810
            .left = left
            .Visible = True
            .ZOrder (0)
            .Caption = ""
        End With
        
        If i = 2 Or i = 3 Then
            left = left + 405
        Else
            left = left + 330
        End If
                
        If i = 1 Then
            Set ToolbarButton(i).Picture = LoadPicture("c:\Program Files\Windows GUI Toolkit\Images\open.bmp")
        ElseIf i = 2 Then
            Set ToolbarButton(i).Picture = LoadPicture("c:\Program Files\Windows GUI Toolkit\Images\save.bmp")
        ElseIf i = 3 Then
            Set ToolbarButton(i).Picture = LoadPicture("c:\Program Files\Windows GUI Toolkit\Images\undo.bmp")
        ElseIf i = 4 Then
            Set ToolbarButton(i).Picture = LoadPicture("c:\Program Files\Windows GUI Toolkit\Images\cut.bmp")
        ElseIf i = 5 Then
            Set ToolbarButton(i).Picture = LoadPicture("c:\Program Files\Windows GUI Toolkit\Images\copy.bmp")
        ElseIf i = 6 Then
            Set ToolbarButton(i).Picture = LoadPicture("c:\Program Files\Windows GUI Toolkit\Images\paste.bmp")
        End If
     
    Next i

    For i = 10 To 1 Step -1
        ComboBox1.AddItem ("Item" & i)
    Next i
    ComboBox1.ListIndex = 0

    For i = 10 To 1 Step -1
        ListBox1.AddItem ("Item" & i)
    Next i

    frmMain.Height = 6690

    ' Restore theme
    SetAppearance (GetSetting(App.EXEName, "appearance", "theme", 1))

End Sub

Private Sub Form_Activate()
    CommandButton1.SetFocus
End Sub

Private Sub Form_Resize()
    
    ' Resize toolbar
    If TitleBar1.Appearance <> Win98 Then
        ToolBar1.Width = frmMain.Width - 10
        ScrollBar1.Top = 1170
        ScrollBar1.left = frmMain.Width - 330
        ScrollBar1.Height = 5100
    Else
        ToolBar1.Width = frmMain.Width - 120
        ScrollBar1.Top = 1100
        ScrollBar1.left = frmMain.Width - 320
        ScrollBar1.Height = 5240
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
        ToolBar1.Top = 810
        ToolBar1.left = 75
        ToolBar1.Height = 375

        For i = 0 To ToolBarButtons
            With ToolbarButton(i)
                .Top = 810
            End With
        Next i
        ToolBarSeparator(0).Top = 810
        ToolBarSeparator(1).Top = 810
        ProgressBar1.Top = 6343
        ProgressBar1.left = 1200
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
            ToolBar1.Top = 705
            ToolBar1.left = 55
            ToolBar1.Height = 390
            For i = 0 To ToolBarButtons
                With ToolbarButton(i)
                    .Top = 730
                End With
            Next i
            
            ToolBarSeparator(0).Top = 740
            ToolBarSeparator(1).Top = 740
            ProgressBar1.Top = 6380
            ProgressBar1.left = 1100
            lblSliderValue.BackColor = &HC8D0D4
          
        End If
        
        Form_Resize

End Sub

Private Sub Slider1_Change()

    lblSliderValue.Caption = Slider1.Value

End Sub

Private Sub CommandButton1_Click()

    If Timer1.Enabled = False Then
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    If ProgressBar1.Value < 100 Then
        ProgressBar1.Value = ProgressBar1.Value + 5
    Else
        ProgressBar1.Value = 0
    End If
End Sub

Private Sub DriveBox1_Change()
    FolderList1.Path = DriveBox1.Drive & "\"
End Sub

Private Sub FolderList1_Change()
    FileList1.Path = FolderList1.Path
End Sub



