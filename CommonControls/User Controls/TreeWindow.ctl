VERSION 5.00
Begin VB.UserControl TreeWindow 
   ClientHeight    =   3555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ScaleHeight     =   3555
   ScaleWidth      =   3795
   Begin VB.PictureBox picBack 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   3060
      Left            =   0
      ScaleHeight     =   3060
      ScaleWidth      =   3210
      TabIndex        =   0
      Top             =   0
      Width           =   3210
   End
End
Attribute VB_Name = "TreeWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_c As cCaptureBF
Private m_sCurrentFolder As String
Private m_bCancel As Boolean

Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Implements ICaptureBF

Public Property Get Cancelled() As Boolean
   Cancelled = m_bCancel
End Property
Public Property Get SelectedFolder() As String
   SelectedFolder = m_sCurrentFolder
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdExtract_Click()
   ' Chosen to extract!
   m_bCancel = False
   Unload Me
End Sub

Private Sub cmdNewFolder_Click()
Dim sI As String
   ' Get a new folder to extract to:
   sI = InputBox("Please enter the folder name.", , m_sCurrentFolder)
   If sI <> "" Then
      On Error Resume Next
      MkDir sI
      If err.Number <> 0 Then
         MsgBox "An error occurred: " & err.Description, vbExclamation
      Else
         ' Reload the browse dialog but point to
         ' the newly created path.  This is much
         ' smoother than the WinZip equivalent!!!
         m_c.Reload sI
      End If
   End If
End Sub

Private Sub cmdPick_Click()
   m_c.Browse.SetFolder cboExtractTo.Text
End Sub

Private Sub Form_Initialize()
   'DebugMsg "frmCapture:Initialize"
   m_bCancel = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' Ensure we have unloaded the dialog:
   m_c.Unload
   ' Important: to ensure this class terminates we
   ' must set to nothing here:
   Set m_c = Nothing
End Sub

Private Sub Form_Terminate()
   'DebugMsg "frmCapture:Terminate"
End Sub

Private Property Let ICaptureBF_CaptureBrowseForFolder(RHS As Object)
   ' Provides you with a reference to the cCaptureBrowseForFolder
   ' object, which you can use to refer to the cBrowseForFolder
   ' dialog:
   Set m_c = RHS
End Property

Private Property Get ICaptureBF_CapturehWnd() As Long
   ' Requests the window you want to capture the folder browse
   ' dialog into.  You must ensure you have shown the form at this stage.
   Me.Show , frmMain
   picBack.BorderStyle = 0
   ICaptureBF_CapturehWnd = picBack.hwnd
End Property

Private Sub ICaptureBF_SelectionChanged(ByVal sPath As String)
   ' Fired when the selection in the folder browse dialog
   ' changes:
   cboExtractTo.Text = sPath
   cboExtractTo.SelStart = Len(sPath)
   If Len(sPath) > 0 Then
      cboExtractTo.SelLength = Len(sPath)
   End If
   m_sCurrentFolder = sPath
End Sub

Private Sub ICaptureBF_Unload()
   ' Fired when the browse for folder dialog
   ' is closed.  Ensures that you clear up at
   ' the right time.
   Unload Me
End Sub


Private Sub UserControl_Initialize()

End Sub
