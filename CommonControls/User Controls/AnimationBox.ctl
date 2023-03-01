VERSION 5.00
Object = "{BBB88E11-FB86-11D3-B06C-00500427A693}#1.0#0"; "VBALAVI6.OCX"
Begin VB.UserControl AnimationBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "AnimationBox.ctx":0000
   Begin vbalAVI6.vbalAVIPlayer AVIMain 
      Height          =   3420
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   6033
      BorderStyle     =   0
      TransparentColor=   0
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "AnimationBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ToolTipIcon = 0
Const m_def_ToolTipStyle = 0
Const m_def_ToolTipTitle = "0"
'Property Variables:
Dim m_BackColor As OLE_COLOR
Dim m_ToolTipIcon As Variant
Dim m_ToolTipStyle As Variant
Dim m_ToolTipTitle As String


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,AVIPlay
Public Function AVIPlay() As Boolean
    AVIPlay = AVIMain.AVIPlay()
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,AVISeek
Public Function AVISeek(ByVal nFrame As Long) As Boolean
    AVISeek = AVIMain.AVISeek(nFrame)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,AVIStop
Public Function AVIStop() As Boolean
    AVIStop = AVIMain.AVIStop()
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,Centre
Public Property Get Centre() As Boolean
    Centre = AVIMain.Centre
End Property

Public Property Let Centre(ByVal New_Centre As Boolean)
    AVIMain.Centre() = New_Centre
    PropertyChanged "Centre"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,FileName
Public Property Get FileName() As String
    FileName = AVIMain.FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    AVIMain.FileName() = New_FileName
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = AVIMain.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set AVIMain.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,Playing
Public Property Get Playing() As Boolean
    Playing = AVIMain.Playing
End Property

Public Property Let Playing(ByVal New_Playing As Boolean)
    AVIMain.Playing() = New_Playing
    PropertyChanged "Playing"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,Transparent
Public Property Get Transparent() As Boolean
    Transparent = AVIMain.Transparent
End Property

Public Property Let Transparent(ByVal New_Transparent As Boolean)
    AVIMain.Transparent() = New_Transparent
    PropertyChanged "Transparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=AVIMain,AVIMain,-1,TransparentColor
Public Property Get TransparentColor() As Long
    TransparentColor = AVIMain.TransparentColor
End Property

Public Property Let TransparentColor(ByVal New_TransparentColor As Long)
    AVIMain.TransparentColor() = New_TransparentColor
    PropertyChanged "TransparentColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipIcon() As Variant
    ToolTipIcon = m_ToolTipIcon
End Property

Public Property Let ToolTipIcon(ByVal New_ToolTipIcon As Variant)
    m_ToolTipIcon = New_ToolTipIcon
    PropertyChanged "ToolTipIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ToolTipStyle() As Variant
    ToolTipStyle = m_ToolTipStyle
End Property

Public Property Let ToolTipStyle(ByVal New_ToolTipStyle As Variant)
    m_ToolTipStyle = New_ToolTipStyle
    PropertyChanged "ToolTipStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipTitle() As String
    ToolTipTitle = m_ToolTipTitle
End Property

Public Property Let ToolTipTitle(ByVal New_ToolTipTitle As String)
    m_ToolTipTitle = New_ToolTipTitle
    PropertyChanged "ToolTipTitle"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ToolTipIcon = m_def_ToolTipIcon
    m_ToolTipStyle = m_def_ToolTipStyle
    m_ToolTipTitle = m_def_ToolTipTitle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    AVIMain.Centre = PropBag.ReadProperty("Centre", False)
    AVIMain.FileName = PropBag.ReadProperty("FileName", "")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    AVIMain.Playing = PropBag.ReadProperty("Playing", False)
    AVIMain.Transparent = PropBag.ReadProperty("Transparent", True)
    AVIMain.TransparentColor = PropBag.ReadProperty("TransparentColor", 0)
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", m_def_ToolTipIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", m_def_ToolTipStyle)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", m_def_ToolTipTitle)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Centre", AVIMain.Centre, False)
    Call PropBag.WriteProperty("FileName", AVIMain.FileName, "")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Playing", AVIMain.Playing, False)
    Call PropBag.WriteProperty("Transparent", AVIMain.Transparent, True)
    Call PropBag.WriteProperty("TransparentColor", AVIMain.TransparentColor, 0)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, m_def_ToolTipIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, m_def_ToolTipStyle)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, m_def_ToolTipTitle)
End Sub

