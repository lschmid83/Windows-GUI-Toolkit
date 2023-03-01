VERSION 5.00
Object = "{EF59A10B-9BC4-11D3-8E24-44910FC10000}#10.0#0"; "VBALEDIT.OCX"
Begin VB.UserControl RichEditBox 
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   ScaleHeight     =   2670
   ScaleWidth      =   3945
   ToolboxBitmap   =   "RichEditBox.ctx":0000
   Begin vbalEdit.vbalRichEdit txtMain 
      Height          =   660
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   1164
      Version         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      ViewMode        =   0
      Border          =   0   'False
      AutoURLDetect   =   0   'False
      ScrollBars      =   0
   End
End
Attribute VB_Name = "RichEditBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_Border = 0
Const m_def_ForeColor = 0
'Property Variables:
Dim m_BackColor As OLE_COLOR
Dim m_Border As Boolean
Dim m_ForeColor As OLE_COLOR


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,AutoURLDetect
Public Property Get AutoURLDetect() As Boolean
Attribute AutoURLDetect.VB_Description = "Gets/sets whether the control will automatically detect hyperlinks prefixed by certain URL identifiers (e.g. http:)"
    AutoURLDetect = txtMain.AutoURLDetect
End Property

Public Property Let AutoURLDetect(ByVal New_AutoURLDetect As Boolean)
    txtMain.AutoURLDetect() = New_AutoURLDetect
    PropertyChanged "AutoURLDetect"
End Property

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
'MemberInfo=0,0,0,0
Public Property Get Border() As Boolean
Attribute Border.VB_Description = "Gets/sets whether the control has a 3D border."
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,CharFormatRange
Public Property Get CharFormatRange() As ERECSetFormatRange
Attribute CharFormatRange.VB_Description = "Gets/sets the range to which font formatting will apply."
    CharFormatRange = txtMain.CharFormatRange
End Property

Public Property Let CharFormatRange(ByVal New_CharFormatRange As ERECSetFormatRange)
    txtMain.CharFormatRange() = New_CharFormatRange
    PropertyChanged "CharFormatRange"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,ControlLeftMargin
Public Property Get ControlLeftMargin() As Long
Attribute ControlLeftMargin.VB_Description = "Gets/sets the margin from the left hand edge of the control to the RichEdit control."
    ControlLeftMargin = txtMain.ControlLeftMargin
End Property

Public Property Let ControlLeftMargin(ByVal New_ControlLeftMargin As Long)
    txtMain.ControlLeftMargin() = New_ControlLeftMargin
    PropertyChanged "ControlLeftMargin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,ControlRightMargin
Public Property Get ControlRightMargin() As Long
Attribute ControlRightMargin.VB_Description = "Gets/sets the margin from the right hand edge of the control to the RichEdit control."
    ControlRightMargin = txtMain.ControlRightMargin
End Property

Public Property Let ControlRightMargin(ByVal New_ControlRightMargin As Long)
    txtMain.ControlRightMargin() = New_ControlRightMargin
    PropertyChanged "ControlRightMargin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,DisableNoScroll
Public Property Get DisableNoScroll() As Boolean
    DisableNoScroll = txtMain.DisableNoScroll
End Property

Public Property Let DisableNoScroll(ByVal New_DisableNoScroll As Boolean)
    txtMain.DisableNoScroll() = New_DisableNoScroll
    PropertyChanged "DisableNoScroll"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Gets/sets the font of the control or selection, depending on the setting of CharFormatRange."
Attribute Font.VB_UserMemId = -512
    Set Font = txtMain.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtMain.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontBackColour
Public Property Get FontBackColour() As Long
Attribute FontBackColour.VB_Description = "Gets/sets the background colour of the control or selection, depending on the setting of CharFormatRange."
    FontBackColour = txtMain.FontBackColour
End Property

Public Property Let FontBackColour(ByVal New_FontBackColour As Long)
    txtMain.FontBackColour() = New_FontBackColour
    PropertyChanged "FontBackColour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Gets/sets whether the font is bold for the control or selection, depending on the setting of CharFormatRange."
    FontBold = txtMain.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    txtMain.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontColour
Public Property Get FontColour() As Long
Attribute FontColour.VB_Description = "Gets/sets the colour of the font for the control or selection, depending on the setting of CharFormatRange."
    FontColour = txtMain.FontColour
End Property

Public Property Let FontColour(ByVal New_FontColour As Long)
    txtMain.FontColour() = New_FontColour
    PropertyChanged "FontColour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Gets/sets whether the font is italic for the control or selection, depending on the setting of CharFormatRange."
    FontItalic = txtMain.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    txtMain.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontLink
Public Property Get FontLink() As Boolean
Attribute FontLink.VB_Description = "Gets/sets whether the selection acts as a hyperlink.  Set CharFormatRange to selection."
    FontLink = txtMain.FontLink
End Property

Public Property Let FontLink(ByVal New_FontLink As Boolean)
    txtMain.FontLink() = New_FontLink
    PropertyChanged "FontLink"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontProtected
Public Property Get FontProtected() As Boolean
Attribute FontProtected.VB_Description = "Gets/sets whether the selection is protected (raises the ModifyRequest event).  Set CharFormatRange to selection."
    FontProtected = txtMain.FontProtected
End Property

Public Property Let FontProtected(ByVal New_FontProtected As Boolean)
    txtMain.FontProtected() = New_FontProtected
    PropertyChanged "FontProtected"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontStrikeOut
Public Property Get FontStrikeOut() As Boolean
Attribute FontStrikeOut.VB_Description = "Gets/sets whether the font is struck out for the control or selection, depending on the setting of CharFormatRange."
    FontStrikeOut = txtMain.FontStrikeOut
End Property

Public Property Let FontStrikeOut(ByVal New_FontStrikeOut As Boolean)
    txtMain.FontStrikeOut() = New_FontStrikeOut
    PropertyChanged "FontStrikeOut"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontSubScript
Public Property Get FontSubScript() As Boolean
Attribute FontSubScript.VB_Description = "Gets/sets whether the font is subscripted for the control or selection, depending on the setting of CharFormatRange."
    FontSubScript = txtMain.FontSubScript
End Property

Public Property Let FontSubScript(ByVal New_FontSubScript As Boolean)
    txtMain.FontSubScript() = New_FontSubScript
    PropertyChanged "FontSubScript"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontSuperScript
Public Property Get FontSuperScript() As Boolean
Attribute FontSuperScript.VB_Description = "Gets/sets whether the font is superscripted for the control or selection, depending on the setting of CharFormatRange."
    FontSuperScript = txtMain.FontSuperScript
End Property

Public Property Let FontSuperScript(ByVal New_FontSuperScript As Boolean)
    txtMain.FontSuperScript() = New_FontSuperScript
    PropertyChanged "FontSuperScript"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Gets/sets whether the font is underlined for the control or selection, depending on the setting of CharFormatRange."
    FontUnderline = txtMain.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    txtMain.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,HideSelection
Public Property Get HideSelection() As Boolean
    HideSelection = txtMain.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    txtMain.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Gets/sets the maximum length of text or RTF loaded into the control."
    MaxLength = txtMain.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtMain.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Modified
Public Property Get Modified() As Boolean
Attribute Modified.VB_Description = "Gets/sets whether the contents of the control have been modified."
    Modified = txtMain.Modified
End Property

Public Property Let Modified(ByVal New_Modified As Boolean)
    txtMain.Modified() = New_Modified
    PropertyChanged "Modified"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,ParagraphAlignment
Public Property Get ParagraphAlignment() As ERECParagraphAlignmentConstants
Attribute ParagraphAlignment.VB_Description = "Gets/Sets the alignment of the selected paragraph."
    ParagraphAlignment = txtMain.ParagraphAlignment
End Property

Public Property Let ParagraphAlignment(ByVal New_ParagraphAlignment As ERECParagraphAlignmentConstants)
    txtMain.ParagraphAlignment() = New_ParagraphAlignment
    PropertyChanged "ParagraphAlignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,ParagraphNumbering
Public Property Get ParagraphNumbering() As ERECParagraphNumberingConstants
Attribute ParagraphNumbering.VB_Description = "Gets/sets whether the selected paragraph has bullets or not."
    ParagraphNumbering = txtMain.ParagraphNumbering
End Property

Public Property Let ParagraphNumbering(ByVal New_ParagraphNumbering As ERECParagraphNumberingConstants)
    txtMain.ParagraphNumbering() = New_ParagraphNumbering
    PropertyChanged "ParagraphNumbering"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,PasswordChar
Public Property Get PasswordChar() As String
    PasswordChar = txtMain.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtMain.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Picture
Public Property Get Picture() As IPicture
Attribute Picture.VB_Description = "Gets/sets the background picture tiled behind the control when Transparent is set to True."
    Set Picture = txtMain.Picture
End Property

Public Property Set Picture(ByVal New_Picture As IPicture)
    Set txtMain.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,ReadOnly
Public Property Get ReadOnly() As Boolean
Attribute ReadOnly.VB_Description = "Gets/sets whether the control is read-only."
    ReadOnly = txtMain.ReadOnly
End Property

Public Property Let ReadOnly(ByVal New_ReadOnly As Boolean)
    txtMain.ReadOnly() = New_ReadOnly
    PropertyChanged "ReadOnly"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Redraw
Public Property Get Redraw() As Boolean
Attribute Redraw.VB_Description = "Gets/sets whether the control will redraw or not."
    Redraw = txtMain.Redraw
End Property

Public Property Let Redraw(ByVal New_Redraw As Boolean)
    txtMain.Redraw() = New_Redraw
    PropertyChanged "Redraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,ScrollBars
Public Property Get ScrollBars() As ERECScrollBarConstants
    ScrollBars = txtMain.ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ERECScrollBarConstants)
    txtMain.ScrollBars() = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,SingleLine
Public Property Get SingleLine() As Boolean
    SingleLine = txtMain.SingleLine
End Property

Public Property Let SingleLine(ByVal New_SingleLine As Boolean)
    txtMain.SingleLine() = New_SingleLine
    PropertyChanged "SingleLine"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Gets the text contained in the control."
    Text = txtMain.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtMain.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,TextLimit
Public Property Get TextLimit() As Long
Attribute TextLimit.VB_Description = "Same as MaxLength (!)"
    TextLimit = txtMain.TextLimit
End Property

Public Property Let TextLimit(ByVal New_TextLimit As Long)
    txtMain.TextLimit() = New_TextLimit
    PropertyChanged "TextLimit"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,TextOnly
Public Property Get TextOnly() As Boolean
Attribute TextOnly.VB_Description = "Gets/sets whether the control acts as a text-only control or not."
    TextOnly = txtMain.TextOnly
End Property

Public Property Let TextOnly(ByVal New_TextOnly As Boolean)
    txtMain.TextOnly() = New_TextOnly
    PropertyChanged "TextOnly"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,Transparent
Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Gets/sets whether the control is transparent and displays the Picture or not."
    Transparent = txtMain.Transparent
End Property

Public Property Let Transparent(ByVal New_Transparent As Boolean)
    txtMain.Transparent() = New_Transparent
    PropertyChanged "Transparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,UseVersion
Public Property Get UseVersion() As ERECControlVersion
Attribute UseVersion.VB_Description = "Gets/sets which version of the RichEdit DLL to use: version 2/3 (RichEd20.DLL) or version 1 (RichEd32.DLL)"
    UseVersion = txtMain.UseVersion
End Property

Public Property Let UseVersion(ByVal New_UseVersion As ERECControlVersion)
    txtMain.UseVersion() = New_UseVersion
    PropertyChanged "UseVersion"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMain,txtMain,-1,ViewMode
Public Property Get ViewMode() As ERECViewModes
Attribute ViewMode.VB_Description = "Gets/sets who the control lays out the text on screen."
    ViewMode = txtMain.ViewMode
End Property

Public Property Let ViewMode(ByVal New_ViewMode As ERECViewModes)
    txtMain.ViewMode() = New_ViewMode
    PropertyChanged "ViewMode"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_Border = m_def_Border
    m_ForeColor = m_def_ForeColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtMain.AutoURLDetect = PropBag.ReadProperty("AutoURLDetect", False)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Border = PropBag.ReadProperty("Border", m_def_Border)
    txtMain.CharFormatRange = PropBag.ReadProperty("CharFormatRange", 4)
    txtMain.ControlLeftMargin = PropBag.ReadProperty("ControlLeftMargin", 0)
    txtMain.ControlRightMargin = PropBag.ReadProperty("ControlRightMargin", 0)
    txtMain.DisableNoScroll = PropBag.ReadProperty("DisableNoScroll", False)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtMain.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtMain.FontBackColour = PropBag.ReadProperty("FontBackColour", 0)
    txtMain.FontBold = PropBag.ReadProperty("FontBold", False)
    txtMain.FontColour = PropBag.ReadProperty("FontColour", 0)
    txtMain.FontItalic = PropBag.ReadProperty("FontItalic", False)
    txtMain.FontLink = PropBag.ReadProperty("FontLink", False)
    txtMain.FontProtected = PropBag.ReadProperty("FontProtected", False)
    txtMain.FontStrikeOut = PropBag.ReadProperty("FontStrikeOut", False)
    txtMain.FontSubScript = PropBag.ReadProperty("FontSubScript", False)
    txtMain.FontSuperScript = PropBag.ReadProperty("FontSuperScript", False)
    txtMain.FontUnderline = PropBag.ReadProperty("FontUnderline", False)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    txtMain.HideSelection = PropBag.ReadProperty("HideSelection", False)
    txtMain.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtMain.Modified = PropBag.ReadProperty("Modified", False)
    txtMain.ParagraphAlignment = PropBag.ReadProperty("ParagraphAlignment", 0)
    txtMain.ParagraphNumbering = PropBag.ReadProperty("ParagraphNumbering", 0)
    txtMain.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    txtMain.ReadOnly = PropBag.ReadProperty("ReadOnly", False)
    txtMain.Redraw = PropBag.ReadProperty("Redraw", True)
    txtMain.ScrollBars = PropBag.ReadProperty("ScrollBars", 0)
    txtMain.SingleLine = PropBag.ReadProperty("SingleLine", False)
    txtMain.Text = PropBag.ReadProperty("Text", "")
    txtMain.TextLimit = PropBag.ReadProperty("TextLimit", 32767)
    txtMain.TextOnly = PropBag.ReadProperty("TextOnly", False)
    txtMain.Transparent = PropBag.ReadProperty("Transparent", False)
    txtMain.UseVersion = PropBag.ReadProperty("UseVersion", 1)
    txtMain.ViewMode = PropBag.ReadProperty("ViewMode", 0)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoURLDetect", txtMain.AutoURLDetect, False)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
    Call PropBag.WriteProperty("CharFormatRange", txtMain.CharFormatRange, 4)
    Call PropBag.WriteProperty("ControlLeftMargin", txtMain.ControlLeftMargin, 0)
    Call PropBag.WriteProperty("ControlRightMargin", txtMain.ControlRightMargin, 0)
    Call PropBag.WriteProperty("DisableNoScroll", txtMain.DisableNoScroll, False)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", txtMain.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBackColour", txtMain.FontBackColour, 0)
    Call PropBag.WriteProperty("FontBold", txtMain.FontBold, False)
    Call PropBag.WriteProperty("FontColour", txtMain.FontColour, 0)
    Call PropBag.WriteProperty("FontItalic", txtMain.FontItalic, False)
    Call PropBag.WriteProperty("FontLink", txtMain.FontLink, False)
    Call PropBag.WriteProperty("FontProtected", txtMain.FontProtected, False)
    Call PropBag.WriteProperty("FontStrikeOut", txtMain.FontStrikeOut, False)
    Call PropBag.WriteProperty("FontSubScript", txtMain.FontSubScript, False)
    Call PropBag.WriteProperty("FontSuperScript", txtMain.FontSuperScript, False)
    Call PropBag.WriteProperty("FontUnderline", txtMain.FontUnderline, False)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("HideSelection", txtMain.HideSelection, False)
    Call PropBag.WriteProperty("MaxLength", txtMain.MaxLength, 0)
    Call PropBag.WriteProperty("Modified", txtMain.Modified, False)
    Call PropBag.WriteProperty("ParagraphAlignment", txtMain.ParagraphAlignment, 0)
    Call PropBag.WriteProperty("ParagraphNumbering", txtMain.ParagraphNumbering, 0)
    Call PropBag.WriteProperty("PasswordChar", txtMain.PasswordChar, "")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ReadOnly", txtMain.ReadOnly, False)
    Call PropBag.WriteProperty("Redraw", txtMain.Redraw, True)
    Call PropBag.WriteProperty("ScrollBars", txtMain.ScrollBars, 0)
    Call PropBag.WriteProperty("SingleLine", txtMain.SingleLine, False)
    Call PropBag.WriteProperty("Text", txtMain.Text, "")
    Call PropBag.WriteProperty("TextLimit", txtMain.TextLimit, 32767)
    Call PropBag.WriteProperty("TextOnly", txtMain.TextOnly, False)
    Call PropBag.WriteProperty("Transparent", txtMain.Transparent, False)
    Call PropBag.WriteProperty("UseVersion", txtMain.UseVersion, 1)
    Call PropBag.WriteProperty("ViewMode", txtMain.ViewMode, 0)
End Sub

