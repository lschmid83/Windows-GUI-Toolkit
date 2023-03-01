VERSION 5.00
Begin VB.UserControl ImageButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   ScaleHeight     =   675
   ScaleWidth      =   2460
   ToolboxBitmap   =   "ImageButton.ctx":0000
   Begin XpCommonControls.MaskBox imgPicture 
      Height          =   405
      Left            =   135
      Top             =   120
      Width           =   420
      _ExtentX        =   0
      _ExtentY        =   0
   End
End
Attribute VB_Name = "ImageButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Citex Software                                                       '
'XP GUI Toolkit v1.0 - Command Button Component v1.0                  '
'Copyright 2001 Lawrence Schmid                                       '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


'General Variables:
Dim UsrRect As RECT
Dim TextAlign As Long
Dim strEnabled As String 'Holds the buttons enabled value
Dim strValue As String 'Holds the current state of the control
Dim CaptionX, CaptionY As Integer 'Holds Caption x,y position before button is pressed
Dim PictureX, PictureY As Integer 'Holds Picture x,y position before button is pressed
Dim TabFocus As Boolean 'Holds buttons tab status

'Property Variables:
Dim m_Alignment As AlignmentConstants
Dim m_Default As Boolean
Dim m_FocusRectangle As Boolean
Dim m_Value As Boolean
Dim m_CheckButton As Boolean
Dim m_MouseOverPicture As Picture
Dim m_ToolTipCaption As String
Dim m_ToolTipIcon As ToolTipIconEnum
Dim m_ToolTipStyle As ToolTipStyleEnum
Dim m_ToolTipTitle As String
Dim m_Caption As String
Dim m_MaskColor As OLE_COLOR
Dim m_DisabledPicture As Picture
Dim m_DownPicture As Picture
Dim m_Picture As Picture
Dim m_PictureAlign As PictureAlignEnum
Dim m_UseMaskColor As Boolean
Dim m_ForeColor As OLE_COLOR

'Event Declarations:
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event MouseLeave()
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
'Default Property Values:
Const m_def_AutoDisabledColor = 0






'---------------------------------------------------------------------------------------
'Public control subs

Public Sub ChangeTheme()

pSetEnabled


If UserControl.Enabled = True Then

    If m_CheckButton = False Then

        If m_Default = True Then
        TabFocus = True
        pSetGraphic "HasFocus"
       
        
        Else
        pSetGraphic "UnPressed"
           
        End If

    Else
    
     If m_Value = True Then
     pSetGraphic "Pressed"
     End If
       
    
    End If

Else

pSetGraphic "Disabled"

End If

End Sub

Public Sub RefreshTheme()

pSetEnabled

If UserControl.Enabled = True Then

    If m_CheckButton = False Then

        If m_Default = True Then
        TabFocus = True
        pSetGraphic "HasFocus"
       
        
        Else
        pSetGraphic "UnPressed"
           
        End If

    Else
    
     If m_Value = True Then
     pSetGraphic "Pressed"
     
     Else
     pSetGraphic "UnPressed"
     
     End If
       
    
    End If

Else
pSetGraphic "Disabled"

End If

End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

Property Let FontBold(Bold As Boolean)
'lblCaption.FontBold = Bold
End Property

Property Get FontBold() As Boolean
Attribute FontBold.VB_MemberFlags = "400"
'FontBold = lblCaption.FontBold

End Property

Property Let FontItalic(Italic As Boolean)
'lblCaption.FontItalic = Italic
End Property

Property Get FontItalic() As Boolean
Attribute FontItalic.VB_MemberFlags = "400"
'FontItalic = lblCaption.FontItalic

End Property

Property Let FontName(Name As String)
'lblCaption.FontName = Name
End Property

Property Get FontName() As String
Attribute FontName.VB_MemberFlags = "400"
'FontName = lblCaption.FontName

End Property

Property Let FontSize(Size As Long)
'lblCaption.FontSize = Size
End Property

Property Get FontSize() As Long
Attribute FontSize.VB_MemberFlags = "400"
'FontSize = lblCaption.FontSize

End Property

Property Let FontStrikeThru(StrikeThru As Boolean)
'lblCaption.FontStrikeThru = StrikeThru
End Property

Property Get FontStrikeThru() As Boolean
Attribute FontStrikeThru.VB_MemberFlags = "400"
'FontStrikeThru = lblCaption.FontStrikeThru

End Property

Property Let FontUnderline(UnderLine As Boolean)
'lblCaption.FontUnderline = UnderLine
End Property

Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_MemberFlags = "400"
'FontUnderline = lblCaption.FontUnderline

End Property

'------------------------------------------------------------------------------
'Private internal subs

Private Sub pSetEnabled()

If m_CheckButton = False Then

    If UserControl.Enabled = True Then
        
                
        Set imgPicture.Picture = m_Picture
        pPositionPic
        
        UserControl.ForeColor = m_ForeColor
            
     
    Else 'button disabled
        
        
    Set UserControl.MaskPicture = DisabledPicture
    
    If UserControl.MaskPicture <> Empty Then
    Set imgPicture.Picture = DisabledPicture
    pPositionPic
    End If
    
        If glbAppearance <> Win98 Then
        UserControl.ForeColor = &H92A1A1
        Else
        UserControl.ForeColor = &H808080
        End If
    
  
    End If

Else

    If UserControl.Enabled = True Then
        
        UserControl.ForeColor = m_ForeColor
        
   
    Else 'button disabled
    
        If glbAppearance <> Win98 Then
        UserControl.ForeColor = &H92A1A1
        Else
        UserControl.ForeColor = &H808080
        End If
    
    End If



End If


End Sub


Private Sub pPositionPic()

If imgPicture.Picture <> Empty Then 'If button has a picture

    If m_PictureAlign = ButtonCenter Then
    imgPicture.Top = ((UserControl.Height / 2) - (imgPicture.Height / 2)) - 110
    imgPicture.Left = (UserControl.Width / 2) - (imgPicture.Width / 2)
    
   ' lblCaption.Top = UserControl.Height - (lblCaption.Height + 100)
   ' lblCaption.Left = 7 * Screen.TwipsPerPixelX
   ' lblCaption.Width = Width - (14 * Screen.TwipsPerPixelX)
        
    Else
       
    
    imgPicture.Top = (UserControl.Height / 2) - (imgPicture.Height / 2)
    imgPicture.Left = 80
    
   ' lblCaption.Top = (UserControl.Height / 2) - (lblCaption.Height / 2)
   ' lblCaption.Left = imgPicture.Left + imgPicture.Width + 90
   ' lblCaption.Width = UserControl.Width - 180
    End If

Else 'Button has no picture center caption

'lblCaption.Top = (Height / 2) - (lblCaption.Height / 2)
'lblCaption.Left = 7 * Screen.TwipsPerPixelX
'lblCaption.Width = Width - (14 * Screen.TwipsPerPixelX)

End If

PictureX = imgPicture.Left
PictureY = imgPicture.Top
'CaptionX = lblCaption.Left
'CaptionY = lblCaption.Top


End Sub


Private Sub pSetGraphic(New_Value As String)

'Draw caption
Select Case m_Alignment
Case vbCenter
TextAlign = DT_CENTER
Case vbLeftJustify
TextAlign = DT_LEFT
Case vbRightJustify
TextAlign = DT_RIGHT
End Select


'Sets current value
strValue = New_Value

UserControl.Cls

'Loads the correct background picture
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\Back", crBitmap), 3 * Screen.TwipsPerPixelX, 3 * Screen.TwipsPerPixelY, Width - (6 * Screen.TwipsPerPixelX), Height - (6 * Screen.TwipsPerPixelY)

UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\TopLeft", crBitmap), 0, 0
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\TopRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 0
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\BottomLeft", crBitmap), 0, Height - (3 * Screen.TwipsPerPixelY)
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\BottomRight", crBitmap), Width - (3 * Screen.TwipsPerPixelX), Height - (3 * Screen.TwipsPerPixelY)

UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\Left", crBitmap), 0, 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelY, Height - (6 * Screen.TwipsPerPixelY)
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\Right", crBitmap), Width - (3 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY, 3 * Screen.TwipsPerPixelY, Height - (6 * Screen.TwipsPerPixelY)
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\Top", crBitmap), 3 * Screen.TwipsPerPixelY, 0, Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY
UserControl.PaintPicture PictureFromResource(ResourceLib.hModule, "CommandButton\" & strValue & "\Bottom", crBitmap), 3 * Screen.TwipsPerPixelY, Height - (3 * Screen.TwipsPerPixelY), Width - (6 * Screen.TwipsPerPixelX), 3 * Screen.TwipsPerPixelY

    If strValue = "Pressed" Then
    
        Call SetRect(UsrRect, 15, (((Height / 2) / Screen.TwipsPerPixelY) - UserControl.Font.Size) + 3, (Width / Screen.TwipsPerPixelX) - 15, Height)
        'Call SetRect(UsrRect, (imgPicture.Width / Screen.TwipsPerPixelX) + 15, (((Height / 2) / Screen.TwipsPerPixelY) - UserControl.Font.Size) + 3, (Width / Screen.TwipsPerPixelX) - ((imgPicture.Width / Screen.TwipsPerPixelX) + 19), Height)
        Call DrawText(UserControl.hdc, m_Caption, -1, UsrRect, TextAlign)
        Call SetRectEmpty(UsrRect)
           
    Else
        Call SetRect(UsrRect, 15, (((Height / 2) / Screen.TwipsPerPixelY) - UserControl.Font.Size) + 2, 15, Height)
        'Call SetRect(UsrRect, (imgPicture.Width / Screen.TwipsPerPixelX) + 14, (((Height / 2) / Screen.TwipsPerPixelY) - UserControl.Font.Size) + 2, (Width / Screen.TwipsPerPixelX) - ((imgPicture.Width / Screen.TwipsPerPixelX) + 20), Height)
        Call DrawText(UserControl.hdc, m_Caption, -1, UsrRect, TextAlign)
        Call SetRectEmpty(UsrRect)
    
   
    End If


If strValue = "Pressed" Then
'Make picture/caption move down-right
imgPicture.Left = imgPicture.Left + 15
imgPicture.Top = imgPicture.Top + 15
'lblCaption.Left = lblCaption.Left + 15
'lblCaption.Top = lblCaption.Top + 15

Else
'Restore picture/caption to their original positions
imgPicture.Left = PictureX
imgPicture.Top = PictureY
'lblCaption.Left = CaptionX
'lblCaption.Top = CaptionY

End If




End Sub


'--------------------------------------------------------------------------------------
'Control properties

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Alignment() As AlignmentConstants
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
m_Alignment = New_Alignment

pSetGraphic strValue

PropertyChanged "Alignment"
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

RefreshTheme

PropertyChanged "Enabled"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
Set UserControl.Font = New_Font
UserControl_Resize

PropertyChanged "Font"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the command button."
Attribute Caption.VB_UserMemId = -518
Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
m_Caption = New_Caption
UserControl.AccessKeys = GetAccessKeyFromString(m_Caption)
    
PropertyChanged "Caption"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get CheckButton() As Boolean
Attribute CheckButton.VB_Description = "Returns/sets whether the button only has two states i.e Checked/UnChecked."
CheckButton = m_CheckButton
End Property

Public Property Let CheckButton(ByVal New_CheckButton As Boolean)
m_CheckButton = New_CheckButton

pSetEnabled

PropertyChanged "CheckButton"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Default() As Boolean
Attribute Default.VB_Description = "Returns/sets whether this button is the default button for the form."
    Default = m_Default
End Property

Public Property Let Default(ByVal New_Default As Boolean)
m_Default = New_Default

RefreshTheme

PropertyChanged "Default"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
MaskColor = m_MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
m_MaskColor = New_MaskColor
imgPicture.SetMaskColor New_MaskColor
Set Picture = Picture


PropertyChanged "MaskColor"
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
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
m_ForeColor = New_ForeColor

'If UserControl.Enabled = True Then
'lblCaption.ForeColor = m_ForeColor

'End If

PropertyChanged "ForeColor"
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
'MemberInfo=14,0,0,0
Public Property Get DownPicture() As Picture
Attribute DownPicture.VB_Description = "Returns/sets a graphic to be displayed when the button is pressed."
Set DownPicture = m_DownPicture
End Property

Public Property Set DownPicture(ByVal New_DownPicture As Picture)
Set m_DownPicture = New_DownPicture

PropertyChanged "DownPicture"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
    hdc = UserControl.hdc
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
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/sets whether this object can act as an OLE drop target, and whether this takes place automatically or under programmatic control."
OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
UserControl.OLEDropMode() = New_OLEDropMode
'lblCaption.OLEDropMode() = New_OLEDropMode
PropertyChanged "OLEDropMode"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get FocusRectangle() As Boolean
Attribute FocusRectangle.VB_Description = "Returns/sets whether a dotted border is drawn in the button when it has focus."
    FocusRectangle = m_FocusRectangle
End Property

Public Property Let FocusRectangle(ByVal New_FocusRectangle As Boolean)
m_FocusRectangle = New_FocusRectangle

RefreshTheme


PropertyChanged "FocusRectangle"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get ToolTipCaption() As String
Attribute ToolTipCaption.VB_Description = "Returns/sets the text displayed in the tooltip."
ToolTipCaption = m_ToolTipCaption
End Property

Public Property Let ToolTipCaption(ByVal New_ToolTipCaption As String)
m_ToolTipCaption = New_ToolTipCaption
pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon


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
pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

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
pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon

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
pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon


PropertyChanged "ToolTipTitle"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets whether the color assignged in the MaskColor property is used as the transparent color in the buttons picture."
UseMaskColor = m_UseMaskColor
End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)
m_UseMaskColor = New_UseMaskColor

PropertyChanged "UseMaskColor"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Value() As Boolean
Attribute Value.VB_Description = "Returns/sets the state of the button if CheckButton = True."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)

If Enabled = True Then

    If m_CheckButton = True Then
    m_Value = New_Value
    
        If m_Value = True Then
            If strValue <> "Pressed" Then
            pSetGraphic "Pressed"
            End If
        
        Else
        
            If strValue <> "UnPressed" Then
            pSetGraphic "UnPressed"
            End If
        
        End If
    
    End If

End If

PropertyChanged "Value"
End Property




Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)


TabFocus = True

'Draw focus rectangle
If TabFocus = True And m_FocusRectangle = True Then
Call SetRect(UsrRect, 4, 4, (Width / Screen.TwipsPerPixelX) - 4, (Height / Screen.TwipsPerPixelY) - 4)
Call DrawFocusRect(UserControl.hdc, UsrRect)
Call SetRectEmpty(UsrRect)
End If

End Sub



'-------------------------------------------------------------------------------------------------------------------------


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    SetDefaultTheme

    m_Alignment = 0
    Set m_DisabledPicture = LoadPicture("")
    Set m_DownPicture = LoadPicture("")
    m_CheckButton = False
    Set m_MouseOverPicture = LoadPicture("")
    m_Default = False
    m_PictureAlign = ButtonCenter
    m_UseMaskColor = True
    m_MaskColor = &HFF00FF
    imgPicture.SetMaskColor m_MaskColor
    m_ToolTipCaption = ""
    m_ToolTipIcon = NoIcon
    m_ToolTipStyle = Standard
    m_ToolTipTitle = ""
    m_Value = False
    m_FocusRectangle = True

    Caption = Ambient.DisplayName

    Width = 110 * Screen.TwipsPerPixelX
    Height = 39 * Screen.TwipsPerPixelY

    ChangeTheme

End Sub





Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


            If m_CheckButton = False Then
    
                Set imgPicture.Picture = m_Picture
                pPositionPic
                 
                    If TabFocus = False Then
                    pSetGraphic "UnPressed"
                    Else
                    pSetGraphic "HasFocus"
                    End If
                
                        'Draw focus rectangle
                        If TabFocus = True And m_FocusRectangle = True Then
                        
                        Call SetRect(UsrRect, 4, 4, (Width / Screen.TwipsPerPixelX) - 4, (Height / Screen.TwipsPerPixelY) - 4)
                        Call DrawFocusRect(UserControl.hdc, UsrRect)
                        Call SetRectEmpty(UsrRect)
                        End If
       
                RaiseEvent MouseLeave
    
    
            Else
            
            
                 If m_Value = False Then
                        Set imgPicture.Picture = m_Picture
                        pPositionPic
                        pSetGraphic "UnPressed"
                    End If
                            
            RaiseEvent MouseLeave
            
            End If





        If m_CheckButton = False Then
        
         Set UserControl.MaskPicture = MouseOverPicture
            
            If UserControl.MaskPicture <> Empty Then
            Set imgPicture.Picture = MouseOverPicture
            pPositionPic
            End If
            
            
                If glbAppearance <> Win98 Then
                    
                    If strValue <> "MouseOver" And strValue <> "Pressed" Then
                    pSetGraphic "MouseOver"
                    End If
                             
                Else
                
                    If TabFocus = True Then
                    pSetGraphic "HasFocus"
                    End If
                
                             
                End If
                
         
             
                    'Draw focus rectangle
                    If TabFocus = True And m_FocusRectangle = True Then
                    
                    Call SetRect(UsrRect, 4, 4, (Width / Screen.TwipsPerPixelX) - 4, (Height / Screen.TwipsPerPixelY) - 4)
                    Call DrawFocusRect(UserControl.hdc, UsrRect)
                    Call SetRectEmpty(UsrRect)
                    End If
                          
            RaiseEvent MouseMove(Button, Shift, X, Y)
            
            
            
        Else
        
        
            If m_Value = False Then
            
            Set UserControl.MaskPicture = MouseOverPicture
            
            If UserControl.MaskPicture <> Empty Then
            Set imgPicture.Picture = MouseOverPicture
            pPositionPic
            End If
            
                If glbAppearance <> Win98 Then
                pSetGraphic "MouseOver"
                End If
            End If
            
            
            RaiseEvent MouseMove(Button, Shift, X, Y)
            
            
            End If


    
    RaiseEvent MouseMove(Button, Shift, X, Y)


End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SetDefaultTheme

    m_Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Caption = PropBag.ReadProperty("Caption", "Command1")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set m_DisabledPicture = PropBag.ReadProperty("DisabledPicture", Nothing)
    Set m_DownPicture = PropBag.ReadProperty("DownPicture", Nothing)
    m_PictureAlign = PropBag.ReadProperty("PictureAlign", ButtonCenter)
    m_UseMaskColor = PropBag.ReadProperty("UseMaskColor", True)
    m_MaskColor = PropBag.ReadProperty("MaskColor", &HFF00FF)
    m_CheckButton = PropBag.ReadProperty("CheckButton", False)
    Set m_MouseOverPicture = PropBag.ReadProperty("MouseOverPicture", Nothing)
    m_Default = PropBag.ReadProperty("Default", False)
    m_ToolTipCaption = PropBag.ReadProperty("ToolTipCaption", "")
    m_ToolTipIcon = PropBag.ReadProperty("ToolTipIcon", NoIcon)
    m_ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", Standard)
    m_ToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
    m_Value = PropBag.ReadProperty("Value", False)
    m_FocusRectangle = PropBag.ReadProperty("FocusRectangle", False)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    
    If m_UseMaskColor = True Then
    imgPicture.SetMaskColor m_MaskColor
    End If
    Set Picture = PropBag.ReadProperty("Picture", Nothing)

    pSetToolTip UserControl.hwnd, m_ToolTipCaption, m_ToolTipTitle, m_ToolTipStyle, m_ToolTipIcon
    UserControl.AccessKeys = GetAccessKeyFromString(m_Caption)
    
  
    
    If glbControlsRefreshed = True Then
    RefreshTheme 'Refresh control because tilebar refresh code has already run before this
                 'This is needed to make sure the correct appearance is displayed
    End If
    
    
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Caption", m_Caption, "Command1")
    Call PropBag.WriteProperty("MaskColor", m_MaskColor, &HFF00FF)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, &H80000012)
    Call PropBag.WriteProperty("DisabledPicture", m_DisabledPicture, Nothing)
    Call PropBag.WriteProperty("DownPicture", m_DownPicture, Nothing)
    Call PropBag.WriteProperty("PictureAlign", m_PictureAlign, ButtonCenter)
    Call PropBag.WriteProperty("UseMaskColor", m_UseMaskColor, True)
    Call PropBag.WriteProperty("ToolTipIcon", m_ToolTipIcon, NoIcon)
    Call PropBag.WriteProperty("ToolTipStyle", m_ToolTipStyle, Standard)
    Call PropBag.WriteProperty("ToolTipTitle", m_ToolTipTitle, "")
    Call PropBag.WriteProperty("CheckButton", m_CheckButton, False)
    Call PropBag.WriteProperty("MouseOverPicture", m_MouseOverPicture, Nothing)
    Call PropBag.WriteProperty("Value", m_Value, False)
    Call PropBag.WriteProperty("ToolTipCaption", m_ToolTipCaption, "")
    Call PropBag.WriteProperty("FocusRectangle", m_FocusRectangle, False)
    Call PropBag.WriteProperty("Alignment", m_Alignment, 2)
    Call PropBag.WriteProperty("Default", m_Default, False)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
 
End Sub


'-------------------------------------------------------------------------------------------------------------------------
'Control Events




Private Sub UserControl_Click()
RaiseEvent Click
  
End Sub

Private Sub UserControl_EnterFocus()

If m_CheckButton = False Then
    TabFocus = True
    
    If strValue <> "Pressed" Then
    'What happens when button gets focus
    pSetGraphic "HasFocus"
    
        'Draw focus rectangle
        If TabFocus = True And m_FocusRectangle = True Then
        
        Call SetRect(UsrRect, 4, 4, (Width / Screen.TwipsPerPixelX) - 4, (Height / Screen.TwipsPerPixelY) - 4)
        Call DrawFocusRect(UserControl.hdc, UsrRect)
        Call SetRectEmpty(UsrRect)
        End If
    
    End If
        

End If

End Sub

Private Sub UserControl_ExitFocus()

If m_CheckButton = False Then

    If Enabled = True Then
    TabFocus = False
    pSetGraphic "UnPressed"
    End If

End If

End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If m_CheckButton = False Then

    If Button = 1 Then
       
       
        
        Set UserControl.MaskPicture = DownPicture
    
        If UserControl.MaskPicture <> Empty Then
        Set imgPicture.Picture = DownPicture
        pPositionPic
        End If
        
        pSetGraphic "Pressed"
  
        'Draw focus rectangle
        If m_FocusRectangle = True Then
        Call SetRect(UsrRect, 4, 4, (Width / Screen.TwipsPerPixelX) - 4, (Height / Screen.TwipsPerPixelY) - 4)
        Call DrawFocusRect(UserControl.hdc, UsrRect)
        Call SetRectEmpty(UsrRect)
        UserControl.Refresh
        End If
   
    
    
    End If


RaiseEvent MouseDown(Button, Shift, X, Y)

Else 'button is a check button

If m_Value = False Then
    
    Set UserControl.MaskPicture = DownPicture
    
    If UserControl.MaskPicture <> Empty Then
    Set imgPicture.Picture = DownPicture
    pPositionPic
    End If
    
pSetGraphic "Pressed"
m_Value = True

Else
    Set imgPicture.Picture = m_Picture
    pPositionPic

pSetGraphic "UnPressed"
m_Value = False

End If


End If



End Sub










Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


If m_CheckButton = False Then

If Button = 1 Then

        
        If TabFocus = True Then
        pSetGraphic "HasFocus"
        End If
        
        Set imgPicture.Picture = m_Picture
        pPositionPic
        
        RaiseEvent MouseUp(Button, Shift, X, Y)

End If


Else 'button is a check button



End If
  

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


Private Sub UserControl_Resize()

On Error GoTo err

pPositionPic

If ResourceLib.hModule <> 0 Then
    If strValue <> "HasFocus" Then
    pSetGraphic "UnPressed"
    Else
    pSetGraphic "HasFocus"
        If TabFocus = True Then
            If m_FocusRectangle = True Then
            UserControl.Refresh
          '  DottedBrush.Rectangle UserControl.hdc, 4, 4, (Width / Screen.TwipsPerPixelX) - 9, (Height / Screen.TwipsPerPixelY) - 9, 1
            End If
            End If
    
    
    End If

End If

err:
Exit Sub


End Sub



Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call UserControl_MouseDown(Button, Shift, X, Y)

End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub


Private Sub lblCaption_Click()
Call UserControl_Click

End Sub




Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub lblCaption_OLECompleteDrag(Effect As Long)
Call UserControl_OLECompleteDrag(Effect)

End Sub

Private Sub lblCaption_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub


Private Sub lblCaption_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Call UserControl_OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub lblCaption_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Call UserControl_OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub lblCaption_OLESetData(Data As DataObject, DataFormat As Integer)
Call UserControl_OLESetData(Data, DataFormat)

End Sub

Private Sub lblCaption_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Call UserControl_OLEStartDrag(Data, AllowedEffects)
End Sub

