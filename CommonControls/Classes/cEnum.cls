VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Citex Software                                                      '
' Windows GUI Toolkit - cEnum.cls                                     '
' Copyright 1999-2001 Lawrence Schmid                                 '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum AppearanceEnum
    Blue = 1
    Green = 2
    Silver = 3
    Win98 = 4
End Enum

Public Enum FocusEnum
    HasFocus = 1
    LostFocus = 2
End Enum

Public Enum ButtonsEnum
    All = 1
    MinClose = 2
    CloseOnly = 3
    WhatsThis = 4
    None = 5
End Enum

Public Enum BorderStyleEnum
    Fixed = 1
    Sizable = 2
End Enum
                 
Public Enum WindowStyleEnum
    StandardWindow = 1
    ToolWindow = 2
End Enum

Public Enum SizeEnum2
    Small = 1
    Large = 2
End Enum

Public Enum OrientationEnum
    Horizontal = 0
    Vertical = 1
End Enum

Public Enum ExtaControlsEnum
    SeparatorHorizontal = 1
    SeparatorVertical = 2
    ToolBarBackground = 3
    PowerButton = 4
    SleepButton = 5
    RestartButton = 6
    GoButton = 7
    TruncateButton = 8
    UpArrow = 9
    DownArrow = 10
    LeftArrow = 11
    RightArrow = 12
    ComboArrow = 13
    TopBorder = 14
    BottomBorder = 15
    LeftBorder = 16
    RightBorder = 17
End Enum

Public Enum PictureAlignEnum
    ButtonCenter = 1
    ButtonLeft = 2
End Enum

Public Enum ButtonTypeEnum
    CommandButton = 1
    ToolbarButton = 2
End Enum

Public Enum BorderEnum
    bdrLeft = 1
    bdrRight = 2
    bdrBottom = 3
End Enum

Public Enum ToolTipIconEnum
    Info = 1
    Warning = 2
    Error = 3
    NoIcon = 4
End Enum

Public Enum ToolTipStyleEnum
    Standard = 1
    Balloon = 2
End Enum

Public Enum ButtonStyleEnum
    StandardButton = 1
    DropDown = 2
    Dual = 3
End Enum
