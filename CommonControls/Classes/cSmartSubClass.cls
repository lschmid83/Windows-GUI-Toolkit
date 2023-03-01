VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmartSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------
' Class SmartSubClass
'
' Version... 1.0
' Date...... 24 April 2001
'
' Copyright (C) 2001 Andrés Pons (andres@vbsmart.com)
' ----------------------------------------------------

Option Explicit

'Public event:
Public Event NewMessage( _
    ByVal hWnd As Long, _
    ByRef uMsg As Long, _
    ByRef wParam As Long, _
    ByRef lParam As Long, _
    ByRef Cancel As Boolean)

'Private variables:
Private m_hWnds() As Long

'API declarations:
Private Const GWL_WNDPROC = (-4)

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" ( _
    ByVal hWnd As Long, _
    ByVal lpString As String) As Long
    
Private Declare Function SetProp Lib "user32" Alias "SetPropA" ( _
    ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal hData As Long) As Long

Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" ( _
    ByVal hWnd As Long, _
    ByVal lpString As String) As Long

Private Declare Function IsWindow Lib "user32" ( _
    ByVal hWnd As Long) As Long

'
' Function SubClassHwnd
'
' This is the core function in this class.
' You can use it to both subclass and unsubclass a window.
' Once a window is subclassed the event NewMessage will
' be raised every time a message is sent to the window.
'
Public Function SubClassHwnd(ByVal hWnd As Long, ByVal bSubClass As Boolean) As Boolean

    Dim lRet As Long
    
    lRet = 0
    
    'Make sure that hWnd is a valid window handler...
    If IsWindow(hWnd) Then
    
        If bSubClass Then
        'We are subclassing a window...
            
            'Make sure that the window wasn't already subclassed...
            If GetProp(hWnd, SSC_OLDPROC) = 0 Then
            
                'Now we subclass the window by changing its windowproc
                lRet = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SmartSubClassWindowProc)
                
                'Check if we've managed to subclass...
                If lRet <> 0 Then
                    'Store the old windowproc and the memory address of this class...
                    SetProp hWnd, SSC_OLDPROC, lRet
                    SetProp hWnd, SSC_OBJADDR, ObjPtr(Me)
                    
                    'Add the window to an internal list of subclassed windows...
                    pAddHwndToList hWnd
                End If
            End If
        Else
        'We are unsubclassing a window...
        
            'Get the old windowproc...
            lRet = GetProp(hWnd, SSC_OLDPROC)
            
            If lRet <> 0 Then
                'Unsubclass the window...
                lRet = SetWindowLong(hWnd, GWL_WNDPROC, lRet)
            End If
            
            'Remove any extra information...
            RemoveProp hWnd, SSC_OLDPROC
            RemoveProp hWnd, SSC_OBJADDR
            
            'Remove the window from the internal list...
            pRemoveHwndFromList hWnd
        End If
    Else
        'If hWnd is not a valid window,
        'make sure that there isn't stored garbage...
        RemoveProp hWnd, SSC_OLDPROC
        RemoveProp hWnd, SSC_OBJADDR
        
        pRemoveHwndFromList hWnd
    End If
      
    SubClassHwnd = (lRet <> 0)
    
End Function

'
' Function WindowProc
'
' This is the link between the windowproc and the class instance.
' Every time SmartSubClassWindowProc receives a window message,
' it will post it to the right class instance.
'
Friend Function WindowProc( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

    Dim lRet As Long
    Dim bCancel As Boolean
    
    bCancel = False
    
    WindowProc = 0
    
    'Raise the event NewMessage...
    'This will tell the owner of the class variable that a
    'new message is ready to be processed.
    'The owner will be able to cancel the message by setting
    'the variable bCancel to True.
    RaiseEvent NewMessage(hWnd, uMsg, wParam, lParam, bCancel)
    
    'If the event hasn't been canceled by the owner
    'we need to send it to the original windowproc
    If Not bCancel Then
    
        lRet = GetProp(hWnd, SSC_OLDPROC)
        
        If lRet <> 0 Then
            'Send the message to the original windowproc...
            WindowProc = CallWindowProc(lRet, hWnd, uMsg, wParam, lParam)
        End If
        
    End If
    
End Function

'
' Every instance of the class mantains an internal
' list of subclassed windows.
'
Private Sub Class_Initialize()

    ReDim m_hWnds(0) As Long
    
End Sub

'
' When the class terminates it makes sure that
' there are no remainig subclassed windows.
'
Private Sub Class_Terminate()

    Dim i As Long
    
    For i = UBound(m_hWnds) To 1 Step -1
        If m_hWnds(i) > 0 Then
            SubClassHwnd m_hWnds(i), False
        End If
    Next i
    
End Sub

'
' Private Function pFindHwndInList()
'
' This functions searches for a specific window
' in its internal list. If it doesn't find the
' window it returns 0.
'
Private Function pFindHwndInList(ByVal hWnd As Long) As Long

    Dim i As Long
    Dim lPos As Long
    
    lPos = 0
    
    For i = 1 To UBound(m_hWnds)
        If m_hWnds(i) = hWnd And m_hWnds(i) > 0 Then
            lPos = i
            Exit For
        End If
    Next i
    
    pFindHwndInList = lPos
    
End Function

'
' Private Sub pAddHwndToList()
'
' This procedure adds a window handle to the internal list...
'
Private Sub pAddHwndToList(ByVal hWnd As Long)

    Dim lPos As Long
    
    If pFindHwndInList(hWnd) = 0 Then
        
        lPos = pFindNextPositionAvailableInList
        
        If lPos <> 0 Then
            m_hWnds(lPos) = hWnd
        Else
            lPos = UBound(m_hWnds) + 1
            ReDim Preserve m_hWnds(lPos) As Long
            
            m_hWnds(lPos) = hWnd
        End If
        
    End If

End Sub

'
' Private Sub pRemoveHwndFromList()
'
' This procedure removes a window handle from the internal list...
'
Private Sub pRemoveHwndFromList(ByVal hWnd As Long)

    Dim lPos As Long
    
    lPos = pFindHwndInList(hWnd)
    
    If lPos <> 0 Then
        If lPos = UBound(m_hWnds) Then
            ReDim Preserve m_hWnds(lPos - 1) As Long
        Else
            m_hWnds(lPos) = -1
        End If
    End If
    
End Sub

'
' Private Function pFindNextPositionAvailableInList()
'
' This functions searches for an "empty" entry in the
' internal list of window handles. When an entry is
' removed its is marked as empty by setting its value to -1.
'
' If there are no positions available, the function returns 0.
'
Private Function pFindNextPositionAvailableInList() As Long

    Dim i As Long
    Dim lPos As Long

    lPos = 0
    
    For i = 1 To UBound(m_hWnds)
        If m_hWnds(i) <= 0 Then
            lPos = i
            Exit For
        End If
    Next i
    
    pFindNextPositionAvailableInList = lPos
    
End Function
