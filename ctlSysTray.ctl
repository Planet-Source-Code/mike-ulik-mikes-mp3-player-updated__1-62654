VERSION 5.00
Begin VB.UserControl ctlSysTrayIcon 
   BackColor       =   &H00C0C0C0&
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   390
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   435
   ScaleWidth      =   390
End
Attribute VB_Name = "ctlSysTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ToolTipText             As String * 64
Private m_Icon                    As Picture
Private m_Visible                 As Boolean
Private Type NOTIFYICONDATA
    cbSize                            As Long
    hwnd                              As Long
    uId                               As Long
    uFlags                            As Long
    ucallbackMessage                  As Long
    hIcon                             As Long
    szTip                             As String * 64
End Type
Private mTrayIcon                 As NOTIFYICONDATA
Private Const NIM_ADD             As Long = &H0
Private Const NIM_MODIFY          As Long = &H1
Private Const NIM_DELETE          As Long = &H2
Private Const NIF_MESSAGE         As Long = &H1
Private Const NIF_ICON            As Long = &H2
Private Const NIF_TIP             As Long = &H4
Private Const WM_MOUSEMOVE        As Long = &H200
Private Const WM_LBUTTONDOWN      As Long = &H201
Private Const WM_LBUTTONUP        As Long = &H202
Private Const WM_LBUTTONDBLCLK    As Long = &H203
Private Const WM_RBUTTONDOWN      As Long = &H204
Private Const WM_RBUTTONUP        As Long = &H205
Private Const WM_RBUTTONDBLCLK    As Long = &H206
'Various events that are generated
Public Event MouseMoved()
Public Event LeftButtonDown()
Public Event LeftButtonUp()
Public Event RightButtonDown()
Public Event RightButtonUp()
Public Event RightButtonDoubleClick()
Public Event LeftButtonDoubleClick()
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                   pnid As NOTIFYICONDATA) As Boolean
Private Sub AddSystemTray() ' Add the Icon in the Tray
    On Error GoTo ErrorTrap
    With mTrayIcon
        .cbSize = Len(mTrayIcon)
        .hwnd = UserControl.hwnd
        .uId = vbNull '1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_MOUSEMOVE
    End With 'mTrayIcon
    If PictureType(m_Icon) = vbPicTypeIcon Then
        mTrayIcon.hIcon = m_Icon.Handle
    End If
    mTrayIcon.szTip = Trim$(m_ToolTipText) & vbNullChar
    Shell_NotifyIcon NIM_ADD, mTrayIcon
Exit Sub
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.AddSystemTray" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Property Get Icon() As Picture
    On Error GoTo ErrorTrap
    Set Icon = m_Icon
Exit Property
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.Icon" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Property
Public Property Set Icon(vNewValue As Picture)
    On Error GoTo ErrorTrap
    If PictureType(vNewValue) = vbPicTypeIcon Then
        Set m_Icon = vNewValue
        PropertyChanged "Icon"
        ModSystemTray
        Set UserControl.Picture = vNewValue
    ElseIf (PictureType(vNewValue) = vbPicTypeNone) Then
        Set m_Icon = LoadPicture("")
        Set UserControl.Picture = m_Icon
        PropertyChanged "Icon"
        RemoveSystemTray
    Else
        MsgBox "Only Icons and Cursors allowed.", vbInformation, UserControl.Ambient.DisplayName
    End If
Exit Property
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.Icon" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Property
Public Property Get IconToolTipText() As String
    On Error GoTo ErrorTrap
    IconToolTipText = Trim$(m_ToolTipText)
Exit Property
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.IconToolTipText" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Property
Public Property Let IconToolTipText(ByVal vNewValue As String)
    On Error GoTo ErrorTrap
    m_ToolTipText = Trim$(vNewValue)
    PropertyChanged "IconToolTipText"
    ModSystemTray
Exit Property
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.IconToolTipText" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Property
Public Property Get IconVisible() As Boolean
    On Error GoTo ErrorTrap
    IconVisible = m_Visible
Exit Property
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.IconVisible" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Property
Public Property Let IconVisible(ByVal vNewValue As Boolean)
    On Error GoTo ErrorTrap
    m_Visible = vNewValue
    PropertyChanged "IconVisible"
    If IsRunTime() = True Then   'Make the Icon visible only in runtime
        If m_Visible Then
            AddSystemTray
        Else
            RemoveSystemTray
        End If
    End If
Exit Property
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.IconVisible" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Property
Private Function IsRunTime() As Boolean
    On Error GoTo ErrorTrap
    IsRunTime = UserControl.Ambient.UserMode
Exit Function
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.IsRunTime" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Private Sub ModSystemTray() ' Modify the Icon in the Tray
    On Error GoTo ErrorTrap
    With mTrayIcon
        .cbSize = Len(mTrayIcon)
        .hwnd = UserControl.hwnd
        .uId = vbNull '1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_MOUSEMOVE
    End With 'mTrayIcon
    If PictureType(m_Icon) = vbPicTypeIcon Then
        mTrayIcon.hIcon = m_Icon.Handle
    End If
    mTrayIcon.szTip = Trim$(m_ToolTipText) & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, mTrayIcon
Exit Sub
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.ModSystemTray" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Function PictureType(p As StdPicture) As PictureTypeConstants
Dim ans As PictureTypeConstants
    On Error GoTo ErrorTrap
    If TypeName(p) = "Nothing" Then
        ans = vbPicTypeNone
    Else
        ans = p.type
    End If
    PictureType = ans
Exit Function
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.PictureType" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Private Sub RemoveSystemTray() ' Remove the Icon in the Tray
    On Error GoTo ErrorTrap
    With mTrayIcon
        .cbSize = Len(mTrayIcon)
        .hwnd = UserControl.hwnd
        .uId = vbNull '1&
    End With 'mTrayIcon
    Shell_NotifyIcon NIM_DELETE, mTrayIcon
Exit Sub
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.RemoveSystemTray" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
    On Error GoTo ErrorTrap
'Raise the Respective events after analysis of X.
    Select Case x \ Screen.TwipsPerPixelX
    Case WM_MOUSEMOVE
        RaiseEvent MouseMoved
    Case WM_LBUTTONDOWN
        RaiseEvent LeftButtonDown
    Case WM_LBUTTONUP
        RaiseEvent LeftButtonUp
    Case WM_LBUTTONDBLCLK
        RaiseEvent LeftButtonDoubleClick
    Case WM_RBUTTONDOWN
        RaiseEvent RightButtonDown
    Case WM_RBUTTONUP
        RaiseEvent RightButtonUp
    Case WM_RBUTTONDBLCLK
        RaiseEvent RightButtonDoubleClick
    End Select
Exit Sub
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.UserControl_MouseMove" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo ErrorTrap
    With PropBag
        Set m_Icon = .ReadProperty("Icon", Nothing)
        m_ToolTipText = Trim$(.ReadProperty("IconToolTipText", " "))
        m_Visible = .ReadProperty("IconVisible", False)
    End With 'PropBag
    If IsRunTime() = True Then   'Make the Icon visible only in runtime
        If m_Visible Then
            AddSystemTray
        Else
            RemoveSystemTray
        End If
    End If
    Set UserControl.Picture = m_Icon
Exit Sub
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.UserControl_ReadProperties" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub UserControl_Terminate()
    On Error GoTo ErrorTrap
    RemoveSystemTray
Exit Sub
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.UserControl_Terminate" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo ErrorTrap
    With PropBag
        .WriteProperty "Icon", m_Icon, Nothing
        .WriteProperty "IconToolTipText", Trim$(m_ToolTipText), " "
        .WriteProperty "IconVisible", m_Visible, False
    End With 'PropBag
Exit Sub
ErrorTrap:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
       Err.Description & vbNewLine & _
       vbNewLine & _
       "Debug Information:" & vbNewLine & _
       "MidiDateBase.SysTrayIcon.UserControl_WriteProperties" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:30:33 PM) 36 + 273 = 309 Lines Thanks Ulli for inspiration and lots of code.


