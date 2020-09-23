Attribute VB_Name = "modMouse_Control"
Option Explicit
Private Const MAX_DEFAULTCHAR            As Integer = 2
Private Const MAX_LEADBYTES              As Integer = 12
Private Const WM_LBUTTONDOWN             As Long = &H201
Private Const WM_LBUTTONUP               As Long = &H202
Private Const SWP_FRAMECHANGED           As Long = &H20    '  The frame changed: send WM_NCCALCSIZE
Private Const MOUSEEVENTF_LEFTDOWN       As Long = &H2
Private Const MOUSEEVENTF_LEFTUP         As Long = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN     As Long = &H20
Private Const MOUSEEVENTF_MIDDLEUP       As Long = &H40
Private Const MOUSEEVENTF_RIGHTDOWN      As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP        As Long = &H10
Private Const MOUSEEVENTF_MOVE           As Long = &H1
Private Const SM_CXSCREEN                As Integer = 0
Private Const SM_CYSCREEN                As Integer = 1
Private Const TWIPS_PER_INCH             As Integer = 1440
Private Const POINTS_PER_INCH            As Integer = 72
Private Const MOUSEEVENTF_ABSOLUTE       As Long = &H8000
Private Const MOUSE_MICKEYS              As Long = 65535
Type RECT
    Left                                     As Long
    Top                                      As Long
    Right                                    As Long
    bottom                                   As Long
End Type
Type POINTAPI
    x                                        As Long
    y                                        As Long
End Type
Type CPINFO
    MaxCharSize                              As Long
    DefaultChar(MAX_DEFAULTCHAR)             As Byte
    LeadByte(MAX_LEADBYTES)                  As Byte
End Type
Public Enum enReportStyle
    rsPixels
    rsTwips
    rsInches
    rsPoints
End Enum
#If False Then
Private rsPixels, rsTwips, rsInches, rsPoints
#End If
Public Enum enButtonToClick
    btcLeft
    btcRight
    btcMiddle
End Enum
#If False Then
Private btcLeft, btcRight, btcMiddle
#End If
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, _
                                                    ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
                                              ByVal dx As Long, _
                                              ByVal dy As Long, _
                                              ByVal cbuttons As Long, _
                                              ByVal dwExtraInfo As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
                                                       ByVal yPoint As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Public Sub GetScreenRes(ByRef x As Long, _
                        ByRef y As Long, _
                        Optional ByVal ReportStyle As enReportStyle)
    On Error GoTo ErrorTrap
    x = GetSystemMetrics(SM_CXSCREEN)
    y = GetSystemMetrics(SM_CYSCREEN)
    If Not IsMissing(ReportStyle) Then
        If ReportStyle <> rsPixels Then
            x = x * Screen.TwipsPerPixelX
            y = y * Screen.TwipsPerPixelY
            If ReportStyle = rsInches Or ReportStyle = rsPoints Then
                x = x \ TWIPS_PER_INCH
                y = y \ TWIPS_PER_INCH
                If ReportStyle = rsPoints Then
                    x = x * POINTS_PER_INCH
                    y = y * POINTS_PER_INCH
                End If
            End If
        End If
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.GetScreenRes" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Function M__CenterMouseOn(ByVal lngHwnd As Long) As Boolean
Dim Xa    As Long
Dim Ya    As Long
Dim maxX  As Long
Dim maxY  As Long
Dim crect As RECT
Dim rc    As Long
Dim x     As Long
    On Error GoTo ErrorTrap
    GetScreenRes maxX, maxY
    rc = GetWindowRect(lngHwnd, crect)
    If rc Then
        Xa = crect.Left + ((crect.Right - crect.Left) / 2)
        Ya = crect.Top + ((crect.bottom - crect.Top) / 2)
        If (x >= 0 And x <= maxX) And (Ya >= 0 And Ya <= maxY) Then
            M__MouseMove Xa, Ya
            M__CenterMouseOn = True
        Else
            M__CenterMouseOn = False
        End If
    End If
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M__CenterMouseOn" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Function M__Get_Window() As Long
Dim cursorPos As Long
Dim Pos       As POINTAPI
    On Error GoTo ErrorTrap
    GetCursorPos Pos
    cursorPos = WindowFromPoint(Pos.x, Pos.y)
    M__Get_Window = cursorPos
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M__Get_Window" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Sub M__MouseMove(ByRef xPixel As Long, _
                        ByRef yPixel As Long)
Dim cbuttons    As Long
Dim dwExtraInfo As Long
    On Error GoTo ErrorTrap
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, PixelXToMickey(xPixel), PixelYToMickey(yPixel), cbuttons, dwExtraInfo
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M__MouseMove" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M__SetMousePos(Xx As Long, _
                          Yy As Long)
    On Error GoTo ErrorTrap
    SetCursorPos Xx, Yy
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M__SetMousePos" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Function M_GetCusorPos() As Long
Dim N   As Long
Dim Pos As POINTAPI
    On Error GoTo ErrorTrap
    N = GetCursorPos(Pos)
    M_GetCusorPos = N
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_GetCusorPos" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Function M_GetX() As Long
Dim N As POINTAPI
    On Error GoTo ErrorTrap
    GetCursorPos N
    M_GetX = N.x
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_GetX" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Function M_GetY() As Long
Dim N As POINTAPI
    On Error GoTo ErrorTrap
    GetCursorPos N
    M_GetY = N.y
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_GetY" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Sub M_LeftClick()
    On Error GoTo ErrorTrap
    M_LeftDown
    M_LeftUp
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_LeftClick" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M_LeftDown()
    On Error GoTo ErrorTrap
    mouse_event WM_LBUTTONDOWN, 0, 0, 0, 0
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_LeftDown" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M_LeftUp()
    On Error GoTo ErrorTrap
    mouse_event WM_LBUTTONUP, 0, 0, 0, 0
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_LeftUp" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M_MiddleClick()
    On Error GoTo ErrorTrap
    M_MiddleDown
    M_MiddleUp
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_MiddleClick" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M_MiddleDown()
    On Error GoTo ErrorTrap
    mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_MiddleDown" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M_MiddleUp()
    On Error GoTo ErrorTrap
    mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_MiddleUp" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M_RightClick()
    On Error GoTo ErrorTrap
    M_RightDown
    M_RightUp
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_RightClick" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M_RightDown()
    On Error GoTo ErrorTrap
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_RightDown" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub M_RightUp()
    On Error GoTo ErrorTrap
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.M_RightUp" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
' Simulates a mouse click
Public Function MouseFullClick(ByVal MBClick As enButtonToClick) As Boolean
Dim cbuttons    As Long
Dim dwExtraInfo As Long
Dim mevent      As Long
    On Error GoTo ErrorTrap
    Select Case MBClick
    Case btcLeft
        mevent = MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP
    Case btcRight
        mevent = MOUSEEVENTF_RIGHTDOWN Or MOUSEEVENTF_RIGHTUP
    Case btcMiddle
        mevent = MOUSEEVENTF_MIDDLEDOWN Or MOUSEEVENTF_MIDDLEUP
    Case Else
        MouseFullClick = False
        Exit Function
    End Select
    mouse_event mevent, 0&, 0&, cbuttons, dwExtraInfo
    MouseFullClick = True
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.MouseFullClick" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Sub MoveMouse(xMove As Long, _
                     yMove As Long)
    On Error GoTo ErrorTrap
    mouse_event MOUSEEVENTF_MOVE, xMove, yMove, 0, 0
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.MoveMouse" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
' Converts pixel X coordinates to mickeys
Public Function PixelXToMickey(ByVal pixX As Long) As Long
Dim x        As Long
Dim y        As Long
Dim tX       As Single
Dim tpixX    As Single
Dim tMickeys As Single
    On Error GoTo ErrorTrap
    GetScreenRes x, y
    tMickeys = MOUSE_MICKEYS
    tX = x
    tpixX = pixX
    PixelXToMickey = CLng((tMickeys / tX) * tpixX)
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.PixelXToMickey" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
' Converts pixel Y coordinates to mickeys
Public Function PixelYToMickey(ByVal pixY As Long) As Long
Dim x        As Long
Dim y        As Long
Dim tY       As Single
Dim tpixY    As Single
Dim tMickeys As Single
    On Error GoTo ErrorTrap
    GetScreenRes x, y
    tMickeys = MOUSE_MICKEYS
    tY = y
    tpixY = pixY
    PixelYToMickey = CLng((tMickeys / tY) * tpixY)
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.Mouse_Control.PixelYToMickey" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
':)Code Fixer V3.0.9 (9/15/2005 1:32:10 PM) 57 + 349 = 406 Lines Thanks Ulli for inspiration and lots of code.


