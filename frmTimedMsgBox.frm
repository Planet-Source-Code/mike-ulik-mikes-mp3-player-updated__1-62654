VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E1FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1950
   ClientLeft      =   2730
   ClientTop       =   2580
   ClientWidth     =   5835
   ControlBox      =   0   'False
   ForeColor       =   &H80000017&
   Icon            =   "frmTimedMsgBox.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUserText 
      Height          =   300
      Left            =   690
      TabIndex        =   6
      Top             =   930
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Help"
      Height          =   360
      Index           =   3
      Left            =   3405
      TabIndex        =   5
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
      Height          =   360
      Index           =   2
      Left            =   2355
      TabIndex        =   4
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
      Height          =   360
      Index           =   1
      Left            =   1305
      TabIndex        =   3
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
      Height          =   360
      Index           =   0
      Left            =   255
      TabIndex        =   2
      Top             =   1530
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Timer tmrMTimer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   1200
   End
   Begin VB.Image imgClose 
      Height          =   315
      Left            =   5430
      ToolTipText     =   "Close"
      Top             =   45
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   2
      Left            =   5295
      Picture         =   "frmTimedMsgBox.frx":000C
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   1
      Left            =   4935
      Picture         =   "frmTimedMsgBox.frx":0410
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgX 
      Height          =   315
      Index           =   0
      Left            =   4575
      Picture         =   "frmTimedMsgBox.frx":0837
      Top             =   1245
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label txtMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1020
      TabIndex        =   1
      Top             =   495
      UseMnemonic     =   0   'False
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H009EF5F3&
      BackStyle       =   0  'Transparent
      Caption         =   "<Title>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   375
      TabIndex        =   0
      Top             =   60
      Width           =   3735
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/*************************************/
'/* Author: Morgan Haueisen
'/*         morganh@hartcom.net
'/* Copyright (c) 2003-2004
'/*************************************/
'Legal:
'        This is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, whole or in part, as source code or in binary form, alone or
'        as part of a larger distribution or product, is forbidden for any commercial or for-profit
'        use without the author's explicit written permission.
'
'        Redistribution of this code, as source code or in binary form, with or without
'        modification, is permitted provided that the following conditions are met:
'
'        Redistributions of source code must include this list of conditions, and the following
'        acknowledgment:
'
'        This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.
'/* Used for Manifest files (Win XP style controls)
'/* Used to keep form always on top
'/* Used to get screen size
Private Type RECT
    Left                                     As Long
    Top                                      As Long
    Right                                    As Long
    bottom                                   As Long
End Type
Private Const SPI_GETWORKAREA            As Long = 48
'/* Used to get positions of cursor
Private Type POINTAPI
    x                                        As Long
    y                                        As Long
End Type
Private CursorXY                         As POINTAPI
'/* Button and Icon types
Public Enum ShowIconTypes
    None_i = 0
    vbCritical = 16         '/* Display Critical Message icon.
    vbQuestion = 32         '/* Display Warning Query icon.
    vbExclamation = 48      '/* Display Warning Message icon.
    vbInformation = 64      '/* Display Information Message icon.
    WinLogo_i = 128         '/* Display WinLogo icon.
    Folder_i = 144          '/* Display Folder icon.
    Printer_i = 160         '/* Display Printer icon.
    Find_i = 176            '/* Display Find icon.
    Save_i = 240            '/* Display Save icon.
    Hourglass_i = 80        '/* Display Hourglass icon.
    vbDefaultButton1 = 0    '/* First button is default.
    vbDefaultButton2 = 256  '/* Second button is default.
    vbDefaultButton3 = 512  '/* Third button is default.
    vbDefaultButton4 = 768  '/* Fourth button is default.
    vbOKCancel = 1          '/* Display OK and Cancel buttons.
    vbAbortRetryIgnore = 2  '/* Display Abort, Retry, and Ignore buttons.
    vbYesNoCancel = 3       '/* Display Yes, No, and Cancel buttons.
    vbYesNo = 4             '/* Display Yes and No buttons.
    vbRetryCancel = 5       '/* Display Retry and Cancel buttons.
    vbOkButton = 6          '/* Display OK button only.
    vbMsgBoxHelpButton = 16384 '/* Display the Help button
    vbHelp = 8              '/* Help button pressed
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private None_i, vbCritical, vbQuestion, vbExclamation, vbInformation, WinLogo_i, Folder_i, Printer_i, Find_i, Save_i, Hourglass_i, vbDefaultButton1, vbDefaultButton2, vbDefaultButton3, vbDefaultButton4
Private vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel, vbOkButton, vbMsgBoxHelpButton, vbHelp
#End If
'/* Used for moving the form around by draging the caption bar
'/* Used to draw the form's border
'/* Used to round the corners of the form and make trasnparent
'/* Used to play system sounds
Private Const MB_IconAsterisk            As Long = &H10&
Private Const MB_IconQuestion            As Long = &H20&
Private Const MB_IconExclamation         As Long = &H30&
Private Const MB_IconInformation         As Long = &H40&
'/* Used to draw system icons
Private Enum SystemIconConstants
    IDI_Application = 32512
    IDI_Error = 32513       'vbCritical (Critical)
    IDI_Question = 32514    'vbQuestion
    IDI_Warning = 32515     'vbExlamation (Exclamation)
    IDI_Information = 32516 'vbInformation (Asterisk)
    IDI_WinLogo = 32517
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private IDI_Application, IDI_Error, IDI_Question, IDI_Warning, IDI_Information, IDI_WinLogo
#End If
'/* Used to draw system icons from Shell32.dll
'/* GradientFill API - Requires Windows 2000 or later; Requires Windows 98 or later
Private Type GRADIENT_TRIANGLE
    Vertex1                                  As Long
    Vertex2                                  As Long
    Vertex3                                  As Long
End Type
Private Type TRIVERTEX
    x                                        As Long
    y                                        As Long
    Red                                      As Integer    '/* Ushort value (-256 to 0)
    Green                                    As Integer    '/* Ushort value (-256 to 0)
    Blue                                     As Integer    '/* Ushort value (-256 to 0)
    Alpha                                    As Integer    '/* Ushort value (-256 to 0)
End Type
Private Const GRADIENT_FILL_TRIANGLE     As Long = &H2&
'/* Operating system version information
Private Type OSVersionInfo
    OSVSize                                  As Long
    dwVerMajor                               As Long
    dwVerMinor                               As Long
    dwBuildNumber                            As Long
    PlatformID                               As Long
    szCSDVersion                             As String * 128
End Type
'/* Form Variables
Private oStandardIcon                    As Long
Private oCaption                         As String
Private oAutoCloseSeconds                As Long
Private oButtonResponse                  As Integer
Private oButtonFocus                     As Byte
Private oNonModal                        As Boolean
Private oInputBox                        As Boolean
Private oCountDown                       As Long
Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cX As Long, _
                                                    ByVal cY As Long, _
                                                    ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal Left As Long, _
                                                ByVal Top As Long, _
                                                ByVal Right As Long, _
                                                ByVal bottom As Long, _
                                                ByVal EllipseWidth As Long, _
                                                ByVal EllipseHeight As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal RectX1 As Long, _
                                                         ByVal RectY1 As Long, _
                                                         ByVal RectX2 As Long, _
                                                         ByVal RectY2 As Long, _
                                                         ByVal EllipseWidth As Long, _
                                                         ByVal EllipseHeight As Long) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Private Declare Function LoadStandardIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, _
                                                                          ByVal lpIconNum As SystemIconConstants) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, _
                                                ByVal x As Long, _
                                                ByVal y As Long, _
                                                ByVal hIcon As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, _
                                                                                 ByVal nIconIndex As Long, _
                                                                                 phiconLarge As Long, _
                                                                                 phiconSmall As Long, _
                                                                                 ByVal nIcons As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                                                                        ByVal nSize As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, _
                                                                                  pVertex As TRIVERTEX, _
                                                                                  ByVal dwNumVertex As Long, _
                                                                                  pMesh As GRADIENT_TRIANGLE, _
                                                                                  ByVal dwNumMesh As Long, _
                                                                                  ByVal dwMode As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long
Private Sub CheckIfLoaded()
Dim Frm As Form
    On Local Error Resume Next
    For Each Frm In Forms
        If LCase$(Frm.Name) = "frmmsgbox" Then
            Unload Frm
            Exit For
        End If
    Next Frm
End Sub
Private Sub cmdButton_Click(index As Integer)
    oButtonResponse = cmdButton(index).Tag
    Me.Hide
    If oNonModal Then
        Unload Me
    End If
End Sub
Private Sub DisplayInputBox(ByVal sPrompt As String, _
                            ByVal sTitle As String, _
                            Optional ByVal sDefault As String = vbNullString, _
                            Optional ByVal bShowClose As Boolean = True, _
                            Optional ByVal bCenter As Boolean = False, _
                            Optional sFont = "Tahoma")
Dim lPosX  As Long
Dim lPosY  As Long
Dim lWidth As Long
Dim I      As Byte
Dim hIcon  As Long
'/* Set defaults
    On Error Resume Next
    With Me
        .ScaleMode = vbPixels
        .DrawWidth = 1
        .FillStyle = 1
        .Font = sFont
    End With 'Me
    txtMessage.Font = sFont
    txtMessage.FontSize = 10
    lblMTitle.Font = sFont
    imgClose.Picture = imgX(0).Picture
    On Error GoTo 0
'/* Get display position from mouse position
    GetCursorPos CursorXY
    lPosX = CursorXY.x * Screen.TwipsPerPixelX
    lPosY = CursorXY.y * Screen.TwipsPerPixelY
    oCaption = sTitle
    txtUserText.Text = sDefault
'/* Resize the Form's width to fit the title bar/messagebox width
    Me.FontSize = 10
    lWidth = 5000
    Me.FontSize = 8
    If lWidth < (Me.TextWidth(sPrompt) + 90) * Screen.TwipsPerPixelX Then
        lWidth = (Me.TextWidth(sPrompt) + 90) * Screen.TwipsPerPixelX
    End If
    lblMTitle.Caption = sTitle
    Me.Width = lWidth
    Me.Height = 800
'/* Resize the Form's height based on the amount of text to display
    txtMessage.Move 8, 40, Me.ScaleWidth - 20, Me.ScaleHeight - 50
    txtMessage.Caption = sPrompt
    If txtMessage.Top + txtMessage.Height >= Me.ScaleHeight - 10 Then
        Me.Height = (txtMessage.Top + txtMessage.Height + 10) * Screen.TwipsPerPixelY
    End If
    txtUserText.Move 25, txtMessage.Top + txtMessage.Height + 10, txtMessage.Width - 25
'/* Locate Buttons and resize Form if required
    If Val(cmdButton(0).Tag) > 0 Or Val(cmdButton(3).Tag) > 0 Then
'/* How many buttons are visible?
        If Val(cmdButton(1).Tag) > 0 Then
            I = 1
        End If
        If Val(cmdButton(2).Tag) > 0 Then
            I = 2
        End If
        If Val(cmdButton(3).Tag) > 0 Then
            I = 3
        End If
        cmdButton(0).Top = txtUserText.Top + txtUserText.Height + 10
        cmdButton(1).Top = txtUserText.Top + txtUserText.Height + 10
        cmdButton(2).Top = txtUserText.Top + txtUserText.Height + 10
        cmdButton(3).Top = txtUserText.Top + txtUserText.Height + 10
        If Me.Width < (cmdButton(I).Left + cmdButton(I).Width + 15) * Screen.TwipsPerPixelX Then
            Me.Width = (cmdButton(I).Left + cmdButton(I).Width + 15) * Screen.TwipsPerPixelX
        End If
        Me.Height = (cmdButton(0).Top + cmdButton(0).Height + 10) * Screen.TwipsPerPixelY
    End If
'/* Show or don't show the close button
    imgClose.Visible = bShowClose
'/* Locate title bar and close button
    imgClose.Move (Me.ScaleWidth - imgClose.Width) - 8, 4
    lblMTitle.Move 2, 5, Me.ScaleWidth, 25
    GradientFill
'/* Draw box around Title Bar
    Me.Line (0, 0)-(Me.ScaleWidth, 25), &HB1FFFF, BF
'/* Draw border around the Form
    With Me
        .ForeColor = &H80000015
        RoundRect .hdc, 0, 0, (.Width / Screen.TwipsPerPixelX) - 1, (.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
        .ForeColor = &H8000000F
        RoundRect .hdc, 1, 1, (.Width / Screen.TwipsPerPixelX) - 2, (.Height / Screen.TwipsPerPixelY) - 2, CLng(25), CLng(25)
'/* Make corners transparent
        SetWindowRgn .hwnd, CreateRoundRectRgn(0, 0, (.Width / Screen.TwipsPerPixelX), (.Height / Screen.TwipsPerPixelY), 25, 25), True
'/* Position form on screen
    End With 'Me
    If Not bCenter Then
        Me.Move lPosX, lPosY
    End If
    PositionForm bCenter
    hIcon = LoadStandardIcon(0&, IDI_Question)
    DrawIcon Me.hdc, 4&, 4&, hIcon
    DestroyIcon hIcon
    txtUserText.SelStart = 0
    txtUserText.SelLength = Len(txtUserText.Text)
End Sub
Private Sub DisplayMessage(ByVal sText As String, _
                           Optional sIcon As ShowIconTypes = None_i, _
                           Optional ByVal sTitle As String = vbNullString, _
                           Optional ByVal lAutoCloseSeconds As Long = 0, _
                           Optional ByVal bShowClose As Boolean = True, _
                           Optional ByVal bCenter As Boolean = False, _
                           Optional lWidth As Long = -1, _
                           Optional sFont = "Tahoma", _
                           Optional OwnerForm As Form)
Dim lPosX       As Long
Dim lPosY       As Long
Dim Shell32Icon As Boolean
Dim I           As Byte
Dim hIcon       As Long
'/* Set defaults
    On Error Resume Next
    With Me
        .ScaleMode = vbPixels
        .DrawWidth = 1
        .FillStyle = 1
        .Font = sFont
    End With 'Me
    txtMessage.Font = sFont
    txtMessage.FontSize = 10
    lblMTitle.Font = sFont
    imgClose.Picture = imgX(0).Picture
    On Error GoTo 0
'/* Get display position from mouse position
    GetCursorPos CursorXY
    lPosX = CursorXY.x * Screen.TwipsPerPixelX
    lPosY = CursorXY.y * Screen.TwipsPerPixelY
'/* Set Title bar
    Select Case sIcon
    Case vbInformation '/* The "i" icon - Information
        If sTitle = vbNullString Then
            sTitle = "Information"
        End If
        MessageBeep MB_IconInformation
        oStandardIcon = IDI_Information
    Case vbCritical '/* The "x" icon - Critical
        If sTitle = vbNullString Then
            sTitle = "ERROR!"
        End If
        MessageBeep MB_IconAsterisk
        oStandardIcon = IDI_Error
    Case vbExclamation '/* The "!" icon - Exclamation
        If sTitle = vbNullString Then
            sTitle = "Warning!"
        End If
        MessageBeep MB_IconExclamation
        oStandardIcon = IDI_Warning
    Case vbQuestion '/* The "?" icon - Question
        If sTitle = vbNullString Then
            sTitle = "Question?"
        End If
        MessageBeep MB_IconQuestion
        oStandardIcon = IDI_Question
    Case WinLogo_i '/* Winlogo icon
        oStandardIcon = IDI_WinLogo
    Case Printer_i '/* Printer icon
        If sTitle = vbNullString Then
            sTitle = "Printing.. Please Wait"
        End If
        MessageBeep MB_IconInformation
        oStandardIcon = 16
        Shell32Icon = True
    Case Folder_i '/* Open folder icon
        MessageBeep MB_IconInformation
        oStandardIcon = 4
        Shell32Icon = True
    Case Find_i '/* Find icon
        MessageBeep MB_IconInformation
        oStandardIcon = 22
        Shell32Icon = True
    Case Save_i '/* Save icon
        MessageBeep MB_IconInformation
        oStandardIcon = 6
        Shell32Icon = True
    Case Hourglass_i '/* Hourglass icon
        If sTitle = vbNullString Then
            sTitle = "Working.. Please Wait"
        End If
        MessageBeep MB_IconInformation
        oStandardIcon = 76
        Shell32Icon = True
    Case Else 'Use no icon
    End Select
    oCaption = sTitle
'/* Resize the Form's width to fit the title bar/messagebox width
    Me.FontSize = 10
    If lWidth = -1 Then
        lWidth = (Me.TextWidth(sText) + 20) * Screen.TwipsPerPixelX
        If lWidth > 5000 Then
            lWidth = 5000
        End If
    End If
    If lWidth < 1500 Then
        lWidth = 1500
    End If
    If lAutoCloseSeconds > 0 Then
        If sTitle > vbNullString Then
            sTitle = sTitle & " -" & CStr(lAutoCloseSeconds)
        Else
            sTitle = CStr(lAutoCloseSeconds)
        End If
    End If
    Me.FontSize = 8
    If lWidth < (Me.TextWidth(sTitle) + 90) * Screen.TwipsPerPixelX Then
        lWidth = (Me.TextWidth(sTitle) + 90) * Screen.TwipsPerPixelX
    End If
    lblMTitle.Caption = sTitle
    Me.Width = lWidth
    Me.Height = 800
'/* Resize the Form's height based on the amount of text to display
    txtMessage.Move 8, 40, Me.ScaleWidth - 20, Me.ScaleHeight - 50
    txtMessage.Caption = sText
    If txtMessage.Top + txtMessage.Height >= Me.ScaleHeight - 10 Then
        Me.Height = (txtMessage.Top + txtMessage.Height + 10) * Screen.TwipsPerPixelY
    End If
'/* Locate Buttons and resize Form if required
    If Val(cmdButton(0).Tag) > 0 Or Val(cmdButton(3).Tag) > 0 Then
'/* How many buttons are visible?
        If Val(cmdButton(1).Tag) > 0 Then
            I = 1
        End If
        If Val(cmdButton(2).Tag) > 0 Then
            I = 2
        End If
        If Val(cmdButton(3).Tag) > 0 Then
            I = 3
        End If
'Me.Height = Me.Height + 500
        cmdButton(0).Top = txtMessage.Top + txtMessage.Height + 10
        cmdButton(1).Top = txtMessage.Top + txtMessage.Height + 10
        cmdButton(2).Top = txtMessage.Top + txtMessage.Height + 10
        cmdButton(3).Top = txtMessage.Top + txtMessage.Height + 10
        If Me.Width < (cmdButton(I).Left + cmdButton(I).Width + 15) * Screen.TwipsPerPixelX Then
            Me.Width = (cmdButton(I).Left + cmdButton(I).Width + 15) * Screen.TwipsPerPixelX
        End If
        Me.Height = (cmdButton(0).Top + cmdButton(0).Height + 10) * Screen.TwipsPerPixelY
'Me.Height + 500
    End If
'/* Show or don't show the close button
    imgClose.Visible = bShowClose
'/* Enable or disable auto close timer
    If lAutoCloseSeconds = 0 Then
        tmrMTimer1.Enabled = False
    Else
        If oCaption > vbNullString Then
            oCaption = oCaption & " -"
        End If
        oAutoCloseSeconds = lAutoCloseSeconds
        tmrMTimer1.Enabled = True
    End If
    GradientFill
'/* Locate title bar and close button
    imgClose.Move (Me.ScaleWidth - imgClose.Width) - 8, 4
    lblMTitle.Move 2, 5, Me.ScaleWidth, 25
'/* Draw box around Title Bar
    Me.Line (0, 0)-(Me.ScaleWidth, 25), &HB1FFFF, BF
'/* Draw border around the Form
    With Me
        .ForeColor = &H80000015
        RoundRect .hdc, 0, 0, (.Width / Screen.TwipsPerPixelX) - 1, (.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
        .ForeColor = &H8000000F
        RoundRect .hdc, 1, 1, (.Width / Screen.TwipsPerPixelX) - 2, (.Height / Screen.TwipsPerPixelY) - 2, CLng(25), CLng(25)
'/* Draw Icon
    End With 'Me
    If Shell32Icon Then
        LoadShell32Icon oStandardIcon
    Else
        hIcon = LoadStandardIcon(0&, oStandardIcon)
        DrawIcon Me.hdc, 4&, 4&, hIcon
        DestroyIcon hIcon
    End If
'/* Make corners transparent
    SetWindowRgn Me.hwnd, CreateRoundRectRgn(0, 0, (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), 25, 25), True
'/* Position form on screen
    If Not bCenter Then
        Me.Move lPosX, lPosY
    End If
    PositionForm bCenter
End Sub
Private Sub Form_Activate()
    With cmdButton(oButtonFocus)
        If .Visible Then
            .SetFocus
            M__CenterMouseOn (.hwnd)
        End If
    End With 'cmdButton(oButtonFocus)
    If oInputBox Then
        txtUserText.SetFocus
    End If
    Me.ZOrder
End Sub
Private Sub Form_Initialize()
'/* Used for Manifest files (Win XP style controls)
    InitCommonControls
End Sub
Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)
    imgClose.Picture = imgX(0).Picture
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmMsgBox = Nothing
End Sub
Private Sub GradientFill()
Dim vert(4) As TRIVERTEX
Dim gTRi(1) As GRADIENT_TRIANGLE
Dim iOSver  As Long
Dim OSV     As OSVersionInfo
'/* Get OS compatability flag
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
'/* Win 98/ME
        If OSV.PlatformID = 1 Then
            If OSV.dwVerMinor >= 10 Then
                iOSver = 1
            End If
        End If
'/* Win 2000/XP
        If OSV.PlatformID = 2 Then
            If OSV.dwVerMajor >= 5 Then
                iOSver = 2
            End If
        End If
    End If
'/* Requires Windows 2000 or later; Requires Windows 98/ME
    If iOSver = 0 Then
        GoTo GradientFillExit
    End If
    Me.AutoRedraw = True
'/* Top Left Trangle
    With vert(0)
        .x = 0
        .y = 0
        .Red = -256&
        .Green = -256&
        .Blue = -256&
        .Alpha = 0&
'/* Top Right Trangle
    End With 'vert(0)
    With vert(1)
        .x = Me.ScaleWidth * 2
        .y = 0
        .Red = -100
        .Green = -256&
        .Blue = -256&
        .Alpha = 0&
'/* Bottom Right Trangle
    End With 'vert(1)
    With vert(2)
        .x = Me.ScaleWidth * 3
        .y = Me.ScaleHeight * 3
        .Red = -100
        .Green = -256&
        .Blue = 0&
        .Alpha = 0&
'/* Bottom Left Trangle
    End With 'vert(2)
    With vert(3)
        .x = 0
        .y = Me.ScaleHeight * 2
        .Red = -256&
        .Green = -256&
        .Blue = -256&
        .Alpha = 0&
    End With 'vert(3)
    With gTRi(0)
        .Vertex1 = 0
        .Vertex2 = 1
        .Vertex3 = 2
    End With 'gTRi(0)
    With gTRi(1)
        .Vertex1 = 0
        .Vertex2 = 2
        .Vertex3 = 3
    End With 'gTRi(1)
    GradientFillTriangle Me.hdc, vert(0), 4, gTRi(0), 2, GRADIENT_FILL_TRIANGLE
GradientFillExit:
End Sub
Private Sub imgClose_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)
    If Button = vbLeftButton Then
        imgClose.Picture = imgX(2).Picture
    End If
End Sub
Private Sub imgClose_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               y As Single)
    If imgClose.Picture <> imgX(2).Picture Then
        imgClose.Picture = imgX(1).Picture
    End If
End Sub
Private Sub imgClose_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
    Unload Me
End Sub
Private Sub lblMTitle_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    If Me.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage Me.hwnd, &HA1, 2, 0&
    End If
End Sub
Private Sub lblMTitle_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
    imgClose.Picture = imgX(0).Picture
End Sub
Private Sub LoadShell32Icon(ByVal index As Long)
Dim SysDir    As String
Dim CurFile   As String
Dim hIcon     As Long
Dim IconCount As Long
Dim Rv        As Long
    SysDir = Space$(260)
    Rv = GetSystemDirectory(SysDir, 260)
    SysDir = Left$(SysDir, Rv) & "\"
    CurFile = SysDir & "Shell32.dll"
    IconCount = ExtractIconEx(CurFile, -1, 0, 0, 0)
    If IconCount >= index Then
        ExtractIconEx CurFile, index, hIcon, 0&, 1&
        DrawIcon Me.hdc, 4&, 4&, hIcon
        DestroyIcon hIcon
    End If
End Sub
Private Sub PositionForm(ByVal center As Boolean)
Dim rc       As RECT
Dim T        As Long
Dim B        As Long
Dim L        As Long
Dim r        As Long
Dim mT       As Long
Dim mL       As Long
Const offset As Long = 150
'/* Get screen size
    SystemParametersInfo SPI_GETWORKAREA, 0&, rc, 0&
    With rc
        T = .Top * Screen.TwipsPerPixelY
        B = .bottom * Screen.TwipsPerPixelY
        L = .Left * Screen.TwipsPerPixelX
        r = .Right * Screen.TwipsPerPixelX
    End With 'Rc
    If center Then
'/* Center Form on screen
        mT = Abs((B / 2) - (Me.Height / 2))
        mL = Abs((r / 2) - (Me.Width / 2))
        If mT < T Then
            mT = T
        End If
        If mT > B - Me.Height Then
            mT = B - Me.Height
        End If
        If mL < L Then
            mL = L
        End If
    Else
'/* Make sure all the Form is on the screen
        mT = Me.Top
        mL = Me.Left
        If Me.Top - offset < T Then
            mT = T + offset
        End If
        If Me.Left - offset < L Then
            mL = L + offset
        End If
        If Me.Top + Me.Height + offset > B Then
            mT = B - Me.Height - offset
        End If
        If Me.Left + Me.Width + offset > r Then
            mL = r - Me.Width - offset
        End If
    End If
    Me.Move mL, mT
End Sub
Public Function SInputBox(ByVal sPrompt As String, _
                          Optional sTitle As String = vbNullString, _
                          Optional sDefault As String = vbNullString, _
                          Optional ByVal bShowClose As Boolean = False, _
                          Optional ByVal bCenter As Boolean = True, _
                          Optional ByVal sFont As String = "Tahoma") As String
    CheckIfLoaded
    With cmdButton(0)
        .Visible = True
        .Caption = "Ok"
        .Tag = vbOK
        .Default = True
    End With
    With cmdButton(1)
        .Visible = True
        .Caption = "Cancel"
        .Tag = vbCancel
        .Cancel = True
    End With
    txtUserText.Visible = True
    oInputBox = True
    If sTitle = vbNullString Then
        sTitle = App.Title
    End If
    DisplayInputBox sPrompt, sTitle, sDefault, bShowClose, bCenter, sFont
    DoEvents
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Me.Show vbModal
    If oButtonResponse = vbCancel Then
        SInputBox = vbNullString
    Else
        SInputBox = txtUserText.Text
    End If
    DoEvents
    Unload Me
End Function
Public Function SMessage(ByVal sText As String, _
                         Optional ByVal sIcon As ShowIconTypes = None_i, _
                         Optional ByVal sTitle As String = vbNullString, _
                         Optional ByVal lAutoCloseSeconds As Long = 0, _
                         Optional ByVal bShowClose As Boolean = True, _
                         Optional ByVal bCenter As Boolean = True, _
                         Optional ByVal lWidth As Long = -1, _
                         Optional ByVal sFont As String = "Tahoma", _
                         Optional OwnerForm As Form) As Integer
Dim MsgType As ShowIconTypes
Dim TesthDC As Boolean
    CheckIfLoaded
'/* Separate Message Icon from input
    MsgType = sIcon And 240
'/* Only the OK button allowed for a non-modal message box
    If (sIcon And 15) = vbOkButton Then
        With cmdButton(0)
            .Visible = True
            .Caption = "Ok"
            .Tag = vbOK
            .Cancel = True
        End With 'cmdButton(0)
        oNonModal = True
    End If
    DisplayMessage sText, MsgType, sTitle, lAutoCloseSeconds, bShowClose, bCenter, lWidth, sFont
    DoEvents
    On Local Error Resume Next
    TesthDC = OwnerForm.HasDC
    If Not TesthDC Then
        SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    End If
    Show , OwnerForm
    DoEvents
    Me.ZOrder
End Function
Public Function SMessageModal(ByVal sText As String, _
                              Optional ByVal sIcon As ShowIconTypes = None_i, _
                              Optional ByVal sTitle As String = vbNullString, _
                              Optional ByVal lAutoCloseSeconds As Long = 0, _
                              Optional ByVal bShowClose As Boolean = True, _
                              Optional ByVal bCenter As Boolean = True, _
                              Optional ByVal lWidth As Long = -1, _
                              Optional ByVal sFont As String = "Tahoma", _
                              Optional OwnerForm As Form) As Integer
Dim MsgType As ShowIconTypes
Dim TesthDC As Boolean
    CheckIfLoaded
'/* Separate Message Icon from input
    MsgType = sIcon And 240
'/* Separate button default from input
    Select Case sIcon And 1792
    Case 256
        oButtonFocus = 1 '/* Second button is default.
    Case 512
        oButtonFocus = 2 '/* Third button is default.
    Case 768
        oButtonFocus = 3 '/* Fourth button is default.
    Case Else
        oButtonFocus = 0 '/* First button is default.
    End Select
'/* Separate Button type from input
    If lAutoCloseSeconds = 0 Then
        bShowClose = True
    End If
    Select Case sIcon And 15
    Case vbRetryCancel
        With cmdButton(0)
            .Visible = True
            .Caption = "Retry"
            .Tag = vbRetry
        End With 'cmdButton(0)
        With cmdButton(1)
            .Visible = True
            .Caption = "Cancel"
            .Tag = vbCancel
            .Cancel = True
        End With 'cmdButton(1)
        bShowClose = False
        lAutoCloseSeconds = 0
    Case vbYesNo
        With cmdButton(0)
            .Visible = True
            .Caption = "Yes"
            .Tag = vbYes
        End With 'cmdButton(0)
        With cmdButton(1)
            .Visible = True
            .Caption = "No"
            .Tag = vbNo
        End With 'cmdButton(1)
        bShowClose = False
        lAutoCloseSeconds = 0
    Case vbYesNoCancel
        With cmdButton(0)
            .Visible = True
            .Caption = "Yes"
            .Tag = vbYes
        End With 'cmdButton(0)
        With cmdButton(1)
            .Visible = True
            .Caption = "No"
            .Tag = vbNo
        End With 'cmdButton(1)
        With cmdButton(2)
            .Visible = True
            .Caption = "Cancel"
            .Tag = vbCancel
            .Cancel = True
        End With 'cmdButton(2)
        bShowClose = False
        lAutoCloseSeconds = 0
    Case vbAbortRetryIgnore
        With cmdButton(0)
            .Visible = True
            .Caption = "Abort"
            .Tag = vbAbort
        End With 'cmdButton(0)
        With cmdButton(1)
            .Visible = True
            .Caption = "Retry"
            .Tag = vbRetry
        End With 'cmdButton(1)
        With cmdButton(2)
            .Visible = True
            .Caption = "Ignore"
            .Tag = vbIgnore
        End With 'cmdButton(2)
        bShowClose = False
    Case vbOKCancel
        With cmdButton(0)
            .Visible = True
            .Caption = "Ok"
            .Tag = vbOK
        End With 'cmdButton(0)
        With cmdButton(1)
            .Visible = True
            .Caption = "Cancel"
            .Tag = vbCancel
            .Cancel = True
        End With 'cmdButton(1)
        bShowClose = False
        lAutoCloseSeconds = 0
    Case Else
        With cmdButton(0)
            .Visible = True
            .Caption = "Ok"
            .Tag = vbOK
            .Cancel = True
        End With 'cmdButton(0)
    End Select
'/* Show Help button?
    If sIcon And 16384 Then
        cmdButton(3).Visible = True
        cmdButton(3).Tag = vbHelp
    End If
    DisplayMessage sText, MsgType, sTitle, lAutoCloseSeconds, bShowClose, bCenter, lWidth, sFont
    DoEvents
    On Local Error Resume Next
    TesthDC = OwnerForm.HasDC
    If Not TesthDC Then
        SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    End If
    Me.Show vbModal
    SMessageModal = oButtonResponse
    DoEvents
    Unload Me
End Function
Private Sub tmrMTimer1_Timer()
    oCountDown = oCountDown + 1
    If oCountDown >= oAutoCloseSeconds Then
        Unload Me
    Else
        lblMTitle.Caption = oCaption & CStr(oAutoCloseSeconds - oCountDown)
    End If
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:30:55 PM) 144 + 730 = 874 Lines Thanks Ulli for inspiration and lots of code.


