VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlayList 
   BorderStyle     =   0  'None
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   ControlBox      =   0   'False
   Icon            =   "frmPlayList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayList.frx":57E2
   ScaleHeight     =   4860
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MP3JukeBox.isButton cmdClearList 
      Height          =   435
      Left            =   2235
      TabIndex        =   4
      Top             =   4320
      Width           =   855
      _extentx        =   1508
      _extenty        =   767
      style           =   7
      caption         =   "Clear List"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmPlayList.frx":49DA6
   End
   Begin MP3JukeBox.isButton cmdSaveList 
      Height          =   435
      Left            =   1200
      TabIndex        =   3
      Top             =   4320
      Width           =   855
      _extentx        =   1508
      _extenty        =   767
      style           =   7
      caption         =   "Save List"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmPlayList.frx":49DCA
   End
   Begin MP3JukeBox.isButton cmdSlidePlay 
      Height          =   435
      Left            =   225
      TabIndex        =   2
      Top             =   4320
      Width           =   855
      _extentx        =   1508
      _extenty        =   767
      style           =   7
      caption         =   "Play"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmPlayList.frx":49DEE
   End
   Begin MP3JukeBox.isButton cmdClose 
      Height          =   435
      Left            =   3240
      TabIndex        =   1
      Top             =   4320
      Width           =   855
      _extentx        =   1508
      _extenty        =   767
      style           =   7
      caption         =   "Close"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmPlayList.frx":49E12
   End
   Begin MSComDlg.CommonDialog cdlCommonDialog3 
      Left            =   1545
      Top             =   1125
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save Play List"
   End
   Begin MSComctlLib.ListView lsvSlideListView 
      Height          =   4110
      Left            =   180
      TabIndex        =   0
      Top             =   165
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   7250
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   65280
      BackColor       =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   13300
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   3422
      EndProperty
   End
End
Attribute VB_Name = "frmPlayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API Delcares
Private bytRegion(191)   As Byte
Private nBytes           As Long
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, _
                                                      ByVal nCount As Long, _
                                                      lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Boolean) As Long
Private Sub cmdClearList_Click()
    On Error GoTo ErrorTrap
    lsvSlideListView.ListItems.Clear
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmSlide.Command3_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdClose_Click()
    On Error GoTo ErrorTrap
    blnPlayOpen = False
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmSlide.Command4_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdSaveList_Click()
Dim FileNum As Long
Dim x       As Long
    On Error GoTo ERRHANDLER
    If lsvSlideListView.ListItems.Count = 0 Then
        frmMsgBox.SMessageModal "Nothing to save!", 48, "Save Error"
        GoTo cmdSaveListExit
    End If
    frmSave.Show vbModal, Me
    If blnCancel Then
        blnCancel = False
        GoTo cmdSaveListExit
    End If
    If FileExists1(Savefilename) = True Then
        If MsgBox("File already exits." & vbNewLine & "Overwrite?", 4, "Save Error") = vbNo Then
            GoTo cmdSaveListExit
        Else
            Kill Savefilename
        End If
    End If
    FileNum = FreeFile()
    Open Savefilename For Output As #FileNum
    Print #FileNum, SongDir
    With lsvSlideListView
        For x = 1 To .ListItems.Count
            Print #FileNum, .ListItems(x).Text
        Next x
    End With
    Close #FileNum
cmdSaveListExit:
Exit Sub
ERRHANDLER:
    frmMsgBox.SMessageModal "Unexpected error in procedure: cmdSaveList_Click" & vbNewLine & _
          "Error #" & Err.Number & ": " & Err.description, vbCritical + vbOKOnly, App.Title
End Sub
Private Sub cmdSlidePlay_Click()
Dim I        As Long
Dim listitem As listitem
Dim Time     As Single
    On Error GoTo ErrorTrap
    If lsvSlideListView.ListItems.Count = 0 Then
        frmRats.Show , Me
        Time = Timer
        Do
            DoEvents
        Loop Until Timer > Time + 1 Or Timer < Time
        frmRats.Hide
        GoTo cmdSlidePlay_ClickExit
    End If
    frmMP3Jukebox.lsvListView1.ListItems.Clear
    For I = 1 To lsvSlideListView.ListItems.Count
        Set listitem = frmMP3Jukebox.lsvListView1.ListItems.Add(, , lsvSlideListView.ListItems(I).Text)
' add a listitem
    Next I
    blnPlayOpen = False
    Zx_Tkey = OldKey
    Unload Me
    frmMP3Jukebox.PlayAll
cmdSlidePlay_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmSlide.Command1_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
Dim rgnMain As Long
    On Error GoTo ErrorTrap
    nBytes = 192
    LoadBytes
    rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
    SetWindowRgn Me.hwnd, rgnMain, True
 
    Me.Top = frmMP3Jukebox.Top
    Me.Left = frmMP3Jukebox.Left + frmMP3Jukebox.Width - 80
    blnPlayOpen = True
    Me.Show
    OldKey = Zx_Tkey
    DoEvents
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmPlayList.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub LoadBytes()
    On Error GoTo ErrorTrap
    bytRegion(0) = 32
    bytRegion(4) = 1
    bytRegion(8) = 10
    bytRegion(12) = 160
    bytRegion(16) = 1
    bytRegion(20) = 1
    bytRegion(24) = 30
    bytRegion(25) = 1
    bytRegion(28) = 66
    bytRegion(29) = 1
    bytRegion(32) = 6
    bytRegion(36) = 1
    bytRegion(40) = 25
    bytRegion(41) = 1
    bytRegion(44) = 2
    bytRegion(48) = 5
    bytRegion(52) = 2
    bytRegion(56) = 26
    bytRegion(57) = 1
    bytRegion(60) = 3
    bytRegion(64) = 4
    bytRegion(68) = 3
    bytRegion(72) = 27
    bytRegion(73) = 1
    bytRegion(76) = 4
    bytRegion(80) = 3
    bytRegion(84) = 4
    bytRegion(88) = 28
    bytRegion(89) = 1
    bytRegion(92) = 5
    bytRegion(96) = 2
    bytRegion(100) = 5
    bytRegion(104) = 29
    bytRegion(105) = 1
    bytRegion(108) = 6
    bytRegion(112) = 1
    bytRegion(116) = 6
    bytRegion(120) = 30
    bytRegion(121) = 1
    bytRegion(124) = 61
    bytRegion(125) = 1
    bytRegion(128) = 2
    bytRegion(132) = 61
    bytRegion(133) = 1
    bytRegion(136) = 29
    bytRegion(137) = 1
    bytRegion(140) = 62
    bytRegion(141) = 1
    bytRegion(144) = 3
    bytRegion(148) = 62
    bytRegion(149) = 1
    bytRegion(152) = 28
    bytRegion(153) = 1
    bytRegion(156) = 64
    bytRegion(157) = 1
    bytRegion(160) = 5
    bytRegion(164) = 64
    bytRegion(165) = 1
    bytRegion(168) = 26
    bytRegion(169) = 1
    bytRegion(172) = 65
    bytRegion(173) = 1
    bytRegion(176) = 6
    bytRegion(180) = 65
    bytRegion(181) = 1
    bytRegion(184) = 25
    bytRegion(185) = 1
    bytRegion(188) = 66
    bytRegion(189) = 1
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmPlayList.LoadBytes" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:25 PM) 6 + 206 = 212 Lines Thanks Ulli for inspiration and lots of code.


