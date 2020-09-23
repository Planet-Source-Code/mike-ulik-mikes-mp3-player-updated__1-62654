VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMP3Jukebox 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Midi Jukebox"
   ClientHeight    =   5625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4275
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMP3Jukebox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin MP3JukeBox.isButton cmdPlayAll 
      Height          =   375
      Left            =   225
      TabIndex        =   22
      Top             =   5055
      Width           =   3855
      _extentx        =   6800
      _extenty        =   661
      icon            =   "frmMP3Jukebox.frx":57E2
      style           =   7
      caption         =   "Play all listed files"
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":57FE
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdReloadFiles 
      Height          =   525
      Left            =   3150
      TabIndex        =   21
      Top             =   4500
      Width           =   930
      _extentx        =   1640
      _extenty        =   926
      icon            =   "frmMP3Jukebox.frx":5822
      style           =   7
      caption         =   "New Directory"
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":583E
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdRadomize 
      Height          =   525
      Left            =   2175
      TabIndex        =   20
      Top             =   4500
      Width           =   930
      _extentx        =   1640
      _extenty        =   926
      icon            =   "frmMP3Jukebox.frx":5862
      style           =   7
      caption         =   "Radomize List"
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":587E
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdLoadPlaylist 
      Height          =   525
      Left            =   1200
      TabIndex        =   19
      Top             =   4500
      Width           =   930
      _extentx        =   1640
      _extenty        =   926
      icon            =   "frmMP3Jukebox.frx":58A2
      style           =   7
      caption         =   "Load Playlist"
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":58BE
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdNewPlaylist 
      Height          =   525
      Left            =   225
      TabIndex        =   18
      Top             =   4500
      Width           =   930
      _extentx        =   1640
      _extenty        =   926
      icon            =   "frmMP3Jukebox.frx":58E2
      style           =   7
      caption         =   "New Playlist"
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":58FE
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdEFFECTS 
      Height          =   300
      Left            =   1680
      TabIndex        =   17
      Top             =   4140
      Width           =   945
      _extentx        =   1667
      _extenty        =   529
      icon            =   "frmMP3Jukebox.frx":5922
      style           =   7
      caption         =   "EFFECTS"
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":593E
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmfForward 
      Height          =   300
      Left            =   2340
      TabIndex        =   16
      Top             =   495
      Width           =   495
      _extentx        =   873
      _extenty        =   529
      icon            =   "frmMP3Jukebox.frx":5962
      style           =   7
      caption         =   " "
      iconalign       =   0
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":5ABE
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdStop 
      Height          =   300
      Left            =   1806
      TabIndex        =   15
      Top             =   495
      Width           =   495
      _extentx        =   873
      _extenty        =   529
      icon            =   "frmMP3Jukebox.frx":5AE6
      style           =   7
      caption         =   " "
      iconalign       =   0
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":5C42
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdPause 
      Height          =   300
      Left            =   1274
      TabIndex        =   14
      Top             =   495
      Width           =   495
      _extentx        =   873
      _extenty        =   529
      icon            =   "frmMP3Jukebox.frx":5C6A
      style           =   7
      caption         =   " "
      iconalign       =   0
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":5DC6
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdJukeBoxPlay 
      Height          =   300
      Left            =   742
      TabIndex        =   13
      Top             =   495
      Width           =   495
      _extentx        =   873
      _extenty        =   529
      icon            =   "frmMP3Jukebox.frx":5DEE
      style           =   7
      caption         =   " "
      iconalign       =   0
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":5F4A
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin MP3JukeBox.isButton cmdReverse 
      Height          =   300
      Left            =   210
      TabIndex        =   12
      Top             =   495
      Width           =   495
      _extentx        =   873
      _extenty        =   529
      icon            =   "frmMP3Jukebox.frx":5F72
      style           =   7
      caption         =   " "
      iconalign       =   0
      inonthemestyle  =   0
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      ttforecolor     =   0
      font            =   "frmMP3Jukebox.frx":60CE
      maskcolor       =   0
      roundedbordersbytheme=   0   'False
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3495
      Picture         =   "frmMP3Jukebox.frx":60F6
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   11
      Top             =   150
      Width           =   270
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3840
      Picture         =   "frmMP3Jukebox.frx":6528
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   10
      Top             =   150
      Width           =   270
   End
   Begin VB.Timer tmrJukeboxTimer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   3015
   End
   Begin VB.Timer tmrJukeboxTimer2 
      Interval        =   100
      Left            =   3600
      Top             =   2985
   End
   Begin VB.PictureBox picSpectrum 
      BackColor       =   &H00000000&
      Height          =   705
      Left            =   210
      ScaleHeight     =   645
      ScaleWidth      =   3795
      TabIndex        =   1
      Top             =   1125
      Width           =   3855
   End
   Begin VB.CheckBox chkLoop 
      BackColor       =   &H00000000&
      Caption         =   "Loop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   3150
      TabIndex        =   0
      Top             =   450
      Width           =   780
   End
   Begin MP3JukeBox.ctlSysTrayIcon ctlSysTrayIcon 
      Left            =   1575
      Top             =   2820
      _extentx        =   1799
      _extenty        =   926
      icon            =   "frmMP3Jukebox.frx":695A
      icontooltiptext =   ""
   End
   Begin MP3JukeBox.ctlEBSlider ctlJukeboxVolumn 
      Height          =   210
      Left            =   480
      TabIndex        =   2
      Top             =   810
      Width           =   3105
      _extentx        =   5477
      _extenty        =   370
      max             =   1024
      value           =   800
      slidercolor     =   255
   End
   Begin SysInfoLib.SysInfo snfSysInfo2 
      Left            =   525
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdlCommonDialog1 
      Left            =   1500
      Top             =   2085
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lsvListView1 
      Height          =   1845
      Left            =   210
      TabIndex        =   3
      Top             =   1905
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3254
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
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
         Object.Width           =   21167
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Powered by Fmod Ex"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   1500
      TabIndex        =   9
      Top             =   5430
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      TabIndex        =   8
      Top             =   3795
      Width           =   3855
   End
   Begin VB.Label lblTimeLenght 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      TabIndex        =   7
      Top             =   4140
      Width           =   1335
   End
   Begin VB.Label lblJukebox 
      BackColor       =   &H00000000&
      Caption         =   " Mike's MP3 Jukebox"
      BeginProperty Font 
         Name            =   "Black Chancery"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   150
      Width           =   2820
   End
   Begin VB.Label lblTimeplayed 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00.000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2745
      TabIndex        =   5
      Top             =   4140
      Width           =   1335
   End
   Begin VB.Label lblVolume 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3630
      TabIndex        =   4
      Top             =   795
      Width           =   420
   End
   Begin VB.Image imgImage1 
      Height          =   240
      Left            =   225
      Picture         =   "frmMP3Jukebox.frx":C14E
      Top             =   795
      Width           =   240
   End
End
Attribute VB_Name = "frmMP3Jukebox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API Delcares
Private Declare Function ExtCreateRegion Lib "gdi32" (lpXform As Any, _
                                                      ByVal nCount As Long, _
                                                      lpRgnData As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Boolean) As Long
Private bytRegion(623)   As Byte
Private nBytes           As Long
Private Type BITMAP
    bmType                   As Long
    bmWidth                  As Long
    bmHeight                 As Long
    bmWidthBytes             As Long
    bmPlanes                 As Integer
    bmBitsPixel              As Integer
    bmBits                   As Long
End Type
Private itmX             As listitem
'============================================================
' Routine Name: cmdEFFECTS_Click
' Description:
' Author: Mike
' Date: 8/25/2005 1:10:21 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Sub cmdEFFECTS_Click()
    On Error GoTo cmdEFFECTS_ClickError
    frmEffects.Show vbModal, Me
cmdEFFECTS_ClickExit:
Exit Sub
cmdEFFECTS_ClickError:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "Command1_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdFullSize_Click()
    On Error GoTo ErrorTrap
    DoEvents
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmJukebox.cmdFullSize_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdJukeBoxPlay_Click()
    On Error GoTo ErrorTrap
    PlaySong
cmdJukeBoxPlayExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.cmdPlay_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdLoadPlaylist_Click()
Dim FName As String
Dim fnum  As Long
Dim Title As String
Dim Tdir  As String
    On Error GoTo ErrorTrap
    With cdlCommonDialog1
        .CancelError = True
        .DefaultExt = "MCU"
        .Filter = "Play List (*.MCU)|*.MCU"
        .InitDir = iniMCUdirectory
        .ShowOpen
        FName = .Filename
    End With
    lsvListView1.ListItems.Clear
    If FileExists1(FName) = True Then
        fnum = FreeFile()
        Open FName For Input As #fnum
        Input #fnum, Tdir
        Do While Not EOF(fnum)
            Input #fnum, Title
            Set itmX = lsvListView1.ListItems.Add(, , Title)   ' add a listitem
        Loop
        iniLastDirectory = Tdir
        Close fnum
        If lsvListView1.ListItems.Count > 0 Then
            lsvListView1.ListItems(1).Selected = True
        End If
    End If
ErrorTrap:
End Sub
Private Sub cmdNewPlaylist_Click()
    On Error GoTo ErrorTrap
    frmPlayList.Show , Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.Command1_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdPause_Click()
Dim paused As Long
    On Error GoTo ErrorTrap
    If ChannelPlaying Then
        result = FMOD_Channel_GetPaused(ChannelPlaying, paused)
        ERRCHECK (result)
        If paused Then
            result = FMOD_Channel_SetPaused(ChannelPlaying, 0)
        Else
            result = FMOD_Channel_SetPaused(ChannelPlaying, 1)
        End If
        ERRCHECK (result)
    End If
CmdPauseExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.CmdPause_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdPlayAll_Click()
    On Error GoTo ErrorTrap
    PlayAll
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.cmdPlayAll_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdRadomize_Click()
Dim x          As Long
Dim lngRdm     As Long
Dim lngMax     As Long
Dim strOldName As String
Dim strOldID   As String
    On Error GoTo ErrorTrap
    Randomize
    lngMax = lsvListView1.ListItems.Count
    For x = 1 To lngMax
        lngRdm = Round((Rnd * lngMax) + 1)
        If lngRdm <= lngMax Then
            With lsvListView1
                If .ListItems(x).Selected = True Then
                    .ListItems(x).Selected = False
                    lngRandom = lngRdm
                    .ListItems(lngRandom).Selected = True
                    .SelectedItem.EnsureVisible
                    DoEvents
                End If
                strOldName = .ListItems(x).Text
                strOldID = .ListItems(x).SubItems(1)
                .ListItems(x).Text = .ListItems(lngRdm).Text
                .ListItems(x).SubItems(1) = .ListItems(lngRdm).SubItems(1)
                .ListItems(lngRdm).Text = strOldName
                .ListItems(lngRdm).SubItems(1) = strOldID
            End With
        End If
    Next x
    lsvListView1.ListItems(1).Selected = True
    lsvListView1.SelectedItem.EnsureVisible
    DoEvents
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.Command3_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdReloadFiles_Click()
Dim nDirs  As Long
Dim nFiles As Long
Dim sDir   As String
    On Error GoTo ErrorTrap
    frmODirectory.Show vbModal, Me
    If blnCancel Then
        blnCancel = False
        GoTo cmdReloadFiles_ClickExit
    End If
    DoEvents
    frmLoading.Show
    DoEvents
    sDir = Zx_Tkey
    Zx_Pattern = ".mid .kar"
    FindFile sDir, "*.*", nDirs, nFiles
    lsvListView1.Sorted = True
    frmLoading.Hide
    DoEvents
    lsvListView1.Sorted = False
    If lsvListView1.ListItems.Count > 0 Then
        lsvListView1.ListItems(1).Selected = True
    End If
cmdReloadFiles_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.Command4_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdReverse_Click()
    On Error GoTo ErrorTrap
    PlayPrevious
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.cmdReverse_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdStop_Click()
    On Error GoTo ErrorTrap
    If ChannelPlaying <> 0 Then
        result = FMOD_Channel_IsPlaying(ChannelPlaying, State)
        ERRCHECK (result)
        If State = 1 Then
            result = FMOD_Channel_Stop(ChannelPlaying)
            ERRCHECK (result)
        End If
        result = FMOD_Sound_Release(SoundPlaying)
        ERRCHECK (result)
        Sound = 0
    End If
    ChannelPlaying = 0
    SoundPlaying = 0
    State = 0
    tmrJukeboxTimer1.Enabled = False
    lblTimeplayed.Caption = "00:00:00.000"
    lblTimeLenght.Caption = "00:00:00.000"
    Label1.Caption = vbNullString
    blnStop = True
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.cmdStop_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmfForward_Click()
    On Error GoTo ErrorTrap
    PlayNext
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.cmfForward_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlJukeboxVolumn_Changed()
Dim SV As Single
    On Error GoTo ErrorTrap
    SV = ctlJukeboxVolumn.Value / 1024
    lblVolume.Caption = ctlJukeboxVolumn.Value
    If ChannelPlaying > 0 Then
        result = FMOD_Channel_SetVolume(ChannelPlaying, SV)
        ERRCHECK (result)
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmJukebox.ctlJukeboxVolumn_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlSysTrayIcon1_LeftButtonDown()
    On Error GoTo ErrorTrap
    ctlSysTrayIcon.IconVisible = False
    frmMP3Jukebox.Visible = True
    If blnPlayOpen Then
        frmPlayList.Visible = True
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.SysTrayIcon1_LeftButtonDown" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlSysTrayIcon1_RightButtonUp()
    On Error GoTo ErrorTrap
    Me.Visible = False
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.SysTrayIcon1_RightButtonUp" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlSysTrayIcon_LeftButtonDoubleClick()
    On Error GoTo ErrorTrap
    ctlSysTrayIcon.IconVisible = False
    Me.Visible = True
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmJukebox.ctlSysTrayIcon_LeftButtonDoubleClick" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
Dim rgnMain As Long
Dim nDirs   As Long
Dim nFiles  As Long
Dim sDir    As String
    On Error GoTo ErrorTrap
    nBytes = 624
    LoadBytes
    rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
    SetWindowRgn Me.hwnd, rgnMain, True
    picSpectrum.ScaleMode = 3
    picSpectrum.DrawMode = 13
    result = FMOD_System_Create(system)
    ERRCHECK (result)
    result = FMOD_System_GetVersion(system, Version)
    ERRCHECK (result)
    If Version <> FMOD_VERSION Then
        frmMsgBox.SMessageModal "Error!  You are using an old version of FMOD " & Hex$(Version) & ". " & _
 "This program requires " & Hex$(FMOD_VERSION)
    End If
    result = FMOD_System_Init(system, 1, FMOD_INIT_NORMAL, 0)
    ERRCHECK (result)
    blnAllStop = False
  
    ctlSysTrayIcon.IconVisible = False
    ctlSysTrayIcon.IconToolTipText = "Midi Jukebox"
    If Me.Width > snfSysInfo2.WorkAreaWidth Or Me.Height > snfSysInfo2.WorkAreaHeight Then
        Me.Width = snfSysInfo2.WorkAreaWidth
        Me.Height = snfSysInfo2.WorkAreaHeight
    End If
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
'aKeepOnTop frmrounded
    ctlJukeboxVolumn.Value = 800
    frmODirectory.Show vbModal, Me
    If blnCancel Then
        Unload Me
        End
    End If
    ReadIni
    frmLoading.Show
    DoEvents
    sDir = Zx_Tkey
    SongDir = Zx_Tkey
    Zx_Pattern = ".mp3"
    FindFile sDir, "*.*", nDirs, nFiles
    frmLoading.Hide
    DoEvents
    lsvListView1.Sorted = True
    DoEvents
    lsvListView1.Sorted = False
    If lsvListView1.ListItems.Count > 0 Then
        lsvListView1.ListItems(1).Selected = True
    End If
SubExit1:
    InitEq
    Me.Show
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)
    On Error GoTo ErrorTrap
    DragX = x
    DragY = y
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.Form_MouseDown" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)
    On Error GoTo ErrorTrap
    If Button = 1 Then
        frmMP3Jukebox.Move frmMP3Jukebox.Left + x - DragX, frmMP3Jukebox.Top + y - DragY
        If blnPlayOpen = True Then
            frmPlayList.Move frmPlayList.Left + x - DragX, frmPlayList.Top + y - DragY
        End If
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.Form_MouseMove" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorTrap
 
' Shut down
    If Sound Then
        result = FMOD_Sound_Release(Sound)
        ERRCHECK (result)
    End If
    If system Then
        result = FMOD_System_Close(system)
        ERRCHECK (result)
        result = FMOD_System_Release(system)
        ERRCHECK (result)
    End If
    Unload Me
    End
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.Form_Unload" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub InitEq()
Dim Mike As Single
    On Error GoTo ERRHANDLER
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq0)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq1)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq2)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq3)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq4)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq5)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq6)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq7)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq8)
    ERRCHECK (result)
    result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_PARAMEQ, Eq9)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq0)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq1)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq2)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq3)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq4)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq5)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq6)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq7)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq8)
    ERRCHECK (result)
    result = FMOD_System_AddDSP(system, Eq9)
    ERRCHECK (result)
    Mike = Val("80.0")
    result = FMOD_DSP_SetParameter(Eq0, FMOD_DSP_PARAMEQ_CENTER, Mike)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq0, FMOD_DSP_PARAMEQ_BANDWIDTH, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq0, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq1, FMOD_DSP_PARAMEQ_CENTER, 170)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq1, FMOD_DSP_PARAMEQ_BANDWIDTH, 2)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq1, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq2, FMOD_DSP_PARAMEQ_CENTER, 310)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq2, FMOD_DSP_PARAMEQ_BANDWIDTH, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq2, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq3, FMOD_DSP_PARAMEQ_CENTER, 600)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq3, FMOD_DSP_PARAMEQ_BANDWIDTH, 2)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq3, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq4, FMOD_DSP_PARAMEQ_CENTER, 1000)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq4, FMOD_DSP_PARAMEQ_BANDWIDTH, 2)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq4, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq5, FMOD_DSP_PARAMEQ_CENTER, 3000)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq5, FMOD_DSP_PARAMEQ_BANDWIDTH, 2)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq5, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq6, FMOD_DSP_PARAMEQ_CENTER, 6000)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq6, FMOD_DSP_PARAMEQ_BANDWIDTH, 2)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq6, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq7, FMOD_DSP_PARAMEQ_CENTER, 12000)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq7, FMOD_DSP_PARAMEQ_BANDWIDTH, 2)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq7, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq8, FMOD_DSP_PARAMEQ_CENTER, 14000)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq8, FMOD_DSP_PARAMEQ_BANDWIDTH, 2)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq8, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_CENTER, 16000)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_BANDWIDTH, 2)
    ERRCHECK (result)
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, 1)
    ERRCHECK (result)
    DoEvents
Exit Sub
ERRHANDLER:
    frmMsgBox.SMessageModal "Unexpected error in procedure: InitEq" & vbNewLine & _
          "Error #" & Err.Number & ": " & Err.description, vbCritical + vbOKOnly, App.Title
End Sub
Private Sub LoadBytes()
    bytRegion(0) = 32
    bytRegion(4) = 1
    bytRegion(8) = 37
    bytRegion(12) = 80
    bytRegion(13) = 2
    bytRegion(16) = 1
    bytRegion(20) = 1
    bytRegion(24) = 29
    bytRegion(25) = 1
    bytRegion(28) = 117
    bytRegion(29) = 1
    bytRegion(32) = 19
    bytRegion(36) = 1
    bytRegion(40) = 11
    bytRegion(41) = 1
    bytRegion(44) = 2
    bytRegion(48) = 17
    bytRegion(52) = 2
    bytRegion(56) = 13
    bytRegion(57) = 1
    bytRegion(60) = 3
    bytRegion(64) = 15
    bytRegion(68) = 3
    bytRegion(72) = 14
    bytRegion(73) = 1
    bytRegion(76) = 4
    bytRegion(80) = 14
    bytRegion(84) = 4
    bytRegion(88) = 16
    bytRegion(89) = 1
    bytRegion(92) = 5
    bytRegion(96) = 13
    bytRegion(100) = 5
    bytRegion(104) = 17
    bytRegion(105) = 1
    bytRegion(108) = 6
    bytRegion(112) = 11
    bytRegion(116) = 6
    bytRegion(120) = 18
    bytRegion(121) = 1
    bytRegion(124) = 7
    bytRegion(128) = 11
    bytRegion(132) = 7
    bytRegion(136) = 19
    bytRegion(137) = 1
    bytRegion(140) = 8
    bytRegion(144) = 10
    bytRegion(148) = 8
    bytRegion(152) = 20
    bytRegion(153) = 1
    bytRegion(156) = 9
    bytRegion(160) = 9
    bytRegion(164) = 9
    bytRegion(168) = 21
    bytRegion(169) = 1
    bytRegion(172) = 10
    bytRegion(176) = 8
    bytRegion(180) = 10
    bytRegion(184) = 22
    bytRegion(185) = 1
    bytRegion(188) = 11
    bytRegion(192) = 7
    bytRegion(196) = 11
    bytRegion(200) = 23
    bytRegion(201) = 1
    bytRegion(204) = 13
    bytRegion(208) = 6
    bytRegion(212) = 13
    bytRegion(216) = 24
    bytRegion(217) = 1
    bytRegion(220) = 14
    bytRegion(224) = 5
    bytRegion(228) = 14
    bytRegion(232) = 25
    bytRegion(233) = 1
    bytRegion(236) = 16
    bytRegion(240) = 4
    bytRegion(244) = 16
    bytRegion(248) = 26
    bytRegion(249) = 1
    bytRegion(252) = 18
    bytRegion(256) = 3
    bytRegion(260) = 18
    bytRegion(264) = 26
    bytRegion(265) = 1
    bytRegion(268) = 19
    bytRegion(272) = 3
    bytRegion(276) = 19
    bytRegion(280) = 27
    bytRegion(281) = 1
    bytRegion(284) = 21
    bytRegion(288) = 2
    bytRegion(292) = 21
    bytRegion(296) = 28
    bytRegion(297) = 1
    bytRegion(300) = 24
    bytRegion(304) = 1
    bytRegion(308) = 24
    bytRegion(312) = 28
    bytRegion(313) = 1
    bytRegion(316) = 25
    bytRegion(320) = 1
    bytRegion(324) = 25
    bytRegion(328) = 29
    bytRegion(329) = 1
    bytRegion(332) = 93
    bytRegion(333) = 1
    bytRegion(336) = 1
    bytRegion(340) = 93
    bytRegion(341) = 1
    bytRegion(344) = 28
    bytRegion(345) = 1
    bytRegion(348) = 94
    bytRegion(349) = 1
    bytRegion(352) = 2
    bytRegion(356) = 94
    bytRegion(357) = 1
    bytRegion(360) = 28
    bytRegion(361) = 1
    bytRegion(364) = 97
    bytRegion(365) = 1
    bytRegion(368) = 3
    bytRegion(372) = 97
    bytRegion(373) = 1
    bytRegion(376) = 27
    bytRegion(377) = 1
    bytRegion(380) = 99
    bytRegion(381) = 1
    bytRegion(384) = 3
    bytRegion(388) = 99
    bytRegion(389) = 1
    bytRegion(392) = 26
    bytRegion(393) = 1
    bytRegion(396) = 100
    bytRegion(397) = 1
    bytRegion(400) = 4
    bytRegion(404) = 100
    bytRegion(405) = 1
    bytRegion(408) = 26
    bytRegion(409) = 1
    bytRegion(412) = 102
    bytRegion(413) = 1
    bytRegion(416) = 5
    bytRegion(420) = 102
    bytRegion(421) = 1
    bytRegion(424) = 25
    bytRegion(425) = 1
    bytRegion(428) = 104
    bytRegion(429) = 1
    bytRegion(432) = 6
    bytRegion(436) = 104
    bytRegion(437) = 1
    bytRegion(440) = 24
    bytRegion(441) = 1
    bytRegion(444) = 105
    bytRegion(445) = 1
    bytRegion(448) = 7
    bytRegion(452) = 105
    bytRegion(453) = 1
    bytRegion(456) = 23
    bytRegion(457) = 1
    bytRegion(460) = 107
    bytRegion(461) = 1
    bytRegion(464) = 8
    bytRegion(468) = 107
    bytRegion(469) = 1
    bytRegion(472) = 22
    bytRegion(473) = 1
    bytRegion(476) = 108
    bytRegion(477) = 1
    bytRegion(480) = 9
    bytRegion(484) = 108
    bytRegion(485) = 1
    bytRegion(488) = 21
    bytRegion(489) = 1
    bytRegion(492) = 109
    bytRegion(493) = 1
    bytRegion(496) = 10
    bytRegion(500) = 109
    bytRegion(501) = 1
    bytRegion(504) = 20
    bytRegion(505) = 1
    bytRegion(508) = 110
    bytRegion(509) = 1
    bytRegion(512) = 11
    bytRegion(516) = 110
    bytRegion(517) = 1
    bytRegion(520) = 19
    bytRegion(521) = 1
    bytRegion(524) = 111
    bytRegion(525) = 1
    bytRegion(528) = 11
    bytRegion(532) = 111
    bytRegion(533) = 1
    bytRegion(536) = 18
    bytRegion(537) = 1
    bytRegion(540) = 112
    bytRegion(541) = 1
    bytRegion(544) = 13
    bytRegion(548) = 112
    bytRegion(549) = 1
    bytRegion(552) = 17
    bytRegion(553) = 1
    bytRegion(556) = 113
    bytRegion(557) = 1
    bytRegion(560) = 14
    bytRegion(564) = 113
    bytRegion(565) = 1
    bytRegion(568) = 16
    bytRegion(569) = 1
    bytRegion(572) = 114
    bytRegion(573) = 1
    bytRegion(576) = 15
    bytRegion(580) = 114
    bytRegion(581) = 1
    bytRegion(584) = 14
    bytRegion(585) = 1
    bytRegion(588) = 115
    bytRegion(589) = 1
    bytRegion(592) = 17
    bytRegion(596) = 115
    bytRegion(597) = 1
    bytRegion(600) = 13
    bytRegion(601) = 1
    bytRegion(604) = 116
    bytRegion(605) = 1
    bytRegion(608) = 19
    bytRegion(612) = 116
    bytRegion(613) = 1
    bytRegion(616) = 11
    bytRegion(617) = 1
    bytRegion(620) = 117
    bytRegion(621) = 1
End Sub
Private Sub lsvListView1_Click()
    On Error GoTo ErrorTrap
    If blnPlayOpen Then
        Set itmX = frmPlayList.lsvSlideListView.ListItems.Add(, , lsvListView1.SelectedItem.Text & "")
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.ListView1_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub lsvListView1_DblClick()
    On Error GoTo ErrorTrap
    PlaySong
    If blnPlayAll Then
        playX = lsvListView1.SelectedItem.index
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmJukebox.lsvListView1_DblClick" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Picture1_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmJukebox.cmdJukeboxClose_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Picture2_Click()
    On Error GoTo cmdMin_ClickError
    ctlSysTrayIcon.IconVisible = True
    Me.Hide
cmdMin_ClickExit:
Exit Sub
cmdMin_ClickError:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "cmdMin_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub PlayAll()
    On Error GoTo ErrorTrap
    lsvListView1.Refresh
    DoEvents
    blnPlayAll = True
    For playX = 1 To lsvListView1.ListItems.Count
        If blnStop Then
            blnStop = False
            GoTo PlayAllExit
        End If
        If blnPrevious Then
            playX = playX - lngPCount
            blnPrevious = False
            lngPCount = 0
        End If
        If blnNext Then
            playX = playX + lngNCount
            lngNCount = 0
            blnNext = False
        End If
        With lsvListView1
            .ListItems(playX).Selected = True
            .SelectedItem.EnsureVisible
            .Refresh
        End With
        DoEvents
        PlaySong
        Do While State = 0
            result = FMOD_Channel_IsPlaying(ChannelPlaying, State)
            ERRCHECK (result)
            DoEvents
        Loop
        Do While State = 1
            result = FMOD_Channel_IsPlaying(ChannelPlaying, State)
            ERRCHECK (result)
            DoEvents
        Loop
        If SoundPlaying > 0 Then
            result = FMOD_Sound_Release(SoundPlaying)
            ERRCHECK (result)
        End If
        Sound = 0
        ChannelPlaying = 0
        channel = 0
        If blnAllStop Then
            GoTo PlayAllExit
        End If
    Next playX
PlayAllExit:
    blnPlayAll = False
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.PlayAll" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub PlayNext()
Dim Xx      As Long
    On Error GoTo ErrorTrap
    blnNext = True
    lngNCount = lngNCount + 1
    Xx = lsvListView1.SelectedItem.index + 1
    If Xx <= lsvListView1.ListItems.Count Then
        lsvListView1.SelectedItem.Selected = False
        lsvListView1.ListItems(Xx).Selected = True
        lsvListView1.SelectedItem.EnsureVisible
        PlaySong
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.PlayNext" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub PlayPrevious()
Dim Xx      As Long
    On Error GoTo ErrorTrap
    Xx = lsvListView1.SelectedItem.index - 1
    If Xx > 0 Then
        blnPrevious = True
        lngPCount = lngPCount + 1
        With lsvListView1
            .SelectedItem.Selected = False
            .ListItems(Xx).Selected = True
            .SelectedItem.EnsureVisible
        End With
        PlaySong
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.PlayPrevious" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub PlaySong()
Dim strID    As String
Dim SV       As Single
Dim MSeconds As Long
Dim st       As Single
Dim tt       As Long
Dim sec      As Single
Dim min      As Long
Dim hour     As Long
    On Error GoTo ErrorTrap
    picSpectrum.Cls
    tmrJukeboxTimer2.Enabled = False
    strID = Zx_Tkey & lsvListView1.SelectedItem.Text
    If ChannelPlaying <> 0 Then
        result = FMOD_Channel_IsPlaying(ChannelPlaying, State)
        ERRCHECK (result)
        If State = 1 Then
            result = FMOD_Channel_Stop(ChannelPlaying)
            ERRCHECK (result)
        End If
        result = FMOD_Sound_Release(SoundPlaying)
        ERRCHECK (result)
        Sound = 0
    End If
    result = FMOD_System_CreateStream(system, strID, (FMOD_2D Or FMOD_software), Sound)
    ERRCHECK (result)
    result = FMOD_System_PlaySound(system, FMOD_CHANNEL_FREE, Sound, 0, channel)
    ERRCHECK (result)
    If chkLoop.Value = 1 Then
        result = FMOD_Channel_SetLoopCount(channel, -1)
        ERRCHECK (result)
        result = FMOD_Channel_GetPosition(channel, position, FMOD_TIMEUNIT_MS)
        result = FMOD_Channel_SetPosition(channel, position, FMOD_TIMEUNIT_MS)
    End If
    SV = ctlJukeboxVolumn.Value / 1024
    lblVolume.Caption = ctlJukeboxVolumn.Value
    result = FMOD_Channel_SetVolume(channel, SV)
    ERRCHECK (result)
    result = FMOD_Sound_GetLength(Sound, MSeconds, FMOD_TIMEUNIT_MS)
    If ((result <> FMOD_OK) And (result <> FMOD_ERR_INVALID_HANDLE) And (result <> FMOD_ERR_CHANNEL_STOLEN)) Then
        ERRCHECK (result)
    End If
    Do Until MSeconds > 0
    Loop
    sec = (MSeconds / 1000)
    st = sec
    Do Until sec < 60
        sec = sec - 60
    Loop
    tt = Int(st / 60)
    min = Int(tt Mod 60)
    hour = Int(min / 60)
    lblTimeLenght.Caption = Format$(hour, "00") & ":" & Format$(min, "00") & ":" & Format$(sec, "00.00")
    ChannelPlaying = channel
    SoundPlaying = Sound
    tmrJukeboxTimer1.Enabled = True
    tmrJukeboxTimer2.Enabled = True
    Label1.Caption = lsvListView1.SelectedItem.Text
    DoEvents
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmNewJukeBox.PlaySong" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ReadIni()
    On Error GoTo ErrorTrap
'read MCU Directory
    KeySection = "Directories"
    KeyKey = "MCUdirectory"
    loadINI
    iniMCUdirectory = KeyValue
'read Equalizer
    KeySection = "Equalizer"
    KeyKey = "Eq0Pos"
    loadINI
    iniEq0Pos = Val(KeyValue)
    KeyKey = "Eq1Pos"
    loadINI
    iniEq1Pos = Val(KeyValue)
    KeyKey = "Eq2Pos"
    loadINI
    iniEq2Pos = Val(KeyValue)
    KeyKey = "Eq3Pos"
    loadINI
    iniEq3Pos = Val(KeyValue)
    KeyKey = "Eq4Pos"
    loadINI
    iniEq4Pos = Val(KeyValue)
    KeyKey = "Eq5Pos"
    loadINI
    iniEq5Pos = Val(KeyValue)
    KeyKey = "Eq6Pos"
    loadINI
    iniEq6Pos = Val(KeyValue)
    KeyKey = "Eq7Pos"
    loadINI
    iniEq7Pos = Val(KeyValue)
    KeyKey = "Eq8Pos"
    loadINI
    iniEq8Pos = Val(KeyValue)
    KeyKey = "Eq9Pos"
    loadINI
    iniEq9Pos = Val(KeyValue)
' read Chorus
    KeySection = "Chorus"
    KeyKey = "DryMix"
    loadINI
    iniChorusDryMix = Val(KeyValue)
    KeyKey = "WetMix1"
    loadINI
    iniChorusWetMix1 = Val(KeyValue)
    KeyKey = "WetMix2"
    loadINI
    iniChorusWetMix2 = Val(KeyValue)
    KeyKey = "WetMix3"
    loadINI
    iniChorusWetMix3 = Val(KeyValue)
    KeyKey = "Delay"
    loadINI
    iniChorusDelay = Val(KeyValue)
    KeyKey = "Rate"
    loadINI
    iniChorusRate = Val(KeyValue)
    KeyKey = "Depth"
    loadINI
    iniChorusDepth = Val(KeyValue)
    KeyKey = "Feedback"
    loadINI
    iniChorusFeedback = Val(KeyValue)
' read Echo
    KeySection = "Echo"
    KeyKey = "Delay"
    loadINI
    iniEchoDelay = Val(KeyValue)
    KeyKey = "DecayRatio"
    loadINI
    iniEchoDecayRatio = Val(KeyValue)
    KeyKey = "MaxChannels"
    loadINI
    iniEchoMaxChannels = Val(KeyValue)
    KeyKey = "DryMix"
    loadINI
    iniEchoDryMix = Val(KeyValue)
    KeyKey = "WetMix"
    loadINI
    iniEchoWetMix = Val(KeyValue)
' read Distortion
    KeySection = "Distortion"
    KeyKey = "Level"
    loadINI
    iniDistortionLevel = Val(KeyValue)
' read Flange
    KeySection = "Flange"
    KeyKey = "DryMix"
    loadINI
    iniFlangeDryMix = Val(KeyValue)
    KeyKey = "WetMix"
    loadINI
    iniFlangeWetMix = Val(KeyValue)
    KeyKey = "Depth"
    loadINI
    iniFlangeDepth = Val(KeyValue)
    KeyKey = "Rate"
    loadINI
    iniFlangeRate = Val(KeyValue)
' read Highpass
    KeySection = "Highpass"
    KeyKey = "Cutoff"
    loadINI
    iniHighpassCutoff = Val(KeyValue)
    KeyKey = "Resonance"
    loadINI
    iniHighpassResonance = Val(KeyValue)
' read Lowpass
    KeySection = "Lowpass"
    KeyKey = "Cutoff"
    loadINI
    iniLowpassCutoff = Val(KeyValue)
    KeyKey = "Resonance"
    loadINI
    iniLowpassResonance = Val(KeyValue)
' read Normalise
    KeySection = "Normalise"
    KeyKey = "FadeTime"
    loadINI
    iniNormaliseFadeTime = Val(KeyValue)
    KeyKey = "Threshhold"
    loadINI
    iniNormaliseThreshhold = Val(KeyValue)
    KeyKey = "MaxAmp"
    loadINI
    iniNormaliseMaxAmp = Val(KeyValue)
' read Reverb
    KeySection = "Reverb"
    KeyKey = "RoomSize"
    loadINI
    iniReverbRoomSize = Val(KeyValue)
    KeyKey = "Damp"
    loadINI
    iniReverbDamp = Val(KeyValue)
    KeyKey = "WetMix"
    loadINI
    iniReverbWetMix = Val(KeyValue)
    KeyKey = "DryMix"
    loadINI
    iniReverbDryMix = Val(KeyValue)
    KeyKey = "Width"
    loadINI
    iniReverbWidth = Val(KeyValue)
    KeyKey = "Mode"
    loadINI
    iniReverbMode = KeyValue
' read Filter status
    KeySection = "Filters"
    KeyKey = "ReverbFilter"
    loadINI
    iniFilterReverb = KeyValue
    KeyKey = "ChorusFilter"
    loadINI
    iniFilterChorus = KeyValue
    KeyKey = "DistortionFilter"
    loadINI
    iniFilterDistortion = KeyValue
    KeyKey = "EchoFilter"
    loadINI
    iniFilterEcho = KeyValue
    KeyKey = "FlangeFilter"
    loadINI
    iniFilterFlange = KeyValue
    KeyKey = "HighpassFilter"
    loadINI
    iniFilterHighpass = KeyValue
    KeyKey = "LowpassFilter"
    loadINI
    iniFilterLowpass = KeyValue
    KeyKey = "NormalizeFilter"
    loadINI
    iniFilterNormalize = KeyValue
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmJukebox.ReadIni" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub tmrJukeboxTimer1_Timer()
Dim st   As Single
Dim T    As Long
Dim tt   As Long
Dim sec  As Single
Dim min  As Long
Dim hour As Long
    On Error GoTo ERRHANDLER
    result = FMOD_Channel_IsPlaying(ChannelPlaying, State)
    ERRCHECK (result)
    If State = 0 Then
        tmrJukeboxTimer1.Enabled = False
        result = FMOD_Sound_Release(SoundPlaying)
        ERRCHECK (result)
        lblTimeplayed.Caption = "00:00:00.000"
        lblTimeLenght.Caption = "00:00:00.000"
        Label1.Caption = vbNullString
        ChannelPlaying = 0
        SoundPlaying = 0
        Sound = 0
        If chkLoop = 1 Then
            PlaySong
        End If
        GoTo Timer1Exit
    End If
    result = FMOD_Channel_GetPosition(ChannelPlaying, T, FMOD_TIMEUNIT_MS)
    ERRCHECK (result)
    If Not T = -1 Then
        sec = T / 1000
        st = sec
        Do Until sec < 60
            sec = sec - 60
        Loop
        tt = Int(st / 60)
        min = Int(tt Mod 60)
        hour = Int(min / 60)
        lblTimeplayed.Caption = Format$(hour, "00") & ":" & Format$(min, "00") & ":" & Format$(sec, "00.000")
    End If
Timer1Exit:
    DoEvents
Exit Sub
ERRHANDLER:
    frmMsgBox.SMessageModal "Unexpected error in procedure: Timer1_Timer" & vbNewLine & _
          "Error #" & Err.Number & ": " & Err.description, vbCritical + vbOKOnly, App.Title
End Sub
Private Sub tmrJukeboxTimer2_Timer()
Dim X1            As Single
Dim X2            As Single
Dim Y1            As Single
Dim Y2            As Single
Dim SpecArr(512)  As Single
Dim SpecArr1(512) As Single
Dim Spectrum      As Single
Dim I             As Long
Dim result        As FMOD_RESULT
    On Error GoTo ErrorTrap
    result = FMOD_Channel_GetSpectrum(channel, SpecArr(0), 512, 0, FMOD_DSP_FFT_WINDOW_MAX)
    result = FMOD_Channel_GetSpectrum(channel, SpecArr1(0), 512, 1, FMOD_DSP_FFT_WINDOW_MAX)
    Do While I < 512
        Spectrum = SpecArr(I) + SpecArr1(I)
        X1 = picSpectrum.ScaleWidth * (I / picSpectrum.ScaleWidth)
        Y1 = Int(picSpectrum.ScaleHeight - picSpectrum.ScaleHeight * (Spectrum * 10))
        X2 = picSpectrum.ScaleWidth * (I / picSpectrum.ScaleWidth)
        Y2 = -1
        picSpectrum.Line (X1, Y1)-(X2, Y2), vbBlack
        X1 = picSpectrum.ScaleWidth * (I / picSpectrum.ScaleWidth)
        Y1 = Int(picSpectrum.ScaleHeight * (1 - Spectrum * 10))
        X2 = picSpectrum.ScaleWidth * (I / picSpectrum.ScaleWidth)
        Y2 = picSpectrum.ScaleHeight
        picSpectrum.Line (X1, Y1)-(X2, Y2), vbYellow
        I = I + 3
    Loop
    DoEvents
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmJukebox.tmrJukeboxTimer2_Timer" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:58 PM) 18 + 1301 = 1319 Lines Thanks Ulli for inspiration and lots of code.


