VERSION 5.00
Begin VB.Form frmChorusSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4140
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChorusSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdChorusCancel 
      Height          =   450
      Left            =   2940
      TabIndex        =   49
      Top             =   3435
      Width           =   975
      _extentx        =   1720
      _extenty        =   794
      style           =   7
      caption         =   "Cancel"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmChorusSetup.frx":000C
   End
   Begin MP3JukeBox.isButton cmdChorusOK 
      Height          =   450
      Left            =   1590
      TabIndex        =   48
      Top             =   3435
      Width           =   975
      _extentx        =   1720
      _extenty        =   794
      style           =   7
      caption         =   "OK"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmChorusSetup.frx":0030
   End
   Begin MP3JukeBox.isButton ChorusDefaultSettings 
      Height          =   615
      Left            =   210
      TabIndex        =   47
      Top             =   3353
      Width           =   1020
      _extentx        =   1799
      _extenty        =   1085
      style           =   7
      caption         =   "Default Settings"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmChorusSetup.frx":0054
   End
   Begin MP3JukeBox.ctlEBSlider ctlChorusDepth 
      Height          =   150
      Left            =   1050
      TabIndex        =   6
      Top             =   2494
      Width           =   2010
      _extentx        =   3545
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlChorusRate 
      Height          =   150
      Left            =   1050
      TabIndex        =   5
      Top             =   2102
      Width           =   2010
      _extentx        =   3545
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlChorusDelay 
      Height          =   150
      Left            =   1050
      TabIndex        =   4
      Top             =   1710
      Width           =   2010
      _extentx        =   3545
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlChorusWetMix3 
      Height          =   150
      Left            =   1050
      TabIndex        =   3
      Top             =   1318
      Width           =   2010
      _extentx        =   3545
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlChorusWetMix2 
      Height          =   150
      Left            =   1050
      TabIndex        =   2
      Top             =   926
      Width           =   2010
      _extentx        =   3545
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlChorusWetMix1 
      Height          =   150
      Left            =   1050
      TabIndex        =   1
      Top             =   534
      Width           =   2010
      _extentx        =   3545
      _extenty        =   265
      slidercolor     =   -2147483635
   End
   Begin MP3JukeBox.ctlEBSlider ctlChorusDryMix 
      Height          =   150
      Left            =   1050
      TabIndex        =   0
      Top             =   142
      Width           =   2010
      _extentx        =   3545
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlChorusFeedback 
      Height          =   150
      Left            =   1050
      TabIndex        =   33
      Top             =   2887
      Width           =   2010
      _extentx        =   3545
      _extenty        =   265
      slidercolor     =   255
   End
   Begin VB.Label lblChorusLabel7 
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
      Left            =   1905
      TabIndex        =   46
      Top             =   3900
      Width           =   1305
   End
   Begin VB.Label lblChorusLabel27 
      BackColor       =   &H00000000&
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1050
      TabIndex        =   45
      Top             =   2265
      Width           =   240
   End
   Begin VB.Label lblChorusLabel28 
      BackColor       =   &H00000000&
      Caption         =   "20.0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2730
      TabIndex        =   44
      Top             =   2265
      Width           =   330
   End
   Begin VB.Label lblChorusLabel6 
      BackColor       =   &H00000000&
      Caption         =   "Hz"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1950
      TabIndex        =   43
      Top             =   2265
      Width           =   210
   End
   Begin VB.Label lblChorusFeedback 
      BackColor       =   &H00000000&
      Caption         =   "Feedback"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   42
      Top             =   2850
      Width           =   825
   End
   Begin VB.Label lblChorusRate 
      BackColor       =   &H00000000&
      Caption         =   "Rate"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   41
      Top             =   2065
      Width           =   390
   End
   Begin VB.Label lblChorusDepth 
      BackColor       =   &H00000000&
      Caption         =   "Depth"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   40
      Top             =   2457
      Width           =   495
   End
   Begin VB.Label lblChorusDelay 
      BackColor       =   &H00000000&
      Caption         =   "Delay"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   39
      Top             =   1673
      Width           =   465
   End
   Begin VB.Label lblChorusWetMix3 
      BackColor       =   &H00000000&
      Caption         =   "Wet Mix 3"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   38
      Top             =   1281
      Width           =   840
   End
   Begin VB.Label lblChorusWetMix2 
      BackColor       =   &H00000000&
      Caption         =   "Wet Mix 2"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   37
      Top             =   889
      Width           =   840
   End
   Begin VB.Label lblChorusWetMix1 
      BackColor       =   &H00000000&
      Caption         =   "Wet Mix 1"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   36
      Top             =   497
      Width           =   840
   End
   Begin VB.Label lblChorusDryMix 
      BackColor       =   &H00000000&
      Caption         =   "Dry Mix"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   35
      Top             =   105
      Width           =   630
   End
   Begin VB.Label lblChorusFeedbackTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   34
      Top             =   2827
      Width           =   705
   End
   Begin VB.Label lblChorusLabel5 
      BackColor       =   &H00000000&
      Caption         =   "Milliseconds"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1560
      TabIndex        =   32
      Top             =   1875
      Width           =   990
   End
   Begin VB.Label lblChorusLabel4 
      BackColor       =   &H00000000&
      Caption         =   "volume"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1748
      TabIndex        =   31
      Top             =   1470
      Width           =   615
   End
   Begin VB.Label lblChorusLabel3 
      BackColor       =   &H00000000&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1763
      TabIndex        =   30
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label lblChorusLabel2 
      BackColor       =   &H00000000&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1763
      TabIndex        =   29
      Top             =   705
      Width           =   585
   End
   Begin VB.Label lblChorusLabel1 
      BackColor       =   &H00000000&
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1763
      TabIndex        =   28
      Top             =   300
      Width           =   585
   End
   Begin VB.Label lblChorusLabel32 
      BackColor       =   &H00000000&
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2850
      TabIndex        =   27
      Top             =   3060
      Width           =   210
   End
   Begin VB.Label lblChorusLabel31 
      BackColor       =   &H00000000&
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1050
      TabIndex        =   26
      Top             =   3060
      Width           =   240
   End
   Begin VB.Label lblChorusLabel30 
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2985
      TabIndex        =   25
      Top             =   2640
      Width           =   75
   End
   Begin VB.Label lblChorusLabel29 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1050
      TabIndex        =   24
      Top             =   2640
      Width           =   105
   End
   Begin VB.Label lblChorusLabel26 
      BackColor       =   &H00000000&
      Caption         =   "100.0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2670
      TabIndex        =   23
      Top             =   1875
      Width           =   390
   End
   Begin VB.Label lblChorusLabel25 
      BackColor       =   &H00000000&
      Caption         =   ".1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1050
      TabIndex        =   22
      Top             =   1905
      Width           =   120
   End
   Begin VB.Label lblChorusLabel24 
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2985
      TabIndex        =   21
      Top             =   1470
      Width           =   75
   End
   Begin VB.Label lblChorusLabel23 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1050
      TabIndex        =   20
      Top             =   1470
      Width           =   105
   End
   Begin VB.Label lblChorusLabel22 
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2955
      TabIndex        =   19
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label lblChorusLabel21 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1050
      TabIndex        =   18
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label lblChorusLabel20 
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2985
      TabIndex        =   17
      Top             =   705
      Width           =   75
   End
   Begin VB.Label lblChorusLabel19 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1050
      TabIndex        =   16
      Top             =   705
      Width           =   105
   End
   Begin VB.Label lblChorusLabel18 
      BackColor       =   &H00000000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   2985
      TabIndex        =   15
      Top             =   315
      Width           =   75
   End
   Begin VB.Label lblChorusLabel17 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   1050
      TabIndex        =   14
      Top             =   315
      Width           =   105
   End
   Begin VB.Label lblChorusDepthTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.03"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   13
      Top             =   2434
      Width           =   705
   End
   Begin VB.Label lblChorusRateTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   12
      Top             =   2042
      Width           =   705
   End
   Begin VB.Label lblChorusDelayTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "40.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   11
      Top             =   1650
      Width           =   705
   End
   Begin VB.Label lblChorusWetMix3Tot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   10
      Top             =   1258
      Width           =   705
   End
   Begin VB.Label lblChorusWetMix2Tot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   9
      Top             =   866
      Width           =   705
   End
   Begin VB.Label lblChorusWetMix1Tot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   8
      Top             =   474
      Width           =   705
   End
   Begin VB.Label lblChorusDryMixTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.50"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3285
      TabIndex        =   7
      Top             =   82
      Width           =   705
   End
End
Attribute VB_Name = "frmChorusSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ChorusDefaultSettings_Click()
    On Error GoTo ErrorTrap
    ctlChorusDryMix.Value = 50
    ctlChorusDelay.Value = 400
    ctlChorusDepth.Value = 3
    ctlChorusFeedback.Value = 0
    ctlChorusRate.Value = 50
    ctlChorusWetMix1.Value = 50
    ctlChorusWetMix2.Value = 50
    ctlChorusWetMix3.Value = 50
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ChorusDefaultSettings_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdChorusCancel_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.cmdChorusCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdChorusOK_Click()
    On Error GoTo ErrorTrap
    KeySection = "Chorus"
    KeyKey = "DryMix"
    KeyValue = Format$(iniChorusDryMix, "0.00")
    saveINI
    KeyKey = "WetMix1"
    KeyValue = Format$(iniChorusWetMix1, "0.00")
    saveINI
    KeyKey = "WetMix2"
    KeyValue = Format$(iniChorusWetMix2, "0.00")
    saveINI
    KeyKey = "WetMix3"
    KeyValue = Format$(iniChorusWetMix3, "0.00")
    saveINI
    KeyKey = "Delay"
    KeyValue = Format$(iniChorusDelay, "0.00")
    saveINI
    KeyKey = "Rate"
    KeyValue = Format$(iniChorusRate, "0.00")
    saveINI
    KeyKey = "Depth"
    KeyValue = Format$(iniChorusDepth, "0.00")
    saveINI
    KeyKey = "Feedback"
    KeyValue = Format$(iniChorusFeedback, "0.00")
    saveINI
 
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.cmdChorusOK_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlChorusDelay_Changed()
    On Error GoTo ErrorTrap
    iniChorusDelay = ctlChorusDelay.Value / 10
    lblChorusDelayTot.Caption = Format$(iniChorusDelay, "##0.0")
    result = FMOD_DSP_SetParameter(DspChorusFilter, FMOD_DSP_CHORUS_DELAY, iniChorusDelay)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ctlChorusDelay_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlChorusDepth_Changed()
    On Error GoTo ErrorTrap
    iniChorusDepth = ctlChorusDepth.Value / 100
    lblChorusDepthTot.Caption = Format$(iniChorusDepth, "0.00")
    result = FMOD_DSP_SetParameter(DspChorusFilter, FMOD_DSP_CHORUS_DEPTH, iniChorusDepth)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ctlChorusDepth_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlChorusDryMix_Changed()
    On Error GoTo ErrorTrap
    iniChorusDryMix = ctlChorusDryMix.Value / 100
    lblChorusDryMixTot.Caption = Format$(iniChorusDryMix, "0.00")
    result = FMOD_DSP_SetParameter(DspChorusFilter, FMOD_DSP_CHORUS_DRYMIX, iniChorusDryMix)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ctlChorusDryMix_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlChorusFeedback_Changed()
    On Error GoTo ErrorTrap
    iniChorusFeedback = ctlChorusFeedback.Value / 100
    lblChorusFeedbackTot.Caption = Format$(iniChorusFeedback, "0.00")
    result = FMOD_DSP_SetParameter(DspChorusFilter, FMOD_DSP_CHORUS_FEEDBACK, iniChorusFeedback)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ctlChorusFeedback_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlChorusRate_Changed()
    On Error GoTo ErrorTrap
    iniChorusRate = ctlChorusRate.Value / 10
    lblChorusRateTot.Caption = Format$(iniChorusRate, "0.0")
    result = FMOD_DSP_SetParameter(DspChorusFilter, FMOD_DSP_CHORUS_RATE, iniChorusRate)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ctlChorusRate_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlChorusWetMix1_Changed()
    On Error GoTo ErrorTrap
    iniChorusWetMix1 = ctlChorusWetMix1.Value / 100
    lblChorusWetMix1Tot.Caption = Format$(iniChorusWetMix1, "0.00")
    result = FMOD_DSP_SetParameter(DspChorusFilter, FMOD_DSP_CHORUS_WETMIX1, iniChorusWetMix1)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ctlChorusWetMix1_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlChorusWetMix2_Changed()
    On Error GoTo ErrorTrap
    iniChorusWetMix2 = ctlChorusWetMix2.Value / 100
    lblChorusWetMix2Tot.Caption = Format$(iniChorusWetMix2, "0.00")
    result = FMOD_DSP_SetParameter(DspChorusFilter, FMOD_DSP_CHORUS_WETMIX2, iniChorusWetMix2)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ctlChorusWetMix2_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlChorusWetmix3_Changed()
    On Error GoTo ErrorTrap
    iniChorusWetMix3 = ctlChorusWetMix3.Value / 100
    lblChorusWetMix3Tot.Caption = Format$(iniChorusWetMix3, "0.00")
    result = FMOD_DSP_SetParameter(DspChorusFilter, FMOD_DSP_CHORUS_WETMIX3, iniChorusWetMix3)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.ctlChorusWetmix3_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
 
    ctlChorusDryMix.Value = iniChorusDryMix * 100
    ctlChorusWetMix1.Value = iniChorusWetMix1 * 100
    ctlChorusWetMix2.Value = iniChorusWetMix2 * 100
    ctlChorusWetMix3.Value = iniChorusWetMix3 * 100
    ctlChorusDelay.Value = iniChorusDelay * 10
    ctlChorusRate.Value = iniChorusRate * 10
    ctlChorusDepth.Value = iniChorusDepth * 100
    ctlChorusFeedback.Value = iniChorusFeedback * 100
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmChorusSetup.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:28 PM) 1 + 214 = 215 Lines Thanks Ulli for inspiration and lots of code.


