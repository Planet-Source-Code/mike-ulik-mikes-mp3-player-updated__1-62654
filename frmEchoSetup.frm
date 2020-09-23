VERSION 5.00
Begin VB.Form frmEchoSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4125
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
   Icon            =   "frmEchoSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdEchoCancel 
      Height          =   375
      Left            =   2685
      TabIndex        =   32
      Top             =   2408
      Width           =   1065
      _extentx        =   1879
      _extenty        =   661
      style           =   7
      caption         =   "Cancel"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEchoSetup.frx":000C
   End
   Begin MP3JukeBox.isButton cmdEchoOK 
      Height          =   375
      Left            =   1365
      TabIndex        =   31
      Top             =   2408
      Width           =   1065
      _extentx        =   1879
      _extenty        =   661
      style           =   7
      caption         =   "OK"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEchoSetup.frx":0030
   End
   Begin MP3JukeBox.isButton cmdEchoDefault 
      Height          =   660
      Left            =   210
      TabIndex        =   30
      Top             =   2265
      Width           =   900
      _extentx        =   1588
      _extenty        =   1164
      style           =   7
      caption         =   "Default Settings"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEchoSetup.frx":0054
   End
   Begin MP3JukeBox.ctlEBSlider ctlEchoWetMix 
      Height          =   135
      Left            =   1410
      TabIndex        =   19
      Top             =   1860
      Width           =   2010
      _extentx        =   3545
      _extenty        =   238
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlEchoDryMix 
      Height          =   135
      Left            =   1410
      TabIndex        =   20
      Top             =   1443
      Width           =   2010
      _extentx        =   3545
      _extenty        =   238
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlEchoMaxChannels 
      Height          =   135
      Left            =   1410
      TabIndex        =   21
      Top             =   1027
      Width           =   2010
      _extentx        =   3545
      _extenty        =   238
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlEchoDecayRatio 
      Height          =   135
      Left            =   1410
      TabIndex        =   22
      Top             =   611
      Width           =   2010
      _extentx        =   3545
      _extenty        =   238
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlEchoDelay 
      Height          =   135
      Left            =   1410
      TabIndex        =   23
      Top             =   195
      Width           =   2010
      _extentx        =   3545
      _extenty        =   238
      slidercolor     =   255
   End
   Begin VB.Label Label5 
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
      Left            =   1815
      TabIndex        =   29
      Top             =   2850
      Width           =   1305
   End
   Begin VB.Label lblEchoLabel7 
      BackColor       =   &H00000000&
      Caption         =   "10"
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
      Left            =   1410
      TabIndex        =   28
      Top             =   345
      Width           =   165
   End
   Begin VB.Label lblEchoLabel9 
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
      Left            =   1410
      TabIndex        =   27
      Top             =   750
      Width           =   105
   End
   Begin VB.Label lblEchoLabel12 
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
      Left            =   1410
      TabIndex        =   26
      Top             =   1185
      Width           =   105
   End
   Begin VB.Label lblEchoLabel14 
      BackColor       =   &H00000000&
      Caption         =   "0.00"
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
      Left            =   1410
      TabIndex        =   25
      Top             =   1590
      Width           =   330
   End
   Begin VB.Label lblEchoLabel19 
      BackColor       =   &H00000000&
      Caption         =   "0.00"
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
      Left            =   1410
      TabIndex        =   24
      Top             =   1995
      Width           =   330
   End
   Begin VB.Label Label4 
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
      Left            =   2123
      TabIndex        =   18
      Top             =   1995
      Width           =   585
   End
   Begin VB.Label Label3 
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
      Left            =   2123
      TabIndex        =   17
      Top             =   1590
      Width           =   585
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Decay per delay"
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
      Left            =   1815
      TabIndex        =   16
      Top             =   750
      Width           =   1200
   End
   Begin VB.Label Label1 
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
      Left            =   1920
      TabIndex        =   15
      Top             =   345
      Width           =   990
   End
   Begin VB.Label lblEchoLabel20 
      BackColor       =   &H00000000&
      Caption         =   "1.00"
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
      Left            =   3120
      TabIndex        =   14
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label lblEchoWetMixTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3495
      TabIndex        =   13
      Top             =   1815
      Width           =   510
   End
   Begin VB.Label lblEchoDryMixTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3495
      TabIndex        =   12
      Top             =   1398
      Width           =   510
   End
   Begin VB.Label lblEchoMaxChannelsTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   225
      Left            =   3495
      TabIndex        =   11
      Top             =   982
      Width           =   510
   End
   Begin VB.Label lblEchoLabel15 
      BackColor       =   &H00000000&
      Caption         =   "1.00"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   1590
      Width           =   300
   End
   Begin VB.Label lblEchoLabel13 
      BackColor       =   &H00000000&
      Caption         =   "16"
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
      Left            =   3255
      TabIndex        =   9
      Top             =   1185
      Width           =   165
   End
   Begin VB.Label lblEchoDecayRatioTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.50"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3495
      TabIndex        =   8
      Top             =   566
      Width           =   510
   End
   Begin VB.Label lblEchoLabel10 
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
      Left            =   3345
      TabIndex        =   7
      Top             =   750
      Width           =   75
   End
   Begin VB.Label lblEchoLabel8 
      BackColor       =   &H00000000&
      Caption         =   "5000"
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
      Left            =   3045
      TabIndex        =   6
      Top             =   345
      Width           =   375
   End
   Begin VB.Label lblEchoDelayTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3495
      TabIndex        =   5
      Top             =   150
      Width           =   510
   End
   Begin VB.Label lblEchoWetMix 
      BackColor       =   &H00000000&
      Caption         =   "Wet Mix"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1815
      Width           =   690
   End
   Begin VB.Label lblEchoDryMix 
      BackColor       =   &H00000000&
      Caption         =   "Dry Mix"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   1398
      Width           =   630
   End
   Begin VB.Label lblEchoMaxChannels 
      BackColor       =   &H00000000&
      Caption         =   "Max Channels"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   982
      Width           =   1200
   End
   Begin VB.Label lblEchoDecayRatio 
      BackColor       =   &H00000000&
      Caption         =   "Decay Ratio"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   566
      Width           =   1005
   End
   Begin VB.Label lblEchoDelay 
      BackColor       =   &H00000000&
      Caption         =   "Delay"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   465
   End
End
Attribute VB_Name = "frmEchoSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEchoCancel_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.cmdEchoCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEchoDefault_Click()
    On Error GoTo ErrorTrap
    ctlEchoDecayRatio.Value = 50
    ctlEchoDelay.Value = 500
    ctlEchoDryMix.Value = 100
    ctlEchoMaxChannels.Value = 0
    ctlEchoWetMix.Value = 100
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.cmdEchoDefault_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEchoOK_Click()
    On Error GoTo ErrorTrap
    KeySection = "Echo"
    KeyKey = "Delay"
    KeyValue = Format$(iniEchoDelay, "###0")
    saveINI
    KeyKey = "DecayRatio"
    KeyValue = Format$(iniEchoDecayRatio, "0.00")
    saveINI
    KeyKey = "MaxChannels"
    KeyValue = Format$(iniEchoMaxChannels, "#0")
    saveINI
    KeyKey = "DryMix"
    KeyValue = Format$(iniEchoDryMix, "0.00")
    saveINI
    KeyKey = "WetMix"
    KeyValue = Format$(iniEchoWetMix, "0.00")
    saveINI
 
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.cmdEchoOK_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlEchoDecayRatio_Changed()
    On Error GoTo ErrorTrap
    iniEchoDecayRatio = ctlEchoDecayRatio.Value / 100
    lblEchoDecayRatioTot.Caption = Format$(iniEchoDecayRatio, "0.00")
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.ctlEchoDecayRatio_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlEchoDelay_Changed()
    On Error GoTo ErrorTrap
    iniEchoDelay = ctlEchoDelay.Value
    lblEchoDelayTot.Caption = Format$(iniEchoDelay, "###0")
    result = FMOD_DSP_SetParameter(DspEchoFilter, FMOD_DSP_ECHO_DELAY, iniEchoDelay)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.ctlEchoDelay_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlEchoDryMix_Changed()
    On Error GoTo ErrorTrap
    iniEchoDryMix = ctlEchoDryMix.Value / 100
    lblEchoDryMixTot.Caption = Format$(iniEchoDryMix, "0.00")
    result = FMOD_DSP_SetParameter(DspEchoFilter, FMOD_DSP_ECHO_DRYMIX, iniEchoDryMix)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.ctlEchoDryMix_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlEchoMaxChannels_Changed()
    On Error GoTo ErrorTrap
    iniEchoMaxChannels = ctlEchoMaxChannels.Value
    lblEchoMaxChannelsTot.Caption = Format$(iniEchoMaxChannels, "#0")
    result = FMOD_DSP_SetParameter(DspEchoFilter, FMOD_DSP_ECHO_MAXCHANNELS, iniEchoMaxChannels)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.ctlEchoMaxChannels_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlEchoWetMix_Changed()
    On Error GoTo ErrorTrap
    iniEchoWetMix = ctlEchoWetMix.Value / 100
    lblEchoWetMixTot.Caption = Format$(iniEchoWetMix, "0.00")
    result = FMOD_DSP_SetParameter(DspEchoFilter, FMOD_DSP_ECHO_WETMIX, iniEchoWetMix)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.ctlEchoWetMix_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
 
    ctlEchoDecayRatio.Value = iniEchoDecayRatio * 100
    ctlEchoDelay.Value = iniEchoDelay
    ctlEchoDryMix.Value = iniEchoDryMix * 100
    ctlEchoMaxChannels.Value = iniEchoMaxChannels
    ctlEchoWetMix.Value = iniEchoWetMix * 100
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEchoSetup.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:33 PM) 1 + 152 = 153 Lines Thanks Ulli for inspiration and lots of code.


