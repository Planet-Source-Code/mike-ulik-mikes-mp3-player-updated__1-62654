VERSION 5.00
Begin VB.Form frmNormalizeSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2355
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7335
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
   Icon            =   "frmNormaliseSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdNormaliseCancel 
      Height          =   435
      Left            =   4500
      TabIndex        =   20
      Top             =   1575
      Width           =   1275
      _extentx        =   2249
      _extenty        =   767
      style           =   7
      caption         =   "Cancel"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmNormaliseSetup.frx":000C
   End
   Begin MP3JukeBox.isButton cmdNormaliseOK 
      Height          =   435
      Left            =   2512
      TabIndex        =   19
      Top             =   1575
      Width           =   1275
      _extentx        =   2249
      _extenty        =   767
      style           =   7
      caption         =   "OK"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmNormaliseSetup.frx":0030
   End
   Begin MP3JukeBox.ctlEBSlider ctlNormaliseMaxAmp 
      Height          =   135
      Left            =   1140
      TabIndex        =   5
      Top             =   1050
      Width           =   5235
      _extentx        =   9234
      _extenty        =   238
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlNormaliseThreshhold 
      Height          =   135
      Left            =   1140
      TabIndex        =   4
      Top             =   622
      Width           =   5235
      _extentx        =   9234
      _extenty        =   238
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlNormaliseFadeTime 
      Height          =   135
      Left            =   1140
      TabIndex        =   3
      Top             =   195
      Width           =   5235
      _extentx        =   9234
      _extenty        =   238
      slidercolor     =   255
   End
   Begin MP3JukeBox.isButton cmdNormaliseDefault 
      Height          =   585
      Left            =   810
      TabIndex        =   21
      Top             =   1500
      Width           =   990
      _extentx        =   1746
      _extenty        =   1032
      style           =   7
      caption         =   "Default Settings"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmNormaliseSetup.frx":0054
   End
   Begin VB.Label Label4 
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
      Left            =   3495
      TabIndex        =   18
      Top             =   2055
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Maximum amplification"
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
      Left            =   2872
      TabIndex        =   17
      Top             =   1215
      Width           =   1770
   End
   Begin VB.Label Label2 
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
      Left            =   3465
      TabIndex        =   16
      Top             =   780
      Width           =   585
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
      Left            =   3262
      TabIndex        =   15
      Top             =   360
      Width           =   990
   End
   Begin VB.Label lblNormaliseLabel12 
      BackColor       =   &H00000000&
      Caption         =   "100000"
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
      Left            =   5850
      TabIndex        =   14
      Top             =   1215
      Width           =   525
   End
   Begin VB.Label lblNormaliseLabel11 
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
      Left            =   1140
      TabIndex        =   13
      Top             =   1215
      Width           =   75
   End
   Begin VB.Label lblNormaliseLabel10 
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
      Left            =   6165
      TabIndex        =   12
      Top             =   780
      Width           =   210
   End
   Begin VB.Label lblNormaliseLabel9 
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
      Left            =   1140
      TabIndex        =   11
      Top             =   780
      Width           =   240
   End
   Begin VB.Label lblNormaliseLabel8 
      BackColor       =   &H00000000&
      Caption         =   "20000"
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
      Left            =   5910
      TabIndex        =   10
      Top             =   360
      Width           =   465
   End
   Begin VB.Label lblNormaliseLabel7 
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
      Left            =   1140
      TabIndex        =   9
      Top             =   360
      Width           =   105
   End
   Begin VB.Label lblNormaliseMaxAmpTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   1005
      Width           =   810
   End
   Begin VB.Label lblNormaliseThreshholdTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.10"
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
      Left            =   6480
      TabIndex        =   7
      Top             =   577
      Width           =   810
   End
   Begin VB.Label lblNormaliseFadeTimeTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   225
      Left            =   6480
      TabIndex        =   6
      Top             =   150
      Width           =   810
   End
   Begin VB.Label lblNormaliseMaxAmp 
      BackColor       =   &H00000000&
      Caption         =   "Max Amp"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   2
      Top             =   1005
      Width           =   795
   End
   Begin VB.Label lblNormaliseThreshhold 
      BackColor       =   &H00000000&
      Caption         =   "Threshhold"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   1
      Top             =   577
      Width           =   960
   End
   Begin VB.Label lblNormaliseFadeTime 
      BackColor       =   &H00000000&
      Caption         =   "Fade Time"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   870
   End
End
Attribute VB_Name = "frmNormalizeSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdNormaliseCancel_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmNormalizeSetup.cmdNormaliseCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdNormaliseDefault_Click()
    On Error GoTo ErrorTrap
    ctlNormaliseFadeTime.Value = 5000
    ctlNormaliseMaxAmp.Value = 20
    ctlNormaliseThreshhold.Value = 1
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmNormalizeSetup.cmdNormaliseDefault_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdNormaliseOK_Click()
    On Error GoTo ErrorTrap
    KeySection = "Normalise"
    KeyKey = "FadeTime"
    KeyValue = Format$(iniNormaliseFadeTime, "####0")
    saveINI
    KeyKey = "MaxAmp"
    KeyValue = Format$(iniNormaliseMaxAmp, "#####0")
    saveINI
    KeyKey = "Threshhold"
    KeyValue = Format$(iniNormaliseThreshhold, "0.00")
    saveINI
 
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmNormalizeSetup.cmdNormaliseOK_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlNormaliseFadeTime_Changed()
    On Error GoTo ErrorTrap
    iniNormaliseFadeTime = ctlNormaliseFadeTime.Value
    lblNormaliseFadeTimeTot.Caption = Format$(iniNormaliseFadeTime, "####0")
    result = FMOD_DSP_SetParameter(DspNormalizeFilter, FMOD_DSP_NORMALIZE_FADETIME, iniNormaliseFadeTime)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmNormalizeSetup.ctlNormaliseFadeTime_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlNormaliseMaxAmp_Changed()
    On Error GoTo ErrorTrap
    iniNormaliseMaxAmp = ctlNormaliseMaxAmp.Value
    lblNormaliseMaxAmpTot.Caption = Format$(iniNormaliseMaxAmp, "#####0")
    result = FMOD_DSP_SetParameter(DspNormalizeFilter, FMOD_DSP_NORMALIZE_MAXAMP, iniNormaliseMaxAmp)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmNormalizeSetup.ctlNormaliseMaxAmp_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlNormaliseThreshhold_Changed()
    On Error GoTo ErrorTrap
    iniNormaliseThreshhold = ctlNormaliseThreshhold.Value / 100
    lblNormaliseThreshholdTot.Caption = Format$(iniNormaliseThreshhold, "0.00")
    result = FMOD_DSP_SetParameter(DspNormalizeFilter, FMOD_DSP_NORMALIZE_THRESHHOLD, iniNormaliseThreshhold)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmNormalizeSetup.ctlNormaliseThreshhold_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
 
    ctlNormaliseFadeTime.Value = iniNormaliseFadeTime
    ctlNormaliseMaxAmp.Value = iniNormaliseMaxAmp
    ctlNormaliseThreshhold.Value = iniNormaliseThreshhold * 100
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmNormalizeSetup.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:42 PM) 1 + 114 = 115 Lines Thanks Ulli for inspiration and lots of code.


