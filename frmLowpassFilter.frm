VERSION 5.00
Begin VB.Form frmLowpassSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1800
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5370
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
   Icon            =   "frmLowpassFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdLowpassCancel 
      Height          =   420
      Left            =   3810
      TabIndex        =   15
      Top             =   1087
      Width           =   1455
      _extentx        =   2566
      _extenty        =   741
      style           =   7
      caption         =   "Cancel"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmLowpassFilter.frx":000C
   End
   Begin MP3JukeBox.isButton cmdLowpassOK 
      Height          =   420
      Left            =   1837
      TabIndex        =   14
      Top             =   1087
      Width           =   1455
      _extentx        =   2566
      _extenty        =   741
      style           =   7
      caption         =   "OK"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmLowpassFilter.frx":0030
   End
   Begin MP3JukeBox.isButton cmdLowpassDefault 
      Height          =   555
      Left            =   345
      TabIndex        =   13
      Top             =   1020
      Width           =   975
      _extentx        =   1720
      _extenty        =   979
      style           =   7
      caption         =   "Default Settings"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmLowpassFilter.frx":0054
   End
   Begin MP3JukeBox.ctlEBSlider ctlLowpassCutoff 
      Height          =   135
      Left            =   705
      TabIndex        =   2
      Top             =   210
      Width           =   3885
      _extentx        =   6853
      _extenty        =   238
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlLowpassResonance 
      Height          =   135
      Left            =   1320
      TabIndex        =   1
      Top             =   570
      Width           =   3270
      _extentx        =   5768
      _extenty        =   238
      slidercolor     =   255
   End
   Begin VB.Label Label3 
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
      Left            =   2925
      TabIndex        =   12
      Top             =   1515
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Q value"
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
      Left            =   2663
      TabIndex        =   11
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label1 
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
      Left            =   2542
      TabIndex        =   10
      Top             =   330
      Width           =   210
   End
   Begin VB.Label lblLowpassResonance 
      BackColor       =   &H00000000&
      Caption         =   "Resonance"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   105
      TabIndex        =   9
      Top             =   525
      Width           =   960
   End
   Begin VB.Label lblLowpassLabel4 
      BackColor       =   &H00000000&
      Caption         =   "1.0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   705
      TabIndex        =   8
      Top             =   330
      Width           =   255
   End
   Begin VB.Label lblLowpassLabel5 
      BackColor       =   &H00000000&
      Caption         =   "22000"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   4065
      TabIndex        =   7
      Top             =   330
      Width           =   525
   End
   Begin VB.Label lblLowpassLabel8 
      BackColor       =   &H00000000&
      Caption         =   "10.0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   4230
      TabIndex        =   6
      Top             =   720
      Width           =   360
   End
   Begin VB.Label lblLowpassLabel7 
      BackColor       =   &H00000000&
      Caption         =   "1.0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblLowpassResonanceTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   225
      Left            =   4665
      TabIndex        =   4
      Top             =   525
      Width           =   615
   End
   Begin VB.Label lblLowpassCutoffTot 
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
      Left            =   4665
      TabIndex        =   3
      Top             =   165
      Width           =   615
   End
   Begin VB.Label lblLowpassCuttoff 
      BackColor       =   &H00000000&
      Caption         =   "Cutoff"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   510
   End
End
Attribute VB_Name = "frmLowpassSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdLowpassCancel_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmLowpassSetup.cmdLowpassCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdLowpassDefault_Click()
    On Error GoTo ErrorTrap
    ctlLowpassCutoff.Value = 5000
    ctlLowpassResonance.Value = 10
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmLowpassSetup.cmdLowpassDefault_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdLowpassOK_Click()
    On Error GoTo ErrorTrap
    KeySection = "Lowpass"
    KeyKey = "Cutoff"
    KeyValue = Format$(iniLowpassCutoff, "###0")
    saveINI
    KeyKey = "Resonance"
    KeyValue = Format$(iniLowpassResonance, "#0.0")
    saveINI
 
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmLowpassSetup.cmdLowpassOK_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlLowpassCutoff_Changed()
    On Error GoTo ErrorTrap
    iniLowpassCutoff = ctlLowpassCutoff.Value
    lblLowpassCutoffTot.Caption = Format$(iniLowpassCutoff, "###0")
    result = FMOD_DSP_SetParameter(DspLowpassFilter, FMOD_DSP_LOWPASS_CUTOFF, iniLowpassCutoff)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmLowpassSetup.ctlLowpassCutoff_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlLowpassResonance_Changed()
    On Error GoTo ErrorTrap
    iniLowpassResonance = ctlLowpassResonance.Value / 10
    lblLowpassResonanceTot.Caption = Format$(iniLowpassResonance, "#0.0")
    result = FMOD_DSP_SetParameter(DspLowpassFilter, FMOD_DSP_LOWPASS_RESONANCE, iniLowpassResonance)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmLowpassSetup.ctlLowpassResonance_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
 
    ctlLowpassCutoff.Value = iniLowpassCutoff
    ctlLowpassResonance.Value = iniLowpassResonance * 10
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmLowpassSetup.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:40 PM) 1 + 94 = 95 Lines Thanks Ulli for inspiration and lots of code.


