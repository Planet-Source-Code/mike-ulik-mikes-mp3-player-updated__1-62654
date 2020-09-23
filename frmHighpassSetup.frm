VERSION 5.00
Begin VB.Form frmHighpassSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1740
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5025
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
   Icon            =   "frmHighpassSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdHighpassCancel 
      Height          =   525
      Left            =   3615
      TabIndex        =   15
      Top             =   870
      Width           =   1035
      _extentx        =   1826
      _extenty        =   926
      style           =   7
      caption         =   "Cancel"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmHighpassSetup.frx":000C
   End
   Begin MP3JukeBox.isButton cmdHighpassOK 
      Height          =   525
      Left            =   1822
      TabIndex        =   14
      Top             =   870
      Width           =   1035
      _extentx        =   1826
      _extenty        =   926
      style           =   7
      caption         =   "OK"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmHighpassSetup.frx":0030
   End
   Begin MP3JukeBox.isButton HighpassDefault 
      Height          =   555
      Left            =   240
      TabIndex        =   13
      Top             =   855
      Width           =   825
      _extentx        =   1455
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
      font            =   "frmHighpassSetup.frx":0054
   End
   Begin MP3JukeBox.ctlEBSlider ctlHighpassCutoff 
      Height          =   120
      Left            =   720
      TabIndex        =   5
      Top             =   127
      Width           =   3390
      _extentx        =   5980
      _extenty        =   212
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlHighpassResonance 
      Height          =   120
      Left            =   1140
      TabIndex        =   9
      Top             =   517
      Width           =   2970
      _extentx        =   5239
      _extenty        =   212
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
      Left            =   2025
      TabIndex        =   12
      Top             =   1425
      Width           =   1305
   End
   Begin VB.Label lblHighpassResonance 
      BackColor       =   &H00000000&
      Caption         =   "Resonance"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   465
      Width           =   960
   End
   Begin VB.Label lblHighpassResonanceTot 
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
      Left            =   4230
      TabIndex        =   10
      Top             =   465
      Width           =   645
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
      Left            =   2333
      TabIndex        =   8
      Top             =   630
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
      Left            =   2310
      TabIndex        =   7
      Top             =   270
      Width           =   210
   End
   Begin VB.Label lblHighpassSetupLabel6 
      BackColor       =   &H00000000&
      Caption         =   "22000"
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
      Left            =   3645
      TabIndex        =   6
      Top             =   270
      Width           =   465
   End
   Begin VB.Label lblHighpassSetupLabel8 
      BackColor       =   &H00000000&
      Caption         =   "10.0"
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
      Left            =   3810
      TabIndex        =   4
      Top             =   630
      Width           =   300
   End
   Begin VB.Label lblHighpassSetupLabel7 
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
      Left            =   1140
      TabIndex        =   3
      Top             =   630
      Width           =   210
   End
   Begin VB.Label lblHighpassSetupLabel5 
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
      Left            =   720
      TabIndex        =   2
      Top             =   270
      Width           =   165
   End
   Begin VB.Label lblHighpassCutoffTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 5000"
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
      Left            =   4230
      TabIndex        =   1
      Top             =   75
      Width           =   645
   End
   Begin VB.Label lblHighpassCutoff 
      BackColor       =   &H00000000&
      Caption         =   "Cutoff"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   510
   End
End
Attribute VB_Name = "frmHighpassSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdHighpassCancel_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmHighpassSetup.cmdHighpassCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdHighpassOK_Click()
    On Error GoTo ErrorTrap
    KeySection = "Highpass"
    KeyKey = "Cutoff"
    KeyValue = Format$(iniHighpassCutoff, "###0")
    saveINI
    KeyKey = "Resonance"
    KeyValue = Format$(iniHighpassResonance, "0.0")
    saveINI
 
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmHighpassSetup.cmdHighpassOK_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlHighpassCutoff_Changed()
    On Error GoTo ErrorTrap
    iniHighpassCutoff = ctlHighpassCutoff.Value
    lblHighpassCutoffTot.Caption = Format$(iniHighpassCutoff, "###0")
    result = FMOD_DSP_SetParameter(DspHighpassFilter, FMOD_DSP_HIGHPASS_CUTOFF, iniHighpassCutoff)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmHighpassSetup.ctlHighpassCutoff_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlHighpassResonance_Changed()
    On Error GoTo ErrorTrap
    iniHighpassResonance = ctlHighpassResonance.Value / 10
    lblHighpassResonanceTot.Caption = Format$(iniHighpassResonance, "#0.0")
    result = FMOD_DSP_SetParameter(DspHighpassFilter, FMOD_DSP_HIGHPASS_RESONANCE, iniHighpassResonance)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmHighpassSetup.ctlHighpassResonance_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
 
    ctlHighpassCutoff.Value = iniHighpassCutoff
    ctlHighpassResonance.Value = iniHighpassResonance * 10
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmHighpassSetup.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub HighpassDefault_Click()
    On Error GoTo ErrorTrap
    ctlHighpassCutoff.Value = 5000
    ctlHighpassResonance.Value = 1
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmHighpassSetup.HighpassDefault_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:38 PM) 1 + 94 = 95 Lines Thanks Ulli for inspiration and lots of code.


