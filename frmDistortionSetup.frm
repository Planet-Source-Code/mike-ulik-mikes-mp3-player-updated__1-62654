VERSION 5.00
Begin VB.Form frmDistortionSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1305
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3345
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
   Icon            =   "frmDistortionSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdDistortionCancel 
      Height          =   345
      Left            =   2340
      TabIndex        =   8
      Top             =   645
      Width           =   900
      _extentx        =   1588
      _extenty        =   609
      style           =   7
      caption         =   "Cancel"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmDistortionSetup.frx":000C
   End
   Begin MP3JukeBox.isButton cmdDistortionOK 
      Height          =   345
      Left            =   1260
      TabIndex        =   7
      Top             =   645
      Width           =   900
      _extentx        =   1588
      _extenty        =   609
      style           =   7
      caption         =   "OK"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmDistortionSetup.frx":0030
   End
   Begin MP3JukeBox.isButton isButton1 
      Height          =   495
      Left            =   165
      TabIndex        =   6
      Top             =   570
      Width           =   915
      _extentx        =   1614
      _extenty        =   873
      style           =   7
      caption         =   "Default Settings"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmDistortionSetup.frx":0054
   End
   Begin MP3JukeBox.ctlEBSlider ctlDistortionLevel 
      Height          =   165
      Left            =   600
      TabIndex        =   3
      Top             =   165
      Width           =   2145
      _extentx        =   3784
      _extenty        =   291
      slidercolor     =   255
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
      Left            =   1440
      TabIndex        =   5
      Top             =   1020
      Width           =   1305
   End
   Begin VB.Label lblDistortionLabel4 
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
      Left            =   2535
      TabIndex        =   4
      Top             =   330
      Width           =   210
   End
   Begin VB.Label lblDistortionLabel3 
      AutoSize        =   -1  'True
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
      Left            =   600
      TabIndex        =   2
      Top             =   330
      Width           =   105
   End
   Begin VB.Label lblDistortionLevelTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.50"
      Height          =   285
      Left            =   2820
      TabIndex        =   1
      Top             =   105
      Width           =   420
   End
   Begin VB.Label lblDistortionLevel 
      BackColor       =   &H00000000&
      Caption         =   "Level"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   450
   End
End
Attribute VB_Name = "frmDistortionSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdDistortionCancel_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmDistortionSetup.cmdDistortionCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdDistortionDefault_Click()
    On Error GoTo ErrorTrap
    ctlDistortionLevel.Value = 50
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmDistortionSetup.cmdDistortionDefault_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdDistortionOK_Click()
    On Error GoTo ErrorTrap
    KeySection = "Distortion"
    KeyKey = "Level"
    KeyValue = Format$(iniDistortionLevel, "0.00")
    saveINI
 
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmDistortionSetup.cmdDistortionOK_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlDistortionLevel_Changed()
    On Error GoTo ErrorTrap
    iniDistortionLevel = ctlDistortionLevel.Value / 100
    lblDistortionLevelTot.Caption = Format$(iniDistortionLevel, "0.00")
    result = FMOD_DSP_SetParameter(DspDistortionFilter, FMOD_DSP_DISTORTION_LEVEL, iniDistortionLevel)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmDistortionSetup.ctlDistortionLevel_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
 
    ctlDistortionLevel.Value = iniDistortionLevel * 100
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmDistortionSetup.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub isButton1_Click()
    ctlDistortionLevel.Value = 50
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:31 PM) 1 + 80 = 81 Lines Thanks Ulli for inspiration and lots of code.


