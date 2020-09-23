VERSION 5.00
Begin VB.Form frmFlangeSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2805
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3690
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
   Icon            =   "frmFlangeSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdFlangeCancel 
      Height          =   465
      Left            =   2490
      TabIndex        =   26
      Top             =   2025
      Width           =   840
      _extentx        =   1482
      _extenty        =   820
      style           =   7
      caption         =   "Cancel"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmFlangeSetup.frx":000C
   End
   Begin MP3JukeBox.isButton cmdFlangeOK 
      Height          =   465
      Left            =   1365
      TabIndex        =   25
      Top             =   2025
      Width           =   840
      _extentx        =   1482
      _extenty        =   820
      style           =   7
      caption         =   "OK"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmFlangeSetup.frx":0030
   End
   Begin MP3JukeBox.isButton cmdFlangeDefault 
      Height          =   645
      Left            =   240
      TabIndex        =   24
      Top             =   1935
      Width           =   840
      _extentx        =   1482
      _extenty        =   1138
      style           =   7
      caption         =   "Default Settings"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmFlangeSetup.frx":0054
   End
   Begin MP3JukeBox.ctlEBSlider ctlFlangeRate 
      Height          =   150
      Left            =   1035
      TabIndex        =   2
      Top             =   1582
      Width           =   1875
      _extentx        =   3307
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlFlangeDepth 
      Height          =   150
      Left            =   1035
      TabIndex        =   1
      Top             =   1112
      Width           =   1875
      _extentx        =   3307
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlFlangeDryMix 
      Height          =   150
      Left            =   1035
      TabIndex        =   0
      Top             =   172
      Width           =   1875
      _extentx        =   3307
      _extenty        =   265
      slidercolor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlFlangeWetMix 
      Height          =   150
      Left            =   1035
      TabIndex        =   21
      Top             =   642
      Width           =   1875
      _extentx        =   3307
      _extenty        =   265
      slidercolor     =   255
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
      Left            =   1740
      TabIndex        =   23
      Top             =   2505
      Width           =   1305
   End
   Begin VB.Label lblFlangeWetMix 
      BackColor       =   &H00000000&
      Caption         =   "Wet Mix"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   150
      TabIndex        =   22
      Top             =   605
      Width           =   690
   End
   Begin VB.Label lblFlangeDryMix 
      BackColor       =   &H00000000&
      Caption         =   "Dry Mix"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   150
      TabIndex        =   20
      Top             =   135
      Width           =   630
   End
   Begin VB.Label lblFlangeDepth 
      BackColor       =   &H00000000&
      Caption         =   "Depth"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   150
      TabIndex        =   19
      Top             =   1075
      Width           =   495
   End
   Begin VB.Label lblFlangeRate 
      BackColor       =   &H00000000&
      Caption         =   "Rate"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   150
      TabIndex        =   18
      Top             =   1545
      Width           =   390
   End
   Begin VB.Label Label3 
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
      Left            =   1867
      TabIndex        =   17
      Top             =   1740
      Width           =   210
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
      Left            =   1680
      TabIndex        =   16
      Top             =   795
      Width           =   585
   End
   Begin VB.Label Label1 
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
      Left            =   1680
      TabIndex        =   15
      Top             =   330
      Width           =   585
   End
   Begin VB.Label lblFlangeLabel16 
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
      Left            =   2580
      TabIndex        =   14
      Top             =   1740
      Width           =   330
   End
   Begin VB.Label lblFlangeLabel15 
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
      Left            =   1035
      TabIndex        =   13
      Top             =   1740
      Width           =   240
   End
   Begin VB.Label lblFlangeLabel14 
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
      Left            =   2610
      TabIndex        =   12
      Top             =   1290
      Width           =   300
   End
   Begin VB.Label lblFlangeLabel13 
      BackColor       =   &H00000000&
      Caption         =   "0.01"
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
      Left            =   1035
      TabIndex        =   11
      Top             =   1290
      Width           =   300
   End
   Begin VB.Label lblFlangeLabel12 
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
      Left            =   2610
      TabIndex        =   10
      Top             =   795
      Width           =   300
   End
   Begin VB.Label lblFlangeLabel11 
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
      Left            =   1035
      TabIndex        =   9
      Top             =   795
      Width           =   330
   End
   Begin VB.Label lblFlangeLabel10 
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
      Left            =   2610
      TabIndex        =   8
      Top             =   330
      Width           =   300
   End
   Begin VB.Label lblFlangeLabel9 
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
      Left            =   1035
      TabIndex        =   7
      Top             =   330
      Width           =   330
   End
   Begin VB.Label lblFlangeRateTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2940
      TabIndex        =   6
      Top             =   1537
      Width           =   600
   End
   Begin VB.Label lblFlangeDepthTot 
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
      Height          =   240
      Left            =   2940
      TabIndex        =   5
      Top             =   1067
      Width           =   600
   End
   Begin VB.Label lblFlangeWetMixTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.55"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2940
      TabIndex        =   4
      Top             =   597
      Width           =   600
   End
   Begin VB.Label lblFlangeDryMixTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.45"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2940
      TabIndex        =   3
      Top             =   127
      Width           =   600
   End
End
Attribute VB_Name = "frmFlangeSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdFlangeCancel_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmFlangeSetup.cmdFlangeCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdFlangeDefault_Click()
    On Error GoTo ErrorTrap
    ctlFlangeDepth.Value = 100
    ctlFlangeDryMix.Value = 45
    ctlFlangeRate.Value = 1
    ctlFlangeWetMix.Value = 55
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmFlangeSetup.cmdFlangeDefault_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdFlangeOK_Click()
    On Error GoTo ErrorTrap
    KeySection = "Flange"
    KeyKey = "Depth"
    KeyValue = Format$(iniFlangeDepth, "0.00")
    saveINI
    KeyKey = "DryMix"
    KeyValue = Format$(iniFlangeDryMix, "0.00")
    saveINI
    KeyKey = "Rate"
    KeyValue = Format$(iniFlangeRate, "0.0")
    saveINI
    KeyKey = "WetMix"
    KeyValue = Format$(iniFlangeWetMix, "0.00")
    saveINI
 
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmFlangeSetup.cmdFlangeOK_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlFlangeDepth_Changed()
    On Error GoTo ErrorTrap
    iniFlangeDepth = ctlFlangeDepth.Value / 100
    lblFlangeDepthTot.Caption = Format$(iniFlangeDepth, "0.00")
    result = FMOD_DSP_SetParameter(DspFlangeFilter, FMOD_DSP_FLANGE_DEPTH, iniFlangeDepth)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmFlangeSetup.ctlFlangeDepth_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlFlangeDryMix_Changed()
    On Error GoTo ErrorTrap
    iniFlangeDryMix = ctlFlangeDryMix.Value / 100
    lblFlangeDryMixTot.Caption = Format$(iniFlangeDryMix, "0.00")
    result = FMOD_DSP_SetParameter(DspFlangeFilter, FMOD_DSP_FLANGE_DRYMIX, iniFlangeDryMix)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmFlangeSetup.ctlFlangeDryMix_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlFlangeRate_Changed()
    On Error GoTo ErrorTrap
    iniFlangeRate = ctlFlangeRate.Value / 10
    lblFlangeRateTot.Caption = Format$(iniFlangeRate, "#0.0")
    result = FMOD_DSP_SetParameter(DspFlangeFilter, FMOD_DSP_FLANGE_RATE, iniFlangeRate)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmFlangeSetup.ctlFlangeRate_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlFlangeWetMix_Changed()
    On Error GoTo ErrorTrap
    iniFlangeWetMix = ctlFlangeWetMix.Value / 100
    lblFlangeWetMixTot.Caption = Format$(iniFlangeWetMix, "0.00")
    result = FMOD_DSP_SetParameter(DspFlangeFilter, FMOD_DSP_FLANGE_WETMIX, iniFlangeWetMix)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmFlangeSetup.ctlFlangeWetMix_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
 
    ctlFlangeDepth.Value = iniFlangeDepth * 100
    ctlFlangeDryMix.Value = iniFlangeDryMix * 100
    ctlFlangeRate.Value = iniFlangeRate * 10
    ctlFlangeWetMix.Value = iniFlangeWetMix * 100
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmFlangeSetup.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:36 PM) 1 + 134 = 135 Lines Thanks Ulli for inspiration and lots of code.


