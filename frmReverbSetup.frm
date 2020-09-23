VERSION 5.00
Begin VB.Form frmReverbSetup 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3885
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmReverbSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MP3JukeBox.isButton cmdReverbCancel 
      Height          =   330
      Left            =   3150
      TabIndex        =   35
      Top             =   3195
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   582
      Style           =   7
      Caption         =   "Cancel"
      USeCustomColors =   -1  'True
      BackColor       =   65535
      HighlightColor  =   255
      FontHighlightColor=   255
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MP3JukeBox.isButton cmdReverbOK 
      Height          =   330
      Left            =   1605
      TabIndex        =   34
      Top             =   3195
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   582
      Style           =   7
      Caption         =   "OK"
      USeCustomColors =   -1  'True
      BackColor       =   65535
      HighlightColor  =   255
      FontHighlightColor=   255
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MP3JukeBox.isButton cmdReverbDefault 
      Height          =   540
      Left            =   405
      TabIndex        =   33
      Top             =   3090
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   953
      Style           =   7
      Caption         =   "Default Settings"
      USeCustomColors =   -1  'True
      BackColor       =   65535
      HighlightColor  =   255
      FontHighlightColor=   255
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkReverbFreeze 
      BackColor       =   &H00000000&
      Caption         =   "Freeze"
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2625
      TabIndex        =   1
      Top             =   2760
      Width           =   1005
   End
   Begin VB.CheckBox chkReverbNormal 
      BackColor       =   &H00000000&
      Caption         =   "Normal"
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   2745
      Value           =   1  'Checked
      Width           =   1005
   End
   Begin MP3JukeBox.ctlEBSlider ctlReverbWidth 
      Height          =   165
      Left            =   1080
      TabIndex        =   2
      Top             =   2310
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   291
      SliderColor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlReverbRoomSize 
      Height          =   165
      Left            =   1080
      TabIndex        =   3
      Top             =   300
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   291
      SliderColor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlReverbWetMix 
      Height          =   165
      Left            =   1080
      TabIndex        =   4
      Top             =   1305
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   291
      SliderColor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlReverbDamp 
      Height          =   165
      Left            =   1080
      TabIndex        =   5
      Top             =   795
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   291
      SliderColor     =   255
   End
   Begin MP3JukeBox.ctlEBSlider ctlReverbDryMix 
      Height          =   165
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   291
      SliderColor     =   255
   End
   Begin VB.Label lblReverbRoomSize 
      BackColor       =   &H00000000&
      Caption         =   "Room Size"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   105
      TabIndex        =   32
      Top             =   270
      Width           =   900
   End
   Begin VB.Label lblReverbWetMix 
      BackColor       =   &H00000000&
      Caption         =   "Wet Mix"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   105
      TabIndex        =   31
      Top             =   1275
      Width           =   690
   End
   Begin VB.Label lblReverbWidth 
      BackColor       =   &H00000000&
      Caption         =   "Width"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   105
      TabIndex        =   30
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblReverbMode 
      BackColor       =   &H00000000&
      Caption         =   "Mode"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   105
      TabIndex        =   29
      Top             =   2790
      Width           =   465
   End
   Begin VB.Label lblReverbLabel8 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3525
      TabIndex        =   28
      Top             =   465
      Width           =   105
   End
   Begin VB.Label lblReverbRoomSizeTot 
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
      Left            =   3690
      TabIndex        =   27
      Top             =   240
      Width           =   570
   End
   Begin VB.Label lblReverbWidthTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1.00"
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
      Left            =   3675
      TabIndex        =   26
      Top             =   2250
      Width           =   570
   End
   Begin VB.Label lblReverbLabel21 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3525
      TabIndex        =   25
      Top             =   2490
      Width           =   105
   End
   Begin VB.Label lblReverbWetMixTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.33"
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
      Left            =   3660
      TabIndex        =   24
      Top             =   1245
      Width           =   570
   End
   Begin VB.Label lblReverbLabel17 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3525
      TabIndex        =   23
      Top             =   1500
      Width           =   105
   End
   Begin VB.Label lblReverbLabel15 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3525
      TabIndex        =   22
      Top             =   1005
      Width           =   105
   End
   Begin VB.Label lblReverbDampTot 
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
      Left            =   3690
      TabIndex        =   21
      Top             =   750
      Width           =   570
   End
   Begin VB.Label lblReverbDamp 
      BackColor       =   &H00000000&
      Caption         =   "Damp"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   105
      TabIndex        =   20
      Top             =   765
      Width           =   495
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
      Left            =   2070
      TabIndex        =   19
      Top             =   1500
      Width           =   585
   End
   Begin VB.Label lblReverbLabel19 
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   3525
      TabIndex        =   18
      Top             =   1995
      Width           =   105
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
      Left            =   2070
      TabIndex        =   17
      Top             =   1995
      Width           =   585
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Stereo width"
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
      Left            =   1875
      TabIndex        =   16
      Top             =   2490
      Width           =   960
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "damping"
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
      Left            =   2040
      TabIndex        =   15
      Top             =   1005
      Width           =   645
   End
   Begin VB.Label lblReverbDryMixTot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.66"
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
      Left            =   3675
      TabIndex        =   14
      Top             =   1755
      Width           =   570
   End
   Begin VB.Label lblReverbDryMix 
      BackColor       =   &H00000000&
      Caption         =   "Dry Mix"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   105
      TabIndex        =   13
      Top             =   1770
      Width           =   630
   End
   Begin VB.Label lblReverbLabel18 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1080
      TabIndex        =   12
      Top             =   1995
      Width           =   105
   End
   Begin VB.Label lblReverbLabel14 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1080
      TabIndex        =   11
      Top             =   1005
      Width           =   105
   End
   Begin VB.Label lblReverbLabel16 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1080
      TabIndex        =   10
      Top             =   1500
      Width           =   105
   End
   Begin VB.Label lblReverbLabel20 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1080
      TabIndex        =   9
      Top             =   2490
      Width           =   105
   End
   Begin VB.Label lblReverbLabel7 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1080
      TabIndex        =   8
      Top             =   465
      Width           =   105
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
      Left            =   2310
      TabIndex        =   7
      Top             =   3570
      Width           =   1305
   End
End
Attribute VB_Name = "frmReverbSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkReverbFreeze_Click()
    On Error GoTo ErrorTrap
    If chkReverbFreeze.Value = 0 Then
        chkReverbNormal.Value = 1
        iniReverbMode = "Normal"
        result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_MODE, 0)
        ERRCHECK (result)
    Else
        chkReverbNormal.Value = 0
        iniReverbMode = "Freeze"
        result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_MODE, 1)
        ERRCHECK (result)
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.chkReverbFreeze_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub chkReverbNormal_Click()
    On Error GoTo ErrorTrap
    If chkReverbNormal.Value = 0 Then
        chkReverbFreeze.Value = 1
        iniReverbMode = "Freeze"
        result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_MODE, 1)
        ERRCHECK (result)
    Else
        chkReverbFreeze.Value = 0
        iniReverbMode = "Normal"
        result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_MODE, 0)
        ERRCHECK (result)
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.chkReverbNormal_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdReverbCancel_Click()
    On Error GoTo ErrorTrap
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.cmdReverbCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdReverbDefault_Click()
    On Error GoTo ErrorTrap
    ctlReverbDamp.Value = 50
    ctlReverbDryMix.Value = 66
    ctlReverbRoomSize.Value = 50
    ctlReverbWetMix.Value = 33
    ctlReverbWidth.Value = 100
    chkReverbNormal.Value = 1
    chkReverbFreeze.Value = 0
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.cmdReverbDefault_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdReverbOK_Click()
    On Error GoTo ErrorTrap
    KeySection = "Reverb"
    KeyKey = "RoomSize"
    KeyValue = Format$(iniReverbRoomSize, "0.00")
    saveINI
    KeyKey = "Damp"
    KeyValue = Format$(iniReverbDamp, "0.00")
    saveINI
    KeyKey = "WetMix"
    KeyValue = Format$(iniReverbWetMix, "0.00")
    saveINI
    KeyKey = "DryMix"
    KeyValue = Format$(iniReverbDryMix, "0.00")
    saveINI
    KeyKey = "Width"
    KeyValue = Format$(iniReverbWidth, "0.00")
    saveINI
    KeyKey = "Mode"
    KeyValue = iniReverbMode
    saveINI
 
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.cmdReverbOK_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlReverbDamp_Changed()
    On Error GoTo ErrorTrap
    iniReverbDamp = ctlReverbDamp.Value / 100
    lblReverbDampTot.Caption = Format$(iniReverbDamp, "0.00")
    result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_DAMP, iniReverbDamp)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.ctlReverbDamp_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlReverbDryMix_Changed()
    On Error GoTo ErrorTrap
    iniReverbDryMix = ctlReverbDryMix.Value / 100
    lblReverbDryMixTot.Caption = Format$(iniReverbDryMix, "0.00")
    result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_DRYMIX, iniReverbDryMix)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.ctlReverbDryMix_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlReverbRoomSize_Changed()
    On Error GoTo ErrorTrap
    iniReverbRoomSize = ctlReverbRoomSize.Value / 100
    lblReverbRoomSizeTot.Caption = Format$(iniReverbRoomSize, "0.00")
    result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_ROOMSIZE, iniReverbRoomSize)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.ctlReverbRoomSize_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlReverbWetMix_Changed()
    On Error GoTo ErrorTrap
    iniReverbWetMix = ctlReverbWetMix.Value / 100
    lblReverbWetMixTot.Caption = Format$(iniReverbWetMix, "0.00")
    result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_WETMIX, iniReverbWetMix)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.ctlReverbWetMix_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub ctlReverbWidth_Changed()
    On Error GoTo ErrorTrap
    iniReverbWidth = ctlReverbWidth.Value / 100
    lblReverbWidthTot.Caption = Format$(iniReverbWidth, "0.00")
    result = FMOD_DSP_SetParameter(DspReverbFilter, FMOD_DSP_REVERB_WIDTH, iniReverbWidth)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.ctlReverbWidth_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
 
    ctlReverbDamp.Value = iniReverbDamp * 100
    ctlReverbDryMix.Value = iniReverbDryMix * 100
    ctlReverbRoomSize.Value = iniReverbRoomSize * 100
    ctlReverbWetMix.Value = iniReverbWetMix * 100
    ctlReverbWidth.Value = iniReverbWidth * 100
    If iniReverbMode = "Normal" Then
        chkReverbNormal.Value = 1
        chkReverbFreeze.Value = 0
    End If
    If iniReverbMode = "Freeze" Then
        chkReverbNormal.Value = 0
        chkReverbFreeze.Value = 1
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmReverbSetup.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:32:18 PM) 1 + 211 = 212 Lines Thanks Ulli for inspiration and lots of code.


