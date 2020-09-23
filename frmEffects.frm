VERSION 5.00
Begin VB.Form frmEffects 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdEFFECTSCommand9 
      Height          =   300
      Left            =   945
      TabIndex        =   70
      Top             =   5190
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      style           =   7
      caption         =   "Configure"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":0000
   End
   Begin MP3JukeBox.isButton cmdEFFECTSCommand8 
      Height          =   300
      Left            =   945
      TabIndex        =   69
      Top             =   4575
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      style           =   7
      caption         =   "Configure"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":0024
   End
   Begin MP3JukeBox.isButton cmdEFFECTSCommand7 
      Height          =   300
      Left            =   945
      TabIndex        =   68
      Top             =   3916
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      style           =   7
      caption         =   "Configure"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":0048
   End
   Begin MP3JukeBox.isButton cmdEFFECTSCommand6 
      Height          =   300
      Left            =   945
      TabIndex        =   67
      Top             =   3249
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      style           =   7
      caption         =   "Configure"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":006C
   End
   Begin MP3JukeBox.isButton cmdEFFECTSCommand5 
      Height          =   300
      Left            =   945
      TabIndex        =   66
      Top             =   2625
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      style           =   7
      caption         =   "Configure"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":0090
   End
   Begin MP3JukeBox.isButton cmdEFFECTSCommand4 
      Height          =   300
      Left            =   945
      TabIndex        =   65
      Top             =   2010
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      style           =   7
      caption         =   "Configure"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":00B4
   End
   Begin MP3JukeBox.isButton cmdEFFECTSCommand3 
      Height          =   300
      Left            =   945
      TabIndex        =   64
      Top             =   1368
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      style           =   7
      caption         =   "Configure"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":00D8
   End
   Begin MP3JukeBox.isButton cmdEFFECTSCommand2 
      Height          =   300
      Left            =   945
      TabIndex        =   63
      Top             =   686
      Width           =   960
      _extentx        =   1693
      _extenty        =   529
      style           =   7
      caption         =   "Configure"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":00FC
   End
   Begin MP3JukeBox.isButton cmdEFFECTSCommand1 
      Height          =   300
      Left            =   3270
      TabIndex        =   62
      Top             =   255
      Width           =   765
      _extentx        =   1349
      _extenty        =   529
      style           =   7
      caption         =   "Exit"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":0120
   End
   Begin MP3JukeBox.isButton Command1 
      Height          =   300
      Left            =   2745
      TabIndex        =   61
      Top             =   4815
      Width           =   765
      _extentx        =   1349
      _extenty        =   529
      style           =   7
      caption         =   "Default"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmEffects.frx":0144
   End
   Begin VB.CheckBox chkEFFECTSCheck5 
      BackColor       =   &H00000000&
      Caption         =   "Flange Filter"
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
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   2938
      Width           =   1410
   End
   Begin VB.CheckBox chkEFFECTSCheck8 
      BackColor       =   &H00000000&
      Caption         =   "Normalize Filter"
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
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   4879
      Width           =   1650
   End
   Begin VB.CheckBox chkEFFECTSCheck7 
      BackColor       =   &H00000000&
      Caption         =   "Lowpass Filter"
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
      Height          =   270
      Left            =   90
      TabIndex        =   5
      Top             =   4227
      Width           =   1635
   End
   Begin VB.CheckBox chkEFFECTSCheck6 
      BackColor       =   &H00000000&
      Caption         =   "Highpass Filter"
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
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   3560
      Width           =   1590
   End
   Begin VB.CheckBox chkEFFECTSCheck4 
      BackColor       =   &H00000000&
      Caption         =   "Echo Filter"
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
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   2316
      Width           =   1290
   End
   Begin VB.CheckBox chkEFFECTSCheck3 
      BackColor       =   &H00000000&
      Caption         =   "Distortion Filter"
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
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   1679
      Width           =   1635
   End
   Begin VB.CheckBox chkEFFECTSCheck2 
      BackColor       =   &H00000000&
      Caption         =   "Chorus Filter"
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
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   997
      Width           =   1425
   End
   Begin VB.CheckBox chkEFFECTSCheck1 
      BackColor       =   &H00000000&
      Caption         =   "Reverb Filter"
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
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Top             =   315
      UseMaskColor    =   -1  'True
      Width           =   1425
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide9 
      Height          =   3075
      Left            =   3975
      TabIndex        =   8
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide7 
      Height          =   3075
      Left            =   3615
      TabIndex        =   9
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide6 
      Height          =   3075
      Left            =   3435
      TabIndex        =   10
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide5 
      Height          =   3075
      Left            =   3255
      TabIndex        =   11
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide4 
      Height          =   3075
      Left            =   3075
      TabIndex        =   12
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide3 
      Height          =   3075
      Left            =   2895
      TabIndex        =   13
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide1 
      Height          =   3075
      Left            =   2535
      TabIndex        =   14
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide0 
      Height          =   3075
      Left            =   2355
      TabIndex        =   15
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide2 
      Height          =   3075
      Left            =   2715
      TabIndex        =   16
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
   End
   Begin MP3JukeBox.ctlEBSlider EqSlide8 
      Height          =   3075
      Left            =   3795
      TabIndex        =   17
      Top             =   765
      Width           =   120
      _extentx        =   212
      _extenty        =   5424
      max             =   300
      value           =   150
      slidercolor     =   255
      orientation     =   1
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
      Left            =   2520
      TabIndex        =   60
      Top             =   5310
      Width           =   1305
   End
   Begin VB.Line linEFFECTSLine2 
      BorderColor     =   &H0000FFFF&
      X1              =   1965
      X2              =   1965
      Y1              =   720
      Y2              =   5610
   End
   Begin VB.Label lblEFFECTSLabel40 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "8"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2355
      TabIndex        =   59
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2355
      TabIndex        =   58
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2535
      TabIndex        =   57
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "7"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2535
      TabIndex        =   56
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "6"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2895
      TabIndex        =   55
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel10 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2895
      TabIndex        =   54
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel12 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3075
      TabIndex        =   53
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3075
      TabIndex        =   52
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel15 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3075
      TabIndex        =   51
      Top             =   4440
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel16 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "3"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3255
      TabIndex        =   50
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel17 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3255
      TabIndex        =   49
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel20 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "6"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3435
      TabIndex        =   48
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel21 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3435
      TabIndex        =   47
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel23 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3435
      TabIndex        =   46
      Top             =   4440
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel24 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3615
      TabIndex        =   45
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel25 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "2"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3615
      TabIndex        =   44
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel28 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3615
      TabIndex        =   43
      Top             =   4620
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel34 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3975
      TabIndex        =   42
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel35 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "6"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3975
      TabIndex        =   41
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel37 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3975
      TabIndex        =   40
      Top             =   4440
      Width           =   135
   End
   Begin VB.Label lblEFFECTSLabel38 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3975
      TabIndex        =   39
      Top             =   4620
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "3"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2715
      TabIndex        =   38
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2715
      TabIndex        =   37
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2535
      TabIndex        =   36
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel11 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2895
      TabIndex        =   35
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel14 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3075
      TabIndex        =   34
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel18 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3255
      TabIndex        =   33
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel22 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3435
      TabIndex        =   32
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel26 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3615
      TabIndex        =   31
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel36 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3975
      TabIndex        =   30
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel8 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   2715
      TabIndex        =   29
      Top             =   4275
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel39 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   " Hz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   28
      Top             =   4815
      Width           =   405
   End
   Begin VB.Line linEFFECTSLine1 
      BorderColor     =   &H0000FFFF&
      X1              =   1950
      X2              =   4275
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Label lblEFFECTSLabel41 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   2130
      TabIndex        =   27
      Top             =   2205
      Width           =   90
   End
   Begin VB.Label lblEFFECTSLabel42 
      BackColor       =   &H00000000&
      Caption         =   "+12"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   2025
      TabIndex        =   26
      Top             =   780
      Width           =   270
   End
   Begin VB.Label lblEFFECTSLabel43 
      BackColor       =   &H00000000&
      Caption         =   "-12"
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   2085
      TabIndex        =   25
      Top             =   3615
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3630
      TabIndex        =   24
      Top             =   4440
      Width           =   105
   End
   Begin VB.Label lblEFFECTSLabel31 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3795
      TabIndex        =   23
      Top             =   4305
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel33 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3795
      TabIndex        =   22
      Top             =   4620
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel30 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "4"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3795
      TabIndex        =   21
      Top             =   4095
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel29 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3795
      TabIndex        =   20
      Top             =   3930
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel19 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3255
      TabIndex        =   19
      Top             =   4440
      Width           =   120
   End
   Begin VB.Label lblEFFECTSLabel27 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   165
      Left            =   3795
      TabIndex        =   18
      Top             =   4440
      Width           =   120
   End
End
Attribute VB_Name = "frmEffects"
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
Private Sub chkEFFECTSCheck1_Click() 'Reverb Filter
    On Error GoTo ErrorTrap
    If chkEFFECTSCheck1.Value = 0 Then
' Unchecked
        iniFilterReverb = "unchecked"
        If DspReverbFilter = 0 Then
            GoTo chkEFFECTSCheck1_ClickExit
        Else
            result = FMOD_DSP_GetActive(DspReverbFilter, active)
            ERRCHECK (result)
            If active Then
                result = FMOD_DSP_Remove(DspReverbFilter)
                ERRCHECK (result)
            End If
            DspReverbFilter = 0
            GoTo chkEFFECTSCheck1_ClickExit
        End If
    Else
' Checked
        iniFilterReverb = "checked"
        If DspReverbFilter > 0 Then
            result = FMOD_DSP_GetActive(DspReverbFilter, active)
            ERRCHECK (result)
            If Not active Then
                result = FMOD_System_AddDSP(system, DspReverbFilter)
                ERRCHECK (result)
            End If
            GoTo chkEFFECTSCheck1_ClickExit
        Else
            result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_REVERB, DspReverbFilter)
            ERRCHECK (result)
            result = FMOD_System_AddDSP(system, DspReverbFilter)
            ERRCHECK (result)
        End If
    End If
chkEFFECTSCheck1_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.chkEFFECTSCheck1_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub chkEFFECTSCheck2_Click() ' Chorus Filter
    On Error GoTo ErrorTrap
    If chkEFFECTSCheck2.Value = 0 Then
' Unchecked
        iniFilterChorus = "unchecked"
        If DspChorusFilter = 0 Then
            GoTo chkEFFECTSCheck2_ClickExit
        Else
            result = FMOD_DSP_GetActive(DspChorusFilter, active)
            ERRCHECK (result)
            If active Then
                result = FMOD_DSP_Remove(DspChorusFilter)
                ERRCHECK (result)
            End If
            DspChorusFilter = 0
            GoTo chkEFFECTSCheck2_ClickExit
        End If
    Else
' Checked
        iniFilterChorus = "checked"
        If DspChorusFilter > 0 Then
            result = FMOD_DSP_GetActive(DspChorusFilter, active)
            ERRCHECK (result)
            If Not active Then
                result = FMOD_System_AddDSP(system, DspChorusFilter)
                ERRCHECK (result)
            End If
            GoTo chkEFFECTSCheck2_ClickExit
        Else
            result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_CHORUS, DspChorusFilter)
            ERRCHECK (result)
            result = FMOD_System_AddDSP(system, DspChorusFilter)
            ERRCHECK (result)
        End If
    End If
chkEFFECTSCheck2_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.chkEFFECTSCheck2_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub chkEFFECTSCheck3_Click() ' Distortion Filter
    On Error GoTo ErrorTrap
    If chkEFFECTSCheck3.Value = 0 Then
' Unchecked
        iniFilterDistortion = "unchecked"
        If DspDistortionFilter = 0 Then
            GoTo chkEFFECTSCheck3_ClickExit
        Else
            result = FMOD_DSP_GetActive(DspDistortionFilter, active)
            ERRCHECK (result)
            If active Then
                result = FMOD_DSP_Remove(DspDistortionFilter)
                ERRCHECK (result)
            End If
            DspDistortionFilter = 0
            GoTo chkEFFECTSCheck3_ClickExit
        End If
    Else
' Checked
        iniFilterDistortion = "checked"
        If DspDistortionFilter > 0 Then
            result = FMOD_DSP_GetActive(DspDistortionFilter, active)
            ERRCHECK (result)
            If Not active Then
                result = FMOD_System_AddDSP(system, DspDistortionFilter)
                ERRCHECK (result)
            End If
            GoTo chkEFFECTSCheck3_ClickExit
        Else
            result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_DISTORTION, DspDistortionFilter)
            ERRCHECK (result)
            result = FMOD_System_AddDSP(system, DspDistortionFilter)
            ERRCHECK (result)
        End If
    End If
chkEFFECTSCheck3_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.chkEFFECTSCheck3_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub chkEFFECTSCheck4_Click() ' Echo Filter
    On Error GoTo ErrorTrap
    If chkEFFECTSCheck4.Value = 0 Then
' Unchecked
        iniFilterEcho = "unchecked"
        If DspEchoFilter = 0 Then
            GoTo chkEFFECTSCheck4_ClickExit
        Else
            result = FMOD_DSP_GetActive(DspEchoFilter, active)
            ERRCHECK (result)
            If active Then
                result = FMOD_DSP_Remove(DspEchoFilter)
                ERRCHECK (result)
            End If
            DspEchoFilter = 0
            GoTo chkEFFECTSCheck4_ClickExit
        End If
    Else
' Checked
        iniFilterEcho = "checked"
        If DspEchoFilter > 0 Then
            result = FMOD_DSP_GetActive(DspEchoFilter, active)
            ERRCHECK (result)
            If Not active Then
                result = FMOD_System_AddDSP(system, DspEchoFilter)
                ERRCHECK (result)
            End If
            GoTo chkEFFECTSCheck4_ClickExit
        Else
            result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_ECHO, DspEchoFilter)
            ERRCHECK (result)
            result = FMOD_System_AddDSP(system, DspEchoFilter)
            ERRCHECK (result)
        End If
    End If
chkEFFECTSCheck4_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.chkEFFECTSCheck4_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub chkEFFECTSCheck5_Click() ' Flange Filter
    On Error GoTo ErrorTrap
    If chkEFFECTSCheck5.Value = 0 Then
' Unchecked
        iniFilterFlange = "unchecked"
        If DspFlangeFilter = 0 Then
            GoTo chkEFFECTSCheck5_ClickExit
        Else
            result = FMOD_DSP_GetActive(DspFlangeFilter, active)
            ERRCHECK (result)
            If active Then
                result = FMOD_DSP_Remove(DspFlangeFilter)
                ERRCHECK (result)
            End If
            DspFlangeFilter = 0
            GoTo chkEFFECTSCheck5_ClickExit
        End If
    Else
' Checked
        iniFilterFlange = "checked"
        If DspFlangeFilter > 0 Then
            result = FMOD_DSP_GetActive(DspFlangeFilter, active)
            ERRCHECK (result)
            If Not active Then
                result = FMOD_System_AddDSP(system, DspFlangeFilter)
                ERRCHECK (result)
            End If
            GoTo chkEFFECTSCheck5_ClickExit
        Else
            result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_FLANGE, DspFlangeFilter)
            ERRCHECK (result)
            result = FMOD_System_AddDSP(system, DspFlangeFilter)
            ERRCHECK (result)
        End If
    End If
chkEFFECTSCheck5_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.chkEFFECTSCheck5_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub chkEFFECTSCheck6_Click() ' Highpass Filter
    On Error GoTo ErrorTrap
    If chkEFFECTSCheck6.Value = 0 Then
' Unchecked
        iniFilterHighpass = "unchecked"
        If DspHighpassFilter = 0 Then
            GoTo chkEFFECTSCheck6_ClickExit
        Else
            result = FMOD_DSP_GetActive(DspHighpassFilter, active)
            ERRCHECK (result)
            If active Then
                result = FMOD_DSP_Remove(DspHighpassFilter)
                ERRCHECK (result)
            End If
            DspHighpassFilter = 0
            GoTo chkEFFECTSCheck6_ClickExit
        End If
    Else
' Checked
        iniFilterHighpass = "checked"
        If DspHighpassFilter > 0 Then
            result = FMOD_DSP_GetActive(DspHighpassFilter, active)
            ERRCHECK (result)
            If Not active Then
                result = FMOD_System_AddDSP(system, DspHighpassFilter)
                ERRCHECK (result)
            End If
            GoTo chkEFFECTSCheck6_ClickExit
        Else
            result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_HIGHPASS, DspHighpassFilter)
            ERRCHECK (result)
            result = FMOD_System_AddDSP(system, DspHighpassFilter)
            ERRCHECK (result)
        End If
    End If
chkEFFECTSCheck6_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.chkEFFECTSCheck6_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub chkEFFECTSCheck7_Click()  ' Lowpass Filter
    On Error GoTo ErrorTrap
    If chkEFFECTSCheck7.Value = 0 Then
' Unchecked
        iniFilterLowpass = "unchecked"
        If DspLowpassFilter = 0 Then
            GoTo chkEFFECTSCheck7_ClickExit
        Else
            result = FMOD_DSP_GetActive(DspLowpassFilter, active)
            ERRCHECK (result)
            If active Then
                result = FMOD_DSP_Remove(DspLowpassFilter)
                ERRCHECK (result)
            End If
            DspLowpassFilter = 0
            GoTo chkEFFECTSCheck7_ClickExit
        End If
    Else
' Checked
        iniFilterLowpass = "checked"
        If DspLowpassFilter > 0 Then
            result = FMOD_DSP_GetActive(DspLowpassFilter, active)
            ERRCHECK (result)
            If Not active Then
                result = FMOD_System_AddDSP(system, DspLowpassFilter)
                ERRCHECK (result)
            End If
            GoTo chkEFFECTSCheck7_ClickExit
        Else
            result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_LOWPASS, DspLowpassFilter)
            ERRCHECK (result)
            result = FMOD_System_AddDSP(system, DspLowpassFilter)
            ERRCHECK (result)
        End If
    End If
chkEFFECTSCheck7_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.chkEFFECTSCheck7_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub chkEFFECTSCheck8_Click() ' Normalize Filter
    On Error GoTo ErrorTrap
    If chkEFFECTSCheck8.Value = 0 Then
' Unchecked
        iniFilterNormalize = "unchecked"
        If DspNormalizeFilter = 0 Then
            GoTo chkEFFECTSCheck8_ClickExit
        Else
            result = FMOD_DSP_GetActive(DspNormalizeFilter, active)
            ERRCHECK (result)
            If active Then
                result = FMOD_DSP_Remove(DspNormalizeFilter)
                ERRCHECK (result)
            End If
            DspNormalizeFilter = 0
            GoTo chkEFFECTSCheck8_ClickExit
        End If
    Else
' Checked
        iniFilterNormalize = "checked"
        If DspNormalizeFilter > 0 Then
            result = FMOD_DSP_GetActive(DspNormalizeFilter, active)
            ERRCHECK (result)
            If Not active Then
                result = FMOD_System_AddDSP(system, DspNormalizeFilter)
                ERRCHECK (result)
            End If
            GoTo chkEFFECTSCheck8_ClickExit
        Else
            result = FMOD_System_CreateDSPByType(system, FMOD_DSP_TYPE_NORMALIZE, DspNormalizeFilter)
            ERRCHECK (result)
            result = FMOD_System_AddDSP(system, DspNormalizeFilter)
            ERRCHECK (result)
        End If
    End If
chkEFFECTSCheck8_ClickExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.chkEFFECTSCheck8_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand1_Click()
    On Error GoTo ErrorTrap
 
    KeySection = "Equalizer"
    KeyKey = "Eq0Pos"
    KeyValue = Format$(iniEq0Pos, "0.00")
    saveINI
    KeyKey = "Eq1Pos"
    KeyValue = Format$(iniEq1Pos, "0.00")
    saveINI
    KeyKey = "Eq2Pos"
    KeyValue = Format$(iniEq2Pos, "0.00")
    saveINI
    KeyKey = "Eq3Pos"
    KeyValue = Format$(iniEq3Pos, "0.00")
    saveINI
    KeyKey = "Eq4Pos"
    KeyValue = Format$(iniEq4Pos, "0.00")
    saveINI
    KeyKey = "Eq5Pos"
    KeyValue = Format$(iniEq5Pos, "0.00")
    saveINI
    KeyKey = "Eq6Pos"
    KeyValue = Format$(iniEq6Pos, "0.00")
    saveINI
    KeyKey = "Eq7Pos"
    KeyValue = Format$(iniEq7Pos, "0.00")
    saveINI
    KeyKey = "Eq8Pos"
    KeyValue = Format$(iniEq8Pos, "0.00")
    saveINI
    KeyKey = "Eq9Pos"
    KeyValue = Format$(iniEq9Pos, "0.00")
    saveINI
    KeySection = "Filters"
    KeyKey = "ReverbFilter"
    KeyValue = iniFilterReverb
    saveINI
    KeyKey = "ChorusFilter"
    KeyValue = iniFilterChorus
    saveINI
    KeyKey = "DistortionFilter"
    KeyValue = iniFilterDistortion
    saveINI
    KeyKey = "EchoFilter"
    KeyValue = iniFilterEcho
    saveINI
    KeyKey = "FlangeFilter"
    KeyValue = iniFilterFlange
    saveINI
    KeyKey = "HighpassFilter"
    KeyValue = iniFilterHighpass
    saveINI
    KeyKey = "LowpassFilter"
    KeyValue = iniFilterLowpass
    saveINI
    KeyKey = "NormalizeFilter"
    KeyValue = iniFilterNormalize
    saveINI
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand1_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand2_Click() ' Reverb
    On Error GoTo ErrorTrap
    If DspReverbFilter > 0 Then
        frmReverbSetup.Show vbModal, Me
    Else
        frmMsgBox.SMessageModal "ERROR: Reverb Filter not active!", 48, "Reverb Filter"
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand2_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand3_Click() ' Chorus
    On Error GoTo ErrorTrap
    If DspChorusFilter > 0 Then
        frmChorusSetup.Show vbModal, Me
    Else
        frmMsgBox.SMessageModal "ERROR: Chorus Filter not active!", 48, "Chorus Filter"
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand3_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand4_Click() ' Distortion
    On Error GoTo ErrorTrap
    If DspDistortionFilter > 0 Then
        frmDistortionSetup.Show vbModal, Me
    Else
        frmMsgBox.SMessageModal "ERROR: Distortion Filter not active!", 48, "Distortion Filter"
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand4_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand5_Click() ' Echo
    On Error GoTo ErrorTrap
    If DspEchoFilter > 0 Then
        frmEchoSetup.Show vbModal, Me
    Else
        frmMsgBox.SMessageModal "ERROR: Echo Filter not active!", 48, "Echo Filter"
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand5_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand6_Click() ' Flange
    On Error GoTo ErrorTrap
    If DspFlangeFilter > 0 Then
        frmFlangeSetup.Show vbModal, Me
    Else
        frmMsgBox.SMessageModal "ERROR: Flange Filter not active!", 48, "Flange Filter"
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand6_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand7_Click() ' Highpass
    On Error GoTo ErrorTrap
    If DspHighpassFilter > 0 Then
        frmHighpassSetup.Show vbModal, Me
    Else
        frmMsgBox.SMessageModal "ERROR: Highpass Filter not active!", 48, "Highpass Filter"
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand7_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand8_Click() ' Lowpass
    On Error GoTo ErrorTrap
    If DspLowpassFilter > 0 Then
        frmLowpassSetup.Show vbModal, Me
    Else
        frmMsgBox.SMessageModal "ERROR: Lowpass Filter not active!", 48, "Lowpass Filter"
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand8_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdEFFECTSCommand9_Click() ' Normalze
    On Error GoTo ErrorTrap
    If DspNormalizeFilter > 0 Then
        frmNormalizeSetup.Show vbModal, Me
    Else
        frmMsgBox.SMessageModal "ERROR: Normalize Filter not active!", 48, "Normalize Filter"
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.cmdEFFECTSCommand9_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Command1_Click()
    On Error GoTo ErrorTrap
    EqSlide0.Value = 150
    EqSlide1.Value = 150
    EqSlide2.Value = 150
    EqSlide3.Value = 150
    EqSlide4.Value = 150
    EqSlide5.Value = 150
    EqSlide6.Value = 150
    EqSlide7.Value = 150
    EqSlide8.Value = 150
    EqSlide9.Value = 150
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.Command1_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide0_Changed()
    On Error GoTo ErrorTrap
    iniEq0Pos = EqSlide0.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq0Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide0_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide1_Changed()
    On Error GoTo ErrorTrap
    iniEq1Pos = EqSlide1.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq1Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide1_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide2_Changed()
    On Error GoTo ErrorTrap
    iniEq2Pos = EqSlide2.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq2Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide2_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide3_Changed()
    On Error GoTo ErrorTrap
    iniEq3Pos = EqSlide3.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq3Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide3_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide4_Changed()
    On Error GoTo ErrorTrap
    iniEq4Pos = EqSlide4.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq4Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide4_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide5_Changed()
    On Error GoTo ErrorTrap
    iniEq5Pos = EqSlide5.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq5Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide5_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide6_Changed()
    On Error GoTo ErrorTrap
    iniEq6Pos = EqSlide6.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq6Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide6_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide7_Changed()
    On Error GoTo ErrorTrap
    iniEq7Pos = EqSlide7.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq7Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide7_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide8_Changed()
    On Error GoTo ErrorTrap
    iniEq8Pos = EqSlide8.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq8Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide8_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub EqSlide9_Changed()
    On Error GoTo ErrorTrap
    iniEq9Pos = EqSlide9.Value / 100
    result = FMOD_DSP_SetParameter(Eq9, FMOD_DSP_PARAMEQ_GAIN, iniEq9Pos)
    ERRCHECK (result)
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.EqSlide9_Changed" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
Dim rgnMain As Long
    On Error GoTo ErrorTrap
    nBytes = 624
    LoadBytes
    rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))
    SetWindowRgn Me.hwnd, rgnMain, True
 
    EqSlide0.Value = iniEq0Pos * 100
    EqSlide1.Value = iniEq1Pos * 100
    EqSlide2.Value = iniEq2Pos * 100
    EqSlide3.Value = iniEq3Pos * 100
    EqSlide4.Value = iniEq4Pos * 100
    EqSlide5.Value = iniEq5Pos * 100
    EqSlide6.Value = iniEq6Pos * 100
    EqSlide7.Value = iniEq7Pos * 100
    EqSlide8.Value = iniEq8Pos * 100
    EqSlide9.Value = iniEq9Pos * 100
    If iniFilterReverb = "unchecked" Then
        chkEFFECTSCheck1.Value = 0
    Else
        chkEFFECTSCheck1.Value = 1
    End If
    If iniFilterChorus = "unchecked" Then
        chkEFFECTSCheck2.Value = 0
    Else
        chkEFFECTSCheck2.Value = 1
    End If
    If iniFilterDistortion = "unchecked" Then
        chkEFFECTSCheck3.Value = 0
    Else
        chkEFFECTSCheck3.Value = 1
    End If
    If iniFilterEcho = "unchecked" Then
        chkEFFECTSCheck4.Value = 0
    Else
        chkEFFECTSCheck4.Value = 1
    End If
    If iniFilterFlange = "unchecked" Then
        chkEFFECTSCheck5.Value = 0
    Else
        chkEFFECTSCheck5.Value = 1
    End If
    If iniFilterHighpass = "unchecked" Then
        chkEFFECTSCheck6.Value = 0
    Else
        chkEFFECTSCheck6.Value = 1
    End If
    If iniFilterLowpass = "unchecked" Then
        chkEFFECTSCheck7.Value = 0
    Else
        chkEFFECTSCheck7.Value = 1
    End If
    If iniFilterNormalize = "unchecked" Then
        chkEFFECTSCheck8.Value = 0
    Else
        chkEFFECTSCheck8.Value = 1
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmEFFECTS.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
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
':)Code Fixer V3.0.9 (9/15/2005 1:31:46 PM) 7 + 1030 = 1037 Lines Thanks Ulli for inspiration and lots of code.


