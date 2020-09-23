VERSION 5.00
Begin VB.UserControl ctlEBSlider 
   BackColor       =   &H00000000&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   ScaleHeight     =   360
   ScaleWidth      =   4440
   ToolboxBitmap   =   "ctlEBSlider.ctx":0000
   Begin VB.PictureBox picSlider 
      BorderStyle     =   0  'None
      Height          =   200
      Left            =   120
      ScaleHeight     =   182
      ScaleMode       =   0  'User
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Line linGroove 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   4380
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line linGroove 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   4380
      Y1              =   180
      Y2              =   180
   End
End
Attribute VB_Name = "ctlEBSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'[Description]
'   EBSlider
'   A stand-alone slider control
'[Author]
'   Richard Allsebrook  <RA>    RichardAllsebrook@earlybirdmarketing.com
'[History]
'   V1.0.0  20/06/2001
'   Initial Release
'[Declarations]
'Property storage
Private lngMin                   As Long                   'Minimum value range
Private lngMax                   As Long                   'Maximum value range
Private lngValue                 As Long                   'Current Value
Private lngSliderWidth           As Long
Private zBorderStyle             As EBSliderBorderStyle
Private zOrientation             As EBSliderOrientation    'Current Orientation
'Event Stubs
Public Event Changed()
'Enums
Public Enum EBSliderOrientation
    EBHorizontal
    EBVertical
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private EBHorizontal, EBVertical
#End If
Public Enum EBSliderBorderStyle
    EBNone = 0
    EBSunkenOuter = &H2
    EBRaisedInner = &H4
    EBEtched = (EBSunkenOuter Or EBRaisedInner)
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private EBNone, EBSunkenOuter, EBRaisedInner, EBEtched
#End If
'API Stubs
'API UDTs
Private Type RECT
    Left                             As Long
    Top                              As Long
    Right                            As Long
    bottom                           As Long
End Type
'API Constants
Private Const BDR_RAISEDINNER    As Long = &H4
Private Const BF_BOTTOM          As Long = &H8
Private Const BF_LEFT            As Long = &H1
Private Const BF_RIGHT           As Long = &H4
Private Const BF_TOP             As Long = &H2
Private Const BF_RECT            As Double = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
                                                qrc As RECT, _
                                                ByVal edge As Long, _
                                                ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, _
                                               ByVal X1 As Long, _
                                               ByVal Y1 As Long, _
                                               ByVal X2 As Long, _
                                               ByVal Y2 As Long) As Long
Public Property Get BorderStyle() As EBSliderBorderStyle
    BorderStyle = zBorderStyle
End Property
Public Property Let BorderStyle(newValue As EBSliderBorderStyle)
    zBorderStyle = newValue
    UserControl_Paint
End Property
Public Property Get Max() As Long
'[Description]
'   Return the current Max property
'[Code]
    Max = lngMax
End Property
Public Property Let Max(ByVal newValue As Long)
'[Description]
'   Set the current max property
'[Code]
    If newValue > lngMin Then
'Max must be greater than Min
        lngMax = newValue
        If lngValue > lngMax Then
'Ensure current value is within new min-max range
            lngValue = lngMax
            PropertyChanged "Value"
        End If
'Re-initialise slider
        PositionSlider
        PropertyChanged "Max"
    End If
End Property
Public Property Get min() As Long
'[Description]
'   Return the current Min property
'[Code]
    min = lngMin
End Property
Public Property Let min(ByVal newValue As Long)
'[Description]
'   Set the Min property
'[Code]
    If newValue <= lngMax Then
'Min must be less than Max
        lngMin = newValue
        If lngValue < lngMin Then
'ensure current value still in min-max range
            lngValue = lngMin
            PropertyChanged "Value"
        End If
        PositionSlider
        PropertyChanged "Min"
    End If
End Property
Public Property Get Orientation() As EBSliderOrientation
    Orientation = zOrientation
End Property
Public Property Let Orientation(newValue As EBSliderOrientation)
    zOrientation = newValue
    SliderWidth = lngSliderWidth 'force resize or slider
    picSlider_Paint
    UserControl_Resize
End Property
Private Sub picSlider_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)
'[Description]
'   Allow the user to reposition the slider by dragging
'[Declarations]
Dim lngPos   As Long     'New position of slider
Dim sglScale As Single   'Calculated scale of slider
'[Code]
    If Button = vbLeftButton Then
'Only move if the button is pressed
        With picSlider
            If zOrientation = EBHorizontal Then
'calulate new position of slider and round to nearest pixel
                lngPos = ((.Left + x - lngSliderWidth / 2) \ 15) * 15
'Constrain to control
                If lngPos < 0 Then
'Attempted to move slider past start
                    lngPos = 0
                ElseIf lngPos > UserControl.Width - lngSliderWidth Then
'Attempted to move slider past end
                    lngPos = UserControl.Width - lngSliderWidth
                End If
'Move slider
                .Left = lngPos
'Re-calculate value based on new position
                sglScale = (UserControl.Width - lngSliderWidth) / (lngMax - lngMin)
                lngValue = (lngPos / sglScale) + lngMin
                RaiseEvent Changed
            Else
'Vertical
'calulate new position of slider and round to nearest pixel
                lngPos = ((.Top + y - lngSliderWidth / 2) \ 15) * 15
'Constrain to control
                If lngPos < 0 Then
'Attempted to move slider past start
                    lngPos = 0
                ElseIf lngPos > UserControl.Height - lngSliderWidth Then
'Attempted to move slider past end
                    lngPos = UserControl.Height - lngSliderWidth
                End If
'Move slider
                .Top = lngPos
'Re-calculate value based on new position
                sglScale = (UserControl.Height - lngSliderWidth) / (lngMax - lngMin)
                lngValue = (lngPos / sglScale) + lngMin
                RaiseEvent Changed
            End If
        End With
    End If
End Sub
Private Sub picSlider_Paint()
'[Description]
'   Draw a raised border round the slider
'[Declarations]
Dim udtRECT                 As RECT         'Slider RECT structure
'[Code]
    With picSlider
        SetRect udtRECT, 0, 0, .Width / 15, .Height / 15
        DrawEdge .hdc, udtRECT, BDR_RAISEDINNER, BF_RECT
    End With
End Sub
Private Sub picSlider_Resize()
    picSlider.Cls
End Sub
Private Sub PositionSlider()
'[Description]
'   Moves the slider to match the current Value property
'[Declarations]
Dim sglScale                As Single       'Calculated scale of slider
'[Code]
    With picSlider
        If lngMax - lngMin <> 0 Then
'Avoid devide by zero error
'Calculate new position
            If zOrientation = EBHorizontal Then
                sglScale = (UserControl.Width - lngSliderWidth) / (lngMax - lngMin)
                .Left = (lngValue - lngMin) * sglScale
            Else
                sglScale = (UserControl.Height - lngSliderWidth) / (lngMax - lngMin)
                .Top = (lngValue - lngMin) * sglScale
            End If
        End If
    End With
End Sub
Public Property Get SliderColor() As OLE_COLOR
'[Description]
'   Return the current slider color
'[Code]
    SliderColor = picSlider.BackColor
End Property
Public Property Let SliderColor(newValue As OLE_COLOR)
'[Description]
'   Set the slider color
'[Code]
    picSlider.BackColor = newValue
'Redraw the slider
    picSlider_Paint
    PropertyChanged "SliderColor"
End Property
Public Property Get SliderWidth() As Long
'[Description]
'   Reurn current slider width
'[Code]
    SliderWidth = lngSliderWidth
End Property
Public Property Let SliderWidth(ByVal newValue As Long)
'[Description]
'   Set slider width
'[Code]
    If (zOrientation = EBHorizontal And newValue < UserControl.Width) Or (zOrientation = EBVertical And newValue < UserControl.Height) Then
'Ensure slider width is less than control
        lngSliderWidth = newValue
        If zOrientation = EBHorizontal Then
            picSlider.Width = lngSliderWidth
'picSlider.Height = UserControl.Height
        Else
'picSlider.Height = lngSliderWidth
            picSlider.Width = UserControl.Width
        End If
'Redraw the slider
        picSlider_Paint
'Reposition the slider
        PositionSlider
        PropertyChanged "SliderWidth"
    End If
End Property
Private Sub UserControl_InitProperties()
'[Description]
'   Set initial values for properties
'[Code]
    lngMin = 0
    lngMax = 100
    lngValue = 50
    lngSliderWidth = 315
    picSlider.BackColor = vb3DFace
    Orientation = EBHorizontal
    BorderStyle = EBNone
'Initialise the slider
    PositionSlider
End Sub
Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  y As Single)
'[Description]
'   Clicking anywhere on the control makes the slider jump to that position
'[Declarations]
Dim lngPos   As Long     'New position of slider
Dim sglScale As Single   'Calculated scale of slider
    With picSlider
        If zOrientation = EBHorizontal Then
'Caluclate new position and round to nearest pixel
            lngPos = ((x - lngSliderWidth / 2) \ 15) * 15
'Constrain to control
            If lngPos < 0 Then
'Attempted to move past start
                lngPos = 0
            ElseIf lngPos > UserControl.Width - lngSliderWidth Then
'Attempted to move past end
                lngPos = UserControl.Width - lngSliderWidth
            End If
'Move slider
            .Left = lngPos
'Calculate value based on new position
            sglScale = (UserControl.Width - .Width) / (lngMax - lngMin)
            lngValue = (lngPos / sglScale) + lngMin
            RaiseEvent Changed
        Else
'Caluclate new position and round to nearest pixel
            lngPos = ((y - lngSliderWidth / 2) \ 15) * 15
'Constrain to control
            If lngPos < 0 Then
'Attempted to move past start
                lngPos = 0
            ElseIf lngPos > UserControl.Height - lngSliderWidth Then
'Attempted to move past end
                lngPos = UserControl.Height - lngSliderWidth
            End If
'Move slider
            .Top = lngPos
'Calculate value based on new position
            sglScale = (UserControl.Height - lngSliderWidth) / (lngMax - lngMin)
            lngValue = (lngPos / sglScale) + lngMin
            RaiseEvent Changed
        End If
    End With
End Sub
Private Sub UserControl_Paint()
Dim udtRECT                 As RECT
    SetRect udtRECT, 0, 0, UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY
    DrawEdge UserControl.hdc, udtRECT, zBorderStyle, BF_RECT
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'[Description]
'   Retrieve stored properties from PropBag
'[Code]
    With PropBag
        lngMin = .ReadProperty("Min", 0)
        lngMax = .ReadProperty("Max", 100)
        lngValue = .ReadProperty("Value", 50)
        lngSliderWidth = .ReadProperty("SliderWidth", 315)
        picSlider.BackColor = .ReadProperty("SliderColor", vb3DFace)
        BorderStyle = .ReadProperty("BorderStyle", EBNone)
        Orientation = .ReadProperty("Orientation", EBHorizontal)
    End With
'Initialise the slider
    PositionSlider
End Sub
Private Sub UserControl_Resize()
'[Description]
'   Resize constituant controls to match new control size
'[Declarations]
Dim lngWidth  As Long  'New control width
Dim lngHeight As Long  'New control height
Dim intIndex  As Integer
'[Code]
    With UserControl
        .Cls
        lngWidth = .Width - Screen.TwipsPerPixelX
        lngHeight = .Height - Screen.TwipsPerPixelY
        If zOrientation = EBHorizontal Then
'Horizontal
            For intIndex = 0 To 1
                With linGroove(intIndex)
                    .X1 = 15
                    .X2 = lngWidth - 15
                    .Y1 = lngHeight / 2
                    .Y2 = lngHeight / 2
                End With 'linGroove(intIndex)
            Next intIndex
            With picSlider
                .Top = 0
                .Height = lngHeight
                .Width = lngSliderWidth
            End With 'picSlider
        Else
'Vertical
            For intIndex = 0 To 1
                With linGroove(intIndex)
                    .X1 = lngWidth / 2
                    .X2 = lngWidth / 2
                    .Y1 = 15
                    .Y2 = lngHeight - 15
                End With 'linGroove(intIndex)
            Next intIndex
            With picSlider
                .Left = 0
                .Width = lngWidth
                .Height = lngSliderWidth
            End With 'picSlider
        End If
    End With
'Initialise the slider
    PositionSlider
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'[Description]
'   Store properties in PropBag
'[Code]
    With PropBag
        .WriteProperty "Min", lngMin, 0
        .WriteProperty "Max", lngMax, 100
        .WriteProperty "Value", lngValue, 50
        .WriteProperty "SliderWidth", lngSliderWidth, 315
        .WriteProperty "SliderColor", picSlider.BackColor, vb3DFace
        .WriteProperty "BorderStyle", zBorderStyle, EBNone
        .WriteProperty "Orientation", zOrientation, EBHorizontal
    End With
End Sub
Public Property Get Value() As Long
'[Description]
'   Return the current Value property
'[Code]
    If zOrientation = EBHorizontal Then
        Value = lngValue
    Else
        Value = lngMax + lngMin - lngValue
    End If
End Property
Public Property Let Value(newValue As Long)
'[Description]
'   Set the current Value property
'[Code]
'Constrain new value to min-max range
    If newValue < lngMin Then
        newValue = lngMin
    ElseIf newValue > lngMax Then
        newValue = lngMax
    End If
    lngValue = newValue
'Reposition slider
    PositionSlider
    PropertyChanged "Value"
    RaiseEvent Changed
End Property
':)Code Fixer V3.0.9 (9/15/2005 1:30:29 PM) 53 + 451 = 504 Lines Thanks Ulli for inspiration and lots of code.


