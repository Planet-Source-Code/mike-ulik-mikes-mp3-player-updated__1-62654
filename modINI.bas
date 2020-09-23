Attribute VB_Name = "modINI"
Option Explicit
Public KeySection                As String
Public KeyKey                    As String
Public KeyValue                  As String
Public iniLastDirectory          As String
Public iniMCUdirectory           As String
Public iniFilterReverb           As String
Public iniFilterChorus           As String
Public iniFilterDistortion       As String
Public iniFilterEcho             As String
Public iniFilterFlange           As String
Public iniFilterHighpass         As String
Public iniFilterLowpass          As String
Public iniFilterNormalize        As String
Public iniEq0Pos                 As Single
Public iniEq1Pos                 As Single
Public iniEq2Pos                 As Single
Public iniEq3Pos                 As Single
Public iniEq4Pos                 As Single
Public iniEq5Pos                 As Single
Public iniEq6Pos                 As Single
Public iniEq7Pos                 As Single
Public iniEq8Pos                 As Single
Public iniEq9Pos                 As Single
Public iniChorusDryMix           As Single
Public iniChorusWetMix1          As Single
Public iniChorusWetMix2          As Single
Public iniChorusWetMix3          As Single
Public iniChorusDelay            As Single
Public iniChorusRate             As Single
Public iniChorusDepth            As Single
Public iniChorusFeedback         As Single
Public iniDistortionLevel        As Single
Public iniEchoDelay              As Long
Public iniEchoDecayRatio         As Single
Public iniEchoMaxChannels        As Long
Public iniEchoWetMix             As Single
Public iniEchoDryMix             As Single
Public iniFlangeDryMix           As Single
Public iniFlangeWetMix           As Single
Public iniFlangeDepth            As Single
Public iniFlangeRate             As Single
Public iniHighpassCutoff         As Long
Public iniHighpassResonance      As Single
Public iniLowpassCutoff          As Long
Public iniLowpassResonance       As Single
Public iniNormaliseFadeTime      As Long
Public iniNormaliseThreshhold    As Single
Public iniNormaliseMaxAmp        As Long
Public iniReverbRoomSize         As Single
Public iniReverbDamp             As Single
Public iniReverbWetMix           As Single
Public iniReverbDryMix           As Single
Public iniReverbWidth            As Single
Public iniReverbMode             As String
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                                                                     ByVal lpKeyName As Any, _
                                                                                                     ByVal lsString As Any, _
                                                                                                     ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, _
                                                                                                 ByVal lpKeyName As String, _
                                                                                                 ByVal lpDefault As String, _
                                                                                                 ByVal lpReturnedString As String, _
                                                                                                 ByVal nSize As Long, _
                                                                                                 ByVal lpFileName As String) As Long
Public Sub loadINI()
Dim lngResult   As Long
Dim strFileName As String
Dim strResult   As String * 100
Dim Temp        As String
    On Error GoTo ErrorTrap
    strResult = Space$(100)
    strFileName = App.Path & "\Jukebox.INI" 'Declare your ini file !
    lngResult = GetPrivateProfileString(KeySection, KeyKey, strFileName, strResult, Len(strResult), strFileName)
    If lngResult = 0 Then
'An error has occurred
        frmMsgBox.SMessageModal "An error has occurred while calling the INI read!", vbExclamation
    Else
        Temp = Trim$(strResult)
        KeyValue = Left$(Trim$(Temp), Len(Temp) - 1)
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.modINI.loadINI" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Sub saveINI()
Dim lngResult   As Long
Dim strFileName As String
    On Error GoTo ErrorTrap
    strFileName = App.Path & "\Jukebox.INI" 'Declare your ini file !
    lngResult = WritePrivateProfileString(KeySection, KeyKey, KeyValue, strFileName)
    If lngResult = 0 Then
'An error has occurred
        frmMsgBox.SMessageModal "An error has occurred while calling the INI write!", vbExclamation
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.modINI.saveINI" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:24 PM) 59 + 50 = 109 Lines Thanks Ulli for inspiration and lots of code.


