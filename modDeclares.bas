Attribute VB_Name = "modDeclares"
Option Explicit
Public FSO                             As New Scripting.FileSystemObject
Public FSOdrive                        As Scripting.drive
Public FSOfolder                       As Scripting.Folder
Public FSOfilename                     As Scripting.File
Public FSOrootNode                     As Node
Public FSOnode                         As Node
Public FSOIcon                         As Long
Public FSOpath                         As String
Public FSOpattern                      As String
Public Zx_Tkey                         As String
Public Zx_Pattern                      As String
Public FSOfilecount                    As Long
Public FSOFileNameOutput               As String
Public Const OnaThousand               As Long = 1000
Public ChannelPlaying                  As Long
Public State                           As Long
Public SoundPlaying                    As Long
Public Eq0                             As Long
Public Eq1                             As Long
Public Eq2                             As Long
Public Eq3                             As Long
Public Eq4                             As Long
Public Eq5                             As Long
Public Eq6                             As Long
Public Eq7                             As Long
Public Eq8                             As Long
Public Eq9                             As Long
Public blnCancel                       As Boolean
Public playX                           As Long
Public blnPlayAll                      As Boolean
Public adoRsSong                       As ADODB.Recordset
Public blnProceed                      As Boolean
Public OldKey                          As String
Public blnPlayOpen                     As Boolean
Public blnAllStop                      As Boolean
Public blnPrevious                     As Boolean
Public blnNext                         As Boolean
Public blnStop                         As Boolean
Public lngPCount                       As Long
Public lngNCount                       As Long
Public lngRandom                       As Long
Public lngMin                          As Long
Public SortKey                         As Long    'Retains the value of the last SortKey
Private CC                             As Long
Public SongDir                         As String
Public DragX                           As Long
Public DragY                           As Long
Public channel                         As Long
Public Sound                           As Long
Public system                          As Long
Public result                          As FMOD_RESULT
Public Version                         As Long
Public exinfo                          As FMOD_CREATESOUNDEXINFO
Public active                          As Long
Public position                        As Long
Public DspReverbFilter                 As Long
Public DspChorusFilter                 As Long
Public DspDistortionFilter             As Long
Public DspEchoFilter                   As Long
Public DspFlangeFilter                 As Long
Public DspHighpassFilter               As Long
Public DspLowpassFilter                As Long
Public DspNormalizeFilter              As Long
Public Savefilename                    As String
Public RecID()                         As String
Public ListID()                        As String
Private Const LVM_FIRST                As Long = &H1000
Private Const LVM_GETNEXTITEM          As Double = (LVM_FIRST + 12)
Private Const LVNI_SELECTED            As Long = &H2
Private Const LVM_GETSELECTEDCOUNT     As Double = (LVM_FIRST + 50)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cX As Long, _
                                                    ByVal cY As Long, _
                                                    ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Function CountSelectedItemsInListview(oListview As ListView) As Long
    On Error GoTo ErrorTrap
    CountSelectedItemsInListview = SendMessage(oListview.hwnd, LVM_GETSELECTEDCOUNT, 0, 0)
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.Module1.CountSelectedItemsInListview" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Sub ERRCHECK(result As FMOD_RESULT)
Dim msgResult As VbMsgBoxResult
    On Error GoTo ErrorTrap
    If result <> FMOD_OK Then
        msgResult = MsgBox("FMOD error! (" & result & ") " & FMOD_ErrorString(result))
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.modDeclares.ERRCHECK" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Public Function FileExists1(ByVal sFileName As String) As Boolean
Dim sFile As String
    On Error GoTo ErrorTrap
    On Error Resume Next
    FileExists1 = False
    sFile = Dir(sFileName)
    If Len(sFile) > 0 Then
        If Err.Number = 0 Then
            FileExists1 = True
        End If
    End If
    On Error GoTo 0
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.modDeclares.FileExists1" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Function FindFile(ByVal sFol As String, _
                         ByVal sFile As String, _
                         nDirs As Long, _
                         ByVal nFiles As Long) As Currency
Dim strFileName As String
Dim fld         As Folder
Dim itmX        As listitem
Dim N           As String
    On Error GoTo Catch
    Set fld = FSO.GetFolder(sFol)
    frmMP3Jukebox.lsvListView1.ListItems.Clear
    strFileName = Dir(FSO.BuildPath(fld.Path, sFile), vbNormal Or vbReadOnly)
    Do While Len(strFileName) <> 0
        If Zx_Pattern <> "*.*" Then
            N = LCase$(Right$(strFileName, 3))
            If InStr(Zx_Pattern, N) > 0 Then
                Set itmX = frmMP3Jukebox.lsvListView1.ListItems.Add(, , strFileName)
            End If
        Else
            Set itmX = frmMP3Jukebox.lsvListView1.ListItems.Add(, , strFileName)
        End If
        strFileName = Dir()
    Loop
    nDirs = nDirs + 1
Exit Function
Catch:
    strFileName = vbNullString
    Resume Next
End Function
Public Function GetSelectedItemsFromListview(oListview As ListView) As Collection
Dim lCurSelectedItemIndex As Long
Dim MyCol                 As Collection
Dim I                     As Long
    On Error GoTo ErrorTrap
'begin/start position in the listview
    lCurSelectedItemIndex = -1
'create collection to hold selected items
    Set MyCol = New Collection
    CC = CountSelectedItemsInListview(oListview)
    ReDim RecID(CC) As String
    ReDim ListID(CC) As String
    For I = 1 To CountSelectedItemsInListview(oListview)
'get the itemx index from the selected (current)item
        With oListview
            lCurSelectedItemIndex = SendMessage(.hwnd, LVM_GETNEXTITEM, lCurSelectedItemIndex, ByVal LVNI_SELECTED)
'add the listitem to the collection
            MyCol.Add .ListItems.Item(lCurSelectedItemIndex + 1)
            RecID(I) = .ListItems.Item(lCurSelectedItemIndex + 1).SubItems(13)
            ListID(I) = .ListItems.Item(lCurSelectedItemIndex + 1).index
        End With 'oListview
    Next I
'return the collection
    Set GetSelectedItemsFromListview = MyCol
Exit Function
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.Module1.GetSelectedItemsFromListview" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Function
Public Sub KeepOnTop(F As Form)
Const SWP_NOMOVE   As Integer = 2
Const SWP_NOSIZE   As Integer = 1
Const HWND_TOPMOST As Integer = -1
    On Error GoTo ErrorTrap
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.modDeclares.KeepOnTop" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
':)Code Fixer V3.0.9 (9/15/2005 1:31:22 PM) 81 + 135 = 216 Lines Thanks Ulli for inspiration and lots of code.


