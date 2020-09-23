VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65EF1815-88A2-4A49-8CAB-38BF7D66AA74}#1.0#0"; "BSE.ocx"
Begin VB.Form frmODirectory 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4125
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSE_Engine.BSE ctlJBBSE16 
      Left            =   180
      Top             =   4275
      _ExtentX        =   6588
      _ExtentY        =   1085
   End
   Begin VB.TextBox txtJBODirectoryText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   382
      TabIndex        =   2
      Top             =   2730
      Width           =   4410
   End
   Begin MidiJukeBox.isButton cmdJBOCancel 
      Height          =   555
      Left            =   3465
      TabIndex        =   0
      Top             =   3360
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   979
      Icon            =   "frmSplash.frx":000C
      Style           =   7
      Caption         =   "Cancel"
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   65535
      HighlightColor  =   255
      FontHighlightColor=   255
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin MidiJukeBox.isButton cmdJBODirectory 
      Height          =   555
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   979
      Icon            =   "frmSplash.frx":0028
      Style           =   7
      Caption         =   "Open"
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   65535
      HighlightColor  =   255
      FontHighlightColor=   255
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin MSComctlLib.TreeView ZxDirTreeview 
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   4683
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   647
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   1905
      TabIndex        =   4
      Top             =   3885
      Width           =   1305
   End
End
Attribute VB_Name = "frmODirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdJBOCancel_Click()
    On Error GoTo ErrorTrap
    blnCancel = True
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmODirectory.cmdJBOCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdJBODirectory_Click()
    On Error GoTo ErrorTrap
    If LenB(txtJBODirectoryText.Text) = 0 Then
        frmMsgBox.SMessageModal "No directory selected!", 16, "Error"
        GoTo SubExit9
    End If
    If ctlJBBSE16.EngineStarted Then
        ctlJBBSE16.EndSubClassing
    End If
    iniLastDirectory = txtJBODirectoryText.Text
'save Last Directory
    KeySection = "Directories"
    KeyKey = "Last"
    KeyValue = iniLastDirectory
    saveINI
    Unload Me
SubExit9:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmODirectory.cmdODirectory_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
    ctlJBBSE16.InitSubClassing
    ZxDirFillTree
'load Last Directory
    KeySection = "Directories"
    KeyKey = "Last"
    loadINI
    iniLastDirectory = KeyValue
    ZxDirTreeView_OpenFolder iniLastDirectory
    txtJBODirectoryText.Text = iniLastDirectory
'Me.show
'ZxDirTreeview.SetFocus
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmODirectory.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorTrap
    If ctlJBBSE16.EngineStarted Then
        ctlJBBSE16.EndSubClassing
    End If
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "MidiDateBase.frmODirectory.Form_Unload" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub txtJBODirectoryText_Change()
    On Error GoTo ErrorTrap
    cmdJBODirectory.Enabled = True
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmODirectory.txtJBODirectoryText_Change" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
'============================================================
' Routine Name: ZxDirAddDummyChild
' Description: Add a dummy node to the treeview node to
'              indicate there are some sub directories.
' Author: Mike
' Date: 8/11/2005 6:55:37 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Sub ZxDirAddDummyChild(TVnode As Node)
    On Error GoTo ErrorTrap
' Add a dummy child Node if necessary.
    If TVnode.Children = 0 Then
        ZxDirTreeview.Nodes.Add TVnode.index, tvwChild, , "   "
    End If
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmODirectory.ZxDirAddDummyChild" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
'============================================================
' Routine Name: ZxDirAddSubdirs
' Description: Add the subdirectories to the treeview node.
' Author: Mike
' Date: 8/11/2005 6:55:37 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Sub ZxDirAddSubdirs(ByVal TVnode As Node)
    On Error GoTo ErrorTrap
' Add all the subdirectories under a Node.
    For Each FSOfolder In FSO.GetFolder(TVnode.Key).SubFolders
        If FSOfolder.Attributes = Hidden Or FSOfolder.Attributes = system Or FSOfolder.Attributes = Volume Or FSOfolder.Attributes = 22 Then
            DoEvents
        Else
            Set FSOnode = ZxDirTreeview.Nodes.Add(TVnode, tvwChild, FSOfolder.Path, FSOfolder.Name, 6)
            FSOnode.ExpandedImage = 7
' If this directory has subfolders, add a plus sign.
            If FSOfolder.SubFolders.Count > 0 Then
                ZxDirAddDummyChild FSOnode
            End If
        End If
    Next FSOfolder
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmODirectory.ZxDirAddSubdirs" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
'============================================================
' Routine Name: ZxDirFillTree
' Description: Fill the treeview with drives
' Author: Mike
' Date: 8/11/2005 6:55:37 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Public Sub ZxDirFillTree()
    
    On Error Resume Next
    frmLoading.Show
    DoEvents
    ' Add the "My Computer" root Node (expanded).
    Set FSOrootNode = ZxDirTreeview.Nodes.Add(, , "\\MyComputer", "My Computer", 1)
    FSOrootNode.Expanded = True
    ' Add all the drives; display a plus sign beside them.
    For Each FSOdrive In FSO.Drives
        If FSOdrive.DriveType = Removable Then
            FSOIcon = 1
            GoTo DriveFound
        End If
        If FSOdrive.DriveType = Fixed Then
            FSOIcon = 2
            GoTo DriveFound
        End If
        If FSOdrive.DriveType = CDRom Then
            FSOIcon = 3
            GoTo DriveFound
        End If
        If FSOdrive.DriveType = RamDisk Then
            FSOIcon = 4
            GoTo DriveFound
        End If
        If FSOdrive.DriveType = UnknownType Then
            FSOIcon = 8
        End If
DriveFound:
        Set FSOnode = ZxDirTreeview.Nodes.Add(FSOrootNode.Key, tvwChild, FSOdrive.Path & "\", FSOdrive.Path, FSOIcon)
        If FSOdrive.Path = "A:" Or FSOdrive.Path = "B:" Then
            DoEvents
        Else
            If FSOdrive.IsReady Then
                ZxDirAddDummyChild FSOnode
            End If
        End If
    Next FSOdrive
FillTreeExit:
    frmLoading.Hide
    DoEvents
    Exit Sub
    
FillTreeError:
    MsgBox "Error Number: " & Err.Number & vbNewLine & _
        Err.Description & vbNewLine & _
        vbNewLine & _
        "Debug Information:" & vbNewLine & _
        "FillTree" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
    On Error GoTo 0
    
End Sub
'============================================================
' Routine Name: ZxDirListview_Click
' Description:
' Author: Mike
' Date: 8/12/2005 7:53:32 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Sub ZxDirListview_Click()
    On Error GoTo ZxDirListview_ClickError
ZxDirListview_ClickExit:
Exit Sub
ZxDirListview_ClickError:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "ZxDirListview_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
'============================================================
' Routine Name: ZxDirTreeView1_Expand
' Description: Treeview node expansion
' Author: Mike
' Date: 8/11/2005 6:55:37 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Sub ZxDirTreeview_Expand(ByVal TVnode As Node)
Dim Temp       As String
Dim Length     As Long
Dim start      As Long
Dim drive      As String
Dim DriveReady As Boolean
    On Error GoTo ErrorTrap
' A Node is being expanded.
' Exit if the Node had already been expanded or has no children.
'frmLoading.show
    DoEvents
    Temp = TVnode.FullPath
    Length = Len(Temp)
    start = InStr(Temp, ":")
    FSOpath = Right$(Temp, Length - (start - 2))
    drive = FSO.GetDriveName(FSOpath)
    DriveReady = FSO.GetDrive(drive).IsReady
    If Not DriveReady Then
        frmMsgBox.SMessageModal "Drive not ready!"
        GoTo ZxDirTreeview_ExpandExit
    End If
    If FSO.GetDrive(drive).IsReady Then
        If Not TVnode.Children = 0 Or TVnode.Children > 1 Then
' Also exit if it doesn't have a dummy child Node.
            If TVnode.Child.Text <> "   " Then
                GoTo ZxDirTreeview_ExpandExit
            End If
' Remove the dummy child item.
            ZxDirTreeview.Nodes.Remove TVnode.Child.index
' Add all the subdirs of this Node object.
            ZxDirAddSubdirs TVnode
            GoTo ZxDirTreeview_ExpandExit
        End If
    End If
    frmMsgBox.SMessageModal "Drive not ready"
ZxDirTreeview_ExpandExit:
'frmLoading.Hide
    DoEvents
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmODirectory.ZxDirTreeview_Expand" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
'============================================================
' Routine Name: ZxDirTreeView_NodeClick
' Description: Treeview node click.
' Author: Mike
' Date: 8/11/2005 6:55:37 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Sub ZxDirTreeview_NodeClick(ByVal TVnode As MSComctlLib.Node)
Dim Temp       As String
Dim Length     As Long
Dim start      As Long
Dim DriveReady As Boolean
Dim drive      As String
    On Error GoTo ErrorTrap
'frmLoading.show
    DoEvents
    Temp = TVnode.FullPath
    Length = Len(Temp)
    start = InStr(Temp, ":")
    FSOpath = Right$(Temp, Length - (start - 2))
    Zx_Tkey = FSOpath & "\"
    drive = FSO.GetDriveName(Zx_Tkey)
    DriveReady = FSO.GetDrive(drive).IsReady
    If Not DriveReady Then
        frmMsgBox.SMessageModal "Drive not ready!", vbCritical
        GoTo ZxDirTreeview_NodeClickExit
    End If
ZxDirTreeview_NodeClickExit:
'frmLoading.Hide
    txtJBODirectoryText.Text = Right$(TVnode.FullPath, Len(TVnode.FullPath) - 12) & "\"
    iniLastDirectory = txtJBODirectoryText.Text
    DoEvents
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmODirectory.ZxDirTreeview_NodeClick" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
'============================================================
' Routine Name: ZxDirTreeView_OpenFolder
' Description:
' Author: Mike
' Date: 8/21/2005 9:29:19 AM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Sub ZxDirTreeView_OpenFolder(ByVal FolderToOpen As String)
Dim y              As Long
Dim x              As Long
Dim Temp           As String
Dim drive          As String
Dim DriveReady     As Boolean
Dim FolderToFind() As String
    On Error GoTo ZxDirTreeView_OpenFolderError
'frmLoading.show
    DoEvents
    drive = FSO.GetDriveName(FolderToOpen) & "\"
    DriveReady = FSO.GetDrive(drive).IsReady
    If Not DriveReady Then
        frmMsgBox.SMessageModal "Drive not ready!", vbCritical
        GoTo ZxDirTreeView_OpenFolderExit
    End If
    For x = 1 To ZxDirTreeview.Nodes.Count
        If ZxDirTreeview.Nodes(x).Key = drive Then
            ZxDirTreeview.Nodes(x).EnsureVisible
            ZxDirTreeview.Nodes(x).Selected = True
            Exit For
        End If
    Next x
    Zx_Tkey = FolderToOpen
    DoEvents
    FolderToFind = Split(FolderToOpen, "\")
    If UBound(FolderToFind) - 1 > 2 Then
        For x = 2 To UBound(FolderToFind) - 1
            If x + 1 <= UBound(FolderToFind) Then
                Temp = FolderToFind(x - 1) & "\" & FolderToFind(x)
                FolderToFind(x) = Temp
            End If
        Next x
    End If
    For x = 1 To UBound(FolderToFind) - 1
        Debug.Print FolderToFind(x)
    Next x
    For x = 1 To UBound(FolderToFind) - 1
        For y = 1 To ZxDirTreeview.Nodes.Count
            If Len(ZxDirTreeview.Nodes(y).Key) > Len(FolderToFind(x)) Then
                Temp = Right$(ZxDirTreeview.Nodes(y).Key, Len(FolderToFind(x)))
            Else
                Temp = ZxDirTreeview.Nodes(y).Key
            End If
            If LCase$(Temp) = LCase$(FolderToFind(x)) Then
                ZxDirTreeview.Nodes(y).EnsureVisible
                ZxDirTreeview.Nodes(y).Selected = True
                Exit For
            End If
        Next y
        DoEvents
    Next x
    DoEvents
'ZxDirLoadListView FolderToOpen
ZxDirTreeView_OpenFolderExit:
'frmLoading.Hide
    DoEvents
Exit Sub
ZxDirTreeView_OpenFolderError:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.Description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "ZxDirTreeView_OpenFolder" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
'============================================================
' Routine Name: ZxDirTreeViewFolderExists
' Description:
' Author: Mike
' Date: 8/12/2005 7:36:01 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Function ZxDirTreeViewFolderExists(strFolderName As String) As Boolean
    On Error Resume Next
    ZxDirTreeViewFolderExists = FSO.FolderExists(strFolderName)
    On Error GoTo 0
End Function




