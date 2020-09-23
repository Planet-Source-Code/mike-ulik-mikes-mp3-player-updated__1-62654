VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSave 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3270
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSave.frx":57E2
   ScaleHeight     =   5190
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MP3JukeBox.isButton cmdCancel 
      Height          =   600
      Left            =   2160
      TabIndex        =   4
      Top             =   4410
      Width           =   750
      _extentx        =   1323
      _extenty        =   1058
      style           =   7
      caption         =   "Cancel"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmSave.frx":66AC
   End
   Begin MP3JukeBox.isButton cmdJBSsave 
      Height          =   600
      Left            =   300
      TabIndex        =   3
      Top             =   4410
      Width           =   750
      _extentx        =   1323
      _extenty        =   1058
      style           =   7
      caption         =   "Save"
      usecustomcolors =   -1  'True
      backcolor       =   65535
      highlightcolor  =   255
      fonthighlightcolor=   255
      tooltiptitle    =   ""
      tooltipicon     =   0
      tooltiptype     =   1
      font            =   "frmSave.frx":66D0
   End
   Begin VB.TextBox txtJBSName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   0
      TabIndex        =   0
      Text            =   "*.mcu"
      Top             =   3870
      Width           =   3300
   End
   Begin MSComctlLib.ImageList ZxDirImageList 
      Left            =   195
      Top             =   3735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":66F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":8EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":91C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":B972
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":BC8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":BFA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":C2C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSave.frx":C41A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ZxDirListview 
      Height          =   1980
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   3493
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Names"
         Object.Width           =   11360
      EndProperty
   End
   Begin MSComctlLib.TreeView ZxDirTreeview 
      Height          =   1800
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   3175
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   647
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ZxDirImageList"
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
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    On Error GoTo ErrorTrap
    blnCancel = True
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmSave.cmdCancel_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub cmdJBSsave_Click()
Dim tname As String
    On Error GoTo ErrorTrap
    If LenB(txtJBSName.Text) = 0 Or txtJBSName.Text = "*.mcu" Then
        frmMsgBox.SMessageModal "Nothing to save!", 48, "Save Error"
        GoTo cmdSaveExit
    End If
    With txtJBSName
        If InStr(.Text, ".mcu") = 0 Then
            tname = .Text & ".mcu"
            .Text = tname
        End If
    End With
'save MCU Directory
    KeySection = "Directories"
    KeyKey = "MCUdirectory"
    KeyValue = Zx_Tkey
    saveINI
    Savefilename = iniMCUdirectory & txtJBSName.Text
    Unload Me
cmdSaveExit:
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmSave.cmdJBSsave_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorTrap
     
    FSOpattern = "*.mcu"
    ZxDirFillTree
    ZxDirTreeView_OpenFolder iniMCUdirectory
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmSave.Form_Load" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorTrap
    
    Unload Me
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmSave.Form_Unload" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
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
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmSave.ZxDirAddDummyChild" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
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
            Set FSOnode = ZxDirTreeview.Nodes.Add(TVnode, tvwChild, FSOfolder.Path, FSOfolder.name, 6)
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
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmSave.ZxDirAddSubdirs" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
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
'frmLoading.show
    DoEvents
' Add the "My Computer" root Node (expanded).
    Set FSOrootNode = ZxDirTreeview.Nodes.Add(, , "\\MyComputer", "My Computer", 1)
    FSOrootNode.Expanded = True
' Add all the drives; display a plus sign beside them.
    For Each FSOdrive In FSO.Drives
        If FSOdrive.DriveType = Removable Then
            FSOIcon = 1
        End If
        If FSOdrive.DriveType = Fixed Then
            FSOIcon = 2
        End If
        If FSOdrive.DriveType = CDRom Then
            FSOIcon = 3
        End If
        If FSOdrive.DriveType = RamDisk Then
            FSOIcon = 4
        End If
        If FSOdrive.DriveType = UnknownType Then
            FSOIcon = 8
        End If
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
'frmLoading.Hide
    DoEvents
Exit Sub
FillTreeError:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
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
    txtJBSName.Text = ZxDirListview.SelectedItem.Text
ZxDirListview_ClickExit:
Exit Sub
ZxDirListview_ClickError:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "ZxDirListview_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
End Sub
'============================================================
' Routine Name: ZxDirLoadListView
' Description:
' Author: Mike
' Date: 8/12/2005 5:42:48 PM
' Copyright © 2005
' Notes:
' Modification History:
'============================================================
Private Sub ZxDirLoadListView(ByVal Tkey As String)
Dim strTemp As String
Dim itmX    As listitem
    On Error GoTo LoadListViewError
    ZxDirListview.ListItems.Clear
    For Each FSOfilename In FSO.GetFolder(Tkey).Files
        strTemp = LCase$(Right$(FSOfilename.name, 4))
        If FSOpattern = "*.*" Then
            Set itmX = ZxDirListview.ListItems.Add(, , FSOfilename.name)
        Else
            If InStr(FSOpattern, strTemp) > 0 Then
                Set itmX = ZxDirListview.ListItems.Add(, , FSOfilename.name)
            End If
        End If
    Next FSOfilename
    FSOfilecount = ZxDirListview.ListItems.Count
LoadListViewExit:
Exit Sub
LoadListViewError:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "LoadListView" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
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
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmSave.ZxDirTreeview_Expand" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
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
    If ZxDirTreeViewFolderExists(Zx_Tkey) Then
        ZxDirLoadListView Zx_Tkey
    End If
ZxDirTreeview_NodeClickExit:
'frmLoading.Hide
    DoEvents
Exit Sub
ErrorTrap:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
          vbNewLine & _
          "Debug Information:" & vbNewLine & _
          "JukeBox.frmSave.ZxDirTreeview_NodeClick" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"
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
Private Sub ZxDirTreeView_OpenFolder(FolderToOpen As String)
Dim y              As Long
Dim x              As Long
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
    For x = 1 To UBound(FolderToFind) - 1
        For y = 1 To ZxDirTreeview.Nodes.Count
            If InStr(LCase$(ZxDirTreeview.Nodes(y).Key), LCase$(FolderToFind(x))) > 0 Then
                ZxDirTreeview.Nodes(y).EnsureVisible
                ZxDirTreeview.Nodes(y).Selected = True
                Exit For
            End If
        Next y
        DoEvents
    Next x
    DoEvents
    ZxDirLoadListView FolderToOpen
ZxDirTreeView_OpenFolderExit:
'frmLoading.Hide
    DoEvents
Exit Sub
ZxDirTreeView_OpenFolderError:
    frmMsgBox.SMessageModal "Error Number: " & Err.Number & vbNewLine & _
          Err.description & vbNewLine & _
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
':)Code Fixer V3.0.9 (9/15/2005 1:30:58 PM) 1 + 426 = 427 Lines Thanks Ulli for inspiration and lots of code.


