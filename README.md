<div align="center">

## A File system object example FSO


</div>

### Description

Example of the File System Object
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[andreas gustafsson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andreas-gustafsson.md)
**Level**          |Intermediate
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andreas-gustafsson-a-file-system-object-example-fso__1-11838/archive/master.zip)





### Source Code

```
'*************************************************
'* This program was created by andreas
'*gustafsson.
'* Please do not change/remove this
'*text      '*
'* Feel free to edit the code as you
'*wish
'* send comments to
'*andreasgustafsson1@hotmail.com
'* References: Microsoft scripting
'*runtime
'************************************************* Option Explicit
 Dim fso As New FileSystemObject
 'The selected drive
 Dim strDrive As String
 'The folderpath
 Dim strFolder As String
 'Collection to store the selected filepaths
Private Sub cmbDrives_Click()
 Dim drive As drive
 Dim File As File
 Dim SubFolder As Folder
 Dim i As Integer
 i = 0
 lstFiles.Clear
 If cmbDrives = "" Then Exit Sub
 strDrive = cmbDrives.Text
 strFolder = ""
 Set drive = fso.GetDrive(cmbDrives.Text)
 If drive.IsReady Then
 For Each File In drive.RootFolder.Files
  lstFiles.AddItem File.Name, i
  i = i + 1
 Next
 i = lstFiles.ListCount
 For Each SubFolder In _ drive.RootFolder.SubFolders
 lstFiles.AddItem SubFolder, i
 i = i + 1
 Next
 Else
 MsgBox "Drives not ready"
 End If
End Sub
'Moves to the parent folder (if any)
Private Sub cmdup_Click()
 Dim Folder As Folder
 Dim File As File
 Dim SubFolder As Folder
 Dim i As Integer
 If strDrive = "" Then Exit Sub
 If strFolder = "" Then Exit Sub
 'Get current folder
 Set Folder = fso.GetFolder(strDrive & _ strFolder)
 'Find parent folder
 strFolder = Left(strFolder, InStrRev _(strFolder, "\") - 1)
 lstFiles.Clear
 'If parent exists
 If Not Folder.ParentFolder Is Nothing Then
 'Add all files in parent
 For Each File In Folder.ParentFolder.Files
  lstFiles.AddItem File.Name, i
  i = i + 1
 Next
 i = lstFiles.ListCount
 'Add all subfolders in parent
 For Each SubFolder In _ Folder.ParentFolder.SubFolders
  lstFiles.AddItem SubFolder, i
  i = i + 1
 Next
 Else 'If it not has parent
 For Each File In Folder.Files
  lstFiles.AddItem File.Name, i
  i = i + 1
 Next
 i = lstFiles.ListCount
 For Each SubFolder In Folder.SubFolders
  lstFiles.AddItem SubFolder, i
  i = i + 1
 Next
 End If
End Sub
Private Sub Form_Load()
 Dim drive As drive
 Dim i As Integer
 i = 0
 'Add all drives to combo
 For Each drive In fso.Drives
 cmbDrives.AddItem drive.Path, i
 i = i + 1
 Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Set fso = Nothing
End Sub
Private Sub lstFiles_Click()
 Dim Folder As Folder
 Dim SubFolder As Folder
 Dim File As File
 Dim i As Integer
 i = 0
 If Not lstFiles.SelCount > 1 Then
 'if its a folder
 If InStr(lstFiles.Text, ":\") Then
  Set Folder = fso.GetFolder _(lstFiles.Text)
  lstFiles.Clear
  strFolder = strFolder & "\" & _ Folder.Name
  'Add all files
  For Each File In Folder.Files
  lstFiles.AddItem File.Name, i
  i = i + 1
  Next
  i = lstFiles.ListCount
  'Add subfolders
  For Each SubFolder In _ Folder.SubFolders
  lstFiles.AddItem SubFolder, i
  i = i + 1
  Next
 End If
 End If
End Sub
```

