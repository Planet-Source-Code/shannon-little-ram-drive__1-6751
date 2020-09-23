Attribute VB_Name = "modTreeView"
Option Explicit


'This file was created 2/22/00
'by Shannon Little
'http://go.to/neotrix
'This little module will handle adding and showing drives/files/folders
'in a treeView control
'All you need to do is to call InitializeTreeView
'Then call tvNodeClick everytime the user clicks on the tree
'You have to pass the TreeView object and Node on every call though
'And you have to add an image list with images in this order

'This requires that you have CommonControls-2 or -3, whichever has the
'treeview control in it, or else errors occur because it doesn't know what
'type a treeview control is

'Unknown
'Removable
'Fixed
'Network
'CD-ROM
'RAM Disk
'Open folder
'Closed folder
'File

Enum DriveType
    dtUnknown = 0
    dtRemovable = 1
    dtFixed = 2
    dtNetwork = 3
    dtDisconnectedNetwork = 4
    dtCDROM = 5
    dtRamDrive = 6
    dtAllDriveTypes = 7
    dtFloppy = 8
End Enum

'Created 2/22/00
'Add drives to a TreeView control
'And a dummy node if the drive has any sub-folders
Public Sub InitializeTreeView(ByVal TreeView As TreeView, Optional bExpandToFiles As Boolean)
    On Error GoTo lError
    Dim xNode As Node, strDriveType As String
    Dim fso, drv, cDrv, tDrv, fld, fil
    'FileSystem Object, Drive, Drive Collection, Temporary Drive Object, Folder
    Dim strOutput As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set cDrv = fso.Drives
    
    If IsMissing(bExpandToFiles) Then bExpandToFiles = False

    TreeView.Nodes.Clear
       
    For Each tDrv In fso.Drives
        Select Case tDrv.DriveType
            Case 0: strDriveType = "Fixed"   'Unknown
            Case 1:
                'Test to see if its a floppy or some other removable type drive
                If tDrv.DriveLetter = "A" Or tDrv.DriveLetter = "B" Then
                    strDriveType = "Floppy"
                Else
                    strDriveType = "Removable"
                End If
            Case 2: strDriveType = "Fixed"
            Case 3:
                If tDrv.IsReady Then
                    strDriveType = "Network"
                Else
                    strDriveType = "DisconnectedNetwork"
                End If
            Case 4: strDriveType = "CD-ROM"
            Case 5: strDriveType = "Ram Disk"
            Case Default:
                strDriveType = "Fixed" 'In case any new drive types are added later
        End Select
        
        strOutput = ""
        'strOutput = FormatProperName(tDrv.VolumeName) & " (" & tDrv.Path & ")"
        strOutput = tDrv.Path & "\"

        Set xNode = TreeView.Nodes.Add(, , tDrv.Path & "\", strOutput, strDriveType)
        xNode.Sorted = True
        
        ' Create a Dummy Node so we get a + beside each drive
        'if it has any sub-folders
        Set fld = fso.GetFolder(tDrv.Path & "\")

        If fld.SubFolders.Count > 0 Then
            Set xNode = TreeView.Nodes.Add(tDrv.Path & "\", tvwChild)
        End If
        
        If bExpandToFiles Then
            If fld.Files.Count > 0 Then
                'Don't create a dummy node if one has already been created about
                If fld.SubFolders.Count = 0 Then
                    TreeView.Nodes.Add xNode, tvwChild
                End If
            End If
        End If
jumpBackIn:
    Next
    
    Exit Sub
lError:
    Select Case Err.Number
        Case 76: 'Floppy Not ready
                GoTo jumpBackIn
    End Select

End Sub

'Created 2/23/00
'Adds just drives to a TreeView
Public Sub InitializeCertainDrives(ByVal TreeView As TreeView, DrvType As DriveType, ByVal DoAddSubFolders As Boolean, Optional bExpandToFiles As Boolean)
    On Error Resume Next
    
    Dim xNode As Node, strDriveType As String, bAddThisDrive As Boolean
    Dim fso, drv, tDrv, fld
    'FileSystem Object, Drive, Temporary Drive Object, Folder
    Dim strOutput As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    If IsMissing(bExpandToFiles) Then bExpandToFiles = False

    TreeView.Nodes.Clear
       
    'DriveType Enumerations
    'All drives
    'Unknown
    'Removable
    'Fixed
    'Network
    'DisconnectedNetwork
    'CDROM
    'RamDisk
    
    For Each tDrv In fso.Drives
        bAddThisDrive = False
        
        Select Case tDrv.DriveType
            Case 0:
                strDriveType = "Fixed" 'Unknown
                If DrvType = dtUnknown Then bAddThisDrive = True
                If DrvType = dtAllDriveTypes Then bAddThisDrive = True
            Case 1:
                strDriveType = "Removable"
                If tDrv.DriveLetter = "A" Or tDrv.DriveLetter = "B" Then
                    strDriveType = "Floppy"
                    If DrvType = dtFloppy Then bAddThisDrive = True
                Else
                    strDriveType = "Removable"
                    If DrvType = dtRemovable Then bAddThisDrive = True
                End If
                If DrvType = dtAllDriveTypes Then bAddThisDrive = True
            Case 2:
                strDriveType = "Fixed"
                If DrvType = dtFixed Then bAddThisDrive = True
                If DrvType = dtAllDriveTypes Then bAddThisDrive = True
            Case 3:
                If tDrv.IsReady Then
                    strDriveType = "Network"
                    'Network
                    If DrvType = dtNetwork Then bAddThisDrive = True
                    If DrvType = dtAllDriveTypes Then bAddThisDrive = True
                Else
                    strDriveType = "DisconnectedNetwork"
                    'Disconnected Network
                    If DrvType = dtDisconnectedNetwork Then bAddThisDrive = True
                    If DrvType = dtAllDriveTypes Then bAddThisDrive = True
                End If
            Case 4:
                strDriveType = "CD-ROM"
                If DrvType = dtCDROM Then bAddThisDrive = True
                If DrvType = dtAllDriveTypes Then bAddThisDrive = True
            Case 5:
                strDriveType = "Ram Disk"
                If DrvType = dtRamDrive Then bAddThisDrive = True
                If DrvType = dtAllDriveTypes Then bAddThisDrive = True
            Case Default:
                strDriveType = "Fixed" 'In case any new drive types are added later
                If DrvType = dtUnknown Then bAddThisDrive = True
                If DrvType = dtAllDriveTypes Then bAddThisDrive = True
        End Select
        
        
        If bAddThisDrive Then
            strOutput = ""
            'strOutput = FormatProperName(tDrv.VolumeName) & " (" & tDrv.Path & ")"
            strOutput = tDrv.Path & "\"

            Set xNode = TreeView.Nodes.Add(, , tDrv.Path & "\", strOutput, strDriveType)
            xNode.Sorted = True

            
            If DoAddSubFolders Then
                ' Create a Dummy Node so we get a + beside each drive
                'if it has any sub-folders
                Set fld = fso.GetFolder(tDrv.Path & "\")
                
                If fld.SubFolders.Count > 0 Then
                    Set xNode = TreeView.Nodes.Add(tDrv.Path & "\", tvwChild)
                End If
            End If
            
            If bExpandToFiles Then
                If fld.Files.Count > 0 Then
                    'Don't create a dummy node if one has already been created about
                    If fld.SubFolders.Count = 0 Then
                        TreeView.Nodes.Add xNode, tvwChild
                    End If
                End If
            End If
            
        End If
        
    Next
    
End Sub

'Created 2/22/00
'Removes all sub nodes on the node clicked, then reads all the subfolders
Public Sub tvNodeClick(ByVal TreeView As TreeView, ByVal Node As MSComctlLib.Node, Optional bExpandToFiles As Boolean)
    On Error Resume Next
    Dim xNode As Node
    Dim fso, fld, tFld, tFil
    'FileSystem Object, Folder, Temporary Folder Object, Temporary Files
    Set fso = CreateObject("Scripting.FileSystemObject")
        
    Set fld = fso.GetFolder(Node.FullPath)
    
    If IsMissing(bExpandToFiles) Then bExpandToFiles = False

    'Remove all sub-nodes
    
    Do While Node.Children > 0
        TreeView.Nodes.Remove Node.Child.Index
    Loop
    
    'Show hourglass because network operations can take a little while
    TreeView.MousePointer = ccArrowHourglass
    'Now add them back
    For Each tFld In fld.SubFolders
        'Creates a new node for each sub-folders
        Set xNode = TreeView.Nodes.Add(Node.Key, tvwChild, Node.FullPath & StrConv(tFld.Name, vbProperCase), StrConv(tFld.Name, vbProperCase), "ClosedFolder")
        xNode.ExpandedImage = "OpenFolder"
        xNode.Sorted = True
        'Create Dummy Node if the folder contain subfolders
        If tFld.SubFolders.Count > 0 Then
            TreeView.Nodes.Add xNode, tvwChild
        End If
        
        If bExpandToFiles Then
            If tFld.Files.Count > 0 Then
                'Don't create a dummy node if one has already been created about
                If tFld.SubFolders.Count = 0 Then
                    TreeView.Nodes.Add xNode, tvwChild
                End If
            End If
        End If
    Next
    If bExpandToFiles Then
        For Each tFil In fld.Files
            'Creates a new node for each sub-folders
            Set xNode = TreeView.Nodes.Add(Node.Key, tvwChild, Node.FullPath & StrConv(tFil.Name, vbProperCase), StrConv(tFil.Name, vbProperCase), "File")
            xNode.Sorted = True
        Next
    End If
    TreeView.MousePointer = ccDefault
End Sub

'Created 2/24/00
'Return the next avaliable drive letter
'Purposely skips A,B, and C
Public Function NextAvailableDriveLetter() As String
    Dim fso, intChar As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For intChar = 68 To 90  'D to Z, case doesn't matter
        If fso.DriveExists(Chr(intChar)) = False Then
            NextAvailableDriveLetter = Chr(intChar)
            Exit Function
        End If
    Next intChar
End Function
