VERSION 5.00
Begin VB.Form frmRemoveRAMDrive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Remove RAM Drive"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmRemoveRAMDrive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   1850
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemoveDrive 
      Caption         =   "&Remove Selected Drives"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1850
      Width           =   2055
   End
   Begin VB.ListBox lstRAMDrives 
      Height          =   1410
      ItemData        =   "frmRemoveRAMDrive.frx":27A2
      Left            =   120
      List            =   "frmRemoveRAMDrive.frx":27A4
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "List of all RAM drives found in the Config.sys file"
      Top             =   360
      Width           =   6255
   End
   Begin VB.Label lblDrivesFoundInCFGFile 
      Caption         =   "Drives found in Config.sys file:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   80
      Width           =   2175
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "Refresh"
      Visible         =   0   'False
      Begin VB.Menu mnuRefreshRefreshList 
         Caption         =   "&Refresh"
      End
   End
End
Attribute VB_Name = "frmRemoveRAMDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload frmRemoveRAMDrive
End Sub

Private Sub cmdRemoveDrive_Click()
    On Error GoTo lError
    
    'This will read the entire config.sys file into an array
    'If will then remove the line numbers of all checked items in the
    'ListBox control
    'It will then write the remaining data back into the config.sys file
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

    'Array is created large enough to handle any size Config.sys file
    Dim strCfgFile(1 To 200) As String, N As Integer
    Dim strTemp As String
    Dim fso, fil
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If lstRAMDrives.ListCount = 0 Then
        TypInfo "No RAM drives located"
        Exit Sub
    End If
    
    If fso.FileExists(GetKey("CfgLocation")) Then
        'Checks the make sure a copy of RamDrive.sys is located in the correct
        'place on the user's computer
        
        Set fil = fso.OpenTextFile(GetKey("CfgLocation"), ForReading, TristateUseDefault)
        
        'Add each line to the array
        N = 1
        Do While fil.AtEndOfStream <> True
            strCfgFile(N) = fil.ReadLine
            N = N + 1
        Loop
        
        N = N + 1
        'To let the writing to file loop know when the file stored in the array has ended
        strCfgFile(N) = "$$ENDOFFILE$$"
        fil.Close

        'Now set the array position that equals the ItemData of the selected
        'item's index to ""
        For N = 0 To lstRAMDrives.ListCount - 1
            If lstRAMDrives.Selected(N) = True Then
                'strCfgFile(5) = ""
                '5 would be the line that a ramdrive.sys was detected and which
                'the user checked to be deleted
                 strCfgFile(lstRAMDrives.ItemData(N)) = ""
            End If
        Next N
        
        'This will write everything back to the file
        'But now the lines that were marked for deletion are just blank lines
        
        Set fil = fso.OpenTextFile(GetKey("CfgLocation"), ForWriting, TristateUseDefault)
        'Repeat from 1 to the array size
        N = 1
        Do While strCfgFile(N) <> "$$ENDOFFILE$$"
            fil.Write strCfgFile(N)
            fil.Write vbNewLine
            N = N + 1
        Loop
        fil.Close
        
        'Reload the list with the new config.sys file
        Form_Load
    Else
        TypError "Cannot locate Config.sys file" & vbNewLine & "Please goto Add Drive and set the path to it"
    End If
    
    Exit Sub
lError:
    GenError
End Sub

Private Sub Form_Load()
    On Error GoTo lError
    
    'This read the Config.sys file and test each line for the string "ramdrive.sys"
    'If it contains it, it is added to the ListBox and the line number it is located
    'on is stored in its ListData property
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

    Dim N As Integer
    Dim strTemp As String
    Dim fso, fil
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    lstRAMDrives.Clear
    If fso.FileExists(GetKey("CfgLocation")) Then
        'Checks the make sure a copy of RamDrive.sys is located in the correct
        'place on the user's computer
        
        Set fil = fso.OpenTextFile(GetKey("CfgLocation"), ForReading, TristateUseDefault)
        
        'Read each line, and if RamDrive.sys is detected in the list then
        'add it to the list box
        N = 1
        Do While fil.AtEndOfStream <> True
            strTemp = fil.ReadLine
            If IsStringContainedIn(strTemp, "RamDrive.sys") Then
                lstRAMDrives.AddItem strTemp
                lstRAMDrives.ItemData(lstRAMDrives.NewIndex) = N    'Store which line this was located at
            End If
            N = N + 1
        Loop
        fil.Close
        
    Else
        TypError "Cannot locate Config.sys file" & vbNewLine & "Please goto Add Drive and set the path to it"
    End If
    
    Exit Sub
lError:
    GenError
End Sub

Private Sub lstRAMDrives_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuRefresh, , , , mnuRefreshRefreshList
    End If
End Sub

Private Sub mnuRefreshRefreshList_Click()
    Form_Load
End Sub
