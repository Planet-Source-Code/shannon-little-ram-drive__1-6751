VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ntRam Drive v1.0"
   ClientHeight    =   2685
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5115
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5115
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   2040
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   2160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iglTreeViewImages 
      Left            =   1800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A2
            Key             =   "Floppy"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D3E
            Key             =   "Fixed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32DA
            Key             =   "CD-ROM"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":343A
            Key             =   "Removable"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39D6
            Key             =   "Ram Disk"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F72
            Key             =   "Network"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":450E
            Key             =   "ClosedFolder"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AAA
            Key             =   "DisconnectNetwork"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5046
            Key             =   "OpenFolder"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmProperties 
      Caption         =   "Properties"
      Height          =   1695
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.Label lblDriveSize 
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblVolumeName 
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblVolumeNameTitle 
         Caption         =   "Volume name:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblDriveFreeSpace 
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblDriveFType 
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblDriveFileSystem 
         Caption         =   "Filesystem:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblFreeSpace 
         Caption         =   "Free space:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblDriveSizeTitle 
         Caption         =   "Drive size:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdRemoveDrive 
      Caption         =   "Remove Drive"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      ToolTipText     =   "Remove the currently selected RAM drive"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddDrive 
      Caption         =   "Add Drive..."
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Add a new RAM drive"
      Top             =   1920
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   2385
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5345
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   582
            MinWidth        =   353
            Picture         =   "frmMain.frx":55E2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   582
            MinWidth        =   441
            Picture         =   "frmMain.frx":72EE
            Key             =   "RamDrive"
            Object.ToolTipText     =   "Ram drive(s) found."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   582
            MinWidth        =   441
            Picture         =   "frmMain.frx":788A
            Key             =   "NoRamDrive"
            Object.ToolTipText     =   "No Ram drives found."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "7:30 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwRamDrives 
      Height          =   2055
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "List of currently running RAM drives"
      Top             =   300
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3625
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      Style           =   5
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "iglTreeViewImages"
      Appearance      =   1
   End
   Begin VB.Line MenuLineDark 
      BorderColor     =   &H00808080&
      X1              =   -3120
      X2              =   6200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line MenuLineLight 
      BorderColor     =   &H00FFFFFF&
      X1              =   -3120
      X2              =   6200
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Label lblRamDrivesInstalled 
      Caption         =   "Ram drives installed:"
      Height          =   255
      Left            =   40
      TabIndex        =   1
      Top             =   80
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpWhatIsARAMDrive 
         Caption         =   "&What is a RAM drive?"
      End
      Begin VB.Menu mnuHelpHyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpGotoWebsite 
         Caption         =   "&Goto NeoTrix Website..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ntRam Drive..."
      End
   End
   Begin VB.Menu mnuRefresh 
      Caption         =   "RefreshMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuRefreshRefreshList 
         Caption         =   "&Refresh"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAddDrive_Click()
    frmAddRAMDrive.Show vbModal, Me
End Sub

Private Sub cmdRemoveDrive_Click()
    'This will try to remove the selected drive from the Config.sys file
    'It will read the config.sys file and show a listbox with every entry
    'of ramdrive.sys
    'Since it is not possible to match the drive size exactly the
    'user will have to select which one to remove
    frmRemoveRAMDrive.Show vbModal, Me
End Sub

Private Sub Form_Load()
    strProgramName = "ntRam Drive"
    'This adds just the RAM drives to the TreeView control
    InitializeCertainDrives tvwRamDrives, dtRamDrive, False
    'Just a nice little icon in the status bar that is visible when
    'Ram drives are present
    If tvwRamDrives.Nodes.Count > 0 Then
        StatusBar.Panels("RamDrive").Visible = True
        StatusBar.Panels("NoRamDrive").Visible = False
    Else
        StatusBar.Panels("RamDrive").Visible = False
        StatusBar.Panels("NoRamDrive").Visible = True
    End If
    
    Top = TestKey("WindowMTop", 1500)
    Left = TestKey("WindowMLeft", 1500)
    Height = 3345
    
    StatusBar.Panels(1).Text = Caption
    
    'Checks for a new version of the program on the net
    If InternetConnectionPresent(Winsock) Then
        CheckForNewVersionOfProgram Inet, "ntRamDrive", 1, 0, True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetKey "WindowMTop", Top
    SetKey "WindowMLeft", Left
End Sub

Private Sub mnuFileExit_Click()
    Unload frmMain
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpGotoWebSite_Click()
    GotoMyWebSite
End Sub

Private Sub mnuHelpWhatIsARAMDrive_Click()
    MsgBox "A RAM drive is a virtual hard drive that is created in the memory of your computer.  In effect you have a hard drive, that shows up as a drive letter, that runs at the speed of your memory.  Which is extremely fast, but costly.", vbOKOnly + vbInformation, "What is a RAM Drive?"
End Sub

Private Sub mnuRefreshRefreshList_Click()
    'This adds just the RAM drives to the TreeView control
    AddOnlyDrives tvwRamDrives, dtRamDrive, False
    lblVolumeName.Caption = ""
    lblDriveSize.Caption = ""
    lblDriveFreeSpace.Caption = ""
    lblDriveFType.Caption = ""
End Sub

Private Sub tvwRamDrives_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuRefresh, , , , mnuRefreshRefreshList
    End If
End Sub

Private Sub tvwRamDrives_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo lError
    Dim fso, drv, cDrv
    'FileSystem Object, Drive, Drive Collection, Temporary Drive Object, Folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set drv = fso.GetDrive(Node.Key)
    
    lblVolumeName.Caption = drv.VolumeName
    lblDriveSize.Caption = FormatBytesToBestSize(drv.Totalsize)
    lblDriveFreeSpace.Caption = FormatBytesToBestSize(drv.Freespace)
    lblDriveFType.Caption = drv.FileSystem
    Exit Sub
lError:
    Select Case Err.Number
        'Case 71: TypError "Disk not ready"
        Case Default: GenError
    End Select
End Sub

'Selects the text in a textbox when clicked
Private Sub SelectText()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtDriveSize_GotFocus()
    SelectText
End Sub
