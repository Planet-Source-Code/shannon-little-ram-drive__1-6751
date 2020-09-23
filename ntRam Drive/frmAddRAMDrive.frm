VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddRAMDrive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add RAM Drive"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmAddRAMDrive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMemoryLocation 
      Caption         =   "RAM Drive Memory Location"
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   4575
      Begin VB.OptionButton optExpanded 
         Caption         =   "Expanded memory"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optExtended 
         Caption         =   "Extended memory"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   2760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdBrowseForConfig 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreateRAMDrive 
      Caption         =   "C&reate"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin MSComCtl2.UpDown udnVolumeSize 
      Height          =   285
      Left            =   1935
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtSize"
      BuddyDispid     =   196615
      OrigLeft        =   2760
      OrigTop         =   480
      OrigRight       =   3000
      OrigBottom      =   1215
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtVolumeName 
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   -120
      X2              =   9200
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   -120
      X2              =   9200
      Y1              =   3615
      Y2              =   3615
   End
   Begin VB.Line MenuLineLight 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9320
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line MenuLineDark 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   9320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblConfigLocation 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   315
      UseMnemonic     =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblConfigLocationTitle 
      Caption         =   "Location of Config.sys:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   80
      Width           =   1815
   End
   Begin VB.Label lblMaxCharsVolumeName 
      Caption         =   "(8 Chars. Max)"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblSizeInfo 
      Caption         =   $"frmAddRAMDrive.frx":27A2
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label lblMemoryAmount 
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblSize 
      Caption         =   "Size:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblVolumeName 
      Caption         =   "Volume name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddRAMDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowseForConfig_Click()
    On Error GoTo lError:
lStartOver:
    With CommonDialog
        .InitDir = "C:\"
        .DefaultExt = "*.sys"
        .DialogTitle = "Locate config.sys"
        .FileName = "*.sys"
        .Filter = "System File (*.sys)|*.sys"
        .FLAGS = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNPathMustExist
        .ShowOpen
        If .FileTitle <> "Config.sys" Then
            TypError "This is not a Config.sys file"
            .FileName = "*.sys"
            GoTo lStartOver
        End If
    End With
    lblConfigLocation.Caption = CommonDialog.FileName
    Exit Sub
lError:
End Sub

Private Sub cmdCancel_Click()
    Unload frmAddRAMDrive
End Sub

Private Sub cmdCreateRAMDrive_Click()
    On Error GoTo lError
    
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    '     SystemDefault,            UNICode,           ASCII
    
    Dim strOutput As String, strAdditionalSwitches As String, strAdditionalLines As String
    Dim fso, fil, txStr, strRamDriveLocation As String, Result
    'FileSystem Object, File, TextStream
    Dim bHimemPresent As Boolean, bEmm386Present As Boolean, strTemp As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(lblConfigLocation.Caption) Then
        'Checks the make sure a copy of RamDrive.sys is located in the correct
        'place on the user's computer
        
        'Checking for the only 3 locations RamDrive.sys could be located
        If fso.FileExists("C:\Windows\Command\RamDrive.sys") Then
            strRamDriveLocation = "C:\Windows\Command\RamDrive.sys"
        Else
            If fso.FileExists("C:\Windows\Command\Ebd\RamDrive.sys") Then
                strRamDriveLocation = "C:\Windows\Command\Ebd\RamDrive.sys"
            Else
                If fso.FileExists("C:\Windows\RamDrive.sys") Then
                    strRamDriveLocation = "C:\Windows\RamDrive.sys"
                Else
                    TypError "Cannot location RamDrive in the standard directories of C:\Windows, C:\Windows\Command, or C:\Windows\Command\Edb " & vbNewLine & "Please move a copy of it there so this program can work."
                End If
            End If
        End If
        
        'Check to make sure Himem.sys and Emm386.exe is loaded
        'As they are required for Ramdrive.sys to work
        Set fil = fso.OpenTextFile(lblConfigLocation.Caption, ForReading, TristateUseDefault)
        
        bEmm386Present = False
        bHimemPresent = False
        
        'Checking for HIMEM.SYS
        Do While fil.AtEndOfStream <> True
            strTemp = fil.ReadLine
            'If the line isn't commented out
            If Left(UCase(strTemp), 3) <> "REM" Then
                If IsStringContainedIn(strTemp, "Himem.sys") Then
                    bHimemPresent = True
                Else    'If its not Himem.sys then maybe its emm386.exe
                    If IsStringContainedIn(strTemp, "Emm386.exe") Then
                        bEmm386Present = True
                    End If
                End If
            End If
        Loop
        fil.Close
        
        'Gets user choice if they want the program to add these 2 missing lines
        strAdditionalLines = ""
        If bHimemPresent = False Then
            Result = MsgBox("Himem.sys is not being loaded. It is either commented out or not present in your Config.sys file." & vbNewLine & "Would you like me to add it in for you?", vbYesNo + vbQuestion, "Himem.sys is not loaded")
            If Result = vbYes Then
lStartOverHimem:
                With CommonDialog
                    .InitDir = "C:\Windows"
                    .DefaultExt = "*.sys"
                    .DialogTitle = "Locate Himem.sys"
                    .FileName = "*.sys"
                    .Filter = "System File (*.sys)|*.sys"
                    .FLAGS = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNPathMustExist
                    .ShowOpen
                    If UCase(.FileTitle) <> UCase("Himem.sys") Then
                        TypError "This is the Himem.sys file. Please locate it."
                        .FileName = "*.sys"
                        GoTo lStartOverHimem
                    End If
                    strAdditionalLines = "device=" & CommonDialog.FileName
                End With
            End If
        End If
        
        If bEmm386Present = False Then
            Result = MsgBox("Emm386.exe is not being loaded. It is either commented out or not present in your Config.sys file." & vbNewLine & "Would you like me to add it in for you?", vbYesNo + vbQuestion, "Emm386.exe is not loaded")
            If Result = vbYes Then
lStartOverEmm:
                With CommonDialog
                    .InitDir = "C:\Windows"
                    .DefaultExt = "*.exe"
                    .DialogTitle = "Locate Emm386.exe"
                    .FileName = "*.exe"
                    .Filter = "Executable (*.exe)|*.exe"
                    .FLAGS = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNPathMustExist
                    .ShowOpen
                    If UCase(.FileTitle) <> UCase("Emm386.exe") Then
                        TypError "This is not a Config.sys file"
                        .FileName = "*.sys"
                        GoTo lStartOverEmm
                    End If
                    strAdditionalLines = strAdditionalLines & vbNewLine & _
                                        "device=" & CommonDialog.FileName & " on"
                End With
            End If
        End If
        
        'Now get all the info together to write to the end of the config.sys file
        Set fil = fso.OpenTextFile(lblConfigLocation.Caption, ForAppending, TristateUseDefault)
    
        If optExtended.Value = True Then
            strAdditionalSwitches = "/e"
        Else
            strAdditionalSwitches = "/a"
        End If
        
        strOutput = strAdditionalLines & vbNewLine & vbNewLine & _
                    "REM This line added by ntRamDrive http://go.to/neotrix" & vbNewLine & _
                    "device=" & strRamDriveLocation & " " & ConvertSizeTo(Val(txtSize.Text), csMb, csKb, 0) & " " & strAdditionalSwitches
                    
        'Write lines to Config.sys file
        fil.Write vbNewLine & vbNewLine     'Put a blank line in there
        fil.Write strOutput     'Write to file
        fil.Close
        
        'Now write to Autoexec.bat, which should be in the same place
        'Label RAMDrive, NewLabel
        'This will set the volume name everytime the computer is reset
        
        If fso.FileExists(StrReturnLeft(lblConfigLocation.Caption, "\") & "\Autoexec.bat") Then
            Set fil = fso.OpenTextFile(StrReturnLeft(lblConfigLocation.Caption, "\") & "\Autoexec.bat", ForAppending, TristateUseDefault)
            
            
            strOutput = "REM This line added by ntRamDrive http://go.to/neotrix" & vbNewLine & _
                        "REM label INSERTDRIVELETTER: " & txtVolumeName.Text
            'The drive letter is only a prediction by getting the next avaliable drive letter
            fil.Write vbNewLine & vbNewLine
            fil.Write strOutput
            fil.Close
        Else
            TypError "The Autoexec.bat file could not be found at that same location as Config.sys. Please click on Browse... and locate it. Or create one if it does not exist."
        Exit Sub
    End If
    Else
        TypError "The Config.sys file could not be found at that location. Please click on Browse... and locate it"
        Exit Sub
    End If
    
    TypInfo "Remember: Anything stored in the RAM drive goes bye-bye if your computer crashes or is shutdown."
    TypInfo "You must reset your computer before you can use the RAM drive." & vbNewLine & _
            "The code to change the Volume Name of the RAM drive has been added to the Autoexec.bat file.  However you must manually edit it and insert the correct Drive Letter and uncomment it by removing the REM in front of the Label command"
    TypInfo "Drive created successfully!"
    
    SetKey "CfgLocation", lblConfigLocation.Caption
    
    Unload frmAddRAMDrive
    Exit Sub
lError:
    Select Case Err.Number
        Case 32755: 'Cancel error from common dialog
        Case Default:   GenError
    End Select
End Sub

Private Sub Form_Load()
    lblMemoryAmount.Caption = FormatByteSize(TotalPhysicalMemory, bsMb, True, False)
    udnVolumeSize.Max = Val(FormatByteSize(TotalPhysicalMemory, bsMb))
    If udnVolumeSize.Max > 32 Then
        txtSize.Text = Format((udnVolumeSize.Max - 32) * 0.2, "####.#")
    End If
    
    lblConfigLocation.Caption = TestKey("CfgLocation", "C:\Config.sys")
End Sub



Private Sub txtSize_Change()
    'Prevent non-numbers from being entered
    txtSize.Text = Val(txtSize.Text)
    'Prevents user from typing over max
    If Val(txtSize.Text) > udnVolumeSize.Max Then
        txtSize.Text = Str(udnVolumeSize.Max)
    End If
    'Prevent negative numbers
    If Val(txtSize.Text) <= 0 Then
        txtSize.Text = Abs(Val(txtSize.Text))
    End If
End Sub

'Selects the text in a textbox when clicked
Private Sub SelectText()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.Text)
End Sub

Private Sub txtSize_GotFocus()
    SelectText
End Sub

Private Sub txtVolumeName_GotFocus()
    SelectText
End Sub
