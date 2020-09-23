Attribute VB_Name = "modWebFunctions"
Option Explicit

'This file was created 2/22/00
'by Shannon Little
'http://go.to/neotrix
'This file contains functions for web operations

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function FindExecutable Lib "shell32.dll" Alias _
    "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
    String, ByVal lpResult As String) As Long
    
'Created 2/22/00
Public Sub OpenWebPage(ByVal strWebPage As String)
    Dim Result
    Result = Shell("start.exe " & strWebPage, vbHide)
End Sub

'Created 2/23/00
Public Sub GotoMyWebSite()
    OpenWebPage "http://go.to/neotrix"
End Sub

'Created 2/25/00
Public Function InternetConnectionPresent(Winsock As Winsock) As Boolean
    If Winsock.LocalIP = "127.0.0.1" Or Winsock.LocalIP = "0.0.0.0" Then
        InternetConnectionPresent = False
    Else
        InternetConnectionPresent = True
    End If
End Function

'Created 2/25/00
'This will check my website to see if the program is up to date
'It will return TRUE if there is a new version of the program
'If bAskUserToDLNewVersion is TRUE then a msgbox will pop-up
'Asking the user if they want to DL the new version now, it will open a browser window
'Requires the InternetTransferControl
'                                          \/ InternetTransferControl \/
Public Sub CheckForNewVersionOfProgram(Inet As Inet, ByVal strProgramName As String, ByVal intCurrentMajorVersion As Integer, ByVal intCurrentMinorVersion As Integer, ByVal bAskUserToDLNewVersion As Boolean)
    On Error GoTo lError
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    
    Dim strTemp As String, strTempName As String, bNewerVersion As Boolean
    Dim strURL2Dl As String, Result
    Dim intNewerMajorVersion As Integer, intNewerMinorVersion As Integer
    Dim fso, fil
    
    Set fso = CreateObject("Scripting.FileSystemObject")

    With Inet
        .Protocol = icHTTP
        .RequestTimeout = Val(TestGlobalKey("WebTimeout", 10))
    End With
    
    DoEvents
    
    strTemp = Inet.OpenURL("http://saturn.spaceports.com/~neotrix/Files/ntPrograms/Versions.txt", icString)
    
    DoEvents
    
    'A check to make sure we recieved the right file
    If Left(strTemp, 8) = "VERSIONS" Then
        'It is eaiser to just write this to a file the use fil.ReadLine to
        'Check each line instead of checking for where vbNewLine is located in the file
        strTempName = fso.GetTempName   'Return a temporary file name
        Set fil = fso.CreateTextFile(App.Path & "\" & strTempName, True)
        fil.Write (strTemp)
        fil.Close
        
        'No read the file and check for the program name
        Set fil = fso.OpenTextFile(App.Path & "\" & strTempName, ForReading, TristateUseDefault)
    
        bNewerVersion = False
        
        'This reads the file and checks for the program name
        'It then compares the major and minor version to see if a newer version is avaliable
        Do While fil.AtEndOfStream <> True
            strTemp = fil.ReadLine
            If UCase(strTemp) = UCase(strProgramName) Then  'Ucase'd just to make sure it works in case of a typo in the versions.txt file
                'Now there are just 3 more lines to read
                'Major version, Minor version, then location to DL new version
                intNewerMajorVersion = Val(fil.ReadLine)
                intNewerMinorVersion = Val(fil.ReadLine)
                If intNewerMajorVersion > intCurrentMajorVersion Then bNewerVersion = True
                If intNewerMinorVersion > intCurrentMinorVersion Then bNewerVersion = True
                strURL2Dl = fil.ReadLine
                If bNewerVersion Then
                    If bAskUserToDLNewVersion Then
                        Result = MsgBox("There is a newer version of " & strProgramName & " avaliable for download" & vbNewLine & _
                                "You are using version " & intCurrentMajorVersion & "." & intCurrentMinorVersion & ", the current version is " & intNewerMajorVersion & "." & intNewerMajorVersion & vbNewLine & _
                                "Would you like to download the newer version now?", vbYesNo + vbQuestion, "Newer version avaliable")
                        If Result = vbYes Then
                            OpenWebPage strURL2Dl
                        End If
                    End If
                End If
                fil.Close
                fso.DeleteFile App.Path & "\" & strTempName, True
                Exit Do
            End If
        Loop
    End If
    Exit Sub
lError:
    Select Case Err.Number
        Case 1: 'Timeout error, just ignore it
        Case Default: GenError
    End Select
End Sub

