Attribute VB_Name = "modCommonProgramFunctions"
Option Explicit


'Created 2/20/00
'This file created by Shannon Little
'http://go.to/neotrix
'This file contains function common to all of my programs
'Such as functions to keep errors looking the same

'Used by Browse for Folder
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" ( _
    lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" ( _
    ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long

Private Declare Function lstrcat Lib "KERNEL32" Alias "lstrcatA" ( _
    ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type
'*****

'Created 2/20/00
'Typical error function
'Keeps my error message looking constant across programs
Public Sub TypError(ByVal strMainMessage As String)
    MsgBox "Error: " & strMainMessage, vbOKOnly + vbExclamation, strProgramName
End Sub

'Created 2/25/00
'Typical information function
'For any general error is any of my error handling routines
Public Sub TypInfo(ByVal strMainMessage As String)
    MsgBox strMainMessage, vbOKOnly + vbInformation, strProgramName
End Sub

'Created 2/20/00
'General error function
'For any general error is any of my error handling routines
Public Sub GenError()
    MsgBox "There was an error" & vbNewLine & "Source: " & Err.Source & vbNewLine & "Code: " & Err.Number & vbNewLine & "Description: " & Err.Description, vbOKOnly + vbExclamation, "Generic Error Handler"
End Sub

'Created 2/22/00
'Shows the Browse For Folder dialog
'Copied from MSDN
Public Function BrowseForFolder(hWnd, Optional szTitle As String) As String
    'Opens a Treeview control that displays the directories in a computer
    Dim lpIDList As Long
    Dim sBuffer As String
    'Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo

    If IsMissing(szTitle) Then szTitle = "Browse for folder:"
    
    With tBrowseInfo
       .hWndOwner = hWnd
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
    BrowseForFolder = sBuffer
End Function

'Created 2/20/00
Private Sub CODEtoCOPYtoFORM()

'Selects the text in a textbox when clicked
'Private Sub SelectText()
    'ActiveControl.SelStart = 0
    'ActiveControl.SelLength = Len(ActiveControl.Text)
'End Sub

End Sub
