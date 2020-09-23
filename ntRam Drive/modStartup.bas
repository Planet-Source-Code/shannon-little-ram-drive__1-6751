Attribute VB_Name = "modStartup"
Option Explicit
Public strDisclaimer As String
Public strProgramName As String
Public strRegPath As String

'Created 1/10/00
'This is the startup file common to all of my programs
'Each program needs to have its own copy of this module

Sub main()
    'All settings are saved under the HKEY_CURRENT_USER
    ' Software\NeoTrix\ProgramName
    strRegPath = "Software\NeoTrix\ProgramName\1.0"
    strProgramName = "ntProgram Name"
    strDisclaimer = _
    "You should carefully read the following terms and conditions before using this software. Your use of this software indicates your acceptance of this license agreement and warranty." & _
    vbNewLine & vbNewLine & "Disclaimer of Warranty:" & _
    vbNewLine & vbNewLine & "THIS SOFTWARE AND THE ACCOMPANYING FILES ARE DISTRIBUTED 'AS IS' AND WITHOUT WARRANTIES AS TO PERFORMANCE OF MERCHANTABILITY OR ANY OTHER WARRANTIES WHETHER EXPRESSED OR IMPLIED." & _
    vbNewLine & vbNewLine & "NO WARRANTY OF FITNESS FOR A PARTICULAR PURPOSE IS OFFERED. THE USER MUST ASSUME THE ENTIRE RISK OF USING THIS PROGRAM." & _
    vbNewLine & vbNewLine & "Distribution:" & _
    vbNewLine & vbNewLine & "You may redistribute copies of this software, and copy this program to as many machines as you like, but you may offer such copies ONLY IDENTICAL TO THE ORIGINAL, including the software, source code and documentation. You are specifically prohibited from charging or requesting donations for any such copies, except for the cost of media. You are also prohibited from distributing the software, source code and documentation with commercial products without prior WRITTEN permission of the author." & _
    vbNewLine & vbNewLine & "Placing this software on any site which charges indirectly or directly for access to file downloading or accessing areas is strictly prohibited. (Except for ISP expenses) " & _
    vbNewLine & vbNewLine & "If you modify the source code to a substantial degree, then you are, of course, exempt from all the above. " & _
    vbNewLine & vbNewLine & "Basically I just don't want you taking my work, slapping your name on it and selling it. Nor blaming me if this program somehow corrupts your mind or computer." & _
    vbNewLine & vbNewLine & "Copyright(c) 2000 Shannon Little" & _
    vbNewLine & "codingman@yahoo.com" & _
    vbNewLine & "http://go.to/neotrix"
    'AgreeToDisclaimer"
    If TestKey("ATD", "FALSE") <> "TRUE" Then        'See if disclaimer has already been agreed to
        frmDisclaimer.Show
    Else    'Already been agreed to
        frmMain.Show
    End If
End Sub
