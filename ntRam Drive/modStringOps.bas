Attribute VB_Name = "modStringOps"
Option Explicit

'This file was created 2/13/00
'by Shannon Little
'http://go.to/neotrix
'This contains all of my functions to format and do operations on strings
'and numbers

'Private Declare Function StrFormatByteSize Lib "shlwapi" Alias "StrFormatByteSizeA" ( _
            'ByVal dw As Long, _
            'ByVal pszBuf As String, _
            'ByRef cchBuf As Long) As String


Enum ByteScale
    bsKb = 0
    bsMb = 1
    bsGb = 2
    bsTb = 3
End Enum

Enum ConvertScale
    csByte = 4
    csKb = 3
    csMb = 2
    csGb = 1
    csTb = 0
End Enum


'Created 2/13/00
Public Function CapFirstLetter(ByVal strString As String) As String
    'Returns a string with just the first letter capitalized
    CapFirstLetter = UCase(Left(strString, 1)) & Right(strString, Len(strString) - 1)
End Function

'Created 2/13/00
Public Function StrReturnRight(ByVal strString As String, ByVal strKeyValue As String) As String
    'Check for a key value, such as "." and returns everything to the
    'right of it, starts at the first occurence
    Dim N As Integer
    N = InStr(1, strString, strKeyValue, vbBinaryCompare)    'Returns the position of first occurence
    
    If N = 0 Then  '0 means is isn't in there, return original string
        StrReturnRight = strString
        Exit Function
    Else
        StrReturnRight = Right(strString, Len(strString) - N)
    End If
    
    'sdfs.ext
    '12345678
    'N=5, return 6 to 8
End Function

'Created 3/10/00
'Check for a key value, such as "." and returns everything to the
'right of it. It starts searching from the end of the string
Public Function StrReturnRightFromEnd(ByVal strString As String, ByVal strKeyValue As String) As String
    Dim N As Integer
      
    N = InStrRev(strString, strKeyValue, , vbBinaryCompare)   'Returns the position of first occurence from the end
    
    If N = 0 Then  '0 means is isn't in there, return original string
        StrReturnRightFromEnd = strString
        Exit Function
    Else
        StrReturnRightFromEnd = Right(strString, Len(strString) - N)
    End If
End Function

'Created 2/13/00
Public Function StrReturnLeft(ByVal strString As String, ByVal strKeyValue As String) As String
    'Check for a key value, such as "." and returns everything to the
    'left of it, starts at the first occurence
    Dim N As Integer
    N = InStr(1, strString, strKeyValue, vbBinaryCompare)
    
    If N = 0 Then  '0 means is isn't in there, return original string
        StrReturnLeft = strString
        Exit Function
    Else
        StrReturnLeft = Left(strString, N - 1)
    End If
End Function

'Created 3/10/00
'Check for a key value, such as "." and returns everything to the
'right of it. It starts searching from the end of the string
Public Function StrReturnLeftFromEnd(ByVal strString As String, ByVal strKeyValue As String) As String
    'Check for a key value, such as "." and returns everything to the
    'left of it, starts at the first occurence
    Dim N As Integer
    N = InStrRev(strString, strKeyValue, , vbBinaryCompare)   'Returns the position of first occurence from the end
    
    If N = 0 Then  '0 means is isn't in there, return original string
        StrReturnLeftFromEnd = strString
        Exit Function
    Else
        StrReturnLeftFromEnd = Left(strString, N - 1)
    End If
End Function

'Created 2/13/00
Public Function DoesCharExist(ByVal strStringToSearch As String, ByVal strKey As String) As Boolean
    Dim Result
    'Its defualt search mode is binary, so case does matter
    'Returns TRUE if it exists
    Result = InStr(1, strStringToSearch, strKey)
    DoesCharExist = IIf(Result > 0, True, False)
End Function

'Created 2/23/00
'Capitializes the first letter in every word
Public Function FormatProperCase(ByVal strString As String) As String
    FormatProperCase = StrConv(strString, vbProperCase)
End Function

'Created 2/26/00
'Formats the string to a good looking file/folder name
Public Function FormatProperName(ByVal strString As String) As String
    FormatProperName = FormatProperCase(LCase(strString))
End Function


'Created 2/13/00
'I just found out there was already a VB function for this, well this really sucks
'Public Function CapAfterEverySpace(ByVal strString As String)
    'Returns a string with every letter after a space captalized
    'Dim strFront As String, strMid As String, strEnd As String
    'Dim N As Integer
    
    'For N = 1 To Len(strString)   'Go through each letter
        'If Mid(strString, N, 1) = " " Then    'If we find a space
            'Caps the letter after the space
            'Example (for me)
            'new folder
            '12345678910
            'N=4, So 1 to 4 Add to Ucase(5) Add to 5 to 10
            'strFront = Left(strString, N)
            'strMid = UCase(Mid(strString, N + 1, 1))
            'strEnd = Right(strString, Len(strString) - N - 1)
            'This was my all in 1 line version, but it was too hard to read at a glance
            'strTemp = Left(strTemp, N) & UCase(Mid(strTemp, N + 1, 1)) & Right(strTemp, Len(strTemp) - N - 2)
            'Combine new name and keep looping through checking for more spaces
            'strString = strFront & strMid & strEnd
            'If you uncomment this, you will see exactly how, on each loop, the program does the caps after each space
            'Debug.Print "strTemp: " & strTemp, "Front: " & strFront, "Mid: " & strMid, "End: " & strEnd, "Comp: " & strFront & strMid & strEnd
        'End If
    'Next
    'CapAfterEverySpace = strString
'End Function

'DO NOT USE, does not handle large file sizes over 1.5 Gb
'Formats a bytes number to Kb or Mb
'Public Function FormatSize(ByVal lngAmount As Long) As String
    'Dim strBuffer As String
    'Dim strReturn As String
'
    'strBuffer = Space$(255)
    'strReturn = StrFormatByteSize(lngAmount, strBuffer, Len(strBuffer))
'
    'If InStr(strReturn, vbNullChar) <> 0 Then
        'FormatSize = Left$(strReturn, InStr(strReturn, vbNullChar) - 1)
    'End If
'End Function



'Created 2/23/00
'I made this because the built in Window version did not handle large sizes
'You must pass a BYTE size only
'It will then convert it to the best size to display it as, ie MB,GB,TB!
Public Function FormatBytesToBestSize(ByVal Amount, Optional bUseCommas As Boolean, Optional intRoundToPlaces As Integer) As String
    'All variables that handle the numbers are left as variant because
    'The file size can get Very large when expressed as bytes
    
    'If no value is specified, the default will be 2 places
    If IsMissing(intRoundToPlaces) Then intRoundToPlaces = 2
    'If no value is specified, the default will be false
    If IsMissing(bUseCommas) Then bUseCommas = False
    
    If Amount < 1024 Then
        Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
        FormatBytesToBestSize = Amount & " Bytes"
    Else
        Amount = Amount / 1024
        If Amount < 1024 Then
            Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
            FormatBytesToBestSize = Amount & " Kb"
        Else
            Amount = Amount / 1024
            If Amount < 1024 Then
                Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
                FormatBytesToBestSize = Amount & " Mb"
            Else
                Amount = Amount / 1024
                If Amount < 1024 Then
                    Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
                    FormatBytesToBestSize = Amount & " Gb"
                Else
                    Amount = Amount / 1024
                    If Amount < 1024 Then
                        Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
                        FormatBytesToBestSize = Amount & " Tb"
                    End If
                End If
            End If
        End If
    End If
End Function

'Created 2/24/00
'This will format a bytes number to the specified type, MB, GB TB!
'Without adding on what type it is, ie Adding MB to 96 MB
'It will just return 96
'I decides to make a function just for Bytes because it is the most (only?)
'Commonly returned size from window API functions
Public Function FormatByteSize(ByVal Amount, bsScale As ByteScale, Optional bFormatToString As Boolean, Optional bUseCommas As Boolean, Optional intRoundToPlaces As Integer) As String
    'All variables that handle the numbers are left as variant because
    'The file size can get Very large when expressed as bytes
    
    'If no value is specified, the default will be 2 places
    If IsMissing(intRoundToPlaces) Then intRoundToPlaces = 2
    'If no value is specified, the deafult will be false
    If IsMissing(bFormatToString) Then bFormatToString = False
    'If no value is specified, the defaul will be false
    If IsMissing(bUseCommas) Then bUseCommas = False
    
    Select Case bsScale
        Case bsKb:
            Amount = Amount / 1024
            Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
            FormatByteSize = IIf(bFormatToString, Amount & " Kb", Amount)
        Case bsMb:
            Amount = Amount / 1024 / 1024
            Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
            FormatByteSize = IIf(bFormatToString, Amount & " Mb", Amount)
        Case bsGb:
            Amount = Amount / 1024 / 1024 / 1024
            Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
            FormatByteSize = IIf(bFormatToString, Amount & " Gb", Amount)
        Case bsTb:
            Amount = Amount / 1024 / 1024 / 1024 / 1024
            Amount = IIf(bUseCommas, Format(Round(Amount, intRoundToPlaces), "###,###,###.##########"), Round(Amount, intRoundToPlaces))
            FormatByteSize = IIf(bFormatToString, Amount & " Tb", Amount)
    End Select
    
End Function

'Created 2/24/00
'This function will convert any size (Byte,Kb,Mb,Gb,Tb) to any size (Byte,Kb,Mb,Gb,Tb)
'Of course some resolution is lost when going from Mb to Byte unless you are using a single number
Public Function ConvertSizeTo(ByVal Amount, bsFromScale As ConvertScale, bsToScale As ConvertScale, Optional intRoundToPlaces As Integer)
    Dim intLevels As Integer
    
    If IsMissing(intRoundToPlaces) Then intRoundToPlaces = 2
    
    'Shows how many levels to go up or down
    'Byte = 4
    'Mb = 2
    '4 - 2 = 2 levels up to go
    'So Amount / 1024 / 1024 = Mb from Bytes
    intLevels = bsFromScale - bsToScale
    
    
    Select Case intLevels
        Case -4: Amount = Amount * 1024 * 1024 * 1024 * 1024
        Case -3: Amount = Amount * 1024 * 1024 * 1024
        Case -2: Amount = Amount * 1024 * 1024
        Case -1: Amount = Amount * 1024
        Case 0: 'Same size in both From and To
        Case 1: Amount = Amount / 1024
        Case 2: Amount = Amount / 1024 / 1024
        Case 3: Amount = Amount / 1024 / 1024 / 1024
        Case 4: Amount = Amount / 1024 / 1024 / 1024 / 1024
    End Select
    ConvertSizeTo = Round(Amount, intRoundToPlaces)
End Function

'Created 2/25/00
'Returns TRUE if the strSearchForString string in somewhere in the strSourceString string
Public Function IsStringContainedIn(ByVal strSourceString As String, ByVal strSearchForString As String, Optional bCaseMatters As Boolean) As Boolean
    Dim intSearchType As Integer
    
    'By default case does not matter when searching
    If IsMissing(bCaseMatters) Then bCaseMatters = False
    
    'vbBinaryCompare 0 Performs a binary comparison.    'Case sensitive
    'vbTextCompare 1 Performs a textual comparison.     'Not case sensitive
    
    If bCaseMatters Then
        intSearchType = 0
    Else
        intSearchType = 1
    End If
    
    If InStr(1, strSourceString, strSearchForString, intSearchType) > 0 Then
        IsStringContainedIn = True
    Else
        IsStringContainedIn = False
    End If
End Function
