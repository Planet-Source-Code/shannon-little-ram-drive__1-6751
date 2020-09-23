Attribute VB_Name = "modHardware"
Option Explicit

'This mod was created 2/24/00
'It contains functions to retrieve info on the computer's hardware
'Such as Gfx or Memory
'This module requires the modStringOps.bas module if formatting of return
'values is to be enabled


Type MEMORYSTATUS   'memory status structure
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

'Return all info on memory
Private Declare Sub GlobalMemoryStatus Lib "kernel32.dll" (lpBuffer As MEMORYSTATUS)

Public Function AvaliablePhysicalMemory(Optional bFormat As Boolean)
    'This function returns a string if bFormat is TRUE
    'else it return a number in bytes
    Dim lpBuffer As MEMORYSTATUS
    GlobalMemoryStatus lpBuffer
    
    'By default it will be false if no value is given
    If IsMissing(bFormat) Then bFormat = False
    
    If bFormat Then
        'This will size it to the most readable form, MB,GB and add that string to the end
        AvaliablePhysicalMemory = FormatBytesToBestSize(lpBuffer.dwAvailPhys, True)
    Else
        AvaliablePhysicalMemory = lpBuffer.dwAvailPhys
    End If
End Function

Public Function AvaliableVirtualMemory(Optional bFormat As Boolean)
    'This function returns a string if bFormat is TRUE
    'else it return a number in bytes
    Dim lpBuffer As MEMORYSTATUS
    GlobalMemoryStatus lpBuffer
    
    'By default it will be false if no value is given
    If IsMissing(bFormat) Then bFormat = False
    
    If bFormat Then
        AvaliableVirtualMemory = FormatBytesToBestSize(lpBuffer.dwAvailVirtual, True)
    Else
        AvaliableVirtualMemory = lpBuffer.dwAvailVirtual
    End If
End Function

Public Function TotalPhysicalMemory(Optional bFormat As Boolean)
    'This function returns a string if bFormat is TRUE
    'else it return a number in bytes
    Dim lpBuffer As MEMORYSTATUS
    GlobalMemoryStatus lpBuffer
    
    'By default it will be false if no value is given
    If IsMissing(bFormat) Then bFormat = False
    
    If bFormat Then
        TotalPhysicalMemory = FormatBytesToBestSize(lpBuffer.dwTotalPhys, True)
    Else
        TotalPhysicalMemory = lpBuffer.dwTotalPhys
    End If
End Function

Public Function TotalVirtualMemory(Optional bFormat As Boolean)
    'This function returns a string if bFormat is TRUE
    'else it return a number in bytes
    Dim lpBuffer As MEMORYSTATUS
    GlobalMemoryStatus lpBuffer
    
    'By default it will be false if no value is given
    If IsMissing(bFormat) Then bFormat = False
    
    If bFormat Then
        TotalVirtualMemory = FormatBytesToBestSize(lpBuffer.dwTotalVirtual, True)
    Else
        TotalVirtualMemory = lpBuffer.dwTotalVirtual
    End If
End Function
