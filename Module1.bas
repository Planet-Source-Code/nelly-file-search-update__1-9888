Attribute VB_Name = "Module1"
Option Explicit

'**********************************************************************************
'FindFirstFile, FindNextFile, FindClose
'**********************************************************************************
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
(ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long


Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
(ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long


Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA '-----------These are self explanitary
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Public Enum FILE_ATTRIBUTES
    FILE_ATTRIBUTE_READONLY = &H1
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    FILE_ATTRIBUTE_ENCRYPTED = &H4000
End Enum

'**********************************************************************************
'GetTickCount, Used for timing events
'**********************************************************************************
Declare Function GetTickCount Lib "kernel32" () As Long


