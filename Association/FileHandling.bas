Attribute VB_Name = "FileHandling"
'general file handling stuff here.
'Nothing to do with Association
Option Explicit
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Function OneGulp(Src As String) As String
    On Error Resume Next
    Dim f As Integer, temp As String
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get #f, , temp
    Close #f
    If Left(temp, 2) = "ÿþ" Or Left(temp, 2) = "þÿ" Then temp = Replace(Right(temp, Len(temp) - 2), Chr(0), "")
    OneGulp = temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function FileExists(sSource As String) As Boolean
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function

