VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Api"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' A.P.I. Class
' .file io.

Const FILE_ATTRIBUTE_TEMPORARY = &H100
Const FILE_ATTRIBUTE_NORMAL = &H80
Const CREATE_ALWAYS = 2
Const FILE_BEGIN = 0
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_NEW = 1
Const OPEN_EXISTING = 3
Const OPEN_ALWAYS = 4
Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private hFile As Long, nSize As Long, Ret As Long

Public Function pReadFile(FileName As String, bBytes() As Byte) As Long

PreLoadFile True, FileName

    ReDim bBytes(1 To nSize - Int(nSize / 2)) As Byte

ReadFile hFile, bBytes(1), UBound(bBytes), Ret, ByVal 0&

pReadFile = Ret
CloseHandle hFile

End Function

Public Function pWriteFile(FileName As String, sText As String) As Long

Dim bBytes() As Byte
Str2Byte sText, bBytes()

PreLoadFile False, FileName

WriteFile hFile, bBytes(0), UBound(bBytes()) + 1, Ret, ByVal 0&

pWriteFile = Ret
CloseHandle hFile

End Function

Private Function PreLoadFile(Read As Boolean, FileName As String)

If Read = False Then
    hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0&)
Else
    hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
End If

nSize = GetFileSize(hFile, 0)
SetFilePointer hFile, Int(nSize / 2), 0, FILE_BEGIN

End Function

Private Sub Str2Byte(sInput As String, bOutput() As Byte)

ReDim bOutput(Len(sInput) - 1)

For i = 0 To Len(sInput) - 1
    bOutput(i) = Asc(Mid(sInput, i + 1, 1))
Next i

End Sub

