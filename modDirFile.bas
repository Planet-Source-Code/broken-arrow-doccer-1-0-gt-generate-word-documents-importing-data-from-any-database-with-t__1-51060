Attribute VB_Name = "modDirFile"
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

Public Function ShortPath(FullPath As String) As String
Dim sPath As String * 256

GetShortPathName FullPath, sPath, Len(sPath)

ShortPath = Trim(Replace(sPath, Chr(0), ""))
End Function

Public Function FullPath(FileNameWithFullPath As String) As String
Dim FilePart As String, Buffer As String * 255

FilePart = GetFileFromFullPath(FileNameWithFullPath)

GetFullPathName FileNameWithFullPath, Len(Buffer), Buffer, FilePart

FullPath = Trim(Replace(Buffer, Chr(0), ""))
End Function

Public Function GetFileFromFullPath(Full_Path As String) As String 'Extract only filename from full
Dim AltStr1 As String, AltStr2 As String, a As Long         'path & filename
For a = Len(Full_Path) To 1 Step -1
    AltStr1 = AltStr1 & Mid(Full_Path, a, 1)
Next
AltStr1 = Left(AltStr1, InStr(AltStr1, "\") - 1)
For a = Len(AltStr1) To 1 Step -1
    AltStr2 = AltStr2 & Mid(AltStr1, a, 1)
Next
GetFileFromFullPath = AltStr2
End Function

Public Function CheckPath(PathString As String, Optional AddSlash As Boolean = True) As String
If PathString = "" Then
    CheckPath = CheckPath(App.Path, True)
    Exit Function
End If

If AddSlash And Right(PathString, 1) <> "\" Then CheckPath = PathString & "\"
If Not AddSlash And Right(PathString, 1) = "\" Then CheckPath = Left(PathString, Len(PathString) - 1)
End Function

