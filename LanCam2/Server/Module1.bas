Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal WindowName As String, ByVal Style As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Parent As Long, ByVal ID As Long) As Long
Public Declare Function ConvertBMPtoJPG Lib "ImageUtils.dll" (ByVal InputFile As String, ByVal OutputFile As String, ByVal OverWrite As Boolean, ByVal JPGCompression As Integer, ByVal SaveBMP As Boolean) As Integer

Public Function SetFile(ByVal StrFileName As String, ByVal FileData As String)
On Error Resume Next
Dim H1 As Long
H1 = FreeFile
Kill StrFileName
Open StrFileName For Binary Access Write As #H1
Put #H1, , FileData
Close #H1
End Function

Public Function GetFile(ByVal StrFileName As String) As String
On Error Resume Next
Dim H1 As Long
Dim GetFil As String
H1 = FreeFile
Open StrFileName For Binary As #H1
GetFil = Space$(LOF(H1))
Get #H1, , GetFil
Close #H1
GetFile = GetFil
End Function

Public Function SetDLL()
On Error Resume Next
Dim H1 As Long, StrFileName As String
Dim DllBuffer() As Byte
DllBuffer = LoadResData(101, "CUSTOM")
StrFileName = App.Path & "\ImageUtils.dll"
H1 = FreeFile
Open StrFileName For Binary Access Write As #H1
Put #H1, , DllBuffer
Close #H1
End Function
