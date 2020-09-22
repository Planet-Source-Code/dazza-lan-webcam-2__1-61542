Attribute VB_Name = "Module1"
Option Explicit

Public Function SetFile(ByVal StrFileName As String, ByVal FileData As String)
On Error Resume Next
Dim H1 As Long
H1 = FreeFile
Kill StrFileName
Open StrFileName For Binary Access Write As #H1
Put #H1, , FileData
Close #H1
End Function
