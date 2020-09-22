VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Webcam Client"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Text            =   "4440"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   600
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   3555
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iData As String

Private Sub Command1_Click()
On Error Resume Next
Winsock1.Close
Winsock1.RemoteHost = Text1.Text
Winsock1.RemotePort = Text2.Text
Winsock1.Connect
End Sub

Private Sub Command2_Click()
On Error Resume Next
Winsock1.Close
Me.Caption = "Connection Closed."
Picture1.Picture = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Dir(App.Path & "\temp.srm", vbNormal) = "temp.srm" Then Kill App.Path & "\temp.srm"
End Sub

Private Sub Winsock1_Close()
On Error Resume Next
Winsock1.Close
Me.Caption = "Connection Closed."
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
Winsock1.SendData "CAMSTREAM"
Me.Caption = "Connection Active."
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim data As String
Winsock1.GetData data
iData = iData + data
DoEvents
SetImg iData
End Sub

Sub SetImg(Imdata As String)
On Error Resume Next
SetFile App.Path & "\temp.srm", Imdata
DoEvents
Picture1.Picture = LoadPicture(App.Path & "\temp.srm")
iData = ""
End Sub
