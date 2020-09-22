VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Webcam Server"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5400
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start Server"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Stop Stream"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pause Stream"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start Stream"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Timer TmrStream 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   120
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stop Cam"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start Cam"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Text            =   "4440"
      Top             =   1200
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TmrPreview 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   3555
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Port To Listen On"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CamHwnd As Long
Private m_Jpeg      As cJpeg
Private m_FileName  As String
Private m_Image                   As New cImage

Private Sub Command1_Click()
On Error Resume Next
Winsock1.Close
Winsock1.LocalPort = Text2.Text
Winsock1.Listen
Me.Caption = "Listening on Port: " & Winsock1.LocalPort
End Sub

Private Sub Command2_Click()
On Error Resume Next
CamHwnd = capCreateCaptureWindow("CamWnd", 0, 0, 0, 320, 240, Me.hwnd, 0)
DoEvents
Call SendMessage(CamHwnd, 1034, 0, 0)
TmrPreview.Enabled = True
End Sub

Private Sub Command3_Click()
On Error Resume Next
Call SendMessage(CamHwnd, 1035, 0, 0)
Picture1.Picture = Nothing
End Sub

Private Sub Command4_Click()
On Error Resume Next
TmrStream.Enabled = True
Me.Caption = "Stream Active."
End Sub

Private Sub Command5_Click()
TmrStream.Enabled = False
Me.Caption = "Stream Paused."
End Sub

Private Sub Command6_Click()
On Error Resume Next
TmrStream.Enabled = False
Winsock1.Close
Me.Caption = "Stream Stopped."
Kill App.Path & "\temp.jsrm"
End Sub

Sub Convert(InputFile As String, OutputFile As String)
On Error Resume Next
Dim MyPic As StdPicture
Set MyPic = LoadPicture(InputFile)
Kill InputFile
Set m_Image = New cImage
m_Image.CopyStdPicture MyPic
m_Jpeg.SampleHDC m_Image.hDC, 320, 240
Kill OutputFile
m_Jpeg.SaveFile OutputFile
Set MyPic = Nothing
End Sub

Private Sub Form_Load()
On Error Resume Next
Set m_Jpeg = New cJpeg
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Call SendMessage(CamHwnd, 1035, 0, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set m_Image = Nothing
Set m_Jpeg = Nothing
If Dir(App.Path & "\temp.jsrm", vbNormal) = "temp.jsrm" Then Kill App.Path & "\temp.jsrm"
If Dir(App.Path & "\temp.bsrm", vbNormal) = "temp.bsrm" Then Kill App.Path & "\temp.bsrm"
End
End Sub

Public Sub TakeFrame()
On Error Resume Next
Dim Bjpg As String
If Winsock1.State = sckConnected Then
SavePicture Picture1.Picture, App.Path & "\temp.bsrm"
Call Convert(App.Path & "\temp.bsrm", App.Path & "\temp.jsrm")
DoEvents
Bjpg = GetFile(App.Path & "\temp.jsrm")
DoEvents
Winsock1.SendData Bjpg
End If
End Sub

Private Sub TmrPreview_Timer()
On Error Resume Next
SendMessage CamHwnd, 1084, 0, 0
SendMessage CamHwnd, 1054, 0, 0
Picture1.Picture = Clipboard.GetData
Clipboard.Clear
End Sub

Private Sub TmrStream_Timer()
On Error Resume Next
TakeFrame
End Sub

Private Sub Winsock1_Close()
On Error Resume Next
Winsock1.Close
Winsock1.LocalPort = Text2.Text
Winsock1.Listen
Me.Caption = "Connection Closed."
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Winsock1.Close
Winsock1.Accept requestID
Me.Caption = "Connection From: " & Winsock1.RemoteHostIP & " Accepted."
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim data As String
Winsock1.GetData data
If data = "CAMSTREAM" Then
Else
Winsock1.Close
End If
End Sub

