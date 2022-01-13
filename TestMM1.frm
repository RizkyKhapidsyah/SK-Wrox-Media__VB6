VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1584
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3132
   LinkTopic       =   "Form1"
   ScaleHeight     =   1584
   ScaleWidth      =   3132
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2640
      Top             =   600
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2172
      _ExtentX        =   3831
      _ExtentY        =   445
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "TestMM1.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   1080
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327680
      FontSize        =   1.17491e-38
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load and play a file"
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2172
   End
   Begin VB.Label Label1 
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2172
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Multimedia As New Mmedia

Private Sub Command1_Click()

    With CommonDialog1
        .Filter = "WaveAudio (*.wav)|*.wav|Midi (*.mid)|*.mid|Video files (*.avi)|*.avi"
        .FilterIndex = 0
        .ShowOpen
    End With

    If CommonDialog1.Filename <> "" Then
        Multimedia.Wait = False
        Multimedia.mmOpen CommonDialog1.Filename
        ProgressBar1.Value = 0
        ProgressBar1.Max = Multimedia.Length
        Timer1.Enabled = True
        Multimedia.mmPlay
    End If



End Sub

Private Sub Timer1_Timer()
   
   ProgressBar1.Value = Multimedia.Position
   Label1 = "Status: " & Multimedia.Status
   If ProgressBar1.Value = ProgressBar1.Max Then
      Multimedia.mmClose
      Timer1.Enabled = False
   End If

End Sub
