VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1092
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   2796
   LinkTopic       =   "Form1"
   ScaleHeight     =   1092
   ScaleWidth      =   2796
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   720
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()

    Dim Multimedia As New Mmedia
    With CommonDialog1
        .Filter = "WaveAudio (*.wav)|*.wav|Midi (*.mid)|*.mid|Video files (*.avi)|*.avi"
        .FilterIndex = 0
        .ShowOpen
    End With

    If CommonDialog1.Filename <> "" Then
        Multimedia.mmOpen CommonDialog1.Filename
        Multimedia.mmPlay
    End If



End Sub


