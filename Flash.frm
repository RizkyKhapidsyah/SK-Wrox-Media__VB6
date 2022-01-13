VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1440
      Top             =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long




Private Sub Timer1_Timer()
    
    Dim nReturnValue As Integer
    nReturnValue = FlashWindow(Form1.hWnd, True)

End Sub
