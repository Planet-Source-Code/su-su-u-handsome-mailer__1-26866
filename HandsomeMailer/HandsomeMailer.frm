VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   StartUpPosition =   3  'Windows Default
   Begin VB.Image bg 
      Height          =   6300
      Left            =   0
      Picture         =   "HandsomeMailer.frx":0000
      Top             =   0
      Width           =   4500
   End
   Begin VB.Line Line3 
      X1              =   56
      X2              =   256
      Y1              =   320
      Y2              =   320
   End
   Begin VB.Line Line2 
      X1              =   256
      X2              =   256
      Y1              =   0
      Y2              =   320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   56
      X2              =   256
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Resize()
bg.Width = frmMain.Width
bg.Height = frmMain.Height
bg.Left = 0
bg.Top = 0

End Sub

