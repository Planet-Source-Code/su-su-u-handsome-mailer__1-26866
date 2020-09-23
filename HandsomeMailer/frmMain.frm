VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Handsome Mailer"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Status"
      Height          =   1215
      Left            =   960
      TabIndex        =   12
      Top             =   4080
      Width           =   3135
      Begin VB.TextBox txtStatus 
         Height          =   855
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image imgFailure 
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":0CCA
         Stretch         =   -1  'True
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Success"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   735
      End
      Begin VB.Image imgSuccess 
         Height          =   600
         Left            =   120
         Picture         =   "frmMain.frx":1594
         Stretch         =   -1  'True
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Body"
      Height          =   1575
      Left            =   960
      TabIndex        =   10
      Top             =   2400
      Width           =   3135
      Begin VB.TextBox txtMsg 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Header"
      Height          =   1815
      Left            =   960
      TabIndex        =   6
      Top             =   480
      Width           =   3135
      Begin VB.TextBox txtServer 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtSubject 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtTo 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtFrom 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Server :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "To      :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "From   :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Image imgAttach 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":245E
      ToolTipText     =   "Attach Files"
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image imgSave 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":3BE0
      ToolTipText     =   "Save Email"
      Top             =   3360
      Width           =   480
   End
   Begin VB.Image imgAddressBook 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":68DA
      ToolTipText     =   "Address Book"
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblCurrentTime 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image imgNew 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":75A4
      Stretch         =   -1  'True
      ToolTipText     =   "New Email"
      Top             =   4800
      Width           =   480
   End
   Begin VB.Image imgHelp 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":85E6
      Stretch         =   -1  'True
      ToolTipText     =   "Help"
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image imgSend 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":B658
      Stretch         =   -1  'True
      ToolTipText     =   "Send"
      Top             =   5400
      Width           =   480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   4200
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   5880
      Y2              =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   4440
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMain.frx":10E3A
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgMin 
      Height          =   255
      Left            =   3720
      Picture         =   "frmMain.frx":111F0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMain.frx":115A6
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Handsome Mailer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image bg 
      Height          =   6300
      Left            =   0
      Picture         =   "frmMain.frx":12270
      Top             =   0
      Width           =   4500
   End
   Begin VB.Line Line3 
      X1              =   840
      X2              =   3840
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   840
      X2              =   3840
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------
'* Coded by Daniel Ho                                    *
'* Code written on 27/08/2001                            *
'* This code is written for educational purposes and     *
'* personal interests only. But it can be used in any of *
'* your projects (including cut/paste for losers). And   *
'* feel to re-distribute it.                             *
'*                                                       *
'* Credits to Daniel Ho.                                 *
'*                                                       *
'* Pls post comments on the download page of this code.  *
'* This code will be upgraded soon. And you are more than*
'* welcome to post comments!                             *
'*                                                       *
'* --@_@--                                               *
'                                                        *
'     For more info on "Sending Attachments", go to      *
'     http://www.vbip.com/winsock/winsock_uucode_01.asp  *
'---------------------------------------------------------





Option Explicit
Dim DataBuffer As String
Private m_strEncodedFiles As String
'Small animation when the form is closed
Sub Startrek(frm As Form)
Dim GotoVal
Dim Gointo


    GotoVal = frm.Height / 2
    For Gointo = 1 To GotoVal
        DoEvents
        frm.Height = frm.Height - 100
        frm.Top = (Screen.Height - frm.Height) \ 2
        If frm.Height <= 500 Then Exit For
    Next Gointo
horiz:
    frm.Height = 30
    GotoVal = frm.Width / 2
    For Gointo = 1 To GotoVal
        DoEvents
        frm.Width = frm.Width - 100
        frm.Left = (Screen.Width - frm.Width) \ 2
        If frm.Width <= 2000 Then Exit For
    Next Gointo

End Sub
'Sub that does the job of waiting for a response from server
Sub WaitForResponse()
Do
DoEvents
Loop Until DataBuffer <> ""
End Sub
'Used for displaying response to the tstStatus text box
Sub DisplayStatus(data As String)
txtStatus.Text = txtStatus.Text & vbNewLine & data
End Sub

'This makes sending data easier
Public Sub SendData(data As String)

Winsock1.SendData data & vbCrLf
DisplayStatus "> " & data

End Sub





Private Sub Form_Load()
imgSuccess.Visible = False
imgFailure.Visible = False
lblStatus.Caption = "None"
'lblCurrentTime.Caption = CStr(Time) & " " & Format(Date, "Short Date")
Kill App.Path & "/AttachedFiles.tmp"
End Sub
'Resize the background
Private Sub Form_Resize()
bg.Width = frmMain.Width
bg.Height = frmMain.Height
bg.Left = 0
bg.Top = 0

End Sub
Private Sub Image2_Click()
frmAbout.Show
End Sub

Private Sub imgAddressBook_Click()
frmAddressBook.Show

End Sub

Private Sub imgAttach_Click()
frmAttachments.Show
End Sub
'Close the form by activating that animation
Private Sub imgClose_Click()
On Error Resume Next
      Call Startrek(Me)
      
Call WSACleanup
      End
End Sub

Private Sub imgHelp_Click()
frmAbout.Show
End Sub
'Minimize the form
Private Sub imgMin_Click()
frmMain.WindowState = 1
End Sub

Private Sub imgNew_Click()
txtServer.Text = ""
txtFrom.Text = ""
txtTo.Text = ""
txtSubject.Text = ""
txtMsg.Text = ""
txtStatus.Text = ""
frmAttachments.lstAttachedFiles.Clear

imgSuccess.Visible = False
imgFailure.Visible = False
lblStatus.Caption = "None"

End Sub
'Saving the Email msg in a text file
Private Sub imgSave_Click()
Dim buffer
'Filter
CommonDialog1.Filter = "Text Files (.txt) | *txt"
CommonDialog1.ShowSave
'Add up the required fields
buffer = "Date: " & Now & vbCrLf & _
         "From: " & txtFrom.Text & vbCrLf & _
         "To: " & txtTo.Text & vbCrLf & _
         "Subject: " & txtSubject.Text & vbCrLf & vbCrLf & _
         "Message: " & vbCrLf & vbCrLf & txtMsg.Text
'Create the file
Open CommonDialog1.filename & ".txt" For Output As #1

    Print #1, buffer
        
Close #1



End Sub

Private Sub imgSend_Click()
On Error Resume Next

Dim i As Integer
    '
    'Prepare attachments (if any)
    '
    For i = 0 To frmAttachments.lstAttachedFiles.ListCount - 1
        frmAttachments.lstAttachedFiles.ListIndex = i
        m_strEncodedFiles = m_strEncodedFiles & _
                         UUEncodeFile(frmAttachments.lstAttachedFiles.Text) & vbCrLf
    Next i

Winsock1.RemoteHost = txtServer.Text
Winsock1.RemotePort = 25
Winsock1.Connect
txtStatus.Text = ""

End Sub


Private Sub Timer2_Timer()
lblCurrentTime.Caption = CStr(Time)
End Sub



Private Sub txtStatus_Change()
txtStatus.SelStart = Len(txtStatus.Text)
End Sub
'Display to the user that the connection is closed
Private Sub Winsock1_Close()
DisplayStatus "<CONNECTION CLOSED>"
End Sub

Private Sub Winsock1_Connect()
Dim SentFrom As String, SendTo As String, Subject As String, FullMsg As String, TimesToSend As Integer
SentFrom = txtFrom.Text
SendTo = txtTo.Text
Subject = txtSubject.Text
FullMsg = txtMsg.Text & vbCrLf & vbCrLf & m_strEncodedFiles

m_strEncodedFiles = ""

'Tell the user that he/she is connected
    DisplayStatus "<CONNECTED>"
'Say "HI" to the server
   SendData "HELO " & Winsock1.LocalIP
   Call WaitForResponse
'Check if the "From" filed is empty
   If txtFrom.Text <> "" Then
      SendData "MAIL FROM: " & SentFrom
   Else
      SentFrom = InputBox("Please enter Your name in the <From> field", "Please complete all Fields")
   End If
'Wait for server's response
   Call WaitForResponse
'Send the recipient's email address to server
   If Right(SendTo, 1) <> "," Then SendTo = SendTo & ","
   Do
       SendData "RCPT TO: " & Mid(SendTo, 1, InStr(SendTo, ",") - 1)
       SendTo = Mid(SendTo, InStr(SendTo, ",") + 1, Len(SendTo))
   Loop Until InStr(SendTo, ",") = 0
'Wait for response
   Call WaitForResponse
'Send the actualy Data
   SendData "DATA"
   Call WaitForResponse
'Send "Subject"
   If txtSubject.Text <> "" Then
      SendData "SUBJECT: " & txtSubject.Text
   Else
      SendData "SUBJECT: " & "None"
   End If
'Wait for response
   Call WaitForResponse
'Send the full string msg and attachments (if any)
   SendData FullMsg & vbCrLf & "." & vbCrLf
   Call WaitForResponse

 
'Close the connection
    Winsock1.Close

End Sub
'Display any data arriving
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData DataBuffer
DisplayStatus DataBuffer
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Error occured"
imgFailure.Visible = True
lblStatus.Caption = "Failure"
End Sub

'Inform the user that his/her msg is delivered to the recipient
Private Sub Winsock1_SendComplete()
DisplayStatus "Sent Successfully"
imgSuccess.Visible = True
lblStatus.Caption = "Success"
End Sub

