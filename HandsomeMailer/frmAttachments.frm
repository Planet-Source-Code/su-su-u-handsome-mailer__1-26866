VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAttachments 
   BorderStyle     =   0  'None
   Caption         =   "Attachments"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommDialog2 
      Left            =   2640
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstAttachedFiles 
      Height          =   1425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label lblDetach 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detach"
      Height          =   345
      Left            =   3960
      TabIndex        =   3
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label lblOK 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O.K."
      Height          =   345
      Left            =   2280
      TabIndex        =   2
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add"
      Height          =   345
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1005
   End
End
Attribute VB_Name = "frmAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
MsgBox "Don't attach too many files! ^_^"
On Error Resume Next
    'Call ReadList routine to input data in the "AttachedFiles.tmp" file
    Call ReadList(lstAttachedFiles, App.Path & "/AttachedFiles.tmp", True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If lstAttachedFiles.ListCount = 0 Then
Kill App.Path & "/AttachedFiles.tmp"
End If
Call WriteList(lstAttachedFiles, App.Path & "/AttachedFiles.tmp")

End Sub

'Add attachments
Private Sub lblAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.BackColor = RGB(32, 32, 32)
lblAdd.ForeColor = vbWhite
With CommDialog2
    .ShowOpen
        If Len(.filename) > 0 Then
        'Add file paths to listbox
            lstAttachedFiles.AddItem .filename
        End If
    End With
End Sub

'Simply for removing attachments
Private Sub lblDetach_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
lblDetach.BackColor = RGB(32, 32, 32)
lblDetach.ForeColor = vbWhite

Dim remove
remove = lstAttachedFiles.ListIndex
lstAttachedFiles.RemoveItem (remove)

End Sub

Private Sub lblOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.BackColor = RGB(32, 32, 32)
lblOK.ForeColor = vbWhite
Me.Hide
'Call WriteList routine to save current "attached files" list
Call WriteList(lstAttachedFiles, App.Path & "/AttachedFiles.tmp")
End Sub
Private Sub LblOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblOK.BackColor = RGB(192, 192, 192)
lblOK.ForeColor = &H404040
Unload Me
End Sub
Private Sub lblAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblAdd.BackColor = RGB(192, 192, 192)
lblAdd.ForeColor = &H404040
Unload Me
End Sub
Private Sub lblDetach_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDetach.BackColor = RGB(192, 192, 192)
lblDetach.ForeColor = &H404040
Unload Me
End Sub
