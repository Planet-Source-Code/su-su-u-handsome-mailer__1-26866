VERSION 5.00
Begin VB.Form frmAddressBook 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Address Book"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H8000000A&
      Caption         =   "Add"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3480
      Width           =   1125
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   3480
      Width           =   1000
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1000
   End
   Begin VB.ListBox lbAddressBook 
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000A&
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Dim NewEntry

NewEntry = InputBox("Please enter desired email address:", "Enter Email Address")

lbAddressBook.AddItem (NewEntry)

End Sub

Private Sub cmdRemove_Click()

Dim remove

remove = lbAddressBook.ListIndex
lbAddressBook.RemoveItem (remove)

End Sub

Private Sub cmdSave_Click()
'Save the data in the AddressBook.tmp by calling WriteList routine
Call WriteList(lbAddressBook, App.Path & "/AddressBook.tmp")
Unload Me
End Sub

Private Sub Form_Load()

On Error Resume Next
    'Call ReadList routine to input data in the "AddressBook.tmp" file
    Call ReadList(lbAddressBook, App.Path & "/AddressBook.tmp", True)
    
End Sub

Private Sub lbAddressBook_DblClick()
Dim EmailInput

EmailInput = lbAddressBook.ListIndex
frmMain.txtTo.Text = lbAddressBook.list(EmailInput)
Unload Me

End Sub
