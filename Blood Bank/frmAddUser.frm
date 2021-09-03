VERSION 5.00
Begin VB.Form frmAddUser 
   Caption         =   "SAMARPAN BLOOD BANK"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   6180
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Caption         =   "ADD USER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton CMDcANCEL 
         Caption         =   "CANCEL"
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdCreateUser 
         Caption         =   "CREATE USER"
         Height          =   495
         Left            =   1440
         TabIndex        =   7
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txtCPassword 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   3000
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCPassword 
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         TabIndex        =   6
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   5
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCreateUser_Click()
Dim strsql As String

If Validatedata = False Then
    Exit Sub
    
ElseIf txtPassword.Text <> txtCPassword.Text Then
txtCPassword.Text = ""
MsgBox "Please Reconfirm Password"
txtCPassword.SetFocus
Exit Sub
Else
Con.BeginTrans
strsql = "INSERT INTO tbl_Login(Username,Password) VALUES ('" & txtUserName.Text & "','" & txtPassword.Text & "') "
Con.Execute strsql
Con.CommitTrans
End If
MsgBox "User Created"

End Sub


Public Function Validatedata() As Boolean
Validatedata = True

If Len(Trim(txtUserName.Text)) = 0 Then
    MsgBox "Please Enter User name"
    Validatedata = False
    txtUserName.SetFocus
    Exit Function
End If



If Len(Trim(txtPassword.Text)) = 0 Then
    MsgBox "Please Enter Password"
    Validatedata = False
    txtPassword.SetFocus
    Exit Function
End If

If Len(Trim(txtCPassword.Text)) = 0 Then
    MsgBox "Please Confirm Password"
    Validatedata = False
    txtCPassword.SetFocus
    Exit Function
End If

End Function

Private Sub Form_Load()
frmAddUser.Height = 4425
frmAddUser.Width = 6300
CenterForm frmAddUser
End Sub
