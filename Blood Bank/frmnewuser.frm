VERSION 5.00
Begin VB.Form frmnewuser 
   Caption         =   "NEW USER"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
   Icon            =   "frmnewuser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   4710
   Begin VB.Frame Frame1 
      Caption         =   "ADD USER"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtUser 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2640
         TabIndex        =   0
         Top             =   555
         Width           =   1815
      End
      Begin VB.TextBox txtPwd 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtCPwd 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdSaveUser 
         Caption         =   "SAVE USER"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   4
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         Caption         =   "User Name               :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2160
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password                 :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2130
      End
      Begin VB.Label lblCPassword 
         Caption         =   "Confirm Password :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmnewuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmnewuser

'Displays:  Form will be displayed when user clicks on the New User Menu of the MDI Form

'Unload  :  When the user has inserted the New User details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmnewuser

'Functions: This form is used to add the New User.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "LOGIN")
Select Case response                           ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                              ' Unload the current form.
        
  Case vbNo:
  
  'Clear all the controls and set default values.
  
        txtUser.Text = ""
        txtPwd.Text = ""
        txtCPwd.Text = ""
        txtUser.SetFocus
   End Select
End Sub

'*********************** Save New User ***********************'

' This procedure is used to insert the new user.

Private Sub cmdSaveUser_Click()
Dim rsnew As New ADODB.Recordset
Dim strquery As String

'Validation for New user name.
If (Len(Trim(txtUser.Text)) = 0) Then
 MsgBox "First enter the User Name", vbInformation, "Username"
 txtUser.SetFocus
 Exit Sub
End If

'Validation for the password.
If (Len(Trim(txtPwd.Text)) = 0) Then
 MsgBox "First enter the Password", vbInformation, "Password"
 txtPwd.SetFocus
 Exit Sub
End If

'Validation to confirm the password.
If (Len(Trim(txtCPwd.Text)) = 0) Then
 MsgBox "Please Confirm the Password", vbInformation, "Password"
 txtCPwd.SetFocus
 Exit Sub
End If
rsnew.Open "SELECT * FROM Login WHERE Username='" & txtUser.Text & "'", Con, adOpenKeyset, adLockOptimistic
If rsnew.BOF And rsnew.EOF Then
 
If Trim(txtPwd.Text) = Trim(txtCPwd.Text) Then
 
 ' To Insert new user details.
 
 strquery = "INSERT INTO Login VALUES('" & txtUser.Text & "','" & txtPwd.Text & "')"
 Con.Execute strquery
 MsgBox "The new user is saved."
 
 ' Clear all the controls and set the default values.
 
 txtUser.Text = ""
 txtPwd.Text = ""
 txtCPwd.Text = ""
 txtUser.SetFocus
Else
 MsgBox "Re-type password again!"
 
 ' Clear all the controls and set the default values.

 txtPwd.Text = ""
 txtCPwd.Text = ""
 txtPwd.SetFocus
End If
Else
  MsgBox "This user already exists"
  
 ' Clear all the controls and set the default values.
 
  txtUser.Text = ""
  txtPwd.Text = ""
  txtCPwd.Text = ""
  txtUser.SetFocus
End If
 Set rsnew = Nothing
 
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'set the dimension for the form.

frmnewuser.Height = 4140
frmnewuser.Width = 4830
CenterForm frmnewuser
End Sub

