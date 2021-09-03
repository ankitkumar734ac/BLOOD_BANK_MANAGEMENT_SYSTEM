VERSION 5.00
Begin VB.Form frmChangePass 
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   Icon            =   "frmChangePass.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   4230
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4215
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
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
         Left            =   2280
         TabIndex        =   8
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Left            =   600
         TabIndex        =   7
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtNPwd 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
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
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   1695
      End
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
         Height          =   405
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label lblCPassword 
         Caption         =   "New Password :"
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
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password          :"
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1710
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         Caption         =   "User Name        :"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmChangePass

'Displays:  Form will be displayed when user clicks on the Change Password Menu of the MDI Form

'Unload  :  When the user has accessed the User details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmChangePass

'Functions: This form is used to Change password.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "LOGIN")
Select Case response                 ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                    ' Unloads the Current Form
        
  Case vbNo:
  
        'Clears the Controls to set the default values.       txtUser.Text = ""
        txtPwd.Text = ""
        txtNPwd.Text = ""
        txtUser.SetFocus
  
 End Select
End Sub

'*********************** Inserts the Values into Database ***********************'

' This procedure is used to insert the password for the Existing user.

Private Sub cmdSave_Click()
Dim rschange As New ADODB.Recordset
Dim str As String

'Validates the Username
If (Len(Trim(txtUser.Text)) = 0) Then
 MsgBox "First enter the User Name", vbInformation, "Username"
 txtUser.SetFocus
 Exit Sub
End If

'Validates the Users Old Password
If (Len(Trim(txtPwd.Text)) = 0) Then
 MsgBox "First enter the Password", vbInformation, "Password"
 txtPwd.SetFocus
 Exit Sub
End If

'Validates the users new Password
If (Len(Trim(txtNPwd.Text)) = 0) Then
  MsgBox "Enter the new Password", vbInformation, "Confirm Password"
  txtNPwd.SetFocus
  Exit Sub
 End If
rschange.Open "SELECT * FROM Login WHERE Username='" & txtUser.Text & "'", Con, adOpenKeyset, adLockOptimistic

If rschange.BOF = True And rschange.EOF = True Then
 MsgBox "Invalid User", vbExclamation, "CHANGE PASSWORD"
 txtUser.Text = ""
 txtPwd.Text = ""
 txtNPwd.Text = ""
 txtUser.SetFocus
 Set rschange = Nothing
 Exit Sub
ElseIf txtUser.Text = rschange("Username") And txtPwd.Text = rschange("Password") Then
 ' rschange.Fields("Password").Value = "'" & txtNPwd.Text & "'"
 
  'Update the Old Password of the Existing User with the new Password
   rschange.Update "Password", txtNPwd.Text
  
   'str = "Update Login set Password='" & txtNPwd.Text & "',Username='" & txtUser.Text & "'"
  MsgBox "User Updated with the new password", vbInformation, "CHANGE PASSWORD"
End If

'Clears the Controls to set the default values.'Clears the Controls to set the default values.'Clears the Controls to set the default values.'Clears the Controls to set the default values.

txtUser.Text = ""
 txtPwd.Text = ""
 txtNPwd.Text = ""
 txtUser.SetFocus
Set rschange = Nothing
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'Sets the dimensions for the Form

frmChangePass.Height = 3915
frmChangePass.Width = 4350
CenterForm frmChangePass
End Sub




