VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFDUser 
   Caption         =   "Find/Delete User"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frmFDUser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmFDUser.frx":25CA
   ScaleHeight     =   4320
   ScaleWidth      =   4560
   Begin VB.Frame Frame1 
      Caption         =   "FIND/DELETE USER"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin MSFlexGridLib.MSFlexGrid msflgDetails 
         Height          =   1455
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdFindUser 
         Caption         =   "FIND USER "
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
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CommandButton cmdDeleteUser 
         BackColor       =   &H80000013&
         Caption         =   "DELETE USER "
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
         Left            =   2280
         TabIndex        =   3
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CommandButton cmdChangePass 
         Caption         =   "CHANGE PASSWORD "
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
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   2895
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
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   3720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmFDUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmFDUser

'Displays:  Form will be displayed when user clicks on the Find/Delete User Menu of the MDI Form

'Unload  :  When the user has accessed the user details , user can click the Cancel command button and choose the Yes option to terminate from the Form frmFDUser

'Functions: This form is used to find and delete the user from the database.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "FIND DELETE USER")
Select Case response                         ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                            ' Unload the current form
  Case vbNo:
   'Clears the Controls to set the default values.
        msflgDetails.Clear
        msflgDetails.Rows = 2
  End Select
End Sub


'*********************** Change the Password of the User ***********************'

' This procedure is used to change the password of the existing user.

Private Sub cmdChangePass_Click()
Dim rsnew As New ADODB.Recordset
If Len(msflgDetails.TextMatrix(msflgDetails.Row, 1)) = 0 Then
 MsgBox "First Find the user for which you want to change password.", vbExclamation, "FDUser"
Else
rsnew.Open "Select * from Login where Username='" & msflgDetails.TextMatrix(msflgDetails.Row, 1) & "' ", Con, adOpenKeyset, adLockOptimistic
If rsnew.BOF And rsnew.EOF Then
MsgBox "This user doesnot exist." & vbCrLf & "Find any existing user.", vbInformation, "FIND DELETE USER"
Else
 strUsername = msflgDetails.TextMatrix(msflgDetails.Row, 1)
 
 ' Clear the controls and set default values.
 
 msflgDetails.Clear
 frmChangePass.Show
 frmChangePass.txtUser.Text = strUsername
 frmChangePass.txtPwd.SetFocus
 Unload Me
End If
Set rsnew = Nothing
End If
End Sub


'*********************** Delete the User ***********************'

' This procedure is used to delete the existing user.

Private Sub cmdDeleteUser_Click()
Dim strsql As String
Dim response As String
Dim i As Integer
If Len(msflgDetails.TextMatrix(msflgDetails.Row, 1)) = 0 Then
 MsgBox "Select Username"
Else
Con.BeginTrans
MsgBox "User is '" & msflgDetails.TextMatrix(msflgDetails.Row, 1) & "'", vbInformation, "FIND DELETE USER"
response = MsgBox("Do you want to delete?", vbQuestion + vbYesNo, "FIND DELETE USER")
Select Case response
 Case vbYes:
 
 ' Delete the existing user from the database
    
    strsql = "DELETE  FROM Login WHERE Username='" & msflgDetails.TextMatrix(msflgDetails.Row, 1) & "' "
    Con.Execute strsql
    Con.CommitTrans
    MsgBox "User Deleted"
    For i = 0 To msflgDetails.Cols - 1
         msflgDetails.TextMatrix(msflgDetails.Row, i) = ""
    Next i
End Select
msflgDetails.Rows = msflgDetails.Rows - 1
If msflgDetails.Rows <= 1 Then
 msflgDetails.Rows = 2
End If
End If
End Sub

'*********************** Find the User ***********************'

' This procedure is used to find the existing user.

Private Sub cmdFindUser_Click()
Dim rssearch As New ADODB.Recordset
Dim srno As Integer
Dim strresponse As String
strresponse = InputBox("Enter the username you want to find", "FDUser")
  rssearch.Open "SELECT * FROM Login WHERE Username='" & strresponse & "'", Con, adOpenDynamic, adLockOptimistic

    srno = 1
    msflgDetails.Clear
    msflgDetails.Rows = 2
    msflgDetails.FormatString = "SRNO |USERNAME | PASSWORD"
    msflgDetails.ColWidth(2) = 2000
    
    If rssearch.BOF And rssearch.EOF Then MsgBox "No records found for this criteria."
        
        'Assigns the values from the database to the corresponding controls.
    
    While Not rssearch.EOF
        msflgDetails.TextMatrix(srno, 0) = srno
        msflgDetails.TextMatrix(srno, 1) = rssearch("Username") & ""
        msflgDetails.TextMatrix(srno, 2) = rssearch("Password") & ""
        rssearch.MoveNext
        srno = srno + 1
        If msflgDetails.Rows = srno Then msflgDetails.Rows = msflgDetails.Rows + 1
    Wend
    rssearch.Close
    Set rssearch = Nothing
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'set the dimension for the form

frmFDUser.Height = 4680
frmFDUser.Width = 4830
CenterForm frmFDUser
End Sub


