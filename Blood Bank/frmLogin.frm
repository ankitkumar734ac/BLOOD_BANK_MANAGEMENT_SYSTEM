VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   Caption         =   "SECURITY LOGIN"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmLogin.frx":25CA
   ScaleHeight     =   6675
   ScaleWidth      =   10695
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
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
         Height          =   360
         Left            =   2400
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LOGIN"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtPwd 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "SECURITY LOGIN"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   600
         TabIndex        =   7
         Top             =   0
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   480
         TabIndex        =   6
         Top             =   840
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Password    :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   480
         TabIndex        =   5
         Top             =   1560
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmLogin

'Displays:  Form will be displayed when user clicks on the Login Menu of the MDI Form

'Unload  :  When the user clicks the Cancel command button and choose the Yes option to terminate from the Form frmLogin

'Functions: This form is used to get authorization to access the software.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is not a Child of a MDI form.

Option Explicit

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "LOGIN")
Select Case response                        ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                           ' Unload the current form.
        
  Case vbNo:
  
  'Clear the controls and assign the default values.
        txtUserName.Text = ""
        txtPwd.Text = ""
        txtUserName.SetFocus
   End Select
End Sub


'*********************** Login ***********************'

' This procedure is used to authentically Log into the software.

Private Sub cmdLogin_Click()
Dim rsLogin As New ADODB.Recordset
Dim storeit As String

' Validation for user name
If Len(Trim(txtUserName.Text)) = 0 Then
    MsgBox "Please enter UserName."
    txtUserName.SetFocus
    Exit Sub
End If

'Validation for Password
If Len(Trim(txtPwd.Text)) = 0 Then
    MsgBox "Please enter password."
    txtPwd.SetFocus
    Exit Sub
End If

rsLogin.Open "SELECT * FROM Login WHERE Username='" & txtUserName.Text & "' ", Con, adOpenKeyset, adLockOptimistic

If rsLogin.BOF = True And rsLogin.EOF = True Then
  MsgBox "Invalid User"
  
  'Clear all the controls and set default values.
  
  txtUserName.Text = ""
  txtPwd.Text = ""
  txtUserName.SetFocus
  Exit Sub
ElseIf txtUserName.Text = rsLogin("Username") And txtPwd.Text = rsLogin("Password") Then

' Enabled all menu controls.

    MDIForm1.mnuDonorDetails.Enabled = True
    MDIForm1.mnuDonorSearch.Enabled = True
    MDIForm1.mnuLogout.Enabled = True
    MDIForm1.mnuMainLogin.Enabled = False
    MDIForm1.mnuHospitalInfo.Enabled = True
    MDIForm1.mnuHospitalLookUp.Enabled = True
    MDIForm1.mnuBloodBagLookUp.Enabled = True
    MDIForm1.mnuBloodBagIssue.Enabled = True
    MDIForm1.mnuPatientInfo.Enabled = True
    MDIForm1.mnuPatientLookUp.Enabled = True
    MDIForm1.mnuDonorReport.Enabled = True
    MDIForm1.mnuHospitalReport.Enabled = True
    MDIForm1.mnuPatient.Enabled = True
    MDIForm1.mnuBloodBag.Enabled = True
    MDIForm1.mnuChange.Enabled = True
    MDIForm1.mnuNewUser.Enabled = True
    MDIForm1.mnudonorInfo.Enabled = True
    MDIForm1.mnuPatientReport.Enabled = True
    MDIForm1.mnuHospitals.Enabled = True
    MDIForm1.mnuReports.Enabled = True
    MDIForm1.mnuFindDeleteUser.Enabled = True
    MDIForm1.mnuHelp.Enabled = True
    MDIForm1.mnuAbout.Enabled = True
    MDIForm1.mnuDatabaseBackUp.Enabled = True
    Unload Me
    CheckExpiry

Else

'Clear all the controls and set default values.
    
    txtUserName.Text = ""
    txtPwd.Text = ""
    txtPwd.SetFocus
    txtUserName.SetFocus
    MsgBox "Either User Name or Password is Incorrect."
    Exit Sub
End If
Set rsLogin = Nothing
End Sub


'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'set the dimension for the form.

    frmLogin.Height = 4095
    frmLogin.Width = 4650
    CenterForm frmLogin
    
End Sub




