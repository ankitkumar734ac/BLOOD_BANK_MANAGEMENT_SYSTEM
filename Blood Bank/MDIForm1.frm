VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "BLOOD BANK MANAGEMENT SYSTEM"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   975
   ClientWidth     =   20250
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":25CA
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   360
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   255
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   20190
      TabIndex        =   1
      Top             =   0
      Width           =   20250
      Begin VB.TextBox Text2 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   4440
         TabIndex        =   2
         Text            =   "BLOOD  FOR HUMANS  COMES ONLY FROM HUMANS"
         Top             =   1320
         Width           =   12975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   39
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2655
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   $"MDIForm1.frx":2F795
         Top             =   0
         Width           =   15135
      End
      Begin VB.Image Image1 
         Height          =   2430
         Left            =   0
         Picture         =   "MDIForm1.frx":2F7EB
         Top             =   0
         Width           =   2730
      End
      Begin VB.Image Image2 
         Height          =   2415
         Left            =   17880
         Picture         =   "MDIForm1.frx":30487
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2610
      End
   End
   Begin VB.Menu mnuLogin 
      Caption         =   "LOGIN"
      Begin VB.Menu mnuMainLogin 
         Caption         =   "&LOGIN"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuChange 
         Caption         =   "CHANGEPASSWORD"
      End
      Begin VB.Menu mnuNewUser 
         Caption         =   "NEWUSER"
      End
      Begin VB.Menu mnuFindDeleteUser 
         Caption         =   "FIND/DELETEUSER"
      End
      Begin VB.Menu mnuLogout 
         Caption         =   "LOG&OUT"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&CLOSE"
      End
   End
   Begin VB.Menu mnudonorInfo 
      Caption         =   "DONOR &INFORMATION"
      Begin VB.Menu mnuDonorDetails 
         Caption         =   "DONOR &DETAILS"
      End
      Begin VB.Menu mnuDonorSearch 
         Caption         =   "DONOR &SEARCH"
      End
      Begin VB.Menu mnuDonorList 
         Caption         =   "DONOR LIST"
      End
   End
   Begin VB.Menu mnuHospitals 
      Caption         =   "HOSPITAL"
      Begin VB.Menu mnuHospitalInfo 
         Caption         =   "&HOSPITAL"
      End
      Begin VB.Menu mnuHospitalLookUp 
         Caption         =   "HOSPITAL &LOOK UP"
      End
      Begin VB.Menu mnuHospitalList 
         Caption         =   "HOSPITAL LIST"
      End
   End
   Begin VB.Menu mnuBloodBag 
      Caption         =   "BLOOD BAG"
      Begin VB.Menu mnuBloodBagLookUp 
         Caption         =   "BLOOD BA&G LOOK UP"
      End
      Begin VB.Menu mnuBloodBagIssue 
         Caption         =   "BLOOD BAG &ISSUE"
      End
      Begin VB.Menu mnuBloodBagList 
         Caption         =   "BLOOD BAG LIST"
      End
   End
   Begin VB.Menu mnuPatient 
      Caption         =   "PATIENT"
      Begin VB.Menu mnuPatientInfo 
         Caption         =   "PATIENT INFO"
      End
      Begin VB.Menu mnuPatientLookUp 
         Caption         =   "PATIENT LOOK UP"
      End
      Begin VB.Menu mnuPatientList 
         Caption         =   "PATIENT LIST"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "REPORTS"
      Begin VB.Menu mnuDonorReport 
         Caption         =   "DONOR REPORT"
      End
      Begin VB.Menu mnuHospitalReport 
         Caption         =   "HOSPITAL REPORT"
      End
      Begin VB.Menu mnuPatientReport 
         Caption         =   "PATIENT REPORT"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "HELP"
      Begin VB.Menu mnuAbout 
         Caption         =   "ABOUT"
      End
      Begin VB.Menu mnuDatabaseBackUp 
         Caption         =   "DATABASE BACKUP"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------Form Summary-----------------------------------------------------------------------------------

'Form Name: MDIForm1.frm

'Displays:  'It Displays After Splash Screen but is accessible only after succesful authentication

'Unloaded: When Appliation Ends

'Functions:

    '1. This is a Parent or Container  form of all Forms in a Project.
    '2. It is the main interface to access the system functions.
    '3. Almost all the forms are accessible through the Menus.
    '4. All the Forms except frmSplash are the Childs of MDI form.

'Note:

    '1. The connection is established at module level.
    '2. Database back up is provided.

'Database Connection:

    ' ADODB Record set object are created.
    ' Also the form level ADODB Recordsets are created at each forms.

'************************************************************************

Option Explicit

'*********************** Loads the MDIForm and performs the given settings ***********************'

Private Sub MDIForm_Load()
    mnuDonorSearch.Enabled = False
    mnuDonorDetails.Enabled = False
    mnuHospitalInfo.Enabled = False
    mnuHospitalLookUp.Enabled = False
    mnuBloodBagLookUp.Enabled = False
    mnuBloodBagIssue.Enabled = False
    mnuPatientInfo.Enabled = False
    mnuPatientLookUp.Enabled = False
    mnuHospitalReport.Enabled = False
    mnuDonorReport.Enabled = False
    mnuPatient.Enabled = False
    mnuBloodBag.Enabled = False
    mnuChange.Enabled = False
    mnuNewUser.Enabled = False
    mnudonorInfo.Enabled = False
    mnuHospitals.Enabled = False
    mnuReports.Enabled = False
    mnuFindDeleteUser.Enabled = False
    mnuPatientReport.Enabled = False
    mnuHelp.Enabled = False
    mnuAbout.Enabled = False
    mnuDatabaseBackUp.Enabled = False
End Sub

Private Sub mnuBloodBagList_Click()
frmBloodBagList.Show

End Sub

Private Sub mnuDonorList_Click()
frmDonorList.Show

End Sub

Private Sub mnuHospitalList_Click()
frmHospitalList.Show

End Sub

'*********************** LOGIN MENU ***********************

' Login Menu --> Login

Private Sub mnuMainLogin_Click()
 frmLogin.Show                    ' Loads and Shows the Login Form for authenticated access
End Sub

' Login Menu --> Change Password

Private Sub mnuChange_Click()
 frmChangePass.Show               ' Loads and Shows the Form for changing the password
End Sub

' Login Menu --> New User

Private Sub mnuNewUser_Click()
 frmnewuser.Show                  ' Loads and Shows the Form to add the New User
End Sub

' Login Menu --> Find/Delete User

Private Sub mnuFindDeleteUser_Click()
 frmFDUser.Show                   ' Loads and Shows to Find / Deletes the user
End Sub

'Login Menu --> Log Out

Private Sub mnuLogout_Click()
 Dim response As Integer
 response = MsgBox("Do you surely want to LOG OUT ?", vbYesNo + vbQuestion, "SECURITY lOG OUT")
  If response = vbYes Then
    mnuLogout.Enabled = False
    mnuMainLogin.Enabled = True
    mnuDonorSearch.Enabled = False
    mnuDonorDetails.Enabled = False
    mnuHospitalInfo.Enabled = False
    mnuHospitalLookUp.Enabled = False
    mnuPatientInfo.Enabled = False
    mnuPatientLookUp.Enabled = False
    mnuBloodBagLookUp.Enabled = False
    mnuBloodBagIssue.Enabled = False
    mnuDonorReport.Enabled = False
    mnuHospitalReport.Enabled = False
    mnuPatient.Enabled = False
    mnuBloodBag.Enabled = False
    mnuChange.Enabled = False
    mnuNewUser.Enabled = False
    mnudonorInfo.Enabled = False
    mnuHospitals.Enabled = False
    mnuPatientReport.Enabled = False
    mnuReports.Enabled = False
    mnuFindDeleteUser.Enabled = False
    mnuHelp.Enabled = False
    mnuAbout.Enabled = False
    mnuDatabaseBackUp.Enabled = False
  Else
    Exit Sub
  End If
    End Sub

' Login Menu --> Close

Private Sub mnuClose_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "BLOOD BANK MANAGEMENT SYSTEM")
Select Case response
  Case vbYes:
        Unload Me
        End
End Select

End Sub

'*********************** DONONR DETAILS MENU ***********************'

' Donor Details --> Donor Details

Private Sub mnuDonorDetails_Click()
 frmDonorDetails.Show              ' Loads and shows the form that tackles with the Donor Information
End Sub

' Donor Details --> Donor Search

Private Sub mnuDonorSearch_Click()
 frmDonorSearch.Show              ' Loads and shows the form that provides the quick access for the Donor Details
End Sub

'*********************** HOSPITAL MENU ***********************'

' Hospital --> Hospital Information

Private Sub mnuHospitalInfo_Click()
 frmHospitalInfo.Show             ' Loads and shows the form that tackles with he hospital Information
End Sub

' Hospital --> Hospital Look Up

Private Sub mnuHospitalLookUp_Click()
 frmHospitalLookUp.Show           ' Loads and shows the form that provides the quick access for the hospital Information
End Sub

'*********************** BLOOD BAG DETAILS MENU ***********************'

'Private Sub mnuBloodBagDetails_Click()
'frmBloodBagDetails.Show
'End Sub

' Blood Bag Details --> Blood Bag Look Up

Private Sub mnuBloodBagLookUp_Click()
 frmBBLookUp.Show                 ' Loads and shows the form that provides the quick access for the Blood Bag Details
End Sub

' Blood Bag Details --> Blood Bag Issue

Private Sub mnuBloodBagIssue_Click()


aaa = InputBox("Enter any valid bloodbagID:", "Blood Bag ID")
If aaa = "" Then
   MsgBox "Enter BloodBagID", vbCritical
Exit Sub
ElseIf IsNumeric(aaa) Then
 frmBloodBagIssue.Show ' Loads and shows the form that issues the blood bags for the needed ones.
 Else
 MsgBox "Enter Valid BloodBagID", vbCritical
 
End If
End Sub

'*********************** PATIENT MENU ***********************'

'Patient --> Patient Details

Private Sub mnuPatientInfo_Click()
 frmPatient.Show                 ' Loads and shows the form that tackles with the patient information
End Sub

Private Sub mnuPatientList_Click()
frmPatientList.Show

End Sub

'Patient --> Patient Look Up

Private Sub mnuPatientLookUp_Click()
 frmPatientQuickInfo.Show        ' Loads and shows the form that provides the quick access for the Patient Details
End Sub

'*********************** REPORT MENU ***********************'

' Report --> Donor Report

Private Sub mnuDonorReport_Click()
 frmdonorreport.Show             ' Provides the report for the Donor
End Sub

' Report --> Hospital Report

Private Sub mnuHospitalReport_Click()
 frmhospitalreport.Show           ' Provides the report for the Hospital Report
End Sub


' Report --> Patient Report

Private Sub mnuPatientReport_Click()
 frmpatientreport.Show            ' Provides the report for the Patient Report
End Sub

'Private Sub mnuBloodBagReport_Click()
'frmBloodBagReport.Show
'End Sub

'*********************** HELP MENU ***********************'

' Help --> About

Private Sub mnuAbout_Click()
 frmAbout.Show                  ' Provides the form that give information about the Software
End Sub

' Help --> Database Back Up

Private Sub mnuDatabaseBackUp_Click() ' Provides the database back up
On Error GoTo CancelErr
    
        Dim Bck As Object, file As Object
        
        Set Bck = CreateObject("Scripting.filesystemobject")
        Me.Dialog1.DialogTitle = "Database Back up"
        Me.Dialog1.Filter = "Backup files (*.bck) |*.bck|"
        Me.Dialog1.ShowSave
            If Me.Dialog1.FileName = "" Then
                MsgBox "Please Enter The File Name."
            Else
                Set file = Bck.GetFile(App.Path & "\Blood.mdb")
                file.Copy Me.Dialog1.FileName
                MsgBox "Back up Completed", vbInformation
                
              ' username = frmLogin.txtUserid1.Text
    'FPath = App.Path
                
        ' Open FPath + "\LOG FILES\csms.log" For Append As #1
        ' Write #1, UserName, "Takes Back Up Of Database.", Format$(Now, "hh:mm:ss dd/mm/yyyy")
        ' Close #1
         
            End If
            Exit Sub
            
CancelErr:
        MsgBox Err.Description & "", vbExclamation, "Attention"
End Sub







