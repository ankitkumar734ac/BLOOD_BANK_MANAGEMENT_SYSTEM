VERSION 5.00
Begin VB.Form frmpatientreport 
   Caption         =   "PATIENT REPORT"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3885
   Icon            =   "frmpatientreport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   3885
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cbopatientid 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdreport 
         Caption         =   "REPORT GENERATION"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   3255
      End
      Begin VB.CommandButton cmdcancel 
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
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Patient ID :"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmpatientreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmpatientreport

'Displays:  Form will be displayed when user clicks patient report menu on the the MDI Form

'Unload  :  When the user has generated the report for patient details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmpatientreport

'Functions: This form is used to generate the report for patient.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit


'*********************** Retrives the Patient ID's ***********************'

' This procedure is used to retrive the Patient ID's from the Database.

Private Sub cbopatientid_GotFocus()
Dim rscbo As New ADODB.Recordset
 Dim strquery As String
 rscbo.Open "Select * from tbl_PatientInfo", Con, adOpenKeyset, adLockOptimistic
 If cbopatientid.ListCount > 0 Then
  Set rscbo = Nothing
  Exit Sub
 Else
  For i = 0 To rscbo.RecordCount
   cbopatientid.AddItem rscbo("PatientID")
   rscbo.MoveNext
   If rscbo.EOF Then
    Exit For
   End If
  Next i
End If
Set rscbo = Nothing
End Sub


'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit?", vbQuestion + vbYesNo, "CANCEL")
Select Case response                 ' Provides the MsgBox response as the expression to match for the Case Structure.
 Case vbYes:
       Unload Me                      ' Unloads the Current Form
 Case vbNo:
 
  'Clears the Controls to set the default values.
        cbopatientid.Text = ""
        cbopatientid.SetFocus

End Select
End Sub

'*********************** Report Generation ***********************'

' This procedure is used to generate the Report.

Private Sub cmdreport_Click()
Dim rsnew As New ADODB.Recordset

'Code to show the Report for Hospital.
If Len(Trim(cbopatientid.Text)) = 0 Then
  MsgBox "Please Enter the Patient ID:", vbInformation, "Patient Report"
Else
 rsnew.Open "select * from tbl_PatientInfo,tbl_PatientIssue where tbl_PatientInfo.PatientID = tbl_PatientIssue.Patientid AND tbl_PatientInfo.PatientID='" & cbopatientid.Text & "'", Con, adOpenKeyset, adLockOptimistic
 If rsnew.RecordCount = 0 Then
    MsgBox "No Data To Show.", vbOKOnly + vbInformation, "Empty Database"
 Else
     Set DRpatientreport.DataSource = rsnew
     DRpatientreport.Show
 End If
End If
Set rsnew = Nothing
End Sub


'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'Set the dimension for the Form
frmpatientreport.Height = 3015
frmpatientreport.Width = 4005
CenterForm frmpatientreport
'txtpatientid.SetFocus
End Sub

