VERSION 5.00
Begin VB.Form frmhospitalreport 
   Caption         =   "HOSPITAL REPORT"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   Icon            =   "frmhospitalreport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   4080
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cbohospitalid 
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
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   1575
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
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   3135
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
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Hospital ID :"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmhospitalreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmhospitalreport

'Displays:  Form will be displayed when user clicks hospital report menu on the the MDI Form

'Unload  :  When the user has generated the report for hospital details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmhospitalreport

'Functions: This form is used to generate the report for Hospital.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Retrives the Hospital ID's ***********************'

' This procedure is used to retrive the Hospital ID's from the Database.

Private Sub cbohospitalid_GotFocus()
Dim rscbo As New ADODB.Recordset
 Dim strquery As String
 rscbo.Open "Select * from tbl_HospitalInfo", Con, adOpenKeyset, adLockOptimistic
 If cbohospitalid.ListCount > 0 Then
  Set rscbo = Nothing
  Exit Sub
 Else
  For i = 0 To rscbo.RecordCount
   cbohospitalid.AddItem rscbo("HospitalID")
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
 
         'Clears the Controls to set the default values.       cbohospitalid.Text = ""
        cbohospitalid.SetFocus
End Select
End Sub

'*********************** Report Generation ***********************'

' This procedure is used to generate the Report.

Private Sub cmdreport_Click()
Dim rsnew As New ADODB.Recordset

'Code to show the Report for Hospital.
If Len(Trim(cbohospitalid.Text)) = 0 Then
  MsgBox "Please Enter the Hospital ID:", vbInformation, "Hospital Report"
Else
 rsnew.Open "select * from tbl_HospitalInfo,tbl_HospitalIssue where tbl_HospitalInfo.HospitalID = tbl_HospitalIssue.HospitalID AND tbl_HospitalInfo.HospitalID='" & cbohospitalid.Text & "'", Con, adOpenKeyset, adLockOptimistic
 If rsnew.RecordCount = 0 Then
    MsgBox "No Data To Show.", vbOKOnly + vbInformation, "Empty Database"
 Else
     Set DRhospitalreport.DataSource = rsnew
     DRhospitalreport.Show
 End If
End If
Set rsnew = Nothing
End Sub


'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

 'Set the dimension of the Form
 frmhospitalreport.Height = 3600
 frmhospitalreport.Width = 4200
 CenterForm frmhospitalreport
End Sub
