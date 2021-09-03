VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPatientLookUp 
   Caption         =   "PATIENT LOOK UP"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   6150
   Begin MSFlexGridLib.MSFlexGrid msflgPatIssueDetails 
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   4200
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   6
      FormatString    =   "SNo | Patient ID | BloodBag ID |             Product |        Issued By |  Issue Date"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msflgPatientDetails 
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   8
      AllowUserResizing=   3
      FormatString    =   "SNo  |Patient Name  | Patient ID  | Address     | Phone      |  Occupation   | Gender   | Age       "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "SEARCH BY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmddelete 
         Caption         =   "Delete "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optPatientID 
         Caption         =   "Patient ID "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optName 
         Caption         =   "Name "
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search "
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtPatientID 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Patient issue"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Patient Details"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "frmPatientLookUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "PATIENT DETIALS")
Select Case (response)
  Case vbYes:
        Unload Me
        
  Case vbNo:
        txtName.Text = ""
        txtName.Enabled = True
        txtPatientID.Text = ""
        txtPatientID.Enabled = True
        optName.Value = False
        optPatientID.Value = False
        optName.Enabled = True
        optPatientID.Enabled = True
        msflgPatientDetails.Clear
        msflgPatientDetails.Rows = 2
        msflgPatIssueDetails.Clear
        msflgPatIssueDetails.Rows = 2
  Case Default
        Unload Me
     
 End Select
End Sub

Private Sub cmddelete_Click()
Dim strquery As String
Dim strresponse As String
Dim i As Integer

If Len(msflgPatientDetails.TextMatrix(msflgPatientDetails.Row, 1)) = 0 Then
  MsgBox "First search  for the Patient You want to Delete", vbExclamation, "PATIENT SEARCH"
Else
strresponse = MsgBox("Are you sure you want to delete", vbQuestion + vbYesNo)
   Select Case strresponse
        Case vbYes:
        strquery = "Insert into tbl_hodpitalwaste values('" & msflgHospitalDetails.TextMatrix(msflgHospitalDetails.Row, 1) & "')"
        Con.Execute strquery
         Con.BeginTrans
         strquery = "DELETE FROM tbl_PatientInfo WHERE PatientID='" & msflgPatientDetails.TextMatrix(msflgPatientDetails.Row, 1) & "' "
         Con.Execute strquery
         For i = 1 To msflgPatientDetails.Cols - 1
         msflgPatientDetails.TextMatrix(msflgPatientDetails.Row, i) = ""
         Next i
         msflgPatientDetails.Rows = msflgPatientDetails.Rows - 1
         Con.CommitTrans
         MsgBox "Record Deleted", vbInformation + vbOKOnly, "Delete"
         msflgPatIssueDetails.Clear
    End Select
End If
End Sub

Private Sub cmdSearch_Click()
Dim strSql As String
Dim srno As Integer
Dim rssearch As New ADODB.Recordset
On Error GoTo errPara
If optPatientID.Value = True Then
     If (Len(Trim(txtPatientID.Text)) = 0) Then
       MsgBox "First enter the patient ID", vbInformation, "PatientID"
       txtPatientID.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT  *  FROM tbl_PatientInfo WHERE PatientID='" & txtPatientID.Text & "'", Con, adOpenKeyset, adLockOptimistic
     End If
ElseIf optName.Value = True Then
     If (Len(Trim(txtName.Text)) = 0) Then
       MsgBox "First enter the Name", vbInformation, "Name"
       txtName.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT * FROM tbl_PatientInfo WHERE Patient='" & txtName.Text & "'", Con, adOpenKeyset, adLockOptimistic
     End If
End If

srno = 1
msflgPatientDetails.Clear
msflgPatientDetails.Rows = 2
msflgPatIssueDetails.Clear
msflgPatIssueDetails.Rows = 2
msflgPatientDetails.FormatString = " SNo  |Patient Name  | Patient ID  | Address     | Phone      |  Occupation   | Gender   | Age       "
msflgPatientDetails.ColWidth(2) = 2000
    
If rssearch.BOF And rssearch.EOF Then
   MsgBox "No records found for this criteria."
Else
    While Not rssearch.EOF
        msflgPatientDetails.TextMatrix(srno, 0) = srno
        msflgPatientDetails.TextMatrix(srno, 1) = rssearch("PatientName") & " " & rssearch("PatientSurname")
        msflgPatientDetails.TextMatrix(srno, 2) = rssearch("PatientID") & ""
        msflgPatientDetails.TextMatrix(srno, 3) = rssearch("Address") & ""
        msflgPatientDetails.TextMatrix(srno, 4) = rssearch("Phone") & ""
        msflgPatientDetails.TextMatrix(srno, 5) = rssearch("Occupation") & ""
        msflgPatientDetails.TextMatrix(srno, 6) = rssearch("Gender") & ""
        msflgPatientDetails.TextMatrix(srno, 7) = rssearch("Age") & ""
        rssearch.MoveNext
        srno = srno + 1
        If msflgPatientDetails.Rows = srno Then msflgPatientDetails.Rows = msflgPatientDetails.Rows + 1
    Wend
 End If
    rssearch.Close
    Set rssearch = Nothing
        txtName.Enabled = True
        txtName.Text = ""
        txtPatientID.Text = ""
        txtPatientID.Enabled = True
        optName.Value = False
        optPatientID.Value = False
        optName.Enabled = True
        optPatientID.Enabled = True

Exit Sub
errPara:
    If rssearch.State = 1 Then
        rssearch.Close
        Set rssearch = Nothing
    End If
    MsgBox "Error in code"
    
    
End Sub

Private Sub Form_Load()
CenterForm frmPatientLookUP
frmPatientLookUP.Height = 6135
frmPatientLookUP.Width = 6270

End Sub

Private Sub msflgPatientDetails_Click()

    Dim srno As Integer
    Dim strSql As String
    Dim rsSearchDetails As New ADODB.Recordset
    srno = 1
    rsSearchDetails.Open "SELECT * FROM tbl_PatientIssue WHERE PatientId='" & msflgPatientDetails.TextMatrix(msflgPatientDetails.Row, 2) & "'", Con, adOpenKeyset, adLockOptimistic
    
    msflgPatIssueDetails.Clear
    msflgPatIssueDetails.Rows = 2
    msflgPatIssueDetails.FormatString = "SNo | Patient ID | BloodBag ID |       Product | Issued By |  Issue Date"
    If rsSearchDetails.BOF And rsSearchDetails.EOF Then MsgBox "No records found for this criteria."
        
    
    While Not rsSearchDetails.EOF
        msflgPatIssueDetails.TextMatrix(srno, 0) = srno
        msflgPatIssueDetails.TextMatrix(srno, 1) = rsSearchDetails("PatientId") & ""
        msflgPatIssueDetails.TextMatrix(srno, 2) = rsSearchDetails("BloodBagID") & ""
        msflgPatIssueDetails.TextMatrix(srno, 3) = rsSearchDetails("Product") & ""
        msflgPatIssueDetails.TextMatrix(srno, 4) = rsSearchDetails("IssuedBy") & ""
        msflgPatIssueDetails.TextMatrix(srno, 5) = rsSearchDetails("IssueDate")
        
        rsSearchDetails.MoveNext
        srno = srno + 1
        If msflgPatIssueDetails.Rows = srno Then msflgPatIssueDetails.Rows = msflgPatIssueDetails.Rows + 1
    Wend
    rsSearchDetails.Close
    Set rsSearchDetails = Nothing
End Sub

Private Sub optName_Click()
optPatientID.Enabled = False
txtPatientID.Text = ""
txtPatientID.Enabled = False
txtName.Text = Clear
txtName.Enabled = True
txtName.SetFocus
       
End Sub

Private Sub optPatientID_Click()
txtPatientID.Enabled = True
txtPatientID.Text = Clear
txtPatientID.SetFocus
optName.Enabled = False
txtName.Text = ""
txtName.Enabled = False

End Sub
