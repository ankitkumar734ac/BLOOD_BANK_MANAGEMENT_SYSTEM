VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPatientQuickInfo 
   Caption         =   "PATIENT LOOK UP"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frmPatientQuickInfo.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   6330
   Begin MSFlexGridLib.MSFlexGrid msflgPatIssueDetails 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   6
      FormatString    =   "SNo | Patient ID | BloodBag ID |             Product |        Issued By |  Issue Date"
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
   Begin MSFlexGridLib.MSFlexGrid msflgPatientDetails 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   8
      AllowUserResizing=   3
      FormatString    =   "SNo  |Patient Name  | Patient ID  | Address     | Phone      |  Occupation   | Gender   | Age       "
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
   Begin VB.Frame Frame1 
      Caption         =   "SEARCH BY"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmddelete 
         Caption         =   "Delete "
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
         Left            =   1920
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel "
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
         Left            =   3600
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optPatientID 
         Caption         =   "Patient ID "
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
         Left            =   3120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optName 
         Caption         =   "Name "
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
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search "
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
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtPatientID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3120
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   4695
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Patient issue"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Patient Details"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "frmPatientQuickInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmPatientQuickInfo

'Displays:  Form will be displayed when user clicks on the Patient Look Up Menu of the MDI Form

'Unload  :  When the user finish searching the Patient details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmpatientQuickInfo

'Functions: This form is used to Search for a specified Patient.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "PATIENT DETIALS")
Select Case response                      ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                         ' Unload the current form.
        
  Case vbNo:
  
  ' Clear all the controls and set the default values.
  
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
  
  End Select
End Sub

'*********************** Delete the Patient Details***********************'

' This procedure is used to come out of the form.

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
        strquery = "Insert into tbl_patientwaste values('" & msflgPatientDetails.TextMatrix(msflgPatientDetails.Row, 2) & "')"
        Con.Execute strquery
         Con.BeginTrans
         strquery = "DELETE FROM tbl_PatientInfo WHERE PatientID='" & msflgPatientDetails.TextMatrix(msflgPatientDetails.Row, 2) & "' "
         Con.Execute strquery
         For i = 0 To msflgPatientDetails.Cols - 1
         msflgPatientDetails.TextMatrix(msflgPatientDetails.Row, i) = ""
         Next i
         msflgPatientDetails.Rows = msflgPatientDetails.Rows - 1
         Con.CommitTrans
         MsgBox "Record Deleted", vbInformation + vbOKOnly, "Delete"
         msflgPatIssueDetails.Clear
    End Select
End If
End Sub

'*********************** Delete the Patient Details***********************'

' This procedure is used to come out of the form.

Private Sub cmdSearch_Click()
Dim strsql As String
Dim srno As Integer
Dim rssearch As New ADODB.Recordset
On Error GoTo errPara

'Validation for Pateint ID
If optName.Value = False And optPatientID.Value = True Then
     If (Len(Trim(txtPatientID.Text)) = 0) Then
       MsgBox "First enter the patient ID", vbInformation, "PatientID"
       txtPatientID.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT  *  FROM tbl_PatientInfo WHERE PatientID='" & txtPatientID.Text & "'", Con, adOpenKeyset, adLockOptimistic
     End If
     
'Validation for Pateint Name
ElseIf optPatientID.Value = False And optName.Value = True Then
     If (Len(Trim(txtName.Text)) = 0) Then
       MsgBox "First enter the Name", vbInformation, "Name"
       txtName.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT * FROM tbl_PatientInfo WHERE PatientName='" & txtName.Text & "'", Con, adOpenKeyset, adLockOptimistic
     End If
Else
   MsgBox "First choose and enter the search data.", vbExclamation, "Patient Details"
   Exit Sub
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
    
    'Assigns the values from the database to the corresponding controls.

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
    
    'Clear all the control and set default values.
    
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

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'Sets the dimension for the form.

frmPatientQuickInfo.Height = 6360
frmPatientQuickInfo.Width = 6450
CenterForm frmPatientQuickInfo
End Sub

'*********************** Shows the Patient Details in Grid Control ***********************'

' This procedure is used to show the Patient Details in Grid Control.

Private Sub msflgPatientDetails_Click()

    Dim srno As Integer
    Dim strsql As String
    Dim rsSearchDetails As New ADODB.Recordset
    srno = 1
    rsSearchDetails.Open "SELECT * FROM tbl_PatientIssue WHERE PatientId='" & msflgPatientDetails.TextMatrix(msflgPatientDetails.Row, 2) & "'", Con, adOpenKeyset, adLockOptimistic
    
    msflgPatIssueDetails.Clear
    msflgPatIssueDetails.Rows = 2
    msflgPatIssueDetails.FormatString = "SNo | Patient ID | BloodBag ID |       Product | Issued By |  Issue Date"
    If rsSearchDetails.BOF And rsSearchDetails.EOF Then MsgBox "No records found for this criteria."
        
    
    While Not rsSearchDetails.EOF
        
    'Assigns the values from the database to the corresponding controls.
      
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

'*********************** Select the Patient Name ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optName_Click()

'Sets the default values.

optPatientID.Enabled = False
txtPatientID.Enabled = False
txtName.Enabled = True

'Clear all the controls.
txtPatientID.Text = ""
txtName.Text = ""
txtName.SetFocus
       
End Sub

'*********************** Select the Patient ID ***********************'

' This procedure is used to set the values of the Control wh

Private Sub optPatientID_Click()

 ' Sets the default values for all the control.
 
txtName.Enabled = False
txtPatientID.Enabled = True
txtPatientID.SetFocus
optName.Enabled = False

'Clear all the control.

txtPatientID.Text = ""
txtName.Text = ""


End Sub
