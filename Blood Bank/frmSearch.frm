VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSearch 
   Caption         =   "DONOR SEARCH"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   8295
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "SEARCH BY"
      Height          =   1335
      Left            =   720
      TabIndex        =   7
      Top             =   0
      Width           =   6375
      Begin VB.ComboBox cmbRh 
         Height          =   315
         ItemData        =   "frmSearch.frx":0000
         Left            =   5160
         List            =   "frmSearch.frx":000A
         TabIndex        =   14
         Top             =   720
         Width           =   615
      End
      Begin VB.ComboBox cmbBloodGroup 
         Height          =   315
         ItemData        =   "frmSearch.frx":0018
         Left            =   4440
         List            =   "frmSearch.frx":0028
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtDonorName 
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtDonorNumber 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optBloodGroup 
         Caption         =   "BloodGroup"
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optDonorName 
         Caption         =   "DonorName"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optDonorNumber 
         Caption         =   "Donor Number"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDonorDetails 
      Caption         =   "Donor Details"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid msflgTestResults 
      Height          =   1455
      Left            =   1080
      TabIndex        =   2
      Top             =   4560
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   8
      AllowUserResizing=   3
      FormatString    =   "SNo | VDRL | HBsAG | M.P. | HCV | HIV | Entered By | Entered Date"
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid msflgDetails 
      Height          =   1335
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   5
      AllowUserResizing=   3
      FormatString    =   "SNo|Donor Id | Donor Name | Blood Group | Donor Age"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "DONOR DETAILS"
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "TEST RESULTS"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim strsql As String
If Len(msflgDetails.TextMatrix(msflgDetails.Row, 1)) = 0 Then
MsgBox "select donor ID"
Exit Sub
End If

Con.BeginTrans
MsgBox "no is" & msflgDetails.TextMatrix(msflgDetails.Row, 1)
strsql = "DELETE  FROM tbl_Donor WHERE DonorID='" & msflgDetails.TextMatrix(msflgDetails.Row, 1) & "' "
'MsgBox strsql

Con.Execute strsql
Con.CommitTrans
MsgBox "record deleted"
 msflgDetails.Clear
 msflgTestResults.Clear
End Sub

Private Sub cmdDonorDetails_Click()
If Len(msflgDetails.TextMatrix(msflgDetails.Row, 1)) = 0 Then Exit Sub
    lngDonorId = msflgDetails.TextMatrix(msflgDetails.Row, 1)
    frmDonorDetails.Show
'    MsgBox intDonorId'
End Sub

Private Sub cmdSearch_Click()
Dim strsql As String
Dim srno As Integer
Dim rssearch As New ADODB.Recordset

On Error GoTo errPara
    If Len(Trim(txtDonorNumber)) <> 0 Then
        strsql = "SELECT DonorID,DonorName,DonorSName,DonorMName,BloodGroup,Age,RH FROM tbl_Donor WHERE DonorID='" & txtDonorNumber & "'"
    ElseIf Len(Trim(txtDonorName)) <> 0 Then
        strsql = "SELECT DonorID,DonorName,DonorSName,DonorMName,BloodGroup,Age,RH  FROM tbl_Donor WHERE DonorName like '" & txtDonorName & "%'"
    ElseIf Len(Trim(cmbBloodGroup.Text)) <> 0 Then
       If Len(Trim(cmbRh.Text)) = 0 Then
           MsgBox "Please select RH Value."
       End If
        strsql = "SELECT DonorID,DonorName,DonorSName,DonorMName,BloodGroup,Age,RH  FROM tbl_Donor WHERE BloodGroup = '" & cmbBloodGroup.Text & "' and RH='" & cmbRh & "'"
        
    End If
    If Len(strsql) = 0 Then
        MsgBox "Please enter search data."
        Exit Sub
    End If
'    MsgBox strsql
    rssearch.Open strsql, Con, adOpenDynamic, adLockOptimistic
'    rssearch.MoveFirst
    srno = 1
    msflgDetails.Clear
    msflgDetails.Rows = 2
    msflgTestResults.Clear
    msflgTestResults.Rows = 2
    msflgDetails.FormatString = "SNo |Donor Id | Donor Name | Blood Group | Donor Age"
    msflgDetails.ColWidth(2) = 2000
    
    If rssearch.BOF And rssearch.EOF Then MsgBox "No records found for this criteria."
        
    While Not rssearch.EOF
        msflgDetails.TextMatrix(srno, 0) = srno
        msflgDetails.TextMatrix(srno, 1) = rssearch("donorID") & ""
        msflgDetails.TextMatrix(srno, 2) = rssearch("DonorSName") & " " & rssearch("DonorName") & " " & rssearch("DonorMName")
        msflgDetails.TextMatrix(srno, 3) = rssearch("BloodGroup") & " " & rssearch("RH") & ""
        msflgDetails.TextMatrix(srno, 4) = rssearch("Age") & ""
        rssearch.MoveNext
        srno = srno + 1
        If msflgDetails.Rows = srno Then msflgDetails.Rows = msflgDetails.Rows + 1
    Wend
    rssearch.Close
    Set rssearch = Nothing
Exit Sub
errPara:
    If rssearch.State = 1 Then
        rssearch.Close
        Set rssearch = Nothing
    End If
    MsgBox "Error in code"
End Sub



Private Sub Form_Load()
    frmSearch.Height = 6525
    frmSearch.Width = 8010
    CenterForm frmSearch
    txtDonorName.Enabled = False
    cmbBloodGroup.Enabled = False
    txtDonorNumber.Enabled = False
    cmbRh.Enabled = False
    
End Sub

Private Sub msflgDetails_Click()
    Dim srno As Integer
    Dim strsql As String
    Dim rsSearchDetails As New Recordset
    
    srno = 1
    strsql = "SELECT * FROM tbl_DonorTestResults WHERE DonorID='" & msflgDetails.TextMatrix(msflgDetails.Row, 4) & "'"
    rsSearchDetails.Open "SELECT * FROM tbl_DonorTestResults WHERE DonorID='" & msflgDetails.TextMatrix(msflgDetails.Row, 1) & "'", Con
    
    msflgTestResults.Clear
    msflgTestResults.Rows = 2
    msflgTestResults.FormatString = "SNo | VDRL | HBsAG | M.P. | HCV | HIV | Entered By | Entered Date"
    
    While Not rsSearchDetails.EOF
        msflgTestResults.TextMatrix(srno, 0) = srno
        msflgTestResults.TextMatrix(srno, 1) = IIf(rsSearchDetails("TestVDRL") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 2) = IIf(rsSearchDetails("TestHBsAG") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 3) = IIf(rsSearchDetails("TestMP") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 4) = IIf(rsSearchDetails("TestHCV") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 5) = IIf(rsSearchDetails("TestHIV") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 6) = rsSearchDetails("EnteredBy")
        msflgTestResults.TextMatrix(srno, 7) = rsSearchDetails("EnteredDate") & ""
        rsSearchDetails.MoveNext
        srno = srno + 1
        If msflgTestResults.Rows = srno Then msflgTestResults.Rows = msflgTestResults.Rows + 1
    Wend
    rsSearchDetails.Close
    Set rsSearchDetails = Nothing
End Sub

Private Sub optBloodGroup_Click()
txtDonorName.Enabled = False
txtDonorNumber.Enabled = False
cmbBloodGroup.Enabled = True
cmbRh.Enabled = True
txtDonorName.Text = ""
txtDonorNumber.Text = ""
cmbBloodGroup.SetFocus
End Sub

Private Sub optDonorName_Click()
txtDonorNumber.Enabled = False
cmbBloodGroup.Enabled = False
txtDonorName.Enabled = True
cmbRh.Enabled = False
txtDonorNumber.Text = ""
cmbBloodGroup.Text = ""
cmbRh.Text = ""
txtDonorName.SetFocus
End Sub

Private Sub optDonorNumber_Click()
txtDonorName.Enabled = False
cmbBloodGroup.Enabled = False
txtDonorNumber.Enabled = True
cmbRh.Enabled = False
cmbBloodGroup.Text = ""
cmbRh.Text = ""
txtDonorName.Text = ""
txtDonorNumber.SetFocus
End Sub
