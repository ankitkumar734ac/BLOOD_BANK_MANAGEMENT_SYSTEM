VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDonorSearch 
   Caption         =   "DONOR SEARCH"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "frmDonorSearch.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   6780
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE "
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
      Left            =   4080
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
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
      Height          =   1695
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox cmbRh 
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
         Height          =   420
         ItemData        =   "frmDonorSearch.frx":25CA
         Left            =   5640
         List            =   "frmDonorSearch.frx":25D4
         TabIndex        =   14
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbBloodGroup 
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
         Height          =   420
         ItemData        =   "frmDonorSearch.frx":25E2
         Left            =   4560
         List            =   "frmDonorSearch.frx":25F2
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtDonorName 
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
         Height          =   405
         Left            =   2520
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtDonorNumber 
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
         Height          =   405
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton optBloodGroup 
         Caption         =   "BloodGroup "
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
         Left            =   4560
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optDonorName 
         Caption         =   "DonorName "
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2520
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optDonorNumber 
         Caption         =   "Donor Number "
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&CLOSE "
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
      Left            =   5400
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDonorDetails 
      Caption         =   "DONOR DETAILS "
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid msflgTestResults 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   8
      AllowUserResizing=   3
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
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&SEARCH "
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
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid msflgDetails 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   17
      AllowUserResizing=   3
      FormatString    =   $"frmDonorSearch.frx":2603
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
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   6615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "DONOR DETAILS "
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
      Top             =   2520
      Width           =   2160
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "TEST RESULTS "
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
      TabIndex        =   3
      Top             =   4080
      Width           =   1875
   End
End
Attribute VB_Name = "frmDonorSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmDonorSearch

'Displays:  Form will be displayed when user clicks on the Donor Search Menu of the MDI Form

'Unload  :  When the user finish searching the donor details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmDonorSearch

'Functions: This form is used to Search for a specified donor.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.


Option Explicit


'*********************** Close ***********************'

' This procedure is used to terminate the form.

Private Sub cmdclose_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "DONOR DETIALS")
Select Case response                           ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                              ' Unload the current form.
        
  Case vbNo:
        
    'Clear all the controls and set the default values.
    
        txtDonorNumber.Text = ""
        txtDonorName.Text = ""
        cmbBloodGroup.Text = ""
        cmbRh.Text = ""
        optDonorNumber.Value = False
        optDonorName.Value = False
        optBloodGroup.Value = False
        optDonorNumber.Enabled = True
        optDonorName.Enabled = True
        optBloodGroup.Enabled = True
        msflgDetails.Clear
        msflgDetails.Rows = 2
        msflgTestResults.Clear
        msflgTestResults.Rows = 2
   End Select
End Sub

'*********************** Delete the Donor Details ***********************'

' This procedure is used to delete the existing Donor Details.

Private Sub cmddelete_Click()
Dim strquery As String
Dim strresponse As String
Dim i As Integer

If Len(msflgDetails.TextMatrix(msflgDetails.Row, 1)) = 0 Then
  MsgBox "First search  for the Donor You want to Delete", vbExclamation, "DONOR SEARCH"
  Exit Sub
End If

strresponse = MsgBox("Are you sure you want to delete", vbQuestion + vbYesNo)
   Select Case strresponse
        Case vbYes:
        strquery = "Insert into tbl_donorwaste values('" & msflgDetails.TextMatrix(msflgDetails.Row, 1) & "')"
        Con.Execute strquery
        Con.BeginTrans
        
        'Delete the Donor Details from the database.
        
        strquery = "DELETE  FROM tbl_Donor WHERE DonorID='" & msflgDetails.TextMatrix(msflgDetails.Row, 1) & "' "
         For i = 0 To msflgDetails.Cols - 1
         msflgDetails.TextMatrix(msflgDetails.Row, i) = ""
         Next i
         Con.Execute strquery
         Con.CommitTrans
         MsgBox "Record Deleted", vbInformation + vbOKOnly, "Delete"
         msflgTestResults.Clear
         msflgDetails.Rows = msflgDetails.Rows - 1
    End Select
  
End Sub

'***********************  Donor Details ***********************'

' This procedure is used to show all Donor Details.

Private Sub cmdDonorDetails_Click()
  Dim rsnew As New ADODB.Recordset
  Dim i As Integer
  Dim srno As Integer
  srno = 1
  rsnew.Open "Select * from tbl_Donor ", Con, adOpenKeyset, adLockOptimistic
  rsnew.MoveFirst
  For i = 0 To rsnew.RecordCount - 1
  
  'Assigns the values from the database to the corresponding controls.
  
        msflgDetails.TextMatrix(srno, 0) = srno
        msflgDetails.TextMatrix(srno, 1) = rsnew("DonorID") & ""
        msflgDetails.TextMatrix(srno, 2) = rsnew("DonorSName") & ""
        msflgDetails.TextMatrix(srno, 3) = rsnew("DonorName") & ""
        msflgDetails.TextMatrix(srno, 4) = rsnew("DonorMName") & ""
        msflgDetails.TextMatrix(srno, 5) = rsnew("Address") & ""
        msflgDetails.TextMatrix(srno, 6) = rsnew("PhoneRes") & ""
        msflgDetails.TextMatrix(srno, 7) = rsnew("PhoneOff") & ""
        msflgDetails.TextMatrix(srno, 8) = rsnew("Mobile") & ""
        msflgDetails.TextMatrix(srno, 9) = rsnew("Gender") & ""
        msflgDetails.TextMatrix(srno, 10) = rsnew("Age") & ""
        msflgDetails.TextMatrix(srno, 11) = rsnew("MaritalStatus") & ""
        msflgDetails.TextMatrix(srno, 12) = rsnew("BloodGroup") & ""
        msflgDetails.TextMatrix(srno, 13) = rsnew("RH") & ""
        msflgDetails.TextMatrix(srno, 14) = rsnew("Occupation") & ""
        msflgDetails.TextMatrix(srno, 15) = rsnew("DonorType") & ""
        msflgDetails.TextMatrix(srno, 16) = rsnew("LastDonateDate") & ""
        rsnew.MoveNext
        srno = srno + 1
        If msflgDetails.Rows = srno Then
         msflgDetails.Rows = msflgDetails.Rows + 1
        End If
   Next i
 Set rsnew = Nothing
      
End Sub

'***********************  Search the Donor Details ***********************'

' This procedure is used to search Donor Details for desired search criteria.

Private Sub cmdSearch_Click()
Dim strsql As String
Dim srno As Integer
Dim rssearch As New ADODB.Recordset
Dim i As Integer
On Error GoTo errPara

   'Validation for Donor Name
    
    If optDonorName.Value = False And optBloodGroup.Value = False And optDonorNumber.Value = True Then
     If (Len(Trim(txtDonorNumber.Text)) = 0) Then
       MsgBox "First enter the Donor ID", vbInformation, "DonorID"
       txtDonorNumber.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT  *  FROM tbl_Donor WHERE DonorID='" & txtDonorNumber.Text & "'", Con, adOpenKeyset, adLockOptimistic
     End If
     
    'Validation for Donor Number
    
    ElseIf optDonorNumber.Value = False And optBloodGroup.Value = False And optDonorName.Value = True Then
     If (Len(Trim(txtDonorName.Text)) = 0) Then
       MsgBox "First enter the Donor Name", vbInformation, "DonorName"
       txtDonorName.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT  *  FROM tbl_Donor WHERE DonorName='" & txtDonorName.Text & "'", Con, adOpenKeyset, adLockOptimistic
     End If
     
    'Validation for Blood Group
    
    ElseIf optDonorNumber.Value = False And optDonorName.Value = False And optBloodGroup.Value = True Then
       If Len(Trim(cmbBloodGroup.Text)) <> 0 Then
         If Len(Trim(cmbRh.Text)) <> 0 Then
           rssearch.Open "SELECT  *  FROM tbl_Donor WHERE BloodGroup='" & cmbBloodGroup.Text & "' And RH='" & cmbRh.Text & "'", Con, adOpenKeyset, adLockOptimistic
         Else
           MsgBox "Please select RH Value."
           Exit Sub
         End If
        Else
           MsgBox "Please select BloodGroup Value."
           Exit Sub
        End If
   Else
       MsgBox "First choose and enter the search data.", vbExclamation, "Donor Search"
       Exit Sub
      End If
    
    srno = 1
    msflgDetails.Clear
    msflgDetails.Rows = 2
    msflgTestResults.Clear
    msflgTestResults.Rows = 2
    msflgDetails.FormatString = "srno |DonorID | DonorSName | DonorName | DonorMName | Address | PhoneRes | PhoneOff | Mobile | Gender | Age | MaritalStatus | BloodGroup | RH | Occupation | DonorType | LastDonateDate"
    'msflgDetails.ColWidth(2) = 2000
    
    If rssearch.BOF And rssearch.EOF Then
     MsgBox "No records found for this criteria."
    Else
        If rssearch.RecordCount > 0 Then
        
        For i = 0 To rssearch.RecordCount - 1
        
     'Assigns the values from the database to the corresponding controls.
     
        msflgDetails.TextMatrix(srno, 0) = srno
        msflgDetails.TextMatrix(srno, 1) = rssearch("DonorID") & ""
        msflgDetails.TextMatrix(srno, 2) = rssearch("DonorSName") & ""
        msflgDetails.TextMatrix(srno, 3) = rssearch("DonorName") & ""
        msflgDetails.TextMatrix(srno, 4) = rssearch("DonorMName") & ""
        msflgDetails.TextMatrix(srno, 5) = rssearch("Address") & ""
        msflgDetails.TextMatrix(srno, 6) = rssearch("PhoneRes") & ""
        msflgDetails.TextMatrix(srno, 7) = rssearch("PhoneOff") & ""
        msflgDetails.TextMatrix(srno, 8) = rssearch("Mobile") & ""
        msflgDetails.TextMatrix(srno, 9) = rssearch("Gender") & ""
        msflgDetails.TextMatrix(srno, 10) = rssearch("Age") & ""
        msflgDetails.TextMatrix(srno, 11) = rssearch("MaritalStatus") & ""
        msflgDetails.TextMatrix(srno, 12) = rssearch("BloodGroup") & ""
        msflgDetails.TextMatrix(srno, 13) = rssearch("RH") & ""
        msflgDetails.TextMatrix(srno, 14) = rssearch("Occupation") & ""
        msflgDetails.TextMatrix(srno, 15) = rssearch("DonorType") & ""
        msflgDetails.TextMatrix(srno, 16) = rssearch("LastDonateDate") & ""
        srno = srno + 1
        If msflgDetails.Rows = srno Then
         msflgDetails.Rows = msflgDetails.Rows + 1
        End If
        rssearch.MoveNext
        If rssearch.EOF Then
          rssearch.MoveFirst
        End If
   Next i
  End If
  End If
    Set rssearch = Nothing
        
  'Clear all the controls and set default values.
  
        txtDonorNumber.Text = ""
        txtDonorName.Text = ""
        cmbBloodGroup.Text = ""
        cmbRh.Text = ""
        txtDonorNumber.Enabled = False
        txtDonorName.Enabled = False
        cmbBloodGroup.Enabled = False
        cmbRh.Enabled = False
        optDonorNumber.Value = False
        optDonorName.Value = False
        optBloodGroup.Value = False
        optDonorNumber.Enabled = True
        optDonorName.Enabled = True
        optBloodGroup.Enabled = True
 Exit Sub

'Code for Error Handling
errPara:
    If rssearch.State = 1 Then
        'rsSearch.Close
        Set rssearch = Nothing
    End If
    MsgBox "Error in code"
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()
    
   'Sets the dimension for the form.
   
    frmDonorSearch.Height = 6150
    frmDonorSearch.Width = 6960
    CenterForm frmDonorSearch
    
   'Sets default value for each control.
   
    txtDonorName.Enabled = False
    cmbBloodGroup.Enabled = False
    txtDonorNumber.Enabled = False
    cmbRh.Enabled = False
    msflgDetails.FormatString = " DonorID | DonorSName | DonorName | DonorMName | Address | PhoneRes | PhoneOff | Mobile | Gender | Age | MaritalStatus | BloodGroup | RH | Occupation | DonorType | LastDonateDate"
    msflgTestResults.FormatString = " DonorID | TestVDRL | TestHBsAG | TestMP | TestHCV | TestHIV | EnteredBy | EnteredDate"
End Sub

'*********************** Shows the Donor Details in Grid Control ***********************'

' This procedure is used to show the Donor Details in Grid Control.

 Private Sub msflgDetails_Click()
    Dim srno As Integer
    Dim strsql As String
    Dim rsSearchDetails As New ADODB.Recordset
    
    srno = 1
    strsql = "SELECT * FROM tbl_DonorTestResults WHERE DonorID='" & msflgDetails.TextMatrix(msflgDetails.Row, 4) & "'"
    rsSearchDetails.Open "SELECT * FROM tbl_DonorTestResults WHERE DonorID='" & msflgDetails.TextMatrix(msflgDetails.Row, 1) & "'", Con
    
    'msflgTestResults.Clear
    'msflgTestResults.Rows = 2
    msflgTestResults.FormatString = "SNo | DonorID | VDRL | HBsAG | M.P. | HCV | HIV | Entered By | Entered Date"
    
    While Not rsSearchDetails.EOF
    
       'Assigns the values from the database to the corresponding controls.
      
        msflgTestResults.TextMatrix(srno, 0) = srno
        msflgTestResults.TextMatrix(srno, 1) = rsSearchDetails("DonorID") & ""
        msflgTestResults.TextMatrix(srno, 2) = IIf(rsSearchDetails("TestVDRL") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 3) = IIf(rsSearchDetails("TestHBsAG") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 4) = IIf(rsSearchDetails("TestMP") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 5) = IIf(rsSearchDetails("TestHCV") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 6) = IIf(rsSearchDetails("TestHIV") = True, "+VE", "-VE")
        msflgTestResults.TextMatrix(srno, 7) = rsSearchDetails("EnteredBy") & ""
        msflgTestResults.TextMatrix(srno, 8) = rsSearchDetails("EnteredDate") & ""
        rsSearchDetails.MoveNext
        srno = srno + 1
        If msflgTestResults.Rows = srno Then msflgTestResults.Rows = msflgTestResults.Rows + 1
    Wend
    rsSearchDetails.Close
    Set rsSearchDetails = Nothing
End Sub

'*********************** Select the Blood Group ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optBloodGroup_Click()
  
 'Clear the control and set the default value.
 
  txtDonorName.Text = ""
  txtDonorNumber.Text = ""
  txtDonorName.Enabled = False
  cmbBloodGroup.Enabled = True
  txtDonorNumber.Enabled = False
  cmbRh.Enabled = True
  optDonorName.Value = False
  optDonorNumber.Value = False
  cmbBloodGroup.Text = ""
  cmbRh.Text = ""
  optDonorNumber.Enabled = False
  optDonorName.Enabled = False
  cmbBloodGroup.SetFocus
End Sub

'*********************** Select the Donor Name ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.


Private Sub optDonorName_Click()
 
 'Clear the control and set the default value.
  
  txtDonorName.Text = ""
  txtDonorNumber.Text = ""
  cmbBloodGroup.Text = ""
  cmbRh.Text = ""
  txtDonorName.Enabled = True
  cmbBloodGroup.Enabled = False
  txtDonorNumber.Enabled = False
  cmbRh.Enabled = False
  optDonorNumber.Value = False
  optBloodGroup.Value = False
  optDonorNumber.Enabled = False
  optBloodGroup.Enabled = False
  txtDonorName.SetFocus
End Sub

'*********************** Select the Donor Number ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optDonorNumber_Click()

 'Clear the control and set the default value.

  txtDonorName.Text = ""
  txtDonorNumber.Text = ""
  cmbBloodGroup.Text = ""
  cmbRh.Text = ""
  txtDonorName.Enabled = False
  cmbBloodGroup.Enabled = False
  txtDonorNumber.Enabled = True
  cmbRh.Enabled = False
  optDonorName.Value = False
  optBloodGroup.Value = False
  optDonorName.Enabled = False
  optBloodGroup.Enabled = False
  txtDonorNumber.SetFocus
End Sub

