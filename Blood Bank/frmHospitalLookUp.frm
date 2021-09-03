VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHospitalLookUp 
   Caption         =   "HOSPITAL LOOK UP"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "frmHospitalLookUp.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   5985
   Begin MSFlexGridLib.MSFlexGrid msflgHospitalDetails 
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   6
      AllowUserResizing=   3
      FormatString    =   "SNo | HospitalID |Hospital Name|Hospital Address |Phone1           | Area                   "
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
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
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
         Left            =   1440
         TabIndex        =   11
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdBBEntry 
         Caption         =   "BloodBagEntry "
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
         Left            =   3720
         TabIndex        =   10
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtHospitalID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2280
         TabIndex        =   9
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton optHospitalID 
         Caption         =   "HospitalID "
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
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optArea 
         Caption         =   "Area "
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
         Left            =   3480
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optHospitalName 
         Caption         =   "HospitalName "
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
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtArea 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3480
         TabIndex        =   4
         ToolTipText     =   "Enter Hospital Area"
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtHospitalName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2175
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
         Left            =   2520
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
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
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmHospitalLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmHospitalLookUp

'Displays:  Form will be displayed when user clicks on the Hospital Look Up Menu of the MDI Form

'Unload  :  When the user finish searching the Hospital details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmDonorSearch

'Functions: This form is used to Search for a specified Hospital.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit
Private Sub cmdBBEntry_Click()
If Len(msflgHospitalDetails.TextMatrix(msflgHospitalDetails.Row, 1)) = 0 Then
   MsgBox "No rows selected."
Else
    strHospitalID = msflgHospitalDetails.TextMatrix(msflgHospitalDetails.Row, 1)
    frmBBFromHos.Show
    frmBBFromHos.txtHospitalID = strHospitalID
    Unload Me
End If
End Sub

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "HOSPITAL DETIALS")
Select Case response                        ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                           ' Unload the current form.
        
  Case vbNo:
  
  'Clear the control and set the default values.
  
        msflgHospitalDetails.Clear
        msflgHospitalDetails.Rows = 2
        optHospitalName.Enabled = True
        optHospitalID.Enabled = True
        optArea.Enabled = True
        optHospitalName.Value = False
        optHospitalID.Value = False
        optArea.Value = False
        txtHospitalID.Enabled = True
        txtHospitalName.Enabled = True
        txtArea.Enabled = True
        txtHospitalName.Text = ""
        txtHospitalID.Text = ""
        txtArea.Text = ""
  
  End Select
End Sub

'*********************** Delete Hospital Details ***********************'

' This procedure is used to Delete Hospital Details.

Private Sub cmddelete_Click()
Dim strquery As String
Dim strresponse As String
Dim i As Integer

If Len(msflgHospitalDetails.TextMatrix(msflgHospitalDetails.Row, 1)) = 0 Then
  MsgBox "First search  for the Hospital You want to Delete", vbExclamation, "HOSPITAL SEARCH"
Else
strresponse = MsgBox("Are you sure you want to delete", vbQuestion + vbYesNo)
   Select Case strresponse
        Case vbYes:
        strquery = "Insert into tbl_hospitalwaste values('" & msflgHospitalDetails.TextMatrix(msflgHospitalDetails.Row, 1) & "')"
        Con.Execute strquery
         Con.BeginTrans
         
        'Delete the Hospital Details
        
         strquery = "DELETE FROM tbl_HospitalInfo WHERE HospitalID='" & msflgHospitalDetails.TextMatrix(msflgHospitalDetails.Row, 1) & "' "
         Con.Execute strquery
         For i = 0 To msflgHospitalDetails.Cols - 1
         msflgHospitalDetails.TextMatrix(msflgHospitalDetails.Row, i) = ""
         Next i
         Con.CommitTrans
         MsgBox "Record Deleted", vbInformation + vbOKOnly, "Delete"
         msflgHospitalDetails.Rows = msflgHospitalDetails.Rows - 1
    End Select
End If
End Sub

'*********************** Search Hospital Details ***********************'

' This procedure is used to Search for Hospital Details for the search criteria.

Private Sub cmdSearch_Click()
Dim strsql As String
Dim srno As Integer
Dim rssearch As New ADODB.Recordset
On Error GoTo errPara

'Validation for Hospital ID
If optHospitalName.Value = False And optArea.Value = False And optHospitalID.Value = True Then
     If (Len(Trim(txtHospitalID.Text)) = 0) Then
       MsgBox "First enter the Hospital ID", vbInformation, "HospitalID"
       txtHospitalID.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT  *  FROM tbl_HospitalInfo WHERE HospitalID='" & txtHospitalID.Text & "'", Con, adOpenKeyset, adLockOptimistic
     End If
     
'Validation for Hospital Name
    ElseIf optHospitalID.Value = False And optArea.Value = False And optHospitalName.Value = True Then
     If (Len(Trim(txtHospitalName.Text)) = 0) Then
       MsgBox "First enter the Hospital Name", vbInformation, "Hospital Name"
       txtHospitalName.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT * FROM tbl_HospitalInfo WHERE Hospital='" & txtHospitalName.Text & "'", Con, adOpenKeyset, adLockOptimistic
     End If
     
'Validation for Hospital Area
    ElseIf optHospitalID.Value = False And optHospitalName.Value = False And optArea.Value = True Then
       If Len(Trim(txtArea.Text)) = 0 Then
         MsgBox "First enter the area.", vbInformation, "Hospital Area"
         txtArea.SetFocus
         Exit Sub
       Else
           rssearch.Open "SELECT  *  FROM tbl_HospitalInfo WHERE Area='" & txtArea.Text & "'", Con, adOpenKeyset, adLockOptimistic
       End If
    Else
      MsgBox "First choose and enter the search data.", vbExclamation, "Hospital Search"
      Exit Sub
   End If

    srno = 1
    msflgHospitalDetails.Clear
    msflgHospitalDetails.Rows = 2
    
    
    msflgHospitalDetails.FormatString = "SNo | HospitalID |Hospital Name   |Hospital Address   | Phone1      | Area            "
   
If rssearch.BOF And rssearch.EOF Then
  MsgBox "No records found for this criteria."
Else
 While Not rssearch.EOF
 
   'Assigns the values from the database to the corresponding controls.

        msflgHospitalDetails.TextMatrix(srno, 0) = srno
        msflgHospitalDetails.TextMatrix(srno, 1) = rssearch("HospitalID") & ""
        msflgHospitalDetails.TextMatrix(srno, 2) = rssearch("Hospital") & ""
        msflgHospitalDetails.TextMatrix(srno, 3) = rssearch("HospitalAddress") & ""
        msflgHospitalDetails.TextMatrix(srno, 4) = rssearch("Phone1") & ""
        msflgHospitalDetails.TextMatrix(srno, 5) = rssearch("Area") & ""
        rssearch.MoveNext
        srno = srno + 1
        If msflgHospitalDetails.Rows = srno Then
         msflgHospitalDetails.Rows = msflgHospitalDetails.Rows + 1
        End If
 Wend
    rssearch.Close
    Set rssearch = Nothing
End If

'Set the default values.

optHospitalName.Enabled = True
optHospitalID.Enabled = True
optArea.Enabled = True
optHospitalName.Value = False
optHospitalID.Value = False
optArea.Value = False
txtHospitalID.Enabled = False
txtHospitalName.Enabled = False
txtArea.Enabled = False

'Clear all the controls
txtHospitalName.Text = ""
txtHospitalID.Text = ""
txtArea.Text = ""
Exit Sub

'Code for Error Handling.
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

 frmHospitalLookUp.Height = 4815
 frmHospitalLookUp.Width = 6105
 CenterForm frmHospitalLookUp
 
End Sub

'*********************** Select the Hospital Area ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optArea_Click()
  
 'Sets the default value for all the control.
 
 optHospitalName.Enabled = False
 optHospitalID.Enabled = False
 txtHospitalName.Enabled = False
 txtArea.Enabled = True
 txtHospitalID.Enabled = False
 
 'Clear all the control
 txtHospitalName.Text = ""
 txtHospitalID.Text = ""
 txtArea.Text = ""
 txtArea.SetFocus
End Sub

'*********************** Select the Hospital ID ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optHospitalID_Click()

 'Sets the default value for all the control.
 
 optHospitalName.Enabled = False
 optArea.Enabled = False
 txtHospitalName.Enabled = False
 txtArea.Enabled = False
 txtHospitalID.Enabled = True
 
 'Clear all the control

 txtArea.Text = ""
 txtHospitalName.Text = ""
 txtHospitalID.Text = ""
 txtHospitalID.SetFocus
End Sub

'*********************** Select the Hospital Name ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optHospitalName_Click()
 
'Sets the default value for all the control.

 optHospitalID.Enabled = False
 optArea.Enabled = False
 txtHospitalName.Enabled = True
 txtArea.Enabled = False
 txtHospitalID.Enabled = False
  
'Clear all the control

 txtArea.Text = ""
 txtHospitalID.Text = ""
 txtHospitalName.Text = ""
 txtHospitalName.SetFocus
End Sub

