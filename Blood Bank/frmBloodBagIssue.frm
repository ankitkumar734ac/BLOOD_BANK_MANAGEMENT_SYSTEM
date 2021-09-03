VERSION 5.00
Begin VB.Form frmBloodBagIssue 
   Caption         =   "Blood Bag Issue"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   Icon            =   "frmBloodBagIssue.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   10845
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "BloodBagIssue"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9135
      Begin VB.ComboBox cboissuedby 
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
         ItemData        =   "frmBloodBagIssue.frx":25CA
         Left            =   1560
         List            =   "frmBloodBagIssue.frx":25DD
         TabIndex        =   2
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox cboproduct 
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
         ItemData        =   "frmBloodBagIssue.frx":2624
         Left            =   1560
         List            =   "frmBloodBagIssue.frx":2634
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdHSearch 
         Caption         =   "Hospital  Search"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5880
         TabIndex        =   17
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdPSearch 
         Caption         =   "Patient  search "
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Cancel 
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
         Height          =   495
         Left            =   7680
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdIssue 
         Caption         =   "Issue "
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
         Left            =   3240
         TabIndex        =   12
         Top             =   2280
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "ISSUED TO"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   3120
         TabIndex        =   7
         Top             =   240
         Width           =   5895
         Begin VB.OptionButton optPatient 
            Caption         =   "Patient "
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optHospital 
            Caption         =   "Hospital "
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
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtPatientID 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   9
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtHospitalID 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   8
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblPatientID 
            AutoSize        =   -1  'True
            Caption         =   "PatientID"
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
            Left            =   3360
            TabIndex        =   11
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label lblHospitalID 
            Caption         =   "HospitalID"
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
            Left            =   360
            TabIndex        =   10
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.TextBox txtBloodBagID 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1560
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   2040
         Width           =   5775
      End
      Begin VB.Label lblIssuedBy 
         AutoSize        =   -1  'True
         Caption         =   "Issued By"
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
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Product"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblBloodBagID 
         AutoSize        =   -1  'True
         Caption         =   "BloodBagID"
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1365
      End
   End
   Begin VB.Label Label2 
      Caption         =   "BloodBagID is not available plz enter Valid BloodBagID"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   1320
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frmBloodBagIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmBloodBagIssue

'Displays:  Form will be displayed when user clicks on the Blood Bag Issue Menu of the MDI Form

'Unload  :  When the user has accessed the blood bag details , user can click the Cancel command button and choose the Yes option to terminate from the Form frmBBLookUp

'Functions: This form is used to issue the Blood bags to the Hospital and Patients

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.


Option Explicit

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub Cancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "BLOOD BAG DETIALS")
Select Case response                   ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                      'Unload the current Form
        
  Case vbNo:
  
        'Clears the Controls to set the default values.
        optHospital.Enabled = True
        optPatient.Enabled = True
        optHospital.Value = False
        optPatient.Value = False
        txtBloodBagID.Text = ""
        txtHospitalID.Text = ""
        cboissuedby.Text = ""
        txtPatientID.Text = ""
        cboproduct.Text = ""
       
 End Select
End Sub

'*********************** Hospital Search ***********************'

' This procedure is used to search the Existing Hospital details.

Private Sub cmdHSearch_Click()
 frmHospitalLookUp.Show           'Loads and Shows the Hospital Look Up Form
End Sub

'*********************** Issue the Blood Bags ***********************'

' This procedure is used to Issue the Blood bags to Patients and Hospitals.

Private Sub cmdIssue_Click()
Dim rsnew As New ADODB.Recordset
Dim rsnew1 As New ADODB.Recordset
Dim strsql As String
Dim strSql1 As String
Dim i As Integer
On Error GoTo errPara
If Validatedata = False Then
    Exit Sub
Else
 rsnew.Open "Select * from tbl_BloodBagDetails where Status='Issued'", Con, adOpenKeyset, adLockOptimistic
  For i = 0 To rsnew.RecordCount - 1
   If rsnew("BloodBagID") = txtBloodBagID.Text Then
    MsgBox "This blood bag is already issued."
    Exit For
   Else
   rsnew.MoveNext
   End If
  Next i
  If rsnew.EOF Then
   If optHospital.Value = True Then
   Con.BeginTrans
    
    'Insert the Existing details in Hospital Issue Table of the Database
    strsql = "INSERT INTO tbl_HospitalIssue(HospitalID,BloodBagID,Product,IssuedBy,IssuedDate) VALUES ('" & txtHospitalID.Text & "','" _
  & txtBloodBagID.Text & "' ,'" & cboproduct.Text & "' , '" & cboissuedby.Text & "' , '" & Now() & "')"
  Con.Execute strsql
  
   'Modify the existing Records of the Database.
  strSql1 = "UPDATE  tbl_BloodBagDetails SET  Status='Issued' WHERE BloodBagID=" & Val(txtBloodBagID.Text)
  Con.Execute strSql1
  Con.CommitTrans
  MsgBox "Blood bag issued"
  
   'Insert the Existing details in Patient Issue Table of the Database.
 ElseIf optPatient.Value = True Then
 Con.BeginTrans
 strsql = "INSERT INTO tbl_PatientIssue(PatientID,BloodBagID,Product,IssuedBy,IssueDate) VALUES ('" & txtPatientID.Text & "','" _
 & txtBloodBagID.Text & "' ,'" & cboproduct.Text & "' , '" & cboissuedby.Text & "' , '" & Now() & "')"
 Con.Execute strsql
 
    'Modify the existing Records of the Database.
 strSql1 = "UPDATE  tbl_BloodBagDetails SET  Status='Issued' WHERE BloodBagID=" & Val(txtBloodBagID.Text)
 Con.Execute strSql1
 Con.CommitTrans
 MsgBox " Blood bag issued"
End If
 End If
End If
 Exit Sub
 
'Code for error handling.
errPara:
MsgBox Err.Description
MsgBox "The entry is not allowed"
Load MDIForm1
Con.RollbackTrans
End Sub

'*********************** Patient Search ***********************'

' This procedure is used to search the Existing Patient details.

Private Sub cmdPSearch_Click()
  frmPatientQuickInfo.Show              'Loads and Shows the Patient Look Up Form
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()
On Error GoTo errPara
'Set the Dimensions of the Form
frmBloodBagIssue.Height = 3840
frmBloodBagIssue.Width = 9360
'CenterForm frmBloodBagIssue
Dim BBID As String
Dim i As Integer
BBID = aaa

'Set the Default settings of the form
open_recordset "tbl_BloodBagDetails"
rs.MoveFirst
For i = 0 To rs.RecordCount - 1
If rs("BloodBagID") = BBID Then
  fillDetails BBID
 ' CenterForm frmBloodBagIssue
  Exit For
End If
rs.MoveNext
Next i
If rs.EOF Then
 MsgBox "This is not a valid BloodBagID.", vbInformation, "BloodBagIssue"
 close_recordset
' lblBloodBagID.Visible = False
' Label1.Visible = False
' lblIssuedBy.Visible = False
' optHospital.Visible = False
' lblHospitalID.Visible = False
' txtHospitalID.Visible = False
' cboissuedby.Visible = False
' cboproduct.Visible = False
' txtBloodBagID.Visible = False
 'cmdIssue.Visible = False
 'cmdPSearch.Visible = False
 'cmdHSearch.Visible = False
 'txtPatientID.Visible = False
 'optPatient.Visible = False
 'lblPatientID.Visible = False
 Frame2.Visible = False
 Frame1.Visible = False
 Label2.Visible = True
 Command1.Visible = True
 
 
 
 'Unload Me
 Exit Sub
End If
close_recordset
Exit Sub

'Code for error Handling
errPara:
 MsgBox Err.Description
  Unload Me
End Sub

'*********************** Select the Hospital ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optHospital_Click()

'Settings for all the control.
optPatient.Enabled = False
txtPatientID.Text = ""
txtPatientID.Enabled = False
txtHospitalID.Text = ""
txtHospitalID.Enabled = True
txtHospitalID.SetFocus

End Sub

'*********************** Select the Patient ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optPatient_Click()

'Settings for all the control.
optHospital.Enabled = False
txtPatientID.Text = ""
txtPatientID.Enabled = True
txtHospitalID.Text = ""
txtHospitalID.Enabled = False
txtPatientID.SetFocus

End Sub

'*********************** Fills the Blood Bag Details into the Form ***********************'

' This procedure is used to fill the Blood Bag Details in the form

Private Sub fillDetails(BID As String)
Dim strsql As String
Dim rsDetail As New ADODB.Recordset

If BID = 0 Then
  Exit Sub
Else
  strsql = "SELECT * FROM tbl_BloodBagDetails WHERE BloodBagID=" & BID
  rsDetail.Open strsql, Con, adOpenKeyset, adLockOptimistic
  txtBloodBagID.Text = rsDetail("BloodBagID")
End If
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate the data that is to be inserted into the Database.

Public Function Validatedata() As Boolean
Validatedata = True
Dim strSql1 As String
Dim rshos As New ADODB.Recordset
strSql1 = "SELECT * from tbl_HospitalInfo WHERE HospitalID='" & txtHospitalID.Text & "' "
rshos.Open strSql1, Con, adOpenDynamic, adLockOptimistic

Dim strsql2 As String
Dim rspat As New ADODB.Recordset
strsql2 = "SELECT * from tbl_PatientInfo WHERE PatientID='" & txtPatientID.Text & "'"
rspat.Open strsql2, Con, adOpenDynamic, adLockOptimistic

'Validate the Blood Bag ID
If Len(txtBloodBagID.Text) = 0 Then
    MsgBox "Please enter BloodBagID"
    Validatedata = False
    txtBloodBagID.SetFocus
    Exit Function
End If

'Validate the Product
If Len(cboproduct.Text) = 0 Then
    MsgBox "Please enter product"
    Validatedata = False
    cboproduct.SetFocus
    Exit Function
End If

'Validate the Hospital ID
If optHospital.Value = True Then
If txtHospitalID.Text = "" Then
MsgBox "Enter Hospital ID"
Validatedata = False
txtHospitalID.SetFocus
Exit Function
End If

'Validate The Patient ID

ElseIf optPatient.Value = True Then
If txtPatientID.Text = "" Then
MsgBox "Enter Patient ID"
Validatedata = False
txtPatientID.SetFocus
Exit Function
End If
End If

'Validate the DEO's name
If Len(cboissuedby.Text) = 0 Then
    MsgBox "Please enter DEO's name"
    Validatedata = False
    cboissuedby.SetFocus
    Exit Function
End If


If optHospital.Value = True Then
  If rshos.BOF And rshos.EOF Then
  
  MsgBox "INVALID HOSPITALID.BLOOD BAG CAN NOT BE ISSUED"
  ans = MsgBox("PLEASE SAVE THE DEATILS OF THE HOSPITAL BEFORE ISSUING A BLOOD BAG", vbYesNo)
  If ans = vbYes Then
 
  frmHospitalInfo.Show
  Validatedata = False
  Exit Function
  Else
  Validatedata = False
  txtHospitalID.Text = ""
  Exit Function
  End If
  rshos.MoveNext
  End If

 rshos.Close
 Set rshos = Nothing

End If

If optPatient.Value = True Then
  If rspat.BOF And rspat.EOF Then
  MsgBox "INVALID PATIENTID.BLOOD BAG CAN NOT BE ISSUED"
  ans = MsgBox("PLEASE SAVE THE DETAILS OF THE PATIENT BEFORE ISSUING A BLOOD BAG", vbYesNo)
  If ans = vbYes Then
 
  frmPatient.Show
  Validatedata = False
  Exit Function
  Else
  Validatedata = False
  txtPatientID.Text = ""
  Exit Function
  End If
  rspat.MoveNext
  End If
 
 rspat.Close
 Set rspat = Nothing

End If
End Function

