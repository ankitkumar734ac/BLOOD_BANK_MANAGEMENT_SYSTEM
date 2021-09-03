VERSION 5.00
Begin VB.Form frmPatient 
   Caption         =   "PATIENT DETAILS ENTRY"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   Icon            =   "frmPatient.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   8250
   Begin VB.Frame fraPatientInfo 
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
      Height          =   5055
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   8295
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel "
         Enabled         =   0   'False
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
         Left            =   2880
         TabIndex        =   27
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Frame framove 
         Caption         =   "Move"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5280
         TabIndex        =   22
         Top             =   240
         Width           =   2535
         Begin VB.CommandButton cmdmove 
            Caption         =   "Last"
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
            Index           =   3
            Left            =   1320
            TabIndex        =   26
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdmove 
            Caption         =   "Next"
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   1095
         End
         Begin VB.CommandButton cmdmove 
            Caption         =   "Previous"
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   1320
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdmove 
            Caption         =   "First"
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
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdmodify 
         Caption         =   "Modify "
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
         Left            =   4200
         TabIndex        =   21
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add "
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
         Left            =   360
         TabIndex        =   20
         Top             =   4320
         Width           =   975
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
         Height          =   390
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C&lear "
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
         Left            =   5640
         TabIndex        =   11
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdclose 
         Cancel          =   -1  'True
         Caption         =   "Close "
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
         Left            =   6840
         TabIndex        =   12
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save "
         Enabled         =   0   'False
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
         Left            =   1560
         TabIndex        =   10
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtAge 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Frame fraGender 
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4680
         TabIndex        =   17
         Top             =   3000
         Width           =   3135
         Begin VB.OptionButton optFemale 
            Caption         =   "Female"
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
            Left            =   1440
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optMale 
            Caption         =   "Male"
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
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbOccupation 
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
         ItemData        =   "frmPatient.frx":25CA
         Left            =   6120
         List            =   "frmPatient.frx":25DA
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtPhone 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         MaxLength       =   17
         TabIndex        =   3
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1560
         TabIndex        =   4
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtSurname 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   4200
         Width           =   7695
      End
      Begin VB.Label lblPatientID 
         AutoSize        =   -1  'True
         Caption         =   "Patient ID"
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
         TabIndex        =   19
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         Caption         =   "Age"
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
         Left            =   5400
         TabIndex        =   18
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lblOccupation 
         AutoSize        =   -1  'True
         Caption         =   "Occupation"
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
         Left            =   4680
         TabIndex        =   16
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
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
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmPatient

'Displays:  Form will be displayed when user clicks on the Patient Info Menu of the MDI Form

'Unload  :  When the user finish accessing the Patient details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmPatient

'Functions: This form is used to access the Patient Details.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Add New Patient Details ***********************'

' This procedure is used to Add New Patient Details.

Private Sub cmdAdd_Click()
Dim rsnew As New ADODB.Recordset
Dim n As String

'Clear all the controls and sets the default value.
        txtSurname.Text = ""
        txtPatientID.Text = ""
        txtName.Text = ""
        txtAddress.Text = ""
        txtPhone.Text = ""
        txtAge.Text = ""
        optMale.Value = False
        optFemale.Value = False
        cmbOccupation.Text = ""
        txtSurname.SetFocus
   rsnew.Open "Select * from tbl_patientwaste", Con, adOpenKeyset, adLockOptimistic
If rsnew.RecordCount > 0 Then
 rsnew.MoveFirst
 txtPatientID.Text = rsnew("PATIENTID")
 strsql = "Delete from tbl_patientwaste where PATIENTID='" & txtPatientID.Text & "'"
 Con.Execute strsql
Else
    open_recordset "tbl_PatientInfo"
   
   'Auto Generation For New Donor Details.
    n = rs.RecordCount
    If n < 8 Then
        n = rs.RecordCount + 1
        txtPatientID.Text = "p0" + Trim(n)
    ElseIf n = 8 Then
        txtPatientID.Text = "p09"
    Else
        n = rs.RecordCount + 1
        txtPatientID.Text = "p" + Trim(n)
    End If
    close_recordset
End If
Set rsnew = Nothing

'Sets the default value for all the controls.
 cmdAdd.Enabled = False
 cmdClear.Enabled = False
 cmdmodify.Enabled = False
 cmdclose.Enabled = False
 cmdSave.Enabled = True
 cmdSave.Default = True
 cmdcancel.Enabled = True
 cmdcancel.Cancel = True
 framove.Enabled = False
End Sub

'*********************** Cancel ***********************'

' This procedure is used to unload the form.

Private Sub cmdcancel_Click()
Dim strsql As String
strsql = "Insert into tbl_patientwaste values('" & txtPatientID.Text & "')"
Con.Execute strsql

'Clear all the controls and sets default values.
 txtSurname.Text = ""
 txtPatientID.Text = ""
 txtName.Text = ""
 txtAddress.Text = ""
 txtPhone.Text = ""
 txtAge.Text = ""
 optMale.Value = False
 optFemale.Value = False
 cmbOccupation.Text = ""
 txtSurname.SetFocus
 cmdAdd.Enabled = True
 cmdClear.Enabled = True
 cmdmodify.Enabled = True
 cmdclose.Enabled = True
 cmdSave.Enabled = True
 cmdSave.Default = False
 cmdcancel.Enabled = True
 cmdmove(1).Enabled = False
 cmdmove(2).Enabled = False
 cmdcancel.Cancel = False
 framove.Enabled = True
End Sub

'*********************** Clear The Donor Details ***********************'

' This procedure is used to clear the Patient Details.
Private Sub cmdclear_Click()

'Clear all the controls and sets default values.
txtSurname.Text = ""
txtPatientID.Text = ""
txtName.Text = ""
txtAddress.Text = ""
txtPhone.Text = ""
txtAge.Text = ""
optMale.Value = False
optFemale.Value = False
cmbOccupation.Text = ""
txtSurname.SetFocus

End Sub

'*********************** Close ***********************'

' This procedure is used to terminate the form.

Private Sub cmdclose_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "Patient Details")
Select Case response                    ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                       ' Unload the current form.
        
  Case vbNo:
  
  'Clear all the controls and sets default values.

        optFemale.Enabled = True
        optFemale.Value = False
        optMale.Enabled = True
        optMale.Value = False
        txtAddress.Text = ""
        txtAge.Text = ""
        txtName.Text = ""
        txtPatientID.Text = ""
        txtPhone.Text = ""
        txtSurname.Text = ""
        cmbOccupation.Text = ""
        cmdmove(1).Enabled = False
        cmdmove(2).Enabled = False
        cmdSave.Enabled = False
        cmdcancel.Enabled = False
 
 End Select
End Sub

'*********************** Update the Donor Details ***********************'

' This procedure is used to update the Patient Details.

Private Sub cmdmodify_Click()
  Dim rsnew As New ADODB.Recordset
  On Error GoTo errPara
  rsnew.Open "select * from tbl_PatientInfo where PatientID='" & txtPatientID.Text & "'", Con, adOpenKeyset, adLockOptimistic
    If Validatedata = False Then
     Exit Sub
  Else
     Con.BeginTrans
    'Update the Patient Details.
      rsnew.Update "PatientSurname", txtSurname.Text
      rsnew.Update "PatientName", txtName.Text
      rsnew.Update "Address", txtAddress.Text
      rsnew.Update "Phone", txtPhone.Text
      rsnew.Update "Occupation", cmbOccupation.Text
      rsnew.Update "Gender", IIf(optMale.Value = True, "Male", "Female")
      rsnew.Update "Age", txtAge.Text
     Con.CommitTrans
  End If
   MsgBox "The Patient Information is Modified", vbInformation, "Hospital Modification"
   Set rsnew = Nothing
   
  'Clear all the controls and sets default values.

   optFemale.Enabled = True
   optFemale.Value = False
   optMale.Enabled = True
   optMale.Value = False
   txtAddress.Text = ""
   txtAge.Text = ""
   txtName.Text = ""
   txtPatientID.Text = ""
   txtPhone.Text = ""
   txtSurname.Text = ""
   cmbOccupation.Text = ""
   cmdmove(1).Enabled = False
   cmdmove(2).Enabled = False

   Exit Sub

'Code for Error Handling.
errPara:
 MsgBox Err.Description
 Con.RollbackTrans
 
End Sub

'*********************** Move ***********************'

' This procedure is used to Move through the records of Donor.

Private Sub cmdmove_Click(Index As Integer)
Dim i As Integer
open_recordset "tbl_PatientInfo"
Select Case Index
 Case 0:
    
    'Shows the First record of Patient.
    'Display the record on various controls.
        rs.MoveFirst
        txtSurname.Text = rs("PatientSurname")
        txtPatientID.Text = rs("PatientID")
        txtName.Text = rs("PatientName")
        txtAddress.Text = rs("Address")
        txtPhone.Text = rs("Phone")
        txtAge.Text = rs("Age")
        If rs("Gender") = "Male" Then
         optMale.Value = True
        Else
         optFemale.Value = True
        End If
        cmbOccupation.Text = rs("Occupation")
        close_recordset
        cmdmove(1).Enabled = True
        cmdmove(2).Enabled = True
        
 Case 1:
     
 ' Shows the Previous Record of the Patient  with respect to the current record.

     For i = 0 To rs.RecordCount
       If rs("PatientID") = txtPatientID.Text Then
           rs.MovePrevious
          If rs.BOF Then
           rs.MoveLast
           Exit For
           End If
           Exit For
         Else
       rs.MoveNext
       End If
      Next i
    
    'Display the record on various controls.
        txtSurname.Text = rs("PatientSurname")
        txtPatientID.Text = rs("PatientID")
        txtName.Text = rs("PatientName")
        txtAddress.Text = rs("Address")
        txtPhone.Text = rs("Phone")
        txtAge.Text = rs("Age")
        If rs("Gender") = "Male" Then
         optMale.Value = True
        Else
         optFemale.Value = True
        End If
        cmbOccupation.Text = rs("Occupation")
        close_recordset
Case 2:

   ' Shows the Next record of the Patient with respect to the current record.

     rs.MoveFirst
     For i = 0 To rs.RecordCount
       If rs("PatientID") = txtPatientID.Text Then
         rs.MoveNext
        If rs.EOF Then
         rs.MoveFirst
       Exit For
       End If
       Exit For
     Else
      rs.MovePrevious
      If rs.BOF Then
       rs.MoveLast
      End If
     End If
     Next i

    'Display the record on various controls.
        txtSurname.Text = rs("PatientSurname")
        txtPatientID.Text = rs("PatientID")
        txtName.Text = rs("PatientName")
        txtAddress.Text = rs("Address")
        txtPhone.Text = rs("Phone")
        txtAge.Text = rs("Age")
        If rs("Gender") = "Male" Then
         optMale.Value = True
        Else
         optFemale.Value = True
        End If
        cmbOccupation.Text = rs("Occupation")
        close_recordset
        
 Case 3:
     ' Shows the last record of Patient
        rs.MoveLast
     'Display the record on various controls.

        txtSurname.Text = rs("PatientSurname")
        txtPatientID.Text = rs("PatientID")
        txtName.Text = rs("PatientName")
        txtAddress.Text = rs("Address")
        txtPhone.Text = rs("Phone")
        txtAge.Text = rs("Age")
        If rs("Gender") = "Male" Then
         optMale.Value = True
        Else
         optFemale.Value = True
        End If
        cmbOccupation.Text = rs("Occupation")
     close_recordset
      cmdmove(1).Enabled = True
      cmdmove(2).Enabled = True
End Select
End Sub

'*********************** Save ***********************'

' This procedure is used to insert the New Patient Details into the Database.

Private Sub cmdSave_Click()
Dim strsql As String
Dim rssearch As New ADODB.Recordset
Dim response As String
Dim n As Integer
On Error GoTo errPara
If Validatedata = False Then
    Exit Sub
Else
Con.BeginTrans

'Insert the New Donor Details into the Database.
strsql = "INSERT INTO tbl_PatientInfo (PatientID,PatientSurname,PatientName,Address," _
& "Phone, Occupation,Gender,Age) VALUES ('" & txtPatientID.Text & "','" & txtSurname.Text & "','" & txtName.Text & "','" & txtAddress.Text & "','" _
& txtPhone.Text & "','" & cmbOccupation.Text & "','" & IIf(optMale.Value = True, "Male", "Female") & "','" _
& txtAge.Text & "')"
Con.Execute strsql
Con.CommitTrans
MsgBox "data saved"
End If

'Clear all the controls and sets default values.
 txtAddress.Text = ""
 txtAge.Text = ""
 txtName.Text = ""
 txtPatientID.Text = ""
 txtPhone.Text = ""
 txtSurname.Text = ""
 cmbOccupation.Text = ""
 cmdAdd.Enabled = True
 cmdClear.Enabled = True
 cmdmodify.Enabled = True
 cmdclose.Enabled = True
 cmdSave.Enabled = False
 cmdcancel.Enabled = False
 cmdSave.Default = False
 cmdcancel.Cancel = False
 framove.Enabled = True
 Exit Sub

'Code for Error Handling.
errPara:
    MsgBox Err.Description
    Con.RollbackTrans
    Load MDIForm1
   
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'Sets the Dimension of the Form

frmPatient.Height = 5415
frmPatient.Width = 8370
CenterForm frmPatient
End Sub

'*********************** Age Validation ***********************'

' This procedure is used to validate the Age entered by the user.

Private Sub txtAge_keypress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If (keyascii < 48 Or keyascii > 57) Then
    MsgBox "Please enter only digits."
    keyascii = 0
End If
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate the data entered by the user.

Public Function Validatedata() As Boolean
Validatedata = True

'Validate the entered Surname
 If Len(Trim(txtSurname.Text)) = 0 Then
    MsgBox "Please enter Sur name"
    Validatedata = False
    txtSurname.SetFocus
    Exit Function
 End If
 
'Validate the entered Name
 If Len(Trim(txtName.Text)) = 0 Then
    MsgBox "Please enter name"
    Validatedata = False
    txtName.SetFocus
   Exit Function
End If

'Validate the entered Patient ID
If Len(Trim(txtPatientID.Text)) = 0 Then
    MsgBox "Please enter PatientID "
    Validatedata = False
    txtPatientID.SetFocus
    Exit Function
End If

'Validate the entered Address
If Len(Trim(txtAddress.Text)) = 0 Then
    MsgBox "Please enter Address"
    Validatedata = False
    txtAddress.SetFocus
   Exit Function
End If

'Validate the entered Occupation
If Len(Trim(cmbOccupation.Text)) = 0 Then
    MsgBox "Please Select Occupation"
    Validatedata = False
    Exit Function
End If

'Validate the entered Age
If Len(Trim(txtAge.Text)) = 0 Then
    MsgBox "Please enter Age"
    Validatedata = False
    txtAge.SetFocus
    Exit Function
End If

'Validate the entered Phone No.
If Len(Trim(txtPhone.Text)) = 0 Then
    MsgBox "Please enter Phone number"
    Validatedata = False
    txtPhone.SetFocus
    Exit Function
End If


End Function

'*********************** Input Validation ***********************'

' This procedure is used to validate the data entered by the user.

'Validate the entered Phone No
Private Sub txtPhone_KeyPress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If (keyascii < 48 Or keyascii > 57) Then
    MsgBox "Please enter only digits."
    keyascii = 0
End If
txtPhone.SetFocus
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate the data entered by the user.
' Validate the Entered Surname

Private Sub txtSurname_keypress(keyascii As Integer)
keyascii = lettercheck(keyascii)
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate the data entered by the user.
' Validate the Entered Name

Private Sub txtName_keypress(keyascii As Integer)
keyascii = lettercheck(keyascii)
End Sub

