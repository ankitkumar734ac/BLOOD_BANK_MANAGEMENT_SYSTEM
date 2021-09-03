VERSION 5.00
Begin VB.Form frmHospitalInfo 
   Caption         =   "HOSPITAL DETAILS  ENTRY"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   Icon            =   "frmHospitalInfo.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   6855
   Begin VB.Frame HospitalDetails 
      Caption         =   "Hospital Details"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6855
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
         Left            =   2280
         TabIndex        =   21
         Top             =   3480
         Width           =   1095
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
         Left            =   3480
         TabIndex        =   20
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear "
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
         Left            =   4680
         TabIndex        =   19
         Top             =   3480
         Width           =   855
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
         TabIndex        =   18
         Top             =   3480
         Width           =   855
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
         Height          =   1455
         Left            =   3960
         TabIndex        =   13
         Top             =   1680
         Width           =   2655
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
            TabIndex        =   17
            Top             =   360
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
            Left            =   1440
            TabIndex        =   16
            Top             =   360
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
            TabIndex        =   15
            Top             =   840
            Width           =   1095
         End
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
            Left            =   1440
            TabIndex        =   14
            Top             =   840
            Width           =   1095
         End
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdClose 
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
         Left            =   5640
         TabIndex        =   7
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save "
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
         Left            =   1320
         TabIndex        =   6
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtArea 
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
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtPhone1 
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
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtHospitalAddress 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2040
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtHospitalName 
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
         Left            =   2040
         MaxLength       =   35
         TabIndex        =   1
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   6615
      End
      Begin VB.Label lblHospitalID 
         AutoSize        =   -1  'True
         Caption         =   "Hospital ID"
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
         TabIndex        =   12
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lblArea 
         Caption         =   "Area"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label lblPhone1 
         AutoSize        =   -1  'True
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
         Height          =   315
         Left            =   4080
         TabIndex        =   10
         Top             =   480
         Width           =   705
      End
      Begin VB.Label lblHospitalAdrress 
         AutoSize        =   -1  'True
         Caption         =   "Hospital Address"
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
         TabIndex        =   9
         Top             =   1680
         Width           =   1890
      End
      Begin VB.Label lblHospitalName 
         AutoSize        =   -1  'True
         Caption         =   "Hospital Name"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmHospitalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmHospitalInfo

'Displays:  Form will be displayed when user clicks on the Hospital Menu of the MDI Form

'Unload  :  When the user finish accessing the Hospital details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmHospitalInfo

'Functions: This form is used to access the Hospital Details.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit


'*********************** Add New Hospital Details ***********************'

' This procedure is used to Add New Hospital Details.

Private Sub cmdAdd_Click()
Dim rsnew As New ADODB.Recordset
Dim n As String

'Clear all the controls and sets default values.

     txtHospitalID.Text = ""
     txtHospitalName.Text = ""
     txtHospitalAddress.Text = ""
     txtPhone1.Text = ""
     txtArea.Text = ""

rsnew.Open "Select * from tbl_hospitalwaste", Con, adOpenKeyset, adLockOptimistic
If rsnew.RecordCount > 0 Then
 rsnew.MoveFirst
 txtHospitalID.Text = rsnew("HOSPITALID")
 strsql = "Delete from tbl_hospitalwaste where HOSPITALID='" & txtHospitalID.Text & "'"
 Con.Execute strsql
Else
     open_recordset "tbl_HospitalInfo"
     
     'Auto Generation For New Hospital Details
      n = rs.RecordCount
      If n < 8 Then
        n = rs.RecordCount + 1
        txtHospitalID.Text = "h0" + Trim(n)
      ElseIf n = 8 Then
        txtHospitalID.Text = "h09"
      Else
        n = rs.RecordCount + 1
        txtHospitalID.Text = "h" + Trim(n)
      End If
     close_recordset
End If
Set rsnew = Nothing

'Clear all the controls and set default values.

 cmdAdd.Enabled = False
 cmdClear.Enabled = False
 cmdmodify.Enabled = False
 cmdSave.Enabled = True
 cmdcancel.Enabled = True
 cmdSave.Default = True
 cmdcancel.Cancel = True
 cmdclose.Enabled = False
 framove.Enabled = False
 cmdcancel.Cancel = True
End Sub

'*********************** Cancel ***********************'

' This procedure is used to unload the form.

Private Sub cmdcancel_Click()
Dim strsql As String
strsql = "Insert into tbl_hospitalwaste values('" & txtHospitalID.Text & "')"
Con.Execute strsql
  
  'Clear all the controls and set default values.

  txtArea.Text = ""
  txtHospitalAddress.Text = ""
  txtHospitalID.Text = ""
  txtHospitalName.Text = ""
  txtPhone1.Text = ""
  cmdAdd.Enabled = True
  cmdClear.Enabled = True
  cmdmodify.Enabled = True
  cmdclose.Enabled = True
  cmdSave.Enabled = False
  cmdcancel.Enabled = False
  cmdSave.Default = True
  cmdcancel.Cancel = True
  framove.Enabled = True
  cmdmove(1).Enabled = False
  cmdmove(2).Enabled = False
  cmdcancel.Cancel = False
End Sub

'*********************** Clear ***********************'

' This procedure is used to terminate the form.

 Private Sub cmdclear_Click()
 
 'Clear all the controls and set default values.

  txtArea.Text = ""
  txtHospitalAddress.Text = ""
  txtHospitalID.Text = ""
  txtHospitalName.Text = ""
  txtPhone1.Text = ""
End Sub

'*********************** Close ***********************'

' This procedure is used to terminate the form.

Private Sub cmdclose_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "DONOR DETIALS")
Select Case response                     ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                        ' Unload the current form.
        
  Case vbNo:
  
   'Clear all the controls and set default values.

        txtArea.Text = ""
        txtHospitalAddress.Text = ""
        txtHospitalID.Text = ""
        txtHospitalName.Text = ""
        txtPhone1.Text = ""
        cmdmove(1).Enabled = False
        cmdmove(2).Enabled = False
        cmdSave.Enabled = False
        cmdcancel.Enabled = False
 End Select
End Sub

'*********************** Update the Hospital Details ***********************'

' This procedure is used to update the Hospital Details.

Private Sub cmdmodify_Click()
Dim rsnew As New ADODB.Recordset
On Error GoTo errPara
  rsnew.Open "Select * from tbl_HospitalInfo where HospitalId='" & txtHospitalID.Text & "'", Con, adOpenKeyset, adLockOptimistic
  If Validatedata = False Then
    Exit Sub
  Else
     Con.BeginTrans
     
    'Update the Hospital Details
      rsnew.Update "Hospital", txtHospitalName.Text
      rsnew.Update "HospitalAddress", txtHospitalAddress.Text
      rsnew.Update "Phone1", txtPhone1.Text
      rsnew.Update "Area", txtArea.Text
     Con.CommitTrans
   MsgBox "The Hospital Information is Modified", vbInformation, "Hospital Modification"
      End If
    Set rsnew = Nothing
    
  'Clear all the controls and set default values.

    txtArea.Text = ""
    txtHospitalAddress.Text = ""
    txtHospitalID.Text = ""
    txtHospitalName.Text = ""
    txtPhone1.Text = ""
    cmdmove(1).Enabled = False
    cmdmove(2).Enabled = False
   Exit Sub

'Code for Error Handling.
errPara:
 MsgBox Err.Description
 Con.RollbackTrans

End Sub

'*********************** Move ***********************'

' This procedure is used to Move through the records of Hospital.

Private Sub cmdmove_Click(Index As Integer)
Dim i As Integer
open_recordset "tbl_HospitalInfo"
Select Case Index
 Case 0:
    
    'Shows the First record of Donor.
    'Display the record on various controls.
       
        rs.MoveFirst
        txtHospitalID.Text = rs("HospitalID")
        txtHospitalName.Text = rs("Hospital")
        txtPhone1.Text = rs("Phone1")
        txtHospitalAddress.Text = rs("HospitalAddress")
        txtArea.Text = rs("Area")
        cmdmove(1).Enabled = True
        cmdmove(2).Enabled = True
        close_recordset
 Case 1:
     
 ' Shows the Previous Record of the Donor  with respect to the current record.
     
     For i = 0 To rs.RecordCount
       If rs("HospitalID") = txtHospitalID.Text Then
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
        txtHospitalID.Text = rs("HospitalID")
        txtHospitalName.Text = rs("Hospital")
        txtPhone1.Text = rs("Phone1")
        txtHospitalAddress.Text = rs("HospitalAddress")
        txtArea.Text = rs("Area")
        close_recordset
        
Case 2:

' Shows the Next record of the Donor with respect to the current record.
     
     rs.MoveFirst
     For i = 0 To rs.RecordCount
       If rs("HospitalID") = txtHospitalID.Text Then
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
        txtHospitalID.Text = rs("HospitalID")
        txtHospitalName.Text = rs("Hospital")
        txtPhone1.Text = rs("Phone1")
        txtHospitalAddress.Text = rs("HospitalAddress")
        txtArea.Text = rs("Area")
     close_recordset
 
 Case 3:
      'Shows the last record of Donor
        rs.MoveLast
        
      'Display the record on various controls.
        txtHospitalID.Text = rs("HospitalID")
        txtHospitalName.Text = rs("Hospital")
        txtPhone1.Text = rs("Phone1")
        txtHospitalAddress.Text = rs("HospitalAddress")
        txtArea.Text = rs("Area")
        cmdmove(1).Enabled = True
        cmdmove(2).Enabled = True
        close_recordset
      
End Select
End Sub

'*********************** Save ***********************'

' This procedure is used to insert the New Hospital Details into the Database.

Private Sub cmdSave_Click()
Dim strsql As String
Dim n As Integer
On Error GoTo errPara

If Validatedata = False Then
    Exit Sub
Else
  Con.BeginTrans
  
  'Insert the New Hospital Details into the Database.
  strsql = "INSERT INTO tbl_HospitalInfo(HospitalID,Hospital,HospitalAddress,Phone1,Area)VALUES('" & txtHospitalID & "','" & txtHospitalName.Text & "','" _
  & txtHospitalAddress.Text & "' , '" & txtPhone1.Text & "' , '" & txtArea.Text & "')"
  Con.Execute strsql
  Con.CommitTrans
End If
 MsgBox "Data saved"
 
'Sets Default values for all controls.
 cmdAdd.Enabled = True
 cmdClear.Enabled = True
 cmdmodify.Enabled = True
 cmdclose.Enabled = True
 cmdSave.Enabled = False
 cmdcancel.Enabled = False
 cmdSave.Default = False
 cmdcancel.Cancel = False
 framove.Enabled = True
 cmdcancel.Cancel = False

'Clear all the Controls.
  txtArea.Text = ""
  txtHospitalAddress.Text = ""
  txtHospitalID.Text = ""
  txtHospitalName.Text = ""
  txtPhone1.Text = ""
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

' Sets the dimension for the form.
 frmHospitalInfo.Height = 4710
 frmHospitalInfo.Width = 6975
 CenterForm frmHospitalInfo
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate the input.

Public Function Validatedata() As Boolean
Validatedata = True

'Validate the Hospital ID Details.
If Len(Trim(txtHospitalID.Text)) = 0 Then
    MsgBox "Please enter Hospital ID"
    Validatedata = False
    txtHospitalID.SetFocus
    Exit Function
End If

'Validate the Hospital Name Details.
If Len(Trim(txtHospitalName.Text)) = 0 Then
    MsgBox "Please enter Hospital Name"
    Validatedata = False
    txtHospitalName.SetFocus
    Exit Function
End If

'Validate the Hospital Address Details.
If Len(Trim(txtHospitalAddress.Text)) = 0 Then
    MsgBox "Please enter Hospital Address"
    Validatedata = False
    txtHospitalAddress.SetFocus
    Exit Function
End If

'Validate the Hospital Phone No. Details.
If Len(Trim(txtPhone1.Text)) = 0 Then
    MsgBox "Please enter Phone Number"
    Validatedata = False
    txtPhone1.SetFocus
    Exit Function
End If

'Validate the Hospital Area Details.
If Len(Trim(txtArea.Text)) = 0 Then
    MsgBox "Please enter Area"
    Validatedata = False
    txtArea.SetFocus
    Exit Function
End If
End Function

'*********************** Input Validation ***********************'

' This procedure is used to validate the input.

Private Sub txtPhone1_KeyPress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If (keyascii < 48 Or keyascii > 57) Then
    MsgBox "Please enter only digits."
    keyascii = 0
End If
txtPhone1.SetFocus
End Sub
