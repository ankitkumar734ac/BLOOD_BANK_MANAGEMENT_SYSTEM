VERSION 5.00
Begin VB.Form frmIssue2Patient 
   Caption         =   "PATIENT ISSUE"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   7230
   Begin VB.Frame fraPatientIssue 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3120
         TabIndex        =   15
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdIssue 
         Caption         =   "Issue"
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtIssuedBy 
         Height          =   285
         Left            =   5160
         TabIndex        =   11
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtIndForTrans 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtProduct 
         Height          =   285
         Left            =   5160
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtBloodBagID 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtPatientID 
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblIssuedBy 
         AutoSize        =   -1  'True
         Caption         =   "Issued By"
         Height          =   195
         Left            =   3840
         TabIndex        =   10
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblIndForTrans 
         Caption         =   "Indication For Transfusion"
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblProduct 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         Height          =   195
         Left            =   3840
         TabIndex        =   6
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lblBloodBagID 
         AutoSize        =   -1  'True
         Caption         =   "BloodBag ID"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblPatientID 
         AutoSize        =   -1  'True
         Caption         =   "Patient ID"
         Height          =   195
         Left            =   3840
         TabIndex        =   2
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmIssue2Patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
txtName.Text = ""
txtPatientID.Text = ""
txtBloodBagID.Text = ""
txtProduct.Text = ""
txtIndForTrans.Text = ""
txtIssuedBy.Text = ""
txtName.SetFocus

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdIssue_Click()
Dim strsql As String

If Validation = False Then
  MsgBox "blood bag cannot be issued"
    
Exit Sub
End If



Con.BeginTrans

strsql = "UPDATE tbl_BloodBagDetails Set Status = 'Issued' Where BloodBagID = '" & txtBloodBagID & "'"
Con.Execute strsql
strsql = "INSERT INTO tbl_PatientIssue(PatientID,BloodBagID,Product,IndForTrans,IssuedBy," _
& "IssueDate) VALUES ('" & txtPatientID.Text & "','" & txtBloodBagID.Text & "','" _
& txtProduct.Text & "','" & txtIndForTrans.Text & "','" & txtIssuedBy.Text & "','" _
& Now() & "')"
Con.Execute strsql
Con.CommitTrans
MsgBox "blood bag issued"
End Sub

Private Sub Form_Load()
frmIssue2Patient.Height = 4080
frmIssue2Patient.Width = 7350
CenterForm frmIssue2Patient
End Sub



Public Function Validation() As Boolean
Validation = True

Dim rssearch As New ADODB.Recordset
Dim strsql As String


strsql = "SELECT Status FROM tbl_BloodBagDetails WHERE BloodBagID='" & txtBloodBagID & "'"
rssearch.Open strsql, Con, adOpenKeyset, adLockOptimistic
If rssearch("Status") = "Issued" Then
   Validation = False
ElseIf rssearch("Status") = "Discarded" Then
 Validation = False
Exit Function
End If
End Function


