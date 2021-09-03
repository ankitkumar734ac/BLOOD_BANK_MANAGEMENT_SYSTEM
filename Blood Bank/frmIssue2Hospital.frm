VERSION 5.00
Begin VB.Form frmIssue2Hospital 
   Caption         =   "HOSPITAL ISSUE "
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   7365
   Begin VB.Frame fraHospitalIssue 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtRh 
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ADD"
         Height          =   495
         Left            =   5760
         TabIndex        =   14
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdIssue 
         Caption         =   "ISSUE"
         Height          =   495
         Left            =   4080
         TabIndex        =   13
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox txtEnteredBy 
         Height          =   375
         Left            =   5160
         TabIndex        =   12
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtProduct 
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtBloodGroup 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtBloodBagID 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtHospitalID 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtHospitalName 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblRh 
         AutoSize        =   -1  'True
         Caption         =   "RH"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label lblEnteredBy 
         AutoSize        =   -1  'True
         Caption         =   "EnteredBy"
         Height          =   195
         Left            =   4080
         TabIndex        =   11
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblProduct 
         Caption         =   "Product"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBloodGroup 
         Caption         =   "Blood Group"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2385
         Width           =   1335
      End
      Begin VB.Label lblBloodBagID 
         AutoSize        =   -1  'True
         Caption         =   "BloodBag ID"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label lblHospitalID 
         AutoSize        =   -1  'True
         Caption         =   "Hospital ID"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label lblHospitalName 
         AutoSize        =   -1  'True
         Caption         =   "HospitalName"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmIssue2Hospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdd_Click()
txtHospitalName.Text = Clear
txtHospitalID.Text = Clear
txtBloodBagID.Text = Clear
txtBloodGroup.Text = Clear
txtRh.Text = Clear
txtProduct.Text = Clear
txtEnteredBy.Text = Clear
End Sub

Private Sub cmdIssue_Click()
Dim strsql As String

If Validation = False Then
MsgBox "blood bag cannot be issued"
Exit Sub
End If



If txtHospitalID.Text <> 0 Then
Con.BeginTrans

strsql = "UPDATE tbl_BloodBagDetails Set Status = 'Issued' Where BloodBagID = '" & txtBloodBagID & "'"
Con.Execute strsql

strsql = "INSERT INTO tbl_BBIssue(BloodBagID,IssueHospitalName,IssueHospitalID,Product,IssueEntryBy,IssueEntryTime)VALUES('" & txtBloodBagID.Text & "','" _
& txtHospitalName.Text & "' , '" & txtHospitalID.Text & "' , '" & txtProduct.Text & "' , '" & txtEnteredBy.Text & "','" & Now() & "')"
Con.Execute strsql
Con.CommitTrans

End If
MsgBox "blood bag issued"

End Sub

Private Sub txtBloodGroup_DblClick()
Dim strsql As String
strsql = "SELECT BloodGroup FROM tbl_BloodBagDetails WHERE BloodBagID='" & txtBloodBagID & "'"

Dim rssearch As New ADODB.Recordset
rssearch.Open strsql, Con, adOpenKeyset, adLockOptimistic

txtBloodGroup.Text = rssearch("BloodGroup") & ""

rssearch.Close
Set rssearch = Nothing
End Sub

Private Sub txtHospitalName_DblClick()
Dim strsql As String
strsql = "SELECT Hospital from tbl_HospitalInfo WHERE HospitalID='" & txtHospitalID & "'"

Dim rssearch As New ADODB.Recordset
rssearch.Open strsql, Con

txtHospitalName.Text = rssearch("Hospital") & ""
rssearch.Close
Set rssearch = Nothing
End Sub

Private Sub txtRh_DblClick()
Dim strsql As String
strsql = "SELECT Rh FROM tbl_BloodBagDetails WHERE BloodBagID='" & txtBloodBagID & "'"

Dim rssearch As New ADODB.Recordset
rssearch.Open strsql, Con, adOpenKeyset, adLockOptimistic

txtRh.Text = rssearch("Rh") & ""

rssearch.Close
Set rssearch = Nothing
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


