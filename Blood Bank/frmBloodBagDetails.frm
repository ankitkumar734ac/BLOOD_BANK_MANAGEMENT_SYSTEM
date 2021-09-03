VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBFromDonor 
   Caption         =   "BLOOD BAG FROM DONOR"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   7650
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txtStatus 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3000
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dprDateOfCol 
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36526
      End
      Begin VB.TextBox txtRh 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox txtBloodGroup 
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtDateOfExp 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   4080
         TabIndex        =   16
         Top             =   360
         Width           =   2895
         Begin VB.TextBox txtDonorID 
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblDonorID 
            AutoSize        =   -1  'True
            Caption         =   "DonorID"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   600
         End
         Begin VB.Label Label1 
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   3600
         Width           =   1215
      End
      Begin VB.ComboBox cmbTestResults 
         Height          =   315
         ItemData        =   "frmBloodBagDetails.frx":0000
         Left            =   1800
         List            =   "frmBloodBagDetails.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtEnteredBy 
         Height          =   285
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtQuantity 
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1000
         Width           =   1575
      End
      Begin VB.TextBox txtBloodBagNo 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblDateOfExpiry 
         Caption         =   "Date Of Expiry"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3000
         Width           =   450
      End
      Begin VB.Label lblEnteredBy 
         AutoSize        =   -1  'True
         Caption         =   "Entered By"
         Height          =   195
         Left            =   4080
         TabIndex        =   15
         Top             =   2880
         Width           =   780
      End
      Begin VB.Label lblTestResults 
         AutoSize        =   -1  'True
         Caption         =   "Test Results"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label lblDateOfCol 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Collection"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1640
         Width           =   1410
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         Caption         =   "Quantity (ml)"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1005
         Width           =   870
      End
      Begin VB.Label lblRh 
         AutoSize        =   -1  'True
         Caption         =   "RH"
         Height          =   195
         Left            =   6000
         TabIndex        =   11
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label lblBloodGroup 
         AutoSize        =   -1  'True
         Caption         =   "Blood Group"
         Height          =   195
         Left            =   4080
         TabIndex        =   10
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label lblBloodBagNo 
         Caption         =   "BloodBag ID"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmBBFromDonor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancel_Click()
Unload Me
strDonorId = ""
End Sub

Private Sub cmdSave_Click()
Dim strsql As String
If Validatedata = False Then
    Exit Sub
Else

 
 
 Con.BeginTrans
 strsql = "INSERT INTO tbl_BloodBagDetails(BloodBagID,BloodGroup,Rh,Quantity,DateOfCollection,DateOfExpiry,TestResults," _
 & "Status,EnteredBy,EntryDate)VALUES('" & txtBloodBagNo.Text & "' , '" & txtBloodGroup.Text & "' , '" & txtRh.Text & "' , '" _
 & txtQuantity.Text & "' , '" & dprDateOfCol.Value & "' , '" & txtDateOfExp.Text & "' , '" & cmbTestResults.Text & "' ,  '" _
 & txtStatus.Text & "' , '" & txtEnteredBy.Text & "' , '" & Now() & "')"
 Con.Execute strsql
 
 strsql = "INSERT INTO tbl_BBDonor(BloodBagID,DonorID) VALUES ('" & txtBloodBagNo.Text & "' ,'" & txtDonorID & "')"
 Con.Execute strsql
 
 Con.CommitTrans
 MsgBox "data saved"
 
' ElseIf optHospital.Value = True Then
'
' Con.BeginTrans
' strsql = "INSERT INTO tbl_BloodBagDetails(BloodBagID,BloodGroup,Rh,Quantity,DateOfCollection,DateOfExpiry,TestResults," _
' & "Status,EnteredBy,EntryDate)VALUES('" & txtBloodBagNo.Text & "' , '" & cmbBloodGroup.Text & "' , '" & cmbRh.Text & "' , '" _
' & txtQuantity.Text & "' , '" & dprDateOfCol.Value & "' , '" & txtDateOfExp.Text & "' , '" & cmbTestResults.Text & "' ,  '" _
' & txtStatus.Text & "' , '" & txtEnteredBy.Text & "' , '" & Now() & "')"
' Con.Execute strsql
'
' strsql = "INSERT INTO tbl_BBHospital(BloodBagID,HospitalID) VALUES ('" & txtBloodBagNo.Text & "' ,'" & txtHospitalID & "')"
' Con.Execute strsql
'
' Con.CommitTrans
' MsgBox "data saved"
'
 
 End If
End Sub

Private Sub Form_Load()
frmBloodBagDetails.Height = 4695
frmBloodBagDetails.Width = 7770
CenterForm frmBloodBagDetails


fillDetails strDonorId

End Sub

Private Sub txtDateOfExp_Click()
If cmbTestResults.Text = "POSITIVE" Then
txtDateOfExp.Text = dprDateOfCol.Value
ElseIf cmbTestResults.Text = "NEGATIVE" Then
txtDateOfExp.Text = dprDateOfCol.Value + 35
End If
End Sub








Private Sub txtStatus_Click()
If cmbTestResults.Text = "POSITIVE" Then
     txtStatus.Text = "Discarded"
ElseIf cmbTestResults.Text = "NEGATIVE" Then
     txtStatus.Text = "Existing"
End If
End Sub


Private Sub txtQuantity_keypress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If (keyascii < 48 Or keyascii > 57) Then
    MsgBox "Please enter only digits."
    keyascii = 0
End If
End Sub


Public Function Validatedata() As Boolean
Validatedata = True

'If optHospital.Value = True Then
'If Len(Trim(cmbBloodGroup.Text)) = 0 Then
'    MsgBox "Please select Blood Group."
'    Validatedata = False
'    cmbBloodGroup.SetFocus
'    Exit Function
'
'ElseIf Len(Trim(cmbRh.Text)) = 0 Then
'    MsgBox "Please select Rh."
'    Validatedata = False
'    cmbRh.SetFocus
'    Exit Function
'End If
'End If

If Len(Trim(txtBloodBagNo.Text)) = 0 Then
    MsgBox "Please enter Blood Bag ID"
    Validatedata = False
    txtBloodBagNo.SetFocus
    Exit Function
End If




If Len(Trim(txtQuantity.Text)) = 0 Then
    MsgBox "Please enter Quantity"
    Validatedata = False
    txtQuantity.SetFocus
    Exit Function
End If

If Len(Trim(cmbTestResults.Text)) = 0 Then
    MsgBox "Please select test results"
    Validatedata = False
    txtBloodBagNo.SetFocus
    Exit Function
End If

If Len(Trim(txtStatus.Text)) = 0 Then
    MsgBox "Please click to get the status of the blood bag "
    Validatedata = False
    txtStatus.SetFocus
    Exit Function
End If

If Len(Trim(txtDateOfExp.Text)) = 0 Then
    MsgBox "Please click to get the date of expiry of the blood bag "
    Validatedata = False
    txtDateOfExp.SetFocus
    Exit Function
End If


If Len(Trim(txtEnteredBy.Text)) = 0 Then
    MsgBox "Please enter DEO's Name."
    Validatedata = False
   txtEnteredBy.SetFocus
    Exit Function
End If






End Function




Private Sub fillDetails(strDonorId As String)
Dim strsql As String
Dim rsDetail As New ADODB.Recordset

If strDonorId = "" Then Exit Sub
 

   strsql = "SELECT * FROM tbl_Donor WHERE DonorID='" & strDonorId & "'"
    rsDetail.Open strsql, Con, adOpenKeyset, adLockOptimistic
'    If rsDetail.BOF And rsDetail.EOF Then Exit Sub
    txtDonorID.Text = rsDetail("DonorID")
    txtBloodGroup.Text = rsDetail("BloodGroup")
    txtRh.Text = rsDetail("RH")
   
End Sub


