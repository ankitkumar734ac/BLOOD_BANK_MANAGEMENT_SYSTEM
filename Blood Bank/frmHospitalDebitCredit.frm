VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmHospitalDebitCredit 
   Caption         =   "HOSPITAL  DEBIT / CREDIT"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   5745
   Begin MSFlexGridLib.MSFlexGrid msflgHosDebit 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid msflgHosCredit 
      Height          =   1935
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewDebit 
         Caption         =   "Veiw Debit"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewCredit 
         Caption         =   "View Credit"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtHospitalID 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblHospitalID 
         AutoSize        =   -1  'True
         Caption         =   "HospitalID"
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmHospitalDebitCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdViewCredit_Click()
Dim strsql As String
Dim srno As Integer
Dim rsvc As New ADODB.Recordset

'On Error GoTo errPara
    If Len(Trim(txtHospitalID)) <> 0 Then
        strsql = "SELECT HospitalID,BloodBagID,IssueDate FROM tbl_HospitalIssue WHERE HospitalID='" & txtHospitalID & "'"
    End If
    If Len(strsql) = 0 Then
        MsgBox "Please enter search data."
        Exit Sub
    End If

    rsvc.Open strsql, Con, adOpenDynamic, adLockOptimistic
'    rssearch.MoveFirst
    srno = 1
    msflgHosCredit.Clear
    msflgHosCredit.Rows = 2
    
    
   msflgHosCredit.FormatString = "SNo |Hospital Id | Blood Bag ID | Issue Date"
    msflgHosCredit.ColWidth(2) = 2000
    
    If rsvc.BOF And rsvc.EOF Then MsgBox "No records found for this criteria."
        
    While Not rsvc.EOF
        msflgHosCredit.TextMatrix(srno, 0) = srno
        msflgHosCredit.TextMatrix(srno, 1) = rsvc("HospitalID") & ""
        msflgHosCredit.TextMatrix(srno, 2) = rsvc("BloodBagID") & ""
        msflgHosCredit.TextMatrix(srno, 3) = rsvc("IssueDate") & ""
       
        rsvc.MoveNext
        srno = srno + 1
        If msflgHosCredit.Rows = srno Then msflgHosCredit.Rows = msflgHosCredit.Rows + 1
    Wend
    rsvc.Close
    Set rsvc = Nothing
Exit Sub

End Sub

Private Sub cmdViewDebit_Click()
Dim strsql As String
Dim srno As Integer
Dim rsvc As New ADODB.Recordset

'On Error GoTo errPara
    If Len(Trim(txtHospitalID)) <> 0 Then
        strsql = "SELECT tbl_BBHospital.HospitalID,tbl_BBHospital.BloodBagID,tbl_BloodBagDetails.DateOfCollection FROM tbl_BBHospital,tbl_BloodBagDetails WHERE tbl_BBHospital.BloodBagID=tbl_BloodBagDetails.BloodBagID"
    End If
    If Len(strsql) = 0 Then
        MsgBox "Please enter search data."
        Exit Sub
    End If

    rsvc.Open strsql, Con, adOpenDynamic, adLockOptimistic
'    rssearch.MoveFirst
    srno = 1
    msflgHosDebit.Clear
    msflgHosDebit.Rows = 2
    
    
   msflgHosDebit.FormatString = "SNo |Hospital Id | Blood Bag ID | Date Of Collection"
    msflgHosDebit.ColWidth(2) = 2000
    
    If rsvc.BOF And rsvc.EOF Then MsgBox "No records found for this criteria."
        
    While Not rsvc.EOF
        msflgHosDebit.TextMatrix(srno, 0) = srno
        msflgHosDebit.TextMatrix(srno, 1) = rsvc("HospitalID") & ""
        msflgHosDebit.TextMatrix(srno, 2) = rsvc("BloodBagID") & ""
        msflgHosDebit.TextMatrix(srno, 3) = rsvc("DateOfCollection") & ""
       
        rsvc.MoveNext
        srno = srno + 1
        If msflgHosDebit.Rows = srno Then msflgHosDebit.Rows = msflgHosDebit.Rows + 1
    Wend
    rsvc.Close
    Set rsvc = Nothing
Exit Sub

End Sub

Private Sub Form_Load()
CenterForm frmHospitalDebitCredit
frmHospitalDebitCredit.Height = 6300
frmHospitalDebitCredit.Width = 5865
End Sub
