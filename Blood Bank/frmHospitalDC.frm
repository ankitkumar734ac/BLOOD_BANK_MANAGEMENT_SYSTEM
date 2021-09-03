VERSION 5.00
Begin VB.Form frmHospitalDC 
   Caption         =   "HOSPITAL DC"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   6900
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtHospital 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtEnteredBy 
         Height          =   285
         Left            =   4560
         TabIndex        =   16
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtIssueDate 
         Height          =   285
         Left            =   4560
         TabIndex        =   15
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton cmdIssue 
         Caption         =   "Issue"
         Height          =   375
         Left            =   5040
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtProduct 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtlQuantity 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cmbBloodGroup 
         Height          =   315
         ItemData        =   "frmHospitalDC.frx":0000
         Left            =   4560
         List            =   "frmHospitalDC.frx":0010
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmbRh 
         Height          =   315
         ItemData        =   "frmHospitalDC.frx":0021
         Left            =   6120
         List            =   "frmHospitalDC.frx":002B
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtBloodBagNo 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblIssueDate 
         AutoSize        =   -1  'True
         Caption         =   "Issue Date"
         Height          =   195
         Left            =   3600
         TabIndex        =   14
         Top             =   840
         Width           =   765
      End
      Begin VB.Label lblHospital 
         AutoSize        =   -1  'True
         Caption         =   "Hospital"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   570
      End
      Begin VB.Label lblEnteredBy 
         AutoSize        =   -1  'True
         Caption         =   "Entered By"
         Height          =   195
         Left            =   3600
         TabIndex        =   11
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label lblProduct 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblRH 
         AutoSize        =   -1  'True
         Caption         =   "RH"
         Height          =   195
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblBloodGroup 
         AutoSize        =   -1  'True
         Caption         =   "Blood Group"
         Height          =   195
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblBloodBagNo 
         AutoSize        =   -1  'True
         Caption         =   "Blood Bag Number"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmHospitalDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdIssue_Click()
Dim strsql As String
Con.BeginTrans

strsql = "UPDATE tbl_BloodBag SET(Status='Issued',Product='" & txtProduct.Text & "',IssueDate='" & txtIssueDate.Text & "',IssueEntryBy='" & txtEnteredBy.Text & "',IssuedToHospital='" & txtHospital.Text & "') WHERE BloodBagID = '" & txtBloodBagNo.Text & "' "

Con.Execute strsql
Con.CommitTrans
MsgBox "Issue Data Saved"
End Sub
'WHERE (BloodBagID = '" & txtBloodBagNo.Text & "')"
