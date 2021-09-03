VERSION 5.00
Begin VB.Form frmdonorreport 
   Caption         =   "Donor Report"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3450
   Icon            =   "frmDonorReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2025
   ScaleWidth      =   3450
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin VB.ComboBox cbodonorid 
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
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdcancel 
         Cancel          =   -1  'True
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdreport 
         Caption         =   "REPORT GENERATION"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Donor ID :"
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
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmdonorreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmdonorreport

'Displays:  Form will be displayed when user clicks donor report menu on the the MDI Form

'Unload  :  When the user has generated the report for donor details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmdonorreport

'Functions: This form is used to generate the report for Donor.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Retrives the Donor ID's ***********************'

' This procedure is used to retrive the Dononr ID's from the Database.

Private Sub cbodonorid_GotFocus()
Dim rscbo As New ADODB.Recordset
Dim strquery As String
 rscbo.Open "Select * from tbl_Donor", Con, adOpenKeyset, adLockOptimistic
 If cbodonorid.ListCount > 0 Then
  Set rscbo = Nothing
  Exit Sub
 Else
  For i = 0 To rscbo.RecordCount
   cbodonorid.AddItem rscbo("DonorID")
   rscbo.MoveNext
   If rscbo.EOF Then
    Exit For
   End If
  Next i
End If
Set rscbo = Nothing
End Sub

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit?", vbQuestion + vbYesNo, "CANCEL")
Select Case response             ' Provides the MsgBox response as the expression to match for the Case Structure.
 Case vbYes:
       Unload Me                 ' Unloads the Current Form
 Case vbNo:
 
        'Clears the Controls to set the default values.      cbodonorid.Text = ""
        cbodonorid.SetFocus
End Select
End Sub

'*********************** Report Generation ***********************'

' This procedure is used to generate the Report.

Private Sub cmdreport_Click()
Dim rsnew As New ADODB.Recordset

'Code to show the Report for Donor.
If Len(Trim(cbodonorid.Text)) = 0 Then
  MsgBox "Please Enter the Donor ID:", vbInformation, "Donor Report"
Else
  rsnew.Open "select * from tbl_Donor,tbl_DonorTestResults where tbl_Donor.DonorID = tbl_DonorTestResults.DonorID AND tbl_Donor.DonorID='" & cbodonorid.Text & "'", Con, adOpenKeyset, adLockOptimistic
 If rsnew.RecordCount = 0 Then
    MsgBox "No Data To Show.", vbOKOnly + vbInformation, "Empty Database"
 Else
     Set DRdonorreport.DataSource = rsnew
     DRdonorreport.Show
 End If
End If
Set rsnew = Nothing
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'set the Dimensions for the Form
frmdonorreport.Height = 2535
frmdonorreport.Width = 3570
CenterForm frmdonorreport

End Sub
