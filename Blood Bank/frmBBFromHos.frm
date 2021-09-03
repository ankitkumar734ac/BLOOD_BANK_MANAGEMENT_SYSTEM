VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBFromHos 
   Caption         =   "BLOOD BAG FROM HOSPITAL"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   Icon            =   "frmBBFromHos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   8475
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   8295
      Begin MSComCtl2.DTPicker dtpexpdate 
         Height          =   375
         Left            =   5520
         TabIndex        =   5
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108986369
         CurrentDate     =   39161
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add "
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
         Left            =   4200
         TabIndex        =   23
         Top             =   3240
         Width           =   975
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
         Left            =   6600
         TabIndex        =   22
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save  "
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
         Left            =   5400
         TabIndex        =   21
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtEnteredBy 
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
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ComboBox cmbRh 
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
         ItemData        =   "frmBBFromHos.frx":25CA
         Left            =   7200
         List            =   "frmBBFromHos.frx":25D4
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbBloodGroup 
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
         ItemData        =   "frmBBFromHos.frx":25E2
         Left            =   5520
         List            =   "frmBBFromHos.frx":25F2
         TabIndex        =   7
         Top             =   1680
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
         Height          =   450
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtStatus 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtTestResults 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpcollectiondate 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   3120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108986369
         CurrentDate     =   36526
      End
      Begin VB.TextBox txtQuantity 
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
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtBloodBagNo 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   3120
         Width           =   4095
      End
      Begin VB.Label lblEnteredBy 
         AutoSize        =   -1  'True
         Caption         =   "Entered By"
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
         Left            =   3960
         TabIndex        =   20
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblRh 
         Caption         =   "RH"
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
         Left            =   6720
         TabIndex        =   19
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label lblBloodGroup 
         Caption         =   "Blood Group"
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
         Left            =   3960
         TabIndex        =   18
         Top             =   1800
         Width           =   1455
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
         Left            =   3960
         TabIndex        =   17
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label lblDateOfExpiry 
         AutoSize        =   -1  'True
         Caption         =   "Expiry Date  "
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
         Left            =   3960
         TabIndex        =   16
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status "
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
         TabIndex        =   15
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblTestResults 
         AutoSize        =   -1  'True
         Caption         =   "Test Results "
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
         TabIndex        =   14
         Top             =   1800
         Width           =   1395
      End
      Begin VB.Label lblDateOfCol 
         AutoSize        =   -1  'True
         Caption         =   "Collection Date"
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
         TabIndex        =   13
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         Caption         =   "Quantity            "
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
         Left            =   75
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblBloodBagNo 
         AutoSize        =   -1  'True
         Caption         =   "Blood Bag ID      "
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
         Left            =   75
         TabIndex        =   11
         Top             =   360
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmBBFromHos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmBBFromHospital

'Displays:  Form will be displayed when user clicks on the Blood Bag Entry command button of the Form frmHospitalLookUp

'Unload:    When the user is saved the blood bag details that is provided by Hospital, he can click the Cancel command button and choose the Yes option to terminate from the Form frmBBFromHos

'Functions: This form is used to save the Blood bag details that are provided by Hospital.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Lets the user add the data to the form ***********************'

' This procedure is used to set the default values to all the controls
' to let the user add the data in the form.

Private Sub cmdAdd_Click()
Dim n, max As Integer
Dim i As Integer

'Clears the Controls to set the default values.

txtBloodBagNo.Text = ""
txtQuantity.Text = ""
txtEnteredBy.Text = ""
cmbBloodGroup.Text = ""
cmbRh.Text = ""
dtpexpdate.Value = Date

'Auto generation code for Blood Bag ID's

max = 0
open_recordset "tbl_BloodBagDetails"
n = rs.RecordCount
If n = 0 Then
 txtBloodBagNo.Text = 1
 close_recordset
Else
 For i = 0 To rs.RecordCount - 1
 If (max < rs("BloodBagID")) Then
  max = rs("BloodBagID")
 End If
 Next i
 txtBloodBagNo.Text = rs.RecordCount + 1
End If

close_recordset
End Sub

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "DONOR DETIALS")
Select Case response                ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                    'Unload the current form
        
  Case vbNo:
  
        'Clears the Controls to set the default values.

        txtBloodBagNo.Text = ""
        dtpexpdate.Value = Date
        dtpcollectiondate.Value = Date
        txtEnteredBy.Text = ""
        txtQuantity.Text = ""
        txtHospitalID.Text = ""
        txtStatus.Text = ""
        txtTestResults.Text = ""
        cmbBloodGroup.Text = ""
        cmbRh.Text = ""
 
 End Select
End Sub

'*********************** Save ***********************'

' This procedure is used to insert
' the Data in the form into the database.

Private Sub cmdSave_Click()
Dim strsql As String
On Error GoTo errPara
If Validatedata = False Then
    Exit Sub
Else
Con.BeginTrans
 
'Insert Blood Bag Details into database.
 strsql = "INSERT INTO tbl_BloodBagDetails(BloodBagID,BloodGroup,Rh,Quantity,DateOfCollection,DateOfExpiry,TestResults," _
 & "Status,EnteredBy,Entrydate)VALUES('" & txtBloodBagNo.Text & "' , '" & cmbBloodGroup.Text & "' , '" & cmbRh.Text & "' , '" _
 & txtQuantity.Text & "' , '" & dtpcollectiondate.Value & "' , '" & dtpexpdate.Value & "' , '" & txtTestResults.Text & "' ,  '" _
 & txtStatus.Text & "' , '" & txtEnteredBy.Text & "' , '" & Now() & "')"
 Con.Execute strsql

 'Insert Blood Bag Details into Database for Quick access.
 strsql = "INSERT INTO tbl_BBHospital(BloodBagID,HospitalID,EnteredDate) VALUES ('" & txtBloodBagNo.Text & "' ,'" & txtHospitalID & "','" & Now() & "')"
 Con.Execute strsql
Con.CommitTrans
 MsgBox "data saved"
End If
 cmdAdd.Enabled = True
 Exit Sub
 
' Code provided to handle the Errors.

errPara:
MsgBox Err.Description
    
    Con.RollbackTrans
    Load MDIForm1
    Exit Sub
End Sub

'*********************** Expiry Date ***********************'

' This procedure is used to calculate and assign the expiry date for each blood bag.

Private Sub dtpexpdate_Click()
 dtpexpdate.Value = dtpcollectiondate.Value + 35  'Expiry Date is 35 days after the Collection Date.
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to load and show the default settings

Private Sub Form_Load()
Dim n, max As Integer
Dim i As Integer

'Setting the dimensions of the form.

frmBBFromHos.Height = 4545
frmBBFromHos.Width = 8595
CenterForm frmBBFromHos

'Fills the existing details of the form.

dtpcollectiondate.Value = Date
dtpexpdate.Value = Date
fillhosdetails strHospitalID
txtTestResults.Text = "NEGATIVE"

'Auto Generation code for Blood Bag ID's
max = 0
open_recordset "tbl_BloodBagDetails"
n = rs.RecordCount
If n = 0 Then
 txtBloodBagNo.Text = 1
 close_recordset
Else
 For i = 0 To rs.RecordCount - 1
 If (max < rs("BloodBagID")) Then
  max = rs("BloodBagID")
 End If
 Next i
 txtBloodBagNo.Text = rs.RecordCount + 1
End If

close_recordset
End Sub

'*********************** Quantity of the Blood Bag ***********************'

' This procedure is used to assign the quantity of the Blood Bag.

Private Sub txtQuantity_GotFocus()
 txtQuantity.Text = "350ml"            'Standard Quantity of the Blood Bag is 350ml.
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate the entries before storing it in database

Public Function Validatedata() As Boolean
Validatedata = True

'Validate the Blood Bag Id
If Len(Trim(txtBloodBagNo.Text)) = 0 Then
    MsgBox "Please enter Blood Bag ID"
    Validatedata = False
    txtBloodBagNo.SetFocus
    Exit Function
End If

'Validates the Quantity
If Len(Trim(txtQuantity.Text)) = 0 Then
    MsgBox "Please enter Quantity"
    Validatedata = False
    txtQuantity.SetFocus
    Exit Function
End If

'Validates the Status
If Len(Trim(txtStatus.Text)) = 0 Then
    MsgBox "Please click to get the status of the blood bag "
    Validatedata = False
    txtStatus.SetFocus
    Exit Function
End If

'Validates the Blood Group.
If Len(Trim(cmbBloodGroup.Text)) = 0 Then
    MsgBox "Please select BloodGroup."
    Validatedata = False
    cmbBloodGroup.SetFocus
    Exit Function
End If

'Validates the RH Factor
If Len(Trim(cmbRh.Text)) = 0 Then
    MsgBox "Please select Rh."
    Validatedata = False
    cmbRh.SetFocus
    Exit Function
End If

'Validates the name of the DEO
If Len(Trim(txtEnteredBy.Text)) = 0 Then
    MsgBox "Please enter DEO's Name."
    Validatedata = False
   txtEnteredBy.SetFocus
    Exit Function
End If
End Function

'*********************** Fills the Hospital Details in the Form ***********************'

' This procedure is used to fill the existing hospital details from database into the Form

Private Sub fillhosdetails(strHospitalID As String)
Dim strsql As String
Dim rsDetail As New ADODB.Recordset

If strHospitalID = "" Then
  Exit Sub
Else
 rsDetail.Open "SELECT * FROM tbl_HospitalInfo WHERE HospitalID='" & strHospitalID & "'", Con, adOpenKeyset, adLockOptimistic
 txtHospitalID.Text = rsDetail("HospitalID")
End If
Set rsDetail = Nothing
End Sub

'*********************** Status of the Blood Bag ***********************'

'This procedure is used to assign the status of the Blood Bag.

Private Sub txtStatus_GotFocus()
 If txtTestResults.Text = "NEGATIVE" Then           ' If Test Results are Negetive then
     txtStatus.Text = "Existing"                    ' the Status is Existing
 Else                                               ' Else
    txtStatus.Text = "Discarded"                    ' Discard the Blood Bags.
End If
End Sub
