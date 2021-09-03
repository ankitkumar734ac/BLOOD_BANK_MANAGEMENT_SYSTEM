VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBBFromDonor 
   Caption         =   "BLOOD BAG FROM DONOR"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "frmBBFromDonor.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4095
   ScaleWidth      =   8160
   Begin VB.Frame Frame1 
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
      Height          =   4215
      Left            =   0
      TabIndex        =   2
      Top             =   -120
      Width           =   8295
      Begin MSComCtl2.DTPicker dtpexpdate 
         Height          =   495
         Left            =   5520
         TabIndex        =   6
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108986369
         CurrentDate     =   39177
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
         TabIndex        =   4
         Top             =   1800
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
         Height          =   435
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2520
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpcollectiondate 
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   3240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   108986369
         CurrentDate     =   39177
      End
      Begin VB.TextBox txtRh 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtBloodGroup 
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   20
         Top             =   960
         Width           =   4215
         Begin VB.TextBox txtDonorID 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblDonorID 
            AutoSize        =   -1  'True
            Caption         =   "DonorID"
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
            TabIndex        =   24
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label1 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   375
         End
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
         Height          =   495
         Left            =   6120
         TabIndex        =   12
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save "
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
         Left            =   4200
         TabIndex        =   11
         Top             =   3360
         Width           =   1455
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
         Height          =   435
         Left            =   5520
         MaxLength       =   20
         TabIndex        =   10
         Top             =   2520
         Width           =   2415
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
         Height          =   435
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1080
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
         Height          =   435
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   3240
         Width           =   3975
      End
      Begin VB.Label lblDateOfExpiry 
         Caption         =   "Expiry Date"
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
         TabIndex        =   22
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         TabIndex        =   21
         Top             =   2640
         Width           =   675
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
         TabIndex        =   19
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblTestResults 
         AutoSize        =   -1  'True
         Caption         =   "Test Results"
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
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblDateOfCol 
         AutoSize        =   -1  'True
         Caption         =   "Collection Date "
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
         TabIndex        =   17
         Top             =   3360
         Width           =   1755
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         Caption         =   "Quantity (ml) "
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
         TabIndex        =   16
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label lblRh 
         AutoSize        =   -1  'True
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
         Height          =   315
         Left            =   6720
         TabIndex        =   15
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblBloodGroup 
         AutoSize        =   -1  'True
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
         Height          =   315
         Left            =   3960
         TabIndex        =   14
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label lblBloodBagNo 
         Caption         =   "BloodBag ID"
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
         Top             =   480
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmBBFromDonor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmBBFromDonor

'Displays:  Form will be displayed when user clicks on the Issue Blood Bag command button of the Form frmDonorDetails

'Unload:    When the user is saved the blood bag details that is donated by Donor, he can click the Cancel command button and choose the Yes option to terminate from the Form frmBBFromDonor

'Functions: This form is used to save the Blood bag details that are donated by Donor.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "DONOR DETAILS")
Select Case response         ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me            ' Unloads the current form
        
  Case vbNo:
          
        'Clears the Controls to provide the Default value.
        
        txtBloodBagNo.Text = ""
        txtBloodGroup.Text = ""
        txtDonorID.Text = ""
        txtEnteredBy.Text = ""
        txtQuantity.Text = ""
        txtRh.Text = ""
        txtStatus.Text = ""
        txtTestResults.Text = ""
        dtpcollectiondate.Value = Date
        dtpexpdate.Value = Date
  
 End Select

End Sub

'*********************** Expiry Date ***********************'

' This procedure is used to calculate the Expiry Date

Private Sub dtpexpdate_Click()
If txtTestResults.Text = "POSITIVE" Then           ' If Test Results are positive then
 dtpexpdate.Value = dtpcollectiondate.Value        ' discards the blood bags
ElseIf txtTestResults.Text = "NEGATIVE" Then       ' else
 dtpexpdate.Value = dtpcollectiondate.Value + 35   ' the Expiry date is 35 days after the Collection date.
End If
End Sub

'*********************** Quantity of Blood Bags ***********************'

' This procedure is used to assign the Quantity of the Blood Bags Collected.

Private Sub txtQuantity_GotFocus()
txtQuantity.Text = "350ml"                        ' Standard quantity of the Blood bag is 350 ml.
txtQuantity.Locked = True
End Sub

'*********************** Status of the Blood Bags ***********************'

' This procedure is used to assign the status to the blood bags.
Private Sub txtStatus_GotFocus()
If txtTestResults.Text = "POSITIVE" Then           ' If Test Results are positive then
          txtStatus.Text = "Discarded"             ' discard the Blood bags
ElseIf txtTestResults.Text = "NEGATIVE" Then       ' Else
      txtStatus.Text = "Existing"                  ' Assign the status as Existing
End If
End Sub

'*********************** Insert Blood Bag Details ***********************'

' This procedure is used to insert the details
' of blood bags donated by the Donor into the Database.

Private Sub cmdSave_Click()
Dim strsql As String
On Error GoTo errPara
If Validatedata = False Then
    Exit Sub
Else
 Con.BeginTrans
 
  'Insert the Blood Bag Details into the Database.
 
 strsql = "INSERT INTO tbl_BloodBagDetails(BloodBagID,BloodGroup,Rh,Quantity,DateOfCollection,DateOfExpiry,TestResults," _
 & "Status,EnteredBy,Entrydate)VALUES('" & txtBloodBagNo.Text & "' , '" & txtBloodGroup.Text & "' , '" & txtRh.Text & "' , '" _
 & txtQuantity.Text & "' , '" & dtpcollectiondate.Value & "' , '" & dtpexpdate.Value & "' , '" & txtTestResults.Text & "' ,  '" _
 & txtStatus.Text & "' , '" & txtEnteredBy.Text & "' , '" & Now() & "')"
 Con.Execute strsql
 
  'Insert the Details into the Database for quick access.
 
 strsql = "INSERT INTO tbl_BBDonor(BloodBagID,DonorID,EnteredDate) VALUES ('" & txtBloodBagNo.Text & "' ,'" & txtDonorID & "', '" & Now() & "')"
 Con.Execute strsql
 Con.CommitTrans
 MsgBox "data saved"
 txtQuantity.Locked = False
End If
 
 ' Clear the controls for default values
 
 txtBloodBagNo.Text = ""
 txtBloodGroup.Text = ""
 txtDonorID.Text = ""
 txtEnteredBy.Text = ""
 txtQuantity.Text = ""
 txtRh.Text = ""
 txtStatus.Text = ""
 txtTestResults.Text = ""
 dtpcollectiondate.Value = Date
 dtpexpdate.Value = Date
 Exit Sub
 
 ' Code provided to handle the Errors.

errPara:
    MsgBox Err.Description
    Con.RollbackTrans
    Load MDIForm1
    Exit Sub
End Sub

'*********************** Loads the form ***********************'

' This procedure is used to load and show the form.

Private Sub Form_Load()
Dim n, max As Integer
Dim i As Integer

'Setting the dimensions of the form.

frmBBFromDonor.Height = 4605
frmBBFromDonor.Width = 8280
CenterForm frmBBFromDonor

'Fills the existing details of the form.

strDonorId = frmDonorDetails.txtDonorID.Text
fillDetails strDonorId
If Resultscheck(strDonorId) = False Then
 txtTestResults.Text = "POSITIVE"
Else
 txtTestResults.Text = "NEGATIVE"
End If

' Auto generation code for the Blood Bag ID's.

max = 0
open_recordset "tbl_BloodBagDetails"

n = rs.RecordCount
If n = 0 Then
 txtBloodBagNo.Text = 1
Else
  For i = 0 To rs.RecordCount - 1
  If (max < rs("BloodBagID")) Then
  max = rs("BloodBagID")
  End If
  Next i
  txtBloodBagNo.Text = rs.RecordCount + 1
End If
dtpexpdate.Value = Date
dtpcollectiondate.Value = Date
close_recordset
frmDonorDetails.Hide
End Sub

'Private Sub txtQuantity_keypress(keyascii As Integer)
'If keyascii = 8 Then Exit Sub
'If (keyascii < 48 Or keyascii > 57) Then
 '   MsgBox "Please enter only digits."
  '  keyascii = 0
'End If
'End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate the Data
' that is to be inserted into the database.

Public Function Validatedata() As Boolean
Validatedata = True

'Validates the Blood Bag ID.
If Len(Trim(txtBloodBagNo.Text)) = 0 Then
    MsgBox "Please enter Blood Bag ID"
    Validatedata = False
    txtBloodBagNo.SetFocus
    Exit Function
End If

'Validates the Quantity of Blood Bag.
If Len(Trim(txtQuantity.Text)) = 0 Then
    MsgBox "Please enter Quantity"
    Validatedata = False
    txtQuantity.SetFocus
    Exit Function
End If

'Validates the Status of the Blood Bag.
If Len(Trim(txtStatus.Text)) = 0 Then
    MsgBox "Please focus on status text box to get the status of the blood bag "
    Validatedata = False
    txtStatus.SetFocus
    Exit Function
End If

'Validates the name of the DEO.
If Len(Trim(txtEnteredBy.Text)) = 0 Then
    MsgBox "Please enter DEO's Name."
    Validatedata = False
   txtEnteredBy.SetFocus
    Exit Function
End If

End Function

'*********************** Fills the Donor Details into the Form ***********************'

' This procedure is used to fill the Donor Details in the form

Private Function fillDetails(strDonorId As String)
Dim rsDetail As New ADODB.Recordset

If strDonorId = "" Then
  Exit Function
Else
   rsDetail.Open "SELECT * FROM tbl_Donor WHERE DonorID='" & strDonorId & "'", Con, adOpenKeyset, adLockOptimistic
    txtDonorID.Text = rsDetail("DonorID")
    txtBloodGroup.Text = rsDetail("BloodGroup")
    txtRh.Text = rsDetail("RH")
End If
Set rsDetail = Nothing
End Function

'*********************** Checks the Blood Test Result ***********************'

' This procedure is used to check whether
' the donated blood is unaffected with the serious diseases.

Public Function Resultscheck(strDonorId As String) As Boolean
Dim strsql As String
Dim rsres As New ADODB.Recordset
Dim count As Integer
count = 0
If strDonorId = "" Then Exit Function

strsql = "select * from tbl_DonorTestResults where DonorID='" & strDonorId & "'"
rsres.Open strsql, Con, adOpenKeyset, adLockOptimistic

If rsres("TestVDRL") = True Then      'VDRL --> Vinryl Disease Reasearch Laboratory.
count = count + 1
End If
If rsres("TestHBsAG") = True Then     'HB's AG --> Hepatatis B surface AnteGen.
count = count + 1
End If

If rsres("TestMP") = True Then        'MP --> Malarial parasites
count = count + 1
End If

If rsres("TestHCV") = True Then       'HCV --> Hepatatis C Virus
count = count + 1
End If
 
If rsres("TestHIV") = True Then       'HIV --> Human Immune Virus
count = count + 1
End If

If count > 0 Then
 Resultscheck = False
Else
 Resultscheck = True
End If
rsres.Close
Set rsres = Nothing
End Function
