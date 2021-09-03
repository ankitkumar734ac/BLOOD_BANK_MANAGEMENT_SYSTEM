VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBBLookUp 
   Caption         =   "Look Up"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBBLookUp.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   9225
   Begin VB.CommandButton cmdSearchExp 
      Caption         =   "BLOOD BAGS EXPIRED TODAY"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   10
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE BLOOD BAG"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdSearchExisting 
      Caption         =   "SEARCH EXISTING"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search By"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox cmbRh 
         Enabled         =   0   'False
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
         ItemData        =   "frmBBLookUp.frx":25CA
         Left            =   6360
         List            =   "frmBBLookUp.frx":25D4
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox cmbBloodGroup 
         Enabled         =   0   'False
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
         ItemData        =   "frmBBLookUp.frx":25E2
         Left            =   4920
         List            =   "frmBBLookUp.frx":25F2
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtBloodBagID 
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   5
         Top             =   735
         Width           =   2415
      End
      Begin VB.OptionButton optBloodGroup 
         Caption         =   "Blood Group "
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optBloodBagID 
         Caption         =   "BloodBag ID "
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "SEARCH "
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
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid msflgBBSearch 
      Height          =   1215
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   7
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBBLookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Form Name: frmBBLookUp

'Displays:  Form will be displayed when user clicks on the Blood Bag Look Up Menu of the MDI Form

'Unload  :  When the user has accessed the blood bag details , user can click the Cancel command button and choose the Yes option to terminate from the Form frmBBLookUp

'Functions: This form is used to have the Quick access to the existing Blood bag details from the database .

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit

'*********************** Cancel ***********************'

' This procedure is used to come out of the form.

Private Sub cmdcancel_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "BLOOD BAG LOOK UP")
Select Case response                   ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                      ' Unload the current form
        
  Case vbNo:
  
        'Clears the Controls to set the default values.
        
        optBloodBagID.Enabled = True
        optBloodGroup.Enabled = True
        optBloodBagID.Value = False
        optBloodGroup.Value = False
        txtBloodBagID.Text = ""
        cmbBloodGroup.Text = ""
        cmbRh.Text = ""
        msflgBBSearch.Clear
        msflgBBSearch.Rows = 2
       
 End Select
End Sub

'*********************** Delete ***********************'

' This procedure is used to delete
' the existing records of the Blood Bags from the database.

Private Sub cmddelete_Click()
Dim strsql As String
Dim strnew As String
Dim response As String
Dim i As Integer
strnew = msflgBBSearch.TextMatrix(msflgBBSearch.Row, 1)
If Len(strnew) = 0 Then
 MsgBox "select BloodBag ID"
Else
response = MsgBox("Do you want to delete the record.", vbQuestion + vbYesNo, "BBLookUp")
Select Case response
Case vbYes:
   Con.BeginTrans
   MsgBox "Record No is" & strnew, vbInformation, "BBLookUp"
   
   'Deletes the Existing Blood Bag Record.
   strsql = "DELETE  FROM tbl_BloodBagDetails WHERE BloodBagID=" & Val(strnew)
   Con.Execute strsql
   Con.CommitTrans
   MsgBox "record deleted"
   For i = 1 To msflgBBSearch.Cols - 1
         msflgBBSearch.TextMatrix(msflgBBSearch.Row, i) = ""
   Next i
   txtBloodBagID.Text = ""
End Select
End If
End Sub

'*********************** Search Existing ***********************'

' This procedure is used to search the blood bags that are still existing.

Private Sub cmdSearchExisting_Click()
cmdDelete.Enabled = False
Dim strsql As String
Dim srno As Integer
Dim rssearch As New ADODB.Recordset
On Error GoTo errPara

'Input Validation for the search criteria
    'Checks for the Blood Bag ID
    If optBloodGroup.Value = False And optBloodBagID.Value = True Then
     If (Len(Trim(txtBloodBagID.Text)) = 0) Then
       MsgBox "First enter the Blood Bag ID", vbInformation, "BloodBagID"
       txtBloodBagID.SetFocus
       Exit Sub
     Else
       rssearch.Open "SELECT * FROM tbl_BloodBagDetails WHERE BloodBagID=" & Val(txtBloodBagID.Text) & " And Status='Existing'", Con, adOpenKeyset, adLockOptimistic
     End If
     'Check for the Blood Group
   ElseIf optBloodBagID.Value = False And optBloodGroup.Value = True Then
       If Len(Trim(cmbBloodGroup.Text)) <> 0 Then
         If Len(Trim(cmbRh.Text)) <> 0 Then
           rssearch.Open "SELECT * FROM tbl_BloodBagDetails WHERE BloodGroup='" & cmbBloodGroup.Text & "' And RH='" & cmbRh.Text & "' And Status='Existing'", Con, adOpenKeyset, adLockOptimistic
         Else
           MsgBox "Please select RH Value."
           Exit Sub
         End If
        Else
           MsgBox "Please select BloodGroup Value."
           Exit Sub
        End If
    Else
       MsgBox "First choose and enter the search data.", vbExclamation, "BloodBag search"
       Exit Sub
      End If
    srno = 1
    msflgBBSearch.Clear
    msflgBBSearch.Rows = 2
  
    msflgBBSearch.FormatString = "SNo |BloodBagId |Blood Group | TestResults | Status | DateOfCollection | DateOfExpiry  "
    msflgBBSearch.ColWidth(2) = 2000
    
    If rssearch.BOF And rssearch.EOF Then MsgBox "No records found for this criteria."
        
    'Assigns the values from the database to the corresponding controls.
    While Not rssearch.EOF
        msflgBBSearch.TextMatrix(srno, 0) = srno
        msflgBBSearch.TextMatrix(srno, 1) = rssearch("BloodBagID") & ""
        msflgBBSearch.TextMatrix(srno, 2) = rssearch("BloodGroup") & " " & rssearch("RH") & ""
        msflgBBSearch.TextMatrix(srno, 3) = rssearch("TestResults") & ""
        
        msflgBBSearch.TextMatrix(srno, 4) = rssearch("Status") & ""
        msflgBBSearch.TextMatrix(srno, 5) = rssearch("DateOfCollection") & ""
        msflgBBSearch.TextMatrix(srno, 6) = rssearch("DateOfExpiry") & ""
        
        rssearch.MoveNext
        srno = srno + 1
        If msflgBBSearch.Rows = srno Then msflgBBSearch.Rows = msflgBBSearch.Rows + 1
    Wend
    rssearch.Close
    Set rssearch = Nothing
    
    'Clears the Control to set the default values
    optBloodBagID.Enabled = True
    optBloodGroup.Enabled = True
    optBloodBagID.Value = False
    optBloodGroup.Value = False
    txtBloodBagID.Text = ""
    cmbBloodGroup.Text = ""
    cmbRh.Text = ""
    txtBloodBagID.Enabled = False
    cmbBloodGroup.Enabled = False
    cmbRh.Enabled = False
    Exit Sub

'Code for error Handling
errPara:
    If rssearch.State = 1 Then
        'rsSearch.Close
        Set rssearch = Nothing
    End If
    MsgBox "Error in code"
End Sub

'*********************** Search ***********************'

' This procedure is used to search the blood bags that are in database.

Private Sub cmdSearch_Click()
cmdDelete.Enabled = True
cmdSearchExisting.Enabled = True
Dim strsql As String
Dim srno As Integer
Dim rssearch As New ADODB.Recordset
On Error GoTo errPara

'Input Validation for the search criteria
    'Checks for the Blood Bag ID
    If optBloodGroup.Value = False And optBloodBagID.Value = True Then
     If (Len(Trim(txtBloodBagID.Text)) = 0) Then
       MsgBox "First enter the Blood Bag ID", vbInformation, "BloodBagID"
       txtBloodBagID.SetFocus
     Else
       rssearch.Open "SELECT  *  FROM tbl_BloodBagDetails WHERE BloodBagID=" & Val(txtBloodBagID.Text), Con, adOpenKeyset, adLockOptimistic
     End If
    
  'Check for the Blood Group
    ElseIf optBloodBagID.Value = False And optBloodGroup.Value = True Then
       If Len(Trim(cmbBloodGroup.Text)) <> 0 Then
         If Len(Trim(cmbRh.Text)) <> 0 Then
           rssearch.Open "SELECT  *  FROM tbl_BloodBagDetails WHERE BloodGroup='" & cmbBloodGroup.Text & "' And RH='" & cmbRh.Text & "'", Con, adOpenKeyset, adLockOptimistic
         Else
           MsgBox "Please select RH Value."
           Exit Sub
         End If
        Else
           MsgBox "Please select BloodGroup Value."
           Exit Sub
        End If
    Else
       MsgBox "First choose and enter the search data.", vbExclamation, "BloodBag search"
       Exit Sub
      End If
   
  
    srno = 1
    msflgBBSearch.Clear
    msflgBBSearch.Rows = 2
  
    msflgBBSearch.FormatString = "SNo |BloodBagId |Blood Group | TestResults | Status | DateOfCollection | DateOfExpiry  "
    msflgBBSearch.ColWidth(2) = 2000
    
   If rssearch.BOF And rssearch.EOF Then MsgBox "No records found for this criteria."
     
     'Assigns the values from the database to the corresponding controls.
    While Not rssearch.EOF
        msflgBBSearch.TextMatrix(srno, 0) = srno
        msflgBBSearch.TextMatrix(srno, 1) = rssearch("BloodBagID") & ""
        msflgBBSearch.TextMatrix(srno, 2) = rssearch("BloodGroup") & "" & rssearch("Rh") & ""
        msflgBBSearch.TextMatrix(srno, 3) = rssearch("TestResults") & ""
        
        msflgBBSearch.TextMatrix(srno, 4) = rssearch("Status") & ""
        msflgBBSearch.TextMatrix(srno, 5) = rssearch("DateOfCollection") & ""
        msflgBBSearch.TextMatrix(srno, 6) = rssearch("DateOfExpiry") & ""
        
        rssearch.MoveNext
        srno = srno + 1
        If msflgBBSearch.Rows = srno Then msflgBBSearch.Rows = msflgBBSearch.Rows + 1
    Wend
    rssearch.Close
    Set rssearch = Nothing
    
     'Clears the Control to set the default values
        optBloodBagID.Enabled = True
        optBloodGroup.Enabled = True
        optBloodBagID.Value = False
        optBloodGroup.Value = False
        txtBloodBagID.Text = ""
        cmbBloodGroup.Text = ""
        cmbRh.Text = ""
        txtBloodBagID.Enabled = False
        cmbBloodGroup.Enabled = False
        cmbRh.Enabled = False
        
Exit Sub

'Code for error Handling
errPara:
    If rssearch.State = 1 Then
        'rsSearch.Close
    Set rssearch = Nothing
    End If
    MsgBox "Error in code"
End Sub

'Not Included in the project'

Private Sub cmdSearchExp_Click()
cmdDelete.Enabled = True
Dim strsql As String
Dim srno As Integer
Dim rssearch As New ADODB.Recordset


    rssearch.Open "SELECT * FROM tbl_BloodBagDetails WHERE DateOfExpiry <='" & Now() & "' ", Con, adOpenKeyset, adLockOptimistic

    srno = 1
    msflgBBSearch.Clear
    msflgBBSearch.Rows = 2
  
    msflgBBSearch.FormatString = "SNo |BloodBagId |Blood Group | DateOfCollection | DateOfExpiry | TestResults | Status "
    msflgBBSearch.ColWidth(2) = 2000
    
    If rssearch.BOF And rssearch.EOF Then MsgBox "NO BLOOD BAGS EXPIRED TODAY"
        
    While Not rssearch.EOF
        msflgBBSearch.TextMatrix(srno, 0) = srno
        msflgBBSearch.TextMatrix(srno, 1) = rssearch("BloodBagID") & ""
        msflgBBSearch.TextMatrix(srno, 2) = rssearch("BloodGroup") & "" & rssearch("RH") & ""
        msflgBBSearch.TextMatrix(srno, 3) = rssearch("DateOfCollection") & ""
        
        msflgBBSearch.TextMatrix(srno, 4) = rssearch("DateOfExpiry") & ""
        msflgBBSearch.TextMatrix(srno, 5) = rssearch("TestResults") & ""
        msflgBBSearch.TextMatrix(srno, 6) = rssearch("Status") & ""
        
        rssearch.MoveNext
        srno = srno + 1
        If msflgBBSearch.Rows = srno Then msflgBBSearch.Rows = msflgBBSearch.Rows + 1
    Wend
    rssearch.Close
    Set rssearch = Nothing
Exit Sub
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()

'Setting the dimensions of the form.

 frmBBLookUp.Height = 4095
 frmBBLookUp.Width = 9345
 CenterForm frmBBLookUp
End Sub

'*********************** Select the Blood Bag ID ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optBloodBagID_Click()

'Settings for all the control.
optBloodGroup.Enabled = False
cmbRh.Enabled = False
cmbBloodGroup.Enabled = False
txtBloodBagID.Enabled = True
txtBloodBagID.Text = ""
cmbBloodGroup.Text = ""
cmbRh.Text = ""
txtBloodBagID.SetFocus
End Sub

'*********************** Select the Blood Group ***********************'

' This procedure is used to set the values of the Control when the user clicks on this Option Button.

Private Sub optBloodGroup_Click()

'Settings for all the control.
optBloodBagID.Enabled = False
cmbRh.Enabled = True
cmbBloodGroup.Enabled = True
txtBloodBagID.Text = ""
cmbBloodGroup.Text = ""
cmbRh.Text = ""
txtBloodBagID.Enabled = False
cmbBloodGroup.SetFocus
End Sub

