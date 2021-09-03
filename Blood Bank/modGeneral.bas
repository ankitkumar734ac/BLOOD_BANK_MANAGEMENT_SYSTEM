Attribute VB_Name = "modGeneral"
'Form Name: modGeneral.bas

'Displays:   No Displays for Class Module

'Function:  All Public Variables & Methods are Declared in this modules

'Note    :   Provides the Services to Other forms


'************************* Class Modules ***************************************'


Option Explicit

'**** Variable Declarations Publicly ****'
Public aaa As String

Public Con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public Positive As String
Public Negative As String
Public strDonorId As String
Public strUsername As String
Public strHospitalID As String
Public strBBID As String
Public c As String
Public s As String
Public ans As String
Public strDontestres As String
Public flag As Boolean
Public strsql As String
Public i As Integer
Public numeric As Integer



'**** Begins the Connection ****'

Public Sub main()
    openConnection
End Sub

'**** Connection termination ****'

Public Sub CloseConnection()
    Con.Close
    Set Con = Nothing
End Sub

'**** Connection Establishment ****'

Public Sub openConnection()
   Set Con = New ADODB.Connection
   Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Blood.mdb;Persist Security Info=False"
   Con.Open
   If Con.State Then
   frmSplash.Show        'Loads and Shows the Splash Screen
   End If
End Sub

'**** Opens the ADODB Object ****'

Public Sub open_recordset(r As String)
Set rs = New ADODB.Recordset
rs.Open "Select * from " & r, Con, adOpenStatic, adLockOptimistic
End Sub

'**** Closes the ADODB Recordset ****'

Public Sub close_recordset()
Set rs = Nothing
End Sub

'**** Provides the Center Screen Display for all the MDI Child Forms ****'

Public Sub CenterForm(frm As Form)
    frm.Top = (Screen.Height - frm.Height) / 2 - 1200
    frm.Left = (Screen.Width - frm.Height) / 2
End Sub

'**** Input Validation for Integers ****'

Function numericcheck(i As Integer) As Integer
If i >= 48 And i <= 57 Then
numeric = i
Else
numeric = 0
MsgBox "Please enter numbers only"
End If
End Function

'**** Input Validation for Alphabets ****'

Function lettercheck(i As Integer) As Integer
If i >= 65 And i <= 90 Or i >= 97 And i <= 122 Or i = 8 Or i = 32 Then
lettercheck = i
Else
lettercheck = 0
MsgBox "Please enter letters only"
End If
End Function
 
'**** Checks the Expiration of Blood Bags,if any ****'

Public Function CheckExpiry()

Dim strsql As String
Dim strSql1 As String
Dim rssearch As New ADODB.Recordset
Dim str As String
strsql = "Select BloodBagID from tbl_BloodBagDetails where DateOfExpiry <='" & Date & "'  and Status IN ('Existing') "
rssearch.Open strsql, Con, adOpenKeyset, adLockOptimistic
If rssearch.BOF = False And rssearch.EOF = False Then
 rssearch.MoveFirst
 For i = 0 To rssearch.RecordCount - 1
  str = str & "  " & "'" & rssearch("BloodBagID") & "'"
  rssearch.MoveNext
 Next i
 strSql1 = "UPDATE tbl_BloodBagDetails SET Status='Expired' where DateOfExpiry<='" & Date & "' "
 Con.Execute strSql1

 MsgBox "SOME BLOOD BAGS HAVE EXPIRED.", vbInformation, "BLOOD BAG DETAILS"
 MsgBox "Their Blood Bag ID's are: " + vbCrLf + str
Else
  Exit Function
End If
End Function
