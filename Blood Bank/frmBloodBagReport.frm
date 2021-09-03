VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBloodBagReport 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36526
      End
      Begin VB.Label lblDate 
         Caption         =   "date"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmBloodBagReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim dt1 As String
dt1 = DTPicker1.Value
MsgBox dt1
DataEnvironment1.Command2_grouping dt1
DataReport2.Show
End Sub
