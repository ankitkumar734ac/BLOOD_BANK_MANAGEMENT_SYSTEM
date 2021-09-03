VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Donor Record"
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtDonorID 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblDonorID 
         AutoSize        =   -1  'True
         Caption         =   "Donor ID"
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
msgbo
Dim strSql As String
strSql = ""

End Sub
