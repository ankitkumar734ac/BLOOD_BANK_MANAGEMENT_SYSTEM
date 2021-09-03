VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7755
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   7755
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton cmdCancel 
         Caption         =   "CANCEL"
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "SAVE"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5520
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5520
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtHospital 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone"
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblPhone1 
         Caption         =   "Phone"
         Height          =   255
         Left            =   4560
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblHospitalAddress 
         Caption         =   "Hospital Address"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblHospital 
         Caption         =   "Hospital"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
