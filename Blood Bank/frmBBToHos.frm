VERSION 5.00
Begin VB.Form frmBBToHos 
   Caption         =   "ISSUE TO HOSPITAL"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4200
         TabIndex        =   12
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton cmdIssue 
         Caption         =   "Issue"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtHospitalName 
         Height          =   285
         Left            =   4560
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtHospitalID 
         Height          =   285
         Left            =   4560
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtIssuedBy 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtProduct 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtBloodBagID 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblHospitalName 
         AutoSize        =   -1  'True
         Caption         =   "Hospital Name"
         Height          =   195
         Left            =   3360
         TabIndex        =   9
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblHospitalID 
         AutoSize        =   -1  'True
         Caption         =   "HospitalID"
         Height          =   195
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblIssuedBy 
         AutoSize        =   -1  'True
         Caption         =   "IssuedBy"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblProduct 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblBloodBagID 
         AutoSize        =   -1  'True
         Caption         =   "Blood Bag ID"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmBBToHos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

