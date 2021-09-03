VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   6000
         TabIndex        =   22
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   495
         Left            =   4560
         TabIndex        =   21
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtEnteredBy 
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.ComboBox cmbRh 
         Height          =   315
         Left            =   6840
         TabIndex        =   18
         Text            =   "Combo1"
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox cmbBloodGroup 
         Height          =   315
         Left            =   5280
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtHospitalID 
         Height          =   375
         Left            =   5640
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtDateOfExp 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox txtStatus 
         Height          =   375
         Left            =   2160
         TabIndex        =   10
         Top             =   3435
         Width           =   2055
      End
      Begin VB.TextBox txtTestResults 
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   2430
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dprDateOfCol 
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1650
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   24576001
         CurrentDate     =   36526
      End
      Begin VB.TextBox txtQuantity 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   885
         Width           =   2055
      End
      Begin VB.TextBox txtBloodBagNo 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblEnteredBy 
         AutoSize        =   -1  'True
         Caption         =   "Entered By"
         Height          =   195
         Left            =   4560
         TabIndex        =   19
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label lblRh 
         Caption         =   "RH"
         Height          =   255
         Left            =   6240
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblBloodGroup 
         Caption         =   "Blood Group"
         Height          =   495
         Left            =   4560
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblHospitalID 
         AutoSize        =   -1  'True
         Caption         =   "Hospital ID"
         Height          =   195
         Left            =   4560
         TabIndex        =   13
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lblDateOfExpiry 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Expiry"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   4440
         Width           =   1020
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   3600
         Width           =   450
      End
      Begin VB.Label lblTestResults 
         AutoSize        =   -1  'True
         Caption         =   "Test Results"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label lblDateOfCol 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Collection"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label lblQuantity 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   585
      End
      Begin VB.Label lblBloodBagNo 
         AutoSize        =   -1  'True
         Caption         =   "Blood Bag ID"
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   945
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

