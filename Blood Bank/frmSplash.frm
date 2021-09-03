VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4095
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.Timer proccesTimer 
         Interval        =   100
         Left            =   7440
         Top             =   3600
      End
      Begin VB.Timer loadTimmer 
         Interval        =   100
         Left            =   7920
         Top             =   3600
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   3600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) : Product is copyrighted in the year 2019"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Warning : Use of pirated copy is illegal."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   4
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0.A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6495
         TabIndex        =   3
         Top             =   1080
         Width           =   1590
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         Height          =   1575
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "               Welcome to Blood Bank"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form name : frmSplash
'This is loading form


Dim i As Integer
Option Explicit


Private Sub Form_KeyPress(keyascii As Integer)
Unload Me

End Sub

Private Sub Form_Load()
  lblVersion.Caption = "Windows Version " & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub


Private Sub proccesTimer_Timer()
i = Rnd() * 20
    If ProgressBar1.Value < 100 Then
        If ProgressBar1.Value + i < 100 Then
            ProgressBar1.Value = ProgressBar1.Value + i
        Else
            ProgressBar1.Value = 100
        End If
    Else
    Unload Me
   MDIForm1.Show
    'MDI_BBMS.Show
    End If

End Sub

