VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   1785
   ClientTop       =   255
   ClientWidth     =   11355
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   5400
      TabIndex        =   50
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3840
      TabIndex        =   49
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   48
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1245
      Index           =   3
      Left            =   2520
      TabIndex        =   47
      Top             =   1320
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   46
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   45
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "SCREENING TESTS RESULTS"
      Height          =   2895
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   5895
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   4320
         TabIndex        =   58
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   4320
         TabIndex        =   57
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   4320
         TabIndex        =   56
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   3120
         TabIndex        =   55
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   3120
         TabIndex        =   54
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   3120
         TabIndex        =   53
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   3120
         TabIndex        =   52
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   3120
         TabIndex        =   51
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   1920
         TabIndex        =   39
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   1920
         TabIndex        =   37
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   1920
         TabIndex        =   35
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   1920
         TabIndex        =   33
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H008080FF&
         Height          =   285
         Left            =   1920
         TabIndex        =   31
         Top             =   600
         Width           =   615
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   4080
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   4080
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   4080
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   4080
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIME"
         Height          =   255
         Left            =   4320
         TabIndex        =   42
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label27 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DATE"
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label26 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ENTERED BY"
         Height          =   255
         Left            =   4320
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line3 
         X1              =   4080
         X2              =   4080
         Y1              =   240
         Y2              =   2880
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HIV"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label24 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HCV"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label23 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MP"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HBsAG"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VDRL"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   3000
         X2              =   3000
         Y1              =   240
         Y2              =   2880
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   1680
         Y1              =   240
         Y2              =   2880
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NEG"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3120
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "POS"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TESTS"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   7320
      TabIndex        =   25
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   5400
      TabIndex        =   23
      Top             =   3840
      Width           =   495
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      ItemData        =   "DONOR INFO.frx":0000
      Left            =   2280
      List            =   "DONOR INFO.frx":0010
      TabIndex        =   21
      Top             =   4560
      Width           =   1455
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "DONOR INFO.frx":0041
      Left            =   2280
      List            =   "DONOR INFO.frx":0051
      TabIndex        =   19
      Top             =   4200
      Width           =   1455
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "DONOR INFO.frx":007C
      Left            =   2280
      List            =   "DONOR INFO.frx":0086
      TabIndex        =   17
      Top             =   3840
      Width           =   1455
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   6120
      TabIndex        =   15
      Text            =   "19"
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "DONOR INFO.frx":009E
      Left            =   5280
      List            =   "DONOR INFO.frx":00C6
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "DONOR INFO.frx":0106
      Left            =   4560
      List            =   "DONOR INFO.frx":0167
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3360
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      DataMember      =   "M"
      Height          =   315
      ItemData        =   "DONOR INFO.frx":01DF
      Left            =   2280
      List            =   "DONOR INFO.frx":01E9
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   7920
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label30 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CLEAR"
      Height          =   375
      Left            =   6480
      TabIndex        =   44
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SAVE"
      Height          =   375
      Left            =   6480
      TabIndex        =   43
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "RH"
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "GROUP"
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "DONOR TYPE"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "OCCUPATION"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "MARITAL STATUS"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SEX"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MOB"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "OFF"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "RES"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MIDDLE NAME"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SURNAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "DONOR'S NAME"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label30_Click()
Text1.Count = Null

End Sub
