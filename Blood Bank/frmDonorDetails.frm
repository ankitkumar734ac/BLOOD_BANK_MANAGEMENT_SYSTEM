VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDonorDetails 
   Caption         =   "DONOR DETAILS ENTRY"
   ClientHeight    =   6360
   ClientLeft      =   -330
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmDonorDetails.frx":0000
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   0
      TabIndex        =   16
      Top             =   -120
      Width           =   11895
      Begin VB.CommandButton cmdupdate 
         Caption         =   "Update "
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
         Left            =   4560
         TabIndex        =   56
         Top             =   5640
         Width           =   1455
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel "
         Enabled         =   0   'False
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
         Left            =   2880
         TabIndex        =   55
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Frame framove 
         Caption         =   "Move"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   10200
         TabIndex        =   50
         Top             =   3600
         Width           =   1575
         Begin VB.CommandButton cmdmove 
            Caption         =   "Last"
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
            Index           =   3
            Left            =   120
            TabIndex        =   54
            Top             =   2040
            Width           =   1335
         End
         Begin VB.CommandButton cmdmove 
            Caption         =   "Next"
            Enabled         =   0   'False
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
            Index           =   2
            Left            =   120
            TabIndex        =   53
            Top             =   1480
            Width           =   1335
         End
         Begin VB.CommandButton cmdmove 
            Caption         =   "Previous"
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   920
            Width           =   1335
         End
         Begin VB.CommandButton cmdmove 
            Caption         =   "First"
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
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdBBIssue 
         Caption         =   "Issue Blood Bag"
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
         Left            =   7560
         TabIndex        =   48
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
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
         Left            =   8040
         TabIndex        =   47
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear All"
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
         Left            =   9720
         TabIndex        =   46
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
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
         Left            =   480
         TabIndex        =   45
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "Close"
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
         Left            =   6240
         TabIndex        =   24
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
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
         Left            =   1680
         TabIndex        =   23
         Top             =   5640
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Screening Test Results"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   39
         Top             =   3600
         Width           =   9975
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
            Height          =   390
            Left            =   7080
            MaxLength       =   20
            TabIndex        =   15
            Top             =   1200
            Width           =   2655
         End
         Begin VB.ComboBox cmbRH 
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
            ItemData        =   "frmDonorDetails.frx":25CA
            Left            =   4440
            List            =   "frmDonorDetails.frx":25D4
            TabIndex        =   14
            Top             =   1200
            Width           =   1095
         End
         Begin VB.ComboBox cmbBloodGroup 
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
            ItemData        =   "frmDonorDetails.frx":25E2
            Left            =   1680
            List            =   "frmDonorDetails.frx":25F2
            TabIndex        =   13
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox chkHiv 
            Caption         =   "HIV"
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
            Left            =   6000
            TabIndex        =   22
            ToolTipText     =   "Human ImmunoDeficiency  Virus"
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkHCV 
            Caption         =   "HCV"
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
            Left            =   4560
            TabIndex        =   21
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkmp 
            Caption         =   "M.P."
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
            Left            =   3240
            TabIndex        =   20
            ToolTipText     =   "Malarial Parasite"
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkHbsag 
            Caption         =   "HBsAG"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   19
            Top             =   780
            Width           =   1215
         End
         Begin VB.CheckBox chkVdrl 
            Caption         =   "VDRL"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label8 
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
            Height          =   375
            Left            =   5760
            TabIndex        =   43
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label7 
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
            Left            =   3960
            TabIndex        =   42
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label Label6 
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
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "Note: Please check the above boxes only if the test results are found to be positive."
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
            TabIndex        =   40
            Top             =   360
            Width           =   8955
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Donor Details"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   11655
         Begin MSComCtl2.DTPicker dtplastdate 
            Height          =   375
            Left            =   6000
            TabIndex        =   12
            Top             =   2760
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
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
            CurrentDate     =   39166
         End
         Begin VB.TextBox txtDonorId 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            TabIndex        =   0
            Top             =   480
            Width           =   2415
         End
         Begin VB.ComboBox cmbDonorType 
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
            ItemData        =   "frmDonorDetails.frx":2603
            Left            =   1560
            List            =   "frmDonorDetails.frx":2613
            TabIndex        =   11
            Top             =   2760
            Width           =   2295
         End
         Begin VB.ComboBox cmbOccupation 
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
            ItemData        =   "frmDonorDetails.frx":2644
            Left            =   9240
            List            =   "frmDonorDetails.frx":2657
            TabIndex        =   10
            Top             =   2040
            Width           =   2175
         End
         Begin VB.ComboBox cmbMarStatus 
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
            ItemData        =   "frmDonorDetails.frx":268B
            Left            =   5520
            List            =   "frmDonorDetails.frx":269B
            TabIndex        =   9
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Frame fraGender 
            Caption         =   "Gender"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   35
            Top             =   1800
            Width           =   2535
            Begin VB.OptionButton optMale 
               Caption         =   "Male"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optFemale 
               Caption         =   "Female"
               BeginProperty Font 
                  Name            =   "Book Antiqua"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   1320
               TabIndex        =   36
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtAge 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   8
            Top             =   2040
            Width           =   495
         End
         Begin VB.TextBox txtMobile 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9240
            MaxLength       =   17
            TabIndex        =   7
            Top             =   1515
            Width           =   2175
         End
         Begin VB.TextBox txtOffPhone 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5040
            MaxLength       =   17
            TabIndex        =   6
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtResPhone 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            MaxLength       =   17
            TabIndex        =   5
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox txtAddress 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5040
            MaxLength       =   150
            TabIndex        =   4
            Top             =   480
            Width           =   6375
         End
         Begin VB.TextBox txtMiddleName 
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
            Left            =   9240
            MaxLength       =   20
            TabIndex        =   3
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5040
            MaxLength       =   20
            TabIndex        =   2
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtSurname 
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   1
            Top             =   960
            Width           =   2415
         End
         Begin VB.Shape Shape1 
            Height          =   735
            Left            =   7800
            Shape           =   4  'Rounded Rectangle
            Top             =   2640
            Width           =   3615
         End
         Begin VB.Label Label4 
            Caption         =   "Last Donated date"
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
            TabIndex        =   49
            Top             =   2760
            Width           =   2415
         End
         Begin VB.Label lblDonorId 
            AutoSize        =   -1  'True
            Caption         =   "Donor ID"
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
            TabIndex        =   44
            Top             =   480
            Width           =   1050
         End
         Begin VB.Label lblDonorType 
            AutoSize        =   -1  'True
            Caption         =   "Donor Type"
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
            TabIndex        =   38
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label lblOccupation 
            AutoSize        =   -1  'True
            Caption         =   "Occupation"
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
            Left            =   7920
            TabIndex        =   37
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Label lblMaritalStatus 
            AutoSize        =   -1  'True
            Caption         =   "Marital Status"
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
            TabIndex        =   34
            Top             =   2040
            Width           =   1530
         End
         Begin VB.Label lblAge 
            AutoSize        =   -1  'True
            Caption         =   "Age"
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
            Left            =   2760
            TabIndex        =   33
            Top             =   2040
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Mobile"
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
            Left            =   7920
            TabIndex        =   32
            Top             =   1560
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Office"
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
            TabIndex        =   31
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Phone (R)"
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
            TabIndex        =   30
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblAddress 
            AutoSize        =   -1  'True
            Caption         =   "Address"
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
            TabIndex        =   29
            Top             =   480
            Width           =   900
         End
         Begin VB.Label lblMiddleNAme 
            Caption         =   "Middle Name"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7920
            TabIndex        =   28
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblName 
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   27
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblSurName 
            AutoSize        =   -1  'True
            Caption         =   "Surname"
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
            TabIndex        =   26
            Top             =   960
            Width           =   990
         End
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   5520
         Width           =   9855
      End
   End
End
Attribute VB_Name = "frmDonorDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Form Name: frmDonorDetails

'Displays:  Form will be displayed when user clicks on the Donor Details Menu of the MDI Form

'Unload  :  When the user finish accessing the donor details ,
            'User can click the Cancel command button and choose the Yes option to terminate from the Form frmDonorDetails

'Functions: This form is used to access the Donor Details.

'Database:  Database Connection is Required for this form. It is provided through the ADODB Connection Object Con

'Note       This form is Child of a MDI form.

Option Explicit
Dim i As Integer

'*********************** Add New Donor Details ***********************'

' This procedure is used to Add New Donor Details.

Private Sub cmdAdd_Click()
Dim rsnew As New ADODB.Recordset
Dim n As String
Dim strsql As String
rsnew.Open "Select * from tbl_donorwaste", Con, adOpenKeyset, adLockOptimistic
If rsnew.RecordCount > 0 Then
 rsnew.MoveFirst
 txtDonorID.Text = rsnew("DONORID")
 strsql = "Delete from tbl_donorwaste where DONORID='" & txtDonorID.Text & "'"
 Con.Execute strsql
Else
open_recordset "tbl_Donor"

'Auto Generation For New Donor Details
n = rs.RecordCount
      If n < 8 Then
        n = rs.RecordCount + 1
        txtDonorID.Text = "d0" + Trim(n)
      ElseIf n = 8 Then
        txtDonorID.Text = "d09"
      Else
        n = rs.RecordCount + 1
        txtDonorID.Text = "d" + Trim(n)
      End If
close_recordset
End If
Set rsnew = Nothing

'Clear all the controls and set default values.

txtName.Text = ""
txtSurname.Text = ""
txtMiddleName.Text = ""
txtAddress.Text = ""
txtResPhone.Text = ""
txtOffPhone.Text = ""
txtMobile.Text = ""
txtAge.Text = ""
cmbMarStatus.Text = ""
cmbOccupation.Text = ""
cmbDonorType.Text = ""
chkVdrl.Value = 0
chkHbsag.Value = 0
chkmp.Value = 0
chkHCV.Value = 0
chkHiv.Value = 0
cmbBloodGroup.Text = ""
cmbRh.Text = ""
txtEnteredBy.Text = ""
dtplastdate.Value = Date
cmdAdd.Enabled = False
cmdClear.Enabled = False
cmdclose.Enabled = False
cmdupdate.Enabled = False
cmdShow.Enabled = False
cmdBBIssue.Enabled = False
framove.Enabled = False
cmdSave.Enabled = True
cmdSave.Default = True
cmdcancel.Enabled = True
cmdcancel.Cancel = True
End Sub

'*********************** Issue Blood Bags to Donor  ***********************'

' This procedure is used to Issue Blood Bags to Donor.

Private Sub cmdBBIssue_Click()
open_recordset "tbl_Donor"
For i = 0 To rs.RecordCount - 1
 If rs("DonorID") = txtDonorID.Text Then
  Exit For
 End If
 rs.MoveNext
Next i
If rs.EOF Then
 MsgBox "This Donor does not exist", vbExclamation, "DONOR DETAILS"
 Exit Sub
End If
close_recordset

'd1 = txtLastDate.Text
'd2 = Date
'If (d2 - d1 < 120) Then
    'MsgBox ("This donor is not eligible to donate the blood")
    'Exit Sub
'End If

If Resultscheck(txtDonorID.Text) = False Then
    MsgBox "This donor is not eligible to donate the blood", vbExclamation, "DONOR DETAILS"
End If

        frmBBFromDonor.Show
        
        'Clear all the controls and sets default values.
        
        txtDonorID.Text = ""
        txtName.Text = ""
        txtSurname.Text = ""
        txtMiddleName.Text = ""
        txtAddress.Text = ""
        txtResPhone.Text = ""
        txtOffPhone.Text = ""
        txtMobile.Text = ""
        txtAge.Text = ""
        cmbMarStatus.Text = ""
        cmbOccupation.Text = ""
        cmbDonorType.Text = ""
        chkVdrl.Value = 0
        chkHbsag.Value = 0
        chkmp.Value = 0
        chkHCV.Value = 0
        chkHiv.Value = 0
        cmbBloodGroup.Text = ""
        cmbRh.Text = ""
        txtEnteredBy.Text = ""
        dtplastdate.Value = Date
 End Sub

'*********************** Cancel ***********************'

' This procedure is used to unload the form.

Private Sub cmdcancel_Click()
Dim strsql As String
strsql = "Insert into tbl_donorwaste values('" & txtDonorID.Text & "')"
Con.Execute strsql

'Clear all the controls and sets default values.

txtDonorID.Text = ""
txtName.Text = ""
txtSurname.Text = ""
txtMiddleName.Text = ""
txtAddress.Text = ""
txtResPhone.Text = ""
txtOffPhone.Text = ""
txtMobile.Text = ""
txtAge.Text = ""
cmbMarStatus.Text = ""
cmbOccupation.Text = ""
cmbDonorType.Text = ""
chkVdrl.Value = 0
chkHbsag.Value = 0
chkmp.Value = 0
chkHCV.Value = 0
chkHiv.Value = 0
cmbBloodGroup.Text = ""
cmbRh.Text = ""
txtEnteredBy.Text = ""
dtplastdate.Value = Date
cmdAdd.Enabled = True
cmdClear.Enabled = True
cmdclose.Enabled = True
cmdShow.Enabled = True
cmdupdate.Enabled = True
cmdBBIssue.Enabled = True
framove.Enabled = True
cmdcancel.Enabled = False
cmdSave.Enabled = False
cmdSave.Default = False
cmdcancel.Cancel = False
End Sub

'*********************** Close ***********************'

' This procedure is used to terminate the form.

Private Sub cmdclose_Click()
Dim response As String
response = MsgBox("Do you want to Quit ?", vbQuestion + vbYesNo, "DONOR DETIALS")
Select Case response                        ' Provides the MsgBox response as the expression to match for the Case Structure.
  Case vbYes:
        Unload Me                           ' Unload the current form
        
  Case vbNo:
  
   'Clear all the controls and sets default values.
   
        txtDonorID.Text = ""
        txtName.Text = ""
        txtSurname.Text = ""
        txtMiddleName.Text = ""
        txtAddress.Text = ""
        txtResPhone.Text = ""
        txtOffPhone.Text = ""
        txtMobile.Text = ""
        txtAge.Text = ""
        cmbMarStatus.Text = ""
        cmbOccupation.Text = ""
        cmbDonorType.Text = ""
        chkVdrl.Value = 0
        chkHbsag.Value = 0
        chkmp.Value = 0
        chkHCV.Value = 0
        chkHiv.Value = 0
        cmbBloodGroup.Text = ""
        cmbRh.Text = ""
        txtEnteredBy.Text = ""
        dtplastdate.Value = Date
        cmdmove(1).Enabled = False
        cmdmove(2).Enabled = False
        cmdSave.Enabled = False
        cmdcancel.Enabled = False
   End Select

End Sub

'*********************** Move ***********************'

' This procedure is used to Move through the records of Donor.

Private Sub cmdmove_Click(Index As Integer)
Dim i As Integer
Dim rsmove As New ADODB.Recordset
open_recordset "tbl_Donor"
rsmove.Open "Select * from tbl_DonorTestResults", Con, adOpenKeyset, adLockOptimistic
Select Case Index
 Case 0:
     
     rs.MoveFirst
     
    'Shows the First record of Donor.
    'Display the record on various controls.
    
        txtDonorID.Text = rs("DonorID")
        txtName.Text = rs("DonorName")
        txtSurname.Text = rs("DonorSName")
        txtMiddleName.Text = rs("DonorMName")
        txtAddress.Text = rs("Address")
        txtResPhone.Text = rs("PhoneRes")
        txtOffPhone.Text = rs("PhoneOff")
        txtMobile.Text = rs("Mobile")

        If rs("Gender") = "Male" Then
         optMale.Value = True
        Else
         optFemale.Value = True
        End If
        txtAge.Text = rs("Age")
        cmbMarStatus.Text = rs("MaritalStatus")
        cmbOccupation.Text = rs("Occupation")
        cmbDonorType.Text = rs("DonorType")
        cmbBloodGroup.Text = rs("BloodGroup")
        cmbRh.Text = rs("RH")
     rsmove.MoveFirst
       If rsmove("TestVDRL") = 0 Then chkVdrl.Value = 0 Else chkVdrl.Value = 1
       If rsmove("TestHBsAG") = 0 Then chkHbsag.Value = 0 Else chkHbsag.Value = 1
       If rsmove("TestMP") = 0 Then chkmp.Value = 0 Else chkmp.Value = 1
       If rsmove("TestHCV") = 0 Then chkHCV.Value = 0 Else chkHCV.Value = 1
       If rsmove("TestHIV") = 0 Then chkHiv.Value = 0 Else chkHiv.Value = 1
       txtEnteredBy.Text = rsmove("EnteredBy")
       dtplastdate.Value = rs("LastDonateDate")
       Set rsmove = Nothing
     close_recordset
     cmdmove(1).Enabled = True
     cmdmove(2).Enabled = True
 
 Case 1:
     
     ' Shows the Previous Record of the Donor  with respect to the current record.
     
     For i = 0 To rs.RecordCount
       If rs("DonorID") = txtDonorID.Text Then
           rs.MovePrevious
          If rs.BOF Then
           rs.MoveLast
           Exit For
           End If
           Exit For
         Else
       rs.MoveNext
       End If
      Next i

     For i = 0 To rsmove.RecordCount
       If rsmove("DonorID") = txtDonorID.Text Then
           rsmove.MovePrevious
          If rsmove.BOF Then
           rsmove.MoveLast
           Exit For
           End If
           Exit For
         Else
       rsmove.MoveNext
       End If
      Next i
                    
       'Display the record on various controls.

        txtDonorID.Text = rs("DonorID")
        txtName.Text = rs("DonorName")
        txtSurname.Text = rs("DonorSName")
        txtMiddleName.Text = rs("DonorMName")
        txtAddress.Text = rs("Address")
        txtResPhone.Text = rs("PhoneRes")
        txtOffPhone.Text = rs("PhoneOff")
        txtMobile.Text = rs("Mobile")

        If rs("Gender") = "Male" Then
         optMale.Value = True
        Else
         optFemale.Value = True
        End If
        txtAge.Text = rs("Age")
        cmbMarStatus.Text = rs("MaritalStatus")
        cmbOccupation.Text = rs("Occupation")
        cmbDonorType.Text = rs("DonorType")
        cmbBloodGroup.Text = rs("BloodGroup")
        cmbRh.Text = rs("RH")
       If rsmove("TestVDRL") = 0 Then chkVdrl.Value = 0 Else chkVdrl.Value = 1
       If rsmove("TestHBsAG") = 0 Then chkHbsag.Value = 0 Else chkHbsag.Value = 1
       If rsmove("TestMP") = 0 Then chkmp.Value = 0 Else chkmp.Value = 1
       If rsmove("TestHCV") = 0 Then chkHCV.Value = 0 Else chkHCV.Value = 1
       If rsmove("TestHIV") = 0 Then chkHiv.Value = 0 Else chkHiv.Value = 1
       txtEnteredBy.Text = rsmove("EnteredBy")
       dtplastdate.Value = rs("LastDonateDate")
       Set rsmove = Nothing
     close_recordset
     
Case 2:

   ' Shows the Next record of the Donor with respect to the current record.
   
   rs.MoveFirst
     For i = 0 To rs.RecordCount
       If rs("DonorID") = txtDonorID.Text Then
         rs.MoveNext
        If rs.EOF Then
         rs.MoveFirst
       Exit For
       End If
       Exit For
     Else
      rs.MovePrevious
      If rs.BOF Then
       rs.MoveLast
      End If
     End If
     Next i
    
     rsmove.MoveFirst
     For i = 0 To rsmove.RecordCount
       If rsmove("DonorID") = txtDonorID.Text Then
         rsmove.MoveNext
        If rsmove.EOF Then
         rsmove.MoveFirst
       Exit For
       End If
       Exit For
     Else
      rsmove.MovePrevious
      If rsmove.BOF Then
       rsmove.MoveLast
      End If
     End If
     Next i
        
      'Display the record on various controls.
 
        txtDonorID.Text = rs("DonorID")
        txtName.Text = rs("DonorName")
        txtSurname.Text = rs("DonorSName")
        txtMiddleName.Text = rs("DonorMName")
        txtAddress.Text = rs("Address")
        txtResPhone.Text = rs("PhoneRes")
        txtOffPhone.Text = rs("PhoneOff")
        txtMobile.Text = rs("Mobile")

        If rs("Gender") = "Male" Then
         optMale.Value = True
        Else
         optFemale.Value = True
        End If
        txtAge.Text = rs("Age")
        cmbMarStatus.Text = rs("MaritalStatus")
        cmbOccupation.Text = rs("Occupation")
        cmbDonorType.Text = rs("DonorType")
        cmbBloodGroup.Text = rs("BloodGroup")
        cmbRh.Text = rs("RH")
       If rsmove("TestVDRL") = 0 Then chkVdrl.Value = 0 Else chkVdrl.Value = 1
       If rsmove("TestHBsAG") = 0 Then chkHbsag.Value = 0 Else chkHbsag.Value = 1
       If rsmove("TestMP") = 0 Then chkmp.Value = 0 Else chkmp.Value = 1
       If rsmove("TestHCV") = 0 Then chkHCV.Value = 0 Else chkHCV.Value = 1
       If rsmove("TestHIV") = 0 Then chkHiv.Value = 0 Else chkHiv.Value = 1
       txtEnteredBy.Text = rsmove("EnteredBy")
       dtplastdate.Value = rs("LastDonateDate")
     Set rsmove = Nothing
     close_recordset
     
 Case 3:
 
     ' Shows the last record of Donor
      rs.MoveLast
      rsmove.MoveLast
        
     'Display the record on various controls.
 
        txtDonorID.Text = rs("DonorID")
        txtName.Text = rs("DonorName")
        txtSurname.Text = rs("DonorSName")
        txtMiddleName.Text = rs("DonorMName")
        txtAddress.Text = rs("Address")
        txtResPhone.Text = rs("PhoneRes")
        txtOffPhone.Text = rs("PhoneOff")
        txtMobile.Text = rs("Mobile")

        If rs("Gender") = "Male" Then
         optMale.Value = True
        Else
         optFemale.Value = True
        End If
        txtAge.Text = rs("Age")
        cmbMarStatus.Text = rs("MaritalStatus")
        cmbOccupation.Text = rs("Occupation")
        cmbDonorType.Text = rs("DonorType")
        cmbBloodGroup.Text = rs("BloodGroup")
        cmbRh.Text = rs("RH")
       If rsmove("TestVDRL") = 0 Then chkVdrl.Value = 0 Else chkVdrl.Value = 1
       If rsmove("TestHBsAG") = 0 Then chkHbsag.Value = 0 Else chkHbsag.Value = 1
       If rsmove("TestMP") = 0 Then chkmp.Value = 0 Else chkmp.Value = 1
       If rsmove("TestHCV") = 0 Then chkHCV.Value = 0 Else chkHCV.Value = 1
       If rsmove("TestHIV") = 0 Then chkHiv.Value = 0 Else chkHiv.Value = 1
       txtEnteredBy.Text = rsmove("EnteredBy")
       dtplastdate.Value = rs("LastDonateDate")
     Set rsmove = Nothing
     close_recordset
      cmdmove(1).Enabled = True
      cmdmove(2).Enabled = True
End Select

End Sub

'*********************** Save ***********************'

' This procedure is used to insert the New Donor Details into the Database.

Private Sub cmdSave_Click()
Dim strsql As String
Dim count As Integer
Dim i As Integer
Dim rsdtr As New ADODB.Recordset
On Error GoTo errPara
open_recordset "tbl_Donor"
rsdtr.Open "Select * from tbl_DonorTestResults", Con, adOpenKeyset, adLockOptimistic
rs.MoveFirst
For i = 0 To rs.RecordCount
If rs("DonorID") = txtDonorID.Text Then
 count = count + 1
End If
 rs.MoveNext
 If rs.EOF Then
  rs.MoveFirst
 End If
Next i
If Validatedata = False Then
    Exit Sub
Else
      If count > 0 Then
        MsgBox "This Donor already exists."
      Else
        Con.BeginTrans
        
        'Insert the New Donor Details into the Database.
        
        strsql = "INSERT INTO tbl_Donor (DonorID, DonorSName,DonorName,DonorMName,Address,PhoneRes,PhoneOff,Mobile," _
        & "Gender , Age, MaritalStatus, BloodGroup, RH, Occupation, DonorType,LastDonateDate) VALUES ('" & txtDonorID.Text & "','" & txtSurname.Text & "','" & txtName.Text & "','" & txtMiddleName.Text & "','" _
        & txtAddress.Text & "','" & txtResPhone.Text & "','" & txtOffPhone.Text & "','" & txtMobile.Text & "','" _
        & IIf(optMale.Value = True, "Male", "Female") & "'," & txtAge.Text & ",'" & cmbMarStatus.Text & "','" & cmbBloodGroup.Text & "','" & cmbRh.Text & "','" _
        & cmbOccupation.Text & "','" & cmbDonorType.Text & "','" & dtplastdate.Value & "')"
        Con.Execute strsql
        
        'Insert the New Donor Test Details into the Database.
        
        strsql = "INSERT INTO tbl_DonorTestResults (DonorID,TestVDRL, TestHBsAG, TestMP, TestHCV, TestHIV,EnteredBy,EnteredDate) VALUES ('" & txtDonorID.Text & "'," _
        & IIf(chkVdrl.Value = 1, 1, 0) & "," & IIf(chkHbsag.Value = 1, 1, 0) & "," _
        & IIf(chkmp.Value = 1, 1, 0) & "," & IIf(chkHCV.Value = 1, 1, 0) & "," _
        & IIf(chkHiv.Value = 1, 1, 0) & ",'" & txtEnteredBy.Text & "','" & Date & "')"
        Con.Execute strsql
        Con.CommitTrans
        MsgBox "Data is Saved.", vbInformation, "Donor Details"
      End If
End If
    
  'Clear all the controls and sets default values.
  
    strDonorId = txtDonorID.Text
    txtDonorID.Text = ""
    txtName.Text = ""
    txtSurname.Text = ""
    txtMiddleName.Text = ""
    txtAddress.Text = ""
    txtResPhone.Text = ""
    txtOffPhone.Text = ""
    txtMobile.Text = ""
    txtAge.Text = ""
    cmbMarStatus.Text = ""
    cmbOccupation.Text = ""
    cmbDonorType.Text = ""
    chkVdrl.Value = 0
    chkHbsag.Value = 0
    chkmp.Value = 0
    chkHCV.Value = 0
    chkHiv.Value = 0
    cmbBloodGroup.Text = ""
    cmbRh.Text = ""
    txtEnteredBy.Text = ""
    dtplastdate.Value = Date
    cmdAdd.Enabled = True
    cmdClear.Enabled = True
    cmdclose.Enabled = True
    cmdShow.Enabled = True
    cmdupdate.Enabled = True
    cmdBBIssue.Enabled = True
    framove.Enabled = True
    cmdSave.Enabled = False
    cmdcancel.Enabled = False
    cmdSave.Default = False
    cmdcancel.Cancel = False

Exit Sub

'Code for Error Handling.
errPara:
    MsgBox Err.Description
    Con.RollbackTrans
    Load MDIForm1
    MsgBox "The entry is not allowed"
    Exit Sub
    
End Sub

'*********************** Validation Of Donor Details ***********************'

' This procedure is used to validate the Donor Details.

Public Function Validatedata() As Boolean
Validatedata = True

'Validate the Donor ID.
If Len(Trim(txtDonorID.Text)) = 0 Then
    MsgBox "Please enter Donor ID"
    Validatedata = False
    txtDonorID.SetFocus
    Exit Function
End If

'Validate the Donor Name.
If Len(Trim(txtName.Text)) = 0 Then
    MsgBox "Please enter name"
    Validatedata = False
    txtName.SetFocus
    Exit Function
End If

'Validate the Donor Address.
If Len(Trim(txtAddress.Text)) = 0 Then
    MsgBox "Please enter Address"
    Validatedata = False
    txtAddress.SetFocus
    Exit Function
End If

'Validate the Donor Age.
If Len(txtAge.Text) = 0 Then
    MsgBox "Please enter age"
    Validatedata = False
    txtAge.SetFocus
    Exit Function
    ElseIf (txtAge.Text < 18 Or txtAge.Text > 60) Then
    MsgBox "Age should be between 18 to 60"
    Validatedata = False
    txtAge.Text = ""
    txtAge.SetFocus
    Exit Function
End If

'Validate the Donor Marital Status.
If Len(cmbMarStatus.Text) = 0 Then
    MsgBox "Please enter Marital status"
    Validatedata = False
    cmbMarStatus.SetFocus
    Exit Function
End If

'Validate the Donor Blood Group Details.
If Len(cmbBloodGroup.Text) = 0 Then
    MsgBox "Please select Blood Group"
    Validatedata = False
    cmbBloodGroup.SetFocus
    Exit Function
End If

'Validate the Donor Rh Factor Details.
If Len(cmbRh.Text) = 0 Then
    MsgBox "Please select RH"
    Validatedata = False
    cmbRh.SetFocus
    Exit Function
End If

'Validate the DEO's Details.
If Len(txtEnteredBy.Text) = 0 Then
    MsgBox "Please enter DEO's Name "
    Validatedata = False
    txtEnteredBy.SetFocus
    Exit Function
End If

'Validate the Donor Occupation Details.
If Len(cmbOccupation.Text) = 0 Then
    MsgBox "Please select Occupation"
    Validatedata = False
    cmbOccupation.SetFocus
    Exit Function
End If

'Validate the Donor Type.
If Len(cmbDonorType.Text) = 0 Then
    MsgBox "Please select Donor Type"
    Validatedata = False
    cmbDonorType.SetFocus
    Exit Function
End If
End Function

'*********************** Clear The Donor Details ***********************'

' This procedure is used to clear the Donor Details.

Private Sub cmdclear_Click()

'Clear all the controls and sets default values.

txtDonorID.Text = ""
txtName.Text = ""
txtSurname.Text = ""
txtMiddleName.Text = ""
txtAddress.Text = ""
txtResPhone.Text = ""
txtOffPhone.Text = ""
txtMobile.Text = ""
txtAge.Text = ""
cmbMarStatus.Text = ""
cmbOccupation.Text = ""
cmbDonorType.Text = ""
chkVdrl.Value = 0
chkHbsag.Value = 0
chkmp.Value = 0
chkHCV.Value = 0
chkHiv.Value = 0
cmbBloodGroup.Text = ""
cmbRh.Text = ""
txtEnteredBy.Text = ""
dtplastdate.Value = Date
 cmdmove(1).Enabled = False
 cmdmove(2).Enabled = False
End Sub

'*********************** Show the Donor Details ***********************'

' This procedure is used to show the Donor Details.

Private Sub cmdShow_Click()

'Clear all the controls and sets default values.

txtName.Text = ""
txtSurname.Text = ""
txtMiddleName.Text = ""
txtAddress.Text = ""
txtResPhone.Text = ""
txtOffPhone.Text = ""
txtMobile.Text = ""
txtAge.Text = ""
cmbMarStatus.Text = ""
cmbOccupation.Text = ""
cmbDonorType.Text = ""
chkVdrl.Value = 0
chkHbsag.Value = 0
chkmp.Value = 0
chkHCV.Value = 0
chkHiv.Value = 0
cmbBloodGroup.Text = ""
cmbRh.Text = ""
txtEnteredBy.Text = ""
dtplastdate.Value = Date

open_recordset "tbl_Donor"
For i = 0 To rs.RecordCount - 1
If rs("DonorID") = txtDonorID.Text Then

'Assign Values to all the controls.

txtName.Text = rs("DonorName")
txtSurname.Text = rs("DonorSName")
txtMiddleName.Text = rs("DonorMName")
txtAddress.Text = rs("Address")
txtResPhone.Text = rs("PhoneRes")
txtOffPhone.Text = rs("PhoneOff")
txtMobile.Text = rs("Mobile")

If rs("Gender") = "Male" Then
        optMale.Value = True
    Else
        optFemale.Value = True
End If

'Assign Values to all the controls.

txtAge.Text = rs("Age")
cmbMarStatus.Text = rs("MaritalStatus")
cmbOccupation.Text = rs("Occupation")
cmbDonorType.Text = rs("DonorType")
cmbBloodGroup.Text = rs("BloodGroup")
cmbRh.Text = rs("RH")
dtplastdate.Value = rs("LastDonateDate")

Exit For
End If
rs.MoveNext
Next i

If rs.EOF Then
MsgBox ("No such entry in Donor_Details")
End If

close_recordset
open_recordset "tbl_DonorTestResults"

For i = 0 To rs.RecordCount - 1
If rs("DonorID") = txtDonorID.Text Then

  'Validate the Test Result Details.
  
  If rs("TestVDRL") = 0 Then chkVdrl.Value = 0 Else chkVdrl.Value = 1
  If rs("TestHBsAG") = 0 Then chkHbsag.Value = 0 Else chkHbsag.Value = 1
  If rs("TestMP") = 0 Then chkmp.Value = 0 Else chkmp.Value = 1
  If rs("TestHCV") = 0 Then chkHCV.Value = 0 Else chkHCV.Value = 1
  If rs("TestHIV") = 0 Then chkHiv.Value = 0 Else chkHiv.Value = 1
  txtEnteredBy.Text = rs("EnteredBy")
  Exit For
End If
rs.MoveNext
Next i

close_recordset
End Sub

'*********************** Update the Donor Details ***********************'

' This procedure is used to update the Donor Details.

Private Sub cmdupdate_Click()
Dim response As String
Dim rsdtr As New ADODB.Recordset
Dim rsnew As New ADODB.Recordset
Dim strsql As String
rsdtr.Open "Select * from tbl_DonorTestResults", Con, adOpenKeyset, adLockOptimistic
rsnew.Open "Select * from tbl_Donor where DonorID='" & txtDonorID.Text & "'", Con, adOpenKeyset, adLockOptimistic
If rsnew.BOF And rsnew.EOF Then
response = MsgBox("This donor does not exist.Do you want to add it in DataBase", vbQuestion + vbYesNo, "Donor Details")
 Select Case response
  Case vbYes:
      Con.BeginTrans
      
       'Insert the Donor Details.
        strsql = "INSERT INTO tbl_Donor (DonorID, DonorSName,DonorName,DonorMName,Address,PhoneRes,PhoneOff,Mobile," _
        & "Gender , Age, MaritalStatus, BloodGroup, RH, Occupation, DonorType,LastDonateDate) VALUES ('" & txtDonorID.Text & "','" & txtSurname.Text & "','" & txtName.Text & "','" & txtMiddleName.Text & "','" _
        & txtAddress.Text & "','" & txtResPhone.Text & "','" & txtOffPhone.Text & "','" & txtMobile.Text & "','" _
        & IIf(optMale.Value = True, "Male", "Female") & "'," & txtAge.Text & ",'" & cmbMarStatus.Text & "','" & cmbBloodGroup.Text & "','" & cmbRh.Text & "','" _
        & cmbOccupation.Text & "','" & cmbDonorType.Text & "','" & dtplastdate.Value & "')"
        Con.Execute strsql
       
       'Insert the Donor Details.
        strsql = "INSERT INTO tbl_DonorTestResults (DonorID,TestVDRL, TestHBsAG, TestMP, TestHCV, TestHIV,EnteredBy,EnteredDate) VALUES ('" & txtDonorID.Text & "'," _
        & IIf(chkVdrl.Value = 1, 1, 0) & "," & IIf(chkHbsag.Value = 1, 1, 0) & "," _
        & IIf(chkmp.Value = 1, 1, 0) & "," & IIf(chkHCV.Value = 1, 1, 0) & "," _
        & IIf(chkHiv.Value = 1, 1, 0) & ",'" & txtEnteredBy.Text & "','" & Date & "')"
        Con.Execute strsql
      Con.CommitTrans
      MsgBox "Data Inserted"
 End Select
Else
    Con.BeginTrans
    
   'Update the Donor Details.
    strsql = "UPDATE tbl_Donor SET DonorSName='" & txtSurname.Text & "',DonorName='" & txtName.Text & "',DonorMName='" & txtMiddleName.Text & "',Address='" & txtAddress.Text & "',PhoneRes='" & txtResPhone.Text _
    & "',PhoneOff='" & txtOffPhone.Text & "',Mobile='" & txtMobile & "',Age=" & txtAge.Text & ",Gender='" & IIf(optMale.Value = True, "Male", "Female") & "',MaritalStatus='" & cmbMarStatus.Text & "',Occupation='" & cmbOccupation.Text _
    & "',DonorType='" & cmbDonorType.Text & "',LastDonateDate='" & dtplastdate.Value & "' WHERE DonorID='" & txtDonorID.Text & "'"
     Con.Execute strsql
     rsdtr.MoveFirst
     'strsql = "UPDATE tbl_DonorTestResults SET TestVDRL=" & IIf(chkVdrl.Value = 1, 1, 0) & ",TestHB&AG=" & IIf(chkHbsag.Value = 1, 1, 0) & ",TestMP=" & IIf(chkmp.Value = 1, 1, 0) & ",TestHCV=" & IIf(chkHCV.Value = 1, 1, 0) & ",TestHIV=" & IIf(chkHiv.Value = 1, 1, 0) & ",EnteredBy='" & txtEnteredBy.Text & "',EnteredDate=" & Date & " WHERE DonorID='" & txtDonorId.Text & "'"
     For i = 0 To rsdtr.RecordCount - 1
     If rsdtr("DonorID") = txtDonorID.Text Then
    
    'Update the Test Result Details.
     rsdtr.Update "TestVDRL", IIf(chkVdrl.Value = 1, 1, 0)
     rsdtr.Update "TestHBsAG", IIf(chkHbsag.Value = 1, 1, 0)
     rsdtr.Update "TestMP", IIf(chkmp.Value = 1, 1, 0)
     rsdtr.Update "TestHCV", IIf(chkHCV.Value = 1, 1, 0)
     rsdtr.Update "TestHIV", IIf(chkHiv.Value = 1, 1, 0)
     rsdtr.Update "EnteredBy", txtEnteredBy.Text
     rsdtr.Update "EnteredDate", Date
     Exit For
     Else
     rsdtr.MoveNext
     If rsdtr.EOF Then
          rsdtr.MoveFirst
     End If
     End If
     Next i
     Con.CommitTrans
     MsgBox "Data Updated."
 End If
 
 'Clear all the controls and sets default values.

  txtDonorID.Text = ""
  txtName.Text = ""
  txtSurname.Text = ""
  txtMiddleName.Text = ""
  txtAddress.Text = ""
  txtResPhone.Text = ""
  txtOffPhone.Text = ""
  txtMobile.Text = ""
  txtAge.Text = ""
  cmbMarStatus.Text = ""
  cmbOccupation.Text = ""
  cmbDonorType.Text = ""
  chkVdrl.Value = 0
  chkHbsag.Value = 0
  chkmp.Value = 0
  chkHCV.Value = 0
  chkHiv.Value = 0
  cmbBloodGroup.Text = ""
  cmbRh.Text = ""
  txtEnteredBy.Text = ""
  dtplastdate.Value = Date
  cmdmove(1).Enabled = False
 cmdmove(2).Enabled = False
 Set rsdtr = Nothing
 Set rsnew = Nothing
End Sub

'*********************** Loads the Form ***********************'

' This procedure is used to Load and show the Form.

Private Sub Form_Load()
    
 'Sets the dimension for the form.
    frmDonorDetails.Height = 6870
    frmDonorDetails.Width = 11880
    CenterForm frmDonorDetails
 
 'Sets the Default value.
    dtplastdate.Value = Date
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate details of Donor.

Private Sub txtAge_keypress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If (keyascii < 48 Or keyascii > 57) Then
    MsgBox "Please enter only digits."
    keyascii = 0
End If
End Sub

'*********************** Fill Details Of Donor ***********************'

' This procedure is used to Fill Details Of Donor .

Private Sub fillDetails(strDonorId As String)
Dim strsql As String
Dim rsDetail As New ADODB.Recordset

If strDonorId = "" Then Exit Sub

    strsql = "SELECT * FROM tbl_Donor WHERE DonorID='" & strDonorId & "'"
    rsDetail.Open strsql, Con, adOpenKeyset, adLockOptimistic
     
   'Assign Values to all the controls.

    txtDonorID.Text = rsDetail("DonorID")
    txtName.Text = rsDetail("DonorName")
    txtSurname.Text = rsDetail("DonorSName")
    txtMiddleName.Text = rsDetail("DonorMName") & ""
    txtAddress.Text = rsDetail("Address") & ""
    txtResPhone.Text = rsDetail("PhoneRes") & ""
    txtOffPhone.Text = rsDetail("PhoneOff") & ""
    If rsDetail("Gender") = "Male" Then
        optMale.Value = True
    Else
        optFemale.Value = True
    End If
    
    'Assign Values to the controls.

    txtAge.Text = rsDetail("Age")
    cmbMarStatus.Text = rsDetail("MaritalStatus")
    cmbBloodGroup.Text = rsDetail("bloodGroup")
    cmbRh.Text = rsDetail("RH")
    cmbOccupation.Text = rsDetail("Occupation")
    cmbDonorType.Text = rsDetail("DonorType")
    
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate Mobile No.Details Of Donor .

Private Sub txtMobile_KeyPress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If (keyascii < 48 Or keyascii > 57) Then
    MsgBox "Please enter only digits."
    keyascii = 0
End If
txtMobile.SetFocus
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate Office Phone No. Details Of Donor .

Private Sub txtOffPhone_KeyPress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If (keyascii < 48 Or keyascii > 57) Then
    MsgBox "Please enter only digits."
    keyascii = 0
End If
txtOffPhone.SetFocus
End Sub

'*********************** Input Validation ***********************'

' This procedure is used to validate Residential Phone No. Details Of Donor .

Private Sub txtResPhone_KeyPress(keyascii As Integer)
If keyascii = 8 Then Exit Sub
If (keyascii < 48 Or keyascii > 57) Then
    MsgBox "Please enter only digits."
    keyascii = 0
End If
txtResPhone.SetFocus
End Sub

'*********************** Test Result Check ***********************'

' This procedure is used to check the test result.

Public Function Resultscheck(strDonorId As String) As Boolean
Dim strsql As String
Dim rsres As New ADODB.Recordset
Dim count As Integer
count = 0
If strDonorId = "" Then Exit Function

strsql = "select * from tbl_DonorTestResults where DonorID='" & strDonorId & "'"
rsres.Open strsql, Con, adOpenKeyset, adLockOptimistic

If rsres("TestVDRL") = True Then       ' VDRL --> Vinryl Disease Reasearch Laboratory.
count = count + 1
End If
If rsres("TestHBsAG") = True Then      ' HBsAG --> Hepitatis B surface AntiGen.
count = count + 1
End If

If rsres("TestMP") = True Then         ' MP --> Malarial Parasite.
count = count + 1
End If

If rsres("TestHCV") = True Then        ' HCV --> Hepatatis C Virus.
count = count + 1
End If
 
If rsres("TestHIV") = True Then        ' HIV --> Human Immune Virus.
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
