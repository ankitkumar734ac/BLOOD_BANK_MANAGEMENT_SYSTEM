VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go >"
         Height          =   375
         Left            =   8040
         TabIndex        =   2
         Top             =   7680
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Height          =   8055
         Left            =   0
         Picture         =   "band.frx":0000
         ScaleHeight     =   7995
         ScaleWidth      =   7875
         TabIndex        =   1
         Top             =   120
         Width           =   7935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

