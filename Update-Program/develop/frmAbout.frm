VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About AutoUpdate Deluxe"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "CLOSE"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label lblhomepage 
         Caption         =   "' Visit our homepage  http://www.itson.nu '"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label lblcopyright 
         Caption         =   "Copyright 2000 ITson [Ronnie Staxborn]"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version 1.0"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblProgram 
         Caption         =   "Auto Update Deluxe"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


