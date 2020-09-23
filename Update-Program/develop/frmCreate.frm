VERSION 5.00
Begin VB.Form frmCreate 
   Caption         =   "Skapa Updatefiler"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "frmCreate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose2 
      Caption         =   "Close"
      Height          =   615
      Left            =   5280
      Picture         =   "frmCreate.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Create RemVersion.dat"
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
      Begin VB.CommandButton cmdCreateRem 
         Caption         =   "Create"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtremversion 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Version:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create CurVersion.dat"
      Height          =   1215
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.CommandButton cmdCreatecur 
         Caption         =   "Create"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCurversion 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Version:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label lblCopyright2 
      Caption         =   "Update Installer Copyright © ITson 1999-2000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmCreate.frx":0454
      Top             =   3240
      Width           =   480
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose2_Click()
Unload Me
End Sub

Private Sub cmdCreatecur_Click()
Dim iFileNum As Integer

'Get a free file handle
iFileNum = FreeFile

'If the file is not there, one will be created
'If the file does exist, this one will
'overwrite it.
If txtCurversion.Text = "" Then MsgBox "Denna inmatning kan EJ vara tom"

Open App.Path & "\CurVersion.dat" For Output As iFileNum

Print #iFileNum, txtCurversion.Text

Close iFileNum

End Sub

Private Sub cmdCreateRem_Click()
Dim Remdat As Integer

'Get a free file handle
Remdat = FreeFile(1)

'If the file is not there, one will be created
'If the file does exist, this one will
'overwrite it.
Open App.Path & "\RemVersion.dat" For Output As Remdat

Print #Remdat, txtremversion.Text

Close #Remdat

End Sub
