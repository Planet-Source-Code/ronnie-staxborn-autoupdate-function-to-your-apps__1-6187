VERSION 5.00
Begin VB.Form frmAutoUpdate 
   Caption         =   "Auto Update Deluxe"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6660
   Icon            =   "AutoUpdate.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKeyFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   1440
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.TextBox txtKeyPhrase 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   1440
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   5160
      TabIndex        =   14
      Top             =   0
      Width           =   1215
      Begin VB.Label lblmail 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblwww 
         Caption         =   "Homepage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         MousePointer    =   2  'Cross
         TabIndex        =   15
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1080
      Width           =   1095
      Visible         =   0   'False
   End
   Begin VB.TextBox txtepost 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Text            =   "txtepost"
      Top             =   1440
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.TextBox txtwww 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Text            =   "txtwww"
      Top             =   1080
      Width           =   735
      Visible         =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Programinformation"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   6255
      Begin VB.CommandButton cmdNext1 
         Caption         =   "Next  >"
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblProgram 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label lblVersion 
         Caption         =   "Label6"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Program name:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Current Version:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Date of this program:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lbldate 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   4920
      Picture         =   "AutoUpdate.frx":27A2
      Top             =   3840
      Width           =   1860
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   360
      X2              =   6240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   1080
      X2              =   6360
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "AutoUpdate.frx":298E
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label lblcopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "AutoUpdate Deluxe Copyright © ITson 1999-2000"
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
      Left            =   840
      TabIndex        =   2
      Top             =   4320
      Width           =   3735
   End
   Begin VB.Label lblwelcome 
      Caption         =   "Welcome. This guide will help you Update your software. Confirm the information below is correct."
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label lblforetag 
      Caption         =   "Companyname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   6240
      X2              =   360
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Menu Aboutmenu 
      Caption         =   "&About"
      Begin VB.Menu mnuabout 
         Caption         =   "About &this program"
      End
   End
End
Attribute VB_Name = "frmAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNext1_Click()
Unload Me
frmUpdate.Show
End Sub

Private Sub Form_Load()
Dim www
Dim epost
'Hämta in info från inifilen :)

TheKey = "123456789"

 'Hämta in info från inifilen :)
 txtKeyPhrase.Text = TheKey
 txtKeyFile.Text = App.Path & "\remoteversion.ini"

x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Foretag"))
lblforetag.Caption = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "ProgNamn"))
lblProgram.Caption = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Version"))
lblVersion.Caption = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Datum"))
Text1.Text = x
lbldate.Caption = ConvertDate(Text1)

x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "WWW"))
txtwww.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Email"))
txtepost.Text = x

End Sub

Private Sub lblmail_Click()
epost = Shell("start.exe mailto:" & txtepost, 0)
End Sub

Private Sub lblwww_Click()
www = Shell("start.exe " & txtwww, 0)
End Sub

Private Sub mnuabout_Click()
frmAbout.Show
End Sub

