VERSION 5.00
Begin VB.Form frmUpdateinstall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Install"
   ClientHeight    =   6810
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "frmUpdateinstall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKeyPhrase 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   840
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.TextBox txtKeyFile 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   360
      Width           =   255
      Visible         =   0   'False
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   615
      Left            =   5160
      Picture         =   "frmUpdateinstall.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   615
      Left            =   3840
      Picture         =   "frmUpdateinstall.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "WWW Information"
      Height          =   3015
      Left            =   480
      TabIndex        =   10
      Top             =   3000
      Width           =   5775
      Begin VB.TextBox txtepost 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "your@company.com"
         Top             =   2400
         Width           =   5535
      End
      Begin VB.TextBox txtwww 
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Text            =   "http://www.your-company-here.com"
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox txturltoremote 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   5535
      End
      Begin VB.TextBox txturlfordownload 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label4 
         Caption         =   "Emailaddress:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Homepage:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "URL to RemVersion.dat"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "URL of program for download:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox txtdatum 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtversion 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtprognamn 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtforetag 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Programinformation"
      Height          =   2655
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtexeprogram 
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox txtexeupdate 
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label lblexeprogram 
         Caption         =   "EXE file of program:"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblexeupdate 
         Caption         =   "EXE file of update:"
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lbldatum 
         Caption         =   "Date (YY-MM-DD) :"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblversion 
         Caption         =   "Version:"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblprognamn 
         Caption         =   "Program Name:"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblForetag 
         Caption         =   "Company:"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label lblCopyright 
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
      Left            =   600
      TabIndex        =   9
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Image imgUpdate 
      Height          =   480
      Left            =   0
      Picture         =   "frmUpdateinstall.frx":06D6
      Top             =   6360
      Width           =   480
   End
   Begin VB.Menu Arkiv 
      Caption         =   "&File"
      Begin VB.Menu Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu Skapa 
         Caption         =   "&Create Updatefiles"
      End
      Begin VB.Menu line 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu End 
         Caption         =   "&End"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu Helping 
         Caption         =   "Hel&p"
      End
      Begin VB.Menu om 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmUpdateinstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdclose_Click()
End
End Sub

Private Sub CmdSave_Click()

'Spara all inskriven text i inifilen remoteversion.ini
OutputMode = "export"
x = PutIni(txtKeyFile.Text, "General", "Pass", crypt(txtKeyPhrase.Text, txtKeyPhrase.Text))

x = PutIni(txtKeyFile.Text, "Data", "Foretag", crypt(txtKeyPhrase.Text, txtforetag.Text))
x = PutIni(txtKeyFile.Text, "Data", "ProgNamn", crypt(txtKeyPhrase.Text, txtprognamn.Text))
x = PutIni(txtKeyFile.Text, "Data", "Version", crypt(txtKeyPhrase.Text, txtversion.Text))
x = PutIni(txtKeyFile.Text, "Data", "Datum", crypt(txtKeyPhrase.Text, txtdatum.Text))
x = PutIni(txtKeyFile.Text, "Data", "Exe-of-Update", crypt(txtKeyPhrase.Text, txtexeupdate.Text))
x = PutIni(txtKeyFile.Text, "Data", "Exe-of-Program", crypt(txtKeyPhrase.Text, txtexeprogram.Text))
x = PutIni(txtKeyFile.Text, "Data", "URL-to-download", crypt(txtKeyPhrase.Text, txturlfordownload.Text))
x = PutIni(txtKeyFile.Text, "Data", "URL-to-remversion", crypt(txtKeyPhrase.Text, txturltoremote.Text))
x = PutIni(txtKeyFile.Text, "Data", "WWW", crypt(txtKeyPhrase.Text, txtwww.Text))
x = PutIni(txtKeyFile.Text, "Data", "Email", crypt(txtKeyPhrase.Text, txtepost.Text))

MsgBox "All info sparad !", vbOKOnly, "Update Installer Deluxe"



End Sub

Private Sub End_Click()
End
End Sub

Private Sub Form_Load()
TheKey = "123456789"

 'Hämta in info från inifilen :)
 txtKeyPhrase.Text = TheKey
 txtKeyFile.Text = App.Path & "\remoteversion.ini"

x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Foretag"))

txtforetag.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "ProgNamn"))
txtprognamn.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Version"))
txtversion.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Datum"))
txtdatum.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Exe-of-Update"))
txtexeupdate.Text = x

x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Exe-of-Program"))
txtexeprogram.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "URL-to-download"))
txturlfordownload.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "URL-to-remversion"))
txturltoremote.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "WWW"))
txtwww.Text = x
x = decrypt(txtKeyPhrase.Text, GetIni(txtKeyFile.Text, "data", "Email"))
txtepost.Text = x

 


End Sub

Private Sub Helping_Click()
frmhelp.Show
End Sub

Private Sub om_Click()
frmAbout.Show

End Sub

Private Sub Save_Click()
'Spara all inskriven text i inifilen remoteversion.ini

OutputMode = "export"
x = PutIni(txtKeyFile.Text, "General", "Pass", crypt(txtKeyPhrase.Text, txtKeyPhrase.Text))

x = PutIni(txtKeyFile.Text, "Data", "Foretag", crypt(txtKeyPhrase.Text, txtforetag.Text))
x = PutIni(txtKeyFile.Text, "Data", "ProgNamn", crypt(txtKeyPhrase.Text, txtprognamn.Text))
x = PutIni(txtKeyFile.Text, "Data", "Version", crypt(txtKeyPhrase.Text, txtversion.Text))
x = PutIni(txtKeyFile.Text, "Data", "Datum", crypt(txtKeyPhrase.Text, txtdatum.Text))
x = PutIni(txtKeyFile.Text, "Data", "Exe-of-Update", crypt(txtKeyPhrase.Text, txtexeupdate.Text))
x = PutIni(txtKeyFile.Text, "Data", "Exe-of-Program", crypt(txtKeyPhrase.Text, txtexeprogram.Text))
x = PutIni(txtKeyFile.Text, "Data", "URL-to-download", crypt(txtKeyPhrase.Text, txturlfordownload.Text))
x = PutIni(txtKeyFile.Text, "Data", "URL-to-remversion", crypt(txtKeyPhrase.Text, txturltoremote.Text))
x = PutIni(txtKeyFile.Text, "Data", "WWW", crypt(txtKeyPhrase.Text, txtwww.Text))
x = PutIni(txtKeyFile.Text, "Data", "Email", crypt(txtKeyPhrase.Text, txtepost.Text))

MsgBox "All info sparad !", vbOKOnly, "Update Installer Deluxe"


End Sub

Private Sub Skapa_Click()
frmCreate.Show
End Sub

Private Sub Update_Click()
frmUpdate.Show
End Sub
