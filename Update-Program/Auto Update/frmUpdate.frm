VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   Caption         =   "Auto Update Deluxe"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5537.202
   ScaleMode       =   0  'User
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet InetUpdate 
      Left            =   1560
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   4335
      Left            =   120
      Picture         =   "frmUpdate.frx":27A2
      ScaleHeight     =   4275
      ScaleWidth      =   1155
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   1680
      X2              =   5880
      Y1              =   5083.333
      Y2              =   5083.333
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1680
      X2              =   5880
      Y1              =   5083.333
      Y2              =   5083.333
   End
   Begin VB.Label lblexeprogram 
      Caption         =   "lblexeprogram"
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblexeupdate 
      Caption         =   "lblexeupdate"
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lbldownload 
      Caption         =   "lbldownload"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblremote 
      Caption         =   "lblremote"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image imgcircle4 
      Height          =   240
      Left            =   3240
      Picture         =   "frmUpdate.frx":30F1
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Loading CurVersion..."
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblcopyright 
      Alignment       =   1  'Right Justify
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
      Left            =   3000
      TabIndex        =   3
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6120
      Picture         =   "frmUpdate.frx":323B
      Top             =   4080
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   5160
      Picture         =   "frmUpdate.frx":59DD
      Top             =   -360
      Width           =   1860
   End
   Begin VB.Image imgcircle3 
      Height          =   240
      Left            =   3240
      Picture         =   "frmUpdate.frx":5BC9
      Top             =   1800
      Width           =   240
   End
   Begin VB.Image imgcircle2 
      Height          =   240
      Left            =   3240
      Picture         =   "frmUpdate.frx":5D13
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image imgcircle1 
      Height          =   240
      Left            =   3240
      Picture         =   "frmUpdate.frx":5E5D
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Updating program..."
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Checking for Updates..."
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Connecting..."
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intLocalVer As Integer
Dim b() As Byte
Dim intRemoteVer As Integer
Dim strRemoteVer As String
Dim doUpdate As Boolean
Dim www

Private Sub cmdDone_Click()
End
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error GoTo err:
cmdDone.Enabled = False
imgcircle1.Visible = False
imgcircle2.Visible = False
imgcircle3.Visible = False
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = False


x = decrypt(frmAutoUpdate.txtKeyPhrase.Text, GetIni(frmAutoUpdate.txtKeyFile.Text, "data", "URL-to-download"))
lbldownload.Caption = x
x = decrypt(frmAutoUpdate.txtKeyPhrase.Text, GetIni(frmAutoUpdate.txtKeyFile.Text, "data", "URL-to-remversion"))
lblremote.Caption = x

x = decrypt(frmAutoUpdate.txtKeyPhrase.Text, GetIni(frmAutoUpdate.txtKeyFile.Text, "data", "Exe-of-Update"))
lblexeupdate.Caption = x
x = decrypt(frmAutoUpdate.txtKeyPhrase.Text, GetIni(frmAutoUpdate.txtKeyFile.Text, "data", "Exe-of-Program"))
lblexeprogram.Caption = x

Let exeupdate = lblexeupdate.Caption
Let exeprogram = lblexeprogram.Caption



'1. Open the local version file and read in the number
Open App.Path & "\curversion.dat" For Input As #1
intLocalVer = CInt(Input(LOF(1), 1))
Close 1

imgcircle4.Visible = False
Label4.Enabled = False
imgcircle1.Visible = True
Label1.Enabled = True

'2. Download the remote version file and read in the number
' Note: This is all one line:


b() = InetUpdate.OpenURL(lblremote.Caption, 1)
'InetUpdate.Execute lblremote.Caption

strRemoteVer = ""

For T = 0 To UBound(b)
strRemoteVer = strRemoteVer + Chr(b(T))
Next

intRemoteVer = Int(strRemoteVer)

'3. Compare numbers

imgcircle1.Visible = False
Label1.Enabled = False
imgcircle2.Visible = True
Label2.Enabled = True


If intRemoteVer > intLocalVer Then
'Note: This is all one line:
If MsgBox("A more recent version of this program exists. Would you like to update it now?", vbYesNo Or vbQuestion) = vbYes Then
doUpdate = True
Else
doUpdate = False
End If
Else
MsgBox "You already have the most recent version of this program."
doUpdate = False
End If

'4. If doupdate = True, then download the latest program exe from the site

If doUpdate Then
'Note: This is all one line:


b() = InetUpdate.OpenURL(lbldownload.Caption, 1)

imgcircle2.Visible = False
Label2.Enabled = False
imgcircle3.Visible = True
Label3.Enabled = True


Open App.Path & "\" & lblexeupdate.Caption For Binary Access Write As #1
Put #1, , b()
Close 1

Kill App.Path & "\" & lblexeprogram.Caption
Name App.Path & "\" & lblexeupdate.Caption As App.Path & "\" & lblexeprogram.Caption

'Now save the current version into the local version file

Open App.Path & "\curversion.dat" For Output As #1
Print #1, strRemoteVer
Close 1

MsgBox "Update Complete!"
cmdDone.Enabled = True
End If
err:
If err.Number = 13 Then
   MsgBox "Please make sure you are connected to the Internet." & vbCrLf & "If not, please (re)connect to the Internet." & vbCrLf & vbCrLf & "If you suspect problems with this program" & vbCrLf & "please contact the author of the program or" & vbCrLf & "at ronnie@itson.nu" & vbCrLf & vbCrLf & "Auto Update Deluxe will shutdown", vbExclamation
       End
        End If

    End Sub


