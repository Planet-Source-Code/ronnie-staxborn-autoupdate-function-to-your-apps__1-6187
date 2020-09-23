VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmhelp 
   Caption         =   "Help"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   Icon            =   "frmhelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8281
      _Version        =   327681
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmhelp.frx":27A2
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lFileLength As Long
    Dim iFileNum As Integer
    
Private Sub Form_Load()
'Get a free file number and open the file
    iFileNum = FreeFile
    Open "help.txt" For Input As iFileNum

    'Get the length of the file and
    'read it into the text box
    lFileLength = LOF(iFileNum)
    RichTextBox1.Text = Input(lFileLength, #iFileNum)

    Close iFileNum
End Sub

