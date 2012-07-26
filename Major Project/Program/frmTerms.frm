VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTerms 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Terms of Use"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox rtfTerms 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7858
      _Version        =   393217
      BackColor       =   15594698
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmTerms.frx":0000
   End
End
Attribute VB_Name = "frmTerms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    rtfTerms.LoadFile App.Path & "\terms.rtf"
End Sub
