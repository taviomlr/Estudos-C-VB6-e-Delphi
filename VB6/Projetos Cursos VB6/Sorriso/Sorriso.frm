VERSION 5.00
Begin VB.Form Sorriso 
   BackColor       =   &H8000000E&
   Caption         =   "Sorriso"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   1410
      Picture         =   "Sorriso.frx":0000
      ScaleHeight     =   4815
      ScaleWidth      =   4290
      TabIndex        =   0
      Top             =   195
      Width           =   4290
   End
End
Attribute VB_Name = "Sorriso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
  Print "Tenha um ";
  FontBold = True
  Print "bom";
  FontBold = False
  Print " dia"
End Sub

Private Sub Form_DblClick()
  Cls
End Sub

