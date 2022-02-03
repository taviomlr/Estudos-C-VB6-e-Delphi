VERSION 5.00
Begin VB.Form Simbolos 
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   1800
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   1800
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image4 
      Height          =   480
      Left            =   705
      Picture         =   "Simbolos.frx":0000
      Top             =   675
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   90
      Picture         =   "Simbolos.frx":0442
      Top             =   690
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   705
      Picture         =   "Simbolos.frx":0884
      Top             =   105
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "Simbolos.frx":0CC6
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "Simbolos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
  Debug.Print "Selecionado Paus"
End Sub

Private Sub Image2_Click()
  Debug.Print "Selecionado Ouro"
End Sub

Private Sub Image3_Click()
  Debug.Print "Selecionado Copas"
End Sub

Private Sub Image4_Click()
  Debug.Print "Selecionado Espada"
End Sub
