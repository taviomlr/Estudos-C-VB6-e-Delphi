VERSION 5.00
Begin VB.Form SeuNomeEh 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Mostrar"
      Height          =   495
      Left            =   1635
      TabIndex        =   1
      Top             =   1305
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AssinalarEMostrar"
      Height          =   495
      Left            =   915
      TabIndex        =   0
      Top             =   555
      Width           =   2895
   End
End
Attribute VB_Name = "SeuNomeEh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeuNome As String
Private Sub Command1_Click()
  'Dim SeuNome As String
  SeuNome = InputBox("Qual o seu nome?")
  MsgBox "O seu nome ? " & SeuNome
End Sub

Private Sub Command2_Click()
  'Dim SeuNome As String
  MsgBox "Al? " & SeuNome
End Sub
