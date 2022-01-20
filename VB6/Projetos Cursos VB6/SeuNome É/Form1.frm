VERSION 5.00
Begin VB.Form Form1 
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
      Left            =   1710
      TabIndex        =   1
      Top             =   1485
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AssinalarEMostrar"
      Height          =   495
      Left            =   1470
      TabIndex        =   0
      Top             =   855
      Width           =   1830
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SeuNome As String

Private Sub Command1_Click()
  'Dim SeuNome As String
  SeuNome = InputBox("Qual é o seu nome?")
  MsgBox "O seu nome é " & SeuNome
End Sub

Private Sub Command2_Click()
  'Dim SeuNome As String
  MsgBox "Alô " & SeuNome
End Sub
