VERSION 5.00
Begin VB.Form VoceMeAma 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   630
      Left            =   1170
      TabIndex        =   0
      Top             =   1050
      Width           =   2235
   End
End
Attribute VB_Name = "VoceMeAma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub Command1_Click()
  Resposta$ = InputBox("Você me ama?")
  If Resposta$ = "Sim" Then
    MsgBox "Ela me ama"
  Else
  If Resposta$ = "Não sei" Or "Não" Then
    MsgBox "Ela não me ama"
  End If
  End If
End Sub
