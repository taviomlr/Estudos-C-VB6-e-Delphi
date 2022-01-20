VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Modo"
      Height          =   1020
      Left            =   105
      TabIndex        =   7
      Top             =   1635
      Width           =   1365
      Begin VB.OptionButton Option4 
         Caption         =   "Experientes"
         Height          =   270
         Left            =   60
         TabIndex        =   9
         Top             =   660
         Width           =   1245
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Iniciantes"
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Som"
      Height          =   210
      Index           =   0
      Left            =   1635
      TabIndex        =   6
      Top             =   1860
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Pontos de Bônus"
      Height          =   210
      Index           =   1
      Left            =   1635
      TabIndex        =   5
      Top             =   2175
      Width           =   1560
   End
   Begin VB.OptionButton Option2 
      Caption         =   "2 Jogadores"
      Height          =   285
      Left            =   135
      TabIndex        =   4
      Top             =   1215
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1 Jogador"
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   705
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   615
      Left            =   1695
      TabIndex        =   2
      Top             =   2730
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Iniciar"
      Default         =   -1  'True
      Height          =   615
      Left            =   210
      TabIndex        =   1
      Top             =   2745
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Devol. Moeda"
      Height          =   1215
      Left            =   1515
      TabIndex        =   0
      Top             =   300
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
  MsgBox "Fim do Jogo"
  Debug.Print Option1.Value
  Debug.Print Option2.Value
  Debug.Print Option3.Value
  Debug.Print Option4.Value
  Debug.Print Check1.Value
  Debug.Print Check2.Value
End Sub

Private Sub Command3_Click()
  End
End Sub
