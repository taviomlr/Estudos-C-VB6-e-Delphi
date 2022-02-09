VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   3405
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Iniciar"
      Default         =   -1  'True
      Height          =   495
      Left            =   280
      TabIndex        =   7
      Top             =   3405
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "4 Jogadores"
      Height          =   495
      Left            =   280
      TabIndex        =   6
      Top             =   2685
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "3 Jogadores"
      Height          =   495
      Left            =   280
      TabIndex        =   5
      Top             =   2220
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "2 Jogadores"
      Height          =   495
      Left            =   280
      TabIndex        =   4
      Top             =   1740
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1 Jogador"
      Height          =   495
      Left            =   280
      TabIndex        =   3
      Top             =   1230
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Pontos de Bônus"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   2670
      Width           =   1620
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Som"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Devol. Moeda"
      Height          =   1470
      Left            =   2190
      TabIndex        =   0
      Top             =   240
      Width           =   1650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
  'MsgBox "Fim de Jogo"
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
