VERSION 5.00
Begin VB.Form cacaniquel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caça Níquel"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   675
      Left            =   795
      TabIndex        =   4
      Top             =   3540
      Width           =   2220
   End
   Begin VB.CommandButton btnSair 
      Caption         =   "Sair"
      Height          =   675
      Left            =   4980
      TabIndex        =   2
      Top             =   3570
      Width           =   2220
   End
   Begin VB.CommandButton btnJogada 
      Caption         =   "Jogada"
      Height          =   750
      Left            =   1725
      TabIndex        =   0
      Top             =   1530
      Width           =   5310
   End
   Begin VB.Label Label1 
      Caption         =   "By Távio"
      Height          =   240
      Left            =   7575
      TabIndex        =   5
      Top             =   4380
      Width           =   780
   End
   Begin VB.Image ImgVisor4 
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   6900
      Stretch         =   -1  'True
      Top             =   480
      Width           =   780
   End
   Begin VB.Label lblTotJogadas 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   5040
      TabIndex        =   3
      Top             =   2625
      Width           =   1935
   End
   Begin VB.Image ImgPaus 
      Height          =   480
      Left            =   2010
      Picture         =   "cacaniquel1.frx":0000
      Top             =   4815
      Width           =   480
   End
   Begin VB.Image ImgOuros 
      Height          =   480
      Left            =   3780
      Picture         =   "cacaniquel1.frx":0442
      Top             =   4770
      Width           =   480
   End
   Begin VB.Image ImgEspadas 
      Height          =   480
      Left            =   5220
      Picture         =   "cacaniquel1.frx":0884
      Top             =   4815
      Width           =   480
   End
   Begin VB.Image ImgCopas 
      Height          =   480
      Left            =   585
      Picture         =   "cacaniquel1.frx":0CC6
      Top             =   4755
      Width           =   480
   End
   Begin VB.Label lblTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   1710
      TabIndex        =   1
      Top             =   2595
      Width           =   1935
   End
   Begin VB.Image ImgVisor3 
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   5055
      Stretch         =   -1  'True
      Top             =   495
      Width           =   780
   End
   Begin VB.Image ImgVisor2 
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   2970
      Stretch         =   -1  'True
      Top             =   480
      Width           =   765
   End
   Begin VB.Image ImgVisor1 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   930
      Stretch         =   -1  'True
      Top             =   480
      Width           =   795
   End
End
Attribute VB_Name = "cacaniquel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const COPAS% = 1
Const PAUS% = 2
Const OUROS% = 3
Const ESPADAS% = 4
Dim Vitórias As Currency
Dim TotJogadas As Integer

Private Sub btnReset_Click()
  MsgBox "Seu Saldo Atual é de " & lblTotal
  Vitórias = 0
  TotJogadas = 0
  lblTotal = ""
  lblTotJogadas = ""
  
End Sub
Private Sub btnJogada_Click()
  Dim P1 As Integer, P2 As Integer, P3 As Integer, P4 As Integer
  Dim Pagar As Currency
  
  
  TotJogadas = TotJogadas + 1
  
  'Cobrar pela jogada
  Vitórias = Vitórias - 1
  
  'Gerar resultados aleatórios
  P1 = Int(4 * Rnd + 1)
  P2 = Int(4 * Rnd + 1)
  P3 = Int(4 * Rnd + 1)
  P4 = Int(4 * Rnd + 1)
  
  'Mostrar ícones no visor 1
  Select Case P1
    Case COPAS:
      ImgVisor1.Picture = ImgCopas.Picture
    Case PAUS
      ImgVisor1.Picture = ImgPaus.Picture
    Case OUROS:
      ImgVisor1.Picture = ImgOuros.Picture
    Case ESPADAS:
      ImgVisor1.Picture = ImgEspadas.Picture
  End Select
  
  'Mostrar ícones no visor 2
  Select Case P2
    Case COPAS:
      ImgVisor2.Picture = ImgCopas.Picture
    Case PAUS
      ImgVisor2.Picture = ImgPaus.Picture
    Case OUROS:
      ImgVisor2.Picture = ImgOuros.Picture
    Case ESPADAS:
      ImgVisor2.Picture = ImgEspadas.Picture
  End Select
  
  'Mostrar ícones no visor 3
  Select Case P3
    Case COPAS:
      ImgVisor3.Picture = ImgCopas.Picture
    Case PAUS
      ImgVisor3.Picture = ImgPaus.Picture
    Case OUROS:
      ImgVisor3.Picture = ImgOuros.Picture
    Case ESPADAS:
      ImgVisor3.Picture = ImgEspadas.Picture
  End Select
  
  'Mostrar ícones no visor 4
  Select Case P4
    Case COPAS:
      ImgVisor4.Picture = ImgCopas.Picture
    Case PAUS
      ImgVisor4.Picture = ImgPaus.Picture
    Case OUROS:
      ImgVisor4.Picture = ImgOuros.Picture
    Case ESPADAS:
      ImgVisor4.Picture = ImgEspadas.Picture
  End Select
  
  'Verificar se o jogador ganhou
  If P1 = P2 And P2 = P3 And P3 = P4 Then
    If P1 = OUROS Then
      Pagar = 250
      MsgBox "Você tirou Sorte Grande"
    Else
      Pagar = 25
      MsgBox "Você venceu"
    End If
  Else
    Pagar = 0
  End If
  
  'Calcular e exibir o total acumulado
  Vitórias = Vitórias + Pagar
  lblTotal.Caption = Format(Vitórias, "R$0.00")
  lblTotJogadas.Caption = TotJogadas
  
  
End Sub


Private Sub btnSair_Click()
  End
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
  Randomize
  Vitórias = 0
End Sub


