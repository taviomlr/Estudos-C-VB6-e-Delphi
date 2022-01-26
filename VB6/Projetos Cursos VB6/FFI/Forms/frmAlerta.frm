VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmAlerta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alerta de Títulos Vencidos e à Vencer"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11055
   Icon            =   "frmAlerta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   5445
      Left            =   30
      TabIndex        =   2
      Top             =   -60
      Width           =   9675
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridAlerta 
         Height          =   5175
         Left            =   90
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9128
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraBotoes 
      Height          =   5460
      Left            =   9720
      TabIndex        =   0
      Top             =   -60
      Width           =   1320
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmAlerta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Davi Brito - #169397 - 22/05/2017
Private Const mColunaPagRec = 0
Private Const mColunaOrigem = 1
Private Const mColunaCodigo = 2
Private Const mColunaParcela = 3
Private Const mColunaEmpresa = 4
Private Const mColunaValor = 5
Private Const mColunaVencimento = 6
Private Const mColunaLiberacao = 7

Private Function GetColunasGrid() As String

    GetColunasGrid = "campo=PagRec;label=Pag/Rec;tamanho=800" & _
                            "|campo=Origem;Label=Origem;tamanho=1100" & _
                            "|campo=Código;label=Número;tamanho=900;tipo=tpColGridInteger" & _
                            "|campo=Parcela;Label=Parc.;tamanho=600" & _
                            "|campo=Empresa;label=Empresa;tamanho=2000" & _
                            "|campo=Valor Original;Label=Valor;tamanho=1500;tipo=tpColGridDouble;formato=###,###,##0.00" & _
                            "|campo=Vencimento;Label=Vencimento;tamanho=1100" & _
                            "|campo=Liberação;Label=Liberação;tamanho=1000"
                            
End Function

Public Sub CarregarGrid(ByRef rs As Object)
    Call CarregaHFlexGrid(gridAlerta, rs, GetColunasGrid)
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = fMain.Icon
End Sub
