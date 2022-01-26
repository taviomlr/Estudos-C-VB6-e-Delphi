VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeracaoDuplicataReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbooPermiteAlteracao        As Boolean
'VARIAVEIS DA PARCELA.
Private mstrP_PagRec                As String
Private mlngP_Nota                  As Long
Private mstrP_Empresa               As String
Private mstrP_Tipo                  As String
Private mintP_Parcela               As Integer
Private mstrP_Descricao             As String
Private mdtaP_Emissao               As Date
Private mdtaP_Vencimento            As Date
Private mdtaP_Pagamento             As String
Private mdblP_Valor_Original        As Double
Private mlngP_Banco                 As Long
Private mlngP_Conta                 As Long
Private mlngP_Centro                As Long
Private mstrP_Moeda                 As String
Private mlngP_cd_operacao_contabil  As Long
Private mobjParcelas As cColecaoGeracaoTituloReceber
'pt. 85684 - Moacir Pfau(11/07/2008) - GERAÇÃO DE TITULOS

'PARCELAS
Public Property Let P_PagRec(ByVal NewVal As String)
    mstrP_PagRec = NewVal
End Property

Public Property Get P_PagRec() As String
    P_PagRec = mstrP_PagRec
End Property

Public Property Let P_Nota(ByVal NewVal As Long)
    mlngP_Nota = NewVal
End Property

Public Property Get P_Nota() As Long
    P_Nota = mlngP_Nota
End Property

Public Property Let P_Empresa(ByVal NewVal As String)
    mstrP_Empresa = NewVal
End Property

Public Property Get P_Empresa() As String
    P_Empresa = mstrP_Empresa
End Property

Public Property Let P_Tipo(ByVal NewVal As String)
    mstrP_Tipo = NewVal
End Property

Public Property Get P_Tipo() As String
    P_Tipo = mstrP_Tipo
End Property

Public Property Let P_Parcela(ByVal NewVal As Integer)
    mintP_Parcela = NewVal
End Property

Public Property Get P_Parcela() As Integer
    P_Parcela = mintP_Parcela
End Property

Public Property Let P_Descricao(ByVal NewVal As String)
    mstrP_Descricao = NewVal
End Property

Public Property Get P_Descricao() As String
    P_Descricao = mstrP_Descricao
End Property

Public Property Let P_Emissao(ByVal NewVal As Date)
    mdtaP_Emissao = NewVal
End Property

Public Property Get P_Emissao() As Date
    P_Emissao = mdtaP_Emissao
End Property

Public Property Let P_Vencimento(ByVal NewVal As Date)
    mdtaP_Vencimento = NewVal
End Property

Public Property Get P_Vencimento() As Date
    P_Vencimento = mdtaP_Vencimento
End Property

Public Property Let P_Pagamento(ByVal NewVal As String)
    mdtaP_Pagamento = NewVal
End Property

Public Property Get P_Pagamento() As String
    P_Pagamento = mdtaP_Pagamento
End Property

Public Property Let P_Valor_Original(ByVal NewVal As Double)
    mdblP_Valor_Original = NewVal
End Property

Public Property Get P_Valor_Original() As Double
    P_Valor_Original = mdblP_Valor_Original
End Property

Public Property Let P_Banco(ByVal NewVal As Long)
    mlngP_Banco = NewVal
End Property

Public Property Get P_Banco() As Long
    P_Banco = mlngP_Banco
End Property

Public Property Let P_Conta(ByVal NewVal As Long)
    mlngP_Conta = NewVal
End Property

Public Property Get P_Conta() As Long
    P_Conta = mlngP_Conta
End Property

Public Property Let P_Centro(ByVal NewVal As Long)
    mlngP_Centro = NewVal
End Property

Public Property Get P_Centro() As Long
    P_Centro = mlngP_Centro
End Property

Public Property Let P_Moeda(ByVal NewVal As String)
    mstrP_Moeda = NewVal
End Property

Public Property Get P_Moeda() As String
    P_Moeda = mstrP_Moeda
End Property

Public Property Let P_cd_operacao_contabil(ByVal NewVal As Long)
    mlngP_cd_operacao_contabil = NewVal
End Property

Public Property Get P_cd_operacao_contabil() As Long
    P_cd_operacao_contabil = mlngP_cd_operacao_contabil
End Property

Public Property Get PermiteAlteracao() As Boolean
    PermiteAlteracao = mbooPermiteAlteracao
End Property

Public Property Let PermiteAlteracao(ByVal NewVal As Boolean)
    mbooPermiteAlteracao = NewVal
End Property

Private Sub Class_Initialize()
    mbooPermiteAlteracao = True
    Set mobjParcelas = New cColecaoGeracaoTituloReceber
End Sub

Public Property Get parcelas() As cColecaoGeracaoTituloReceber
    Set parcelas = mobjParcelas
End Property

Public Property Let parcelas(ByVal NewVal As cColecaoGeracaoTituloReceber)
    Set mobjParcelas = NewVal
End Property

Private Sub Class_Terminate()
    Set mobjParcelas = Nothing
End Sub

Public Function equals(obj As Object) As Boolean
    Dim booIgual As Boolean
    Dim newObj As cGeracaoDuplicataReceber
    
    Set newObj = obj
    booIgual = (newObj.P_Nota = mlngP_Nota)
    booIgual = booIgual And (newObj.P_Parcela = mintP_Parcela)
    equals = booIgual
End Function
