VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoImpDigExtratoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 61827 - Demanda: 61828
'---|   Ueder Budni (12/01/2015)
'---|---------------------------------------------------------------------------------------------------------------------------
Private mlngEnterpriseId        As Long
Private mlngCdEstabelecimento   As Long
Private mlngCdExtrato           As Long
Private mlngCdBanco             As Long
Private mlngSeqLancExtrato      As Long
Private mlngCdHistorico         As Long
Private mdteDataExtrato         As Date
Private mstrDescricao           As String
Private mstrDocumento           As String
Private mdblValor               As Double
Private mstrTipoOperacao        As String
Private mdblValorInterno        As Double
Private mblnConciliado          As Boolean
Private mdteDtConciliacao       As Date
Private mstrDescricaoHistorico  As String

Public Property Let EnterpriseId(ByVal NewVal As Long)
    mlngEnterpriseId = NewVal
End Property

Public Property Get EnterpriseId() As Long
    EnterpriseId = mlngEnterpriseId
End Property

Public Property Let CdEstabelecimento(ByVal NewVal As Long)
    mlngCdEstabelecimento = NewVal
End Property

Public Property Get CdEstabelecimento() As Long
    CdEstabelecimento = mlngCdEstabelecimento
End Property

Public Property Let CdExtrato(ByVal NewVal As Long)
    mlngCdExtrato = NewVal
End Property

Public Property Get CdExtrato() As Long
    CdExtrato = mlngCdExtrato
End Property

Public Property Let CdBanco(ByVal NewVal As Long)
    mlngCdBanco = NewVal
End Property

Public Property Get CdBanco() As Long
    CdBanco = mlngCdBanco
End Property

Public Property Let SeqLancExtrato(ByVal NewVal As Long)
    mlngSeqLancExtrato = NewVal
End Property

Public Property Get SeqLancExtrato() As Long
    SeqLancExtrato = mlngSeqLancExtrato
End Property

Public Property Let CdHistorico(ByVal NewVal As Long)
    mlngCdHistorico = NewVal
End Property

Public Property Get CdHistorico() As Long
    CdHistorico = mlngCdHistorico
End Property

Public Property Let DataExtrato(ByVal NewVal As Date)
    mdteDataExtrato = NewVal
End Property

Public Property Get DataExtrato() As Date
    DataExtrato = mdteDataExtrato
End Property

Public Property Let Descricao(ByVal NewVal As String)
    mstrDescricao = NewVal
End Property

Public Property Get Descricao() As String
    Descricao = mstrDescricao
End Property

Public Property Let DescricaoHistorico(ByVal NewVal As String)
    mstrDescricaoHistorico = NewVal
End Property

Public Property Get DescricaoHistorico() As String
    DescricaoHistorico = mstrDescricaoHistorico
End Property

Public Property Let Documento(ByVal NewVal As String)
    mstrDocumento = NewVal
End Property

Public Property Get Documento() As String
    Documento = mstrDocumento
End Property

Public Property Let valor(ByVal NewVal As Double)
    mdblValor = NewVal
End Property

Public Property Get valor() As Double
    valor = mdblValor
End Property

Public Property Let TipoOperacao(ByVal NewVal As String)
    mstrTipoOperacao = NewVal
End Property

Public Property Get TipoOperacao() As String
    TipoOperacao = mstrTipoOperacao
End Property

Public Property Let ValorInterno(ByVal NewVal As Double)
    mdblValorInterno = NewVal
End Property

Public Property Get ValorInterno() As Double
    ValorInterno = mdblValorInterno
End Property

Public Property Let Conciliado(ByVal NewVal As Boolean)
    mblnConciliado = NewVal
End Property

Public Property Get Conciliado() As Boolean
    Conciliado = mblnConciliado
End Property

Public Property Let DataConciliacao(ByVal NewVal As Date)
    mdteDtConciliacao = NewVal
End Property

Public Property Get DataConciliacao() As Date
    DataConciliacao = mdteDtConciliacao
End Property

Public Function Equals(objItem As VoImpDigExtratoBancario) As Boolean
    With objItem
        Equals = IIf(.CdBanco = mlngCdBanco And .SeqLancExtrato = mlngSeqLancExtrato, True, False)
    End With
End Function
