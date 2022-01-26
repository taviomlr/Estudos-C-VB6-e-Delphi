VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoCadHistBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 61827 - Demanda: 62317
'---|   Ueder Budni (13/01/2015)
'---|---------------------------------------------------------------------------------------------------------------------------
Private mlngEnterpriseId        As Long
Private mlngCdEstabelecimento   As Long
Private mlngCdBanco             As Long
Private mlngCdHistorico         As Long
Private mstrDescricaoExtrato    As String
Private mstrComplDescr          As String
Private mstrTipoOperacao        As String

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

Public Property Let CdBanco(ByVal NewVal As Long)
    mlngCdBanco = NewVal
End Property

Public Property Get CdBanco() As Long
    CdBanco = mlngCdBanco
End Property

Public Property Let CdHistorico(ByVal NewVal As Long)
    mlngCdHistorico = NewVal
End Property

Public Property Get CdHistorico() As Long
    CdHistorico = mlngCdHistorico
End Property

Public Property Let DescricaoExtrato(ByVal NewVal As String)
    mstrDescricaoExtrato = NewVal
End Property

Public Property Get DescricaoExtrato() As String
    DescricaoExtrato = mstrDescricaoExtrato
End Property

Public Property Let ComplementoDescricao(ByVal NewVal As String)
    mstrComplDescr = NewVal
End Property

Public Property Get ComplementoDescricao() As String
    ComplementoDescricao = mstrComplDescr
End Property

Public Property Let TipoOperacao(ByVal NewVal As String)
    mstrTipoOperacao = NewVal
End Property

Public Property Get TipoOperacao() As String
    TipoOperacao = mstrTipoOperacao
End Property

Public Function Equals(objItem As VoCadHistBancario) As Boolean
    With objItem
        Equals = IIf(.CdBanco = mlngCdBanco And .CdHistorico = mlngCdHistorico, True, False)
    End With
End Function

