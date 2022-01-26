VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoRateioLancamentoDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrpag_rec_origem As String
Private mstrcodigo_nota_origem As String
Private mlngnr_parcela_origem As Long
Private mstrpag_rec_destino As String
Private mstrcodigo_nota_destino As String
Private mlngnr_parcela_destino As Long
Private mlngcd_centro As Long
Private mlngCd_conta As Long
Private mdtadt_pagamento As Date
Private mdblvl_valor As Double
Private mstrcd_empresa_origem As String
Private mstrtp_registro_origem As String
Private mstrcd_empresa_destino As String
Private mstrtp_registro_destino As String

Public Property Get pag_rec_origem() As String
    pag_rec_origem = mstrpag_rec_origem
End Property

Public Property Let pag_rec_origem(ByVal Valor As String)
    mstrpag_rec_origem = Valor
End Property

Public Property Get codigo_nota_origem() As String
    codigo_nota_origem = mstrcodigo_nota_origem
End Property

Public Property Let codigo_nota_origem(ByVal Valor As String)
    mstrcodigo_nota_origem = Valor
End Property

Public Property Get nr_parcela_origem() As Long
    nr_parcela_origem = mlngnr_parcela_origem
End Property

Public Property Let nr_parcela_origem(ByVal Valor As Long)
    mlngnr_parcela_origem = Valor
End Property

Public Property Get pag_rec_destino() As String
    pag_rec_destino = mstrpag_rec_destino
End Property

Public Property Let pag_rec_destino(ByVal Valor As String)
    mstrpag_rec_destino = Valor
End Property

Public Property Get codigo_nota_destino() As String
    codigo_nota_destino = mstrcodigo_nota_destino
End Property

Public Property Let codigo_nota_destino(ByVal Valor As String)
    mstrcodigo_nota_destino = Valor
End Property

Public Property Get nr_parcela_destino() As Long
    nr_parcela_destino = mlngnr_parcela_destino
End Property

Public Property Let nr_parcela_destino(ByVal Valor As Long)
    mlngnr_parcela_destino = Valor
End Property

Public Property Get cd_centro() As Long
    cd_centro = mlngcd_centro
End Property

Public Property Let cd_centro(ByVal Valor As Long)
    mlngcd_centro = Valor
End Property

Public Property Get Cd_conta() As Long
    Cd_conta = mlngCd_conta
End Property

Public Property Let Cd_conta(ByVal Valor As Long)
    mlngCd_conta = Valor
End Property

Public Property Get dt_pagamento() As Date
    dt_pagamento = mdtadt_pagamento
End Property

Public Property Let dt_pagamento(ByVal Valor As Date)
    mdtadt_pagamento = Valor
End Property

Public Property Get vl_valor() As Double
    vl_valor = mdblvl_valor
End Property

Public Property Let vl_valor(ByVal Valor As Double)
    mdblvl_valor = Valor
End Property

Public Property Get cd_empresa_origem() As String
    cd_empresa_origem = mstrcd_empresa_origem
End Property

Public Property Let cd_empresa_origem(ByVal Valor As String)
    mstrcd_empresa_origem = Valor
End Property

Public Property Get tp_registro_origem() As String
    tp_registro_origem = mstrtp_registro_origem
End Property

Public Property Let tp_registro_origem(ByVal Valor As String)
    mstrtp_registro_origem = Valor
End Property

Public Property Get tp_registro_destino() As String
    tp_registro_destino = mstrtp_registro_destino
End Property

Public Property Let tp_registro_destino(ByVal Valor As String)
    mstrtp_registro_destino = Valor
End Property

Public Property Get cd_empresa_destino() As String
    cd_empresa_destino = mstrcd_empresa_destino
End Property

Public Property Let cd_empresa_destino(ByVal Valor As String)
    mstrcd_empresa_destino = Valor
End Property


