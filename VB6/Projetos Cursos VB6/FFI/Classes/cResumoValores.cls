VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResumoValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''
''  Implementation of the  Class cResumoValores
''  Generated by Enterprise Architect
''  Created on:      05-jun-2007 17:23:19
''  Original author: Dulcino J�nior
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private mcurBaseIpi As Currency
Private mcurBaseICMS As Currency
Private mcurBaseIcmsSubstituicao As Currency
Private mcurBasePisCofins As Currency
Private mdblQuantidade As Double
Private mcurValorIpi As Currency
Private mcurValorICMS As Currency
Private mcurValorIpiBaseIcms As Currency
Private mcurValorIcmsSubstituicao As Currency
Private mcurValorPIS As Currency
Private mcurValorCOFINS As Currency
Private mcurValorCsll As Currency
Private mcurValorISSRF As Currency
Private mcurValorIrrf As Currency
Private mcurValorINSS As Currency
Private mcurValorISS As Currency
Private mcurValorImpostoImportacao As Currency
Private mcurValorOutros As Currency
Private mcurValorFrete As Currency
Private mdblPercentualDesconto As Double
Private mcurValorDesconto As Currency
Private mcurValorSeguro As Currency
Private mcurValorDespesaFinanceira As Currency
Private mcurValorOutrasDespesasAcessorias As Currency
Private mcurTotalMercadorias As Currency
Private mcurTotalServicos As Currency
Private mcurTotalGeral As Currency
Private mcurBasePisNc As Currency
Private mcurValorPisNc As Currency
Private mcurBaseCOFINSNc As Currency
Private mcurValorCOFINSNc As Currency

Public Property Get BaseICMS() As Currency
    BaseICMS = mcurBaseICMS
End Property

Public Property Let BaseICMS(ByVal NewVal As Currency)
    mcurBaseICMS = NewVal
End Property

Public Property Get BaseIcmsSubstituicao() As Currency
    BaseIcmsSubstituicao = mcurBaseIcmsSubstituicao
End Property

Public Property Let BaseIcmsSubstituicao(ByVal NewVal As Currency)
    mcurBaseIcmsSubstituicao = NewVal
End Property

Public Property Let BaseIpi(ByVal NewVal As Currency)
    mcurBaseIpi = NewVal
End Property

Public Property Get BaseIpi() As Currency
    BaseIpi = mcurBaseIpi
End Property

Public Property Get BasePisCofins() As Currency
    BasePisCofins = mcurBasePisCofins
End Property

Public Property Let BasePisCofins(ByVal NewVal As Currency)
    mcurBasePisCofins = NewVal
End Property

Public Property Let PercentualDesconto(ByVal NewVal As Double)
    mdblPercentualDesconto = NewVal
End Property

Public Property Get PercentualDesconto() As Double
    PercentualDesconto = mdblPercentualDesconto
End Property

Public Property Get Quantidade() As Double
    Quantidade = mdblQuantidade
End Property

Public Property Let Quantidade(ByVal NewVal As Double)
    mdblQuantidade = NewVal
End Property

Public Property Get TotalGeral() As Currency
    TotalGeral = mcurTotalGeral
End Property

Public Property Let TotalGeral(ByVal NewVal As Currency)
    mcurTotalGeral = NewVal
End Property

Public Property Get TotalMercadorias() As Currency
    TotalMercadorias = mcurTotalMercadorias
End Property

Public Property Let TotalMercadorias(ByVal NewVal As Currency)
    mcurTotalMercadorias = NewVal
End Property

Public Property Let TotalServicos(ByVal NewVal As Currency)
    mcurTotalServicos = NewVal
End Property

Public Property Get TotalServicos() As Currency
    TotalServicos = mcurTotalServicos
End Property

Public Property Get ValorCOFINS() As Currency
    ValorCOFINS = mcurValorCOFINS
End Property

Public Property Let ValorCOFINS(ByVal NewVal As Currency)
    mcurValorCOFINS = NewVal
End Property

Public Property Get ValorCsll() As Currency
    ValorCsll = mcurValorCsll
End Property

Public Property Let ValorCsll(ByVal NewVal As Currency)
    mcurValorCsll = NewVal
End Property

Public Property Get ValorDesconto() As Currency
    ValorDesconto = mcurValorDesconto
End Property

Public Property Let ValorDesconto(ByVal NewVal As Currency)
    mcurValorDesconto = NewVal
End Property

Public Property Get ValorDespesaFinanceira() As Currency
    ValorDespesaFinanceira = mcurValorDespesaFinanceira
End Property

Public Property Let ValorDespesaFinanceira(ByVal NewVal As Currency)
    mcurValorDespesaFinanceira = NewVal
End Property

Public Property Let ValorFrete(ByVal NewVal As Currency)
    mcurValorFrete = NewVal
End Property

Public Property Get ValorFrete() As Currency
    ValorFrete = mcurValorFrete
End Property

Public Property Get ValorICMS() As Currency
    ValorICMS = mcurValorICMS
End Property

Public Property Let ValorICMS(ByVal NewVal As Currency)
    mcurValorICMS = NewVal
End Property

Public Property Let ValorIcmsSubstituicao(ByVal NewVal As Currency)
    mcurValorIcmsSubstituicao = NewVal
End Property

Public Property Get ValorIcmsSubstituicao() As Currency
    ValorIcmsSubstituicao = mcurValorIcmsSubstituicao
End Property

Public Property Get ValorImpostoImportacao() As Currency
    ValorImpostoImportacao = mcurValorImpostoImportacao
End Property

Public Property Let ValorImpostoImportacao(ByVal NewVal As Currency)
    mcurValorImpostoImportacao = NewVal
End Property

Public Property Get ValorINSS() As Currency
    ValorINSS = mcurValorINSS
End Property

Public Property Let ValorINSS(ByVal NewVal As Currency)
    mcurValorINSS = NewVal
End Property

Public Property Let ValorIpi(ByVal NewVal As Currency)
    mcurValorIpi = NewVal
End Property

Public Property Get ValorIpi() As Currency
    ValorIpi = mcurValorIpi
End Property

Public Property Get ValorIpiBaseIcms() As Currency
    ValorIpiBaseIcms = mcurValorIpiBaseIcms
End Property

Public Property Let ValorIpiBaseIcms(ByVal NewVal As Currency)
    mcurValorIpiBaseIcms = NewVal
End Property

Public Property Get ValorIRRF() As Currency
    ValorIRRF = mcurValorIrrf
End Property

Public Property Let ValorIRRF(ByVal NewVal As Currency)
    mcurValorIrrf = NewVal
End Property

Public Property Let ValorISS(ByVal NewVal As Currency)
    mcurValorISS = NewVal
End Property

Public Property Get ValorISS() As Currency
    ValorISS = mcurValorISS
End Property

Public Property Get ValorISSRF() As Currency
    ValorISSRF = mcurValorISSRF
End Property

Public Property Let ValorISSRF(ByVal NewVal As Currency)
    mcurValorISSRF = NewVal
End Property

Public Property Get ValorOutrasDespesasAcessorias() As Currency
    ValorOutrasDespesasAcessorias = mcurValorOutrasDespesasAcessorias
End Property

Public Property Let ValorOutrasDespesasAcessorias(ByVal NewVal As Currency)
    mcurValorOutrasDespesasAcessorias = NewVal
End Property

Public Property Let ValorOutros(ByVal NewVal As Currency)
    mcurValorOutros = NewVal
End Property

Public Property Get ValorOutros() As Currency
    ValorOutros = mcurValorOutros
End Property

Public Property Let ValorPIS(ByVal NewVal As Currency)
    mcurValorPIS = NewVal
End Property

Public Property Get ValorPIS() As Currency
    ValorPIS = mcurValorPIS
End Property

Public Property Get ValorSeguro() As Currency
    ValorSeguro = mcurValorSeguro
End Property

Public Property Let ValorSeguro(ByVal NewVal As Currency)
    mcurValorSeguro = NewVal
End Property

Public Property Get BasePisNc() As Currency
    BasePisNc = mcurBasePisNc
End Property

Public Property Let BasePisNc(ByVal NewVal As Currency)
    mcurBasePisNc = NewVal
End Property

Public Property Get ValorPisNc() As Currency
    ValorPisNc = mcurValorPisNc
End Property

Public Property Let ValorPisNc(ByVal NewVal As Currency)
    mcurValorPisNc = NewVal
End Property

Public Property Get BaseCOFINSNc() As Currency
    BaseCOFINSNc = mcurBaseCOFINSNc
End Property

Public Property Let BaseCOFINSNc(ByVal NewVal As Currency)
    mcurBaseCOFINSNc = NewVal
End Property

Public Property Get ValorCOFINSNc() As Currency
    ValorCOFINSNc = mcurValorCOFINSNc
End Property

Public Property Let ValorCOFINSNc(ByVal NewVal As Currency)
    mcurValorCOFINSNc = NewVal
End Property