VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParcela"
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
''  Implementation of the  Class cParcela
''  Generated by Enterprise Architect
''  Created on:      25-ago-2006 16:51:48
''  Original author: Dulcino J�nior
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private mintParcela As Integer
Private mdatVencimento As Date
Private mdblValor As Double
Private mlngCentroCusto As Long
Private mcurValorIpi As Currency
Private mcurValorIcmsSubstituicao As Currency
Private mcurValorDespesasFinanceiras As Currency

Public Property Let Parcela(ByVal NewVal As Integer)
        mintParcela = NewVal
End Property

Public Property Get Parcela() As Integer
        Parcela = mintParcela
End Property

Public Property Let Valor(ByVal NewVal As Double)
        mdblValor = NewVal
End Property

Public Property Get Valor() As Double
        Valor = mdblValor
End Property

Public Property Let Vencimento(ByVal NewVal As Date)
        mdatVencimento = NewVal
End Property

Public Property Get Vencimento() As Date
        Vencimento = mdatVencimento
End Property

Public Property Let CentroCusto(ByVal NewVal As Long)
        mlngCentroCusto = NewVal
End Property

Public Property Get CentroCusto() As Long
        CentroCusto = mlngCentroCusto
End Property

Public Property Get ValorIpi() As Currency
    ValorIpi = mcurValorIpi
End Property

Public Property Let ValorIpi(ByVal curValorIpi As Currency)
    mcurValorIpi = curValorIpi
End Property

Public Property Get ValorIcmsSubstituicao() As Currency
    ValorIcmsSubstituicao = mcurValorIcmsSubstituicao
End Property

Public Property Let ValorIcmsSubstituicao(ByVal curValorIcmsSubstituicao As Currency)
    mcurValorIcmsSubstituicao = curValorIcmsSubstituicao
End Property

Public Property Get ValorDespesasFinanceiras() As Currency
    ValorDespesasFinanceiras = mcurValorDespesasFinanceiras
End Property

Public Property Let ValorDespesasFinanceiras(ByVal curValorDespesasFinanceiras As Currency)
    mcurValorDespesasFinanceiras = curValorDespesasFinanceiras
End Property

Public Function equals(ByVal obj As Object) As Boolean
        Dim objParc As cParcela
        Set objParc = obj
        equals = mintParcela = objParc.Parcela
        Set objParc = Nothing
End Function

Public Function CriarObjeto(rd As IDBReader) As cParcela
    With rd
        mintParcela = .GetInteger("Parcela")
        mdatVencimento = .GetDate("vencimento")
        mdblValor = .GetDouble("Valor Original")
        mlngCentroCusto = .GetLong("Centro")
        mcurValorIpi = .GetCurrency("vl_ipi_parcela")
        mcurValorIcmsSubstituicao = .GetCurrency("vl_icms_substituicao_parcela")
        mcurValorDespesasFinanceiras = .GetCurrency("vl_despesas_financeiras_parcela")
    End With
    Set CriarObjeto = Me
End Function