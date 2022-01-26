VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoRateio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngCentro              As Long
Private mlngConta                   As Long
Private mdblPercentual              As Double
Private mdblValor                   As Double
Private mdblAcrescimo               As Double
Private mdblAbatimento              As Double

Public Property Get Centro() As Long
    Centro = mlngCentro
End Property

Public Property Let Centro(ByVal valor As Long)
    mlngCentro = valor
End Property

Public Property Get conta() As Long
    conta = mlngConta
End Property

Public Property Let conta(ByVal valor As Long)
    mlngConta = valor
End Property

Public Property Get Percentual() As Double
    Percentual = mdblPercentual
End Property

Public Property Let Percentual(ByVal valor As Double)
    mdblPercentual = valor
End Property

Public Property Get valor() As Double
    valor = mdblValor
End Property

Public Property Let valor(ByVal valor As Double)
    mdblValor = valor
End Property

Public Property Get Acrescimo() As Double
    Acrescimo = mdblAcrescimo
End Property

Public Property Let Acrescimo(ByVal valor As Double)
    mdblAcrescimo = valor
End Property

Public Property Get Abatimento() As Double
    Abatimento = mdblAbatimento
End Property

Public Property Let Abatimento(ByVal valor As Double)
    mdblAbatimento = valor
End Property

Public Property Get ValorTotal() As Double
    ValorTotal = mdblValor + mdblAcrescimo - mdblAbatimento
End Property
