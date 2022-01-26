VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mlngBanco As Long
Private mlngCheque As Long
Private mstrSituacao As String
Private mstrNominal As String
Private mstrHistorico As String

Public Property Get Banco() As Long
    Banco = mlngBanco
End Property

Public Property Let Banco(ByVal lngBanco As Long)
    mlngBanco = lngBanco
End Property

Public Property Get Cheque() As Long
    Cheque = mlngCheque
End Property

Public Property Let Cheque(ByVal lngCheque As Long)
    mlngCheque = lngCheque
End Property

Public Property Get Situacao() As String
    Situacao = mstrSituacao
End Property

Public Property Let Situacao(ByVal strSituacao As String)
    mstrSituacao = strSituacao
End Property

Public Property Get Nominal() As String
    Nominal = mstrNominal
End Property

Public Property Let Nominal(ByVal strNominal As String)
    mstrNominal = strNominal
End Property

Public Property Get Historico() As String
    Historico = mstrHistorico
End Property

Public Property Let Historico(ByVal strHistorico As String)
    mstrHistorico = strHistorico
End Property
