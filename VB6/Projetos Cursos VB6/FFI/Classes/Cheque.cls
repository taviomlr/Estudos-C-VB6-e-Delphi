VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Autor.......: Gustavo Cuman
'Data........: 08/12/2008
'Descrição...:

Private mlngCodigoBanco As Long
Private mstrNumeroAgencia As String
Private mstrContaCorrente As String
Private mstrNomeCorrentista As String

Public Property Get CodigoBanco() As Long
    CodigoBanco = mlngCodigoBanco
End Property

Public Property Let CodigoBanco(ByVal lngCodigoBanco As Long)
    mlngCodigoBanco = lngCodigoBanco
End Property

Public Property Get NumeroAgencia() As String
    NumeroAgencia = mstrNumeroAgencia
End Property

Public Property Let NumeroAgencia(ByVal strNumeroAgencia As String)
    mstrNumeroAgencia = strNumeroAgencia
End Property

Public Property Get ContaCorrente() As String
    ContaCorrente = mstrContaCorrente
End Property

Public Property Let ContaCorrente(ByVal strContaCorrente As String)
    mstrContaCorrente = strContaCorrente
End Property

Public Property Get NomeCorrentista() As String
    NomeCorrentista = mstrNomeCorrentista
End Property

Public Property Let NomeCorrentista(ByVal strNomeCorrentista As String)
    mstrNomeCorrentista = strNomeCorrentista
End Property

'Descrição..:
'Parametros.: [String]
Public Sub Parser(strLinha As String)
    If UCase(Left(strLinha, 1)) = "C" Then
        mlngCodigoBanco = Mid(strLinha, 2, 9)
        mstrNumeroAgencia = Mid(strLinha, 11, 10)
        mstrContaCorrente = Mid(strLinha, 21, 20)
        mstrNomeCorrentista = Mid(strLinha, 41, 60)
    End If
End Sub
