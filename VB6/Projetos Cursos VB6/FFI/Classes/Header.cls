VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Autor.......: Gustavo Cuman
'Data........: 08/12/2008
'Descrição...:

Private mstrCnpjEmpresa As String
Private mdatDtInicial As Date
Private mdatDtFinal As Date

Public Property Get cnpjEmpresa() As String
    cnpjEmpresa = mstrCnpjEmpresa
End Property

Public Property Let cnpjEmpresa(ByVal strCnpjEmpresa As String)
    mstrCnpjEmpresa = strCnpjEmpresa
End Property

Public Property Get dtInicial() As Date
    dtInicial = mdatDtInicial
End Property

Public Property Let dtInicial(ByVal datDtInicial As Date)
    mdatDtInicial = datDtInicial
End Property

Public Property Get dtFinal() As Date
    dtFinal = mdatDtFinal
End Property

Public Property Let dtFinal(ByVal datDtFinal As Date)
    mdatDtFinal = datDtFinal
End Property

'Descrição..:
'Parametros.: [String]
Public Sub Parser(strLinha As String)
    If UCase(Left(strLinha, 1)) = "H" Then
        mstrCnpjEmpresa = Mid(strLinha, 6, 14)
        mdatDtInicial = Mid(strLinha, 20, 2) & "/" & Mid(strLinha, 22, 2) & "/" & Mid(strLinha, 24, 4)
        mdatDtFinal = Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2) & "/" & Mid(strLinha, 32, 4)
    End If
End Sub

Public Function Validar(strCnpj As String) As Boolean
    Validar = (strCnpj = mstrCnpjEmpresa)
End Function



