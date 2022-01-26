VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BizCentroCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function validarDataLimiteCentroCusto(ByVal CentroCusto As Long, ByVal Data As String, ByRef col As Collection) As Boolean
    Dim dao                 As New DaoCentroCusto
    Dim objMensagemErro     As New clsMensagemErro
    Dim DataLimite          As String
    
    DataLimite = dao.retornaDataLimite(CentroCusto)
    If EData(DataLimite) Then
        If CDateDef(DataLimite) < CDateDef(Data) Then
            objMensagemErro.mensagem = "A Data do lançamento ultrapassa a 'Data Limite' para movimentação do Centro de Custo."
            objMensagemErro.status = atencao
            col.add objMensagemErro
        End If
    End If
End Function

