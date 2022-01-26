VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DaoExtratoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 56943 - Hist: 56927
'---|   Vinicius Elyseu - 20/11/2014
'---|---------------------------------------------------------------------------------------------------------------------------

Private mApp            As IApplication
Private Const NomeTabela$ = "FFIExtratoBancario"


Public Function ConciliaExtrato(ByVal blnconcilia As Boolean, ByVal strBanco As String, ByVal lngExtrato As Integer, ByVal lngSeqExtrato As Long) As Boolean
    Dim cmd            As IDBUpdateCommand
    Dim rdResult       As IDBReader
    Dim strCodigo      As String
On Error GoTo erro_alterando
   
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateUpdateCommand
    cmd.Table = NomeTabela$
    
    Call cmd.Filter.Append("cd_Extrato = @pCdExtr")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCdExtr", lngExtrato, dbFieldTypeLong))
    
    Call cmd.Filter.Append("cd_banco = @pBanco")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBanco", strBanco, dbFieldTypeString))
    
    Call cmd.Filter.Append("seq_lanc_extrato = @pSeqExt")
    Call cmd.Parameters.add(cmd.CreateParameter("@pSeqExt", lngSeqExtrato, dbFieldTypeLong))
    
    Call cmd.AddValue("conciliado", "@pConciliado")
    Call cmd.Parameters.add(cmd.CreateParameter("@pConciliado", blnconcilia, dbFieldTypeBool))
    
    Call cmd.AddValue("dt_conciliacao", "@pDtConc")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDtConc", Now(), dbFieldTypeDateTime))
   
    ConciliaExtrato = Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd)
    Aplicacao.Disconnect
    Exit Function

erro_alterando:
    err.Clear
    ConciliaExtrato = False
    Aplicacao.Disconnect
End Function

Public Function ExisteExtratoVinculado(ByVal lngCdHistorico As Long, ByVal strBanco As String) As Boolean
    Dim cmd            As IDBSelectCommand
    Dim rdResult       As IDBReader
On Error GoTo erro_consultando
        
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "FFIExtratoBancario"
    
    Call cmd.Filter.Append("cd_historico = @pCodHist")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCodHist", lngCdHistorico, dbFieldTypeString))
    
    Call cmd.Filter.Append("cd_banco = @pBanco")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBanco", strBanco, dbFieldTypeString))
        
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    
    If Not rdResult.EOF Then
        ExisteExtratoVinculado = True
    End If
    
    Aplicacao.Disconnect
    Exit Function
erro_consultando:
    err.Clear
    ExisteExtratoVinculado = False
    Aplicacao.Disconnect
End Function


