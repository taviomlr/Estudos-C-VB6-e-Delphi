VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DaoImpDigExtratoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 61827 - Demanda: 61828
'---|   Ueder Budni (12/01/2015)
'---|---------------------------------------------------------------------------------------------------------------------------

Private mApp            As IApplication

Private Const NOME_TABELA = "FFIExtratoBancario"

Public Sub init(App As IApplication)
    Set mApp = App
End Sub

Public Sub Initialize()
    If (mApp Is Nothing) Then
        Aplicacao.Connect
        Set mApp = Aplicacao
    End If
End Sub

Public Sub Terminate()
    If (Not mApp Is Nothing) Then
        Set mApp = Nothing
        Aplicacao.Disconnect
    End If
End Sub


Public Function CarregaColecao(ByVal lngBanco As String, Optional ByVal intMes As Integer, Optional ByVal intAno As Integer, Optional ByVal lngNumExtrato As Long, Optional blnMostraConciliado As Boolean, Optional ByVal strOptTipo As String) As ColImpDigExtratoBancario
    Dim objExtratoBanc  As VoImpDigExtratoBancario
    Dim selCmd          As IDBSelectCommand
    Dim rdResult        As IDBReader
    Dim colTmp          As ColImpDigExtratoBancario
           
On Error GoTo Erro
    Aplicacao.Connect
            
    Set selCmd = Aplicacao.CreateSelectCommand
    Set colTmp = New ColImpDigExtratoBancario
    
    With selCmd
        
        .Table.TableName = NOME_TABELA
        
        If lngBanco > 0 Then
            Call .Filter.Append("cd_banco = @pBanco")
            Call .Parameters.add(selCmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        End If
        
        If Not IsMissing(intAno) And intAno <> 1899 Then
            Call .Filter.Append("YEAR(data_extrato) = @pAno")
            Call .Parameters.add(selCmd.CreateParameter("@pAno", intAno, dbFieldTypeInt))
        End If
        
        If Not IsMissing(intMes) And intAno <> 1899 Then
            Call .Filter.Append("MONTH(data_extrato) = @pMes")
            Call .Parameters.add(selCmd.CreateParameter("@pMes", intMes, dbFieldTypeInt))
        End If
        
        If Not IsMissing(lngNumExtrato) Then
            Call .Filter.Append("cd_extrato = @pNumExtrato")
            Call .Parameters.add(selCmd.CreateParameter("@pNumExtrato", lngNumExtrato, dbFieldTypeLong))
        End If
        
        If Not blnMostraConciliado Then
            Call .Filter.Append("conciliado = @pConciliado")
            Call .Parameters.add(selCmd.CreateParameter("@pConciliado", 0, dbFieldTypeBool))
        End If
        
        If strOptTipo <> "" And strOptTipo <> "A" Then
            Call .Filter.Append("tipo_operacao = @pTpOP")
            Call .Parameters.add(selCmd.CreateParameter("@pTpOP", IIf(strOptTipo = "P", "D", "C"), dbFieldTypeString))
        End If
        
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
        While Not rdResult.EOF
            Set objExtratoBanc = New VoImpDigExtratoBancario
            Set objExtratoBanc = CarregaVO(rdResult)
            Call colTmp.add(objExtratoBanc)
            Set objExtratoBanc = Nothing
            rdResult.MoveNext
        Wend
    End With
    Set CarregaColecao = colTmp
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    
    Aplicacao.Disconnect
    Exit Function
Erro:
    FinallyConnection Aplicacao
End Function

Private Function CarregaVO(ByVal rs As IDBReader) As VoImpDigExtratoBancario
    Dim objTmp As New VoImpDigExtratoBancario

On Error GoTo Erro:
    If Not rs Is Nothing Then
        With objTmp
            .EnterpriseId = rs.GetLong("Enterprise_id")
            .CdEstabelecimento = rs.GetLong("Cd_estabelecimento")
            .CdExtrato = rs.GetLong("cd_extrato")
            .CdBanco = rs.GetLong("Cd_banco")
            .SeqLancExtrato = rs.GetLong("seq_lanc_extrato")
            .CdHistorico = rs.GetLong("cd_historico")
            .DataExtrato = rs.GetDate("data_extrato")
            .Descricao = rs.GetString("Descricao")
            .Documento = rs.GetString("Documento")
            .Valor = rs.GetDouble("Valor")
            .TipoOperacao = rs.GetString("tipo_operacao")
            .ValorInterno = rs.GetDouble("valor_interno")
            .Conciliado = rs.GetBoolean("Conciliado")
            .DataConciliacao = rs.GetDate("dt_conciliacao")
        End With
    End If
    Set CarregaVO = objTmp
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi possível carregar o objeto! (DaoImpDigExtratoBancario.CarregaVO)"
    
End Function

Public Function GravaColecao(ByVal col As ColImpDigExtratoBancario, lngBanco As Long, intMes As Integer, intAno As Integer, Optional lngExtrato As Long)
    Dim obj         As VoImpDigExtratoBancario
    Dim Gravando    As Boolean
    
    Gravando = True
On Error GoTo err
    Aplicacao.Connect
    Set mApp = Aplicacao
    mApp.BeginTransaction
    If col.Count > 0 Then
        col.MoveFirst
        Gravando = ExcluirTodos(lngBanco, intMes, intAno, lngExtrato)
        While Not col.EOF And Gravando
            Set obj = col.CurrentObject
            Gravando = Gravando And GravarVO(obj, mApp, lngExtrato)
            col.MoveNext
        Wend
    End If
    mApp.CommitTransaction
    Set mApp = Nothing
    
    GravaColecao = Gravando
    Aplicacao.Disconnect
    
    Exit Function
err:
    mApp.RollbackTransaction
    Aplicacao.Disconnect
    GravaColecao = False
    Call Throw(err)
End Function


Public Function ExcluirTodos(ByVal lngBanco As Long, intMes As Integer, intAno As Integer, Optional lngExtrato As Long) As Boolean
    Dim cmd     As IDBDeleteCommand
    
On Error GoTo erro_excluindo
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateDeleteCommand
    With cmd
        .Table = NOME_TABELA
    
        Call .Filter.Append("cd_banco = @pBanco")
        Call .Parameters.add(cmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
        Call .Filter.Append("YEAR(data_extrato) = @pAno")
        Call .Parameters.add(cmd.CreateParameter("@pAno", intAno, dbFieldTypeInt))
        
        Call .Filter.Append("MONTH(data_extrato) = @pMes")
        Call .Parameters.add(cmd.CreateParameter("@pMes", intMes, dbFieldTypeInt))
        
        If lngExtrato > 0 Then
            Call .Filter.Append("cd_extrato = @pExtrato")
            Call .Parameters.add(cmd.CreateParameter("@pExtrato", lngExtrato, dbFieldTypeLong))
        End If
    End With
    Call Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd)
    ExcluirTodos = True
    Aplicacao.Disconnect
    Set cmd = Nothing
    
    Exit Function
erro_excluindo:
    Set cmd = Nothing
    ExcluirTodos = False
    Call Throw(err)
End Function


Public Function GravarVO(ByVal obj As VoImpDigExtratoBancario, mApp As IApplication, Optional lngExtrato As Long) As Boolean
    Dim cmd        As IDBInsertCommand
    
On Error GoTo erro_gravando
    
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = NOME_TABELA
        
        Call .AddValue("[enterprise_id]", "@pEnterpriseID")
        Call .Parameters.add(.CreateParameter("@pEnterpriseID", ModGeral.EnterpriseId, dbFieldTypeLong))
    
        Call .AddValue("[cd_estabelecimento]", "@pCdEst")
        Call .Parameters.add(.CreateParameter("@pCdEst", ModGeral.CdEstabelecimento, dbFieldTypeLong))
        
        Call .AddValue("[cd_extrato]", "@pCdExtr")
        Call .Parameters.add(.CreateParameter("@pCdExtr", lngExtrato, dbFieldTypeLong))
                
        Call .AddValue("[cd_banco]", "@pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", obj.CdBanco, dbFieldTypeLong))
        
        Call .AddValue("[seq_lanc_extrato]", "@pSeqLanc")
        Call .Parameters.add(.CreateParameter("@pSeqLanc", obj.SeqLancExtrato, dbFieldTypeLong))
        
        Call PrepareCommand(cmd, obj)
    End With
    
    GravarVO = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    Set cmd = Nothing
    Exit Function
erro_gravando:
    GravarVO = False
    Set cmd = Nothing
    Call Throw(err)
End Function

Public Function GravarVOHistorico(ByVal obj As VoImpDigExtratoBancario, strDescHist As String, mApp As IApplication) As Boolean
    Dim cmd        As IDBInsertCommand
    
On Error GoTo erro_gravando
    
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = "FFIExtratoBancarioHistorico"
        
        Call .AddValue("[enterprise_id]", "@pEnterpriseID")
        Call .Parameters.add(.CreateParameter("@pEnterpriseID", ModGeral.EnterpriseId, dbFieldTypeLong))
    
        Call .AddValue("[cd_estabelecimento]", "@pCdEst")
        Call .Parameters.add(.CreateParameter("@pCdEst", ModGeral.CdEstabelecimento, dbFieldTypeLong))
        
        Call .AddValue("[cd_banco]", "@pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", obj.CdBanco, dbFieldTypeLong))
        
        Call .AddValue("[cd_historico]", "@pCdHist")
        Call .Parameters.add(.CreateParameter("@pCdHist", obj.CdHistorico, dbFieldTypeLong))
        
        Call PrepareCommandHistorico(cmd, strDescHist, obj)
    End With
    
    GravarVOHistorico = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    Set cmd = Nothing
    Exit Function
erro_gravando:
    GravarVOHistorico = False
    Set cmd = Nothing
    Call Throw(err)
End Function

Private Sub PrepareCommandHistorico(ByVal cmd As Object, strDescHist As String, ByVal obj As VoImpDigExtratoBancario)
    With cmd
        Call .AddValue("[descricao_extrato]", "@pDescrEx")
        Call .Parameters.add(.CreateParameter("@pDescrEx", strDescHist, dbFieldTypeString))
        
        Call .AddValue("[complemento_descricao]", "@pComplDescHist")
        Call .Parameters.add(.CreateParameter("@pComplDescHist", "", dbFieldTypeString))
        
        Call .AddValue("[tipo_operacao]", "@pTpOp")
        Call .Parameters.add(.CreateParameter("@pTpOp", obj.TipoOperacao, dbFieldTypeString))
        
    End With
End Sub
Private Sub PrepareCommand(ByVal cmd As Object, ByVal obj As VoImpDigExtratoBancario)
    With cmd
        Call .AddValue("[cd_historico]", "@pCdHist")
        Call .Parameters.add(.CreateParameter("@pCdHist", obj.CdHistorico, dbFieldTypeLong))
        
        Call .AddValue("[data_extrato]", "@pDtExtr")
        Call .Parameters.add(.CreateParameter("@pDtExtr", obj.DataExtrato, dbFieldTypeDateTime))
        
        Call .AddValue("[descricao]", "@pDescr")
        Call .Parameters.add(.CreateParameter("@pDescr", obj.Descricao, dbFieldTypeString))
        
        Call .AddValue("[documento]", "@pDoc")
        Call .Parameters.add(.CreateParameter("@pDoc", obj.Documento, dbFieldTypeString))
        
        Call .AddValue("[valor]", "@pValor")
        Call .Parameters.add(.CreateParameter("@pValor", obj.Valor, dbFieldTypeDouble))
        
        Call .AddValue("[tipo_operacao]", "@pTpOp")
        Call .Parameters.add(.CreateParameter("@pTpOp", obj.TipoOperacao, dbFieldTypeString))
        
        Call .AddValue("[valor_interno]", "@pVlInterno")
        Call .Parameters.add(.CreateParameter("@pVlInterno", obj.ValorInterno, dbFieldTypeDouble))
        
        Call .AddValue("[conciliado]", "@pConciliado")
        Call .Parameters.add(.CreateParameter("@pConciliado", obj.Conciliado, dbFieldTypeBool))
        
        Call .AddValue("[dt_conciliacao]", "@pDtConc")
        Call .Parameters.add(.CreateParameter("@pDtConc", obj.DataConciliacao, dbFieldTypeDateTime))

    End With
End Sub

Public Function ExisteHistorico(ByVal lngBanco As String, ByVal strDescExtrato As String, ByVal strTipoOperacao As String) As Long
    Dim selCmd          As IDBSelectCommand
    Dim rdResult        As IDBReader
               
On Error GoTo Erro
    Aplicacao.Connect
            
    Set selCmd = Aplicacao.CreateSelectCommand
    
    With selCmd
        
        .Table.TableName = "FFIExtratoBancarioHistorico"
        
        Call .Filter.Append("cd_banco = @pBanco")
        Call .Parameters.add(selCmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
        Call .Filter.Append("descricao_extrato = @pDescExt")
        Call .Parameters.add(selCmd.CreateParameter("@pDescExt", strDescExtrato, dbFieldTypeString))
        
        Call .Filter.Append("tipo_operacao = @pTpOperacao")
        Call .Parameters.add(selCmd.CreateParameter("@pTpOperacao", strTipoOperacao, dbFieldTypeString))
        
        Call .Filter.Append("enterprise_id = @pEnterpriseID")
        Call .Parameters.add(.CreateParameter("@pEnterpriseID", ModGeral.EnterpriseId, dbFieldTypeLong))
    
        Call .Filter.Append("cd_estabelecimento = @pCdEst")
        Call .Parameters.add(.CreateParameter("@pCdEst", ModGeral.CdEstabelecimento, dbFieldTypeLong))
    End With
    
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    
    ExisteHistorico = rdResult.GetValue("cd_historico")
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    
    Aplicacao.Disconnect
    Exit Function
Erro:
    ExisteHistorico = 0
    FinallyConnection Aplicacao
End Function


Public Function BuscaDescricaoHistorico(ByVal lngBanco As Long, ByVal lngCdHistorico As Long) As String
    Dim selCmd            As IDBSelectCommand
    Dim rdResult          As IDBReader
    Dim strDescHistorico  As String
           
On Error GoTo Erro
            
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    
    With selCmd
        
        .Table.TableName = NOME_TABELA & "Historico"
        
        Call .Filter.Append("cd_banco = @pBanco")
        Call .Parameters.add(selCmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
                
        Call .Filter.Append("cd_historico = @pCdHist")
        Call .Parameters.add(selCmd.CreateParameter("@pCdHist", lngCdHistorico, dbFieldTypeLong))
                                           
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
        
        If Not rdResult.EOF Then
            strDescHistorico = rdResult.GetValue("descricao_extrato")
        Else
            strDescHistorico = ""
        End If
    End With
    
    BuscaDescricaoHistorico = strDescHistorico
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    Aplicacao.Disconnect
    
    Exit Function
Erro:
    Aplicacao.Disconnect
    BuscaDescricaoHistorico = ""
End Function


Public Function ExcluirRegistros(ByVal lngBanco As Long, ByVal strDataInicial As String, ByVal strDataFinal As String, Optional lngExtrato As Long) As Boolean
    Dim cmd     As IDBDeleteCommand
    
On Error GoTo erro_excluindo
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateDeleteCommand
    With cmd
        .Table = NOME_TABELA
    
        Call .Filter.Append("cd_banco = @pBanco")
        Call .Parameters.add(cmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
        Call .Filter.Append("data_extrato >= @pDataInicial")
        Call .Parameters.add(cmd.CreateParameter("@pDataInicial", Format(strDataInicial, "yyyy/mm/dd") & " 00:00:01", dbFieldTypeDateTime))
        
        Call .Filter.Append("data_extrato <= @pDataFinal")
        Call .Parameters.add(cmd.CreateParameter("@pDataFinal", Format(strDataFinal, "yyyy/mm/dd") & " 23:59:59", dbFieldTypeDateTime))
        
        If lngExtrato > 0 Then
            Call .Filter.Append("cd_extrato = @pExtrato")
            Call .Parameters.add(cmd.CreateParameter("@pExtrato", lngExtrato, dbFieldTypeLong))
        End If
    End With
    Call Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd)
    ExcluirRegistros = True
    Aplicacao.Disconnect
    Set cmd = Nothing
    
    Exit Function
erro_excluindo:
    Set cmd = Nothing
    ExcluirRegistros = False
    Call Throw(err)
End Function
