VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DaoCadHistBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 61827 - Demanda: 61828
'---|   Ueder Budni (13/01/2015)
'---|---------------------------------------------------------------------------------------------------------------------------

Private mApp            As IApplication

Private Const NOME_TABELA = "FFIExtratoBancarioHistorico"

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

Public Function CarregaColecao(ByVal lngBanco As Long) As ColCadHistBancario
    Dim objHistBanc     As VoCadHistBancario
    Dim selCmd          As IDBSelectCommand
    Dim rdResult        As IDBReader
    Dim colTmp          As ColCadHistBancario
           
On Error GoTo Erro
    Aplicacao.Connect
            
    Set selCmd = Aplicacao.CreateSelectCommand
    Set colTmp = New ColCadHistBancario
    
    With selCmd
        
        .Table.TableName = NOME_TABELA
        
        Call .Filter.Append("cd_banco = @pBanco")
        Call .Parameters.add(selCmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
        .OrderByClause = "cd_historico"
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
        While Not rdResult.EOF
            Set objHistBanc = New VoCadHistBancario
            Set objHistBanc = CarregaVO(rdResult)
            Call colTmp.add(objHistBanc)
            Set objHistBanc = Nothing
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

Private Function CarregaVO(ByVal rs As IDBReader) As VoCadHistBancario
    Dim objTmp As New VoCadHistBancario

On Error GoTo Erro:
    If Not rs Is Nothing Then
        With objTmp
            .EnterpriseId = rs.GetLong("Enterprise_id")
            .CdEstabelecimento = rs.GetLong("Cd_estabelecimento")
            .CdBanco = rs.GetLong("Cd_banco")
            .CdHistorico = rs.GetLong("cd_historico")
            .DescricaoExtrato = rs.GetString("descricao_extrato")
            .ComplementoDescricao = rs.GetString("complemento_descricao")
            .TipoOperacao = rs.GetString("tipo_operacao")
        End With
    End If
    Set CarregaVO = objTmp
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi possível carregar o objeto! (DaoCadHistBancario.CarregaVO)"
    
End Function

Public Function GravaColecao(ByVal col As ColCadHistBancario, lngBanco As Long)
    Dim obj         As VoCadHistBancario
    Dim Gravando    As Boolean
    
    Gravando = True
On Error GoTo err
    Aplicacao.Connect
    Set mApp = Aplicacao
    mApp.BeginTransaction
    If col.Count > 0 Then
        col.MoveFirst
        Gravando = ExcluirTodos(col.CurrentObject.CdBanco)
        While Not col.EOF And Gravando
            Set obj = col.CurrentObject
            Gravando = Gravando And GravarVO(obj)
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


Public Function ExcluirTodos(ByVal lngBanco As Long) As Boolean
    Dim cmd     As IDBDeleteCommand
    
On Error GoTo erro_excluindo
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateDeleteCommand
    With cmd
        .Table = NOME_TABELA
    
        Call .Filter.Append("cd_banco = @pBanco")
        Call .Parameters.add(cmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
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


Public Function GravarVO(ByVal obj As VoCadHistBancario) As Boolean
    Dim cmd        As IDBInsertCommand
    
On Error GoTo erro_gravando
    
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = NOME_TABELA
        
        Call .AddValue("[enterprise_id]", "@pEnterpriseID")
        Call .Parameters.add(.CreateParameter("@pEnterpriseID", obj.EnterpriseId, dbFieldTypeLong))
    
        Call .AddValue("[cd_estabelecimento]", "@pCdEst")
        Call .Parameters.add(.CreateParameter("@pCdEst", obj.CdEstabelecimento, dbFieldTypeLong))
        
        Call .AddValue("[cd_banco]", "@pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", obj.CdBanco, dbFieldTypeLong))
        
        Call .AddValue("[cd_historico]", "@pCdHist")
        Call .Parameters.add(.CreateParameter("@pCdHist", obj.CdHistorico, dbFieldTypeLong))
        
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

Private Sub PrepareCommand(ByVal cmd As Object, ByVal obj As VoCadHistBancario)
    With cmd
        Call .AddValue("[descricao_extrato]", "@pDescrExtr")
        Call .Parameters.add(.CreateParameter("@pDescrExtr", obj.DescricaoExtrato, dbFieldTypeString))
        
        Call .AddValue("[complemento_descricao]", "@pComplDescr")
        Call .Parameters.add(.CreateParameter("@pComplDescr", obj.ComplementoDescricao, dbFieldTypeString))
               
        Call .AddValue("[tipo_operacao]", "@pTpOp")
        Call .Parameters.add(.CreateParameter("@pTpOp", obj.TipoOperacao, dbFieldTypeString))
    End With
End Sub

Public Function NovoCodigo(lngBanco As Long) As Long
    Dim selCmd          As IDBSelectCommand
    Dim rdResult        As IDBReader
On Error GoTo Erro
    Aplicacao.Connect
            
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        
        .Table.TableName = NOME_TABELA
        .SelectClause = "MAX(cd_historico) + 1 as NovoCodigo"
        
        Call .Filter.Append("cd_banco = @pBanco")
        Call .Parameters.add(selCmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
           
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
        If Not rdResult.EOF Then
            NovoCodigo = IIf(rdResult.GetLong("NovoCodigo") = 0, 1, rdResult.GetLong("NovoCodigo"))
        End If
        
    End With
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    
    Aplicacao.Disconnect
    Exit Function
Erro:
    Aplicacao.Disconnect
    NovoCodigo = 0
End Function

Public Function CarregaVOpelaPK(lngEnterpriseId As Long, lngCdEstab As Long, lngBanco As Long, lngCodHist As Long, strTipoOperacao As String, Optional blnFiltraAmbos As Boolean) As VoCadHistBancario
    Dim objHistBanc     As VoCadHistBancario
    Dim selCmd          As IDBSelectCommand
    Dim rdResult        As IDBReader
           
On Error GoTo Erro
    Aplicacao.Connect
            
    Set selCmd = Aplicacao.CreateSelectCommand
    
    With selCmd
        
        .Table.TableName = NOME_TABELA
        Call .Filter.Append("enterprise_id = @pEntId")
        Call .Parameters.add(selCmd.CreateParameter("@pEntId", lngEnterpriseId, dbFieldTypeLong))
        
        Call .Filter.Append("cd_estabelecimento = @pCdEstab")
        Call .Parameters.add(selCmd.CreateParameter("@pCdEstab", lngCdEstab, dbFieldTypeLong))
        
        Call .Filter.Append("cd_banco = @pBanco")
        Call .Parameters.add(selCmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
        Call .Filter.Append("cd_historico = @pCdHist")
        Call .Parameters.add(selCmd.CreateParameter("@pCdHist", lngCodHist, dbFieldTypeLong))
        
        Call .Filter.Append(IIf(blnFiltraAmbos, "(", "") & "tipo_operacao = @pTpOperacao")
        Call .Parameters.add(selCmd.CreateParameter("@pTpOperacao", Trim(strTipoOperacao), dbFieldTypeString))
        
        If blnFiltraAmbos Then
            Call .Filter.Append("tipo_operacao = @pAmbos)", dbLogicOperatorOR)
            Call .Parameters.add(selCmd.CreateParameter("@pAmbos", "A", dbFieldTypeString))
        End If
        
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
        If Not rdResult.EOF Then
            Set CarregaVOpelaPK = CarregaVO(rdResult)
        End If
    End With
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    
    Aplicacao.Disconnect
    Exit Function
Erro:
    FinallyConnection Aplicacao
End Function
