VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeracaoTituloPagarDuplicataDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const nomeTabela$ = "FVFTituloPagarDuplicata"

Public Function carregar(ByVal lngCodigo As Integer) As cGeracaoTituloPagar
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Dim GeracaoTituloPagar As New cGeracaoTituloPagar

On Error GoTo erro_carregando
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = nomeTabela
    Call cmd.Filter.Append("cd_titulo = @pNumero")
    Call cmd.Parameters.Add(cmd.CreateParameter("@pNumero", lngCodigo, dbFieldTypeLong))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set carregar = GeracaoTituloPagar.CriarObjeto(rdResult)
    End If
    Aplicacao.Disconnect
    Exit Function
erro_carregando:
    Set carregar = Nothing
End Function

Public Function persistir(ByVal obj As cGeracaoTituloPagar, ByVal objAplicacao As IApplication, Optional blnAtualizar As Boolean) As Boolean
    Dim cmd As IDBInsertCommand
    Dim objGeracaoTituloPagarDAO As New cGeracaoTituloPagarDAO
    
On Error GoTo erro_persistindo
    Set cmd = objAplicacao.CreateInsertCommand
    persistir = True
    cmd.Table = nomeTabela
    Call cmd.AddValue("cd_titulo", "@pcd_titulo")
    If blnAtualizar Then
        Call cmd.Parameters.Add(cmd.CreateParameter("@pcd_titulo", obj.Cd_titulo, dbFieldTypeInt))
    Else
        Call cmd.Parameters.Add(cmd.CreateParameter("@pcd_titulo", fProximoNumero, dbFieldTypeInt))
    End If
    Call prepareCommand(cmd, obj)
    Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
    
    Set objGeracaoTituloPagarDAO = Nothing
    Set cmd = Nothing
    Exit Function
erro_persistindo:
    persistir = False
    Throw err
End Function


Private Sub prepareCommand(ByVal cmd As Object, ByVal obj As cGeracaoTituloPagar)

    With cmd
        Call .AddValue("Descricao", "@pDescricao")
        Call .Parameters.Add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString, 50))
        
        Call .AddValue("Numero_nota", "@pNumero_nota")
        Call .Parameters.Add(.CreateParameter("@pNumero_nota", obj.Numero_nota, dbFieldTypeInt))
        
        Call .AddValue("Tipo_registro", "@pTipo_registro")
        Call .Parameters.Add(.CreateParameter("@pTipo_registro", obj.Tipo_registro, dbFieldTypeString, 20))
        
        Call .AddValue("Empresa", "@pEmpresa")
        Call .Parameters.Add(.CreateParameter("@pEmpresa", obj.Empresa, dbFieldTypeString, 15))
        
        Call .AddValue("Vl_valor_nota", "@pVl_valor_nota")
        Call .Parameters.Add(.CreateParameter("@pVl_valor_nota", obj.Vl_valor_nota, dbFieldTypeDouble))
        
        Call .AddValue("Intervalo_vencimento", "@pIntervalo_vencimento")
        Call .Parameters.Add(.CreateParameter("@pIntervalo_vencimento", obj.Intervalo_vencimento, dbFieldTypeInt))
                                       
        Call .AddValue("Dt_data_emissao", "@pDt_data_emissao")
        Call .Parameters.Add(.CreateParameter("@pDt_data_emissao", obj.Dt_data_emissao, dbFieldTypeDate))
        
        Call .AddValue("Cd_banco", "@pCd_banco")
        Call .Parameters.Add(.CreateParameter("@pCd_banco", obj.Cd_banco, dbFieldTypeInt))
        
        Call .AddValue("Cd_conta", "@pCd_conta")
        Call .Parameters.Add(.CreateParameter("@pCd_conta", obj.Cd_conta, dbFieldTypeInt))
        
        Call .AddValue("Cd_centro_custo", "@pCd_centro_custo")
        Call .Parameters.Add(.CreateParameter("@pCd_centro_custo", obj.Cd_centro_custo, dbFieldTypeInt))
        
        Call .AddValue("Cd_moeda", "@pCd_moeda")
        Call .Parameters.Add(.CreateParameter("@pCd_moeda", obj.Cd_moeda, dbFieldTypeString, 10))
        
        Call .AddValue("cd_operacao_contabil", "@pcd_operacao_contabil")
        Call .Parameters.Add(.CreateParameter("@pcd_operacao_contabil", obj.cd_operacao_contabil, dbFieldTypeInt))
        
        Call .AddValue("nr_parcela", "@pnr_parcela")
        Call .Parameters.Add(.CreateParameter("@pnr_parcela", obj.nr_parcela, dbFieldTypeInt))
    End With
End Sub

Public Function excluir(ByVal obj As cGeracaoTituloPagar, ByVal objAplicacao As IApplication) As Boolean
    Dim cmd                As IDBDeleteCommand
    'Dim objNotasFiscaisDAO As New cFretePagarNotaFiscalDAO
    'Dim objTituloDAO       As New cTituloDAO
 
    
On Error GoTo erro_excluindo
    Set cmd = objAplicacao.CreateDeleteCommand
    cmd.Table = nomeTabela
    Call cmd.Filter.Append("Cd_titulo = @pCd_titulo")
    'Call cmd.Filter.Append("tp_registro = @pTipoRegistro")
    'Call cmd.Filter.Append("cd_transportadora = @pTransportadora")
    Call cmd.Parameters.Add(cmd.CreateParameter("@pCd_titulo", obj.Cd_titulo, dbFieldTypeLong))
    'Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", obj.TipoRegistro, dbFieldTypeString, 20))
    'Call cmd.Parameters.add(cmd.CreateParameter("@pTransportadora", obj.codigoTransportadora, dbFieldTypeLong))
'    objTituloDAO.numeroConhecimento = obj.numeroConhecimento
'    objTituloDAO.tipoConhecimento = obj.TipoRegistro
'    objTituloDAO.PagRec = "P"
'    objTituloDAO.Tabela = "FreteEntradaDuplicata"
'    objTituloDAO.Transportadora = obj.codigoTransportadora
'    Set objTituloDAO = Nothing
    Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
    excluir = True
    Set cmd = Nothing
    Exit Function
erro_excluindo:
    excluir = False
    Throw err
End Function

Public Function existir(ByVal lngCodigo As Long) As Boolean
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = nomeTabela
    Call cmd.Filter.Append("Cd_titulo = @pCd_titulo")
    Call cmd.Parameters.Add(cmd.CreateParameter("@pCd_titulo", lngCodigo, dbFieldTypeLong))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    existir = Not rdResult.EOF
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

Public Function Atualizar(ByVal obj As cGeracaoTituloPagar, ByVal objAplicacao As IApplication) As Boolean
    Dim objTitPagar As cGeracaoTituloPagar
    Dim booErro As Boolean
On Error GoTo erro_atualizando
    
    Set objTitPagar = carregar(obj.Cd_titulo)
'    If objTitPagar.PermiteAlteracao Then
        If excluir(objTitPagar, objAplicacao) Then
            Atualizar = persistir(obj, objAplicacao, True)
        Else
            Atualizar = False
        End If
'    Else
'        Dim daoTitulo As New cTituloDAO
'
'        daoTitulo.numeroConhecimento = obj.numeroConhecimento
'        daoTitulo.tipoConhecimento = obj.TipoRegistro
'        daoTitulo.PagRec = "P"
'        'pt. 82395 - Dulcino Júnior
'        daoTitulo.OperacaoContabil = obj.OperacaoContabil
'        'pt. 86013 - Ivo Sousa(20/05/2008)
'        daoTitulo.Transportadora = obj.codigoTransportadora
'        Atualizar = daoTitulo.alterar(obj.Titulo, objAplicacao)
'        Set daoTitulo = Nothing
'    End If
    Exit Function

erro_atualizando:
    Atualizar = False
    Call Throw(err) 'repasso o erro
End Function

