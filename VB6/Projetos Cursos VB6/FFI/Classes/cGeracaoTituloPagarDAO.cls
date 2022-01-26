VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeracaoTituloPagarDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const NomeTabela$ = "FFITituloPagar"
Private Const nomeTabelaRateio$ = "FFITituloPagarRateio"
Private Const nomeTabelaDuplicata$ = "Duplicatas"
Private Const nomeTabelaFVFDup$ = "FVFTituloPagarDuplicata"

Public Function Carregar(ByVal lngCodigo As Long) As cGeracaoTituloPagar
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Dim GeracaoTituloPagar As New cGeracaoTituloPagar

On Error GoTo erro_carregando
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = NomeTabela
    Call cmd.Filter.Append("cd_titulo = @pNumero")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumero", lngCodigo, dbFieldTypeLong))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set Carregar = GeracaoTituloPagar.CriarObjeto(rdResult)
    End If
    Aplicacao.Disconnect
    Exit Function
erro_carregando:
    Set Carregar = Nothing
End Function

Public Function persistir(ByVal obj As cGeracaoTituloPagar, ByVal objAplicacao As IApplication, ByVal objRateio As cGeracaoTituloPagar, Optional ByVal objDuplicata As cGeracaoDuplicataPagar, Optional blnAtualizar As Boolean) As Boolean
    Dim cmd As IDBInsertCommand
    Dim objGeracaoTituloPagarDAO As New cGeracaoTituloPagarDAO
    
On Error GoTo erro_persistindo
    Set cmd = objAplicacao.CreateInsertCommand
    persistir = True
    cmd.Table = NomeTabela
    Call cmd.AddValue("cd_titulo", "@pcd_titulo")
    If Not blnAtualizar Then
        obj.Cd_Titulo = fProximoNumero
    End If
    Call cmd.Parameters.add(cmd.CreateParameter("@pcd_titulo", obj.Cd_Titulo, dbFieldTypeLong))
    Call PrepareCommand(cmd, obj)
    Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
    
    'Inicio Rateio.
    objRateio.Rateio.MoveFirst
    While Not objRateio.Rateio.EOF
        If objGeracaoTituloPagarDAO.persistirRateio(objRateio.Rateio.CurrentObject, objAplicacao, obj.Cd_Titulo) Then
            objRateio.Rateio.MoveNext
        Else
            GoTo erro_persistindo
        End If
    Wend
    'Final Rateio.
    
    'Inicio Duplicata.
    objDuplicata.parcelas.MoveFirst
    While Not objDuplicata.parcelas.EOF
        If objGeracaoTituloPagarDAO.persistirDuplicata(objDuplicata.parcelas.CurrentObject, objAplicacao, obj.Cd_Titulo) Then
            objDuplicata.parcelas.MoveNext
        Else
            GoTo erro_persistindo
        End If
    Wend
    'Final Duplicata.
        
    Set objGeracaoTituloPagarDAO = Nothing
    Set cmd = Nothing
    Exit Function
erro_persistindo:
    persistir = False
End Function

Private Sub PrepareCommand(ByVal cmd As Object, ByVal obj As cGeracaoTituloPagar)

    With cmd
        Call .AddValue("Descricao", "@pDescricao")
        Call .Parameters.add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString, 50))
        
        Call .AddValue("Numero_nota", "@pNumero_nota")
        Call .Parameters.add(.CreateParameter("@pNumero_nota", obj.Numero_nota, dbFieldTypeLong))
        
        Call .AddValue("Tipo_registro", "@pTipo_registro")
        Call .Parameters.add(.CreateParameter("@pTipo_registro", obj.Tipo_registro, dbFieldTypeString, 20))
        
        Call .AddValue("Empresa", "@pEmpresa")
        Call .Parameters.add(.CreateParameter("@pEmpresa", obj.Empresa, dbFieldTypeString, 15))
        
        Call .AddValue("Vl_valor_nota", "@pVl_valor_nota")
        Call .Parameters.add(.CreateParameter("@pVl_valor_nota", obj.Vl_valor_nota, dbFieldTypeDouble))
        
        Call .AddValue("Intervalo_vencimento", "@pIntervalo_vencimento")
        Call .Parameters.add(.CreateParameter("@pIntervalo_vencimento", obj.Intervalo_vencimento, dbFieldTypeInt))
                                       
        Call .AddValue("Dt_data_emissao", "@pDt_data_emissao")
        Call .Parameters.add(.CreateParameter("@pDt_data_emissao", obj.Dt_data_emissao, dbFieldTypeDate))
        
        Call .AddValue("Cd_banco", "@pCd_banco")
        Call .Parameters.add(.CreateParameter("@pCd_banco", obj.Cd_banco, dbFieldTypeLong))
        
        Call .AddValue("Cd_conta", "@pCd_conta")
        Call .Parameters.add(.CreateParameter("@pCd_conta", obj.Cd_conta, dbFieldTypeLong))
        
        Call .AddValue("Cd_centro_custo", "@pCd_centro_custo")
        Call .Parameters.add(.CreateParameter("@pCd_centro_custo", obj.Cd_centro_custo, dbFieldTypeLong))
        
        Call .AddValue("Cd_moeda", "@pCd_moeda")
        Call .Parameters.add(.CreateParameter("@pCd_moeda", obj.Cd_moeda, dbFieldTypeString, 10))
        
        Call .AddValue("cd_operacao_contabil", "@pcd_operacao_contabil")
        Call .Parameters.add(.CreateParameter("@pcd_operacao_contabil", obj.cd_operacao_contabil, dbFieldTypeLong))
        
        Call .AddValue("nr_parcela", "@pnr_parcela")
        Call .Parameters.add(.CreateParameter("@pnr_parcela", obj.nr_parcela, dbFieldTypeInt))
        
        Call .AddValue("Status", "@pStatus")
        Call .Parameters.add(.CreateParameter("@pStatus", obj.Status, dbFieldTypeString, 1))
    End With
End Sub

Private Sub prepareCommandDuplicata(ByVal cmd As Object, ByVal obj As cGeracaoDuplicataPagar)
            
    With cmd
        Call .AddValue("Descrição", "@pDescricao")
        Call .Parameters.add(.CreateParameter("@pDescricao", obj.P_Descricao, dbFieldTypeString, 50))
        
        Call .AddValue("[Valor Original]", "@pValor_Original")
        Call .Parameters.add(.CreateParameter("@pValor_Original", obj.P_Valor_Original, dbFieldTypeDouble))
        
        Call .AddValue("Banco", "@pCd_banco")
        Call .Parameters.add(.CreateParameter("@pCd_banco", obj.P_Banco, dbFieldTypeLong))
        
        Call .AddValue("Conta", "@pCd_conta")
        Call .Parameters.add(.CreateParameter("@pCd_conta", obj.P_Conta, dbFieldTypeLong))
        
        Call .AddValue("Centro", "@pCd_centro_custo")
        Call .Parameters.add(.CreateParameter("@pCd_centro_custo", obj.P_Centro, dbFieldTypeLong))
        
        Call .AddValue("Moeda", "@pCd_moeda")
        Call .Parameters.add(.CreateParameter("@pCd_moeda", obj.P_Moeda, dbFieldTypeString, 10))
        
        Call .AddValue("cd_operacao_contabil", "@pcd_operacao_contabil")
        Call .Parameters.add(.CreateParameter("@pcd_operacao_contabil", obj.P_cd_operacao_contabil, dbFieldTypeLong))
               
        Call .AddValue("Vencimento", "@pVencimento")
        Call .Parameters.add(.CreateParameter("@pVencimento", obj.P_Vencimento, dbFieldTypeDate))
        
        'pt. 85684 - Ivo Sousa(14/07/2008)
        Call .AddValue("Liberação", "@pLiberação")
        Call .Parameters.add(.CreateParameter("@pLiberação", obj.P_Vencimento, dbFieldTypeDate))
        
        Call .AddValue("Emissão", "@pEmissao")
        Call .Parameters.add(.CreateParameter("@pEmissao", obj.P_Emissao, dbFieldTypeDate))
        
    End With
End Sub

Public Function Excluir(ByVal obj As cGeracaoTituloPagar, ByVal objAplicacao As IApplication) As Boolean
    Dim cmd                As IDBDeleteCommand
    
On Error GoTo erro_excluindo
    Set cmd = objAplicacao.CreateDeleteCommand
    cmd.Table = NomeTabela
    Call cmd.Filter.Append("Cd_titulo = @pCd_titulo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCd_titulo", obj.Cd_Titulo, dbFieldTypeLong))
    Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
    
    Set cmd = objAplicacao.CreateDeleteCommand
    cmd.Table = nomeTabelaRateio
    Call cmd.Filter.Append("Cd_titulo = @pCd_titulo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCd_titulo", obj.Cd_Titulo, dbFieldTypeLong))
    Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
    
    Set cmd = objAplicacao.CreateDeleteCommand
    cmd.Table = nomeTabelaFVFDup
    Call cmd.Filter.Append("Cd_titulo = @pCd_titulo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCd_titulo", obj.Cd_Titulo, dbFieldTypeLong))
    Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
    
    Excluir = True
    Set cmd = Nothing
    Exit Function
erro_excluindo:
    Excluir = False
    Throw err
End Function

Public Function excluirDuplicata(ByVal obj As cGeracaoDuplicataPagar, ByVal objAplicacao As IApplication) As Boolean
    Dim cmd                As IDBDeleteCommand
    
On Error GoTo erro_excluindo
    Set cmd = objAplicacao.CreateDeleteCommand
    cmd.Table = "Duplicatas"
    Call cmd.Filter.Append("PagRec = @pPagRec")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPagRec", obj.P_PagRec, dbFieldTypeString, 1))
    Call cmd.Filter.Append("Nota = @pNota")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNota", obj.P_Nota, dbFieldTypeLong))
    Call cmd.Filter.Append("Empresa = @pEmpresa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", obj.P_Empresa, dbFieldTypeString, 15))
    Call cmd.Filter.Append("Tipo = @pTipo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipo", obj.P_Tipo, dbFieldTypeString, 30))
    Call cmd.Filter.Append("Parcela = @pParcela")
    Call cmd.Parameters.add(cmd.CreateParameter("@pParcela", obj.P_Parcela, dbFieldTypeInt))
    Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
    
    excluirDuplicata = True
    Set cmd = Nothing
    Exit Function
erro_excluindo:
    excluirDuplicata = False
    Throw err
End Function

Public Function existir(ByVal lngCodigo As Long) As Boolean
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = NomeTabela
    Call cmd.Filter.Append("Cd_titulo = @pCd_titulo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCd_titulo", lngCodigo, dbFieldTypeLong))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    existir = Not rdResult.EOF
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

Public Function Atualizar(ByVal obj As cGeracaoTituloPagar, ByVal objAplicacao As IApplication, ByVal objRateio As cGeracaoTituloPagar, ByVal objDuplicata As cGeracaoDuplicataPagar) As Boolean
    Dim objTitPagar As cGeracaoTituloPagar
    Dim booErro As Boolean
On Error GoTo erro_atualizando
    
    Set objTitPagar = Carregar(obj.Cd_Titulo)
        If Excluir(objTitPagar, objAplicacao) Then
            Atualizar = persistir(obj, objAplicacao, objRateio, objDuplicata, True)
        Else
            Atualizar = False
        End If
    Exit Function

erro_atualizando:
    Atualizar = False
End Function

Private Function fProximoNumero() As Long
    Dim strSql As String
    Dim rstTab As Object
    
    strSql = "SELECT max(cd_titulo) as Codigo FROM FFITituloPagar;"
    If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
        fProximoNumero = GetValue(rstTab, "Codigo") + 1
    Else
        fProximoNumero = 0
    End If
    FechaRecordset (rstTab)
End Function


Public Function persistirRateio(ByVal obj As cGeracaoTituloPagar, ByVal objAplicacao As IApplication, Cd_Titulo As Long) As Boolean
    Dim cmd As IDBInsertCommand
   
On Error GoTo erro_persistindo
    Set cmd = objAplicacao.CreateInsertCommand
    persistirRateio = True
    cmd.Table = nomeTabelaRateio
    Call cmd.AddValue("cd_titulo", "@pcd_titulo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pcd_titulo", Cd_Titulo, dbFieldTypeLong))
    Call cmd.AddValue("cd_centro_custo", "@pcd_centro_custo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pcd_centro_custo", obj.R_Cd_centro_custo, dbFieldTypeLong))
    Call cmd.AddValue("cd_conta_financeira", "@pcd_conta_financeira")
    Call cmd.Parameters.add(cmd.CreateParameter("@pcd_conta_financeira", obj.R_Cd_conta, dbFieldTypeLong))
    Call cmd.AddValue("pr_percentual", "@ppr_percentual")
    Call cmd.Parameters.add(cmd.CreateParameter("@ppr_percentual", obj.R_Percentual, dbFieldTypeDouble))
    Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
    
    Set cmd = Nothing
    Exit Function
erro_persistindo:
    persistirRateio = False
End Function

Public Function persistirDuplicata(ByVal obj As cGeracaoDuplicataPagar, ByVal objAplicacao As IApplication, Optional Cod_Titulo As Long) As Boolean
    Dim cmd As IDBInsertCommand
    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
    Dim objLogLancDup       As New clsLogLancamentosDuplicatas
    
On Error GoTo erro_persistindo
    
    If obj.P_Pagamento = "" Then
        excluirDuplicata obj, objAplicacao
        
        Set cmd = objAplicacao.CreateInsertCommand
        persistirDuplicata = True
        cmd.Table = nomeTabelaDuplicata
        Call cmd.AddValue("PagRec", "@pPagRec")
        Call cmd.Parameters.add(cmd.CreateParameter("@pPagRec", obj.P_PagRec, dbFieldTypeString, 1))
        Call cmd.AddValue("Nota", "@pNumero_nota")
        Call cmd.Parameters.add(cmd.CreateParameter("@pNumero_nota", obj.P_Nota, dbFieldTypeLong))
        Call cmd.AddValue("Tipo", "@pTipo_registro")
        Call cmd.Parameters.add(cmd.CreateParameter("@pTipo_registro", obj.P_Tipo, dbFieldTypeString, 20))
        Call cmd.AddValue("Empresa", "@pEmpresa")
        Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", obj.P_Empresa, dbFieldTypeString, 15))
        Call cmd.AddValue("Parcela", "@pParcela")
        Call cmd.Parameters.add(cmd.CreateParameter("@pParcela", obj.P_Parcela, dbFieldTypeInt))
        Call prepareCommandDuplicata(cmd, obj)
        Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
        'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
        With objLogLancDup
            Call .SetKey(obj.P_PagRec, CDbl(obj.P_Nota), obj.P_Empresa, obj.P_Tipo, obj.P_Parcela, Duplicata)
            Call .InsertMsg("Título criado através da rotina de Geração de Títulos a Pagar.")
        End With
    End If
    
        Set cmd = objAplicacao.CreateInsertCommand
        cmd.Table = "FVFTituloPagarDuplicata"
        Call cmd.AddValue("cd_titulo", "@pcd_titulo")
        Call cmd.Parameters.add(cmd.CreateParameter("@pcd_titulo", Cod_Titulo, dbFieldTypeLong))
        Call cmd.AddValue("Nota", "@pNumero_nota")
        Call cmd.Parameters.add(cmd.CreateParameter("@pNumero_nota", obj.P_Nota, dbFieldTypeLong))
        Call cmd.AddValue("tipo_registro", "@pTipo_registro")
        Call cmd.Parameters.add(cmd.CreateParameter("@pTipo_registro", obj.P_Tipo, dbFieldTypeString, 20))
        Call cmd.AddValue("Empresa", "@pEmpresa")
        Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", obj.P_Empresa, dbFieldTypeString, 15))
        Call cmd.AddValue("Parcela", "@pParcela")
        Call cmd.Parameters.add(cmd.CreateParameter("@pParcela", obj.P_Parcela, dbFieldTypeInt))
        Call cmd.AddValue("PagRec", "@pPagRec")
        Call cmd.Parameters.add(cmd.CreateParameter("@pPagRec", obj.P_PagRec, dbFieldTypeString, 1))
        Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
        
        Set cmd = Nothing
    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
    Set objLogLancDup = Nothing
    
    persistirDuplicata = True
    Exit Function
erro_persistindo:
    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
    Set objLogLancDup = Nothing
    
    persistirDuplicata = False
End Function
