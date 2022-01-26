VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCondicaoPagamentoDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mobjMatrizParcelaDAO As cMatrizParcelaDAO
Const NomeTabela As String = "Condições de Pagamento"

Private Sub Class_Initialize()
    Set mobjMatrizParcelaDAO = New cMatrizParcelaDAO
End Sub

Public Function Gravar(ByVal objCondicaoPagamento As cCondicaoPagamento, ByRef objAplicacao As IApplication, ByRef strErro As String, blnAlteracao As Boolean, Optional lngCodigo As Long = 0) As Boolean
    Dim cmd As Object
    
On Error GoTo erro_gravando
    objAplicacao.Connect
    objAplicacao.BeginTransaction
    With objCondicaoPagamento
        If blnAlteracao Then
            Set cmd = objAplicacao.CreateUpdateCommand
            Call cmd.Filter.Append("Código = @pCodigo")
        Else
            Set cmd = objAplicacao.CreateInsertCommand
            Call cmd.AddValue("Código", "@pCodigo")
        End If
        cmd.Table = "[" & NomeTabela & "]"
        Call cmd.Parameters.add(cmd.CreateParameter("@pCodigo", .Codigo, dbFieldTypeLong))
        'Monta restante da SQL de gravação
        Call prepareCommand(cmd, objCondicaoPagamento)
        'Grava a Condição
        Call objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd)
        'Exclui as parcelas caso as mesmas existam
        If mobjMatrizParcelaDAO.Existe(lngCodigo) Then
            Call mobjMatrizParcelaDAO.Excluir(lngCodigo)
        End If
        'Grava as Parcelas
        If mobjMatrizParcelaDAO.Gravar(objCondicaoPagamento.MatrizParcelas, lngCodigo) Then
            objAplicacao.CommitTransaction
        Else
            objAplicacao.RollbackTransaction
        End If
    End With
    Gravar = True
    objAplicacao.Disconnect
    Exit Function
erro_gravando:
    strErro = "Erro ao gravar a Condição de Pagamento: " & Err.Description
    objAplicacao.RollbackTransaction
    objAplicacao.Disconnect
    Gravar = False
End Function

Public Function Excluir(ByVal objCondicaoPagamento As cCondicaoPagamento, ByRef objAplicacao As IApplication, ByRef strErro As String, Optional lngCodigo As Long = 0) As Boolean
    Dim cmd        As IDBDeleteCommand
    Dim blnExcluiu As Boolean
    
On Error GoTo Erro_Excluir
    objAplicacao.Connect
    objAplicacao.BeginTransaction
    With objCondicaoPagamento
        Set cmd = objAplicacao.CreateDeleteCommand
        cmd.Table = "[" & NomeTabela & "]"
        Call cmd.Filter.Append("Código = @pCodigo")
        Call cmd.Parameters.add(cmd.CreateParameter("@pCodigo", .Codigo, dbFieldTypeLong))
        blnExcluiu = (objAplicacao.ExecuteUpdate(objAplicacao.GetInternalAuthorization, cmd) > 0)
        'Exclui as parcelas caso as mesmas existam
        If mobjMatrizParcelaDAO.Existe(lngCodigo) Then
            blnExcluiu = mobjMatrizParcelaDAO.Excluir(lngCodigo)
        End If
        If blnExcluiu Then
            objAplicacao.CommitTransaction
        Else
            objAplicacao.RollbackTransaction
        End If
    End With
    Excluir = blnExcluiu
    objAplicacao.Disconnect
    Exit Function
Erro_Excluir:
    strErro = "Erro ao excluir a Condição de Pagamento: " & Err.Description
    objAplicacao.RollbackTransaction
    objAplicacao.Disconnect
    Excluir = False
End Function

Private Sub prepareCommand(ByRef cmd As Object, ByVal objCondicaoPagamento As cCondicaoPagamento)
    
    With objCondicaoPagamento
        'Descrição
        Call cmd.AddValue("[Descrição]", "@pDescricao")
        Call cmd.Parameters.add(cmd.CreateParameter("@pDescricao", .Descricao, dbFieldTypeString, 120))
        'Despesas Financeras
        Call cmd.AddValue("[Despesa Financeira]", "@pDespesaFinanceira")
        Call cmd.Parameters.add(cmd.CreateParameter("@pDespesaFinanceira", .despesaFinanceira, dbFieldTypeString))
        'Dias de referencia
        Call cmd.AddValue("[DiasReferencia]", "@pDiasReferencia")
        Call cmd.Parameters.add(cmd.CreateParameter("@pDiasReferencia", .diasReferencia, dbFieldTypeInt))
        'Valor maximo
        Call cmd.AddValue("[Máximo]", "@pMaximo")
        Call cmd.Parameters.add(cmd.CreateParameter("@pMaximo", .Maximo, dbFieldTypeDecimal))
        'Valor minimo
        Call cmd.AddValue("[Mínimo]", "@pMinimo")
        Call cmd.Parameters.add(cmd.CreateParameter("@pMinimo", .Minimo, dbFieldTypeDecimal))
        'Distribuição do IPI
        Call cmd.AddValue("[IPI]", "@pIPI")
        Call cmd.Parameters.add(cmd.CreateParameter("@pIPI", .ipi, dbFieldTypeString))
        'Numero de Parcelas
        Call cmd.AddValue("[Número de Parcelas]", "@pNumeroParcelas")
        Call cmd.Parameters.add(cmd.CreateParameter("@pNumeroParcelas", .numeroParcelas, dbFieldTypeByte))
        'Taxa
        Call cmd.AddValue("[Taxa]", "@pTaxa")
        Call cmd.Parameters.add(cmd.CreateParameter("@pTaxa", .taxa, dbFieldTypeSingle))
        'Tipo de Data Inicial
        Call cmd.AddValue("[Tipo da Data Inicial]", "@pTipoDataInicial")
        Call cmd.Parameters.add(cmd.CreateParameter("@pTipoDataInicial", .tipoDataInicial, dbFieldTypeString))
        'Tipo de Dia
        Call cmd.AddValue("[Tipo de Dia]", "@pTipoDia")
        Call cmd.Parameters.add(cmd.CreateParameter("@pTipoDia", .TipoDia, dbFieldTypeString))
        'Venda a Vista
        Call cmd.AddValue("[VendaAVista]", "@pVendaAVista")
        Call cmd.Parameters.add(cmd.CreateParameter("@pVendaAVista", .VendaaVista, dbFieldTypeBool))
        'Apenas dias Úteis
        Call cmd.AddValue("[Considerar apenas dias úteis]", "@pApenasDiasUteis")
        Call cmd.Parameters.add(cmd.CreateParameter("@pApenasDiasUteis", .ApenasDiasUteis, dbFieldTypeBool))
        'Iguais
        Call cmd.AddValue("[Iguais]", "@pIguais")
        Call cmd.Parameters.add(cmd.CreateParameter("@pIguais", .iguais, dbFieldTypeBool))
        'Inativa
        Call cmd.AddValue("[Inativa]", "@pInativa")
        Call cmd.Parameters.add(cmd.CreateParameter("@pInativa", .inativa, dbFieldTypeBool))
        'Mes Atual
        Call cmd.AddValue("[Mês Atual]", "@pMesAtual")
        Call cmd.Parameters.add(cmd.CreateParameter("@pMesAtual", .mesAtual, dbFieldTypeBool))
        'Percentual de comissão
        Call cmd.AddValue("[PerCom]", "@pPerCom")
        Call cmd.Parameters.add(cmd.CreateParameter("@pPerCom", .percentualComissao, dbFieldTypeDouble))
        'Ditribuição do ICMSST
        Call cmd.AddValue("[gera_valor_icms_st]", "@pIcmsSt")
        Call cmd.Parameters.add(cmd.CreateParameter("@pIcmsSt", .IcmsSt, dbFieldTypeString))
        'Dia Definido
        Call cmd.AddValue("[diaDefinido]", "@pDiaDefinido")
        Call cmd.Parameters.add(cmd.CreateParameter("@pDiaDefinido", .DiaDefinido, dbFieldTypeString))
        'Taxa de Juros Mensal
        Call cmd.AddValue("[tx_juro_mensal]", "@pTxJuroMensal")
        Call cmd.Parameters.add(cmd.CreateParameter("@pTxJuroMensal", .JurosMensal, dbFieldTypeDouble))
        'pt. 00000 - Ivo Sousa (20/01/2009)
        'Condição de Pagamento para NFe
        Call cmd.AddValue("[cond_pagt_nfe]", "@pCondPagtNfe")
        Call cmd.Parameters.add(cmd.CreateParameter("@pCondPagtNfe", .CondicaoPagamentoNFe, dbFieldTypeInt))
        'Pt. 92271 - Moacir Pfau(15/05/2009)
        'Tipo de operação
        Call cmd.AddValue("[tipo_operacao]", "@pTipoOperacao")
        Call cmd.Parameters.add(cmd.CreateParameter("@pTipoOperacao", .TipoOperacao, dbFieldTypeString))
        'Pt. 88817 - Moacir Pfau(05/11/2009)
        'Forma Pagamento
        Call cmd.AddValue("[codigo_formapagamento]", "@pFormapagamento")
        Call cmd.Parameters.add(cmd.CreateParameter("@pFormapagamento", .Codigo_formapagamento, dbFieldTypeLong))
        
        'Projeto: #35764 - História: #27764 - Desenvolvimento: #39950 - Ivo Sousa(24/06/2014)
        Call cmd.AddValue("[integracao_portal]", "@pIntegracaoPortal")
        Call cmd.Parameters.add(cmd.CreateParameter("@pIntegracaoPortal", 0, dbFieldTypeInt))
        
        Call cmd.AddValue("[rejeicao_integracao_portal]", "@pRejIntPortal")
        Call cmd.Parameters.add(cmd.CreateParameter("@pRejIntPortal", Empty, dbFieldTypeString))
        
        'Ivo Sousa(16/05/2016) - Projeto: #93871 | Demanda: #120151 | História: #120198
        Call cmd.AddValue("[integrou_pdv]", "@pIntegrouPDV")
        Call cmd.Parameters.add(cmd.CreateParameter("@pIntegrouPDV", 0, dbFieldTypeInt))
        
        'Luiz Satto(18/05/2017) - Projeto: #167145
        Call cmd.AddValue("[integrou_mobile]", "@pIntegrouMobile")
        Call cmd.Parameters.add(cmd.CreateParameter("@pIntegrouMobile", 2, dbFieldTypeInt))
    End With
End Sub
