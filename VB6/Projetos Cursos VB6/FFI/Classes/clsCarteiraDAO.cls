VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarteiraDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mApp                            As IApplication
Private Const mstrTabela = "[FFICarteira]"

Public Sub init(ByVal objApp As IApplication)
    Set mApp = objApp
End Sub

Public Function Gravar(ByVal obj As clsCarteira) As Boolean
    Dim cmd                 As IDBInsertCommand
    Dim blnGravar           As Boolean
    Dim daoCpEspecial       As New clsCarteiraCpEspecialDAO

On Error GoTo erro_gravando
    obj.Id_carteira = fProximoNumero(obj)
    'Cria o camando apra inserir
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = mstrTabela
        Call .AddValue("[Enterprise_id]", "@pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", obj.Enterprise_id, dbFieldTypeInt))
                    
        Call .AddValue("[Cd_estabelecimento]", "@pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", obj.Cd_estabelecimento, dbFieldTypeInt))
        
        Call .AddValue("[id_carteira]", "@pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", fProximoNumero(obj), dbFieldTypeLong))
                    
        Call PrepareCommand(cmd, obj)
    End With
    
    blnGravar = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    If blnGravar Then
        Call daoCpEspecial.init(mApp)
        blnGravar = daoCpEspecial.Gravar_Atualizar(obj)
    End If
    Gravar = blnGravar
    Set cmd = Nothing
    Set daoCpEspecial = Nothing
    Exit Function

erro_gravando:
    Set cmd = Nothing
    Set daoCpEspecial = Nothing
    Gravar = False
End Function

Public Function Atualizar(ByVal obj As clsCarteira) As Boolean
    Dim cmd                 As IDBUpdateCommand
    Dim blnAtualizado       As Boolean
    Dim daoCpEspecial       As New clsCarteiraCpEspecialDAO
    
On Error GoTo erro_atualizado
    Set cmd = mApp.CreateUpdateCommand
     With cmd
        .Table = mstrTabela
            
        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", obj.Enterprise_id, dbFieldTypeInt))
                    
        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", obj.Cd_estabelecimento, dbFieldTypeInt))
        
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", obj.Id_carteira, dbFieldTypeLong))
        
        Call PrepareCommand(cmd, obj)
    End With
    
    blnAtualizado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    If blnAtualizado Then
        Call daoCpEspecial.init(mApp)
        blnAtualizado = daoCpEspecial.Gravar_Atualizar(obj)
    End If
   
    Set cmd = Nothing
    Atualizar = blnAtualizado
    Exit Function
erro_atualizado:
    Set cmd = Nothing
End Function

Public Function Carregar(ByVal lngEnterprise_id As Integer, ByVal lngCd_estabelecimento As Integer, ByVal lngId_carteira As Long) As clsCarteira
    Dim cmd                     As IDBSelectCommand
    Dim rdResult                As IDBReader
    Dim result                  As clsCarteira
    Dim objCpEmpeciaisDAO       As New clsCarteiraCpEspecialDAO

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = mstrTabela
        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", lngEnterprise_id, dbFieldTypeLong))
                    
        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", lngCd_estabelecimento, dbFieldTypeLong))
        
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngId_carteira, dbFieldTypeLong))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set result = New clsCarteira
        Set Carregar = result.CriarObjeto(rdResult)
        Call objCpEmpeciaisDAO.init(mApp)
        Carregar.ColCpEspeciais = objCpEmpeciaisDAO.Carregar(lngEnterprise_id, lngCd_estabelecimento, lngId_carteira)
        Set result = Nothing
    Else
        Set Carregar = Nothing
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function

erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    Set Carregar = Nothing
End Function

Public Function Excluir(ByVal obj As clsCarteira) As Boolean
    Dim cmd                             As IDBDeleteCommand
    Dim blnExcluido                     As Boolean
    
On Error GoTo erro_excluindo
    
    blnExcluido = True
    Set cmd = mApp.CreateDeleteCommand
    With cmd
        .Table = mstrTabela
        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", obj.Enterprise_id, dbFieldTypeInt))
                    
        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", obj.Cd_estabelecimento, dbFieldTypeInt))
        
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", obj.Id_carteira, dbFieldTypeLong))
    End With
    
    blnExcluido = CBool(mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = CLng(1))
    
    Excluir = blnExcluido
    
    Exit Function
erro_excluindo:
    Excluir = False
End Function

Public Function Existe(ByVal obj As clsCarteira) As Boolean
    Dim selCmd                      As IDBSelectCommand
    Dim rdResult                    As IDBReader

    Set selCmd = mApp.CreateSelectCommand
    With selCmd
        .Table.TableName = mstrTabela
        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", obj.Enterprise_id, dbFieldTypeInt))
                    
        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", obj.Cd_estabelecimento, dbFieldTypeInt))
        
        Call .Filter.Append("[Id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", obj.Id_carteira, dbFieldTypeLong))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, selCmd)
    If rdResult.Recordcount > 0 Then
        rdResult.MoveFirst
    End If
    Existe = Not rdResult.EOF
    Set selCmd = Nothing
    rdResult.CloseReader
    Set rdResult = Nothing
End Function

Private Sub PrepareCommand(ByRef cmd As Object, ByVal obj As clsCarteira)
    With cmd
        Call .AddValue("desc_carteira", "@pDesc_carteira")
        Call .Parameters.add(.CreateParameter("@pDesc_carteira", obj.Desc_carteira, dbFieldTypeString, 70))
    
        Call .AddValue("codigo_cedente", "@pCodigo_cedente")
        Call .Parameters.add(.CreateParameter("@pCodigo_cedente", obj.Codigo_cedente, dbFieldTypeString, 70))

        Call .AddValue("inicio_nosso_numero", "@pInicio_nosso_numero")
        Call .Parameters.add(.CreateParameter("@pInicio_nosso_numero", obj.Inicio_nosso_numero, dbFieldTypeString, 30))
    
        Call .AddValue("fim_nosso_numero", "@pFim_nosso_numero")
        Call .Parameters.add(.CreateParameter("@pFim_nosso_numero", obj.Fim_nosso_numero, dbFieldTypeString, 30))
    
        If obj.Proximo_nosso_numero <> "" Then
            Call .AddValue("proximo_nosso_numero", "@pProximo_nosso_numero")
            Call .Parameters.add(.CreateParameter("@pProximo_nosso_numero", obj.Proximo_nosso_numero, dbFieldTypeString, 30))
        End If
    
        Call .AddValue("demonstrativo", "@pDemonstrativo")
        Call .Parameters.add(.CreateParameter("@pDemonstrativo", obj.Demonstrativo, dbFieldTypeString, 255))
        
        Call .AddValue("instrucoesCaixa", "@pInstrucoes_caixa")
        Call .Parameters.add(.CreateParameter("@pInstrucoes_caixa", obj.Instrucoes_caixa, dbFieldTypeString, 255))
        
        Call .AddValue("tipo_layout_boleto", "@pTipo_layout_boleto")
        Call .Parameters.add(.CreateParameter("@pTipo_layout_boleto", obj.Tipo_layout_boleto, dbFieldTypeString, 30))
        
        Call .AddValue("tipo_layout_remessa", "@pTipo_layout_remessa")
        Call .Parameters.add(.CreateParameter("@pTipo_layout_remessa", obj.Tipo_layout_remessa, dbFieldTypeString, 30))
        
        Call .AddValue("tipo_layout_retorno", "@pTipo_layout_retorno")
        Call .Parameters.add(.CreateParameter("@pTipo_layout_retorno", obj.Tipo_layout_retorno, dbFieldTypeString, 30))
        
        Call .AddValue("arquivo_licenca", "@pArquivo_licenca")
        Call .Parameters.add(.CreateParameter("@pArquivo_licenca", obj.Arquivo_licenca, dbFieldTypeString, 250))
        
        Call .AddValue("logo_empresa", "@pLogo_empresa")
        Call .Parameters.add(.CreateParameter("@pLogo_empresa", obj.Logo_empresa, dbFieldTypeString, 250))
        
        Call .AddValue("caminho_arquivo_remessa_padrao", "@pCaminho_arquivo_remessa_padrao")
        Call .Parameters.add(.CreateParameter("@pCaminho_arquivo_remessa_padrao", obj.Caminho_arquivo_remessa_padrao, dbFieldTypeString, 250))

        Call .AddValue("caminho_arquivo_retorno_padrao", "@pCaminho_arquivo_retorno_padrao")
        Call .Parameters.add(.CreateParameter("@pCaminho_arquivo_retorno_padrao", obj.Caminho_arquivo_retorno_padrao, dbFieldTypeString, 250))

        Call .AddValue("margem_superior", "@pMargem_superior_boleto")
        Call .Parameters.add(.CreateParameter("@pMargem_superior_boleto", obj.Margem_superior_boleto, dbFieldTypeInt))

        Call .AddValue("OutroDadoConfiguracao1", "@pOutro_dado_configuracao1")
        Call .Parameters.add(.CreateParameter("@pOutro_dado_configuracao1", obj.Outro_dado_configuracao1, dbFieldTypeString, 50))

        Call .AddValue("OutroDadoConfiguracao2", "@pOutro_dado_configuracao2")
        Call .Parameters.add(.CreateParameter("@pOutro_dado_configuracao2", obj.Outro_dado_configuracao2, dbFieldTypeString, 50))

        Call .AddValue("sequencial_remessa", "@pSequencial_remessa")
        Call .Parameters.add(.CreateParameter("@pSequencial_remessa", obj.Sequencial_remessa, dbFieldTypeLong))
        
        'Pt. 96180 - Moacir Pfau(08/12/2009)
        Call .AddValue("Especie", "@pEspecie")
        Call .Parameters.add(.CreateParameter("@pEspecie", obj.Especie, dbFieldTypeString, 10))
        
        'Pt. 96589 - Moacir Pfau(05/02/2010)
        Call .AddValue("Banco_gera_nosso_numero", "@pBanco_gera_nosso_numero")
        Call .Parameters.add(.CreateParameter("@pBanco_gera_nosso_numero", obj.Banco_gera_nosso_numero, dbFieldTypeBool))
        
        'Pt. 97161 - Moacir Pfau(05/02/2010)
        Call .AddValue("Dias_protesto", "@pDias_protesto")
        Call .Parameters.add(.CreateParameter("@pDias_protesto", obj.Dias_protesto, dbFieldTypeInt))
        
        'pt.98446 - Fernando Paludo(26/04/2010)
        Call .AddValue("banco_emite_boleto", "@pbanco_emite_boleto")
        Call .Parameters.add(.CreateParameter("@pbanco_emite_boleto", obj.banco_Emite_boleto, dbFieldTypeBool))
        
        'pt.99257 - Moacir Pfau(30/06/2010)
        Call .AddValue("Per_multa", "@pPer_multa")
        Call .Parameters.add(.CreateParameter("@pPer_multa", obj.Per_multa, dbFieldTypeDouble))
        
        Call .AddValue("Per_mora", "@pPer_mora")
        Call .Parameters.add(.CreateParameter("@pPer_mora", obj.Per_mora, dbFieldTypeDouble))
        
        'pt.98929 - Fernando Paludo(02/08/2010)
        Call .AddValue("outra_especie", "@pOutra_especie")
        Call .Parameters.add(.CreateParameter("@pOutra_especie", obj.Outra_especie, dbFieldTypeString))
        
        'Pt. 102459 - Moacir Pfau(29/10/2010)
        Call .AddValue("NaoGerarRegistroCabecalho1", "@pNaoGerarRegistroCabecalho1")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroCabecalho1", obj.NaoGerarRegistroCabecalho1, dbFieldTypeBool))
        
        Call .AddValue("NaoGerarRegistroCabecalho2", "@pNaoGerarRegistroCabecalho2")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroCabecalho2", obj.NaoGerarRegistroCabecalho2, dbFieldTypeBool))
        
        Call .AddValue("NaoGerarRegistroDetalhe1", "@pNaoGerarRegistroDetalhe1")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroDetalhe1", obj.NaoGerarRegistroDetalhe1, dbFieldTypeBool))
        
        Call .AddValue("NaoGerarRegistroDetalhe2", "@pNaoGerarRegistroDetalhe2")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroDetalhe2", obj.NaoGerarRegistroDetalhe2, dbFieldTypeBool))
        
        Call .AddValue("NaoGerarRegistroDetalhe3", "@pNaoGerarRegistroDetalhe3")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroDetalhe3", obj.NaoGerarRegistroDetalhe3, dbFieldTypeBool))
        
        Call .AddValue("NaoGerarRegistroDetalhe4", "@pNaoGerarRegistroDetalhe4")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroDetalhe4", obj.NaoGerarRegistroDetalhe4, dbFieldTypeBool))
        
        Call .AddValue("NaoGerarRegistroDetalhe5", "@pNaoGerarRegistroDetalhe5")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroDetalhe5", obj.NaoGerarRegistroDetalhe5, dbFieldTypeBool))
        
        Call .AddValue("NaoGerarRegistroRodape1", "@pNaoGerarRegistroRodape1")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroRodape1", obj.NaoGerarRegistroRodape1, dbFieldTypeBool))
        
        Call .AddValue("NaoGerarRegistroRodape2", "@pNaoGerarRegistroRodape2")
        Call .Parameters.add(.CreateParameter("@pNaoGerarRegistroRodape2", obj.NaoGerarRegistroRodape2, dbFieldTypeBool))
        
        'Pt. 105912 - Moacir Pfau(17/03/2011)
        Call .AddValue("tipoimpressao", "@pTipoImpressao")
        Call .Parameters.add(.CreateParameter("@pTipoImpressao", obj.TipoImpressao, dbFieldTypeString))
        
        'Pt. 106012 - Moacir Pfau(28/09/2011)
        Call .AddValue("BairroSacado", "@pBairroSacado")
        Call .Parameters.add(.CreateParameter("@pBairroSacado", obj.BairroSacado, dbFieldTypeString))
        
        'Pt. 106012 - Moacir Pfau(28/09/2011)
        Call .AddValue("CodigoPracaSacado", "@pCodigoPracaSacado")
        Call .Parameters.add(.CreateParameter("@pCodigoPracaSacado", obj.CodigoPracaSacado, dbFieldTypeString))
        
        'Pt. 114032 - Moacir Pfau(23/02/2012)
        Call .AddValue("IdentificacaoCedente", "@pIdentificacaoCedente")
        Call .Parameters.add(.CreateParameter("@pIdentificacaoCedente", obj.IdentificacaoCedente, dbFieldTypeString))
        
        Call .AddValue("HTMLReciboPersonalizado", "@pHTMLReciboPersonalizado")
        Call .Parameters.add(.CreateParameter("@pHTMLReciboPersonalizado", obj.HTMLReciboPersonalizado, dbFieldTypeString))
        
        'Projeto: #4350 - História: # - Desenvolvimento# - Moacir Pfau(09/04/2013)
        Call .AddValue("data_baixa_retorno", "@pData_baixa_retorno")
        Call .Parameters.add(.CreateParameter("@pData_baixa_retorno", obj.Data_baixa_retorno, dbFieldTypeInt))
        
        'Projeto: #17081 - História: # - Desenvolvimento# - Moacir Pfau(02/12/2013)
        Call .AddValue("LocalPagamento", "@pLocalPagamento")
        Call .Parameters.add(.CreateParameter("@pLocalPagamento", obj.localPagamento, dbFieldTypeString))

        Call .AddValue("InstrucaoCobranca1", "@pInstrucaoCobranca_1")
        Call .Parameters.add(.CreateParameter("@pInstrucaoCobranca_1", obj.InstrucaoCobranca1, dbFieldTypeString))

        Call .AddValue("InstrucaoCobranca2", "@pInstrucaoCobranca_2")
        Call .Parameters.add(.CreateParameter("@pInstrucaoCobranca_2", obj.InstrucaoCobranca2, dbFieldTypeString))

        Call .AddValue("InstrucaoCobranca3", "@pInstrucaoCobranca_3")
        Call .Parameters.add(.CreateParameter("@pInstrucaoCobranca_3", obj.InstrucaoCobranca3, dbFieldTypeString))

        Call .AddValue("ValorInstrucaoCobranca1", "@pValorInstrucCobranc1")
        Call .Parameters.add(.CreateParameter("@pValorInstrucCobranc1", obj.ValorInstrucaoCobranca1, dbFieldTypeString))

        Call .AddValue("ValorInstrucaoCobranca2", "@pValorInstrucCobranc2")
        Call .Parameters.add(.CreateParameter("@pValorInstrucCobranc2", obj.ValorInstrucaoCobranca2, dbFieldTypeString))

        Call .AddValue("ValorInstrucaoCobranca3", "@pValorInstrucCobranc3")
        Call .Parameters.add(.CreateParameter("@pValorInstrucCobranc3", obj.ValorInstrucaoCobranca3, dbFieldTypeString))

        Call .AddValue("UtilizaNumeroControle", "@pUtilizaNumeroControle")
        Call .Parameters.add(.CreateParameter("@pUtilizaNumeroControle", obj.UtilizaNumeroControle, dbFieldTypeBool))
        
        'Vinicius Elyseu(06/10/2015) - Projeto: #0 - História: #0 - Desenv: #0
        Call .AddValue("seq_remessa_nr_doc", "@pSeqRemessaNrDoc")
        Call .Parameters.add(.CreateParameter("@pSeqRemessaNrDoc", obj.SeqRemessaNrDoc, dbFieldTypeBool))
    End With
End Sub

Private Function fProximoNumero(ByVal obj As clsCarteira) As Long
'    Dim strSql As String
'    Dim rstTab As Object
    Dim rdResult As IDBReader
    Dim cmd As IDBSelectCommand
    Dim result As clsCarteira
   
On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = mstrTabela
        .SelectClause = "MAX(id_carteira)+1 AS Sequencia"
        Call cmd.Filter.Append("[enterprise_id] = @pEnterprise_id")
        Call cmd.Parameters.add(cmd.CreateParameter("@pEnterprise_id", obj.Enterprise_id, dbFieldTypeLong))
    
        Call cmd.Filter.Append("[cd_estabelecimento] = @pCd_estabelecimento")
        Call cmd.Parameters.add(cmd.CreateParameter("@pCd_estabelecimento", obj.Cd_estabelecimento, dbFieldTypeLong))
        
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set result = New clsCarteira
        fProximoNumero = rdResult.GetLong("Sequencia")
        If fProximoNumero = 0 Then
            fProximoNumero = 1
        End If
        Set result = Nothing
    Else
        fProximoNumero = 1
    End If
    
    Exit Function
erro_carregando:
    Set result = Nothing
    Set cmd = Nothing
    Set rdResult = Nothing
End Function


