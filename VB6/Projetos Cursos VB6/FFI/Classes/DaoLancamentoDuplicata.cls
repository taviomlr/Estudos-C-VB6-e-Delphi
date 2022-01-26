VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DaoLancamentoDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mApp                            As IApplication

Public Sub init(ByVal objApp As IApplication)
    Set mApp = objApp
End Sub

Public Function Carregar(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup) As VoLancamentoDuplicata
    Dim rdResult            As IDBReader
    Dim cmd                 As IDBSelectCommand
    Dim result              As VoLancamentoDuplicata

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = fTabela(LancDup)
        
        Call .Filter.Append("[PagRec] = @pPagRec")
        Call .Parameters.add(cmd.CreateParameter("@pPagRec", IIf(PagRec = Pagamento, "P", "R"), dbFieldTypeString))
    
        If LancDup = Lancamento Then
            Call .Filter.Append("[Código] = @pCodigoNota")
        Else
            Call .Filter.Append("[Nota] = @pCodigoNota")
        End If
        Call .Parameters.add(cmd.CreateParameter("@pCodigoNota", lngNumero, dbFieldTypeDouble))
            
        Call .Filter.Append("[Parcela] = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", lngParcela, dbFieldTypeLong))
        
        If LancDup = Duplicata Then
            Call .Filter.Append("[Empresa] = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", strEmpresa, dbFieldTypeString))
            
            Call .Filter.Append("[Tipo] = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", strTipo, dbFieldTypeString))
        End If
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set Carregar = CriarObjeto(rdResult, LancDup)
    Else
        Set Carregar = Nothing
    End If
    
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
erro_carregando:
    Set Carregar = Nothing
    Set cmd = Nothing
    Set rdResult = Nothing
End Function

Public Function Gravar(ByVal obj As VoLancamentoDuplicata) As Boolean
    Dim cmd        As IDBInsertCommand
    Dim bolGravado As Boolean
    
On Error GoTo erro_gravando
    
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = fTabela(obj.LancDup)
        
        Call .AddValue("[PagRec]", "@pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", obj.PagRec, dbFieldTypeString))

        If obj.LancDup = Lancamento Then
            Call .AddValue("[Código]", "@pCodigoNota")
        Else
            Call .AddValue("[Nota]", "@pCodigoNota")
        End If
        If obj.Codigo_Nota = 0 Then
            obj.Codigo_Nota = fProximoNumero(obj.LancDup, obj.PagRec, obj.Tipo)
        End If
        
        Call .Parameters.add(cmd.CreateParameter("@pCodigoNota", obj.Codigo_Nota, dbFieldTypeDouble))
        
        If obj.Parcela = 0 Then
            obj.Parcela = fProximoParcela(obj.Codigo_Nota, obj.LancDup, obj.PagRec)
        End If
        Call .AddValue("[Parcela]", "@pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", obj.Parcela, dbFieldTypeLong))

        Call .AddValue("[Empresa]", "@pEmpresa")
        Call .Parameters.add(.CreateParameter("@pEmpresa", obj.Empresa, dbFieldTypeString))

        Call .AddValue("[Tipo]", "@pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", obj.Tipo, dbFieldTypeString))
        
        Call prepareCommand(cmd, obj)
    End With
    
    bolGravado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    If bolGravado Then
        Call ConfigSys.GravaUltimoLancDup(obj.LancDup, obj.Pagamento)
    End If
    
    Gravar = bolGravado
    
    Set cmd = Nothing
    Exit Function
erro_gravando:
    Gravar = False
    Set cmd = Nothing
    Call Throw(err)
End Function

Public Function Atualizar(ByVal obj As VoLancamentoDuplicata) As Boolean
    Dim cmd        As IDBUpdateCommand
    Dim blnGravado As Boolean
    
On Error GoTo erro_gravando
    Set cmd = mApp.CreateUpdateCommand
    With cmd
        .Table = fTabela(obj.LancDup)
        
        Call .Filter.Append("[PagRec] = @pPagRec")
        Call .Parameters.add(cmd.CreateParameter("@pPagRec", obj.PagRec, dbFieldTypeString))
        
        If obj.LancDup = Lancamento Then
            Call .Filter.Append("[Código] = @pCodigoNota")
        Else
            Call .Filter.Append("[Nota] = @pCodigoNota")
        End If
        Call .Parameters.add(cmd.CreateParameter("@pCodigoNota", obj.Codigo_Nota, dbFieldTypeString))
        
        Call .Filter.Append("[Parcela] = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", obj.Parcela, dbFieldTypeLong))
        
        If obj.LancDup = Duplicata Then
            If Trim(obj.Empresa) <> "" Then
                Call .Filter.Append("[Empresa] = @pEmpresa")
                Call .Parameters.add(.CreateParameter("@pEmpresa", obj.Empresa, dbFieldTypeString))
            End If
            
            Call .Filter.Append("[Tipo] = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", obj.Tipo, dbFieldTypeString))
        Else
            Call .AddValue("[Empresa]", "@pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", obj.Empresa, dbFieldTypeString))
        End If
        
        Call prepareCommand(cmd, obj)
    End With
    
    blnGravado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    If blnGravado Then
        Call ConfigSys.GravaUltimoLancDup(obj.LancDup, obj.Pagamento)
    End If
    
    Atualizar = blnGravado
    
    Set cmd = Nothing
    
    Exit Function
erro_gravando:
    Atualizar = False
    Set cmd = Nothing
    Call Throw(err)
End Function

Public Function Excluir(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal LancDup As enuLancDup, Optional ByVal strEmpresa As String, Optional strDataPagamento As String) As Boolean
    Dim cmd                             As IDBDeleteCommand
    Dim blnExcluido                     As Boolean
    Dim strTabela                       As String
    
On Error GoTo erro_excluindo
    
    blnExcluido = True
    
    Set cmd = mApp.CreateDeleteCommand
    With cmd
        .Table = fTabela(LancDup)
    
        Call .Filter.Append("[PagRec] = @pPagRec")
        Call .Parameters.add(cmd.CreateParameter("@pPagRec", IIf(PagRec = Pagamento, "P", "R"), dbFieldTypeString))
    
        If LancDup = Lancamento Then
            Call .Filter.Append("[Código] = @pCodigoNota")
        Else
            Call .Filter.Append("[Nota] = @pCodigoNota")
        End If
        Call .Parameters.add(cmd.CreateParameter("@pCodigoNota", lngNumero, dbFieldTypeDouble))
            
        Call .Filter.Append("[Parcela] = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", lngParcela, dbFieldTypeLong))
        
        If LancDup = Duplicata Then
            Call .Filter.Append("[Empresa] = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", strEmpresa, dbFieldTypeString))

            Call .Filter.Append("[Tipo] = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", strTipo, dbFieldTypeString))
        End If
    End With
        
    blnExcluido = CBool(mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = CLng(1))
    
    If blnExcluido Then
        Call ConfigSys.GravaUltimoLancDup(IIf(fTabela(LancDup) = "Lançamentos" = 0, 0, 1), strDataPagamento)
    End If
    
    If blnExcluido Then
        If LancDup = Duplicata Then
            strTabela = "FFIDuplicHistorico"
        Else
            strTabela = "FFILancHistorico"
        End If
        blnExcluido = blnExcluido And GravarHistorico(strTabela, IIf(PagRec = Pagamento, "P", "R"), strTipo, strEmpresa, lngNumero, lngParcela)
    End If
        
    Excluir = blnExcluido
    
    Set cmd = Nothing
    Exit Function
erro_excluindo:
    Set cmd = Nothing
    Excluir = False
    Call Throw(err)
End Function

Private Function fProximoNumero(ByVal LancDup As enuLancDup, ByVal strPagRec As String, Optional ByVal strTipo As String) As Double
    Dim rdResult As IDBReader
    Dim cmd As IDBSelectCommand

On Error GoTo erro_carregando

    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = fTabela(LancDup)
        
        Call .Filter.Append("[PagRec] = @pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", strPagRec, dbFieldTypeString))
        
        If LancDup = Duplicata Then
            .SelectClause = "MAX(Nota)+1 AS Sequencia"
        Else
            .SelectClause = "MAX(Código)+1 AS Sequencia"
        End If
        
        If LancDup = Duplicata And Trim(strTipo) <> "" Then
            Call .Filter.Append("[Tipo] = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", strTipo, dbFieldTypeString))
        End If
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    
    If Not rdResult.EOF Then
        fProximoNumero = rdResult.GetDouble("Sequencia")
        If fProximoNumero = 0 Then
            fProximoNumero = 1
        End If
        'Vinicius Elyseu(07/06/2016) - Projeto: #100340 SP6
        If Len(CDec(fProximoNumero)) > 15 Then
            fProximoNumero = ProximoGapDeNumero(strPagRec)
        End If
    Else
        fProximoNumero = 1
    End If
    
    Set cmd = Nothing
    Set rdResult = Nothing

    Exit Function
erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
End Function

Public Function fProximoParcela(ByVal lngCodigo_Nota As Double, ByVal LancDup As enuLancDup, ByVal strPagRec As String) As Long
    Dim rdResult As IDBReader
    Dim cmd As IDBSelectCommand

On Error GoTo erro_carregando

    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = fTabela(LancDup)
        
        Call .Filter.Append("[PagRec] = @pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", strPagRec, dbFieldTypeString))
                
        .SelectClause = "MAX(parcela)+1 AS Sequencia"
        
        If LancDup = Lancamento Then
            Call .Filter.Append("[Código] = @pCodigoNota")
        Else
            Call .Filter.Append("[Nota] = @pCodigoNota")
        End If
        Call .Parameters.add(cmd.CreateParameter("@pCodigoNota", lngCodigo_Nota, dbFieldTypeDouble))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    
    If Not rdResult.EOF Then
        fProximoParcela = rdResult.GetLong("Sequencia")
        If fProximoParcela = 0 Then
            fProximoParcela = 1
        End If
    Else
        fProximoParcela = 1
    End If
    
    Set cmd = Nothing
    Set rdResult = Nothing

    Exit Function
erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
End Function

Private Sub prepareCommand(ByRef cmd As Object, ByVal obj As VoLancamentoDuplicata)
    With cmd
        Call .AddValue("[Emissão]", "@pEmissao")
        Call .Parameters.add(.CreateParameter("@pEmissao", obj.Emissao, dbFieldTypeDate))

        Call .AddValue("[Vencimento]", "@pVencimento")
        Call .Parameters.add(.CreateParameter("@pVencimento", obj.Vencimento, dbFieldTypeDate))

        Call .AddValue("[Liberação]", "@pLiberacao")
        Call .Parameters.add(.CreateParameter("@pLiberacao", obj.Liberacao, dbFieldTypeDate))

        Call .AddValue("[Valor Original]", "@pValorOriginal")
        Call .Parameters.add(.CreateParameter("@pValorOriginal", obj.ValorOriginal, dbFieldTypeDouble))
        
        Call .AddValue("[Situação]", "@pSituacao")
        Call .Parameters.add(.CreateParameter("@pSituacao", obj.Situacao, dbFieldTypeString))

        Call .AddValue("[Descrição]", "@pDescricao")
        Call .Parameters.add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString))

        Call .AddValue("[Pagamento]", "@pPagamento")
        Call .Parameters.add(.CreateParameter("@pPagamento", obj.Pagamento, dbFieldTypeDate))

        Call .AddValue("[Acréscimo]", "@pAcrescimo")
        Call .Parameters.add(.CreateParameter("@pAcrescimo", obj.Acrescimo, dbFieldTypeDouble))

        Call .AddValue("[Abatimento]", "@pAbatimento")
        Call .Parameters.add(.CreateParameter("@pAbatimento", obj.Abatimento, dbFieldTypeDouble))

        Call .AddValue("[Banco]", "@pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", obj.Banco, dbFieldTypeDouble))

        Call .AddValue("[Conta]", "@pConta")
        Call .Parameters.add(.CreateParameter("@pConta", obj.conta, dbFieldTypeDouble))

        Call .AddValue("[Centro]", "@pCentro")
        Call .Parameters.add(.CreateParameter("@pCentro", obj.Centro, dbFieldTypeDouble))

        Call .AddValue("[Cheque]", "@pCheque")
        Call .Parameters.add(.CreateParameter("@pCheque", obj.Cheque, dbFieldTypeDouble))

        Call .AddValue("[Moeda]", "@pMoeda")
        Call .Parameters.add(.CreateParameter("@pMoeda", obj.Moeda, dbFieldTypeString))

        Call .AddValue("[Valor da Moeda]", "@pValordaMoeda")
        Call .Parameters.add(.CreateParameter("@pValordaMoeda", obj.ValordaMoeda, dbFieldTypeDouble))

        Call .AddValue("[Controle]", "@pControle")
        Call .Parameters.add(.CreateParameter("@pControle", obj.Controle, dbFieldTypeString))

        Call .AddValue("[Marcação]", "@pMarcacao")
        Call .Parameters.add(.CreateParameter("@pMarcacao", obj.Marcacao, dbFieldTypeBool))

        Call .AddValue("[Obs]", "@pObs")
        Call .Parameters.add(.CreateParameter("@pObs", obj.Obs, dbFieldTypeString))

        Call .AddValue("[Borderô]", "@pBordero")
        Call .Parameters.add(.CreateParameter("@pBordero", obj.Bordero, dbFieldTypeLong))

        Call .AddValue("[Data Integração Fato Gerador]", "@pDataIntegracaoFatoGerador")
        Call .Parameters.add(.CreateParameter("@pDataIntegracaoFatoGerador", obj.DataIntegracaoFatoGerador, dbFieldTypeDate))

        Call .AddValue("[Data Integração Fato Pagamento]", "@pDataIntegracaoFatoPagamento")
        Call .Parameters.add(.CreateParameter("@pDataIntegracaoFatoPagamento", obj.DataIntegracaoFatoPagamento, dbFieldTypeDate))

        Call .AddValue("[Usuário]", "@pUsuario")
        Call .Parameters.add(.CreateParameter("@pUsuario", obj.usuario, dbFieldTypeString))

        Call .AddValue("[Alteração]", "@pAlteracao")
        Call .Parameters.add(.CreateParameter("@pAlteracao", obj.Alteracao, dbFieldTypeDate))

        Call .AddValue("[LINDIG]", "@pLINDIG")
        Call .Parameters.add(.CreateParameter("@pLINDIG", obj.LINDIG, dbFieldTypeString))

        Call .AddValue("[Conciliado]", "@pConciliado")
        Call .Parameters.add(.CreateParameter("@pConciliado", obj.Conciliado, dbFieldTypeBool))

        Call .AddValue("[Carteira]", "@pCarteira")
        Call .Parameters.add(.CreateParameter("@pCarteira", obj.Carteira, dbFieldTypeString))
      
        Call .AddValue("[VlrMrD]", "@pVlrMrD")
        Call .Parameters.add(.CreateParameter("@pVlrMrD", obj.VlrMrD, dbFieldTypeDouble))

        Call .AddValue("[PerMrD]", "@pPerMrD")
        Call .Parameters.add(.CreateParameter("@pPerMrD", obj.PerMrD, dbFieldTypeDouble))

        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        Call .AddValue("[SeqNossoNumero]", "@pSeqNossoNumero")
        Call .Parameters.add(.CreateParameter("@pSeqNossoNumero", obj.SeqNossoNumero, dbFieldTypeString, 40))

        Call .AddValue("[VlrDsP]", "@pVlrDsP")
        Call .Parameters.add(.CreateParameter("@pVlrDsP", obj.VlrDsP, dbFieldTypeDouble))
        
        Call .AddValue("[PerJur]", "@pPerJur")
        Call .Parameters.add(.CreateParameter("@pPerJur", obj.PerJur, dbFieldTypeDouble))
    
        Call .AddValue("[VlrJur]", "@pVlrJur")
        Call .Parameters.add(.CreateParameter("@pVlrJur", obj.VlrJur, dbFieldTypeDouble))
        
        Call .AddValue("[CODPRJ]", "@pCODPRJ")
        Call .Parameters.add(.CreateParameter("@pCODPRJ", obj.CODPRJ, dbFieldTypeLong))

        Call .AddValue("[CODFPG]", "@pCODFPG")
        Call .Parameters.add(.CreateParameter("@pCODFPG", obj.CODFPG, dbFieldTypeLong))

        Call .AddValue("[CheBan]", "@pCheBan")
        Call .Parameters.add(.CreateParameter("@pCheBan", obj.CheBan, dbFieldTypeLong))

        Call .AddValue("[CheAge]", "@pCheAge")
        Call .Parameters.add(.CreateParameter("@pCheAge", obj.CheAge, dbFieldTypeString))

        Call .AddValue("[CheEmi]", "@pCheEmi")
        Call .Parameters.add(.CreateParameter("@pCheEmi", obj.CheEmi, dbFieldTypeString))

        Call .AddValue("[CheCco]", "@pCheCco")
        Call .Parameters.add(.CreateParameter("@pCheCco", obj.CheCco, dbFieldTypeString))
        
        Call .AddValue("[PerMul]", "@pPerMul")
        Call .Parameters.add(.CreateParameter("@pPerMul", obj.PerMul, dbFieldTypeDouble))

        Call .AddValue("[VlrMul]", "@pVlrMul")
        Call .Parameters.add(.CreateParameter("@pVlrMul", obj.VlrMul, dbFieldTypeDouble))

        Call .AddValue("[AGECCE]", "@pAGECCE")
        Call .Parameters.add(.CreateParameter("@pAGECCE", obj.AGECCE, dbFieldTypeString))

        Call .AddValue("[NOSNUM]", "@pNOSNUM")
        Call .Parameters.add(.CreateParameter("@pNOSNUM", obj.NOSNUM, dbFieldTypeString))
        
        Call .AddValue("[CODBAR]", "@pCODBAR")
        Call .Parameters.add(.CreateParameter("@pCODBAR", obj.CodBar, dbFieldTypeString))

        Call .AddValue("[LOCPAG]", "@pLOCPAG")
        Call .Parameters.add(.CreateParameter("@pLOCPAG", obj.LOCPAG, dbFieldTypeString))

        Call .AddValue("[INSTRU]", "@pINSTRU")
        Call .Parameters.add(.CreateParameter("@pINSTRU", obj.INSTRU, dbFieldTypeString))

        Call .AddValue("[cd_operacao_contabil]", "@pcd_operacao_contabil")
        Call .Parameters.add(.CreateParameter("@pcd_operacao_contabil", obj.cd_operacao_contabil, dbFieldTypeLong))

        Call .AddValue("[registro_integrado]", "@pregistro_integrado")
        Call .Parameters.add(.CreateParameter("@pregistro_integrado", obj.registro_integrado, dbFieldTypeBool))
        
        Call .AddValue("cd_operacao_baixa", "@pcd_operacao_baixa")
        Call .Parameters.add(.CreateParameter("@pcd_operacao_baixa", obj.cd_operacao_baixa, dbFieldTypeLong))
        
        Call .AddValue("baixa_integrado", "@pbaixa_integrado")
        Call .Parameters.add(.CreateParameter("@pbaixa_integrado", obj.baixa_integrado, dbFieldTypeBool))
        
        Call .AddValue("[proveniente_rateio]", "@pproveniente_rateio")
        Call .Parameters.add(.CreateParameter("@pproveniente_rateio", obj.proveniente_rateio, dbFieldTypeBool))
        
        Call .AddValue("parc_origem_baixa", "@pparc_origem_baixa")
        Call .Parameters.add(.CreateParameter("@pparc_origem_baixa", obj.parc_origem_baixa, dbFieldTypeLong))

        Call .AddValue("id_carteira", "@pid_carteira")
        Call .Parameters.add(.CreateParameter("@pid_carteira", obj.Id_carteira, dbFieldTypeLong))

        Call .AddValue("cod_motivo", "@pcod_motivo")
        Call .Parameters.add(.CreateParameter("@pcod_motivo", obj.cod_motivo, dbFieldTypeString))

        Call .AddValue("desc_motivo", "@pdesc_motivo")
        Call .Parameters.add(.CreateParameter("@pdesc_motivo", obj.desc_motivo, dbFieldTypeString))

        'Projeto: #218 - História: #268 - Desenvolvimento#589 - Moacir Pfau(21/09/2012)
        Call .AddValue("ultima_mov_cnab", "@pultimamovcnab")
        Call .Parameters.add(.CreateParameter("@pultimamovcnab", obj.ultima_mov_cnab, dbFieldTypeLong))
        
        Call .AddValue("ultima_data_cnab", "@pultima_data_cnab")
        Call .Parameters.add(.CreateParameter("@pultima_data_cnab", obj.ultima_data_cnab, dbFieldTypeDate))

        Call .AddValue("cd_cobranca", "@pcd_cobranca")
        Call .Parameters.add(.CreateParameter("@pcd_cobranca", obj.cd_cobranca, dbFieldTypeLong))

        Call .AddValue("seq_remessa", "@pseq_remessa")
        Call .Parameters.add(.CreateParameter("@pseq_remessa", obj.seq_remessa, dbFieldTypeLong))
        
        'Projeto: #35566 - História: # - Desenvolvimento#43993 -  Vinicius Alexandre Elyseu (06/08/2014)
        Call .AddValue("integracao_bi", "@pIntegracaoBI")
        Call .Parameters.add(.CreateParameter("@pIntegracaoBI", 0, dbFieldTypeLong))

        If obj.LancDup = Lancamento Then
            Call .AddValue("[Previsao]", "@pPrevisao")
            Call .Parameters.add(.CreateParameter("@pPrevisao", obj.previsao, dbFieldTypeBool))
    
            Call .AddValue("data_integracao", "@pdata_integracao")
            Call .Parameters.add(.CreateParameter("@pdata_integracao", obj.data_integracao, dbFieldTypeDate))
            
            Call .AddValue("[Enviado]", "@pEnviado")
            Call .Parameters.add(.CreateParameter("@pEnviado", obj.Enviado, dbFieldTypeString))
        End If
    End With
End Sub

Public Function CriarObjeto(ByVal rd As IDBReader, ByVal LancDup As enuLancDup) As VoLancamentoDuplicata
    Dim obj As New VoLancamentoDuplicata
    
    With rd
        obj.PagRec = .GetString("PagRec")
        If LancDup = Lancamento Then
            obj.Codigo_Nota = .GetString("Código")
            obj.Enviado = .GetString("Enviado")
            obj.previsao = .GetBoolean("Previsao")
            obj.data_integracao = .GetDate("data_integracao")
        ElseIf LancDup = Duplicata Then
            obj.Codigo_Nota = .GetString("Nota")
        End If
        obj.Parcela = .GetLong("Parcela")
        obj.Empresa = .GetString("Empresa")
        obj.Tipo = .GetString("Tipo")
        obj.Descricao = .GetString("Descrição")
        obj.Emissao = .GetDate("Emissão")
        obj.Vencimento = .GetDate("Vencimento")
        obj.Pagamento = .GetDate("Pagamento")
        obj.Liberacao = .GetDate("Liberação")
        obj.ValorOriginal = .GetDouble("Valor Original")
        obj.Acrescimo = .GetDouble("Acréscimo")
        obj.Abatimento = .GetDouble("Abatimento")
        obj.Banco = .GetLong("Banco")
        obj.conta = .GetLong("Conta")
        obj.Centro = .GetLong("Centro")
        obj.Cheque = .GetLong("Cheque")
        obj.Moeda = .GetString("Moeda")
        obj.ValordaMoeda = .GetDouble("Valor da Moeda")
        obj.Controle = .GetString("Controle")
        obj.Situacao = .GetString("Situação")
        obj.Marcacao = .GetBoolean("Marcação")
        obj.Obs = .GetString("Obs")
        obj.Bordero = .GetLong("Borderô")
        obj.DataIntegracaoFatoGerador = .GetDate("Data Integração Fato Gerador")
        obj.DataIntegracaoFatoPagamento = .GetDate("Data Integração Fato Pagamento")
        obj.usuario = .GetString("Usuário")
        obj.Alteracao = .GetDate("Alteração")
        obj.LINDIG = .GetString("LINDIG")
        obj.Conciliado = .GetBoolean("Conciliado")
        obj.Carteira = .GetString("Carteira")
        obj.VlrMrD = .GetDouble("VlrMrD")
        obj.PerMrD = .GetDouble("PerMrD")
        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        obj.SeqNossoNumero = .GetString("SeqNossoNumero")
        obj.VlrDsP = .GetDouble("VlrDsP")
        obj.PerJur = .GetDouble("PerJur")
        obj.VlrJur = .GetDouble("VlrJur")
        obj.CODPRJ = .GetLong("CODPRJ")
        obj.CODFPG = .GetLong("CODFPG")
        obj.CheBan = .GetLong("CheBan")
        obj.CheAge = .GetString("CheAge")
        obj.CheEmi = .GetString("CheEmi")
        obj.CheCco = .GetString("CheCco")
        obj.PerMul = .GetDouble("PerMul")
        obj.VlrMul = .GetDouble("VlrMul")
        obj.AGECCE = .GetString("AGECCE")
        obj.NOSNUM = .GetString("NOSNUM")
        obj.CodBar = .GetString("CODBAR")
        obj.LOCPAG = .GetString("LOCPAG")
        obj.INSTRU = .GetString("INSTRU")
        obj.cd_operacao_contabil = .GetLong("cd_operacao_contabil")
        obj.registro_integrado = .GetBoolean("registro_integrado")
        obj.cd_operacao_baixa = .GetLong("cd_operacao_baixa")
        obj.baixa_integrado = .GetBoolean("baixa_integrado")
        obj.proveniente_rateio = .GetBoolean("proveniente_rateio")
        obj.parc_origem_baixa = .GetLong("parc_origem_baixa")
        obj.Id_carteira = .GetLong("id_carteira")
        obj.cod_motivo = .GetString("cod_motivo")
        obj.desc_motivo = .GetString("desc_motivo")
        obj.ultima_mov_cnab = .GetLong("ultima_mov_cnab")
        obj.ultima_data_cnab = .GetDate("ultima_data_cnab")
        obj.cd_cobranca = .GetLong("cd_cobranca")
        obj.seq_remessa = .GetLong("seq_remessa")
        obj.LancDup = LancDup
        'Vinicius Elyseu(01/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
        obj.Remessa = .GetBoolean("remessa")
        obj.Boleto = .GetBoolean("boleto")
        
    End With
    Set CriarObjeto = obj
End Function

Public Function CriarObjetoRateioLancamentoDuplicata(ByVal rd As IDBReader, ByVal eLancDup As enuLancDup) As VoRateioLancamentoDuplicata
    Dim obj As New VoRateioLancamentoDuplicata
    With rd
        obj.pag_rec_origem = .GetString("pag_rec_origem")
        If eLancDup = Duplicata Then
            obj.codigo_nota_origem = .GetLong("nr_nota_origem")
            obj.codigo_nota_destino = .GetLong("nr_nota_destino")
            obj.cd_empresa_origem = .GetString("cd_empresa_origem")
            obj.tp_registro_origem = .GetString("tp_registro_origem")
            obj.cd_empresa_destino = .GetString("cd_empresa_destino")
            obj.tp_registro_destino = .GetString("tp_registro_destino")
        Else
            obj.codigo_nota_origem = .GetLong("cd_lancamento_origem")
            obj.codigo_nota_destino = .GetLong("cd_lancamento_destino")
        End If
        obj.nr_parcela_origem = .GetLong("nr_parcela_origem")
        obj.pag_rec_destino = .GetString("pag_rec_destino")
        obj.nr_parcela_destino = .GetLong("nr_parcela_destino")
        obj.cd_centro = .GetLong("cd_centro")
        obj.Cd_conta = .GetLong("cd_conta")
        obj.dt_pagamento = .GetDate("dt_pagamento")
        obj.vl_valor = .GetDouble("vl_valor")
    End With
    Set CriarObjetoRateioLancamentoDuplicata = obj
End Function

Private Function fTabela(ByVal intValor As enuLancDup) As String
    Dim strTabela                 As String
    
    strTabela = ""
    If intValor = Lancamento Then
        strTabela = "[Lançamentos]"
    ElseIf intValor = Duplicata Then
        strTabela = "[Duplicatas]"
    End If
    fTabela = strTabela
End Function

Private Function fTabelaRateio(ByVal intValor As enuLancDup) As String
    Dim strTabela                 As String
    
    strTabela = ""
    If intValor = Lancamento Then
        strTabela = "[FFIRateioLancamento]"
    ElseIf intValor = Duplicata Then
        strTabela = "[FFIRateioDuplicata]"
    End If
    fTabelaRateio = strTabela
End Function

Public Function GravarRateioLancamentoDuplicata(ByVal obj As VoRateioLancamentoDuplicata, ByVal eLancDup As enuLancDup) As Boolean
    Dim cmd        As IDBInsertCommand
    Dim bolGravado As Boolean
    
On Error GoTo erro_gravando
    
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = fTabelaRateio(eLancDup)

        Call .AddValue("[pag_rec_origem]", "@pPag_rec_origem")
        Call .Parameters.add(.CreateParameter("@pPag_rec_origem", obj.pag_rec_origem, dbFieldTypeString))
                
        If eLancDup = Lancamento Then
            Call .AddValue("[cd_lancamento_origem]", "@pcodigo_nota_origem")
            Call .Parameters.add(.CreateParameter("@pcodigo_nota_origem", obj.codigo_nota_origem, dbFieldTypeString))
        Else
            Call .AddValue("[nr_nota_origem]", "@pcodigo_nota_origem")
            Call .Parameters.add(.CreateParameter("@pcodigo_nota_origem", obj.codigo_nota_origem, dbFieldTypeLong))
        End If
        

        Call .AddValue("[nr_parcela_origem]", "@pNr_parcela_origem")
        Call .Parameters.add(.CreateParameter("@pNr_parcela_origem", obj.nr_parcela_origem, dbFieldTypeLong))

        Call .AddValue("[pag_rec_destino]", "@pPag_rec_destino")
        Call .Parameters.add(.CreateParameter("@pPag_rec_destino", obj.pag_rec_destino, dbFieldTypeString))

        If eLancDup = Lancamento Then
            Call .AddValue("[cd_lancamento_destino]", "@pcodigo_nota_destino")
            Call .Parameters.add(.CreateParameter("@pcodigo_nota_destino", obj.codigo_nota_destino, dbFieldTypeString))
        Else
            Call .AddValue("[nr_nota_destino]", "@pcodigo_nota_destino")
            Call .Parameters.add(.CreateParameter("@pcodigo_nota_destino", obj.codigo_nota_destino, dbFieldTypeLong))
        End If
        

        Call .AddValue("[nr_parcela_destino]", "@pNr_parcela_destino")
        Call .Parameters.add(.CreateParameter("@pNr_parcela_destino", obj.nr_parcela_destino, dbFieldTypeLong))

        Call .AddValue("[cd_centro]", "@pCd_centro")
        Call .Parameters.add(.CreateParameter("@pCd_centro", obj.cd_centro, dbFieldTypeLong))

        Call .AddValue("[cd_conta]", "@pCd_conta")
        Call .Parameters.add(.CreateParameter("@pCd_conta", obj.Cd_conta, dbFieldTypeLong))

        Call .AddValue("[dt_pagamento]", "@pDt_pagamento")
        Call .Parameters.add(.CreateParameter("@pDt_pagamento", obj.dt_pagamento, dbFieldTypeDate))

        Call .AddValue("[vl_valor]", "@pVl_valor")
        Call .Parameters.add(.CreateParameter("@pVl_valor", obj.vl_valor, dbFieldTypeDouble))
        
        If eLancDup = Duplicata Then
            Call .AddValue("[cd_empresa_origem]", "@pcd_empresa_origem")
            Call .Parameters.add(.CreateParameter("@pcd_empresa_origem", obj.cd_empresa_origem, dbFieldTypeString))
            
            Call .AddValue("[cd_empresa_destino]", "@pcd_empresa_destino")
            Call .Parameters.add(.CreateParameter("@pcd_empresa_destino", obj.cd_empresa_destino, dbFieldTypeString))
            
            Call .AddValue("[tp_registro_origem]", "@ptp_registro_origem")
            Call .Parameters.add(.CreateParameter("@ptp_registro_origem", obj.tp_registro_origem, dbFieldTypeString))
            
            Call .AddValue("[tp_registro_destino]", "@ptp_registro_destino")
            Call .Parameters.add(.CreateParameter("@ptp_registro_destino", obj.tp_registro_destino, dbFieldTypeString))
        End If
    End With
    
    bolGravado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    GravarRateioLancamentoDuplicata = bolGravado
    
    Set cmd = Nothing
    Exit Function
erro_gravando:
    GravarRateioLancamentoDuplicata = False
    Set cmd = Nothing
    Call Throw(err)
End Function

Public Function CarregaRateioLancamentoDuplicataOrigem(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup) As colRateioLancamentoDuplicata
    Dim rdResult            As IDBReader
    Dim cmd                 As IDBSelectCommand
    Dim result              As VoRateioLancamentoDuplicata

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = fTabelaRateio(LancDup)
        
        Call .Filter.Append("[pag_rec_origem] = @pPagRec")
        Call .Parameters.add(cmd.CreateParameter("@pPagRec", IIf(PagRec = Pagamento, "P", "R"), dbFieldTypeString))
    
        If LancDup = Lancamento Then
            Call .Filter.Append("[cd_lancamento_origem] = @pCodigoNota")
        Else
            Call .Filter.Append("[nr_nota_origem] = @pCodigoNota")
        End If
        Call .Parameters.add(cmd.CreateParameter("@pCodigoNota", lngNumero, dbFieldTypeDouble))
        
        Call .Filter.Append("[nr_parcela_origem] = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", lngParcela, dbFieldTypeLong))
        
        If LancDup = Duplicata Then
            Call .Filter.Append("[cd_empresa_origem] = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", strEmpresa, dbFieldTypeString))
            
            Call .Filter.Append("[tp_registro_origem] = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", strTipo, dbFieldTypeString))
        End If
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    
    If Not rdResult.EOF Then
        Set CarregaRateioLancamentoDuplicataOrigem = New colRateioLancamentoDuplicata
    End If
    While Not rdResult.EOF
        Set result = New VoRateioLancamentoDuplicata
        Call CarregaRateioLancamentoDuplicataOrigem.add(CriarObjetoRateioLancamentoDuplicata(rdResult, LancDup))
        
        rdResult.MoveNext
    Wend
    
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
End Function

Public Function CarregaRateioLancamentoDuplicataDestino(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup) As VoRateioLancamentoDuplicata
    Dim rdResult            As IDBReader
    Dim cmd                 As IDBSelectCommand

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = fTabelaRateio(LancDup)
        
        Call .Filter.Append("[pag_rec_destino] = @pPagRec")
        Call .Parameters.add(cmd.CreateParameter("@pPagRec", IIf(PagRec = Pagamento, "P", "R"), dbFieldTypeString))
    
        If LancDup = Lancamento Then
            Call .Filter.Append("[cd_lancamento_destino] = @pCodigoNota")
        Else
            Call .Filter.Append("[nr_nota_destino] = @pCodigoNota")
        End If
        Call .Parameters.add(cmd.CreateParameter("@pCodigoNota", lngNumero, dbFieldTypeDouble))
        
        Call .Filter.Append("[nr_parcela_destino] = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", lngParcela, dbFieldTypeLong))
        
        If LancDup = Duplicata Then
            Call .Filter.Append("[cd_empresa_destino] = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", strEmpresa, dbFieldTypeString))
            
            Call .Filter.Append("[tp_registro_destino] = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", strTipo, dbFieldTypeString))
        End If
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set CarregaRateioLancamentoDuplicataDestino = CriarObjetoRateioLancamentoDuplicata(rdResult, LancDup)
    Else
        Set CarregaRateioLancamentoDuplicataDestino = Nothing
    End If
    
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
erro_carregando:
    Set CarregaRateioLancamentoDuplicataDestino = Nothing
    Set cmd = Nothing
    Set rdResult = Nothing
End Function


Public Function ExcluirRateioDestino(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal LancDup As enuLancDup, Optional ByVal strEmpresa As String) As Boolean
    Dim cmd                             As IDBDeleteCommand
    Dim blnExcluido                     As Boolean
    
On Error GoTo erro_excluindo
    
    blnExcluido = True
    
    Set cmd = mApp.CreateDeleteCommand
    With cmd
        .Table = fTabelaRateio(LancDup)
    
        Call .Filter.Append("[pag_rec_destino] = @pPagRec")
        Call .Parameters.add(cmd.CreateParameter("@pPagRec", IIf(PagRec = Pagamento, "P", "R"), dbFieldTypeString))
    
        If LancDup = Lancamento Then
            Call .Filter.Append("[cd_lancamento_destino] = @pCodigoNota")
        Else
            Call .Filter.Append("[nr_nota_destino] = @pCodigoNota")
        End If
        Call .Parameters.add(cmd.CreateParameter("@pCodigoNota", lngNumero, dbFieldTypeDouble))
            
        Call .Filter.Append("[nr_parcela_destino] = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", lngParcela, dbFieldTypeLong))
        
        If LancDup = Duplicata Then
            Call .Filter.Append("[cd_empresa_destino] = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", strEmpresa, dbFieldTypeString))
            
            Call .Filter.Append("[tp_registro_destino] = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", strTipo, dbFieldTypeString))
        End If
    End With
        
    blnExcluido = CBool(mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = CLng(1))
        
    ExcluirRateioDestino = blnExcluido
    
    Set cmd = Nothing
    Exit Function
erro_excluindo:
    Set cmd = Nothing
    ExcluirRateioDestino = False
    Call Throw(err)
End Function

Public Function ExisteRateioLancamentoDuplicataOrigem(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup) As Boolean
    Dim col                 As New colRateioLancamentoDuplicata
    Set col = CarregaRateioLancamentoDuplicataOrigem(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup)
    If col.Count > 0 Then
        ExisteRateioLancamentoDuplicataOrigem = True
    End If
End Function


Public Function GravarHistorico(strTabela As String, strPagRec As String, strTpRegistro As String, strEmpresa As String, intNum As String, intNumParcela As Long) As Boolean
    Dim cmd         As IDBInsertCommand
    Dim booExcluido As Boolean
    
On Error GoTo erro_inserindo
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateInsertCommand
    With cmd
        .Table = strTabela
                
        Call .AddValue("[id_seq]", "@pIdSequencial")
        Call .Parameters.add(.CreateParameter("@pIdSequencial", ProximoNumero("id_seq", strTabela, NUL)))
        
        Call .AddValue("[enterprise_id]", "@pEnterpriseId")
        Call .Parameters.add(.CreateParameter("@pEnterpriseId", EnterpriseId, dbFieldTypeInt))
        
        Call .AddValue("[cd_estabelecimento]", "@pCdEstabelecimento")
        Call .Parameters.add(.CreateParameter("@pCdEstabelecimento", CdEstabelecimento, dbFieldTypeInt))
        
        Call .AddValue("[pagRec]", "@pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", strPagRec, dbFieldTypeString))
        
        Call .AddValue("[tp_registro]", "@pTipoRegistro")
        Call .Parameters.add(.CreateParameter("@pTipoRegistro", strTpRegistro, dbFieldTypeString))
        
        Call .AddValue("[empresa]", "@pEmpresa")
        Call .Parameters.add(.CreateParameter("@pEmpresa", strEmpresa, dbFieldTypeString))
                       
        If strTabela = "FFIDuplicHistorico" Then
            Call .AddValue("[nr_duplic]", "@pNumDuplicLanc")
        Else
            Call .AddValue("[nr_lanc]", "@pNumDuplicLanc")
        End If
        Call .Parameters.add(.CreateParameter("@pNumDuplicLanc", intNum, dbFieldTypeString))
        
        Call .AddValue("[nr_parcela]", "@pNumParcela")
        Call .Parameters.add(.CreateParameter("@pNumParcela", intNumParcela, dbFieldTypeLong))
        
        Call .AddValue("[usuario]", "@pUsuario")
        Call .Parameters.add(.CreateParameter("@pUsuario", UserName, dbFieldTypeString))
        
        Call .AddValue("[dataHora]", "@pDataHora")
        Call .Parameters.add(.CreateParameter("@pDataHora", Now, dbFieldTypeDateTime))
        
        Call .AddValue("[integracao_bi]", "@pIntegracaoBI")
        Call .Parameters.add(.CreateParameter("@pIntegracaoBI", 0, dbFieldTypeInt))
       
    End With
    
    booExcluido = Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd) > 0
    
    GravarHistorico = booExcluido
    Aplicacao.Disconnect
    Exit Function
erro_inserindo:
    MsgBox "Erro ao inserir histórico de exclusão de lançamento: " & err.Description
    Call Throw(err)
    Aplicacao.Disconnect
End Function
'Autor.......: Vinicius
'Data........: 21/01/2015
'Descrição...: Função utilizada para retornar se um extrato tem conciliação em duplicata ou lançamento
Public Function ExisteExtratoConciliado(strTabela As String, lngExtrato As Long, intBanco As Integer, Optional intSeqExtrato As Integer) As Boolean
    Dim cmdSelect As New IDBSelectCommand
    Dim rdResult  As New IDBReader
    
On Error GoTo err_Handler
    
    Aplicacao.Connect
    Set cmdSelect = Aplicacao.CreateSelectCommand
    With cmdSelect
        
        .Table.TableName = strTabela
        
        Call .Filter.Append("conciliacao_banco = @pConciliadoBanco")
        Call .Parameters.add(.CreateParameter("@pConciliadoBanco", intBanco, dbFieldTypeInt))
        
        Call .Filter.Append("conciliacao_extrato = @pExtrato")
        Call .Parameters.add(.CreateParameter("@pExtrato", lngExtrato, dbFieldTypeLong))
        
        If intSeqExtrato > 0 Then
            Call .Filter.Append("conciliacao_sequencial_extrato = @pSeqExtrato")
            Call .Parameters.add(.CreateParameter("@pSeqExtrato", intSeqExtrato, dbFieldTypeInt))
        End If
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmdSelect)
    If Not rdResult.EOF Then
        ExisteExtratoConciliado = True
    End If
    Aplicacao.Disconnect
    Exit Function
    
err_Handler:
    ExisteExtratoConciliado = False
    Aplicacao.Disconnect
End Function
