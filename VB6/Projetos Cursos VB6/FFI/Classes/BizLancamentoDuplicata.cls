VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BizLancamentoDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Carregar(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup) As VoLancamentoDuplicata
    Dim dao As New DaoLancamentoDuplicata
    Aplicacao.Connect
    Call dao.init(Aplicacao)
    Set Carregar = dao.Carregar(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup)
    Set dao = Nothing
    Aplicacao.Disconnect
End Function

Public Function Gravar(ByVal obj As VoLancamentoDuplicata, ByVal objCheque As VoCheque) As Boolean
    Dim dao                 As New DaoLancamentoDuplicata
    Dim BCheque             As New bizCheque
    Dim blnGravar           As Boolean
    
On Error GoTo err
    Aplicacao.Connect
    Aplicacao.BeginTransaction
    
    Call dao.init(Aplicacao)
    blnGravar = dao.Gravar(obj)
    
    If blnGravar And obj.Cheque > 0 And Not (objCheque Is Nothing) Then
        blnGravar = BCheque.Gravar_Atualizar(objCheque)
    End If
    
    If blnGravar Then
        Aplicacao.CommitTransaction
    Else
        Aplicacao.RollbackTransaction
    End If
    Set dao = Nothing
    Aplicacao.Disconnect
    Gravar = blnGravar
    Exit Function
err:
    Set dao = Nothing
    Aplicacao.RollbackTransaction
    Aplicacao.Disconnect
    Gravar = False
End Function

Public Function Atualizar(ByVal obj As VoLancamentoDuplicata, Optional ByVal objCheque As VoCheque) As Boolean
    Dim dao                 As New DaoLancamentoDuplicata
    Dim BCheque             As New bizCheque
    Dim blnAtualizar        As Boolean
    
On Error GoTo err
    Aplicacao.Connect
    Aplicacao.BeginTransaction
    Call dao.init(Aplicacao)
    blnAtualizar = dao.Atualizar(obj)
    
    If blnAtualizar And obj.Cheque > 0 And Not (objCheque Is Nothing) Then
        blnAtualizar = BCheque.Gravar_Atualizar(objCheque)
    End If
    
    If blnAtualizar Then
        Aplicacao.CommitTransaction
    Else
        Aplicacao.RollbackTransaction
    End If
    
    Set dao = Nothing
    Aplicacao.Disconnect
    Atualizar = blnAtualizar
    Exit Function
err:
    Aplicacao.RollbackTransaction
    Set dao = Nothing
    Aplicacao.Disconnect
    Atualizar = False
End Function

Public Function Excluir(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup, Optional blnExcluirRateioVinculado As Boolean, Optional strDataPagamento As String) As Boolean
    Dim dao                 As New DaoLancamentoDuplicata
    Dim blnExcluir          As Boolean
On Error GoTo err
    Aplicacao.Connect
    Aplicacao.BeginTransaction
    Call dao.init(Aplicacao)
    
    blnExcluir = ExcluirRateioLancamentoDuplicata(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup, blnExcluirRateioVinculado)
    
    If blnExcluir Then
        blnExcluir = blnExcluir And dao.Excluir(PagRec, lngNumero, strTipo, lngParcela, LancDup, strEmpresa, strDataPagamento)
    End If
    Excluir = blnExcluir
    Set dao = Nothing
    Aplicacao.CommitTransaction
    Aplicacao.Disconnect
    Exit Function
err:
    Aplicacao.RollbackTransaction
    Aplicacao.Disconnect
End Function

Public Function ExcluirRateioLancamentoDuplicata(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup, Optional blnExcluirRateioVinculado As Boolean) As Boolean
    Dim dao                             As New DaoLancamentoDuplicata
    Dim blnExcluir                      As Boolean
    Dim colLancamentoDuplicata          As colRateioLancamentoDuplicata
    Dim voRLancamentoDuplicata          As VoRateioLancamentoDuplicata
    Dim voLancDup                       As VoLancamentoDuplicata
    Dim dblValor                        As Double
    Dim dblAcrescimo                    As Double
    Dim dblAbatimento                   As Double
 
 On Error GoTo err
    Aplicacao.Connect
    Aplicacao.BeginTransaction
    Call dao.init(Aplicacao)
    
    'Verificar o pai
    blnExcluir = True
    If blnExcluirRateioVinculado Then
        Set colLancamentoDuplicata = New colRateioLancamentoDuplicata
        Set colLancamentoDuplicata = dao.CarregaRateioLancamentoDuplicataOrigem(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup)
        If Not colLancamentoDuplicata Is Nothing Then
            If colLancamentoDuplicata.Count > 0 Then
                colLancamentoDuplicata.MoveFirst
                While Not colLancamentoDuplicata.EOF
                    Set voRLancamentoDuplicata = colLancamentoDuplicata.CurrentObject
                    
                    blnExcluir = blnExcluir And AtualizaRegistroOriginal(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup, voRLancamentoDuplicata.nr_parcela_destino)
        
                    blnExcluir = blnExcluir And dao.ExcluirRateioDestino(PagRec, voRLancamentoDuplicata.codigo_nota_destino, voRLancamentoDuplicata.tp_registro_destino, voRLancamentoDuplicata.nr_parcela_destino, LancDup, voRLancamentoDuplicata.cd_empresa_destino)
                    blnExcluir = blnExcluir And dao.Excluir(PagRec, voRLancamentoDuplicata.codigo_nota_destino, voRLancamentoDuplicata.tp_registro_destino, voRLancamentoDuplicata.nr_parcela_destino, LancDup, voRLancamentoDuplicata.cd_empresa_destino)
                    colLancamentoDuplicata.MoveNext
                    Set voRLancamentoDuplicata = Nothing
                Wend
            End If
        End If
    End If
    
    'Verificar o filho.
    Set voRLancamentoDuplicata = New VoRateioLancamentoDuplicata
    Set voRLancamentoDuplicata = dao.CarregaRateioLancamentoDuplicataDestino(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup)
    
    If Not voRLancamentoDuplicata Is Nothing Then
    
        blnExcluir = blnExcluir And AtualizaRegistroOriginal(PagRec, voRLancamentoDuplicata.codigo_nota_origem, voRLancamentoDuplicata.tp_registro_origem, voRLancamentoDuplicata.nr_parcela_origem, voRLancamentoDuplicata.cd_empresa_origem, LancDup, lngParcela)

        blnExcluir = blnExcluir And dao.ExcluirRateioDestino(PagRec, voRLancamentoDuplicata.codigo_nota_destino, voRLancamentoDuplicata.tp_registro_destino, voRLancamentoDuplicata.nr_parcela_destino, LancDup, voRLancamentoDuplicata.cd_empresa_destino)
    End If
     
    Set voRLancamentoDuplicata = Nothing
    ExcluirRateioLancamentoDuplicata = blnExcluir
    Set dao = Nothing
    Set colLancamentoDuplicata = Nothing
    Aplicacao.CommitTransaction
    Aplicacao.Disconnect
Exit Function
err:
    Resume
    ExcluirRateioLancamentoDuplicata = False
    Set dao = Nothing
    Set colLancamentoDuplicata = Nothing
    Aplicacao.RollbackTransaction
    Aplicacao.Disconnect
End Function

Private Function AtualizaRegistroOriginal(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup, ByVal lngParcelaDestino As Long) As Boolean
    Dim voLancDup                       As VoLancamentoDuplicata
    Dim dblValor                        As Double
    Dim dblAcrescimo                    As Double
    Dim dblAbatimento                   As Double
    
    Set voLancDup = Carregar(PagRec, lngNumero, strTipo, lngParcelaDestino, strEmpresa, LancDup)
    If Not voLancDup Is Nothing Then
        dblValor = voLancDup.ValorOriginal
        dblAcrescimo = voLancDup.Acrescimo
        dblAbatimento = voLancDup.Abatimento
    End If
    Set voLancDup = Nothing
        
    Set voLancDup = Carregar(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup)
    If Not voLancDup Is Nothing Then
        voLancDup.ValorOriginal = voLancDup.ValorOriginal + dblValor
        voLancDup.Acrescimo = voLancDup.Acrescimo + dblAcrescimo
        voLancDup.Abatimento = voLancDup.Abatimento + dblAbatimento
    End If
    AtualizaRegistroOriginal = Atualizar(voLancDup)
    Set voLancDup = Nothing
End Function

Private Function LancamentoDuplicataCadastrado(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup) As Boolean
    Dim dao As New DaoLancamentoDuplicata
    Dim vo As VoLancamentoDuplicata
    
    Aplicacao.Connect
    Call dao.init(Aplicacao)
    Set vo = dao.Carregar(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup)
    If Not vo Is Nothing Then
        LancamentoDuplicataCadastrado = True
    End If
    Set dao = Nothing
    Set vo = Nothing
    Aplicacao.Disconnect
End Function

Public Sub validarLancamentoDuplicataExiste(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    Dim obj             As New BizConta
    
        If LancamentoDuplicataCadastrado(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup) Then
            objMensagemErro.mensagem = "O registro de " & IIf(LancDup = Lancamento, "'Lançamento'", "Duplicata") & " já existe no sistema. Favor verificar."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    Set objMensagemErro = Nothing
End Sub


' Valida campos obrigatórios.
Public Function validarCampoObrigatorio(ByVal Emissao As Date, ByVal Vencimento As Date, ByVal Banco As Long, ByVal conta As Long, _
                                        ByVal CentroCusto As Long, ByVal ValorOriginal As Double, ByVal Parcela As Long, _
                                        ByVal OperacaoContabil As Long, ByVal Liberacao As Date, ByVal Empresa As String, _
                                        ByVal Pagamento As Date, ByVal OperacaoContabilBaixa As Long, ByVal Situacao As String, ByVal CodigoNota As String, enumPagRec As enuPagRec, enumLancDup As enuLancDup, ByRef col As Collection) As Boolean
    On Error GoTo err

    'Projeto: #1332 - História: #0 - Desenvolvimento#0 - Moacir Pfau(01/11/2012)
    If enumPagRec = Pagamento And enumLancDup = Duplicata Then
        Call validarCampoObrigatorioCodigoNota(CodigoNota, col)
        Call validarCampoObrigatorioParcela(Parcela, col)
    End If
    Call validarCampoObrigatorioEmpresa(Empresa, col)
    Call validarCampoObrigatorioBanco(Banco, col)
    Call validarCampoObrigatorioConta(conta, col)
    Call validarCampoObrigatorioCentroCusto(CentroCusto, col)
    If IsEmptyDate(Pagamento) Then
        Call validarCampoObrigatorioOperacaoContabil(OperacaoContabil, Pagamento, col)
    End If

    Call validarCampoObrigatorioValorOriginal(ValorOriginal, col)
    Call validarCampoObrigatorioEmissao(Emissao, col)
    Call validarCampoObrigatorioVencimento(Vencimento, col)
    Call validarCampoObrigatorioLiberacao(Liberacao, col)
    'Projeto: #1332 - História: #0 - Desenvolvimento#0 - Moacir Pfau(01/11/2012)
    Call validarCampoObrigatorioSituacao(Situacao, col)
   
    validarCampoObrigatorio = (col.Count = 0)
    Exit Function
err:
    validarCampoObrigatorio = False
End Function

Private Sub validarInformacaoBaixaNaoPoderaSerPreenchidoOperacaoContabilEOperacaoContabilBaixa(ByVal OperacaoContabilBaixa As Long, ByVal OperacaoContabil As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If OperacaoContabilBaixa <> 0 And OperacaoContabil <> 0 Then
        objMensagemErro.mensagem = "O campo 'Operação Contabil' não podera ser preenchido quando é realizado a baixa, somente poderá ser informado o campo 'Operação Contabil para Baixa'."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'01. Verificar se a data de emissão foi prenchida e se é uma data valida.
Private Sub validarCampoObrigatorioEmissao(ByVal Emissao As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If (Not IsDate(Emissao)) Or IsEmptyDate(Emissao) Then  ' Campo sem preenchimento.
        objMensagemErro.mensagem = "O campo 'Emissão' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

Private Sub validarCampoObrigatorioLiberacao(ByVal Liberacao As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If (Not IsDate(Liberacao)) Or IsEmptyDate(Liberacao) Then  ' Campo sem preenchimento.
        objMensagemErro.mensagem = "O campo 'Liberação' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

Private Sub validarCampoObrigatorioSituacao(ByVal Situacao As String, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Len(Trim(Situacao)) = 0 Then ' Campo sem preenchimento.
        objMensagemErro.mensagem = "O campo 'Situação' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'02. Verificar se a data de vencimento foi prenchida e se é uma data valida.
Private Sub validarCampoObrigatorioVencimento(ByVal Vencimento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If (Not IsDate(Vencimento)) Or IsEmptyDate(Vencimento) Then  ' Campo sem preenchimento.
        objMensagemErro.mensagem = "O campo 'Vencimento' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

Private Sub validarCampoObrigatorioEmpresa(ByVal Empresa As String, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Trim(Empresa) = "" Then
        objMensagemErro.mensagem = "O campo 'Empresa' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'05. É obrigatório o preenchimento do banco.
Private Sub validarCampoObrigatorioBanco(ByVal Banco As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Banco = 0 Then
        objMensagemErro.mensagem = "O campo 'Banco' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'06. É obrigatório o preenchimento da conta.
Private Sub validarCampoObrigatorioConta(ByVal conta As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If conta = 0 Then
        objMensagemErro.mensagem = "O campo 'Conta' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'08. É obrigatório o preenchimento do Centro de Custo.
Private Sub validarCampoObrigatorioCentroCusto(ByVal CentroCusto As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If CentroCusto = 0 And ConfigSys.ControlarCentrodeCusto Then
        objMensagemErro.mensagem = "O campo 'Centro de Custo' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'11. É obrigatório o preenchimento do Valor Original.
Private Sub validarCampoObrigatorioValorOriginal(ByVal ValorOriginal As Double, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If ValorOriginal = 0 Then
        objMensagemErro.mensagem = "O campo 'Valor Original' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'12. É obrigatório o preenchimento do parcela.
Private Sub validarCampoObrigatorioParcela(ByVal Parcela As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Parcela = 0 Then
        objMensagemErro.mensagem = "O campo 'Parcela' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'Projeto: #1332 - História: #0 - Desenvolvimento#0 - Moacir Pfau(01/11/2012)
'É obrigatório o preenchimento da conta.
Private Sub validarCampoObrigatorioCodigoNota(ByVal CodigoNota As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If CodigoNota = 0 Then
        objMensagemErro.mensagem = "O campo 'Código/Nota' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'15. Preenchimento é obrigatório da Operação Contábil.
Private Sub validarCampoObrigatorioOperacaoContabil(ByVal OperacaoContabil As Long, ByVal Pagamento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If OperacaoContabil = 0 And ConfigSys.UtilizaIntegracaoContabil Then
        objMensagemErro.mensagem = "O campo 'Operação Contábil' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

Public Function validarInformacaoGeral(ByVal Emissao As Date, ByVal Vencimento As Date, ByVal CentroCusto As Long, ByVal conta As Long, ByVal Liberacao As Date, ByVal OperacaoContabilBaixa As Long, ByVal Pagamento As Date, ByVal Modulo As EnuModulos, ByRef col As Collection) As Boolean
    On Error GoTo err
    
    Call validarCampoVencimentoAnteriorEmissao(Emissao, Vencimento, col)
    Call validarCampoLiberacaoAnteriorVencimento(Vencimento, Liberacao, Pagamento, col)
    Call validarDataLimiteCentroCusto(CentroCusto, Emissao, col)
    Call validarCampoContaAtiva(conta, col)
    Call validarOperacaoContabilSemDataBaixa(OperacaoContabilBaixa, Pagamento, col)
    'Projeto: #7904 - História: #7886 - Desenvolvimento#8103 - João Henrique(21/05/2013)
    'Call validarPermissaoMovimentoConferido(Emissao, Modulo, col)
    Call validarPermissaoMovimentoConferido(Liberacao, Modulo, col)
    Exit Function
err:
    validarInformacaoGeral = False
End Function

'03. Verificar se a data de vencimento é menor que a data de emissão.
Private Sub validarCampoVencimentoAnteriorEmissao(ByVal Emissao As Date, ByVal Vencimento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Not (IsEmptyDate(Emissao) And IsEmptyDate(Vencimento)) Then  '
        If DateDiff("d", Emissao, Vencimento) < 0 Then
            objMensagemErro.mensagem = "A data de 'Vencimento' é anterior a data de 'Emissão'."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

'04. Verificar se a data de liberação do centro de custo está dentro da data limite para utilização.
Private Sub validarDataLimiteCentroCusto(ByVal CentroCusto As Long, ByVal Emissao As String, ByRef col As Collection)
    Dim biz         As New BizCentroCusto
    
    Call biz.validarDataLimiteCentroCusto(CentroCusto, Emissao, col)
    Set biz = Nothing
End Sub

'07. A conta é obrigatorio estar ativa.
Private Sub validarCampoContaAtiva(ByVal conta As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    Dim obj             As New BizConta
    
    If conta > 0 Then   '
        If Not obj.verificaContaAtiva(conta) Then
            objMensagemErro.mensagem = "A 'Conta' não está ativa, somente poderá ser preenchida uma 'Conta Ativa'."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

Private Sub validarOperacaoContabilSemDataBaixa(ByVal OperacaoContabilBaixa As Long, ByVal Pagamento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    Dim obj             As New BizConta
    
    If IsEmptyDate(Pagamento) Then
        If OperacaoContabilBaixa > 0 Then
            objMensagemErro.mensagem = "Quando for preenchido a 'Operação Contábil de Baixa' é obrigatório o preenchimento do campo 'Pagamento'."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

'Caso aconteça a baixa.
Public Function validarInformacaoBaixa(ByVal OperacaoContabilBaixa As Long, ByVal Pagamento As Date, ByVal Vencimento As Date, ByVal Liberacao As Date, ByVal Cheque As String, ByVal Emissao As Date, ByVal OperacaoContabil As Long, ByVal enumLancDup As enuLancDup, ByRef col As Collection) As Boolean
    On Error GoTo err
    
    Call validarCampoPagamentoAnteriorEmissao(Emissao, Pagamento, col)
    Call validarCampoObrigatorioPagamentoquandoChequeForPreenchido(Cheque, Pagamento, col)
    Call validarCampoPagamentoQuantidadeDiasAtraso(Pagamento, Vencimento, col)
    Call validarCampoLiberacaoAnteriorPagamento(Pagamento, Liberacao, col)
    Call validarCampoObrigatorioPagamentoQuandoOperacaoContabilBaixaForPreenchido(OperacaoContabilBaixa, Pagamento, col)
    Call validarCampoObrigatorioOperacaoContabilBaixaQuandoPagamentoForPreenchido(OperacaoContabilBaixa, Pagamento, col)
        
    validarInformacaoBaixa = (col.Count = 0)
    Exit Function
err:
    validarInformacaoBaixa = False
End Function

'16. Verificar se a data de Pagamento não é anterior a emissão.
Private Sub validarCampoPagamentoAnteriorEmissao(ByVal Emissao As Date, ByVal Pagamento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Not (IsEmptyDate(Emissao) And IsEmptyDate(Pagamento)) Then  '
        If DateDiff("d", Emissao, Pagamento) < 0 Then
            'objMensagemErro.mensagem = "A data de 'Pagamento' é anterior a data de 'Emissão'."
            objMensagemErro.mensagem = "A data de 'Pagamento' deverá ser posterior a data de 'Emissão'."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

'17. Se o campo Cheque estiver preenchido é obrigatório esta preenchido a data de pagamento.(Somente Pagamento)
Private Sub validarCampoObrigatorioPagamentoquandoChequeForPreenchido(ByVal Cheque As String, ByVal Pagamento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Trim(Cheque) <> "" Then
        If (Not IsDate(Pagamento)) Or IsEmptyDate(Pagamento) Then
            objMensagemErro.mensagem = "O campo 'Pagamento' é de preenchimento obrigatório, quando a informação do 'Cheque' for preenchida."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

''18. Exibe a mensagem caso a Data de Pagamento seja posteior a Data de Vencimento.
'Private Sub validarCampoPagamentoAnteriorVencimento(ByVal Pagamento As Date, ByVal Vencimento As Date, ByRef col As Collection)
'    Dim objMensagemErro As New clsMensagemErro
'
'    If Not (IsEmptyDate(Pagamento) And IsEmptyDate(Vencimento)) Then  '
'        If DateDiff("d", Pagamento, Vencimento) < 0 Then
'            objMensagemErro.mensagem = "A data de 'Vencimento' é anterior a data de 'Pagamento'."
'            objMensagemErro.status = atencao
'            col.add objMensagemErro
'        End If
'    End If
'    Set objMensagemErro = Nothing
'End Sub

'19. Mostrar a mensagem "A Data de Pagamento informada está em atraso há 999 dia(s). Informe 'Acréscimo' ou 'Multa' se necessário.
Private Sub validarCampoPagamentoQuantidadeDiasAtraso(ByVal Pagamento As Date, ByVal Vencimento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Not (IsEmptyDate(Pagamento) And IsEmptyDate(Vencimento)) Then  '
        If DateDiff("d", Vencimento, Pagamento) > 0 Then
            objMensagemErro.mensagem = "A Data de Pagamento informada está em atraso há " & DateDiff("d", Vencimento, Pagamento) & " dia(s). Informe 'Acréscimo' ou 'Multa' se necessário."
            objMensagemErro.Status = Alerta
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

'20. Verificar se a data de Liberação não pode ser inferior a data de Pagamento.
Private Sub validarCampoLiberacaoAnteriorPagamento(ByVal Pagamento As Date, ByVal Liberacao As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Not (IsEmptyDate(Pagamento) And IsEmptyDate(Liberacao)) Then  '
        If DateDiff("d", Pagamento, Liberacao) < 0 Then
            objMensagemErro.mensagem = "A data de 'Liberação' é anterior a data de 'Pagamento'."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

'Projeto: #1332 - História: #0 - Desenvolvimento#0 - Moacir Pfau(01/11/2012)
'21. Verificar se a data de Liberação não pode ser inferior a data de Vencimento.
Private Sub validarCampoLiberacaoAnteriorVencimento(ByVal Vencimento As Date, ByVal Liberacao As Date, ByVal Pagamento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro

    If Not (IsEmptyDate(Vencimento) And IsEmptyDate(Liberacao)) Then  '
        If IsEmptyDate(Pagamento) Then
            If DateDiff("d", Vencimento, Liberacao) < 0 Then
                objMensagemErro.mensagem = "A data de 'Liberação' é anterior a data de 'Vencimento'."
                objMensagemErro.Status = atencao
                col.add objMensagemErro
            End If
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

'22. Caso seja preenchido a operação contábil de baixa será necessário informar a data de pagamento.
Private Sub validarCampoObrigatorioPagamentoQuandoOperacaoContabilBaixaForPreenchido(ByVal OperacaoContabilBaixa As Long, ByVal Pagamento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If ConfigSys.UtilizaIntegracaoContabil Then
        If OperacaoContabilBaixa = 0 Then
            If (Not IsDate(Pagamento)) Or IsEmptyDate(Pagamento) Then
                objMensagemErro.mensagem = "O campo 'Pagamento' é de preenchimento obrigatório, quando a informação da 'Operação Contábil de Baixa' for preenchida."
                objMensagemErro.Status = atencao
                col.add objMensagemErro
            End If
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

'23. Quando pagamento esta preenchido é necessário informar a Operação de Baixa
Private Sub validarCampoObrigatorioOperacaoContabilBaixaQuandoPagamentoForPreenchido(ByVal OperacaoContabilBaixa As Long, ByVal Pagamento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If ConfigSys.UtilizaIntegracaoContabil Then
        If Not (IsEmptyDate(Pagamento)) Then   '
            If OperacaoContabilBaixa = 0 Then
                objMensagemErro.mensagem = "O campo 'Operação Contábil de Baixa' é de preenchimento obrigatório, quando a informação de 'Pagamento' for preenchida."
                objMensagemErro.Status = atencao
                col.add objMensagemErro
            End If
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

Private Sub validarCampoObrigatorioOperacaoContabilBaixaQuandoPagamentoNaoForPreenchido(ByVal OperacaoContabilBaixa As Long, ByVal Pagamento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If OperacaoContabilBaixa <> 0 And (IsEmptyDate(Pagamento)) And ConfigSys.UtilizaIntegracaoContabil Then
        objMensagemErro.mensagem = "O campo 'Operação Contabil Baixa' não podera ser preenchido quando não é realizado a baixa, somente poderá ser informado na Baixa."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

Public Function ProximoCheque(ByVal Banco As Long) As Long
    Dim biz As New bizCheque
    
    ProximoCheque = biz.ProximoCheque(Banco)
    Set biz = Nothing
End Function

Public Function ProcessoRateioLancamentoDuplicata(ByVal col As colRateio, ByVal vo As VoLancamentoDuplicata) As Boolean
    Dim objVoRateio                     As VoRateio
    Dim i                               As Integer
    Dim dao                             As DaoLancamentoDuplicata
    Dim blnProcesso                     As Boolean
    Dim voRLancamentoDuplicata          As VoRateioLancamentoDuplicata
    Dim intParcelaOriginal              As Integer
    Dim dblValorOriginal                As Double
    Dim dblValorAcrescimo               As Double
    Dim dblValorAbatimento              As Double
    Dim intQuantidadeAcumulado          As Integer
    Dim dblValorAcumulado               As Double
    Dim dblAbatimentoAcumulado          As Double
    Dim dblAcrescimoAcumulado           As Double
        
    intQuantidadeAcumulado = 0
    dblValorAcumulado = 0
    dblAbatimentoAcumulado = 0
    dblAcrescimoAcumulado = 0

On Error GoTo err
    i = 0
    Aplicacao.Connect
    Aplicacao.BeginTransaction
    Set dao = New DaoLancamentoDuplicata
    Call dao.init(Aplicacao)
    blnProcesso = True
    If col.Count > 0 Then
        col.MoveFirst
        While Not col.EOF
            If blnProcesso Then
                i = i + 1
                Set objVoRateio = col.CurrentObject
                If i = 1 Then
                    dblValorOriginal = vo.ValorOriginal
                    dblValorAbatimento = vo.Abatimento
                    dblValorAcrescimo = vo.Acrescimo
                    vo.conta = objVoRateio.conta
                    vo.Centro = objVoRateio.Centro
                    If objVoRateio.Valor > 0 Then
                        vo.ValorOriginal = objVoRateio.Valor
                        vo.Abatimento = objVoRateio.Abatimento
                        vo.Acrescimo = objVoRateio.Acrescimo
                    Else
                        intQuantidadeAcumulado = intQuantidadeAcumulado + 1
                        vo.ValorOriginal = fRateioPercentualparaValor(dblValorOriginal, objVoRateio.Percentual, dblValorAcumulado, col.Count, intQuantidadeAcumulado)
                        vo.Abatimento = fRateioPercentualparaValor(dblValorAbatimento, objVoRateio.Percentual, dblAbatimentoAcumulado, col.Count, intQuantidadeAcumulado)
                        vo.Acrescimo = fRateioPercentualparaValor(dblValorAcrescimo, objVoRateio.Percentual, dblAcrescimoAcumulado, col.Count, intQuantidadeAcumulado)
                    End If
                    
                    blnProcesso = dao.Atualizar(vo)
                    intParcelaOriginal = vo.Parcela
                Else
                    vo.conta = objVoRateio.conta
                    vo.Centro = objVoRateio.Centro
                    If objVoRateio.Valor > 0 Then
                        vo.ValorOriginal = objVoRateio.Valor
                        vo.Abatimento = objVoRateio.Abatimento
                        vo.Acrescimo = objVoRateio.Acrescimo
                    Else
                        intQuantidadeAcumulado = intQuantidadeAcumulado + 1
                        vo.ValorOriginal = fRateioPercentualparaValor(dblValorOriginal, objVoRateio.Percentual, dblValorAcumulado, col.Count, intQuantidadeAcumulado)
                        vo.Abatimento = fRateioPercentualparaValor(dblValorAbatimento, objVoRateio.Percentual, dblAbatimentoAcumulado, col.Count, intQuantidadeAcumulado)
                        vo.Acrescimo = fRateioPercentualparaValor(dblValorAcrescimo, objVoRateio.Percentual, dblAcrescimoAcumulado, col.Count, intQuantidadeAcumulado)
                    End If
                    vo.Parcela = dao.fProximoParcela(vo.Codigo_Nota, vo.LancDup, vo.PagRec)
                    blnProcesso = dao.Gravar(vo)
                    
                    Set voRLancamentoDuplicata = New VoRateioLancamentoDuplicata
                    voRLancamentoDuplicata.cd_centro = objVoRateio.Centro
                    voRLancamentoDuplicata.Cd_conta = objVoRateio.conta
                    voRLancamentoDuplicata.codigo_nota_origem = vo.Codigo_Nota
                    voRLancamentoDuplicata.codigo_nota_destino = vo.Codigo_Nota
                    voRLancamentoDuplicata.dt_pagamento = vo.Pagamento
                    voRLancamentoDuplicata.nr_parcela_origem = intParcelaOriginal
                    voRLancamentoDuplicata.nr_parcela_destino = vo.Parcela
                    voRLancamentoDuplicata.pag_rec_origem = vo.PagRec
                    voRLancamentoDuplicata.pag_rec_destino = vo.PagRec
                    voRLancamentoDuplicata.vl_valor = vo.ValorOriginal
                    If vo.LancDup = Duplicata Then
                        voRLancamentoDuplicata.tp_registro_origem = vo.Tipo
                        voRLancamentoDuplicata.tp_registro_destino = vo.Tipo
                        voRLancamentoDuplicata.cd_empresa_origem = vo.Empresa
                        voRLancamentoDuplicata.cd_empresa_destino = vo.Empresa
                    End If
                    blnProcesso = blnProcesso And dao.GravarRateioLancamentoDuplicata(voRLancamentoDuplicata, vo.LancDup)
                    Set voRLancamentoDuplicata = Nothing
                    
                End If
            End If
            col.MoveNext
        Wend
    End If
    
    Set objVoRateio = Nothing
    Set dao = Nothing
    Set vo = Nothing
    If blnProcesso Then
        Aplicacao.CommitTransaction
    Else
        Aplicacao.RollbackTransaction
    End If
    Aplicacao.Disconnect
    ProcessoRateioLancamentoDuplicata = blnProcesso
    Exit Function
err:
    Set objVoRateio = Nothing
    Set dao = Nothing
    Aplicacao.RollbackTransaction
    Aplicacao.Disconnect
    ProcessoRateioLancamentoDuplicata = False
End Function

'***'09. Caso seja preenchido o campo cheque devera ser informar um número do banco.
'10. Verifica se a carteira esta cadastrada no Banco (não é utilizado pela rotina nova)
'13. Valida o método para empresa bloqueada(fEmpresaBloqueada) (sera necessário inserir na tela, devido o tratamento da regra antiga)
'***'14. Para Lançamentos é necessário informar a Operação Contábil


'Projeto: #1203 - História: #10564 - Desenvolvimento#10575 - João Henrique(05/04/2012)
Public Function validarCampoObrigatorioRateio(ByVal conta As Long, ByVal CentroCusto As Long, ByVal Porcentagem As Double, ByVal Valor As Double, ByRef col As Collection) As Boolean
    On Error GoTo err
    
    Call validarCampoObrigatorioRateioCentroCusto(CentroCusto, col)
    Call validarCampoObrigatorioRateioConta(conta, col)
    Call validarCampoObrigatorioRateioPercentualValor(Porcentagem, Valor, col)
        
    validarCampoObrigatorioRateio = (col.Count = 0)
    Exit Function
err:
    validarCampoObrigatorioRateio = False
End Function

'É obrigatório o preenchimento do Centro de Custo.
Private Sub validarCampoObrigatorioRateioCentroCusto(ByVal CentroCusto As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If CentroCusto = 0 Then
        objMensagemErro.mensagem = "O campo 'Centro de Custo' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'É obrigatório o preenchimento da conta.
Private Sub validarCampoObrigatorioRateioConta(ByVal conta As Long, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If conta = 0 Then
        objMensagemErro.mensagem = "O campo 'Conta Financeira' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'É obrigatório o preenchimento da porcentagem ou valor.
Private Sub validarCampoObrigatorioRateioPercentualValor(ByVal Porcentagem As Double, ByVal Valor As Double, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Porcentagem = 0 And Valor = 0 Then
        objMensagemErro.mensagem = "O campo 'Porcentagem' ou o campo 'Valor' é de preenchimento obrigatório."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'Projeto: #1203 - História: #10564 - Desenvolvimento#10575 - João Henrique(05/04/2012)
Public Function validarCampoDiversosRateio(ByVal Porcentagem As Double, ByVal ValorItem As Double, ByVal Acrescimo As Double, _
                                           ByVal Abatimento As Double, ByVal TotalPercentual As Double, ByVal ValorTotalRateio As Double, _
                                           ByVal conta As Double, ByVal Centro As Double, ByVal Col_Rateio As colRateio, _
                                           ByVal ValorTotal As Double, ByVal ValorItemLiquido As Double, ByRef col As Collection) As Boolean
    
    On Error GoTo err
        
    Call validarCampoDiversosRateioPorcentagem_Valor(Porcentagem, ValorItem, col)
    Call validarCampoDiversosRateioPorcentagem_Acrescimo(Porcentagem, Acrescimo, col)
    Call validarCampoDiversosRateioPorcentagem_Abatimento(Porcentagem, Abatimento, col)
    Call validarCampoDiversosRateioPorcentagem(Porcentagem, col)
    Call validarCampoDiversosRateioSomenteValor(ValorTotalRateio, Porcentagem, col)
    Call validarCampoDiversosRateioSomentePercentual(TotalPercentual, ValorItem, col)
    Call validarCampoDiversosRateioContaCentroDuplicidade(conta, Centro, Col_Rateio, col)
    
    'Projeto: #1203 - História: #10564 - Problema#11881 - João Henrique(12/04/2012)
    Call validarCampoDiversosRateioTotalPorcentagem(Porcentagem, TotalPercentual, col)
    Call validarCampoDiversosRateioTotalValor(Col_Rateio.ValorTotalRateio, ValorTotal, ValorItemLiquido, col)
    Call validarCampoDiversosRateioValorItemLiquido(ValorItemLiquido, col)
    
    validarCampoDiversosRateio = (col.Count = 0)
    Exit Function
err:
    validarCampoDiversosRateio = False
End Function

'Verificar se o Campo Porcentagem e Valor estão preenchidos.
Private Sub validarCampoDiversosRateioPorcentagem_Valor(ByVal Porcentagem As Double, ByVal ValorItem As Double, ByRef col As Collection)
    Dim objMensagemErro         As New clsMensagemErro
        
    If Porcentagem > 0 And ValorItem > 0 Then
        objMensagemErro.mensagem = "O campo 'Porcentagem' e o campo 'Valor' não podem ser preenchidos juntos."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'Verificar se somente o Campo Porcentagem pode ser utilizado.
Private Sub validarCampoDiversosRateioSomentePercentual(ByVal TotalPercentual As Double, ByVal ValorItem As Double, ByRef col As Collection)
    Dim objMensagemErro         As New clsMensagemErro
        
    If TotalPercentual > 0 And ValorItem > 0 Then
        objMensagemErro.mensagem = "Somente pode ser usado 'Porcentagem'."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'Verificar se somente o Campo Valor pode ser utilizado.
Private Sub validarCampoDiversosRateioSomenteValor(ByVal ValorTotalRateio As Double, ByVal Percentual As Double, ByRef col As Collection)
    Dim objMensagemErro         As New clsMensagemErro
        
    If ValorTotalRateio > 0 And Percentual > 0 Then
        objMensagemErro.mensagem = "Somente pode ser usado 'Valor'."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    Set objMensagemErro = Nothing
End Sub

'Verificar se o Campo Porcentagem e Acrecimo estão preenchidos.
Private Sub validarCampoDiversosRateioPorcentagem_Acrescimo(ByVal Porcentagem As Double, ByVal Acrescimo As Double, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
        If Porcentagem > 0 And Acrescimo > 0 Then
            objMensagemErro.mensagem = "O campo 'Porcentagem' e o campo 'Acréscimo' não podem ser preenchidos juntos."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    Set objMensagemErro = Nothing
End Sub

'Verificar se o Campo Porcentagem e Abatimento estão preenchidos.
Private Sub validarCampoDiversosRateioPorcentagem_Abatimento(ByVal Porcentagem As Double, ByVal Abatimento As Double, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
        If Porcentagem > 0 And Abatimento > 0 Then
            objMensagemErro.mensagem = "O campo 'Porcentagem' e o campo 'Abatimento' não podem ser preenchidos juntos."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    Set objMensagemErro = Nothing
End Sub

'Verificar se o Campo Porcentagem é maior que 100%.
Private Sub validarCampoDiversosRateioPorcentagem(ByVal Porcentagem As Double, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
        If Not Porcentagem <= 100 Then
            objMensagemErro.mensagem = "O Valor do campo 'Porcentagem' é maior que 100%."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    Set objMensagemErro = Nothing
End Sub

'Verificar se há duplicidade no Campo Conta e Centro de Custo.
Private Sub validarCampoDiversosRateioContaCentroDuplicidade(ByVal conta As Double, ByVal Centro As Double, ByVal Col_Rateio As colRateio, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Col_Rateio.Count > 0 Then
        If ExisteDuplicidadeContaCentro(conta, Centro, Col_Rateio) Then
            objMensagemErro.mensagem = "Conta e Centro de custo já utilizados."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

Private Function ExisteDuplicidadeContaCentro(ByVal conta As Double, ByVal Centro As Double, ByVal Col_Rateio As colRateio) As Boolean
    Dim vo As VoRateio
    
    If Col_Rateio.Count > 0 Then
        Col_Rateio.MoveFirst
        While Not Col_Rateio.EOF
            Set vo = Col_Rateio.CurrentObject
            If vo.conta = conta And vo.Centro = Centro Then
                ExisteDuplicidadeContaCentro = True
                Exit Function
            End If
            Col_Rateio.MoveNext
        Wend
    End If
End Function

Private Function fRateioPercentualparaValor(ByVal dblValorTotal As Double, ByVal dblPercentual As Double, ByRef dblValorAcumulado As Double, ByVal lngTotalItens As Long, ByVal intSumItemAtual As Integer) As Double
    Dim dblVlrCalculado         As Double
    If lngTotalItens = intSumItemAtual Then
        fRateioPercentualparaValor = dblValorTotal - dblValorAcumulado
    Else
        dblVlrCalculado = dblValorTotal * dblPercentual / 100
        dblValorAcumulado = dblValorAcumulado + dblVlrCalculado
        fRateioPercentualparaValor = dblVlrCalculado
    End If
End Function

Public Function mensagemAposBaixaInformativo(ByVal Pagamento As Date, ByVal Vencimento As Date, ByVal Empresa As String, ByRef col As Collection) As Boolean
    On Error GoTo err
    
    Call mensagemAposBaixaInformativoAtraso(Pagamento, Vencimento, col)
    Call mensagemAposBaixaInformativoEmpresa(Empresa, col)
    
    mensagemAposBaixaInformativo = True
    Exit Function
err:
    mensagemAposBaixaInformativo = False
End Function

Private Sub mensagemAposBaixaInformativoAtraso(ByVal Pagamento As Date, ByVal Vencimento As Date, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    
    If Not (IsEmptyDate(Vencimento) And IsEmptyDate(Pagamento)) And Pagamento <> "00:00:00" Then
        If DateDiff("d", Vencimento, Pagamento) > 0 Then
            objMensagemErro.mensagem = "A Data de Pagamento informada está em atraso há " & DateDiff("d", Vencimento, Pagamento) & " dia(s)." & _
            objMensagemErro.Status = branco
            col.add objMensagemErro
            objMensagemErro.mensagem = "Informe 'Acréscimo' ou 'Multa' se necessário."
            objMensagemErro.Status = Alerta
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub

Private Sub mensagemAposBaixaInformativoEmpresa(ByVal Empresa As String, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
    Dim objEmpresa      As New CEmpresas
    
    If objEmpresa.CarregarRegistro(Empresa) Then
        If Trim(objEmpresa.ObsFinanceira) <> "" Then
            objMensagemErro.mensagem = objEmpresa.ObsFinanceira
            objMensagemErro.Status = Alerta
            col.add objMensagemErro
        End If
    End If
    Set objMensagemErro = Nothing
End Sub


'Verificar se a soma do valor total de Porcentagem é > 100.
Private Sub validarCampoDiversosRateioTotalPorcentagem(ByVal Porcentagem As Double, ByVal TotalPercentual As Double, ByRef col As Collection)
    Dim objMensagemErro As New clsMensagemErro
        If Porcentagem > 0 Then
            TotalPercentual = TotalPercentual + Porcentagem
            If TotalPercentual > 100 Then
                objMensagemErro.mensagem = "A Soma da 'Porcentagem' não pode ser maior que 100%."
                objMensagemErro.Status = atencao
                col.add objMensagemErro
            End If
        End If
    Set objMensagemErro = Nothing
End Sub

'Verificar se a soma total do valor de Rateio é > que o Valor total.
Private Sub validarCampoDiversosRateioTotalValor(ByVal ValorTotalRateio As Double, ByVal ValorTotal As Double, ByVal ValorItemLiquido As Double, ByRef col As Collection)
    Dim objMensagemErro     As New clsMensagemErro

        If ValorItemLiquido > 0 Then
            ValorTotalRateio = ValorTotalRateio + ValorItemLiquido
            'Projeto: #0 - História: #18915 - Desenvolvimento#18944 - Moacir Pfau(10/12/2012)
            If Round(ValorTotalRateio, 8) > Round(ValorTotal, 8) Then
                objMensagemErro.mensagem = "A Soma do 'Valor' não pode ser maior que o Valor Original."
                objMensagemErro.Status = atencao
                col.add objMensagemErro
            End If
        End If
    Set objMensagemErro = Nothing
End Sub

'Verificar se a soma total do valor de Rateio é < que Zero.
Private Sub validarCampoDiversosRateioValorItemLiquido(ByVal ValorItemLiquido As Double, ByRef col As Collection)
    Dim objMensagemErro     As New clsMensagemErro
        
        If ValorItemLiquido < 0 Then
            objMensagemErro.mensagem = "A Soma Total do 'Valor' não pode ser menor que o Zero."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
       
    Set objMensagemErro = Nothing
End Sub

Public Function validarPermissaoExclusao(ByVal Codigo As String, ByVal Parcela As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal LancDup As enuLancDup, ByVal PagRec As enuPagRec, ByVal Data As Date, ByVal Modulo As EnuModulos, ByRef col As Collection) As Boolean
    On Error GoTo err
    
    Call validarPermissaoExclusaoGeracaoTitulo(Codigo, Parcela, Tipo, Empresa, LancDup, PagRec, col)
    Call validarPermissaoExclusaoPertenceNota(Codigo, Parcela, Tipo, Empresa, LancDup, PagRec, col)
    Call validarPermissaoExclusaoPertencePedido(Codigo, Parcela, Tipo, Empresa, LancDup, PagRec, col)
    Call validarPermissaoMovimentoConferido(Data, Modulo, col)
    'Mauricio - 09/04/2015 - PT 217865
    Call validarPermissaoExclusaoSinalNegocio(Codigo, Parcela, Tipo, Empresa, LancDup, PagRec, col)
    'Mauricio - 09/04/2015 - Projeto: #67649 Problema #73672
    Call validarPermissaoExclusaoCTE(Codigo, Parcela, Tipo, Empresa, LancDup, PagRec, col)
    
    Exit Function
err:
    validarPermissaoExclusao = False
End Function

Private Sub validarPermissaoExclusaoGeracaoTitulo(ByVal Codigo As String, ByVal Parcela As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal LancDup As enuLancDup, ByVal PagRec As enuPagRec, ByRef col As Collection)
    Dim biz                 As New BizGeracaoTitulo
    Dim objMensagemErro     As New clsMensagemErro
    If LancDup = Duplicata Then
        If biz.GeracaoTituloExiste(Codigo, Parcela, Tipo, Empresa, PagRec) Then
            objMensagemErro.mensagem = "Duplicata gerada pelo processo de 'Geração de Títulos', não pode ser excluida por esta rotina."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    
    Set biz = Nothing
    Set objMensagemErro = Nothing
End Sub

Private Sub validarPermissaoExclusaoPertenceNota(ByVal Codigo As String, ByVal Parcela As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal LancDup As enuLancDup, ByVal PagRec As enuPagRec, ByRef col As Collection)
    Dim biz                 As New BizGeracaoTitulo
    Dim objMensagemErro     As New clsMensagemErro
    If LancDup = Duplicata Then
        'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
        'Se a parcela é uma baixa parcial, o sistema deverá permitir a exclusão
        If PertenceNota(Codigo, Tipo, PagRec, Empresa) And BaixaParcial(Codigo, Parcela, LancDup) = 0 Then
            objMensagemErro.mensagem = "Duplicata gerada pela 'Nota Fiscal', não pode ser excluida por esta rotina."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    
    Set biz = Nothing
    Set objMensagemErro = Nothing
End Sub

Public Function PertenceNota(ByVal Codigo As String, ByVal Tipo As String, ByVal PagRec As enuPagRec, ByVal Empresa As String) As Boolean
    Dim selCmd    As IDBSelectCommand
    Dim rdResult  As IDBReader
    Dim strTabela As String
    
On Error GoTo error_handler
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "Número"
        
        If PagRec = Pagamento Then
            strTabela = "[Notas Fiscais de Entrada]"
        Else
            strTabela = "[Notas Fiscais de Saída]"
        End If
        
        .Table.TableName = strTabela
        
        Call .Filter.Append("Número = @pNumero")
        Call .Parameters.add(.CreateParameter("@pNumero", Codigo, dbFieldTypeDouble))
        
        Call .Filter.Append("[Tipo de Registro] = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", Tipo, dbFieldTypeString, 30))
        
        If PagRec = Pagamento Then
            Call .Filter.Append("Fornecedor = @pFornecedor")
            Call .Parameters.add(.CreateParameter("@pFornecedor", Empresa, dbFieldTypeString, 15))
        Else
            Call .Filter.Append("Empresa = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", Empresa, dbFieldTypeString, 15))
        End If
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    PertenceNota = Not rdResult.EOF
    rdResult.CloseReader
    Set selCmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
    
    Exit Function
error_handler:
    FinallyConnection Aplicacao
    err.Clear
    PertenceNota = False
End Function
Private Sub validarPermissaoExclusaoSinalNegocio(ByVal Codigo As Long, ByVal Parcela As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal LancDup As enuLancDup, ByVal PagRec As enuPagRec, ByRef col As Collection)
    Dim biz                 As New BizGeracaoTitulo
    Dim objMensagemErro     As New clsMensagemErro
    If LancDup = Duplicata Then
        If PertenceSinalNegocio(Codigo, Tipo, Empresa, Parcela) Then
            objMensagemErro.mensagem = "Duplicata gerada através de um sinal de negócio, não pode ser excluida por esta rotina."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    
    Set biz = Nothing
    Set objMensagemErro = Nothing
End Sub

Public Function PertenceSinalNegocio(ByVal Codigo As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal Parcela As Long) As Boolean
    Dim selCmd    As IDBSelectCommand
    Dim rdResult  As IDBReader
    Dim strTabela As String
    
On Error GoTo error_handler
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        
        .SelectClause = "NUMPED"
        
        strTabela = "SinaisNegocio"
                
        .Table.TableName = strTabela
        
        Call .Filter.Append("NUMPED = @pNumero")
        Call .Parameters.add(.CreateParameter("@pNumero", Codigo, dbFieldTypeLong))
        
        Call .Filter.Append("TIPREG = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", Tipo, dbFieldTypeString, 30))
        
        Call .Filter.Append("FORNEC = @pFornecedor")
        Call .Parameters.add(.CreateParameter("@pFornecedor", Empresa, dbFieldTypeString, 30))
        
        Call .Filter.Append("NUMPAR = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", Parcela, dbFieldTypeLong, 5))
        

    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    PertenceSinalNegocio = Not rdResult.EOF
    rdResult.CloseReader
    Set selCmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
    
    Exit Function
error_handler:
    FinallyConnection Aplicacao
    err.Clear
    PertenceSinalNegocio = False
End Function
Private Sub validarPermissaoExclusaoCTE(ByVal Codigo As Long, ByVal Parcela As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal LancDup As enuLancDup, ByVal PagRec As enuPagRec, ByRef col As Collection)
    Dim biz                 As New BizGeracaoTitulo
    Dim objMensagemErro     As New clsMensagemErro
    If LancDup = Duplicata Then
        If PertenceCTE(Codigo, Tipo, Empresa, Parcela, PagRec) Then
            objMensagemErro.mensagem = "Duplicata gerada através de um Conhecimento de Frete, não pode ser excluida por esta rotina."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    
    Set biz = Nothing
    Set objMensagemErro = Nothing
End Sub


Public Function PertenceCTE(ByVal Codigo As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal Parcela As Long, ByVal PagRec As enuPagRec) As Boolean
    Dim selCmd    As IDBSelectCommand
    Dim rdResult  As IDBReader
    Dim strTabela As String
    
On Error GoTo error_handler
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "nr_conhecimento"
        
        If PagRec = Pagamento Then
            strTabela = "FreteEntradaDuplicata"
        Else
            strTabela = "FreteSaidaDuplicata"
        End If
                        
        .Table.TableName = strTabela
        
        Call .Filter.Append("nr_conhecimento = @pNumero")
        Call .Parameters.add(.CreateParameter("@pNumero", Codigo, dbFieldTypeLong))
        
        Call .Filter.Append("tp_registro = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", Tipo, dbFieldTypeString, 30))
        
        Call .Filter.Append("empresa = @pEmpresa")
        Call .Parameters.add(.CreateParameter("@pEmpresa", Empresa, dbFieldTypeString, 30))
        
        Call .Filter.Append("parcela = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", Parcela, dbFieldTypeLong, 5))
        

    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    PertenceCTE = Not rdResult.EOF
    rdResult.CloseReader
    Set selCmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
    
    Exit Function
error_handler:
    FinallyConnection Aplicacao
    err.Clear
    PertenceCTE = False
End Function

Private Sub validarPermissaoExclusaoPertencePedido(ByVal Codigo As Long, ByVal Parcela As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal LancDup As enuLancDup, ByVal PagRec As enuPagRec, ByRef col As Collection)
    Dim biz                 As New BizGeracaoTitulo
    Dim objMensagemErro     As New clsMensagemErro
    If LancDup = Duplicata Then
        If PertencePedido(Codigo, Tipo, PagRec, Empresa) Then
            objMensagemErro.mensagem = "Duplicata gerada pelo 'Pedido', não pode ser excluida por esta rotina."
            objMensagemErro.Status = atencao
            col.add objMensagemErro
        End If
    End If
    
    Set biz = Nothing
    Set objMensagemErro = Nothing
End Sub

Public Function PertencePedido(ByVal Codigo As String, ByVal Tipo As String, ByVal PagRec As enuPagRec, ByVal Empresa As String) As Boolean
    Dim selCmd    As IDBSelectCommand
    Dim rdResult  As IDBReader
    Dim strTabela As String
    
On Error GoTo error_handler
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "Número"
        
        If PagRec = Pagamento Then
            strTabela = "[Pedidos de Compra]"
        Else
            strTabela = "[Pedidos de Venda]"
        End If
        
        .Table.TableName = strTabela
        
        Call .Filter.Append("Número = @pNumero")
        Call .Parameters.add(.CreateParameter("@pNumero", Codigo, dbFieldTypeDouble))
        
        Call .Filter.Append("[Tipo de Registro] = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", Tipo, dbFieldTypeString, 30))
        
        If PagRec = Pagamento Then
            Call .Filter.Append("Fornecedor = @pFornecedor")
            Call .Parameters.add(.CreateParameter("@pFornecedor", Empresa, dbFieldTypeString, 15))
        Else
            Call .Filter.Append("Empresa = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", Empresa, dbFieldTypeString, 15))
        End If
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    PertencePedido = Not rdResult.EOF
    rdResult.CloseReader
    Set selCmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
    
    Exit Function
error_handler:
    FinallyConnection Aplicacao
    err.Clear
    PertencePedido = False
End Function

Private Sub validarPermissaoMovimentoConferido(ByVal Data As Date, ByVal Modulo As EnuModulos, ByRef col As Collection)
    Dim biz                 As New BizMovimentoConferido
    Dim objMensagemErro     As New clsMensagemErro
    If biz.ValidaMovimentoConferido(Data, Modulo) Then
        objMensagemErro.mensagem = "Movimento conferido foi definido, não pode ser realizado esta operação."
        objMensagemErro.Status = atencao
        col.add objMensagemErro
    End If
    
    Set biz = Nothing
    Set objMensagemErro = Nothing
End Sub

Public Function ExisteRateioLancamentoDuplicataOrigem(ByVal PagRec As enuPagRec, ByVal lngNumero As String, ByVal strTipo As String, ByVal lngParcela As Long, ByVal strEmpresa As String, ByVal LancDup As enuLancDup) As Boolean
    Dim dao As New DaoLancamentoDuplicata
    Aplicacao.Connect
    Call dao.init(Aplicacao)
    ExisteRateioLancamentoDuplicataOrigem = dao.ExisteRateioLancamentoDuplicataOrigem(PagRec, lngNumero, strTipo, lngParcela, strEmpresa, LancDup)
    Set dao = Nothing
    Aplicacao.Disconnect
End Function

'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
Public Function BaixaParcial(ByVal Nota As String, ByVal Parcela As Long, ByVal Tabela As enuLancDup, Optional ByRef NovoAbatimento As Variant) As Integer
'------------------------------------------------------------------------------------------------------------
'RETORNA: Nr da parcela que originou a baixa parcial e o valor do abatimento sem a parcela em questão.
'ENTRADA: Nr da nota, Nr da parcela e codigo para identificar a Tabela
'--------------------------- QUERY EXECUTADA ----------------------------------------------------------------
'SELECT POrig.Parcela,
'       POrig.Abatimento - (D.[Valor Original] + D.Acréscimo - D.Abatimento)  Novo_Abatimento
'FROM Duplicatas D JOIN Duplicatas POrig ON d.parc_origem_baixa = Porig.Parcela
'                                            and d.Nota = POrig.Nota
'                                            and D.Tipo = POrig.Tipo
'                                            and D.PagRec = POrig.PagRec
'                                            and D.Empresa = POrig.Empresa
'WHERE d.nota = <nota> And d.parcela = <parcela>
'------------------------------------------------------------------------------------------------------------

    Dim selCmd    As IDBSelectCommand
    Dim rdResult  As IDBReader
    Dim strTabela As String
    Dim objJoin   As CDBJoin
    
On Error GoTo error_handler
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        
        .SelectClause = "POrig.Parcela, POrig.Abatimento - (D.[Valor Original] + D.Acréscimo - D.Abatimento) AS NovoAbat"
        strTabela = IIf(Tabela = Duplicata, "[Duplicatas]", "[Lançamentos]")
        
        .Table.TableName = strTabela
        .Table.Alias = "D"
        Set objJoin = New CDBJoin
        objJoin.init
        objJoin.JoinType = dbJoinTypeInner
        objJoin.LeftTable.TableName = strTabela           'Duplicata
        objJoin.LeftTable.Alias = "D"
        objJoin.RightTable.TableName = strTabela     'Parcela de Origem
        objJoin.RightTable.Alias = "POrig"
        Call objJoin.AddJoinField("parc_origem_baixa", "Parcela")
        If Tabela = Duplicata Then
            Call objJoin.AddJoinField("Nota", "Nota")
        Else
            Call objJoin.AddJoinField("Código", "Código")
        End If
        Call objJoin.AddJoinField("Tipo", "Tipo")
        Call objJoin.AddJoinField("PagRec", "PagRec")
        Call objJoin.AddJoinField("Empresa", "Empresa")
        Call .AddJoin(objJoin)
                
        If Tabela = Duplicata Then
            Call .Filter.Append(objJoin.LeftTable.Alias & ".Nota = @pNota")
        Else
            Call .Filter.Append(objJoin.LeftTable.Alias & ".Código = @pNota")
        End If
        Call .Parameters.add(.CreateParameter("@pNota", Nota, dbFieldTypeDouble))
        
        Call .Filter.Append(objJoin.LeftTable.Alias & ".Parcela = @pParc")
        Call .Parameters.add(.CreateParameter("@pParc", Parcela, dbFieldTypeLong))
        
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    
    If rdResult.EOF Then
        BaixaParcial = 0
    Else
        BaixaParcial = rdResult.GetInteger("Parcela")
        If Not IsMissing(NovoAbatimento) Then
            NovoAbatimento = rdResult.GetDouble("NovoAbat")
        End If
    End If
    
    rdResult.CloseReader
    Set selCmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
    
    Exit Function
error_handler:
    FinallyConnection Aplicacao
    err.Clear
    BaixaParcial = False
Aplicacao.Disconnect
End Function

'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
'No momento de uma exclusão de uma parcela identificada como baixa parcial, o valor da parcela é diminuído do valor do Abatimento da parcela original
Public Function DevolveValorParaParcOrig(ByVal Nota As String, ByVal Parcela As Long, ByVal strEmpresa As String, ByVal strTipo As String, PagRec As enuPagRec, Tabela As enuLancDup) As Boolean
    Dim cmd As IDBUpdateCommand
    Dim dblNovoAbat As Double
    Dim intParcOrig As Integer
    Dim strPagRec As String
On Error GoTo err_Handler
        
    intParcOrig = BaixaParcial(Nota, Parcela, Tabela, dblNovoAbat)
    If intParcOrig = 0 Then
        DevolveValorParaParcOrig = False
        Exit Function
    End If
    strPagRec = IIf(PagRec = Pagamento, "P", "R")
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateUpdateCommand
    With cmd
        .Table = IIf(Tabela = Duplicata, "[Duplicatas]", "[Lançamentos]")
        
        If Tabela = Duplicata Then
            Call .Filter.Append("Nota = @pNota")
        Else
            Call .Filter.Append("Código = @pNota")
        End If
        Call .Parameters.add(.CreateParameter("@pNota", Nota, dbFieldTypeString))
        Call .Filter.Append("Parcela = @pParcOrig")
        Call .Parameters.add(.CreateParameter("@pParcOrig", intParcOrig, dbFieldTypeInt))
        Call .Filter.Append("PagRec = @pPR")
        Call .Parameters.add(.CreateParameter("@pPR", strPagRec, dbFieldTypeString))
        Call .Filter.Append("Empresa = @pEmp")
        Call .Parameters.add(.CreateParameter("@pEmp", strEmpresa, dbFieldTypeString))
        Call .Filter.Append("Tipo = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", strTipo, dbFieldTypeString))
        
        Call .AddValue("Abatimento", "@pAbat")
        Call .Parameters.add(.CreateParameter("@pAbat", dblNovoAbat, dbFieldTypeDouble))
       
    End With
    DevolveValorParaParcOrig = Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd) > 0
    Aplicacao.Disconnect
    Exit Function
err_Handler:
    FinallyConnection Aplicacao
    DevolveValorParaParcOrig = False
End Function

'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
Public Function AdicionaValorParcOrig(ByVal objBaixaParcial As VoLancamentoDuplicata, ByVal dblValor As Double, ByVal strColumn As String, ByVal Tabela As enuLancDup) As Boolean
    Dim selCmd      As IDBSelectCommand
    Dim rdResult    As IDBReader
    Dim updCmd      As IDBUpdateCommand
    Dim strTable    As String
    Dim strCol      As String

On Error GoTo err_Handler
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    strTable = IIf(Tabela = Duplicata, "[Duplicatas]", "[Lançamentos]")
    strCol = "[" & strColumn & "]"
    
    With selCmd
        
        .SelectClause = strCol
        .Table.TableName = strTable
        
        If Tabela = Duplicata Then
            Call .Filter.Append("Nota = @pNota")
        Else
            Call .Filter.Append("Código = @pNota")
        End If
        Call .Parameters.add(.CreateParameter("@pNota", objBaixaParcial.Codigo_Nota, dbFieldTypeString))
        Call .Filter.Append("Parcela = @pParcOrig")
        Call .Parameters.add(.CreateParameter("@pParcOrig", objBaixaParcial.parc_origem_baixa, dbFieldTypeLong))
        Call .Filter.Append("PagRec = @pPR")
        Call .Parameters.add(.CreateParameter("@pPR", objBaixaParcial.PagRec, dbFieldTypeString))
        Call .Filter.Append("Empresa = @pEmp")
        Call .Parameters.add(.CreateParameter("@pEmp", objBaixaParcial.Empresa, dbFieldTypeString))
        Call .Filter.Append("Tipo = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", objBaixaParcial.Tipo, dbFieldTypeString))
    
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)

    If Not rdResult.EOF Then
        Set updCmd = Aplicacao.CreateUpdateCommand
        With updCmd
            .Table = strTable
                           
            If Tabela = Duplicata Then
                Call .Filter.Append("Nota = @pNota")
            Else
                Call .Filter.Append("Código = @pNota")
            End If
            Call .Parameters.add(.CreateParameter("@pNota", objBaixaParcial.Codigo_Nota, dbFieldTypeString))
            Call .Filter.Append("Parcela = @pParcOrig")
            Call .Parameters.add(.CreateParameter("@pParcOrig", objBaixaParcial.parc_origem_baixa, dbFieldTypeLong))
            Call .Filter.Append("PagRec = @pPR")
            Call .Parameters.add(.CreateParameter("@pPR", objBaixaParcial.PagRec, dbFieldTypeString))
            Call .Filter.Append("Empresa = @pEmp")
            Call .Parameters.add(.CreateParameter("@pEmp", objBaixaParcial.Empresa, dbFieldTypeString))
            Call .Filter.Append("Tipo = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", objBaixaParcial.Tipo, dbFieldTypeString))
           
            Call .AddValue(strCol, "@pCol")
            Call .Parameters.add(.CreateParameter("@pCol", rdResult.GetDouble(strColumn) - dblValor, dbFieldTypeDouble))
            
            AdicionaValorParcOrig = Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, updCmd) > 0
        End With
    End If
    Aplicacao.Disconnect
    Exit Function

err_Handler:
    FinallyConnection Aplicacao
    AdicionaValorParcOrig = False
    
End Function

'Projeto: 17081 - Sugestão de Melhoria: 23370 - Ueder Budni (20/01/2014)
Public Function CarregarColBaixasParc(ByVal parcOrig As VoLancamentoDuplicata, ByVal Tabela As enuLancDup) As cColLancamentoDuplicata
    Dim selCmd      As IDBSelectCommand
    Dim rdResult    As IDBReader
    Dim strCol      As String
    Dim strTable    As String
    Dim colTmp      As cColLancamentoDuplicata
    Dim objTmp      As VoLancamentoDuplicata
    
On Error GoTo err_Handler
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    strTable = IIf(Tabela = Duplicata, "[Duplicatas]", "[Lançamentos]")
    strCol = IIf(Tabela = Duplicata, "Nota", "Código")
    
    With selCmd
        .SelectClause = "PagRec, Parcela, Empresa, Tipo, " & strCol
        .Table.TableName = strTable
        
        If Tabela = Duplicata Then
            Call .Filter.Append("Nota = @pNota")
        Else
            Call .Filter.Append("Código = @pNota")
        End If
        Call .Parameters.add(.CreateParameter("@pNota", parcOrig.Codigo_Nota, dbFieldTypeDouble))
        Call .Filter.Append("PagRec = @pPR")
        Call .Parameters.add(.CreateParameter("@pPR", parcOrig.PagRec, dbFieldTypeString))
        Call .Filter.Append("Empresa = @pEmp")
        Call .Parameters.add(.CreateParameter("@pEmp", parcOrig.Empresa, dbFieldTypeString))
        Call .Filter.Append("Tipo = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", parcOrig.Tipo, dbFieldTypeString))
        Call .Filter.Append("parc_origem_baixa = @pPOrig")
        Call .Parameters.add(.CreateParameter("@pPOrig", parcOrig.Parcela, dbFieldTypeLong))
    
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    
    If Not rdResult.EOF Then
        Set colTmp = New cColLancamentoDuplicata
        With rdResult
            While Not .EOF
                Set objTmp = New VoLancamentoDuplicata
                Set objTmp = Carregar(IIf(.GetString("PagRec") = "P", enuPagRec.Pagamento, enuPagRec.Recebimento), _
                                  .GetDouble(strCol), _
                                  .GetString("Tipo"), _
                                  .GetLong("Parcela"), _
                                  .GetString("Empresa"), _
                                  Tabela)
                Call colTmp.add(objTmp)
                Set objTmp = Nothing
                .MoveNext
            Wend
        End With
        Set CarregarColBaixasParc = colTmp
    End If
    Aplicacao.Disconnect
    Exit Function
    
err_Handler:
    FinallyConnection Aplicacao
    CarregarColBaixasParc = Nothing
    
End Function

