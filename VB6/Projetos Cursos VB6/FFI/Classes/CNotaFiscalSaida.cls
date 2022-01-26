VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNotaFiscalSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Private sTipoRegistro As String
Private lNumero As Long
Private sFornecedor As String
Private dEmissao As Date
Private sTipoEmpresa As String
Private sEmpresa As String
Private sMoeda As String
Private lCondicaoPagamento As Long
Private dDespesaFinanceira As Double
Private dDescontoValor As Double
Private dDescontoPorcentagem As Double
Private lTransportadora As Long
Private sFreteConta As String
Private dValorFrete As Double
Private sUsuario As String
Private sUsuarioLberacao As String
Private bRegistroImpresso As Boolean
Private dValidade As Date
Private sSituacao As String
Private sMotivo As String
Private dDataCadastro As Date
Private dHoraCadastro As Date
Private sSerie As String
Private sPedidoCliente As String
Private dTaxasExtras As Double
Private dValorSinal As Double
Private dDataSinal As Date
Private sViaTransporte As String
Private dPesoLiquido As Double
Private dPesoBruto As Double
Private dVolumes As Double
Private sEspecieVolume As String
Private dValorSeguro As Double
Private sMarcaNota As String
Private dDescontoEspecial As Double
Private sContato As String
Private sDepartamentoContato As String
Private sFoneContato As String
Private sFaxContato As String
Private sObservacoes As String
Private sObservacoesInternas As String
Private sEnderecoEntrega As String
Private sBairroEntrega As String
Private sCepEntrega As String
Private sCidadeEntrega As String
Private sEstadoEntrega As String
Private sPessoa As String
Private sCnpjEntrega As String
Private sIEEntrega As String
Private sFoneEntrega As String
Private sRamalEntrega As String
Private sEnderecoCobranca As String
Private sBairroCobranca As String
Private sCepCobranca As String
Private sCidadeCobranca As String
Private sEstadoCobranca As String
Private sFoneCobranca As String
Private sRamalCobranca As String
Private sEmitenteCarga As String
Private sNaturezaCarga As String
Private lNotaFiscalCarga As Long
Private dValorMercadoriasCarga As Double
Private dBaseIcms As Double
Private dBaseIpi As Double
Private dTotalIcms As Double
Private dBaseIcmsSubstituicao As Double
Private dTotalIcmsSubstituicao As Double
Private dTotalDesconto As Double
Private dTotalFrete As Double
Private dReembolsoFrete As Double
Private dTotalSeguro As Double
Private dTotalDespesas As Double
Private dTotalTaxas As Double
Private dIpiBaseIcms As Double
Private dTotalIpi As Double
Private dTotalMaoObra As Double
Private dTotalInss As Double
Private dTotalIss As Double
Private dTotalIrf As Double
Private dTotalOutros As Double
Private dTotalMercadorias As Double
Private dValorTotal As Double
Private dDescontoFinanceiro As Double
Private lBanco As Long
Private lConta As Long
Private sDescricaoTaxa As String
Private dDespesaImportacao As Double
Private dCotacaoDi As Double
Private sPlacaVeiculo As String
Private sLocalEmbarque As String
Private sLocalDesembarque As String
Private dTotalQuantidade As Double
Private sNaturezaOperacao As String
Private sComplementoNaturezaOperacao As String
Private sTipoRegistroNotaFiscalComplementar As String
Private lNumeroNotaComplementar As Long
Private bNotaVendasVarejo As Boolean
Private dTotalIssRetido As Double
Private dDataProgramada As Date
Private dPercentualDescontoGeral As Double
Private dTotalImpostoImportacao As Double
Private iCaixa As Integer
Private dDescontoLista As Double
Private dPercentualFrete As Double
Private dPercentualEmbalagem As Double
Private dDescontoMedio As Double
Private dMargemMedia As Double
Private dMargemFinal As Double
Private bMercadoriaTerceiros As Boolean
Private sFantasiaProprietariaMercadoria As String
Private lCodigoProjeto As Long
Private sRegistroIntegrado As String
Private dBasePisConfins As Double
Private dValorPis As Double
Private dValorCOFINS As Double
Private dValorCsl As Double
Private dSaldoSinal As Double
Private Const NomeTabela$ = "[Notas Fiscais de Saída]"

Public Property Get TipoRegistro() As String
    TipoRegistro = sTipoRegistro
End Property

Public Property Let TipoRegistro(ByVal sNewTipoRegistro As String)
    sTipoRegistro = sNewTipoRegistro
End Property

Public Property Get Numero() As Long
    Numero = lNumero
End Property

Public Property Let Numero(ByVal lNewNumero As Long)
    lNumero = lNewNumero
End Property

Public Property Get Fornecedor() As String
    Fornecedor = sFornecedor
End Property

Public Property Let Fornecedor(ByVal sNewFornecedor As String)
    sFornecedor = sNewFornecedor
End Property

Public Property Get Emissao() As Date
    Emissao = dEmissao
End Property

Public Property Let Emissao(ByVal dNewEmissao As Date)
    dEmissao = dNewEmissao
End Property

Public Property Get tipoEmpresa() As String
    tipoEmpresa = sTipoEmpresa
End Property

Public Property Let tipoEmpresa(ByVal sNewTipoEmpresa As String)
    sTipoEmpresa = sNewTipoEmpresa
End Property

Public Property Get Empresa() As String
    Empresa = sEmpresa
End Property

Public Property Let Empresa(ByVal sNewEmpresa As String)
    sEmpresa = sNewEmpresa
End Property

Public Property Get Moeda() As String
    Moeda = sMoeda
End Property

Public Property Let Moeda(ByVal sNewMoeda As String)
    sMoeda = sNewMoeda
End Property

Public Property Get CondicaoPagamento() As Long
    CondicaoPagamento = lCondicaoPagamento
End Property

Public Property Let CondicaoPagamento(ByVal lNewCondicaoPagamento As Long)
    lCondicaoPagamento = lNewCondicaoPagamento
End Property

Public Property Get despesaFinanceira() As Double
    despesaFinanceira = dDespesaFinanceira
End Property

Public Property Let despesaFinanceira(ByVal dNewDespesaFinanceira As Double)
    dDespesaFinanceira = dNewDespesaFinanceira
End Property

Public Property Get descontoValor() As Double
    descontoValor = dDescontoValor
End Property

Public Property Let descontoValor(ByVal dNewDescontoValor As Double)
    dDescontoValor = dNewDescontoValor
End Property

Public Property Get descontoPorcentagem() As Double
    descontoPorcentagem = dDescontoPorcentagem
End Property

Public Property Let descontoPorcentagem(ByVal dNewDescontoPorcentagem As Double)
    dDescontoPorcentagem = dNewDescontoPorcentagem
End Property

Public Property Get Transportadora() As Long
    Transportadora = lTransportadora
End Property

Public Property Let Transportadora(ByVal lNewTransportadora As Long)
    lTransportadora = lNewTransportadora
End Property

Public Property Get FreteConta() As String
    FreteConta = sFreteConta
End Property

Public Property Let FreteConta(ByVal sNewFreteConta As String)
    sFreteConta = sNewFreteConta
End Property

Public Property Get ValorFrete() As Double
    ValorFrete = dValorFrete
End Property

Public Property Let ValorFrete(ByVal dNewValorFrete As Double)
    dValorFrete = dNewValorFrete
End Property

Public Property Get usuario() As String
    usuario = sUsuario
End Property

Public Property Let usuario(ByVal sNewUsuario As String)
    sUsuario = sNewUsuario
End Property

Public Property Get usuarioLiberacao() As String
    usuarioLiberacao = sUsuarioLberacao
End Property

Public Property Let usuarioLiberacao(ByVal sNewUsuarioLiberacao As String)
    sUsuarioLberacao = sNewUsuarioLiberacao
End Property

Public Property Get registroImpresso() As Boolean
    registroImpresso = bRegistroImpresso
End Property

Public Property Let registroImpresso(ByVal bNewRegistroImpresso As Boolean)
    bRegistroImpresso = bNewRegistroImpresso
End Property

Public Property Get validade() As Date
    validade = dValidade
End Property

Public Property Let validade(ByVal dNewValidade As Date)
    dValidade = dNewValidade
End Property

Public Property Get situacao() As String
    situacao = sSituacao
End Property

Public Property Let situacao(ByVal sNewSituacao As String)
    sSituacao = sNewSituacao
End Property

Public Property Get Motivo() As String
    Motivo = sMotivo
End Property

Public Property Let Motivo(ByVal sNewMotivo As String)
    sMotivo = sNewMotivo
End Property

Public Property Get dataCadastro() As Date
    dataCadastro = dDataCadastro
End Property

Public Property Let dataCadastro(ByVal dNewDataCadastro As Date)
    dDataCadastro = dNewDataCadastro
End Property

Public Property Get horaCadastro() As Date
    horaCadastro = dHoraCadastro
End Property

Public Property Let horaCadastro(ByVal dNewHoraCadastro As Date)
    dHoraCadastro = dNewHoraCadastro
End Property

Public Property Get serie() As String
    serie = sSerie
End Property

Public Property Let serie(ByVal sNewSerie As String)
    sSerie = sNewSerie
End Property

Public Property Get pedidoCliente() As String
    pedidoCliente = sPedidoCliente
End Property

Public Property Let pedidoCliente(ByVal sNewPedidoCliente As String)
    sPedidoCliente = sNewPedidoCliente
End Property

Public Property Get taxasExtras() As Double
    taxasExtras = dTaxasExtras
End Property

Public Property Let taxasExtras(ByVal dNewTaxasExtras As Double)
    dTaxasExtras = dNewTaxasExtras
End Property

Public Property Get valorSinal() As Double
    valorSinal = dValorSinal
End Property

Public Property Let valorSinal(ByVal dNewValorSinal As Double)
    dValorSinal = dNewValorSinal
End Property

Public Property Get dataSinal() As Date
    dataSinal = dDataSinal
End Property

Public Property Let dataSinal(ByVal dNewDataSinal As Date)
    dDataSinal = dNewDataSinal
End Property

Public Property Get ViaTransporte() As String
    ViaTransporte = sViaTransporte
End Property

Public Property Let ViaTransporte(ByVal sNewViaTransporte As String)
    sViaTransporte = sNewViaTransporte
End Property

Public Property Get PesoLiquido() As Double
    PesoLiquido = dPesoLiquido
End Property

Public Property Let PesoLiquido(ByVal dNewPesoLiquido As Double)
    dPesoLiquido = dNewPesoLiquido
End Property

Public Property Get PesoBruto() As Double
    PesoBruto = dPesoBruto
End Property

Public Property Let PesoBruto(ByVal dNewPesoBruto As Double)
    dPesoBruto = dNewPesoBruto
End Property

Public Property Get Volumes() As Double
    Volumes = dVolumes
End Property

Public Property Let Volumes(ByVal dNewVolumes As Double)
    dVolumes = dNewVolumes
End Property

Public Property Get especieVolume() As String
    especieVolume = sEspecieVolume
End Property

Public Property Let especieVolume(ByVal sNewEspecieVolume As String)
    sEspecieVolume = sNewEspecieVolume
End Property

Public Property Get ValorSeguro() As Double
    ValorSeguro = dValorSeguro
End Property

Public Property Let ValorSeguro(ByVal dNewValorSeguro As Double)
    dValorSeguro = dNewValorSeguro
End Property

Public Property Get MarcaNota() As String
    MarcaNota = sMarcaNota
End Property

Public Property Let MarcaNota(ByVal sNewMarcaNota As String)
    sMarcaNota = sNewMarcaNota
End Property

Public Property Get descontoEspecial() As Double
    descontoEspecial = dDescontoEspecial
End Property

Public Property Let descontoEspecial(ByVal dNewDescontoEspecial As Double)
    dDescontoEspecial = dNewDescontoEspecial
End Property

Public Property Get Contato() As String
    Contato = sContato
End Property

Public Property Let Contato(ByVal sNewContato As String)
    sContato = sNewContato
End Property

Public Property Get departamentoContato() As String
    departamentoContato = sDepartamentoContato
End Property

Public Property Let departamentoContato(ByVal sNewDepartamentoContato As String)
    sDepartamentoContato = sNewDepartamentoContato
End Property

Public Property Get foneContato() As String
    foneContato = sFoneContato
End Property

Public Property Let foneContato(ByVal sNewFoneContato As String)
    sFoneContato = sNewFoneContato
End Property

Public Property Get faxContato() As String
    faxContato = sFaxContato
End Property

Public Property Let faxContato(ByVal sNewFaxContato As String)
    sFaxContato = sNewFaxContato
End Property

Public Property Get Observacoes() As String
    Observacoes = sObservacoes
End Property

Public Property Let Observacoes(ByVal sNewObservacoes As String)
    sObservacoes = sNewObservacoes
End Property

Public Property Get observacoesInternas() As String
    observacoesInternas = sObservacoesInternas
End Property

Public Property Let observacoesInternas(ByVal sNewObservacoesInternas As String)
    sObservacoesInternas = sNewObservacoesInternas
End Property

Public Property Get EnderecoEntrega() As String
    EnderecoEntrega = sEnderecoEntrega
End Property

Public Property Let EnderecoEntrega(ByVal sNewEnderecoEntrega As String)
    sEnderecoEntrega = sNewEnderecoEntrega
End Property

Public Property Get bairroEntrega() As String
    bairroEntrega = sBairroEntrega
End Property

Public Property Let bairroEntrega(ByVal sNewBairroEntrega As String)
    sBairroEntrega = sNewBairroEntrega
End Property

Public Property Get CepEntrega() As String
    CepEntrega = sCepEntrega
End Property

Public Property Let CepEntrega(ByVal sNewCepEntrega As String)
    sCepEntrega = sNewCepEntrega
End Property

Public Property Get cidadeEntrega() As String
    cidadeEntrega = sCidadeEntrega
End Property

Public Property Let cidadeEntrega(ByVal sNewCidadeEntrega As String)
    sCidadeEntrega = sNewCidadeEntrega
End Property

Public Property Get estadoEntrega() As String
    estadoEntrega = sEstadoEntrega
End Property

Public Property Let estadoEntrega(ByVal sNewEstadoEntrega As String)
    sEstadoEntrega = sNewEstadoEntrega
End Property

Public Property Get Pessoa() As String
    Pessoa = sPessoa
End Property

Public Property Let Pessoa(ByVal sNewPessoa As String)
    sPessoa = sNewPessoa
End Property

Public Property Get CnpjEntrega() As String
    CnpjEntrega = sCnpjEntrega
End Property

Public Property Let CnpjEntrega(ByVal sNewCnpjEntrega As String)
    sCnpjEntrega = sNewCnpjEntrega
End Property

Public Property Get ieEntrega() As String
    ieEntrega = sIEEntrega
End Property

Public Property Let ieEntrega(ByVal sNewIeEntrega As String)
    sIEEntrega = sNewIeEntrega
End Property

Public Property Get foneEntrega() As String
    foneEntrega = sFoneEntrega
End Property

Public Property Let foneEntrega(ByVal sNewFoneEntrega As String)
    sFoneEntrega = sNewFoneEntrega
End Property

Public Property Get ramalEntrega() As String
    ramalEntrega = sRamalEntrega
End Property

Public Property Let ramalEntrega(ByVal sNewRamalEntrega As String)
    sRamalEntrega = sNewRamalEntrega
End Property

Public Property Get EnderecoCobranca() As String
    EnderecoCobranca = sEnderecoCobranca
End Property

Public Property Let EnderecoCobranca(ByVal sNewEnderecoCobranca As String)
    sEnderecoCobranca = sNewEnderecoCobranca
End Property

Public Property Get bairroCobranca() As String
    bairroCobranca = sBairroCobranca
End Property

Public Property Let bairroCobranca(ByVal sNewBairroCobranca As String)
    sBairroCobranca = sNewBairroCobranca
End Property

Public Property Get CepCobranca() As String
    CepCobranca = sCepCobranca
End Property

Public Property Let CepCobranca(ByVal sNewCepCobranca As String)
    sCepCobranca = sNewCepCobranca
End Property

Public Property Get cidadeCobranca() As String
    cidadeCobranca = sCidadeCobranca
End Property

Public Property Let cidadeCobranca(ByVal sNewCidadeCobranca As String)
    sCidadeCobranca = sNewCidadeCobranca
End Property

Public Property Get estadoCobranca() As String
    estadoCobranca = sEstadoCobranca
End Property

Public Property Let estadoCobranca(ByVal sNewEstadoCobranca As String)
    sEstadoCobranca = sNewEstadoCobranca
End Property

Public Property Get foneCobranca() As String
    foneCobranca = sFoneCobranca
End Property

Public Property Let foneCobranca(ByVal sNewFoneCobranca As String)
    sFoneCobranca = sNewFoneCobranca
End Property

Public Property Get ramalCobranca() As String
    ramalCobranca = sRamalCobranca
End Property

Public Property Let ramalCobranca(ByVal sNewRamalCobranca As String)
    sRamalCobranca = sNewRamalCobranca
End Property

Public Property Get emitenteCarga() As String
    emitenteCarga = sEmitenteCarga
End Property

Public Property Let emitenteCarga(ByVal sNewEmitenteCarga As String)
    sEmitenteCarga = sNewEmitenteCarga
End Property

Public Property Get naturezaCarga() As String
    naturezaCarga = sNaturezaCarga
End Property

Public Property Let naturezaCarga(ByVal sNewNaturezaCarga As String)
    sNaturezaCarga = sNewNaturezaCarga
End Property

Public Property Get notaFiscalCarga() As Long
    notaFiscalCarga = lNotaFiscalCarga
End Property

Public Property Let notaFiscalCarga(ByVal lNewNotaFiscalCarga As Long)
    lNotaFiscalCarga = lNewNotaFiscalCarga
End Property

Public Property Get valorMercadoriasCarga() As Double
    valorMercadoriasCarga = dValorMercadoriasCarga
End Property

Public Property Let valorMercadoriasCarga(ByVal dNewValorMercadoriasCarga As Double)
    dValorMercadoriasCarga = dNewValorMercadoriasCarga
End Property

Public Property Get BaseICMS() As Double
    BaseICMS = dBaseIcms
End Property

Public Property Let BaseICMS(ByVal dNewBaseIcms As Double)
    dBaseIcms = dNewBaseIcms
End Property

Public Property Get BaseIpi() As Double
    BaseIpi = dBaseIpi
End Property

Public Property Let BaseIpi(ByVal dNewBaseIpi As Double)
    dBaseIpi = dNewBaseIpi
End Property

Public Property Get totalIcms() As Double
    totalIcms = dTotalIcms
End Property

Public Property Let totalIcms(ByVal dNewTotalIcms As Double)
    dTotalIcms = dNewTotalIcms
End Property

Public Property Get BaseIcmsSubstituicao() As Double
    BaseIcmsSubstituicao = dBaseIcmsSubstituicao
End Property

Public Property Let BaseIcmsSubstituicao(ByVal dNewBaseIcmsSubstituicao As Double)
    dBaseIcmsSubstituicao = dNewBaseIcmsSubstituicao
End Property

Public Property Get totalIcmsSubstituicao() As Double
    totalIcmsSubstituicao = dTotalIcmsSubstituicao
End Property

Public Property Let totalIcmsSubstituicao(ByVal dNewTotalIcmsSubstituicao As Double)
    dTotalIcmsSubstituicao = dNewTotalIcmsSubstituicao
End Property

Public Property Get totalDesconto() As Double
    totalDesconto = dTotalDesconto
End Property

Public Property Let totalDesconto(ByVal dNewTotalDesconto As Double)
    dTotalDesconto = dNewTotalDesconto
End Property

Public Property Get totalFrete() As Double
    totalFrete = dTotalFrete
End Property

Public Property Let totalFrete(ByVal dNewTotalFrete As Double)
    dTotalFrete = dNewTotalFrete
End Property

Public Property Get reembolsoFrete() As Double
    reembolsoFrete = dReembolsoFrete
End Property

Public Property Let reembolsoFrete(ByVal dNewReembolsoFrete As Double)
    dReembolsoFrete = dNewReembolsoFrete
End Property

Public Property Get totalSeguro() As Double
    totalSeguro = dTotalSeguro
End Property

Public Property Let totalSeguro(ByVal dNewTotalSeguro As Double)
    dTotalSeguro = dNewTotalSeguro
End Property

Public Property Get totalDespesas() As Double
    totalDespesas = dTotalDespesas
End Property

Public Property Let totalDespesas(ByVal dNewTotalDespesas As Double)
    dTotalDespesas = dNewTotalDespesas
End Property

Public Property Get totalTaxas() As Double
    totalTaxas = dTotalTaxas
End Property

Public Property Let totalTaxas(ByVal dNewTotalTaxas As Double)
    dTotalTaxas = dNewTotalTaxas
End Property

Public Property Get ipiBaseIcms() As Double
    ipiBaseIcms = dIpiBaseIcms
End Property

Public Property Let ipiBaseIcms(ByVal dNewIpiBaseIcms As Double)
    dIpiBaseIcms = dNewIpiBaseIcms
End Property

Public Property Get totalIpi() As Double
    totalIpi = dTotalIpi
End Property

Public Property Let totalIpi(ByVal dNewTotalIpi As Double)
    dTotalIpi = dNewTotalIpi
End Property

Public Property Get totalMaoObra() As Double
    totalMaoObra = dTotalMaoObra
End Property

Public Property Let totalMaoObra(ByVal dNewTotalMaoObra As Double)
    dTotalMaoObra = dNewTotalMaoObra
End Property

Public Property Get totalInss() As Double
    totalInss = dTotalInss
End Property

Public Property Let totalInss(ByVal dNewTotalInss As Double)
    dTotalInss = dNewTotalInss
End Property

Public Property Get totalIss() As Double
    totalIss = dTotalIss
End Property

Public Property Let totalIss(ByVal dNewTotalIss As Double)
    dTotalIss = dNewTotalIss
End Property

Public Property Get totalIrf() As Double
    totalIrf = dTotalIrf
End Property

Public Property Let totalIrf(ByVal dNewTotalIrf As Double)
    dTotalIrf = dNewTotalIrf
End Property

Public Property Get totalOutros() As Double
    totalOutros = dTotalOutros
End Property

Public Property Let totalOutros(ByVal dNewTotalOutros As Double)
    dTotalOutros = dNewTotalOutros
End Property

Public Property Get TotalMercadorias() As Double
    TotalMercadorias = dTotalMercadorias
End Property

Public Property Let TotalMercadorias(ByVal dNewTotalMercadorias As Double)
    dTotalMercadorias = dNewTotalMercadorias
End Property

Public Property Get ValorTotal() As Double
    ValorTotal = dValorTotal
End Property

Public Property Let ValorTotal(ByVal dNewValorTotal As Double)
    dValorTotal = dNewValorTotal
End Property

Public Property Get DescontoFinanceiro() As Double
    DescontoFinanceiro = dDescontoFinanceiro
End Property

Public Property Let DescontoFinanceiro(ByVal dNewDescontoFinanceiro As Double)
    dDescontoFinanceiro = dNewDescontoFinanceiro
End Property

Public Property Get Banco() As Long
    Banco = lBanco
End Property

Public Property Let Banco(ByVal lNewBanco As Long)
    lBanco = lNewBanco
End Property

Public Property Get Conta() As Long
    Conta = lConta
End Property

Public Property Let Conta(ByVal lNewConta As Long)
    lConta = lNewConta
End Property

Public Property Get descricaoTaxa() As String
    descricaoTaxa = sDescricaoTaxa
End Property

Public Property Let descricaoTaxa(ByVal sNewDescricaoTaxa As String)
    sDescricaoTaxa = sNewDescricaoTaxa
End Property

Public Property Get despesaImportacao() As Double
    despesaImportacao = dDespesaImportacao
End Property

Public Property Let despesaImportacao(ByVal dNewDespesaImportacao As Double)
    dDespesaImportacao = dNewDespesaImportacao
End Property

Public Property Get cotacaoDi() As Double
    cotacaoDi = dCotacaoDi
End Property

Public Property Let cotacaoDi(ByVal dNewCotacaoDi As Double)
    dCotacaoDi = dNewCotacaoDi
End Property

Public Property Get PlacaVeiculo() As String
    PlacaVeiculo = sPlacaVeiculo
End Property

Public Property Let PlacaVeiculo(ByVal sNewPlacaVeiculo As String)
    sPlacaVeiculo = sNewPlacaVeiculo
End Property

Public Property Get localEmbarque() As String
    localEmbarque = sLocalEmbarque
End Property

Public Property Let localEmbarque(ByVal sNewLocalEmbarque As String)
    sLocalEmbarque = sNewLocalEmbarque
End Property

Public Property Get localDesembarque() As String
    localDesembarque = sLocalDesembarque
End Property

Public Property Let localDesembarque(ByVal sNewLocalDesembarque As String)
    sLocalDesembarque = sNewLocalDesembarque
End Property

Public Property Get totalQuantidade() As Double
    totalQuantidade = dTotalQuantidade
End Property

Public Property Let totalQuantidade(ByVal dNewTotalQuantidade As Double)
    dTotalQuantidade = dNewTotalQuantidade
End Property

Public Property Get NaturezaOperacao() As String
    NaturezaOperacao = sNaturezaOperacao
End Property

Public Property Let NaturezaOperacao(ByVal sNewNaturezaOperacao As String)
    sNaturezaOperacao = sNewNaturezaOperacao
End Property

Public Property Get complementoNaturezaOperacao() As String
    complementoNaturezaOperacao = sComplementoNaturezaOperacao
End Property

Public Property Let complementoNaturezaOperacao(ByVal sNewComplementoNaturezaOperacao As String)
    sComplementoNaturezaOperacao = sNewComplementoNaturezaOperacao
End Property

Public Property Get tipoRegistroNotaFiscalComplementar() As String
    tipoRegistroNotaFiscalComplementar = sTipoRegistroNotaFiscalComplementar
End Property

Public Property Let tipoRegistroNotaFiscalComplementar(ByVal sNewTipoRegistroNotaFiscalComplementar As String)
    sTipoRegistroNotaFiscalComplementar = sNewTipoRegistroNotaFiscalComplementar
End Property

Public Property Get numeroNotaComplementar() As Long
    numeroNotaComplementar = lNumeroNotaComplementar
End Property

Public Property Let numeroNotaComplementar(ByVal lNewNumeroNotaComplementar As Long)
    lNumeroNotaComplementar = lNewNumeroNotaComplementar
End Property

Public Property Get notaVendasVarejo() As Boolean
    notaVendasVarejo = bNotaVendasVarejo
End Property

Public Property Let notaVendasVarejo(ByVal bNewNotaVendasVarejo As Boolean)
    bNotaVendasVarejo = bNewNotaVendasVarejo
End Property

Public Property Get totalIssRetido() As Double
    totalIssRetido = dTotalIssRetido
End Property

Public Property Let totalIssRetido(ByVal dNewTotalIssRetido As Double)
    dTotalIssRetido = dNewTotalIssRetido
End Property

Public Property Get dataProgramada() As Date
    dataProgramada = dDataProgramada
End Property

Public Property Let dataProgramada(ByVal dNewDataProgramada As Date)
    dDataProgramada = dNewDataProgramada
End Property

Public Property Get percentualDescontoGeral() As Double
    percentualDescontoGeral = dPercentualDescontoGeral
End Property

Public Property Let percentualDescontoGeral(ByVal dNewPercentualDescontoGeral As Double)
    dPercentualDescontoGeral = dNewPercentualDescontoGeral
End Property

Public Property Get totalImpostoImportacao() As Double
    totalImpostoImportacao = dTotalImpostoImportacao
End Property

Public Property Let totalImpostoImportacao(ByVal dNewTotalImpostoImportacao As Double)
    dTotalImpostoImportacao = dNewTotalImpostoImportacao
End Property

Public Property Get caixa() As Integer
    caixa = iCaixa
End Property

Public Property Let caixa(ByVal iNewCaixa As Integer)
    iCaixa = iNewCaixa
End Property

Public Property Get descontoLista() As Double
    descontoLista = dDescontoLista
End Property

Public Property Let descontoLista(ByVal dNewDescontoLista As Double)
    dDescontoLista = dNewDescontoLista
End Property

Public Property Get percentualFrete() As Double
    percentualFrete = dPercentualFrete
End Property

Public Property Let percentualFrete(ByVal dNewPercentualFrete As Double)
    dPercentualFrete = dNewPercentualFrete
End Property

Public Property Get percentualEmbalagem() As Double
    percentualEmbalagem = dPercentualEmbalagem
End Property

Public Property Let percentualEmbalagem(ByVal dNewPercentualEmbalagem As Double)
    dPercentualEmbalagem = dNewPercentualEmbalagem
End Property

Public Property Get descontoMedio() As Double
    descontoMedio = dDescontoMedio
End Property

Public Property Let descontoMedio(ByVal dNewDescontoMedio As Double)
    dDescontoMedio = dNewDescontoMedio
End Property

Public Property Get margemMedia() As Double
    margemMedia = dMargemMedia
End Property

Public Property Let margemMedia(ByVal dNewMargemMedia As Double)
    dMargemMedia = dNewMargemMedia
End Property

Public Property Get margemFinal() As Double
    margemFinal = dMargemFinal
End Property

Public Property Let margemFinal(ByVal dNewMargemFinal As Double)
    dMargemFinal = dNewMargemFinal
End Property

Public Property Get mercadoriaTerceiros() As Boolean
    mercadoriaTerceiros = bMercadoriaTerceiros
End Property

Public Property Let mercadoriaTerceiros(ByVal bNewMercadoriaTerceiros As Boolean)
    bMercadoriaTerceiros = bNewMercadoriaTerceiros
End Property

Public Property Get fantasiaProprietariaMercadoria() As String
    fantasiaProprietariaMercadoria = sFantasiaProprietariaMercadoria
End Property

Public Property Let fantasiaProprietariaMercadoria(ByVal sNewFantasiaProprietariaMercadoria As String)
    sFantasiaProprietariaMercadoria = sNewFantasiaProprietariaMercadoria
End Property

Public Property Get codigoProjeto() As Long
    codigoProjeto = lCodigoProjeto
End Property

Public Property Let codigoProjeto(ByVal lNewCodigoProjeto As Long)
    lCodigoProjeto = lNewCodigoProjeto
End Property

Public Property Get registroIntegrado() As String
    registroIntegrado = sRegistroIntegrado
End Property

Public Property Let registroIntegrado(ByVal sNewRegistroIntegrado As String)
    sRegistroIntegrado = sNewRegistroIntegrado
End Property

Public Property Get basePisConfins() As Double
    basePisConfins = dBasePisConfins
End Property

Public Property Let basePisConfins(ByVal dNewBasePisConfins As Double)
    dBasePisConfins = dNewBasePisConfins
End Property

Public Property Get ValorPIS() As Double
    ValorPIS = dValorPis
End Property

Public Property Let ValorPIS(ByVal dNewValorPis As Double)
    dValorPis = dNewValorPis
End Property

Public Property Get ValorCOFINS() As Double
    ValorCOFINS = dValorCOFINS
End Property

Public Property Let ValorCOFINS(ByVal dNewValorCOFINS As Double)
    dValorCOFINS = dNewValorCOFINS
End Property

Public Property Get valorCsl() As Double
    valorCsl = dValorCsl
End Property

Public Property Let valorCsl(ByVal dNewValorCsl As Double)
    dValorCsl = dNewValorCsl
End Property

Public Property Get saldoSinal() As Double
    saldoSinal = dSaldoSinal
End Property

Public Property Let saldoSinal(ByVal dNewSaldoSinal As Double)
    dSaldoSinal = dNewSaldoSinal
End Property

Public Sub incluir()
    Dim cmd As IDBInsertCommand
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateInsertCommand
    cmd.Table = NomeTabela
    Call PrepareCommand(cmd)
    Call Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd)
    Set cmd = Nothing
    Aplicacao.Disconnect
End Sub

Public Sub Alterar()
    Dim cmd As IDBUpdateCommand
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateUpdateCommand
    cmd.Table = NomeTabela
    Call cmd.Filter.Append("Tipo de Registro = @pTipoRegistro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", sTipoRegistro, dbFieldTypeString))
    Call cmd.Filter.Append("Número = @pNumero")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumero", lNumero, dbFieldTypeLong))
    Call cmd.Filter.Append("Fornecedor = @pFornecedor")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFornecedor", sFornecedor, dbFieldTypeString))
    Call PrepareCommand(cmd)
    Set cmd = Nothing
    Aplicacao.Disconnect
End Sub

Private Sub PrepareCommand(ByRef cmd As Object)
    Call cmd.AddValue("[Tipo de Registro]", "@pTipoRegistro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", sTipoRegistro, dbFieldTypeString))
    
    Call cmd.AddValue("[Número]", "@pNumero")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumero", lNumero, dbFieldTypeLong))
    
    Call cmd.AddValue("[Fornecedor]", "@pFornecedor")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFornecedor", sFornecedor, dbFieldTypeString))
    
    Call cmd.AddValue("[Emissão]", "@pEmissao")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmissao", dEmissao, dbFieldTypeDate))
    
    Call cmd.AddValue("[Tipo da Empresa]", "@pTipoEmpresa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoEmpresa", sTipoEmpresa, dbFieldTypeString))
    
    Call cmd.AddValue("Empresa", "@pEmpresa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", sEmpresa, dbFieldTypeString))
    
    Call cmd.AddValue("Moeda", "@pMoeda")
    Call cmd.Parameters.add(cmd.CreateParameter("@pMoeda", sMoeda, dbFieldTypeString))
    
    Call cmd.AddValue("[Condição de Pagamento]", "@CondicaoPagamento")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCondicaoPagamento", lCondicaoPagamento, dbFieldTypeInt))
    
    Call cmd.AddValue("[Despesa Financeira]", "@pDespesaFinanceira")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDespesaFinanceira", dDespesaFinanceira, dbFieldTypeInt))
    
    Call cmd.AddValue("[Desconto em Valor]", "@pDescontoValor")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDescontoValor", dDescontoValor, dbFieldTypeDecimal))
    
    Call cmd.AddValue("[Desconto em porcentagem]", "@pDescontoPorcentagem")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDescontoPorcentagem", dDescontoPorcentagem, dbFieldTypeInt))
    
    Call cmd.AddValue("[Transportadora]", "@pTransportadora")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTransportadora", lTransportadora, dbFieldTypeLong))
    
    Call cmd.AddValue("[Frete por Conta]", "@pFreteConta")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFreteConta", sFreteConta, dbFieldTypeString))
    
    Call cmd.AddValue("[Valor do Frete]", "@pValorFrete")
    Call cmd.Parameters.add(cmd.CreateParameter("@pValorFrete", dValorFrete, dbFieldTypeDecimal))

    Call cmd.AddValue("[Cadastrado por]", "@pCadastradoPor")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCadastradoPor", sUsuario, dbFieldTypeString))

    Call cmd.AddValue("[Duplicatas Liberado por]", "@pDuplicatasLiberadoPor")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDuplicatasLiberadoPor", sUsuarioLberacao, dbFieldTypeString))

    Call cmd.AddValue("[Registro Impresso]", "@pRegistroImpresso")
    Call cmd.Parameters.add(cmd.CreateParameter("@pRegistroImpresso", bRegistroImpresso, dbFieldTypeBool))
    
    Call cmd.AddValue("[Validade]", "@pValidade")
    Call cmd.Parameters.add(cmd.CreateParameter("@pValidade", dValidade, dbFieldTypeDate))

    Call cmd.AddValue("[Situação]", "@pSituacao")
    Call cmd.Parameters.add(cmd.CreateParameter("@pSituacao", sSituacao, dbFieldTypeString))

    Call cmd.AddValue("[Motivo]", "@pMotivo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pMotivo", sMotivo, dbFieldTypeString))

    Call cmd.AddValue("[Data]", "@pData")
    Call cmd.Parameters.add(cmd.CreateParameter("@pData", dDataCadastro, dbFieldTypeDate))

    Call cmd.AddValue("[Hora]", "@pHora")
    Call cmd.Parameters.add(cmd.CreateParameter("@pHora", dHoraCadastro, dbFieldTypeDateTime))

    Call cmd.AddValue("[Série]", "@pSerie")
    Call cmd.Parameters.add(cmd.CreateParameter("@pSerie", sSerie, dbFieldTypeString))
    
    Call cmd.AddValue("[Pedido do Cliente]", "@pPedidoCliente")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPedidoCliente", sPedidoCliente, dbFieldTypeString))
    
    Call cmd.AddValue("[Taxas Extras]", "@pTaxasExtras")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTaxasExtras", dTaxasExtras, dbFieldTypeDecimal))
    
    Call cmd.AddValue("[Valor do Sinal]", "@pValorSinal")
    Call cmd.Parameters.add(cmd.CreateParameter("@pValorSinal", dValorSinal, dbFieldTypeDecimal))
    
    Call cmd.AddValue("[Data do Sinal]", "@pDataSinal")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDataSinal", dDataSinal, dbFieldTypeDate))
    
    Call cmd.AddValue("[Via de Transporte]", "@pViaTransporte")
    Call cmd.Parameters.add(cmd.CreateParameter("@pViaTransporte", sViaTransporte, dbFieldTypeString))
    
    Call cmd.AddValue("[Peso Líquido]", "@pPesoLiquido")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPesoLiquido", dPesoLiquido, dbFieldTypeDecimal))
    
    Call cmd.AddValue("[Peso Bruto]", "@pPesoBruto")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPesoBruto", dPesoBruto, dbFieldTypeDecimal))

    Call cmd.AddValue("[Volumes]", "@pVolumes")
    Call cmd.Parameters.add(cmd.CreateParameter("@pVolumes", dVolumes, dbFieldTypeDecimal))

    Call cmd.AddValue("[Espécie]", "@pEspecie")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEspecie", sEspecieVolume, dbFieldTypeString))

    Call cmd.AddValue("[Valor do Seguro]", "@pValorSeguro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pValorSeguro", dValorSeguro, dbFieldTypeDecimal))
    
    Call cmd.AddValue("[Marca da Nota]", "@pMarcaNota")
    Call cmd.Parameters.add(cmd.CreateParameter("@pMarcaNota", sMarcaNota, dbFieldTypeString))

    Call cmd.AddValue("[Desconto Especial]", "@pDescontoEspecial")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDescontoEspecial", dDescontoEspecial, dbFieldTypeDecimal))
    
    Call cmd.AddValue("[Contato]", "@pContato")
    Call cmd.Parameters.add(cmd.CreateParameter("@pContato", sContato, dbFieldTypeString))

    Call cmd.AddValue("[Departamento]", "@pDepartamento")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDepartamento", sDepartamentoContato, dbFieldTypeString))

    Call cmd.AddValue("[Fone]", "@pFone")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFone", sFoneContato, dbFieldTypeString))

    Call cmd.AddValue("[Fax]", "@pFax")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFax", sFaxContato, dbFieldTypeString))

    Call cmd.AddValue("[Observações]", "@pObservacoes")
    Call cmd.Parameters.add(cmd.CreateParameter("@pObservacoes", sObservacoes, dbFieldTypeText))

    Call cmd.AddValue("[Observações Internas]", "@pObservacoesInternas")
    Call cmd.Parameters.add(cmd.CreateParameter("@pObservacoesInternas", sObservacoesInternas, dbFieldTypeText))

    Call cmd.AddValue("[Endereço de Entrega]", "@pEnderecoEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEnderecoEntrega", sEnderecoEntrega, dbFieldTypeString))
    
    Call cmd.AddValue("[Bairro de Entrega]", "@pBairroEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBairroEntrega", sBairroEntrega, dbFieldTypeString))
    
    Call cmd.AddValue("[CEP de Entrega]", "@pCepEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCepEntrega", sCepEntrega, dbFieldTypeString))

    Call cmd.AddValue("[Cidade de Entrega]", "@pCidadeEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCidadeEntrega", sCidadeEntrega, dbFieldTypeString))

    Call cmd.AddValue("[Estado de Entrega]", "@pEstadoEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEstadoEntrega", sEstadoEntrega, dbFieldTypeString))

    Call cmd.AddValue("[Pessoa]", "@pPessoa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPessoa", sPessoa, dbFieldTypeString))

    Call cmd.AddValue("[CNPJ de Entrega]", "@pCnpjEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCnpjEntrega", sCnpjEntrega, dbFieldTypeString))
    
    Call cmd.AddValue("[IE de Entrega]", "@pIeEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pIeEntrega", sIEEntrega, dbFieldTypeString))

    Call cmd.AddValue("[Fone de Entrega]", "@pFoneEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFoneEntrega", sFoneEntrega, dbFieldTypeString))

    Call cmd.AddValue("[Ramal de Entrega]", "@pRamalEntrega")
    Call cmd.Parameters.add(cmd.CreateParameter("@pRamalEntrega", sRamalEntrega, dbFieldTypeString))

    Call cmd.AddValue("[Endereço de Cobrança]", "@pEnderecoCobranca")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEnderecoCobranca", sEnderecoCobranca, dbFieldTypeString))

    Call cmd.AddValue("[Bairro de Cobrança]", "@pBairroCobranca")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBairroCobranca", sBairroCobranca, dbFieldTypeString))

    Call cmd.AddValue("[CEP de Cobrança]", "@pCepCobranca")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCepCobranca", sCepCobranca, dbFieldTypeString))

    Call cmd.AddValue("[Cidade de Cobrança]", "@pCidadeCobranca")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCidadeCobranca", sCidadeCobranca, dbFieldTypeString))

    Call cmd.AddValue("[Estado de Cobrança]", "@pEstadoCobranca")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEstadoCobranca", sEstadoCobranca, dbFieldTypeString))

    Call cmd.AddValue("[Fone de Cobrança]", "@pFoneCobranca")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFoneCobranca", sFoneCobranca, dbFieldTypeString))

    Call cmd.AddValue("[Ramal de Cobrança]", "@pRamalCobranca")
    Call cmd.Parameters.add(cmd.CreateParameter("@pRamalCobranca", sRamalCobranca, dbFieldTypeString))

    Call cmd.AddValue("[Emitente da Carga]", "@pEmitenteCarga")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmitenteCarga", sEmitenteCarga, dbFieldTypeString))

    Call cmd.AddValue("[Natureza da Carga]", "@pNaturezaCarga")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNaturezaCarga", sNaturezaCarga, dbFieldTypeString))

    Call cmd.AddValue("[Nota Fiscal da Carga]", "@pNotaFiscalCarga")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNotaFiscalCarga", lNotaFiscalCarga, dbFieldTypeLong))

    Call cmd.AddValue("[Valor das Mercadorias da Carga]", "@pValorMercadoriasCarga")
    Call cmd.Parameters.add(cmd.CreateParameter("@pValorMercadoriasCarga", dValorMercadoriasCarga, dbFieldTypeDouble))

    Call cmd.AddValue("[Base de ICMS]", "@pBaseIcms")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBaseIcms", dBaseIcms, dbFieldTypeDouble))

    Call cmd.AddValue("[Base de IPI]", "@pBaseIpi")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBaseIpi", dBaseIpi, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de ICMS]", "@pTotalIcms")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalIcms", dTotalIcms, dbFieldTypeDouble))

    Call cmd.AddValue("[Base de ICMS de Substituição]", "@pBaseIcmsSubstituicao")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBaseIcmsSubstituicao", dBaseIcmsSubstituicao, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de ICMS de Substituição]", "@pTotalIcmsSubstituicao")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalIcmsSubstituicao", dTotalIcmsSubstituicao, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de Descontos]", "@pTotalDescontos")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalDescontos", dTotalDesconto, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de Frete]", "@pTotalFrete")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalFrete", dTotalFrete, dbFieldTypeDouble))

    Call cmd.AddValue("[Reembolso_Frete]", "@pReembolsoFrete")
    Call cmd.Parameters.add(cmd.CreateParameter("@pReembolsoFrete", dReembolsoFrete, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de Seguro]", "@pTotalSeguro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalSeguro", dTotalSeguro, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de Despesas]", "@pTotalDespesas")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalDespesas", dTotalDespesas, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de Taxas]", "@pTotalTaxas")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalTaxas", dTotalTaxas, dbFieldTypeDouble))

    Call cmd.AddValue("[IPI na Base de ICMS]", "@pIpiBaseIcms")
    Call cmd.Parameters.add(cmd.CreateParameter("@pIpiBaseIcms", dIpiBaseIcms, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de IPI]", "@pTotalIpi")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalIpi", dTotalIpi, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de Mão de Obra]", "@pTotalMaoObra")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalMaoObra", dTotalMaoObra, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de INSS]", "@pTotalInss")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalInss", dTotalInss, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de IRF]", "@pTotalIrf")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalIrf", dTotalIrf, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de OUTROS]", "@pTotalOutros")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalOutros", dTotalOutros, dbFieldTypeDouble))

    Call cmd.AddValue("[Total de Mercadorias]", "@pTotalMercadorias")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalMercadorias", dTotalMercadorias, dbFieldTypeDouble))

    Call cmd.AddValue("[Valor Total]", "@pValorTotal")
    Call cmd.Parameters.add(cmd.CreateParameter("@pValorTotal", dValorTotal, dbFieldTypeDouble))

    Call cmd.AddValue("[Desconto Financeiro]", "@pDescontoFinanceiro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDescontoFinanceiro", dDescontoFinanceiro, dbFieldTypeDecimal))

    Call cmd.AddValue("[Banco]", "@pBanco")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBanco", lBanco, dbFieldTypeLong))

    Call cmd.AddValue("[Conta]", "@pConta")
    Call cmd.Parameters.add(cmd.CreateParameter("@pConta", lConta, dbFieldTypeLong))

    Call cmd.AddValue("[Descrição da Taxa]", "@pDescricaoTaxa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDescricaoTaxa", sDescricaoTaxa, dbFieldTypeString))

    Call cmd.AddValue("[Despesa de Importação]", "@pDespesaImportacao")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDespesaImportacao", dDespesaImportacao, dbFieldTypeDouble))

    Call cmd.AddValue("[Cotação da DI]", "@pCotacaoDi")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCotacaoDi", dCotacaoDi, dbFieldTypeDouble))

    Call cmd.AddValue("[Placa do Veículo]", "@pPlacaVeiculo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPlacaVeiculo", sPlacaVeiculo, dbFieldTypeString))

    Call cmd.AddValue("[Local de Embarque]", "@pLocalEmbarque")
    Call cmd.Parameters.add(cmd.CreateParameter("@pLocalEmbarque", sLocalEmbarque, dbFieldTypeString))

    Call cmd.AddValue("[Local de Desembarque]", "@pLocalDesembarque")
    Call cmd.Parameters.add(cmd.CreateParameter("@pLocalDesembarque", sLocalDesembarque, dbFieldTypeString))

    Call cmd.AddValue("[Total de Quantidade]", "@pTotalQuantidade")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalQuantidade", dTotalQuantidade, dbFieldTypeDouble))

    Call cmd.AddValue("[Natureza de Operação]", "@pNaturezaOperacao")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNaturezaOperacao", sNaturezaOperacao, dbFieldTypeString))

    Call cmd.AddValue("[Complemento da Natureza de Operação]", "@pComplementoNaturezaOperacao")
    Call cmd.Parameters.add(cmd.CreateParameter("@pComplementoNaturezaOperacao", sComplementoNaturezaOperacao, dbFieldTypeString))

    Call cmd.AddValue("[Tipo de Registro da Nota Complementar]", "@pTipoRegistroNotaComplementar")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistroNotaComplementar", sTipoRegistroNotaFiscalComplementar, dbFieldTypeString))

    Call cmd.AddValue("[Número da Nota Complementar]", "@pNumeroRegistroNotaComplementar")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumeroRegistroNotaComplementar", lNumeroNotaComplementar, dbFieldTypeLong))

    Call cmd.AddValue("[Nota proveniente do Vendas a Varejo]", "@pNotaVendaVarejo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNotaVendaVarejo", bNotaVendasVarejo, dbFieldTypeBool))

    Call cmd.AddValue("[Total de ISS Retido na Fonte]", "@pTotaIssRetido")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotaIssRetido", dTotalIssRetido, dbFieldTypeDouble))

    Call cmd.AddValue("[Data Programada]", "@pDataProgramada")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDataProgramada", dDataProgramada, dbFieldTypeDate))

    Call cmd.AddValue("[Percentual de Desconto Geral]", "@pPercentualDescontoGeral")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPercentualDescontoGeral", dPercentualDescontoGeral, dbFieldTypeDecimal))

    Call cmd.AddValue("[Total de Imposto de Importação]", "@pTotalImpostoImportacao")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalImpostoImportacao", dTotalImpostoImportacao, dbFieldTypeDouble))

    Call cmd.AddValue("[Caixa]", "@pCaixa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCaixa", iCaixa, dbFieldTypeInt))

    Call cmd.AddValue("[DescontoLista]", "@pDescontoLista")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDescontoLista", dDescontoLista, dbFieldTypeDecimal))

    Call cmd.AddValue("[PercFrete]", "@pPercentualFrete")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPercentualFrete", dPercentualFrete, dbFieldTypeInt))

    Call cmd.AddValue("[PercEmbalagem]", "@pPercentualEmbalagem")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPercentualEmbalagem", dPercentualEmbalagem, dbFieldTypeInt))

    Call cmd.AddValue("[DescMedio]", "@pDescontoMedio")
    Call cmd.Parameters.add(cmd.CreateParameter("@pDescontoMedio", dDescontoMedio, dbFieldTypeInt))

    Call cmd.AddValue("[MargemMedia]", "@pMargemMedia")
    Call cmd.Parameters.add(cmd.CreateParameter("@pMargemMedia", dMargemMedia, dbFieldTypeInt))

    Call cmd.AddValue("[MargemFinal]", "@pMargemFinal")
    Call cmd.Parameters.add(cmd.CreateParameter("@pMargemFinal", dMargemFinal, dbFieldTypeInt))

    Call cmd.AddValue("[MerTerceiro]", "@pMercadoriaTerceiro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pMercadoriaTerceiro", bMercadoriaTerceiros, dbFieldTypeBool))

    Call cmd.AddValue("EmpProprietaria", "@pEmpresaProprietaria")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresaProprietaria", sFantasiaProprietariaMercadoria, dbFieldTypeString))

    Call cmd.AddValue("CODPRJ", "@pCodigoProjeto")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCodigoProjeto", lCodigoProjeto, dbFieldTypeLong))

    Call cmd.AddValue("RegInt", "@pRegistroIntegrado")
    Call cmd.Parameters.add(cmd.CreateParameter("@pRegistroIntegrado", sRegistroIntegrado, dbFieldTypeString))

    Call cmd.AddValue("TotBpc", "@pBasePisCofins")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBasePisCofins", dBasePisConfins, dbFieldTypeDouble))

    Call cmd.AddValue("TotPis", "@pTotalPis")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalPis", dValorPis, dbFieldTypeDouble))

    Call cmd.AddValue("TotCof", "@pTotalCofins")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalCofins", dValorCOFINS, dbFieldTypeDouble))

    Call cmd.AddValue("TOTCSL", "@pTotalCsl")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTotalCsl", dValorCsl, dbFieldTypeDouble))

    Call cmd.AddValue("SLDSIN", "@pSaldoSinal")
    Call cmd.Parameters.add(cmd.CreateParameter("@pSaldoSinal", dSaldoSinal, dbFieldTypeDouble))
End Sub

Public Function Existe(sTipoRegistro As String, lNumero As Long, sFornecedor As String)
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = NomeTabela
    Call cmd.Filter.Append("Tipo de Registro = @pTipoRegistro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", sTipoRegistro, dbFieldTypeString))
    Call cmd.Filter.Append("Número = @pNumero")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumero", lNumero, dbFieldTypeLong))
    Call cmd.Filter.Append("Fornecedor = @pFornecedor")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFornecedor", sFornecedor, dbFieldTypeString))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    Existe = Not rdResult.EOF
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

Public Function lastCodigo() As Long
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = NomeTabela
    cmd.SelectClause = "MAX([Número]) AS LASTCOD"
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        lastCodigo = rdResult.GetLong("LASTCOD") + 1
    Else
        lastCodigo = 1
    End If
    Aplicacao.Disconnect
End Function

Public Function cotacaoMoeda() As Double
    If dCotacaoDi > 0 Then
        cotacaoMoeda = dCotacaoDi
    Else
        cotacaoMoeda = Cotacao(sMoeda, Date)
    End If
End Function
