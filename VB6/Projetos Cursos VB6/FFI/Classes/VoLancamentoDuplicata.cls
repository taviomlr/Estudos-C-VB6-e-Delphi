VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VoLancamentoDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mstrPagRec                          As String
Private mlngCodigo_Nota                     As String
Private mlngPARCELA                         As Long
Private mstrEmpresa                         As String
Private mstrTipo                            As String
Private mstrDescricao                       As String
Private mdtaEmissao                         As Date
Private mdtaVencimento                      As Date
Private mdtaPagamento                       As Date
Private mdtaLiberacao                       As Date
Private mdblValorOriginal                   As Double
Private mdblAcrescimo                       As Double
Private mdblAbatimento                      As Double
Private mlngBanco                           As Double
Private mlngConta                           As Double
Private mlngCentro                          As Double
Private mlngCheque                          As Double
Private mstrMoeda                           As String
Private mdblValordaMoeda                    As Double
Private mstrControle                        As String
Private mstrSituacao                        As String
Private mstrEnviado                         As String
Private mblnMarcacao                        As Boolean
Private mstrObs                             As String
Private mlngBordero                         As Long
Private mdtaDataIntegracaoFatoGerador       As Date
Private mdtaDataIntegracaoFatoPagamento     As Date
Private mstrUsuario                         As String
Private mdtaAlteracao                       As Date
Private mstrLINDIG                          As String
Private mblnConciliado                      As Boolean
Private mstrCarteira                        As String
Private mblnPrevisao                        As Boolean
Private mdblVlrMrD                          As Double
Private mdblPerMrD                          As Double
'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
Private mstrSeqNossoNumero                  As String
Private mdblVlrDsP                          As Double
Private mdblPerJur                          As Double
Private mdblVlrJur                          As Double
Private mlngCODPRJ                          As Long
Private mlngCODFPG                          As Long
Private mlngCheBan                          As Long
Private mstrCheAge                          As String
Private mstrCheEmi                          As String
Private mstrCheCco                          As String
Private mdblPerMul                          As Double
Private mdblVlrMul                          As Double
Private mstrAGECCE                          As String
Private mstrNOSNUM                          As String
Private mstrCODBAR                          As String
Private mstrLOCPAG                          As String
Private mstrINSTRU                          As String
Private mlngCd_operacao_contabil            As Long
Private mblnregistro_integrado              As Boolean
Private mlngcd_operacao_baixa               As Long
Private mblnbaixa_integrado                 As Boolean
Private mblnproveniente_rateio              As Boolean
Private mlngparc_origem_baixa               As Long
Private mlngId_carteira                     As Long
Private mstrcod_motivo                      As String
Private mstrdesc_motivo                     As String
Private mlngultima_mov_cnab                 As Long
Private mdtaultima_data_cnab                As Date
Private mlngcd_cobranca                     As Long
Private mlngseq_remessa                     As Long
Private mdtadata_integracao                 As Date
Private mcolRateio                          As colRateio
Private menumLancDup                        As enuLancDup
'Vinicius Elyseu(01/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
Private mblnRemessa                         As Boolean
Private mblnBoleto                          As Boolean

Public Property Get PagRec() As String
    PagRec = mstrPagRec
End Property

Public Property Let PagRec(ByVal Valor As String)
    mstrPagRec = Valor
End Property

Public Property Get Codigo_Nota() As String
    Codigo_Nota = mlngCodigo_Nota
End Property

Public Property Let Codigo_Nota(ByVal Valor As String)
    mlngCodigo_Nota = Valor
End Property

Public Property Get Parcela() As Long
    Parcela = mlngPARCELA
End Property

Public Property Let Parcela(ByVal Valor As Long)
    mlngPARCELA = Valor
End Property

Public Property Get Empresa() As String
    Empresa = mstrEmpresa
End Property

Public Property Let Empresa(ByVal Valor As String)
    mstrEmpresa = Valor
End Property

Public Property Get Tipo() As String
    Tipo = mstrTipo
End Property

Public Property Let Tipo(ByVal Valor As String)
    mstrTipo = Valor
End Property

Public Property Get Descricao() As String
    Descricao = mstrDescricao
End Property

Public Property Let Descricao(ByVal Valor As String)
    mstrDescricao = Valor
End Property

Public Property Get Emissao() As Date
    Emissao = mdtaEmissao
End Property

Public Property Let Emissao(ByVal Valor As Date)
    mdtaEmissao = Valor
End Property

Public Property Get Vencimento() As Date
    Vencimento = mdtaVencimento
End Property

Public Property Let Vencimento(ByVal Valor As Date)
    mdtaVencimento = Valor
End Property

Public Property Get Pagamento() As Date
    Pagamento = mdtaPagamento
End Property

Public Property Let Pagamento(ByVal Valor As Date)
    mdtaPagamento = Valor
End Property

Public Property Get Liberacao() As Date
    Liberacao = mdtaLiberacao
End Property

Public Property Let Liberacao(ByVal Valor As Date)
    mdtaLiberacao = Valor
End Property

Public Property Get ValorOriginal() As Double
    ValorOriginal = mdblValorOriginal
End Property

Public Property Let ValorOriginal(ByVal Valor As Double)
    mdblValorOriginal = Valor
End Property

Public Property Get Acrescimo() As Double
    Acrescimo = mdblAcrescimo
End Property

Public Property Let Acrescimo(ByVal Valor As Double)
    mdblAcrescimo = Valor
End Property

Public Property Get Abatimento() As Double
    Abatimento = mdblAbatimento
End Property

Public Property Let Abatimento(ByVal Valor As Double)
    mdblAbatimento = Valor
End Property

Public Property Get Banco() As Long
    Banco = mlngBanco
End Property

Public Property Let Banco(ByVal Valor As Long)
    mlngBanco = Valor
End Property

Public Property Get conta() As Long
    conta = mlngConta
End Property

Public Property Let conta(ByVal Valor As Long)
    mlngConta = Valor
End Property

Public Property Get Centro() As Long
    Centro = mlngCentro
End Property

Public Property Let Centro(ByVal Valor As Long)
    mlngCentro = Valor
End Property

Public Property Get Cheque() As Long
    Cheque = mlngCheque
End Property

Public Property Let Cheque(ByVal Valor As Long)
    mlngCheque = Valor
End Property

Public Property Get Moeda() As String
    Moeda = mstrMoeda
End Property

Public Property Let Moeda(ByVal Valor As String)
    mstrMoeda = Valor
End Property

Public Property Get ValordaMoeda() As Double
    ValordaMoeda = mdblValordaMoeda
End Property

Public Property Let ValordaMoeda(ByVal Valor As Double)
    mdblValordaMoeda = Valor
End Property

Public Property Get Controle() As String
    Controle = mstrControle
End Property

Public Property Let Controle(ByVal Valor As String)
    mstrControle = Valor
End Property

Public Property Get Situacao() As String
    Situacao = mstrSituacao
End Property

Public Property Let Situacao(ByVal Valor As String)
    mstrSituacao = Valor
End Property

Public Property Get Enviado() As String
    Enviado = mstrEnviado
End Property

Public Property Let Enviado(ByVal Valor As String)
    mstrEnviado = Valor
End Property

Public Property Get Marcacao() As Boolean
    Marcacao = mblnMarcacao
End Property

Public Property Let Marcacao(ByVal Valor As Boolean)
    mblnMarcacao = Valor
End Property

Public Property Get Obs() As String
    Obs = mstrObs
End Property

Public Property Let Obs(ByVal Valor As String)
    mstrObs = Valor
End Property

Public Property Get Bordero() As Long
    Bordero = mlngBordero
End Property

Public Property Let Bordero(ByVal Valor As Long)
    mlngBordero = Valor
End Property

Public Property Get DataIntegracaoFatoGerador() As Date
    DataIntegracaoFatoGerador = mdtaDataIntegracaoFatoGerador
End Property

Public Property Let DataIntegracaoFatoGerador(ByVal Valor As Date)
    mdtaDataIntegracaoFatoGerador = Valor
End Property

Public Property Get DataIntegracaoFatoPagamento() As Date
    DataIntegracaoFatoPagamento = mdtaDataIntegracaoFatoPagamento
End Property

Public Property Let DataIntegracaoFatoPagamento(ByVal Valor As Date)
    mdtaDataIntegracaoFatoPagamento = Valor
End Property

Public Property Get usuario() As String
    usuario = mstrUsuario
End Property

Public Property Let usuario(ByVal Valor As String)
    mstrUsuario = Valor
End Property

Public Property Get Alteracao() As Date
    Alteracao = mdtaAlteracao
End Property

Public Property Let Alteracao(ByVal Valor As Date)
    mdtaAlteracao = Valor
End Property

Public Property Get LINDIG() As String
    LINDIG = mstrLINDIG
End Property

Public Property Let LINDIG(ByVal Valor As String)
    mstrLINDIG = Valor
End Property

Public Property Get Conciliado() As Boolean
    Conciliado = mblnConciliado
End Property

Public Property Let Conciliado(ByVal Valor As Boolean)
    mblnConciliado = Valor
End Property

Public Property Get Carteira() As String
    Carteira = mstrCarteira
End Property

Public Property Let Carteira(ByVal Valor As String)
    mstrCarteira = Valor
End Property

Public Property Get previsao() As Boolean
    previsao = mblnPrevisao
End Property

Public Property Let previsao(ByVal Valor As Boolean)
    mblnPrevisao = Valor
End Property

Public Property Get VlrMrD() As Double
    VlrMrD = mdblVlrMrD
End Property

Public Property Let VlrMrD(ByVal Valor As Double)
    mdblVlrMrD = Valor
End Property

Public Property Get PerMrD() As Double
    PerMrD = mdblPerMrD
End Property

Public Property Let PerMrD(ByVal Valor As Double)
    mdblPerMrD = Valor
End Property

'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
Public Property Get SeqNossoNumero() As String
    SeqNossoNumero = mstrSeqNossoNumero
End Property

Public Property Let SeqNossoNumero(ByVal Valor As String)
    mstrSeqNossoNumero = Valor
End Property

Public Property Get VlrDsP() As Double
    VlrDsP = mdblVlrDsP
End Property

Public Property Let VlrDsP(ByVal Valor As Double)
    mdblVlrDsP = Valor
End Property

Public Property Get PerJur() As Double
    PerJur = mdblPerJur
End Property

Public Property Let PerJur(ByVal Valor As Double)
    mdblPerJur = Valor
End Property

Public Property Get VlrJur() As Double
    VlrJur = mdblVlrJur
End Property

Public Property Let VlrJur(ByVal Valor As Double)
    mdblVlrJur = Valor
End Property

Public Property Get CODPRJ() As Long
    CODPRJ = mlngCODPRJ
End Property

Public Property Let CODPRJ(ByVal Valor As Long)
    mlngCODPRJ = Valor
End Property

Public Property Get CODFPG() As Long
    CODFPG = mlngCODFPG
End Property

Public Property Let CODFPG(ByVal Valor As Long)
    mlngCODFPG = Valor
End Property

Public Property Get CheBan() As Long
    CheBan = mlngCheBan
End Property

Public Property Let CheBan(ByVal Valor As Long)
    mlngCheBan = Valor
End Property

Public Property Get CheAge() As String
    CheAge = mstrCheAge
End Property

Public Property Let CheAge(ByVal Valor As String)
    mstrCheAge = Valor
End Property

Public Property Get CheEmi() As String
    CheEmi = mstrCheEmi
End Property

Public Property Let CheEmi(ByVal Valor As String)
    mstrCheEmi = Valor
End Property

Public Property Get CheCco() As String
    CheCco = mstrCheCco
End Property

Public Property Let CheCco(ByVal Valor As String)
    mstrCheCco = Valor
End Property

Public Property Get PerMul() As Double
    PerMul = mdblPerMul
End Property

Public Property Let PerMul(ByVal Valor As Double)
    mdblPerMul = Valor
End Property

Public Property Get VlrMul() As Double
    VlrMul = mdblVlrMul
End Property

Public Property Let VlrMul(ByVal Valor As Double)
    mdblVlrMul = Valor
End Property

Public Property Get AGECCE() As String
    AGECCE = mstrAGECCE
End Property

Public Property Let AGECCE(ByVal Valor As String)
    mstrAGECCE = Valor
End Property

Public Property Get NOSNUM() As String
    NOSNUM = mstrNOSNUM
End Property

Public Property Let NOSNUM(ByVal Valor As String)
    mstrNOSNUM = Valor
End Property

Public Property Get CodBar() As String
    CodBar = mstrCODBAR
End Property

Public Property Let CodBar(ByVal Valor As String)
    mstrCODBAR = Valor
End Property

Public Property Get LOCPAG() As String
    LOCPAG = mstrLOCPAG
End Property

Public Property Let LOCPAG(ByVal Valor As String)
    mstrLOCPAG = Valor
End Property

Public Property Get INSTRU() As String
    INSTRU = mstrINSTRU
End Property

Public Property Let INSTRU(ByVal Valor As String)
    mstrINSTRU = Valor
End Property

Public Property Get cd_operacao_contabil() As Long
    cd_operacao_contabil = mlngCd_operacao_contabil
End Property

Public Property Let cd_operacao_contabil(ByVal Valor As Long)
    mlngCd_operacao_contabil = Valor
End Property

Public Property Get registro_integrado() As Boolean
    registro_integrado = mblnregistro_integrado
End Property

Public Property Let registro_integrado(ByVal Valor As Boolean)
    mblnregistro_integrado = Valor
End Property

Public Property Get cd_operacao_baixa() As Long
    cd_operacao_baixa = mlngcd_operacao_baixa
End Property

Public Property Let cd_operacao_baixa(ByVal Valor As Long)
    mlngcd_operacao_baixa = Valor
End Property

Public Property Get baixa_integrado() As Boolean
    baixa_integrado = mblnbaixa_integrado
End Property

Public Property Let baixa_integrado(ByVal Valor As Boolean)
    mblnbaixa_integrado = Valor
End Property

Public Property Get proveniente_rateio() As Boolean
    proveniente_rateio = mblnproveniente_rateio
End Property

Public Property Let proveniente_rateio(ByVal Valor As Boolean)
    mblnproveniente_rateio = Valor
End Property

Public Property Get parc_origem_baixa() As Long
    parc_origem_baixa = mlngparc_origem_baixa
End Property

Public Property Let parc_origem_baixa(ByVal Valor As Long)
    mlngparc_origem_baixa = Valor
End Property

Public Property Get Id_carteira() As Long
    Id_carteira = mlngId_carteira
End Property

Public Property Let Id_carteira(ByVal Valor As Long)
    mlngId_carteira = Valor
End Property

Public Property Get cod_motivo() As String
    cod_motivo = mstrcod_motivo
End Property

Public Property Let cod_motivo(ByVal Valor As String)
    mstrcod_motivo = Valor
End Property

Public Property Get desc_motivo() As String
    desc_motivo = mstrdesc_motivo
End Property

Public Property Let desc_motivo(ByVal Valor As String)
    mstrdesc_motivo = Valor
End Property

Public Property Get ultima_mov_cnab() As Long
    ultima_mov_cnab = mlngultima_mov_cnab
End Property

Public Property Let ultima_mov_cnab(ByVal Valor As Long)
    mlngultima_mov_cnab = Valor
End Property

Public Property Get ultima_data_cnab() As Date
    ultima_data_cnab = mdtaultima_data_cnab
End Property

Public Property Let ultima_data_cnab(ByVal Valor As Date)
    mdtaultima_data_cnab = Valor
End Property

Public Property Get cd_cobranca() As Long
    cd_cobranca = mlngcd_cobranca
End Property

Public Property Let cd_cobranca(ByVal Valor As Long)
    mlngcd_cobranca = Valor
End Property

Public Property Get seq_remessa() As Long
    seq_remessa = mlngseq_remessa
End Property

Public Property Let seq_remessa(ByVal Valor As Long)
    mlngseq_remessa = Valor
End Property

Public Property Get data_integracao() As Date
    data_integracao = mdtadata_integracao
End Property

Public Property Let data_integracao(ByVal Valor As Date)
    mdtadata_integracao = Valor
End Property

Public Property Get ValorTotal() As Double
    ValorTotal = mdblValorOriginal + mdblAcrescimo - mdblAbatimento
End Property

Public Property Get LancDup() As enuLancDup
    LancDup = menumLancDup
End Property

Public Property Let LancDup(ByVal enumLancDup As enuLancDup)
    menumLancDup = enumLancDup
End Property

Public Property Get Col_Rateio() As colRateio
    Set Col_Rateio = mcolRateio
End Property

Public Property Let Col_Rateio(ByVal Valor As colRateio)
    Set mcolRateio = Valor
End Property
'Vinicius Elyseu(01/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
Public Property Get Remessa() As Long
    Remessa = mblnRemessa
End Property

Public Property Let Remessa(ByVal Valor As Long)
    mblnRemessa = Valor
End Property
Public Property Get Boleto() As Long
    Boleto = mblnBoleto
End Property

Public Property Let Boleto(ByVal Valor As Long)
    mblnBoleto = Valor
End Property


Private Sub Class_Initialize()
    Set mcolRateio = New colRateio
End Sub

Private Sub Class_Terminate()
    Set mcolRateio = Nothing
End Sub
