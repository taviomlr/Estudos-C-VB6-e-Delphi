VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "voPAGFORBradescoDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mintIdentificacaoDetalhe    As Integer
Private mintTipoInscricao           As Integer
Private mstrNumeroInscricaoCNPJ     As String
Private mintFilialCNPJ              As Integer
Private mintControleCNPJ            As Integer
Private mstrRazaoSocial             As String
Private mstrEnderecoFornec          As String
Private mlngCEPFornec               As Long
Private mlngCEPCompFornec           As Long
Private mintCodBancoFornec          As Integer
Private mlngCodAgenciaFornec        As Long
Private mstrDigitoAgenciaFornec     As String
Private mstrContaCorrenteFornec     As String
Private mstrDigContaCorrenteFornec  As String
Private mstrNumeroPagt              As String
Private mintCarteira                As Integer
Private mstrNossoNumero             As String
Private mstrDataVencimento          As String
Private mstrDataEmissao             As String
Private mstrDataLimiteDesconto      As String
Private mintFatorVencimento         As Integer
Private mcurValorDocumento          As Currency
Private mcurValorPagamento          As Currency
Private mcurValorDesconto           As Currency
Private mcurValorAcrescimo          As Currency
Private mintTipoDocumento           As Integer
Private mdblNumeroNF                As Double
Private mstrSerieNF                 As String
Private mintModalidadePagt          As Integer
Private mstrDataEfetivacaoPagt      As String
Private mstrMoeda                   As String 'Mandar 3 brancos
Private mstrSituacaoAgendamento     As String 'Preencher fixo 01
Private mstrTipoMovimento           As String
Private mintCodigoMovimento         As Integer
Private mstrInfComplementares       As String
Private mstrTipoDocCompensado       As String
Private mintCodigoFinalidade        As Integer
Private mintCodigoLancamento        As Integer
Private mintTipoContaForn           As Integer
Private mlngSequencialRegistro      As Long
Private mstrContaComplementar       As String
Private mstrOrigemDoc               As String
Private mstrCodigoBarras            As String
Private mintTipoConta               As Integer

Public Property Get IdentificacaoDetalhe() As Integer
    IdentificacaoDetalhe = mintIdentificacaoDetalhe
End Property

Public Property Let IdentificacaoDetalhe(ByVal value As Integer)
    mintIdentificacaoDetalhe = value
End Property

Public Property Get TipoInscricao() As Integer
    TipoInscricao = mintTipoInscricao
End Property

Public Property Let TipoInscricao(ByVal value As Integer)
    mintTipoInscricao = value
End Property

Public Property Get NumeroInscricaoCNPJ() As String
    NumeroInscricaoCNPJ = mstrNumeroInscricaoCNPJ
End Property

Public Property Let NumeroInscricaoCNPJ(ByVal value As String)
    mstrNumeroInscricaoCNPJ = value
End Property

Public Property Get FilialCNPJ() As Integer
    FilialCNPJ = mintFilialCNPJ
End Property

Public Property Let FilialCNPJ(ByVal value As Integer)
    mintFilialCNPJ = value
End Property

Public Property Get ControleCNPJ() As Integer
    ControleCNPJ = mintControleCNPJ
End Property

Public Property Let ControleCNPJ(ByVal value As Integer)
    mintControleCNPJ = value
End Property

Public Property Get RazaoSocial() As String
    RazaoSocial = mstrRazaoSocial
End Property

Public Property Let RazaoSocial(ByVal value As String)
    mstrRazaoSocial = value
End Property

Public Property Get EnderecoFornec() As String
    EnderecoFornec = mstrEnderecoFornec
End Property

Public Property Let EnderecoFornec(ByVal value As String)
    mstrEnderecoFornec = value
End Property

Public Property Get CEPFornec() As Long
    CEPFornec = mlngCEPFornec
End Property

Public Property Let CEPFornec(ByVal value As Long)
    mlngCEPFornec = value
End Property

Public Property Get CEPCompFornec() As Long
    CEPCompFornec = mlngCEPCompFornec
End Property

Public Property Let CEPCompFornec(ByVal value As Long)
    mlngCEPCompFornec = value
End Property

Public Property Get CodBancoFornec() As Integer
    CodBancoFornec = mintCodBancoFornec
End Property

Public Property Let CodBancoFornec(ByVal value As Integer)
    mintCodBancoFornec = value
End Property

Public Property Get CodAgenciaFornec() As Long
    CodAgenciaFornec = mlngCodAgenciaFornec
End Property

Public Property Let CodAgenciaFornec(ByVal value As Long)
    mlngCodAgenciaFornec = value
End Property

Public Property Get DigitoAgenciaFornec() As String
    DigitoAgenciaFornec = mstrDigitoAgenciaFornec
End Property

Public Property Let DigitoAgenciaFornec(ByVal value As String)
    mstrDigitoAgenciaFornec = value
End Property

Public Property Get ContaCorrenteFornec() As String
    ContaCorrenteFornec = mstrContaCorrenteFornec
End Property

Public Property Let ContaCorrenteFornec(ByVal value As String)
    mstrContaCorrenteFornec = value
End Property

Public Property Get DigContaCorrenteFornec() As String
    DigContaCorrenteFornec = mstrDigContaCorrenteFornec
End Property

Public Property Let DigContaCorrenteFornec(ByVal value As String)
    mstrDigContaCorrenteFornec = value
End Property

Public Property Get NumeroPagt() As String
    NumeroPagt = mstrNumeroPagt
End Property

Public Property Let NumeroPagt(ByVal value As String)
    mstrNumeroPagt = value
End Property

Public Property Get Carteira() As Integer
    Carteira = mintCarteira
End Property

Public Property Let Carteira(ByVal value As Integer)
    mintCarteira = value
End Property

Public Property Get NossoNumero() As String
    NossoNumero = mstrNossoNumero
End Property

Public Property Let NossoNumero(ByVal value As String)
    mstrNossoNumero = value
End Property

Public Property Get DataVencimento() As String
    DataVencimento = mstrDataVencimento
End Property

Public Property Let DataVencimento(ByVal value As String)
    mstrDataVencimento = value
End Property

Public Property Get dataEmissao() As String
    dataEmissao = mstrDataEmissao
End Property

Public Property Let dataEmissao(ByVal value As String)
    mstrDataEmissao = value
End Property

Public Property Get DataLimiteDesconto() As String
    DataLimiteDesconto = mstrDataLimiteDesconto
End Property

Public Property Let DataLimiteDesconto(ByVal value As String)
    mstrDataLimiteDesconto = value
End Property

Public Property Get FatorVencimento() As Integer
    FatorVencimento = mintFatorVencimento
End Property

Public Property Let FatorVencimento(ByVal value As Integer)
    mintFatorVencimento = value
End Property

Public Property Get ValorDocumento() As Currency
    ValorDocumento = mcurValorDocumento
End Property

Public Property Let ValorDocumento(ByVal value As Currency)
    mcurValorDocumento = value
End Property

Public Property Get ValorPagamento() As Currency
    ValorPagamento = mcurValorPagamento
End Property

Public Property Let ValorPagamento(ByVal value As Currency)
    mcurValorPagamento = value
End Property

Public Property Get ValorDesconto() As Currency
    ValorDesconto = mcurValorDesconto
End Property

Public Property Let ValorDesconto(ByVal value As Currency)
    mcurValorDesconto = value
End Property

Public Property Get valorAcrescimo() As Currency
    valorAcrescimo = mcurValorAcrescimo
End Property

Public Property Let valorAcrescimo(ByVal value As Currency)
    mcurValorAcrescimo = value
End Property

Public Property Get TipoDocumento() As Integer
    TipoDocumento = mintTipoDocumento
End Property

Public Property Let TipoDocumento(ByVal value As Integer)
    mintTipoDocumento = value
End Property

Public Property Get NumeroNF() As Double
    NumeroNF = mdblNumeroNF
End Property

Public Property Let NumeroNF(ByVal value As Double)
    mdblNumeroNF = value
End Property

Public Property Get SerieNF() As String
    SerieNF = mstrSerieNF
End Property

Public Property Let SerieNF(ByVal value As String)
    mstrSerieNF = value
End Property

Public Property Get ModalidadePagt() As Integer
    ModalidadePagt = mintModalidadePagt
End Property

Public Property Let ModalidadePagt(ByVal value As Integer)
    mintModalidadePagt = value
End Property

Public Property Get DataEfetivacaoPagt() As String
    DataEfetivacaoPagt = mstrDataEfetivacaoPagt
End Property

Public Property Let DataEfetivacaoPagt(ByVal value As String)
    mstrDataEfetivacaoPagt = value
End Property

Public Property Get Moeda() As String
    Moeda = mstrMoeda
End Property

Public Property Let Moeda(ByVal value As String)
    mstrMoeda = value
End Property

Public Property Get SituacaoAgendamento() As String
    SituacaoAgendamento = mstrSituacaoAgendamento
End Property

Public Property Let SituacaoAgendamento(ByVal value As String)
    mstrSituacaoAgendamento = value
End Property

Public Property Get TipoMovimento() As String
    TipoMovimento = mstrTipoMovimento
End Property

Public Property Let TipoMovimento(ByVal value As String)
    mstrTipoMovimento = value
End Property

Public Property Get CodigoMovimento() As Integer
    CodigoMovimento = mintCodigoMovimento
End Property

Public Property Let CodigoMovimento(ByVal value As Integer)
    mintCodigoMovimento = value
End Property

Public Property Get InfComplementares() As String
    InfComplementares = mstrInfComplementares
End Property

Public Property Let InfComplementares(ByVal value As String)
    mstrInfComplementares = value
End Property

Public Property Get TipoDocCompensado() As String
    TipoDocCompensado = mstrTipoDocCompensado
End Property

Public Property Let TipoDocCompensado(ByVal value As String)
    mstrTipoDocCompensado = value
End Property

Public Property Get CodigoFinalidade() As Integer
    CodigoFinalidade = mintCodigoFinalidade
End Property

Public Property Let CodigoFinalidade(ByVal value As Integer)
    mintCodigoFinalidade = value
End Property

Public Property Get CodigoLancamento() As Integer
    CodigoLancamento = mintCodigoLancamento
End Property

Public Property Let CodigoLancamento(ByVal value As Integer)
    mintCodigoLancamento = value
End Property

Public Property Get TipoContaForn() As Integer
    TipoContaForn = mintTipoContaForn
End Property

Public Property Let TipoContaForn(ByVal value As Integer)
    mintTipoContaForn = value
End Property

Public Property Get TipoConta() As Integer
    TipoConta = mintTipoConta
End Property

Public Property Let TipoConta(ByVal value As Integer)
    mintTipoConta = value
End Property

Public Property Get SequencialRegistro() As Long
    SequencialRegistro = mlngSequencialRegistro
End Property

Public Property Let SequencialRegistro(ByVal value As Long)
    mlngSequencialRegistro = value
End Property

Public Property Get ContaComplementar() As String
    ContaComplementar = mstrContaComplementar
End Property

Public Property Let ContaComplementar(ByVal value As String)
    mstrContaComplementar = value
End Property

Public Property Let OrigemDoc(ByVal value As String)
    mstrOrigemDoc = value
End Property

Public Property Let CodigoBarras(ByVal value As String)
    mstrCodigoBarras = value
End Property

Public Property Get CodigoBarras() As String
    CodigoBarras = mstrCodigoBarras
End Property

Public Function Equals(obj As voPAGFORBradescoDetail) As Boolean
    Equals = (obj.SequencialRegistro = mlngSequencialRegistro)
End Function
