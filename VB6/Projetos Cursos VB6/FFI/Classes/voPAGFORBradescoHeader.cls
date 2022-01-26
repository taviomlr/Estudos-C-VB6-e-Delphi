VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "voPAGFORBradescoHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mintIdentificacaoRegistro   As Integer
Private mlngCodigoComunicacao       As Long
Private mintTipoInscricao           As Integer
Private mstrNumeroInscricaoCNPJ     As String
Private mintFilialCNPJ              As Integer
Private mintControleCNPJ            As Integer
Private mstrRazaoSocial             As String
Private mintTipoServico             As Integer
Private mintControle                As Integer
Private mintOrigemArquivo           As Integer
Private mlngNumeroRemessa           As Long
Private mlngNumeroRetorno           As Long
Private mstrDataGravacao            As String
Private mstrHoraGravacao            As String
Private mstrUnidadeDensidade        As String
Private mstrDensidadeGravacao       As String
Private mstrIdentificacaoModulo     As String
Private mlngTipoProcessamento       As Long
Private mstrReservadoEmpresa        As String
Private mstrNumeroLista             As String
Private mlngSequencial              As Long

Public Property Get IdentificacaoRegistro() As Integer
    IdentificacaoRegistro = mintIdentificacaoRegistro
End Property

Public Property Let IdentificacaoRegistro(ByVal Valor As Integer)
    mintIdentificacaoRegistro = Valor
End Property

Public Property Get CodigoComunicacao() As Long
    CodigoComunicacao = mlngCodigoComunicacao
End Property

Public Property Let CodigoComunicacao(ByVal Valor As Long)
    mlngCodigoComunicacao = Valor
End Property

Public Property Get TipoInscricao() As Integer
    TipoInscricao = mintTipoInscricao
End Property

Public Property Let TipoInscricao(ByVal Valor As Integer)
    mintTipoInscricao = Valor
End Property

Public Property Get NumeroInscricaoCNPJ() As String
    NumeroInscricaoCNPJ = mstrNumeroInscricaoCNPJ
End Property

Public Property Let NumeroInscricaoCNPJ(ByVal Valor As String)
    mstrNumeroInscricaoCNPJ = Valor
End Property

Public Property Get FilialCNPJ() As Integer
    FilialCNPJ = mintFilialCNPJ
End Property

Public Property Let FilialCNPJ(ByVal Valor As Integer)
    mintFilialCNPJ = Valor
End Property

Public Property Get ControleCNPJ() As Integer
    ControleCNPJ = mintControleCNPJ
End Property

Public Property Let ControleCNPJ(ByVal Valor As Integer)
    mintControleCNPJ = Valor
End Property

Public Property Get RazaoSocial() As String
    RazaoSocial = mstrRazaoSocial
End Property

Public Property Let RazaoSocial(ByVal Valor As String)
    mstrRazaoSocial = Valor
End Property

Public Property Get TipoServico() As Integer
    TipoServico = mintTipoServico
End Property

Public Property Let TipoServico(ByVal Valor As Integer)
    mintTipoServico = Valor
End Property

Public Property Get Controle() As Integer
    Controle = mintControle
End Property

Public Property Let Controle(ByVal Valor As Integer)
    mintControle = Valor
End Property

Public Property Get OrigemArquivo() As Integer
    OrigemArquivo = mintOrigemArquivo
End Property

Public Property Let OrigemArquivo(ByVal Valor As Integer)
    mintOrigemArquivo = Valor
End Property

Public Property Get NumeroRemessa() As Long
    NumeroRemessa = mlngNumeroRemessa
End Property

Public Property Let NumeroRemessa(ByVal Valor As Long)
    mlngNumeroRemessa = Valor
End Property

Public Property Get NumeroRetorno() As Long
    NumeroRetorno = mlngNumeroRetorno
End Property

Public Property Let NumeroRetorno(ByVal Valor As Long)
    mlngNumeroRetorno = Valor
End Property

Public Property Get DataGravacao() As String
    DataGravacao = mstrDataGravacao
End Property

Public Property Let DataGravacao(ByVal Valor As String)
    mstrDataGravacao = Valor
End Property

Public Property Get HoraGravacao() As String
    HoraGravacao = mstrHoraGravacao
End Property

Public Property Let HoraGravacao(ByVal Valor As String)
    mstrHoraGravacao = Valor
End Property

Public Property Get UnidadeDensidade() As String
    UnidadeDensidade = mstrUnidadeDensidade
End Property

Public Property Let UnidadeDensidade(ByVal Valor As String)
    mstrUnidadeDensidade = Valor
End Property

Public Property Get DensidadeGravacao() As String
    DensidadeGravacao = mstrDensidadeGravacao
End Property

Public Property Let DensidadeGravacao(ByVal Valor As String)
    mstrDensidadeGravacao = Valor
End Property

Public Property Get IdentificacaoModulo() As String
    IdentificacaoModulo = mstrIdentificacaoModulo
End Property

Public Property Let IdentificacaoModulo(ByVal Valor As String)
    mstrIdentificacaoModulo = Valor
End Property

Public Property Get TipoProcessamento() As Long
    TipoProcessamento = mlngTipoProcessamento
End Property

Public Property Let TipoProcessamento(ByVal Valor As Long)
    mlngTipoProcessamento = Valor
End Property

Public Property Get ReservadoEmpresa() As String
    ReservadoEmpresa = mstrReservadoEmpresa
End Property

Public Property Let ReservadoEmpresa(ByVal Valor As String)
    mstrReservadoEmpresa = Valor
End Property

Public Property Get NumeroLista() As String
    NumeroLista = mstrNumeroLista
End Property

Public Property Let NumeroLista(ByVal Valor As String)
    mstrNumeroLista = Valor
End Property

Public Property Get Sequencial() As Long
    Sequencial = mlngSequencial
End Property

Public Property Let Sequencial(ByVal Valor As Long)
    mlngSequencial = Valor
End Property
