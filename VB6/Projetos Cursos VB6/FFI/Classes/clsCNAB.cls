VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCNAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateBoleto Lib "EBSBoleto.dll" (iBanco As Long, iAgencia As Long, _
          iDVAgencia As Long, iConta As Long, iDVConta As Long, ByVal sCodCedente As String, _
          ByVal sNossoNum As String, ByVal sCarteira As String, nValor As Double, _
          dDataVenc As Date, iMoeda As Long, ByVal sRegistro As String, ByVal sConvenio As String, _
          ByVal sNrDocumento As String, ByVal sCodEmpresa As String, ByVal outLinhaDigitavel As String, _
          ByVal outCodBarras As String, ByVal outNossoNumeroImp As String, ByVal outAgeCodCedImp As String) As Long
Private Declare Sub SetConfigFileToLoggerEBSBoleto Lib "EBSBoleto.dll" (ByVal configFile As String)
Private mlngIdentificadorBanco   As Long
Private mstrAgencia              As String
Private mlngDigitoAgencia        As Long
Private mstrNumeroConta          As String
Private mlngDigitoConta          As Long
Private mstrCodigoCedente        As String
Private mstrCarteira             As String
Private mdblValorDocumento       As Double
Private mdatDataVencimento       As Date
Private mstrLinhaDigitavel       As String
Private mstrCodigoBarras         As String
Private mstrNossoNumero          As String
Private mstrAgenciaCodigoCedente As String
Private mstrConvenio             As String
Private mstrNumeroDocumento      As String
Private mstrCodigoEmpresa        As String
Private mstrTipoCobranca         As String
Private mobjCNAB240              As clsCNAB240Caixa
Private mlngPadrao               As Long

Public Property Get IdentificadorBanco() As Long
    IdentificadorBanco = mlngIdentificadorBanco
End Property

Public Property Let IdentificadorBanco(ByVal lngIdentificadorBanco As Long)
    mlngIdentificadorBanco = lngIdentificadorBanco
End Property

Public Property Get Agencia() As String
    Agencia = mstrAgencia
End Property

Public Property Let Agencia(ByVal strAgencia As String)
    mstrAgencia = strAgencia
End Property

Public Property Get DigitoAgencia() As Long
    DigitoAgencia = mlngDigitoAgencia
End Property

Public Property Let DigitoAgencia(ByVal lngDigitoAgencia As Long)
    mlngDigitoAgencia = lngDigitoAgencia
End Property

Public Property Get NumeroConta() As String
    NumeroConta = mstrNumeroConta
End Property

Public Property Let NumeroConta(ByVal strNumeroConta As String)
    mstrNumeroConta = strNumeroConta
End Property

Public Property Get DigitoConta() As Long
    DigitoConta = mlngDigitoConta
End Property

Public Property Let DigitoConta(ByVal lngDigitoConta As Long)
    mlngDigitoConta = lngDigitoConta
End Property

Public Property Get CodigoCedente() As String
    CodigoCedente = mstrCodigoCedente
End Property

Public Property Let CodigoCedente(ByVal strCodigoCedente As String)
    mstrCodigoCedente = strCodigoCedente
End Property

Public Property Get Carteira() As String
    Carteira = mstrCarteira
End Property

Public Property Let Carteira(ByVal strCarteira As String)
    mstrCarteira = strCarteira
End Property

Public Property Get ValorDocumento() As Double
    ValorDocumento = mdblValorDocumento
End Property

Public Property Let ValorDocumento(ByVal dblValorDocumento As Double)
    mdblValorDocumento = dblValorDocumento
End Property

Public Property Get DataVencimento() As Date
    DataVencimento = mdatDataVencimento
End Property

Public Property Let DataVencimento(ByVal datDataVencimento As Date)
    mdatDataVencimento = datDataVencimento
End Property

Public Property Get LinhaDigitavel() As String
    LinhaDigitavel = mstrLinhaDigitavel
End Property

Public Property Let LinhaDigitavel(ByVal strLinhaDigitavel As String)
    mstrLinhaDigitavel = strLinhaDigitavel
End Property

Public Property Get CodigoBarras() As String
    CodigoBarras = mstrCodigoBarras
End Property

Public Property Let CodigoBarras(ByVal strCodigoBarras As String)
    mstrCodigoBarras = strCodigoBarras
End Property

Public Property Get NossoNumero() As String
    NossoNumero = mstrNossoNumero
End Property

Public Property Let NossoNumero(ByVal strNossoNumero As String)
    mstrNossoNumero = strNossoNumero
End Property

Public Property Get AgenciaCodigoCedente() As String
    AgenciaCodigoCedente = mstrAgenciaCodigoCedente
End Property

Public Property Let AgenciaCodigoCedente(ByVal strAgenciaCodigoCedente As String)
    mstrAgenciaCodigoCedente = strAgenciaCodigoCedente
End Property

Public Property Get TipoCobranca() As String
    TipoCobranca = mstrTipoCobranca
End Property

Public Property Let TipoCobranca(ByVal strTipoCobranca As String)
    mstrTipoCobranca = strTipoCobranca
End Property

Public Property Get Convenio() As String
    Convenio = mstrConvenio
End Property

Public Property Let Convenio(ByVal strConvenio As String)
    mstrConvenio = strConvenio
End Property

Public Property Get NumeroDocumento() As String
    NumeroDocumento = mstrNumeroDocumento
End Property

Public Property Let NumeroDocumento(ByVal strNumeroDocumento As String)
    mstrNumeroDocumento = strNumeroDocumento
End Property

Public Property Get CodigoEmpresa() As String
    CodigoEmpresa = mstrCodigoEmpresa
End Property

Public Property Let CodigoEmpresa(ByVal strCodigoEmpresa As String)
    mstrCodigoEmpresa = strCodigoEmpresa
End Property

Public Property Let PadraoCNAB(ByVal NewVal As Long)
    mlngPadrao = NewVal
End Property

Public Function CriarBoleto(ByVal dblProximoSequencial As Double) As Boolean
    If mlngIdentificadorBanco = 104 Then
        If mlngPadrao = 240 Then
            Set mobjCNAB240 = New clsCNAB240Caixa
            With mobjCNAB240
                .IdentificacaoBanco = "104"
                .DataVencimento = CStr(mdatDataVencimento)
                .ValorDocumento = CStr(mdblValorDocumento)
                .Agencia = mstrAgencia
                .NumeroContaCorrente = mstrNumeroConta
                .ProximoNossoNumero = dblProximoSequencial
                If mstrCarteira = "CR" Then
                    .Carteira = 12
                Else
                    .Carteira = 14
                End If
                CriarBoleto = .CriarBoleto
                mstrNossoNumero = .NossoNumero
                mstrCodigoBarras = .CodigoBarras
                mstrLinhaDigitavel = .LinhaDigitavelFormatada
            End With
        Else
            mstrLinhaDigitavel = Space(59)
            mstrCodigoBarras = Space(51)
            mstrNossoNumero = Space(25)
            mstrAgenciaCodigoCedente = Space(25)
            CriarBoleto = (CreateBoleto(mlngIdentificadorBanco, CLng(mstrAgencia), mlngDigitoAgencia, CLng(mstrNumeroConta), mlngDigitoConta, mstrCodigoCedente, CStr(dblProximoSequencial), mstrCarteira, mdblValorDocumento, mdatDataVencimento, 9, mstrTipoCobranca, mstrConvenio, mstrNumeroDocumento, mstrCodigoEmpresa, mstrLinhaDigitavel, mstrCodigoBarras, mstrNossoNumero, mstrAgenciaCodigoCedente) = 0)
            Call RemoveUltimo(mstrLinhaDigitavel)
            Call RemoveUltimo(mstrCodigoBarras)
            Call RemoveUltimo(mstrNossoNumero)
            Call RemoveUltimo(mstrAgenciaCodigoCedente)
        End If
    Else
        mstrLinhaDigitavel = Space(59)
        mstrCodigoBarras = Space(51)
        mstrNossoNumero = Space(25)
        mstrAgenciaCodigoCedente = Space(25)
        CriarBoleto = (CreateBoleto(mlngIdentificadorBanco, CLng(mstrAgencia), mlngDigitoAgencia, CLng(mstrNumeroConta), mlngDigitoConta, mstrCodigoCedente, CStr(dblProximoSequencial), mstrCarteira, mdblValorDocumento, mdatDataVencimento, 9, mstrTipoCobranca, mstrConvenio, mstrNumeroDocumento, mstrCodigoEmpresa, mstrLinhaDigitavel, mstrCodigoBarras, mstrNossoNumero, mstrAgenciaCodigoCedente) = 0)
        Call RemoveUltimo(mstrLinhaDigitavel)
        Call RemoveUltimo(mstrCodigoBarras)
        Call RemoveUltimo(mstrNossoNumero)
        Call RemoveUltimo(mstrAgenciaCodigoCedente)
    End If
End Function

'Descrição..: Sub utilizada para remover o ultimo caracter de uma String
'               a sub foi criada para resolver o problema da EBSBoleto que
'               aloca caracteres loucos no fim da string, e é necessário
'               remove-los.
'Parametros.: [String] Texto retornado pela chamada da função em Delphi.
Private Sub RemoveUltimo(ByRef sTexto As String)
    sTexto = Trim$(sTexto)
    sTexto = Left(sTexto, Len(sTexto) - 1)
End Sub

