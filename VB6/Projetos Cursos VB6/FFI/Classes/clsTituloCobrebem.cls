VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTituloCobrebem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrNumeroDocumento                     As String
Private mstrParcela                             As String
Private mstrTipo                                As String
Private mstrNomesacado                          As String
Private mstrCpfsacado                           As String
Private mstrCnpjsacado                          As String
Private mstrEnderecosacado                      As String
Private mstrBairroSacado                        As String
Private mstrCidadesacado                        As String
Private mstrEstadosacado                        As String
Private mstrCepsacado                           As String
Private mdteDatadocumento                       As Date
Private mdteDatavencimento                      As Date
Private mdteDataprocessamento                   As Date
Private mdblValorDocumento                      As Double
Private mdblPercentualjurosdiaatraso            As Double
Private mdblPercentualmultaatraso               As Double
Private mdblPercentualdesconto                  As Double
Private mdblValoroutrosacrescimos               As Double
Private mdblValorAbatimento                     As Double
Private mdblValorOriginal                       As Double
Private mstrLinhaDigitavel                      As String
Private mstrCodigoBarras                        As String
Private mstrNossoNumero                         As String
'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
Private mstrSequencianossonumero                As String
Private mstrLocalPagamento                      As String
Private mstrEmpresa                             As String
Private mstrDemonstrativo                       As String
Private mstrInstrucoesCaixa                     As String
Private mstrOrigem                              As String
Private mstrDiretorio                           As String
Private mstrArquivo                             As String
Private mdteDataPagamentoRealizado              As Date
Private mstrMensagemValidacao                   As String
Private mstrMensagemBaixados                    As String
Private mlngEnterpriseId                        As Long
Private mlngCdEstabelecimento                   As Long
Private mlngSequencial_remessa                  As Long
'Pt. 94406 - Moacir Pfau 18/08/2009
Private mstrMensagemMotivo                      As String
Private mstrMensagemCodigo                      As String
Private mstrRelacaoNossoNumero                  As String
'Pt. 96589 - Moacir Pfau(08/02/2010)
Private mlngSequencial_remessa_Titulo           As Long
'pt.99257 - Moacir Pfau(30/06/2010)
Private mdblValorMulta                          As Double
Private mdblValorMora                           As Double
Private mlngCarteira                            As Long
Private mlngBanco                               As Long
Private mobjCobreBem                            As CobreBemX.ContaCorrente
Private mApp                                    As IApplication
'Projeto: #7373 - História: #6135 - Desenvolvimento: #7434 - Ivo Sousa(13/05/2013)
Private mstrRemessaStatus                       As String
'Projeto: #17081 - História: # - Desenvolvimento# - Moacir Pfau(20/01/2014)
Private mblnEmailboleto_enviado                 As Boolean
Private mstrCaminhoPDF                          As String
Private mobjAnexos                              As clscolAnexosEmail
'Vinicius Elyseu(02/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
Private mblnRemessa                             As Boolean
'Projeto: 100340 - Desenv.: 142893 - Ueder Budni (23/09/2016)
Private mstrPagRec                              As String

Const PDF_PRINTERNAME = "PDF Writer - bioPDF"
Const PRINTER_PROGID = "bioPDF.PDFPrinterSettings"

Public Enum ENUMTipoTitulo
    e_Boleto = 1
    e_Remessa = 2
    e_Retorno = 3
End Enum

Public Sub init(ByVal objApp As IApplication)
    Set mApp = objApp
End Sub

Public Property Let NumeroDocumento(Valor As String)
   mstrNumeroDocumento = Valor
End Property

Public Property Get NumeroDocumento() As String
   NumeroDocumento = mstrNumeroDocumento
End Property

Public Property Let Parcela(Valor As String)
   mstrParcela = Valor
End Property

Public Property Get Parcela() As String
   Parcela = mstrParcela
End Property

Public Property Let tipo(Valor As String)
   mstrTipo = Valor
End Property

Public Property Get tipo() As String
   tipo = mstrTipo
End Property

Public Property Let Nomesacado(Valor As String)
   mstrNomesacado = Valor
End Property

Public Property Get Nomesacado() As String
   Nomesacado = mstrNomesacado
End Property

Public Property Let Cpfsacado(Valor As String)
   mstrCpfsacado = Valor
End Property

Public Property Get Cpfsacado() As String
   Cpfsacado = mstrCpfsacado
End Property

Public Property Let Cnpjsacado(Valor As String)
   mstrCnpjsacado = Valor
End Property

Public Property Get Cnpjsacado() As String
   Cnpjsacado = mstrCnpjsacado
End Property

Public Property Let Enderecosacado(Valor As String)
   mstrEnderecosacado = Valor
End Property

Public Property Get Enderecosacado() As String
   Enderecosacado = mstrEnderecosacado
End Property

Public Property Let BairroSacado(Valor As String)
   mstrBairroSacado = Valor
End Property

Public Property Get BairroSacado() As String
   BairroSacado = mstrBairroSacado
End Property

Public Property Let Cidadesacado(Valor As String)
   mstrCidadesacado = Valor
End Property

Public Property Get Cidadesacado() As String
   Cidadesacado = mstrCidadesacado
End Property

Public Property Let Estadosacado(Valor As String)
   mstrEstadosacado = Valor
End Property

Public Property Get Estadosacado() As String
   Estadosacado = mstrEstadosacado
End Property

Public Property Let Cepsacado(Valor As String)
   mstrCepsacado = Valor
End Property

Public Property Get Cepsacado() As String
   Cepsacado = mstrCepsacado
End Property

Public Property Let Datadocumento(Valor As Date)
   mdteDatadocumento = Valor
End Property

Public Property Get Datadocumento() As Date
   Datadocumento = mdteDatadocumento
End Property

Public Property Let DataVencimento(Valor As Date)
   mdteDatavencimento = Valor
End Property

Public Property Get DataVencimento() As Date
   DataVencimento = mdteDatavencimento
End Property

Public Property Let Dataprocessamento(Valor As Date)
   mdteDataprocessamento = Valor
End Property

Public Property Get Dataprocessamento() As Date
   Dataprocessamento = mdteDataprocessamento
End Property

Public Property Let ValorDocumento(Valor As Double)
   mdblValorDocumento = Valor
End Property

Public Property Get ValorDocumento() As Double
   ValorDocumento = mdblValorDocumento
End Property

Public Property Let Percentualjurosdiaatraso(Valor As Double)
   mdblPercentualjurosdiaatraso = Valor
End Property

Public Property Get Percentualjurosdiaatraso() As Double
   Percentualjurosdiaatraso = mdblPercentualjurosdiaatraso
End Property

Public Property Let Percentualmultaatraso(Valor As Double)
   mdblPercentualmultaatraso = Valor
End Property

Public Property Get Percentualmultaatraso() As Double
   Percentualmultaatraso = mdblPercentualmultaatraso
End Property

Public Property Let Percentualdesconto(Valor As Double)
   mdblPercentualdesconto = Valor
End Property

Public Property Get Percentualdesconto() As Double
   Percentualdesconto = mdblPercentualdesconto
End Property

Public Property Let Valoroutrosacrescimos(Valor As Double)
   mdblValoroutrosacrescimos = Valor
End Property

Public Property Get Valoroutrosacrescimos() As Double
   Valoroutrosacrescimos = mdblValoroutrosacrescimos
End Property

Public Property Let ValorAbatimento(Valor As Double)
   mdblValorAbatimento = Valor
End Property

Public Property Get ValorAbatimento() As Double
   ValorAbatimento = mdblValorAbatimento
End Property

Public Property Let ValorOriginal(Valor As Double)
   mdblValorOriginal = Valor
End Property

Public Property Get ValorOriginal() As Double
   ValorOriginal = mdblValorOriginal
End Property

Public Property Let LinhaDigitavel(Valor As String)
   mstrLinhaDigitavel = Valor
End Property

Public Property Get LinhaDigitavel() As String
   LinhaDigitavel = mstrLinhaDigitavel
End Property

Public Property Let CodigoBarras(Valor As String)
   mstrCodigoBarras = Valor
End Property

Public Property Get CodigoBarras() As String
   CodigoBarras = mstrCodigoBarras
End Property

Public Property Let NossoNumero(Valor As String)
   mstrNossoNumero = Valor
End Property

Public Property Get NossoNumero() As String
   NossoNumero = mstrNossoNumero
End Property

Public Property Let Sequencianossonumero(Valor As String)
   mstrSequencianossonumero = Valor
End Property

Public Property Get Sequencianossonumero() As String
   Sequencianossonumero = mstrSequencianossonumero
End Property

Public Property Let localPagamento(Valor As String)
   mstrLocalPagamento = Valor
End Property

Public Property Get localPagamento() As String
   localPagamento = mstrLocalPagamento
End Property

Public Property Let Empresa(Valor As String)
   mstrEmpresa = Valor
End Property

Public Property Get Empresa() As String
   Empresa = mstrEmpresa
End Property

Public Property Let Origem(Valor As String)
   mstrOrigem = Valor
End Property

Public Property Get Origem() As String
   Origem = mstrOrigem
End Property

Public Property Get MensagemValidacao() As String
   MensagemValidacao = mstrMensagemValidacao
End Property

Public Property Get MensagemBaixado() As String
   MensagemBaixado = mstrMensagemBaixados
End Property

Public Property Let Sequencial_remessa(Valor As Long)
   mlngSequencial_remessa = Valor
End Property

Public Property Get Sequencial_remessa() As Long
   Sequencial_remessa = mlngSequencial_remessa
End Property

'Pt. 94406 - Moacir Pfau 18/08/2009
Public Property Let RelacaoNossoNumero(Valor As String)
   mstrRelacaoNossoNumero = Valor
End Property

Public Property Get RelacaoNossoNumero() As String
   RelacaoNossoNumero = mstrRelacaoNossoNumero
End Property

'Pt. 96589 - Moacir Pfau(08/02/2010)
Public Property Let Sequencial_remessa_Titulo(Valor As Long)
   mlngSequencial_remessa_Titulo = Valor
End Property

Public Property Get Sequencial_remessa_Titulo() As Long
   Sequencial_remessa_Titulo = mlngSequencial_remessa_Titulo
End Property

'pt.99257 - Moacir Pfau(30/06/2010)
Public Property Let valorMulta(Valor As Double)
   mdblValorMulta = Valor
End Property

Public Property Get valorMulta() As Double
   valorMulta = mdblValorMulta
End Property

Public Property Let ValorMora(Valor As Double)
   mdblValorMora = Valor
End Property

Public Property Get ValorMora() As Double
   ValorMora = mdblValorMora
End Property

'Projeto: #7373 - História: #6135 - Desenvolvimento: #7434 - Ivo Sousa(13/05/2013)
Public Property Let RemessaStatus(Valor As String)
   mstrRemessaStatus = Valor
End Property

Public Property Get RemessaStatus() As String
   RemessaStatus = mstrRemessaStatus
End Property

'Vinicius Elyseu(02/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
Public Property Let Remessa(Valor As Boolean)
   mblnRemessa = Valor
End Property

Public Property Get Remessa() As Boolean
   Remessa = mblnRemessa
End Property


'Projeto: #17081 - História: # - Desenvolvimento# - Moacir Pfau(20/01/2014)
Public Property Let Emailboleto_enviado(Valor As Boolean)
   mblnEmailboleto_enviado = Valor
End Property

Public Property Get Emailboleto_enviado() As Boolean
   Emailboleto_enviado = mblnEmailboleto_enviado
End Property

'Projeto: #17081 - História: # - Desenvolvimento# - Moacir Pfau(20/01/2014)
Public Property Let CaminhoPDF(Valor As String)
   mstrCaminhoPDF = Valor
End Property

Public Property Get CaminhoPDF() As String
   CaminhoPDF = mstrCaminhoPDF
End Property

Public Property Get Anexos() As clscolAnexosEmail
    Set Anexos = mobjAnexos
End Property

Public Property Let Anexos(ByVal Valor As clscolAnexosEmail)
    Set mobjAnexos = Valor
End Property

Public Function CriarObjeto(ByVal rd As IDBReader, Optional blnRemStatus As Boolean = False) As clsTituloCobrebem
    Dim objEmpresa              As New CEmpresas
    'Pt. 94752 - Moacir Pfau(21/10/2009)
    Dim objEnderecoCobranca     As New CEmpresasEnderecos

    With rd
        Call objEmpresa.CarregarRegistro(.GetString("Empresa"))
        If mstrOrigem = "Lançamentos" Then
            mstrNumeroDocumento = .GetString("Código")
        ElseIf mstrOrigem = "Duplicatas" Then
            mstrNumeroDocumento = .GetString("Nota")
        End If
        mstrParcela = .GetString("Parcela")
        mstrTipo = .GetString("Tipo")
        mstrNomesacado = objEmpresa.Razao
        If objEmpresa.Pessoa = "Jurídica" Then
            mstrCnpjsacado = objEmpresa.CNPJ_CPF
            mstrCpfsacado = ""
        ElseIf objEmpresa.Pessoa = "Física" Then
            mstrCpfsacado = objEmpresa.CNPJ_CPF
            mstrCnpjsacado = ""
        End If
                
        'Pt. 94752 - Moacir Pfau(21/10/2009)
        If objEnderecoCobranca.CarregarRegistro(.GetString("Empresa"), .GetLong("cd_cobranca")) Then
            If objEnderecoCobranca.NroEndereco <> Empty Then
                mstrEnderecosacado = objEnderecoCobranca.Endereco & ", " & objEnderecoCobranca.NroEndereco
            Else
                mstrEnderecosacado = objEnderecoCobranca.Endereco
            End If
            mstrBairroSacado = objEnderecoCobranca.Bairro
            mstrCidadesacado = objEnderecoCobranca.Cidade
            mstrEstadosacado = objEnderecoCobranca.Estado
            mstrCepsacado = Trim(Replace(objEnderecoCobranca.Cep, "-", ""))
        Else
            If objEmpresa.NroEndereco <> Empty Then
                mstrEnderecosacado = objEmpresa.Endereco & ", " & objEmpresa.NroEndereco
            Else
                mstrEnderecosacado = objEmpresa.Endereco
            End If
            mstrBairroSacado = objEmpresa.Bairro
            mstrCidadesacado = objEmpresa.Cidade
            mstrEstadosacado = objEmpresa.Estado
            mstrCepsacado = Trim(Replace(objEmpresa.Cep, "-", ""))
        End If
        mdteDatadocumento = .GetDate("Emissão")
        mdteDatavencimento = .GetDate("Vencimento")
        mdteDataprocessamento = Format(Now, "dd/mm/yyyy")
        mdblValorDocumento = (.GetDouble("Valor Original") + .GetDouble("Acréscimo")) - .GetDouble("Abatimento")
        mdblValorOriginal = .GetDouble("Valor Original")
        mdblPercentualjurosdiaatraso = .GetDouble("PerMrD")
        mdblPercentualmultaatraso = .GetDouble("PerMul")
        mdblPercentualdesconto = 0
        mdblValoroutrosacrescimos = .GetDouble("Acréscimo")
        mdblValorAbatimento = .GetDouble("Abatimento")
        mstrLinhaDigitavel = .GetString("LINDIG")
        mstrCodigoBarras = .GetString("CODBAR")
        mstrNossoNumero = .GetString("NOSNUM")
        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        mstrSequencianossonumero = .GetString("SeqNossoNumero")
        mstrEmpresa = .GetString("Empresa")
        mlngSequencial_remessa_Titulo = .GetLong("seq_remessa")
        'pt.99257 - Moacir Pfau(30/06/2010)
        mdblValorMulta = .GetDouble("VlrMul")
        mdblValorMora = .GetDouble("VlrMrD")
        'Projeto: #7373 - História: #6135 - Desenvolvimento: #7434 - Ivo Sousa(13/05/2013)
        If blnRemStatus Then
            mstrRemessaStatus = .GetString("RemessaStatus")
        End If
        mblnEmailboleto_enviado = .GetBoolean("Emailboleto_enviado")
        mblnRemessa = .GetBoolean("remessa")
    End With
    Set CriarObjeto = Me
End Function

Public Function Equals(obj As Object) As Boolean
    Dim booIgual As Boolean
    Dim newObj As clsTituloCobrebem
    
    Set newObj = obj
    booIgual = (newObj.NumeroDocumento = mstrNumeroDocumento)
    booIgual = booIgual And (newObj.Parcela = mstrParcela)
    Equals = booIgual
End Function

Public Function equalsEmpresa(obj As Object) As Boolean
    Dim booIgual As Boolean
    Dim newObj As clsTituloCobrebem
    
    Set newObj = obj
    booIgual = (newObj.Empresa = mstrEmpresa)
    equalsEmpresa = booIgual
End Function

Private Function fTiraCaracter(ByVal strCampo As String)
    strCampo = Replace(strCampo, ".", "")
    strCampo = Replace(strCampo, ",", "")
    strCampo = Replace(strCampo, "-", "")
    strCampo = Replace(strCampo, "/", "")
    fTiraCaracter = strCampo
End Function

Private Function fColocaMascara(ByVal strValor As String, ByVal strMascara As String) As String
    Dim Index As Integer
    Dim temp As String
    Dim i As Integer
    
    For i = 1 To Len(strMascara)
        If (Mid(strMascara, i, 1) = "#") Or (Mid(strMascara, i, 1) = ".") Or (Mid(strMascara, i, 1) = ",") Or (Mid(strMascara, i, 1) = "-") Or (Mid(strMascara, i, 1) = "/") Then
            temp = temp & Mid(strMascara, i, 1)
            Index = Index + 1
        Else
            If Index = 0 Then
                temp = temp & Mid(strValor, i, 1)
            Else
                temp = temp & Mid(strValor, i - Index, 1)
            End If
        End If
    Next
    fColocaMascara = temp
End Function

Private Function fColocaZero(ByVal strValor As String, ByVal strMascara As String) As String
    Dim lngQtdeZero             As Long
    Dim lngQtdeMascara          As Long
    Dim lngQtdeValor            As Long

On Error GoTo err
    lngQtdeMascara = Len(fTiraCaracter(strMascara))
    lngQtdeValor = Len(CStr(strValor))
    lngQtdeZero = lngQtdeMascara - lngQtdeValor

    fColocaZero = String(lngQtdeZero, "0") & strValor
    Exit Function
err:
End Function

Private Function carregarBanco(ByVal lngCodigoBanco As Long) As clsBanco
    Dim selCmd             As IDBSelectCommand
    Dim rdResult        As IDBReader
On Error GoTo err
    
    Set selCmd = mApp.CreateSelectCommand
    With selCmd
        .Table.TableName = "[Bancos]"
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", lngCodigoBanco, dbFieldTypeLong))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, selCmd)
    Set carregarBanco = New clsBanco
    If Not rdResult.EOF Then
        Call carregarBanco.CriarObjeto(rdResult, True)
    Else
        Set carregarBanco = Nothing
    End If
        
    Exit Function
err:
    Set selCmd = Nothing
    Set rdResult = Nothing
    Set carregarBanco = Nothing
End Function

Private Function fSeparaDiretorioArquivo(ByVal strDiretorioArquivo As String) As Boolean
    Dim i                   As Integer
    Dim intTemp             As Integer
    
    i = 1: mstrDiretorio = "":  mstrArquivo = ""
    For i = 0 To Len(strDiretorioArquivo)
        intTemp = Len(strDiretorioArquivo) - i
        If Mid(strDiretorioArquivo, intTemp, 1) = "\" Then
            mstrArquivo = Mid(strDiretorioArquivo, intTemp + 1, Len(strDiretorioArquivo) - intTemp)
            mstrDiretorio = Mid(strDiretorioArquivo, 1, intTemp)
            i = Len(strDiretorioArquivo)
        End If
    Next
End Function

'Emissao de boleto
Public Function EmissaoBoleto(ByVal colTitulo As clsColTituloCobrebem, ByVal lngCarteira As Long, ByVal lngBanco As Long)
    If InformacoesContaCorrente(lngCarteira, lngBanco, e_Boleto) Then
        If InformacoesBoleto(colTitulo, lngCarteira, e_Boleto) Then
            Call mobjCobreBem.ImprimeBoletos
        End If
    Else
        MsgBox "Problema ao carregar as informações da Conta Corrente/Carteira.", vbCritical, NomeModulo
    End If
End Function

'Emissao de boleto por email
Public Function EmissaoBoletoEmail(ByVal colTitulo As clsColTituloCobrebem, ByVal lngCarteira As Long, ByVal lngBanco As Long, ByVal CaminhoPDF As String)
    Dim objBioPDF                   As Object
    
    If InformacoesContaCorrente(lngCarteira, lngBanco, e_Boleto) Then
        If InformacoesBoleto(colTitulo, lngCarteira, e_Boleto, True) Then
            Call configurarImpressoraPDF(CaminhoPDF, objBioPDF)
            mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.NomeImpressora = PDF_PRINTERNAME
            mobjCobreBem.ImprimeBoletosSemPreview
            Call finalizaImpressoraPDF(objBioPDF)
        End If
    Else
        MsgBox "Problema ao carregar as informações da Conta Corrente/Carteira.", vbCritical, NomeModulo
    End If
End Function

'Geração de remessa
Public Function GeracaoRemessa(ByVal colTitulo As clsColTituloCobrebem, ByVal lngCarteira As Long, ByVal lngBanco As Long, ByVal strDiretorioArquivo As String) As Boolean
    Call fSeparaDiretorioArquivo(strDiretorioArquivo)
    If InformacoesContaCorrente(lngCarteira, lngBanco, e_Remessa) Then
        If InformacoesBoleto(colTitulo, lngCarteira, e_Remessa) Then
            Call mobjCobreBem.GravaArquivoRemessa
            GeracaoRemessa = True
        End If
    Else
        MsgBox "Problema ao carregar as informações da Conta Corrente/Carteira.", vbCritical, NomeModulo
    End If
End Function

'Carregar retorno
Public Function CarregarRetorno(ByVal lngCarteira As Long, ByVal lngBanco As Long, ByVal strDiretorioArquivo As String) As Boolean
    CarregarRetorno = True
    Call fSeparaDiretorioArquivo(strDiretorioArquivo)
    mstrMensagemValidacao = ""
    If InformacoesContaCorrente(lngCarteira, lngBanco, e_Retorno) Then
        mApp.BeginTransaction
        If CarregaRetorno(lngCarteira) Then
            mApp.CommitTransaction
        Else
            mApp.RollbackTransaction
        End If
    Else
        CarregarRetorno = False
        MsgBox "Problema ao carregar as informações da Conta Corrente/Carteira.", vbCritical, NomeModulo
    End If
End Function

Private Sub configurarImpressoraPDF(strCaminhoArquivo As String, ByRef objBioPDF As Object)
    Dim intIndexPrint As Integer
    Dim runonce As String
    
    Set objBioPDF = CreateObject(PRINTER_PROGID)
    objBioPDF.setValue "Output", strCaminhoArquivo
    objBioPDF.setValue "ConfirmOverwrite", "yes"
    objBioPDF.setValue "ShowSaveAS", "never"
    objBioPDF.setValue "ShowSettings", "never"
    objBioPDF.setValue "ShowPDF", "no"
    objBioPDF.setValue "RememberLastFileName", "no"
    objBioPDF.setValue "RememberLastFolderName", "no"
    objBioPDF.setValue "ShowProgressFinished", "never"
    objBioPDF.setValue "ShowProgress", "never"
    
    objBioPDF.WriteSettings True
        
    intIndexPrint = PrinterIndex(PDF_PRINTERNAME)
    If intIndexPrint >= 0 Then
        Set Printer = Printers(intIndexPrint)
    End If
End Sub

Private Sub finalizaImpressoraPDF(ByRef objBioPDF As Object)
     Dim runonce As String
        
    runonce = objBioPDF.GetSettingsFileName(True)
    While Dir(runonce, vbNormal) <> ""
        Sleep 100
    Wend
End Sub

Public Function PrinterIndex(ByVal printername As String) As Integer
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).DeviceName) Like LCase(printername) Then
            PrinterIndex = i
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function

Private Function InformacoesContaCorrente(ByVal lngCarteira As Long, ByVal lngBanco As Long, ByVal intTipo As ENUMTipoTitulo) As Boolean
    Dim objCarteira                     As clsCarteira
    Dim objCarteiraDao                  As clsCarteiraDAO
    Dim objBanco                        As clsBanco
    
On Error GoTo err

    Set objCarteira = New clsCarteira
    Set objCarteiraDao = New clsCarteiraDAO
    Call objCarteiraDao.init(mApp)
    
    Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, lngCarteira)
    Set mobjCobreBem = New CobreBemX.ContaCorrente
    Set objBanco = New clsBanco
    mlngSequencial_remessa = 0
    
    'Início dos parâmetros obrigatórios da conta corrente.
    mobjCobreBem.ArquivoLicenca = objCarteira.Arquivo_licenca
    Set objBanco = carregarBanco(lngBanco)
    If objBanco Is Nothing Then
        MsgBox "Erro ao carregar o banco", vbInformation, NomeModulo
        Exit Function
    End If
    ' 08/01/2018 - FBMI:62 - Yuji F. - Ajuste na geração do Banrisul, informações de agência e cedente possuem máscara diferente
    If Left$(mobjCobreBem.NumeroBanco, 3) = "041" Then
        mobjCobreBem.CodigoAgencia = fColocaMascara(fColocaZero(fTiraCaracter(objBanco.NumeroAgencia), _
            "0000.00"), "0000.00")
        mobjCobreBem.CodigoCedente = fColocaMascara(fColocaZero(fTiraCaracter(objCarteira.Codigo_cedente), _
            "0000000000.0.00"), "0000000000.0.00")
    Else
        If Trim(mobjCobreBem.MascaraCodigoAgencia) <> "" Then
            mobjCobreBem.CodigoAgencia = fColocaMascara(fTiraCaracter(objBanco.NumeroAgencia), mobjCobreBem.MascaraCodigoAgencia)
        End If
        mobjCobreBem.CodigoCedente = objCarteira.Codigo_cedente
    End If
    If Trim(mobjCobreBem.MascaraContaCorrente) <> "" Then
        mobjCobreBem.NumeroContaCorrente = fColocaMascara(fColocaZero(fTiraCaracter(objBanco.NumeroConta), mobjCobreBem.MascaraContaCorrente), mobjCobreBem.MascaraContaCorrente)
    End If
    mobjCobreBem.InicioNossoNumero = fColocaZero(objCarteira.Inicio_nosso_numero, mobjCobreBem.MascaraNossoNumero)
    mobjCobreBem.FimNossoNumero = fColocaZero(objCarteira.Fim_nosso_numero, mobjCobreBem.MascaraNossoNumero)
    mobjCobreBem.ProximoNossoNumero = fColocaZero(objCarteira.Proximo_nosso_numero, mobjCobreBem.MascaraNossoNumero)
    If Trim(objCarteira.Outro_dado_configuracao1) <> "" Then
        mobjCobreBem.OutroDadoConfiguracao1 = objCarteira.Outro_dado_configuracao1
    End If
    If Trim(objCarteira.Outro_dado_configuracao2) <> "" Then
        mobjCobreBem.OutroDadoConfiguracao2 = objCarteira.Outro_dado_configuracao2
    End If
    
    'mobjCobreBem.AcoesCobrancaRemessa.item(0).Codigo = 18
    mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.ArquivoLogotipo = objCarteira.Logo_empresa
    mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.CaminhoImagensCodigoBarras = App.Path & "\Cobrebem\Imagens\"
    mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.LayoutBoleto = objCarteira.Tipo_layout_boleto
    mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = objCarteira.Margem_superior_boleto
    mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.CaminhoImagensCodigoBarras = App.Path & "\Imagens\"
    mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = objCarteira.Margem_superior_boleto
    
    'Boleto
    'mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.LayoutBoleto = "PadraoReciboPersonalizado"
    'Pt. 114032 - Moacir Pfau(23/02/2012)
    If Trim(objCarteira.HTMLReciboPersonalizado) <> "" Then
        mobjCobreBem.PadroesBoleto.PadroesBoletoImpresso.HTMLReciboPersonalizado = objCarteira.HTMLReciboPersonalizado
    End If
    'Pt. 114032 - Moacir Pfau(23/02/2012)
    If Trim(objCarteira.IdentificacaoCedente) <> "" Then
        mobjCobreBem.PadroesBoleto.IdentificacaoCedente = objCarteira.IdentificacaoCedente
    End If

    If Not IsNothing(mobjCobreBem.ArquivoRemessa) Then
        mobjCobreBem.ArquivoRemessa.Layout = objCarteira.Tipo_layout_remessa
        'Pt. 105912 - Moacir Pfau(17/03/2011)
        If Trim(objCarteira.TipoImpressao) <> "" Then
            mobjCobreBem.ArquivoRemessa.TipoImpressao = objCarteira.TipoImpressao
        End If
        mobjCobreBem.ArquivoRetorno.Layout = objCarteira.Tipo_layout_retorno
        
        'Pt. 102459 - Moacir Pfau(29/10/2010)
        'Configurações Especiais.
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroCabecalho1 = objCarteira.NaoGerarRegistroCabecalho1
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroCabecalho2 = objCarteira.NaoGerarRegistroCabecalho2
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroDetalhe1 = objCarteira.NaoGerarRegistroDetalhe1
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroDetalhe2 = objCarteira.NaoGerarRegistroDetalhe2
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroDetalhe3 = objCarteira.NaoGerarRegistroDetalhe3
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroDetalhe4 = objCarteira.NaoGerarRegistroDetalhe4
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroDetalhe5 = objCarteira.NaoGerarRegistroDetalhe5
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroRodape1 = objCarteira.NaoGerarRegistroRodape1
        mobjCobreBem.ArquivoRemessa.NaoGerarRegistroRodape2 = objCarteira.NaoGerarRegistroRodape2
    End If
               
    If objCarteira.localPagamento <> "" Then
        mobjCobreBem.localPagamento = objCarteira.localPagamento
    End If
               
    If mstrArquivo <> "" And mstrDiretorio <> "" Then
        'Remessa
        mobjCobreBem.ArquivoRemessa.Arquivo = mstrArquivo
        mobjCobreBem.ArquivoRemessa.Diretorio = mstrDiretorio
        If intTipo = e_Remessa Then
            If objCarteira.Sequencial_remessa > 0 Then
                mlngSequencial_remessa = objCarteira.Sequencial_remessa + 1
                mobjCobreBem.ArquivoRemessa.Sequencia = mlngSequencial_remessa
            End If
        End If
        'Retorno
        mobjCobreBem.ArquivoRetorno.Arquivo = mstrArquivo
        mobjCobreBem.ArquivoRetorno.Diretorio = mstrDiretorio
    End If
    
    'Carrega variavel
    mstrDemonstrativo = objCarteira.Demonstrativo
    mstrInstrucoesCaixa = objCarteira.Instrucoes_caixa
    
    InformacoesContaCorrente = True
    mlngCarteira = lngCarteira
    
    Exit Function
err:
    InformacoesContaCorrente = False
End Function

Private Function InformacoesBoleto(ByVal colTitulo As clsColTituloCobrebem, ByVal lngCarteira As Long, ByVal intTipo As ENUMTipoTitulo, Optional ByVal Email As Boolean) As Boolean
    Dim objBoleto                               As Object
    Dim objTitulo                               As clsTituloCobrebem
    Dim strLinhaDigitavel                       As String
    Dim strCodigoBarras                         As String
    Dim strNossoNumero                          As String
    Dim strProximoNossoNumero                   As String
    Dim blnAtualiza                             As Boolean
    Dim objCarteira                             As clsCarteira
    Dim objCarteiraDao                          As clsCarteiraDAO
    Dim mlngNrRemessa                           As Long
    Dim objMDados                               As Object
    Dim objEmpresa As CEmpresas
    ' Demanda 187926
    ' Autor: Yuji
    ' Criação de variáveis para manipular a mensagem
    Dim strObservacao As String
    Dim strNumeroBanco As String
    Dim arrObservacao() As String
    
    Dim objPagFor As clsPagFor

On Error GoTo err

    Set objCarteira = New clsCarteira
    Set objCarteiraDao = New clsCarteiraDAO
    Set objPagFor = New clsPagFor
    Call objCarteiraDao.init(mApp)
    
    Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, lngCarteira)
    
    
    blnAtualiza = True
    mApp.BeginTransaction
    mobjCobreBem.DocumentosCobranca.Clear
    If colTitulo.Count > 0 Then
        colTitulo.MoveFirst
        While Not colTitulo.EOF
            Set objTitulo = colTitulo.CurrentObject
            
            Set objBoleto = mobjCobreBem.DocumentosCobranca.add
            ' Demanda 187926
            ' Autor: Yuji
            ' Obtendo o número do banco
            strNumeroBanco = Left$(mobjCobreBem.NumeroBanco, 3)
            
            'Pt. 96589 - Moacir Pfau(05/02/2010)
            If objCarteira.Banco_gera_nosso_numero Then
                'Verifica o último número gerado de remessa.
                If objTitulo.Sequencial_remessa_Titulo = 0 Then
                    mlngNrRemessa = fProximoNumeroRemessa
                    objTitulo.Sequencial_remessa_Titulo = mlngNrRemessa
                Else
                    mlngNrRemessa = objTitulo.Sequencial_remessa_Titulo
                End If
                'Vinicius Elyseu(06/10/2015) - Projeto: #0 - História: #0 - Desenv: #0
                If objCarteira.SeqRemessaNrDoc Then
                    objBoleto.NumeroDocumento = objTitulo.NumeroDocumento & "/" & objTitulo.Parcela & "/" & mlngNrRemessa
                Else
                    objBoleto.NumeroDocumento = objTitulo.NumeroDocumento & "/" & objTitulo.Parcela
                End If
            Else
                objBoleto.NumeroDocumento = objTitulo.NumeroDocumento & "/" & objTitulo.Parcela
            End If
           
            'pt.98929 - Fernando Paludo(02/08/2010)
            If Trim(objCarteira.Outra_especie) <> "" Then
                objBoleto.TipoDocumentoCobranca = objCarteira.Outra_especie
            Else
                objBoleto.TipoDocumentoCobranca = objCarteira.Especie
            End If
            
            'Pt. 96210 - Moacir Pfau(08/12/2009)
            objBoleto.Nomesacado = objTitulo.Nomesacado
            If Len(objTitulo.Cnpjsacado) > 11 Then
                objBoleto.Cnpjsacado = objTitulo.Cnpjsacado
            Else
                objBoleto.Cpfsacado = objTitulo.Cpfsacado
            End If
            
            objBoleto.Enderecosacado = objTitulo.Enderecosacado
            
            ' 22/07/2019 - FBMI:798 - Yuji F. - Ajuste na geração da remessa do Sicredi
            If strNumeroBanco = "748" And intTipo = e_Remessa And objCarteira.Tipo_layout_remessa = "CNAB400" Then
                objBoleto.BairroSacado = "0"
            Else
                'Pt. 106012 - Moacir Pfau(28/09/2011)
                objBoleto.BairroSacado = IIf(Trim(objCarteira.BairroSacado) <> "", _
                    objCarteira.BairroSacado, objTitulo.BairroSacado)
            End If
            
            objBoleto.Cidadesacado = objTitulo.Cidadesacado
            objBoleto.Estadosacado = objTitulo.Estadosacado
            objBoleto.Cepsacado = objTitulo.Cepsacado
            objBoleto.Datadocumento = objTitulo.Datadocumento
            objBoleto.DataVencimento = objTitulo.DataVencimento
            objBoleto.Dataprocessamento = objTitulo.Dataprocessamento
            objBoleto.ValorDocumento = objTitulo.ValorOriginal
            'pt.99257 - Moacir Pfau(30/06/2010)
            If objTitulo.Percentualjurosdiaatraso = 0 Then
                objTitulo.Percentualjurosdiaatraso = objCarteira.Per_mora
            End If
            If objTitulo.Percentualmultaatraso = 0 Then
                objTitulo.Percentualmultaatraso = objCarteira.Per_multa
            End If
            
            objBoleto.Percentualjurosdiaatraso = objTitulo.Percentualjurosdiaatraso
            objBoleto.Percentualmultaatraso = objTitulo.Percentualmultaatraso
            objBoleto.Percentualdesconto = objTitulo.Percentualdesconto
            objBoleto.Valoroutrosacrescimos = objTitulo.Valoroutrosacrescimos
            objBoleto.PadroesBoleto.Demonstrativo = mstrDemonstrativo
            
            'Demanda 161746 - Davi Brito(08/03/2017)
            If mobjCache Is Nothing Then
                Set mobjCache = New clsCache
            End If
            
            Set objEmpresa = mobjCache.GetCacheEmpresa(objTitulo.Empresa)
            
            ' Demanda 187926
            ' Autor: Yuji
            ' Modificando a lógica para concatenar as msg's 'Instruções Caixa' + 'Obs Remessa'
            strObservacao = ""
            If mstrInstrucoesCaixa <> "" Then
                strObservacao = mstrInstrucoesCaixa & "  "
            End If
            
            If objEmpresa.ObsRemessaBoleto <> "" Then
                strObservacao = strObservacao & Replace(objEmpresa.ObsRemessaBoleto, vbCrLf, " ")
                
                If (strNumeroBanco = "341") And (intTipo = e_Remessa) Then
                
                    ReDim arrObservacao(Int(Len(strObservacao) / 69))
                    Dim i As Integer, cont As Integer
                    cont = 0
                    For i = LBound(arrObservacao) To UBound(arrObservacao)
                        arrObservacao(i) = Left$(strObservacao, 69)
                        strObservacao = Replace(strObservacao, arrObservacao(i), NUL, , 1)
                        cont = cont + 1
                        If cont >= 3 Then Exit For
                    Next
                    
                    strObservacao = "<br>" & Join(arrObservacao, "<br>")
                
                End If
                                
            End If
            
            objBoleto.PadroesBoleto.InstrucoesCaixa = strObservacao
            
            Set mobjCache = Nothing
            
            objBoleto.ValorDesconto = objTitulo.ValorAbatimento
            'Pt. 97161 - Moacir Pfau(05/02/2010)
            objBoleto.DiasProtesto = objCarteira.Dias_protesto
            
            'Pt. 96589 - Moacir Pfau(05/02/2010)
            objBoleto.BancoGeraNossoNumero = objCarteira.Banco_gera_nosso_numero
            'pt.98445  - Fernando Paludo(26/04/2010)
            objBoleto.BancoEmiteBoleto = objCarteira.banco_Emite_boleto
            
            'Pt. 106012 - Moacir Pfau(28/09/2011)
            If Trim(objCarteira.CodigoPracaSacado) <> "" Then
                Set objMDados = objBoleto.MeusDados.add
                    objMDados.Nome = "CodigoPracaSacado"
                    objMDados.Valor = objCarteira.CodigoPracaSacado
            End If
                                                           
            If objCarteira.InstrucaoCobranca1 <> "" And objCarteira.ValorInstrucaoCobranca1 <> "" Then
                objBoleto.InstrucaoCobranca1 = objCarteira.InstrucaoCobranca1
                objBoleto.ValorInstrucaoCobranca1 = objCarteira.ValorInstrucaoCobranca1
            End If

            If objCarteira.InstrucaoCobranca2 <> "" And objCarteira.ValorInstrucaoCobranca2 <> "" Then
                objBoleto.InstrucaoCobranca2 = objCarteira.InstrucaoCobranca2
                objBoleto.ValorInstrucaoCobranca2 = objCarteira.ValorInstrucaoCobranca2
            End If

            If objCarteira.InstrucaoCobranca3 <> "" And objCarteira.ValorInstrucaoCobranca3 <> "" Then
                objBoleto.InstrucaoCobranca3 = objCarteira.InstrucaoCobranca3
                objBoleto.ValorInstrucaoCobranca3 = objCarteira.ValorInstrucaoCobranca3
            End If
            
            If objCarteira.UtilizaNumeroControle Then
                objBoleto.NumeroControle = objBoleto.NumeroDocumento
            End If
                                              
            'Projeto: #4350 - História: #4336 - Problema: #5436 - Ivo Sousa(01/03/2013)
            If objTitulo.Sequencianossonumero <> "0" And objTitulo.Sequencianossonumero <> Empty Then
                objBoleto.NossoNumero = objTitulo.Sequencianossonumero
            Else
                objBoleto.NossoNumero = objBoleto.NossoNumero
            End If


            If Not objCarteira.ColCpEspeciais Is Nothing Then
                If objCarteira.ColCpEspeciais.Count > 0 Then
                    objCarteira.ColCpEspeciais.MoveFirst
    
                    While Not objCarteira.ColCpEspeciais.EOF
                        Set objMDados = objBoleto.MeusDados.add
                        objMDados.Nome = objCarteira.ColCpEspeciais.CurrentObject.CpEspNome
                        objMDados.Valor = objCarteira.ColCpEspeciais.CurrentObject.CpEspValor
                        objCarteira.ColCpEspeciais.MoveNext
                    Wend
                End If
            End If
            
            'Calcula boleto.
            'objBoleto.CalculaDacNossoNumero = False
            Call mobjCobreBem.CalcularDadosBoletos
            strLinhaDigitavel = "": strCodigoBarras = "":   strNossoNumero = "": strProximoNossoNumero = ""
            strLinhaDigitavel = objBoleto.LinhaDigitavel
            strCodigoBarras = objBoleto.CodigoBarras
            If mobjCobreBem.NumeroBanco = "033-7" Then
                strNossoNumero = Right(objBoleto.NossoNumero, Len(objBoleto.NossoNumero) - 1) & objPagFor.DigitoVerificador(Right(objBoleto.NossoNumero, Len(objBoleto.NossoNumero) - 1), 11, False, False, 2, 3, 4, 5, 6, 7, 8, 9)
            Else
                strNossoNumero = objBoleto.NossoNumero
            End If

            'Projeto: #4350 - História: #4336 - Problema: #5436 - Ivo Sousa(01/03/2013)
            If objTitulo.Sequencianossonumero <> "0" And objTitulo.Sequencianossonumero <> "" Then
                strProximoNossoNumero = objTitulo.Sequencianossonumero
            Else
                strProximoNossoNumero = fColocaZero(mobjCobreBem.ProximoNossoNumero, mobjCobreBem.MascaraNossoNumero)
                If val(mobjCobreBem.ProximoNossoNumero) >= val(mobjCobreBem.FimNossoNumero) Then
                    MsgBox "Sequência do nosso número chegou ao final, favor verificar a carteira.", vbInformation, NomeModulo
                    colTitulo.MoveLast
                End If
            End If
            
            'Função para gravar no banco
            blnAtualiza = blnAtualiza And AtualizarTitulo(strLinhaDigitavel, strCodigoBarras, strNossoNumero, strProximoNossoNumero, objTitulo, lngCarteira, intTipo, Email)
      
            Set objTitulo = Nothing
            Set objBoleto = Nothing
            colTitulo.MoveNext
        Wend
    End If
    
    blnAtualiza = blnAtualiza And AtualizarCarteira(lngCarteira, fColocaZero(mobjCobreBem.ProximoNossoNumero, mobjCobreBem.MascaraNossoNumero))

    If blnAtualiza Then
        mApp.CommitTransaction
        InformacoesBoleto = True
    Else
        mApp.RollbackTransaction
        MsgBox "Erro ao gerar título", vbInformation, NomeModulo
        InformacoesBoleto = False
    End If
    Exit Function
err:
    mApp.RollbackTransaction
    MsgBox "Erro ao gerar título", vbInformation, NomeModulo
    InformacoesBoleto = False
End Function

Private Function CarregaRetorno(ByVal lngCarteira As Long) As Boolean
    Dim dblValorPago                As Double
    Dim dblValorAcrescimoComp       As Double
    Dim dblValorDescontoComp        As Double
    Dim dtaPagamento                As Date
    Dim dtaLiberacao                As Date
    Dim blnRetorno                  As Boolean
    Dim i                           As Integer
    Dim H                           As Integer
    Dim dblDifValor                 As Double
    Dim dblAcrescimoValor           As Double
    Dim dblDescontoValor            As Double
    Dim strNumeroDocumento          As String
    Dim dblValorCredito             As String
    Dim objCarteira                 As clsCarteira
    Dim objCarteiraDao              As clsCarteiraDAO
    Dim dblValorBaixa               As Double
    
    'Pt. 96589 - Moacir Pfau(08/02/2010)
    Set objCarteira = New clsCarteira
    Set objCarteiraDao = New clsCarteiraDAO
    Call objCarteiraDao.init(mApp)
    Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, lngCarteira)
    
    mobjCobreBem.OcorrenciasCobranca.Clear
    mobjCobreBem.CarregaArquivosRetorno
    blnRetorno = True
    mstrMensagemBaixados = ""
    mstrRelacaoNossoNumero = ""
    
    For i = 0 To mobjCobreBem.OcorrenciasCobranca.Count - 1
        dblValorPago = 0: dblValorAcrescimoComp = 0: dblValorDescontoComp = 0: dtaPagamento = Empty
        
        '*** Somente para verificação. ***
'        Debug.Print mobjCobreBem.OcorrenciasCobranca.item(i).Pagamento
'        Debug.Print mobjCobreBem.OcorrenciasCobranca(i).DataOcorrencia & " - " & mobjCobreBem.OcorrenciasCobranca(i).ValorPago & " - " & i
'
'        If mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia.Count > 0 Then
'            For h = 0 To mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia.Count - 1
'                Debug.Print "Código: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(h).Codigo & " Descrição: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(h).Descricao
'            Next
'        End If
        '*** Somente para verificação. ***
        
        'Relação de Nosso Numero dos Titulos baixados.
        If mstrRelacaoNossoNumero = "" Then
            mstrRelacaoNossoNumero = "'" & mobjCobreBem.OcorrenciasCobranca(i).NossoNumero & "'"
        Else
            mstrRelacaoNossoNumero = mstrRelacaoNossoNumero & ", " & "'" & mobjCobreBem.OcorrenciasCobranca(i).NossoNumero & "'"
        End If
        
        'pt.100834 - Moacir Pfau(18/08/2010)
        dtaPagamento = RetornaDataPagamento(i, objCarteira)
        dblValorPago = mobjCobreBem.OcorrenciasCobranca(i).ValorPago
        dblValorCredito = mobjCobreBem.OcorrenciasCobranca(i).ValorCredito
        dblValorAcrescimoComp = mobjCobreBem.OcorrenciasCobranca(i).ValorJurosPago + mobjCobreBem.OcorrenciasCobranca(i).ValorMultaPaga + mobjCobreBem.OcorrenciasCobranca(i).Valoroutrosacrescimos
        dblValorDescontoComp = mobjCobreBem.OcorrenciasCobranca(i).ValorDesconto
        #If DESENV Then
            Call ListaConteudoRetorno(i)
        #End If
        
        dblValorBaixa = IIf(dblValorPago > dblValorCredito, dblValorPago, dblValorCredito)
        
        'Pt. 96589 - Moacir Pfau(08/02/2010)
        strNumeroDocumento = ""
        If mobjCobreBem.OcorrenciasCobranca(i).NumeroDocumento = "" Then
            strNumeroDocumento = mobjCobreBem.OcorrenciasCobranca(i).NumeroControle
        Else
            If Len(mobjCobreBem.OcorrenciasCobranca(i).NumeroControle) > Len(mobjCobreBem.OcorrenciasCobranca(i).NumeroDocumento) Then
                strNumeroDocumento = mobjCobreBem.OcorrenciasCobranca(i).NumeroControle
            Else
                strNumeroDocumento = mobjCobreBem.OcorrenciasCobranca(i).NumeroDocumento
            End If
        End If
        
        If LocalizaDuplicata(Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero), objCarteira, strNumeroDocumento) Then
            mstrOrigem = "Duplicatas"
            If Trim(mdteDataPagamentoRealizado) <> "00:00:00" Then
                mstrMensagemValidacao = mstrMensagemValidacao & "O Título com o Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & " com o valor de " & Format(dblValorBaixa, "#,##0.00") & " já foi baixado." & vbCrLf
                'Projeto: 100340 - Desenv.: 142893  - Ueder Budni (10/10/2016)
                Call RegistraLogLancDup(CDbl(mstrNumeroDocumento), mstrNomesacado, mstrTipo, CLng(mstrParcela), mstrPagRec, Duplicata, "Retorno de Remessa Bancária: " & mobjCobreBem.ArquivoRetorno.Arquivo & " - O Título com o Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & " com o valor de " & Format(dblValorBaixa, "#,##0.00") & " já foi baixado.")
            Else
                'Pt. 94271 - Moacir Pfau(12/08/2009)
                 Call fTratamentoMotivo(i)
                 
                dblDifValor = 0:    dblAcrescimoValor = 0: dblDescontoValor = 0
                
                dblDifValor = dblValorBaixa - mdblValorDocumento
                If dblDifValor < 0 Then
                    dblDescontoValor = dblDifValor * -1
                End If
                If dblDifValor > 0 Then
                    dblAcrescimoValor = dblDifValor
                End If
                
                'Projeto: #0 - História: #18930 - Desenvolvimento#18970 - Moacir Pfau(11/12/2012)
                dtaLiberacao = SugestaoDataLiberacao(dtaPagamento, mlngBanco)
                blnRetorno = blnRetorno And AtualizarBaixaTitulo(Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero), lngCarteira, dblValorBaixa, dblAcrescimoValor, dblDescontoValor, dtaPagamento, e_Retorno, mstrMensagemCodigo, mstrMensagemMotivo, mobjCobreBem.OcorrenciasCobranca(i).Pagamento, dtaLiberacao, strNumeroDocumento)
                If blnRetorno And mobjCobreBem.OcorrenciasCobranca(i).Pagamento Then
                    mstrMensagemBaixados = mstrMensagemBaixados & "O Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & ", da Empresa: " & mstrNomesacado & ", com o Nr.Doc.: " & mstrNumeroDocumento & "/" & mstrParcela & ", com o valor: " & Format(dblValorBaixa, "#,##0.00") & "." & vbCrLf
                    'Projeto: 100340 - Desenv.: 142893  - Ueder Budni (10/10/2016)
                    Call RegistraLogLancDup(CDbl(mstrNumeroDocumento), mstrNomesacado, mstrTipo, CLng(mstrParcela), mstrPagRec, Duplicata, "Retorno de Remessa Bancária: " & mobjCobreBem.ArquivoRetorno.Arquivo & ": O Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & ", da Empresa: " & mstrNomesacado & ", com o Nr.Doc.: " & mstrNumeroDocumento & "/" & mstrParcela & ", com o valor: " & Format(dblValorBaixa, "#,##0.00") & ".")
                End If
            End If
        ElseIf LocalizaLancamento(Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero), objCarteira, strNumeroDocumento) Then
            mstrOrigem = "Lançamentos"
            If Trim(mdteDataPagamentoRealizado) <> "00:00:00" Then
                mstrMensagemValidacao = mstrMensagemValidacao & "O Título com o Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & " com o valor de " & Format(dblValorBaixa, "#,##0.00") & " já foi baixado." & vbCrLf
                'Projeto: 100340 - Desenv.: 142893  - Ueder Budni (10/10/2016)
                Call RegistraLogLancDup(CDbl(mstrNumeroDocumento), mstrNomesacado, mstrTipo, CLng(mstrParcela), mstrPagRec, Lancamento, "Retorno de Remessa Bancária: " & mobjCobreBem.ArquivoRetorno.Arquivo & ": O Título com o Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & " com o valor de " & Format(dblValorBaixa, "#,##0.00") & " já foi baixado.")
            Else
                'Pt. 94271 - Moacir Pfau(12/08/2009)
                 Call fTratamentoMotivo(i)
            
                dblDifValor = 0:    dblAcrescimoValor = 0: dblDescontoValor = 0
                
                dblDifValor = dblValorBaixa - mdblValorDocumento
                If dblDifValor < 0 Then
                    dblDescontoValor = dblDifValor * -1
                End If
                If dblDifValor > 0 Then
                    dblAcrescimoValor = dblDifValor
                End If
                'Projeto: #0 - História: #18930 - Desenvolvimento#18970 - Moacir Pfau(11/12/2012)
                dtaLiberacao = SugestaoDataLiberacao(dtaPagamento, mlngBanco)
                blnRetorno = blnRetorno And AtualizarBaixaTitulo(Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero), lngCarteira, dblValorBaixa, dblAcrescimoValor, dblDescontoValor, dtaPagamento, e_Retorno, mstrMensagemCodigo, mstrMensagemMotivo, mobjCobreBem.OcorrenciasCobranca(i).Pagamento, dtaLiberacao, strNumeroDocumento)
                If blnRetorno And mobjCobreBem.OcorrenciasCobranca(i).Pagamento Then
                    mstrMensagemBaixados = mstrMensagemBaixados & "O Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & ", da Empresa: " & mstrNomesacado & ", com o Nr.Doc.: " & mstrNumeroDocumento & "/" & mstrParcela & ", com o valor: " & Format(dblValorBaixa, "#,##0.00") & "." & vbCrLf
                    'Projeto: 100340 - Desenv.: 142893  - Ueder Budni (10/10/2016)
                    Call RegistraLogLancDup(CDbl(mstrNumeroDocumento), mstrNomesacado, mstrTipo, CLng(mstrParcela), mstrPagRec, Lancamento, "Retorno de Remessa Bancária: " & mobjCobreBem.ArquivoRetorno.Arquivo & ": O Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & ", da Empresa: " & mstrNomesacado & ", com o Nr.Doc.: " & mstrNumeroDocumento & "/" & mstrParcela & ", com o valor: " & Format(dblValorBaixa, "#,##0.00") & ".")
                End If
            End If
        Else
            mstrMensagemValidacao = mstrMensagemValidacao & "O Título com o Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & " com o valor de " & Format(dblValorBaixa, "#,##0.00") & " não foi localizado." & vbCrLf
        End If
    Next
    CarregaRetorno = True

    MsgBox "Cobranças carregadas --> " & mobjCobreBem.OcorrenciasCobranca.Count

End Function

Public Function LocalizaDuplicata(ByVal strNossoNumero As String, ByVal objCarteira As clsCarteira, ByVal strNumeroDocumento As String) As Boolean
    Dim cmd                                     As IDBSelectCommand
    Dim rdResult                                As IDBReader
    ReDim strArrDoc(3) As String

On Error GoTo erro_carregando
    mdteDataPagamentoRealizado = "00:00:00"
    
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = "[Duplicatas]"
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", objCarteira.Id_carteira, dbFieldTypeLong))
        
        'Pt. 96589 - Moacir Pfau(08/02/2010)
        'Sera necessario validação junto ao cliente.
        If objCarteira.Banco_gera_nosso_numero Then
            strArrDoc = Split(strNumeroDocumento, "/")
            Call .Filter.Append("[Nota] = @pNota")
            Call .Parameters.add(.CreateParameter("@pNota", strArrDoc(0), dbFieldTypeString))
            
            Call .Filter.Append("[Parcela] = @pParcela")
            Call .Parameters.add(.CreateParameter("@pParcela", strArrDoc(1), dbFieldTypeLong))
                        
            If UBound(strArrDoc) > 1 Then
                Call .Filter.Append("[seq_remessa] = @pSeq_remessa")
                Call .Parameters.add(.CreateParameter("@pSeq_remessa", strArrDoc(2), dbFieldTypeLong))
            End If
        Else
            ' 17/06/2019 - FBMI:728 - Yuji F. - Localizando duplicata sem o DV do Database
            If Left$(mobjCobreBem.NumeroBanco, 3) = "033" Then
                Call .Filter.Append("([NOSNUM] = @pNOSNUM Or (Ltrim(Str(Nota, 30)) + '/' + Cast(Parcela As Varchar(10)) = @pNumDoc " & _
                    "And (CASE WHEN ISNUMERIC(LEFT([NOSNUM],LEN(NOSNUM) - 1)) = 1 then cast(LEFT([NOSNUM],LEN(NOSNUM) - 1) as numeric(30,0)) else -1 end) = " & _
                    "(CASE WHEN ISNUMERIC(@pNOSNUM) = 1 then cast(@pNOSNUM as numeric(30,0)) else -2 end)))")
                Call .Parameters.add(.CreateParameter("@pNumDoc", strNumeroDocumento, dbFieldTypeString, 70))
            Else
                Call .Filter.Append("[NOSNUM] = @pNOSNUM")
            End If
            Call .Parameters.add(.CreateParameter("@pNOSNUM", strNossoNumero, dbFieldTypeString, 70))
        End If
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        LocalizaDuplicata = True
        mdteDataPagamentoRealizado = rdResult.GetDate("pagamento")
        mstrNumeroDocumento = rdResult.GetString("Nota")
        mstrParcela = rdResult.GetString("Parcela")
        mdteDatavencimento = rdResult.GetDate("Vencimento")
        mdblValorDocumento = (rdResult.GetDouble("Valor Original") + rdResult.GetDouble("Acréscimo")) - rdResult.GetDouble("Abatimento")
        mstrNomesacado = rdResult.GetString("Empresa")
        mlngBanco = rdResult.GetLong("banco")
        mstrTipo = rdResult.GetString("tipo")
        'Projeto: 100340 - Desenv.: 142893 - Ueder Budni (23/09/2016)
        mstrPagRec = rdResult.GetString("PagRec")
    Else
        LocalizaDuplicata = False
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
    
erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    LocalizaDuplicata = False
End Function

Public Function LocalizaLancamento(ByVal strNossoNumero As String, ByVal objCarteira As clsCarteira, ByVal strNumeroDocumento As String) As Boolean
    Dim cmd             As IDBSelectCommand
    Dim rdResult        As IDBReader
    'Pt. 96589 - Moacir Pfau(08/02/2010)
    ReDim strArrDoc(3) As String
    

On Error GoTo erro_carregando
    mdteDataPagamentoRealizado = "00:00:00"

    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = "[Lançamentos]"
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", objCarteira.Id_carteira, dbFieldTypeLong))
        
        'Pt. 96589 - Moacir Pfau(08/02/2010)
        If objCarteira.Banco_gera_nosso_numero Then
            strArrDoc = Split(strNumeroDocumento, "/")
            Call .Filter.Append("[Código] = @pCodigo")
            Call .Parameters.add(.CreateParameter("@pCodigo", strArrDoc(0), dbFieldTypeString))
            
            Call .Filter.Append("[Parcela] = @pParcela")
            Call .Parameters.add(.CreateParameter("@pParcela", strArrDoc(1), dbFieldTypeLong))
            
            If UBound(strArrDoc) > 1 Then
                Call .Filter.Append("[seq_remessa] = @pSeq_remessa")
                Call .Parameters.add(.CreateParameter("@pSeq_remessa", strArrDoc(2), dbFieldTypeLong))
            End If
        Else
            ' 19/07/2019 - FBMI:811 - Yuji F. - Correção na importação de retorno dos lançamentos
            If Left$(mobjCobreBem.NumeroBanco, 3) = "033" Then
                Call .Filter.Append("([NOSNUM] = @pNOSNUM Or (Ltrim(Str(Código, 30)) + '/' + Cast(Parcela As Varchar(10)) = @pNumDoc " & _
                    "And (CASE WHEN ISNUMERIC(LEFT([NOSNUM],LEN(NOSNUM) - 1)) = 1 then cast(LEFT([NOSNUM],LEN(NOSNUM) - 1) as numeric(30,0)) else -1 end) = " & _
                    "(CASE WHEN ISNUMERIC(@pNOSNUM) = 1 then cast(@pNOSNUM as numeric(30,0)) else -2 end)))")
            Else
                Call .Filter.Append("([NOSNUM] = @pNOSNUM Or (Ltrim(Str(Código, 30)) + '/' + Cast(Parcela As Varchar(10)) = @pNumDoc And [NOSNUM] = @pNosNum1))")
                Call .Parameters.add(.CreateParameter("@pNosNum1", "0" & Mid(strNossoNumero, 1, Len(strNossoNumero) - 1), dbFieldTypeString, 70))
            End If
            Call .Parameters.add(.CreateParameter("@pNOSNUM", strNossoNumero, dbFieldTypeString, 70))
            Call .Parameters.add(.CreateParameter("@pNumDoc", strNumeroDocumento, dbFieldTypeString, 70))
        End If
    End With
    
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        LocalizaLancamento = True
        mdteDataPagamentoRealizado = rdResult.GetDate("pagamento")
        mstrNumeroDocumento = rdResult.GetString("Código")
        mstrParcela = rdResult.GetString("Parcela")
        mdteDatavencimento = rdResult.GetDate("Vencimento")
        mdblValorDocumento = (rdResult.GetDouble("Valor Original") + rdResult.GetDouble("Acréscimo")) - rdResult.GetDouble("Abatimento")
        mstrNomesacado = rdResult.GetString("Empresa")
        mlngBanco = rdResult.GetLong("banco")
        mstrTipo = rdResult.GetString("Tipo")
        'Projeto: 100340 - Desenv.: 142893 - Ueder Budni (23/09/2016)
        mstrPagRec = rdResult.GetString("PagRec")
    Else
        LocalizaLancamento = False
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
    
erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    LocalizaLancamento = False
End Function


Public Function AtualizarTitulo(ByVal strLinhaDigitavel As String, ByVal strCodigoBarras As String, ByVal strNossoNumero As String, ByVal strProximoNossoNumero As String, ByVal obj As clsTituloCobrebem, ByVal lngCarteira, ByVal intTipo As ENUMTipoTitulo, Optional ByVal Email As Boolean) As Boolean
    Dim cmd                 As IDBUpdateCommand
    Dim blnAtualizado       As Boolean
    Dim strMsg              As String
    
On Error GoTo erro_atualizado
    Set cmd = mApp.CreateUpdateCommand
     With cmd
        .Table = mstrOrigem
        If mstrOrigem = "Duplicatas" Then
            Call .Filter.Append("Nota = @pNota")
        ElseIf mstrOrigem = "Lançamentos" Then
            Call .Filter.Append("Código = @pNota")
        End If
        Call .Parameters.add(.CreateParameter("@pNota", obj.NumeroDocumento, dbFieldTypeDouble))
                    
        Call .Filter.Append("Parcela = @pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", obj.Parcela, dbFieldTypeLong))

        Call .Filter.Append("PagRec = @pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", "R", dbFieldTypeString, 1))
        
        Call .Filter.Append("Empresa = @pEmpresa")
        Call .Parameters.add(.CreateParameter("@pEmpresa", obj.Empresa, dbFieldTypeString, 15))
        
        Call .Filter.Append("Tipo = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", obj.tipo, dbFieldTypeString, 20))
        
        Call .AddValue("LINDIG", "@pLINDIG")
        Call .Parameters.add(.CreateParameter("@pLINDIG", Trim(strLinhaDigitavel), dbFieldTypeString, 70))
    
        Call .AddValue("CODBAR", "@pCODBAR")
        Call .Parameters.add(.CreateParameter("@pCODBAR", strCodigoBarras, dbFieldTypeString, 44))

        Call .AddValue("NOSNUM", "@pNOSNUM")
        Call .Parameters.add(.CreateParameter("@pNOSNUM", strNossoNumero, dbFieldTypeString))
        
        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        Call .AddValue("SeqNossoNumero", "@pSeqNossoNumero")
        Call .Parameters.add(.CreateParameter("@pSeqNossoNumero", strProximoNossoNumero, dbFieldTypeString))
        
        Call .AddValue("Id_carteira", "@pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngCarteira, dbFieldTypeLong))
        
        Call .AddValue("ultima_mov_cnab", "@pUltima_mov_cnab")
        Call .Parameters.add(.CreateParameter("@pUltima_mov_cnab", intTipo, dbFieldTypeInt))
        
        Call .AddValue("ultima_data_cnab", "@pUltima_data_cnab")
        Call .Parameters.add(.CreateParameter("@pUltima_data_cnab", Format(Now, "DD/MM/YYYY"), dbFieldTypeDate))
        
        'Pt. 96589 - Moacir Pfau(08/02/2010)
        Call .AddValue("seq_remessa", "@pSeq_remessa")
        If obj.Sequencial_remessa_Titulo > 0 Then
            Call .Parameters.add(.CreateParameter("@pSeq_remessa", obj.Sequencial_remessa_Titulo, dbFieldTypeLong))
        Else
            Call .Parameters.add(.CreateParameter("@pSeq_remessa", 0, dbFieldTypeLong))
        End If
        
        'pt.99257 - Moacir Pfau(30/06/2010)
        Call .AddValue("PerMul", "@pPerMul")
        Call .Parameters.add(.CreateParameter("@pPerMul", Format(obj.Percentualmultaatraso, "#0.00"), dbFieldTypeDouble))
        
        Call .AddValue("PerMrD", "@pPerMrD")
        Call .Parameters.add(.CreateParameter("@pPerMrD", Format(obj.Percentualjurosdiaatraso, "#0.00"), dbFieldTypeDouble))
        
        If obj.Percentualmultaatraso > 0 And obj.valorMulta = 0 Then
            obj.valorMulta = CDbl(obj.ValorDocumento) * CDbl(obj.Percentualmultaatraso) / 100
            Call .AddValue("VlrMul", "@pVlrMul")
            Call .Parameters.add(.CreateParameter("@pVlrMul", Format(obj.valorMulta, "#0.00"), dbFieldTypeDouble))
        End If
                
        If obj.Percentualjurosdiaatraso > 0 And obj.ValorMora = 0 Then
            obj.ValorMora = CDbl(obj.ValorDocumento) * CDbl(obj.Percentualjurosdiaatraso) / 100
            Call .AddValue("VlrMrD", "@pVlrMrD")
            Call .Parameters.add(.CreateParameter("@pVlrMrD", Format(obj.ValorMora, "#0.00"), dbFieldTypeDouble))
        End If
        
        If Email Then
            Call .AddValue("emailboleto_enviado", "@pEmail")
            Call .Parameters.add(.CreateParameter("@pEmail", Email, dbFieldTypeBool))
        End If
        
        'Vinicius Elyseu(01/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
        If intTipo = e_Boleto Then
            Call .AddValue("boleto", "@pBoleto")
            Call .Parameters.add(.CreateParameter("@pBoleto", True, dbFieldTypeBool))
        ElseIf intTipo = e_Remessa Then
            Call .AddValue("remessa", "@pRemessa")
            Call .Parameters.add(.CreateParameter("@pRemessa", True, dbFieldTypeBool))
        End If
    End With
    
    blnAtualizado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
       
    'Projeto: 100340 - Desenv.: 142890 - Ueder Budni (23/09/2016)
    If blnAtualizado Then
        With obj
            strMsg = IIf(intTipo = e_Boleto, "Boleto gerado.", "Remessa gerada.")
            Call RegistraLogLancDup(.NumeroDocumento, .Empresa, .tipo, .Parcela, "R", IIf(mstrOrigem = "Duplicatas", enuLancDup.Duplicata, enuLancDup.Lancamento), strMsg)
        End With
    End If
   
    Set cmd = Nothing
    AtualizarTitulo = blnAtualizado
    Exit Function
erro_atualizado:
    Set cmd = Nothing
End Function

Public Function AtualizarCarteira(ByVal lngCarteira As Long, ByVal strProximoNossoNumero As String) As Boolean
    Dim cmd                 As IDBUpdateCommand
    Dim blnAtualizado       As Boolean
    Dim objCarteira         As New clsCarteira
    Dim objCarteiraDao      As New clsCarteiraDAO
    Dim blnUpdate           As Boolean
    
    
On Error GoTo erro_atualizado

    Call objCarteiraDao.init(Aplicacao)
    Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, lngCarteira)
    
    Set cmd = mApp.CreateUpdateCommand
     With cmd
        .Table = "[FFICarteira]"

        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngCarteira, dbFieldTypeLong))
                                                      
        If Not objCarteira.Banco_gera_nosso_numero Then
            Call .AddValue("proximo_nosso_numero", "@pProximo_nosso_numero")
            Call .Parameters.add(.CreateParameter("@pProximo_nosso_numero", strProximoNossoNumero, dbFieldTypeString, 30))
            blnUpdate = True
        End If
        
        'Pt. 112805 - Moacir Pfau(23/02/2012)
        If mlngSequencial_remessa > 0 Then
            Call .AddValue("Sequencial_remessa", "@pSequencial_remessa")
            Call .Parameters.add(.CreateParameter("@pSequencial_remessa", mlngSequencial_remessa, dbFieldTypeLong))
            blnUpdate = True
        End If
        
    End With
    
    If blnUpdate Then
        blnAtualizado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    Else
        blnAtualizado = True
    End If
   
    Set cmd = Nothing
    AtualizarCarteira = blnAtualizado
    Exit Function
erro_atualizado:
    Set cmd = Nothing
End Function

Public Function AtualizarBaixaTitulo(ByVal strNossoNumero As String, ByVal lngCarteira As String, ByVal dblValorPago As Double, ByVal dblValorAcrescimo As Double, ByVal dblValorDesconto As Double, ByVal dtaPagamento As Date, ByVal intTipo As ENUMTipoTitulo, ByVal strCod_Motivo, ByVal strDesc_Motivo, ByVal blnPagamento As Boolean, ByVal dtaLiberacao As Date, Optional ByVal strNumeroDocumento As String) As Boolean
    Dim cmd                 As IDBUpdateCommand
    Dim blnAtualizado       As Boolean
    Dim objCarteira         As New clsCarteira
    Dim objCarteiraDao      As New clsCarteiraDAO
    ReDim strArrDoc(3) As String
    Dim objMatrizCont       As cMatrizContabilizacao
    Dim objMatrizContDAO    As New cMatrizContabilizacaoDAO
    
On Error GoTo erro_atualizado
    Set cmd = mApp.CreateUpdateCommand
     With cmd
        .Table = mstrOrigem
        'Filtro.
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngCarteira, dbFieldTypeLong))
                    
        Call objCarteiraDao.init(Aplicacao)
        Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, lngCarteira)

        If objCarteira.Banco_gera_nosso_numero Then
            strArrDoc = Split(strNumeroDocumento, "/")

            'Pt. 101693 - Moacir Pfau(28/09/2010)
            If mstrOrigem = "Lançamentos" Then
                Call .Filter.Append("[Código] = @pCodigo")
                Call .Parameters.add(.CreateParameter("@pCodigo", strArrDoc(0), dbFieldTypeString))
            ElseIf mstrOrigem = "Duplicatas" Then
                Call .Filter.Append("[Nota] = @pNota")
                Call .Parameters.add(.CreateParameter("@pNota", strArrDoc(0), dbFieldTypeString))
            End If
            
            Call .Filter.Append("[Parcela] = @pParcela")
            Call .Parameters.add(.CreateParameter("@pParcela", strArrDoc(1), dbFieldTypeLong))
            
            
            If UBound(strArrDoc) > 1 Then
                Call .Parameters.add(.CreateParameter("@pSeq_remessa", strArrDoc(2), dbFieldTypeLong))
                Call .Filter.Append("[seq_remessa] = @pSeq_remessa")
            'Vinicius Elyseu(27/04/2016) - Retirado para correçao do problema do PT - 415440 (CARPA)
            'Else
            '    Call .Parameters.add(.CreateParameter("@pSeq_remessa", 0, dbFieldTypeLong))
            End If
            
            'Atualiza o Nosso Numero para situações em que o banco gere o nosso numero.
            Call .AddValue("NOSNUM", "@pNOSNUM")
            Call .Parameters.add(.CreateParameter("@pNOSNUM", strNossoNumero, dbFieldTypeString))
        Else
            If mstrOrigem = "Lançamentos" Then
                Call .Filter.Append("([NOSNUM] = @pNOSNUM Or (Ltrim(Str(Código, 30)) + '/' + Cast(Parcela As Varchar(10)) = @pNumDoc And Empresa = @pEmpresa And PagRec = 'R' And Tipo = @pTipo))")
            ElseIf mstrOrigem = "Duplicatas" Then
                Call .Filter.Append("([NOSNUM] = @pNOSNUM Or (Ltrim(Str(Nota, 30)) + '/' + Cast(Parcela As Varchar(10)) = @pNumDoc And Empresa = @pEmpresa And PagRec = 'R' And Tipo = @pTipo))")
            End If
            Call .Parameters.add(.CreateParameter("@pNOSNUM", strNossoNumero, dbFieldTypeString, 70))
            Call .Parameters.add(.CreateParameter("@pNumDoc", strNumeroDocumento, dbFieldTypeString, 70))
            Call .Parameters.add(.CreateParameter("@pEmpresa", mstrNomesacado, dbFieldTypeString, 70))
            Call .Parameters.add(.CreateParameter("@pTipo", mstrTipo, dbFieldTypeString, 70))
        End If
        
        'Campos para serem atualizados.
        Call .AddValue("Cod_motivo", "@pCod_motivo_rejeicao")
        Call .Parameters.add(.CreateParameter("@pCod_motivo_rejeicao", strCod_Motivo, dbFieldTypeString, 10))
        
        Call .AddValue("Desc_motivo", "@pDesc_motivo_rejeicao")
        Call .Parameters.add(.CreateParameter("@pDesc_motivo_rejeicao", strDesc_Motivo, dbFieldTypeString, 255))
        
        Call .AddValue("ultima_mov_cnab", "@pUltima_mov_cnab")
        Call .Parameters.add(.CreateParameter("@pUltima_mov_cnab", intTipo, dbFieldTypeInt))
        
        Call .AddValue("ultima_data_cnab", "@pUltima_data_cnab")
        Call .Parameters.add(.CreateParameter("@pUltima_data_cnab", Format(Now, "DD/MM/YYYY"), dbFieldTypeDate))
        
        'Pt. 94406 - Moacir Pfau 18/08/2009
        If blnPagamento Then
            Call .AddValue("Abatimento", "Abatimento + @pAbatimento")
            Call .Parameters.add(.CreateParameter("@pAbatimento", dblValorDesconto, dbFieldTypeDouble))
        
            Call .AddValue("Acréscimo", "Acréscimo + @pAcrescimo")
            Call .Parameters.add(.CreateParameter("@pAcrescimo", dblValorAcrescimo, dbFieldTypeDouble))
        
            Call .AddValue("Pagamento", "@pPagamento")
            Call .Parameters.add(.CreateParameter("@pPagamento", dtaPagamento, dbFieldTypeDate))
            
            'Projeto: #0 - História: #18930 - Desenvolvimento#18970 - Moacir Pfau(11/12/2012)
            Call .AddValue("Liberação", "@pLiberacao")
            Call .Parameters.add(.CreateParameter("@pLiberacao", dtaLiberacao, dbFieldTypeDate))
                        
            If ConfigSys.UtilizaIntegracaoContabil Then
                Set objMatrizCont = objMatrizContDAO.Carregar(mstrTipo)
                If Not objMatrizCont Is Nothing Then
                    Call .AddValue("cd_operacao_baixa", "@pOpercaoBaixa")
                    Call .Parameters.add(.CreateParameter("@pOpercaoBaixa", objMatrizCont.BaixaDuplicatasReceber, dbFieldTypeInt))
                End If
            End If
        End If
    End With
    
    blnAtualizado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
   
    Set cmd = Nothing
    AtualizarBaixaTitulo = blnAtualizado
    Exit Function
erro_atualizado:
    AtualizarBaixaTitulo = False
    Set cmd = Nothing
End Function

Private Function SugestaoDataLiberacao(ByVal datData As Date, ByVal Banco As Long) As Date
    Dim datLiberacao        As Date
    Dim biz                 As New BizBanco
    
    datLiberacao = datData
    datLiberacao = DateAdd("d", biz.DiasLiberacao(Banco), datLiberacao)
    If calendario.PermiteLancamento(datLiberacao, , False) <> "A" Then
        datLiberacao = datLiberacao + NumeroDiasUteisNaoUteis(datLiberacao, 0)
    End If
    SugestaoDataLiberacao = datLiberacao
End Function

'Tratamento por motivo, com este tratamento é possivel identificar todas as mensagens de motivo referente ao título.
'Pt. 94406 - Moacir Pfau(18/08/2009)
Private Function fTratamentoMotivo(ByVal i As Integer)
    Dim H                           As Integer
    Dim strMensagem                 As String
    Dim strMsgLog                   As String
    
    fTratamentoMotivo = True
    strMensagem = ""
    'Pt. 94406 - Moacir Pfau 18/08/2009
    mstrMensagemMotivo = ""
    mstrMensagemCodigo = ""
    
    If mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia.Count > 0 Then
        For H = 0 To mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia.Count - 1
            Debug.Print "Código: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Codigo & " Descrição: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Descricao
            strMensagem = strMensagem & "Código: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Codigo & " Descrição: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Descricao & " "
            
            'Projeto: 100340 - Desenv.: 142889 - Ueder Budni (22/09/2016)
            With mobjCobreBem
                strMsgLog = "Retorno de Remessa Bancária: " & .ArquivoRetorno.Arquivo & ": Código: " & .OcorrenciasCobranca(i).MotivosOcorrencia(H).Codigo & " - Descrição: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Descricao & "."
                Call RegistraLogLancDup(strToDbl(mstrNumeroDocumento), mstrNomesacado, mstrTipo, strToLng(mstrParcela), mstrPagRec, IIf(mstrOrigem = "Duplicatas", enuLancDup.Duplicata, enuLancDup.Lancamento), strMsgLog)
            End With
            'Pt. 94406 - Moacir Pfau 18/08/2009
            If mstrMensagemCodigo = "" Then
                mstrMensagemCodigo = mstrMensagemCodigo & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Codigo
                mstrMensagemMotivo = mstrMensagemMotivo & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Descricao
            Else
                mstrMensagemCodigo = mstrMensagemCodigo & ", " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Codigo & " "
                mstrMensagemMotivo = mstrMensagemMotivo & ", " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(H).Descricao & " "
            End If
        Next
    End If
    If strMensagem <> "" Then
        'Pt. 94406 - Moacir Pfau 18/08/2009
        'Propriedade que identifica o pagamento do título.
        'Verificado que mesmo com o título pago é descrito o motivo.
        If Not mobjCobreBem.OcorrenciasCobranca(i).Pagamento Then
            mstrMensagemValidacao = mstrMensagemValidacao & "O Título com o Nosso Número: " & Trim(mobjCobreBem.OcorrenciasCobranca(i).NossoNumero) & " com o valor de " & Format(mobjCobreBem.OcorrenciasCobranca(i).ValorPago, "#,##0.00") & " possui motivo, " & strMensagem & vbCrLf
        End If
    End If
End Function

Private Sub fLoadEnterprise_estabelecimento()
        mlngEnterpriseId = GetFieldValue("enterprise_id", "Usuários", "usuário = '" & UserName & "'")
        mlngCdEstabelecimento = GetFieldValue("cd_estabelecimento", "Usuários", "usuário = '" & UserName & "'")
End Sub

Private Sub Class_Initialize()
    Call fLoadEnterprise_estabelecimento
    Set mobjAnexos = New clscolAnexosEmail
End Sub

'Pt. 96589 - Moacir Pfau(08/02/2010)
Private Function fProximoNumeroRemessa() As Long
    Dim rdResult As IDBReader
    Dim cmd As IDBSelectCommand
   
On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = mstrOrigem
        .SelectClause = "MAX(seq_remessa) AS Sequencia"
        
        If mlngCarteira > 0 Then
            Call .Filter.Append("[id_carteira] = @pId_carteira")
            Call .Parameters.add(.CreateParameter("@pId_carteira", mlngCarteira, dbFieldTypeLong))
        End If
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        fProximoNumeroRemessa = rdResult.GetLong("Sequencia") + 1
        If fProximoNumeroRemessa = 0 Then
            fProximoNumeroRemessa = 1
        End If
    Else
        fProximoNumeroRemessa = 1
    End If
    
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
End Function

'---------------------------------------------------------------------------------------
'Procedure..: verificaProximoNN
'Data.......: 05/11/2010
'Autor......: Moacir Pfau
'Descrição..: Verifica se o proximo NN já foi gerado, para lançamento e duplicata.
'Retorno....: Se localizar um titulo com o proximo numero retorna false
'Protocolo..: 102597
'---------------------------------------------------------------------------------------
Public Function verificaProximoNN(ByVal lngCarteira As Long) As Boolean
    Dim blnAlteracao            As Boolean
    mstrMensagemValidacao = ""
    blnAlteracao = verificaProximoNN_Lancamento(lngCarteira)
    blnAlteracao = blnAlteracao And verificaProximoNN_Duplicata(lngCarteira)
    verificaProximoNN = blnAlteracao
End Function

Private Function verificaProximoNN_Lancamento(ByVal lngCarteira As Long) As Boolean
    Dim cmd                                     As IDBSelectCommand
    Dim rdResult                                As IDBReader
    Dim objCarteira                             As clsCarteira
    Dim objCarteiraDao                          As clsCarteiraDAO
    Dim strMensagem                             As String
    
On Error GoTo erro_carregando
    
    strMensagem = ""
    Set objCarteira = New clsCarteira
    Set objCarteiraDao = New clsCarteiraDAO
    Call objCarteiraDao.init(mApp)
    Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, lngCarteira)
        
    'Pt. 102814 - Moacir Pfau(18/11/2010)
    If CDbl(objCarteira.Proximo_nosso_numero) = 0 Then
        verificaProximoNN_Lancamento = True
        Exit Function
    End If

    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = "[Lançamentos]"
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngCarteira, dbFieldTypeLong))
        
        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        #If FOXSQL Then
            Call .Filter.Append("CONVERT(Float,[SeqNossoNumero]) >= @pSeqNossoNumero")
        #Else
            Call .Filter.Append("IIF(NOT SeqNossoNumero IS NULL, CDBL([SeqNossoNumero]),0) >= @pSeqNossoNumero")
        #End If
        Call .Parameters.add(.CreateParameter("@pSeqNossoNumero", CDbl(objCarteira.Proximo_nosso_numero) + 1, dbFieldTypeDouble))
        
        Call .Filter.Append("[PagRec] = @pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", "R", dbFieldTypeString))
    End With
    
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        rdResult.MoveFirst
        While Not rdResult.EOF
            If strMensagem = "" Then
                strMensagem = strMensagem & rdResult.GetLong("Código") & "-" & rdResult.GetLong("Parcela")
            Else
                strMensagem = strMensagem & ", " & rdResult.GetLong("Código") & "-" & rdResult.GetLong("Parcela")
            End If
            rdResult.MoveNext
        Wend
    End If
    
    If strMensagem <> "" Then
        mstrMensagemValidacao = mstrMensagemValidacao & "Lançamento(s) possui o sequencial já gerado: " & strMensagem & vbCrLf
        verificaProximoNN_Lancamento = False
    Else
        verificaProximoNN_Lancamento = True
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
    
erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    mstrMensagemValidacao = mstrMensagemValidacao & "Erro ao executar verificação do próximo sequêncial do Lançamento: " & err.Description
    verificaProximoNN_Lancamento = False
End Function

Private Function verificaProximoNN_Duplicata(ByVal lngCarteira As Long) As Boolean
    Dim cmd                                     As IDBSelectCommand
    Dim rdResult                                As IDBReader
    Dim objCarteira                             As clsCarteira
    Dim objCarteiraDao                          As clsCarteiraDAO
    Dim strMensagem                             As String
    
On Error GoTo erro_carregando
    
    strMensagem = ""
    Set objCarteira = New clsCarteira
    Set objCarteiraDao = New clsCarteiraDAO
    Call objCarteiraDao.init(mApp)
    Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, lngCarteira)
    
    'Pt. 102814 - Moacir Pfau(18/11/2010)
    If CDbl(objCarteira.Proximo_nosso_numero) = 0 Then
        verificaProximoNN_Duplicata = True
        Exit Function
    End If
    
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = "[Duplicatas]"
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngCarteira, dbFieldTypeLong))
        
        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        #If FOXSQL Then
            Call .Filter.Append("CONVERT(Float,[SeqNossoNumero]) >= @pSeqNossoNumero")
        #Else
            Call .Filter.Append("IIF(NOT SeqNossoNumero IS NULL, CDBL([SeqNossoNumero]),0) >= @pSeqNossoNumero")
        #End If
        Call .Parameters.add(.CreateParameter("@pSeqNossoNumero", CDbl(objCarteira.Proximo_nosso_numero) + 1, dbFieldTypeDouble))
                
        Call .Filter.Append("[PagRec] = @pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", "R", dbFieldTypeString))
    End With
    
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        rdResult.MoveFirst
        While Not rdResult.EOF
            If strMensagem = "" Then
                strMensagem = strMensagem & rdResult.GetLong("nota") & "-" & rdResult.GetLong("Parcela")
            Else
                strMensagem = strMensagem & ", " & rdResult.GetLong("nota") & "-" & rdResult.GetLong("Parcela")
            End If
            rdResult.MoveNext
        Wend
    End If
    
    If strMensagem <> "" Then
        mstrMensagemValidacao = mstrMensagemValidacao & "Duplicata(s) possui o sequencial já gerado: " & strMensagem & vbCrLf
        verificaProximoNN_Duplicata = False
    Else
        verificaProximoNN_Duplicata = True
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
    
erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    mstrMensagemValidacao = mstrMensagemValidacao & "Erro ao executar verificação do próximo sequêncial da Duplicata: " & err.Description
    verificaProximoNN_Duplicata = False
End Function

Private Sub ListaConteudoRetorno(ByVal i As Integer)
    Dim Agencia, Banco, Carteira, CodigoCedente, CodigoOcorrencia, ContaCorrente, Nome, ValorNumerico, valorTexto
    Dim DataCredito, DataOcorrencia, Codigo, Descricao, NossoNumero, NumeroControle, NumeroDocumento, Pagamento
    Dim ValorAbatimento, ValorCredito, ValorDesconto, ValorIOF, ValorJurosPago, ValorMultaPaga, ValorOutrasDespesas
    Dim Valoroutrosacrescimos, ValorPago, ValorTaxaCobranca, ValorDocumento
    Dim ii As Integer
    Dim strMensagem As String
    Dim intFreeFile As Integer
   
    intFreeFile = FreeFile
    
    strMensagem = ""
    strMensagem = strMensagem & "Agencia: " & mobjCobreBem.OcorrenciasCobranca(i).Agencia & vbCrLf
    strMensagem = strMensagem & "Banco: " & mobjCobreBem.OcorrenciasCobranca(i).Banco & vbCrLf
    strMensagem = strMensagem & "Carteira: " & mobjCobreBem.OcorrenciasCobranca(i).Carteira & vbCrLf
    strMensagem = strMensagem & "CodigoCedente: " & mobjCobreBem.OcorrenciasCobranca(i).CodigoCedente & vbCrLf
    strMensagem = strMensagem & "CodigoOcorrencia: " & mobjCobreBem.OcorrenciasCobranca(i).CodigoOcorrencia & vbCrLf
    strMensagem = strMensagem & "ContaCorrente: " & mobjCobreBem.OcorrenciasCobranca(i).ContaCorrente & vbCrLf

    If mobjCobreBem.OcorrenciasCobranca(i).DadosOcorrencia.Count > 0 Then
        For ii = 0 To mobjCobreBem.OcorrenciasCobranca(i).DadosOcorrencia.Count - 1
            strMensagem = strMensagem & "Nome: " & mobjCobreBem.OcorrenciasCobranca(i).DadosOcorrencia(ii).Nome & vbCrLf
            strMensagem = strMensagem & "ValorNumerico: " & ValorNumerico = mobjCobreBem.OcorrenciasCobranca(i).DadosOcorrencia(ii).ValorNumerico & vbCrLf
            strMensagem = strMensagem & "valorTexto: " & mobjCobreBem.OcorrenciasCobranca(i).DadosOcorrencia.item(ii).valorTexto & vbCrLf
        Next
    End If
    
    strMensagem = strMensagem & "DataCredito: " & mobjCobreBem.OcorrenciasCobranca(i).DataCredito & vbCrLf
    strMensagem = strMensagem & "DataOcorrencia: " & mobjCobreBem.OcorrenciasCobranca(i).DataOcorrencia & vbCrLf
    strMensagem = strMensagem & "Codigo: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(0).Codigo & vbCrLf
    strMensagem = strMensagem & "Descricao: " & mobjCobreBem.OcorrenciasCobranca(i).MotivosOcorrencia(0).Descricao & vbCrLf
    strMensagem = strMensagem & "NossoNumero: " & mobjCobreBem.OcorrenciasCobranca(i).NossoNumero & vbCrLf
    strMensagem = strMensagem & "NumeroControle: " & mobjCobreBem.OcorrenciasCobranca(i).NumeroControle & vbCrLf
    strMensagem = strMensagem & "NumeroDocumento: " & mobjCobreBem.OcorrenciasCobranca(i).NumeroDocumento & vbCrLf
    strMensagem = strMensagem & "Pagamento: " & mobjCobreBem.OcorrenciasCobranca(i).Pagamento & vbCrLf
    strMensagem = strMensagem & "ValorAbatimento: " & mobjCobreBem.OcorrenciasCobranca(i).ValorAbatimento & vbCrLf
    strMensagem = strMensagem & "ValorCredito: " & mobjCobreBem.OcorrenciasCobranca(i).ValorCredito & vbCrLf
    strMensagem = strMensagem & "ValorDesconto: " & mobjCobreBem.OcorrenciasCobranca(i).ValorDesconto & vbCrLf
    strMensagem = strMensagem & "ValorIOF: " & mobjCobreBem.OcorrenciasCobranca(i).ValorIOF & vbCrLf
    strMensagem = strMensagem & "ValorJurosPago: " & mobjCobreBem.OcorrenciasCobranca(i).ValorJurosPago & vbCrLf
    strMensagem = strMensagem & "ValorMultaPaga: " & mobjCobreBem.OcorrenciasCobranca(i).ValorMultaPaga & vbCrLf
    strMensagem = strMensagem & "ValorOutrasDespesas: " & mobjCobreBem.OcorrenciasCobranca(i).ValorOutrasDespesas & vbCrLf
    strMensagem = strMensagem & "Valoroutrosacrescimos: " & mobjCobreBem.OcorrenciasCobranca(i).Valoroutrosacrescimos & vbCrLf
    strMensagem = strMensagem & "ValorPago: " & mobjCobreBem.OcorrenciasCobranca(i).ValorPago & vbCrLf
    strMensagem = strMensagem & "ValorTaxaCobranca: " & mobjCobreBem.OcorrenciasCobranca(i).ValorTaxaCobranca & vbCrLf
    strMensagem = strMensagem & "Item: " & i & vbCrLf & String(30, "_") & vbCrLf & vbCrLf

    If mobjCobreBem.OcorrenciasCobranca(i).Pagamento Then
        Open App.Path & "\Pagamento.txt" For Append As #intFreeFile
        Print #intFreeFile, strMensagem
        Close #intFreeFile
    Else
        Open App.Path & "\Confirmacao.txt" For Append As #intFreeFile
        Print #intFreeFile, strMensagem
        Close #intFreeFile
    End If

End Sub

Private Function RetornaDataPagamento(ByVal i As Integer, ByVal objCarteira As clsCarteira) As Date
    Dim dtaPagamento                    As Date
    
    If mobjCobreBem.OcorrenciasCobranca.item(i).Pagamento Then
        If objCarteira.Data_baixa_retorno = eData_baixa_retorno.DataCredito Then
            dtaPagamento = mobjCobreBem.OcorrenciasCobranca(i).DataCredito
        Else
            dtaPagamento = mobjCobreBem.OcorrenciasCobranca(i).DataOcorrencia
        End If
    Else
        dtaPagamento = mobjCobreBem.OcorrenciasCobranca(i).DataOcorrencia
    End If
    RetornaDataPagamento = dtaPagamento
End Function

Private Sub Class_Terminate()
    Set mobjAnexos = Nothing
End Sub

'Projeto: 100340 - Desenv.: 142890 - Ueder Budni (23/09/2016)
Private Sub RegistraLogLancDup(dblNumero As Double, strEmpresa As String, strTipo As String, lngParcela As Long, strPagRec As String, enuTabela As enuLancDup, strMsg As String)
    Dim objLogLancDup   As New clsLogLancamentosDuplicatas

On Error GoTo erro
    Call objLogLancDup.SetKey(strPagRec, dblNumero, strEmpresa, strTipo, lngParcela, enuTabela)
    Call objLogLancDup.InsertMsg(strMsg)
erro:
    Set objLogLancDup = Nothing
End Sub
