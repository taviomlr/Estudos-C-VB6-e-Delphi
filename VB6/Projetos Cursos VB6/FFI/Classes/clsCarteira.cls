VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarteira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngEnterprise_id                       As Long
Private mlngCd_estabelecimento                  As Long
Private mlngId_carteira                         As Long
Private mstrDesc_carteira                       As String
Private mstrCodigo_cedente                      As String
Private mstrInicio_nosso_numero                 As String
Private mstrFim_nosso_numero                    As String
Private mstrProximo_nosso_numero                As String
Private mstrDemonstrativo                       As String
Private mstrInstrucoes_caixa                    As String
Private mstrTipo_layout_boleto                  As String
Private mstrTipo_layout_remessa                 As String
Private mstrTipo_layout_retorno                 As String
Private mstrArquivo_licenca                     As String
Private mstrLogo_empresa                        As String
Private mstrCaminho_arquivo_remessa_padrao      As String
Private mstrCaminho_arquivo_retorno_padrao      As String
Private mlngMargem_superior_boleto              As Long
Private mstrOutro_dado_configuracao1            As String
Private mstrOutro_dado_configuracao2            As String
Private mlngSequencial_remessa                  As Long
'Pt. 96180 - Moacir Pfau(08/12/2009)
Private mstrEspecie                             As String
'Pt. 96589 - Moacir Pfau(05/02/2010)
Private mblnBanco_gera_nosso_numero             As Boolean
'Pt. 97161 - Moacir Pfau(05/02/2010)
Private mintDias_protesto                       As Integer
'pt.98446 - Fernando Paludo(26/04/2010)
Private mblnBancoEmiteBoleto                    As Boolean
'pt.99257 - Moacir Pfau(30/06/2010)
Private mdblPer_multa                           As Double
Private mdblPer_mora                            As Double
'pt.98929 - Fernando Paludo(02/08/2010)
Private mstrOutra_especie                       As String
'Pt. 102459 - Moacir Pfau(29/10/2010)
Private mblnNaoGerarRegistroCabecalho1          As Boolean
Private mblnNaoGerarRegistroCabecalho2          As Boolean
Private mblnNaoGerarRegistroDetalhe1            As Boolean
Private mblnNaoGerarRegistroDetalhe2            As Boolean
Private mblnNaoGerarRegistroDetalhe3            As Boolean
Private mblnNaoGerarRegistroDetalhe4            As Boolean
Private mblnNaoGerarRegistroDetalhe5            As Boolean
Private mblnNaoGerarRegistroRodape1             As Boolean
Private mblnNaoGerarRegistroRodape2             As Boolean
'Pt. 105912  - Moacir Pfau(17/03/2011)
Private mstrTipoImpressao                       As String
'Pt. 106012 - Moacir Pfau(28/09/2011)
Private mstrBairroSacado                        As String
Private mstrCodigoPracaSacado                   As String
 'Pt. 114032 - Moacir Pfau(23/02/2012)
Private mstrIdentificacaoCedente                As String
Private mstrHTMLReciboPersonalizado             As String
'Projeto: #4350 - História: # - Desenvolvimento# - Moacir Pfau(09/04/2013)
Private mintData_baixa_retorno                  As Integer
'Projeto: #17081 - História: #0 - Desenvolvimento#0 - Moacir Pfau(29/11/2013)
Private mstrInstrucaoCobranca1                  As String
Private mstrInstrucaoCobranca2                  As String
Private mstrInstrucaoCobranca3                  As String
Private mstrValorInstrucaoCobranca1             As String
Private mstrValorInstrucaoCobranca2             As String
Private mstrValorInstrucaoCobranca3             As String
Private mstrLocalPagamento                      As String
Private mblnUtilizaNumeroControle               As Boolean
Private mcolCpEspeciais                         As clscolCarteiraCpEspecial
'Vinicius Elyseu(06/10/2015) - Projeto: #0 - História: #0 - Desenv: #0
Private mblnSeqRemessaNrDoc                     As Boolean

Public Property Let Enterprise_id(Valor As Long)
   mlngEnterprise_id = Valor
End Property

Public Property Get Enterprise_id() As Long
   Enterprise_id = mlngEnterprise_id
End Property

Public Property Let Cd_estabelecimento(Valor As Long)
   mlngCd_estabelecimento = Valor
End Property

Public Property Get Cd_estabelecimento() As Long
   Cd_estabelecimento = mlngCd_estabelecimento
End Property

Public Property Let Id_carteira(Valor As Long)
   mlngId_carteira = Valor
End Property

Public Property Get Id_carteira() As Long
   Id_carteira = mlngId_carteira
End Property

Public Property Let Desc_carteira(Valor As String)
   mstrDesc_carteira = Valor
End Property

Public Property Get Desc_carteira() As String
   Desc_carteira = mstrDesc_carteira
End Property

Public Property Let Codigo_cedente(Valor As String)
   mstrCodigo_cedente = Valor
End Property

Public Property Get Codigo_cedente() As String
   Codigo_cedente = mstrCodigo_cedente
End Property

Public Property Let Inicio_nosso_numero(Valor As String)
   mstrInicio_nosso_numero = Valor
End Property

Public Property Get Inicio_nosso_numero() As String
   Inicio_nosso_numero = mstrInicio_nosso_numero
End Property

Public Property Let Fim_nosso_numero(Valor As String)
   mstrFim_nosso_numero = Valor
End Property

Public Property Get Fim_nosso_numero() As String
   Fim_nosso_numero = mstrFim_nosso_numero
End Property

Public Property Let Proximo_nosso_numero(Valor As String)
   mstrProximo_nosso_numero = Valor
End Property

Public Property Get Proximo_nosso_numero() As String
   Proximo_nosso_numero = mstrProximo_nosso_numero
End Property

Public Property Let Demonstrativo(Valor As String)
   mstrDemonstrativo = Valor
End Property

Public Property Get Demonstrativo() As String
   Demonstrativo = mstrDemonstrativo
End Property

Public Property Let Instrucoes_caixa(Valor As String)
   mstrInstrucoes_caixa = Valor
End Property

Public Property Get Instrucoes_caixa() As String
   Instrucoes_caixa = mstrInstrucoes_caixa
End Property

Public Property Let Tipo_layout_boleto(Valor As String)
   mstrTipo_layout_boleto = Valor
End Property

Public Property Get Tipo_layout_boleto() As String
   Tipo_layout_boleto = mstrTipo_layout_boleto
End Property

Public Property Let Tipo_layout_remessa(Valor As String)
   mstrTipo_layout_remessa = Valor
End Property

Public Property Get Tipo_layout_remessa() As String
   Tipo_layout_remessa = mstrTipo_layout_remessa
End Property

Public Property Let Tipo_layout_retorno(Valor As String)
   mstrTipo_layout_retorno = Valor
End Property

Public Property Get Tipo_layout_retorno() As String
   Tipo_layout_retorno = mstrTipo_layout_retorno
End Property

Public Property Let Arquivo_licenca(Valor As String)
   mstrArquivo_licenca = Valor
End Property

Public Property Get Arquivo_licenca() As String
   Arquivo_licenca = mstrArquivo_licenca
End Property

Public Property Let Logo_empresa(Valor As String)
   mstrLogo_empresa = Valor
End Property

Public Property Get Logo_empresa() As String
   Logo_empresa = mstrLogo_empresa
End Property

Public Property Let Caminho_arquivo_remessa_padrao(Valor As String)
   mstrCaminho_arquivo_remessa_padrao = Valor
End Property

Public Property Get Caminho_arquivo_remessa_padrao() As String
   Caminho_arquivo_remessa_padrao = mstrCaminho_arquivo_remessa_padrao
End Property

Public Property Let Caminho_arquivo_retorno_padrao(Valor As String)
   mstrCaminho_arquivo_retorno_padrao = Valor
End Property

Public Property Get Caminho_arquivo_retorno_padrao() As String
   Caminho_arquivo_retorno_padrao = mstrCaminho_arquivo_retorno_padrao
End Property

Public Property Let Margem_superior_boleto(Valor As Long)
   mlngMargem_superior_boleto = Valor
End Property

Public Property Get Margem_superior_boleto() As Long
   Margem_superior_boleto = mlngMargem_superior_boleto
End Property

Public Property Let Outro_dado_configuracao1(Valor As String)
   mstrOutro_dado_configuracao1 = Valor
End Property

Public Property Get Outro_dado_configuracao1() As String
   Outro_dado_configuracao1 = mstrOutro_dado_configuracao1
End Property

Public Property Let Outro_dado_configuracao2(Valor As String)
   mstrOutro_dado_configuracao2 = Valor
End Property

Public Property Get Outro_dado_configuracao2() As String
   Outro_dado_configuracao2 = mstrOutro_dado_configuracao2
End Property

Public Property Let Sequencial_remessa(Valor As Long)
   mlngSequencial_remessa = Valor
End Property

Public Property Get Sequencial_remessa() As Long
   Sequencial_remessa = mlngSequencial_remessa
End Property

'Pt. 96180 - Moacir Pfau(08/12/2009)
Public Property Let Especie(Valor As String)
   mstrEspecie = Valor
End Property

Public Property Get Especie() As String
   Especie = mstrEspecie
End Property

'Pt. 96589 - Moacir Pfau(05/02/2010)
Public Property Let Banco_gera_nosso_numero(Valor As Boolean)
   mblnBanco_gera_nosso_numero = Valor
End Property

Public Property Get Banco_gera_nosso_numero() As Boolean
   Banco_gera_nosso_numero = mblnBanco_gera_nosso_numero
End Property

'Pt. 97161 - Moacir Pfau(05/02/2010)
Public Property Let Dias_protesto(Valor As Integer)
   mintDias_protesto = Valor
End Property

Public Property Get Dias_protesto() As Integer
   Dias_protesto = mintDias_protesto
End Property

'pt. 98446 - Fernando Paludo(26/04/2010)
Public Property Let banco_Emite_boleto(Valor As Boolean)
   mblnBancoEmiteBoleto = Valor
End Property

Public Property Get banco_Emite_boleto() As Boolean
   banco_Emite_boleto = mblnBancoEmiteBoleto
End Property

'pt.99257 - Moacir Pfau(30/06/2010)
Public Property Let Per_multa(Valor As Double)
   mdblPer_multa = Valor
End Property

Public Property Get Per_multa() As Double
   Per_multa = mdblPer_multa
End Property

Public Property Let Per_mora(Valor As Double)
   mdblPer_mora = Valor
End Property

Public Property Get Per_mora() As Double
   Per_mora = mdblPer_mora
End Property

'pt.98292 - Fernando Paludo(02/08/2010)
Public Property Let Outra_especie(Valor As String)
   mstrOutra_especie = Valor
End Property

Public Property Get Outra_especie() As String
   Outra_especie = mstrOutra_especie
End Property

'Pt. 102459 - Moacir Pfau(29/10/2010)
Public Property Let NaoGerarRegistroCabecalho1(Valor As Boolean)
   mblnNaoGerarRegistroCabecalho1 = Valor
End Property

Public Property Get NaoGerarRegistroCabecalho1() As Boolean
   NaoGerarRegistroCabecalho1 = mblnNaoGerarRegistroCabecalho1
End Property

Public Property Let NaoGerarRegistroCabecalho2(Valor As Boolean)
   mblnNaoGerarRegistroCabecalho2 = Valor
End Property

Public Property Get NaoGerarRegistroCabecalho2() As Boolean
   NaoGerarRegistroCabecalho2 = mblnNaoGerarRegistroCabecalho2
End Property

Public Property Let NaoGerarRegistroDetalhe1(Valor As Boolean)
   mblnNaoGerarRegistroDetalhe1 = Valor
End Property

Public Property Get NaoGerarRegistroDetalhe1() As Boolean
   NaoGerarRegistroDetalhe1 = mblnNaoGerarRegistroDetalhe1
End Property

Public Property Let NaoGerarRegistroDetalhe2(Valor As Boolean)
   mblnNaoGerarRegistroDetalhe2 = Valor
End Property

Public Property Get NaoGerarRegistroDetalhe2() As Boolean
   NaoGerarRegistroDetalhe2 = mblnNaoGerarRegistroDetalhe2
End Property

Public Property Let NaoGerarRegistroDetalhe3(Valor As Boolean)
   mblnNaoGerarRegistroDetalhe3 = Valor
End Property

Public Property Get NaoGerarRegistroDetalhe3() As Boolean
   NaoGerarRegistroDetalhe3 = mblnNaoGerarRegistroDetalhe3
End Property

Public Property Let NaoGerarRegistroDetalhe4(Valor As Boolean)
   mblnNaoGerarRegistroDetalhe4 = Valor
End Property

Public Property Get NaoGerarRegistroDetalhe4() As Boolean
   NaoGerarRegistroDetalhe4 = mblnNaoGerarRegistroDetalhe4
End Property

Public Property Let NaoGerarRegistroDetalhe5(Valor As Boolean)
   mblnNaoGerarRegistroDetalhe5 = Valor
End Property

Public Property Get NaoGerarRegistroDetalhe5() As Boolean
   NaoGerarRegistroDetalhe5 = mblnNaoGerarRegistroDetalhe5
End Property

Public Property Let NaoGerarRegistroRodape1(Valor As Boolean)
   mblnNaoGerarRegistroRodape1 = Valor
End Property

Public Property Get NaoGerarRegistroRodape1() As Boolean
   NaoGerarRegistroRodape1 = mblnNaoGerarRegistroRodape1
End Property

Public Property Let NaoGerarRegistroRodape2(Valor As Boolean)
   mblnNaoGerarRegistroRodape2 = Valor
End Property

Public Property Get NaoGerarRegistroRodape2() As Boolean
   NaoGerarRegistroRodape2 = mblnNaoGerarRegistroRodape2
End Property

'Pt. 105912 - Moacir Pfau(17/03/2011)
Public Property Let TipoImpressao(Valor As String)
   mstrTipoImpressao = Valor
End Property

Public Property Get TipoImpressao() As String
   TipoImpressao = mstrTipoImpressao
End Property

'Pt. 106012 - Moacir Pfau(28/09/2011)
Public Property Let BairroSacado(Valor As String)
   mstrBairroSacado = Valor
End Property

Public Property Get BairroSacado() As String
   BairroSacado = mstrBairroSacado
End Property

Public Property Let CodigoPracaSacado(Valor As String)
   mstrCodigoPracaSacado = Valor
End Property

Public Property Get CodigoPracaSacado() As String
   CodigoPracaSacado = mstrCodigoPracaSacado
End Property

'Pt. 114032  - Moacir Pfau(23/02/2012)
Public Property Let IdentificacaoCedente(Valor As String)
   mstrIdentificacaoCedente = Valor
End Property

Public Property Get IdentificacaoCedente() As String
   IdentificacaoCedente = mstrIdentificacaoCedente
End Property

Public Property Let HTMLReciboPersonalizado(Valor As String)
   mstrHTMLReciboPersonalizado = Valor
End Property

Public Property Get HTMLReciboPersonalizado() As String
   HTMLReciboPersonalizado = mstrHTMLReciboPersonalizado
End Property

Public Property Let Data_baixa_retorno(Valor As Integer)
   mintData_baixa_retorno = Valor
End Property

Public Property Get Data_baixa_retorno() As Integer
   Data_baixa_retorno = mintData_baixa_retorno
End Property

Public Property Let localPagamento(Valor As String)
   mstrLocalPagamento = Valor
End Property

Public Property Get localPagamento() As String
   localPagamento = mstrLocalPagamento
End Property

Public Property Let InstrucaoCobranca1(Valor As String)
   mstrInstrucaoCobranca1 = Valor
End Property

Public Property Get InstrucaoCobranca1() As String
   InstrucaoCobranca1 = mstrInstrucaoCobranca1
End Property

Public Property Let InstrucaoCobranca2(Valor As String)
   mstrInstrucaoCobranca2 = Valor
End Property

Public Property Get InstrucaoCobranca2() As String
   InstrucaoCobranca2 = mstrInstrucaoCobranca2
End Property

Public Property Let InstrucaoCobranca3(Valor As String)
   mstrInstrucaoCobranca3 = Valor
End Property

Public Property Get InstrucaoCobranca3() As String
   InstrucaoCobranca3 = mstrInstrucaoCobranca3
End Property

Public Property Let ValorInstrucaoCobranca1(Valor As String)
   mstrValorInstrucaoCobranca1 = Valor
End Property

Public Property Get ValorInstrucaoCobranca1() As String
   ValorInstrucaoCobranca1 = mstrValorInstrucaoCobranca1
End Property

Public Property Let ValorInstrucaoCobranca2(Valor As String)
   mstrValorInstrucaoCobranca2 = Valor
End Property

Public Property Get ValorInstrucaoCobranca2() As String
   ValorInstrucaoCobranca2 = mstrValorInstrucaoCobranca2
End Property

Public Property Let ValorInstrucaoCobranca3(Valor As String)
   mstrValorInstrucaoCobranca3 = Valor
End Property

Public Property Get ValorInstrucaoCobranca3() As String
   ValorInstrucaoCobranca3 = mstrValorInstrucaoCobranca3
End Property

Public Property Let UtilizaNumeroControle(Valor As Boolean)
   mblnUtilizaNumeroControle = Valor
End Property

Public Property Get UtilizaNumeroControle() As Boolean
   UtilizaNumeroControle = mblnUtilizaNumeroControle
End Property

Public Property Get ColCpEspeciais() As clscolCarteiraCpEspecial
    Set ColCpEspeciais = mcolCpEspeciais
End Property

Public Property Let ColCpEspeciais(ByVal Valor As clscolCarteiraCpEspecial)
    Set mcolCpEspeciais = Valor
End Property

'Vinicius Elyseu(06/10/2015) - Projeto: #0 - História: #0 - Desenv: #0
Public Property Get SeqRemessaNrDoc() As Boolean
    SeqRemessaNrDoc = mblnSeqRemessaNrDoc
End Property

Public Property Let SeqRemessaNrDoc(ByVal Valor As Boolean)
    mblnSeqRemessaNrDoc = Valor
End Property

Private Sub Class_Initialize()
    Set mcolCpEspeciais = New clscolCarteiraCpEspecial
End Sub

Private Sub Class_Terminate()
    Set mcolCpEspeciais = Nothing
End Sub

'Projeto: #17081 - História: #0 - Desenvolvimento#0 - Moacir Pfau(29/11/2013)
Public Function CriarObjeto(ByVal rd As IDBReader) As clsCarteira
    With rd
        mlngEnterprise_id = .GetLong("enterprise_id")
        mlngCd_estabelecimento = .GetLong("cd_estabelecimento")
        mlngId_carteira = .GetLong("id_carteira")
        mstrDesc_carteira = .GetString("desc_carteira")
        mstrCodigo_cedente = .GetString("codigo_cedente")
        mstrInicio_nosso_numero = .GetString("inicio_nosso_numero")
        mstrFim_nosso_numero = .GetString("fim_nosso_numero")
        mstrProximo_nosso_numero = .GetString("proximo_nosso_numero")
        mstrDemonstrativo = .GetString("demonstrativo")
        mstrInstrucoes_caixa = .GetString("instrucoescaixa")
        mstrTipo_layout_boleto = .GetString("tipo_layout_boleto")
        mstrTipo_layout_remessa = .GetString("tipo_layout_remessa")
        mstrTipo_layout_retorno = .GetString("tipo_layout_retorno")
        mstrArquivo_licenca = .GetString("arquivo_licenca")
        mstrLogo_empresa = .GetString("logo_empresa")
        mstrCaminho_arquivo_remessa_padrao = .GetString("caminho_arquivo_remessa_padrao")
        mstrCaminho_arquivo_retorno_padrao = .GetString("caminho_arquivo_retorno_padrao")
        mlngMargem_superior_boleto = .GetLong("margem_superior")
        mstrOutro_dado_configuracao1 = .GetString("outrodadoconfiguracao1")
        mstrOutro_dado_configuracao2 = .GetString("outrodadoconfiguracao2")
        mlngSequencial_remessa = .GetLong("Sequencial_remessa")
        mstrEspecie = .GetString("especie")
        'Pt. 96589 - Moacir Pfau(05/02/2010)
        mblnBanco_gera_nosso_numero = .GetBoolean("Banco_gera_nosso_numero")
        'Pt. 97161 - Moacir Pfau(05/02/2010)
        mintDias_protesto = .GetInteger("Dias_protesto")
        'pt.98446 - Fernando Paludo(26/04/2010)
        mblnBancoEmiteBoleto = .GetBoolean("banco_emite_boleto")
        'pt.99257 - Moacir Pfau(30/06/2010)
        mdblPer_multa = .GetDouble("Per_multa")
        mdblPer_mora = .GetDouble("Per_mora")
        'pt.98929 - Fernando Paludo(02/08/2010)
        mstrOutra_especie = .GetString("outra_especie")
        'Pt. 102459 - Moacir Pfau(29/10/2010)
        mblnNaoGerarRegistroCabecalho1 = .GetBoolean("NaoGerarRegistroCabecalho1")
        mblnNaoGerarRegistroCabecalho2 = .GetBoolean("NaoGerarRegistroCabecalho2")
        mblnNaoGerarRegistroDetalhe1 = .GetBoolean("NaoGerarRegistroDetalhe1")
        mblnNaoGerarRegistroDetalhe2 = .GetBoolean("NaoGerarRegistroDetalhe2")
        mblnNaoGerarRegistroDetalhe3 = .GetBoolean("NaoGerarRegistroDetalhe3")
        mblnNaoGerarRegistroDetalhe4 = .GetBoolean("NaoGerarRegistroDetalhe4")
        mblnNaoGerarRegistroDetalhe5 = .GetBoolean("NaoGerarRegistroDetalhe5")
        mblnNaoGerarRegistroRodape1 = .GetBoolean("NaoGerarRegistroRodape1")
        mblnNaoGerarRegistroRodape2 = .GetBoolean("NaoGerarRegistroRodape2")
        'Pt. 105912 - Moacir Pfau(17/03/2011)
        mstrTipoImpressao = .GetString("TipoImpressao")
        'Pt. 106012 - Moacir Pfau(28/09/2011)
        mstrBairroSacado = .GetString("BairroSacado")
        mstrCodigoPracaSacado = .GetString("CodigoPracaSacado")
        'Pt. 114032 - Moacir Pfau(23/02/2012)
        mstrIdentificacaoCedente = .GetString("IdentificacaoCedente")
        mstrHTMLReciboPersonalizado = .GetString("HTMLReciboPersonalizado")
        'Projeto: #4350 - História: # - Desenvolvimento# - Moacir Pfau(09/04/2013)
        mintData_baixa_retorno = .GetInteger("data_baixa_retorno")
        'Projeto: #17081 - História: # - Desenvolvimento# - Moacir Pfau(02/12/2013)
        mstrLocalPagamento = .GetString("LocalPagamento")
        mstrInstrucaoCobranca1 = .GetString("InstrucaoCobranca1")
        mstrInstrucaoCobranca2 = .GetString("InstrucaoCobranca2")
        mstrInstrucaoCobranca3 = .GetString("InstrucaoCobranca3")
        mstrValorInstrucaoCobranca1 = .GetString("ValorInstrucaoCobranca1")
        mstrValorInstrucaoCobranca2 = .GetString("ValorInstrucaoCobranca2")
        mstrValorInstrucaoCobranca3 = .GetString("ValorInstrucaoCobranca3")
        mblnUtilizaNumeroControle = .GetBoolean("UtilizaNumeroControle")
        'Vinicius Elyseu(06/10/2015) - Projeto: #0 - História: #0 - Desenv: #0
        mblnSeqRemessaNrDoc = .GetBoolean("seq_remessa_nr_doc")
    End With
    Set CriarObjeto = Me
End Function
