VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCobreBemX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mobjCobreBemX        As CobreBemX.ContaCorrente
Private mobjBanco            As clsBanco
Private mlngCodigoBanco      As Long
Private mstrDiretorioArquivo As String
Private mstrNomeArquivo      As String

Public Property Let Banco(ByVal lngBanco As Long)
    mlngCodigoBanco = lngBanco
End Property
    
Public Property Let DiretorioArquivo(ByVal NewVal As String)
    mstrDiretorioArquivo = NewVal
End Property

Public Property Let NomeArquivo(ByVal NewVal As String)
    mstrNomeArquivo = NewVal
End Property

Private Sub Class_Initialize()
    mstrCodigoMoeda = "09"
    mstrOperacao = "000"
    Set mobjBanco = New clsBanco
    Set mobjCobreBemX = New ContaCorrente
End Sub

'Data.......: 19/01/2009
'Autor......: Ivo Sousa
'Descrição..: Passar os dados da conta corrente para a DLL de geração do arquivo
'OBS........: As linhas comentadas fazem parte do processo de geração, porém só são utilizadas caso
'             o sistema permita a impressão de boletos ou o envio dos dados por email.
Private Function SetDadosContaCorrente(objAplicacao As IApplication) As Boolean
    Dim selCmd       As IDBSelectCommand
    Dim rdResult     As IDBReader
    Dim strCodCedente As String
    
On Error GoTo ErroSetDados
    Set selCmd = objAplicacao.CreateSelectCommand
    SetDadosContaCorrente = True
    With selCmd
        .Table.TableName = "[Bancos]"
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", mlngCodigoBanco, dbFieldTypeLong))
    End With
    Set rdResult = objAplicacao.ExecuteReader(objAplicacao.GetInternalAuthorization, selCmd)
    If Not rdResult.EOF Then
        Set mobjBanco = New clsBanco
        'Carrega as informações do banco
        Call mobjBanco.CriarObjeto(rdResult, True)
        'Atribui os valores para a dll de geração
        With mobjBanco
            mobjCobreBemX.ArquivoLicenca = .ArquivoLicenca
            mobjCobreBemX.ArquivoRemessa.Layout = "FEBRABAN240"
            mobjCobreBemX.ArquivoRemessa.Diretorio = mstrDiretorioArquivo
            mobjCobreBemX.ArquivoRemessa.Arquivo = mstrNomeArquivo
            mobjCobreBemX.CodigoAgencia = Left(formatCampoInt(.Agencia, 5), 3) & "." & Right(formatCampoInt(.Agencia, 5), 2)
            mobjCobreBemX.NumeroContaCorrente = formatCampoInt(.Conta, 12) & "." & .DVConta
            'Formata o código do cedente somente para o Banrisul
            strCodCedente = formatCampoInt(Replace(.CodCedente, ".", ""), 12)
            strCodCedente = Left(strCodCedente, 9) & "." & Mid(strCodCedente, 10, 1) & "." & Right(strCodCedente, 2)
            mobjCobreBemX.CodigoCedente = strCodCedente
            mobjCobreBemX.ProximoNossoNumero = Format(.NrArquivoRemessa, mobjCobreBemX.MascaraNossoNumero)
            'Término dos parâmetros obrigatórios da conta corrente
    
            'Início da configuração dos dados do Cedente para envio de boletos por email
            mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.SMTP.Servidor = ConfigSys.S_SMTP
            mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.SMTP.Porta = ConfigSys.S_PORT
            mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.SMTP.Usuario = ConfigSys.S_USER
            mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.SMTP.Senha = ConfigSys.S_PASS
            'mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.URLImagensCodigoBarras = "C:\Fox\Programas\Codigo de Barras"
            'mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.URLLogotipo = "C:\Fox\Programas\Imagens"
            mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.Assunto = "Boleto de Cobrança"
            mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Endereco = EmpresaUsuaria.Email
            mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.EmailFrom.Nome = EmpresaUsuaria.Fantasia
            mobjCobreBemX.PadroesBoleto.PadroesBoletoEmail.PadroesEmail.FormaEnvio = feeSMTPBoletoHTML
            'Término da configuração dos dados do Cedente para envio de boletos por email
'
'            'Início da configuração dos dados do Cedente para impressão de boletos
'
'            'A próxima linha só é necessária se for solicitada a impressão de boletos e
'            'se desejar colocar o logotipo do Cedente na parte superior do boleto
'
            If mobjBanco.LogoEmpresa <> "" Then
                mobjCobreBemX.PadroesBoleto.PadroesBoletoImpresso.ArquivoLogotipo = mobjBanco.LogoEmpresa
            End If
            mobjCobreBemX.PadroesBoleto.PadroesBoletoImpresso.CaminhoImagensCodigoBarras = app.Path & "\Imagens\"
'
'            'Utilize o parâmetro abaixo para efetuar ajustes na impressão do boleto subindo ou descendo o mesmo na folha de papel
'            'Os valores devem ser informados em milímetros e quanto maior o valor mais para baixo será iniciado o boleto
'            'Se este parâmetro não for passado será assumido o valor 15 que é o indicado para a maioria das impressoras Jato de Tinta }
'
            mobjCobreBemX.PadroesBoleto.PadroesBoletoImpresso.MargemSuperior = 15
'
'            'Término da configuração dos dados do Cedente para impressão de boletos
'
'            'A próxima linha é utilizada para exibir um texto do lado direito do logotipo nos boletos impressos ou
'            'enviados por email
'            mobjCobreBemX.PadroesBoleto.IdentificacaoCedente = VerificaNulo(datPrimaryRS.Recordset("IdentificacaoCedenteBolet"))
        End With
    End If
    Exit Function
ErroSetDados:
    objAplicacao = False
    Set rdResult = Nothing
End Function

'Data.......: 19/01/2009
'Autor......: Ivo Sousa
'Descrição..: Passar os dados do boleto para a DLL de geração do arquivo
'OBS........: As linhas comentadas fazem parte do processo de geração, porém só são utilizadas caso
'             o sistema permita a impressão de boletos ou o envio dos dados por email.
Public Function SetDadosBoletos(rdResult As IDBReader, objAplicacao As IApplication) As Boolean
    Dim objBoleto  As Object
    Dim Email      As Object
    Dim rs         As ADODB.Recordset
    Dim objEmpresa As CEmpresas
    Dim lngNosNum  As String
    Dim strSql     As String
    
On Error GoTo ErroSet
    SetDadosBoletos = True
    If Not SetDadosContaCorrente(objAplicacao) Then
        SetDadosBoletos = False
        Exit Function
    End If
    mobjCobreBemX.DocumentosCobranca.clear
    'Carrega as informações sobre a conta corrente
    rdResult.MoveFirst
    lngNosNum = Format(GetFieldValue("PROSEQ", "BANCOS", "BANCO = " & mlngCodigoBanco, , 1) + 1, mobjCobreBemX.MascaraNossoNumero)
    While Not rdResult.EOF
        With rdResult
            'Carrega os dados do sacado na classe empresa
            Set objEmpresa = New CEmpresas
            'Adiciona um novo registro na classe dos documentos de cobrança
            Set objBoleto = mobjCobreBemX.DocumentosCobranca.add
            objBoleto.NumeroDocumento = .GetLong("Nota")
            objBoleto.NomeSacado = .GetString("Empresa")
            Call objEmpresa.CarregarRegistro(objBoleto.NomeSacado)
            If Len(limpaDOC(objEmpresa.CNPJ_CPF)) > 11 Then
                objBoleto.CNPJSacado = limpaDOC(objEmpresa.CNPJ_CPF)
            Else
                objBoleto.CPFSacado = limpaDOC(objEmpresa.CNPJ_CPF)
            End If
            objBoleto.EnderecoSacado = objEmpresa.Endereco
            objBoleto.BairroSacado = objEmpresa.Bairro
            objBoleto.CidadeSacado = objEmpresa.Cidade
            objBoleto.EstadoSacado = objEmpresa.Estado
            objBoleto.CepSacado = limpaDOC(objEmpresa.Cep)
            objBoleto.DataDocumento = .GetDate("Emissão")
            If Not IsNull(.GetDate("Vencimento")) Then
                objBoleto.DataVencimento = .GetDate("Vencimento")
            End If
            objBoleto.DataProcessamento = Format$(Date, "dd/mm/yyyy")
            objBoleto.ValorDocumento = .GetCurrency("Valor Original")
            objBoleto.PercentualJurosDiaAtraso = .GetDouble("PerJur")
            objBoleto.PercentualMultaAtraso = .GetCurrency("VlrMul")
            objBoleto.PercentualDesconto = 0
            objBoleto.ValorOutrosAcrescimos = .GetCurrency("Acréscimo")
            If Trim(.GetString("NosNum")) <> "" Then
                If CLng(.GetString("NosNum")) > 0 Then
                    objBoleto.NossoNumero = Format(.GetString("NosNum"), mobjCobreBemX.MascaraNossoNumero)
                Else
                    objBoleto.NossoNumero = lngNosNum
                    lngNosNum = lngNosNum + 1
                End If
            Else
                objBoleto.NossoNumero = lngNosNum
                lngNosNum = lngNosNum + 1
            End If
            
            'objBoleto.PadroesBoleto.Demonstrativo = .GetString("Tipo")
            objBoleto.PadroesBoleto.InstrucoesCaixa = mobjBanco.Instrucoes
            
            'Informações sobre o email do sacado
            Set Email = objBoleto.EnderecosEmailSacado.add
            Email.Nome = objBoleto.NomeSacado
            Email.Endereco = objEmpresa.e_mail
                        
            objBoleto.ControleProcessamentoDocumento.Imprime = scpExecutar
            objBoleto.ControleProcessamentoDocumento.EnviaEmail = scpExecutar
            objBoleto.ControleProcessamentoDocumento.GravaRemessa = scpExecutar
            
            mobjCobreBemX.CalcularDadosBoletos
                
            strSql = "UPDATE Duplicatas SET LINDIG = '" & objBoleto.LinhaDigitavel & "', CODBAR = '" & objBoleto.CodigoBarras
            strSql = strSql & "', NOSNUM = '" & objBoleto.NossoNumero & "', INSTRU = '" & objBoleto.PadroesBoleto.InstrucoesCaixa
            strSql = strSql & "' WHERE Nota = " & rdResult.GetLong("Nota") & " AND Empresa = '" & rdResult.GetString("Empresa")
            strSql = strSql & "' AND Parcela = " & rdResult.GetLong("Parcela") & " AND Tipo = '" & rdResult.GetString("Tipo")
            strSql = strSql & "' AND PagRec = 'R'"
            Call ExecuteSQL(strSql)
            strSql = ""
            Set objEmpresa = Nothing
            rdResult.MoveNext
        End With
    Wend
    strSql = "UPDATE Bancos SET PROSEQ =" & lngNosNum & " WHERE Banco=" & mlngCodigoBanco
    Call ExecuteSQL(strSql)
    Exit Function
ErroSet:
    SetDadosBoletos = False
End Function

'Data.......: 19/01/2009
'Autor......: Ivo Sousa
'Descrição..: Verirfica se um campo é nulo e retorna um valor default
'Parametro..: [Field] Campo que será analisado
Private Function VerificaNulo(Campo As Field) As Variant
    If IsNull(Campo) Then
        If Campo.Type = adVarChar Then
            VerificaNulo = ""
        Else
            VerificaNulo = 0
        End If
    Else
        VerificaNulo = Campo
    End If
End Function

'Data.......: 02/02/2009
'Autor......: Ivo Sousa
'Descrição..: Grava o arquivo com os registros e atualiza a tabela com os mesmos
Public Function GravaArquivoRemessa(rdResult As IDBReader, blnImprimir As Boolean, objAplicacao As Object)
    
On Error GoTo ErrHadler
    GravaArquivoRemessa = True
    If blnImprimir Then
        mobjCobreBemX.ImprimeBoletos
    End If
    mobjCobreBemX.GravaArquivoRemessa
    mstrDiretorioArquivo = mobjCobreBemX.ArquivoRemessa.Diretorio
ErrHadler:
    GravaArquivoRemessa = True
End Function
