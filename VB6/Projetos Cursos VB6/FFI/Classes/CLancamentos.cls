VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLancamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************
'    Classe Criada com o Construtor de Classes EBS Sistemas
'*************************************************************
'
'    Nome da Classe: CLancamentos
'    Escrito por: Edilberto Sandro da Silva
'
'*************************************************************
Private oCONEXAO   As Adodb.Connection
Private rs         As Adodb.Recordset
Private bVAZIO     As Boolean


Private sPagRec As String
Private sCodigo As String
Private sEmpresa As String
Private sTipo As String
Private sDescricao As String
Private sEmissao As String
Private sVencimento As String
Private sPagamento As String
Private sLiberacao As String
Private sValorOriginal As String
Private sAcrescimo As String
Private sAbatimento As String
Private sBanco As String
Private sConta As String
Private sCentro As String
Private sCheque As String
Private sMoeda As String
Private sValordaMoeda As String
Private sControle As String
Private sSituacao As String
Private sEnviado As String
Private sMarcacao As String
Private sObs As String
Private sBordero As String
Private sDataIntegracaoFatoGerador As String
Private sDataIntegracaoFatoPagamento As String
Private sLinhaDigitavel As String
Private sUsuario As String
Private sAlteracao As String
Private sConciliado As String
Private sCarteira As String
Private sSeqNossoNumero As String
Private sPrevisao As String
Private sVlrMrD As String
Private sVlrDsp As String
Private sPerJur As String
Private sVlrJur As String
Private sCODPRJ As String
Private sCODFPG As String
Private sCheBan As String
Private sCheAge As String
Private sCheEmi As String
Private sCheCco As String
Private sPerMul As String
Private sVlrMul As String

Public Property Get CON() As Adodb.Connection
  Set CON = oCONEXAO
End Property

Public Property Get VAZIO() As Boolean
  VAZIO = bVAZIO
End Property

Public Property Let PagRec(NovoValor As String)
    sPagRec = NovoValor
End Property

Public Property Get PagRec() As String
   PagRec = sPagRec
End Property

Public Property Let Codigo(NovoValor As String)
    sCodigo = NovoValor
End Property

Public Property Get Codigo() As String
   Codigo = sCodigo
End Property

Public Property Let Empresa(NovoValor As String)
    sEmpresa = NovoValor
End Property

Public Property Get Empresa() As String
   Empresa = sEmpresa
End Property

Public Property Let Tipo(NovoValor As String)
    sTipo = NovoValor
End Property

Public Property Get Tipo() As String
   Tipo = sTipo
End Property

Public Property Let Descricao(NovoValor As String)
    sDescricao = NovoValor
End Property

Public Property Get Descricao() As String
   Descricao = sDescricao
End Property

Public Property Let Emissao(NovoValor As String)
    sEmissao = NovoValor
End Property

Public Property Get Emissao() As String
   Emissao = sEmissao
End Property

Public Property Let Vencimento(NovoValor As String)
    sVencimento = NovoValor
End Property

Public Property Get Vencimento() As String
   Vencimento = sVencimento
End Property

Public Property Let Pagamento(NovoValor As String)
    sPagamento = NovoValor
End Property

Public Property Get Pagamento() As String
   Pagamento = sPagamento
End Property

Public Property Let Liberacao(NovoValor As String)
    sLiberacao = NovoValor
End Property

Public Property Get Liberacao() As String
   Liberacao = sLiberacao
End Property

Public Property Let ValorOriginal(NovoValor As String)
    sValorOriginal = NovoValor
End Property

Public Property Get ValorOriginal() As String
   ValorOriginal = sValorOriginal
End Property

Public Property Let Acrescimo(NovoValor As String)
    sAcrescimo = NovoValor
End Property

Public Property Get Acrescimo() As String
   Acrescimo = sAcrescimo
End Property

Public Property Let Abatimento(NovoValor As String)
    sAbatimento = NovoValor
End Property

Public Property Get Abatimento() As String
   Abatimento = sAbatimento
End Property

Public Property Let Banco(NovoValor As String)
    sBanco = NovoValor
End Property

Public Property Get Banco() As String
   Banco = sBanco
End Property

Public Property Let Conta(NovoValor As String)
    sConta = NovoValor
End Property

Public Property Get Conta() As String
   Conta = sConta
End Property

Public Property Let Centro(NovoValor As String)
    sCentro = NovoValor
End Property

Public Property Get Centro() As String
   Centro = sCentro
End Property

Public Property Let Cheque(NovoValor As String)
    sCheque = NovoValor
End Property

Public Property Get Cheque() As String
   Cheque = sCheque
End Property

Public Property Let Moeda(NovoValor As String)
    sMoeda = NovoValor
End Property

Public Property Get Moeda() As String
   Moeda = sMoeda
End Property

Public Property Let ValordaMoeda(NovoValor As String)
    sValordaMoeda = NovoValor
End Property

Public Property Get ValordaMoeda() As String
   ValordaMoeda = sValordaMoeda
End Property

Public Property Let Controle(NovoValor As String)
    sControle = NovoValor
End Property

Public Property Get Controle() As String
   Controle = sControle
End Property

Public Property Let Situacao(NovoValor As String)
    sSituacao = NovoValor
End Property

Public Property Get Situacao() As String
   Situacao = sSituacao
End Property

Public Property Let Enviado(NovoValor As String)
    sEnviado = NovoValor
End Property

Public Property Get Enviado() As String
   Enviado = sEnviado
End Property

Public Property Let Marcacao(NovoValor As String)
    sMarcacao = NovoValor
End Property

Public Property Get Marcacao() As String
   Marcacao = sMarcacao
End Property

Public Property Let Obs(NovoValor As String)
    sObs = NovoValor
End Property

Public Property Get Obs() As String
   Obs = sObs
End Property

Public Property Let Bordero(NovoValor As String)
    sBordero = NovoValor
End Property

Public Property Get Bordero() As String
   Bordero = sBordero
End Property

Public Property Let DataIntegracaoFatoGerador(NovoValor As String)
    sDataIntegracaoFatoGerador = NovoValor
End Property

Public Property Get DataIntegracaoFatoGerador() As String
   DataIntegracaoFatoGerador = sDataIntegracaoFatoGerador
End Property

Public Property Let DataIntegracaoFatoPagamento(NovoValor As String)
    sDataIntegracaoFatoPagamento = NovoValor
End Property

Public Property Get DataIntegracaoFatoPagamento() As String
   DataIntegracaoFatoPagamento = sDataIntegracaoFatoPagamento
End Property

Public Property Let LinhaDigitavel(NovoValor As String)
    sLinhaDigitavel = NovoValor
End Property

Public Property Get LinhaDigitavel() As String
   LinhaDigitavel = sLinhaDigitavel
End Property

Public Property Let Usuario(NovoValor As String)
    sUsuario = NovoValor
End Property

Public Property Get Usuario() As String
   Usuario = sUsuario
End Property

Public Property Let Alteracao(NovoValor As String)
    sAlteracao = NovoValor
End Property

Public Property Get Alteracao() As String
   Alteracao = sAlteracao
End Property

Public Property Let Conciliado(NovoValor As String)
    sConciliado = NovoValor
End Property

Public Property Get Conciliado() As String
   Conciliado = sConciliado
End Property

Public Property Let Carteira(NovoValor As String)
    sCarteira = NovoValor
End Property

Public Property Get Carteira() As String
   Carteira = sCarteira
End Property

Public Property Let SeqNossoNumero(NovoValor As String)
    sSeqNossoNumero = NovoValor
End Property

Public Property Get SeqNossoNumero() As String
   SeqNossoNumero = sSeqNossoNumero
End Property

Public Property Let Previsao(NovoValor As String)
    sPrevisao = NovoValor
End Property

Public Property Get Previsao() As String
   Previsao = sPrevisao
End Property

Public Property Let VlrMrD(NovoValor As String)
    sVlrMrD = NovoValor
End Property

Public Property Get VlrMrD() As String
   VlrMrD = sVlrMrD
End Property

Public Property Let VlrDsp(NovoValor As String)
    sVlrDsp = NovoValor
End Property

Public Property Get VlrDsp() As String
   VlrDsp = sVlrDsp
End Property

Public Property Let PerJur(NovoValor As String)
    sPerJur = NovoValor
End Property

Public Property Get PerJur() As String
   PerJur = sPerJur
End Property

Public Property Let VlrJur(NovoValor As String)
    sVlrJur = NovoValor
End Property

Public Property Get VlrJur() As String
   VlrJur = sVlrJur
End Property

Public Property Let CODPRJ(NovoValor As String)
    sCODPRJ = NovoValor
End Property

Public Property Get CODPRJ() As String
   CODPRJ = sCODPRJ
End Property

Public Property Let CODFPG(NovoValor As String)
    sCODFPG = NovoValor
End Property

Public Property Get CODFPG() As String
   CODFPG = sCODFPG
End Property

Public Property Let CheBan(NovoValor As String)
    sCheBan = NovoValor
End Property

Public Property Get CheBan() As String
   CheBan = sCheBan
End Property

Public Property Let CheAge(NovoValor As String)
    sCheAge = NovoValor
End Property

Public Property Get CheAge() As String
   CheAge = sCheAge
End Property

Public Property Let CheEmi(NovoValor As String)
    sCheEmi = NovoValor
End Property

Public Property Get CheEmi() As String
   CheEmi = sCheEmi
End Property

Public Property Let CheCco(NovoValor As String)
    sCheCco = NovoValor
End Property

Public Property Get CheCco() As String
   CheCco = sCheCco
End Property

Public Property Let PerMul(NovoValor As String)
    sPerMul = NovoValor
End Property

Public Property Get PerMul() As String
   PerMul = sPerMul
End Property

Public Property Let VlrMul(NovoValor As String)
    sVlrMul = NovoValor
End Property

Public Property Get VlrMul() As String
   VlrMul = sVlrMul
End Property

Public Function Inserir() As Boolean
Dim SQL1 As String

On Error GoTo ErroInserir
   SQL1 = "INSERT INTO [Lançamentos] ("
   If sPagRec <> "" Then SQL1 = SQL1 + "[PagRec], "
   If sCodigo <> "" Then SQL1 = SQL1 + "[Código], "
   If sEmpresa <> "" Then SQL1 = SQL1 + "[Empresa], "
   If sTipo <> "" Then SQL1 = SQL1 + "[Tipo], "
   If sDescricao <> "" Then SQL1 = SQL1 + "[Descrição], "
   If sEmissao <> "" Then SQL1 = SQL1 + "[Emissão], "
   If sVencimento <> "" Then SQL1 = SQL1 + "[Vencimento], "
   If sPagamento <> "" Then SQL1 = SQL1 + "[Pagamento], "
   If sLiberacao <> "" Then SQL1 = SQL1 + "[Liberação], "
   If sValorOriginal <> "" Then SQL1 = SQL1 + "[Valor Original], "
   If sAcrescimo <> "" Then SQL1 = SQL1 + "[Acréscimo], "
   If sAbatimento <> "" Then SQL1 = SQL1 + "[Abatimento], "
   If sBanco <> "" Then SQL1 = SQL1 + "[Banco], "
   If sConta <> "" Then SQL1 = SQL1 + "[Conta], "
   If sCentro <> "" Then SQL1 = SQL1 + "[Centro], "
   If sCheque <> "" Then SQL1 = SQL1 + "[Cheque], "
   If sMoeda <> "" Then SQL1 = SQL1 + "[Moeda], "
   If sValordaMoeda <> "" Then SQL1 = SQL1 + "[Valor da Moeda], "
   If sControle <> "" Then SQL1 = SQL1 + "[Controle], "
   If sSituacao <> "" Then SQL1 = SQL1 + "[Situação], "
   If sEnviado <> "" Then SQL1 = SQL1 + "[Enviado], "
   If sMarcacao <> "" Then SQL1 = SQL1 + "[Marcação], "
   If sObs <> "" Then SQL1 = SQL1 + "[Obs], "
   If sBordero <> "" Then SQL1 = SQL1 + "[Borderô], "
   If sDataIntegracaoFatoGerador <> "" Then SQL1 = SQL1 + "[Data Integração Fato Gerador], "
   If sDataIntegracaoFatoPagamento <> "" Then SQL1 = SQL1 + "[Data Integração Fato Pagamento], "
   If sLinhaDigitavel <> "" Then SQL1 = SQL1 + "[Linha Digitável], "
   If sUsuario <> "" Then SQL1 = SQL1 + "[Usuário], "
   If sAlteracao <> "" Then SQL1 = SQL1 + "[Alteração], "
   If sConciliado <> "" Then SQL1 = SQL1 + "[Conciliado], "
   If sCarteira <> "" Then SQL1 = SQL1 + "[Carteira], "
   If sSeqNossoNumero <> "" Then SQL1 = SQL1 + "[SeqNossoNumero], "
   If sPrevisao <> "" Then SQL1 = SQL1 + "[Previsao], "
   If sVlrMrD <> "" Then SQL1 = SQL1 + "[VlrMrD], "
   If sVlrDsp <> "" Then SQL1 = SQL1 + "[VlrDsp], "
   If sPerJur <> "" Then SQL1 = SQL1 + "[PerJur], "
   If sVlrJur <> "" Then SQL1 = SQL1 + "[VlrJur], "
   If sCODPRJ <> "" Then SQL1 = SQL1 + "[CODPRJ], "
   If sCODFPG <> "" Then SQL1 = SQL1 + "[CODFPG], "
   If sCheBan <> "" Then SQL1 = SQL1 + "[CheBan], "
   If sCheAge <> "" Then SQL1 = SQL1 + "[CheAge], "
   If sCheEmi <> "" Then SQL1 = SQL1 + "[CheEmi], "
   If sCheCco <> "" Then SQL1 = SQL1 + "[CheCco], "
   If sPerMul <> "" Then SQL1 = SQL1 + "[PerMul], "
   If sVlrMul <> "" Then SQL1 = SQL1 + "[VlrMul]) "
   If Right(RTrim(SQL1), 1) = "," Then SQL1 = Left$(SQL1, Len(SQL1) - 2) + ") "

   SQL1 = SQL1 + "VALUES ("
   
   If sPagRec <> "" Then SQL1 = SQL1 + Quote(sPagRec, "'") + ", "
   If sCodigo <> "" Then SQL1 = SQL1 + str(sCodigo) + ", "
   If sEmpresa <> "" Then SQL1 = SQL1 + Quote(sEmpresa, "'") + ", "
   If sTipo <> "" Then SQL1 = SQL1 + Quote(sTipo, "'") + ", "
   If sDescricao <> "" Then SQL1 = SQL1 + Quote(sDescricao, "'") + ", "
   If sEmissao <> "" Then SQL1 = SQL1 + InverteData(sEmissao) + ", "
   If sVencimento <> "" Then SQL1 = SQL1 + InverteData(sVencimento) + ", "
   If sPagamento <> "" Then SQL1 = SQL1 + InverteData(sPagamento) + ", "
   If sLiberacao <> "" Then SQL1 = SQL1 + InverteData(sLiberacao) + ", "
   If sValorOriginal <> "" Then SQL1 = SQL1 + str(sValorOriginal) + ", "
   If sAcrescimo <> "" Then SQL1 = SQL1 + str(sAcrescimo) + ", "
   If sAbatimento <> "" Then SQL1 = SQL1 + str(sAbatimento) + ", "
   If sBanco <> "" Then SQL1 = SQL1 + str(sBanco) + ", "
   If sConta <> "" Then SQL1 = SQL1 + str(sConta) + ", "
   If sCentro <> "" Then SQL1 = SQL1 + str(sCentro) + ", "
   If sCheque <> "" Then SQL1 = SQL1 + str(sCheque) + ", "
   If sMoeda <> "" Then SQL1 = SQL1 + Quote(sMoeda, "'") + ", "
   If sValordaMoeda <> "" Then SQL1 = SQL1 + str(sValordaMoeda) + ", "
   If sControle <> "" Then SQL1 = SQL1 + Quote(sControle, "'") + ", "
   If sSituacao <> "" Then SQL1 = SQL1 + Quote(sSituacao, "'") + ", "
   If sEnviado <> "" Then SQL1 = SQL1 + Quote(sEnviado, "'") + ", "
   If sMarcacao <> "" Then SQL1 = SQL1 + str(sMarcacao) + ", "
   If sObs <> "" Then SQL1 = SQL1 + Quote(sObs, "'") + ", "
   If sBordero <> "" Then SQL1 = SQL1 + str(sBordero) + ", "
   If sDataIntegracaoFatoGerador <> "" Then SQL1 = SQL1 + InverteData(sDataIntegracaoFatoGerador) + ", "
   If sDataIntegracaoFatoPagamento <> "" Then SQL1 = SQL1 + InverteData(sDataIntegracaoFatoPagamento) + ", "
   If sLinhaDigitavel <> "" Then SQL1 = SQL1 + Quote(sLinhaDigitavel, "'") + ", "
   If sUsuario <> "" Then SQL1 = SQL1 + Quote(sUsuario, "'") + ", "
   If sAlteracao <> "" Then SQL1 = SQL1 + InverteData(sAlteracao) + ", "
   If sConciliado <> "" Then SQL1 = SQL1 + str(sConciliado) + ", "
   If sCarteira <> "" Then SQL1 = SQL1 + str(sCarteira) + ", "
   If sSeqNossoNumero <> "" Then SQL1 = SQL1 + str(sSeqNossoNumero) + ", "
   If sPrevisao <> "" Then SQL1 = SQL1 + str(sPrevisao) + ", "
   If sVlrMrD <> "" Then SQL1 = SQL1 + str(sVlrMrD) + ", "
   If sVlrDsp <> "" Then SQL1 = SQL1 + str(sVlrDsp) + ", "
   If sPerJur <> "" Then SQL1 = SQL1 + str(sPerJur) + ", "
   If sVlrJur <> "" Then SQL1 = SQL1 + str(sVlrJur) + ", "
   If sCODPRJ <> "" Then SQL1 = SQL1 + str(sCODPRJ) + ", "
   If sCODFPG <> "" Then SQL1 = SQL1 + str(sCODFPG) + ", "
   If sCheBan <> "" Then SQL1 = SQL1 + str(sCheBan) + ", "
   If sCheAge <> "" Then SQL1 = SQL1 + Quote(sCheAge, "'") + ", "
   If sCheEmi <> "" Then SQL1 = SQL1 + Quote(sCheEmi, "'") + ", "
   If sCheCco <> "" Then SQL1 = SQL1 + Quote(sCheCco, "'") + ", "
   If sPerMul <> "" Then SQL1 = SQL1 + str(sPerMul) + ", "
   If sVlrMul <> "" Then SQL1 = SQL1 + str(sVlrMul) + ") "
   If Right(RTrim(SQL1), 1) = "," Then SQL1 = Left$(SQL1, Len(SQL1) - 2) + ")"
   oCONEXAO.Execute SQL1
   Inserir = True
   Exit Function
ErroInserir:
   Err.Raise 1 + vbObjectError, , "Não foi possível cadastrar pagrec!"
   Inserir = False
End Function

Public Function Atualizar(pPagRec As String, pCodigo As String) As Boolean
Dim SQL1 As String
On Error GoTo ErroAtualizar

   SQL1 = ""
   SQL1 = "UPDATE [Lançamentos] SET "
   'If sPagRec <> "" Then SQL1 = SQL1 + "[PagRec] = " + Quote(sPagRec,"'") + ", "
   'If sCodigo <> "" Then SQL1 = SQL1 + "[Código] = " + str(sCodigo) + ", "
   If sEmpresa <> "" Then SQL1 = SQL1 + "[Empresa] = " + Quote(sEmpresa, "'") + ", "
   If sTipo <> "" Then SQL1 = SQL1 + "[Tipo] = " + Quote(sTipo, "'") + ", "
   If sDescricao <> "" Then SQL1 = SQL1 + "[Descrição] = " + Quote(sDescricao, "'") + ", "
   If sEmissao <> "" Then SQL1 = SQL1 + "[Emissão] = " + InverteData(sEmissao) + ", "
   If sVencimento <> "" Then SQL1 = SQL1 + "[Vencimento] = " + InverteData(sVencimento) + ", "
   If sPagamento <> "" Then SQL1 = SQL1 + "[Pagamento] = " + InverteData(sPagamento) + ", "
   If sLiberacao <> "" Then SQL1 = SQL1 + "[Liberação] = " + InverteData(sLiberacao) + ", "
   If sValorOriginal <> "" Then SQL1 = SQL1 + "[Valor Original] = " + str(sValorOriginal) + ", "
   If sAcrescimo <> "" Then SQL1 = SQL1 + "[Acréscimo] = " + str(sAcrescimo) + ", "
   If sAbatimento <> "" Then SQL1 = SQL1 + "[Abatimento] = " + str(sAbatimento) + ", "
   If sBanco <> "" Then SQL1 = SQL1 + "[Banco] = " + str(sBanco) + ", "
   If sConta <> "" Then SQL1 = SQL1 + "[Conta] = " + str(sConta) + ", "
   If sCentro <> "" Then SQL1 = SQL1 + "[Centro] = " + str(sCentro) + ", "
   If sCheque <> "" Then SQL1 = SQL1 + "[Cheque] = " + str(sCheque) + ", "
   If sMoeda <> "" Then SQL1 = SQL1 + "[Moeda] = " + Quote(sMoeda, "'") + ", "
   If sValordaMoeda <> "" Then SQL1 = SQL1 + "[Valor da Moeda] = " + str(sValordaMoeda) + ", "
   If sControle <> "" Then SQL1 = SQL1 + "[Controle] = " + Quote(sControle, "'") + ", "
   If sSituacao <> "" Then SQL1 = SQL1 + "[Situação] = " + Quote(sSituacao, "'") + ", "
   If sEnviado <> "" Then SQL1 = SQL1 + "[Enviado] = " + Quote(sEnviado, "'") + ", "
   If sMarcacao <> "" Then SQL1 = SQL1 + "[Marcação] = " + str(sMarcacao) + ", "
   If sObs <> "" Then SQL1 = SQL1 + "[Obs] = " + Quote(sObs, "'") + ", "
   If sBordero <> "" Then SQL1 = SQL1 + "[Borderô] = " + str(sBordero) + ", "
   If sDataIntegracaoFatoGerador <> "" Then SQL1 = SQL1 + "[Data Integração Fato Gerador] = " + InverteData(sDataIntegracaoFatoGerador) + ", "
   If sDataIntegracaoFatoPagamento <> "" Then SQL1 = SQL1 + "[Data Integração Fato Pagamento] = " + InverteData(sDataIntegracaoFatoPagamento) + ", "
   If sLinhaDigitavel <> "" Then SQL1 = SQL1 + "[Linha Digitável] = " + Quote(sLinhaDigitavel, "'") + ", "
   If sUsuario <> "" Then SQL1 = SQL1 + "[Usuário] = " + Quote(sUsuario, "'") + ", "
   If sAlteracao <> "" Then SQL1 = SQL1 + "[Alteração] = " + InverteData(sAlteracao) + ", "
   If sConciliado <> "" Then SQL1 = SQL1 + "[Conciliado] = " + str(sConciliado) + ", "
   If sCarteira <> "" Then SQL1 = SQL1 + "[Carteira] = " + str(sCarteira) + ", "
   If sSeqNossoNumero <> "" Then SQL1 = SQL1 + "[SeqNossoNumero] = " + str(sSeqNossoNumero) + ", "
   If sPrevisao <> "" Then SQL1 = SQL1 + "[Previsao] = " + str(sPrevisao) + ", "
   If sVlrMrD <> "" Then SQL1 = SQL1 + "[VlrMrD] = " + str(sVlrMrD) + ", "
   If sVlrDsp <> "" Then SQL1 = SQL1 + "[VlrDsp] = " + str(sVlrDsp) + ", "
   If sPerJur <> "" Then SQL1 = SQL1 + "[PerJur] = " + str(sPerJur) + ", "
   If sVlrJur <> "" Then SQL1 = SQL1 + "[VlrJur] = " + str(sVlrJur) + ", "
   If sCODPRJ <> "" Then SQL1 = SQL1 + "[CODPRJ] = " + str(sCODPRJ) + ", "
   If sCODFPG <> "" Then SQL1 = SQL1 + "[CODFPG] = " + str(sCODFPG) + ", "
   If sCheBan <> "" Then SQL1 = SQL1 + "[CheBan] = " + str(sCheBan) + ", "
   If sCheAge <> "" Then SQL1 = SQL1 + "[CheAge] = " + Quote(sCheAge, "'") + ", "
   If sCheEmi <> "" Then SQL1 = SQL1 + "[CheEmi] = " + Quote(sCheEmi, "'") + ", "
   If sCheCco <> "" Then SQL1 = SQL1 + "[CheCco] = " + Quote(sCheCco, "'") + ", "
   If sPerMul <> "" Then SQL1 = SQL1 + "[PerMul] = " + str(sPerMul) + ", "
   If sVlrMul <> "" Then SQL1 = SQL1 + "[VlrMul] = " + str(sVlrMul) + " "
   If Right(RTrim(SQL1), 1) = "," Then SQL1 = Left$(SQL1, Len(SQL1) - 2) + " "

   SQL1 = SQL1 + "WHERE [PagRec] = '" & pPagRec & "' AND [Código] = " & pCodigo

   oCONEXAO.Execute SQL1
   Atualizar = True
   Exit Function
ErroAtualizar:
   Err.Raise 1 + vbObjectError, , "Não foi possível atualizar o registro!"
   Atualizar = False
End Function

Public Function Remover(pPagRec As String, pCodigo As String) As Boolean
   On Error GoTo ErroRemover
   oCONEXAO.Execute "DELETE FROM [Lançamentos] WHERE [PagRec] = '" & pPagRec & "' AND [Código] = " & pCodigo
   Remover = True
   Exit Function
ErroRemover:
   Err.Raise 1 + vbObjectError, , "O banco de dados não permitiu " + vbCrLf + "excluir o registro!"
   Remover = False
End Function

Public Function Existe(pPagRec As String, pCodigo As String) As Boolean
'Esta função carrega o registro no caso do pagrec existir.
On Error GoTo ErroExiste
   Set rs = New Adodb.Recordset
   rs.Open "SELECT * FROM [Lançamentos] WHERE [PagRec] = '" & pPagRec & "' AND [Código] = " & pCodigo, oCONEXAO, adOpenKeyset, adLockPessimistic
   If Not rs.BOF Then
      Existe = True
   Else
      Existe = False
   End If
   rs.Close
   Set rs = Nothing
   Exit Function
ErroExiste:
   Err.Raise 1 + vbObjectError, , "Não foi possível localizar pagrec!"
   rs.Close
   Set rs = Nothing
End Function

Public Function CarregarRegistro(pPagRec As String, pCodigo As String) As Boolean
'Esta função carrega o registro no caso do pagrec existir.
On Error GoTo ErroCarregarRegistro
   Set rs = New Adodb.Recordset
   rs.Open "SELECT * FROM [Lançamentos] WHERE [PagRec] = '" & pPagRec & "' AND [Código] = " & pCodigo, oCONEXAO, adOpenKeyset, adLockPessimistic
   If Not rs.BOF Then
         sPagRec = SeCampoAdoNulo(rs![PagRec])
         sCodigo = SeCampoAdoNulo(rs![Código])
         sEmpresa = SeCampoAdoNulo(rs![Empresa])
         sTipo = SeCampoAdoNulo(rs![Tipo])
         sDescricao = SeCampoAdoNulo(rs![Descrição])
         sEmissao = SeCampoAdoNulo(rs![Emissão])
         sVencimento = SeCampoAdoNulo(rs![Vencimento])
         sPagamento = SeCampoAdoNulo(rs![Pagamento])
         sLiberacao = SeCampoAdoNulo(rs![Liberação])
         sValorOriginal = SeCampoAdoNulo(rs![Valor Original])
         sAcrescimo = SeCampoAdoNulo(rs![Acréscimo])
         sAbatimento = SeCampoAdoNulo(rs![Abatimento])
         sBanco = SeCampoAdoNulo(rs![Banco])
         sConta = SeCampoAdoNulo(rs![Conta])
         sCentro = SeCampoAdoNulo(rs![Centro])
         sCheque = SeCampoAdoNulo(rs![Cheque])
         sMoeda = SeCampoAdoNulo(rs![Moeda])
         sValordaMoeda = SeCampoAdoNulo(rs![Valor da Moeda])
         sControle = SeCampoAdoNulo(rs![Controle])
         sSituacao = SeCampoAdoNulo(rs![Situação])
         sEnviado = SeCampoAdoNulo(rs![Enviado])
         sMarcacao = SeCampoAdoNulo(rs![Marcação])
         sObs = SeCampoAdoNulo(rs![Obs])
         sBordero = SeCampoAdoNulo(rs![Borderô])
         sDataIntegracaoFatoGerador = SeCampoAdoNulo(rs![Data Integração Fato Gerador])
         sDataIntegracaoFatoPagamento = SeCampoAdoNulo(rs![Data Integração Fato Pagamento])
         sLinhaDigitavel = SeCampoAdoNulo(rs![Linha Digitável])
         sUsuario = SeCampoAdoNulo(rs![Usuário])
         sAlteracao = SeCampoAdoNulo(rs![Alteração])
         sConciliado = SeCampoAdoNulo(rs![Conciliado])
         sCarteira = SeCampoAdoNulo(rs![Carteira])
         sSeqNossoNumero = SeCampoAdoNulo(rs![SeqNossoNumero])
         sPrevisao = SeCampoAdoNulo(rs![Previsao])
         sVlrMrD = SeCampoAdoNulo(rs![VlrMrD])
         sVlrDsp = SeCampoAdoNulo(rs![VlrDsp])
         sPerJur = SeCampoAdoNulo(rs![PerJur])
         sVlrJur = SeCampoAdoNulo(rs![VlrJur])
         sCODPRJ = SeCampoAdoNulo(rs![CODPRJ])
         sCODFPG = SeCampoAdoNulo(rs![CODFPG])
         sCheBan = SeCampoAdoNulo(rs![CheBan])
         sCheAge = SeCampoAdoNulo(rs![CheAge])
         sCheEmi = SeCampoAdoNulo(rs![CheEmi])
         sCheCco = SeCampoAdoNulo(rs![CheCco])
         sPerMul = SeCampoAdoNulo(rs![PerMul])
         sVlrMul = SeCampoAdoNulo(rs![VlrMul])
      CarregarRegistro = True
   Else
      CarregarRegistro = False
   End If
   rs.Close
   Set rs = Nothing
   Exit Function
ErroCarregarRegistro:
   Err.Raise 1 + vbObjectError, , "Não foi possível localizar pagrec!"
   rs.Close
   Set rs = Nothing
End Function

Private Sub Class_Initialize()
On Error GoTo ErroInicializar
   Set oCONEXAO = conexao
   Set rs = New Adodb.Recordset
   rs.Open "SELECT Count([PagRec]) as Registros FROM [Lançamentos]", oCONEXAO, adOpenKeyset, adLockPessimistic
   If rs!registros > 0 Then bVAZIO = False Else bVAZIO = True
   rs.Close
   Set rs = Nothing
   Exit Sub
ErroInicializar:
   Err.Raise 1 + vbObjectError, , " Erro ao inicializar classe lançamentos!"
End Sub

Private Sub Class_Terminate()
   Set oCONEXAO = Nothing
End Sub
