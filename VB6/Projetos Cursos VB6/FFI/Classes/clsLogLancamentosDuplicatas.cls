VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogLancamentosDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 100340
'---|   Ueder Budni (16/09/2016)
'---|---------------------------------------------------------------------------------------------------------------------------
Private Const NOME_TABELA = "LogLancamentosDuplicatas"

'Para que essa classe de Log verifique novas propriedades eh necessario modificar as constantes abaixo.
Private Const PROPERTIES_VO_LANC_DUP = "PagRec;Codigo_Nota;Parcela;Empresa;Tipo;Descricao;Emissao;Vencimento;Pagamento;Liberacao;ValorOriginal;Acrescimo;Abatimento;Banco;conta;Centro;Cheque;Moeda;ValordaMoeda;Controle;Situacao;Enviado;Marcacao;Obs;Bordero;DataIntegracaoFatoGerador;DataIntegracaoFatoPagamento;usuario;LINDIG;Conciliado;Carteira;previsao;VlrMrD;PerMrD;SeqNossoNumero;VlrDsP;PerJur;VlrJur;CODPRJ;CODFPG;CheBan;CheAge;CheEmi;CheCco;PerMul;VlrMul;AGECCE;NOSNUM;CodBar;LOCPAG;INSTRU;cd_operacao_contabil;cd_operacao_baixa;proveniente_rateio;parc_origem_baixa;Id_carteira;desc_motivo;cd_cobranca;seq_remessa;data_integracao;ValorTotal;LancDup;Col_Rateio;Remessa;Boleto"

Private mstrPagRec      As String
Private mdblNumero      As Double
Private mstrEmpresa     As String
Private mstrTipo        As String
Private mlngParcela     As Long
Private menuTabela      As enuLancDup

Private mstrErro        As String

Public Sub SetKey(strPagRec As String, dblNumero As Double, strEmpresa As String, strTipo As String, lngParcela As Long, enuTabela As enuLancDup)
    mstrPagRec = strPagRec
    mdblNumero = dblNumero
    mstrEmpresa = strEmpresa
    mstrTipo = strTipo
    mlngParcela = lngParcela
    menuTabela = enuTabela
End Sub

Public Property Get MsgErro() As String
    MsgErro = mstrErro
End Property

'|-------------------------------------------------------------------------------------------------------
'|   Função para inserir mensagem de log customizada de acordo com os parametros especificados. Exemplo:
'|   InsertCustomMsg("Alterado de {0} para {1}",1,2) retorna
'|      "Alterado de 1 para 2"
'|-------------------------------------------------------------------------------------------------------
Public Function InsertCustomMsg(ByVal strMsg As String, ParamArray varCampos() As Variant) As Boolean
    Dim i       As Integer

    For i = 0 To UBound(varCampos)
        strMsg = Replace(strMsg, "{" & i & "}", varCampos(i))
    Next i
    
    If InsertMsg(strMsg) Then
        InsertCustomMsg = True
    End If
End Function

Public Function InsertMsg(strMensagem As String) As Boolean
    Dim cmd        As IDBInsertCommand
    
On Error GoTo erro
    
    If ChaveDefinida Then
        Aplicacao.Connect
        Set cmd = Aplicacao.CreateInsertCommand
        With cmd
            .Table = NOME_TABELA
            
            Call .AddValue("[PagRec]", "@pPagRec")
            Call .Parameters.add(.CreateParameter("@pPagRec", mstrPagRec, dbFieldTypeString))
        
            Call .AddValue("[Numero]", "@pNumero")
            Call .Parameters.add(.CreateParameter("@pNumero", mdblNumero, dbFieldTypeDouble))
            
            Call .AddValue("[Empresa]", "@pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", mstrEmpresa, dbFieldTypeString))
            
            Call .AddValue("[Tipo]", "@pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", mstrTipo, dbFieldTypeString))
            
            Call .AddValue("[Parcela]", "@pParcela")
            Call .Parameters.add(.CreateParameter("@pParcela", mlngParcela, dbFieldTypeInt))
                                    
            Call .AddValue("[Descricao]", "@pDescr")
            Call .Parameters.add(.CreateParameter("@pDescr", strMensagem, dbFieldTypeString))
                                    
            Call .AddValue("[Tabela]", "@pTabela")
            Call .Parameters.add(.CreateParameter("@pTabela", menuTabela, dbFieldTypeInt))
            
            Call AdditionalInfo(cmd)
        End With
        
        InsertMsg = (Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd) = 1)
        Set cmd = Nothing
    End If
    Aplicacao.Disconnect
    
    Exit Function
erro:
    mstrErro = "Erro ao inserir Log: " & Err.Description
    If Aplicacao.isConnected Then
        Aplicacao.Disconnect
    End If
    InsertMsg = False
    Set cmd = Nothing
    Debug.Print mstrErro
    
End Function

Public Function CarregarLog() As Collection
    Dim rdResult            As IDBReader
    Dim cmd                 As IDBSelectCommand
    Dim strLinha            As String

On Error GoTo erro
    
    If ChaveDefinida Then
        mstrErro = Empty
        Aplicacao.Connect
        Set cmd = Aplicacao.CreateSelectCommand
        With cmd
            .Table.TableName = NOME_TABELA
            
            Call .Filter.Append("[PagRec] = @pPagRec")
            Call .Parameters.add(cmd.CreateParameter("@pPagRec", mstrPagRec, dbFieldTypeString))
               
            Call .Filter.Append("[Numero] = @pNumero")
            Call .Parameters.add(.CreateParameter("@pNumero", mdblNumero, dbFieldTypeDouble))
               
            Call .Filter.Append("[Parcela] = @pParcela")
            Call .Parameters.add(.CreateParameter("@pParcela", mlngParcela, dbFieldTypeInt))
            
            Call .Filter.Append("[Empresa] = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", mstrEmpresa, dbFieldTypeString))
            
            Call .Filter.Append("[Tipo] = @pTipo")
            Call .Parameters.add(.CreateParameter("@pTipo", mstrTipo, dbFieldTypeString))
            
            .OrderByClause = "dt_registro desc"
        End With
        
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
        With rdResult
            If Not .EOF Then
                Set CarregarLog = New Collection
                While Not .EOF
                    strLinha = .GetString("fox_user") & ";" & .GetDate("dt_registro") & ";" & .GetString("Descricao")
                    Call CarregarLog.add(strLinha)
                    .MoveNext
                Wend
                Set cmd = Nothing
            End If
        End With
        Set rdResult = Nothing
        Aplicacao.Disconnect
    End If
    Exit Function
erro:
    mstrErro = "Erro ao carregar Log: " & Err.Description
    If Aplicacao.isConnected Then
        Aplicacao.Disconnect
    End If
    Set CarregarLog = Nothing
    Set cmd = Nothing
    Set rdResult = Nothing
    Debug.Print mstrErro
End Function

Public Function InsertDiffObject(objOldState As Object, objNewState As Object, Optional strRotina As String = Empty) As Boolean
    Dim strProperty         As Variant
    Dim strArrProp()        As String
    Dim strStdMsg           As String
    Dim strStdMsgBln        As String
    
On Error GoTo erro
    mstrErro = Empty
    
    strStdMsg = "Alterado campo {0} de '{1}' para '{2}'."
    strStdMsgBln = "Campo {0} foi {1}."
    If strRotina <> Empty Then
        strStdMsg = strStdMsg & " Alterado por: " & strRotina
        strStdMsgBln = strStdMsgBln & " Alterado por: " & strRotina
    End If
    
    If Not objOldState Is Nothing And Not objNewState Is Nothing Then
        If TypeName(objOldState) = TypeName(objNewState) Then
            Select Case TypeName(objOldState)
                Case "VoLancamentoDuplicata": strArrProp = Split(PROPERTIES_VO_LANC_DUP, ";")
            End Select
            
            InsertDiffObject = True
            For Each strProperty In strArrProp
                If CallByName(objOldState, strProperty, VbGet) <> CallByName(objNewState, strProperty, VbGet) And InsertDiffObject = True Then
                    If TypeName(CallByName(objOldState, strProperty, VbGet)) = "Boolean" Then
                        InsertDiffObject = InsertDiffObject And InsertCustomMsg(strStdMsgBln, NomeAmigavel(CStr(strProperty)), IIf(CallByName(objNewState, strProperty, VbGet) = True, "marcado", "desmarcado"))
                    Else
                        InsertDiffObject = InsertDiffObject And InsertCustomMsg(strStdMsg, NomeAmigavel(CStr(strProperty)), CallByName(objOldState, strProperty, VbGet), CallByName(objNewState, strProperty, VbGet))
                    End If
                End If
            Next
        End If
    End If
    
    Exit Function
erro:
    mstrErro = "Erro no método 'InsertDiffObject': " & Err.Description
    InsertDiffObject = False
    Debug.Print mstrErro
End Function

Private Function ChaveDefinida() As Boolean
    ChaveDefinida = (mstrPagRec <> Empty) And (mdblNumero <> 0) And (mstrEmpresa <> Empty) And (mstrTipo <> Empty) And (mlngParcela <> 0)
End Function

Public Sub AdditionalInfo(ByRef cmd As Object)
On Error GoTo fim
    With cmd
        
        Call .AddValue("[dt_registro]", "@pDHReg")
        Call .Parameters.add(.CreateParameter("@pDHReg", Now, dbFieldTypeDateTime))
        
        Call .AddValue("[nome_maquina]", "@pNomeMaq")
        Call .Parameters.add(.CreateParameter("@pNomeMaq", Environ("COMPUTERNAME"), dbFieldTypeString))
        
        Call .AddValue("[windows_user]", "@pWinUser")
        Call .Parameters.add(.CreateParameter("@pWinUser", Environ("USERNAME"), dbFieldTypeString))
        
        Call .AddValue("[fox_user]", "@pFoxUser")
        Call .Parameters.add(.CreateParameter("@pFoxUser", fMain.stbMain.Panels(2).Text, dbFieldTypeString))
    End With
fim:
End Sub

Private Function NomeAmigavel(strPropName As String) As String
    Select Case strPropName
        Case "Codigo_Nota": NomeAmigavel = "Codigo/Nota"
        Case "Descricao": NomeAmigavel = "Descrição"
        Case "Emissao": NomeAmigavel = "Emissão"
        Case "Liberacao": NomeAmigavel = "Liberação"
        Case "ValorOriginal": NomeAmigavel = "Valor Original"
        Case "Acrescimo": NomeAmigavel = "Acréscimo"
        Case "Centro": NomeAmigavel = "C. Custo"
        Case "Obs": NomeAmigavel = "Observação"
        Case "LINDIG": NomeAmigavel = "Linha Digitável - Outros"
        Case "SeqNossoNumero": NomeAmigavel = "Nr. Sequencial"
        Case "CODFPG": NomeAmigavel = "Forma Pagto"
        Case "CheBan": NomeAmigavel = "Banco - Outros"
        Case "CheAge": NomeAmigavel = "Agência - Outros"
        Case "CheEmi": NomeAmigavel = "Correntista - Outros"
        Case "CheCco": NomeAmigavel = "Conta Corrente - Outros"
        Case "VlrMul": NomeAmigavel = "Vlr. Multa"
        Case "VlrMrD": NomeAmigavel = "Vlr. Mora Diária"
        Case "PerMul": NomeAmigavel = "Perc. Multa"
        Case "PerMrD": NomeAmigavel = "Perc. Mora"
        Case "VlrDsP": NomeAmigavel = "Vlr. Desc Pontualidade"
        Case "NOSNUM": NomeAmigavel = "Nosso Número"
        Case "cd_operacao_contabil": NomeAmigavel = "Op. Contábil"
        Case "cd_operacao_baixa": NomeAmigavel = "Op. Contábil - Baixa"
        Case "Id_carteira": NomeAmigavel = "Carteira"
        Case "Situacao": NomeAmigavel = "Situação"
        Case "cd_cobranca": NomeAmigavel = "Código do Endereço de Cobrança"
        Case "Conta": NomeAmigavel = "Conta"
        Case "usuario": NomeAmigavel = "Usuário"
        Case "desc_motivo": NomeAmigavel = "Status - Remessa Bancária"
        Case "CodBar": NomeAmigavel = "Linha Digitável"
        Case "ValorTotal": NomeAmigavel = "Valor Total"
        Case Else: NomeAmigavel = strPropName
    End Select
End Function
