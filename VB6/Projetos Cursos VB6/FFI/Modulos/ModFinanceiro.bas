Attribute VB_Name = "modFinanceiro"
Option Explicit

Enum FIN_TIP_DESC_PONTUALIDADE
    tdpCadCliente = 1
    tdpVlrFixo = 2
    tdpPercentual = 3
End Enum

'Projeto: #1203 - Hist�ria: #10564 - Desenvolvimento#10570 - Moacir Pfau(16/03/2012)
Enum enuPagRec
    Pagamento = 0
    Recebimento = 1
End Enum

Enum enuBancos
    BANCODOBRASIL = 1
    BANCOSANTANDER = 33
    BANCOBMG = 318
    BANCOITAU = 341
    BANCOABNAMROREAL = 356
    BANCOMERCANTILDOBRASIL = 389
    BANCOBMC = 394
    HSBCBANKBRASILBANCOMULTIPLO = 399
    UNIBANCO = 409
    BANCOCAPITAL = 412
    BANCOSAFRA = 422
    BANCORURAL = 453
End Enum

'CONVERTE PARA MOEDA BASE
'Tenta converter uma moeda baseado numa cota��o da moeda base.
'Quando identificada uma cota��o at� a data de pagamento ou emiss�o retorna ZERO
Public Function ConvMoedaBase(nValor As Currency, MoedaDoc As String, DtEmi As Date, Optional MoedaBase As String = "", Optional DtPag As Date = Empty) As Currency
    Dim dblCotMoedaDoc  As Currency   'Cota��o da moeda do documento
    Dim dblCotMoedaBase As Currency   'Cota��o da moeda base
       
    If MoedaDoc <> MoedaBase Then
        If Not IsNull(DtPag) And (DtPag <> Empty) Then
           dblCotMoedaDoc = UltimaCotacao(MoedaDoc, DtPag)
           dblCotMoedaBase = UltimaCotacao(MoedaBase, DtPag)
        Else
           dblCotMoedaDoc = UltimaCotacao(MoedaDoc, DtEmi)
           dblCotMoedaBase = UltimaCotacao(MoedaBase, DtEmi)
        End If
        If dblCotMoedaBase = 0 Then
           'Quando n�o houver cota��o, para evitar erro de divisao por zero
           nValor = 0
        Else
           'Essa c�lculo garante a convers�o entre moedas: Exemplo Euro para Dolar ou Dolar para Peso, etc
           'pois converte primeiro o valor para reais e depois converte para a moeda base
           nValor = nValor * dblCotMoedaDoc / dblCotMoedaBase
        End If
    End If
    ConvMoedaBase = nValor
End Function


Public Function CarregaDuplicatas(pPagRec As String, _
    Optional pTipReg As String, _
    Optional pEmp As String = Empty, _
    Optional pNroIni As Long = 0, _
    Optional pNroFim As Long = 0, _
    Optional pEmiIni As Date = 0, _
    Optional pEmiFim As Date = 0, _
    Optional pVctIni As Date = 0, _
    Optional pVctFim As Date = 0, _
    Optional pPgtIni As Date = 0, _
    Optional pPgtFim As Date = 0, _
    Optional pBan As Long = 0, _
    Optional pSomenteNaoPagas As Boolean = False) As ADODB.Recordset
    
Dim s As String

s = "SELECT * FROM DUPLICATAS"
s = s + " WHERE PagRec = " & Quote(pPagRec, "'")


If pTipReg <> Empty Then
    s = s + " AND Tipo = " & Quote(pTipReg, "'")
End If

If pEmp <> Empty Then
    s = s + " AND Empresa = " & Quote(pEmp, "'")
End If

If pNroIni <> 0 Then
    s = s + " AND Nota >= " & str(pNroIni)
End If

If pNroFim <> 0 Then
    s = s + " AND Nota <= " & str(pNroFim)
End If

If pEmiIni <> 0 Then
    s = s + " AND [Emiss�o] >= " & InverteData(pEmiIni, True)
End If

If pEmiFim <> 0 Then
    s = s + " AND [Emiss�o] <= " & InverteData(pEmiFim, True)
End If

If pVctIni <> 0 Then
    s = s + " AND [Vencimento] >= " & InverteData(pVctIni, True)
End If

If pVctFim <> 0 Then
    s = s + " AND [Vencimento] <= " & InverteData(pVctFim, True)
End If

If pPgtIni <> 0 Then
    s = s + " AND [Pagamento] >= " & InverteData(pPgtIni, True)
End If

If pPgtFim <> 0 Then
    s = s + " AND [Pagamento] <= " & InverteData(pPgtFim, True)
End If

If pBan <> 0 Then
    s = s + " AND [Banco] = " & str(pBan)
End If

If pSomenteNaoPagas Then
    s = s + " AND Pagamento IS NULL"
End If

Set CarregaDuplicatas = conexao.Execute(s)

End Function


Public Function CarregaLancamentos(pPagRec As String, _
    Optional pTipReg As String, _
    Optional pEmp As String = Empty, _
    Optional pNroIni As Long = 0, _
    Optional pNroFim As Long = 0, _
    Optional pEmiIni As Date = 0, _
    Optional pEmiFim As Date = 0, _
    Optional pVctIni As Date = 0, _
    Optional pVctFim As Date = 0, _
    Optional pPgtIni As Date = 0, _
    Optional pPgtFim As Date = 0, _
    Optional pBan As Long = 0, _
    Optional pSomenteNaoPagas As Boolean = False) As ADODB.Recordset
    
Dim s As String

s = "SELECT * FROM [LAN�AMENTOS]"
s = s + " WHERE PagRec = " & Quote(pPagRec, "'")


If pTipReg <> Empty Then
    s = s + " AND Tipo = " & Quote(pTipReg, "'")
End If

If pEmp <> Empty Then
    s = s + " AND Empresa = " & Quote(pEmp, "'")
End If

If pNroIni <> 0 Then
    s = s + " AND [C�digo] >= " & str(pNroIni)
End If

If pNroFim <> 0 Then
    s = s + " AND [C�digo] <= " & str(pNroFim)
End If

If pEmiIni <> 0 Then
    s = s + " AND [Emiss�o] >= " & InverteData(pEmiIni, True)
End If

If pEmiFim <> 0 Then
    s = s + " AND [Emiss�o] <= " & InverteData(pEmiFim, True)
End If

If pVctIni <> 0 Then
    s = s + " AND [Vencimento] >= " & InverteData(pVctIni, True)
End If

If pVctFim <> 0 Then
    s = s + " AND [Vencimento] <= " & InverteData(pVctFim, True)
End If

If pPgtIni <> 0 Then
    s = s + " AND [Pagamento] >= " & InverteData(pPgtIni, True)
End If

If pPgtFim <> 0 Then
    s = s + " AND [Pagamento] <= " & InverteData(pPgtFim, True)
End If

If pBan <> 0 Then
    s = s + " AND [Banco] = " & str(pBan)
End If

If pSomenteNaoPagas Then
    s = s + " AND Pagamento IS NULL"
End If

Set CarregaLancamentos = conexao.Execute(s)

End Function


Public Sub AtualizaDescPorPontualidadeDuplicata(pPagRec As String, _
            pTipReg As String, _
            pEmp As String, _
            pNro As Long, _
            pParcela As Integer, _
            pTipoDesc As FIN_TIP_DESC_PONTUALIDADE, _
            pVlrDesc As Double)

    Dim percDescPontCli As Double
    Dim vlrDescPont As Double
    Dim s As String
    Dim sw As String

    sw = " PagRec = " & Quote(pPagRec, "'") & _
        " AND Tipo = " & Quote(pTipReg, "'") & _
        " AND Empresa = " & Quote(pEmp, "'") & _
        " AND [Nota] = " & str(pNro) & _
        " AND Parcela = " & str(pParcela)
        
    If pTipoDesc = tdpCadCliente Then
        'percentual de desconto
        percDescPontCli = GetFieldValue("[Desconto por Pontualidade]", "Empresas", "Apel = " & Quote(pEmp, "'"), default:=0)
        'valor da duplicata
        vlrDescPont = GetFieldValue("[Valor Original]", "Duplicatas", sw, default:=0)
        'valor do desconto
        vlrDescPont = Round(CDec((percDescPontCli / 100) * vlrDescPont), 2)
    
    ElseIf pTipoDesc = tdpVlrFixo Then
        vlrDescPont = pVlrDesc
    
    Else
        vlrDescPont = GetFieldValue("[Valor Original]", "Duplicatas", sw, default:=0)
        vlrDescPont = Round(CDec((pVlrDesc / 100) * vlrDescPont), 2)
    End If
    
    s = "UPDATE DUPLICATAS SET VlrDsP = " & str(vlrDescPont) & _
        " WHERE " & sw
        
    conexao.Execute s
End Sub


Public Sub AtualizaDescPorPontualidadeLancamentos(pPagRec As String, _
            pNro As Long, _
            pTipoDesc As FIN_TIP_DESC_PONTUALIDADE, _
            pVlrDesc As Double)

    Dim percDescPontCli As Double
    Dim vlrDescPont As Double
    Dim sEmp As String
    Dim s As String
    Dim sw As String

    sw = " PagRec = " & Quote(pPagRec, "'") & _
        " AND [C�digo] = " & str(pNro)

    If pTipoDesc = tdpCadCliente Then
        sEmp = GetFieldValue("Empresa", "Lan�amentos", sw, default:=0)
        'percentual de desconto
        percDescPontCli = GetFieldValue("[Desconto por Pontualidade]", "Empresas", "Apel = " & Quote(sEmp, "'"), default:=0)
        'valor do lan�amento
        vlrDescPont = GetFieldValue("[Valor Original]", "Lan�amentos", sw, default:=0)
        'valor do desconto
        vlrDescPont = Round(CDec((percDescPontCli / 100) * vlrDescPont), 2)
        
    ElseIf pTipoDesc = tdpVlrFixo Then
        vlrDescPont = pVlrDesc
        
    Else
        vlrDescPont = GetFieldValue("[Valor Original]", "Lan�amentos", sw, default:=0)
        vlrDescPont = Round(CDec((pVlrDesc / 100) * vlrDescPont), 2)
    End If
    
    s = "UPDATE LAN�AMENTOS SET VlrDsP = " & str(vlrDescPont) & _
        " WHERE " & sw
        
    conexao.Execute s
End Sub

' FUNCTION..: KDecimais
' Objetivo..: Diminui as casas decimais de um n�mero sem arredondamento
' Argumentos: [vNumero]: N�mero.
'             [lDec   ]: N�mero de casas decimais
' Retorna...: UM Double com o valor do n�mero.
' ----------------------------------------------------------
Public Function KDecimais(vNumero, Optional ByVal lDec As Long) As Double
Dim dReturn As Double           '// Valor retornado

  If (lDec) Then                '// Se for maior que zero
    lDec = 10 ^ lDec
    dReturn = Fix(vNumero * lDec)
    dReturn = dReturn / lDec
  Else
    dReturn = Fix(vNumero)
  End If
  KDecimais = dReturn

End Function

Public Function CMoedaFormatoAmericano(Valor As String) As Currency
    #If FOXSQL = 1 Then
        Valor = Replace(Valor, ",", "")
        Valor = Replace(Valor, ".", ",")
        CMoedaFormatoAmericano = CCurDef(Valor)
    #Else
        CMoedaFormatoAmericano = CMoeda(Valor)
    #End If
End Function

' FUNCTION..: LoadResCursor
' Objetivo..: Carrega um cursor do arquivo de recursos do aplicativo.
' Argumento.: [CursorID]: �ndice do recurso.
' Retorna...: Um objeto Picture.
' ------------------------------------------------------------
Public Function LoadResCursor(ByVal CursorID As Integer) As Picture
  On Error Resume Next
  Set LoadResCursor = LoadResPicture(CursorID, vbResCursor)
  If (err().Number) Then
    MsgErro NUL
  End If
End Function


' FUNCTION..: WaitWindowClose
' Objetivo..: Cria um loop aguardando at� que uma determinada janela seja fechada
' Argumentos: [hWnd]: Handle da janela.
' Retorno...: A fun��o retorna True quando consegue mapear a janela solicitada
'             corretamente, False se algum erro for retornado da API do Windows.
' ---------------------------------------------------------------------------------
Public Function WaitWindowClose(hWnd As Long) As Boolean

  While (IsWindow(hWnd))
    DoEvents
  Wend

  WaitWindowClose = True

End Function

' FUNCTION: Kif_Valor
' Calcula o valor total da duplicata ou lan�amento.
' Argumento: [rstKif]: Recordset.
' Retorna  : Um valor currency com o total encontrado.
' --------------------------------------------------------------------
Public Function Kif_Valor(rstKif As Object) As Currency
Dim curValor As Currency

  If (Not EstaVazio(rstKif)) Then
    curValor = GetValue(rstKif, "Valor Original")
    curValor = curValor + GetValue(rstKif, "Acr�scimo", ZERO)
    curValor = curValor - GetValue(rstKif, "Abatimento", ZERO)
    Kif_Valor = curValor
  Else
    Kif_Valor = ZERO
  End If
End Function


' FUNCTION..: DataLimiteCentroCusto
' Objetivo..: Verifica se a data em quest�o passa do limite do centro de custo
' Argumentos: [lngcentrocusto]: Centro de Custo
'             [strData]: Data a ser conferida
' Retorna...: True se passou da Data Limite, False se n�o.
' ------------------------------------------------------------------------------
Public Function DataLimiteCentroCusto(lngCentroCusto As Long, strData As String) As Boolean
  Dim DataLimite    As String
  
  DataLimiteCentroCusto = False
  
  DataLimite = GetFieldValue("[Data Limite]", "[Centros]", "C�digo = " & lngCentroCusto, , strData)
  
  If EData(DataLimite) Then
    If CDateDef(DataLimite) < CDateDef(strData) Then
      MsgBox "A Data " & strData & " ultrapassa a Data Limite para Movimenta��o do Centro de Custo.", vbInformation, MsgBoxCaption
      DataLimiteCentroCusto = True
    End If
  End If
  
End Function

#If ESP = 0 Then
' FUNCTION..: fMemo
' Objetivo..: Exibe um formul�rio com o campo Memo de alguma tabela.
' Argumentos: [strTitulo  ]: T�tulo da janela.
'             [strTabela  ]: Nome da tabela onde se encontra o campo.
'             [strCampo   ]: Nome do campo tipo memo.
'             [strClausula]: Cl�usula Where das compara��es, sem a palavra chave
'                            WHERE.
' Retorna...: True se houver algum dado no campo, False se n�o. A fun��o retorna
'             imediatamente, isto �, n�o aguarda at� que a janela seja fechada.
' ------------------------------------------------------------------------------
Public Function fMemo(strTitulo As String, strTabela As String, strCampo As String, strClausula As String) As Boolean
    Dim rstDados As Object
    Dim strMemo  As String
    
    fMemo = False
    If (Len(strClausula)) Then
        strMemo = "SELECT " & strCampo & " FROM " & strTabela & " WHERE " & strClausula & ";"
    Else
        strMemo = "SELECT " & strCampo & " FROM " & strTabela & ";"
    End If
    
    If (WL_OK = AbreRecordset(rstDados, strMemo, dbOpenSnapshot)) Then
        FechaRecordset rstDados
        
        Load frmObsEmp
        frmObsEmp.Caption = strTitulo
        frmObsEmp.InstSelect = strMemo
        CenterForm frmObsEmp
        frmObsEmp.Show vbModal
        
        fMemo = True
    Else
        MsgBox LoadResString(IDS_NORECORD)
    End If
End Function
#End If

' FUNCTION..: ConfDataCheque
' Objetivo..: Conferir o cheque indicado nos cadastros de Duplicata, Lan�amento e
'             Transfer�ncia Banc�ria, verificando se n�o h� datas diferentes para
'             um mesmo cheque.
' Argumentos: [strCodBco]: C�digo do Banco.
'             [strNumChq]: N�mero do Cheque.
'             [strData  ]: Data indicada pelo usu�rio.
'             [uControle]: Vari�vel de controle de a��es do usu�rio.
' Retorna...: True se n�o houver datas diferetes para este cheque. False se n�o.
' --------------------------------------------------------------------------------
Public Function ConfDataCheque(strCodBco As String, strNumChq As String, strData As String, uControle As Long) As Boolean
    Dim lBanco      As Long         '// C�digo do Banco
    Dim lCheque     As Long         '// N�mero do Cheque
    Dim dData       As Date         '// Data que deve ser conferida para este cheque
    Dim strConsulta As String
    Dim rstConsulta As Object

On Error GoTo ConfDataCheque_Erro

    lBanco = CLngDef(strCodBco)
    lCheque = CLngDef(strNumChq)
    dData = CDateDef(strData)
    
    If ((Not CBool(lBanco)) Or (Not CBool(lCheque)) Or (IsEmptyDate(dData))) Then
        ConfDataCheque = True: Exit Function
    End If

    'A instru��o a seguir obt�m todas as datas do cheque passado
    strConsulta = "SELECT Pagamento AS Data FROM Lan�amentos WHERE Banco = %l AND Cheque = %l " & _
                  "UNION ALL SELECT Pagamento AS Data FROM Duplicatas WHERE Banco = %l AND Cheque = %l " & _
                  "UNION ALL SELECT Data FROM [Transf Banc�ria] WHERE Origem = %l AND Cheque = %l;"

    '// Completa as lacunas da instru��o
    wvsprintf strConsulta, strConsulta, lBanco, lCheque, lBanco, lCheque, lBanco, lCheque

    If (AbreRecordset(rstConsulta, strConsulta, dbOpenSnapshot) = WL_OK) Then
        If (EEdicao(uControle)) Then
        
            '// Se o usu�rio estiver alterando o registro atual devo verificar se
            '// a data informada neste registro n�o � diferente da que j� est� na
            '// consulta, por�m, somente se a consulta contiver mais de um registro
            '// porque se somente um registro foi retornado, � o mesmo que o usu�rio
            '// est� alterando.
            If (rstConsulta.Recordcount > UM) Then
                If (DateDiff(DD_DIA, dData, GetValue(rstConsulta, 0, Date))) Then
                    GoTo ConfDataCheque_ShowRecords
                End If
            End If
        
        ElseIf (EAdicao(uControle)) Then
        
            '// Se o usu�rio estiver adicionando, a data informada n�o pode ser diferente
            '// na encontrada na consulta, mesmo que houver apenas um registro.
            If (DateDiff(DD_DIA, dData, GetValue(rstConsulta, 0, Date))) Then
                GoTo ConfDataCheque_ShowRecords
            End If
        End If
    End If
    FechaRecordset rstConsulta
    ConfDataCheque = True
    Exit Function

ConfDataCheque_ShowRecords:
    If (MsgFunc(LoadResString(142) & vbCrLf & LoadResString(144), _
                vbQuestion Or vbYesNo) = vbYes) Then
        strConsulta = "SELECT Nota As C�digo, 'Duplicata' As Tipo, Pagamento AS Data, " & _
                  "([Valor Original] + Acr�scimo - Abatimento) As Valor " & _
                  "FROM Duplicatas WHERE Banco = %l AND Cheque = %l UNION ALL " & _
                  "SELECT C�digo, 'Lan�amento' AS Tipo, Pagamento As Data, " & _
                  "([Valor Original] + Acr�scimo - Abatimento) As Valor " & _
                  "FROM Lan�amentos WHERE Banco = %l AND Cheque = %l UNION ALL " & _
                  "SELECT C�digo, 'Transfer�ncia' AS Tipo, Data, Valor " & _
                  "FROM [Transf Banc�ria] WHERE Origem = %l AND Cheque = %l;"

        wvsprintf strConsulta, strConsulta, lBanco, lCheque, lBanco, lCheque, lBanco, lCheque
        PRegistro Nothing, Nothing, "Cheques", NUL, strConsulta, NUL, WL_USEREDITNONE
    End If
    FechaRecordset rstConsulta

ConfDataCheque_Erro:
    If (err().Number) Then
        #If (DebugInfo) Then
        MsgErro wsprintf("Erro: %l\n%s\nConfDataCheque", err.Number, err.Description)
        #Else
        DAOErros NUL
        #End If
    End If
    FechaRecordset rstConsulta
End Function


' FUNCTION..: ExisteCheque
' Objetivo..: Verifica se existe algum lan�amento com um determinado cheque
'             nos cadastros de Duplicatas, Lan�amentos e Transfer�ncias.
' Argumentos: [nBanco ]: C�digo do Banco.
'             [nCheque]: N�mero do Cheque.
' Retorna...: O n�mero de registro em que o referido cheque aparece.
' ------------------------------------------------------------------------------------
Public Function ExisteCheque(nBanco As Long, nCheque As Long) As Long
    Dim strDupl As String              '// Express�o de consulta
    Dim strLanc As String
    Dim StrTran As String
    Dim qdfTmp  As QueryDef            '// Objeto QueryDef tempor�rio
    Dim rsCount As Object           '// Resultado da conta

    strDupl = wsprintf("SELECT Banco FROM Duplicatas WHERE Banco = %l AND Cheque = %l", nBanco, nCheque)
    strLanc = wsprintf("SELECT Banco FROM Lan�amentos WHERE Banco = %l AND Cheque = %l", nBanco, nCheque)
    StrTran = wsprintf("SELECT Origem FROM [Transf Banc�ria] WHERE Origem = %l AND Cheque = %l", nBanco, nCheque)
    
    strDupl = wsprintf("%s UNION ALL %s UNION ALL %s", strDupl, strLanc, StrTran)

    If (AbreRecordset(rsCount, strDupl, dbOpenSnapshot) = WL_OK) Then
        If Not rsCount.EOF Then
            ExisteCheque = GetValue(rsCount, ZERO, ZERO)
        End If
    End If
    FechaRecordset rsCount
End Function

' FUNCTION..: ListViewAddItem
' Objetivo..: Carrega um controle ListView com os dados de uma tabela base ou
'             consulta SQL.
' Argumentos: [Controle]: Refer�ncia ao controle ListView.
'             [Origem  ]: Nome da tabela ou consulta para obter os dados.
'             [Icone   ]: �ndice ou Chave do �cone que ser� exibido.
' Retorna...: O n�mero de registros adicionados.
' ------------------------------------------------------------------------------
Public Function ListViewAddItem(Controle As Object, Origem As String, Optional icone) As Long
    Dim rstLvw   As Object
    Dim lngItems As Long                  'Conta os �tens do controle ListView
    Dim iCampos  As Integer               'N�mero de campos acrescentados ao controle
    Dim vValue As Variant
    '
    ' Abrindo a tabela ou consulta
    '
    If AbreRecordset(rstLvw, Origem, dbOpenForwardOnly) <> WL_ERRO Then
        lngItems = Controle.ListItems.Count
        Do Until rstLvw.EOF
            Inc lngItems
            If (IsMissing(icone)) Then
                vValue = GetValue(rstLvw, 0, NUL)
                Controle.ListItems.add lngItems, , vValue
            Else
                vValue = GetValue(rstLvw, 0, NUL)
                Controle.ListItems.add lngItems, , vValue, , icone
            End If
            For iCampos = 1 To rstLvw.Fields.Count - 1
                vValue = GetValue(rstLvw, iCampos, NUL)
                Controle.ListItems(lngItems).SubItems(iCampos) = vValue
            Next
            rstLvw.MoveNext
        Loop
        ListViewAddItem = rstLvw.Recordcount
    End If
    FechaRecordset rstLvw
End Function

' FUNCTION..: CompDatas
' Objetivo..: Compara duas datas, digitadas pelo usu�rio, e verifica se est�o
'             corretas.
' Argumentos: [Controle1]: Refer�ncia ao controle que tem a primeira data.
'             [Controle2]: Refer�ncia ao controle que tem a segunda  data.
'             [Data1]    : Refer�ncia a vari�vel que ir� receber a primeira data.
'             [Data2]    : Refer�ncia a vari�vel que ir� receber a segunda data.
' Retorna...: True se o usu�rio n�o digitou uma data erra. Caso contr�rio False.
' ---------------------------------------------------------------------------------
Public Function CompDatas(Controle1 As Object, Controle2 As Object, Data1 As Date, Data2 As Date) As Boolean
    Const TEXT_CONTROL$ = "TextBox ComboBox"
    
    If InStr(1, TEXT_CONTROL, TypeName(Controle1)) And _
       InStr(1, TEXT_CONTROL, TypeName(Controle2)) Then
    
        If (Len(Controle1.Text) = 0) And (Len(Controle2.Text) = 0) Then
            Exit Function
        End If
        
        If Len(Controle1.Text) Then
            If EData(Controle1.Text) Then
                Data1 = CDate(Controle1.Text)
            Else
                MsgBox ResolveResString(26, resUM, Controle1.Text), vbInformation, MsgBoxCaption
                Exit Function
            End If
        Else
            Data1 = CDate(#1/1/1000#)
        End If
        
        If Len(Controle2.Text) Then
            If EData(Controle2.Text) Then
                Data2 = CDate(Controle2.Text)
            Else
                MsgBox ResolveResString(26, resUM, Controle2.Text), vbInformation, MsgBoxCaption
                Exit Function
            End If
        Else
            Data2 = Date
        End If
        CompDatas = True
    End If

End Function

' FUNCTION..: SaldoInicialGeral
' Objetivo..: Calcula o saldo geral da empresa em uma determinada data.
' Argumentos: [dtSaldo  ]: Data em que se precisa do saldo inicial.
'             [cSaldo   ]: Vari�vel que ir� receber o saldo.
'             [bPrevisao]: Se a fun��o deve considerar o campo Previs�o dos bancos
'             [sConciliado]: Optional Sim, N�o ou Todos. Padr�o Todos
' Retorna...: WL_OK se obtiver sucesso, WL_ERRO se algum erro impedir a
'             fun��o de terminar o c�lculo. WL_CANCEL se o usu�rio precionar
'             a tecla ESC antes da fun��o terminar.
' Nota......: A fun��o calcula o saldo somando todos os saldos dos bancos
'             cadastrados na empresa. Note que, se o banco n�o possuir o saldo
'             no cadastro de Saldos a fun��o tentar� calcular todos os lan�amentos
'             feitos com este banco desde o in�cio da utiliza��o do Sistema.
' ---------------------------------------------------------------------------------
Public Function SaldoInicialGeral(dtSaldo As Date, cSaldo As Currency, bPrevisao As Boolean, Optional strMoeda As String, Optional StrDescMoeda As String, Optional sConciliado As String = "Sim") As Long
    Dim rstBancos As Object        '// Seleciona todos os Bancos cadastrados
    Dim lResult   As Long
    Dim lBanco    As Long             '// C�digo do Banco
    Dim strBancos As String
    Dim strConciliar As String
  
    Select Case sConciliado
        Case "Todos"
            strConciliar = ""
        Case "Sim"
            strConciliar = " AND Conciliado = TRUE "
        Case "N�o"
            strConciliar = " AND Conciliado = FALSE "
    End Select
  
    Call InKey(vbKeyEscape)         '// Limpa o buffer anterior
    
    If (bPrevisao) Then
        strBancos = "SELECT Banco, Nome FROM Bancos WHERE Previs�o = True;"
    Else
        strBancos = "SELECT Banco, Nome FROM Bancos;"
    End If

    lResult = WL_OK
    If (AbreRecordset(rstBancos, strBancos, dbOpenForwardOnly) = WL_OK) Then
        Do
            If (InKey(vbKeyEscape)) Then
                lResult = WL_CANCEL
                GoTo SaldoInicialGeral_Erro
            End If
            
            lBanco = GetValue(rstBancos, "Banco", ZERO)
            SimpleMsgBar ResolveResString(1023, resUM, CStr(lBanco), resDOIS, _
                                          GetValue(rstBancos, "Nome", NUL), resTRES, DataToStr(dtSaldo))
            cSaldo = cSaldo + SaldoInicial(GetValue(rstBancos, "Banco", ZERO), dtSaldo, strMoeda:=strMoeda, StrDescMoeda:=StrDescMoeda, sConciliado:=sConciliado)
            rstBancos.MoveNext
        
        Loop Until (rstBancos.EOF)
    End If

SaldoInicialGeral_Erro:
    If (err().Number) Then
        DAOErros NUL
        lResult = WL_ERRO
    End If
    FechaRecordset rstBancos
    SaldoInicialGeral = lResult
End Function

' FUNCTION..: DataLongaExt
' Objetivo..: Cria a Data Longa para recibos, cheques e outros.
' Argumentos: [Data]: Qualquer express�o de data v�lida.
' -----------------------------------------------------------------------------
Public Function DataLongaExt(Data) As String
    If (Not IsEmpty(Data)) Then
        Dim strResult As String
        If (EData(Data)) Then
            strResult = CidadePadrao() & ", " & StrZero(Day(Data), 2) & _
                  " de " & MesExt(Data) & " de " & Format$(Data, "yyyy")
        End If
    End If
    DataLongaExt = strResult
End Function

' FUNCTION..: ConfereDuplicidade
' Objetivo..: Verifica atrav�s de uma express�o SQL se h� duplicidade de
'             �ndices no Recordset. S� deve ser utilizada quando da inclus�o
'             de dados.
' Argumentos: [Campos ]: String com os campos que devem ser verificados. No
'                        caso de serem v�rios campos separar os nomes com
'                        v�rgulas. Campos com espa�os devem estar entre
'                        colchetes.
'            [Tabela  ]: Nome da Tabela que deve ser verificada.
'            [Clausula]: Parte de cl�usula WHERE de uma consulta sem a
'                        palavra WHERE. Explicitamente as compara��es que
'                        devem ser feitas.
' Retorna...: 0 se n�o houver duplicidade, ou o n�mero de registros encontrados.
' ------------------------------------------------------------------------------
Public Function ConfereDuplicidade(Campos As String, Tabela As String, Clausula As String) As Long
'#EDU-11/03/02#
'Altera�oes efetuadas, debugar...

Dim strExpressaoSQL As String
Dim rstConfere As Object

  ' A express�o foi guardada no arquivo de recurso
  '
  strExpressaoSQL = "SELECT " & Campos & " FROM [" & Tabela & "]" & " WHERE " & Clausula & ";"

  On Error GoTo ErroConfere
  
'-----------------------------------------------------------------------------------------------------------
  #If FOXSQL Then
  AbreRecordset rstConfere, strExpressaoSQL, dbOpenSnapshot
  #Else
  If gTipoDB = Access Then
    Set rstConfere = mdbDatabase.OpenRecordset(strExpressaoSQL, dbOpenSnapshot)
  Else
    Set rstConfere = CreateObject("ADODB.Recordset")
    rstConfere.Open strExpressaoSQL, mdbDatabase, adOpenDynamic
  End If
  #End If
'-----------------------------------------------------------------------------------------------------------
  If EstaVazio(rstConfere) Then
    '
    ' Se estiver vazia retorna zero
    ConfereDuplicidade = 0
  Else
    '
    ' Retorna o n�mero de registros encontrados
    rstConfere.MoveLast
    ConfereDuplicidade = rstConfere.Recordcount
  End If

  rstConfere.Close

ErroConfere:

  If err.Number <> 0 Then
    DAOErros vbNullString
    ConfereDuplicidade = 0
  End If

  Set rstConfere = Nothing

End Function

'Davi Brito - #169397 - 22/05/2017
Public Function AbrirAlertaContasVencidas()
    Dim rstMod As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    If (AbreRecordset(rstMod, "Configura��o") = WL_OK) Then
      If (GetValue(rstMod, "exibir_titulos_ja_vencidos", NUL)) Or (GetValue(rstMod, "Contas a Pagar", False)) Then
        If VerContasVencindas(rstMod, rs) Then
            Load frmAlerta
            Call frmAlerta.CarregarGrid(rs)
            frmAlerta.Show
        End If
      End If
    End If
End Function
