VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Autor.......: Gustavo Cuman
'Data........: 08/12/2008
'Descrição...: Classe representando um registro tipo Lancamento e Observacao do arquivo de importação Duplicata.
Private mstrPagarReceber As String
Private mlngNumDuplicata As Long
Private mstrTipoGlobal As String
Private mintParcela As Integer
Private mstrEmpresa As String
Private mstrDescricao As String
Private mlngBanco As Long
Private mlngConta As Long
Private mlngCentro As Long
Private mlngCheque As Long
Private mstrMoeda As String
Private mdblValor As Double
Private mdblAcrescimo As Double
Private mdblAbatimento As Double
Private mdatEmissao As Date
Private mdatVencimento As Date
Private mdatPagamento As Date
Private mdatLiberacao As Date
Private mdblPercMulta As Double
Private mdblValorMulta As Double
Private mdblPercMora As Double
Private mdblValorMora As Double
Private mdblDescontoPontualidade As Double
Private mstrObservacao As String
Private mobjCheque As clsCheque
Private mobjArquivoTexto As cArquivoTexto

Private Sub Class_Initialize()
    Set mobjArquivoTexto = New cArquivoTexto
End Sub

Private Sub Class_Terminate()
    Set mobjCheque = Nothing
End Sub

Public Property Get Observacao() As String
    Observacao = mstrObservacao
End Property

Public Property Let Observacao(ByVal strObservacao As String)
    mstrObservacao = strObservacao
End Property

Public Property Get PagarReceber() As String
    PagarReceber = mstrPagarReceber
End Property

Public Property Let PagarReceber(ByVal strPagarReceber As String)
    mstrPagarReceber = strPagarReceber
End Property

Public Property Get NumDuplicata() As Long
    NumDuplicata = mlngNumDuplicata
End Property

Public Property Let NumDuplicata(ByVal lngNumDuplicata As Long)
    mlngNumDuplicata = lngNumDuplicata
End Property

Public Property Get TipoGlobal() As String
    TipoGlobal = mstrTipoGlobal
End Property

Public Property Let TipoGlobal(ByVal strTipoGlobal As String)
    mstrTipoGlobal = strTipoGlobal
End Property

Public Property Get parcela() As Integer
    parcela = mintParcela
End Property

Public Property Let parcela(ByVal intParcela As Integer)
    mintParcela = intParcela
End Property

Public Property Get Empresa() As String
    Empresa = mstrEmpresa
End Property

Public Property Let Empresa(ByVal strEmpresa As String)
    mstrEmpresa = strEmpresa
End Property

Public Property Get Descricao() As String
    Descricao = mstrDescricao
End Property

Public Property Let Descricao(ByVal strDescricao As String)
    mstrDescricao = strDescricao
End Property

Public Property Get Banco() As Long
    Banco = mlngBanco
End Property

Public Property Let Banco(ByVal lngBanco As Long)
    mlngBanco = lngBanco
End Property

Public Property Get conta() As Long
    conta = mlngConta
End Property

Public Property Let conta(ByVal lngConta As Long)
    mlngConta = lngConta
End Property

Public Property Get Centro() As Long
    Centro = mlngCentro
End Property

Public Property Let Centro(ByVal lngCentro As Long)
    mlngCentro = lngCentro
End Property

Public Property Get Cheque() As Long
    Cheque = mlngCheque
End Property

Public Property Let Cheque(ByVal intCheque As Long)
    mlngCheque = intCheque
End Property

Public Property Get Moeda() As String
    Moeda = mstrMoeda
End Property

Public Property Let Moeda(ByVal strMoeda As String)
    mstrMoeda = strMoeda
End Property

Public Property Get valor() As Double
    valor = mdblValor
End Property

Public Property Let valor(ByVal dblValor As Double)
    mdblValor = dblValor
End Property

Public Property Get Acrescimo() As Double
    Acrescimo = mdblAcrescimo
End Property

Public Property Let Acrescimo(ByVal dblAcrescimo As Double)
    mdblAcrescimo = dblAcrescimo
End Property

Public Property Get Abatimento() As Double
    Abatimento = mdblAbatimento
End Property

Public Property Let Abatimento(ByVal dblAbatimento As Double)
    mdblAbatimento = dblAbatimento
End Property

Public Property Get Emissao() As Date
    Emissao = mdatEmissao
End Property

Public Property Let Emissao(ByVal datEmissao As Date)
    mdatEmissao = datEmissao
End Property

Public Property Get Vencimento() As Date
    Vencimento = mdatVencimento
End Property

Public Property Let Vencimento(ByVal datVencimento As Date)
    mdatVencimento = datVencimento
End Property

Public Property Get Pagamento() As Date
    Pagamento = mdatPagamento
End Property

Public Property Let Pagamento(ByVal datPagamento As Date)
    mdatPagamento = datPagamento
End Property

Public Property Get Liberacao() As Date
    Liberacao = mdatLiberacao
End Property

Public Property Let Liberacao(ByVal datLiberacao As Date)
    mdatLiberacao = datLiberacao
End Property

Public Property Get PercMulta() As Double
    PercMulta = mdblPercMulta
End Property

Public Property Let PercMulta(ByVal dblPercMulta As Double)
    mdblPercMulta = dblPercMulta
End Property

Public Property Get valorMulta() As Double
    valorMulta = mdblValorMulta
End Property

Public Property Let valorMulta(ByVal dblValorMulta As Double)
    mdblValorMulta = dblValorMulta
End Property

Public Property Get PercMora() As Double
    PercMora = mdblPercMora
End Property

Public Property Let PercMora(ByVal dblPercMora As Double)
    mdblPercMora = dblPercMora
End Property

Public Property Get ValorMora() As Double
    ValorMora = mdblValorMora
End Property

Public Property Let ValorMora(ByVal dblValorMora As Double)
    mdblValorMora = dblValorMora
End Property

Public Property Get DescontoPontualidade() As Double
    DescontoPontualidade = mdblDescontoPontualidade
End Property

Public Property Let DescontoPontualidade(ByVal dblDescontoPontualidade As Double)
    mdblDescontoPontualidade = dblDescontoPontualidade
End Property

Public Property Get objCheque() As clsCheque
    Set objCheque = mobjCheque
End Property

Public Property Set objCheque(objCheque As clsCheque)
    Set mobjCheque = objCheque
End Property

'Descrição..: Função que irá quebrar uma linha de0,0, registro do tipo Lancamento ou Observacao para preencher os atributos
'Parametros.: [String]
Public Sub Parser(strLinha As String)
    If UCase(Left(strLinha, 1)) = "L" Then
        mstrPagarReceber = Mid(strLinha, 2, 1)
        mlngNumDuplicata = CLng(Mid(strLinha, 3, 6))
        mstrTipoGlobal = RTrim(Mid(strLinha, 9, 20))
        mintParcela = CInt(Mid(strLinha, 29, 3))
        mstrEmpresa = Mid(strLinha, 32, 15)
        mstrDescricao = Mid(strLinha, 47, 80)
        mlngBanco = CLng(Mid(strLinha, 127, 9))
        mlngConta = CLng(Mid(strLinha, 136, 9))
        mlngCentro = CLng(Mid(strLinha, 145, 9))
        mlngCheque = CLng(Mid(strLinha, 154, 6))
        mstrMoeda = Mid(strLinha, 160, 9)
        mdblValor = CDbl(Mid(strLinha, 169, 9)) / 100
        mdblAcrescimo = CDbl(Mid(strLinha, 178, 7)) / 100
        mdblAbatimento = CDbl(Mid(strLinha, 185, 7)) / 100
        mdatEmissao = Mid(strLinha, 192, 2) & "/" & Mid(strLinha, 194, 2) & "/" & Mid(strLinha, 196, 4)
        mdatVencimento = Mid(strLinha, 200, 2) & "/" & Mid(strLinha, 202, 2) & "/" & Mid(strLinha, 204, 4)
        If CInt(Mid(strLinha, 208, 2)) <> 0 Then
            mdatPagamento = Mid(strLinha, 208, 2) & "/" & Mid(strLinha, 210, 2) & "/" & Mid(strLinha, 212, 4)
        End If
        If CInt(Mid(strLinha, 216, 2)) <> 0 Then
            mdatLiberacao = Mid(strLinha, 216, 2) & "/" & Mid(strLinha, 218, 2) & "/" & Mid(strLinha, 220, 4)
        End If
        mdblPercMulta = CDbl(Mid(strLinha, 224, 5)) / 100
        mdblValorMulta = CDbl(Mid(strLinha, 229, 7)) / 100
        mdblPercMora = CDbl(Mid(strLinha, 236, 5)) / 100
        mdblValorMora = CDbl(Mid(strLinha, 241, 7)) / 100
        mdblDescontoPontualidade = CDbl(Mid(strLinha, 248, 7)) / 100
    Else
        mstrObservacao = Mid(strLinha, 2, 253)
    End If
End Sub

Public Function Gravar() As Boolean
    Dim strCampos As String
    Dim strValores As String
    
On Error GoTo Error_Handler
    strCampos = " PagRec, Nota, Tipo, Parcela, Empresa, Descrição, Banco, Conta, Centro, Cheque, " & _
            " Moeda, [Valor Original], Acréscimo, Abatimento, Emissão, Vencimento, " & _
            " PerMul, VlrMul, PerMrd, VlrMrd, VlrDsP, Obs "
    
    strValores = " '" & mstrPagarReceber & "', " & mlngNumDuplicata & ", '" & mstrTipoGlobal & "', " & _
             mintParcela & ", '" & mstrEmpresa & "', '" & mstrDescricao & "', " & mlngBanco & ", " & _
             mlngConta & ", " & mlngCentro & ", " & mlngCheque & ", '" & mstrMoeda & "', " & _
             Replace(mdblValor, ",", ".") & ", " & Replace(mdblAcrescimo, ",", ".") & ", " & _
             Replace(mdblAbatimento, ",", ".") & ", " & InverteData(mdatEmissao, True) & ", " & _
             InverteData(mdatVencimento, True) & ", " & Replace(mdblPercMulta, ",", ".") & ", " & _
             Replace(mdblValorMulta, ",", ".") & ", " & Replace(mdblPercMora, ",", ".") & ", " & _
             Replace(mdblValorMora, ",", ".") & ", " & Replace(mdblDescontoPontualidade, ",", ".") & ", '" & _
             mstrObservacao & "' "
             
    If CStr(mdatPagamento) <> "00:00:00" Then
        strCampos = strCampos + ", Pagamento "
        strValores = strValores + ", " & InverteData(mdatPagamento, True)
    End If
    
    If CStr(mdatPagamento) <> "00:00:00" Then
        strCampos = strCampos + ", Liberação "
        strValores = strValores + ", " & InverteData(mdatLiberacao, True)
    End If
             
    If Not mobjCheque Is Nothing Then
        strCampos = strCampos + ", CheBan, CheAge, CheEmi, CheCco"
        
        strValores = strValores + ", " & mobjCheque.CodigoBanco & ", '" & mobjCheque.NumeroAgencia & "', '" & _
                     mobjCheque.NomeCorrentista & "', '" & mobjCheque.ContaCorrente & "'"
    End If
    Gravar = ExecuteSQL("INSERT INTO Duplicatas(" & strCampos & ") VALUES (" & strValores & ")") > 0
    Exit Function
    
Error_Handler:
    Call Err.Raise(Err.Number, TypeName(Me) & ".Gravar", Err.Description)
    Gravar = False
End Function

' 23/04/2019 - FBMI:618 - Yuji F. - Método para gravar as duplicatas a partir do objeto obtido do leitor
Public Function GravarDuplicata(oDuplicata As FoxArquivos.ImportDupDet) As Boolean
    Dim strCampos As String
    Dim strValores As String
    
    On Error GoTo Error_Handler
    
    strCampos = " PagRec, Nota, Tipo, Parcela, Empresa, Descrição, Banco, Conta, Centro, Cheque, " & _
        " Moeda, [Valor Original], Acréscimo, Abatimento, Emissão, Vencimento, " & _
        " PerMul, VlrMul, PerMrd, VlrMrd, VlrDsP, Obs "
        
    With oDuplicata
        strValores = " '" & .PagRec & "', " & CLng(.nrDuplicata) & ", '" & .TpGlobal & "', " & _
            CInt(.nrParcela) & ", '" & .Empresa & "', '" & .descDuplicata & "', " & CLng(.Banco) & ", " & _
            CLng(.conta) & ", " & CLng(.CentroCusto) & ", " & CLng(.nrCheque) & ", '" & .Moeda & "', " & _
            .vlDuplicata & ", " & .VLACRESCIMO & ", " & .VLABATIMENTO & ", " & InverteData(.dtEmissao, True) & ", " & _
            InverteData(.dtVencimento, True) & ", " & .PercMulta & ", " & .vlMulta & ", " & .PercMora & ", " & _
            .vlMora & ", " & .vlDescPontualidade & ", '"
        If .Observacao Is Nothing Then
            strValores = strValores & ""
        Else
            strValores = strValores & .Observacao.Descricao
        End If
        
        strValores = strValores & "' "
        
        If .dtPagamento <> "00000000" Then
            strCampos = strCampos + ", Pagamento "
            strValores = strValores + ", " & InverteData(.dtPagamento, True)
        End If
        
        If .dtLiberacao <> "00000000" Then
            strCampos = strCampos + ", Liberação "
            strValores = strValores + ", " & InverteData(.dtLiberacao, True)
        End If
        
        If Not .Cheque Is Nothing Then
            strCampos = strCampos + ", CheBan, CheAge, CheEmi, CheCco"
            
            strValores = strValores + ", " & CLng(.Cheque.Banco) & ", '" & .Cheque.Agencia & "', '" & _
            .Cheque.nmEmitente & "', '" & .Cheque.conta & "'"
        End If
    End With
    
    GravarDuplicata = ExecuteSQL("INSERT INTO Duplicatas(" & strCampos & ") VALUES (" & strValores & ")") > 0
    Exit Function
    
Error_Handler:
    Call Err.Raise(Err.Number, TypeName(Me) & ".GravarDuplicata", Err.Description)
    GravarDuplicata = False
End Function

Public Function Validar(objHeader As clsHeader) As Boolean
    If mdatEmissao >= objHeader.dtInicial And mdatEmissao <= objHeader.dtFinal Then
        Validar = True
    Else
        Validar = False
    End If
End Function
