VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCNAB240Caixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mstrIdentificacaoBanco  As String
Private mstrCodigoMoeda         As String
Private mstrValorDocumento      As String
Private mstrProximoNossoNumero  As String
Private mintCarteira            As Integer
Private mstrNossoNumero         As String
Private mstrCampoLivre          As String
Private mstrCodigoBarras        As String
Private mstrDigitoCodigoBarras  As String
Private mstrDigitoNossoNumero   As String
Private mstrAgencia             As String
Private mstrOperacao            As String
Private mstrNumeroContaCorrente As String
Private mstrDigitoContaCorrente As String
Private mstrFatorVencimento     As String
Private mstrDataVencimento      As String
Private mstrLinhaDigitavel      As String

Private Sub Class_Initialize()
    mstrCodigoMoeda = "9"
    mstrOperacao = "000"
End Sub

Public Property Let IdentificacaoBanco(ByVal NewVal As String)
    mstrIdentificacaoBanco = NewVal
End Property

Public Property Let ValorDocumento(ByVal NewVal As String)
    mstrValorDocumento = NewVal
End Property

Public Property Let ProximoNossoNumero(ByVal NewVal As String)
    mstrProximoNossoNumero = NewVal
End Property

Public Property Let Carteira(ByVal NewVal As Integer)
    mintCarteira = NewVal
End Property

Public Property Get CampoLivre() As String
    CampoLivre = mstrCampoLivre
End Property

Public Property Get NossoNumero() As String
    NossoNumero = mstrNossoNumero
End Property

Public Property Get DigitoNossoNumero() As String
    DigitoNossoNumero = mstrDigitoNossoNumero
End Property

Public Property Get CodigoBarras() As String
    CodigoBarras = mstrCodigoBarras
End Property

Public Property Get Agencia() As String
    Agencia = mstrAgencia
End Property

Public Property Let Agencia(ByVal strAgencia As String)
    mstrAgencia = strAgencia
End Property

Public Property Get Operacao() As String
    Operacao = mstrOperacao
End Property

Public Property Let Operacao(ByVal strOperacao As String)
    mstrOperacao = strOperacao
End Property

Public Property Get NumeroContaCorrente() As String
    NumeroContaCorrente = mstrNumeroContaCorrente
End Property

Public Property Let NumeroContaCorrente(ByVal strNumeroContaCorrente As String)
    mstrNumeroContaCorrente = strNumeroContaCorrente
End Property

Public Property Get DigitoContaCorrente() As String
    DigitoContaCorrente = mstrDigitoContaCorrente
End Property

Public Property Let DigitoContaCorrente(ByVal strDigitoContaCorrente As String)
    mstrDigitoContaCorrente = strDigitoContaCorrente
End Property

Public Property Let DataVencimento(ByVal NewVal As String)
    mstrDataVencimento = NewVal
End Property

Public Property Get LinhaDigitavel() As String
    Dim strLinhaDigitavel As String
    
    strLinhaDigitavel = Replace(mstrLinhaDigitavel, ".", "")
    strLinhaDigitavel = Replace(strLinhaDigitavel, " ", "")
    LinhaDigitavel = strLinhaDigitavel
End Property

Public Property Get LinhaDigitavelFormatada() As String
    LinhaDigitavelFormatada = mstrLinhaDigitavel
End Property

Public Function CriarBoleto() As Boolean
    If mintCarteira > 10 And mintCarteira < 15 Then
        If mintCarteira <> 13 Then
            Call CalculaNossoNumero
            Call CalculaCampoLivre
            If CalculaFatorVencimento Then
                Call CalculaCodigoBarras
                Call CalculaLinhaDigitavel
                CriarBoleto = True
            Else
                CriarBoleto = False
            End If
        Else
            CriarBoleto = False
        End If
    Else
        CriarBoleto = False
    End If
End Function

Private Sub CalculaNossoNumero()
    Select Case mintCarteira
        Case 11 'Cobrança Simples
            mstrNossoNumero = Format(mstrProximoNossoNumero, "0000000000")
        Case 12 'Cobrança Rápida
            mstrNossoNumero = "9" & Format(mstrProximoNossoNumero, "000000000")
        Case 14 'Cobrança Sem Registro
            mstrNossoNumero = "82" & Format(mstrProximoNossoNumero, "00000000")
    End Select
    mstrDigitoNossoNumero = DigitoVerificador(mstrNossoNumero, 11, False, False, 2, 3, 4, 5, 6, 7, 8, 9)
End Sub

Private Sub CalculaCampoLivre()
    Select Case mintCarteira
        Case 11, 12, 14 'Cobrança Simples, Cobrança Rápida, Cobrança Sem Registro
            mstrCampoLivre = Format(mstrNossoNumero, "0000000000")
            mstrCampoLivre = mstrCampoLivre & Format(mstrAgencia, "0000")
            mstrCampoLivre = mstrCampoLivre & Format(mstrOperacao, "000")
            mstrCampoLivre = mstrCampoLivre & Format(mstrNumeroContaCorrente, "00000000")
    End Select
End Sub

Private Function CalculaFatorVencimento() As Boolean
    Dim datDataBase As Date
    
    If mstrDataVencimento <> "" Then
        datDataBase = CDate("07/10/1997")
        mstrFatorVencimento = DateDiff("d", datDataBase, mstrDataVencimento)
        CalculaFatorVencimento = True
    Else
        mstrFatorVencimento = ""
        CalculaFatorVencimento = False
    End If
End Function

Private Sub CalculaCodigoBarras()
    Dim strCodigo As String
    
    strCodigo = Format(mstrIdentificacaoBanco, "000")
    strCodigo = strCodigo & mstrCodigoMoeda
    strCodigo = strCodigo & Format(mstrFatorVencimento, "0000")
    strCodigo = strCodigo & Format(ValorUnformated, "0000000000")
    strCodigo = strCodigo & Format(mstrCampoLivre, "0000000000000000000000000")
    mstrCodigoBarras = Left(strCodigo, 4)
    mstrDigitoCodigoBarras = DigitoVerificador(strCodigo, 11, False, True, 2, 3, 4, 5, 6, 7, 8, 9)
    mstrCodigoBarras = mstrCodigoBarras & mstrDigitoCodigoBarras
    mstrCodigoBarras = mstrCodigoBarras & Right(strCodigo, Len(strCodigo) - (Len(mstrCodigoBarras) - 1))
End Sub

Private Function DigitoVerificador(strNumero As String, lngModulo As Long, blnAgrupaNumero As Boolean, blnCodigoBarras As Boolean, ParamArray lngPeso() As Variant) As String
    Dim strMultiplicadores As String
    Dim intIndex     As Integer
    Dim strResultado As String
    Dim lngTotal     As Long
    
    If blnCodigoBarras Then
        blnAgrupaNumero = False
    End If
    While Len(strMultiplicadores) < Len(strNumero)
        For intIndex = 0 To UBound(lngPeso)
            If Len(strMultiplicadores) < Len(strNumero) Then
                strMultiplicadores = lngPeso(intIndex) & strMultiplicadores
            End If
        Next
    Wend
    For intIndex = Len(strNumero) To 1 Step -1
        strResultado = (Mid(strNumero, intIndex, 1) * Mid(strMultiplicadores, intIndex, 1))
        If Len(strResultado) > 1 And blnAgrupaNumero Then
            strResultado = CInt(Left(strResultado, 1)) + CInt(Right(strResultado, 1))
        End If
        lngTotal = lngTotal + strResultado
    Next
    lngTotal = lngTotal Mod lngModulo
    'Tratamento específico para o código de barras conforme o manual.
    If Not blnCodigoBarras Then
        If lngTotal < 10 Then
            lngTotal = lngModulo - lngTotal
        End If
        If lngTotal > 9 Then
            lngTotal = 0
        End If
    Else
        If lngTotal = 0 Or lngTotal = 10 Or lngTotal = 1 Then
            lngTotal = 1
        End If
    End If
    DigitoVerificador = lngTotal
End Function

Private Sub CalculaLinhaDigitavel()
    Dim strCampoDigitavel As String
    Dim lngIntervalo      As Long
    
    '1º Campo da linha digitável
    strCampoDigitavel = Format(mstrIdentificacaoBanco, "000") & mstrCodigoMoeda
    strCampoDigitavel = strCampoDigitavel & Left(mstrCampoLivre, 5)
    strCampoDigitavel = strCampoDigitavel & DigitoVerificador(strCampoDigitavel, 10, True, False, 2, 1)
    mstrLinhaDigitavel = Left(strCampoDigitavel, 5) & "." & Right(strCampoDigitavel, 5) & " "
    
    '2º Campo da linha digitável
    strCampoDigitavel = Mid(mstrCampoLivre, 6, 10)
    strCampoDigitavel = strCampoDigitavel & DigitoVerificador(strCampoDigitavel, 10, True, False, 2, 1)
    mstrLinhaDigitavel = mstrLinhaDigitavel & Left(strCampoDigitavel, 5) & "."
    mstrLinhaDigitavel = mstrLinhaDigitavel & Right(strCampoDigitavel, 6) & " "
    
    '3º Campo da linha digitável
    strCampoDigitavel = Mid(mstrCampoLivre, 16, 10)
    strCampoDigitavel = strCampoDigitavel & DigitoVerificador(strCampoDigitavel, 10, True, False, 2, 1)
    mstrLinhaDigitavel = mstrLinhaDigitavel & Left(strCampoDigitavel, 5) & "."
    mstrLinhaDigitavel = mstrLinhaDigitavel & Right(strCampoDigitavel, 6) & " "
    
    '4º Campo da linha digitável
    strCampoDigitavel = mstrDigitoCodigoBarras
    mstrLinhaDigitavel = mstrLinhaDigitavel & strCampoDigitavel & " "
    
    '5º Campo da linha digitável
    strCampoDigitavel = mstrFatorVencimento
    lngIntervalo = 14 - (Len(mstrFatorVencimento) + Len(ValorUnformated))
    strCampoDigitavel = strCampoDigitavel & String(lngIntervalo, "0")
    strCampoDigitavel = strCampoDigitavel & ValorUnformated
    mstrLinhaDigitavel = mstrLinhaDigitavel & strCampoDigitavel
End Sub

Private Function ValorUnformated() As String
    Dim strValor As String
    
    strValor = Replace(mstrValorDocumento, "R$", "")
    strValor = Replace(strValor, ",", "")
    strValor = Replace(strValor, ".", "")
    strValor = Trim(strValor)
    ValorUnformated = strValor
End Function
