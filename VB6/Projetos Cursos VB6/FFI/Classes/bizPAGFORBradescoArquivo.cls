VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bizPAGFORBradescoArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GeraArquivopaPAGFOR(objHeader As voPAGFORBradescoHeader, objDetailCol As voPAGFORBradescoDetailCol, objTrailler As voPAGFORBradescoTrailler, ByRef objArquivoRemessa As cArquivoTexto, ByRef strErros As String) As Boolean
    Dim blnAux As Boolean
On Error GoTo err_Handler
    If (objHeader Is Nothing Or objDetailCol Is Nothing Or objTrailler Is Nothing) Then
        GeraArquivopaPAGFOR = False
    Else
        GeraArquivopaPAGFOR = geraHeaderArquivo(objHeader, objArquivoRemessa, strErros)
        GeraArquivopaPAGFOR = GeraArquivopaPAGFOR And geraDetailArquivo(objDetailCol, objArquivoRemessa, strErros)
        GeraArquivopaPAGFOR = GeraArquivopaPAGFOR And geraTraillerArquivo(objTrailler, objArquivoRemessa, strErros)
    End If
    Exit Function
err_Handler:
    GeraArquivopaPAGFOR = False
End Function

Private Function geraHeaderArquivo(objHeader As voPAGFORBradescoHeader, ByRef objArquivoRemessa As cArquivoTexto, ByRef strErros As String) As Boolean
    Dim strLinha  As String
    
On Error GoTo ErroHeader
    If Not objArquivoRemessa Is Nothing Then
        With objHeader
            strLinha = .IdentificacaoRegistro
            strLinha = strLinha & formatCampoInt(.CodigoComunicacao, 8)
            strLinha = strLinha & formatCampoInt(.TipoInscricao, 1)
            strLinha = strLinha & formatCampoInt(.NumeroInscricaoCNPJ, 9)
            strLinha = strLinha & formatCampoInt(.FilialCNPJ, 4)
            strLinha = strLinha & formatCampoInt(.ControleCNPJ, 2)
            strLinha = strLinha & formatCampoStr(.RazaoSocial, 40)
            strLinha = strLinha & formatCampoInt(.TipoServico, 2)
            strLinha = strLinha & formatCampoInt(.OrigemArquivo, 1)
            strLinha = strLinha & formatCampoInt(.NumeroRemessa, 5)
            strLinha = strLinha & formatCampoInt("0", 5)
            strLinha = strLinha & formatCampoInt(.DataGravacao, 8)
            strLinha = strLinha & formatCampoInt(.HoraGravacao, 6)
            strLinha = strLinha & Space(13)
            strLinha = strLinha & formatCampoInt(.TipoProcessamento, 1)
            strLinha = strLinha & Space(371)
            strLinha = strLinha & formatCampoInt(.NumeroLista, 9)
            strLinha = strLinha & Space(8)
            strLinha = strLinha & formatCampoInt(.Sequencial, 6)
            Call objArquivoRemessa.add(strLinha)
        End With
    End If
    geraHeaderArquivo = True
    Exit Function
ErroHeader:
    geraHeaderArquivo = False
    strErro = err.Description
End Function

Private Function geraDetailArquivo(objDetailCol As voPAGFORBradescoDetailCol, ByRef objArquivoRemessa As cArquivoTexto, ByRef strErros As String) As Boolean
    Dim strLinha            As String
    Dim intCont             As Integer
    Dim objDetail           As voPAGFORBradescoDetail
    Dim strLinhaDigitavel   As String
    
On Error GoTo ErroHeader

    If Not objArquivoRemessa Is Nothing And objDetailCol.Count > 0 Then
        objDetailCol.MoveFirst
        While Not objDetailCol.EOF
            Set objDetail = objDetailCol.CurrentObject
            strLinha = Empty
            With objDetail
                strLinha = formatCampoInt(.IdentificacaoDetalhe, 1)
                strLinha = strLinha & formatCampoInt(.TipoInscricao, 1)
                strLinha = strLinha & formatCampoInt(.NumeroInscricaoCNPJ, 9)
                strLinha = strLinha & formatCampoInt(.FilialCNPJ, 4)
                strLinha = strLinha & formatCampoInt(.ControleCNPJ, 2)
                strLinha = strLinha & formatCampoStr(.RazaoSocial, 30)
                strLinha = strLinha & formatCampoStr(.EnderecoFornec, 40)
                strLinha = strLinha & formatCampoInt(.CEPFornec, 5)
                strLinha = strLinha & formatCampoInt(.CEPCompFornec, 3)
                strLinha = strLinha & formatCampoInt(.CodBancoFornec, 3)
                strLinha = strLinha & formatCampoInt(.CodAgenciaFornec, 5)
                strLinha = strLinha & formatCampoStr(.DigitoAgenciaFornec, 1)
                strLinha = strLinha & formatCampoInt(.ContaCorrenteFornec, 13)
                strLinha = strLinha & formatCampoStr(.DigContaCorrenteFornec, 2)
                strLinha = strLinha & formatCampoStr(.NumeroPagt, 16)
                strLinha = strLinha & formatCampoInt(.Carteira, 3)
                strLinha = strLinha & formatCampoInt(.NossoNumero, 12)
                strLinha = strLinha & Space(15)
                strLinha = strLinha & formatCampoInt(.DataVencimento, 8)
                strLinha = strLinha & formatCampoInt(.dataEmissao, 8)
                strLinha = strLinha & formatCampoInt(.DataLimiteDesconto, 8)
                strLinha = strLinha & formatCampoInt("0", 1)
                strLinha = strLinha & formatCampoInt(.FatorVencimento, 4)
                strLinha = strLinha & formatCampoDbl(Round(.ValorDocumento, 2), 8, 2)
                strLinha = strLinha & formatCampoDbl(Round(.ValorPagamento, 2), 13, 2)
                strLinha = strLinha & formatCampoDbl(Round(.ValorDesconto, 2), 13, 2)
                strLinha = strLinha & formatCampoDbl(Round(.valorAcrescimo, 2), 13, 2)
                strLinha = strLinha & formatCampoInt(.TipoDocumento, 2)
                strLinha = strLinha & formatCampoInt("0", 1)
                strLinha = strLinha & formatCampoInt(.NumeroNF, 9)
                strLinha = strLinha & formatCampoStr(.SerieNF, 2)
                strLinha = strLinha & formatCampoInt(.ModalidadePagt, 2)
                strLinha = strLinha & formatCampoInt(.DataEfetivacaoPagt, 8)
                strLinha = strLinha & Space(3)
                strLinha = strLinha & formatCampoInt(.SituacaoAgendamento, 2)
                strLinha = strLinha & Space(10)
                strLinha = strLinha & formatCampoInt(.TipoMovimento, 1)
                strLinha = strLinha & formatCampoInt(.CodigoMovimento, 2)
                strLinha = strLinha & Space(82)
                
                'Informações Complementares
                Select Case .ModalidadePagt
                    Case 31 '-----------------Título de Terceiros-----------------------
                        'Ivo Sousa (28/07/2017) - Alterada a forma de pegar o código de barras
                        strLinhaDigitavel = Mid(.CodigoBarras, 5, 27)
                        strLinha = strLinha & Left(strLinhaDigitavel, 5) & Mid(strLinhaDigitavel, 7, 10) & Mid(strLinhaDigitavel, 18, 10)
                        strLinha = strLinha & formatCampoInt(Mid(.CodigoBarras, 33, 1), 1)
                        strLinha = strLinha & formatCampoInt(Mid(.CodigoBarras, 4, 1), 1)
                        strLinha = strLinha & Space(13)
                    Case Else '-----------------------------DOC COMPE/TED-----------------------------
                        strLinha = strLinha & formatCampoStr(.TipoDocCompensado, 1)
                        strLinha = strLinha & formatCampoInt(0, 6)
                        strLinha = strLinha & formatCampoInt(.CodigoFinalidade, 2)
                        strLinha = strLinha & formatCampoInt(.TipoConta, 2)
                        strLinha = strLinha & Space(29)
                End Select
                
                strLinha = strLinha & formatCampoInt("0", 2)
                strLinha = strLinha & Space(57)
                strLinha = strLinha & formatCampoInt(Right(.CodigoLancamento, 5), 5)
                strLinha = strLinha & Space(1)
                strLinha = strLinha & formatCampoInt(.TipoContaForn, 1)
                strLinha = strLinha & formatCampoInt(.ContaComplementar, 7)
                strLinha = strLinha & Space(8)
                strLinha = strLinha & formatCampoInt(.SequencialRegistro, 6)
                Call objArquivoRemessa.add(strLinha)
                objDetailCol.MoveNext
            End With
        Wend
    End If
    geraDetailArquivo = True
    Exit Function
ErroHeader:
    geraDetailArquivo = False
    strErro = err.Description
End Function

Private Function geraTraillerArquivo(objTrailler As voPAGFORBradescoTrailler, ByRef objArquivoRemessa As cArquivoTexto, ByRef strErros As String) As Boolean
    Dim strLinha  As String
    
On Error GoTo ErroHeader
    If Not objArquivoRemessa Is Nothing Then
        With objTrailler
            strLinha = formatCampoInt(.IdentificacaoRegistro, 1)
            strLinha = strLinha & formatCampoInt(.QuantidadeRegistros, 6)
            strLinha = strLinha & formatCampoDbl(Round(.TotalValores, 2), 15, 2)
            strLinha = strLinha & Space(470)
            strLinha = strLinha & formatCampoInt(.SequencialArquivo, 6)
            Call objArquivoRemessa.add(strLinha)
        End With
        geraTraillerArquivo = True
    Else
        geraTraillerArquivo = False
    End If
    Exit Function
ErroHeader:
    geraTraillerArquivo = False
    strErro = err.Description
End Function
