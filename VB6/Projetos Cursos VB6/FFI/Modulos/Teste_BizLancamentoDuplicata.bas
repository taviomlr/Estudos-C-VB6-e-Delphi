Attribute VB_Name = "Teste_BizLancamentoDuplicata"
#If TESTE Then
Option Explicit

Public Function BizLancamentoDuplicata_validarCampoObrigatorioEmissao_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection
    On Error GoTo err
    
    Call obj.validarCampoObrigatorio(Empty, Date, 1, 1, 1, 1.01, 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "O campo 'Emiss�o' � de preenchimento obrigat�rio." Then
            BizLancamentoDuplicata_validarCampoObrigatorioEmissao_RetornaMensagemErro = True
        End If
    End If
    Exit Function
err:
    
End Function

Public Function BizLancamentoDuplicata_validarCampoObrigatorioVencimento_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection
    
    Call obj.validarCampoObrigatorio(Date, Empty, 1, 1, 1, 1.01, 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "O campo 'Vencimento' � de preenchimento obrigat�rio." Then
            BizLancamentoDuplicata_validarCampoObrigatorioVencimento_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarCampoObrigatorioBanco_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection
    
    Call obj.validarCampoObrigatorio(Date, Date, 0, 1, 1, 1.01, 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "O campo 'Banco' � de preenchimento obrigat�rio." Then
            BizLancamentoDuplicata_validarCampoObrigatorioBanco_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarCampoObrigatorioConta_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection
    
    Call obj.validarCampoObrigatorio(Date, Date, 1, 0, 1, 1.01, 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "O campo 'Conta' � de preenchimento obrigat�rio." Then
            BizLancamentoDuplicata_validarCampoObrigatorioConta_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarCampoObrigatorioCentrodeCusto_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection
    
    Call obj.validarCampoObrigatorio(Date, Date, 1, 1, 0, 1.01, 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "O campo 'Centro de Custo' � de preenchimento obrigat�rio." Then
            BizLancamentoDuplicata_validarCampoObrigatorioCentrodeCusto_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarCampoObrigatorioValorOriginal_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection
    
    Call obj.validarCampoObrigatorio(Date, Date, 1, 1, 1, 0, 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "O campo 'Valor Original' � de preenchimento obrigat�rio." Then
            BizLancamentoDuplicata_validarCampoObrigatorioValorOriginal_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarCampoObrigatorioParcela_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection
    
    Call obj.validarCampoObrigatorio(Date, Date, 1, 1, 1, 1.01, 0, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "O campo 'Parcela' � de preenchimento obrigat�rio." Then
            BizLancamentoDuplicata_validarCampoObrigatorioParcela_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarCampoObrigatorioOperacaoContabil_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection
    
    Call obj.validarCampoObrigatorio(Date, Date, 1, 1, 1, 1.01, 1, 0, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "O campo 'Opera��o Cont�bil' � de preenchimento obrigat�rio." Then
            BizLancamentoDuplicata_validarCampoObrigatorioOperacaoContabil_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarCampoVencimentoAnteriorEmissao_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection

    Call obj.validarInformacaoGeral(Date, Format(DateSerial(Year(Date), Month(Date), Day(Date) - 1), "dd/MM/yyyy"), 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "A data de 'Vencimento' � anterior a data de 'Emiss�o'." Then
            BizLancamentoDuplicata_validarCampoVencimentoAnteriorEmissao_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarDataLimiteCentroCusto_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection

    Call obj.validarInformacaoGeral(Date, Format(DateSerial(Year(Date), Month(Date), Day(Date) - 1), "dd/MM/yyyy"), 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "A Data do lan�amento ultrapassa a 'Data Limite' para movimenta��o do Centro de Custo." Then
            BizLancamentoDuplicata_validarDataLimiteCentroCusto_RetornaMensagemErro = True
        End If
    End If
End Function

Public Function BizLancamentoDuplicata_validarCampoContaAtiva_RetornaMensagemErro() As Boolean
    Dim obj As New BizLancamentoDuplicata
    Dim col As New Collection

    Call obj.validarInformacaoGeral(Date, Date, 1, 1, col)
    If col.Count = 1 Then
        If col.Item(1).mensagem = "A 'Conta' n�o est� ativa, somente poder� ser preenchida uma 'Conta Ativa'." Then
            BizLancamentoDuplicata_validarCampoContaAtiva_RetornaMensagemErro = True
        End If
    End If
End Function

#End If
