VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArquivoDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Autor.......: Gustavo Cuman
'Data........: 09/12/2008
'Descrição...:

Private mobjCheque As clsCheque
Private mobjArquivoTexto As cArquivoTexto
Private mintIndex As Integer
Private mobjHeader As clsHeader
Private mobjLancamento As clsLancamento

Private Property Get ArquivoTexto() As cArquivoTexto
    Set ArquivoTexto = mobjArquivoTexto
End Property

Private Property Set ArquivoTexto(objArquivoTexto As cArquivoTexto)
    Set mobjArquivoTexto = objArquivoTexto
End Property

Private Property Get Header() As clsHeader
    Set Header = mobjHeader
End Property

Private Property Set Header(objHeader As clsHeader)
    Set mobjHeader = objHeader
End Property

Private Property Get Lancamento() As clsLancamento
    Set Lancamento = mobjLancamento
End Property

Private Property Set Lancamento(objLancamento As clsLancamento)
    Set mobjLancamento = objLancamento
End Property

Private Property Get Cheque() As clsCheque
    Set Cheque = mobjCheque
End Property

Private Property Set Cheque(objCheque As clsCheque)
    Set mobjCheque = objCheque
End Property

Private Sub Class_Initialize()
    Set mobjArquivoTexto = New cArquivoTexto
End Sub

' 23/04/2019 - FBMI:618 - Yuji F. - Novo parâmetro de encoding, para a leitura do arquivo
Public Function Carregar(strArquivo As String, Optional sEncoding As String = "x-ansi") As Boolean
    Set mobjHeader = Nothing
    Carregar = mobjArquivoTexto.Carregar(strArquivo, sEncoding)
End Function

Public Function TotalLinhas() As Integer
    TotalLinhas = mobjArquivoTexto.TotalLinhas
End Function

Public Sub ProximaLinha()
    mintIndex = mintIndex + 1
End Sub

Public Function getLinha() As clsLancamento
    Set mobjLancamento = New clsLancamento
    If mintIndex <= mobjArquivoTexto.TotalLinhas Then
        Call mobjLancamento.Parser(mobjArquivoTexto.getLinha(mintIndex))
        If mintIndex <> mobjArquivoTexto.TotalLinhas Then
            If Left(mobjArquivoTexto.getLinha(mintIndex + 1), 1) = "O" Then
                Call ProximaLinha
                mobjLancamento.Observacao = Mid(mobjArquivoTexto.getLinha(mintIndex), 2, 251)
            End If
            If mintIndex + 1 <= mobjArquivoTexto.TotalLinhas Then
                If Left(mobjArquivoTexto.getLinha(mintIndex + 1), 1) = "C" Then
                    Call ProximaLinha
                    Set mobjLancamento.objCheque = New clsCheque
                    Call mobjLancamento.objCheque.Parser(mobjArquivoTexto.getLinha(mintIndex))
                End If
            End If
        End If
        Set getLinha = mobjLancamento
    End If
End Function

Public Function getHeader() As clsHeader
    If mobjHeader Is Nothing Then
        Set mobjHeader = New clsHeader
        If Left(mobjArquivoTexto.getLinha(1), 1) = "H" Then
            Call mobjHeader.Parser(mobjArquivoTexto.getLinha(1))
        End If
    End If
    Set getHeader = mobjHeader
End Function

'Correção Importação Duplicata - PT 92531 - Gustavo
Public Function UltimaLinha() As Boolean
    UltimaLinha = (mintIndex > mobjArquivoTexto.TotalLinhas)
End Function

Public Sub PrimeiraLinha()
    mintIndex = 1
End Sub

Private Sub Class_Terminate()
    Set mobjCheque = Nothing
    Set mobjArquivoTexto = Nothing
    Set mobjHeader = Nothing
    Set mobjLancamento = Nothing
End Sub
