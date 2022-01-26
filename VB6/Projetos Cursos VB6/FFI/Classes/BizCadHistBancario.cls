VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BizCadHistBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 61827 - Demanda: 61828
'---|   Ueder Budni (13/01/2015)
'---|---------------------------------------------------------------------------------------------------------------------------

Private objDAO As DaoCadHistBancario


Private Sub Class_Initialize()
    Set objDAO = New DaoCadHistBancario
    Call objDAO.init(Aplicacao)
End Sub

Private Sub Class_Terminate()
    Set objDAO = Nothing
End Sub

Public Function carregarColecao(ByVal lngBanco As Long) As ColCadHistBancario

On Error GoTo Erro
    Set carregarColecao = objDAO.CarregaColecao(lngBanco)
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi possível carregar Coleção! (BizCadHistBancario.CarregaColecao)"
End Function

Public Function SalvaColecao(ByVal col As ColCadHistBancario, lngBanco As Long) As Boolean
On Error GoTo Erro

    If col.Count > 0 Then
        SalvaColecao = objDAO.GravaColecao(col, lngBanco)
    Else
        SalvaColecao = objDAO.ExcluirTodos(lngBanco)
    End If
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi possível gravar a coleção! (BizCadHistBancario.SalvaColecao)"
End Function

Public Function NovoCodigo(col As ColCadHistBancario, lngBanco As Long) As Long
    Dim Max     As Integer
On Error GoTo Erro
    
    Max = objDAO.NovoCodigo(lngBanco)
    With col
        .MoveFirst
        While Not .EOF
            If .CurrentObject.CdHistorico >= Max Then
                Max = .CurrentObject.CdHistorico + 1
            End If
            .MoveNext
        Wend
    End With
    NovoCodigo = Max
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi possível gerar um novo código! (BizCadHistBancario.NovoCodigo)"
End Function

Public Function CarregarHistorico(lngBanco As Long, lngCodHist As Long, strTipoOperacao As String, Optional blnFiltraAmbos As Boolean) As VoCadHistBancario
    Dim objVO   As VoCadHistBancario
On Error GoTo Erro
    
    Set objVO = objDAO.CarregaVOpelaPK(ModGeral.EnterpriseId, ModGeral.CdEstabelecimento, lngBanco, lngCodHist, strTipoOperacao, blnFiltraAmbos)
    Set CarregarHistorico = objVO
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi possível carregar o histórico! (BizCadHistBancario.CarregarHistorico)"
End Function
