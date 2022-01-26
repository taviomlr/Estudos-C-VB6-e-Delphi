VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BizImpDigExtratoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 61827 - Demanda: 61828
'---|   Ueder Budni (12/01/2015)
'---|---------------------------------------------------------------------------------------------------------------------------

Private objDao As DaoImpDigExtratoBancario


Private Sub Class_Initialize()
    Set objDao = New DaoImpDigExtratoBancario
    Call objDao.init(Aplicacao)
End Sub

Private Sub Class_Terminate()
    Set objDao = Nothing
End Sub

Public Function carregarColecao(ByVal lngBanco As Long, Optional ByVal dteMesAno As Date, Optional ByVal lngNumExtrato As Long, Optional blnMostraConciliado As Boolean, Optional ByVal strOptTipo As String) As ColImpDigExtratoBancario

On Error GoTo Erro
    Set carregarColecao = objDao.CarregaColecao(lngBanco, Month(dteMesAno), Year(dteMesAno), lngNumExtrato, blnMostraConciliado, strOptTipo)
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi possível carregar Coleção! (BizImpDigExtratoBancario.CarregaColecao)"
End Function


Public Function SalvaColecao(ByVal col As ColImpDigExtratoBancario, lngBanco As Long, dteMesAno As Date, Optional lngExtrato As Long) As Boolean
    Dim dao As New DaoImpDigExtratoBancario

On Error GoTo Erro

    If col.Count > 0 Then
        SalvaColecao = dao.GravaColecao(col, lngBanco, Month(dteMesAno), Year(dteMesAno), lngExtrato)
        Set dao = Nothing
    Else
        SalvaColecao = dao.ExcluirTodos(lngBanco, Month(dteMesAno), Year(dteMesAno), lngExtrato)
    End If
    
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi possível gravar a coleção! (BizImpDigExtratoBancario.SalvaColecao)"
    Set dao = Nothing
End Function

Public Function ExisteDescricao(strDescricao As String, lngBanco As Long) As Boolean
    Dim col As New ColImpDigExtratoBancario

On Error GoTo Erro

    Set col = objDao.CarregaColecao(lngBanco)
    
    ExisteDescricao = False
    While Not col.EOF
        If strDescricao = col.CurrentObject.Descricao Then
            ExisteDescricao = True
            GoTo fim
        End If
        col.MoveNext
    Wend
fim:
    Exit Function
Erro:
    err.Raise 1 + vbObjectError, , "Não foi verificar se a descrição está duplicada! (BizImpDigExtratoBancario.ExisteDescricao)"
End Function

