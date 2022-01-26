VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bizCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ProximoCheque(ByVal Banco As Long) As Long
    Dim dao As New DaoCheque
    If Banco <> 0 Then
        ProximoCheque = dao.ProximoCheque(Banco)
    End If
    Set dao = Nothing
End Function

Public Function Carregar(ByVal lngBanco As Long, ByVal lngCheque As Long) As VoCheque
    Dim dao As New DaoCheque
    Aplicacao.Connect
    Call dao.init(Aplicacao)
    Set Carregar = dao.Carregar(lngBanco, lngCheque)
    Set dao = Nothing
    Aplicacao.Disconnect
End Function

Public Function Gravar(ByVal obj As VoCheque) As Boolean
    Dim dao As New DaoCheque
    Dim blnGravar       As Boolean
On Error GoTo Err
    Aplicacao.Connect
    Aplicacao.BeginTransaction
    Call dao.init(Aplicacao)
    blnGravar = dao.Gravar(obj)
    Set dao = Nothing
    If blnGravar Then
        Aplicacao.CommitTransaction
    Else
        Aplicacao.RollbackTransaction
    End If
    Aplicacao.Disconnect
    Gravar = blnGravar
    Exit Function
Err:
    Set dao = Nothing
    Aplicacao.RollbackTransaction
    Aplicacao.Disconnect
    Gravar = False
End Function

Public Function Atualizar(ByVal obj As VoCheque) As Boolean
    Dim dao As New DaoCheque
    Dim blnAtualizar As Boolean

On Error GoTo Err
    Aplicacao.Connect
    Aplicacao.BeginTransaction
    Call dao.init(Aplicacao)
    blnAtualizar = dao.Atualizar(obj)
    If blnAtualizar Then
        Aplicacao.CommitTransaction
    Else
        Aplicacao.RollbackTransaction
    End If
    Set dao = Nothing
    Aplicacao.Disconnect
    Atualizar = blnAtualizar
    Exit Function
Err:
    Set dao = Nothing
    Aplicacao.RollbackTransaction
    Aplicacao.Disconnect
    Atualizar = False
End Function

Public Function Excluir(ByVal lngBanco As Long, ByVal lngCheque As Long) As Boolean
    Dim dao As New DaoCheque
    Aplicacao.Connect
    Call dao.init(Aplicacao)
    Excluir = dao.Excluir(lngBanco, lngCheque)
    Set dao = Nothing
    Aplicacao.Disconnect
End Function

'Projeto: #0 - História: #18924 - Desenvolvimento#18962 - Moacir Pfau(11/12/2012)
Public Function Gravar_Atualizar(ByVal obj As VoCheque) As Boolean
    Dim blnGravarAtualizar      As Boolean
On Error GoTo Err
    If Carregar(obj.Banco, obj.Cheque) Is Nothing Then
        blnGravarAtualizar = Gravar(obj)
    Else
        blnGravarAtualizar = Atualizar(obj)
    End If
    Gravar_Atualizar = blnGravarAtualizar
    Exit Function
Err:
    Gravar_Atualizar = False
End Function

