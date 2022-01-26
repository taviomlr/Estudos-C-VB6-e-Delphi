VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeracaoTituloReceberNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rdGerTitReceber As IDBReader
Private Const nomeTabela$ = "FFITituloReceber"

Private Sub Class_Initialize()
    Dim cmd As IDBSelectCommand
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "cd_titulo"
    cmd.Table.TableName = nomeTabela
    cmd.OrderByClause = "cd_titulo"
    Set rdGerTitReceber = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
End Sub

Public Sub MoveFirst()
    rdGerTitReceber.MoveFirst
End Sub

Public Sub MovePrevious()
    If Not rdGerTitReceber.GetRecordset.BOF Then
        rdGerTitReceber.MovePrevious
    Else
        rdGerTitReceber.MoveFirst
    End If
End Sub

Public Sub MoveNext()
    If Not rdGerTitReceber.EOF Then
        rdGerTitReceber.MoveNext
    Else
        rdGerTitReceber.MoveLast
    End If
End Sub

Public Sub MoveLast()
    rdGerTitReceber.MoveLast
End Sub

Public Property Get EOF() As Boolean
    EOF = rdGerTitReceber.EOF
End Property

Public Property Get BOF() As Boolean
    BOF = rdGerTitReceber.GetRecordset.BOF
End Property

Public Function CurrentObject() As cGeracaoTituloReceber
    Dim dao As New cGeracaoTituloReceberDAO
    If Not rdGerTitReceber.EOF Or Not rdGerTitReceber.GetRecordset.BOF Then
        Set CurrentObject = dao.carregar(rdGerTitReceber.GetLong("Cd_titulo"))
    Else
        Set CurrentObject = New cGeracaoTituloReceber
    End If
    Set dao = Nothing
End Function

Private Sub Class_Terminate()
    Set cmd = Nothing
    rdGerTitReceber.CloseReader
    Set rdGerTitReceber = Nothing
    Aplicacao.Disconnect
End Sub

Public Function FindObject(ByVal lngCodigo As Long) As cGeracaoTituloReceber
    Dim dao As New cGeracaoTituloReceberDAO
    'If Not rdGerTitPagar.EOF Or Not rdGerTitPagar.GetRecordset.BOF Then
        Set FindObject = dao.carregar(lngCodigo)
    'Else
    '    Set CurrentObject = New cGeracaoTituloPagar
    'End If
    Set dao = Nothing
End Function

