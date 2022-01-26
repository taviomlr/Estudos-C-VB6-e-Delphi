VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeracaoTituloPagarNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private rdGerTitPagar As IDBReader
Private Const nomeTabela$ = "FFITituloPagar"

Private Sub Class_Initialize()
    Dim cmd As IDBSelectCommand
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "cd_titulo"
    cmd.Table.TableName = nomeTabela
    cmd.OrderByClause = "cd_titulo"
    Set rdGerTitPagar = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
End Sub

Public Sub MoveFirst()
    rdGerTitPagar.MoveFirst
End Sub

Public Sub MovePrevious()
    If Not rdGerTitPagar.GetRecordset.BOF Then
        rdGerTitPagar.MovePrevious
    Else
        rdGerTitPagar.MoveFirst
    End If
End Sub

Public Sub MoveNext()
    If Not rdGerTitPagar.EOF Then
        rdGerTitPagar.MoveNext
    Else
        rdGerTitPagar.MoveLast
    End If
End Sub

Public Sub MoveLast()
    rdGerTitPagar.MoveLast
End Sub

Public Property Get EOF() As Boolean
    EOF = rdGerTitPagar.EOF
End Property

Public Property Get BOF() As Boolean
    BOF = rdGerTitPagar.GetRecordset.BOF
End Property

Public Function CurrentObject() As cGeracaoTituloPagar
    Dim dao As New cGeracaoTituloPagarDAO
    If Not rdGerTitPagar.EOF Or Not rdGerTitPagar.GetRecordset.BOF Then
        Set CurrentObject = dao.carregar(rdGerTitPagar.GetLong("Cd_titulo"))
    Else
        Set CurrentObject = New cGeracaoTituloPagar
    End If
    Set dao = Nothing
End Function

Private Sub Class_Terminate()
    Set cmd = Nothing
    rdGerTitPagar.CloseReader
    Set rdGerTitPagar = Nothing
    Aplicacao.Disconnect
End Sub

Public Function FindObject(ByVal lngCodigo As Long) As cGeracaoTituloPagar
    Dim dao As New cGeracaoTituloPagarDAO
    'If Not rdGerTitPagar.EOF Or Not rdGerTitPagar.GetRecordset.BOF Then
        Set FindObject = dao.carregar(lngCodigo)
    'Else
    '    Set CurrentObject = New cGeracaoTituloPagar
    'End If
    Set dao = Nothing
End Function
