VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFreteReceberNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private rdConhecimentos As IDBReader

Private Sub Class_Initialize()
    Dim cmd As IDBSelectCommand
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "nr_conhecimento, seq_conhecimento"
    cmd.Table.TableName = "FreteSaida"
    cmd.OrderByClause = "nr_conhecimento"
    Set rdConhecimentos = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
End Sub

Public Sub MoveFirst()
    rdConhecimentos.MoveFirst
End Sub

Public Sub MovePrevious()
    If Not rdConhecimentos.GetRecordset.BOF Then
        rdConhecimentos.MovePrevious
    Else
        rdConhecimentos.MoveFirst
    End If
End Sub

Public Sub MoveNext()
    If Not rdConhecimentos.EOF Then
        rdConhecimentos.MoveNext
    Else
        rdConhecimentos.MoveLast
    End If
End Sub

Public Sub MoveLast()
    rdConhecimentos.MoveLast
End Sub

Public Property Get EOF() As Boolean
    EOF = rdConhecimentos.EOF
End Property

Public Property Get BOF() As Boolean
    BOF = rdConhecimentos.GetRecordset.BOF
End Property

Public Function CurrentObject() As cFreteReceber
    Dim dao As New cFreteReceberDAO
    If Not rdConhecimentos.EOF Or Not rdConhecimentos.GetRecordset.BOF Then
        ' 30/11/2018 - D. 234266 - H. 240726 - Yuji F. - Correção na passagem do parâmetro; é esperado long
        Set CurrentObject = dao.Carregar(rdConhecimentos.GetLong("nr_conhecimento"), , rdConhecimentos.GetLong("seq_conhecimento"))
    Else
        Set CurrentObject = New cFreteReceber
    End If
    Set dao = Nothing
End Function

Private Sub Class_Terminate()
    rdConhecimentos.CloseReader
    Set rdConhecimentos = Nothing
    Aplicacao.Disconnect
End Sub
