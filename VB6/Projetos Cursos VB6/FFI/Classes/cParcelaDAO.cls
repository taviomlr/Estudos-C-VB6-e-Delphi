VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParcelaDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const nomeTabela$ = "Duplicatas"

Public Function carregar_lista(objDuplicata As CDuplicata) As cColecaoParcela
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Dim objParcela As cParcela
On Error GoTo erro_carregando
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = nomeTabela
    Call cmd.Filter.Append("pagrec = @pPagRec")
    Call cmd.Parameters.add(cmd.CreateParameter("@pPagRec", objDuplicata.PagRec, dbFieldTypeString, 1))
    Call cmd.Filter.Append("Nota = @pNota")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNota", objDuplicata.nota, dbFieldTypeLong))
    Call cmd.Filter.Append("Empresa = @pEmpresa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", objDuplicata.Empresa, dbFieldTypeString, 15))
    Call cmd.Filter.Append("tipo = @pTipo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipo", objDuplicata.tipo, dbFieldTypeString, 20))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set carregar_lista = New cColecaoParcela
        While Not rdResult.EOF
            Set objParcela = New cParcela
            Call carregar_lista.add(objParcela.CriarObjeto(rdResult))
            rdResult.MoveNext
        Wend
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
    Exit Function
erro_carregando:
    Err.Clear
    Set carregar_lista = Nothing
End Function
