VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColecaoSinaisNegocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements FXT.IIterator

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''  Classe cColecaoPedidosVendaItens
''  Criado em: 01/06/2007
''  Autor: Dulcino Júnior
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private mcolSinaisNegocio As Collection
Private mintIndex         As Integer
Private mcurTotalSinais   As Double

Private Sub Class_Initialize()
    Set mcolSinaisNegocio = New Collection
    mintIndex = 0
    mcurTotalSinais = 0
End Sub

Private Function IIterator_Count() As Long
    IIterator_Count = mcolSinaisNegocio.Count
End Function

Private Property Get IIterator_CurrentObject() As Object
    Set IIterator_CurrentObject = mcolSinaisNegocio.item(mintIndex)
End Property

Private Function IIterator_EOF() As Boolean
On Error GoTo FimArquivo
    'Se não existe item na coleção é o fim do arquivo
    If mintIndex = 0 Then IIterator_EOF = True
    'Se o ponteiro estiver com uma quantidade maior do que a coleção
    'é o fim do arquivo
    If mintIndex > mcolSinaisNegocio.Count Then IIterator_EOF = True
    Exit Function

FimArquivo:
    IIterator_EOF = True
End Function

Private Sub IIterator_MoveNext(bInterrompe As Boolean)
    mintIndex = mintIndex + 1
End Sub

Public Sub add(obj As Object)
    Dim item As cSinalNegocio

    If Not Find(obj) Then
        Set item = obj
        Call mcolSinaisNegocio.add(item)
        mcurTotalSinais = mcurTotalSinais + item.valorParcela
        Set item = Nothing
    End If
    If mcolSinaisNegocio.Count > 0 Then
        mintIndex = 1
    End If
End Sub

Public Function Find(ByVal obj As Object) As Integer
    Dim i As Integer
    Dim item As cSinalNegocio

On Error GoTo nao_encontrado
    For i = 1 To mcolSinaisNegocio.Count
        Set item = mcolSinaisNegocio(i)
        If item.Equals(obj) Then
            Find = i
            Exit Function
        End If
    Next
    Exit Function

nao_encontrado:
    Find = 0
End Function

Public Sub Remove(ByVal obj As Object)
    Dim i    As Integer
    Dim item As cSinalNegocio

    i = Find(obj)
    If i > 0 Then
        Set item = mcolSinaisNegocio.item(i)
        Call mcolSinaisNegocio.Remove(i)
        mcurTotalSinais = mcurTotalSinais - item.valorParcela
        Set item = Nothing
        If mcolSinaisNegocio.Count > 0 Then
            mintIndex = 1
        Else
            mintIndex = 0
        End If
    End If
End Sub

Public Sub update(ByVal obj As Object)
    Dim colTemp As Collection
    Dim i       As Integer
    Dim objItem As cSinalNegocio
    
    Set colTemp = New Collection
    mcurTotalSinais = 0
    For i = 1 To mcolSinaisNegocio.Count
        Set objItem = mcolSinaisNegocio.item(i)
        If objItem.Equals(obj) Then
            Set objItem = obj
        End If
        Call colTemp.add(objItem)
        mcurTotalSinais = mcurTotalSinais + objItem.valorParcela
    Next
    Set mcolSinaisNegocio = colTemp
    Set colTemp = Nothing
End Sub

Public Function GetItem(lngPedido As Long, strTipo As String, strFornecedor As String, intParcela As Integer) As cSinalNegocio
    Dim objResult As New cSinalNegocio
    Dim intIndex  As Integer
    
    objResult.numeroParcela = intParcela
    objResult.Fornecedor = strFornecedor
    objResult.numeroPedido = lngPedido
    objResult.TipoRegistro = strTipo
    intIndex = Find(objResult)
    If intIndex > 0 Then
        Set GetItem = mcolSinaisNegocio.item(intIndex)
    End If
    Set objResult = Nothing
End Function

Public Property Get TotalSinais() As Double
    TotalSinais = mcurTotalSinais
End Property

Public Sub CarregaSinais(lngPedido As Long, strTipo As String)
    Dim selCmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Dim objResult As cSinalNegocio

On Error GoTo error_handler
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "NUMPED, TIPREG, NUMPAR"
        
        .Table.TableName = "SinaisNegocio"
        .OrderByClause = "NUMPAR desc"
        Call .Filter.Append("NUMPED = @pNumPedido")
        Call .Parameters.add(.CreateParameter("@pNumPedido", lngPedido, dbFieldTypeLong))
        
        Call .Filter.Append("TIPREG = @pTipoRegistro")
        Call .Parameters.add(.CreateParameter("@pTipoRegistro", strTipo, dbFieldTypeString))
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    If Not rdResult.EOF Then
        While Not rdResult.EOF
            Set objResult = New cSinalNegocio
            Call objResult.Carregar(rdResult.GetLong("NUMPED"), rdResult.GetString("TIPREG"), rdResult.GetInteger("NUMPAR"))
            Call add(objResult)
            rdResult.MoveNext
        Wend
    End If
    Aplicacao.Disconnect
    
    Exit Sub
error_handler:
    FinallyConnection Aplicacao
    Call Throw(err)
End Sub

Public Function Count() As Long
    Count = IIterator_Count
End Function

Public Sub MoveFirst()
    mintIndex = 1
End Sub

Public Sub MoveNext()
    Call IIterator_MoveNext(False)
End Sub

Public Function EOF() As Boolean
    EOF = IIterator_EOF
End Function

Public Property Get CurrentObject() As cSinalNegocio
    Set CurrentObject = IIterator_CurrentObject
End Property

'Data.......: 05/06/2007
'Autor......: Dulcino Júnior
'Descrição..: Função que retorna uma nova instancia do objeto coleção de sinais
'               com os mesmos valores.
'Retorno....: [cColecaoSinaisNegocio] Copia do objeto coleção de sinais de negócio
'               com os mesmos valores.
Public Function Clone() As cColecaoSinaisNegocio
    Set Clone = New cColecaoSinaisNegocio
        
    Me.MoveFirst
    While Not Me.EOF
        Call Clone.add(CurrentObject.Clone)
        Me.MoveNext
    Wend
End Function

Private Sub Class_Terminate()
    Set mcolSinaisNegocio = Nothing
End Sub
