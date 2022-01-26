VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colCamposEspeciais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements FXT.IIterator

Private mCol                                As Collection
Private mintIndex                           As Integer

Private Sub Class_Initialize()
    Set mCol = New Collection
    mintIndex = 0
End Sub

Public Sub Adicionar(ByVal obj As Object)
    If Find(obj) = 0 Then
        Call add(obj)
    Else
        Call update(obj)
    End If
End Sub

Public Sub add(ByVal obj As Object)
    Dim item         As voCamposEspeciais
        
    If Not Find(obj) Then
        Set item = obj
        Call mCol.add(item)
        Set item = Nothing
    End If
    If mCol.Count > 0 Then
        mintIndex = 1
    End If
End Sub

Public Sub update(ByVal obj As Object)
    Dim colTemp         As Collection
    Dim i               As Integer
    Dim objItem         As voCamposEspeciais
       
    Set colTemp = New Collection
    For i = 1 To mCol.Count
        Set objItem = mCol.item(i)
        If objItem.Equals(obj) Then
            Set objItem = obj
        End If
        Call colTemp.add(objItem)
    Next
    Set mCol = colTemp
    Set colTemp = Nothing
End Sub

Public Function Find(ByVal obj As Object) As Integer
    Dim i As Integer
    Dim item As voCamposEspeciais

On Error GoTo nao_encontrado
    For i = 1 To mCol.Count
        Set item = mCol(i)
        If item.Equals(obj) Then
            Find = i
            Exit Function
        End If
    Next
    Exit Function
nao_encontrado:
    Find = 0
End Function

Private Sub IIterator_MoveNext(bInterrompe As Boolean)
    mintIndex = mintIndex + 1
End Sub

Private Function IIterator_Count() As Long
    IIterator_Count = mCol.Count
End Function

Private Property Get IIterator_CurrentObject() As Object
On Error GoTo sem_objeto
    Set IIterator_CurrentObject = mCol.item(mintIndex)
    Exit Property
sem_objeto:
    Set IIterator_CurrentObject = Nothing
End Property

Private Function IIterator_EOF() As Boolean
On Error GoTo FimArquivo
    'Se não existe item na coleção é o fim do arquivo
    If mintIndex = 0 Then IIterator_EOF = True
    'Se o ponteiro estiver com uma quantidade maior do que a coleção
    'é o fim do arquivo
    If mintIndex > mCol.Count Then IIterator_EOF = True
    Exit Function
FimArquivo:
    IIterator_EOF = True
End Function

Public Sub MoveNext(Optional bInterrompe As Boolean = False)
    Call IIterator_MoveNext(bInterrompe)
End Sub

Public Function EOF() As Boolean
    EOF = IIterator_EOF
End Function

Public Function Count() As Long
    Count = IIterator_Count
End Function

Public Property Get CurrentObject() As Object
    Set CurrentObject = IIterator_CurrentObject
End Property

Public Sub MoveFirst()
    mintIndex = 1
End Sub


