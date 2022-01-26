VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colReajusteDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IIterator

Private mCol                                As Collection
Private mintIndex                           As Integer

Private Sub Class_Initialize()
    Set mCol = New Collection
    mintIndex = 0
End Sub

Public Sub Adicionar(ByVal obj As Object)
    If find(obj) = 0 Then
        Call add(obj)
    Else
        Call update(obj)
    End If
End Sub

Public Sub add(ByVal obj As Object)
    Dim Item         As voReajusteDuplicatas
        
    If Not find(obj) Then
        Set Item = obj
        Call mCol.add(Item)
        Set Item = Nothing
    End If
    If mCol.Count > 0 Then
        mintIndex = 1
    End If
End Sub

Public Sub update(ByVal obj As Object)
    Dim colTemp         As Collection
    Dim i               As Integer
    Dim objItem         As voReajusteDuplicatas
       
    Set colTemp = New Collection
    For i = 1 To mCol.Count
        Set objItem = mCol.Item(i)
        If objItem.equals(obj) Then
            Set objItem = obj
        End If
        Call colTemp.add(objItem)
    Next
    Set mCol = colTemp
    Set colTemp = Nothing
End Sub

Public Function find(ByVal obj As Object) As Integer
    Dim i As Integer
    Dim Item As voReajusteDuplicatas

On Error GoTo nao_encontrado
    For i = 1 To mCol.Count
        Set Item = mCol(i)
        If Item.equals(obj) Then
            find = i
            Exit Function
        End If
    Next
    Exit Function
nao_encontrado:
    find = 0
End Function

Public Sub remove(ByVal obj As Object)
    Dim i    As Integer
    Dim Item As voReajusteDuplicatas
    
    i = find(obj)
    If i > 0 Then
        Set Item = mCol.Item(i)
        Call mCol.remove(i)
        Set Item = Nothing
        If mCol.Count > 0 Then
            mintIndex = 1
        Else
            mintIndex = 0
        End If
    End If
End Sub
Private Sub IIterator_MoveNext(bInterrompe As Boolean)
    mintIndex = mintIndex + 1
End Sub

Private Function IIterator_Count() As Long
    IIterator_Count = mCol.Count
End Function

Private Property Get IIterator_CurrentObject() As Object
On Error GoTo sem_objeto
    Set IIterator_CurrentObject = mCol.Item(mintIndex)
    Exit Property
sem_objeto:
    Set IIterator_CurrentObject = Nothing
End Property

Private Function IIterator_EOF() As Boolean
On Error GoTo fimArquivo
    If mintIndex = 0 Then
        IIterator_EOF = True
    End If
    If mintIndex > mCol.Count Then
        IIterator_EOF = True
    End If
    Exit Function
fimArquivo:
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

Public Function GetItem(objReajuste As voReajusteDuplicatas) As voReajusteDuplicatas
    Dim intPos  As Integer
    
    intPos = find(objReajuste)
    If intPos > 0 Then
        Set GetItem = mCol.Item(intPos)
    End If
End Function

Private Sub IIterator_Clear()
    While Not mCol.Count = 0
        Call mCol.remove(mCol.Count)
    Wend
End Sub

Public Sub Clear()
    Call IIterator_Clear
End Sub



