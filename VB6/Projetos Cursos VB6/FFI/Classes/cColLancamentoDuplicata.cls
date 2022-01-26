VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColLancamentoDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''  Implementation of the  Class cColLancamentoDuplicata
''  Created on:      20/01/2014
''  Original author: Ueder Budni
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Implements FXT.IIterator

Private itens As Collection
Private iIndex As Integer

Private Sub Class_Initialize()
    Set itens = New Collection

End Sub

Private Function IIterator_Count() As Long
    IIterator_Count = itens.Count
End Function

Private Property Get IIterator_CurrentObject() As Object
On Error GoTo SemObjeto
    Set IIterator_CurrentObject = itens.item(iIndex)
    Exit Property

SemObjeto:
    Set IIterator_CurrentObject = Nothing
End Property

Private Function IIterator_EOF() As Boolean
On Error GoTo FimArquivo
    'Se não existe item na coleção é o fim do arquivo
    If itens.Count = 0 Then IIterator_EOF = True
    'Se o ponteiro estiver com uma quantidade maior do que a coleção
    'é o fim do arquivo
    If iIndex > itens.Count Then IIterator_EOF = True
    Exit Function

FimArquivo:
    IIterator_EOF = True
End Function

Private Sub IIterator_MoveNext(bInterrompe As Boolean)
    iIndex = iIndex + 1
End Sub

Public Sub add(ByVal obj As Object)
    Dim LancDup  As VoLancamentoDuplicata
    If Not Find(obj) Then
        Set LancDup = obj
        Call itens.add(LancDup)
    End If
    If itens.Count > 0 Then
        iIndex = 1
    End If
End Sub

Public Function Find(ByVal obj As Object) As Integer
On Error GoTo SemRegistro
    Dim i As Integer
    Dim Registro As VoLancamentoDuplicata
    
    For i = 1 To itens.Count
        Set Registro = itens(i)
        If Equals(obj, Registro) Then
            Find = i
            Exit Function
        End If
    Next
    Exit Function
SemRegistro:
    Find = 0
End Function

Public Sub Remove(ByVal obj As Object)
    Dim i As Integer
    i = Find(obj)
    If i > 0 Then
        Call itens.Remove(i)
        If itens.Count > 0 Then
            iIndex = 1
        Else
            iIndex = 0
        End If
    End If
End Sub

'Public Sub update(ByVal obj As Object)
'    Dim i As Integer
'    i = Find(obj)
'    If i > 0 Then
'        itens.item(i).Vencimento = obj.Vencimento
'        mdblTotal = mdblTotal - itens.item(i).valor
'        itens.item(i).valor = obj.valor
'        mdblTotal = mdblTotal + obj.valor
'    End If
'End Sub

Public Function EOF() As Boolean
    EOF = IIterator_EOF
End Function

Public Function Count() As Long
    Count = IIterator_Count
End Function

Public Property Get CurrentObject() As Object
    Set CurrentObject = IIterator_CurrentObject
End Property

Public Sub MoveNext(Optional bInterrompe As Boolean = False)
    Call IIterator_MoveNext(bInterrompe)
End Sub

Public Sub MoveFirst()
    iIndex = 1
End Sub

Public Function Equals(ByVal obj1 As Object, ByVal obj2 As Object) As Boolean
    Dim elem1 As VoLancamentoDuplicata
    Dim elem2 As VoLancamentoDuplicata
    
On Error GoTo erro
    
    Equals = (elem1.PagRec = elem2.PagRec) And _
             (elem1.Codigo_Nota = elem2.Codigo_Nota) And _
             (elem1.Parcela = elem2.Parcela) And _
             (elem1.Tipo = elem2.Tipo) And _
             (elem1.Empresa = elem2.Empresa)
    Exit Function
erro:
    Equals = False

End Function

'
'Public Property Get ValorTotal() As Double
'    ValorTotal = mdblTotal
'End Property

