VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColecaoRateioTituloReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements FXT.IIterator

Private itens As Collection
Private iIndex As Integer
'Private mdblVolume As Double
Private mdblTotalValor As Double

Public Sub Class_Initialize()
    Set itens = New Collection
    iIndex = 0
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
    If iIndex = 0 Then IIterator_EOF = True
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

Public Sub add(obj As Object)
    Dim merc As cGeracaoTituloReceber
    If Not Find(obj) Then
        Set merc = obj
        Call itens.add(merc)
        'mdblVolume = mdblVolume + merc.Quantidade
        mdblTotalValor = mdblTotalValor + merc.R_Percentual
    End If
    If itens.Count > 0 Then
        iIndex = 1
    End If
End Sub

Public Function Find(obj As Object) As Integer
On Error GoTo SemRegistro
    Dim i As Integer
    Dim merc As cGeracaoTituloReceber
    
    For i = 1 To itens.Count
        Set merc = itens(i)
        If merc.Equals(obj) Then
            Find = i
            Exit Function
        End If
    Next
    Exit Function
SemRegistro:
    Find = 0
End Function

Public Sub Remove(obj As Object)
    Dim i As Integer
    i = Find(obj)
    If i > 0 Then
        Call itens.Remove(i)
        'mdblVolume = mdblVolume - obj.Quantidade
        mdblTotalValor = mdblTotalValor - obj.R_Percentual
        If itens.Count > 0 Then
            iIndex = 1
        Else
            iIndex = 0
        End If
    End If
End Sub

Public Sub MoveFirst()
    iIndex = 1
End Sub

Public Sub MoveNext(Optional bInterrompe As Boolean = False)
    Call IIterator_MoveNext(bInterrompe)
End Sub

Public Function EOF() As Boolean
    EOF = IIterator_EOF
End Function

Public Function Count() As Long
    Count = IIterator_Count
End Function

'Public Property Get volume() As Double
'    volume = mdblVolume
'End Property

Public Property Get totalValor() As Double
    totalValor = mdblTotalValor
End Property

Public Property Get CurrentObject() As Object
    Set CurrentObject = IIterator_CurrentObject
End Property
