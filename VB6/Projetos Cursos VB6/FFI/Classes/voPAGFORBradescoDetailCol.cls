VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "voPAGFORBradescoDetailCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements FXT.IIterator

Private itens                           As Collection
Private iIndex                          As Integer

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
    
    If iIndex = 0 Then
        IIterator_EOF = True
    End If
    If iIndex > itens.Count Then
        IIterator_EOF = True
    End If
    Exit Function
FimArquivo:
    IIterator_EOF = True
End Function

Private Sub IIterator_MoveNext(bInterrompe As Boolean)
    iIndex = iIndex + 1
End Sub

Public Sub add(obj As Object)
    Dim objPAGFORBradescoDetail As voPAGFORBradescoDetail
    
    If Not Find(obj) Then
        Set objPAGFORBradescoDetail = obj
        Call itens.add(objPAGFORBradescoDetail)
    End If
    
    If itens.Count > 0 Then
        iIndex = 1
    End If
End Sub

Public Function Find(obj As Object) As Integer
    Dim i                       As Integer
    Dim objPAGFORBradescoDetail As voPAGFORBradescoDetail
    
On Error GoTo SemRegistro

    For i = 1 To itens.Count
        Set objPAGFORBradescoDetail = itens(i)
        If objPAGFORBradescoDetail.Equals(obj) Then
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
    Dim objPAGFORBradescoDetail As voPAGFORBradescoDetail
    
    i = Find(obj)
    If i > 0 Then
        Set objPAGFORBradescoDetail = itens.item(i)
        Call itens.Remove(i)
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

Public Sub MoveLast()
    iIndex = itens.Count
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

Public Property Get CurrentObject() As Object
    Set CurrentObject = IIterator_CurrentObject
End Property

Public Sub update(ByVal obj As Object)
    Dim colTemp                 As Collection
    Dim i                       As Integer
    Dim objPAGFORBradescoDetail As voPAGFORBradescoDetail
    
    Set colTemp = New Collection
    For i = 1 To itens.Count
        Set objPAGFORBradescoDetail = itens.item(i)
        If objPAGFORBradescoDetail.Equals(obj) Then
            Set objPAGFORBradescoDetail = obj
        End If
        Call colTemp.add(objPAGFORBradescoDetail)
    Next
    Set itens = colTemp
    Set colTemp = Nothing
End Sub



