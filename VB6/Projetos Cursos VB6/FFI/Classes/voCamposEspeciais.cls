VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "voCamposEspeciais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mstrDescricao As String
Private mstrValor As String

Public Property Get Descricao() As String
    Descricao = mstrDescricao
End Property

Public Property Let Descricao(ByVal value As String)
    mstrDescricao = value
End Property

Public Property Get Valor() As String
    Valor = mstrValor
End Property

Public Property Let Valor(ByVal value As String)
    mstrValor = value
End Property

Public Function Equals(ByVal obj As voCamposEspeciais) As Boolean
    
On Error GoTo erro
    
    Equals = (obj.Descricao = Me.Descricao) And (obj.Valor = Me.Valor)
    Exit Function
    
erro:
    Equals = False

End Function


