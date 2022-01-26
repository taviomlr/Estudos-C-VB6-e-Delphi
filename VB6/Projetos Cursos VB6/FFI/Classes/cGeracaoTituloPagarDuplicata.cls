VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeracaoTituloPagarDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngCd_titulo               As Integer
Private mstrPagRec                  As String
Private mintNumero_nota             As Integer
Private mstrTipo_registro           As String
Private mstrEmpresa                 As String
Private mintParcela                 As String

Public Property Let Cd_titulo(ByVal NewVal As Long)
    mlngCd_titulo = NewVal
End Property

Public Property Get Cd_titulo() As Long
    Cd_titulo = mlngCd_titulo
End Property

Public Property Let PagRec(ByVal NewVal As String)
    mstrPagRec = NewVal
End Property

Public Property Get PagRec() As String
    PagRec = mstrPagRec
End Property

Public Property Let Numero_nota(ByVal NewVal As Integer)
    mintNumero_nota = NewVal
End Property

Public Property Get Numero_nota() As Integer
    Numero_nota = mintNumero_nota
End Property

Public Property Let Tipo_registro(ByVal NewVal As String)
    mstrTipo_registro = NewVal
End Property

Public Property Get Tipo_registro() As String
    Tipo_registro = mstrTipo_registro
End Property

Public Property Let Empresa(ByVal NewVal As String)
    mstrEmpresa = NewVal
End Property

Public Property Get Empresa() As String
    Empresa = mstrEmpresa
End Property

Public Property Let Parcela(ByVal NewVal As String)
    mintParcela = NewVal
End Property

Public Property Get Parcela() As String
    Parcela = mintParcela
End Property
