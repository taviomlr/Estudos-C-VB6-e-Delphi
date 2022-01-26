VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "voPAGFORBradescoTrailler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mintIdentificacaoRegistro   As Integer
Private mlngQuantidadeRegistros     As Long
Private mdblTotalValores            As Double
Private mstrBrancos                 As String
Private mlngSequencialArquivo       As Long

Public Property Get IdentificacaoRegistro() As Integer
    IdentificacaoRegistro = mintIdentificacaoRegistro
End Property

Public Property Let IdentificacaoRegistro(ByVal Valor As Integer)
    mintIdentificacaoRegistro = Valor
End Property

Public Property Get QuantidadeRegistros() As Long
    QuantidadeRegistros = mlngQuantidadeRegistros
End Property

Public Property Let QuantidadeRegistros(ByVal Valor As Long)
    mlngQuantidadeRegistros = Valor
End Property

Public Property Get TotalValores() As Double
    TotalValores = mdblTotalValores
End Property

Public Property Let TotalValores(ByVal Valor As Double)
    mdblTotalValores = Valor
End Property

Public Property Get Brancos() As String
    Brancos = mstrBrancos
End Property

Public Property Let Brancos(ByVal Valor As String)
    mstrBrancos = Valor
End Property

Public Property Get SequencialArquivo() As Long
    SequencialArquivo = mlngSequencialArquivo
End Property

Public Property Let SequencialArquivo(ByVal Valor As Long)
    mlngSequencialArquivo = Valor
End Property
