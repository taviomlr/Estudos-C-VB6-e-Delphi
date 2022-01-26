VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarteiraCpEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngEnterprise_id                       As Long
Private mlngCd_estabelecimento                  As Long
Private mlngId_carteira                         As Long
Private mstrCpEspNome                           As String
Private mstrCpEspValor                          As String

Public Property Let Enterprise_id(valor As Long)
   mlngEnterprise_id = valor
End Property

Public Property Get Enterprise_id() As Long
   Enterprise_id = mlngEnterprise_id
End Property

Public Property Let Cd_estabelecimento(valor As Long)
   mlngCd_estabelecimento = valor
End Property

Public Property Get Cd_estabelecimento() As Long
   Cd_estabelecimento = mlngCd_estabelecimento
End Property

Public Property Let Id_carteira(valor As Long)
   mlngId_carteira = valor
End Property

Public Property Get Id_carteira() As Long
   Id_carteira = mlngId_carteira
End Property

Public Property Get CpEspNome() As String
    CpEspNome = mstrCpEspNome
End Property

Public Property Let CpEspNome(ByVal valor As String)
    mstrCpEspNome = valor
End Property

Public Property Get CpEspValor() As String
    CpEspValor = mstrCpEspValor
End Property

Public Property Let CpEspValor(ByVal valor As String)
    mstrCpEspValor = valor
End Property

Public Function equals(obj As Object) As Boolean
    Dim booIgual As Boolean
    Dim newObj As clsCarteiraCpEspecial
    
    Set newObj = obj
    booIgual = (newObj.CpEspNome = mstrCpEspNome)
    equals = booIgual
End Function

Public Function CriarObjeto(ByVal rd As IDBReader) As clsCarteiraCpEspecial
    With rd
        mlngEnterprise_id = .GetLong("enterprise_id")
        mlngCd_estabelecimento = .GetLong("cd_estabelecimento")
        mlngId_carteira = .GetLong("id_carteira")
        mstrCpEspNome = .GetString("nome")
        CpEspValor = .GetString("valor")
    End With
    Set CriarObjeto = Me
End Function

