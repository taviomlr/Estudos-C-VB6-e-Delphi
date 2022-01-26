VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "voReajusteDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mlngCdReajuste As Long
Private mstrPagRec As String
Private mlngNota As Long
Private mstrEmpresa As String
Private mstrTipo As String
Private mintParcela As Integer
Private mdtaDataReajuste As Date
Private mdblValorAnterior As Double
Private mdblValorReajustado As Double
Private mdblPercINCC As Double
Private mdblPercIGPM As Double
Private mdblPercAdicional As Double
Private mstrUsuario As String

Public Property Get cd_reajuste() As Long
    cd_reajuste = mlngCdReajuste
End Property

Public Property Let cd_reajuste(ByVal lngCdReajuste As Long)
    mlngCdReajuste = lngCdReajuste
End Property

Public Property Get PagRec() As String
    PagRec = mstrPagRec
End Property

Public Property Let PagRec(ByVal strPagRec As String)
    mstrPagRec = strPagRec
End Property

Public Property Get nota() As Long
    nota = mlngNota
End Property

Public Property Let nota(ByVal lngNota As Long)
    mlngNota = lngNota
End Property

Public Property Get Empresa() As String
    Empresa = mstrEmpresa
End Property

Public Property Let Empresa(ByVal strEmpresa As String)
    mstrEmpresa = strEmpresa
End Property

Public Property Get Tipo() As String
    Tipo = mstrTipo
End Property

Public Property Let Tipo(ByVal strTipo As String)
    mstrTipo = strTipo
End Property

Public Property Get Parcela() As Integer
    Parcela = mintParcela
End Property

Public Property Let Parcela(ByVal intParcela As Integer)
    mintParcela = intParcela
End Property

Public Property Get data_reajuste() As Date
    data_reajuste = mdtaDataReajuste
End Property

Public Property Let data_reajuste(ByVal dtaDataReajuste As Date)
    mdtaDataReajuste = dtaDataReajuste
End Property

Public Property Get valor_anterior() As Double
    valor_anterior = mdblValorAnterior
End Property

Public Property Let valor_anterior(ByVal dblValorAnterior As Double)
    mdblValorAnterior = dblValorAnterior
End Property

Public Property Get valor_reajustado() As Double
    valor_reajustado = mdblValorReajustado
End Property

Public Property Let valor_reajustado(ByVal dblValorReajustado As Double)
    mdblValorReajustado = dblValorReajustado
End Property

Public Property Get perc_1() As Double
    perc_1 = mdblPercINCC
End Property

Public Property Let perc_1(ByVal dblPercINCC As Double)
    mdblPercINCC = dblPercINCC
End Property

Public Property Get perc_2() As Double
    perc_2 = mdblPercIGPM
End Property

Public Property Let perc_2(ByVal dblPercIGPM As Double)
    mdblPercIGPM = dblPercIGPM
End Property

Public Property Get perc_adicionais() As Double
    perc_adicionais = mdblPercAdicional
End Property

Public Property Let perc_adicionais(ByVal dblPercAdicional As Double)
    mdblPercAdicional = dblPercAdicional
End Property

Public Property Get usuario() As String
    usuario = mstrUsuario
End Property

Public Property Let usuario(ByVal strUsuario As String)
    mstrUsuario = strUsuario
End Property

Public Function equals(ByVal Item As voReajusteDuplicatas) As Boolean
    equals = Item.cd_reajuste = mlngCdReajuste
    equals = equals And Item.PagRec = mstrPagRec
    equals = equals And Item.nota = mlngNota
    equals = equals And Item.Empresa = mstrEmpresa
    equals = equals And Item.Tipo = mstrTipo
    equals = equals And Item.Parcela = mintParcela
End Function

Public Function CriarObjeto(ByVal rd As IDBReader) As Boolean

On Error GoTo err
    With rd
        mlngCdReajuste = .GetLong("cd_reajuste")
        mstrPagRec = .GetString("PagRec")
        mlngNota = .GetLong("Nota")
        mstrEmpresa = .GetString("Empresa")
        mstrTipo = .GetString("Tipo")
        mintParcela = .GetLong("Parcela")
        mdtaDataReajuste = .GetDate("data_reajuste")
        mdblValorAnterior = .GetDouble("valor_anterior")
        mdblValorReajustado = .GetDouble("valor_reajustado")
        If Not IsNull(.GetDouble("perc_1")) Then
            mdblPercINCC = .GetDouble("perc_1")
        End If
        If Not IsNull(.GetDouble("perc_2")) Then
            mdblPercIGPM = .GetDouble("perc_2")
        End If
        If Not IsNull(.GetDouble("perc_adicionais")) Then
            mdblPercAdicional = .GetDouble("perc_adicionais")
        End If
        mstrUsuario = .GetString("usuario")
    End With
    CriarObjeto = True
    Exit Function
err:
    CriarObjeto = False
End Function

