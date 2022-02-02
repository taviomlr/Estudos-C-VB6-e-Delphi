VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFretePagarNotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''
''  Implementation of the  Class cFretePagarNotaFiscal
''  Generated by Enterprise Architect
''  Created on:      12-set-2006 11:46:28
''  Original author: Dulcino J�nior
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private mlngNota As Long
Private mstrTipo As String
Private mdatEmissao As Date
Private mstrApel As String
Private mstrEmpresa As String
Private mlngTransportadora As Long
Private mstrNomeTransportadora As String
Private mdblValor As Double
Private mbooEntrada As Boolean
Private mstrSitu As String

Public Property Get entrada() As String
        entrada = mbooEntrada
End Property

Public Property Get Empresa() As String
        Empresa = mstrEmpresa
End Property

Public Property Let Empresa(ByVal NewVal As String)
        mstrEmpresa = NewVal
End Property

Public Property Get Emissao() As Date
        Emissao = mdatEmissao
End Property

Public Property Let Emissao(ByVal NewVal As Date)
        mdatEmissao = NewVal
End Property

Public Property Get Nota() As Long
        Nota = mlngNota
End Property

Public Property Let Nota(ByVal NewVal As Long)
        mlngNota = NewVal
End Property

Public Property Get tipo() As String
        tipo = mstrTipo
End Property

Public Property Let tipo(ByVal NewVal As String)
        mstrTipo = NewVal
End Property

Public Property Get valor() As Double
        valor = mdblValor
End Property

Public Property Let valor(ByVal NewVal As Double)
        mdblValor = NewVal
End Property

Public Property Get Apel() As String
    Apel = mstrApel
End Property

Public Property Let Apel(ByVal NewVal As String)
    mstrApel = NewVal
End Property

Public Property Get Transportadora() As Long
    Transportadora = mlngTransportadora
End Property

Public Property Let Transportadora(ByVal NewVal As Long)
    mlngTransportadora = NewVal
End Property

Public Property Get nomeTransportadora() As String
    nomeTransportadora = mstrNomeTransportadora
End Property

Public Property Let nomeTransportadora(ByVal NewVal As String)
    mstrNomeTransportadora = NewVal
End Property

Public Property Get Situacao() As String
        Situacao = mstrSitu
End Property

Public Property Let Situacao(ByVal NewVal As String)
        mstrSitu = NewVal
End Property

Public Function equals(ByVal obj As Object) As Boolean
    Dim objNotaPagar As cFretePagarNotaFiscal
On Error GoTo erro_comparando
    Set objNotaPagar = obj
    equals = (equals Or objNotaPagar.Nota = mlngNota)
    equals = (equals And objNotaPagar.tipo = mstrTipo)
    equals = (equals And objNotaPagar.Emissao = mdatEmissao)
    Exit Function
erro_comparando:
    equals = False
End Function

Public Function CriarObjeto(rd As IDBReader) As cFretePagarNotaFiscal
    Dim strCriterio As String
    Dim strTabela As String
    With rd
        mlngNota = .GetLong("nr_nota")
        mstrTipo = .GetString("tp_registro_nf")
        mlngTransportadora = .GetLong("cd_transportadora")
        mstrApel = .GetString("empresa")
        mstrNomeTransportadora = GetFieldValue("Raz�o", "Transportadoras", "C�digo = " & mlngTransportadora)
        mstrEmpresa = GetFieldValue("Raz�o", "Empresas", "Apel = '" & mstrApel & "'")
        strCriterio = "[N�mero] = " & mlngNota & " AND [Tipo de Registro] = '" & mstrTipo & "'"
        If .GetBoolean("registro_entrada") Then
            strTabela = "[Notas Fiscais de Entrada]"
            strCriterio = strCriterio & " AND Fornecedor = '" & mstrApel & "'"
        Else
            strTabela = "[Notas Fiscais de Sa�da]"
            strCriterio = strCriterio & " AND Fornecedor = '" & Left(DonaSistema, 15) & "'"
        End If
        mdatEmissao = GetFieldValue("[Emiss�o]", strTabela, strCriterio, , 0)
        mdblValor = GetFieldValue("[Valor Total]", strTabela, strCriterio, , 0)
        mbooEntrada = .GetBoolean("registro_entrada")
        mstrSitu = GetFieldValue("[Situa��o]", strTabela, strCriterio, , 0)
    End With
    Set CriarObjeto = Me
End Function