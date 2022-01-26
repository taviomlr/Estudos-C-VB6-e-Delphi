VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "daoReajusteDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mstrTabela = "FFIReajusteDuplicatasReceber"
Private mobjApp As IApplication

Public Sub init(objApp As IApplication)
    Set mobjApp = objApp
End Sub

Public Function salvar(objReajuste As voReajusteDuplicatas) As Boolean
    Dim cmd        As IDBInsertCommand
    
On Error GoTo erro_gravando

    Set cmd = mobjApp.CreateInsertCommand
    With cmd
        .Table = mstrTabela
        
        Call .AddValue("[cd_reajuste]", "@pCdReajuste")
        Call .Parameters.add(.CreateParameter("@pCdReajuste", GetNextCode, dbFieldTypeLong))
        
        Call .AddValue("[PagRec]", "@pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", objReajuste.PagRec, dbFieldTypeString))
        
        Call .AddValue("[Nota]", "@pNota")
        Call .Parameters.add(.CreateParameter("@pNota", objReajuste.Nota, dbFieldTypeLong))
        
        Call .AddValue("[Empresa]", "@pEmpresa")
        Call .Parameters.add(.CreateParameter("@pEmpresa", objReajuste.Empresa, dbFieldTypeString))
        
        Call .AddValue("[Tipo]", "@pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", objReajuste.Tipo, dbFieldTypeString))
        
        Call .AddValue("[Parcela]", "@pParcela")
        Call .Parameters.add(.CreateParameter("@pParcela", objReajuste.Parcela, dbFieldTypeLong))
        
        Call .AddValue("[data_reajuste]", "@pDataReajuste")
        Call .Parameters.add(.CreateParameter("@pDataReajuste", objReajuste.data_reajuste, dbFieldTypeDate))
        
        Call .AddValue("[valor_anterior]", "@pValorAnterior")
        Call .Parameters.add(.CreateParameter("@pValorAnterior", objReajuste.valor_anterior, dbFieldTypeDouble))
        
        Call .AddValue("[valor_reajustado]", "@pValorReajustado")
        Call .Parameters.add(.CreateParameter("@pValorReajustado", objReajuste.valor_reajustado, dbFieldTypeDouble))
        
        Call .AddValue("[perc_1]", "@pPercIncc")
        Call .Parameters.add(.CreateParameter("@pPercIncc", objReajuste.perc_1, dbFieldTypeDouble))
        
        Call .AddValue("[perc_2]", "@pPercIgpm")
        Call .Parameters.add(.CreateParameter("@pPercIgpm", objReajuste.perc_2, dbFieldTypeDouble))
        
        Call .AddValue("[perc_adicionais]", "@pPercAdic")
        Call .Parameters.add(.CreateParameter("@pPercAdic", objReajuste.perc_adicionais, dbFieldTypeDouble))
        
        Call .AddValue("[usuario]", "@pUsuario")
        Call .Parameters.add(.CreateParameter("@pUsuario", objReajuste.usuario, dbFieldTypeString))
    End With
    
    salvar = (mobjApp.ExecuteUpdate(mobjApp.GetInternalAuthorization, cmd) = 1)
    
    Set cmd = Nothing
    
    Exit Function
erro_gravando:
    salvar = False
    Resume
    Set cmd = Nothing
End Function

Private Function GetNextCode() As Long
    GetNextCode = ProximoNumero("cd_reajuste", mstrTabela, Empty)
End Function
