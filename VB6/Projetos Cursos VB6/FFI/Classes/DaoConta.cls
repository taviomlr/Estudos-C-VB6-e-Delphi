VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DaoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function verificaContaAtiva(ByVal conta As Long) As Boolean
    Dim selCmd                      As IDBSelectCommand
    Dim rdResult                    As IDBReader
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .Table.TableName = "[Contas]"

        Call .Filter.Append("[Código] = @pCodigo")
        Call .Parameters.add(.CreateParameter("@pCodigo", conta, dbFieldTypeLong))

    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    
    If Not rdResult.EOF Then
        verificaContaAtiva = IIf(rdResult.GetString("Ctaati") = "S", True, False)
    End If
    
    Set selCmd = Nothing
    rdResult.CloseReader
    Set rdResult = Nothing
    Aplicacao.Disconnect
End Function
