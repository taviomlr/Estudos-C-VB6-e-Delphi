VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DaoCentroCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function retornaDataLimite(ByVal CentroCusto As Long) As Date
    Dim selCmd                      As IDBSelectCommand
    Dim rdResult                    As IDBReader
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .Table.TableName = "[Centros]"

        Call .Filter.Append("[Código] = @pCodigo")
        Call .Parameters.add(.CreateParameter("@pCodigo", CentroCusto, dbFieldTypeLong))

    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    
    If Not rdResult.EOF Then
        retornaDataLimite = rdResult.GetDate("Data Limite")
    End If
    
    Set selCmd = Nothing
    rdResult.CloseReader
    Set rdResult = Nothing
    Aplicacao.Disconnect
End Function

