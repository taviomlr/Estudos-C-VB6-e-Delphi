VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DaoBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mApp            As IApplication

Private Const strNomeTabela = "Bancos"

Public Function DiasLiberacao(ByVal Banco As Long) As Double
    Dim selCmd   As IDBSelectCommand
    Dim rdResult As IDBReader
        
On Error GoTo err
    Aplicacao.Connect
    If Banco > 0 Then
        Set selCmd = Aplicacao.CreateSelectCommand
        With selCmd
            .SelectClause = "[Dias para Libera��o]"
            
            .Table.TableName = strNomeTabela
            
            Call .Filter.Append("Banco = @pNumero")
            Call .Parameters.add(.CreateParameter("@pNumero", Banco, dbFieldTypeLong))
        End With
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
        If Not rdResult.EOF Then
            DiasLiberacao = rdResult.GetDouble("Dias para Libera��o")
        End If
        rdResult.CloseReader
        Set rdResult = Nothing
        Set selCmd = Nothing
    Else
        DiasLiberacao = 0
    End If
    Aplicacao.Disconnect
    Exit Function
err:
    DiasLiberacao = 0
    rdResult.CloseReader
    Set rdResult = Nothing
    Set selCmd = Nothing
        Aplicacao.Disconnect
End Function

Public Function Existe(ByVal lngBanco As String) As Boolean
    Dim selCmd          As IDBSelectCommand
    Dim rdResult        As IDBReader
               
On Error GoTo Erro
    Aplicacao.Connect
            
    Set selCmd = Aplicacao.CreateSelectCommand
    
    With selCmd
        
        .Table.TableName = strNomeTabela
        
        Call .Filter.Append("Banco = @pBanco")
        Call .Parameters.add(selCmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
    End With
    
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    
    Existe = Not rdResult.EOF
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    
    Aplicacao.Disconnect
    Exit Function
Erro:
    Existe = False
    FinallyConnection Aplicacao
End Function

'Vinicius Alexandre Elyseu - 16/01/2015 - PJ 81627
'Esta fun��o foi desenvolvida para atender a importa��o de extrato banc�rio e inserir automaticamente o banco de acordo com as boas pr�ticas atuais.
'Deixei preparado para podermos utilizar a fun��o em outros lugares, por�m ainda falta implementar o preparedCommand.
'A variavel bolImportacao deve controlar se ser� uma grava��o simples para importa��o ou uma grava��o complexa com todos os dados do banco.
    
Public Function Gravar(ByVal lngBanco As Integer, ByVal strNomeBanco As String, Optional ByVal bolImportacao As Boolean) As Boolean
    Dim cmd        As IDBInsertCommand
    
On Error GoTo erro_gravando

    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = strNomeTabela
        
        Call .AddValue("[Banco]", "@pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
    
        Call .AddValue("[Nome]", "@pNomeBanco")
        Call .Parameters.add(.CreateParameter("@pCdEst", strNomeBanco, dbFieldTypeString))
        
        If Not bolImportacao Then
            'Call PrepareCommand(cmd, obj)
        End If
    End With
    
    Gravar = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    Set cmd = Nothing
    Exit Function
erro_gravando:
    Gravar = False
    Set cmd = Nothing
    Call Throw(err)
End Function


