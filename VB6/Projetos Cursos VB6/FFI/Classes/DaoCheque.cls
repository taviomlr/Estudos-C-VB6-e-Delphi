VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DaoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mApp                            As IApplication

Public Sub init(ByVal objApp As IApplication)
    Set mApp = objApp
End Sub

Public Function ProximoCheque(ByVal Banco As Long) As Long
    Dim selCmd                      As IDBSelectCommand
    Dim rdResult                    As IDBReader
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .Table.TableName = "[Cheque]"

        Call .Filter.Append("[Situação] = @pCodigo")
        Call .Parameters.add(.CreateParameter("@pCodigo", "Normal", dbFieldTypeString))
        
        Call .Filter.Append("(Cheque not in (SELECT Cheque FROM Duplicatas WHERE Banco = Cheque.Banco)")
        
        Call .Filter.Append("Cheque NOT IN (SELECT Cheque FROM Lançamentos WHERE Banco = Cheque.Banco)")
        
        Call .Filter.Append("Cheque NOT IN (SELECT Cheque FROM [Transf Bancária] WHERE Banco = Cheque.Banco))")
        
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", Banco, dbFieldTypeLong))
        
        .OrderByClause = "Cheque"
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    
    If Not rdResult.EOF Then
        ProximoCheque = rdResult.GetLong("Cheque")
    End If
    
    Set selCmd = Nothing
    rdResult.CloseReader
    Set rdResult = Nothing
    Aplicacao.Disconnect
End Function

Public Function Carregar(ByVal lngBanco As Long, ByVal lngCheque As Long) As VoCheque
    Dim rdResult            As IDBReader
    Dim cmd                 As IDBSelectCommand
    Dim result              As VoCheque

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = "[cheque]"
        
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(cmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
            
        Call .Filter.Append("[Cheque] = @pCheque")
        Call .Parameters.add(.CreateParameter("@pCheque", lngCheque, dbFieldTypeLong))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set Carregar = CriarObjeto(rdResult)
    Else
        Set Carregar = Nothing
    End If
    
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function
erro_carregando:
    Set Carregar = Nothing
    Set cmd = Nothing
    Set rdResult = Nothing
End Function

Public Function Gravar(ByVal obj As VoCheque) As Boolean
    Dim cmd        As IDBInsertCommand
    Dim bolGravado As Boolean
    
On Error GoTo erro_gravando
    
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = "[Cheque]"
        
        Call .AddValue("[Banco]", "@pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", obj.Banco, dbFieldTypeLong))
    
        Call .AddValue("[Cheque]", "@pCheque")
        Call .Parameters.add(.CreateParameter("@pCheque", obj.Cheque, dbFieldTypeLong))

        Call .AddValue("[Situação]", "@pSituacao")
        Call .Parameters.add(.CreateParameter("@pSituacao", obj.Situacao, dbFieldTypeString))

        Call .AddValue("[Nominal]", "@pNominal")
        Call .Parameters.add(.CreateParameter("@pNominal", obj.Nominal, dbFieldTypeString))
                
        Call .AddValue("[Histórico]", "@pHistorico")
        Call .Parameters.add(.CreateParameter("@pHistorico", obj.Historico, dbFieldTypeString))
    End With
    
    bolGravado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    Gravar = bolGravado
    
    Set cmd = Nothing
    Exit Function
erro_gravando:
    Gravar = False
    Set cmd = Nothing
    Call Throw(Err)
End Function

Public Function Atualizar(ByVal obj As VoCheque) As Boolean
    Dim cmd        As IDBUpdateCommand
    Dim blnGravado As Boolean
    
On Error GoTo erro_gravando
    Set cmd = mApp.CreateUpdateCommand
    With cmd
        .Table = "[cheque]"
        
        
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(cmd.CreateParameter("@pBanco", obj.Banco, dbFieldTypeLong))
            
        Call .Filter.Append("[Cheque] = @pCheque")
        Call .Parameters.add(.CreateParameter("@pCheque", obj.Cheque, dbFieldTypeLong))
        
        Call .AddValue("[Situação]", "@pSituacao")
        Call .Parameters.add(.CreateParameter("@pSituacao", obj.Situacao, dbFieldTypeString))

        Call .AddValue("[Nominal]", "@pNominal")
        Call .Parameters.add(.CreateParameter("@pNominal", obj.Nominal, dbFieldTypeString))
                
        Call .AddValue("[Histórico]", "@pHistorico")
        Call .Parameters.add(.CreateParameter("@pHistorico", obj.Historico, dbFieldTypeString))
    End With
    
    blnGravado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    Atualizar = blnGravado
    
    Set cmd = Nothing
    
    Exit Function
erro_gravando:
    Atualizar = False
    Set cmd = Nothing
    Call Throw(Err)
    Resume
End Function

Public Function Excluir(ByVal lngBanco As Long, ByVal lngCheque As Long) As Boolean
    Dim cmd                             As IDBDeleteCommand
    Dim blnExcluido                     As Boolean
    
On Error GoTo erro_excluindo
    
    blnExcluido = True
    
    Set cmd = mApp.CreateDeleteCommand
    With cmd
        .Table = "[cheque]"
    
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(cmd.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
            
        Call .Filter.Append("[Cheque] = @pCheque")
        Call .Parameters.add(.CreateParameter("@pCheque", lngCheque, dbFieldTypeLong))
    End With
        
    blnExcluido = CBool(mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = CLng(1))
        
    Excluir = blnExcluido
    
    Set cmd = Nothing
    Exit Function
erro_excluindo:
    Set cmd = Nothing
    Excluir = False
    Call Throw(Err)
End Function

Public Function CriarObjeto(ByVal rd As IDBReader) As VoCheque
    Dim obj As New VoCheque
    
    With rd
        obj.Banco = .GetLong("Banco")
        obj.Cheque = .GetLong("Cheque")
        obj.Situacao = .GetString("Situação")
        obj.Nominal = .GetString("Nominal")
        obj.Historico = .GetString("Histórico")
    End With
    Set CriarObjeto = obj
End Function

