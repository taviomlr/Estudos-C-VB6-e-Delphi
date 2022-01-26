VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "voCamposEspeciaisDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Projeto: 100340 - Demanda: 120997
'---|   Davi Brito (09/05/2016)
'---|---------------------------------------------------------------------------------------------------------------------------

Private mApp                As IApplication

Private Const TABLE_NAME = "FFICamposEspeciais"

Public Sub init(App As IApplication)
    Set mApp = App
End Sub

Public Sub Initialize()
    If (mApp Is Nothing) Then
        Aplicacao.Connect
        Set mApp = Aplicacao
    End If
End Sub

Public Sub Terminate()
    If (Not mApp Is Nothing) Then
        Set mApp = Nothing
        Aplicacao.Disconnect
    End If
End Sub

Public Function Carregar() As colCamposEspeciais
    Dim selCmd          As IDBSelectCommand
    Dim rdResult        As IDBReader
    Dim item            As voCamposEspeciais
    Dim colecao         As colCamposEspeciais
           
On Error GoTo erro
    mApp.Connect
            
    Set selCmd = mApp.CreateSelectCommand
    
    With selCmd
        .Table.TableName = TABLE_NAME
        
        Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, selCmd)
        If Not rdResult Is Nothing Then
        rdResult.MoveFirst
            If Not rdResult.EOF Then
                Set colecao = New colCamposEspeciais
                rdResult.MoveFirst
                While Not rdResult.EOF
                    Set item = New voCamposEspeciais
                    item.Descricao = rdResult.GetString("descricao")
                    item.Valor = rdResult.GetString("valor")
                    
                    colecao.add (item)
                Wend
                Carregar = colecao
            Else
                Set Carregar = Nothing
            End If
        End If
    End With
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    
    mApp.Disconnect
    Exit Function
erro:
    Carregar = Nothing
    FinallyConnection mApp
End Function

Private Function Inserir(ByVal obj As voCamposEspeciais) As Boolean
    Dim cmd        As IDBInsertCommand
    
On Error GoTo erro_gravando
    mApp.Connect
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = TABLE_NAME
     
        Call .AddValue("[descricao]", "@pDescricao")
        Call .Parameters.add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString))
        
        Call .AddValue("[valor]", "@pValor")
        Call .Parameters.add(.CreateParameter("@pValor", obj.Valor, dbFieldTypeString))
                
    End With
    
    Inserir = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    mApp.Disconnect
    Set cmd = Nothing
    Exit Function
    
erro_gravando:
    Inserir = False
    Set cmd = Nothing
    Call Throw(err)

End Function

Public Function ExisteCamposEspecial(ByVal obj As voCamposEspeciais) As Boolean
    Dim selCmd As IDBSelectCommand
    Dim rdResult As IDBReader

On Error GoTo erro
        Aplicacao.Connect
            
        Set selCmd = Aplicacao.CreateSelectCommand
    
        With selCmd
            .Table.TableName = TABLE_NAME
        
            Call .Filter.Append("[descricao]", "@pDescricao")
            Call .Parameters.add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString))
        
            Call .Filter.Append("[valor]", "@pValor")
            Call .Parameters.add(.CreateParameter("@pValor", obj.Valor, dbFieldTypeString))
        
                
            Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
            ExisteCamposEspecial = rdResult.Recordcount > 0
        End With
    
        Set rdResult = Nothing
        Set selCmd = Nothing
    
        Aplicacao.Disconnect
        Exit Function
        
erro:
        FinallyConnection Aplicacao
        ExisteCamposEspecial = False
       
        Set rdResult = Nothing
        Set selCmd = Nothing
End Function

Public Function Gravar(ByVal obj As voCamposEspeciais) As Boolean
    Dim cmd        As IDBUpdateCommand
    
On Error GoTo erro
    
    mApp.Connect
    mApp.BeginTransaction
    
    If ExisteCamposEspecial(obj) Then
        If Excluir(obj) Then
            Call Me.Inserir(obj)
        End If
    Else
        Call Me.Inserir(obj)
    End If
    
    
    mApp.CommitTransaction
    mApp.Disconnect
    
    Set cmd = Nothing
    Exit Function
    
erro:
    mApp.RollbackTransaction
    mApp.Disconnect
    
    Gravar = False
    Set cmd = Nothing
    Call Throw(err)

End Function

Public Function Excluir(ByVal obj As voCamposEspeciais) As Boolean
    Dim cmd        As IDBDeleteCommand

On Error GoTo erro_excluindo
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateDeleteCommand
    With cmd
        .Table = TABLE_NAME
    
        Call .Filter.Append("[descricao]", "@pDescricao")
        Call .Parameters.add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString))
        
        Call .Filter.Append("[valor]", "@pValor")
        Call .Parameters.add(.CreateParameter("@pValor", obj.Valor, dbFieldTypeString))
    End With
    
    Call Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd)
    ExcluiVO = True
    Aplicacao.Disconnect
    Set cmd = Nothing
    
    Exit Function
    
erro_excluindo:
    Set cmd = Nothing
    ExcluiVO = False
End Function

Public Function ExisteReg(ByVal strTipoLancamento As String, ByVal lngCdTipoLanc As Long) As Boolean
    Dim selCmd          As IDBSelectCommand
    Dim rdResult        As IDBReader
           
On Error GoTo erro
    mApp.Connect
            
    Set selCmd = mApp.CreateSelectCommand
    With selCmd
        .Table.TableName = TABLE_NAME
        
        Call .Filter.Append("enterprise_id = @pEntID")
        Call .Parameters.add(.CreateParameter("@pEntID", ModGeral.EnterpriseID, dbFieldTypeLong))
        
        Call .Filter.Append("cd_estabelecimento = @pCdEst")
        Call .Parameters.add(.CreateParameter("@pCdEst", ModGeral.CdEstabelecimento, dbFieldTypeLong))
        
        Call .Filter.Append("tipo = @pTipoLanc")
        Call .Parameters.add(.CreateParameter("@pTipoLanc", strTipoLancamento, dbFieldTypeString))
        
        Call .Filter.Append("cd_codigo = @pCod")
        Call .Parameters.add(.CreateParameter("@pCod", lngCdTipoLanc, dbFieldTypeLong))
        
        Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, selCmd)
        
        If Not rdResult.EOF Then
            ExisteReg = True
        End If
    End With
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    mApp.Disconnect
    Exit Function
erro:
    ExisteReg = False
    FinallyConnection mApp
End Function

