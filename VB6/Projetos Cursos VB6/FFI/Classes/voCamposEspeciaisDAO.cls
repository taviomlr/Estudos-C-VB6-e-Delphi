VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CamposEspeciaisDAO"
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
        
        If rdResult.Recordcount > 0 Then
            
            Set colecao = New colCamposEspeciais
            rdResult.MoveFirst
            While Not rdResult.EOF
                Set item = New voCamposEspeciais
                item.Descricao = rdResult.GetString("descricao")
                item.Valor = rdResult.GetString("valor")
                    
                Call colecao.add(item)
                Call rdResult.MoveNext
            Wend
            Set Carregar = colecao
        Else
           Set Carregar = Nothing
        End If
    End With
    
    Set rdResult = Nothing
    Set selCmd = Nothing
    
    Call mApp.Disconnect
    Exit Function
erro:
    Carregar = Nothing
End Function

Private Function Inserir(ByVal obj As voCamposEspeciais) As Boolean
    Dim cmd        As IDBInsertCommand
    
On Error GoTo erro_gravando
    Call mApp.Connect
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = TABLE_NAME
     
        Call .AddValue("[descricao]", "@pDescricao")
        Call .Parameters.add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString))
        
        Call .AddValue("[valor]", "@pValor")
        Call .Parameters.add(.CreateParameter("@pValor", obj.Valor, dbFieldTypeString))
                
    End With
    
    Inserir = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    
    Call mApp.Disconnect
    Set cmd = Nothing
    Exit Function
    
erro_gravando:
    Inserir = False
    Set cmd = Nothing
End Function

Public Function ExisteRegistro(ByVal obj As voCamposEspeciais) As Boolean
    Dim selCmd As IDBSelectCommand
    Dim rdResult As IDBReader

On Error GoTo erro
        Set selCmd = Aplicacao.CreateSelectCommand
    
        With selCmd
            .Table.TableName = TABLE_NAME
            Call .Filter.Append("[descricao] = @pDescricao")
            Call .Parameters.add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString))
            
            Call .Filter.Append("[valor] = @pValor")
            Call .Parameters.add(.CreateParameter("@pValor", obj.Valor, dbFieldTypeString))
                
            Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
            ExisteRegistro = rdResult.Recordcount > 0
        End With
    
        Set rdResult = Nothing
        Set selCmd = Nothing
        
        Exit Function
        
erro:
        ExisteRegistro = False
       
        Set rdResult = Nothing
        Set selCmd = Nothing
End Function
Public Function ExisteRegistroDuplicado(strDescricao As String, strNome As String) As Boolean
    Dim selCmd As IDBSelectCommand
    Dim rdResult As IDBReader

On Error GoTo erro
        Set selCmd = Aplicacao.CreateSelectCommand
    
        With selCmd
            .Table.TableName = TABLE_NAME
            
        Call .Filter.Append("[descricao] = @pDescricao")
        Call .Parameters.add(.CreateParameter("@pDescricao", strDescricao, dbFieldTypeString))
        
        Call .Filter.Append("[valor] = @pValor")
        Call .Parameters.add(.CreateParameter("@pValor", strNome, dbFieldTypeString))
                
            Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
            ExisteRegistroDuplicado = rdResult.Recordcount > 0
        End With
    
        Set rdResult = Nothing
        Set selCmd = Nothing
        
        Exit Function
        
erro:
        ExisteRegistroDuplicado = False
       
        Set rdResult = Nothing
        Set selCmd = Nothing
End Function


Public Function Gravar(ByVal obj As voCamposEspeciais) As Boolean
    
On Error GoTo erro
    
    Gravar = Inserir(obj)
    Exit Function
    
erro:
    Gravar = False
End Function

Public Function Excluir(ByVal obj As voCamposEspeciais) As Boolean
    Dim cmd        As IDBDeleteCommand

On Error GoTo erro_excluindo
    
    Call Aplicacao.Connect
    Set cmd = Aplicacao.CreateDeleteCommand
    With cmd
        .Table = TABLE_NAME
    
        Call .Filter.Append("[descricao] = @pDescricao")
        Call .Parameters.add(.CreateParameter("@pDescricao", obj.Descricao, dbFieldTypeString))
        
        Call .Filter.Append("[valor] = @pValor")
        Call .Parameters.add(.CreateParameter("@pValor", obj.Valor, dbFieldTypeString))
    End With
        
    Excluir = Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd) > 0
    Call Aplicacao.Disconnect
    Set cmd = Nothing
    
    Exit Function
    
erro_excluindo:
    Set cmd = Nothing
    Excluir = False
End Function
