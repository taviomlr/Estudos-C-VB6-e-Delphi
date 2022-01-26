VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarteiraCpEspecialDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mApp                            As IApplication
Private Const mstrTabela = "[FFICarteiraCamposEspeciais]"

Public Sub init(ByVal objApp As IApplication)
    Set mApp = objApp
End Sub

Public Function Gravar_Atualizar(ByVal obj As clsCarteira) As Boolean
    Dim blnGravar           As Boolean
    Dim objCpEspeciais      As clsCarteiraCpEspecial
    blnGravar = True
    blnGravar = blnGravar And Excluir(obj)
    If Not obj.ColCpEspeciais Is Nothing Then
        If obj.ColCpEspeciais.Count > 0 Then
            obj.ColCpEspeciais.MoveFirst
            While Not obj.ColCpEspeciais.EOF
                Set objCpEspeciais = obj.ColCpEspeciais.CurrentObject
                blnGravar = blnGravar And GravarUm(obj, objCpEspeciais)
                Set objCpEspeciais = Nothing
                obj.ColCpEspeciais.MoveNext
            Wend
        End If
    End If
    Gravar_Atualizar = blnGravar
End Function

Private Function GravarUm(ByVal objCarteira As clsCarteira, ByVal obj As clsCarteiraCpEspecial) As Boolean
    Dim cmd                 As IDBInsertCommand
    Dim blnGravar           As Boolean

On Error GoTo erro_gravando
    Set cmd = mApp.CreateInsertCommand
    With cmd
        .Table = mstrTabela
        Call .AddValue("[Enterprise_id]", "@pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", objCarteira.Enterprise_id, dbFieldTypeInt))
                    
        Call .AddValue("[Cd_estabelecimento]", "@pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", objCarteira.Cd_estabelecimento, dbFieldTypeInt))
        
        Call .AddValue("[id_carteira]", "@pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", objCarteira.Id_carteira, dbFieldTypeLong))
                    
        Call .AddValue("[nome]", "@pNome")
        Call .Parameters.add(.CreateParameter("@pNome", obj.CpEspNome, dbFieldTypeString, 100))
                    
        Call PrepareCommand(cmd, obj)
    End With
    
    blnGravar = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
    GravarUm = blnGravar
    Set cmd = Nothing
    Exit Function

erro_gravando:
    Set cmd = Nothing
    GravarUm = False
End Function

'Public Function AtualizarUm(ByVal obj As clsCarteira) As Boolean
'    Dim cmd                 As IDBUpdateCommand
'    Dim blnAtualizado       As Boolean
'
'On Error GoTo erro_atualizado
'    Set cmd = mApp.CreateUpdateCommand
'     With cmd
'        .Table = mstrTabela
'
'        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
'        Call .Parameters.add(.CreateParameter("@pEnterprise_id", obj.Enterprise_id, dbFieldTypeInt))
'
'        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
'        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", obj.Cd_estabelecimento, dbFieldTypeInt))
'
'        Call .Filter.Append("[id_carteira] = @pId_carteira")
'        Call .Parameters.add(.CreateParameter("@pId_carteira", obj.Id_carteira, dbFieldTypeLong))
'
'        Call .Filter.Append("[nome] = @pNome")
'        Call .Parameters.add(.CreateParameter("@pNome", obj.CpEspNome, dbFieldTypeString))
'
'        Call prepareCommand(cmd, obj)
'    End With
'
'    blnAtualizado = (mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = 1)
'
'    Set cmd = Nothing
'    Atualizar = blnAtualizado
'    Exit Function
'erro_atualizado:
'    Set cmd = Nothing
'End Function

Public Function Carregar(ByVal lngEnterprise_id As Integer, ByVal lngCd_estabelecimento As Integer, ByVal lngId_carteira As Long) As clscolCarteiraCpEspecial
    Dim cmd             As IDBSelectCommand
    Dim rdResult        As IDBReader
    Dim result          As clsCarteiraCpEspecial

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = mstrTabela
        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", lngEnterprise_id, dbFieldTypeLong))
                    
        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", lngCd_estabelecimento, dbFieldTypeLong))
        
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngId_carteira, dbFieldTypeLong))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Set Carregar = New clscolCarteiraCpEspecial
        rdResult.MoveFirst
        While Not rdResult.EOF
            Set result = New clsCarteiraCpEspecial
            Call Carregar.add(result.CriarObjeto(rdResult))
            Set result = Nothing
            rdResult.MoveNext
        Wend
    Else
        Set Carregar = Nothing
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function

erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    Set Carregar = Nothing
End Function

Public Function Excluir(ByVal obj As clsCarteira) As Boolean
    Dim cmd                             As IDBDeleteCommand
    Dim blnExcluido                     As Boolean
    
On Error GoTo erro_excluindo
    
    blnExcluido = True
    Set cmd = mApp.CreateDeleteCommand
    With cmd
        .Table = mstrTabela
        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", obj.Enterprise_id, dbFieldTypeInt))
                    
        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", obj.Cd_estabelecimento, dbFieldTypeInt))
        
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", obj.Id_carteira, dbFieldTypeLong))
        
    End With
    
    blnExcluido = CBool(mApp.ExecuteUpdate(mApp.GetInternalAuthorization, cmd) = CLng(1))
    
    Excluir = True
    
    Exit Function
erro_excluindo:
    Excluir = False
End Function

Public Function Existe(ByVal obj As clsCarteiraCpEspecial) As Boolean
    Dim selCmd                      As IDBSelectCommand
    Dim rdResult                    As IDBReader

    Set selCmd = mApp.CreateSelectCommand
    With selCmd
        .Table.TableName = mstrTabela
        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", obj.Enterprise_id, dbFieldTypeInt))
                    
        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", obj.Cd_estabelecimento, dbFieldTypeInt))
        
        Call .Filter.Append("[Id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", obj.Id_carteira, dbFieldTypeLong))
        
        Call .Filter.Append("[nome] = @pNome")
        Call .Parameters.add(.CreateParameter("@pNome", obj.CpEspNome, dbFieldTypeString))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, selCmd)
    If rdResult.Recordcount > 0 Then
        rdResult.MoveFirst
    End If
    Existe = Not rdResult.EOF
    Set selCmd = Nothing
    rdResult.CloseReader
    Set rdResult = Nothing
End Function

Private Sub PrepareCommand(ByRef cmd As Object, ByVal obj As clsCarteiraCpEspecial)
    With cmd
        Call .AddValue("[valor]", "@pValor_")
        Call .Parameters.add(.CreateParameter("@pValor_", obj.CpEspValor, dbFieldTypeString, 100))
    End With
End Sub
Public Function ExisteVinculacaoCampo(ByVal lngEnterprise_id As Integer, ByVal lngCd_estabelecimento As Integer, ByVal strCampo As String) As Boolean
    Dim cmd             As IDBSelectCommand
    Dim rdResult        As IDBReader
    Dim result          As clsCarteiraCpEspecial

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = mstrTabela
        Call .Filter.Append("[Enterprise_id] = @pEnterprise_id")
        Call .Parameters.add(.CreateParameter("@pEnterprise_id", lngEnterprise_id, dbFieldTypeLong))
                    
        Call .Filter.Append("[Cd_estabelecimento] = @pCd_estabelecimento")
        Call .Parameters.add(.CreateParameter("@pCd_estabelecimento", lngCd_estabelecimento, dbFieldTypeLong))
        
        Call .Filter.Append("[nome] = @pNome")
        Call .Parameters.add(.CreateParameter("@pNome", Trim(strCampo), dbFieldTypeString))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        ExisteVinculacaoCampo = True
    Else
        ExisteVinculacaoCampo = False
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function

erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    ExisteVinculacaoCampo = False
End Function
