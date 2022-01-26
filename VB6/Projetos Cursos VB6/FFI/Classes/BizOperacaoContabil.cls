VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BizOperacaoContabil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub validarCampoVinculadoOperContCusto(ByVal OperContCusto As Long, ByRef col As Collection)
    
    Dim objMensagemErro As New clsMensagemErro
    
    If OperContCusto <> 0 Then ' Campo preenchido.
        If Not OperacaoContabilExiste(OperContCusto) Then
            objMensagemErro.mensagem = "O valor informado no campo 'Operação Contábil' não existe."
            objMensagemErro.status = novo
            objMensagemErro.menu = Cadastro_OperacaoContabil
            col.add objMensagemErro
        End If
    End If
    
    Set objMensagemErro = Nothing
    
End Sub

Public Function OperacaoContabilExiste(ByVal OperContCusto As Long) As Boolean
    
    Dim objDAO As New daoVerificacaoBase
    
    On Error GoTo err
    
    If objDAO.verificaRegistroExiste("[OperacaoContabil]", "cd_operacao", OperContCusto) Then
        OperacaoContabilExiste = True
    Else
        OperacaoContabilExiste = False
    End If
    
    Set objDAO = Nothing
    Exit Function
    
err:
    
    Set objDAO = Nothing
    
End Function

