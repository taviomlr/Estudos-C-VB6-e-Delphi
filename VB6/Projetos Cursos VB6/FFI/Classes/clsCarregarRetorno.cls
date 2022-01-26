VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarregarRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngEnterprise_id                       As Long
Private mlngCd_estabelecimento                  As Long
Private mlngBanco                               As Long
Private mlngId_carteira                         As Long
Private mstrMensagemValidacao                   As String
Private mstrCaminhoRetorno                      As String
Private mApp                                    As IApplication
Private mcolTitulo                              As clsColTituloCobrebem

Public Sub init(ByVal objApp As IApplication)
    Set mApp = objApp
End Sub

Public Property Let Enterprise_id(valor As Long)
   mlngEnterprise_id = valor
End Property

Public Property Get Enterprise_id() As Long
   Enterprise_id = mlngEnterprise_id
End Property

Public Property Let Cd_estabelecimento(valor As Long)
   mlngCd_estabelecimento = valor
End Property

Public Property Get Cd_estabelecimento() As Long
   Cd_estabelecimento = mlngCd_estabelecimento
End Property

Public Property Let Banco(valor As Long)
   mlngBanco = valor
End Property

Public Property Get Banco() As Long
   Banco = mlngBanco
End Property

Public Property Let Id_carteira(valor As Long)
   mlngId_carteira = valor
End Property

Public Property Get Id_carteira() As Long
   Id_carteira = mlngId_carteira
End Property

Public Property Let CaminhoRetorno(valor As String)
   mstrCaminhoRetorno = valor
End Property

Public Property Get CaminhoRetorno() As String
   CaminhoRetorno = mstrCaminhoRetorno
End Property

Public Property Get MensagemValidacao() As String
   MensagemValidacao = mstrMensagemValidacao
End Property

Public Property Get colTitulo() As clsColTituloCobrebem
    Set colTitulo = mcolTitulo
End Property

Public Property Let colTitulo(ByVal NewVal As clsColTituloCobrebem)
    Set mcolTitulo = NewVal
End Property

Public Function ExisteCarteira(ByVal lngBanco As Long, ByVal lngId_carteira As Long) As Boolean
    Dim cmd             As IDBSelectCommand
    Dim rdResult        As IDBReader
    Dim result          As clsCarteira

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = "[FFIBanco_carteira]"
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngId_carteira, dbFieldTypeLong))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        ExisteCarteira = True
    Else
        ExisteCarteira = False
    End If
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function

erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    ExisteCarteira = False
End Function

Private Sub Class_Initialize()
    Set mcolTitulo = New clsColTituloCobrebem
End Sub

Private Sub Class_Terminate()
    Set mcolTitulo = Nothing
End Sub
