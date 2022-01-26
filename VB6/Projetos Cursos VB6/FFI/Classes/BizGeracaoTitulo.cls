VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BizGeracaoTitulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Retorna True se sigla for encontrada e false em caso contrário.
Public Function GeracaoTituloExiste(ByVal Codigo As String, ByVal Parcela As Long, ByVal Tipo As String, ByVal Empresa As String, ByVal PagRec As enuPagRec) As Boolean
    Dim objDAO              As New daoVerificacaoBase
    Dim strPagRec           As String
    Dim strTabela           As String
    
    On Error GoTo err
    
    strPagRec = IIf(PagRec = Pagamento, "P", "R")
    strTabela = IIf(PagRec = Pagamento, "[FVFTituloPagarDuplicata]", "[FVFTituloReceberDuplicata]")
    
    If objDAO.verificaRegistroExiste(strTabela, "nota,tipo_registro,empresa,pagRec,parcela", Codigo, Tipo, Empresa, strPagRec, Parcela) Then
        GeracaoTituloExiste = True
    Else
        GeracaoTituloExiste = False
    End If
    
    Set objDAO = Nothing
    Exit Function
err:
    Set objDAO = Nothing
    GeracaoTituloExiste = False
End Function
