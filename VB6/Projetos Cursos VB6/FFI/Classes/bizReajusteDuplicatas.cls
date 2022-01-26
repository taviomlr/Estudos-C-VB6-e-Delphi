VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bizReajusteDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjApp As IApplication
Private mstrErroReajuste As String

Public Sub init(objApp As IApplication)
    Set mobjApp = objApp
End Sub

Public Property Get ErroReajuste() As String
    ErroReajuste = mstrErroReajuste
End Property

Public Function ReajustaDuplicata(objReajuste As voReajusteDuplicatas) As Boolean
    Dim objDuplicata As New CDuplicata
    Dim objEmpresa As CEmpresas
    Dim objDaoReajuste As New daoReajusteDuplicatas
        
On Error GoTo err_Handler
        
    ReajustaDuplicata = True
    With objReajuste
        Call objDuplicata.Carregar(.PagRec, .Nota, .Empresa, .Tipo, .Parcela)
        If Not objDuplicata Is Nothing Then
            Call objDaoReajuste.init(mobjApp)
            Set objEmpresa = mobjCache.GetCacheEmpresa(.Empresa)
            objReajuste.valor_anterior = objDuplicata.ValorOriginal
            If DateTime.Date >= objEmpresa.DataBase1 And DateTime.Date < objEmpresa.DataBase2 Then
                If objReajuste.perc_2 > 0 Then
                    objReajuste.valor_reajustado = objDuplicata.ValorOriginal + (objDuplicata.ValorOriginal * (objReajuste.perc_1 / 100))
                End If
                objReajuste.perc_2 = 0
            ElseIf DateTime.Date >= objEmpresa.DataBase2 Then
                If objReajuste.perc_2 > 0 Then
                    objReajuste.valor_reajustado = objDuplicata.ValorOriginal + (objDuplicata.ValorOriginal * (objReajuste.perc_2 / 100))
                End If
                objReajuste.perc_1 = 0
            End If
            If objReajuste.perc_adicionais > 0 Then
                objReajuste.valor_reajustado = objDuplicata.ValorOriginal + (objReajuste.valor_reajustado * (objReajuste.perc_adicionais / 100))
            End If
            objDuplicata.ValorOriginal = objReajuste.valor_reajustado
            ReajustaDuplicata = objDaoReajuste.salvar(objReajuste)
            Call objDuplicata.AlterarDuplicata(False, objDuplicata.Nota, objDuplicata.Empresa, objDuplicata.Tipo, objDuplicata.Parcela)
        End If
    End With
    
    Exit Function
err_Handler:
    mstrErroReajuste = err.Description
    ReajustaDuplicata = False
End Function

Private Sub Class_Terminate()
    Set mobjApp = Nothing
End Sub
