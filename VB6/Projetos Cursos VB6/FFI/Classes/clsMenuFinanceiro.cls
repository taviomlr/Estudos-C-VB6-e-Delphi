VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IclsMenu

Public Function IclsMenu_retornaFormulario(ByVal Menu As enuMenu) As Form
    Dim Base As IclsMenu
    
    Select Case Menu
        Case enuMenu.Cadastro_Estado
            Set IclsMenu_retornaFormulario = frmEstados
        Case enuMenu.Cadastro_Municipio
            Set IclsMenu_retornaFormulario = frmCadMunicipio
        Case Else
            Set Base = New ClsMenu
            Set IclsMenu_retornaFormulario = Base.retornaFormulario(Menu)
    End Select
End Function
