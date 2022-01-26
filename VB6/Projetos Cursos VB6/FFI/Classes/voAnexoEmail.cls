VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "voAnexoEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrAnexoEmail  As String
Private mstrDocumentoEmail  As String

Public Property Get AnexoEmail() As String
    AnexoEmail = mstrAnexoEmail
End Property

Public Property Let AnexoEmail(ByVal valor As String)
    mstrAnexoEmail = valor
End Property

Public Property Get DocumentoEmail() As String
    DocumentoEmail = mstrDocumentoEmail
End Property

Public Property Let DocumentoEmail(ByVal valor As String)
    mstrDocumentoEmail = valor
End Property

