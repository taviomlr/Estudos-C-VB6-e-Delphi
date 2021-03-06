VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub AtualizarValorMoraDiaria()
    Dim sSql As String
    If gTipoDB = Access Then
        sSql = ""
        sSql = sSql & "UPDATE Lan�amentos L "
        sSql = sSql & "INNER JOIN "
        sSql = sSql & "Bancos B "
        sSql = sSql & "ON "
        sSql = sSql & "L.Banco = B.Banco "
        sSql = sSql & "SET "
        sSql = sSql & "L.VLRMRD = CCUR( ( (L.[VALOR ORIGINAL] * B.MORA)/100 ) / 30 ) "
        ExecuteSQL sSql
    Else
        sSql = "UPDATE Lan�amentos SET Lan�amentos.VLRMRD =  ( (Lan�amentos.[VALOR ORIGINAL] * Bancos.MORA)/100 ) / 30  FROM Lan�amentos INNER JOIN Bancos ON Lan�amentos.Banco = Bancos.Banco"
        ExecuteSQL sSql
    End If
End Sub

Public Function SelecionaBordero(Codigo As Long, PagRec As String) As String
    If gTipoDB = MsSql Then
        SelecionaBordero = wsprintf("SELECT [C�digo], Empresa, [Emiss�o], Vencimento, " & _
                        "CONVERT(varchar,CAST(([Valor Original] + [Acr�scimo] - Abatimento) AS MONEY),1), 'L' as Origem " & _
                        "FROM [Lan�amentos] WHERE [Border�] = %l AND PagRec = '%s'", Codigo, PagRec)
    Else
        SelecionaBordero = wsprintf("SELECT [C�digo], Empresa, [Emiss�o], Vencimento, " & _
                            "FORMAT([Valor Original] + [Acr�scimo] - Abatimento, \'#,##0.00\'), 'L' as Origem " & _
                            "FROM [Lan�amentos] WHERE [Border�] = %l AND PagRec = '%s'", Codigo, PagRec)
    End If
End Function
