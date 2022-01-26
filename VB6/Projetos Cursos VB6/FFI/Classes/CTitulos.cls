VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************
'
'    Nome da Classe: CTitulos
'    Escrito por: Edilberto Sandro da Silva

'    Utilizada para consultar lançamentos e duplicatas
'    simultaneamente
'
'*************************************************************

Private oConexao      As ADODB.Connection
Private rs            As ADODB.Recordset
Private dTotFinal As Double

Public Enum OrigemTitulo
    ori_Ambos = 0
    ori_Duplicatas = 1
    ori_Lancamentos = 2
End Enum

Public Enum TipoTitulo
    tip_Todos = 0
    tip_A_Receber = 1
    tip_A_Pagar = 2
    tip_Pago = 3
    tip_Recebido = 4
End Enum

Public Enum TipoTituloAtrasado
    tip_Todos_Vencidos = 0
    tip_A_Receber_Vencido = 1
    tip_A_Pagar_Vencido = 2
End Enum

Public Property Get CON() As ADODB.Connection
  Set CON = oConexao
End Property

Public Property Get TotFinal() As Double
    TotFinal = dTotFinal
End Property

Public Function ConsultaTitulosAtrasados(pOrigem As OrigemTitulo, pTipo As TipoTituloAtrasado, Optional pEmpresa As String, Optional pDtBaseCalculo As Date, Optional pJuroMora As Single, Optional pBanco As Long) As Object
'Esta função carrega o registro no caso do título existir.
On Error GoTo ErroConsultaTitulosAtrasados
   Dim sSQLFinal As String
   Dim sSQLDup As String
   Dim sSQLLan As String
   Dim sWhere As String
   
   Set rs = New ADODB.Recordset
   If gTipoDB = mssql Then
    sWhere = "WHERE Vencimento <= GetDate() And Pagamento is NULL "
   Else
    sWhere = "WHERE Vencimento <= Date() And Pagamento is NULL "
   End If
   
   Select Case pTipo
       Case tip_A_Receber_Vencido
           sWhere = sWhere & "AND PagRec = 'R' "
       Case tip_A_Pagar_Vencido
           sWhere = sWhere & "AND PagRec = 'P' "
   End Select

   If pEmpresa <> Empty Then
      sWhere = sWhere & "AND Empresa = '" & pEmpresa & "' "
   End If
   
   'PT 77198
   ' FILTRO PARA VER SE TEM BANCO NA CHAMADA DO METODO,
   ' PARA ISSO FOI ADICIONADO UM PARAMENTRO OPCIONAL pBanco

   If pBanco <> Empty Then
      sWhere = sWhere & "AND BANCO = " & pBanco
   End If

   If gTipoDB = mssql Then
        Dim sDias As String
        Dim sVlrMoraCalc As String
        sDias = "DATEDIFF(d,Vencimento," & InverteData(DataToStr(pDtBaseCalculo), True) & ")"
        sVlrMoraCalc = "CASE WHEN VlrMRD > 0 THEN VlrMRD * " & sDias & " ELSE [valor original] * ( " & str(pJuroMora / 100) & " * " & sDias & " ) END"
        sSQLDup = "SELECT 'D' As Origem, PagRec, Nota as Titulo, Empresa, Tipo, Parcela, Descrição, [Valor Original], VlrMul, VlrMRD , Vencimento, " & sDias & " as Dias, " & sVlrMoraCalc & " as VlrMoraCalc, Banco, ([Valor Original] + VlrMul + " & sVlrMoraCalc & ") as VlrFinal FROM [Duplicatas] " & sWhere
        sSQLLan = "SELECT 'L' As Origem, PagRec, Código as Titulo, Empresa, Tipo, '0' as Parcela, Descrição, [Valor Original], VlrMul, VlrMRD,  Vencimento, " & sDias & " as Dias, " & sVlrMoraCalc & " as VlrMoraCalc , Banco, ([Valor Original] + VlrMul + " & sVlrMoraCalc & ") as VlrFinal FROM [Lançamentos] " & sWhere
   Else
        sSQLDup = "SELECT 'D' As Origem, PagRec, Nota as Titulo, Empresa, Tipo, Parcela, Descrição, [Valor Original], VlrMul, VlrMRD , Vencimento, (" & InverteData(DataToStr(pDtBaseCalculo), True) & " - Vencimento) as Dias, ((IIf(VlrMRD > 0, (VlrMRD * dias), ([valor original] * ( " & str(pJuroMora / 100) & " * Dias ))))) as VlrMoraCalc, Banco, ([Valor Original] + VlrMul + VlrMoraCalc) as VlrFinal FROM [Duplicatas] " & sWhere
        sSQLLan = "SELECT 'L' As Origem, PagRec, Código as Titulo, Empresa, Tipo, '0' as Parcela, Descrição, [Valor Original], VlrMul, VlrMRD,  Vencimento, (" & InverteData(DataToStr(pDtBaseCalculo), True) & " - Vencimento) as Dias, ((IIf(VlrMRD > 0, (VlrMRD * dias), ([valor original] * ( " & str(pJuroMora / 100) & " * Dias ))))) as VlrMoraCalc , Banco, ([Valor Original] + VlrMul + VlrMoraCalc) as VlrFinal FROM [Lançamentos] " & sWhere
   End If
     
   Select Case pOrigem
        Case ori_Duplicatas
             sSQLFinal = sSQLDup
        Case ori_Lancamentos
             sSQLFinal = sSQLLan
        Case Else
             sSQLFinal = sSQLDup & " UNION ALL " & sSQLLan
   End Select
   
   rs.Open sSQLFinal, oConexao, adOpenKeyset, adLockPessimistic
   
   If Not rs.BOF Then
     Call CalcTotFinal(rs)
     Set ConsultaTitulosAtrasados = rs
     
   Else
     Set ConsultaTitulosAtrasados = Nothing
   End If
  ' rs.Close
  ' Set rs = Nothing
   Exit Function
ErroConsultaTitulosAtrasados:
   err.Raise 1 + vbObjectError, , "Não foi possível localizar títulos!"
   rs.Close
   Set rs = Nothing
   Resume
End Function

Private Sub CalcTotFinal(ByVal rsRegistros As ADODB.Recordset)
  If Not rsRegistros.BOF Then
     rsRegistros.MoveFirst
     Do
       dTotFinal = dTotFinal + rsRegistros("VlrFinal").Value
       rsRegistros.MoveNext
     Loop Until rsRegistros.EOF
     rsRegistros.MoveFirst
  Else
     dTotFinal = 0
  End If
End Sub

Private Sub Class_Initialize()
On Error GoTo ErroInicializar
   Set oConexao = conexao
   Exit Sub
ErroInicializar:
   err.Raise 1 + vbObjectError, , " Erro ao inicializar classe lançamentos!"
End Sub

Private Sub Class_Terminate()
   Set oConexao = Nothing
End Sub

