VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRatFinanceiro 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rateio Financeiro"
   ClientHeight    =   3495
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnCancelCentro 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4980
      Picture         =   "frmRatFinanceiro.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   345
   End
   Begin VB.CommandButton btnOkCentro 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4980
      Picture         =   "frmRatFinanceiro.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Confirma Rateio por Centro de custos <F6>"
      Top             =   930
      Width           =   345
   End
   Begin VB.CommandButton btnCentro 
      Height          =   375
      Left            =   4980
      Picture         =   "frmRatFinanceiro.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Rateio por Centro de Custo - <F5>"
      Top             =   540
      Width           =   345
   End
   Begin MSDataGridLib.DataGrid dgCentro 
      Height          =   2055
      Left            =   5340
      TabIndex        =   5
      Top             =   510
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Centro de Custo"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgConta 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   510
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Projeto/Conta Financeira"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtVlrOriginal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   90
      Width           =   2025
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   2970
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   7020
      TabIndex        =   0
      Top             =   2970
      Width           =   1215
   End
   Begin VB.Label lblVlrRatCentro 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7980
      TabIndex        =   10
      Top             =   2610
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Rateado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5880
      TabIndex        =   9
      Top             =   2610
      Width           =   2025
   End
   Begin VB.Label lblVlrRatConta 
      Alignment       =   1  'Right Justify
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3390
      TabIndex        =   8
      Top             =   2610
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Valor Rateado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   7
      Top             =   2610
      Width           =   2205
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Vlr. Original:"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmRatFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const TamCol1 = 10
Const TamColCodigo = 1000
Const TamColValor = 1100
Const TamColPercent = 1000


Public vTipTit As String 'Titulo do título L-Lançamento D-Duplicata
Public vPagRec As String 'Tipo do titulo R-Receber, P-Pagar
Public vNumdoc As String 'Numero do lancamento ou Numero da nota

'somente para duplicatas
Public vCodEmp As String 'codigo da empresa
Public vTipReg As String 'tipo de registro - tipo global
Public vNumPar As String 'número da parcela

Private rstConta            As Object
Private fdsConta(3)         As FieldStruct
Private rstCentro           As Object
Private fdsCentro(4)        As FieldStruct

Private TabConta As String 'Nome da tabela de contas
Private TabCentro As String 'Nome da tabela de centros de custo

Private ConPrj As Boolean

Private CodConta As String 'código da conta sendo rateada por centros de custo
Private CodProjeto As String 'código do projeto sendo rateado por centros de custo


'Armazenamento de valores
'
Public ValRateio As Double 'Recebido da rotina externa.
Private ValRateioCentro As Double 'Recebido da rotina interna de contas

Private VlrRatConta As Double 'Valor rateado da conta
Private VlrRatCentro As Double 'Valor rateado do centro de custo

Private PerRatConta As Single
Private PerRatCentro As Single

Dim rsADOCentro As ADODB.Recordset
Dim rsADOConta As ADODB.Recordset

Private Sub btnCancelCentro_Click()
  
  If MsgBox("Confirma Cancelamento/Exclusão do Rateio por Centro de Custos", vbYesNo) = vbYes Then
    Call EliminaRatCentro(CodConta, CodProjeto)
    
    'envia comando inválido para "limpar" a grid.
    AtualizaTotalCentro (-100)
    Call FiltrosCentros(-100, -100, 0)
      
    dgCentro.Caption = "Centro de Custo"
    btnOkCentro.Enabled = False
    btnCancelCentro.Enabled = False
    btnCentro.Enabled = True
    dgCentro.Enabled = False
    dgConta.Enabled = True
    dgConta.SetFocus
  End If
  
End Sub

Private Sub btnCentro_Click()
  'atribui o controle para rateio por centro de custo.
  CodConta = rsADOConta.Fields("Conta").Value
  CodProjeto = rsADOConta.Fields("Projeto").Value
  ValRateioCentro = rsADOConta.Fields("Valor").Value
  
  If CodConta <> "" Then
    btnCentro.Enabled = False
    btnOkCentro.Enabled = True
    btnCancelCentro.Enabled = True
    dgConta.Enabled = False
    dgCentro.Enabled = True
    dgCentro.SetFocus
    FiltrosCentros CodConta, CodProjeto, ValRateioCentro
  End If
End Sub


Private Sub btnOkCentro_Click()
  
  If Round(ValRateioCentro, 2) <> Round(VlrRatCentro, 2) Then
    MsgBox "Valor total do rateio incorreto"
  Else
    'envia comando inválido para "limpar" a grid.
    AtualizaTotalCentro (-100)
    Call FiltrosCentros(-100, -100, 0)
      
    dgCentro.Caption = "Centro de Custo"
    btnOkCentro.Enabled = False
    btnCancelCentro.Enabled = False
    btnCentro.Enabled = True
    dgCentro.Enabled = False
    dgConta.Enabled = True
    dgConta.SetFocus
  End If
End Sub

Private Sub dgCentro_AfterColUpdate(ByVal ColIndex As Integer)
  'o usuário informou o percentual.. calcula-se o valor
  If ColIndex = 3 Then
    rsADOCentro.Fields("Valor").Value = ValRateioCentro * rsADOCentro.Fields("Percentual").Value / 100
  End If
  
  'o usuário informou o valor... calcula-se o percentual
  If ColIndex = 4 Then
    rsADOCentro.Fields("Percentual").Value = rsADOCentro.Fields("Valor").Value * 100 / ValRateioCentro
  End If
End Sub

Private Sub dgCentro_AfterUpdate()
  AtualizaTotalCentro (rsADOConta.Fields("Conta").Value)
End Sub

Private Sub dgCentro_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim vRet As String
  
  'If KeyCode = 13 And dgCentro.col <> 3 Then
  '  dgCentro.DataChanged = True
  '  dgCentro.col = dgCentro.col + 1
  'End If
  
  'Pesquisa de Centros de Custos
  If Shift = 0 And KeyCode = vbKeyPageDown Then
    If dgCentro.col = 1 Then
      PCampo "Centros", "Centros", pbCampo, vRet, "Código"
      rsADOCentro.Fields("Centro").Value = vRet
      dgCentro.Columns(1).Value = vRet
    End If
  End If
  
  'chama da função do botão de confirmação do rateio por ccu.
  If KeyCode = vbKeyF6 Then
    btnOkCentro_Click
  End If
End Sub

Private Sub dgCentro_OnAddNew()
  Dim dVal As Single
  
  rsADOCentro.Fields("Conta").Value = CodConta
  rsADOCentro.Fields("Projeto").Value = CodProjeto
    
  AtualizaTotalCentro (CodConta)
  
  If ValRateioCentro = VlrRatCentro Then
    MsgBox "Valor total já rateado."
    rsADOConta.Cancel
    Exit Sub
  End If
    
  If Not (rsADOCentro.EOF) And Not (rsADOCentro.BOF) Then
    dVal = 100 - PerRatCentro
    If dVal < 0 Then dVal = 0
    rsADOCentro.Fields("Percentual") = dVal
    
    dVal = ValRateioCentro - VlrRatCentro
    If dVal < 0 Then dVal = 0
    rsADOCentro.Fields("Valor") = dVal
  End If
End Sub

Private Sub dgConta_AfterColUpdate(ByVal ColIndex As Integer)

  
  'o usuário informou o percentual.. calcula-se o valor
  If ColIndex = 2 Then
    rsADOConta.Fields("Valor").Value = ValRateio * rsADOConta.Fields("Percentual").Value / 100
  End If
  
  'o usuário informou o valor... calcula-se o percentual
  If ColIndex = 3 Then
    rsADOConta.Fields("Percentual").Value = rsADOConta.Fields("Valor").Value * 100 / ValRateio
  End If
  
End Sub

Private Sub dgConta_AfterUpdate()
  AtualizaTotalConta
  
  'verifica existência da conta
  Dim rs As New ADODB.Recordset
  Dim strSql As String
  
  strSql = "SELECT 1 FROM CONTAS WHERE [CÓDIGO] = "
  
  'Verifica se a conta informada existe
  If rsADOConta.State = adAddNew Or rsADOConta.State = adUpdate Then
    strSql = strSql + CStr(rsADOConta.Fields("Conta").Value)
    rs.Open strSql, conexao, adOpenKeyset, adLockPessimistic
    If rs.BOF Then
       MsgBox "Código da conta inexistênte"
       rsADOConta.CancelUpdate
    End If
  End If
  
  Set rs = Nothing
End Sub

Private Sub dgConta_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim vRet As String
  
  If KeyCode = 13 And dgConta.col <> 3 Then
    dgConta.DataChanged = True
    dgConta.col = dgConta.col + 1
  End If
  
  'Pesquisa de Contas Financeiras
  If Shift = 0 And KeyCode = vbKeyPageDown Then
    If dgConta.col = 1 Then
      PCampo "Contas", "Contas", pbCampo, vRet, "Código"
      'dgConta.Columns("Conta").Value = vRet
      rsADOConta.Fields("Conta").Value = vRet
      dgConta.Columns(1).Value = vRet
    End If
  End If
  
  'Pesquisa de Projetos
  If Shift = 0 And KeyCode = vbKeyPageDown Then
    If dgConta.col = 0 Then
      PCampo "PROJETO", "PROJETO", pbCampo, vRet, "CODPRJ"
      dgConta.Columns("PROJETO").Value = vRet
      dgConta.Columns(0).Value = vRet
    End If
  End If
  
  If KeyCode = vbKeyF5 Then
    btnCentro_Click
  End If
End Sub

Private Sub dgConta_OnAddNew()
  Dim dVal As Single
  
  AtualizaTotalConta
  
  If Not ConPrj Then
    rsADOConta.Fields("Projeto").Value = 0
  End If
  
  If ValRateio = VlrRatConta Then
    MsgBox "Valor total já rateado."
    rsADOConta.Cancel
    Exit Sub
  End If
    
  If Not (rsADOConta.EOF) And Not (rsADOConta.BOF) Then
    dVal = 100 - PerRatConta
    If dVal < 0 Then dVal = 0
    rsADOConta.Fields("Percentual") = dVal
    
    dVal = ValRateio - VlrRatConta
    If dVal < 0 Then dVal = 0
    rsADOConta.Fields("Valor") = dVal
  End If
  
  dgConta.col = 1
End Sub

Private Sub Form_Load()

  Set rsADOCentro = New ADODB.Recordset
  Set rsADOConta = New ADODB.Recordset
  
  'Cria as tabelas temporárias para manutenção dos registros
  CriaConta
  CriaCentro
   
  'Apresenta a grid com os registros
  'MostraConta
  MostraCentro
  
  ConfiguraGridCentro
   
  txtVlrOriginal.Text = Format(ValRateio, "#,##0.0000")
  
  ConPrj = False
  
  CarregaRateio
  
End Sub


'Cria a tabela temporária para as contas
'
Sub CriaConta()
  Dim strTabela  As String
   
  ' Tabela de Contas
  ' Tabela de Centros de Custos
   
  'Define a estrutura da tabela temporáriA de contas.
  AppendVar fdsConta(0), "Projeto", dbLong, 6, True
  AppendVar fdsConta(1), "Conta", dbLong, 6, True
  AppendVar fdsConta(2), "Percentual", dbDouble, 5, True
  AppendVar fdsConta(3), "Valor", dbDouble, 18, True
  
  CrieAux rstConta, fdsConta
  TabConta = NomeTabeladoRST(rstConta)
End Sub

'Cria a tabela temporária para os centros de custos
'
Sub CriaCentro()
  Dim strTabela As String

  
  'Tabela de Centros de Custos
    
  'define a estrura da tabela temporária de centros de custos
  AppendVar fdsCentro(0), "Conta", dbLong, 6
  AppendVar fdsCentro(1), "Projeto", dbLong, 6
  AppendVar fdsCentro(2), "Centro", dbLong, 6
  AppendVar fdsCentro(3), "Percentual", dbDouble, 18
  AppendVar fdsCentro(4), "Valor", dbDouble, 18

  CrieAux rstCentro, fdsCentro
  TabCentro = NomeTabeladoRST(rstCentro)
End Sub

' filtra a apresentação do centros de custos
' por conta financeira
Private Sub FiltrosCentros(pConta As String, pProjeto As String, pValor As Double)
  
  dgCentro.Caption = "Centro de Custo  - " + Format(pValor, "#,##0.00")
  
  VlrRatCentro = 0
  PerRatCentro = 0
  
  rsADOCentro.Filter = "CONTA = " + pConta + " and PROJETO = " + pProjeto
  rsADOCentro.Requery
  dgCentro.Refresh

  ConfiguraGridCentro
  
  AtualizaTotalCentro (pConta)
   
  
End Sub


' Configura a grid de apresentação dos centros de custo
'
Private Sub ConfiguraGridCentro()
  With dgCentro
    .Columns(0).Width = TamColCodigo
    .Columns(1).Width = TamColCodigo
    .Columns(2).Width = TamColCodigo
    .Columns(3).Width = TamColPercent
    .Columns(4).Width = TamColValor
  
    .Columns(3).Alignment = dbgRight
    .Columns(4).Alignment = dbgRight
    
    .Columns("Percentual").NumberFormat = "#.##0,00"
    .Columns("Valor").NumberFormat = "#.##0,00"
  
    .Columns(0).Visible = False
    .Columns(1).Visible = False
  End With

End Sub


'apresenta a grid das centros de custo
Private Sub MostraCentro()
On Error GoTo errormostracentros
  
  rsADOCentro.CursorLocation = adUseClient
  rsADOCentro.Open TabCentro, conexao, adOpenKeyset, adLockOptimistic
  Set dgCentro.DataSource = rsADOCentro
  dgCentro.Refresh
  
Exit Sub
errormostracentros:
  Err.Raise Err.Number, Err.Source, Err.Description

End Sub


'Apresenta a grid das contas financeiras
Sub MostraConta()
  Dim sSQL As String
  'Prepara o Recordset
  
  sSQL = TabConta
  
  rsADOConta.CursorLocation = adUseClient
  rsADOConta.Open sSQL, conexao, adOpenKeyset, adLockOptimistic
  Set dgConta.DataSource = rsADOConta
  dgConta.Refresh
  
  
  'Configura a apresentação da grid
  '
  With dgConta
    .ClearFields
          
    .Columns(0).Width = TamColCodigo
    .Columns(1).Width = TamColCodigo
    .Columns(2).Width = TamColPercent
    .Columns(3).Width = TamColValor
    
    .Columns(2).Alignment = dbgRight
    .Columns(3).Alignment = dbgRight
  
    If Not ConPrj Then 'ControlaProjeto Then
      .Columns(0).Visible = False
    End If
    
    .Columns("Percentual").NumberFormat = "#.##0,00"
    .Columns("Valor").NumberFormat = "#.##0,00"
  End With
End Sub


'Atualiza os totais já rateado da conta.
'
Private Sub AtualizaTotalConta()
  Dim rs As New ADODB.Recordset
  
  Dim strSql As String
  
  strSql = "SELECT SUM(PERCENTUAL) AS TOTPER , SUM(VALOR) AS TOTVAL FROM " + TabConta

  rs.Open strSql, conexao, adOpenKeyset, adLockPessimistic
  If Not rs.BOF Then
    PerRatConta = IIf(IsNull(rs![TOTPER]), 0, rs![TOTPER])
    VlrRatConta = IIf(IsNull(rs![TOTVAL]), 0, rs![TOTVAL])
  End If
      
  lblVlrRatConta.Caption = Format(VlrRatConta, "#,#0.0000")
  
  Set rs = Nothing
End Sub

'Atualiza o total já ratedo por centro de custo
'
Private Sub AtualizaTotalCentro(pConta As String)
  Dim rs As New ADODB.Recordset
  
  Dim strSql As String
  
  strSql = "SELECT SUM(PERCENTUAL) AS TOTPER , SUM(VALOR) AS TOTVAL FROM " + TabCentro
  strSql = strSql + " WHERE CONTA = " + pConta

  rs.Open strSql, conexao, adOpenKeyset, adLockPessimistic
  If Not rs.BOF Then
    PerRatCentro = IIf(IsNull(rs![TOTPER]), 0, rs![TOTPER])
    VlrRatCentro = IIf(IsNull(rs![TOTVAL]), 0, rs![TOTVAL])
  End If
      
  lblVlrRatCentro.Caption = Format(VlrRatCentro, "#,#0.0000")
  
  Set rs = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
  
  'Apaga as tabelas temporárias
  Call DeleteAux(rstConta, TabConta)
  Call DeleteAux(rstCentro, TabCentro)
  
  Unload Me
End Sub


'Eliminar rateio por centros de custo de uma determinada conta.
'
Private Sub EliminaRatCentro(pConta As String, pProjeto As String)
  Dim strSql As String
  
  rsADOCentro.Cancel
  
  strSql = "DELETE FROM " + TabCentro + " WHERE 1=1 "
  strSql = strSql + " AND CONTA = " + pConta
  strSql = strSql + " AND PROJETO = " + pProjeto
  
  conexao.Execute (strSql)
  
  
  
End Sub

' Le os registro armazenado e grava na tabela destino
'
Private Sub CarregaRateio()
  Dim rs As New ADODB.Recordset
  Dim strSql As String
  
  strSql = "SELECT *  "
  strSql = strSql + " FROM RATFINANCEIRO "
  
  'trecho especifico por tipo de titulo
  If vTipTit = "L" Then
    strSql = strSql + " WHERE TIPTIT = 'L' "
    strSql = strSql + " AND PAGREC = " + Quote(vPagRec, "'")
    strSql = strSql + " AND NUMDOC = " + vNumdoc
  Else
    strSql = strSql + " WHERE TIPTIT = 'D' "
    strSql = strSql + " AND PAGREC = " + Quote(vPagRec, "'")
    strSql = strSql + " AND NUMDOC = " + vNumdoc
    strSql = strSql + " AND CODEMP = " + Quote(vCodEmp, "'")
    strSql = strSql + " AND NUMDOC = " + Quote(vTipReg, "'")
    strSql = strSql + " AND NUMPAR = " + vNumPar
  End If
  
  strSql = strSql + " order by seqrat"
   
  'busca os registro armazenados na base
  rs.Open strSql, conexao, adOpenKeyset, adLockOptimistic
  If Not rs.BOF Then
    Dim vCodCtf As String
    Dim vCodPrj As String
    Dim sSqlCtf As String
    Dim sSqlCcu As String
    
    vCodCtf = -1
    vCodPrj = -1
    
    
    While Not rs.EOF
      sSqlCtf = "INSERT INTO " + TabConta + "(CONTA, PROJETO, VALOR, PERCENTUAL) VALUES ("
      sSqlCcu = "INSERT INTO " + TabCentro + "(CONTA, PROJETO, CENTRO, VALOR, PERCENTUAL) VALUES ("
      
      'verifica necessidade de incluir a conta financeira
      If vCodCtf <> rs![codctf] Or vCodPrj <> rs![codprj] Then
        vCodCtf = rs![codctf]
        vCodPrj = rs![codprj]
        
        sSqlCtf = sSqlCtf + CStr(rs![codctf]) + "," + CStr(rs![codprj]) + "," + CStr(rs![VLRCTF]) + "," + CStr(rs![PERCTF]) + ")"
        conexao.Execute (sSqlCtf)
      End If
      
      'centro de custo sempre inclui.
      sSqlCcu = sSqlCcu + CStr(rs![codctf]) + "," + CStr(rs![codprj]) + "," + CStr(rs![CODCCU]) + "," + CStr(rs![VLRCCU]) + "," + CStr(rs![PERCCU]) + ")"
      conexao.Execute (sSqlCcu)
         
      rs.MoveNext
    Wend
  End If
  Set rs = Nothing
  
  
  MostraConta
  
  AtualizaTotalConta
End Sub

' Valida Rateio
'
Private Function ValidaRateio() As Boolean
  Dim rs As ADODB.Recordset
  Dim strSql As String
  
  ValidaRateio = True
  
  If ValRateio <> VlrRatConta Then
    MsgBox "Valor Rateado para as Contas Financeiras incorreto "
    ValidaRateio = False
    Exit Function
  End If
    
  'aqui validamos o valor do rateio por centro de custo
  Set rs = New ADODB.Recordset

  strSql = "SELECT CTF.CONTA as CONTA , CTF.PROJETO as PROJETO, SUM(CTF.VALOR) AS VLRCTF, SUM(CCU.VALOR) AS VLRCCU "
  strSql = strSql + " FROM " + TabConta + " AS CTF, " + TabCentro + " AS CCU "
  strSql = strSql + " WHERE CTF.CONTA = CCU.CONTA AND CTF.PROJETO = CCU.PROJETO "
  strSql = strSql + " GROUP BY CTF.CONTA, CTF.PROJETO "
  
  'percorre as contas informadas validando o valor do rateio por centro de custo.
  rs.Open strSql, conexao, adOpenKeyset, adLockOptimistic
  If Not rs.BOF Then
    Do While Not rs.EOF
      If rs![VLRCTF] <> rs![VLRCCU] Then
        MsgBox "Rateio por Centro de Custo para a conta " + CStr(rs![Conta]) + " incorreto."
        ValidaRateio = False
        Exit Do
      Else
        rs.MoveNext
      End If
    Loop
  End If
  
  Set rs = Nothing
End Function


' Gera os Registros na tabela de Rateio
'
Private Sub GravaRateio()
  Dim strSql As String 'comando para recupera das informações.
  Dim strInsert As String 'string com o comando para inclusão do rateio
  Dim strDelete As String
  Dim rs As ADODB.Recordset
   
            strSql = " SELECT "
  strSql = strSql + " CTF.CONTA AS CODCTF, "
  strSql = strSql + " CTF.PROJETO AS CODPRJ, "
  strSql = strSql + " CTF.VALOR AS VLRCTF, "
  strSql = strSql + " CTF.PERCENTUAL AS PERCTF, "
  strSql = strSql + " CCU.CENTRO AS CODCCU, "
  strSql = strSql + " CTF.VALOR AS VLRCCU, "
  strSql = strSql + " CTF.PERCENTUAL AS PERCCU "
  strSql = strSql + " FROM " + TabConta + " AS CTF, " + TabCentro + " AS CCU"
  strSql = strSql + " WHERE CTF.CONTA = CCU.CONTA AND CTF.PROJETO = CCU.PROJETO "
  
On Error GoTo Trata_Erro
  
  conexao.BeginTrans
  
  'elimina todas os registro anteriores
  '
  strDelete = "DELETE FROM RATFINANCEIRO WHERE 1=1 "
  'trecho especifico por tipo de titulo
  If vTipTit = "L" Then
    strDelete = strDelete + " AND TIPTIT = 'L' "
    strDelete = strDelete + " AND PAGREC = " + Quote(vPagRec, "'")
    strDelete = strDelete + " AND NUMDOC = " + vNumdoc
  Else
    strDelete = strDelete + " AND TIPTIT = 'D' "
    strDelete = strDelete + " AND PAGREC = " + Quote(vPagRec, "'")
    strDelete = strDelete + " AND NUMDOC = " + vNumdoc
    strDelete = strDelete + " AND CODEMP = " + Quote(vCodEmp, "'")
    strDelete = strDelete + " AND NUMDOC = " + Quote(vTipReg, "'")
    strDelete = strDelete + " AND NUMPAR = " + vNumPar
  End If
  conexao.Execute (strDelete)
    
  'gera novamente todos os registro
  '
  Set rs = New ADODB.Recordset

  rs.Open strSql, conexao, adOpenStatic, adLockOptimistic
  Dim i As Integer
  i = 1
  Do While Not rs.EOF
    
    strInsert = "INSERT INTO RATFINANCEIRO ( "
    strInsert = strInsert + " TIPTIT, PAGREC, NUMDOC, "
    strInsert = strInsert + " CODEMP, TIPREG, NUMPAR, "
    strInsert = strInsert + " SEQRAT, CODCTF, CODPRJ, "
    strInsert = strInsert + " PERCTF, VLRCTF, CODCCU, "
    strInsert = strInsert + " PERCCU, VLRCCU ) "
    strInsert = strInsert + " VALUES ("
    
    strInsert = strInsert + Quote(vTipTit, "'") + ", "
    strInsert = strInsert + Quote(vPagRec, "'") + ", "
    strInsert = strInsert + CStr(vNumdoc) + ", "
    
    'especifico por tipo de titulo
    If vTipTit = "L" Then
      strInsert = strInsert + Quote(" ", "'") + ", "
      strInsert = strInsert + Quote(" ", "'") + ", "
      strInsert = strInsert + "0,"
    Else
      strInsert = strInsert + Quote(vCodEmp, ",") + ", "
      strInsert = strInsert + Quote(vTipReg, ",") + ", "
      strInsert = strInsert + CStr(vNumPar) + ", "
    End If
    
    strInsert = strInsert + CStr(i) + ", "
    strInsert = strInsert + CStr(rs![codctf]) + ", "
    strInsert = strInsert + CStr(rs![codprj]) + ", "
    strInsert = strInsert + CStr(rs![PERCTF]) + ", "
    strInsert = strInsert + CStr(rs![VLRCTF]) + ", "
    strInsert = strInsert + CStr(rs![CODCCU]) + ", "
    strInsert = strInsert + CStr(rs![PERCCU]) + ", "
    strInsert = strInsert + CStr(rs![VLRCCU]) + " "
    
    strInsert = strInsert + ")"
    
    conexao.Execute (strInsert)
    
    rs.MoveNext
    i = i + 1
  Loop
  
  Set rs = Nothing
  
  conexao.CommitTrans
  
  Exit Sub
Trata_Erro:
    conexao.RollbackTrans
    Err.Raise Err.Number, " GravaRateio " + Err.Source
  
End Sub


Private Sub OKButton_Click()
   If ValidaRateio Then
     GravaRateio
     Unload Me
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oHelpHtml As New clsHelp
    If KeyCode = vbKeyF1 Then
        oHelpHtml.Origem = 0
        oHelpHtml.hWnd = Me.hWnd
        oHelpHtml.HelpContext = Me.HelpContextID
        Call oHelpHtml.ShowHelp
        Set oHelpHtml = Nothing
    End If
End Sub
