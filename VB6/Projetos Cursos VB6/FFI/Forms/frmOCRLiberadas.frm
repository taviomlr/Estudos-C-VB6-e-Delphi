VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmOCRLiberadas 
   KeyPreview      =   -1  'True
   Caption         =   "Ordens de Carregamento Concluidas"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   11220
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3285
      Top             =   3285
   End
   Begin VB.Frame Frame2 
      Height          =   3210
      Left            =   0
      TabIndex        =   3
      Top             =   -45
      Width           =   11220
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdOrdens 
         Height          =   2805
         Left            =   135
         TabIndex        =   4
         Top             =   270
         Width           =   10950
         _ExtentX        =   19315
         _ExtentY        =   4948
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   3105
      Width           =   11220
      Begin VB.CommandButton cmdAlterarDados 
         Caption         =   "Altera Dados"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1260
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdNF 
         Caption         =   "Gerar NF"
         Enabled         =   0   'False
         Height          =   375
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   1099
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   10035
         TabIndex        =   1
         Top             =   180
         Width           =   1099
      End
   End
End
Attribute VB_Name = "frmOCRLiberadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strCabecalho As String

Private Sub cmdAlterarDados_Click()
    Dim lngOrdem As Long
    lngOrdem = CLng(grdOrdens.TextMatrix(grdOrdens.Row, 0))
    Call frmManutencaoOCR.setNumeroCarregamento(lngOrdem)
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    Call mostrarForm(frmManutencaoOCR, 2943, True)
    '-------------------------------------------------------------------------------------
    Call PreencheGrid
End Sub

Private Sub cmdNF_Click()
    Dim lngOrdem As Long
    Dim oOrdemCarregamento As New COrdCarregamento
    
    'pt. 86728 - Moacir Pfau(11/06/2008) - CLIENTE BLOQUEADO.
    If Not (fEmpresaBloqueada(CStr(grdOrdens.TextMatrix(grdOrdens.Row, 2)), CDate(Format(Now, "DD/MM/YYYY")))) Then
        Exit Sub
    End If
    
    lngOrdem = CLng(grdOrdens.TextMatrix(grdOrdens.Row, 0))
    oOrdemCarregamento.Carregar (lngOrdem)
    Call frmConfirmaItensNF.setNumeroOrdem(lngOrdem)
    
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    If ConfigSys.TipoFaturamento = "Faturamento Antigo" Then
        Call mostrarForm(frmConfirmaItensNF, 2942, False)
        Call PreencheGrid
    ElseIf GetFieldValue("nr_nota_fiscal_inicial", "FVFControle_NotaFiscal", "tp_registro = '" & oOrdemCarregamento.tipoPedidoVenda & "'") > 0 Then
        Call mostrarForm(frmConfirmaItensNF, 2942, False)
        Call PreencheGrid
    Else
        MsgBox "Número sequencial do tipo global " & oOrdemCarregamento.tipoPedidoVenda & " não cadastrado!" _
                & Chr(13) & " Cadastre um o número sequencial", vbInformation, "Verifica Número Sequencial"
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    strCabecalho = "campo=NUMOCR;label=Numero;tamanho=600|" & _
                    "campo=PDVNUM;label=Pedido;tamanho=600|" & _
                    "campo=EMPRESA;label=Cliente;tamanho=1800|" & _
                    "campo=CODPRO;label=Cod;tamanho=700|" & _
                    "campo=linha01;label=Produto;tamanho=2500|" & _
                    "campo=PESTAR;label=Tara(KG);tamanho=1100|" & _
                    "campo=PESLIQ;label=Pes.Líquido(KG);tamanho=1300|" & _
                    "campo=PESBRU;label=Pes.Bruto(KG);tamanho=1200|" & _
                    "campo=PLAVEI;label=Placa;tamanho=1000|" & _
                    "campo=NOMMOT;label=Motorista;tamanho=1500|" & _
                    "campo=UNIMED;label=U.M;tamanho=900|" & _
                    "campo=QTDPRO;label=Quantidade;tamanho=900"
    Me.Height = 4170
    Me.Width = 11340
    CenterForm Me
    Call PreencheGrid
End Sub

Private Sub PreencheGrid()
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "ORDCARREGAMENTO, Produtos, [Pedidos de Venda] as PDV"
    cmd.SelectClause = "NUMOCR,PLAVEI,NOMMOT,linha01,UNIMED,QTDPRO,EMPRESA,PESBRU,PESLIQ,PESTAR,PDVNUM, CODPRO"
    Call cmd.Filter.Append("ORDCARREGAMENTO.CODPRO = Produtos.[Código]")
    Call cmd.Filter.Append("ORDCARREGAMENTO.PDVNUM = PDV.[Número]")
    Call cmd.Filter.Append("ORDCARREGAMENTO.PDVTIP = PDV.[Tipo de Registro]")
    Call cmd.Filter.Append("ORDCARREGAMENTO.PDVFOR = PDV.Fornecedor")
    Call cmd.Filter.Append("ORDCARREGAMENTO.SITOCR=3")
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
 
    Call CarregaHFlexGrid(grdOrdens, rdResult.GetRecordset, strCabecalho)
    grdOrdens.SelectionMode = flexSelectionByRow
    
    rdResult.CloseReader
    Set rdResult = Nothing
    Aplicacao.Disconnect
End Sub

Private Sub grdOrdens_Click()
    With grdOrdens
        cmdNF.Enabled = .TextMatrix(.Row, 0) <> "" And .TextMatrix(.Row, 0) <> "Numero"
        cmdAlterarDados.Enabled = .TextMatrix(.Row, 0) <> "" And .TextMatrix(.Row, 0) <> "Numero"
    End With
End Sub

Private Sub Timer1_Timer()
    Dim intLinha As Integer
    intLinha = grdOrdens.Row
    Call PreencheGrid
    If grdOrdens.Rows > intLinha Then
        grdOrdens.Row = intLinha
    Else
        grdOrdens.Row = 0
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
