VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmRetornoArquivoPagamento 
   Caption         =   "Retorno do Arquivo de Pagamento"
   ClientHeight    =   7365
   ClientLeft      =   4050
   ClientTop       =   4245
   ClientWidth     =   13305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   13305
   Begin VB.Frame fraBaixa 
      Caption         =   "Operação Contábil - Baixa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   45
      TabIndex        =   30
      Top             =   1890
      Width           =   11775
      Begin Fox.EBSText etxOperContabilDupl 
         Height          =   330
         Left            =   1410
         TabIndex        =   10
         Top             =   240
         Width           =   3495
         _extentx        =   429392
         _extenty        =   582
         font            =   "frmRetornoArquivoPagamento.frx":0000
         tipotexto       =   0
         maxlength       =   5
         possuidescricao =   -1  'True
         campocriterio   =   "cd_operacao"
         tipocriterio    =   4
         campodescricao  =   "descricao"
         tabelaconsulta  =   "OperacaoContabil"
         tamanhodescricao=   2800
         alinhamento     =   1
      End
      Begin Fox.EBSText etxOperContabilLanc 
         Height          =   330
         Left            =   6060
         TabIndex        =   11
         Top             =   240
         Width           =   3495
         _extentx        =   429392
         _extenty        =   582
         font            =   "frmRetornoArquivoPagamento.frx":002C
         tipotexto       =   0
         maxlength       =   5
         possuidescricao =   -1  'True
         campocriterio   =   "cd_operacao"
         tipocriterio    =   4
         campodescricao  =   "descricao"
         tabelaconsulta  =   "OperacaoContabil"
         tamanhodescricao=   2800
         alinhamento     =   1
      End
      Begin VB.Label lblBanco 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lançamentos"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   4
         Left            =   5010
         TabIndex        =   32
         Top             =   300
         Width           =   960
      End
      Begin VB.Label lblBanco 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Duplicatas"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   3
         Left            =   570
         TabIndex        =   31
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.Frame Frame 
      Height          =   7305
      Left            =   11850
      TabIndex        =   28
      Top             =   30
      Width           =   1425
      Begin VB.CommandButton cmdBaixar 
         Caption         =   "&Baixar Títulos"
         Height          =   375
         Left            =   80
         TabIndex        =   14
         Top             =   570
         Width           =   1275
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   80
         TabIndex        =   15
         Top             =   960
         Width           =   1275
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   80
         TabIndex        =   16
         Top             =   1350
         Width           =   1275
      End
      Begin VB.CommandButton cmdAnalisar 
         Caption         =   "&Validar arquivo"
         Height          =   375
         Left            =   80
         TabIndex        =   9
         Top             =   180
         Width           =   1275
      End
      Begin MSComDlg.CommonDialog cmDialog 
         Left            =   450
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imgCheck 
         Left            =   390
         Top             =   4260
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRetornoArquivoPagamento.frx":0058
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRetornoArquivoPagamento.frx":03AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRetornoArquivoPagamento.frx":06FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRetornoArquivoPagamento.frx":0A16
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraSelecao 
      Caption         =   "Seleção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   45
      TabIndex        =   26
      Top             =   2550
      Width           =   11775
      Begin VB.CommandButton cmdMarcarTodos 
         Caption         =   "&Todos"
         Height          =   375
         Left            =   8550
         TabIndex        =   12
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdDesmarcarTodos 
         Caption         =   "&Nenhum"
         Height          =   375
         Left            =   10050
         TabIndex        =   13
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informações do Arquivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   1845
      Left            =   45
      TabIndex        =   17
      Top             =   30
      Width           =   11775
      Begin VB.TextBox txtCarteira 
         Height          =   330
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1050
         Width           =   675
      End
      Begin VB.TextBox txtDVConta 
         Height          =   330
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   690
         Width           =   375
      End
      Begin VB.TextBox txtDVAgencia 
         Height          =   330
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   690
         Width           =   375
      End
      Begin VB.TextBox txtConta 
         Height          =   330
         Left            =   8430
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   690
         Width           =   1005
      End
      Begin VB.TextBox txtAgencia 
         Height          =   330
         Left            =   4470
         TabIndex        =   2
         Top             =   690
         Width           =   1005
      End
      Begin Fox.EBSText etxEmpresa 
         Height          =   330
         Left            =   4470
         TabIndex        =   7
         Top             =   1050
         Width           =   5430
         _extentx        =   440531
         _extenty        =   582
         font            =   "frmRetornoArquivoPagamento.frx":0D30
         tipo            =   4
         tipotexto       =   0
         maxlength       =   15
         possuidescricao =   -1  'True
         campocriterio   =   "Apel"
         campodescricao  =   "Razão"
         tabelaconsulta  =   "Empresas"
         tamanhodescricao=   4000
      End
      Begin Fox.EBSText etxBanco 
         Height          =   330
         Left            =   1410
         TabIndex        =   0
         Top             =   330
         Width           =   6960
         _extentx        =   160073
         _extenty        =   582
         font            =   "frmRetornoArquivoPagamento.frx":0D5C
         tipotexto       =   0
         maxlength       =   9
         possuidescricao =   -1  'True
         campocriterio   =   "Banco"
         tipocriterio    =   4
         campodescricao  =   "Nome"
         tabelaconsulta  =   "Bancos"
         tamanhodescricao=   6000
         alinhamento     =   1
      End
      Begin Fox.EBSText etxCamara 
         Height          =   330
         Left            =   1410
         TabIndex        =   1
         Top             =   690
         Width           =   690
         _extentx        =   265
         _extenty        =   582
         font            =   "frmRetornoArquivoPagamento.frx":0D88
         tipotexto       =   0
         maxlength       =   3
         tipocriterio    =   0
         alinhamento     =   1
      End
      Begin Fox.EBSArquivo etxDiretorio 
         Height          =   330
         Left            =   1410
         TabIndex        =   8
         Top             =   1410
         Width           =   8475
         _extentx        =   14949
         _extenty        =   582
         tipotratamento  =   2
         filter          =   ""
      End
      Begin VB.Label lblBanco 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   29
         Top             =   390
         Width           =   1155
      End
      Begin VB.Label lblCarteira 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carteira"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   24
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lblEmpresa 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código Empresa"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   3
         Left            =   3225
         TabIndex        =   23
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lblConta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   2
         Left            =   7185
         TabIndex        =   22
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label lblAgencia 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agência"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   1
         Left            =   3225
         TabIndex        =   21
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label lblBanco 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Câmara"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   20
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label lblBanco 
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Width           =   3705
      End
      Begin VB.Label lblArquivo 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo Retorno"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   165
         TabIndex        =   18
         Top             =   1470
         Width           =   1155
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResultado 
      Height          =   3525
      Left            =   45
      TabIndex        =   25
      Top             =   3210
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6218
      _Version        =   393216
      FixedCols       =   0
      BackColorUnpopulated=   -2147483625
      GridColor       =   -2147483631
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Image imgInformativa 
      Height          =   480
      Left            =   105
      Picture         =   "frmRetornoArquivoPagamento.frx":0DB4
      Top             =   6795
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmRetornoArquivoPagamento.frx":19F6
      Height          =   525
      Left            =   60
      TabIndex        =   27
      Top             =   6780
      Width           =   11745
   End
End
Attribute VB_Name = "frmRetornoArquivoPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mobjPagFor As clsPagFor
Private mobjArquivo As cArquivoTexto
Private mobjMatrizContabil As cMatrizContabilizacao
Private mobjMatrizContabilDAO As cMatrizContabilizacaoDAO
Private Const grdChecked = 2
Private Const grdUnchecked = 1
Private Const grdAsc = 4
Private Const grdDesc = 3

Private Sub cmdDirGeracao_Click()
    Dim strDiretorio As String
    
    strDiretorio = FolderDialogBox(Me.hWnd, "Diretório que será criado o arquivo de Remessa", rfcDesktop, bfReturnDirs)
    txtDiretorio.Text = strDiretorio
    txtDiretorio.SetFocus
End Sub

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdAnalisar_Click()
    'If UCase(Right(etxDiretorio.Valor, 10)) = "PAGFOR.TXT" Then
        If ValidaDocumento Then
            Call MostraRegistros
            If etxCamara.valorInteiro <> 237 Then
                cmdBaixar.Enabled = True
                fraBaixa.Enabled = True
            End If
        End If
'    Else
'        MsgBox "O arquivo selecionado não se refere ao retorno de Pagamento de Fornecedores.", vbInformation, NomeModulo
'        cmdBaixar.Enabled = False
'        fraBaixa.Enabled = False
'    End If
End Sub

Private Sub cmdBaixar_Click()
    Dim strErro As String
    Dim blnOperCont As Boolean
    
    If ValidaBaixas(blnOperCont) Then
        If BaixarTitulos(strErro) Then
            MsgBox "Título(s) baixado(s) com sucesso.", vbInformation, NomeModulo
        ElseIf strErro <> "" Then
            MsgBox "Erro ao Baixar Título(s): " & strErro, vbInformation, NomeModulo
        Else
            MsgBox "As transferências não podem ser baixadas." & strErro, vbInformation, NomeModulo
        End If
    Else
        If Not blnOperCont Then
            MsgBox "Selecione pelo menos um registro para baixar.", vbInformation, NomeModulo
        End If
    End If
End Sub

Private Sub cmdDesmarcarTodos_Click()
    Dim intCont As Integer
    
    With grdResultado
        If .TextMatrix(1, 2) <> "" Then
            For intCont = 1 To .Rows - 1
                .Row = intCont
                Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            Next
        End If
    End With
End Sub

Private Sub cmdMarcarTodos_Click()
    Dim intCont As Integer
    
    With grdResultado
        If .TextMatrix(1, 2) <> "" Then
            For intCont = 1 To .Rows - 1
                .Row = intCont
                Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
            Next
        End If
    End With
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub etxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxBanco.ValorDescricao = "" Then
            etxBanco.valorInteiro = 0
        End If
        Call PCampo("Bancos", "SELECT Banco, Nome FROM Bancos", pbCampo, etxBanco, "Banco")
    End If
End Sub

Private Sub etxBanco_LostFocus()
    Dim rstResult As Object
    
    Call LimpaCampos
    If etxBanco.ValorDescricao <> "" Then
        If AbreRecordset(rstResult, "SELECT * FROM Bancos WHERE Banco = " & etxBanco.valorInteiro) = WL_OK Then
            etxCamara.valorInteiro = GetValue(rstResult, "Câmara", 0)
            txtCarteira.Text = GetValue(rstResult, "Carteira", "")
            txtAgencia.Text = GetValue(rstResult, "Número Agência", 0)
            txtDVAgencia.Text = GetValue(rstResult, "Dígito da Agência", 0)
            txtConta.Text = GetValue(rstResult, "Número Conta", 0)
            txtDVConta.Text = GetValue(rstResult, "Dígito da Conta", 0)
            etxEmpresa.valorTexto = GetValue(rstResult, "Cedente", 0)
        End If
    End If
End Sub

Private Sub etxOperContabilDupl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxOperContabilDupl.ValorDescricao = "" Then
            etxOperContabilDupl.valorInteiro = 0
        End If
        Call PCampo("Operação Contábil", "SELECT * FROM OperacaoContabil", pbCampo, etxOperContabilDupl, "cd_operacao")
    End If
End Sub

Private Sub etxOperContabilLanc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxOperContabilLanc.ValorDescricao = "" Then
            etxOperContabilLanc.valorInteiro = 0
        End If
        Call PCampo("Operação Contábil", "SELECT * FROM OperacaoContabil", pbCampo, etxOperContabilLanc, "cd_operacao")
    End If
End Sub

Private Sub Form_Load()
    Call etxEmpresa.AddConexao(Aplicacao)
    Call etxBanco.AddConexao(Aplicacao)
    Call etxCamara.AddConexao(Aplicacao)
    Call etxOperContabilDupl.AddConexao(Aplicacao)
    Call etxOperContabilLanc.AddConexao(Aplicacao)
    Call ConfigureGrid
    etxOperContabilDupl.Enabled = ConfigSys.UtilizaIntegracaoContabil
    etxOperContabilLanc.Enabled = ConfigSys.UtilizaIntegracaoContabil
    If ConfigSys.UtilizaIntegracaoContabil Then
        Set mobjMatrizContabil = New cMatrizContabilizacao
        Set mobjMatrizContabilDAO = New cMatrizContabilizacaoDAO
    End If
    cmdBaixar.Enabled = False
    fraBaixa.Enabled = False
End Sub

Private Sub LimpaCampos()
    etxCamara.Clear
    txtCarteira.Text = ""
    txtAgencia.Text = ""
    txtDVAgencia.Text = ""
    txtConta.Text = ""
    txtDVConta.Text = ""
    etxEmpresa.Clear
End Sub

Private Sub ConfigureGrid()
    Dim intColuna As Integer

    With grdResultado
        .Rows = 2
        .FixedRows = 1
        .Cols = 12
        .FixedCols = 1
        
        'Coluna Fixa
        .ColWidth(0) = 150
        .TextMatrix(0, 0) = ""
        
        'Configura a coluna de seleção
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        .ColAlignment(1) = flexAlignCenterCenter
        
        'Coluna Documento
        .ColWidth(2) = 500
        .TextMatrix(0, 2) = "Doc."
        .ColAlignment(2) = flexAlignLeftCenter
        
        'Coluna Tipo de Registro
        .ColWidth(3) = 1200
        .TextMatrix(0, 3) = "Tipo"
        .ColAlignment(3) = flexAlignLeftCenter
        
        'Coluna Número
        .ColWidth(4) = 800
        .TextMatrix(0, 4) = "Número"
        .ColAlignment(4) = flexAlignRightCenter
        
        'Coluna Parcela
        .ColWidth(5) = 600
        .TextMatrix(0, 5) = "Parc."
        .ColAlignment(5) = flexAlignRightCenter
        
        'Coluna Empresa
        .ColWidth(6) = 2300
        .TextMatrix(0, 6) = "Empresa"
        .ColAlignment(6) = flexAlignLeftCenter

        'Coluna Emissão
        .ColWidth(7) = 1050
        .TextMatrix(0, 7) = "Vencimento"
        .ColAlignment(7) = flexAlignCenterCenter
        
        'Coluna Emissão
        .ColWidth(8) = 1050
        .TextMatrix(0, 8) = "Pagamento"
        .ColAlignment(8) = flexAlignCenterCenter
        
        'Coluna Valor
        .ColWidth(9) = 1150
        .TextMatrix(0, 9) = "Valor"
        .ColAlignment(9) = flexAlignRightCenter
        
        'Proveniente Rateio
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = "Rateio"
        
        .ColWidth(11) = 2350
        .TextMatrix(0, 11) = "Status Banco"
        .ColAlignment(11) = flexAlignLeftCenter
        
        For intColuna = 0 To .Cols - 1
            If intColuna = 1 Then
                Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            End If
            .TextMatrix(1, intColuna) = ""
        Next
    End With
End Sub

Private Function ValidaDocumento() As Boolean
    Dim strErro As String
    Dim colErros As New Collection
    Dim intCont As Integer
    
    Set mobjPagFor = New clsPagFor
    Set mobjArquivo = New cArquivoTexto
    
    mobjPagFor.objArquivo = mobjArquivo
    If etxCamara.valorInteiro > 0 Then
        mobjPagFor.CamaraOrigem = etxCamara.valorInteiro
    End If
    
    'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(25/04/2012)
    If etxDiretorio.Valor <> "" Then
        mobjPagFor.BancoDestino = etxBanco.valorInteiro
        ValidaDocumento = mobjPagFor.ValidaArquivo(strToLng(txtAgencia.Text), txtDVAgencia.Text, strToLng(txtConta.Text), txtDVConta.Text, txtCarteira.Text, etxDiretorio.Valor, strErro)
        Set colErros = mobjPagFor.colErros
        If Len(Trim(strErro)) > 0 Then
            MsgBox strErro, vbInformation, NomeModulo
        End If
        If colErros.Count > 0 Then
            For intCont = 1 To colErros.Count
                strErro = strErro & vbNewLine & colErros.item(intCont)
            Next
            MsgBox "Foram encontradas inconsistências em lotes do arquivo:" & strErro
        End If
    Else
        'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(25/04/2012)
        MsgBox "O campo 'arquivo retorno' deve possuir um caminho válido:" & strErro
    End If
End Function

Private Sub MostraRegistros()
    Dim intCont         As Integer
    Dim ColDocumentos   As Collection
    Dim strCodigos()    As String
    Dim rstResult       As Object
    Dim rstResultCC     As Object
    Dim strTabela       As String
    Dim strRegistro     As String
    Dim strCampoCodigo  As String
    Dim blnRateio       As Boolean
    Dim blnDuplicatas   As Boolean
    Dim blnLancamentos  As Boolean
    Dim strTipoGlobal   As String
    Dim strStatusBanco  As String
    
    Set ColDocumentos = mobjPagFor.ColDocumentos
    strTipoGlobal = ""
    Call ConfigureGrid
    For intCont = 1 To ColDocumentos.Count
        strCodigos = Split(ColDocumentos(intCont), ";")
        If UBound(strCodigos) > 4 Then
            strStatusBanco = IIf(strCodigos(5) = "Agendado", "Confirmação de Agendamento", "Confirmação de Pagamento")
        Else
            strStatusBanco = "Confirmação de Pagamento"
        End If
        Call AbreRecordset(rstResult, "SELECT tipo_registro,nr_documento,nr_parcela,cd_empresa,tp_documento,dt_vencimento FROM FFIItemPagamento WHERE cd_arquivoPagamento = " & strCodigos(0) & " AND cd_lotePagamento = " & strCodigos(1) & " AND cd_itemPagamento = " & strCodigos(2))
        
        If Not rstResult.EOF Then
            Select Case UCase(GetValue(rstResult, "tp_documento"))
                Case "DUP"
                    strTabela = "Duplicatas"
                    strCampoCodigo = "Nota"
                    blnDuplicatas = True
                Case "LAN"
                    strTabela = "[Lançamentos]"
                    strCampoCodigo = "Código"
                    blnLancamentos = True
            End Select
            If Trim(strTipoGlobal) = "" Then
                strTipoGlobal = GetValue(rstResult, "Tipo_registro", "")
            End If
            blnRateio = CBool(GetFieldValue("proveniente_rateio", strTabela, "PagRec = 'P' AND " & strCampoCodigo & "=" & GetValue(rstResult, "nr_documento") & " AND Empresa = '" & GetValue(rstResult, "cd_empresa") & "' AND Tipo = '" & GetValue(rstResult, "tipo_registro") & "' AND Parcela = " & GetValue(rstResult, "nr_parcela")))
            strRegistro = "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & rstResult.Fields("tp_documento").value & Chr(vbKeyTab) & rstResult.Fields("tipo_registro").value & _
                    Chr(vbKeyTab) & rstResult.Fields("nr_documento").value & Chr(vbKeyTab) & Format(rstResult.Fields("nr_parcela").value, "000") & _
                    Chr(vbKeyTab) & rstResult.Fields("cd_empresa").value & Chr(vbKeyTab) & rstResult.Fields("dt_vencimento").value & Chr(vbKeyTab) & strCodigos(3) & Chr(vbKeyTab) & _
                    Format(strCodigos(4), "#,##0.00") & Chr(vbKeyTab) & blnRateio & Chr(vbKeyTab) & strStatusBanco
                    
            Call grdResultado.AddItem(strRegistro)
            grdResultado.Row = grdResultado.Rows - 1
            grdResultado.col = 1
            Set grdResultado.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            strRegistro = ""
            If etxCamara.valorInteiro = 237 Then
                cmdBaixar.Enabled = (strCodigos(5) <> "Agendado")
                fraBaixa.Enabled = cmdBaixar.Enabled
            End If
        End If
    Next
    If grdResultado.Rows > 2 And grdResultado.TextMatrix(1, 1) = "" Then
        Call grdResultado.RemoveItem(1)
    End If
    If ConfigSys.UtilizaIntegracaoContabil Then
        If Trim(strTipoGlobal) <> "" Then
            Set mobjMatrizContabil = mobjMatrizContabilDAO.Carregar(strTipoGlobal)
            If Not mobjMatrizContabil Is Nothing Then
                If blnDuplicatas And etxOperContabilDupl.Enabled Then
                    etxOperContabilDupl.valorInteiro = mobjMatrizContabil.BaixaDuplicatasPagar
                End If
                If blnLancamentos And etxOperContabilLanc.Enabled Then
                    etxOperContabilLanc.valorInteiro = mobjMatrizContabil.BaixaLancamentosPagar
                End If
            Else
                Call MsgBox("Não foi possível localizar a Matriz de Contabilização para o Tipo Global " & strTipoGlobal & ".", vbInformation, NomeModulo)
            End If
        End If
    End If
End Sub

Private Sub grdResultado_Click()
    With grdResultado
        If .col = 1 Then
            If .TextMatrix(.Row, 2) <> "" Then
                If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                    Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                Else
                    Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                End If
            End If
        End If
    End With
End Sub

Private Function ValidaBaixas(ByRef blnOperCont As Boolean) As Boolean
    Dim intCont As Integer
    
    If etxOperContabilDupl.Enabled Then
        If etxOperContabilDupl.ValorDescricao = "" Then
            MsgBox "Para baixas é necessário informar uma operção contábil.", vbInformation, NomeModulo
            blnOperCont = True
            ValidaBaixas = False
            Exit Function
        End If
    End If
    With grdResultado
        For intCont = 1 To .Rows - 1
            .Row = intCont
            .col = 1
            If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                ValidaBaixas = True
                Exit Function
            End If
        Next
    End With
End Function

Private Function BaixarTitulos(ByRef strErro As String) As Boolean
    Dim intCont As Integer
    Dim lngOperBaixa As Long
    Dim strSQLUpdate As String
    Dim strTabela As String
    Dim strCampoCod As String
    Dim intContRegistros As Integer
    
On Error GoTo err_Handler
    With grdResultado
        .col = 1
        intContRegistros = 0
        For intCont = 1 To .Rows - 1
            .Row = intCont
            If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                If .TextMatrix(.Row, 2) <> "Tra" Then
                    lngOperBaixa = 0
                    intContRegistros = intContRegistros + 1
                    If .TextMatrix(.Row, 2) = "Dup" Then
                        strTabela = "Duplicatas"
                        strCampoCod = "Nota"
                        If ConfigSys.UtilizaIntegracaoContabil Then
                            lngOperBaixa = etxOperContabilDupl.valorInteiro
                        End If
                    Else
                        strTabela = "Lançamentos"
                        strCampoCod = "Código"
                        If .TextMatrix(.Row, 2) = "Lan" Then
                            lngOperBaixa = etxOperContabilLanc.valorInteiro
                        End If
                    End If
                    If CBool(.TextMatrix(.Row, 10)) Then
                        Call BaixaAgrupadosCC(.Row, lngOperBaixa)
                    Else
                        strSQLUpdate = "UPDATE " & strTabela & " SET Pagamento = " & InverteData(.TextMatrix(.Row, 8), True)
                        If lngOperBaixa > 0 Then
                            strSQLUpdate = strSQLUpdate & ", cd_operacao_baixa = " & lngOperBaixa
                        End If
                        strSQLUpdate = strSQLUpdate & " WHERE PagRec = 'P' AND " & strCampoCod & " = " & .TextMatrix(.Row, 4) & _
                        " AND Empresa = '" & .TextMatrix(.Row, 6) & "' AND Tipo = '" & .TextMatrix(.Row, 3) & _
                        "' AND Parcela = " & .TextMatrix(.Row, 5)
                        Call ExecuteSQL(strSQLUpdate)
                    End If
                End If
            End If
        Next
    End With
    If intContRegistros > o Then
        BaixarTitulos = True
    Else
        BaixarTitulos = False
    End If
    Exit Function

err_Handler:
    strErro = err.Description
    BaixarTitulos = False
End Function

Private Sub BaixaAgrupadosCC(intRow As Integer, lngOperBaixa As Long)
    Dim rstResult    As Object
    Dim strTabela    As String
    Dim strCampoCod  As String
    Dim strSQLUpdate As String
    
    With grdResultado
        If .TextMatrix(intRow, 2) = "Dup" Then
            strTabela = "Duplicatas"
            strCampoCod = "Nota"
        Else
            strTabela = "Lançamentos"
            strCampoCod = "Código"
        End If
        If AbreRecordset(rstResult, "SELECT * FROM " & strTabela & " WHERE PagRec = 'P' AND " & strCampoCod & " = " & .TextMatrix(intRow, 4) & " AND Empresa = '" & .TextMatrix(intRow, 6) & "' AND Tipo = '" & .TextMatrix(intRow, 3) & "' AND Vencimento = #" & InverteData(.TextMatrix(intRow, 7)) & "#") = WL_OK Then
            rstResult.MoveFirst
            While Not rstResult.EOF
                strSQLUpdate = "UPDATE " & strTabela & " SET Pagamento = #" & InverteData(.TextMatrix(intRow, 8)) & "#"
                If lngOperBaixa > 0 Then
                    strSQLUpdate = strSQLUpdate & ", cd_operacao_baixa = " & lngOperBaixa
                End If
                strSQLUpdate = strSQLUpdate & " WHERE PagRec = 'P' AND " & strCampoCod & " = " & .TextMatrix(intRow, 4) & _
                               " AND Empresa = '" & .TextMatrix(intRow, 6) & "' AND Tipo = '" & .TextMatrix(intRow, 3) & "' AND Parcela = " & GetValue(rstResult, "Parcela")
                Call ExecuteSQL(strSQLUpdate)
                rstResult.MoveNext
            Wend
        End If
    End With
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

