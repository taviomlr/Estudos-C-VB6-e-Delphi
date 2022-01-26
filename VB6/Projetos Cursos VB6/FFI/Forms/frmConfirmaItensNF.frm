VERSION 5.00
Begin VB.Form frmConfirmaItensNF 
   Caption         =   "Confirma Itens de Notas Fiscais de Saída"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumeroNota 
      BackColor       =   &H80000018&
      Height          =   330
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   1185
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   375
      Left            =   6435
      TabIndex        =   7
      Top             =   4140
      Width           =   1099
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   7650
      TabIndex        =   8
      Top             =   4140
      Width           =   1099
   End
   Begin VB.TextBox txtEmissao 
      Height          =   330
      Left            =   5130
      TabIndex        =   5
      Top             =   540
      Width           =   1455
   End
   Begin VB.TextBox txtTransportadora 
      Height          =   330
      Left            =   5130
      MaxLength       =   5
      TabIndex        =   4
      Top             =   135
      Width           =   555
   End
   Begin VB.TextBox txtFornecedor 
      BackColor       =   &H80000018&
      Height          =   330
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   990
      Width           =   2085
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   225
      Width           =   1680
   End
   Begin Fox.WDBGrid wdbItens 
      Height          =   2310
      Left            =   135
      TabIndex        =   0
      Top             =   1665
      Width           =   8610
      _extentx        =   15187
      _extenty        =   4075
      corfontefixa    =   -2147483634
      backcolor       =   -2147483636
      rowheightmin    =   225
      colwidthmin     =   1440
      rowheight       =   225
      font            =   "frmConfirmaItensNF.frx":0000
      font            =   "frmConfirmaItensNF.frx":002C
      fontefixa       =   "frmConfirmaItensNF.frx":0058
      fontefixa       =   "frmConfirmaItensNF.frx":0086
      scrollbars      =   3
   End
   Begin Fox.EBSText txtOperacaoContabil 
      Height          =   330
      Left            =   5130
      TabIndex        =   6
      Top             =   990
      Width           =   825
      _extentx        =   265
      _extenty        =   582
      font            =   "frmConfirmaItensNF.frx":00B4
      tipotexto       =   0
      maxlength       =   5
   End
   Begin VB.Label lblOperacaoContabil 
      AutoSize        =   -1  'True
      Caption         =   "lblOperacaoContabil"
      Height          =   195
      Left            =   5985
      TabIndex        =   16
      Top             =   1035
      Width           =   1425
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Op. Contábil:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3940
      TabIndex        =   15
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Número:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   495
      TabIndex        =   14
      Top             =   675
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Emissão:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4300
      TabIndex        =   13
      Top             =   630
      Width           =   765
   End
   Begin VB.Label lblTransportadora 
      Caption         =   "lblTransportadora"
      Height          =   240
      Left            =   5805
      TabIndex        =   12
      Top             =   180
      Width           =   2940
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Transportadora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3700
      TabIndex        =   11
      Top             =   225
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1035
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   765
      TabIndex        =   9
      Top             =   270
      Width           =   450
   End
End
Attribute VB_Name = "frmConfirmaItensNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lOrdCarregamento As Long
Private oOrdemCarregamento As New COrdCarregamento
Private rsItens As Object
Private colUnidade As Integer
Private ColQuantidade As Integer
Private colProduto As Integer
Private ColNatOperacao As Integer
Private ColNatComplemento As Integer
Private ColCODLOC As Integer
Private ColCODPRJ As Integer
Private ColSeqItem As Integer
Private ColFornecedor As Integer
Private colEmpresa As Integer
Private colVlrUnitario As Integer
Private ColCodCCU As Integer


Private ColUnEnt As Integer
Private ColUnEntCampo1 As Integer
Private ColUnEntCampo2 As Integer
Private ColUnEntCampo3 As Integer
Private ColUnEntQtd As Integer
Private ColUnEntDesc As Integer

Private ColUnAux As Integer
Private ColUnAuxCampo1 As Integer
Private ColUnAuxCampo2 As Integer
Private ColUnAuxCampo3 As Integer
Private ColUnAuxQtd As Integer
Private ColUnAuxDesc As Integer

Private ColUnAux1 As Integer
Private ColUnAux1Campo1 As Integer
Private ColUnAux1Campo2 As Integer
Private ColUnAux1Campo3 As Integer
Private ColUnAux1Qtd As Integer
Private ColUnAux1Desc As Integer

Private mintColQuantidade As Integer

Public Sub setNumeroOrdem(lNewNumeroOrdem As Long)
    lOrdCarregamento = lNewNumeroOrdem
    If oOrdemCarregamento.Carregar(lOrdCarregamento) Then
        Set rsItens = oOrdemCarregamento.item
        
        'Pt.97090  - Fernando Paludo - (22/02/2010)
        'Inserido para corrigir o problema de não atualização dos dados na grid.
        Sleep (2000)
        DoEvents
        
        Call PreparaGrid
        With wdbItens
          #If FOXSQL = 0 Then
            .RecordRelation = rsItens.name
            .RecordSource = rsItens.name
            ConfigGrid wdbItens, rsItens.name
          #Else
             'Projeto: #1085 - Problema#4939 - Moacir Pfau(07/07/2011)
            .RecordRelation = Replace(rsItens.Source, "SELECT * FROM ", "")
            .RecordSource = Replace(rsItens.Source, "SELECT * FROM ", "")
            ConfigGrid wdbItens, Replace(rsItens.Source, "SELECT * FROM ", "")
          #End If
          .TipoAcesso = wdbCompleto
          .Refresh
        End With
    End If
End Sub

Private Sub PreparaGrid()

With wdbItens
      Dim i As Integer
      
      Dim CampoItem As Object
      Dim HabUnPersonalizadas As Boolean
      
      i = 1
      
      For Each CampoItem In rsItens.Fields
        
          
        .AddCol CampoItem.name
        .ColCaption(i) = CampoItem.name
        .ColEnabled(i) = False
        .ColFormato(i) = GetFormat(rsItens, CampoItem.name)
        .ColStyle(i) = wdbEdite
                
        '
        ' Númericos são alinhados a direita
        '
        If EIgual(TransDBTypeRetInt(CampoItem.Type), dbSingle, dbDouble, dbByte, dbInteger, dbLong) Then .ColAlin(i) = wdbdireito
                
  
        Select Case CStr(CampoItem.name)
        
        Case "Quantidade"
          .ColRequerida(i) = True
          'pt. 83255 - Dulcino Júnior(Alteração para exibir a quantidade de casas configurada)
          .ColFormato(i) = FCASAS(Configuracao("Decimais de Quantidade para Vendas", strModulo:="KIV"))
          mintColQuantidade = i
          ColQuantidade = i
          i = i + 1
          
        Case "CODLOC"
          .ColRequerida(i) = True
          .ColCaption(i) = "Local"
          ColCODLOC = i
          i = i + 1
          
        Case "CODPRJ"
          .ColRequerida(i) = True
          ColCODPRJ = i
          i = i + 1
          
        Case "Tabela de Preços"
          If Not TabelasPrecos Then
            .RemoveCol i
          Else
            i = i + 1
          End If
        Case "Produto"
            colProduto = i
            i = i + 1
            
        Case "Descrição", "Complemento", "Unidade", "Marca", "Situação Tributária"
            i = i + 1
            
        
        Case "Quantidade Baixada", "Data da Baixa"
          .RemoveCol i
        
        Case "ISS"
            i = i + 1
          
        Case "IRF"
            i = i + 1
  
        Case "Serviço", "Descrição do Serviço"
            i = i + 1
          
        Case "Fornecedor"
          ColFornecedor = i
           i = i + 1
               
        Case "Centro de Custo"
          ColCodCCU = i
          i = i + 1

          
        Case "Tipo do Pedido/Orçamento", "Número do Pedido/Orçamento", "Item do Pedido/Orçamento"
          .RemoveCol i
          
        Case "Cadastrado por", "Duplicatas Liberado por"
          .RemoveCol i
          
        Case "VlrUUV"
          If Not HabUnPersonalizadas Then
            .RemoveCol i
          Else
            .ColCaption(i) = "Vlr. Unitário Un. Venda"
            .ColWidth(i) = 1800
            .ColEnabled(i) = False
            
            i = i + 1
          End If
          
        Case "UnEnt", "UnEntQtd", "UnEntCampo1", "UnEntCampo2", "UnEntCampo3", "UnEntDesc", "UnAux", "UnAuxQtd", "UnAuxCampo1", "UnAuxCampo2", "UnAuxCampo3", "UnAuxDesc", "UnAux1", "UnAux1Qtd", "UnAux1Campo1", "UnAux1Campo2", "UnAux1Campo3", "UnAux1Desc"
        
          If Not HabUnPersonalizadas Then
            .RemoveCol i
          Else
            Dim minColWidth As Double
            minColWidth = .ColWidthMin
            
            .ColWidthMin = 0.1
            Select Case CStr(CampoItem.name)
                Case "UnEnt"
                    .ColCaption(i) = "Un. Entrada"
                    .ColEnabled(i) = False
                Case "UnEntQtd"
                    .ColCaption(i) = "Qtd. Un. Ent."
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnEntCampo1"
                    .ColCaption(i) = "Un. Ent. Campo 1"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnEntCampo2"
                    .ColCaption(i) = "Un. Ent. Campo 2"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnEntCampo3"
                    .ColCaption(i) = "Un. Ent. Campo 3"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnEntDesc"
                    .ColCaption(i) = "Desc. Un. Ent."
                    .ColEnabled(i) = False
                Case "UnAux"
                    .ColCaption(i) = "Un. Aux."
                    .ColEnabled(i) = False
                    .ColWidth(i) = 1440
                Case "UnAuxQtd"
                    .ColCaption(i) = "Qtd. Un. Aux."
                    .ColEnabled(i) = False
                Case "UnAuxCampo1"
                    .ColCaption(i) = "Un. Aux. Campo 1"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnAuxCampo2"
                    .ColCaption(i) = "Un. Aux. Campo 2"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnAuxCampo3"
                    .ColCaption(i) = "Un. Aux. Campo 3"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnAuxDesc"
                    .ColCaption(i) = "Desc. Un. Aux."
                    .ColEnabled(i) = False
                Case "UnAux1"
                    .ColCaption(i) = "Un. Aux1."
                    .ColEnabled(i) = False
                Case "UnAux1Qtd"
                    .ColCaption(i) = "Qtd. Un. Aux1."
                    .ColEnabled(i) = False
                Case "UnAux1Campo1"
                    .ColCaption(i) = "Un. Aux1. Campo 1"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnAux1Campo2"
                    .ColCaption(i) = "Un. Aux1. Campo 2"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnAux1Campo3"
                    .ColCaption(i) = "Un. Aux1. Campo 3"
                    .ColEnabled(i) = False
                    .ColWidth(i) = 0.1
                Case "UnAux1Desc"
                    .ColCaption(i) = "Desc. Un. Aux1."
                    .ColEnabled(i) = False
            End Select
            i = i + 1
            
            .ColWidthMin = minColWidth
            
          End If
        Case "Natureza de Operação"
        
          .ColEnabled(i) = True
          ColNatOperacao = i
          i = i + 1
        Case "Complemento da Natureza"
          ColNatComplemento = i
          i = i + 1
        Case "Item"
          ColSeqItem = i
          i = i + 1
        Case "Valor Original"
          colVlrUnitario = i
          i = i + 1
          
        Case "Empresa"
          colEmpresa = i
          i = i + 1
          
        Case Else
          i = i + 1
        End Select
        
      Next
    End With
End Sub

Private Sub cboTipo_Click()
    Dim DAOMatriz As cMatrizContabilizacaoDAO
    Dim objMatriz As cMatrizContabilizacao
    
    If txtOperacaoContabil.Enabled Then
        If cboTipo.ListIndex <> -1 Then
            Set DAOMatriz = New cMatrizContabilizacaoDAO
            Set objMatriz = DAOMatriz.Carregar(cboTipo.Text)
            If Not objMatriz Is Nothing Then
                txtOperacaoContabil.valorInteiro = objMatriz.NotaFiscalSaida
            Else
                txtOperacaoContabil.valorInteiro = 0
            End If
        End If
    Else
        txtOperacaoContabil.valorInteiro = 0
    End If
End Sub

Private Sub cboTipo_Validate(Cancel As Boolean)
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    'Carrega o próximo número disponível para o tipo global selecionado.
    Cancel = False
    If ConfigSys.TipoFaturamento = "Faturamento Antigo" Then
        txtNumeroNota.Text = CStr(ProximoNumero("Número", "Notas Fiscais de Saída", "[Tipo de Registro] = '" & cboTipo.Text & "'"))
    ElseIf GetFieldValue("nr_nota_fiscal_inicial", "FVFControle_NotaFiscal", "tp_registro = '" & cboTipo.Text & "'") > 0 Then
        txtNumeroNota.Text = CStr(GetFieldValue("nr_nota_fiscal_atual", "FVFControle_NotaFiscal", "tp_registro = '" & cboTipo.Text & "'") + 1)
    Else
         MsgBox "Número sequencial do tipo global " & cboTipo.Text & " não cadastrado!" _
                & Chr(13) & " Cadastre um o número sequencial", vbInformation, "Verifica Número Sequencial"
        Cancel = True
    End If
    '-------------------------------------------------------------------------------------
End Sub

Private Sub cmdCancelar_Click()
    Set rsItens = Nothing
    Unload Me
    Exit Sub
End Sub

Private Sub cmdConfirmar_Click()
    Dim lngNota As Long
    Dim rstNFS As Object
    Dim strCriterio As String
    Dim blnGerouNF As Boolean
    'Pt.97090  - Fernando Paludo - (19/02/2010)
    Dim oOrdemCarregamentoNovo As COrdCarregamento_Novo
    Dim strCriterio1    As String
    '-------------------------------------------------------------------------------------
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    'Função criada para validar os campos obrigatórios
    '-------------------------------------------------------------------------------------
    If ValidaCampos Then
       
        '-----------------------------------------------------------------------------------
        'Pt.97090  - Fernando Paludo - (19/02/2010)
        'Criado o tratamento para verificar qual rotina deve ser chamada na geração da nota.
        '-----------------------------------------------------------------------------------
        
        'Se o sistema esta configurado para utilizar a rotina antiga.
        If ConfigSys.TipoFaturamento = "Faturamento Antigo" Then
        
            lngNota = oOrdemCarregamento.gerarNFS(cboTipo.Text, CLng(strToDbl(txtNumeroNota.Text)), txtOperacaoContabil.valorInteiro, CDate(txtEmissao.Text), CLng(txtTransportadora.Text))
            If lngNota > 0 Then
                strCriterio1 = "[Número] = " & oOrdemCarregamento.NumeroNFS
                strCriterio = "[Número] = " & oOrdemCarregamento.NumeroNFS & " AND [Tipo de Registro] = '" & _
                    oOrdemCarregamento.tipoNFS & "' AND Fornecedor='" & oOrdemCarregamento.fornecedorNFS & "'"
                With wdbItens
                    .RecordRelation = ""
                    .RecordSource = ""
                  .Refresh
                End With
                Call oOrdemCarregamento.LimpaItens
                If MsgBox("Criada a nota fiscal de número '" & lngNota & "'" & vbNewLine & _
                "Deseja visualizar essa Nota Fiscal?", vbInformation + vbYesNo, Me.Caption) = vbYes Then
                    Me.Hide
                    frmOCRLiberadas.Hide
'                    Call frmGBLNFS.ConfigureGlobal(GetItemDataFromStr(LoadResString(1040), "Notas Fiscais de Saída"))
'                    Load frmGBLNFS
                    
                    Call AbreRecordset(rstNFS, "SELECT * FROM [Notas Fiscais de Saída]")
                    'pt. 83956 - Dulcino Júnior (18/10/2007)
                    'rstNFS.FindFirst strCriterio
                    'Pt. 95368 - Moacir Pfau(03/11/2009)
                    rstNFS.Find strCriterio1, 0, adSearchForward
                    
'                    frmGBLNFS.DefineRegistro rstNFS, True, Acesso(LoadResString(2174)), bEdicao:=True
                    
                    Unload frmOCRLiberadas
                    Unload Me
'                    frmGBLNFS.Show
                End If
            Else
                MsgBox "Erro ao gerar a nota fiscal de saída", vbCritical, Me.Caption
            End If
        'Se o sistema esta configurado para utilizar a rotina nova.
        Else
                                   
            'Projeto: #218 - História: #195 - Desenvolvimento#428 - João Henrique(18/09/2012)
            Set oOrdemCarregamentoNovo = New COrdCarregamento_Novo
            
            Call oOrdemCarregamentoNovo.Carregar(oOrdemCarregamento.numeroOCR)
                                   
            'Chama a rotina de geração da pré-nota
            blnGerouNF = oOrdemCarregamentoNovo.GeracaoNF(cboTipo.Text, CLng(strToDbl(txtNumeroNota.Text)), txtOperacaoContabil.valorInteiro, CDate(txtEmissao.Text), CLng(txtTransportadora.Text), txtFornecedor.Text, rsItens)
            
            'Verifica se o retorno da variável e verdadeiro ou falso
            If blnGerouNF Then
                               
                'Verifica se o usuário possui permissão para visualizar a tela de pré-nota.
                If verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_VENDAS, 2819) Then
                    
                    'Exibe a mensagem de geração da pré-nota e solicita se o usuário gostaria de visualizar a mesma.
                    If MsgBox("Criada a pré-nota de número '" & oOrdemCarregamentoNovo.NumeroNFS & "'" & vbNewLine & _
                    "Deseja visualizar essa Pré-Nota?", vbInformation + vbYesNo, Me.Caption) = vbYes Then
                                                                    
                        'Descarrega os forms
                        Unload frmOCRLiberadas
                        Unload Me
                                                                        
                        'Chama o método para exibir o form de pré-nota
                        Call mostrarForm(frmPreNotaFiscalSaida, 2819, False)
                        
                        'Carrega o dados no form de pré-nota
                        Call frmPreNotaFiscalSaida.CarregaNotaSaida(oOrdemCarregamentoNovo.NumeroNFS, oOrdemCarregamentoNovo.tipoNFS)
                    Else
                        'Descarrega os forms
                        Unload Me
                    End If
                Else
                    MsgBox ("Criada a pré-nota de número '" & txtNumeroNota.Text & "'")
                    Unload Me
                End If
            End If
        End If
    End If  'Fim de valida campos
End Sub

Private Sub Form_Load()
    Dim cmd As IDBSelectCommand
    Dim oNota As New CNotaFiscalSaida
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "[Tipos Globais]"
    cmd.OrderByClause = "Tipo"
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    cboTipo.Clear
    
    rdResult.MoveFirst
    While Not rdResult.EOF
        cboTipo.AddItem rdResult.GetString("Tipo")
        rdResult.MoveNext
    Wend
    
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    
    Aplicacao.Disconnect
    
    txtFornecedor.Text = oOrdemCarregamento.fornecedorPedidoVenda
    txtTransportadora.Text = oOrdemCarregamento.Transportadora
    txtEmissao.Text = Date
    cboTipo.Text = oOrdemCarregamento.tipoPedidoVenda
    
    If ConfigSys.TipoFaturamento = "Faturamento Antigo" Then
        txtNumeroNota.Text = CStr(ProximoNumero("Número", "Notas Fiscais de Saída", "[Tipo de Registro] = '" & oOrdemCarregamento.tipoPedidoVenda & "'"))
    Else
        txtNumeroNota.Text = CStr(GetFieldValue("nr_nota_fiscal_atual", "FVFControle_NotaFiscal", "tp_registro = '" & oOrdemCarregamento.tipoPedidoVenda & "'") + 1)
    End If
    
    Set oNota = Nothing
    Me.Height = 4995
    Me.Width = 8985
    CenterForm Me
    'PT. 81189 - Dulcino Júnior
    'Desabilita o campo operação contábil
    lblOperacaoContabil.Caption = ""
    txtOperacaoContabil.valorInteiro = 0
    Label6.Enabled = Configuracao("Utiliza Integração Contábil", False)
    txtOperacaoContabil.Enabled = Configuracao("Utiliza Integração Contábil", False)
    lblOperacaoContabil.Enabled = Configuracao("Utiliza Integração Contábil", False)
        
End Sub

Private Sub txtEmissao_KeyPress(KeyAscii As Integer)
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    SetMascara KeyAscii, txtEmissao.SelStart, MASK_DATA
    '----------------------------------------------------------------------------------
End Sub

Private Sub txtEmissao_Validate(Cancel As Boolean)
    If txtEmissao.Text <> "" Then
        If Not IsDate(txtEmissao) Then
            MsgBox "Insira uma data valida!", vbInformation, Me.Caption
            Cancel = True
        End If
    End If
End Sub

Private Sub txtNumeroNota_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtOperacaoContabil_Change()
    Call GetAssocValue("SELECT descricao FROM OperacaoContabil WHERE cd_operacao=" & txtOperacaoContabil.valorInteiro, lblOperacaoContabil)
End Sub

Private Sub txtOperacaoContabil_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        Call PCampo("Operações Contábeis", "OperacaoContabil", pbCampo, txtOperacaoContabil, "cd_operacao")
    End If
End Sub

Private Sub txtOperacaoContabil_Validate(Cancel As Boolean)
     If txtOperacaoContabil.valorInteiro <> 0 Then
        If GetFieldValue("cd_operacao", "OperacaoContabil", "[cd_operacao] = " & txtOperacaoContabil.valorInteiro) = 0 Then
            MsgBox "Código da operação contábil " & CStr(txtOperacaoContabil.valorInteiro) & " não encontrado", vbInformation, Me.Caption
            lblOperacaoContabil.Caption = ""
            Cancel = True
        End If
    Else
        lblOperacaoContabil.Caption = ""
    End If
End Sub

Private Sub txtTransportadora_Change()
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    If txtTransportadora.Text <> "" Then
        Aplicacao.Connect
        Set cmd = Aplicacao.CreateSelectCommand
        cmd.Table.TableName = "Transportadoras"
        Call cmd.Filter.Append("Código = @pCodigo")
        Call cmd.Parameters.add(cmd.CreateParameter("@pCodigo", txtTransportadora.Text, dbFieldTypeLong))
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
        If Not rdResult.EOF Then
            lblTransportadora.Caption = rdResult.GetString("Apel")
        Else
            lblTransportadora.Caption = ""
        End If
        rdResult.CloseReader
        Set rdResult = Nothing
        Set cmd = Nothing
        Aplicacao.Disconnect
    End If
End Sub

'Pt.97090  - Fernando Paludo - (22/02/2010)
'Função responsável por validar os campos obrigatórios.
Private Function ValidaCampos() As Boolean
        
    If txtTransportadora.Text = "" Then
        MsgBox "O campo " & Chr(34) & "Transportadora" & Chr(34) & " deve ser preenchido.", vbInformation, "Validação de Campos"
        txtTransportadora.SetFocus
        Exit Function
    End If
    If txtOperacaoContabil.Enabled Then
        If txtOperacaoContabil.valorInteiro = 0 Then
            MsgBox "O campo " & Chr(34) & "Operação Contábil" & Chr(34) & " deve ser preenchido.", vbInformation, "Validação de Campos"
            txtOperacaoContabil.SetFocus
            Exit Function
        End If
    End If
    If txtEmissao.Text = "" Then
        MsgBox "O campo " & Chr(34) & "Emissão" & Chr(34) & " deve ser preenchido.", vbInformation, "Validação de Campos"
        txtEmissao.SetFocus
        Exit Function
    End If
    
    ValidaCampos = True
End Function

Private Sub txtTransportadora_KeyDown(KeyCode As Integer, Shift As Integer)
   'Pt.97090  - Fernando Paludo - (22/02/2010)
   If Shift = 0 And KeyCode = vbKeyPageDown Then
        If txtTransportadora.Text = "" Then
            txtTransportadora.Text = 0
        End If
        Call PCampo("Consulta de Transportadoras", "SELECT Código, Razão, Apel FROM Transportadoras", pbCampo, txtTransportadora, "Código")
    End If
    '-------------------------------------------------------------------------------
End Sub

Private Sub txtTransportadora_KeyPress(KeyAscii As Integer)
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    Call validaNumeros(KeyAscii)
    '-------------------------------------------------------------------------------
End Sub

'--------------------------------------------------------------------------------------
'Pt.97090  - Fernando Paludo - (22/02/2010)
'Retorna o próximo número diponível referente ao tipo global selecionado.
'--------------------------------------------------------------------------------------
Private Function ProximoDoc() As Long
    Dim selCmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "MAX(Número) AS Total"
        
        .Table.TableName = "[Pedidos de Venda]"
        
        Call .Filter.Append("[Tipo de Registro] = @pTipoRegistro")
        Call .Parameters.add(.CreateParameter("@pTipoRegistro", cboTipo.Text, dbFieldTypeString))
        
        Call .Filter.Append("Fornecedor = @pFornecedor")
        Call .Parameters.add(.CreateParameter("@pFornecedor", txtFornecedor.Text, dbFieldTypeString))
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    If Not rdResult.EOF Then
        ProximoDoc = rdResult.GetLong("Total") + 1
    Else
        ProximoDoc = 1
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Set selCmd = Nothing
End Function

'Pt.  - Fernando Paludo - (24/02/2010)
'-------------------------------------------------------------------------------------
Private Sub txtTransportadora_Validate(Cancel As Boolean)
     If txtTransportadora.Text <> "" Then
        If GetFieldValue("Código", "Transportadoras", "[Código] = " & val(txtTransportadora.Text)) = 0 Then
            MsgBox "Código da transportadora " & txtTransportadora.Text & " não encontrado", vbInformation, Me.Caption
            lblTransportadora.Caption = ""
            Cancel = True
        End If
    Else
        lblTransportadora.Caption = ""
    End If
End Sub
'-------------------------------------------------------------------------------------

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
