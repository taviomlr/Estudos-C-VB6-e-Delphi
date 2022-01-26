VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmGeracaoArquivoRemessaPagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração do Arquivo de Remessa"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13425
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   13425
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLancamentos 
      Height          =   2895
      Left            =   90
      TabIndex        =   16
      Top             =   930
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   5106
      _Version        =   393216
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Height          =   6045
      Left            =   11940
      TabIndex        =   15
      Top             =   -45
      Width           =   1455
      Begin VB.CommandButton cmdGerar 
         Caption         =   "&Gerar"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin MSComctlLib.ImageList imgGrid 
         Left            =   180
         Top             =   1140
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GeracaoArquivoRemessaPagamento.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "GeracaoArquivoRemessaPagamento.frx":0352
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6045
      Left            =   30
      TabIndex        =   14
      Top             =   -45
      Width           =   11865
      Begin Fox.EBSCombo ecbSegmento 
         Height          =   315
         Left            =   8880
         TabIndex        =   3
         Top             =   570
         Width           =   1500
         _extentx        =   2646
         _extenty        =   556
         origemdados     =   2
         dados           =   ""
         dadosassist     =   ""
         caption         =   "Segmento"
         font            =   "GeracaoArquivoRemessaPagamento.frx":06A4
      End
      Begin VB.Frame fraGeralBanco 
         Height          =   2175
         Left            =   60
         TabIndex        =   20
         Top             =   3810
         Width           =   11745
         Begin VB.Frame fraContas 
            Caption         =   "Contas da Empresa"
            Height          =   1965
            Left            =   60
            TabIndex        =   24
            Top             =   150
            Width           =   5655
            Begin VB.CommandButton cmdNovo 
               Caption         =   "&Novo"
               Enabled         =   0   'False
               Height          =   375
               Left            =   4350
               TabIndex        =   4
               Top             =   240
               Width           =   1215
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFavorecidos 
               Height          =   1665
               Left            =   60
               TabIndex        =   25
               Top             =   240
               Width           =   4245
               _ExtentX        =   7488
               _ExtentY        =   2937
               _Version        =   393216
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.Frame fraInformacoesBancos 
            Caption         =   "Informações Banco"
            Height          =   1965
            Left            =   5760
            TabIndex        =   21
            Top             =   150
            Width           =   4635
            Begin Fox.EBSText etxMovimento 
               Height          =   330
               Left            =   2640
               TabIndex        =   5
               Top             =   210
               Width           =   1905
               _extentx        =   100727
               _extenty        =   582
               font            =   "GeracaoArquivoRemessaPagamento.frx":06D0
               tipo            =   4
               tipotexto       =   0
               possuidescricao =   -1  'True
               campocriterio   =   "cd_tipo_movimento"
               campodescricao  =   "desc_tipo_movimento"
               tabelaconsulta  =   "FFICamaraTipoMovimento"
               tamanhodescricao=   1400
            End
            Begin Fox.EBSText etxCodigoMovimento 
               Height          =   330
               Left            =   2640
               TabIndex        =   6
               Top             =   540
               Width           =   1905
               _extentx        =   100727
               _extenty        =   582
               font            =   "GeracaoArquivoRemessaPagamento.frx":06FC
               tipo            =   4
               tipotexto       =   0
               possuidescricao =   -1  'True
               campocriterio   =   "cd_movimento"
               campodescricao  =   "desc_cd_movimento"
               tabelaconsulta  =   "FFICamaraCodigoMovimento"
               tamanhodescricao=   1400
            End
            Begin Fox.EBSText etxTipoCompTED 
               Height          =   330
               Left            =   2640
               TabIndex        =   7
               Top             =   870
               Width           =   1905
               _extentx        =   100727
               _extenty        =   582
               font            =   "GeracaoArquivoRemessaPagamento.frx":0728
               tipo            =   4
               tipotexto       =   0
               possuidescricao =   -1  'True
               campocriterio   =   "cd_tipo_doc"
               campodescricao  =   "desc_tipo_doc"
               tabelaconsulta  =   "FFICamaraTipoDoc"
               tamanhodescricao=   1400
            End
            Begin Fox.EBSText etxFinalidadeTED 
               Height          =   330
               Left            =   2640
               TabIndex        =   8
               Top             =   1200
               Width           =   1920
               _extentx        =   100753
               _extenty        =   582
               font            =   "GeracaoArquivoRemessaPagamento.frx":0754
               tipo            =   4
               tipotexto       =   0
               possuidescricao =   -1  'True
               campocriterio   =   "cd_finalidade_doc"
               campodescricao  =   "desc_finalidade_doc"
               tabelaconsulta  =   "FFICamaraFinalidadeDoc"
               tamanhodescricao=   1400
            End
            Begin Fox.EBSText etxTipoConta 
               Height          =   330
               Left            =   2640
               TabIndex        =   9
               Top             =   1530
               Width           =   1920
               _extentx        =   103188
               _extenty        =   582
               font            =   "GeracaoArquivoRemessaPagamento.frx":0780
               tipo            =   4
               tipotexto       =   0
               possuidescricao =   -1  'True
               campocriterio   =   "cd_tipo_conta"
               campodescricao  =   "desc_tipo_conta"
               tabelaconsulta  =   "FFICamaraTipoConta"
               tamanhodescricao=   1400
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo de Conta (DOC/TED)"
               Height          =   195
               Left            =   480
               TabIndex        =   30
               Top             =   1590
               Width           =   2085
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Finalidade (DOC/TED)"
               Height          =   195
               Left            =   480
               TabIndex        =   29
               Top             =   1260
               Width           =   2085
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo de Compensação (DOC/TED)"
               Height          =   195
               Left            =   60
               TabIndex        =   28
               Top             =   930
               Width           =   2505
            End
            Begin VB.Label lblDescCamaraCentral 
               Height          =   225
               Left            =   2400
               TabIndex        =   27
               Top             =   1185
               Width           =   825
            End
            Begin VB.Label lblMovimento 
               Alignment       =   1  'Right Justify
               Caption         =   "Tipo Movimento"
               Height          =   195
               Left            =   930
               TabIndex        =   23
               Top             =   270
               Width           =   1635
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Código do Movimento"
               Height          =   195
               Left            =   930
               TabIndex        =   22
               Top             =   600
               Width           =   1635
            End
         End
         Begin VB.CommandButton cmdConfirmar 
            Caption         =   "&Confirmar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   10440
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "C&ancelar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   10440
            TabIndex        =   11
            Top             =   630
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdDirGeracao 
         Height          =   330
         Left            =   8400
         Picture         =   "GeracaoArquivoRemessaPagamento.frx":07AC
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   570
         Width           =   315
      End
      Begin VB.TextBox txtDiretorio 
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Top             =   570
         Width           =   6585
      End
      Begin Fox.EBSText etxNumeroLote 
         Height          =   330
         Left            =   1800
         TabIndex        =   18
         Top             =   210
         Width           =   2115
         _extentx        =   265
         _extenty        =   582
         font            =   "GeracaoArquivoRemessaPagamento.frx":08F6
         tipotexto       =   0
         enabled         =   0   'False
         tipocriterio    =   0
         alinhamento     =   1
      End
      Begin Fox.EBSText etxBancoDestinoArquivo 
         Height          =   330
         Left            =   5220
         TabIndex        =   0
         Top             =   210
         Width           =   6390
         _extentx        =   442992
         _extenty        =   582
         font            =   "GeracaoArquivoRemessaPagamento.frx":0922
         tipotexto       =   0
         maxlength       =   9
         possuidescricao =   -1  'True
         campocriterio   =   "Banco"
         tipocriterio    =   4
         campodescricao  =   "Nome"
         tabelaconsulta  =   "Bancos"
         tamanhodescricao=   5400
         alinhamento     =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco Destino"
         Height          =   195
         Left            =   4080
         TabIndex        =   26
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Diretório para Geração"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número do Lote"
         Enabled         =   0   'False
         Height          =   195
         Left            =   600
         TabIndex        =   17
         Top             =   270
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmGeracaoArquivoRemessaPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mblnConfirmou     As Boolean
Private mlngCodigoArquivo As Long
Private mlngCamaraOrigem  As Long
Private mintQtdLotes      As Integer
Private Const colUnCheck = 1
Private Const colCheck = 2
Private mblnExecuteClick As Boolean

Public Property Let CamaraOrigem(ByVal lngCamaraOrigem As Long)
    mlngCamaraOrigem = lngCamaraOrigem
End Property

Public Property Let QtdLotes(ByVal NewVal As Integer)
    mintQtdLotes = NewVal
End Property

'Data.......: 26/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para configurar os campos da grid para exibição dos registros a
'             serem integrados.
Private Sub ConfigureGrid()
    Dim intColuna As Integer

    With grdLancamentos
        .Rows = 2
        .FixedRows = 1
        .Cols = 21
        .FixedCols = 1
        
        'Coluna Fixa
        .ColWidth(0) = 150
        .TextMatrix(0, 0) = ""
        .ColAlignment(0) = flexAlignLeftCenter
        
        'Coluna do seqüêncial
        .ColWidth(1) = 1
        .TextMatrix(0, 1) = "Sequencial"
        .ColAlignment(1) = flexAlignLeftCenter
        
        'Coluna Documento
        .ColWidth(2) = 600
        .TextMatrix(0, 2) = "Doc."
        .ColAlignment(2) = flexAlignLeftCenter
        
        'Coluna Tipo de Registro
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = "Tipo"
        .ColAlignment(3) = flexAlignLeftCenter
        
        'Coluna Número
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Número"
        .ColAlignment(4) = flexAlignLeftCenter
        
        'Coluna Parcela
        .ColWidth(5) = 600
        .TextMatrix(0, 5) = "Parc."
        .ColAlignment(5) = flexAlignLeftCenter
        
        'Coluna Empresa
        .ColWidth(6) = 1800
        .TextMatrix(0, 6) = "Empresa"
        .ColAlignment(6) = flexAlignLeftCenter
        
        
        'Coluna Número Banco
        .ColWidth(7) = 1
        .TextMatrix(0, 7) = "Código do Banco"
        .ColAlignment(7) = flexAlignLeftCenter
        
        'Coluna Câmara
        .ColWidth(8) = 700
        .TextMatrix(0, 8) = "Câmara"
        .ColAlignment(8) = flexAlignLeftCenter
        
        'Coluna Agência
        .ColWidth(9) = 1000
        .TextMatrix(0, 9) = "Agência"
        .ColAlignment(9) = flexAlignLeftCenter
        
        'Coluna Digito da Agência
        .ColWidth(10) = 600
        .TextMatrix(0, 10) = "DV"
        .ColAlignment(10) = flexAlignLeftCenter
        
        'Coluna Conta Corrente
        .ColWidth(11) = 1500
        .TextMatrix(0, 11) = "Conta Corrente"
        .ColAlignment(11) = flexAlignLeftCenter
        
        'Coluna Digito da Conta Corrente
        .ColWidth(12) = 600
        .TextMatrix(0, 12) = "DV"
        .ColAlignment(12) = flexAlignLeftCenter
        
        'Coluna Tipo de Serviço
        .ColWidth(13) = 1000
        .TextMatrix(0, 13) = "Tipo Serviço"
        .ColAlignment(13) = flexAlignLeftCenter
        
        'Coluna Forma de Lançamento
        .ColWidth(14) = 1450
        .TextMatrix(0, 14) = "Forma Lançamento"
        .ColAlignment(14) = flexAlignLeftCenter
        
        'Coluna Tipo do Movimento
        .ColWidth(15) = 1000
        .TextMatrix(0, 15) = "Tipo Movimento"
        .ColAlignment(15) = flexAlignLeftCenter
        
        'Coluna Código do Movimento
        .ColWidth(16) = 1450
        .TextMatrix(0, 16) = "Código Movimento"
        .ColAlignment(16) = flexAlignLeftCenter
        
        'Coluna Camera Centralizadora
        .ColWidth(17) = 1450
        .TextMatrix(0, 17) = "Cam. Central"
        .ColAlignment(17) = flexAlignLeftCenter
        
        .ColWidth(18) = 1450
        .TextMatrix(0, 18) = "Tipo Compensação DOC/TED"
        .ColAlignment(18) = flexAlignLeftCenter
        
        .ColWidth(19) = 1450
        .TextMatrix(0, 19) = "Finalidade DOC/TED"
        .ColAlignment(19) = flexAlignLeftCenter
        
        .ColWidth(20) = 1450
        .TextMatrix(0, 20) = "Tipo de Conta DOC/TED"
        .ColAlignment(20) = flexAlignLeftCenter
        
        For intColuna = 0 To .Cols - 1
            .TextMatrix(1, intColuna) = ""
        Next
    End With
End Sub

'Data.......: 29/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para listar todos os dados dos favorecidos.
Private Sub ConfigureGridFavorecidos()
    Dim intColuna As Integer
    
    With grdFavorecidos
        .Rows = 2
        .Cols = 7
        
        'Coluna Fixa
        .ColWidth(0) = 150
        .TextMatrix(0, 0) = ""
        
        'Coluna de seleção
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        .ColAlignment(1) = flexAlignCenterCenter
        
        'Coluna Câmara
        .ColWidth(2) = 650
        .TextMatrix(0, 2) = "Câmara"
        .ColAlignment(2) = flexAlignLeftCenter
        
        'Coluna Agência
        .ColWidth(3) = 900
        .TextMatrix(0, 3) = "Agência"
        .ColAlignment(3) = flexAlignLeftCenter
        
        'Coluna Digito da Agência
        .ColWidth(4) = 350
        .TextMatrix(0, 4) = "DV"
        .ColAlignment(4) = flexAlignLeftCenter
        
        'Coluna Conta Corrente
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = "Conta Corrente"
        .ColAlignment(5) = flexAlignLeftCenter
        
        'Coluna Digito da Conta Corrente
        .ColWidth(6) = 350
        .TextMatrix(0, 6) = "DV"
        .ColAlignment(6) = flexAlignLeftCenter
        
        For intColuna = 0 To .Cols - 1
            .TextMatrix(1, intColuna) = ""
            If intColuna = 1 Then
                .col = intColuna
                Set .CellPicture = imgGrid.ListImages(colUnCheck).Picture
            End If
        Next
    End With
End Sub

Private Sub cmdCancelar_Click()
    Call PreencheGridFavorecidos
    cmdConfirmar.Enabled = False
    cmdCancelar.Enabled = False
End Sub

Private Sub cmdConfirmar_Click()
    Dim intLinha       As Integer
    Dim blnSelecionado As Boolean
    Dim strSql         As String

    With grdFavorecidos
        .col = 1
        blnSelecionado = False
        For intLinha = 1 To .Rows - 1
            .Row = intLinha
            If .CellPicture = imgGrid.ListImages(colCheck).Picture Or .TextMatrix(.Row, 2) = Empty Then
                blnSelecionado = True
                Exit For
            End If
        Next
        If blnSelecionado And ValidaInformacoesBanco Then
            'Atualiza o campo da câmara
            grdLancamentos.TextMatrix(grdLancamentos.Row, 8) = .TextMatrix(.Row, 2)
            'Atualiza o campo Agência
            grdLancamentos.TextMatrix(grdLancamentos.Row, 9) = .TextMatrix(.Row, 3)
            'Atualiza o campo Dígito verificador da agência
            grdLancamentos.TextMatrix(grdLancamentos.Row, 10) = .TextMatrix(.Row, 4)
            'Atualiza o campo Conta Corrente
            grdLancamentos.TextMatrix(grdLancamentos.Row, 11) = .TextMatrix(.Row, 5)
            'Atualiza o campo Dígito verificador da Conta Corrente
            grdLancamentos.TextMatrix(grdLancamentos.Row, 12) = .TextMatrix(.Row, 6)
            
            'Informações do Banco para geração do arquivo
            grdLancamentos.TextMatrix(grdLancamentos.Row, 15) = etxMovimento.valorTexto
            grdLancamentos.TextMatrix(grdLancamentos.Row, 16) = etxCodigoMovimento.valorTexto
            'Ivo Sousa (23/04/2017) - Integração Bradesco
            grdLancamentos.TextMatrix(grdLancamentos.Row, 18) = etxTipoCompTED.valorTexto
            grdLancamentos.TextMatrix(grdLancamentos.Row, 19) = etxFinalidadeTED.valorTexto
            grdLancamentos.TextMatrix(grdLancamentos.Row, 20) = etxTipoConta.valorTexto
            
            strSql = "UPDATE FFIItemPagamento SET"
            strSql = strSql & " nr_agencia = " & strToLng(grdLancamentos.TextMatrix(grdLancamentos.Row, 9))
            strSql = strSql & ", nr_digitoAgencia = '" & grdLancamentos.TextMatrix(grdLancamentos.Row, 10) & "'"
            strSql = strSql & ", nr_contaCorrente = " & strToLng(grdLancamentos.TextMatrix(grdLancamentos.Row, 11))
            strSql = strSql & ", nr_digitoConta = '" & grdLancamentos.TextMatrix(grdLancamentos.Row, 12) & "'"
            strSql = strSql & ", nr_camara = " & strToLng(grdLancamentos.TextMatrix(grdLancamentos.Row, 8))
            strSql = strSql & ", cd_tipo_movimento = " & Quote(grdLancamentos.TextMatrix(grdLancamentos.Row, 15), "'")
            strSql = strSql & ", cd_movimento = " & Quote(grdLancamentos.TextMatrix(grdLancamentos.Row, 16), "'")
            strSql = strSql & ", cd_camara_centralizadora = " & grdLancamentos.TextMatrix(grdLancamentos.Row, 17)
            'Ivo Sousa (23/04/2017) - Integração Bradesco
            If grdLancamentos.TextMatrix(grdLancamentos.Row, 18) <> Empty Then
                strSql = strSql & ", cd_tipo_doc = '" & grdLancamentos.TextMatrix(grdLancamentos.Row, 18) & "'"
            End If
            If grdLancamentos.TextMatrix(grdLancamentos.Row, 19) <> Empty Then
                strSql = strSql & ", cd_finalidade_doc = " & grdLancamentos.TextMatrix(grdLancamentos.Row, 19)
            End If
            If grdLancamentos.TextMatrix(grdLancamentos.Row, 20) <> Empty Then
                strSql = strSql & ", cd_tipo_conta = " & grdLancamentos.TextMatrix(grdLancamentos.Row, 20)
            End If
            
            strSql = strSql & " WHERE cd_arquivoPagamento = " & mlngCodigoArquivo & " AND cd_itemPagamento = " & StrToInt(grdLancamentos.TextMatrix(grdLancamentos.Row, 1))
            Call ExecuteSQL(strSql)
        ElseIf Not blnSelecionado Then
            MsgBox "Selecione os dados para alteração do registro.", vbInformation
        End If
    End With
End Sub

Private Sub cmdDirGeracao_Click()
    Dim strDiretorio As String
    
    strDiretorio = FolderDialogBox(Me.hWnd, "Diretório que será criado o arquivo de Remessa", rfcDesktop, bfReturnDirs)
    txtDiretorio.Text = strDiretorio
    txtDiretorio.SetFocus
End Sub

Private Sub cmdGerar_Click()
    Dim objArquivo As New cArquivoTexto
    Dim objPagFor As New clsPagFor
    Dim strErro As String
    Dim intAux As Integer
    
On Error GoTo ErroGeracao
    If ValidaGeracao Then
        BeginTrans
            
        Call WriteSettings("CaminhoArquivo", "RemessaBancaria", txtDiretorio.Text)
        
        'Gera o Cabeçalho do Arquivo
        'Ivo Sousa (17/04/2017) - Tratamento especifico seguindo o padrão estabelecido pelo banco Bradesco
        If mlngCamaraOrigem = 237 Then
            intAux = mlngCodigoArquivo
            Call objArquivo.novo("PG" & Format(Now, "ddMM") & formatCampoInt(CStr(intAux), 2) & ".REM", txtDiretorio.Text)
        Else
            Call objArquivo.novo("PAGFOR.TXT", txtDiretorio.Text)
        End If
        objPagFor.objArquivo = objArquivo
        objPagFor.CamaraOrigem = mlngCamaraOrigem
        objPagFor.NumeroArquivo = mlngCodigoArquivo
            
        objPagFor.QtdLotes = mintQtdLotes
        objPagFor.BancoDestino = etxBancoDestinoArquivo.valorInteiro
        objPagFor.SegmentoOrigem = ecbSegmento.SelectedItem 'Demanda 222036 - Autor Yuji - Mapeamento do segmento J para o Itaú
        If objPagFor.geraArquivo(strErro) Then
            If objArquivo.salvar Then
                mblnConfirmou = True
                Call ExecuteSQL("UPDATE Bancos SET nr_Remessa_PagFor = " & mlngCodigoArquivo & " WHERE Banco = " & etxBancoDestinoArquivo.valorInteiro)
                CommitTrans
                MsgBox "O arquivo foi gerado com sucesso.", vbInformation, NomeModulo
                Unload Me
            End If
        Else
            MsgBox strErro & ".", vbInformation, NomeModulo
            Rollback
        End If
    End If
    Exit Sub
ErroGeracao:
    Rollback
End Sub

Private Sub cmdNovo_Click()
    Call mostrarForm(frmDadosFavorecido, 2845, False)
    Call frmDadosFavorecido.InserirDados(grdLancamentos.TextMatrix(grdLancamentos.Row, 6), Me)
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub etxAgencia_KeyPress(KeyAscii As Integer)
    If Not TeclaEspecial(KeyAscii) Then
        If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub etxContaCorrente_KeyPress(KeyAscii As Integer)
    If Not TeclaEspecial(KeyAscii) Then
        If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub etxBancoDestinoArquivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxBancoDestinoArquivo.ValorDescricao = "" Then
            etxBancoDestinoArquivo.valorInteiro = 0
        End If
        Call PCampo("Bancos", "SELECT Banco,Nome FROM Bancos", pbCampo, etxBancoDestinoArquivo, "Banco")
    End If
End Sub

Private Sub etxCodigoMovimento_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intLinhaSelecionada As Integer
    
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxCodigoMovimento.ValorDescricao = "" Then
            etxCodigoMovimento.valorTexto = ""
        End If
        intLinhaSelecionada = LinhaSelecionada
        mblnExecuteClick = False
        Call PCampo("Código do Movimento", "SELECT * FROM FFICamaraCodigoMovimento WHERE cd_camara = " & mlngCamaraOrigem, pbCampo, etxCodigoMovimento, "cd_movimento")
        DoEvents
    End If
End Sub

Private Sub etxFinalidadeTED_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intLinhaSelecionada As Integer
    
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxFinalidadeTED.ValorDescricao = "" Then
            etxFinalidadeTED.valorTexto = ""
            DoEvents
        End If
        intLinhaSelecionada = LinhaSelecionada
        mblnExecuteClick = False
        Call PCampo("Finalidade", "SELECT * FROM FFICamaraFinalidadeDoc WHERE cd_camara = " & mlngCamaraOrigem, pbCampo, etxFinalidadeTED, "cd_finalidade_doc")
        DoEvents
    End If
End Sub

Private Sub etxMovimento_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intLinhaSelecionada As Integer
    
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxMovimento.ValorDescricao = "" Then
            etxMovimento.valorTexto = ""
            DoEvents
        End If
        intLinhaSelecionada = LinhaSelecionada
        mblnExecuteClick = False
        Call PCampo("Tipo do Movimento", "SELECT * FROM FFICamaraTipoMovimento WHERE cd_camara = " & mlngCamaraOrigem, pbCampo, etxMovimento, "cd_tipo_movimento")
        DoEvents
    End If
End Sub

Private Sub etxTipoCompTED_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intLinhaSelecionada As Integer
    
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxTipoCompTED.ValorDescricao = "" Then
            etxTipoCompTED.valorTexto = ""
            DoEvents
        End If
        intLinhaSelecionada = LinhaSelecionada
        mblnExecuteClick = False
        Call PCampo("Tipo do Documento", "SELECT * FROM FFICamaraTipoDoc WHERE cd_camara = " & mlngCamaraOrigem, pbCampo, etxTipoCompTED, "cd_tipo_doc")
        DoEvents
    End If
End Sub

Private Sub etxTipoConta_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intLinhaSelecionada As Integer
    
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxTipoConta.ValorDescricao = "" Then
            etxTipoConta.valorTexto = ""
            DoEvents
        End If
        intLinhaSelecionada = LinhaSelecionada
        mblnExecuteClick = False
        Call PCampo("Tipo do Conta", "SELECT * FROM FFICamaraTipoConta WHERE cd_camara = " & mlngCamaraOrigem, pbCampo, etxTipoConta, "cd_tipo_conta")
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    Call ConfigureGrid
    Call ConfigureGridFavorecidos
    Call etxCodigoMovimento.AddConexao(Aplicacao)
    Call etxMovimento.AddConexao(Aplicacao)
    Call etxBancoDestinoArquivo.AddConexao(Aplicacao)
    Call etxTipoCompTED.AddConexao(Aplicacao)
    Call etxFinalidadeTED.AddConexao(Aplicacao)
    Call etxTipoConta.AddConexao(Aplicacao)
    mblnConfirmou = False
    mblnExecuteClick = True
    txtDiretorio.Text = ReadSettings("CaminhoArquivo", "RemessaBancaria", "")
    'Demanda 222036 - Autor Yuji - Mapeamento do segmento J para o Itaú
    ecbSegmento.AddItem "A"
    ecbSegmento.SelectItem "A"
End Sub

'Data.......: 26/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para exibir os registros de um determinado Lote de pagamento digital
'               para o fornecedor.
'Parametros.: [Long] Número do lote de pagamento digital gerado na tela anterior.
Public Sub MostraRegistro(lngNumero As Long)
    Dim strSql     As String
    Dim rstResult  As Object
    Dim strResult  As String
    
    strSql = "SELECT * FROM FFIItemPagamento WHERE cd_arquivoPagamento = " & lngNumero & " ORDER BY cd_arquivoPagamento"
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        etxNumeroLote.valorInteiro = lngNumero
        With rstResult
            While Not .EOF
                strResult = "" & Chr(vbKeyTab) & .Fields("cd_itemPagamento").value & Chr(vbKeyTab) & .Fields("tp_documento").value & _
                Chr(vbKeyTab) & .Fields("tipo_registro").value & Chr(vbKeyTab) & .Fields("nr_documento").value & Chr(vbKeyTab) & _
                .Fields("nr_parcela").value & Chr(vbKeyTab) & .Fields("cd_empresa").value & Chr(vbKeyTab) & _
                .Fields("cd_banco")
                Call CompletaCampos(.Fields("cd_empresa").value, strResult, .Fields("cd_tipo_servico").value, .Fields("cd_forma_lancamento").value, .Fields("cd_camara_centralizadora").value)
                'strResult = strResult & Chr(vbKeyTab) & Chr(vbKeyTab) & Chr(vbKeyTab) & .Fields("cd_tipo_conta").value
                Call grdLancamentos.AddItem(strResult)
                .MoveNext
                mlngCodigoArquivo = lngNumero
            Wend
            If grdLancamentos.TextMatrix(1, 1) = "" Then
                Call grdLancamentos.RemoveItem(1)
            End If
        End With
        If grdLancamentos.TextMatrix(1, 1) <> "" Then
            grdLancamentos.Row = 1
            Call grdLancamentos_Click
        End If
    End If
    Call FechaRecordset(rstResult)

    'Demanda 222036 - Autor Yuji - Mapeamento do segmento J para o Itaú
    Select Case mlngCamaraOrigem
        Case 237
            ecbSegmento.Visible = False
        Case 341
            ecbSegmento.AddItem "J"
    End Select
End Sub

'Data.......: 26/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para completar a string que deve ser utilizada como linha para a
'               exibição da grid.
'Parametros.: [String] Empresa a quem o documento pertence.
'             [String] Linha com as informações do registro que devem ser exibidas no grid.
Private Sub CompletaCampos(strEmpresa As String, ByRef strRegistro As String, strTipoServico As String, strFormaLancamento As String, strCamCentralizadora As String)
    Dim strSql    As String
    Dim rstResult As Object

    strSql = "SELECT nr_banco, nr_agencia, dv_agencia, nr_conta_corrente, dv_conta_corrente FROM FFIDadosFavorecidos" & _
             " WHERE cd_empresa='" & strEmpresa & "'"
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        With rstResult
            .MoveLast
            If .Recordcount = 1 Then
                strRegistro = strRegistro & Chr(vbKeyTab) & .Fields("nr_banco").value & Chr(vbKeyTab) & .Fields("nr_agencia").value & _
                        Chr(vbKeyTab) & .Fields("dv_agencia").value & Chr(vbKeyTab) & .Fields("nr_conta_corrente").value & _
                        Chr(vbKeyTab) & .Fields("dv_conta_corrente").value & Chr(vbKeyTab) & Format(strTipoServico, "00") & Chr(vbKeyTab) & _
                        Format(strFormaLancamento, "00") & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & Format(strCamCentralizadora, "000")
            Else
                .MoveFirst
                strRegistro = "+" & strRegistro & Chr(vbKeyTab) & .Fields("nr_banco").value & Chr(vbKeyTab) & .Fields("nr_agencia").value & _
                        Chr(vbKeyTab) & .Fields("dv_agencia").value & Chr(vbKeyTab) & .Fields("nr_conta_corrente").value & _
                        Chr(vbKeyTab) & .Fields("dv_conta_corrente").value & Chr(vbKeyTab) & Format(strTipoServico, "00") & Chr(vbKeyTab) & _
                        Format(strFormaLancamento, "00") & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & Format(strCamCentralizadora, "000")
            End If
        End With
    Else
        strRegistro = strRegistro & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & Format(strTipoServico, "00") & Chr(vbKeyTab) & Format(strFormaLancamento, "00") & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & Format(strCamCentralizadora, "000")
    End If
    Call FechaRecordset(rstResult)
End Sub

'Data.......: 29/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento iniciado através do click da grid superior para preencher todos os registros de
'               contas e bancos cadastrados para o favorecido.
Private Sub PreencheGridFavorecidos()
    Dim strSql    As String
    Dim rstResult As Object
    Dim strLinha  As String
    
    With grdLancamentos
        Call ConfigureGridFavorecidos
        Call LimpaCamposInfBancos
        fraContas.Caption = "Contas da Empresa"
        If .TextMatrix(.Row, 1) <> "" Then
            cmdNovo.Enabled = True
            cmdConfirmar.Enabled = True
            cmdCancelar.Enabled = True
            fraContas.Caption = fraContas.Caption & " " & .TextMatrix(.Row, 6)
            strSql = "SELECT nr_banco, nr_agencia, dv_agencia, nr_conta_corrente, dv_conta_corrente " & _
                    " FROM FFIDadosFavorecidos WHERE cd_empresa='" & .TextMatrix(.Row, 6) & "'"
            If AbreRecordset(rstResult, strSql) = WL_OK Then
                While Not rstResult.EOF
                    strLinha = "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & rstResult.Fields("nr_banco").value & Chr(vbKeyTab) & _
                                rstResult.Fields("nr_agencia").value & Chr(vbKeyTab) & rstResult.Fields("dv_agencia").value & _
                                Chr(vbKeyTab) & rstResult.Fields("nr_conta_corrente").value & Chr(vbKeyTab) & _
                                rstResult.Fields("dv_conta_corrente").value
                    Call grdFavorecidos.AddItem(strLinha)
                    grdFavorecidos.Row = grdFavorecidos.Rows - 1
                    grdFavorecidos.col = 1
                    If .TextMatrix(.Row, 8) = rstResult.Fields("nr_banco").value And .TextMatrix(.Row, 9) = rstResult.Fields("nr_agencia").value _
                        And .TextMatrix(.Row, 10) = rstResult.Fields("dv_agencia").value And .TextMatrix(.Row, 11) = rstResult.Fields("nr_conta_corrente").value _
                        And .TextMatrix(.Row, 12) = rstResult.Fields("dv_conta_corrente").value Then
                        Set grdFavorecidos.CellPicture = imgGrid.ListImages(colCheck).Picture
                    Else
                        Set grdFavorecidos.CellPicture = imgGrid.ListImages(colUnCheck).Picture
                    End If
                    rstResult.MoveNext
                Wend
                If grdFavorecidos.Rows > 2 And grdFavorecidos.TextMatrix(1, 1) = "" Then
                    Call grdFavorecidos.RemoveItem(1)
                End If
            End If
        Else
            cmdNovo.Enabled = False
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intRow As Integer
    
    If Not mblnConfirmou Then
        Call ExecuteSQL("DELETE FROM FFIItemPagamento WHERE cd_arquivoPagamento = " & mlngCodigoArquivo)
        Call ExecuteSQL("DELETE FROM FFIArquivoPagamento WHERE cd_arquivoPagamento = " & mlngCodigoArquivo)
    End If
    frmPagamentoDigitalFornecedores.cmdVisualizar_Click
    frmPagamentoDigitalFornecedores.Visible = True
End Sub

Private Sub grdFavorecidos_Click()
    Dim lngLinhaSelecionada As Long
    Dim lngLinhas           As Long
    
    With grdFavorecidos
        If .col = 1 Then
            If .TextMatrix(.Row, 2) <> "" Then
                If .CellPicture = imgGrid.ListImages(colCheck).Picture Then
                    cmdConfirmar.Enabled = False
                    cmdCancelar.Enabled = False
                    Set .CellPicture = imgGrid.ListImages(colUnCheck).Picture
                Else
                    lngLinhaSelecionada = .Row
                    For lngLinhas = 1 To .Rows - 1
                        .Row = lngLinhas
                        Set .CellPicture = imgGrid.ListImages(colUnCheck).Picture
                    Next
                    .Row = lngLinhaSelecionada
                    Set .CellPicture = imgGrid.ListImages(colCheck).Picture
                    lngLinhaSelecionada = 0
                    cmdConfirmar.Enabled = True
                    cmdCancelar.Enabled = True
                End If
            End If
        End If
    End With
End Sub

Private Sub grdLancamentos_Click()
    If mblnExecuteClick Then
        Call PreencheGridFavorecidos
        Call PreencheCamposFavorecido
    Else
        mblnExecuteClick = True
    End If
End Sub

'Data.......: 29/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para atualizar a grid de dados do favorecido conforme chamada
'               da tela de cadastro de favorecidos.
Public Sub AtualizaLista()
    Call grdLancamentos_Click
End Sub

Private Function LinhaSelecionada() As Integer
    Dim intCont As Integer
    
    For intCont = 1 To grdFavorecidos.Rows - 1
        grdFavorecidos.Row = intCont
        If grdFavorecidos.CellPicture = imgGrid.ListImages(colCheck).Picture Then
            LinhaSelecionada = intCont
            Exit Function
        End If
    Next intCont
End Function

Private Sub LimpaCamposInfBancos()
    etxMovimento.Clear
    etxCodigoMovimento.Clear
End Sub

Private Function ValidaGeracao() As Boolean
    Dim lngLinha As Long
    
    'With grdLancamentos
    '    For lngLinha = 1 To .Rows - 1
    '        If Not .TextMatrix(lngLinha, 8) <> "" Or Not .TextMatrix(lngLinha, 13) <> "" Or Not .TextMatrix(lngLinha, 15) <> "" Then
    '            MsgBox "O registro " & .TextMatrix(lngLinha, 4) & " parcela " & .TextMatrix(lngLinha, 5) & " está sem as informções referentes à Câmara.", vbInformation, NomeModulo
    '            ValidaGeracao = False
    '            Exit Function
    '        End If
    '    Next
    'End With
    If Not etxBancoDestinoArquivo.valorInteiro > 0 Then
        MsgBox "Informe o Banco de destino do arquivo.", vbInformation, NomeModulo
        etxBancoDestinoArquivo.SetFocus
        ValidaGeracao = False
        Exit Function
    End If
    If Not Len(Trim(txtDiretorio.Text)) > 0 Then
        MsgBox "Informe o Diretório de destino do arquivo.", vbInformation, NomeModulo
        txtDiretorio.SetFocus
        ValidaGeracao = False
        Exit Function
    End If
    
    ValidaGeracao = True
End Function

Private Sub PreencheCamposFavorecido()
    With grdLancamentos
        If Trim(.TextMatrix(.Row, 2)) <> "" Then
            If Trim(.TextMatrix(.Row, 15)) <> "" Then
                etxMovimento.valorTexto = .TextMatrix(.Row, 15)
            End If
            If Trim(.TextMatrix(.Row, 16)) <> "" Then
                etxCodigoMovimento.valorTexto = .TextMatrix(.Row, 16)
            End If
        End If
    End With
End Sub

Private Function ValidaInformacoesBanco() As Boolean
    If Not Trim(etxMovimento.valorTexto) <> "" Then
        MsgBox "O Tipo de Movimento do documento não foi preenchido.", vbInformation, NomeModulo
        etxMovimento.SetFocus
        Exit Function
    End If
    If Not Trim(etxCodigoMovimento.valorTexto) <> "" Then
        MsgBox "O Código do Movimento do documento não foi preenchido.", vbInformation, NomeModulo
        etxCodigoMovimento.SetFocus
        Exit Function
    End If
    ValidaInformacoesBanco = True
End Function

Private Sub grdLancamentos_EnterCell()
    Call PreencheGridFavorecidos
    Call PreencheCamposFavorecido
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
