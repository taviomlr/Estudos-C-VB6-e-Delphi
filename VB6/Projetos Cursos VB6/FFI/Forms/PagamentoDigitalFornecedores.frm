VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSComctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmPagamentoDigitalFornecedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagamento Digital de Fornecedores "
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   13095
   Begin VB.Frame fraBotoes 
      Height          =   7515
      Left            =   11610
      TabIndex        =   14
      Top             =   -40
      Width           =   1455
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton cmdVisualizar 
         Caption         =   "&Visualizar"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   1215
      End
      Begin MSComctlLib.ImageList imgCheck 
         Left            =   420
         Top             =   1980
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
               Picture         =   "PagamentoDigitalFornecedores.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PagamentoDigitalFornecedores.frx":0352
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PagamentoDigitalFornecedores.frx":06A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PagamentoDigitalFornecedores.frx":094F
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraControles 
      Height          =   7515
      Left            =   30
      TabIndex        =   13
      Top             =   -40
      Width           =   11535
      Begin VB.PictureBox imgCol8 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   10830
         Picture         =   "PagamentoDigitalFornecedores.frx":0BFA
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   43
         Top             =   2670
         Width           =   270
      End
      Begin VB.PictureBox imgCol7 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   9360
         Picture         =   "PagamentoDigitalFornecedores.frx":0E95
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   42
         Top             =   2670
         Width           =   270
      End
      Begin VB.PictureBox imgCol6 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   8130
         Picture         =   "PagamentoDigitalFornecedores.frx":1130
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   41
         Top             =   2670
         Width           =   270
      End
      Begin VB.PictureBox imgCol5 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   7260
         Picture         =   "PagamentoDigitalFornecedores.frx":13CB
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   40
         Top             =   2670
         Width           =   270
      End
      Begin VB.PictureBox imgCol4 
         BorderStyle     =   0  'None
         Height          =   270
         Index           =   0
         Left            =   4260
         Picture         =   "PagamentoDigitalFornecedores.frx":1666
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   39
         Top             =   2670
         Width           =   270
      End
      Begin VB.PictureBox imgCol3 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   3510
         Picture         =   "PagamentoDigitalFornecedores.frx":1901
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   38
         Top             =   2670
         Width           =   270
      End
      Begin VB.PictureBox imgCol2 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   2500
         Picture         =   "PagamentoDigitalFornecedores.frx":1B9C
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   37
         Top             =   2670
         Width           =   270
      End
      Begin VB.PictureBox imgCol1 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   900
         Picture         =   "PagamentoDigitalFornecedores.frx":1E37
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   36
         Top             =   2670
         Width           =   270
      End
      Begin VB.Frame Frame1 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   60
         TabIndex        =   27
         Top             =   6840
         Width           =   11415
         Begin Fox.EBSText etxRegistrosMarcados 
            Height          =   330
            Left            =   4380
            TabIndex        =   31
            Top             =   165
            Width           =   705
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            Enabled         =   0   'False
            TipoCriterio    =   0
            Alinhamento     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSText etxTotalMarcados 
            Height          =   330
            Left            =   6540
            TabIndex        =   32
            Top             =   165
            Width           =   1425
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
            Enabled         =   0   'False
            TipoCriterio    =   6
            Alinhamento     =   1
            Mascara         =   "##,##0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSText etxTotalRegistros 
            Height          =   330
            Left            =   9690
            TabIndex        =   33
            Top             =   165
            Width           =   1305
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
            Enabled         =   0   'False
            TipoCriterio    =   6
            Alinhamento     =   1
            Mascara         =   "##,##0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSText etxRegistrosLocalizados 
            Height          =   330
            Left            =   1920
            TabIndex        =   34
            Top             =   180
            Width           =   705
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            Enabled         =   0   'False
            TipoCriterio    =   0
            Alinhamento     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Registros Localizados:"
            Height          =   225
            Left            =   270
            TabIndex        =   35
            Top             =   255
            Width           =   1605
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Total dos Registros:"
            Height          =   225
            Left            =   8190
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Marcados:"
            Height          =   225
            Left            =   5040
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblRegMarcados 
            Alignment       =   1  'Right Justify
            Caption         =   "Registros Marcados:"
            Height          =   225
            Left            =   2880
            TabIndex        =   28
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraFiltros 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   60
         TabIndex        =   19
         Top             =   990
         Width           =   11415
         Begin VB.CheckBox chkExibeIntegrados 
            Caption         =   "Mostrar registros já integrados"
            Height          =   315
            Left            =   8400
            TabIndex        =   5
            Top             =   570
            Width           =   2475
         End
         Begin Fox.EBSCombo ecbBancos 
            Height          =   315
            Left            =   780
            TabIndex        =   6
            Top             =   540
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            Dados           =   ""
            DadosAssist     =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSData edtPagamentoInicial 
            Height          =   330
            Left            =   8400
            TabIndex        =   3
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSData edtPagamentoFinal 
            Height          =   330
            Left            =   9960
            TabIndex        =   4
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSCombo ecbOrigemDocumentos 
            Height          =   315
            Left            =   780
            TabIndex        =   1
            Top             =   210
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            OrigemDados     =   2
            Dados           =   "Todos;Lançamentos;Duplicatas;Transferências"
            DadosAssist     =   ""
            DefaultValue    =   "Todos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSCombo ecbTipoRegistro 
            Height          =   315
            Left            =   5280
            TabIndex        =   2
            Top             =   210
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            Dados           =   ""
            DadosAssist     =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSText etxCamara 
            Height          =   330
            Left            =   5280
            TabIndex        =   25
            Top             =   540
            Width           =   5625
            _ExtentX        =   442278
            _ExtentY        =   582
            TipoTexto       =   0
            Enabled         =   0   'False
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "cd_camara"
            TipoCriterio    =   3
            CampoDescricao  =   "desc_camara"
            TabelaConsulta  =   "FFICamaras"
            TamanhoDescricao=   5000
            Alinhamento     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblBancosAtendidos 
            AutoSize        =   -1  'True
            Caption         =   "Layouts de Bancos Atendidos"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   780
            TabIndex        =   44
            Top             =   870
            Width           =   2115
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Câmara"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4650
            TabIndex        =   26
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lblBanco 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   225
            TabIndex        =   24
            Top             =   600
            Width           =   495
         End
         Begin VB.Label lblDataPagamento 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento"
            Height          =   195
            Left            =   7470
            TabIndex        =   23
            Top             =   270
            Width           =   840
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Index           =   0
            Left            =   9690
            TabIndex        =   22
            Top             =   270
            Width           =   225
         End
         Begin VB.Label lblOrigem 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Origem"
            Height          =   195
            Left            =   225
            TabIndex        =   21
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Registro"
            Height          =   195
            Index           =   1
            Left            =   4020
            TabIndex        =   20
            Top             =   270
            Width           =   1170
         End
      End
      Begin VB.CommandButton cmdDesmarcarTodos 
         Caption         =   "&Desmarcar Todos"
         Height          =   375
         Left            =   9990
         TabIndex        =   9
         Top             =   2190
         Width           =   1455
      End
      Begin VB.CommandButton cmdMarcarTodos 
         Caption         =   "&Marcar Todos"
         Height          =   375
         Left            =   8490
         TabIndex        =   8
         Top             =   2190
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResultados 
         Height          =   4155
         Left            =   60
         TabIndex        =   15
         Top             =   2610
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7329
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Fox.EBSText etxNumeroLote 
         Height          =   330
         Left            =   1470
         TabIndex        =   17
         Top             =   240
         Width           =   1215
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         Enabled         =   0   'False
         TipoCriterio    =   0
         Alinhamento     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Fox.EBSText etxDescricaoLote 
         Height          =   330
         Left            =   1470
         TabIndex        =   0
         Top             =   600
         Width           =   7635
         _ExtentX        =   265
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   660
         TabIndex        =   18
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo Remessa"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmPagamentoDigitalFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngTotalRegistros As Long
Private mlngNumeroArquivo  As Long
Private mstrBanco()        As String
Private Const grdUnchecked = 1
Private Const grdChecked = 2
Private Const grdAsc = 3
Private Const grdDesc = 4
Private Const colCheck = 1
Private mcurTotalRegistros As Currency

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdConfirmar_Click()
    Dim blnGravar            As Boolean
    Dim lngLinhas            As Long
    Dim strSql               As String
    Dim lngSeq               As Long
    Dim strEmpresa           As String
    Dim strTipo              As String
    Dim dblNumero            As Double
    Dim intParcela           As Integer
    Dim datVencimento        As Date
    Dim strAgencia           As String
    Dim strDvAgencia         As String
    Dim strConta             As String
    Dim strDvConta           As String
    Dim strOrigemDoc         As String
    Dim lngCodigoBanco       As Long
    Dim dblValorTitulo       As Double
    Dim dblValorAcrescimo    As Double
    Dim dblValorAbatimento   As Double
    Dim dblValorMulta        As Double
    Dim strLote              As String
    Dim strTabela            As String
    Dim strCampoCodigo       As String
    Dim strClausula          As String
    Dim strCamCentralizadora As String
    
On Error GoTo error_handler
    BeginTrans
    Call CriaLote
    For lngLinhas = 1 To grdResultados.Rows - 1
        If blnGravar Or lngSeq = 0 Then
            With grdResultados
                If LinhaSelecionada(lngLinhas) Then
                    strEmpresa = .TextMatrix(lngLinhas, 6)
                    strTipo = .TextMatrix(lngLinhas, 3)
                    dblNumero = .TextMatrix(lngLinhas, 4)
                    intParcela = .TextMatrix(lngLinhas, 5)
                    datVencimento = CDate(.TextMatrix(lngLinhas, 9))
                    strOrigemDoc = .TextMatrix(lngLinhas, 2)
                    lngCodigoBanco = .TextMatrix(lngLinhas, 7)
                    dblValorTitulo = .TextMatrix(lngLinhas, 10)
                    strCamCentralizadora = IIf(dblValorTitulo < 5000, "700", "018")
                    Select Case strOrigemDoc
                        Case "Dup"
                            strTabela = "Duplicatas"
                            strCampoCodigo = "Nota"
                        Case "Lan"
                            strTabela = "Lançamentos"
                            strCampoCodigo = "Código"
                        Case "Tra"
                            strTabela = ""
                    End Select
                    If Len(Trim(strTabela)) > 0 Then
                        strClausula = strCampoCodigo & " = " & dblNumero & " AND PagRec = 'P' AND Empresa = '" & strEmpresa & "' AND Tipo = '" & strTipo & "' AND Parcela = " & intParcela
                        dblValorAcrescimo = GetFieldValue("Acréscimo", strTabela, strClausula, , 0)
                        dblValorAbatimento = GetFieldValue("Abatimento", strTabela, strClausula, , 0)
                        dblValorMulta = GetFieldValue("VlrMul", strTabela, strClausula, , 0)
                    Else
                        dblValorAcrescimo = 0
                        dblValorAbatimento = 0
                        dblValorMulta = 0
                    End If
                    lngSeq = lngSeq + 1
                    strSql = "INSERT INTO FFIItemPagamento(cd_arquivoPagamento, cd_itemPagamento, tp_documento, " & _
                        "tipo_registro, nr_documento, nr_parcela, cd_empresa, cd_banco, dt_vencimento, vlr_titulo, vlr_abatimento, vlr_acrescimo, vlr_multa, cd_camara_centralizadora) VALUES(" & mlngNumeroArquivo & ", " & _
                        lngSeq & ", '" & strOrigemDoc & "', '" & strTipo & "', " & _
                        dblNumero & ", " & intParcela & ", '" & strEmpresa & "', " & _
                        lngCodigoBanco & ", '" & InverteData(datVencimento) & "', " & Replace(dblValorTitulo, ",", ".") & ", " & Replace(dblValorAbatimento, ",", ".") & ", " & Replace(dblValorAcrescimo, ",", ".") & _
                        ", " & Replace(dblValorMulta, ",", ".") & ", " & strCamCentralizadora & ")"
                    blnGravar = (ExecuteSQL(strSql) > 0)
                End If
            End With
        Else
            Exit For
        End If
    Next
    If Not blnGravar Then
        Rollback
        MsgBox "Não foi possível gravar o lote.", vbInformation, NomeModulo
    Else
        CommitTrans
        Load frmAlteracaoTipoLancamento
        frmAlteracaoTipoLancamento.Camara = etxCamara.valorInteiro
        frmAlteracaoTipoLancamento.Descricao = etxDescricaoLote.valorTexto
        frmAlteracaoTipoLancamento.NumeroArquivo = mlngNumeroArquivo
        Call frmAlteracaoTipoLancamento.CarregaRegistrosGrid
        Me.Hide
        Call mostrarForm(frmAlteracaoTipoLancamento, 2865, False)
    End If
    Exit Sub

error_handler:
    err.Clear
    Rollback
    MsgBox "Não foi possivel gravar o lote.", vbInformation, NomeModulo
End Sub

Private Sub cmdDesmarcarTodos_Click()
    Dim lngLinhas As Long
    
    With grdResultados
        For lngLinhas = 1 To .Rows - 1
            .col = 1
            .Row = lngLinhas
            Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
        Next
    End With
    etxTotalMarcados.valorMoeda = 0
    etxRegistrosMarcados.valorInteiro = 0
End Sub

Private Sub cmdMarcarTodos_Click()
    Dim lngLinhas As Long
    Dim curTotalMarcados As Currency
    Dim intContItens As Integer
        
    If grdResultados.TextMatrix(1, 2) <> "" Then
        With grdResultados
            For lngLinhas = 1 To .Rows - 1
                .col = 1
                .Row = lngLinhas
                Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                curTotalMarcados = curTotalMarcados + .TextMatrix(.Row, 10)
            Next
        End With
        etxTotalMarcados.valorMoeda = curTotalMarcados
        If lngLinhas > 1 Then
            etxRegistrosMarcados.valorInteiro = lngLinhas - 1
        End If
    End If
End Sub

Private Sub cmdSair_Click()
    Call Unload(Me)
End Sub

Public Sub cmdVisualizar_Click()
    If ValidaCampos Then
        mstrBanco = Split(ecbBancos.SelectedItem, " - ")
        mcurTotalRegistros = 0
        etxRegistrosMarcados.Clear
        etxTotalMarcados.Clear
        Call ExibeRegistros(Trim(mstrBanco(0)))
    End If
End Sub

Private Sub ecbBancos_LostFocus()
    Dim strBanco() As String
    
    If ecbBancos.SelectedItem <> "" Then
        strBanco = Split(ecbBancos.SelectedItem, " - ")
        etxCamara.valorInteiro = GetFieldValue("Câmara", "Bancos", "Banco = " & strBanco(0), , 0)
    End If
End Sub

'Private Sub etxCamara_Change()
'    Dim strBanco() As String
'
'    If etxCamara.valorInteiro > 0 Then
'        If etxCamara.valorInteiro = 1 Then
'            cmdVisualizar.Enabled = True
'        Else
'            cmdVisualizar.Enabled = False
'        End If
'    End If
'End Sub

Private Sub Form_Load()
    Call etxCamara.AddConexao(Aplicacao)
    Call LimpaCampos
End Sub

'Data.......: 19/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para configurar os campos da grid para exibição dos registros a
'               serem integrados.
Private Sub ConfigureGrid()
    Dim intColuna As Integer

    With grdResultados
        .Rows = 2
        .FixedRows = 1
        .Cols = 11
        .FixedCols = 1
        
        .RowHeight(0) = 320
        
        'Coluna Fixa
        .ColWidth(0) = 150
        .TextMatrix(0, 0) = ""
        
        'Configura a coluna de seleção
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        .ColAlignment(1) = flexAlignCenterCenter
        
        'Coluna Documento
        .ColWidth(2) = 700
        .TextMatrix(0, 2) = "Doc."
        .ColAlignment(2) = flexAlignLeftCenter
        
        'Coluna Tipo de Registro
        .ColWidth(3) = 1600
        .TextMatrix(0, 3) = "Tipo"
        .ColAlignment(3) = flexAlignLeftCenter
        
        'Coluna Número
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Número"
        .ColAlignment(4) = flexAlignLeftCenter
        
        'Coluna Parcela
        .ColWidth(5) = 750
        .TextMatrix(0, 5) = "Parc."
        .ColAlignment(5) = flexAlignLeftCenter
        
        'Coluna Empresa
        .ColWidth(6) = 3000
        .TextMatrix(0, 6) = "Empresa"
        .ColAlignment(6) = flexAlignLeftCenter
        
        'Coluna Número Banco
        .ColWidth(7) = 1
        .TextMatrix(0, 7) = "Código do Banco"
        .ColAlignment(7) = flexAlignLeftCenter
        
        'Coluna Banco
        .ColWidth(8) = 850
        .TextMatrix(0, 8) = "Banco"
        .ColAlignment(8) = flexAlignLeftCenter

        'Coluna Emissão
        .ColWidth(9) = 1250
        .TextMatrix(0, 9) = "Vencimento"
        .ColAlignment(9) = flexAlignLeftCenter
        
        'Coluna Valor
        .ColWidth(10) = 1450
        .TextMatrix(0, 10) = "Valor"
        .ColAlignment(10) = flexAlignRightCenter
                
        For intColuna = 0 To .Cols - 1
            If intColuna = 1 Then
                Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            End If
            .TextMatrix(1, intColuna) = ""
        Next
    End With
End Sub

'Data.......: 19/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para limpar e preencher todos os campos da tela conforme
'               valores padrões.
Private Sub LimpaCampos()
    Dim rstRegistro As Object
    Dim strSql      As String
    Dim strAuxiliar As String
    
    'Preenche a combo de tipos de registros.
    strSql = "SELECT Tipo FROM [Tipos Globais] ORDER BY Tipo"
    If AbreRecordset(rstRegistro, strSql) = WL_OK Then
        strAuxiliar = rstRegistro.Fields("Tipo").value
        While Not rstRegistro.EOF
            If rstRegistro.Fields("Tipo").value = "Fatura" Then
                strAuxiliar = rstRegistro.Fields("Tipo").value
            End If
            Call ecbTipoRegistro.AddItem(rstRegistro.Fields("Tipo").value)
            rstRegistro.MoveNext
        Wend
        Call ecbTipoRegistro.AddItem("Todos")
        Call ecbTipoRegistro.SelectItem("Todos")
    End If
    Call FechaRecordset(rstRegistro)
    
    'Preenche a combo de Bancos
    strAuxiliar = ""
    strSql = "SELECT DISTINCT Banco, Nome FROM Bancos WHERE Câmara <> 0 ORDER BY Banco"
    If AbreRecordset(rstRegistro, strSql) = WL_OK Then
        While Not rstRegistro.EOF
            If strAuxiliar = "" Then
                strAuxiliar = rstRegistro.Fields("Banco").value & " - " & rstRegistro.Fields("Nome").value
            End If
            Call ecbBancos.AddItem(rstRegistro.Fields("Banco").value & " - " & rstRegistro.Fields("Nome").value)
            rstRegistro.MoveNext
        Wend
        Call ecbBancos.SelectItem(strAuxiliar)
        Call ecbBancos_LostFocus
        'etxDescricaoLote.SetFocus
    End If
    Call FechaRecordset(rstRegistro)
    
    'Preenche as datas de vencimento padrão
    edtPagamentoInicial.Data = FirstDay(Date)
    edtPagamentoFinal.Data = LastDay(Date)
    
    'Preenche a combo de Origem
    ecbOrigemDocumentos.preencher
    Call ecbOrigemDocumentos.SelectItem("Todos")
    etxNumeroLote.valorInteiro = ProximoNumero("cd_arquivoPagamento", "FFIArquivoPagamento", "")
    'Prepara a grid
    Call ConfigureGrid
End Sub

'Data.......: 19/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento responsável por consultar os registros nas tabelas conforme os
'               parâmetros informados na tela.
'Parâmetros.: [String] Lista dos bancos que devem ser exibidos.
Private Sub ExibeRegistros(strBancos As String)
    mlngTotalRegistros = 0
    Call ConfigureGrid
    If ecbOrigemDocumentos.SelectedItem = "Todos" Or ecbOrigemDocumentos.SelectedItem = "Duplicatas" Then
        Call ExibeDuplicatas(strBancos)
    End If
    If ecbOrigemDocumentos.SelectedItem = "Todos" Or ecbOrigemDocumentos.SelectedItem = "Lançamentos" Then
        Call ExibeLancamentos(strBancos)
    End If
    If ecbOrigemDocumentos.SelectedItem = "Todos" Or ecbOrigemDocumentos.SelectedItem = "Transferências" Then
        Call ExibeTransferencias(strBancos)
    End If
    If grdResultados.Rows > 2 And grdResultados.TextMatrix(1, 1) = "" Then
        Call grdResultados.RemoveItem(1)
    End If
    If mlngTotalRegistros = 0 Then
        cmdConfirmar.Enabled = False
    Else
        etxRegistrosLocalizados.valorInteiro = mlngTotalRegistros
        cmdConfirmar.Enabled = True
        etxTotalRegistros.valorMoeda = mcurTotalRegistros
    End If
End Sub

'Data.......: 19/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento responsável por consultar os registros na tabela de duplicatas
'               conforme os parâmtros informados na tela.
'Parâmetros.: [String] Lista dos bancos a serem exibidos.
Private Sub ExibeDuplicatas(strBancos As String)
    Dim strSql      As String
    Dim rstResult   As Object
    Dim strRegistro As String
    Dim strSqlCC    As String
    Dim curValorOriginal As Currency
    
    strSql = " Banco = " & strBancos & " AND Pagamento IS NULL AND PagRec='P'"
    strSql = strSql & " AND Vencimento BETWEEN " & InverteData(edtPagamentoInicial.Data, True) & " AND " & InverteData(edtPagamentoFinal.Data, True)
    If ecbTipoRegistro.SelectedItem <> "Todos" Then
        strSql = strSql & " AND Tipo='" & ecbTipoRegistro.SelectedItem & "'"
    End If
    If strSql <> "" Then
        strSql = "SELECT Nota, Tipo, Empresa, Parcela, Banco, Vencimento, [Valor Original],proveniente_rateio FROM Duplicatas WHERE " & strSql
        strSql = strSql & " ORDER BY Tipo, Nota, Parcela"
    End If
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        While Not rstResult.EOF
            curValorOriginal = 0
            If Not CBool(GetValue(rstResult, "proveniente_rateio")) Then
                strRegistro = "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "Dup" & Chr(vbKeyTab) & rstResult.Fields("Tipo").value & _
                                Chr(vbKeyTab) & rstResult.Fields("Nota").value & Chr(vbKeyTab) & Format(rstResult.Fields("Parcela").value, "000") & _
                                Chr(vbKeyTab) & rstResult.Fields("Empresa").value & Chr(vbKeyTab) & rstResult.Fields("Banco").value & Chr(vbKeyTab) & _
                                strBancos & Chr(vbKeyTab) & rstResult.Fields("Vencimento").value & Chr(vbKeyTab) & _
                                FormatCurrency(rstResult.Fields("Valor Original").value) & Chr(vbKeyTab) & CBool(rstResult.Fields("proveniente_rateio").value) & Chr(vbKeyTab) & "False" & Chr(vbKeyTab) & "False"
            Else
                'Verificação para não replicar registros na Grid
                If chkExibeIntegrados.value = vbChecked Or Not ExisteRegistroRateio(GetValue(rstResult, "Tipo", ""), GetValue(rstResult, "Nota", 0), GetValue(rstResult, "Empresa", ""), GetValue(rstResult, "Vencimento", 0), "Dup") Then
                    strSqlCC = "SELECT Nota, Tipo, Empresa, Parcela, Banco, Vencimento, [Valor Original],proveniente_rateio FROM Duplicatas WHERE " & _
                               "Nota = " & GetValue(rstResult, "Nota", 0) & " AND Tipo = '" & GetValue(rstResult, "Tipo", 0) & "' AND Empresa = '" & _
                               GetValue(rstResult, "Empresa", 0) & "' AND Vencimento = #" & InverteData(CDate(GetValue(rstResult, "Vencimento", 0))) & "#" & _
                               "ORDER BY Nota,Parcela"
                    Call SQLAgrupaRegistrosCC(strSqlCC, strRegistro, rstResult, "Dup", "Nota", curValorOriginal)
                End If
            End If
            If strRegistro <> "" Then
                If chkExibeIntegrados.value = vbChecked Or Not ExisteRegistro(strToLng(rstResult.Fields("Nota").value), rstResult.Fields("Tipo").value, rstResult.Fields("Parcela").value, "Dup", rstResult.Fields("Empresa").value) Then
                    grdResultados.AddItem strRegistro
                    grdResultados.Row = grdResultados.Rows - 1
                    grdResultados.col = 1
                    Set grdResultados.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                    mlngTotalRegistros = mlngTotalRegistros + 1
                    If curValorOriginal > 0 Then
                        mcurTotalRegistros = mcurTotalRegistros + curValorOriginal
                    Else
                        mcurTotalRegistros = mcurTotalRegistros + rstResult.Fields("Valor Original").value
                    End If
                End If
            End If
            rstResult.MoveNext
            strRegistro = ""
        Wend
    End If
    Call FechaRecordset(rstResult)
End Sub

'Data.......: 22/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento responsável por consultar os registros na tabela de lançamentos
'               conforme os parâmetros informados na tela.
'Parâmetros.: [String] Lista dos bancos a serem exibidos.
Private Sub ExibeLancamentos(strBancos As String)
    Dim strSql      As String
    Dim rstResult   As Object
    Dim strRegistro As String
    Dim strSqlCC    As String
    Dim curValorOriginal As Currency
    
    strSql = " Banco IN(" & strBancos & ") AND Pagamento IS NULL AND PagRec='P'"
    strSql = strSql & " AND Vencimento BETWEEN " & InverteData(edtPagamentoInicial.Data, True) & " AND " & InverteData(edtPagamentoFinal.Data, True)
    If ecbTipoRegistro.SelectedItem <> "Todos" Then
        strSql = strSql & " AND Tipo='" & ecbTipoRegistro.SelectedItem & "'"
    End If
    If strSql <> "" Then
        strSql = "SELECT Código, Tipo, Empresa, Banco, Parcela, Vencimento, [Valor Original], proveniente_rateio FROM Lançamentos WHERE " & strSql
        strSql = strSql & " ORDER BY Tipo, Código, Parcela"
    End If
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        While Not rstResult.EOF
            curValorOriginal = 0
            If Not CBool(GetValue(rstResult, "proveniente_rateio")) Then
                strRegistro = "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "Lan" & Chr(vbKeyTab) & rstResult.Fields("Tipo").value & _
                    Chr(vbKeyTab) & rstResult.Fields("Código").value & Chr(vbKeyTab) & Format(rstResult.Fields("Parcela").value, "000") & _
                    Chr(vbKeyTab) & rstResult.Fields("Empresa").value & Chr(vbKeyTab) & rstResult.Fields("Banco").value & Chr(vbKeyTab) & _
                    strBancos & Chr(vbKeyTab) & rstResult.Fields("Vencimento").value & Chr(vbKeyTab) & _
                    FormatCurrency(rstResult.Fields("Valor Original").value) & Chr(vbKeyTab) & CBool(rstResult.Fields("proveniente_rateio").value) & Chr(vbKeyTab) & "False" & Chr(vbKeyTab) & "False"
            Else
                'Verificação para não replicar registros na Grid
                If chkExibeIntegrados.value = vbChecked Or Not ExisteRegistroRateio(GetValue(rstResult, "Tipo", ""), GetValue(rstResult, "Código", 0), GetValue(rstResult, "Empresa", ""), GetValue(rstResult, "Vencimento", 0), "Lan") Then
                    strSqlCC = "SELECT Código, Tipo, Empresa, Parcela, Banco, Vencimento, [Valor Original],proveniente_rateio FROM Lançamentos WHERE " & _
                               "Código = " & GetValue(rstResult, "Código", 0) & " AND Tipo = '" & GetValue(rstResult, "Tipo", 0) & "' AND Empresa = '" & _
                               GetValue(rstResult, "Empresa", 0) & "' AND Vencimento = #" & InverteData(CDate(GetValue(rstResult, "Vencimento", 0))) & "#" & _
                               "ORDER BY Código,Parcela"
                    Call SQLAgrupaRegistrosCC(strSqlCC, strRegistro, rstResult, "Lan", "Código", curValorOriginal)
                End If
            End If
            If strRegistro <> "" Then
                If chkExibeIntegrados.value = vbChecked Or Not ExisteRegistro(strToDbl(rstResult.Fields("Código").value), rstResult.Fields("Tipo").value, rstResult.Fields("Parcela").value, "Lan", rstResult.Fields("Empresa").value) Then
                    grdResultados.AddItem strRegistro
                    grdResultados.Row = grdResultados.Rows - 1
                    grdResultados.col = 1
                    Set grdResultados.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                    mlngTotalRegistros = mlngTotalRegistros + 1
                    If curValorOriginal > 0 Then
                        mcurTotalRegistros = mcurTotalRegistros + curValorOriginal
                    Else
                        mcurTotalRegistros = mcurTotalRegistros + rstResult.Fields("Valor Original").value
                    End If
                End If
            End If
            rstResult.MoveNext
            strRegistro = ""
        Wend
    End If
    Call FechaRecordset(rstResult)
End Sub

'Data.......: 22/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento responsável por consultar os registros na tabela de Transferência bancária.
'Parâmetros.: [String] Lista dos bancos a serem listados.
Private Sub ExibeTransferencias(strBancos As String)
    Dim strSql      As String
    Dim rstResult   As Object
    Dim strRegistro As String
    
    strSql = " Origem = " & strBancos & ""
    strSql = strSql & " AND Data BETWEEN " & InverteData(edtPagamentoInicial.Data, True) & " AND " & InverteData(edtPagamentoFinal.Data, True)
    If ecbTipoRegistro.SelectedItem <> "Todos" Then
        strSql = strSql & " AND Tipo_registro='" & ecbTipoRegistro.SelectedItem & "'"
    End If
    If strSql <> "" Then
        strSql = "SELECT Código, Origem, Tipo_registro, Destino, Data, Valor, empresa_favorecida  FROM [Transf Bancária] WHERE " & strSql
        strSql = strSql & " ORDER BY Tipo_registro, Código"
    End If
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        While Not rstResult.EOF
            strRegistro = "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & "Tra" & Chr(vbKeyTab) & rstResult.Fields("Tipo_registro").value & _
                Chr(vbKeyTab) & rstResult.Fields("Código").value & Chr(vbKeyTab) & Format("0", "000") & Chr(vbKeyTab) & _
                rstResult.Fields("empresa_favorecida").value & Chr(vbKeyTab) & rstResult.Fields("Origem").value & Chr(vbKeyTab) & _
                strBancos & Chr(vbKeyTab) & rstResult.Fields("Data").value & Chr(vbKeyTab) & _
                FormatCurrency(rstResult.Fields("Valor").value) & Chr(vbKeyTab) & "False" & Chr(vbKeyTab) & "False" & Chr(vbKeyTab) & "False"
            If chkExibeIntegrados.value = vbChecked Or Not ExisteRegistro(strToLng(rstResult.Fields("Código").value), rstResult.Fields("Tipo_registro").value, 0, "Tra", "") Then
                grdResultados.AddItem strRegistro
                grdResultados.Row = grdResultados.Rows - 1
                grdResultados.col = 1
                Set grdResultados.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                mlngTotalRegistros = mlngTotalRegistros + 1
                mcurTotalRegistros = mcurTotalRegistros + rstResult.Fields("Valor").value
            End If
            rstResult.MoveNext
        Wend
    End If
End Sub

'Data.......: 22/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento responsável por traduzir o código do banco para a descrição do mesmo
'               utilizado no preenchimento da grid dos registros de transferências.
'Parametros.: [Long] Código do banco a ser traduzido.
'Retorno....: [String] Nome do banco correspondente ao código passado.
Private Function NomeBanco(lngCodigo As Long) As String
    Dim strSql   As String
    Dim rstBanco As Object
    
    strSql = "SELECT Nome FROM Bancos WHERE Banco=" & lngCodigo
    If AbreRecordset(rstBanco, strSql) = WL_OK Then
        NomeBanco = rstBanco.Fields("Nome").value
    End If
    Call FechaRecordset(rstBanco)
End Function

'Data.......: 22/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Função responsável por validar o preenchimento dos campos antes de mostrar os registros
'               a serem gerados no arquivo.
'Retorno....: [Boolean] Retorna se o processo pode ser excutado ou não.
Private Function ValidaCampos() As Boolean
    If Not edtPagamentoInicial.IsValidDate Then
        MsgBox "A data de Vencimento inicial é de preenchimento obrigatório.", vbInformation, NomeModulo
        edtPagamentoInicial.SetFocus
    ElseIf Not edtPagamentoFinal.IsValidDate Then
        MsgBox "A data de Vencimento final é de preenchimento obrigatório.", vbInformation, NomeModulo
        edtPagamentoFinal.SetFocus
    ElseIf etxCamara.valorInteiro = 0 Then
        MsgBox "O banco selecionado não tem câmara cadastrado para ele.", vbInformation, NomeModulo
    Else
        ValidaCampos = True
    End If
End Function

Private Sub grdResultados_Click()
    With grdResultados
        If .TextMatrix(.Row, 2) <> "" Then
            .CellPictureAlignment = flexAlignCenterCenter
            If LinhaSelecionada(.Row) Then
                Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                etxRegistrosMarcados.valorInteiro = etxRegistrosMarcados.valorInteiro - 1
                etxTotalMarcados.valorMoeda = etxTotalMarcados.valorMoeda - .TextMatrix(.Row, 10)
            Else
                Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                etxRegistrosMarcados.valorInteiro = etxRegistrosMarcados.valorInteiro + 1
                etxTotalMarcados.valorMoeda = etxTotalMarcados.valorMoeda + .TextMatrix(.Row, 10)
            End If
        End If
    End With
End Sub

'Data.......: 25/09/2007
'Autor......: Dulcino Júnior
'Descrição..: Função responsável por verificar se o registro está selecionado
'               na grid de resultados.
'Parametros.: [Long] Número da linha a ser verificada.
'Retorno....: [Boolean] Retorna o estado de seleção da linha.
Private Function LinhaSelecionada(lngLinha As Long) As Boolean
    If lngLinha <= grdResultados.Rows - 1 Then
        grdResultados.Row = lngLinha
        grdResultados.col = colCheck
        LinhaSelecionada = (grdResultados.CellPicture = imgCheck.ListImages(grdChecked).Picture)
    Else
        LinhaSelecionada = False
    End If
End Function

'Data.......: 26/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para gerar o cabeçalho do lote a ser gerado no arquivo de remessa
'               do banco.
'Retorno....: [Boolean] Retorna o resultado da execução da função, se conseguiu inserir o registro
'               ou não.
Private Function CriaLote() As Boolean
    Dim strSql    As String
    Dim rstResult As Object
    Dim lngCodigo As Long
    
    strSql = "SELECT MAX(nr_Remessa_PagFor) AS LastCod FROM Bancos"
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        If Not rstResult.EOF Then
            lngCodigo = strToLng(rstResult.Fields("LastCod").value & "") + 1
        Else
            lngCodigo = 1
        End If
    End If
    mlngNumeroArquivo = lngCodigo
End Function

'Data.......: 26/09/2008
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para verificar se o registro já está cadastrado para outro lote de  pagamento
'Parametros.: [Long] Número do documento que está sendo verificado.
'             [String] Tipo de registro do documento que está sendo verificado.
'             [Long] Número da parcela do documento que está sendo verificado.
'             [String] Tipo de documento que está sendo verificado Tra, Dup, Lan.
'             [String] Empresa a quem o documento pertence.
'Retorno....: [Boolean] Retorna se o registro já está cadastrado em outro lote.
Private Function ExisteRegistro(lngNumeroDocumento As Double, strTipoRegistro As String, lngNumeroParcela As Long, strDocumento As String, strEmpresa As String) As Boolean
    Dim strSql    As String
    Dim rstResult As Object
    
On Error GoTo error_handler
    With grdResultados
        strSql = "SELECT cd_arquivoPagamento FROM FFIItemPagamento WHERE nr_documento=" & lngNumeroDocumento & _
            " AND tipo_registro='" & strTipoRegistro & "' AND nr_parcela=" & lngNumeroParcela & _
            " AND ((tp_documento='" & strDocumento & "' AND cd_empresa='" & strEmpresa & _
            "') OR tp_documento='" & strDocumento & "')"
        If AbreRecordset(rstResult, strSql) = WL_OK Then
            ExisteRegistro = Not rstResult.EOF
        End If
        Call FechaRecordset(rstResult)
    End With
    Exit Function
error_handler:
    err.Clear
    ExisteRegistro = True
End Function

'Data.......: 09/10/2008
'Autor......: Ivo Sousa
'Descrição..: Função utilizada para compor a SQL para registros reteados por Centro de Custo
'Parametros.: [Long] Número do Lote.
'             [Long] Sequencia do item no Lote
'             [Integer] Linha onde foi encontrado o primeiro registro identificado como rateio.
'Retorno....: [String] SQL com os registros agrupados.
'Private Function SQLAgrupaRegistrosCC(lngSeq As Long, intRow As Integer) As String
Private Function SQLAgrupaRegistrosCC(strSql As String, ByRef strRegistro As String, ByRef rstResult As Object, strOrigem As String, strCampoCodigo As String, ByRef curValorOriginal As Currency) As String
    Dim intCont   As Integer
    Dim dblTotal  As Double
    Dim rstRateioCC As Object
        
    If AbreRecordset(rstRateioCC, strSql) = WL_OK Then
        rstRateioCC.MoveFirst
        While Not rstRateioCC.EOF
            dblTotal = dblTotal + GetValue(rstRateioCC, "Valor Original", 0)
            rstRateioCC.MoveNext
        Wend
        curValorOriginal = dblTotal
        strRegistro = "" & Chr(vbKeyTab) & "" & Chr(vbKeyTab) & strOrigem & Chr(vbKeyTab) & rstResult.Fields("Tipo").value & _
                                Chr(vbKeyTab) & rstResult.Fields(strCampoCodigo).value & Chr(vbKeyTab) & Format(rstResult.Fields("Parcela").value, "000") & _
                                Chr(vbKeyTab) & rstResult.Fields("Empresa").value & Chr(vbKeyTab) & rstResult.Fields("Banco").value & Chr(vbKeyTab) & _
                                mstrBanco(0) & Chr(vbKeyTab) & rstResult.Fields("Vencimento").value & Chr(vbKeyTab) & _
                                FormatCurrency(dblTotal) & Chr(vbKeyTab) & CBool(rstResult.Fields("proveniente_rateio").value) & Chr(vbKeyTab) & "False" & Chr(vbKeyTab) & "False"
    End If
End Function

Private Function ExisteRegistroRateio(strTipo As String, lngNumero As Long, strEmpresa As String, datVencimento As Date, strDocumento As String) As Boolean
    Dim intCont As Integer
    
    If GetFieldValue("nr_documento", "FFIItemPagamento", "tipo_registro = '" & strTipo & "' AND nr_documento = " & lngNumero & " AND cd_empresa = '" & strEmpresa & "' AND dt_vencimento = #" & InverteData(datVencimento) & "# AND tp_documento = '" & strDocumento & "'", , 0) > 0 Then
        ExisteRegistroRateio = True
        Exit Function
    End If
    With grdResultados
        For intCont = 1 To .Rows - 1
            If (.TextMatrix(intCont, 3) = strTipo) And (strToLng(.TextMatrix(intCont, 4)) = lngNumero) And (.TextMatrix(intCont, 6) = strEmpresa) Then
                If (.TextMatrix(intCont, 9) <> "") Then
                    If (.TextMatrix(intCont, 9) = datVencimento) Then
                        ExisteRegistroRateio = True
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    ExisteRegistroRateio = False
End Function

Private Sub imgCol1_Click()
    If imgCol1.Picture = imgCheck.ListImages(grdAsc).Picture Then
        imgCol1.Picture = imgCheck.ListImages(grdDesc).Picture
        Call ConfigureGrid
        'Call MostraRegistro("cd_tipo_servico DESC", True, False, False, False, False)
    Else
        imgCol1.Picture = imgCheck.ListImages(grdAsc).Picture
        Call ConfigureGrid
        'Call MostraRegistro("cd_tipo_servico ASC", True, False, False, False, False)
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

'Projeto: - Desenv.: - Ueder Budni (14/12/2017)
Private Sub lblBancosAtendidos_Click()
    frmLayoutsBancosAtendidos.Show vbModal
End Sub

'Projeto: - Desenv.: - Ueder Budni (14/12/2017)
Private Sub lblBancosAtendidos_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    SetCursor (LoadCursor(0, IDC_HAND))
End Sub
