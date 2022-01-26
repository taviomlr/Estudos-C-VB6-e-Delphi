VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmConciliacaoTitulosAutomatica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliação Bancária Automática"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   13935
   Begin VB.CommandButton cmdConciliacaoAutomatica 
      Caption         =   "C&onciliação Automática"
      Height          =   375
      Left            =   9720
      TabIndex        =   70
      Top             =   3150
      Width           =   2505
   End
   Begin VB.Frame fraBotoes 
      Height          =   10350
      Left            =   12450
      TabIndex        =   24
      Top             =   -45
      Width           =   1455
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   130
         TabIndex        =   81
         Top             =   1020
         Width           =   1185
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   135
         TabIndex        =   20
         Top             =   225
         Width           =   1185
      End
      Begin VB.CommandButton cmdConciliar 
         Caption         =   "&Conciliar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   135
         TabIndex        =   21
         Top             =   630
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   135
         TabIndex        =   22
         Top             =   1410
         Width           =   1185
      End
      Begin ComctlLib.ImageList imgCheck 
         Left            =   450
         Top             =   1770
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   3
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmConciliacaoTitulosAutomatica.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmConciliacaoTitulosAutomatica.frx":0352
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmConciliacaoTitulosAutomatica.frx":06A4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraControles 
      Height          =   10350
      Left            =   30
      TabIndex        =   23
      Top             =   -45
      Width           =   12375
      Begin VB.Frame fraExtrato 
         Caption         =   "Extrato"
         Height          =   615
         Left            =   60
         TabIndex        =   66
         Top             =   2490
         Width           =   12255
         Begin VB.CommandButton cmdImportaExtrato 
            Caption         =   "&Importar/Digitar Extrato"
            Enabled         =   0   'False
            Height          =   375
            Left            =   9630
            TabIndex        =   69
            Top             =   160
            Width           =   2505
         End
         Begin Fox.EBSText etxExtratoBancario 
            Height          =   330
            Left            =   1575
            TabIndex        =   67
            Top             =   180
            Width           =   1260
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            TipoCriterio    =   4
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
            ExibeDescricao  =   0   'False
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Extrato Bancário"
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
            Left            =   90
            TabIndex        =   68
            Top             =   270
            Width           =   1425
         End
      End
      Begin TabDlg.SSTab tabConciliados 
         Height          =   6975
         Left            =   60
         TabIndex        =   48
         Top             =   3300
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   12303
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Lançamentos/Duplicatas não Conciliados"
         TabPicture(0)   =   "frmConciliacaoTitulosAutomatica.frx":0C36
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label14"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "etxDiferenca"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fraResultados"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Lançamentos/Duplicatas Conciliados"
         TabPicture(1)   =   "frmConciliacaoTitulosAutomatica.frx":0C52
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame3 
            Caption         =   "Lançamentos/Duplicatas Conciliados"
            Height          =   6510
            Left            =   -74940
            TabIndex        =   73
            Top             =   390
            Width           =   12075
            Begin VB.CommandButton cmdDesconciliarTodos 
               Caption         =   "Desconciliar Todos"
               Height          =   375
               Left            =   10380
               TabIndex        =   76
               Top             =   210
               Width           =   1575
            End
            Begin VB.CommandButton cmdDesconciliar 
               Caption         =   "Desconciliar"
               Height          =   375
               Left            =   8700
               TabIndex        =   75
               Top             =   210
               Width           =   1575
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConciliados 
               Height          =   5775
               Left            =   60
               TabIndex        =   74
               Top             =   645
               Width           =   11940
               _ExtentX        =   21061
               _ExtentY        =   10186
               _Version        =   393216
               FixedCols       =   0
               FocusRect       =   0
               AllowUserResizing=   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Lançamentos Extrato Bancário"
            Height          =   6090
            Left            =   60
            TabIndex        =   57
            Top             =   390
            Width           =   6045
            Begin VB.CommandButton cmdInserirLancamento 
               Caption         =   "Incluir Lançamento Sistema"
               Height          =   375
               Left            =   1320
               TabIndex        =   78
               Top             =   240
               Width           =   2145
            End
            Begin VB.CommandButton cmdSelecionaTodosExtrato 
               Caption         =   "Todos"
               Height          =   375
               Left            =   3510
               TabIndex        =   59
               Top             =   240
               Width           =   1185
            End
            Begin VB.CommandButton cmdSelecionaNenhumExtrato 
               Caption         =   "&Nenhum"
               Height          =   375
               Left            =   4740
               TabIndex        =   58
               Top             =   240
               Width           =   1215
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdExtrato 
               Height          =   4935
               Left            =   60
               TabIndex        =   60
               Top             =   690
               Width           =   5910
               _ExtentX        =   10425
               _ExtentY        =   8705
               _Version        =   393216
               FixedCols       =   0
               FocusRect       =   0
               SelectionMode   =   1
               AllowUserResizing=   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin Fox.EBSText etxQuantidadeSelecionadaExtrato 
               Height          =   330
               Left            =   2715
               TabIndex        =   61
               Top             =   5670
               Width           =   870
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
            Begin Fox.EBSText etxTotalValorExtrato 
               Height          =   330
               Left            =   4515
               TabIndex        =   62
               Top             =   5670
               Width           =   1455
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
            Begin VB.Label lblTipoOperacaoExtrato 
               BackColor       =   &H8000000A&
               Height          =   285
               Left            =   390
               TabIndex        =   79
               Top             =   270
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label lblBancoExtrato 
               BackColor       =   &H8000000A&
               Height          =   285
               Left            =   90
               TabIndex        =   77
               Top             =   270
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total"
               Enabled         =   0   'False
               Height          =   195
               Left            =   3660
               TabIndex        =   64
               Top             =   5715
               Width           =   765
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Selecionados"
               Enabled         =   0   'False
               Height          =   195
               Left            =   1680
               TabIndex        =   63
               Top             =   5715
               Width           =   960
            End
         End
         Begin VB.Frame fraResultados 
            Caption         =   "Lançamentos/Duplicatas"
            Height          =   6090
            Left            =   6150
            TabIndex        =   49
            Top             =   390
            Width           =   6015
            Begin VB.CommandButton cmdSelecionaTodos 
               Caption         =   "Todos "
               Height          =   375
               Left            =   3420
               TabIndex        =   51
               Top             =   240
               Width           =   1215
            End
            Begin VB.CommandButton cmdSelecionaNenhum 
               Caption         =   "&Nenhum"
               Height          =   375
               Left            =   4680
               TabIndex        =   50
               Top             =   240
               Width           =   1215
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConcTitulo 
               Height          =   4935
               Left            =   60
               TabIndex        =   52
               Top             =   690
               Width           =   5880
               _ExtentX        =   10372
               _ExtentY        =   8705
               _Version        =   393216
               FixedCols       =   0
               FocusRect       =   0
               AllowUserResizing=   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin Fox.EBSText etxQuantidadeSelecionada 
               Height          =   330
               Left            =   3195
               TabIndex        =   53
               Top             =   5670
               Width           =   600
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
            Begin Fox.EBSText etxTotalValor 
               Height          =   330
               Left            =   4725
               TabIndex        =   54
               Top             =   5670
               Width           =   1215
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
            Begin VB.Label lblTipoOperacaoLanc 
               BackColor       =   &H8000000A&
               Height          =   285
               Left            =   120
               TabIndex        =   80
               Top             =   270
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Selecionados"
               Enabled         =   0   'False
               Height          =   195
               Left            =   2160
               TabIndex        =   56
               Top             =   5715
               Width           =   960
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total"
               Enabled         =   0   'False
               Height          =   195
               Left            =   3870
               TabIndex        =   55
               Top             =   5715
               Width           =   765
            End
         End
         Begin Fox.EBSText etxDiferenca 
            Height          =   330
            Left            =   10875
            TabIndex        =   71
            Top             =   6570
            Width           =   1215
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
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Diferença"
            Enabled         =   0   'False
            Height          =   195
            Left            =   10110
            TabIndex        =   72
            Top             =   6615
            Width           =   690
         End
      End
      Begin VB.Frame fraFiltros 
         Caption         =   "Filtros"
         Height          =   1965
         Left            =   60
         TabIndex        =   25
         Top             =   540
         Width           =   12255
         Begin VB.Frame Frame2 
            Caption         =   "Pagamento / Recebimento"
            Height          =   555
            Left            =   7845
            TabIndex        =   45
            Top             =   1335
            Width           =   4330
            Begin VB.OptionButton optOrdenarLiberacao 
               Caption         =   "Ordenar por Liberação"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   270
               Value           =   -1  'True
               Width           =   1905
            End
            Begin VB.OptionButton optOrdenarPagamento 
               Caption         =   "Ordenar por Pagamento"
               Height          =   195
               Left            =   2160
               TabIndex        =   46
               Top             =   270
               Width           =   1995
            End
         End
         Begin VB.Frame fraPagamentoRecebimento 
            Caption         =   "Pagamento / Recebimento"
            Height          =   555
            Left            =   75
            TabIndex        =   40
            Top             =   1335
            Width           =   3840
            Begin VB.OptionButton optAmbosPagRec 
               Caption         =   "Ambos"
               Height          =   195
               Left            =   120
               TabIndex        =   65
               Top             =   270
               Value           =   -1  'True
               Width           =   870
            End
            Begin VB.OptionButton optPagamento 
               Caption         =   "Pagamento"
               Height          =   195
               Left            =   1140
               TabIndex        =   7
               Top             =   270
               Width           =   1230
            End
            Begin VB.OptionButton optRecebimento 
               Caption         =   "Recebimento"
               Height          =   195
               Left            =   2415
               TabIndex        =   8
               Top             =   270
               Width           =   1275
            End
         End
         Begin VB.Frame fraOrigem 
            Caption         =   "Origem dos Registros"
            Height          =   555
            Left            =   3960
            TabIndex        =   39
            Top             =   1335
            Width           =   3840
            Begin VB.OptionButton optAmbos 
               Caption         =   "Ambos"
               Height          =   195
               Left            =   135
               TabIndex        =   17
               Top             =   270
               Value           =   -1  'True
               Width           =   870
            End
            Begin VB.OptionButton optLancamentos 
               Caption         =   "Lançamentos"
               Height          =   195
               Left            =   1275
               TabIndex        =   18
               Top             =   270
               Width           =   1320
            End
            Begin VB.OptionButton optDuplicatas 
               Caption         =   "Duplicatas"
               Height          =   195
               Left            =   2700
               TabIndex        =   19
               Top             =   270
               Width           =   1050
            End
         End
         Begin Fox.EBSData edtEmissaoInicial 
            Height          =   330
            Left            =   1545
            TabIndex        =   0
            Top             =   225
            Width           =   1245
            _ExtentX        =   2196
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
         Begin Fox.EBSData edtEmissaoFinal 
            Height          =   330
            Left            =   3015
            TabIndex        =   1
            Top             =   225
            Width           =   1245
            _ExtentX        =   2196
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
         Begin Fox.EBSData edtPagamentoInicial 
            Height          =   330
            Left            =   1545
            TabIndex        =   2
            Top             =   585
            Width           =   1245
            _ExtentX        =   2196
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
            Left            =   3015
            TabIndex        =   3
            Top             =   585
            Width           =   1245
            _ExtentX        =   2196
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
         Begin Fox.EBSData edtLiberacaoInicial 
            Height          =   330
            Left            =   1545
            TabIndex        =   4
            Top             =   945
            Width           =   1245
            _ExtentX        =   2196
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
         Begin Fox.EBSData edtLiberacaoFinal 
            Height          =   330
            Left            =   3015
            TabIndex        =   5
            Top             =   945
            Width           =   1245
            _ExtentX        =   2196
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
         Begin Fox.EBSText etxValorInicial 
            Height          =   330
            Left            =   8805
            TabIndex        =   9
            Top             =   225
            Width           =   1275
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
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
         Begin Fox.EBSText etxValorFinal 
            Height          =   330
            Left            =   10455
            TabIndex        =   10
            Top             =   225
            Width           =   1230
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
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
         Begin Fox.EBSText etxBancoInicial 
            Height          =   330
            Left            =   8805
            TabIndex        =   11
            Top             =   585
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Banco"
            TipoCriterio    =   4
            CampoDescricao  =   "Nome"
            TabelaConsulta  =   "Bancos"
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxBancoFinal 
            Height          =   330
            Left            =   10455
            TabIndex        =   12
            Top             =   585
            Width           =   1230
            _ExtentX        =   2090
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Banco"
            TipoCriterio    =   4
            CampoDescricao  =   "Nome"
            TabelaConsulta  =   "Bancos"
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxContaInicial 
            Height          =   330
            Left            =   5115
            TabIndex        =   13
            Top             =   585
            Width           =   1275
            _ExtentX        =   2090
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Contas"
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxContaFinal 
            Height          =   330
            Left            =   6645
            TabIndex        =   14
            Top             =   585
            Width           =   1230
            _ExtentX        =   1931
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Contas"
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSCombo ecbTipo 
            Height          =   315
            Left            =   5115
            TabIndex        =   6
            Top             =   225
            Width           =   2760
            _ExtentX        =   4868
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
         Begin Fox.EBSText etxChequeInicial 
            Height          =   330
            Left            =   5115
            TabIndex        =   15
            Top             =   945
            Width           =   1275
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            TipoCriterio    =   4
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
         Begin Fox.EBSText etxChequeFinal 
            Height          =   330
            Left            =   6645
            TabIndex        =   16
            Top             =   945
            Width           =   1230
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            TipoCriterio    =   4
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   6450
            TabIndex        =   43
            Top             =   990
            Width           =   90
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cheque"
            Height          =   195
            Left            =   4470
            TabIndex        =   42
            Top             =   990
            Width           =   555
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reg."
            Height          =   195
            Left            =   4335
            TabIndex        =   38
            Top             =   315
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   6450
            TabIndex        =   37
            Top             =   675
            Width           =   90
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   4620
            TabIndex        =   36
            Top             =   675
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   10200
            TabIndex        =   35
            Top             =   675
            Width           =   90
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   8250
            TabIndex        =   34
            Top             =   675
            Width           =   465
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   10200
            TabIndex        =   33
            Top             =   315
            Width           =   90
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   8340
            TabIndex        =   32
            Top             =   315
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   2850
            TabIndex        =   31
            Top             =   1035
            Width           =   90
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Liberação"
            Height          =   195
            Left            =   765
            TabIndex        =   30
            Top             =   1035
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   2850
            TabIndex        =   29
            Top             =   675
            Width           =   90
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Pagamento"
            Height          =   195
            Left            =   675
            TabIndex        =   28
            Top             =   675
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   2850
            TabIndex        =   27
            Top             =   270
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Left            =   885
            TabIndex        =   26
            Top             =   315
            Width           =   585
         End
      End
      Begin Fox.EBSText etxEmpUser 
         Height          =   330
         Left            =   1605
         TabIndex        =   44
         Top             =   180
         Width           =   7005
         _ExtentX        =   441404
         _ExtentY        =   582
         Tipo            =   4
         Enabled         =   0   'False
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Apel"
         CampoDescricao  =   "Razão"
         TabelaConsulta  =   "Empresas"
         TamanhoDescricao=   4500
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
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Enabled         =   0   'False
         Height          =   195
         Left            =   840
         TabIndex        =   41
         Top             =   225
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmConciliacaoTitulosAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const grdChecked = 2
Private Const grdUnchecked = 1
Private Const grdFind = 3
Private Const colCheck = 1
Private Const colFind = 2
Private Const colConcLiberacao = 3
Private Const colConcDescricao = 4
Private Const colConcValor = 5
Private Const colConcDebCred = 6
Private Const colConcEmpresa = 7
Private Const colConcNumero = 8
Private Const colConcParcela = 9
Private Const colConcTipo = 10
Private Const colConcOrigem = 11
Private Const colConcBanco = 12
Private Const colConcExtrato = 13
Private Const colConcSeqExtrato = 14
Private mlngTListado                As Long
Private mdblTValor                  As Double
Private mbizExtratoBanc             As BizImpDigExtratoBancario
Private mcolLancamentos             As New ColImpDigExtratoBancario
Private mlngSeq                     As Long
Private mblnOrigemTelaConciliacao   As Boolean
Public mblnTemAcessoTela            As Boolean
Private mblnJaRespondeuMsgBox       As Boolean
Private mblnExitSub                 As Boolean

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdConciliacaoAutomatica_Click()
    Dim i                  As Integer
    Dim j                  As Integer
    Dim k                  As Integer
    Dim intContConciliados As Integer
    Dim objDuplicata       As CDuplicata
    Dim lngNumeroLanc      As Long
    Dim dblValorSemRateio  As Double
    
    intContConciliados = 0
    
    With grdConcTitulo
    'Verifica se tem algo na grid de extrato e na grid de lançamentos/duplicatas
    If (grdExtrato.Rows >= 2 And grdExtrato.TextMatrix(grdExtrato.Row, 3) <> "") And (.Rows >= 2 And .TextMatrix(.Row, 3) <> "") Then
        If etxQuantidadeSelecionada.valorInteiro = 0 And etxQuantidadeSelecionadaExtrato.valorInteiro = 0 Then
            If .TextMatrix(.Row, 12) = lblBancoExtrato.Caption Then
               For i = 1 To grdExtrato.Rows - 1
                    For j = 1 To grdConcTitulo.Rows - 1
                        If Format(grdExtrato.TextMatrix(i, 2), "dd/mm/yyyy") = Format(.TextMatrix(j, 3), "dd/mm/yyyy") Then
                            If grdExtrato.TextMatrix(i, 5) = .TextMatrix(j, colConcDebCred) Then
                                If grdExtrato.TextMatrix(i, 4) = .TextMatrix(j, 5) Then
                                        grdExtrato.col = colCheck
                                        grdExtrato.Row = i
                                        Set grdExtrato.CellPicture = imgCheck.ListImages(grdChecked).Picture
                                        .col = colCheck
                                        .Row = j
                                        Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                                        Call ConciliaOuDesconcilia(True, False, grdConcTitulo, True)
                                        intContConciliados = intContConciliados + 1
                                    Exit For
                                Else
                                    Set objDuplicata = New CDuplicata
                                    dblValorSemRateio = objDuplicata.ValorSemRateio(.TextMatrix(j, colConcOrigem), .TextMatrix(j, colConcNumero), .TextMatrix(j, colConcEmpresa), IIf(.TextMatrix(j, colConcDebCred) = "Débito", "P", "R"), .TextMatrix(j, colConcTipo), .TextMatrix(j, colConcLiberacao))
                                    'Arrumar aqui - quando seleciona mais de um extrato para um lançamento Vinicius 23/01/2015
                                    If grdExtrato.TextMatrix(i, 4) = dblValorSemRateio Then
                                        lngNumeroLanc = .TextMatrix(j, colConcNumero)
                                        .col = colCheck
                                        .Row = j
                                        Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                                        
                                        grdExtrato.col = colCheck
                                        grdExtrato.Row = i
                                        Set grdExtrato.CellPicture = imgCheck.ListImages(grdChecked).Picture
                                        'Atualizar os lançamentos com número X
                                        Call ConciliaOuDesconcilia(True, False, grdConcTitulo, True, True)
                                        intContConciliados = intContConciliados + 1
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next
               Next
            Else
                MsgBox "Não foi possível fazer a conciliação automática pois o banco do extrato não corresponde ao banco do lançamento.", vbInformation, "Conciliação Bancária"
            End If
        Else
            MsgBox "Para fazer a conciliação automática é necessário desmarcar todos os lançamentos.", vbInformation, NomeModulo
        End If
    End If
    End With
        
    If intContConciliados > 0 Then
        MsgBox "Conciliação automática feita com sucesso. " & vbNewLine & "Foi(ram) conciliado(s) automaticamente " & intContConciliados & " lançamento(s).", vbInformation, "Conciliação Bancária"
        tabConciliados.Tab = 0
        Call RecarregaGrids(True, False)
    Else
        MsgBox "Não foi encontrado lançamento e/ou duplicata para realizar a conciliação automática.", vbInformation, "Conciliação Bancária Automática"
    End If
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: cmdConciliar_Click
'Data.......: 06/05/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para realizar  conciliação bancária para os títulos selecionados.
'Parametros.:
'Retorno....:
'Protocolo..: 82528
'---------------------------------------------------------------------------------------
Private Sub cmdConciliar_Click()
    Call ConciliaOuDesconcilia(True, False, grdConcTitulo, False)
    cmdConciliar.Enabled = False
End Sub

Private Sub cmdDesconciliar_Click()
'          Dim objDuplicata As CDuplicata
'          Dim i As Integer
'          Dim dblValorSemRateio As Double
'          Dim intNumLanc As Integer
'          Dim blnSemRateio As Boolean
'          Dim intResposta As Integer
          
'        Set objDuplicata = New CDuplicata
'        With grdConciliados
'            For i = 1 To .Rows - 1
'                .Row = i
'                If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
'                   dblValorSemRateio = objDuplicata.ValorSemRateio(.TextMatrix(i, colConcOrigem), .TextMatrix(i, colConcNumero), .TextMatrix(i, colConcEmpresa), IIf(.TextMatrix(i, colConcDebCred) = "Débito", "P", "R"), .TextMatrix(i, colConcTipo), .TextMatrix(i, colConcLiberacao))
'               If .TextMatrix(i, colConcValor) <> dblValorSemRateio Then
'                  intNumLanc = .TextMatrix(i, colConcNumero)
'               Else
'                  blnSemRateio = False
'               End If
'            Else
'                If intNumLanc = .TextMatrix(i, colConcNumero) Then
'                    intResposta = MsgBox("Há mais lançamento(s) conciliado(s) para duplicata/lançamento com número " & intNumLanc & "." & vbNewLine & "Deseja desconciliar todos os lançamentos para o mesmo?", vbYesNoCancel)
'                    If intResposta = vbYes Then
'                       blnSemRateio = True
'                    ElseIf intResposta = vbNo Then
'                       blnSemRateio = False
'                    Else
'                       Exit For
'                    End If
'                End If
'            End If
'        Next
'        End With
    Call ConciliaOuDesconcilia(False, False, grdConciliados, False, True)
End Sub

Private Sub cmdDesconciliarTodos_Click()
    Call ConciliaOuDesconcilia(False, True, grdConciliados, False)
End Sub

Private Sub cmdInserirLancamento_Click()
    Dim i As Integer
    Dim blnTemCheckado As Boolean
    
On Error GoTo err_Handler
    
    For i = 0 To grdExtrato.Rows - 1
        grdExtrato.Row = i
        If grdExtrato.CellPicture = imgCheck.ListImages(grdChecked).Picture Then
            If grdExtrato.TextMatrix(i, 5) = "Débito" Then
                frmLancamentoDuplicata.PagRec = Pagamento
            Else
                frmLancamentoDuplicata.PagRec = Recebimento
            End If
            blnTemCheckado = True
            Exit For
        End If
    Next
    
    If blnTemCheckado Then
        frmLancamentoDuplicata.LancDup = Lancamento
        Call mostrarForm(frmLancamentoDuplicata, frmLancamentoDuplicata.HelpContextID)
        frmLancamentoDuplicata.mblnOrigemTelaConciliacao = True
        mblnOrigemTelaConciliacao = True
        If etxQuantidadeSelecionadaExtrato.valorInteiro > 0 Then
            'If etxQuantidadeSelecionadaExtrato.valorInteiro = 1 Then
            frmLancamentoDuplicata.etxValorOriginal.valorDecimal = etxTotalValorExtrato.valorMoeda
            'End If
            With grdExtrato
                frmLancamentoDuplicata.etxBanco.valorInteiro = lblBancoExtrato.Caption
                frmLancamentoDuplicata.etxEmissao.Data = .TextMatrix(.Row, 2)
                frmLancamentoDuplicata.etxVencimento.Data = .TextMatrix(.Row, 2)
                frmLancamentoDuplicata.etxPagamento.Data = .TextMatrix(.Row, 2)
                frmLancamentoDuplicata.etxLiberacao.Data = .TextMatrix(.Row, 2)
            End With
        End If
    Else
        MsgBox "Favor selecionar um lançamento de extrato bancário.", vbInformation, "Conciliação Bancária Automática"
    End If
err_Handler:

End Sub

Private Sub cmdPesquisar_Click()
    Call PreparaGrid
    Call PreparaGridConciliados
    If tabConciliados.Tab = 0 Then
        Call FiltraLista(False, True)
        etxExtratoBancario_LostFocus
    Else
        Call FiltraLista(True, True)
    End If
    
End Sub
'---------------------------------------------------------------------------------------
'Procedure..: preparaGrid
'Data.......: 06/05/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para realizar a preparação do grid, formatação de cabeçalho e colunas.
'Parametros.:
'Retorno....:
'Protocolo..: 82528
'---------------------------------------------------------------------------------------
Private Sub PreparaGrid()
    Dim intIndex As Integer

    With grdConcTitulo
        .Cols = 15
        .FixedCols = 1
        .Rows = 2
            
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 150
        
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        
        .TextMatrix(0, 2) = ""
        .ColWidth(2) = 300
        
        .TextMatrix(0, 3) = "Liberação"
        .ColWidth(3) = 1000
        .ColAlignment(3) = flexAlignCenterCenter
        
        .TextMatrix(0, 4) = "Descrição"
        .ColWidth(4) = 2000
        .ColAlignment(4) = flexAlignLeftCenter
        
        .TextMatrix(0, 5) = "Valor"
        .ColWidth(5) = 1200
        .ColAlignment(5) = flexAlignRightCenter
        
        .TextMatrix(0, 6) = "Débito/Crédito"
        .ColWidth(6) = 1200
        .ColAlignment(6) = flexAlignLeftCenter
        
        .TextMatrix(0, 7) = "Empresa"
        .ColWidth(7) = 2000
        .ColAlignment(7) = flexAlignLeftCenter
        
        .TextMatrix(0, 8) = "Número"
        .ColWidth(8) = 700
        .ColAlignment(8) = flexAlignRightCenter
        
        .TextMatrix(0, 9) = "Parcela"
        .ColWidth(9) = 700
        .ColAlignment(9) = flexAlignRightCenter
        
        .TextMatrix(0, 10) = "Tipo"
        .ColWidth(10) = 0
        .ColAlignment(10) = flexAlignLeftCenter
        
        .TextMatrix(0, 11) = "Origem"
        .ColWidth(11) = 1200
        .ColAlignment(11) = flexAlignLeftCenter
        
        'Banco escondido
        .TextMatrix(0, 12) = "Banco"
        .ColWidth(12) = 0
        .ColAlignment(12) = flexAlignLeftCenter
        
        'Extrato escondido
        .TextMatrix(0, 13) = "Extrato"
        .ColWidth(13) = 0
        .ColAlignment(13) = flexAlignLeftCenter
        
        'Seq. Extrato escondido
        .TextMatrix(0, 14) = "Seq. Extrato"
        .ColWidth(14) = 0
        .ColAlignment(14) = flexAlignLeftCenter
        
        For intIndex = 0 To .Cols - 1
            .TextMatrix(1, intIndex) = ""
        Next
        .col = colCheck
        .Row = 1
        Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
        
        .col = colFind
        .Row = 1
        Set .CellPicture = imgCheck.ListImages(grdFind).Picture
    End With
End Sub

Private Sub PreparaGridExtrato()
    Dim intIndex As Integer

    With grdExtrato
        .Cols = 8
        .FixedCols = 1
        .Rows = 2
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 150
        
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        
        .TextMatrix(0, 2) = "Data"
        .ColWidth(2) = 1000
        .ColAlignment(2) = flexAlignCenterCenter
        
        .TextMatrix(0, 3) = "Descrição"
        .ColWidth(3) = 2000
        .ColAlignment(3) = flexAlignLeftCenter
        
        .TextMatrix(0, 4) = "Valor"
        .ColWidth(4) = 800
        .ColAlignment(4) = flexAlignRightCenter
        
        .TextMatrix(0, 5) = "Débito/Crédito"
        .ColWidth(5) = 1200
        .ColAlignment(5) = flexAlignLeftCenter
        
        .TextMatrix(0, 6) = "Sequencial"
        .ColWidth(6) = 0
        .ColAlignment(6) = flexAlignLeftCenter
        
        .TextMatrix(0, 7) = "Banco"
        .ColWidth(7) = 0
        .ColAlignment(7) = flexAlignLeftCenter
        
        For intIndex = 0 To .Cols - 1
            .TextMatrix(1, intIndex) = ""
        Next
        .col = colCheck
        .Row = 1
        Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture

    End With
End Sub

Private Sub PreparaGridConciliados()
    Dim intIndex As Integer

    With grdConciliados
        .Cols = 15
        .FixedCols = 1
        .Rows = 2
            
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 150
        
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        
        .TextMatrix(0, 2) = ""
        .ColWidth(2) = 300
        
        .TextMatrix(0, 3) = "Liberação"
        .ColWidth(3) = 1000
        .ColAlignment(3) = flexAlignCenterCenter
        
        .TextMatrix(0, 4) = "Descrição"
        .ColWidth(4) = 2800
        .ColAlignment(4) = flexAlignLeftCenter
        
        .TextMatrix(0, 5) = "Valor"
        .ColWidth(5) = 1200
        .ColAlignment(5) = flexAlignRightCenter
        
        .TextMatrix(0, 6) = "Débito/Crédito"
        .ColWidth(6) = 1200
        .ColAlignment(6) = flexAlignLeftCenter
        
        .TextMatrix(0, 7) = "Empresa"
        .ColWidth(7) = 2000
        .ColAlignment(7) = flexAlignLeftCenter
        
        .TextMatrix(0, 8) = "Número"
        .ColWidth(8) = 700
        .ColAlignment(8) = flexAlignRightCenter
        
        .TextMatrix(0, 9) = "Parcela"
        .ColWidth(9) = 700
        .ColAlignment(9) = flexAlignRightCenter
        
        .TextMatrix(0, 10) = "Tipo"
        .ColWidth(10) = 0
        .ColAlignment(10) = flexAlignLeftCenter
        
        .TextMatrix(0, 11) = "Origem"
        .ColWidth(11) = 1200
        .ColAlignment(11) = flexAlignLeftCenter
        
        'Banco escondido
        .TextMatrix(0, 12) = "Banco"
        .ColWidth(12) = 0
        .ColAlignment(12) = flexAlignLeftCenter
        
        'Extrato escondido
        .TextMatrix(0, 13) = "Extrato"
        .ColWidth(13) = 0
        .ColAlignment(13) = flexAlignLeftCenter
        
        'Seq. Extrato escondido
        .TextMatrix(0, 14) = "Seq. Extrato"
        .ColWidth(14) = 0
        .ColAlignment(14) = flexAlignLeftCenter
        
        For intIndex = 0 To .Cols - 1
            .TextMatrix(1, intIndex) = ""
        Next
        .col = colCheck
        .Row = 1
        Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
        
        .col = colFind
        .Row = 1
        Set .CellPicture = imgCheck.ListImages(grdFind).Picture
    End With
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: FiltraLista
'Data.......: 06/05/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para pegar as informações inseridas pelo usuario e realiza a seleção.
'Parametros.:
'Retorno....:
'Protocolo..: 82528
'---------------------------------------------------------------------------------------
Private Sub FiltraLista(ByVal blnTrazConciliado As Boolean, Optional ByVal blnMsgNaoExisteReg As Boolean)
    Dim strSql      As String
    Dim strSqlD     As String
    Dim strFiltro   As String
    Dim strOrdem    As String
    
    'Função para validar campos.
    If Validacao = False Then
        Exit Sub
    End If
    
    'Limpa variáveis.
    strSql = ""
    strSqlD = ""
    strFiltro = ""
    strOrdem = ""
    etxTotalValor.valorMoeda = 0#
    etxQuantidadeSelecionada.valorInteiro = 0
    If Not mblnOrigemTelaConciliacao Then
        etxTotalValorExtrato.valorMoeda = 0#
        etxQuantidadeSelecionadaExtrato.valorInteiro = 0
    End If
    etxDiferenca.valorMoeda = 0
    lblTipoOperacaoExtrato.Caption = ""
    lblTipoOperacaoLanc.Caption = ""
    DoEvents
    
    'Filtragem
    'Data
    If edtEmissaoInicial.Data <> Trim("00:00:00") Then
        strFiltro = strFiltro & " AND Lançamentos.Emissão>=#" & Format(edtEmissaoInicial.Data, "MM/DD/YYYY") & "# "
    End If
    If edtEmissaoFinal.Data <> Trim("00:00:00") Then
        strFiltro = strFiltro & " AND Lançamentos.Emissão<=#" & Format(edtEmissaoFinal.Data, "MM/DD/YYYY") & "# "
    End If
    If edtPagamentoInicial.Data <> Trim("00:00:00") Then
        strFiltro = strFiltro & " AND Lançamentos.Pagamento>=#" & Format(edtPagamentoInicial.Data, "MM/DD/YYYY") & "# "
    End If
    If edtPagamentoFinal.Data <> Trim("00:00:00") Then
        strFiltro = strFiltro & " AND Lançamentos.Pagamento<=#" & Format(edtPagamentoFinal.Data, "MM/DD/YYYY") & "# "
    End If
    If edtLiberacaoInicial.Data <> Trim("00:00:00") Then
        strFiltro = strFiltro & " AND Lançamentos.Liberação>=#" & Format(edtLiberacaoInicial.Data, "MM/DD/YYYY") & "# "
    End If
    If edtLiberacaoFinal.Data <> Trim("00:00:00") Then
        strFiltro = strFiltro & " AND Lançamentos.Liberação<=#" & Format(edtLiberacaoFinal.Data, "MM/DD/YYYY") & "# "
    End If
    'Combo
    If ecbTipo.SelectedItem <> "" And ecbTipo.SelectedItem <> "Todos" Then
        strFiltro = strFiltro & " AND Lançamentos.Tipo='" & ecbTipo.SelectedItem & "' "
    End If
    'Text
    If etxValorInicial.valorMoeda > 0 Then
        strFiltro = strFiltro & " AND Lançamentos.[Valor Original]>=" & etxValorInicial.valorMoeda & " "
    End If
    If etxValorFinal.valorMoeda > 0 Then
        strFiltro = strFiltro & " AND Lançamentos.[Valor Original]<=" & etxValorFinal.valorMoeda & " "
    End If
    If etxBancoInicial.valorInteiro > 0 Then
        strFiltro = strFiltro & " AND Lançamentos.Banco>=" & etxBancoInicial.valorInteiro & " "
    End If
    If etxBancoFinal.valorInteiro > 0 Then
        strFiltro = strFiltro & " AND Lançamentos.Banco<=" & etxBancoFinal.valorInteiro & " "
    End If
    If etxContaInicial.valorInteiro > 0 Then
        strFiltro = strFiltro & " AND Lançamentos.Conta>=" & etxContaInicial.valorInteiro & " "
    End If
    If etxContaFinal.valorInteiro > 0 Then
        strFiltro = strFiltro & " AND Lançamentos.Conta<=" & etxContaFinal.valorInteiro & " "
    End If
    If etxChequeInicial.valorInteiro > 0 Then
        strFiltro = strFiltro & " AND Lançamentos.Cheque>=" & etxChequeInicial.valorInteiro & " "
    End If
    If etxChequeFinal.valorInteiro > 0 Then
        strFiltro = strFiltro & " AND Lançamentos.Cheque<=" & etxChequeFinal.valorInteiro & " "
    End If
    'Option
    If Not optAmbosPagRec.value Then
        If optPagamento.value = True Then
            strFiltro = strFiltro & " AND Lançamentos.PagRec='P' "
        ElseIf optRecebimento.value = True Then
            strFiltro = strFiltro & " AND Lançamentos.PagRec='R' "
        End If
    End If
    'Ordem
    If optOrdenarLiberacao.value = True Then
        strOrdem = " ORDER BY Liberação "
    End If
    If optOrdenarPagamento.value = True Then
        strOrdem = " ORDER BY Pagamento "
    End If
    'Check
    If blnTrazConciliado Then
        strFiltro = strFiltro & " AND Lançamentos.conciliado= True "
    Else
        strFiltro = strFiltro & " AND Lançamentos.conciliado= False "
    End If
    
    'Seleção da tabela de lançamentos, com a condição de pagamentos terem sidos realizados e o campo conciliado igual a falso.
    strSql = "SELECT Lançamentos.[PagRec] ,Lançamentos.[Abatimento] ,Lançamentos.[Acréscimo] ,Lançamentos.[Código], Lançamentos.[Parcela], Lançamentos.[Empresa], Lançamentos.[Tipo], Lançamentos.[Descrição], Lançamentos.[Emissão], Lançamentos.[Vencimento], Lançamentos.[Pagamento], Lançamentos.[Liberação], Lançamentos.[Valor Original], Lançamentos.[Banco], Lançamentos.[Conta], Lançamentos.[Centro], Lançamentos.[Cheque],  Lançamentos.[Controle], Lançamentos.[Situação], Lançamentos.[Alteração], Lançamentos.[Conciliado], 'Lançamentos' as L_D, Lançamentos.[conciliacao_banco], Lançamentos.[conciliacao_extrato], Lançamentos.[conciliacao_sequencial_extrato] "
    strSql = strSql & "FROM Lançamentos WHERE not isnull(Lançamentos.pagamento) "
    strSql = strSql & strFiltro
    
    'Referente a tabela "Lançamentos" e/ou "Duplicatas"
    If optAmbos.value = True Then
        strSqlD = Replace(Replace(strSql, "Lançamentos", "Duplicatas"), "Duplicatas.[Código]", "Duplicatas.[Nota]")
        strSql = strSql & " UNION " & strSqlD & strOrdem
    ElseIf optLancamentos.value = True Then
        strSql = strSql & strOrdem
    ElseIf optDuplicatas.value = True Then
        'Se for escolhido "Duplicatas" é realizado um replace no nome da tabela, já que os campos são os mesmo.
        strSqlD = Replace(Replace(strSql, "Lançamentos", "Duplicatas"), "Duplicatas.[Código]", "Duplicatas.[Nota] as Código")
        strSql = strSqlD & strOrdem
    End If
    'Chama a função para popular o grid.
    AtualizaLista strSql, IIf(blnTrazConciliado, grdConciliados, grdConcTitulo), blnTrazConciliado, blnMsgNaoExisteReg
End Sub
'---------------------------------------------------------------------------------------
'Procedure..: AtualizaLista
'Data.......: 06/05/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para popular o grid com as informações da seleção.
'Parametros.:
'Retorno....:
'Protocolo..: 82528
'---------------------------------------------------------------------------------------
Private Sub AtualizaLista(strSql As String, grdGenerica As MSHFlexGrid, Optional blnTrazConciliado As Boolean, Optional blnMsgNaoExisteReg As Boolean)
    Dim rsConciliacao  As Object
    Dim i              As Integer
    
    mlngTListado = 0
    mdblTValor = 0
    If AbreRecordset(rsConciliacao, strSql) = WL_OK Then
        rsConciliacao.MoveFirst
        i = 1
        While Not rsConciliacao.EOF
            grdGenerica.AddItem ("")
            grdGenerica.col = colCheck
            grdGenerica.Row = grdGenerica.Rows - 1
            Set grdGenerica.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            grdGenerica.col = colFind
            Set grdGenerica.CellPicture = imgCheck.ListImages(grdFind).Picture
            grdGenerica.TextMatrix(i, 3) = GetValue(rsConciliacao, "Liberação", "")
            grdGenerica.TextMatrix(i, 4) = GetValue(rsConciliacao, "Descrição", "")
            grdGenerica.TextMatrix(i, 5) = Format(GetValue(rsConciliacao, "Valor Original", ZERO) - GetValue(rsConciliacao, "Abatimento", ZERO) + GetValue(rsConciliacao, "Acréscimo", ZERO), "###,0.00")
            grdGenerica.TextMatrix(i, 6) = IIf(GetValue(rsConciliacao, "PagRec", "") = "R", "Crédito", "Débito")
            grdGenerica.TextMatrix(i, 7) = GetValue(rsConciliacao, "Empresa", "")
            grdGenerica.TextMatrix(i, 8) = GetValue(rsConciliacao, "Código", "")
            grdGenerica.TextMatrix(i, 9) = GetValue(rsConciliacao, "Parcela", "")
            grdGenerica.TextMatrix(i, 10) = GetValue(rsConciliacao, "Tipo", "")
            grdGenerica.TextMatrix(i, 11) = GetValue(rsConciliacao, "L_D", "")
            grdGenerica.TextMatrix(i, 12) = GetValue(rsConciliacao, "Banco", "")
            grdGenerica.TextMatrix(i, 13) = GetValue(rsConciliacao, "conciliacao_extrato", "")
            grdGenerica.TextMatrix(i, 14) = GetValue(rsConciliacao, "conciliacao_sequencial_extrato", "")
            mlngTListado = mlngTListado + 1
            i = i + 1
            rsConciliacao.MoveNext
        Wend
        If grdGenerica.Rows > 2 Then
            grdGenerica.RemoveItem (grdGenerica.Rows - 1)
        End If
    
    Else
        If Not blnTrazConciliado Then
            If blnMsgNaoExisteReg Then
                MsgBox "Não há registros à conciliar.", vbInformation, NomeModulo
            End If
        End If
    End If
End Sub

Private Function ExisteRegSelecionado() As Boolean
    Dim i As Integer
    
    With grdConcTitulo
        For i = 1 To .Rows - 1
            .Row = i
            .col = colCheck
            If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                ExisteRegSelecionado = True
                Exit For
            End If
        Next
    End With
End Function
Private Sub cmdSair_Click()
    Unload Me
End Sub
'---------------------------------------------------------------------------------------
'Procedure..: cmdSelecionaNenhum_Click
'Data.......: 06/05/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para tirar a seleção dos itens no grid.
'Parametros.:
'Retorno....:
'Protocolo..: 82528
'---------------------------------------------------------------------------------------
Private Sub cmdSelecionaNenhum_Click()
    Dim intIndex           As Integer
    etxQuantidadeSelecionada.valorInteiro = 0
    etxTotalValor.valorMoeda = 0#
    With grdConcTitulo
        For intIndex = 1 To .Rows - 1
            .Row = intIndex
            .col = colCheck
            Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
        Next
    End With
    etxDiferenca.valorMoeda = (etxTotalValor.valorMoeda - etxTotalValorExtrato.valorMoeda)
    cmdConciliar.Enabled = False
    lblTipoOperacaoLanc.Caption = ""
End Sub

Private Sub cmdSelecionaNenhumExtrato_Click()
    Dim intIndex           As Integer
    etxQuantidadeSelecionadaExtrato.valorInteiro = 0
    etxTotalValorExtrato.valorMoeda = 0#
    With grdExtrato
        For intIndex = 1 To .Rows - 1
            .Row = intIndex
            .col = colCheck
            Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
        Next
    End With
    etxDiferenca.valorMoeda = (etxTotalValor.valorMoeda - etxTotalValorExtrato.valorMoeda)
    cmdConciliar.Enabled = False
    lblTipoOperacaoExtrato.Caption = ""
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: cmdSelecionaTodos_Click
'Data.......: 06/05/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para selecionar todos os itens do grid.
'Parametros.:
'Retorno....:
'Protocolo..: 82528
'---------------------------------------------------------------------------------------
Private Sub cmdSelecionaTodos_Click()
    Dim intIndex            As Integer
        
    If (etxQuantidadeSelecionadaExtrato.valorInteiro <= 1) And ((etxQuantidadeSelecionada.valorInteiro > 1 And etxQuantidadeSelecionadaExtrato.valorInteiro = 1) Or etxQuantidadeSelecionada.valorInteiro = 0 Or etxQuantidadeSelecionadaExtrato.valorInteiro = 0 Or (etxQuantidadeSelecionadaExtrato.valorInteiro = 1 And etxQuantidadeSelecionada.valorInteiro = 1)) Then
        etxQuantidadeSelecionada.valorInteiro = 0
        With grdConcTitulo
            If .TextMatrix(1, 3) <> "" Then
                If optAmbosPagRec.value = False Then
                    For intIndex = 1 To .Rows - 1
                        .Row = intIndex
                        .col = colCheck
                        Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                        lblTipoOperacaoLanc.Caption = .TextMatrix(intIndex, colConcDebCred)
                        etxQuantidadeSelecionada.valorInteiro = etxQuantidadeSelecionada.valorInteiro + 1
                        etxTotalValor.valorMoeda = etxTotalValor.valorMoeda + grdConcTitulo.TextMatrix(grdConcTitulo.Row, 5)
                        etxDiferenca.valorMoeda = (etxTotalValor.valorMoeda - etxTotalValorExtrato.valorMoeda)
                    Next
                End If
            End If
        End With
        If etxDiferenca.valorMoeda = "0,00" And etxQuantidadeSelecionada.valorInteiro > 0 And etxQuantidadeSelecionadaExtrato.valorInteiro > 0 Then
            cmdConciliar.Enabled = True
        Else
            cmdConciliar.Enabled = False
        End If
    Else
       MsgBox "Não é possível vincular vários lançamentos/duplicatas a mais de um lançamento de extrato bancário.", vbInformation, "Conciliação Bancária"
    End If
End Sub

Private Sub cmdSelecionaTodosExtrato_Click()
    Dim intIndex            As Integer
    If (etxQuantidadeSelecionada.valorInteiro <= 1) And ((etxQuantidadeSelecionada.valorInteiro = 1 And etxQuantidadeSelecionadaExtrato.valorInteiro > 1) Or etxQuantidadeSelecionada.valorInteiro = 0 Or etxQuantidadeSelecionadaExtrato.valorInteiro = 0 Or (etxQuantidadeSelecionadaExtrato.valorInteiro = 1 And etxQuantidadeSelecionada.valorInteiro = 1)) Then
        etxQuantidadeSelecionadaExtrato.valorInteiro = 0
        With grdExtrato
            If .TextMatrix(1, 2) <> "" Then
                If optAmbosPagRec.value = False Then
                    For intIndex = 1 To .Rows - 1
                        .Row = intIndex
                        .col = colCheck
                        Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                        lblTipoOperacaoExtrato.Caption = .TextMatrix(intIndex, 5)
                        etxQuantidadeSelecionadaExtrato.valorInteiro = etxQuantidadeSelecionadaExtrato.valorInteiro + 1
                        etxTotalValorExtrato.valorMoeda = etxTotalValorExtrato.valorMoeda + grdExtrato.TextMatrix(grdExtrato.Row, 4)
                        etxDiferenca.valorMoeda = (etxTotalValor.valorMoeda - etxTotalValorExtrato.valorMoeda)
                    Next
                End If
            End If
        End With
        
        If etxDiferenca.valorMoeda = "0,00" And etxQuantidadeSelecionada.valorInteiro > 0 And etxQuantidadeSelecionadaExtrato.valorInteiro > 0 Then
            cmdConciliar.Enabled = True
        Else
            cmdConciliar.Enabled = False
        End If
    Else
        MsgBox "Não é possível vincular vários lançamentos de extrato bancário a mais de um lançamento/duplicata", vbInformation, "Conciliação Bancária"
    End If
End Sub

Private Sub cmdImportaExtrato_Click()
    
On Error GoTo err_Handler
    Call mostrarForm(frmImpDigExtratoBancario, frmImpDigExtratoBancario.HelpContextID)
    frmImpDigExtratoBancario.lblOrigemConciliacao.Caption = "1"
err_Handler:
    'MsgBox err.Description
End Sub

Private Sub etxBancoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        PCampo "Bancos", "Select * from [Bancos]", pbCampo, etxBancoFinal, "Banco"
    End If
End Sub

Private Sub etxBancoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        PCampo "Bancos", "Select * from [Bancos]", pbCampo, etxBancoInicial, "Banco"
    End If
End Sub

Private Sub etxContaFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        PCampo "Bancos", "SELECT Contas.Código as Conta, Contas.Descrição as [Descrição da Conta], Grupos.Código as Grupo, Grupos.Descrição as [Descrição do Grupo] " & _
                       " FROM Grupos INNER JOIN Contas ON Grupos.Código = Contas.Grupo where Contas.Ctaati='S' " & _
                       " ORDER BY Grupos.Código,Contas.Código", pbCampo, etxContaFinal, "Conta"
    End If
End Sub

Private Sub etxContaInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        PCampo "Conta", "SELECT Contas.Código as Conta, Contas.Descrição as [Descrição da Conta], Grupos.Código as Grupo, Grupos.Descrição as [Descrição do Grupo] " & _
                       " FROM Grupos INNER JOIN Contas ON Grupos.Código = Contas.Grupo where Contas.Ctaati='S' " & _
                       " ORDER BY Grupos.Código,Contas.Código", pbCampo, etxContaInicial, "Conta"
    End If
End Sub

Private Sub etxExtratoBancario_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    If KeyCode = vbKeyPageDown Then
        'If etxExtratoBancario.valorInteiro <> 0 Then
        
        
        strSql = ""
        strSql = strSql & "SELECT DISTINCT eb.cd_extrato, eb.cd_banco "
        strSql = strSql & "FROM   FFIExtratoBancario eb, FFIExtratoBancarioHistorico ebh "
        strSql = strSql & "WHERE  eb.cd_historico = ebh.cd_historico and eb.conciliado = 0 "
        strSql = strSql & "ORDER BY eb.cd_extrato, eb.cd_banco "

        Call PMultiCampo("Extrato Bancário", strSql, pbCampo, "cd_extrato", etxExtratoBancario)
        'End If
        CarregaGridExtrato
    End If
End Sub

Private Sub etxExtratoBancario_LostFocus()
    CarregaGridExtrato
End Sub

Private Sub etxTotalValor_Change()
    etxDiferenca.valorMoeda = (etxTotalValor.valorMoeda - etxTotalValorExtrato.valorMoeda)
End Sub

Private Sub etxTotalValorExtrato_Change()
    etxDiferenca.valorMoeda = (etxTotalValor.valorMoeda - etxTotalValorExtrato.valorMoeda)
End Sub

Private Sub Form_Load()
    Dim strMsg            As String
    Dim intIdForm         As Integer
    'Dim intDiasFaltantes  As Integer
    'Dim blnPrimeiroAcesso As Boolean
    
    intIdForm = 3016
    
    If optAmbosPagRec.value = True Then
        cmdSelecionaTodosExtrato.Enabled = False
        cmdSelecionaTodos.Enabled = False
    Else
        cmdSelecionaTodosExtrato.Enabled = True
        cmdSelecionaTodos.Enabled = True
    End If
    
    'Valida permissão botão importar extrato
    cmdImportaExtrato.Enabled = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 3014)
    
    'If ValidaPrimeiroAcessoRotina(intIdForm) Then
    '    InserePrimeiroAcessoRotina (intIdForm)
    '    MsgBox "A tela de Conciliação Bancária Automática estará disponível de forma demonstrativa por 60 dias.", vbInformation, "Conciliação Bancária Automática"
    '    blnPrimeiroAcesso = True
    'End If
    'intDiasFaltantes = ValidaDiasAcessoRotina(intIdForm)
    
    'If intDiasFaltantes > 0 Then
    '    If Not blnPrimeiroAcesso Then
            'Se estiver faltando mais de 10 dias
    '        If intDiasFaltantes >= 10 Then
    '            If ValidaAlertaAcessoRotina(intIdForm) Then
    '                If MsgBox("Você ainda tem " & intDiasFaltantes & " dia(s) para utilizar esta rotina de forma demonstrativa. " & vbNewLine & "Deseja desativar este alerta?", vbYesNo, "Conciliação Bancária Automática") = vbYes Then
    '                    Call AtualizaAlertaAcessoRotina(intIdForm)
    '                End If
    '            End If
    '        'Se estiver nos últimos 10 dias sempre mostra o alerta
    '        Else
    '            MsgBox "Você ainda tem " & intDiasFaltantes & " dia(s) para utilizar esta rotina de forma demonstrativa. ", vbInformation, "Conciliação Bancária Automática"
    '        End If
    '    End If
        
    '    mblnTemAcessoTela = True
        CenterForm Me
        Call etxEmpUser.AddConexao(Aplicacao)
        etxEmpUser.valorTexto = DonaSistema
        Call etxValorInicial.AddConexao(Aplicacao)
        Call etxValorFinal.AddConexao(Aplicacao)
        Call etxBancoInicial.AddConexao(Aplicacao)
        Call etxBancoFinal.AddConexao(Aplicacao)
        Call etxContaInicial.AddConexao(Aplicacao)
        Call etxContaFinal.AddConexao(Aplicacao)
        Call etxChequeInicial.AddConexao(Aplicacao)
        Call etxChequeFinal.AddConexao(Aplicacao)
        mlngTListado = 0
        mdblTValor = 0
        Call PreparaGrid
        Call PreparaGridExtrato
        Call PreparaGridConciliados
        CarregaCombo
        cmdConciliar.Enabled = False
        ecbTipo.SelectItem "Todos"
     'Else
     '   MsgBox "O seu acesso a esta rotina de forma demonstrativa expirou. " & vbNewLine & "Para utilizar novamente entre em contato com o Suporte.", vbInformation, "Conciliação Bancária Automática"
     '   mblnTemAcessoTela = False
     'End If
End Sub


Private Sub grdConciliados_Click()
    On Error GoTo err
        With grdConciliados
            If .col <> 1 Then
                Exit Sub
            End If
            If .col = colCheck Then
                .CellPictureAlignment = flexAlignCenterCenter
                If LinhaSelecionada(.Row, grdConciliados) Then
                    'etxQuantidadeSelecionadaExtrato.valorInteiro = etxQuantidadeSelecionadaExtrato.valorInteiro - 1
                    'etxTotalValorExtrato.valorMoeda = etxTotalValorExtrato.valorMoeda - grdExtrato.TextMatrix(grdExtrato.Row, 4)
                    Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                Else
                    'etxQuantidadeSelecionadaExtrato.valorInteiro = etxQuantidadeSelecionadaExtrato.valorInteiro + 1
                    'etxTotalValorExtrato.valorMoeda = etxTotalValorExtrato.valorMoeda + grdExtrato.TextMatrix(grdExtrato.Row, 4)
                    Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                End If
            End If
        End With
    Exit Sub
err:
End Sub

Private Sub grdConciliados_DblClick()
    Dim blnPodeAcessar As Boolean
    With grdConciliados
        If .col = 2 Then
            'Valida permissões
            If .TextMatrix(.Row, colConcDebCred) <> "" Then
                If .TextMatrix(.Row, colConcDebCred) = "Débito" Then
                    If .TextMatrix(.Row, colConcTipo) = "Lançamentos" Then
                        blnPodeAcessar = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2061)
                        frmLancamentoDuplicata.LancDup = Lancamento
                        frmLancamentoDuplicata.PagRec = Pagamento
                    Else
                        blnPodeAcessar = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2062)
                        frmLancamentoDuplicata.LancDup = Duplicata
                        frmLancamentoDuplicata.PagRec = Pagamento
                    End If
                ElseIf .TextMatrix(.Row, colConcDebCred) = "Crédito" Then
                    If .TextMatrix(.Row, colConcTipo) = "Lancamentos" Then
                        blnPodeAcessar = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2057)
                        frmLancamentoDuplicata.LancDup = Lancamento
                        frmLancamentoDuplicata.PagRec = Recebimento
                    Else
                        blnPodeAcessar = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2058)
                        frmLancamentoDuplicata.LancDup = Duplicata
                        frmLancamentoDuplicata.PagRec = Recebimento
                    End If
                End If
                
                If blnPodeAcessar Then
                    Call mostrarForm(frmLancamentoDuplicata, frmLancamentoDuplicata.HelpContextID)
                    frmLancamentoDuplicata.mblnOrigemTelaConciliacao = True
                    mblnOrigemTelaConciliacao = True
                    Call frmLancamentoDuplicata.CarregarLancamentoDuplicataOutrasRotinas(.TextMatrix(.Row, 8), .TextMatrix(.Row, 10), .TextMatrix(.Row, 9), .TextMatrix(.Row, 7), IIf(.TextMatrix(.Row, 6) = "Débito", 0, 1), IIf(.TextMatrix(.Row, 11) = "Duplicatas", 1, 0))
                    Exit Sub
                Else
                    MsgBox "Usuário sem permissão de acesso a rotina.", vbInformation, "Lançamento/Duplicata"
                End If
            Else
                MsgBox "Não há lançamento/duplicata para visualizar", vbInformation, "Visualização de Lançamento/Duplicata"
            End If
        End If
    End With
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: grdConcTitulo_Click
'Data.......: 06/05/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para marcar e desmarcar os itens dentro do grid(função do grid).
'Parametros.:
'Retorno....:
'Protocolo..: 82528
'---------------------------------------------------------------------------------------
Private Sub grdConcTitulo_Click()
    Dim intLinhaClique As Integer
    Dim i As Integer
    Dim intContCheckado As Integer
    
    On Error GoTo err
        With grdConcTitulo
        
        'Guarda e valida tipo de operação selecionado
        intLinhaClique = .Row
        intContCheckado = 0
        
        'Verifica se existe algum checkado
        For i = 1 To .Rows - 1
            .Row = i
            If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                intContCheckado = intContCheckado + 1
            End If
        Next
        If intContCheckado = 0 Then
            lblTipoOperacaoLanc.Caption = ""
        End If
        'se não existe guarda o tipo da operação
        If intContCheckado = 0 And lblTipoOperacaoLanc.Caption <> .TextMatrix(intLinhaClique, colConcDebCred) Then
            lblTipoOperacaoLanc.Caption = .TextMatrix(intLinhaClique, colConcDebCred)
        Else
            'Valida tipo de operação igual
            If lblTipoOperacaoLanc.Caption <> .TextMatrix(intLinhaClique, colConcDebCred) Then
               MsgBox "Não é possível marcar lançamento com tipo de operação diferente (Débito/Crédito).", vbInformation, "Conciliação Bancária Automática"
               Exit Sub
            End If
        End If
    
        .Row = intLinhaClique
        
        If lblTipoOperacaoExtrato.Caption <> "" And lblTipoOperacaoExtrato.Caption <> lblTipoOperacaoLanc.Caption And .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
           MsgBox "Não é possível marcar um lançamento com tipo de operação diferente do tipo de operação do extrato bancário (Débito/Crédito).", vbInformation, "Conciliação Bancária Automática"
           Exit Sub
        End If
        
        If lblTipoOperacaoExtrato.Caption <> "" And (lblTipoOperacaoExtrato.Caption <> lblTipoOperacaoLanc.Caption) Then
            MsgBox "Não é possível marcar lançamento(s)/duplicata(s) com tipo de operação diferente ao do extrato bancário (Débito/Crédito).", vbInformation, "Conciliação Bancária Automática"
            If .CellPicture = imgCheck.ListImages(grdUnchecked).Picture Then
                lblTipoOperacaoLanc.Caption = IIf(lblTipoOperacaoLanc.Caption = "Débito", "Crédito", "Débito")
            End If
            Exit Sub
        End If
        
        If .CellPicture = imgCheck.ListImages(grdChecked).Picture Or (etxQuantidadeSelecionada.valorInteiro > 1 And etxQuantidadeSelecionadaExtrato.valorInteiro = 1) Or etxQuantidadeSelecionada.valorInteiro = 0 Or etxQuantidadeSelecionadaExtrato.valorInteiro = 0 Or (etxQuantidadeSelecionadaExtrato.valorInteiro = 1 And etxQuantidadeSelecionada.valorInteiro = 1) Then
            If .col <> 1 Then
                Exit Sub
            End If
            
            If .TextMatrix(.Row, 12) = lblBancoExtrato.Caption Or lblBancoExtrato.Caption = "" Then
            Else
                If .TextMatrix(.Row, 12) <> "" And .CellPicture = imgCheck.ListImages(grdUnchecked).Picture Then
                    If MsgBox("Tem certeza que deseja fazer esta conciliação?" & vbNewLine & "O banco do lançamento/duplicata (" & .TextMatrix(.Row, 12) & ") selecionado não corresponde ao banco do extrato (" & lblBancoExtrato.Caption & ").", vbYesNo, "Conciliação Bancária") = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            If .col = colCheck Then
                .CellPictureAlignment = flexAlignCenterCenter
                If LinhaSelecionada(.Row, grdConcTitulo) Then
                    etxQuantidadeSelecionada.valorInteiro = etxQuantidadeSelecionada.valorInteiro - 1
                    etxTotalValor.valorMoeda = etxTotalValor.valorMoeda - grdConcTitulo.TextMatrix(grdConcTitulo.Row, 5)
                    Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                    If intContCheckado = 1 Then
                        lblTipoOperacaoLanc.Caption = ""
                    End If
                Else
                    etxQuantidadeSelecionada.valorInteiro = etxQuantidadeSelecionada.valorInteiro + 1
                    etxTotalValor.valorMoeda = etxTotalValor.valorMoeda + grdConcTitulo.TextMatrix(grdConcTitulo.Row, 5)
                    Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                End If
                etxDiferenca.valorMoeda = (etxTotalValor.valorMoeda - etxTotalValorExtrato.valorMoeda)
                If etxDiferenca.valorMoeda = "0,00" Then
                    cmdConciliar.Enabled = True
                Else
                    cmdConciliar.Enabled = False
                End If
            End If
            If etxQuantidadeSelecionadaExtrato.valorInteiro = 0 Or etxQuantidadeSelecionada.valorInteiro = 0 Then
                cmdConciliar.Enabled = False
            End If
        Else
            MsgBox "Não é possível vincular vários lançamentos/duplicatas a mais de um lançamento de extrato bancário.", vbInformation, "Conciliação Bancária"
        End If
    End With
    Exit Sub
err:
End Sub

Private Function LinhaSelecionada(lngLinha As Long, grdGenerica As MSHFlexGrid) As Boolean
    If lngLinha <= grdGenerica.Rows - 1 Then
        grdGenerica.Row = lngLinha
        grdGenerica.col = colCheck
        LinhaSelecionada = (grdGenerica.CellPicture = imgCheck.ListImages(2).Picture)
    Else
        LinhaSelecionada = False
    End If
End Function

'---------------------------------------------------------------------------------------
'Procedure..: Validacao
'Data.......: 06/05/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para validar os campos do formulário, verifica se a data final é menor que a data inicial.
'Parametros.:
'Retorno....: [Boolean] Valor referente se a validação esta correta.
'Protocolo..: 82528
'---------------------------------------------------------------------------------------
Private Function Validacao() As Boolean
    Dim strValida As String
    
    strValida = ""
    'Data Emissão.
    If edtEmissaoInicial.Data <> Trim("00:00:00") And edtEmissaoFinal.Data <> Trim("00:00:00") Then
        If DateDiff("d", edtEmissaoInicial.Data, edtEmissaoFinal.Data) < 0 Then
            strValida = strValida & "Data de emissão inicial é maior que a data de emissão final." & vbCrLf
        End If
    End If
    'Data Pagamento.
    If edtPagamentoInicial.Data <> Trim("00:00:00") And edtPagamentoFinal.Data <> Trim("00:00:00") Then
        If DateDiff("d", edtPagamentoInicial.Data, edtPagamentoFinal.Data) < 0 Then
            strValida = strValida & "Data de pagamento inicial é maior que a data de pagamento final." & vbCrLf
        End If
    End If
    'Data Liberação.
    If edtLiberacaoInicial.Data <> Trim("00:00:00") And edtLiberacaoFinal.Data <> Trim("00:00:00") Then
        If DateDiff("d", edtLiberacaoInicial.Data, edtLiberacaoFinal.Data) < 0 Then
            strValida = strValida & "Data de liberação inicial é maior que a data de liberação final." & vbCrLf
        End If
    End If
    If Trim(strValida) <> "" Then
        MsgBox strValida, vbInformation, NomeModulo
        Validacao = False
    Else
        Validacao = True
    End If
End Function

Private Sub CarregaCombo()
    Dim rstTipo As Object
    Dim strSql  As String
    
    ecbTipo.Clear
    strSql = ""
    strSql = "Select Tipo from [Tipos Globais]"
    AbreRecordset rstTipo, strSql
    If rstTipo.Recordcount > 0 Then
        rstTipo.MoveFirst
        While Not rstTipo.EOF
            ecbTipo.AddItem rstTipo!Tipo
            rstTipo.MoveNext
        Wend
    End If
    ecbTipo.AddItem "Todos"
    FechaRecordset (rstTipo)
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
Public Sub CarregaGridExtrato()
    etxQuantidadeSelecionadaExtrato.valorInteiro = 0
    etxTotalValorExtrato.valorMoeda = 0
    
    Dim dtaFixa    As Date
    Dim strOptTipo As String
    
    dtaFixa = "01/01/1899"
    If optAmbosPagRec.value Then
        strOptTipo = "A"
    ElseIf optPagamento Then
        strOptTipo = "P"
    Else
        strOptTipo = "R"
    End If
    
    Set mbizExtratoBanc = New BizImpDigExtratoBancario
    Set mcolLancamentos = mbizExtratoBanc.CarregarColecao(0, dtaFixa, etxExtratoBancario.valorInteiro, , strOptTipo)
    If mcolLancamentos.Count > 0 Then
        lblBancoExtrato.Caption = mcolLancamentos.CurrentObject.CdBanco
    'Else
        'MsgBox "O código informado é inválido", vbInformation, "Atenção"
    End If
    If Not mcolLancamentos Is Nothing Then
        Call CarregaGrid
    End If
End Sub
Private Sub CarregaGrid()
    Dim objVO        As VoImpDigExtratoBancario
    Dim daoExtrato   As DaoImpDigExtratoBancario
    Dim strItem      As String
    Dim i            As Integer
    Dim strHistorico As String

On Error GoTo erro
    grdExtrato.Clear
    'data desc valor operacao
    PreparaGridExtrato
    If Not mcolLancamentos Is Nothing Then
        If mcolLancamentos.Count > 0 Then
            grdExtrato.Rows = 1
            mcolLancamentos.MoveFirst
            etxExtratoBancario.valorInteiro = mcolLancamentos.CurrentObject.CdExtrato
            While Not mcolLancamentos.EOF
                Set objVO = mcolLancamentos.CurrentObject
                Set daoExtrato = New DaoImpDigExtratoBancario
                strHistorico = daoExtrato.BuscaDescricaoHistorico(objVO.CdBanco, objVO.CdHistorico)
                With objVO
                    strItem = "" & vbTab & "" & vbTab & Format(.DataExtrato, "dd/mm/yyyy") & vbTab & strHistorico & vbTab & Format(.Valor, "##,##0.00") & vbTab & IIf(.TipoOperacao = "D", "Débito", "Crédito") & vbTab & .SeqLancExtrato & vbTab & .CdBanco
                    grdExtrato.AddItem strItem
                    grdExtrato.col = colCheck
                    grdExtrato.Row = grdExtrato.Rows - 1
                    Set grdExtrato.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                End With
                Set objVO = Nothing
                mcolLancamentos.MoveNext
            Wend
        End If
    End If
    grdExtrato.FixedRows = 1
    grdExtrato.Sort = flexSortNumericAscending
    Exit Sub
erro:
    MsgBox "Erro ao carregar tabela: " & err.Description
End Sub


Private Sub grdConcTitulo_DblClick()
    Dim blnPodeAcessar As Boolean
    
    With grdConcTitulo
        If .col = 2 Then
            'Valida permissões
            If .TextMatrix(.Row, colConcDebCred) <> "" Then
                If .TextMatrix(.Row, colConcDebCred) = "Débito" Then
                    If .TextMatrix(.Row, colConcTipo) = "Lançamentos" Then
                        blnPodeAcessar = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2061)
                        frmLancamentoDuplicata.LancDup = Lancamento
                        frmLancamentoDuplicata.PagRec = Pagamento
                    Else
                        blnPodeAcessar = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2062)
                        frmLancamentoDuplicata.LancDup = Duplicata
                        frmLancamentoDuplicata.PagRec = Pagamento
                    End If
                ElseIf .TextMatrix(.Row, colConcDebCred) = "Crédito" Then
                    If .TextMatrix(.Row, colConcTipo) = "Lancamentos" Then
                        blnPodeAcessar = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2057)
                        frmLancamentoDuplicata.LancDup = Lancamento
                        frmLancamentoDuplicata.PagRec = Recebimento
                    Else
                        blnPodeAcessar = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2058)
                        frmLancamentoDuplicata.LancDup = Duplicata
                        frmLancamentoDuplicata.PagRec = Recebimento
                    End If
                End If
                            
                If blnPodeAcessar Then
                    Call mostrarForm(frmLancamentoDuplicata, frmLancamentoDuplicata.HelpContextID)
                    Call frmLancamentoDuplicata.CarregarLancamentoDuplicataOutrasRotinas(.TextMatrix(.Row, 8), .TextMatrix(.Row, 10), .TextMatrix(.Row, 9), .TextMatrix(.Row, 7), IIf(.TextMatrix(.Row, 6) = "Débito", 0, 1), IIf(.TextMatrix(.Row, 11) = "Duplicatas", 1, 0))
                    frmLancamentoDuplicata.mblnOrigemTelaConciliacao = True
                    mblnOrigemTelaConciliacao = True
                    Exit Sub
                Else
                    MsgBox "Usuário sem permissão de acesso a rotina.", vbInformation, "Lançamento/Duplicata"
                End If
            Else
                MsgBox "Não há lançamento/duplicata para visualizar", vbInformation, "Visualização de Lançamento/Duplicata"
            End If
            
        End If
    End With
End Sub

Private Sub grdExtrato_Click()
    Dim intLinhaClique As Integer
    Dim i As Integer
    Dim intContCheckado As Integer
    
    On Error GoTo err
        With grdExtrato
            'Guarda e valida tipo de operação selecionado
            intLinhaClique = .Row
            intContCheckado = 0
            
            'Valida permissão botão inserir lançamento
            frmLancamentoDuplicata.LancDup = Lancamento
            If .TextMatrix(.Row, 5) = "Débito" Then
                cmdInserirLancamento.Enabled = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2061)
                frmLancamentoDuplicata.PagRec = Pagamento
            Else
                cmdInserirLancamento.Enabled = verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, 2057)
                frmLancamentoDuplicata.PagRec = Recebimento
            End If
            If Not cmdInserirLancamento.Enabled Then
                MsgBox "Usuário sem permissão de acesso a rotina.", vbInformation, "Lançamento/Duplicata"
            End If
            
            'Verifica se existe algum checkado
            For i = 1 To .Rows - 1
                .Row = i
                If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                    intContCheckado = intContCheckado + 1
                End If
            Next
            If intContCheckado = 0 Then
                lblTipoOperacaoExtrato.Caption = ""
            End If
            'se não existe guarda o tipo da operação
            If intContCheckado = 0 And lblTipoOperacaoExtrato.Caption <> .TextMatrix(intLinhaClique, 5) Then
                lblTipoOperacaoExtrato.Caption = .TextMatrix(intLinhaClique, 5)
            Else
                'Valida tipo de operação igual
                If lblTipoOperacaoExtrato.Caption <> .TextMatrix(intLinhaClique, 5) Then
                   MsgBox "Não é possível marcar lançamentos com tipo de operação diferente (Débito/Crédito).", vbInformation, "Conciliação Bancária Automática"
                   Exit Sub
                End If
            End If
            
            .Row = intLinhaClique
            
            If lblTipoOperacaoLanc.Caption <> "" And lblTipoOperacaoExtrato.Caption <> lblTipoOperacaoLanc.Caption Then
                MsgBox "Não é possível marcar lançamentos de extrato bancário com tipo de operação diferente do tipo de operação de lançamento/duplicata (Débito/Crédito).", vbInformation, "Conciliação Bancária Automática"
                If .CellPicture = imgCheck.ListImages(grdUnchecked).Picture Then
                    lblTipoOperacaoExtrato.Caption = IIf(lblTipoOperacaoExtrato.Caption = "Débito", "Crédito", "Débito")
                End If
                Exit Sub
            End If
            
            If .CellPicture = imgCheck.ListImages(grdChecked).Picture Or (etxQuantidadeSelecionada.valorInteiro = 1 And etxQuantidadeSelecionadaExtrato.valorInteiro > 1) Or etxQuantidadeSelecionada.valorInteiro = 0 Or etxQuantidadeSelecionadaExtrato.valorInteiro = 0 Or (etxQuantidadeSelecionadaExtrato.valorInteiro = 1 And etxQuantidadeSelecionada.valorInteiro = 1) Then
                If .col = colCheck Then
                    .CellPictureAlignment = flexAlignCenterCenter
                    If LinhaSelecionada(.Row, grdExtrato) Then
                        etxQuantidadeSelecionadaExtrato.valorInteiro = etxQuantidadeSelecionadaExtrato.valorInteiro - 1
                        etxTotalValorExtrato.valorMoeda = etxTotalValorExtrato.valorMoeda - grdExtrato.TextMatrix(grdExtrato.Row, 4)
                        Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                        If intContCheckado = 1 Then
                            lblTipoOperacaoExtrato.Caption = ""
                        End If
                    Else
                        etxQuantidadeSelecionadaExtrato.valorInteiro = etxQuantidadeSelecionadaExtrato.valorInteiro + 1
                        etxTotalValorExtrato.valorMoeda = etxTotalValorExtrato.valorMoeda + grdExtrato.TextMatrix(grdExtrato.Row, 4)
                        Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                    End If
                    etxDiferenca.valorMoeda = (etxTotalValor.valorMoeda - etxTotalValorExtrato.valorMoeda)
                    If etxDiferenca.valorMoeda = "0,00" Then
                        cmdConciliar.Enabled = True
                    Else
                        cmdConciliar.Enabled = False
                    End If
                End If
            Else
                MsgBox "Não é possível vincular vários lançamentos de extrato bancário a mais de um lançamento/duplicata.", vbInformation, "Conciliação Bancária"
            End If
        End With
        If etxQuantidadeSelecionadaExtrato.valorInteiro = 0 Or etxQuantidadeSelecionada.valorInteiro = 0 Then
            cmdConciliar.Enabled = False
        End If
    Exit Sub
err:
End Sub
Private Sub ConciliaOuDesconcilia(ByVal blnconcilia As Boolean, ByVal blnTodos As Boolean, ByVal grdGenerica As MSHFlexGrid, Optional blnConciliacaoAutomatica As Boolean, Optional blnSemRateio As Boolean)
    Dim i                     As Integer
    Dim j                     As Integer
    Dim k                     As Integer
    Dim strSql                As String
    Dim intSeqExtrato()       As Integer
    Dim strBanco              As String
    Dim blnsucesso            As Boolean
    Dim strTabela             As String
    Dim strSeqExtrato         As String
    Dim blnCredito            As Boolean
    Dim lngSeqExtrato         As Long
    Dim intContVetor          As Integer
    Dim intContExtrato        As Integer
    Dim arrSeqExtrato         As Variant
    Dim intIndex              As Integer
    Dim objDaoDuplic          As CDuplicata
    Dim blnOk                 As Boolean
    Dim objDaoExtrato         As DaoExtratoBancario
    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
    Dim objLogLancDup         As New clsLogLancamentosDuplicatas
    
    'Valida campos
    Call Validacoes(blnconcilia, blnTodos, grdGenerica, blnConciliacaoAutomatica, blnSemRateio)
       
    intContVetor = 1
    intContExtrato = 0
    If Not mblnExitSub Then
        'Verifica se é para conciliar
        If blnconcilia Then
            ReDim intSeqExtrato(TotalLinhasSelecionadas(grdExtrato)) As Integer
            'Percorre o grid de extrato
            For j = 1 To grdExtrato.Rows - 1
                grdExtrato.Row = j
                'Verifica qual linha esta marcada
                If grdExtrato.CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                    intSeqExtrato(intContVetor) = grdExtrato.TextMatrix(j, 6)
                    'Armazena o banco em uma variável
                    If blnconcilia Then
                        strBanco = grdExtrato.TextMatrix(j, 7)
                    End If
                    Set objDaoExtrato = New DaoExtratoBancario
                    blnOk = objDaoExtrato.ConciliaExtrato(blnconcilia, IIf(blnconcilia, strBanco, grdGenerica.TextMatrix(i, 12)), IIf(blnconcilia, etxExtratoBancario.valorInteiro, grdGenerica.TextMatrix(i, 13)), IIf(blnconcilia, intSeqExtrato(intContVetor), grdGenerica.TextMatrix(i, 14)))
                    
                    Set grdExtrato.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                    'Atualiza situação extrato bancário
                    If Not blnOk Then
                        MsgBox "Problema ao " & IIf(blnconcilia, "conciliar", "desconciliar") & " extrato bancário.", vbInformation, "Conciliação Bancária"
                        blnsucesso = False
                        Exit Sub
                    Else
                        blnsucesso = True
                        intContVetor = intContVetor + 1
                        intContExtrato = intContExtrato + 1
                    End If
                End If
            Next
        End If
        
        'Verifica se deu certo o processo anterior ou se é desconciliação
        If blnsucesso Or Not blnconcilia Then
            intContVetor = 1
            If Not blnconcilia Then
                ReDim intSeqExtrato(TotalLinhasSelecionadas(grdGenerica)) As Integer
            End If
            'Percorre grid de lançamentos/duplicatas
            For i = 1 To grdGenerica.Rows - 1
                grdGenerica.Row = i
                'Verifica se esta marcado ou se é pra gerar todos
                If grdGenerica.CellPicture = imgCheck.ListImages(grdChecked).Picture Or (Not blnconcilia And blnTodos) Then
                    strTabela = grdGenerica.TextMatrix(i, 11)
                    If grdGenerica.TextMatrix(i, colConcDebCred) = "Crédito" Then
                        blnCredito = True
                    Else
                        blnCredito = False
                    End If
                    'Se esta conciliando mais de um extrato, concatena os sequenciais do extrato para salvar na duplicata/lancamento
                    If intContExtrato > 1 Then
                        For k = 1 To intContExtrato
                            strSeqExtrato = strSeqExtrato & intSeqExtrato(k) & IIf(k = intContExtrato, "", ";")
                        Next
                    Else
                        If blnconcilia Then
                            If intContExtrato > 1 Then
                                strSeqExtrato = intSeqExtrato(intContVetor)
                            Else
                                strSeqExtrato = intSeqExtrato(intContExtrato)
                            End If
                        Else
                            strSeqExtrato = grdGenerica.TextMatrix(i, colConcSeqExtrato)
                        End If
                    End If
                    
                    Set objDaoDuplic = New CDuplicata
                    blnOk = objDaoDuplic.ConciliaDuplicLanc(blnconcilia, IIf(blnconcilia, grdGenerica.TextMatrix(i, colConcBanco), 0), strTabela, IIf(blnconcilia, strSeqExtrato, "0"), IIf(blnCredito, "R", "P"), grdGenerica.TextMatrix(i, 8), grdGenerica.TextMatrix(i, colConcEmpresa), grdGenerica.TextMatrix(i, colConcTipo), grdGenerica.TextMatrix(i, colConcParcela), grdGenerica.TextMatrix(i, colConcLiberacao), blnSemRateio, IIf(blnconcilia, etxExtratoBancario.valorInteiro, "0"))
                    If blnconcilia And Not blnTodos Then
                        Set grdGenerica.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                    End If
        
                    'Monta e concilia lançamento/duplicata ao extrato
                    If blnOk Then
                        If blnconcilia Then
                            intContVetor = intContVetor + 1
                            blnsucesso = True
                        Else
                            'Caso seja desconciliação, busca os sequenciais relacionados e volta o extrato bancário para desconcilizado
                            arrSeqExtrato = Split(strSeqExtrato, ";")
                            For intIndex = 0 To UBound(arrSeqExtrato)
                                Set objDaoExtrato = New DaoExtratoBancario
                                blnOk = objDaoExtrato.ConciliaExtrato(blnconcilia, IIf(blnconcilia, strBanco, grdGenerica.TextMatrix(i, 12)), IIf(blnconcilia, etxExtratoBancario.valorInteiro, grdGenerica.TextMatrix(i, 13)), IIf(blnconcilia, intSeqExtrato(intContVetor), arrSeqExtrato(intIndex)))
                                If blnOk Then
                                    blnsucesso = True
                                Else
                                    blnsucesso = False
                                End If
                            Next
                        End If
                    Else
                        MsgBox "Problema ao " & IIf(blnconcilia, "conciliar", "desconciliar") & " duplicata/lançamento.", vbInformation, "Conciliação Bancária"
                        Exit Sub
                    End If
                End If
            Next
        End If
        'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
        Set objLogLancDup = Nothing
        
        If blnsucesso And Not blnConciliacaoAutomatica Then
            MsgBox IIf(blnconcilia, "Conciliação", "Desconciliação") & " feita com sucesso.", vbInformation, "Conciliação Bancária"
        End If
        
        If Not blnConciliacaoAutomatica Then
            Call RecarregaGrids(True, False)
            mblnJaRespondeuMsgBox = False
        End If
    End If
End Sub
Public Sub RecarregaGrids(ByVal blnCarregaGridExtrato As Boolean, Optional ByVal blnMsgNaoExisteReg As Boolean)
    
    lblTipoOperacaoLanc.Caption = ""
    Call PreparaGrid
    Call PreparaGridConciliados
    
    If tabConciliados.Tab = 0 Then
        Call FiltraLista(False, blnMsgNaoExisteReg)
    Else
        Call FiltraLista(True, blnMsgNaoExisteReg)
    End If
    
    If blnCarregaGridExtrato Then
        lblTipoOperacaoExtrato.Caption = ""
        Call PreparaGridExtrato
        Call CarregaGridExtrato
        etxTotalValorExtrato.valorMoeda = 0
        etxQuantidadeSelecionadaExtrato.valorInteiro = 0
    End If
    'tabConciliados.Tab = 0
    etxTotalValor.valorMoeda = 0
    etxQuantidadeSelecionada.valorInteiro = 0
End Sub
Private Sub Validacoes(ByVal blnconcilia As Boolean, ByVal blnTodos As Boolean, ByVal grdGenerica As MSHFlexGrid, Optional blnConciliacaoAutomatica As Boolean, Optional blnSemRateio As Boolean)
    Dim i As Integer
    Dim blnTemCheckado As Boolean
    
    If blnconcilia Then
        If Not blnConciliacaoAutomatica And (etxQuantidadeSelecionada.valorInteiro = 0 Or etxQuantidadeSelecionadaExtrato.valorInteiro = 0) Then
            MsgBox "Favor preencher extrato(s) e lançamento(s)/duplicata(s) a ser(em) conciliado(s).", vbInformation, "Conciliação Bancária"
            mblnExitSub = True
            Exit Sub
        ElseIf etxDiferenca.valorMoeda <> "0,00" Then
            mblnExitSub = True
            MsgBox "Não foi possível fazer a conciliação. " & vbNewLine & "O valor total de extrato (R$ " & Format(etxTotalValorExtrato.valorMoeda, "##,0.00") & ") não é igual ao valor total do(s) lançamento(s)/duplicata(s) (R$ " & Format(etxTotalValor.valorMoeda, "##,0.00") & ") selecionado(s).", vbInformation, "Conciliação Bancária"
            Exit Sub
        End If
    Else
        If Not blnTodos Then
            For i = 0 To grdGenerica.Rows - 1
                grdGenerica.Row = i
                If grdGenerica.CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                    blnTemCheckado = True
                    Exit For
                End If
            Next
        Else
            blnTemCheckado = True
        End If
        'If grdConciliados.TextMatrix(1, 4) = "" And grdConciliados.TextMatrix(1, 5) = "" Then
        If Not blnTemCheckado Then
            mblnExitSub = True
            Exit Sub
        ElseIf MsgBox("Confirma a desconciliação dos lançamentos/duplicatas?", vbYesNo, "Desconciliação Bancária") = vbYes Then
            mblnJaRespondeuMsgBox = True
            mblnExitSub = False
        Else
            mblnExitSub = True
            Exit Sub
        End If
    End If
    
    If Not mblnJaRespondeuMsgBox And Not blnConciliacaoAutomatica Then
        If MsgBox("Confirma a " & IIf(blnconcilia, "conciliação", "desconciliação") & " bancária para " & IIf(blnTodos, "todos ", "") & "o(s) título(s) " & IIf(blnTodos, "", "selecionado(s)") & "?", vbYesNo) = vbNo Then
            mblnExitSub = True
            Exit Sub
        Else
            mblnExitSub = False
        End If
    End If
End Sub


Private Sub optAmbosPagRec_Click()
    If optAmbosPagRec.value = True Then
        cmdSelecionaTodosExtrato.Enabled = False
        cmdSelecionaTodos.Enabled = False
    End If
    cmdConciliar.Enabled = False
End Sub

Private Sub optPagamento_Click()
    If optPagamento.value Then
        cmdSelecionaTodosExtrato.Enabled = True
        cmdSelecionaTodos.Enabled = True
    End If
    cmdConciliar.Enabled = False
End Sub

Private Sub optRecebimento_Click()
    If optRecebimento.value Then
        cmdSelecionaTodosExtrato.Enabled = True
        cmdSelecionaTodos.Enabled = True
    End If
    cmdConciliar.Enabled = False
End Sub

Private Sub tabConciliados_Click(PreviousTab As Integer)
    Call RecarregaGrids(True)
End Sub

Private Function TotalLinhasSelecionadas(grdGenerica As MSHFlexGrid) As Integer
    Dim intCont, i As Integer
    
    For i = 1 To grdGenerica.Rows - 1
        grdGenerica.Row = i
        If grdGenerica.CellPicture = imgCheck.ListImages(grdChecked).Picture Then
            intCont = intCont + 1
        End If
    Next
    TotalLinhasSelecionadas = intCont
End Function
