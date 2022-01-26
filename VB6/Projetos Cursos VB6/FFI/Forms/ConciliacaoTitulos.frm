VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmConciliacaoTitulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conciliação Bancária"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   10965
   Begin VB.Frame fraBotoes 
      Height          =   7620
      Left            =   9495
      TabIndex        =   24
      Top             =   -45
      Width           =   1455
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
         Height          =   375
         Left            =   135
         TabIndex        =   21
         Top             =   630
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sai&r"
         Height          =   375
         Left            =   135
         TabIndex        =   22
         Top             =   1035
         Width           =   1185
      End
      Begin ComctlLib.ImageList imgCheck 
         Left            =   360
         Top             =   1800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ConciliacaoTitulos.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "ConciliacaoTitulos.frx":0352
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraControles 
      Height          =   7620
      Left            =   0
      TabIndex        =   23
      Top             =   -45
      Width           =   9465
      Begin VB.OptionButton optOrdenarPagamento 
         Caption         =   "Ordenar por Pagamento"
         Height          =   195
         Left            =   4095
         TabIndex        =   55
         Top             =   3285
         Width           =   1995
      End
      Begin VB.OptionButton optOrdenarLiberacao 
         Caption         =   "Ordenar por Liberação"
         Height          =   195
         Left            =   1305
         TabIndex        =   54
         Top             =   3285
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.CommandButton cmdSelecionaNenhum 
         Caption         =   "&Nenhum"
         Height          =   375
         Left            =   8145
         TabIndex        =   44
         Top             =   3195
         Width           =   1215
      End
      Begin VB.CommandButton cmdSelecionaTodos 
         Caption         =   "&Todos"
         Height          =   375
         Left            =   6885
         TabIndex        =   43
         Top             =   3195
         Width           =   1215
      End
      Begin VB.Frame fraResultados 
         Caption         =   "&Resultados"
         Height          =   3630
         Left            =   90
         TabIndex        =   41
         Top             =   3510
         Width           =   9285
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConcTitulo 
            Height          =   3315
            Left            =   90
            TabIndex        =   42
            Top             =   225
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   5847
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraFiltros 
         Caption         =   "Filtros"
         Height          =   2625
         Left            =   90
         TabIndex        =   25
         Top             =   540
         Width           =   9285
         Begin VB.CheckBox chkConciliado 
            Caption         =   "Mostrar itens conciliados"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   2295
         End
         Begin VB.Frame fraPagamentoRecebimento 
            Caption         =   "Pagamento / Recebimento"
            Height          =   555
            Left            =   135
            TabIndex        =   40
            Top             =   2025
            Width           =   3805
            Begin VB.OptionButton optPagamento 
               Caption         =   "Pagamento"
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   270
               Value           =   -1  'True
               Width           =   1230
            End
            Begin VB.OptionButton optRecebimento 
               Caption         =   "Recebimento"
               Height          =   195
               Left            =   2265
               TabIndex        =   8
               Top             =   270
               Width           =   1275
            End
         End
         Begin VB.Frame fraOrigem 
            Caption         =   "Origem dos Registros"
            Height          =   555
            Left            =   5355
            TabIndex        =   39
            Top             =   2025
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
               Left            =   1215
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
            Left            =   1035
            TabIndex        =   0
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
         Begin Fox.EBSData edtEmissaoFinal 
            Height          =   330
            Left            =   2655
            TabIndex        =   1
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
         Begin Fox.EBSData edtPagamentoInicial 
            Height          =   330
            Left            =   1035
            TabIndex        =   2
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
         Begin Fox.EBSData edtPagamentoFinal 
            Height          =   330
            Left            =   2655
            TabIndex        =   3
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
         Begin Fox.EBSData edtLiberacaoInicial 
            Height          =   330
            Left            =   1035
            TabIndex        =   4
            Top             =   1305
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
            Left            =   2655
            TabIndex        =   5
            Top             =   1305
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
            Left            =   6165
            TabIndex        =   9
            Top             =   585
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
            Left            =   7965
            TabIndex        =   10
            Top             =   585
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
            Left            =   6165
            TabIndex        =   11
            Top             =   945
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
            Left            =   7965
            TabIndex        =   12
            Top             =   945
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
            Left            =   6165
            TabIndex        =   13
            Top             =   1305
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
            Left            =   7965
            TabIndex        =   14
            Top             =   1305
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
            Left            =   1035
            TabIndex        =   6
            Top             =   1665
            Width           =   2910
            _ExtentX        =   5133
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
            Left            =   6165
            TabIndex        =   15
            Top             =   1665
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
            Left            =   7965
            TabIndex        =   16
            Top             =   1665
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
            Caption         =   "até"
            Height          =   195
            Left            =   7560
            TabIndex        =   53
            Top             =   1710
            Width           =   225
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Cheque"
            Height          =   195
            Left            =   5400
            TabIndex        =   52
            Top             =   1710
            Width           =   555
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reg."
            Height          =   195
            Left            =   135
            TabIndex        =   38
            Top             =   1710
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   7560
            TabIndex        =   37
            Top             =   1395
            Width           =   225
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Conta"
            Height          =   195
            Left            =   5400
            TabIndex        =   36
            Top             =   1395
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   7560
            TabIndex        =   35
            Top             =   1035
            Width           =   225
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            Height          =   195
            Left            =   5400
            TabIndex        =   34
            Top             =   1035
            Width           =   465
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   7560
            TabIndex        =   33
            Top             =   675
            Width           =   225
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            Height          =   195
            Left            =   5400
            TabIndex        =   32
            Top             =   675
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   2340
            TabIndex        =   31
            Top             =   1395
            Width           =   225
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Liberação"
            Height          =   195
            Left            =   135
            TabIndex        =   30
            Top             =   1395
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   2340
            TabIndex        =   29
            Top             =   1035
            Width           =   225
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Pagamento"
            Height          =   195
            Left            =   135
            TabIndex        =   28
            Top             =   1035
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "até"
            Height          =   195
            Left            =   2340
            TabIndex        =   27
            Top             =   630
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            Height          =   195
            Left            =   135
            TabIndex        =   26
            Top             =   675
            Width           =   585
         End
      End
      Begin Fox.EBSText etxQuantidadeListada 
         Height          =   330
         Left            =   810
         TabIndex        =   46
         Top             =   7215
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
      Begin Fox.EBSText etxQuantidadeSelecionada 
         Height          =   330
         Left            =   4500
         TabIndex        =   48
         Top             =   7215
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
      Begin Fox.EBSText etxTotalValor 
         Height          =   330
         Left            =   7830
         TabIndex        =   49
         Top             =   7215
         Width           =   1545
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
      Begin Fox.EBSText etxEmpUser 
         Height          =   330
         Left            =   1125
         TabIndex        =   56
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
         Left            =   180
         TabIndex        =   51
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total"
         Enabled         =   0   'False
         Height          =   195
         Left            =   6975
         TabIndex        =   50
         Top             =   7260
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Selecionados"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3465
         TabIndex        =   47
         Top             =   7260
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Listados"
         Enabled         =   0   'False
         Height          =   195
         Left            =   135
         TabIndex        =   45
         Top             =   7260
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmConciliacaoTitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const grdChecked = 2
Private Const grdUnchecked = 1
Private Const colCheck = 1
Private mlngTListado      As Long
Private mdblTValor        As Double

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
    Dim i                     As Integer
    Dim rstEmp                As Object
    Dim strSql                As String
    Dim strPagRec             As String
    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
    Dim objLogLancDup         As New clsLogLancamentosDuplicatas
    
    
    If MsgBox("Confirma a conciliação bancária para o(s) título(s) selecionado(s) ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    strPagRec = ""
    i = 1
    If optPagamento.value = True Then
        strPagRec = "P"
    ElseIf optRecebimento.value = True Then
        strPagRec = "R"
    End If
    'Realiza uma verificação em todo o grid, os itens marcados é realizado um update no banco e mudado o status do campo "CONCILIADO" para TRUE.
    For i = 1 To grdConcTitulo.Rows - 1
        grdConcTitulo.Row = i
        If grdConcTitulo.CellPicture = imgCheck.ListImages(grdChecked).Picture Then
            If grdConcTitulo.TextMatrix(i, 15) = "Lançamentos" Then
                strSql = "UPDATE Lançamentos SET Lançamentos.Conciliado = True "
                strSql = strSql & "WHERE Lançamentos.PagRec ='" & strPagRec & _
                "' And Lançamentos.Código =" & grdConcTitulo.TextMatrix(i, 3) & _
                " And Lançamentos.parcela =" & grdConcTitulo.TextMatrix(i, 4) & _
                " And Lançamentos.Tipo ='" & grdConcTitulo.TextMatrix(i, 5) & "'"
            ElseIf grdConcTitulo.TextMatrix(i, 15) = "Duplicatas" Then
                strSql = "UPDATE Duplicatas SET Duplicatas.Conciliado = True "
                strSql = strSql & "WHERE Duplicatas.PagRec ='" & strPagRec & _
                "' And Duplicatas.Nota =" & grdConcTitulo.TextMatrix(i, 3) & _
                " And Duplicatas.parcela =" & grdConcTitulo.TextMatrix(i, 4) & _
                " And Duplicatas.Tipo ='" & grdConcTitulo.TextMatrix(i, 5) & "'"
            End If
            ExecuteSQL strSql
            'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
            With grdConcTitulo
                Call objLogLancDup.SetKey(strPagRec, CDbl(.TextMatrix(i, 3)), .TextMatrix(i, 7), .TextMatrix(i, 5), CLng(.TextMatrix(i, 4)), IIf(.TextMatrix(i, 15) = "Lançamentos", Lancamento, Duplicata))
                Call objLogLancDup.InsertMsg("Campo Conciliado foi marcado através da rotina " & Me.Caption & ".")
            End With
        End If
    Next
    'Prepara o grid novamente sem os itens que foram conciliados.
    Call PreparaGrid
    Call FiltraLista
    cmdConciliar.Enabled = False
    Set objLogLancDup = Nothing
End Sub

Private Sub cmdPesquisar_Click()
    Call PreparaGrid
    Call FiltraLista
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
        .Cols = 16
        .FixedCols = 1
        .Rows = 2
        'Configura a coluna fixa
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 150
        'Configura a coluna de seleção
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        'Configura a coluna de Cheque
        .TextMatrix(0, 2) = "Cheque"
        .ColWidth(2) = 700
        .ColAlignment(2) = flexAlignLeftCenter
        'Configura a coluna de Número
        .TextMatrix(0, 3) = "Número"
        .ColWidth(3) = 700
        .ColAlignment(3) = flexAlignLeftCenter
        'Configura a coluna Parcela
        .TextMatrix(0, 4) = "Parcela"
        .ColWidth(4) = 700
        .ColAlignment(4) = flexAlignLeftCenter
        'Configura a coluna de Tipo
        .TextMatrix(0, 5) = "Tipo"
        .ColWidth(5) = 700
        .ColAlignment(5) = flexAlignRightCenter
        'Configura a coluna de Descrição
        .TextMatrix(0, 6) = "Descrição"
        .ColWidth(6) = 1800
        .ColAlignment(6) = flexAlignRightCenter
        'Configura a coluna de Empresa
        .TextMatrix(0, 7) = "Empresa"
        .ColWidth(7) = 2000
        .ColAlignment(7) = flexAlignRightCenter
        'Configura a coluna de Banco
        .TextMatrix(0, 8) = "Banco"
        .ColWidth(8) = 600
        .ColAlignment(8) = flexAlignLeftCenter
        'Configura a coluna de Conta
        .TextMatrix(0, 9) = "Conta"
        .ColWidth(9) = 600
        .ColAlignment(9) = flexAlignLeftCenter
        'Configura a coluna de C.C.
        .TextMatrix(0, 10) = "Centro Custo"
        .ColWidth(10) = 1200
        .ColAlignment(10) = flexAlignRightCenter
        'Configura a coluna de Vencimento
        .TextMatrix(0, 11) = "Vencimento"
        .ColWidth(11) = 1200
        .ColAlignment(11) = flexAlignRightCenter
        'Configura a coluna de Valor
        .TextMatrix(0, 12) = "Valor"
        .ColWidth(12) = 800
        .ColAlignment(12) = flexAlignRightCenter
        'Configura a coluna de Emissão
        .TextMatrix(0, 13) = "Emissão"
        .ColWidth(13) = 1200
        .ColAlignment(13) = flexAlignRightCenter
        'Configura a coluna de Controle
        .TextMatrix(0, 14) = "Controle"
        .ColWidth(14) = 800
        .ColAlignment(14) = flexAlignRightCenter
        'Configura a coluna de L_D(identifica qual é a tabela de origem)
        .TextMatrix(0, 15) = "L_D"
        .ColWidth(15) = 1200
        .ColAlignment(15) = flexAlignLeftCenter
            
        For intIndex = 0 To .Cols - 1
            .TextMatrix(1, intIndex) = ""
        Next
        .col = colCheck
        .Row = 1
        Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
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
Private Sub FiltraLista()
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
    etxQuantidadeListada.valorInteiro = 0
    etxTotalValor.valorMoeda = 0#
    etxQuantidadeSelecionada.valorInteiro = 0
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
    If optPagamento.value = True Then
        strFiltro = strFiltro & " AND Lançamentos.PagRec='P' "
    ElseIf optRecebimento.value = True Then
        strFiltro = strFiltro & " AND Lançamentos.PagRec='R' "
    End If
    'Ordem
    If optOrdenarLiberacao.value = True Then
        strOrdem = " ORDER BY Liberação "
    End If
    If optOrdenarPagamento.value = True Then
        strOrdem = " ORDER BY Pagamento "
    End If
    'Check
    If ChkConciliado.value = 1 Then
        strFiltro = strFiltro & " AND Lançamentos.conciliado= True "
    Else
        strFiltro = strFiltro & " AND Lançamentos.conciliado= False "
    End If
    
    'Seleção da tabela de lançamentos, com a condição de pagamentos terem sidos realizados e o campo conciliado igual a falso.
    strSql = "SELECT Lançamentos.[Abatimento] ,Lançamentos.[Acréscimo] ,Lançamentos.[Código], Lançamentos.[Parcela], Lançamentos.[Empresa], Lançamentos.[Tipo], Lançamentos.[Descrição], Lançamentos.[Emissão], Lançamentos.[Vencimento], Lançamentos.[Pagamento], Lançamentos.[Liberação], Lançamentos.[Valor Original], Lançamentos.[Banco], Lançamentos.[Conta], Lançamentos.[Centro], Lançamentos.[Cheque],  Lançamentos.[Controle], Lançamentos.[Situação], Lançamentos.[Alteração], Lançamentos.[Conciliado], 'Lançamentos' as L_D "
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
    AtualizaLista strSql
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
Private Sub AtualizaLista(strSql As String)
    Dim rsConciliacao  As Object
    Dim i              As Integer
    
    mlngTListado = 0
    mdblTValor = 0
    If AbreRecordset(rsConciliacao, strSql) = WL_OK Then
        rsConciliacao.MoveFirst
        i = 1
        While Not rsConciliacao.EOF
            grdConcTitulo.AddItem ("")
            grdConcTitulo.col = 1
            grdConcTitulo.Row = grdConcTitulo.Rows - 1
            Set grdConcTitulo.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            grdConcTitulo.TextMatrix(i, 2) = GetValue(rsConciliacao, "Cheque", "")
            grdConcTitulo.ColAlignment(2) = flexAlignRightCenter
            grdConcTitulo.TextMatrix(i, 3) = GetValue(rsConciliacao, "Código", "")
            grdConcTitulo.ColAlignment(3) = flexAlignRightCenter
            grdConcTitulo.TextMatrix(i, 4) = GetValue(rsConciliacao, "Parcela", "")
            grdConcTitulo.ColAlignment(4) = flexAlignRightCenter
            grdConcTitulo.TextMatrix(i, 5) = GetValue(rsConciliacao, "Tipo", "")
            grdConcTitulo.ColAlignment(5) = flexAlignLeftCenter
            grdConcTitulo.TextMatrix(i, 6) = GetValue(rsConciliacao, "Descrição", "")
            grdConcTitulo.ColAlignment(6) = flexAlignLeftCenter
            grdConcTitulo.TextMatrix(i, 7) = GetValue(rsConciliacao, "Empresa", "")
            grdConcTitulo.ColAlignment(7) = flexAlignLeftCenter
            grdConcTitulo.TextMatrix(i, 8) = GetValue(rsConciliacao, "Banco", "")
            grdConcTitulo.ColAlignment(8) = flexAlignRightCenter
            grdConcTitulo.TextMatrix(i, 9) = GetValue(rsConciliacao, "Conta", "")
            grdConcTitulo.ColAlignment(9) = flexAlignRightCenter
            grdConcTitulo.TextMatrix(i, 10) = GetFieldValue("Descrição", "Centros", "Código = " & GetValue(rsConciliacao, "Centro", ""), , "")
            grdConcTitulo.ColAlignment(10) = flexAlignLeftCenter
            grdConcTitulo.TextMatrix(i, 11) = GetValue(rsConciliacao, "Vencimento", "")
            grdConcTitulo.ColAlignment(11) = flexAlignLeftCenter
            'pt. 106685 - Ivo Sousa (06/05/2011)
            'Conforme solicitação, somado o valor de Acréscimo e subtraído o valor de Abatimentos
            grdConcTitulo.TextMatrix(i, 12) = Format(GetValue(rsConciliacao, "Valor Original", ZERO) - GetValue(rsConciliacao, "Abatimento", ZERO) + GetValue(rsConciliacao, "Acréscimo", ZERO), "#0.00")
            grdConcTitulo.TextMatrix(i, 13) = GetValue(rsConciliacao, "Emissão", "")
            grdConcTitulo.ColAlignment(13) = flexAlignLeftCenter
            grdConcTitulo.TextMatrix(i, 14) = GetValue(rsConciliacao, "Controle", "")
            grdConcTitulo.ColAlignment(14) = flexAlignLeftCenter
            grdConcTitulo.TextMatrix(i, 15) = GetValue(rsConciliacao, "L_D", "")
            
            mlngTListado = mlngTListado + 1
            'mdblTValor = dblTValor + GetValue(rsConciliacao, "Valor Original", ZERO)
            i = i + 1
            rsConciliacao.MoveNext
        Wend
        If grdConcTitulo.Rows > 2 Then
            grdConcTitulo.RemoveItem (grdConcTitulo.Rows - 1)
        End If
    
    Else
        MsgBox "Não há registros à conciliar.", vbInformation, NomeModulo
    End If
    etxQuantidadeListada.valorInteiro = mlngTListado
    'etxTotalValor.valorMoeda = mdblTValor
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
    cmdConciliar.Enabled = False
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
        
    etxQuantidadeSelecionada.valorInteiro = 0
    With grdConcTitulo
        If .TextMatrix(1, 2) <> "" Then
            For intIndex = 1 To .Rows - 1
                .Row = intIndex
                .col = colCheck
                Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                etxQuantidadeSelecionada.valorInteiro = etxQuantidadeSelecionada.valorInteiro + 1
                etxTotalValor.valorMoeda = etxTotalValor.valorMoeda + grdConcTitulo.TextMatrix(grdConcTitulo.Row, 12)
            Next
        End If
    End With
    If ChkConciliado.value = 0 Then
        cmdConciliar.Enabled = True
    End If
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

Private Sub Form_Load()
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
    PreparaGrid
    CarregaCombo
    cmdConciliar.Enabled = False
    ecbTipo.SelectItem "Todos"
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
    On Error GoTo err
        With grdConcTitulo
            .CellPictureAlignment = flexAlignCenterCenter
            If LinhaSelecionada(.Row) Then
                etxQuantidadeSelecionada.valorInteiro = etxQuantidadeSelecionada.valorInteiro - 1
                etxTotalValor.valorMoeda = etxTotalValor.valorMoeda - grdConcTitulo.TextMatrix(grdConcTitulo.Row, 12)
                                Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            Else
                etxQuantidadeSelecionada.valorInteiro = etxQuantidadeSelecionada.valorInteiro + 1
                etxTotalValor.valorMoeda = etxTotalValor.valorMoeda + grdConcTitulo.TextMatrix(grdConcTitulo.Row, 12)
                Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
            End If
        End With
        'Se não tiver nenhum item marcado no grid desabilita o botão conciliar.
        If etxQuantidadeSelecionada.valorInteiro > 0 Then
            If ChkConciliado.value = 0 Then
                cmdConciliar.Enabled = True
            End If
        Else
            cmdConciliar.Enabled = False
        End If
    Exit Sub
err:
End Sub

Private Function LinhaSelecionada(lngLinha As Long) As Boolean
    If lngLinha <= grdConcTitulo.Rows - 1 Then
        grdConcTitulo.Row = lngLinha
        grdConcTitulo.col = colCheck
        LinhaSelecionada = (grdConcTitulo.CellPicture = imgCheck.ListImages(2).Picture)
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
            strValida = strValida & "- Data de emissão inicial é menor que a data de emissão final." & vbCrLf
        End If
    End If
    'Data Pagamento.
    If edtPagamentoInicial.Data <> Trim("00:00:00") And edtPagamentoFinal.Data <> Trim("00:00:00") Then
        If DateDiff("d", edtPagamentoInicial.Data, edtPagamentoFinal.Data) < 0 Then
            strValida = strValida & "- Data de pagamento inicial é menor que a data de pagamento final." & vbCrLf
        End If
    End If
    'Data Liberação.
    If edtLiberacaoInicial.Data <> Trim("00:00:00") And edtLiberacaoFinal.Data <> Trim("00:00:00") Then
        If DateDiff("d", edtLiberacaoInicial.Data, edtLiberacaoFinal.Data) < 0 Then
            strValida = strValida & "- Data de liberação inicial é menor que a data de liberação final." & vbCrLf
        End If
    End If
    If Trim(strValida) <> "" Then
        MsgBox strValida
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
