VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmReajusteDupLan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reajuste de Duplicatas a Receber"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   13215
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   11835
      TabIndex        =   73
      Top             =   -40
      Width           =   1360
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   80
         TabIndex        =   38
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   80
         TabIndex        =   39
         Top             =   580
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   80
         TabIndex        =   40
         Top             =   990
         Width           =   1215
      End
   End
   Begin VB.Frame fraGeral 
      Height          =   8535
      Left            =   30
      TabIndex        =   0
      Top             =   -40
      Width           =   11780
      Begin VB.Frame Frame2 
         Height          =   675
         Left            =   60
         TabIndex        =   65
         Top             =   7800
         Width           =   11660
         Begin Fox.EBSText etxQtTitulo 
            Height          =   330
            Left            =   3240
            TabIndex        =   66
            Top             =   210
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            TipoCriterio    =   0
            Alinhamento     =   1
            Locked          =   -1  'True
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
         Begin Fox.EBSText etxVlSaldo 
            Height          =   330
            Left            =   6480
            TabIndex        =   67
            Top             =   210
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
            TipoCriterio    =   6
            Alinhamento     =   1
            Mascara         =   "##,##0.00"
            Locked          =   -1  'True
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
         Begin Fox.EBSText etxVlTotal 
            Height          =   330
            Left            =   9960
            TabIndex        =   68
            Top             =   210
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
            TipoCriterio    =   6
            Alinhamento     =   1
            Mascara         =   "##,##0.00"
            Locked          =   -1  'True
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
         Begin VB.Label lblQtTitulo 
            Alignment       =   1  'Right Justify
            Caption         =   "Qt.Título"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   2040
            TabIndex        =   71
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label lblVlSaldo 
            Alignment       =   1  'Right Justify
            Caption         =   "Vl.Saldo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   5300
            TabIndex        =   70
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label lblVlTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "Vl.Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   8760
            TabIndex        =   69
            Top             =   255
            Width           =   1095
         End
      End
      Begin VB.Frame fraFiltro 
         Caption         =   "Filtro"
         Height          =   2040
         Left            =   60
         TabIndex        =   45
         Top             =   500
         Width           =   11655
         Begin Fox.EBSData edtLibarecaoFin 
            Height          =   330
            Left            =   2595
            TabIndex        =   2
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Fox.EBSData edtLibarecaoIni 
            Height          =   330
            Left            =   1155
            TabIndex        =   1
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Fox.EBSData edtVencimentoFin 
            Height          =   330
            Left            =   2595
            TabIndex        =   4
            Top             =   580
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Fox.EBSData edtVencimentoIni 
            Height          =   330
            Left            =   1155
            TabIndex        =   3
            Top             =   580
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Fox.EBSData edtEmissaoFin 
            Height          =   330
            Left            =   2595
            TabIndex        =   6
            Top             =   930
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Fox.EBSData edtEmissaoIni 
            Height          =   330
            Left            =   1155
            TabIndex        =   5
            Top             =   930
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Fox.EBSText etxBancoIni 
            Height          =   330
            Left            =   1150
            TabIndex        =   9
            Top             =   1620
            Width           =   1215
            _ExtentX        =   344
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
         Begin Fox.EBSText etxBancoFin 
            Height          =   330
            Left            =   2595
            TabIndex        =   10
            Top             =   1620
            Width           =   1215
            _ExtentX        =   344
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
         Begin Fox.EBSText etxContaIni 
            Height          =   330
            Left            =   5115
            TabIndex        =   11
            Top             =   240
            Width           =   1215
            _ExtentX        =   344
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
         Begin Fox.EBSText etxContaFin 
            Height          =   330
            Left            =   6555
            TabIndex        =   12
            Top             =   240
            Width           =   1215
            _ExtentX        =   344
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
         Begin Fox.EBSText etxCentroCustoIni 
            Height          =   330
            Left            =   5115
            TabIndex        =   13
            Top             =   585
            Width           =   1215
            _ExtentX        =   344
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Centros"
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
         Begin Fox.EBSText etxCentroCustoFin 
            Height          =   330
            Left            =   6570
            TabIndex        =   14
            Top             =   585
            Width           =   1215
            _ExtentX        =   344
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Centros"
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
         Begin Fox.EBSText etxValOriginalIni 
            Height          =   330
            Left            =   5115
            TabIndex        =   15
            Top             =   930
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
            MaxLength       =   9
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxValOriginalFin 
            Height          =   330
            Left            =   6555
            TabIndex        =   16
            Top             =   930
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
            MaxLength       =   9
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxNotaCodigoIni 
            Height          =   330
            Left            =   5115
            TabIndex        =   17
            Top             =   1275
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxNotaCodigoFin 
            Height          =   330
            Left            =   6555
            TabIndex        =   18
            Top             =   1275
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxEmpresa 
            Height          =   330
            Left            =   8130
            TabIndex        =   21
            Top             =   240
            Width           =   2985
            _ExtentX        =   125889
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   15
            Caption         =   "Empresa"
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Apel"
            CampoDescricao  =   "Razão"
            TabelaConsulta  =   "Empresas"
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
         Begin Fox.EBSText etxNossoNr 
            Height          =   330
            Left            =   8040
            TabIndex        =   22
            Top             =   600
            Width           =   3435
            _ExtentX        =   242861
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   20
            Caption         =   "Nosso Nr."
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
         Begin Fox.EBSText etxControle 
            Height          =   330
            Left            =   8160
            TabIndex        =   24
            Top             =   1290
            Width           =   2220
            _ExtentX        =   202830
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   18
            Caption         =   "Controle"
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
         Begin Fox.EBSCombo cboTipo 
            Height          =   315
            Left            =   8835
            TabIndex        =   23
            Top             =   945
            Width           =   1335
            _ExtentX        =   2355
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
         Begin Fox.EBSData edtPagamentoFin 
            Height          =   330
            Left            =   2595
            TabIndex        =   8
            Top             =   1280
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Fox.EBSData edtPagamentoIni 
            Height          =   330
            Left            =   1155
            TabIndex        =   7
            Top             =   1280
            Width           =   1215
            _ExtentX        =   2143
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
         Begin Fox.EBSText etxParcelaIni 
            Height          =   330
            Left            =   5115
            TabIndex        =   19
            Top             =   1620
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxParcelaFin 
            Height          =   330
            Left            =   6555
            TabIndex        =   20
            Top             =   1620
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
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
            ExibeDescricao  =   0   'False
         End
         Begin VB.Label Label13 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   6390
            TabIndex        =   76
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Parcela"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   4050
            TabIndex        =   75
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblLiberacao 
            Alignment       =   1  'Right Justify
            Caption         =   "Liberação"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   220
            TabIndex        =   64
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   2430
            TabIndex        =   63
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lblVencimento 
            Alignment       =   1  'Right Justify
            Caption         =   "Vencimento"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   220
            TabIndex        =   62
            Top             =   640
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   2430
            TabIndex        =   61
            Top             =   640
            Width           =   135
         End
         Begin VB.Label lblEmissao 
            Alignment       =   1  'Right Justify
            Caption         =   "Emissão"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   345
            TabIndex        =   60
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   2430
            TabIndex        =   59
            Top             =   990
            Width           =   135
         End
         Begin VB.Label lblCodLote 
            Alignment       =   1  'Right Justify
            Caption         =   "Banco"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   340
            TabIndex        =   58
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   2430
            TabIndex        =   57
            Top             =   1680
            Width           =   135
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Conta"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   4305
            TabIndex        =   56
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   6390
            TabIndex        =   55
            Top             =   300
            Width           =   135
         End
         Begin VB.Label lblCentroCusto 
            Alignment       =   1  'Right Justify
            Caption         =   "C.Custo"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   4290
            TabIndex        =   54
            Top             =   645
            Width           =   735
         End
         Begin VB.Label lblCentroA 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   6390
            TabIndex        =   53
            Top             =   645
            Width           =   135
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Vl.Original"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   4290
            TabIndex        =   52
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   6390
            TabIndex        =   51
            Top             =   990
            Width           =   135
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Nota"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   4050
            TabIndex        =   50
            Top             =   1335
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   6390
            TabIndex        =   49
            Top             =   1335
            Width           =   135
         End
         Begin VB.Label lblTipo 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   7995
            TabIndex        =   48
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "a"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   2430
            TabIndex        =   47
            Top             =   1335
            Width           =   135
         End
         Begin VB.Label lblPagamento 
            Alignment       =   1  'Right Justify
            Caption         =   "Pagamento"
            ForeColor       =   &H80000007&
            Height          =   255
            Left            =   225
            TabIndex        =   46
            Top             =   1335
            Width           =   855
         End
      End
      Begin VB.Frame fraOrdem 
         Caption         =   "Ordem"
         Height          =   825
         Left            =   60
         TabIndex        =   44
         Top             =   2565
         Width           =   6465
         Begin VB.OptionButton optEmpresa 
            Caption         =   "Empresa"
            Height          =   195
            Left            =   5130
            TabIndex        =   32
            Top             =   510
            Width           =   1215
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Tipo"
            Height          =   195
            Left            =   210
            TabIndex        =   29
            Top             =   500
            Width           =   1335
         End
         Begin VB.OptionButton optNotaCodigo 
            Caption         =   "Nota/Código"
            Height          =   195
            Left            =   210
            TabIndex        =   25
            Top             =   260
            Width           =   1215
         End
         Begin VB.OptionButton optEmissao 
            Caption         =   "Emissão"
            Height          =   195
            Left            =   3570
            TabIndex        =   27
            Top             =   255
            Width           =   1215
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Valor"
            Height          =   195
            Left            =   5130
            TabIndex        =   28
            Top             =   270
            Width           =   1215
         End
         Begin VB.OptionButton optLiberacao 
            Caption         =   "Liberação"
            Height          =   195
            Left            =   1815
            TabIndex        =   30
            Top             =   500
            Width           =   1335
         End
         Begin VB.OptionButton optVencimento 
            Caption         =   "Vencimento"
            Height          =   195
            Left            =   1815
            TabIndex        =   26
            Top             =   260
            Width           =   1215
         End
         Begin VB.OptionButton optControle 
            Caption         =   "Controle"
            Height          =   195
            Left            =   3555
            TabIndex        =   31
            Top             =   500
            Width           =   1215
         End
      End
      Begin VB.Frame fraSituacao 
         Caption         =   "Situação"
         Height          =   825
         Left            =   6570
         TabIndex        =   43
         Top             =   2565
         Width           =   5130
         Begin VB.OptionButton optAndamento 
            Caption         =   "Data Base 1"
            Height          =   195
            Left            =   210
            TabIndex        =   33
            Top             =   360
            Width           =   1365
         End
         Begin VB.OptionButton optFinalizadas 
            Caption         =   "Data Base 2"
            Height          =   195
            Left            =   1890
            TabIndex        =   34
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "Todas"
            Height          =   195
            Left            =   3540
            TabIndex        =   35
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Reajuste"
         Height          =   645
         Left            =   60
         TabIndex        =   42
         Top             =   3420
         Width           =   11655
         Begin VB.CommandButton cmdExecutar 
            Caption         =   "&Executar"
            Height          =   375
            Left            =   10050
            TabIndex        =   37
            Top             =   180
            Width           =   1215
         End
         Begin Fox.EBSText etxPercAdicional 
            Height          =   330
            Left            =   1170
            TabIndex        =   36
            Top             =   210
            Width           =   1200
            _ExtentX        =   609
            _ExtentY        =   582
            Tipo            =   2
            CasasDecimais   =   4
            TipoTexto       =   0
            MaxLength       =   9
            TipoCriterio    =   5
            Alinhamento     =   1
            Mascara         =   "##,##0.0000"
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
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Adicional (%)"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   150
            TabIndex        =   74
            Top             =   270
            Width           =   945
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDuplLanc 
         Height          =   3675
         Left            =   60
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   4125
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   6482
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Fox.EBSText etxEmpUser 
         Height          =   330
         Left            =   360
         TabIndex        =   72
         Top             =   180
         Width           =   10665
         _ExtentX        =   447040
         _ExtentY        =   582
         Tipo            =   4
         Caption         =   "Empresa Usuária"
         Enabled         =   0   'False
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Apel"
         CampoDescricao  =   "Razão"
         TabelaConsulta  =   "Empresas"
         TamanhoDescricao=   7700
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
   End
End
Attribute VB_Name = "frmReajusteDupLan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lWnd                As Long
Private mrsRegistros        As Object
Private mcurTotalGeral      As Currency
Private mlngQtdTitulos      As Long
Private mcurSaldoTotal      As Currency
Private mstrPagRec          As String

Private Const mstrTabelaCotacao = "Cotações"

Private Sub cmdConsultar_Click()
    If AbreRecordset(mrsRegistros, MontaSqlDuplicatas & " ORDER BY Duplicatas." & Mid(BuscaOrderBy, 11, Len(BuscaOrderBy))) Then
        Call CarregaGrid
        etxVlSaldo.valorMoeda = Format(mcurSaldoTotal, "R$ 00.00#,##")
        etxVlTotal.valorMoeda = Format(mcurTotalGeral, "R$ 00.00#,##")
        etxQtTitulo.valorInteiro = mlngQtdTitulos
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpaCampos
End Sub

Private Sub cmdExecutar_Click()
    Dim objReajusteVo   As voReajusteDuplicatas
    Dim objReajusteBiz  As New bizReajusteDuplicatas
    Dim intCont         As Integer
    Dim strTitulos      As String
    Dim blnReajustou    As Boolean
    'Projeto: 100340 - Desenv.: 142889 - Ueder Budni (22/09/2016)
    Dim objLogDuplicata As New clsLogLancamentosDuplicatas
    
    
On Error GoTo err_Handler
    Aplicacao.Connect
    If grdDuplLanc.TextMatrix(1, 2) <> Empty Then
        If MsgBox("Atenção! Ao executar o reajuste, o valor original da(s) Duplicata(s) será alterado, confirma a operação?", vbYesNo, NomeModulo) = vbYes Then
            If mobjCache Is Nothing Then
                Set mobjCache = New clsCache
            End If
            For intCont = 1 To grdDuplLanc.Rows - 1
                Set objReajusteVo = New voReajusteDuplicatas
                With objReajusteVo
                    .data_reajuste = DateTime.Now
                    .Empresa = grdDuplLanc.TextMatrix(intCont, 7)
                    .Nota = grdDuplLanc.TextMatrix(intCont, 4)
                    .PagRec = grdDuplLanc.TextMatrix(intCont, 24)
                    .Parcela = grdDuplLanc.TextMatrix(intCont, 5)
                    If grdDuplLanc.TextMatrix(intCont, 2) <> Empty Then
                        .perc_1 = strToDbl(Replace(grdDuplLanc.TextMatrix(intCont, 2), " %", ""))
                    End If
                    If grdDuplLanc.TextMatrix(intCont, 3) <> Empty Then
                        .perc_2 = strToDbl(Replace(grdDuplLanc.TextMatrix(intCont, 3), " %", ""))
                    End If
                    .perc_adicionais = etxPercAdicional.valorDecimal
                    .Tipo = grdDuplLanc.TextMatrix(intCont, 6)
                    .usuario = UserName
                End With
                Call objReajusteBiz.init(Aplicacao)
                If Not objReajusteBiz.ReajustaDuplicata(objReajusteVo) Then
                    strTitulos = strTitulos & objReajusteVo.Tipo & " | " & objReajusteVo.Nota & " | " & objReajusteVo.Parcela & " - Descrição: " & objReajusteBiz.ErroReajuste & vbNewLine
                Else
                    blnReajustou = True
                    'Projeto: 100340 - Desenv.: 142889 - Ueder Budni (22/09/2016)
                    With objReajusteVo
                        Call RegistraLogLancDup(.Nota, .Empresa, .Tipo, .Parcela, .PagRec, Duplicata, .perc_adicionais)
                    End With
                    
                End If
                Set objReajusteVo = Nothing
            Next
            Set mobjCache = Nothing
            If strTitulos <> Empty Then
                MsgBox strTitulos
            ElseIf blnReajustou Then
                MsgBox "Duplicata(s) reajustada(s) com sucesso!"
                Call cmdConsultar_Click
            End If
        End If
    End If
    Set objLogDuplicata = Nothing
    Exit Sub
err_Handler:
    Set objLogDuplicata = Nothing
    Set mobjCache = Nothing
    Aplicacao.Disconnect
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub LimpaCampos()
    cboTipo.SelectItem "Todos"
    edtEmissaoIni.Clear
    edtEmissaoFin.Clear
    edtLibarecaoIni.Clear
    edtLibarecaoFin.Clear
    edtVencimentoIni.Clear
    edtVencimentoFin.Clear
    edtPagamentoIni.Clear
    edtPagamentoFin.Clear
    etxBancoIni.Clear
    etxBancoFin.Clear
    etxContaIni.Clear
    etxContaFin.Clear
    etxCentroCustoIni.Clear
    etxCentroCustoFin.Clear
    etxValOriginalIni.Clear
    etxValOriginalFin.Clear
    etxNotaCodigoIni.Clear
    etxNotaCodigoFin.Clear
    etxParcelaIni.Clear
    etxParcelaFin.Clear
    etxEmpresa.Clear
    etxNossoNr.Clear
    etxControle.Clear
    etxVlSaldo.Clear
    etxVlTotal.Clear
    etxQtTitulo.Clear
    optTodas.value = True
    optNotaCodigo.value = True
    Call CarregaColunasGrid
End Sub

Private Sub etxBancoIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxBancoIni.valorInteiro > 0 Then
            etxBancoIni.valorInteiro = 0
        End If
        Call PCampo("Bancos", "SELECT * FROM Bancos", pbCampo, etxBancoIni, "Banco")
    End If
End Sub

Private Sub etxBancoFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxBancoFin.valorInteiro > 0 Then
            etxBancoFin.valorInteiro = 0
        End If
        Call PCampo("Bancos", "SELECT * FROM Bancos", pbCampo, etxBancoFin, "Banco")
    End If
End Sub

Private Sub etxCentroCustoIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxCentroCustoIni.valorInteiro > 0 Then
            etxCentroCustoIni.valorInteiro = 0
        End If
        Call PCampo("Centro de Custo", "SELECT * FROM Centros", pbCampo, etxCentroCustoIni, "Código")
    End If
End Sub

Private Sub etxCentroCustoFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxCentroCustoFin.valorInteiro > 0 Then
            etxCentroCustoFin.valorInteiro = 0
        End If
        Call PCampo("Centro de Custo", "SELECT * FROM Centros", pbCampo, etxCentroCustoFin, "Código")
    End If
End Sub

Private Sub etxContaIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxContaIni.valorInteiro > 0 Then
            etxContaIni.valorInteiro = 0
        End If
        Call PCampo("Contas", "SELECT * FROM Contas", pbCampo, etxContaIni, "Código")
    End If
End Sub

Private Sub etxContaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxContaFin.valorInteiro > 0 Then
            etxContaFin.valorInteiro = 0
        End If
        Call PCampo("Contas", "SELECT * FROM Contas", pbCampo, etxContaFin, "Código")
    End If
End Sub

Private Sub etxEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxEmpresa.valorTexto <> "" Then
            etxEmpresa.valorTexto = Empty
        End If
        Call PCampo("Empresas", "SELECT * FROM Empresas", pbCampo, etxEmpresa, "Apel")
    End If
End Sub

Private Sub Form_Load()
    Call CarregaColunasGrid
    Call etxEmpUser.AddConexao(Aplicacao)
    Call etxBancoIni.AddConexao(Aplicacao)
    Call etxBancoFin.AddConexao(Aplicacao)
    Call etxContaIni.AddConexao(Aplicacao)
    Call etxContaFin.AddConexao(Aplicacao)
    Call etxCentroCustoIni.AddConexao(Aplicacao)
    Call etxCentroCustoFin.AddConexao(Aplicacao)
    Call etxEmpresa.AddConexao(Aplicacao)
    etxEmpUser.valorTexto = DonaSistema
    etxCentroCustoIni.Enabled = ConfigSys.ControlarCentrodeCusto
    etxCentroCustoFin.Enabled = ConfigSys.ControlarCentrodeCusto
    lblCentroA.Enabled = ConfigSys.ControlarCentrodeCusto
    lblCentroCusto.Enabled = ConfigSys.ControlarCentrodeCusto
    optTodas.value = True
    optNotaCodigo.value = True
    Aplicacao.Connect
    Call preencheComboTipos
    Aplicacao.Disconnect
End Sub

Private Sub grdDuplLanc_DblClick()
    Dim frmForm                 As Form
    Dim strTabela               As String
    Dim PagRec                  As String
    Dim Campo                   As String
    Dim cod                     As String
    Dim lngHelpContextID        As Long
    Dim strParcela              As String
    Dim strTipo                 As String
    Dim intSetRegistro          As Integer
    Dim strOrigem               As String
    Dim blnEscreve              As Boolean
    Dim lngCodigo               As Long
    Dim lngParcela              As Long
    Dim strEmpresa              As String
    Dim enumPagRec              As enuPagRec
    Dim enumLancDup             As enuLancDup
    
    
    If IsValid(grdDuplLanc.TextMatrix(grdDuplLanc.Row, 1)) Then
        If grdDuplLanc.TextMatrix(grdDuplLanc.Row, 1) = "Dup" Then
            strTabela = "Duplicatas"
        Else
            strTabela = "Lançamentos"
        End If
        'Projeto: #1203 - História: #10582 - Desenvolvimento#12134 - João Henrique(18/04/2012)
        lngCodigo = grdDuplLanc.TextMatrix(grdDuplLanc.Row, 2)
        strTipo = grdDuplLanc.TextMatrix(grdDuplLanc.Row, 4)
        lngParcela = grdDuplLanc.TextMatrix(grdDuplLanc.Row, 3)
        strEmpresa = grdDuplLanc.TextMatrix(grdDuplLanc.Row, 5)
        PagRec = grdDuplLanc.TextMatrix(grdDuplLanc.Row, 22)
        
        If strTabela = "Duplicatas" Then
            enumLancDup = Duplicata
        Else
            enumLancDup = Lancamento
        End If
        
        If PagRec = "R" Then
            enumPagRec = Recebimento
        Else
            enumPagRec = Pagamento
        End If
        
        frmLancamentoDuplicata.LancDup = enumLancDup
        frmLancamentoDuplicata.PagRec = enumPagRec
        blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, 2061, frmLancamentoDuplicata.name, "Lançamentos a Pagar ou Pagos")
        Call mostrarForm(frmLancamentoDuplicata, 2061)
        Call frmLancamentoDuplicata.CarregarLancamentoDuplicataOutrasRotinas(lngCodigo, strTipo, lngParcela, strEmpresa, enumPagRec, enumLancDup)
        
        'pt. 89833 - Ivo Sousa (04/11/2008)
        If TemPermissao(grupoUsuario, NumeroModulo, opAlterar, lngHelpContextID, False) Then
            Me.Hide
            'Projeto: #1203 - História: #10582 - Desenvolvimento#12134 - João Henrique(18/04/2012)
            lWnd = frmLancamentoDuplicata.hWnd
            WaitWindowClose lWnd 'Esperar até que a janela seja fechada
            Me.Show
        Else
            MsgBox "O usuário " & UserName & " não tem permissão para alterações na rotina de " & strTabela, vbInformation, NomeModulo
        End If
On Error Resume Next
        If (err.Number) Then
           err.Clear
        End If
        Set frmForm = Nothing
    End If
End Sub

Private Sub CarregaColunasGrid()
    Dim intIndex As Long

    With grdDuplLanc
        .Cols = 25
        .FixedCols = 1
        .Rows = 2
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 120
                
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 0.1
        
        .TextMatrix(0, 2) = "Índice Data Base 1"
        .ColWidth(2) = 1500
        .ColAlignment(2) = flexAlignRightCenter
        
        .TextMatrix(0, 3) = "Índice Data Base 2"
        .ColWidth(3) = 1500
        .ColAlignment(3) = flexAlignRightCenter
        
        .TextMatrix(0, 4) = "Nota/Código"
        .ColWidth(4) = 1000
        .ColAlignment(4) = flexAlignRightCenter
        
        .TextMatrix(0, 5) = "Parcela"
        .ColWidth(5) = 700
        .ColAlignment(5) = flexAlignRightCenter
        
        .TextMatrix(0, 6) = "Tipo"
        .ColWidth(6) = 700
        .ColAlignment(6) = flexAlignLeftCenter
        
        .TextMatrix(0, 7) = "Empresa"
        .ColWidth(7) = 1200
        .ColAlignment(7) = flexAlignLeftCenter
        
        .TextMatrix(0, 8) = "Descrição"
        .ColWidth(8) = 1500
        .ColAlignment(8) = flexAlignLeftCenter
        
        .TextMatrix(0, 9) = "C.C."
        .ColWidth(9) = 400
        .ColAlignment(9) = flexAlignRightCenter
        
        .TextMatrix(0, 10) = "Dt.Emissão"
        .ColWidth(10) = 1200
        .ColAlignment(10) = flexAlignRightCenter
        
        .TextMatrix(0, 11) = "Dt.Vencimento"
        .ColWidth(11) = 1200
        .ColAlignment(11) = flexAlignRightCenter
        
        .TextMatrix(0, 12) = "Dt.Pagamento"
        .ColWidth(12) = 1200
        .ColAlignment(12) = flexAlignRightCenter
        
        .TextMatrix(0, 13) = "Dt.Liberação"
        .ColWidth(13) = 1200
        .ColAlignment(13) = flexAlignRightCenter
        
        .TextMatrix(0, 14) = "Dias Atraso"
        .ColWidth(14) = 900
        .ColAlignment(14) = flexAlignRightCenter
        
        .TextMatrix(0, 15) = "Vl.Original"
        .ColWidth(15) = 1200
        .ColAlignment(15) = flexAlignRightCenter
        
        .TextMatrix(0, 16) = "Vl.Acrescimo"
        .ColWidth(16) = 1200
        .ColAlignment(16) = flexAlignRightCenter
        
        .TextMatrix(0, 17) = "Vl.Desconto"
        .ColWidth(17) = 1200
        .ColAlignment(17) = flexAlignRightCenter
        
        .TextMatrix(0, 18) = "Vl.Saldo"
        .ColWidth(18) = 1200
        .ColAlignment(18) = flexAlignRightCenter
        
        .TextMatrix(0, 19) = "% Multa"
        .ColWidth(19) = 700
        .ColAlignment(19) = flexAlignRightCenter
        
        .TextMatrix(0, 20) = "Vl.Multa"
        .ColWidth(20) = 1200
        .ColAlignment(20) = flexAlignRightCenter
        
        .TextMatrix(0, 21) = "Vl.Mora Diária"
        .ColWidth(21) = 1200
        .ColAlignment(21) = flexAlignRightCenter
        
        .TextMatrix(0, 22) = "Desc.Pontual"
        .ColWidth(22) = 1200
        .ColAlignment(22) = flexAlignRightCenter
        
        .TextMatrix(0, 23) = "Vl.Total"
        .ColWidth(23) = 1200
        .ColAlignment(23) = flexAlignRightCenter
                
        .TextMatrix(0, 24) = "P/R"
        .ColWidth(24) = 700
        .ColAlignment(24) = flexAlignRightCenter
        
        For intIndex = 0 To .Cols - 1
            .TextMatrix(1, intIndex) = ""
        Next
    End With
End Sub

Private Sub CalculaTotal(SQL As String)
    Dim Total              As Double
    Dim TotalValorOriginal As Double
    Dim TotalAcrescimo     As Double
    Dim TotalAbate         As Double
    Dim rs                 As Object
    
    Total = 0
    TotalValorOriginal = 0
    TotalAcrescimo = 0
    TotalAbate = 0
    If AbreRecordset(rs, SQL, dbOpenDynaset) = WL_OK Then
        If Not rs.EOF Then
            rs.MoveFirst
            While Not rs.EOF
                Total = Total + rs("Valor Total")
                TotalValorOriginal = TotalValorOriginal + rs("Valor Original")
                TotalAcrescimo = TotalAcrescimo + rs("Acréscimo")
                TotalAbate = TotalAbate + rs("Abatimento")
                rs.MoveNext
            Wend
        Else
            Total = 0
            TotalValorOriginal = 0
            TotalAcrescimo = 0
            TotalAbate = 0
        End If
    End If
'    lblCalcTotFinal.Caption = Format(Total, "###,###,##0.00")
'    lblTotalAbate.Caption = Format(TotalAbate, "###,###,##0.00")
'    lblTotalAcrescimo.Caption = Format(TotalAcrescimo, "###,###,##0.00")
'    lblTotalValorOriginal.Caption = Format(TotalValorOriginal, "###,###,##0.00")
End Sub

Private Function MontaSqlDuplicatas() As String
    Dim strSql     As String
    Dim strTipo    As String
    Dim strFiltro  As String
    'Protocolo Nr 102184 - Carlos Felippe Vernizze - 22/11/2010
    strSql = "SELECT 'Dup' AS Origem, Nota as cod_id, Parcela, Duplicatas.Tipo, Empresa, Descrição, Centro, Emissão, Vencimento, " & _
             " Pagamento, Liberação, [Valor Original], Acréscimo, Abatimento, VlrMul, VlrMrd, PerMul, VlrDsP , Controle, PagRec " & _
             "FROM Duplicatas LEFT JOIN Empresas ON Duplicatas.Empresa = Empresas.Apel "
    
    strFiltro = BuscaFiltro("Duplicatas")
    If Len(strFiltro) > 0 Then
        strSql = strSql & "WHERE" & strFiltro
    End If
    MontaSqlDuplicatas = strSql
End Function
Private Function MontaSqlLancamentos() As String
    Dim strSql    As String
    Dim strTipo   As String
    Dim strFiltro As String
    
    If cboTipo.SelectedItem = "Todos" Then
        strTipo = ""
    Else
        strTipo = cboTipo.SelectedItem
    End If
    'Protocolo Nr 102184 - Carlos Felippe Vernizze - 22/11/2010
    strSql = "SELECT 'Lan' AS Origem, Código as cod_id, Parcela, Lançamentos.Tipo, Empresa, Descrição, Centro, Emissão, Vencimento, " & _
             " Pagamento, Liberação, [Valor Original], Acréscimo, Abatimento, VlrMul, VlrMrd, PerMul, VlrDsP , Controle, PagRec " & _
             "FROM Lançamentos LEFT JOIN Empresas ON Lançamentos.Empresa = Empresas.Apel "
    strFiltro = BuscaFiltro("Lançamentos")
    If Len(strFiltro) > 1 Then
        strSql = strSql & "WHERE" & strFiltro
    End If
    MontaSqlLancamentos = strSql
End Function

Private Function BuscaSituacao() As String
    If optAndamento.value Then
        BuscaSituacao = "Andamento"
    ElseIf optFinalizadas.value Then
        BuscaSituacao = "Finalizadas"
    ElseIf optTodas.value Then
        BuscaSituacao = "Todas"
    End If
End Function

Private Function BuscaPagarReceber() As String
    Dim strPagarReceber As String
            
    strPagarReceber = strPagarReceber & " PagRec = 'R' AND Pagamento IS NULL"
    BuscaPagarReceber = strPagarReceber
End Function

Private Function BuscaFiltro(strTipo As String) As String
    Dim strSql                As String
    Dim strTipoGlobal         As String
    Dim strFiltraPagarReceber As String
    
    strSql = BuscaPagarReceber
    If cboTipo.SelectedItem = "Todos" Then
        strTipoGlobal = ""
    Else
        strTipoGlobal = cboTipo.SelectedItem
    End If
    If strTipo = "Duplicatas" Then
        If Len(strTipoGlobal) > 1 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Tipo = '" & strTipoGlobal & "'"
            Else
                strSql = strSql & " Duplicatas.Tipo = '" & strTipoGlobal & "'"
            End If
        End If
        'Nota Código
        If etxNotaCodigoIni.valorInteiro > 0 And etxNotaCodigoFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Nota BETWEEN " & etxNotaCodigoIni.valorInteiro & " AND " & etxNotaCodigoFin.valorInteiro
            Else
                strSql = " Duplicatas.Nota BETWEEN " & etxNotaCodigoIni.valorInteiro & " AND " & etxNotaCodigoFin.valorInteiro
            End If
        ElseIf etxNotaCodigoIni.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Nota >= " & etxNotaCodigoIni.valorInteiro
            Else
                strSql = " Duplicatas.Nota >= " & etxNotaCodigoIni.valorInteiro
            End If
        ElseIf etxNotaCodigoFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Nota <= " & etxNotaCodigoFin.valorInteiro
            Else
                strSql = " Duplicatas.Nota <= " & etxNotaCodigoFin.valorInteiro
            End If
        End If
        'Banco
        If etxBancoIni.valorInteiro > 0 And etxBancoFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Banco BETWEEN " & etxBancoIni.valorInteiro & " AND " & etxBancoFin.valorInteiro
            Else
                strSql = " Duplicatas.Banco BETWEEN " & etxBancoIni.valorInteiro & " AND " & etxBancoFin.valorInteiro
            End If
        ElseIf etxBancoIni.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Banco >= " & etxBancoIni.valorInteiro
            Else
                strSql = " Duplicatas.Banco >= " & etxBancoIni.valorInteiro
            End If
        ElseIf etxBancoFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Banco <= " & etxBancoFin.valorInteiro
            Else
                strSql = " Duplicatas.Banco <= " & etxBancoFin.valorInteiro
            End If
        End If
        'Conta
        If etxContaIni.valorInteiro > 0 And etxContaFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Conta BETWEEN " & etxContaIni.valorInteiro & " AND " & etxContaFin.valorInteiro
            Else
                strSql = " Duplicatas.Conta BETWEEN " & etxContaIni.valorInteiro & " AND " & etxContaFin.valorInteiro
            End If
        ElseIf etxContaIni.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Conta >= " & etxContaIni.valorInteiro
            Else
                strSql = " Duplicatas.Conta >= " & etxContaIni.valorInteiro
            End If
        ElseIf etxContaFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Duplicatas.Conta <= " & etxContaFin.valorInteiro
            Else
                strSql = " Duplicatas.Conta <= " & etxContaFin.valorInteiro
            End If
        End If
        If optAndamento.value Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Empresas.data_base_1 <= " & InverteData(DateTime.Date, True) & " AND Empresas.data_base_2 > " & InverteData(DateTime.Date, True)
            Else
                strSql = " Empresas.data_base_1 <= " & InverteData(DateTime.Date, True) & " AND Empresas.data_base_2 > " & InverteData(DateTime.Date, True)
            End If
        ElseIf optFinalizadas.value Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Empresas.data_base_2 <= " & InverteData(DateTime.Date, True)
            Else
                strSql = " Empresas.data_base_2 <= " & InverteData(DateTime.Date, True)
            End If
        End If
    Else
        If Len(strTipoGlobal) > 1 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Tipo = '" & strTipoGlobal & "'"
            Else
                strSql = " Lançamentos.Tipo = '" & strTipoGlobal & "'"
            End If
        End If
        'Nota Código
        If etxNotaCodigoIni.valorInteiro > 0 And etxNotaCodigoFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Código BETWEEN " & etxNotaCodigoIni.valorInteiro & " AND " & etxNotaCodigoFin.valorInteiro
            Else
                strSql = " Lançamentos.Código BETWEEN " & etxNotaCodigoIni.valorInteiro & " AND " & etxNotaCodigoFin.valorInteiro
            End If
        ElseIf etxNotaCodigoIni.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Código >= " & etxNotaCodigoIni.valorInteiro
            Else
                strSql = " Lançamentos.Código >= " & etxNotaCodigoIni.valorInteiro
            End If
        ElseIf etxNotaCodigoFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Código <= " & etxNotaCodigoFin.valorInteiro
            Else
                strSql = " Lançamentos.Código <= " & etxNotaCodigoFin.valorInteiro
            End If
        End If
        'Banco
        If etxBancoIni.valorInteiro > 0 And etxBancoFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Banco BETWEEN " & etxBancoIni.valorInteiro & " AND " & etxBancoFin.valorInteiro
            Else
                strSql = " Lançamentos.Banco BETWEEN " & etxBancoIni.valorInteiro & " AND " & etxBancoFin.valorInteiro
            End If
        ElseIf etxBancoIni.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Banco >= " & etxBancoIni.valorInteiro
            Else
                strSql = " Lançamentos.Banco >= " & etxBancoIni.valorInteiro
            End If
        ElseIf etxBancoFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Banco <= " & etxBancoFin.valorInteiro
            Else
                strSql = " Lançamentos.Banco <= " & etxBancoFin.valorInteiro
            End If
        End If
        'Conta
        If etxContaIni.valorInteiro > 0 And etxContaFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Conta BETWEEN " & etxContaIni.valorInteiro & " AND " & etxContaFin.valorInteiro
            Else
                strSql = " Lançamentos.Conta BETWEEN " & etxContaIni.valorInteiro & " AND " & etxContaFin.valorInteiro
            End If
        ElseIf etxContaIni.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Conta >= " & etxContaIni.valorInteiro
            Else
                strSql = " Lançamentos.Conta >= " & etxContaIni.valorInteiro
            End If
        ElseIf etxContaFin.valorInteiro > 0 Then
            If Len(strSql) > 1 Then
                strSql = strSql & " AND Lançamentos.Conta <= " & etxContaFin.valorInteiro
            Else
                strSql = " Lançamentos.Conta <= " & etxContaFin.valorInteiro
            End If
        End If
    End If
    strFiltraPagarReceber = BuscaPagarReceber
    If Len(strFiltraPagarReceber) > 1 Then
        strFiltraPagarReceber = strFiltraPagarReceber
    End If
    
    If Len(strSql) > 1 Then
        strSql = strSql & " AND Situação = 'Normal'"
    Else
        strSql = " Situação = 'Normal'"
    End If
    
    'Liberação
    If edtLibarecaoIni.Data > 0 And edtLibarecaoFin.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Liberação BETWEEN #" & Format(edtLibarecaoIni.Data, "mm/dd/yyyy") & "# AND #" & Format(edtLibarecaoFin.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Liberação BETWEEN #" & Format(edtLibarecaoIni.Data, "mm/dd/yyyy") & "# AND #" & Format(edtLibarecaoFin.Data, "mm/dd/yyyy") & "#"
        End If
    ElseIf edtLibarecaoIni.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Liberação >= #" & Format(edtLibarecaoIni.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Liberação >= #" & Format(edtLibarecaoIni.Data, "mm/dd/yyyy") & "#"
        End If
    ElseIf edtLibarecaoFin.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Liberação <= #" & Format(edtLibarecaoFin.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Liberação <= #" & Format(edtLibarecaoFin.Data, "mm/dd/yyyy") & "#"
        End If
    End If
    'Vencimento
    If edtVencimentoIni.Data > 0 And edtVencimentoFin.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Vencimento BETWEEN #" & Format(edtVencimentoIni.Data, "mm/dd/yyyy") & "# AND #" & Format(edtVencimentoFin.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Vencimento BETWEEN #" & Format(edtVencimentoIni.Data, "mm/dd/yyyy") & "# AND #" & Format(edtVencimentoFin.Data, "mm/dd/yyyy") & "#"
        End If
    ElseIf edtVencimentoIni.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Vencimento >= #" & Format(edtVencimentoIni.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Vencimento >= #" & Format(edtVencimentoIni.Data, "mm/dd/yyyy") & "#"
        End If
    ElseIf edtVencimentoFin.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Vencimento <= #" & Format(edtVencimentoFin.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Vencimento <= #" & Format(edtVencimentoFin.Data, "mm/dd/yyyy") & "#"
        End If
    End If
    'Emissão
    If edtEmissaoIni.Data > 0 And edtEmissaoFin.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Emissão BETWEEN #" & Format(edtEmissaoIni.Data, "mm/dd/yyyy") & "# AND #" & Format(edtEmissaoFin.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Emissão BETWEEN #" & Format(edtEmissaoIni.Data, "mm/dd/yyyy") & "# AND #" & Format(edtEmissaoFin.Data, "mm/dd/yyyy") & "#"
        End If
    ElseIf edtEmissaoIni.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Emissão >= #" & Format(edtEmissaoIni.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Emissão >= #" & Format(edtEmissaoIni.Data, "mm/dd/yyyy") & "#"
        End If
    ElseIf edtEmissaoFin.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Emissão <= #" & Format(edtEmissaoFin.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Emissão <= #" & Format(edtEmissaoFin.Data, "mm/dd/yyyy") & "#"
        End If
    End If
    'Pagamento
    If edtPagamentoIni.Data > 0 And edtPagamentoFin.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Pagamento BETWEEN #" & Format(edtPagamentoIni.Data, "mm/dd/yyyy") & "# AND #" & Format(edtPagamentoFin.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Pagamento BETWEEN #" & Format(edtPagamentoIni.Data, "mm/dd/yyyy") & "# AND #" & Format(edtPagamentoFin.Data, "mm/dd/yyyy") & "#"
        End If
    ElseIf edtPagamentoIni.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Pagamento >= #" & Format(edtPagamentoIni.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Pagamento >= #" & Format(edtPagamentoIni.Data, "mm/dd/yyyy") & "#"
        End If
    ElseIf edtPagamentoFin.Data > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Pagamento >= #" & Format(edtPagamentoFin.Data, "mm/dd/yyyy") & "#"
        Else
            strSql = " Pagamento >= #" & Format(edtPagamentoFin.Data, "mm/dd/yyyy") & "#"
        End If
    End If
    'Centro de Custo
    If etxCentroCustoIni.valorInteiro > 0 And etxCentroCustoFin.valorInteiro > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Centro BETWEEN " & etxCentroCustoIni.valorInteiro & " AND " & etxCentroCustoFin.valorInteiro
        Else
            strSql = " Centro BETWEEN " & etxCentroCustoIni.valorInteiro & " AND " & etxCentroCustoFin.valorInteiro
        End If
    ElseIf etxCentroCustoIni.valorInteiro > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Centro >= " & etxCentroCustoIni.valorInteiro
        Else
            strSql = " Centro >= " & etxCentroCustoIni.valorInteiro
        End If
    ElseIf etxCentroCustoFin.valorInteiro > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Centro <= " & etxCentroCustoFin.valorInteiro
        Else
            strSql = " Centro <= " & etxCentroCustoFin.valorInteiro
        End If
    End If
    'Valor Original
    If etxValOriginalIni.valorMoeda > 0 And etxValOriginalFin.valorMoeda > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND [Valor Original] BETWEEN " & Replace(etxValOriginalIni.valorMoeda, ",", ".") & " AND " & Replace(etxValOriginalFin.valorMoeda, ",", ".")
        Else
            strSql = " [Valor Original] BETWEEN " & Replace(etxValOriginalIni.valorMoeda, ",", ".") & " AND " & Replace(etxValOriginalFin.valorMoeda, ",", ".")
        End If
    ElseIf etxValOriginalIni.valorMoeda > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND [Valor Original] >= " & Replace(etxValOriginalIni.valorMoeda, ",", ".")
        Else
            strSql = " [Valor Original] >= " & Replace(etxValOriginalIni.valorMoeda, ",", ".")
        End If
    ElseIf etxValOriginalFin.valorMoeda > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND [Valor Original] <= " & Replace(etxValOriginalFin.valorMoeda, ",", ".")
        Else
            strSql = " [Valor Original] <= " & Replace(etxValOriginalFin.valorMoeda, ",", ".")
        End If
    End If
    'Parcela Inicial
    If etxParcelaIni.valorInteiro > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Parcela >= '" & etxParcelaIni.valorInteiro & "'"
        Else
            strSql = " Parcela = '" & etxParcelaIni.valorInteiro & "'"
        End If
    End If
    'Parcela Final
    If etxParcelaFin.valorInteiro > 0 Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Parcela <= '" & etxParcelaFin.valorInteiro & "'"
        Else
            strSql = " Parcela = '" & etxParcelaFin.valorInteiro & "'"
        End If
    End If
    'Empresa
    If etxEmpresa.valorTexto <> "" Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Empresa = '" & etxEmpresa.valorTexto & "'"
        Else
            strSql = " Empresa = '" & etxEmpresa.valorTexto & "'"
        End If
    End If
    'Nosso Numero
    If etxNossoNr.valorTexto <> "" Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND NOSNUM = '" & etxNossoNr.valorTexto & "'"
        Else
            strSql = " NOSNUM = '" & etxNossoNr.valorTexto & "'"
        End If
    End If
    'Controle
    If etxControle.valorTexto <> "" Then
        If Len(strSql) > 1 Then
            strSql = strSql & " AND Controle = '" & etxControle.valorTexto & "'"
        Else
            strSql = " Controle = '" & etxControle.valorTexto & "'"
        End If
    End If
    
'    'Projeto: #7373 - História: #6135 - Desenvolvimento: #7434 - Ivo Sousa(10/05/2013)
'    If Not chkRemTodos.value = vbChecked Then
'        If chkRemLiquidados.value = vbChecked Then
'            If chkRecebidas.value = vbChecked Then
'                If Len(strSql) > 1 Then
'                    strSql = strSql & " AND Id_carteira > 0"
'                Else
'                    strSql = strSql & " Id_carteira > 0"
'                End If
'            Else
'                If Len(strSql) > 1 Then
'                    strSql = strSql & " AND NOT [Pagamento] IS NULL AND Id_carteira > 0"
'                Else
'                    strSql = strSql & " NOT [Pagamento] IS NULL AND Id_carteira > 0"
'                End If
'            End If
'        ElseIf chkRemEnviados.value = vbChecked Then
'            If Len(strSql) > 1 Then
'                strSql = strSql & " AND Id_carteira > 0"
'            Else
'                strSql = strSql & " Id_carteira > 0"
'            End If
'        ElseIf chkRemNaoEnviados.value = vbChecked Then
'            If Len(strSql) > 1 Then
'                strSql = strSql & " AND Id_carteira = 0"
'            Else
'                strSql = strSql & " Id_carteira = 0"
'            End If
'        End If
'    End If
    
    BuscaFiltro = strSql
End Function

Private Sub preencheComboTipos()
    Dim cmd        As IDBSelectCommand
    Dim rdResult   As IDBReader
    Dim strDefault As String
    
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "Tipo"
    cmd.Table.TableName = "[Tipos Globais]"
    cmd.OrderByClause = "Tipo"
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    cboTipo.RemoveAll
    While Not rdResult.EOF
        If strDefault = "" Then
            strDefault = rdResult.GetString("Tipo")
        End If
        If rdResult.GetString("Tipo") = "Fatura" Then
            strDefault = rdResult.GetString("Tipo")
        End If
        cboTipo.AddItem rdResult.GetString("Tipo")
        rdResult.MoveNext
    Wend
    cboTipo.AddItem "Todos"
    rdResult.CloseReader
    cboTipo.SelectItem "Todos"
    Set rdResult = Nothing
    Set cmd = Nothing
End Sub

Private Function BuscaOrderBy() As String
    Dim strORDERBY As String
    
    If optNotaCodigo.value Then
        strORDERBY = " ORDER BY Nota"
    ElseIf optTipo.value Then
        strORDERBY = " ORDER BY Tipo"
    ElseIf optEmpresa.value Then
        strORDERBY = " ORDER BY Empresa"
    ElseIf optEmissao.value Then
        strORDERBY = " ORDER BY Emissão"
    ElseIf optVencimento.value Then
        strORDERBY = " ORDER BY Vencimento"
    ElseIf optLiberacao.value Then
        strORDERBY = " ORDER BY Liberação"
    ElseIf optValor.value Then
        strORDERBY = " ORDER BY [Valor Original]"
    ElseIf optControle.value Then
        strORDERBY = " ORDER BY Controle"
    End If
    BuscaOrderBy = strORDERBY
End Function

Private Sub DesabilitaSituacao()
    fraSituacao.Enabled = False
    optAndamento.Enabled = False
    optFinalizadas.Enabled = False
    optTodas.Enabled = False
End Sub

Private Sub HabilitaSituacao()
    fraSituacao.Enabled = True
    optAndamento.Enabled = True
    optFinalizadas.Enabled = True
    optTodas.Enabled = True
End Sub

Private Sub CarregaGrid()
    'Pt. 95368 - Moacir Pfau(11/11/2009)
    Dim i               As Long
    Dim strLinha        As String
    Dim curTotal        As Currency
    Dim curSaldo        As Currency
    Dim lngDiasAtraso   As Long
    Dim objEmp          As CEmpresas
    Dim strIndice1      As Double
    Dim dblPercIndice1  As Double
    Dim dblPercIndice2  As Double
    
On Error GoTo err_Handler
    If mobjCache Is Nothing Then
        Set mobjCache = New clsCache
    End If
    Call CarregaColunasGrid
    mcurSaldoTotal = 0
    mcurTotalGeral = 0
    mlngQtdTitulos = 0
    With mrsRegistros
        If Not .EOF Then
            .MoveFirst
            i = 1
            While Not .EOF
                Set objEmp = mobjCache.GetCacheEmpresa(.Fields("Empresa").value)
                If objEmp.IndiceDataBase1 <> Empty Or objEmp.IndiceDataBase2 <> Empty Then
                    If objEmp.IndiceDataBase1 <> Empty Then
                        dblPercIndice1 = SugereValorTaxas(objEmp.IndiceDataBase1)
                    End If
                    If objEmp.IndiceDataBase2 <> Empty Then
                        dblPercIndice2 = SugereValorTaxas(objEmp.IndiceDataBase2)
                    End If
                    If .Fields("Pagamento").value > 0 Then
                        If CDate(.Fields("Pagamento").value) > CDate(.Fields("Vencimento").value) Then
                            lngDiasAtraso = CDate(.Fields("Pagamento").value) - CDate(.Fields("Vencimento").value)
                        Else
                            lngDiasAtraso = 0
                        End If
                    ElseIf Date > CDate(.Fields("Vencimento").value) Then
                        lngDiasAtraso = Date - CDate(.Fields("Vencimento").value)
                    Else
                        lngDiasAtraso = 0
                    End If
                    'Pt. 95368 - Moacir Pfau(21/10/2009)
                    curSaldo = Round(.Fields("Valor Original").value - .Fields("Abatimento").value + .Fields("Acréscimo").value, 2)
                    curTotal = Round(curSaldo + (.Fields("VlrMrd").value * lngDiasAtraso) + ((.Fields("PerMul").value / 100) * curSaldo), 2)
                    strLinha = "" & Chr(vbKeyTab) & "" & _
                                    Chr(vbKeyTab) & dblPercIndice1 & " % " & _
                                    Chr(vbKeyTab) & dblPercIndice2 & " % " & _
                                    Chr(vbKeyTab) & .Fields("cod_id").value & _
                                    Chr(vbKeyTab) & .Fields("Parcela").value & _
                                    Chr(vbKeyTab) & .Fields("Tipo").value & _
                                    Chr(vbKeyTab) & .Fields("Empresa").value & _
                                    Chr(vbKeyTab) & .Fields("Descrição").value & _
                                    Chr(vbKeyTab) & .Fields("Centro").value & _
                                    Chr(vbKeyTab) & .Fields("Emissão").value & _
                                    Chr(vbKeyTab) & .Fields("Vencimento").value & _
                                    Chr(vbKeyTab) & .Fields("Pagamento").value & _
                                    Chr(vbKeyTab) & .Fields("Liberação").value & _
                                    Chr(vbKeyTab) & lngDiasAtraso & _
                                    Chr(vbKeyTab) & FormatNumber(.Fields("Valor Original").value, 2) & _
                                    Chr(vbKeyTab) & FormatNumber(.Fields("Acréscimo").value, 2) & _
                                    Chr(vbKeyTab) & FormatNumber(.Fields("Abatimento").value, 2) & _
                                    Chr(vbKeyTab) & FormatNumber(curSaldo, 2) & _
                                    Chr(vbKeyTab) & .Fields("PerMul").value & _
                                    Chr(vbKeyTab) & FormatNumber(.Fields("VlrMul").value, 2) & _
                                    Chr(vbKeyTab) & FormatNumber(.Fields("VlrMrd").value, 2) & _
                                    Chr(vbKeyTab) & FormatNumber(.Fields("VlrDsP").value, 2) & _
                                    Chr(vbKeyTab) & FormatNumber(curTotal, 2) & _
                                    Chr(vbKeyTab) & .Fields("PagRec").value
                    grdDuplLanc.AddItem (strLinha)
                    i = i + 1
                    mcurTotalGeral = mcurTotalGeral + curTotal
                    mcurSaldoTotal = mcurSaldoTotal + curSaldo
                End If
                .MoveNext
            Wend
            If grdDuplLanc.Rows > 2 Then
                If grdDuplLanc.TextMatrix(1, 1) = "" Then
                    grdDuplLanc.RemoveItem (1)
                End If
            End If
            mlngQtdTitulos = grdDuplLanc.Rows - 1
        Else
            etxVlSaldo.Clear
            etxVlTotal.Clear
            etxQtTitulo.Clear
            MsgBox "Não há registros para o filtro selecionado.", vbOKOnly, NomeModulo
            Call CarregaColunasGrid
        End If
    End With
    Set mobjCache = Nothing
    Set mrsRegistros = Nothing
    Exit Sub
err_Handler:
    MsgBox "Erro ao carregar registro: " + err.Description
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

Private Function SugereValorTaxas(strMoeda As String) As Double
    SugereValorTaxas = Round(SugereValorTaxasMoeda(strMoeda), 4)
End Function

Private Function SugereValorTaxasMoeda(strMoeda As String) As Double
    Dim selCmd   As IDBSelectCommand
    Dim rdResult As IDBReader
    
On Error GoTo Error_Handler

    SugereValorTaxasMoeda = 0
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "TOP 1 Valor"
        .Table.TableName = mstrTabelaCotacao
        .OrderByClause = "Data DESC"
            
        Call .Filter.Append("Moeda = @pMoeda")
        Call .Parameters.add(.CreateParameter("@pMoeda", strMoeda, dbFieldTypeString))
        
        Call .Filter.Append("Data <= @pData")
        Call .Parameters.add(.CreateParameter("@pData", DateTime.Now, dbFieldTypeDate))
        
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    If Not rdResult.EOF Then
        SugereValorTaxasMoeda = rdResult.GetDouble("Valor")
    End If
Error_Handler:
    Aplicacao.Disconnect
End Function

'Projeto: 100340 - Desenv.: 142889 - Ueder Budni (22/09/2016)
Private Sub RegistraLogLancDup(dblNumero As Double, strEmpresa As String, strTipo As String, lngParcela As Long, strPagRec As String, enuTabela As enuLancDup, dblReajuste As Double)
    Dim objLogLancDup   As New clsLogLancamentosDuplicatas

On Error GoTo erro
    Call objLogLancDup.SetKey(strPagRec, dblNumero, strEmpresa, strTipo, lngParcela, enuTabela)
    Call objLogLancDup.InsertMsg("Reajustado o valor da Duplicata em " & Format(dblReajuste, "##,##0.####") & " %.")
erro:
    Set objLogLancDup = Nothing
End Sub
