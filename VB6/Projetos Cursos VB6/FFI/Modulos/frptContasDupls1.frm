VERSION 5.00
Begin VB.Form frptContasDupls1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Duplicatas e Lançamentos"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9975
   Begin VB.Frame Frame 
      Height          =   765
      Index           =   3
      Left            =   30
      TabIndex        =   50
      Top             =   6540
      Width           =   9915
      Begin VB.CommandButton cmdVisualizar 
         Caption         =   "Visualizar"
         Height          =   405
         Left            =   8010
         TabIndex        =   24
         Top             =   210
         Width           =   1785
      End
   End
   Begin VB.Frame Frame 
      Height          =   6585
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   9915
      Begin VB.Frame Frame 
         Caption         =   "Filtros"
         Height          =   3555
         Index           =   2
         Left            =   30
         TabIndex        =   33
         Top             =   2940
         Width           =   9825
         Begin Fox.EBSText etxBancoI 
            Height          =   330
            Left            =   1290
            TabIndex        =   10
            Top             =   930
            Width           =   3690
            _ExtentX        =   252465
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Banco"
            TipoCriterio    =   4
            CampoDescricao  =   "Nome"
            TabelaConsulta  =   "Bancos"
            TamanhoDescricao=   2500
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
         Begin Fox.EBSText etxContaI 
            Height          =   330
            Left            =   1290
            TabIndex        =   12
            Top             =   1320
            Width           =   3690
            _ExtentX        =   252465
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Contas"
            TamanhoDescricao=   2500
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
         Begin Fox.EBSText etxCentroCustoI 
            Height          =   330
            Left            =   1290
            TabIndex        =   14
            Top             =   1680
            Width           =   3690
            _ExtentX        =   120703
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Centros"
            TamanhoDescricao=   2500
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
         Begin Fox.EBSData etxEmissaoI 
            Height          =   330
            Left            =   1290
            TabIndex        =   16
            Top             =   2040
            Width           =   1515
            _ExtentX        =   2672
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
         Begin Fox.EBSData etxVencimentoI 
            Height          =   330
            Left            =   1290
            TabIndex        =   18
            Top             =   2400
            Width           =   1515
            _ExtentX        =   2672
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
         Begin Fox.EBSData etxLiberacaoI 
            Height          =   330
            Left            =   1290
            TabIndex        =   22
            Top             =   3120
            Width           =   1515
            _ExtentX        =   2672
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
         Begin Fox.EBSData etxPagamentoI 
            Height          =   330
            Left            =   1290
            TabIndex        =   20
            Top             =   2760
            Width           =   1515
            _ExtentX        =   2672
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
         Begin Fox.EBSText etxControle 
            Height          =   330
            Left            =   1290
            TabIndex        =   9
            Top             =   570
            Width           =   2715
            _ExtentX        =   344
            _ExtentY        =   582
            Tipo            =   4
            MaxLength       =   15
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
         Begin Fox.EBSText etxEmpresa 
            Height          =   330
            Left            =   1290
            TabIndex        =   8
            Top             =   210
            Width           =   5535
            _ExtentX        =   403886
            _ExtentY        =   582
            Tipo            =   4
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Apel"
            CampoDescricao  =   "Razão"
            TabelaConsulta  =   "Empresas"
            TamanhoDescricao=   4000
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
         Begin Fox.EBSText etxBancoF 
            Height          =   330
            Left            =   5850
            TabIndex        =   11
            Top             =   960
            Width           =   3690
            _ExtentX        =   256858
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Banco"
            TipoCriterio    =   4
            CampoDescricao  =   "Nome"
            TabelaConsulta  =   "Bancos"
            TamanhoDescricao=   2500
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
         Begin Fox.EBSData etxEmissaoF 
            Height          =   330
            Left            =   3390
            TabIndex        =   17
            Top             =   2040
            Width           =   1515
            _ExtentX        =   2672
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
         Begin Fox.EBSData etxVencimentoF 
            Height          =   330
            Left            =   3390
            TabIndex        =   19
            Top             =   2400
            Width           =   1515
            _ExtentX        =   2672
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
         Begin Fox.EBSData etxLiberacaoF 
            Height          =   330
            Left            =   3390
            TabIndex        =   23
            Top             =   3120
            Width           =   1515
            _ExtentX        =   2672
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
         Begin Fox.EBSData etxPagamentoF 
            Height          =   330
            Left            =   3390
            TabIndex        =   21
            Top             =   2760
            Width           =   1515
            _ExtentX        =   2672
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
         Begin Fox.EBSText etxContaF 
            Height          =   330
            Left            =   5850
            TabIndex        =   13
            Top             =   1320
            Width           =   3690
            _ExtentX        =   256858
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Contas"
            TamanhoDescricao=   2500
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
         Begin Fox.EBSText etxCentroCustoF 
            Height          =   330
            Left            =   5850
            TabIndex        =   15
            Top             =   1680
            Width           =   3690
            _ExtentX        =   256858
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Centros"
            TamanhoDescricao=   2500
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
         Begin Fox.EBSReport ertRelatorio 
            Height          =   795
            Left            =   8970
            TabIndex        =   51
            Top             =   2640
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1402
            NomeRelatorio   =   "FOXFVF10174.ERC"
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   3030
            TabIndex        =   49
            Top             =   3195
            Width           =   90
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   3030
            TabIndex        =   48
            Top             =   2835
            Width           =   90
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   3030
            TabIndex        =   47
            Top             =   2475
            Width           =   90
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   3030
            TabIndex        =   46
            Top             =   2115
            Width           =   90
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   5280
            TabIndex        =   45
            Top             =   1755
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   5280
            TabIndex        =   44
            Top             =   1395
            Width           =   90
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   5280
            TabIndex        =   43
            Top             =   1005
            Width           =   90
         End
         Begin VB.Label lblEmpresa 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Left            =   555
            TabIndex        =   42
            Top             =   285
            Width           =   660
         End
         Begin VB.Label lblControle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Controle:"
            Height          =   195
            Left            =   585
            TabIndex        =   41
            Top             =   638
            Width           =   630
         End
         Begin VB.Label lblPagamento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Pagamento:"
            Height          =   195
            Left            =   360
            TabIndex        =   40
            Top             =   2828
            Width           =   855
         End
         Begin VB.Label lblLiberacao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Liberação:"
            Height          =   195
            Left            =   465
            TabIndex        =   39
            Top             =   3195
            Width           =   750
         End
         Begin VB.Label lblVencimento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            Height          =   195
            Left            =   330
            TabIndex        =   38
            Top             =   2475
            Width           =   885
         End
         Begin VB.Label lblEmissao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   585
            TabIndex        =   37
            Top             =   2115
            Width           =   630
         End
         Begin VB.Label lblBanco 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            Height          =   195
            Left            =   705
            TabIndex        =   36
            Top             =   1005
            Width           =   510
         End
         Begin VB.Label lblConta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
            Height          =   195
            Left            =   750
            TabIndex        =   35
            Top             =   1395
            Width           =   465
         End
         Begin VB.Label lblCentroCusto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Centro Custo:"
            Height          =   195
            Left            =   255
            TabIndex        =   34
            Top             =   1755
            Width           =   960
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Gerais"
         Height          =   2775
         Index           =   1
         Left            =   60
         TabIndex        =   25
         Top             =   120
         Width           =   9795
         Begin Fox.EBSCombo cboOrigem 
            Height          =   315
            Left            =   1560
            TabIndex        =   1
            Top             =   210
            Width           =   1815
            _ExtentX        =   3201
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
         Begin Fox.EBSCombo cboTipo 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Top             =   900
            Width           =   1815
            _ExtentX        =   3201
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
         Begin Fox.EBSText etxFormaPagto 
            Height          =   330
            Left            =   1560
            TabIndex        =   6
            Top             =   1980
            Width           =   4290
            _ExtentX        =   126153
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Nome"
            TabelaConsulta  =   "Formas de Pagamento"
            TamanhoDescricao=   2500
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
         Begin Fox.EBSCombo cboSituacao 
            Height          =   315
            Left            =   1560
            TabIndex        =   7
            Top             =   2340
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            Dados           =   "Todas;Normal;Descontada;Caução;Parcial;Em Cartório;Protestada;Em Cobrança;Jurídico;Devolvida;Cancelada"
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
         Begin Fox.EBSCombo cboTipoDoc 
            Height          =   315
            Left            =   1560
            TabIndex        =   2
            Top             =   540
            Width           =   1815
            _ExtentX        =   3201
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
         Begin Fox.EBSCombo cboConciliado 
            Height          =   315
            Left            =   1560
            TabIndex        =   4
            Top             =   1260
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            Dados           =   "Todos;Sim;Não"
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
         Begin Fox.EBSCombo cboQuebraGrupo 
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   1620
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            Dados           =   "Vencimento;Emissão;Pagamento;Liberação;Banco;Número;Empresa;Centro de Custo;Conta Financeira;Controle"
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
         Begin VB.Label lblSituacao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Situação:"
            Height          =   195
            Left            =   810
            TabIndex        =   32
            Top             =   2400
            Width           =   675
         End
         Begin VB.Label lblLancDupl 
            AutoSize        =   -1  'True
            Caption         =   "&Quebra dos Grupos:"
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   31
            Top             =   1680
            Width           =   1425
         End
         Begin VB.Label lblLancDupl 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Pagto.:"
            Height          =   195
            Index           =   30
            Left            =   270
            TabIndex        =   30
            Top             =   2048
            Width           =   1215
         End
         Begin VB.Label lblLancDupl 
            AutoSize        =   -1  'True
            Caption         =   "Conciliado:"
            Height          =   195
            Index           =   26
            Left            =   705
            TabIndex        =   29
            Top             =   1320
            Width           =   780
         End
         Begin VB.Label lblLancDupl 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Reg.:"
            Height          =   195
            Index           =   23
            Left            =   735
            TabIndex        =   28
            Top             =   960
            Width           =   750
         End
         Begin VB.Label lblLancDupl 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Doc.:"
            Height          =   195
            Index           =   1
            Left            =   735
            TabIndex        =   27
            Top             =   600
            Width           =   750
         End
         Begin VB.Label lblLancDupl 
            AutoSize        =   -1  'True
            Caption         =   "Ori&gem:"
            Height          =   195
            Index           =   0
            Left            =   945
            TabIndex        =   26
            Top             =   270
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "frptContasDupls1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdVisualizar_Click()
    Dim strRelatorio As String
    
    ertRelatorio.ClearParametro
    'If cboOrigem.SelectedItem <> "Ambos" Then
        ertRelatorio.AddParametro "VORI", cboOrigem.SelectedItem
    'End If

    'If cboTipoDoc.SelectedItem <> "Todos os Tipos" Then
        ertRelatorio.AddParametro "VTIPO", cboTipoDoc.SelectedItem
    'End If

    ertRelatorio.AddParametro "vtipreg", CStr(cboTipo.SelectedItem)

    If etxEmissaoI.IsValidDate Then
        ertRelatorio.AddParametro "vemiini", CStr(etxEmissaoI.Data)
    Else
        ertRelatorio.AddParametro "vemiini", ""
    End If

    If etxEmissaoF.IsValidDate Then
        ertRelatorio.AddParametro "vemifin", CStr(etxEmissaoF.Data)
    Else
        ertRelatorio.AddParametro "vemifin", ""
    End If

    If etxPagamentoI.IsValidDate Then
        ertRelatorio.AddParametro "vpagini", CStr(etxPagamentoI.Data)
    Else
        ertRelatorio.AddParametro "vpagini", ""
    End If

    If etxPagamentoF.IsValidDate Then
        ertRelatorio.AddParametro "vpagfin", etxPagamentoF.Data
    Else
        ertRelatorio.AddParametro "vpagfin", ""
    End If

    ertRelatorio.AddParametro "vcenini", CStr(etxCentroCustoI.valorInteiro)
    ertRelatorio.AddParametro "vcenfin", etxCentroCustoF.valorInteiro
    ertRelatorio.AddParametro "vbanini", etxBancoI.valorInteiro
    ertRelatorio.AddParametro "vbanfin", etxBancoF.valorInteiro
    ertRelatorio.AddParametro "vconini", etxContaI.valorInteiro
    ertRelatorio.AddParametro "vconfin", etxContaF.valorInteiro
    ertRelatorio.AddParametro "vapeini", etxEmpresa.valorTexto
    ertRelatorio.AddParametro "vapefin", etxEmpresa.valorTexto
    
    If cboConciliado.SelectedItem <> "Todos" Then
        ertRelatorio.AddParametro "VCON", cboConciliado.SelectedItem
    End If
    
    ertRelatorio.AddParametro "vcontrole", etxControle.valorTexto
    ertRelatorio.AddParametro "VQUEBRA", cboQuebraGrupo.SelectedItem
    ertRelatorio.AddParametro "vcodfor", etxFormaPagto.valorInteiro
    
    If cboSituacao.SelectedItem <> "Todas" Then
        ertRelatorio.AddParametro "vsit", cboSituacao.SelectedItem
    End If
    
    If etxVencimentoI.IsValidDate Then
        ertRelatorio.AddParametro "vvenini", etxVencimentoI.Data
    Else
        ertRelatorio.AddParametro "vvenini", ""
    End If
    
    If etxVencimentoF.IsValidDate Then
        ertRelatorio.AddParametro "vvenfin", CStr(etxVencimentoF.Data)
    Else
        ertRelatorio.AddParametro "vvenfin", ""
    End If
    
    If etxLiberacaoI.IsValidDate Then
        ertRelatorio.AddParametro "vlibini", CStr(etxLiberacaoI.Data)
    Else
        ertRelatorio.AddParametro "vlibini", ""
    End If
    
    If etxLiberacaoF.IsValidDate Then
        ertRelatorio.AddParametro "vlibfin", CStr(etxLiberacaoF.Data)
    Else
        ertRelatorio.AddParametro "vlibfin", ""
    End If
    
    ertRelatorio.EnterpriseId = EnterpriseId
    ertRelatorio.UserGroup = GetFieldValue("grupo", "Usuários", "usuário = '" & UserName & "'", , "")

    If ReadSettings("PARAMETROS", "app_remoto", "") <> "" Then
        strRelatorio = ReadSettings("PARAMETROS", "app_remoto", "") & "Programas\ERC\RELS\FOXFIN00216.ERC"
    Else
        strRelatorio = ReadSettings("PARAMETROS", "app_local", "") & "Programas\ERC\RELS\FOXFIN00216.ERC"
    End If
    ertRelatorio.NomeRelatorio = "FOXFIN00216.ERC"
    ertRelatorio.CaminhoConfiguracao = ArquivoConfiguracao
       
    ertRelatorio.NumeroCopias = 1
    ertRelatorio.CaminhoImpressora = "PDF Writer - bioPDF"
    ertRelatorio.EscModel = emNone
    ertRelatorio.OEMConvert = False
    'ertRelatorio.NumeroCopias = lngNrCopia
    ertRelatorio.Visualizador = ReadSettings("PARAMETROS", "app_local", "") & "Programas\fre.exe"
    ertRelatorio.ArquivoExecucao = ReadSettings("PARAMETROS", "app_local", "") & "Programas\LancamentoDuplicata.xml"
    ertRelatorio.LoginUsuario = UserName
    ertRelatorio.SenhaUsuario = GetFieldValue("senha", "Usuários", "usuário = '" & UserName & "'", , "")
    ertRelatorio.Visualizar
End Sub

Private Sub etxBancoF_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Banco, Nome " _
        & "FROM Bancos"
        PCampo "Bancos", strSql, PB_CAMPO, etxBancoF, "Banco"
    End If
End Sub

Private Sub etxBancoI_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Banco, Nome " _
        & "FROM Bancos"
        PCampo "Bancos", strSql, PB_CAMPO, etxBancoI, "Banco"
    End If
End Sub

Private Sub etxCentroCustoF_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Código, Descrição, [Data Limite],[cd_conta_contabil], [cd_centro_crd] " _
        & "FROM Centros"
        PCampo "C.Custo", strSql, PB_CAMPO, etxCentroCustoF, "Código"
    End If
End Sub

Private Sub etxCentroCustoI_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Código, Descrição, [Data Limite],[cd_conta_contabil], [cd_centro_crd] " _
        & "FROM Centros"
        PCampo "C.Custo", strSql, PB_CAMPO, etxCentroCustoI, "Código"
    End If
End Sub

Private Sub etxContaF_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Contas.Código as Conta, Contas.Descrição as [Descrição da Conta], " _
        & "Grupos.Código as Grupo, Grupos.Descrição as [Descrição do Grupo] " _
        & "FROM Grupos INNER JOIN Contas ON Grupos.Código = Contas.Grupo where Contas.Ctaati='S' " _
        & "ORDER BY Grupos.Código,Contas.Código"
        PCampo "Conta", strSql, PB_CAMPO, etxContaF, "Conta"
    End If
End Sub

Private Sub etxContaI_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Contas.Código as Conta, Contas.Descrição as [Descrição da Conta], " _
        & "Grupos.Código as Grupo, Grupos.Descrição as [Descrição do Grupo] " _
        & "FROM Grupos INNER JOIN Contas ON Grupos.Código = Contas.Grupo where Contas.Ctaati='S' " _
        & "ORDER BY Grupos.Código,Contas.Código"
        PCampo "Conta", strSql, PB_CAMPO, etxContaI, "Conta"
    End If
End Sub

Private Sub etxFormaPagto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Código, Nome, Tipo , Banco, Conta,[Tipo de Exportação],[Gerar KIF]," _
        & "[per_despesa_financeira]" _
        & "FROM [Formas de Pagamento]"
        PCampo "Forma Pagto", strSql, PB_CAMPO, etxFormaPagto, "Código"
    End If
End Sub

Private Sub etxEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Apel, Razão, Pessoa, Tipo, [CNPJ/CPF], [IEst/RG], CCM, " _
        & "Ramo, Endereço, Bairro, CEP, Cidade, Estado, " _
        & "Região, País, Fone1, Ramal1, Contato, Dpto " _
        & "FROM Empresas"
        ' Verifica a configuração para separar as empresas por tipo
        PCampo "Empresas", strSql, PB_CAMPO, etxEmpresa, "Apel"
    End If
End Sub

Private Sub etxEmpresa_LostFocus()
    If Trim(etxEmpresa.valorTexto) <> "" Then
       ' Call DemonstrarInformacaoAdicional
    End If
End Sub

Private Sub preencheComboTipo()
    Dim cmd        As IDBSelectCommand
    Dim rdResult   As IDBReader
    Dim strDefault As String
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "Tipo"
    cmd.Table.TableName = "[Tipos Globais]"
    cmd.OrderByClause = "Tipo"
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    cboTipo.RemoveAll
    While Not rdResult.EOF
        If strDefault = "" Then strDefault = rdResult.GetString("Tipo")
        If rdResult.GetString("Tipo") = "Fatura" Then
            strDefault = rdResult.GetString("Tipo")
        End If
        cboTipo.AddItem rdResult.GetString("Tipo")
        rdResult.MoveNext
    Wend
    rdResult.CloseReader
    cboTipo.SelectItem strDefault
    
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Sub

Private Function preencheCombo()
    Call preencheComboTipo
    Call preencheComboOrigem
    Call preencheComboSituacao
    Call preencheComboTipoDoc
    Call preencheComboConciliado
    Call preencheComboQuebraGrupo
   ' Call preencheComboSituacao
End Function

Private Sub Form_Load()
    Aplicacao.Connect
    preencheCombo
    Call etxEmpresa.AddConexao(Aplicacao)
    Call etxFormaPagto.AddConexao(Aplicacao)
    Call etxControle.AddConexao(Aplicacao)
    Call etxBancoI.AddConexao(Aplicacao)
    Call etxBancoF.AddConexao(Aplicacao)
    Call etxContaI.AddConexao(Aplicacao)
    Call etxContaF.AddConexao(Aplicacao)
    Call etxCentroCustoI.AddConexao(Aplicacao)
    Call etxCentroCustoF.AddConexao(Aplicacao)
    Aplicacao.Disconnect
End Sub

Private Sub preencheComboOrigem()
    Dim strDefault          As String
    Dim i                   As Integer
    Dim ArrOrigem()         As String
    
    strDefault = "Ambos"
    ArrOrigem = Split("Ambos;Duplicatas;Lançamentos", ";")
    For i = 0 To UBound(ArrOrigem)
        cboOrigem.AddItem ArrOrigem(i)
    Next
    
    cboOrigem.SelectItem strDefault
    
End Sub

Private Sub preencheComboSituacao()
    Dim strDefault          As String
    Dim i                   As Integer
    Dim ArrSituacao()       As String
    
    strDefault = "Todas"
    ArrSituacao = Split("Todas;Normal;Descontada;Caução;Parcial;Em Cartório;Protestada;Em Cobrança;Jurídico;Devolvida;Cancelada", ";")
    For i = 0 To UBound(ArrSituacao)
        cboSituacao.AddItem ArrSituacao(i)
    Next
    
    cboSituacao.SelectItem strDefault
End Sub

Private Sub preencheComboTipoDoc()
    Dim strDefault          As String
    Dim i                   As Integer
    Dim ArrTipoDoc()        As String
    
    strDefault = "Todos os Tipos"
    ArrTipoDoc = Split("A Pagar;Pagas;A Receber;Recebidas;A Receber em Atraso;A Receber e a Pagar;A Receber e Recebidas;A Pagar e Pagas;Todos os Tipos", ";")
    For i = 0 To UBound(ArrTipoDoc)
        cboTipoDoc.AddItem ArrTipoDoc(i)
    Next
    
    cboTipoDoc.SelectItem strDefault
End Sub

Private Sub preencheComboConciliado()
    Dim strDefault          As String
    Dim i                   As Integer
    Dim ArrConciliado()     As String
    
    strDefault = "Todos"
    ArrConciliado = Split("Todos;Sim;Não", ";")
    For i = 0 To UBound(ArrConciliado)
        cboConciliado.AddItem ArrConciliado(i)
    Next
    
    cboConciliado.SelectItem strDefault
End Sub

Private Sub preencheComboQuebraGrupo()
    Dim strDefault          As String
    Dim i                   As Integer
    Dim ArrQuebraGrupo()     As String
    
    strDefault = "Por Empresa"
    ArrQuebraGrupo = Split("Por Empresa;Por Banco;Por Conta;Por Controle;Por Data;Por Centro de Custo;Por Vendedor;Sem Quebra", ";")
    For i = 0 To UBound(ArrQuebraGrupo)
        cboQuebraGrupo.AddItem ArrQuebraGrupo(i)
    Next
    
    cboQuebraGrupo.SelectItem strDefault
End Sub
