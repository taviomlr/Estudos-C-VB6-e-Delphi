VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmCarteira 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carteira"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11475
   Icon            =   "frmCarteira.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   11475
   Begin VB.Frame Frame 
      Height          =   7095
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   10065
      Begin TabDlg.SSTab SSTab 
         Height          =   6855
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   12091
         _Version        =   393216
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Informações da Carteira"
         TabPicture(0)   =   "frmCarteira.frx":038A
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fraOutras"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Informações ao Banco"
         TabPicture(1)   =   "frmCarteira.frx":03A6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Config. Especial Remessa"
         TabPicture(2)   =   "frmCarteira.frx":03C2
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Frame(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Configuração Especial"
         TabPicture(3)   =   "frmCarteira.frx":03DE
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame(4)"
         Tab(3).ControlCount=   1
         Begin VB.Frame Frame 
            Height          =   6405
            Index           =   4
            Left            =   -74940
            TabIndex        =   89
            Top             =   345
            Width           =   9735
            Begin VB.Frame Frame 
               Caption         =   "Emissão Boleto"
               Height          =   3315
               Index           =   6
               Left            =   90
               TabIndex        =   93
               Top             =   150
               Width           =   9555
               Begin VB.TextBox etxHTMLReciboPersonalizado 
                  Height          =   2505
                  Left            =   1740
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   94
                  Top             =   720
                  Width           =   7695
               End
               Begin Fox.EBSText etxIdentificacaoCedente 
                  Height          =   330
                  Left            =   1740
                  TabIndex        =   95
                  Top             =   240
                  Width           =   3405
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   250
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
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Identificação Cedente"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   97
                  Top             =   308
                  Width           =   1560
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "HTML Personalizado"
                  Height          =   195
                  Left            =   165
                  TabIndex        =   96
                  Top             =   810
                  Width           =   1485
               End
            End
            Begin VB.Frame Frame 
               Caption         =   "Arquivo Retorno"
               Height          =   945
               Index           =   5
               Left            =   90
               TabIndex        =   90
               Top             =   3480
               Width           =   9555
               Begin Fox.EBSCombo cboDataRetorno 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   91
                  Top             =   480
                  Width           =   3030
                  _ExtentX        =   5345
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
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Realizar a baixa pela :"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   92
                  Top             =   270
                  Width           =   1560
               End
            End
         End
         Begin VB.Frame Frame 
            Height          =   6405
            Index           =   2
            Left            =   60
            TabIndex        =   87
            Top             =   345
            Width           =   9735
            Begin VB.Frame Frame 
               Caption         =   "Arquivo Remessa"
               Height          =   6165
               Index           =   3
               Left            =   90
               TabIndex        =   88
               Top             =   150
               Width           =   9555
               Begin VB.CheckBox chkSeqRemessaNrDoc 
                  Caption         =   "Gera Sequencial de Remessa no Número do Documento"
                  Height          =   195
                  Left            =   3210
                  TabIndex        =   115
                  Top             =   1620
                  Width           =   4485
               End
               Begin VB.CheckBox chkUtilizaNumeroControle 
                  Caption         =   "Utiliza Número Controle"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   114
                  Top             =   1620
                  Width           =   3075
               End
               Begin VB.Frame Frame 
                  Caption         =   "Campos Especiais"
                  Height          =   2265
                  Index           =   7
                  Left            =   90
                  TabIndex        =   104
                  Top             =   3810
                  Width           =   9375
                  Begin VB.Frame Frame2 
                     Height          =   2055
                     Left            =   7950
                     TabIndex        =   108
                     Top             =   120
                     Width           =   1350
                     Begin VB.CommandButton cmdCENovo 
                        Caption         =   "&Novo"
                        Height          =   375
                        Left            =   90
                        TabIndex        =   112
                        Top             =   180
                        Width           =   1185
                     End
                     Begin VB.CommandButton cmdCEIncluir 
                        Caption         =   "&Incluir"
                        Height          =   375
                        Left            =   90
                        TabIndex        =   111
                        Top             =   570
                        Width           =   1185
                     End
                     Begin VB.CommandButton cmdCEExcluir 
                        Caption         =   "&Excluir"
                        Height          =   375
                        Left            =   90
                        TabIndex        =   110
                        Top             =   960
                        Width           =   1185
                     End
                     Begin VB.CommandButton cmdCECancelar 
                        Caption         =   "&Cancelar"
                        Height          =   375
                        Left            =   90
                        TabIndex        =   109
                        Top             =   1350
                        Width           =   1185
                     End
                  End
                  Begin Fox.EBSCombo cboCpEspNome 
                     Height          =   315
                     Left            =   600
                     TabIndex        =   75
                     Top             =   390
                     Width           =   3120
                     _ExtentX        =   5503
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
                  Begin Fox.EBSText etxCpEspValor 
                     Height          =   330
                     Left            =   4770
                     TabIndex        =   76
                     Top             =   390
                     Width           =   3120
                     _ExtentX        =   265
                     _ExtentY        =   582
                     Tipo            =   4
                     TipoTexto       =   0
                     MaxLength       =   30
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
                  Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgCpEspeciais 
                     Height          =   1365
                     Left            =   90
                     TabIndex        =   107
                     Top             =   780
                     Width           =   7785
                     _ExtentX        =   13732
                     _ExtentY        =   2408
                     _Version        =   393216
                     _NumberOfBands  =   1
                     _Band(0).Cols   =   2
                  End
                  Begin VB.Label Label30 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Valor"
                     Height          =   195
                     Left            =   4335
                     TabIndex        =   106
                     Top             =   465
                     Width           =   360
                  End
                  Begin VB.Label Label29 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Nome"
                     Height          =   195
                     Left            =   105
                     TabIndex        =   105
                     Top             =   450
                     Width           =   420
                  End
               End
               Begin VB.CheckBox chkNaoGerarRegistroRodape2 
                  Caption         =   "Não Gerar Registro Rodapé 2 "
                  Height          =   195
                  Left            =   6630
                  TabIndex        =   65
                  Top             =   1230
                  Width           =   2715
               End
               Begin VB.CheckBox chkNaoGerarRegistroRodape1 
                  Caption         =   "Não Gerar Registro Rodapé 1 "
                  Height          =   195
                  Left            =   3210
                  TabIndex        =   64
                  Top             =   1230
                  Width           =   3075
               End
               Begin VB.CheckBox chkNaoGerarRegistroDetalhe5 
                  Caption         =   "Não Gerar Registro Detalhe 5"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   63
                  Top             =   1230
                  Width           =   3075
               End
               Begin VB.CheckBox chkNaoGerarRegistroDetalhe4 
                  Caption         =   "Não Gerar Registro Detalhe 4"
                  Height          =   195
                  Left            =   6630
                  TabIndex        =   62
                  Top             =   780
                  Width           =   2775
               End
               Begin VB.CheckBox chkNaoGerarRegistroDetalhe3 
                  Caption         =   "Não Gerar Registro Detalhe 3"
                  Height          =   195
                  Left            =   3210
                  TabIndex        =   61
                  Top             =   780
                  Width           =   3075
               End
               Begin VB.CheckBox chkNaoGerarRegistroDetalhe2 
                  Caption         =   "Não Gerar Registro Detalhe 2"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   60
                  Top             =   780
                  Width           =   3075
               End
               Begin VB.CheckBox chkNaoGerarRegistroDetalhe1 
                  Caption         =   "Não Gerar Registro Detalhe 1"
                  Height          =   195
                  Left            =   6630
                  TabIndex        =   59
                  Top             =   360
                  Width           =   2745
               End
               Begin VB.CheckBox chkNaoGerarRegistroCabecalho2 
                  Caption         =   "Não Gerar Registro Cabeçalho 2"
                  Height          =   195
                  Left            =   3210
                  TabIndex        =   58
                  Top             =   360
                  Width           =   3075
               End
               Begin VB.CheckBox chkNaoGerarRegistroCabecalho1 
                  Caption         =   "Não Gerar Registro Cabeçalho 1"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   57
                  Top             =   360
                  Width           =   3075
               End
               Begin Fox.EBSText etxTipoImpressao 
                  Height          =   330
                  Left            =   7740
                  TabIndex        =   74
                  Top             =   3030
                  Width           =   1695
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   2
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
               Begin Fox.EBSText etxBairroSacado 
                  Height          =   330
                  Left            =   1650
                  TabIndex        =   72
                  Top             =   3030
                  Width           =   4155
                  _ExtentX        =   582
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   100
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
               Begin Fox.EBSText etxCodigoPracaSacado 
                  Height          =   330
                  Left            =   1650
                  TabIndex        =   73
                  Top             =   3390
                  Width           =   4155
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   100
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
               Begin Fox.EBSText etxInstrucaoCobranca1 
                  Height          =   330
                  Left            =   1665
                  TabIndex        =   66
                  Top             =   1950
                  Width           =   2805
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   30
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
               Begin Fox.EBSText etxInstrucaoCobranca2 
                  Height          =   330
                  Left            =   1665
                  TabIndex        =   68
                  Top             =   2310
                  Width           =   2805
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   30
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
               Begin Fox.EBSText etxInstrucaoCobranca3 
                  Height          =   330
                  Left            =   1665
                  TabIndex        =   70
                  Top             =   2670
                  Width           =   2805
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   30
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
               Begin Fox.EBSText etxValorInstrucaoCobranca1 
                  Height          =   330
                  Left            =   6600
                  TabIndex        =   67
                  Top             =   1950
                  Width           =   2835
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   30
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
               Begin Fox.EBSText etxValorInstrucaoCobranca2 
                  Height          =   330
                  Left            =   6600
                  TabIndex        =   69
                  Top             =   2310
                  Width           =   2835
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   30
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
               Begin Fox.EBSText etxValorInstrucaoCobranca3 
                  Height          =   330
                  Left            =   6600
                  TabIndex        =   71
                  Top             =   2700
                  Width           =   2835
                  _ExtentX        =   265
                  _ExtentY        =   582
                  Tipo            =   4
                  TipoTexto       =   0
                  MaxLength       =   30
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
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Valor Instrução Cobrança 3"
                  Height          =   195
                  Left            =   4620
                  TabIndex        =   103
                  Top             =   2745
                  Width           =   1935
               End
               Begin VB.Label Label27 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Valor Instrução Cobrança 2"
                  Height          =   195
                  Left            =   4620
                  TabIndex        =   102
                  Top             =   2385
                  Width           =   1935
               End
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Valor Instrução Cobrança 1"
                  Height          =   195
                  Left            =   4620
                  TabIndex        =   101
                  Top             =   2025
                  Width           =   1935
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Instrução Cobrança 3"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   100
                  Top             =   2745
                  Width           =   1530
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Instrução Cobrança 2"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   99
                  Top             =   2385
                  Width           =   1530
               End
               Begin VB.Label InstrucaoCobranca 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Instrução Cobrança 1"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   98
                  Top             =   2025
                  Width           =   1530
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Código Praca Sacado"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   78
                  Top             =   3465
                  Width           =   1560
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Bairro do Sacado"
                  Height          =   195
                  Left            =   390
                  TabIndex        =   77
                  Top             =   3105
                  Width           =   1230
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo Impressão"
                  Height          =   195
                  Left            =   6600
                  TabIndex        =   79
                  Top             =   3105
                  Width           =   1080
               End
            End
         End
         Begin VB.Frame Frame 
            Height          =   6405
            Index           =   1
            Left            =   -74940
            TabIndex        =   2
            Top             =   345
            Width           =   9735
            Begin VB.TextBox etxDemonstativo 
               Height          =   1725
               Left            =   2010
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   39
               Top             =   240
               Width           =   7575
            End
            Begin VB.TextBox etxInstrucao 
               Height          =   1725
               Left            =   2010
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   41
               Top             =   2040
               Width           =   7575
            End
            Begin VB.CheckBox chkNNBanco 
               Caption         =   "Nosso Número Gerado pelo Banco"
               Height          =   195
               Left            =   2025
               TabIndex        =   55
               Top             =   6120
               Width           =   2775
            End
            Begin VB.CheckBox chkBancoEmiteBoleto 
               Caption         =   "Banco Emite o Boleto Bancário"
               Height          =   195
               Left            =   5430
               TabIndex        =   56
               Top             =   6120
               Width           =   2685
            End
            Begin Fox.EBSText etxSequencialRemessa 
               Height          =   330
               Left            =   2010
               TabIndex        =   48
               Top             =   4950
               Width           =   1575
               _ExtentX        =   265
               _ExtentY        =   582
               TipoTexto       =   0
               MaxLength       =   7
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
            Begin Fox.EBSCombo cboEspecieDoc 
               Height          =   315
               Left            =   2010
               TabIndex        =   50
               Top             =   5310
               Width           =   1575
               _ExtentX        =   2778
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
            Begin Fox.EBSText etxDiasProtesto 
               Height          =   330
               Left            =   2040
               TabIndex        =   54
               Top             =   5670
               Width           =   1575
               _ExtentX        =   265
               _ExtentY        =   582
               TipoTexto       =   0
               MaxLength       =   2
               TipoCriterio    =   3
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
            Begin Fox.EBSText etxPerMulta 
               Height          =   330
               Left            =   2010
               TabIndex        =   44
               Top             =   4200
               Width           =   1575
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   1
               CasasDecimais   =   2
               TipoTexto       =   0
               MaxLength       =   6
               TipoCriterio    =   5
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
            Begin Fox.EBSText etxPerMora 
               Height          =   330
               Left            =   2010
               TabIndex        =   46
               Top             =   4560
               Width           =   1575
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   1
               CasasDecimais   =   2
               TipoTexto       =   0
               MaxLength       =   6
               TipoCriterio    =   5
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
            Begin Fox.EBSText etxOutraEspecie 
               Height          =   330
               Left            =   6360
               TabIndex        =   52
               Top             =   5310
               Width           =   1560
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               TipoTexto       =   0
               MaxLength       =   3
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
            Begin Fox.EBSText etxLocalPagamento 
               Height          =   330
               Left            =   2010
               TabIndex        =   42
               Top             =   3810
               Width           =   7590
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               TipoTexto       =   0
               MaxLength       =   100
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
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Local Pagamento"
               Height          =   195
               Left            =   690
               TabIndex        =   113
               Top             =   3870
               Width           =   1245
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Outra Espécie"
               Height          =   195
               Left            =   5280
               TabIndex        =   51
               Top             =   5370
               Width           =   1005
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Percentual Mora"
               Height          =   195
               Left            =   795
               TabIndex        =   45
               Top             =   4635
               Width           =   1170
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Percentual Multa"
               Height          =   195
               Left            =   765
               TabIndex        =   43
               Top             =   4275
               Width           =   1200
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Instruções ao Caixa"
               Height          =   195
               Left            =   570
               TabIndex        =   40
               Top             =   2040
               Width           =   1395
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Demonstrativo"
               Height          =   195
               Left            =   945
               TabIndex        =   38
               Top             =   240
               Width           =   1020
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Sequencial de Remessa"
               Height          =   195
               Left            =   240
               TabIndex        =   47
               Top             =   5025
               Width           =   1725
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Espécie Doc."
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
               Left            =   810
               TabIndex        =   49
               Top             =   5370
               Width           =   1155
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Dias para Protesto"
               Height          =   195
               Left            =   660
               TabIndex        =   53
               Top             =   5745
               Width           =   1305
            End
         End
         Begin VB.Frame fraOutras 
            Height          =   6405
            Left            =   -74940
            TabIndex        =   3
            Top             =   345
            Width           =   9735
            Begin VB.Frame Frame1 
               Caption         =   "Layout"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   915
               Left            =   210
               TabIndex        =   30
               Top             =   5370
               Width           =   9435
               Begin Fox.EBSCombo cboLayoutBoleto 
                  Height          =   315
                  Left            =   105
                  TabIndex        =   34
                  Top             =   450
                  Width           =   3030
                  _ExtentX        =   5345
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
               Begin Fox.EBSCombo cboLayoutRemessa 
                  Height          =   315
                  Left            =   3180
                  TabIndex        =   35
                  Top             =   450
                  Width           =   3030
                  _ExtentX        =   5345
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
               Begin Fox.EBSCombo cboLayoutRetorno 
                  Height          =   315
                  Left            =   6270
                  TabIndex        =   36
                  Top             =   450
                  Width           =   3030
                  _ExtentX        =   5345
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
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Boleto"
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
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Width           =   555
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Remessa"
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
                  Left            =   3180
                  TabIndex        =   32
                  Top             =   240
                  Width           =   780
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Retorno"
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
                  Left            =   6270
                  TabIndex        =   33
                  Top             =   240
                  Width           =   690
               End
            End
            Begin Fox.EBSText etxDesc 
               Height          =   330
               Left            =   2010
               TabIndex        =   7
               Top             =   630
               Width           =   7575
               _ExtentX        =   2328
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   70
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
            Begin Fox.EBSArquivo etxArquivoLicenca 
               Height          =   330
               Left            =   2010
               TabIndex        =   9
               Top             =   1020
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   582
               TipoTratamento  =   2
               Filter          =   ""
            End
            Begin Fox.EBSText etxOutro1 
               Height          =   330
               Left            =   2010
               TabIndex        =   13
               Top             =   1800
               Width           =   2295
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
            Begin Fox.EBSText etxOutro2 
               Height          =   330
               Left            =   2010
               TabIndex        =   15
               Top             =   2190
               Width           =   2295
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
            Begin Fox.EBSText etxIdCarteira 
               Height          =   330
               Left            =   2010
               TabIndex        =   5
               Top             =   240
               Width           =   1575
               _ExtentX        =   265
               _ExtentY        =   582
               TipoTexto       =   0
               MaxLength       =   6
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxCendente 
               Height          =   330
               Left            =   2010
               TabIndex        =   11
               Top             =   1410
               Width           =   2295
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               TipoTexto       =   0
               MaxLength       =   70
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
            Begin Fox.EBSText etxMargem 
               Height          =   330
               Left            =   2010
               TabIndex        =   29
               Top             =   4890
               Width           =   1575
               _ExtentX        =   265
               _ExtentY        =   582
               TipoTexto       =   0
               MaxLength       =   3
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
            Begin Fox.EBSText etxInicioNN 
               Height          =   330
               Left            =   2010
               TabIndex        =   17
               Top             =   2580
               Width           =   2295
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               TipoTexto       =   0
               MaxLength       =   30
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
            Begin Fox.EBSText etxFimNN 
               Height          =   330
               Left            =   2010
               TabIndex        =   19
               Top             =   2970
               Width           =   2295
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               TipoTexto       =   0
               MaxLength       =   30
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
            Begin Fox.EBSText etxProximoNN 
               Height          =   330
               Left            =   2010
               TabIndex        =   21
               Top             =   3360
               Width           =   2295
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               TipoTexto       =   0
               MaxLength       =   30
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
            Begin Fox.EBSArquivo etxCaminhoRemessa 
               Height          =   330
               Left            =   2010
               TabIndex        =   23
               Top             =   3750
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   582
               Filter          =   ""
            End
            Begin Fox.EBSArquivo etxCaminhoRetorno 
               Height          =   330
               Left            =   2010
               TabIndex        =   25
               Top             =   4140
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   582
               Filter          =   ""
            End
            Begin Fox.EBSArquivo etxLogoEmpresa 
               Height          =   330
               Left            =   2010
               TabIndex        =   27
               Top             =   4500
               Width           =   7575
               _ExtentX        =   13361
               _ExtentY        =   582
               TipoTratamento  =   2
               Filter          =   ""
            End
            Begin VB.Label lblOutro1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Outro Dado Conf 1"
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
               Left            =   345
               TabIndex        =   12
               Top             =   1875
               Width           =   1605
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Carteira"
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
               Left            =   1275
               TabIndex        =   4
               Top             =   315
               Width           =   675
            End
            Begin VB.Label lbllicenca 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Arquivo de Licença"
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
               Left            =   285
               TabIndex        =   8
               Top             =   1095
               Width           =   1665
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Margem Superior Boleto"
               Height          =   195
               Left            =   255
               TabIndex        =   28
               Top             =   4965
               Width           =   1695
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Descrição"
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
               Left            =   1080
               TabIndex        =   6
               Top             =   705
               Width           =   870
            End
            Begin VB.Label lblCendente 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Código Cedente"
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
               Left            =   585
               TabIndex        =   10
               Top             =   1485
               Width           =   1365
            End
            Begin VB.Label lblInicioNN 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Início Nos. Número"
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
               Left            =   285
               TabIndex        =   16
               Top             =   2655
               Width           =   1665
            End
            Begin VB.Label lblFimNN 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Fim Nos. Número"
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
               TabIndex        =   18
               Top             =   3045
               Width           =   1455
            End
            Begin VB.Label lblProximoNN 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Atual Nos. Número"
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
               Left            =   345
               TabIndex        =   20
               Top             =   3435
               Width           =   1605
            End
            Begin VB.Label lblOutro2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Outro Dado Conf 2"
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
               Left            =   345
               TabIndex        =   14
               Top             =   2265
               Width           =   1605
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Arq. Remessa (Padrão)"
               Height          =   195
               Left            =   315
               TabIndex        =   22
               Top             =   3825
               Width           =   1635
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Arq. Retorno (Padrão)"
               Height          =   195
               Left            =   405
               TabIndex        =   24
               Top             =   4215
               Width           =   1545
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logo Empresa"
               Height          =   195
               Left            =   930
               TabIndex        =   26
               Top             =   4575
               Width           =   1020
            End
         End
      End
   End
   Begin VB.Frame fraBotoes 
      Height          =   7095
      Left            =   10110
      TabIndex        =   37
      Top             =   -60
      Width           =   1350
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   84
         Top             =   1770
         Width           =   1185
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   85
         Top             =   2160
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   82
         Top             =   1350
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   86
         Top             =   2550
         Width           =   1185
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   83
         Top             =   960
         Width           =   1185
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   80
         Top             =   570
         Width           =   1185
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   81
         Top             =   180
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmCarteira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mobjCarteiraDAO                 As clsCarteiraDAO
Private mobjCarteira                    As clsCarteira
Private mobjCobreBem                    As CobreBemX.ContaCorrente
Private mlngEnterpriseId                As Long
Private mlngCdEstabelecimento           As Long
Private mblnAlterando                   As Boolean
'Projeto: #4350 - História: # - Desenvolvimento# - Moacir Pfau(09/04/2013)
Private Enum eData_baixa_retorno
    DataCredito = 0
    DataOcorrencia = 1
End Enum
Private mobjCpEspecial                  As clsCarteiraCpEspecial
Private mcolCarteiraCpEspecial          As clscolCarteiraCpEspecial
Private mblnAlterandoCpEspecial         As Boolean

Private Const strCpEspeciais = "campo=#vazio;label=;tamanho=150|" & _
                               "campo=CpEspNome;label=Nome;tamanho=3700|" & _
                               "campo=CpEspvalor;label=Valor;tamanho=3700"


'Pt. 96589 - Moacir Pfau(05/02/2010)
Private Sub chkNNBanco_Click()
    Dim blnHabilita             As Boolean
    
    If chkNNBanco.value = 0 Then
        blnHabilita = True
    Else
        blnHabilita = False
    End If

    etxInicioNN.Enabled = blnHabilita
    etxFimNN.Enabled = blnHabilita
    etxProximoNN.Enabled = blnHabilita
    lblInicioNN.Enabled = blnHabilita
    lblFimNN.Enabled = blnHabilita
    lblProximoNN.Enabled = blnHabilita
    
    If Not blnHabilita Then
        etxInicioNN.Clear:         etxFimNN.Clear:              etxProximoNN.Clear
    End If
End Sub

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

'CARREGA O ENTERPRISE_ID E ESTABELECIMENTO.
Private Sub fLoadEnterprise_estabelecimento()
        mlngEnterpriseId = GetFieldValue("enterprise_id", "Usuários", "usuário = '" & UserName & "'")
        mlngCdEstabelecimento = GetFieldValue("cd_estabelecimento", "Usuários", "usuário = '" & UserName & "'")
End Sub

Private Sub cmdCancelar_Click()
    Call LibProc(WL_CANCELAR)
End Sub

Private Sub cmdCECancelar_Click()
    Call LimpaCamposCpEspecial
    mblnAlterandoCpEspecial = False
    cmdCEExcluir.Enabled = False
End Sub

Private Sub cmdCEExcluir_Click()
    Call CarregaClasseCpEspecial
    Call mcolCarteiraCpEspecial.Remove(mobjCpEspecial)
    Call CarregaGridCpEspecial
    Call LimpaCamposCpEspecial
    cmdCEExcluir.Enabled = False
End Sub

Private Sub cmdCEIncluir_Click()
    If mcolCarteiraCpEspecial Is Nothing Then
        Set mcolCarteiraCpEspecial = New clscolCarteiraCpEspecial
    End If
    If mblnAlterandoCpEspecial Then
        Call fAtualizarCpEspecial
    Else
        Call fAdicionarCpEspecial
    End If
    cmdCEExcluir.Enabled = False
End Sub

Private Function fAdicionarCpEspecial() As Boolean
    Dim blnAdicionar        As Boolean

    Call CarregaClasseCpEspecial
    If ValidaCamposCpEspecial Then
        Call mcolCarteiraCpEspecial.add(mobjCpEspecial)
    End If
    CarregaGridCpEspecial
    LimpaCamposCpEspecial
    fAdicionarCpEspecial = True
End Function

Private Function fAtualizarCpEspecial() As Boolean
    Call CarregaClasseCpEspecial
    Call mcolCarteiraCpEspecial.update(mobjCpEspecial)
    CarregaGridCpEspecial
    LimpaCamposCpEspecial
    fAtualizarCpEspecial = True
End Function

Private Sub LimpaCamposCpEspecial()
    cboCpEspNome.Clear
    etxCpEspValor.Clear
    mblnAlterandoCpEspecial = False
    cmdCEExcluir.Enabled = False
End Sub

Private Sub CarregaRegistroCpEspecial()
    With mobjCpEspecial
        cboCpEspNome.SelectItem .CpEspNome
        etxCpEspValor.valorTexto = .CpEspValor
    End With
    mblnAlterandoCpEspecial = True
End Sub

Private Sub CarregaGridCpEspecial()
    fgCpEspeciais.Clear
    If mcolCarteiraCpEspecial Is Nothing Then
        Call CarregaHFlexGrid(fgCpEspeciais, Nothing, strCpEspeciais)
    Else
        If mcolCarteiraCpEspecial.Count = 0 Then
            Call CarregaHFlexGrid(fgCpEspeciais, Nothing, strCpEspeciais)
        Else
            mcolCarteiraCpEspecial.MoveFirst
            Call CarregaHFlexGrid(fgCpEspeciais, , strCpEspeciais, , , mcolCarteiraCpEspecial)
        End If
    End If
End Sub

Private Sub CarregaClasseCpEspecial()
    
    Set mobjCpEspecial = New clsCarteiraCpEspecial
    With mobjCpEspecial
        .CpEspNome = cboCpEspNome.SelectedItem
        .CpEspValor = etxCpEspValor.valorTexto
    End With
End Sub

Private Function ValidaCamposCpEspecial() As Boolean
    Dim strMensagem                 As String
    
    If Trim(mobjCpEspecial.CpEspNome) = "" Then
        strMensagem = strMensagem & "Favor selecionar um item no campo 'Nome'." & vbCrLf
    End If
    
    If Trim(mobjCpEspecial.CpEspValor) = "" Then
        strMensagem = strMensagem & "Favor preencher o campo 'Valor'." & vbCrLf
    End If
    
    If Not mcolCarteiraCpEspecial Is Nothing Then
        If mcolCarteiraCpEspecial.Find(mobjCpEspecial) > 0 Then
            strMensagem = strMensagem & "Registro já cadastrado, favor informar outro." & vbCrLf
        End If
    End If
    If strMensagem = "" Then
        ValidaCamposCpEspecial = True
    Else
        MsgBox strMensagem, vbInformation, NomeModulo
    End If
End Function

Private Sub cmdCENovo_Click()
    Call LimpaCamposCpEspecial
End Sub

Private Sub cmdExcluir_Click()
        Call LibProc(WL_DELETAR)
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
End Sub

Private Sub cmdPesquisar_Click()
    Call LibProc(WL_PESQUISAR)
End Sub

Private Sub cmdNovo_Click()
    Call LibProc(WL_NOVO)
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
End Sub


Private Sub etxArquivoLicenca_LostFocus()
    If Trim(Len(etxArquivoLicenca.Valor)) > 0 Then
        fCarregaCobrebem
    End If
End Sub

Private Sub etxDemonstativo_GotFocus()
    SSTab.Tab = 1
End Sub


Private Sub fgCpEspeciais_DblClick()
    If itemSelecionado(1) <> "" Then
        Set mobjCpEspecial = mcolCarteiraCpEspecial.GetItem((itemSelecionado(1)))
        If Not mobjCpEspecial Is Nothing Then
            Call CarregaRegistroCpEspecial
        End If
        cmdCEExcluir.Enabled = True
    End If
End Sub

Private Function itemSelecionado(col As Integer) As String
    With fgCpEspeciais
       If .Row > 0 Then
          If .TextMatrix(.Row, col) <> "" Then
            itemSelecionado = .TextMatrix(.Row, col)
          Else
            itemSelecionado = ""
          End If
       End If
    End With
End Function

Private Sub Form_Load()
    Aplicacao.Connect
    Set mobjCarteira = New clsCarteira
    Set mobjCarteiraDAO = New clsCarteiraDAO
    mblnAlterando = False
    fBotaoNovo
    Call CenterForm(Me)
    Call mobjCarteiraDAO.init(Aplicacao)
    fLimpaCampos
    fLoadEnterprise_estabelecimento
    'Projeto: #4350 - História: # - Desenvolvimento# - Moacir Pfau(09/04/2013)
    cboDataRetorno.AddItem "Data da Ocorrência"
    cboDataRetorno.AddItem "Data do Crédito"
    cboDataRetorno.SelectItem "Data do Crédito"
    SSTab.Tab = 0
    Call CarregarComboCamposEmpeciaisNome
    Call CarregaGridCpEspecial
End Sub

Private Sub CarregarComboCamposEmpeciaisNome()
    Dim colecao As colCamposEspeciais
    Dim dao As CamposEspeciaisDAO
        
    Set dao = New CamposEspeciaisDAO
    dao.Initialize
    
    Set colecao = dao.Carregar
    
    'Vinicius Elyseu(30/05/2016) - Demanda: #120997
    If Not IsNothing(colecao) Then
        colecao.MoveFirst
        
        While Not colecao.EOF
            Call cboCpEspNome.AddItem(colecao.CurrentObject.Valor)
            colecao.MoveNext
        Wend
    End If
    
    dao.Terminate
    
    Set colecao = Nothing
    Set dao = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjCobreBem = Nothing
    Aplicacao.Disconnect
End Sub

Private Function fcarregaClasse()
    Set mobjCarteira = New clsCarteira
    With mobjCarteira
        .Enterprise_id = mlngEnterpriseId
        .Cd_estabelecimento = mlngCdEstabelecimento
        .Id_carteira = etxIdCarteira.valorInteiro
        .Desc_carteira = etxDesc.valorTexto
        .Codigo_cedente = etxCendente.valorTexto
        .Inicio_nosso_numero = etxInicioNN.valorTexto
        .Fim_nosso_numero = etxFimNN.valorTexto
        .Proximo_nosso_numero = etxProximoNN.valorTexto
        .Demonstrativo = etxDemonstativo.Text
        .Instrucoes_caixa = etxInstrucao.Text
        .Tipo_layout_boleto = cboLayoutBoleto.SelectedItem
        .Tipo_layout_remessa = cboLayoutRemessa.SelectedItem
        .Tipo_layout_retorno = cboLayoutRetorno.SelectedItem
        .Arquivo_licenca = etxArquivoLicenca.Valor
        .Logo_empresa = etxLogoEmpresa.Valor
        .Caminho_arquivo_remessa_padrao = etxCaminhoRemessa.Valor
        .Caminho_arquivo_retorno_padrao = etxCaminhoRetorno.Valor
        .Margem_superior_boleto = etxMargem.valorInteiro
        .Outro_dado_configuracao1 = etxOutro1.valorTexto
        .Outro_dado_configuracao2 = etxOutro2.valorTexto
        .Sequencial_remessa = etxSequencialRemessa.valorInteiro
        'Pt. 96180 - Moacir Pfau(08/12/2009)
        .Especie = cboEspecieDoc.SelectedItem
        'Pt. 96589 - Moacir Pfau(08/02/2010)
        .Banco_gera_nosso_numero = chkNNBanco.value
        'Pt. 97161 - Moacir Pfau(08/02/2010)
        .Dias_protesto = etxDiasProtesto.valorInteiro
        'pt.98446 - Fernando Paludo(26/04/2010)
        .banco_Emite_boleto = chkBancoEmiteBoleto.value
        'pt.99257 - Moacir Pfau(30/06/2010)
        .Per_multa = etxPerMulta.valorMoeda
        .Per_mora = etxPerMora.valorMoeda
        'pt.98929 - Fernando Paludo(02/08/2010)
        .Outra_especie = etxOutraEspecie.valorTexto
        'Pt. 102459 - Moacir Pfau(29/10/2010)
        .NaoGerarRegistroCabecalho1 = chkNaoGerarRegistroCabecalho1.value
        .NaoGerarRegistroCabecalho2 = chkNaoGerarRegistroCabecalho2.value
        .NaoGerarRegistroDetalhe1 = chkNaoGerarRegistroDetalhe1.value
        .NaoGerarRegistroDetalhe2 = chkNaoGerarRegistroDetalhe2.value
        .NaoGerarRegistroDetalhe3 = chkNaoGerarRegistroDetalhe3.value
        .NaoGerarRegistroDetalhe4 = chkNaoGerarRegistroDetalhe4.value
        .NaoGerarRegistroDetalhe5 = chkNaoGerarRegistroDetalhe5.value
        .NaoGerarRegistroRodape1 = chkNaoGerarRegistroRodape1.value
        .NaoGerarRegistroRodape2 = chkNaoGerarRegistroRodape2.value
        .TipoImpressao = etxTipoImpressao.valorTexto
        .BairroSacado = etxBairroSacado.valorTexto
        .CodigoPracaSacado = etxCodigoPracaSacado.valorTexto
        'Pt. 114032 - Moacir Pfau(23/02/2012)
        .IdentificacaoCedente = etxIdentificacaoCedente.valorTexto
        .HTMLReciboPersonalizado = etxHTMLReciboPersonalizado.Text
        'Projeto: #4350 - História: # - Desenvolvimento# - Moacir Pfau(09/04/2013)
        .Data_baixa_retorno = IIf(cboDataRetorno.SelectedItem = "Data da Ocorrência", eData_baixa_retorno.DataOcorrencia, eData_baixa_retorno.DataCredito)
        'Projeto: #17081 - História: # - Desenvolvimento# - Moacir Pfau(30/12/2013)
        .localPagamento = etxLocalPagamento.valorTexto
        .InstrucaoCobranca1 = etxInstrucaoCobranca1.valorTexto
        .InstrucaoCobranca2 = etxInstrucaoCobranca2.valorTexto
        .InstrucaoCobranca3 = etxInstrucaoCobranca3.valorTexto
        .ValorInstrucaoCobranca1 = etxValorInstrucaoCobranca1.valorTexto
        .ValorInstrucaoCobranca2 = etxValorInstrucaoCobranca2.valorTexto
        .ValorInstrucaoCobranca3 = etxValorInstrucaoCobranca3.valorTexto
        .UtilizaNumeroControle = chkUtilizaNumeroControle.value
        .ColCpEspeciais = mcolCarteiraCpEspecial
        'Vinicius Elyseu(06/10/2015) - Projeto: #0 - História: #0 - Desenv: #0
        .SeqRemessaNrDoc = chkSeqRemessaNrDoc.value
    End With
End Function

Private Sub fLimpaCampos()
    etxIdCarteira.Clear
    etxDesc.Clear
    etxCendente.Clear
    etxInicioNN.Clear
    etxFimNN.Clear
    etxProximoNN.Clear
    etxDemonstativo.Text = ""
    etxInstrucao.Text = ""
    cboLayoutBoleto.Clear
    cboLayoutRemessa.Clear
    cboLayoutRetorno.Clear
    etxArquivoLicenca.Clear
    etxLogoEmpresa.Clear
    etxCaminhoRemessa.Clear
    etxCaminhoRetorno.Clear
    etxOutro1.Clear
    etxOutro2.Clear
    etxMargem.valorInteiro = 15
    etxSequencialRemessa.Clear
    'Pt. 96180 - Moacir Pfau(08/12/2009)
    cboEspecieDoc.Clear
    'pt.99257 - Moacir Pfau(30/06/2010)
    etxPerMulta.Clear
    etxPerMora.Clear
    'pt.98929 - Fernando Paludo(02/08/2010)
    etxOutraEspecie.Clear
    'Pt. 102459 - Moacir Pfau(03/11/2010)
    chkNaoGerarRegistroCabecalho1.value = 0
    chkNaoGerarRegistroCabecalho2.value = 0
    chkNaoGerarRegistroDetalhe1.value = 0
    chkNaoGerarRegistroDetalhe2.value = 0
    chkNaoGerarRegistroDetalhe3.value = 0
    chkNaoGerarRegistroDetalhe4.value = 0
    chkNaoGerarRegistroDetalhe5.value = 0
    chkNaoGerarRegistroRodape1.value = 0
    chkNaoGerarRegistroRodape2.value = 0
    etxTipoImpressao.Clear
    'Pt. 106012 - Moacir Pfau(28/09/2011)
    etxBairroSacado.Clear
    etxCodigoPracaSacado.Clear
    'Pt. 114032 - Moacir Pfau(23/02/2012)
    etxIdentificacaoCedente.valorTexto = ""
    etxHTMLReciboPersonalizado.Text = ""
    'Projeto: #4350 - História: # - Desenvolvimento# - Moacir Pfau(09/04/2013)
    cboDataRetorno.SelectItem "Data do Crédito"
    Call LimpaCamposCpEspecial
    Set mcolCarteiraCpEspecial = New clscolCarteiraCpEspecial
    etxLocalPagamento.Clear
    etxInstrucaoCobranca1.Clear
    etxInstrucaoCobranca2.Clear
    etxInstrucaoCobranca3.Clear
    etxValorInstrucaoCobranca1.Clear
    etxValorInstrucaoCobranca2.Clear
    etxValorInstrucaoCobranca3.Clear
    chkUtilizaNumeroControle.value = 0
    'Vinicius Elyseu(06/10/2015) - Projeto: #0 - História: #0 - Desenv: #0
    chkSeqRemessaNrDoc.value = 0
    CarregaGridCpEspecial
End Sub

'Preenchimento de valores padrão para a tela.
Private Sub fPreencheCampos()
    With mobjCarteira
         mlngEnterpriseId = .Enterprise_id
         mlngCdEstabelecimento = .Cd_estabelecimento
         etxIdCarteira.valorInteiro = .Id_carteira
         etxDesc.valorTexto = .Desc_carteira
         etxCendente.valorTexto = .Codigo_cedente
         etxInicioNN.valorTexto = .Inicio_nosso_numero
         etxFimNN.valorTexto = .Fim_nosso_numero
         etxProximoNN.valorTexto = .Proximo_nosso_numero
         etxDemonstativo.Text = .Demonstrativo
         etxInstrucao.Text = .Instrucoes_caixa
         cboLayoutBoleto.SelectItem .Tipo_layout_boleto
         cboLayoutRemessa.SelectItem .Tipo_layout_remessa
         cboLayoutRetorno.SelectItem .Tipo_layout_retorno
         etxArquivoLicenca.Valor = .Arquivo_licenca
         etxLogoEmpresa.Valor = .Logo_empresa
         etxCaminhoRemessa.Valor = .Caminho_arquivo_remessa_padrao
         etxCaminhoRetorno.Valor = .Caminho_arquivo_retorno_padrao
         etxMargem.valorInteiro = .Margem_superior_boleto
         etxOutro1.valorTexto = .Outro_dado_configuracao1
         etxOutro2.valorTexto = .Outro_dado_configuracao2
         etxSequencialRemessa.valorInteiro = .Sequencial_remessa
         'Pt. 96180 - Moacir Pfau(08/12/2009)
         cboEspecieDoc.SelectItem .Especie
        'Pt. 96589 - Moacir Pfau(08/02/2010)
        chkNNBanco.value = IIf(.Banco_gera_nosso_numero, 1, 0)
        'Pt. 97161 - Moacir Pfau(08/02/2010)
        etxDiasProtesto.valorInteiro = .Dias_protesto
        'pt.98446 - Fernando Paludo(26/04/2010)
        chkBancoEmiteBoleto.value = IIf(.banco_Emite_boleto, 1, 0)
        'pt.99257 - Moacir Pfau(30/06/2010)
        etxPerMulta.valorMoeda = .Per_multa
        etxPerMora.valorMoeda = .Per_mora
        'pt.98929 - Fernando Paludo(02/08/2010)
        etxOutraEspecie.valorTexto = .Outra_especie
        'Pt. 102431 - Moacir Pfau(03/11/2010)
        chkNaoGerarRegistroCabecalho1.value = IIf(.NaoGerarRegistroCabecalho1, 1, 0)
        chkNaoGerarRegistroCabecalho2.value = IIf(.NaoGerarRegistroCabecalho2, 1, 0)
        chkNaoGerarRegistroDetalhe1.value = IIf(.NaoGerarRegistroDetalhe1, 1, 0)
        chkNaoGerarRegistroDetalhe2.value = IIf(.NaoGerarRegistroDetalhe2, 1, 0)
        chkNaoGerarRegistroDetalhe3.value = IIf(.NaoGerarRegistroDetalhe3, 1, 0)
        chkNaoGerarRegistroDetalhe4.value = IIf(.NaoGerarRegistroDetalhe4, 1, 0)
        chkNaoGerarRegistroDetalhe5.value = IIf(.NaoGerarRegistroDetalhe5, 1, 0)
        chkNaoGerarRegistroRodape1.value = IIf(.NaoGerarRegistroRodape1, 1, 0)
        chkNaoGerarRegistroRodape2.value = IIf(.NaoGerarRegistroRodape2, 1, 0)
        'Pt. 105912 - Moacir Pfau(17/03/2011)
        etxTipoImpressao.valorTexto = .TipoImpressao
        etxBairroSacado.valorTexto = .BairroSacado
        etxCodigoPracaSacado.valorTexto = .CodigoPracaSacado
        'Pt. 114032 - Moacir Pfau(23/02/2012)
        etxIdentificacaoCedente.valorTexto = .IdentificacaoCedente
        etxHTMLReciboPersonalizado.Text = .HTMLReciboPersonalizado
        'Projeto: #4350 - História: # - Desenvolvimento# - Moacir Pfau(09/04/2013)
        cboDataRetorno.SelectItem IIf(.Data_baixa_retorno = eData_baixa_retorno.DataOcorrencia, "Data da Ocorrência", "Data do Crédito")
        etxLocalPagamento.valorTexto = .localPagamento
        etxInstrucaoCobranca1.valorTexto = .InstrucaoCobranca1
        etxInstrucaoCobranca2.valorTexto = .InstrucaoCobranca2
        etxInstrucaoCobranca3.valorTexto = .InstrucaoCobranca3
        etxValorInstrucaoCobranca1.valorTexto = .ValorInstrucaoCobranca1
        etxValorInstrucaoCobranca2.valorTexto = .ValorInstrucaoCobranca2
        etxValorInstrucaoCobranca3.valorTexto = .ValorInstrucaoCobranca3
        chkUtilizaNumeroControle.value = IIf(.UtilizaNumeroControle, 1, 0)
        chkSeqRemessaNrDoc.value = IIf(.SeqRemessaNrDoc, 1, 0)
        Set mcolCarteiraCpEspecial = .ColCpEspeciais
        CarregaGridCpEspecial
    End With
End Sub

Public Function LibProc(strFuncao As String) As Boolean
    Dim blnRetorno                      As Boolean
    
    Select Case strFuncao
        Case WL_NOVO
            Call fLimpaCampos
            mblnAlterando = False
            Call fBotaoNovo
        Case WL_SALVAR
            Call mobjCarteiraDAO.init(Aplicacao)
            If fValidaCampos Then
                fcarregaClasse
                fAplicaMascara
                Aplicacao.BeginTransaction
                If mblnAlterando Then
                    mobjCarteira.Proximo_nosso_numero = ""
                    blnRetorno = mobjCarteiraDAO.Atualizar(mobjCarteira)
                Else
                    blnRetorno = mobjCarteiraDAO.Gravar(mobjCarteira)
                End If
                'Se ocorreu erro.
                If blnRetorno = False Then
                    MsgBox "Erro ao realizar a gravação.", vbInformation, NomeModulo
                    Aplicacao.RollbackTransaction
                Else 'senão limpa os campos e atualiza flag.
                    mblnAlterando = True
                    MsgBox "Registro gravado com sucesso.", vbInformation, NomeModulo
                    Aplicacao.CommitTransaction
                    etxIdCarteira.valorInteiro = mobjCarteira.Id_carteira
                End If
                etxIdCarteira.SetFocus
                etxProximoNN.Enabled = False
            End If
            
        Case WL_DELETAR
            If (vbYes = MsgFunc("Tem certeza que deseja excluir este cadastro ?", _
                             vbQuestion Or vbYesNo Or vbDefaultButton2)) Then
                Aplicacao.BeginTransaction
                blnRetorno = mobjCarteiraDAO.Excluir(mobjCarteira)
                
                If blnRetorno = False Then
                    MsgBox "Erro ao realizar a exclusão.", vbInformation, NomeModulo
                    Aplicacao.RollbackTransaction
                Else
                    Call LibProc(WL_NOVO)
                    Aplicacao.CommitTransaction
                End If
            End If
            
        Case WL_CANCELAR
            If mblnAlterando Then
                Call fPreencheCampos
            Else
                Call LibProc(WL_NOVO)
            End If
            
        Case WL_PESQUISAR
            'Chama a tela de pesquisa
            Call fBotaoPesquisar
            Load frmCarteiraConsulta
            Call mostrarForm(frmCarteiraConsulta, 1234, True)
        Case WL_AJUDA
            Call fchamaAjuda
            
        Case WL_SAIR
            Unload Me
    End Select
End Function

Public Function fValidaCampos() As Boolean
    Dim strMensagem                 As String
    
    On Error GoTo err
    
    strMensagem = ""
    'Caminho
    If Trim(etxArquivoLicenca.Valor) <> "" Then
        If mobjCobreBem.ArquivoLicenca = "" Then
            strMensagem = strMensagem & "Não foi possível carregar o arquivo de configuração." & vbCrLf
        End If
    End If
    
    'Descrição                                                                  (01)
    If Trim(etxDesc.valorTexto) = "" Then
        strMensagem = strMensagem & "Descrição da carteira é obrigatória." & vbCrLf
    End If
   
    'Cendente                                                                   (02)
    If etxCendente.Enabled And Trim(etxCendente.valorTexto) = "" Then
        strMensagem = strMensagem & lblCendente.Caption & " é obrigatório." & vbCrLf
    Else
        If Len(Trim(fTiraCaracter(etxCendente.valorTexto))) <> Len(Trim(fTiraCaracter(mobjCobreBem.MascaraCodigoCedente))) Then
             strMensagem = strMensagem & "O " & mobjCobreBem.CabecalhoCodigoCedente & " é formado por " & Len(Trim(fTiraCaracter(mobjCobreBem.MascaraCodigoCedente))) & " caracteres, informação inválida." & vbCrLf
        End If
    End If
   
    'Inicio NN                                                                  (03)
    If etxInicioNN.Enabled Then
        If Trim(etxInicioNN.valorTexto) = "" Then
            strMensagem = strMensagem & "Início do nosso número é obrigatório." & vbCrLf
        Else
            If Len(Trim(fTiraCaracter(etxInicioNN.valorTexto))) <> Len(Trim(fTiraCaracter(mobjCobreBem.MascaraNossoNumero))) Then
                 strMensagem = strMensagem & "O código do início nosso número é formado por " & Len(Trim(fTiraCaracter(mobjCobreBem.MascaraNossoNumero))) & " caracteres, informação inválida." & vbCrLf
            End If
        End If
    End If
    
    'Fim NN                                                                     (04)
    If etxFimNN.Enabled Then
        If Trim(etxFimNN.valorTexto) = "" Then
            strMensagem = strMensagem & "Fim do nosso número é obrigatório." & vbCrLf
        Else
            If Len(Trim(fTiraCaracter(etxFimNN.valorTexto))) <> Len(Trim(fTiraCaracter(mobjCobreBem.MascaraNossoNumero))) Then
                 strMensagem = strMensagem & "O código do fim nosso número é formado por " & Len(Trim(fTiraCaracter(mobjCobreBem.MascaraNossoNumero))) & " caracteres, informação inválida." & vbCrLf
            End If
        End If
    End If
    
    'Proximo NN                                                                 (05)
    If etxFimNN.Enabled Then
        If Trim(etxProximoNN.valorTexto) = "" Then
            strMensagem = strMensagem & "Atual nosso número é obrigatório." & vbCrLf
        Else
            If Len(Trim(fTiraCaracter(etxProximoNN.valorTexto))) <> Len(Trim(fTiraCaracter(mobjCobreBem.MascaraNossoNumero))) Then
                 strMensagem = strMensagem & "O código do atual nosso número é formado por " & Len(Trim(fTiraCaracter(mobjCobreBem.MascaraNossoNumero))) & " caracteres, informação inválida." & vbCrLf
            End If
        End If
    End If
    
    'LayoutBoleto                                                               (06)
    If Trim(cboLayoutBoleto.SelectedItem) = "" Then
        strMensagem = strMensagem & "Layout Boleto é obrigatório." & vbCrLf
    End If
    
    'LayoutRemessa                                                              (07)
    If Trim(cboLayoutRemessa.SelectedItem) = "" Then
        strMensagem = strMensagem & "Layout Remessa é obrigatório." & vbCrLf
    End If
    
    'LayoutRetorno                                                              (08)
    If Trim(cboLayoutRetorno.SelectedItem) = "" Then
        strMensagem = strMensagem & "Layout Retorno é obrigatório." & vbCrLf
    End If
    
    'Margem                                                                     (09)
    If etxMargem.valorInteiro = 0 Then
        etxMargem.valorInteiro = 15
    End If
   
    'Outro1                                                                     (10)
    If Not (mobjCobreBem.NumeroBanco = "341-7") Then
        If Trim(etxOutro1.valorTexto) = "" And etxOutro1.Enabled Then
            strMensagem = strMensagem & lblOutro1.Caption & " é obrigatório." & vbCrLf
        Else
            If Len(Trim(fTiraCaracter(etxOutro1.valorTexto))) <> Len(Trim(fTiraCaracter(mobjCobreBem.MascaraOutroDadoConfiguracao1))) And etxOutro1.Enabled Then
                 strMensagem = strMensagem & "O " & lblOutro1.Caption & " é formado por " & Len(Trim(fTiraCaracter(mobjCobreBem.MascaraOutroDadoConfiguracao1))) & " caracteres, informação inválida." & vbCrLf
            End If
        End If
    End If
    
    'Outro2                                                                     (11)
    If Trim(etxOutro2.valorTexto) = "" And etxOutro2.Enabled Then
        strMensagem = strMensagem & lblOutro2.Caption & " é obrigatório." & vbCrLf
    Else
        If Len(Trim(fTiraCaracter(etxOutro2.valorTexto))) <> Len(Trim(fTiraCaracter(mobjCobreBem.MascaraOutroDadoConfiguracao2))) And etxOutro1.Enabled Then
             strMensagem = strMensagem & "O " & lblOutro2.Caption & " é formado por " & Len(Trim(fTiraCaracter(mobjCobreBem.MascaraOutroDadoConfiguracao2))) & " caracteres, informação inválida." & vbCrLf
        End If
    End If
    
    'Validando inicio e fim do nosso número.                                    (12)
    If Trim(etxInicioNN.valorTexto) <> "" And Trim(etxFimNN.valorTexto) <> "" Then
        If val(etxInicioNN.valorTexto) > val(etxFimNN.valorTexto) Then
            strMensagem = strMensagem & "Valor do fim nosso número não pode ser menor que o valor do início nosso número." & vbCrLf
        End If
    End If
    
    'Validando próximo nosso número.                                            (13)
    If Trim(etxFimNN.valorTexto) <> "" And Trim(etxProximoNN.valorTexto) <> "" Then
        If val(etxFimNN.valorTexto) < val(etxProximoNN.valorTexto) Then
            strMensagem = strMensagem & "Valor do próximo nosso número não pode ser maior que o valor do fim nosso número." & vbCrLf
        End If
    End If
    
    'Validando próximo nosso número com o início nosso numero.                  (14)
    If Trim(etxInicioNN.valorTexto) <> "" And Trim(etxProximoNN.valorTexto) <> "" Then
        If val(etxInicioNN.valorTexto) > val(etxProximoNN.valorTexto) Then
            If Not (val(etxInicioNN.valorTexto) = 1 And val(etxProximoNN.valorTexto) = 0) Then
                strMensagem = strMensagem & "Valor do próximo nosso número não pode ser maior que o valor do início nosso número." & vbCrLf
            End If
        End If
    End If
    
    'Pt. 96180 - Moacir Pfau(08/12/2009)
    'Validando especie.                                                         (15)
    If Trim(cboEspecieDoc.SelectedItem) = "" Then
        strMensagem = strMensagem & "Espécie Documento é obrigatório." & vbCrLf
    End If
    
    'pt.99257 - Moacir Pfau(30/06/2010)
    'Tratamento para o percentual de multa e mora. Não podendo ser maior que 99,99%.
    If etxPerMulta.valorMoeda > 99.99 Then
        etxPerMulta.valorMoeda = 0
        strMensagem = strMensagem & "Percentual de multa não pode ser igual ou maior que 100%." & vbCrLf
    End If
    
    If etxPerMora.valorMoeda > 99.99 Then
        etxPerMora.valorMoeda = 0
        strMensagem = strMensagem & "Percentual de mora não pode ser igual ou maior que 100%." & vbCrLf
    End If
    
    If strMensagem = "" Then
        fValidaCampos = True
    Else
        strMensagem = Replace(strMensagem, "&", "")
        MsgBox strMensagem, vbInformation, NomeModulo
    End If
    
    Exit Function
err:

End Function

'Função que a tela de pesquisa chama para carregar o registro.
Public Sub fCarregaPesquisa(lngIdCarteira As Integer)
    Dim objCarteira As clsCarteira
    Call fLimpaCampos
    
    Set objCarteira = mobjCarteiraDAO.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, lngIdCarteira)
    If Not objCarteira Is Nothing Then
        Set mobjCarteira = objCarteira
        Call fCarregaCobrebem(True)
        Call fPreencheCampos
        Call fBotaoPesquisar
        mblnAlterando = True
        etxProximoNN.Enabled = False
    End If
End Sub

Private Function fCarregaCobrebem(Optional blnPesquisa As Boolean)
    Dim i                       As Integer
    Dim blnEspecieDefault       As Boolean
    
    Set mobjCobreBem = New ContaCorrente
    If blnPesquisa Then
        mobjCobreBem.ArquivoLicenca = mobjCarteira.Arquivo_licenca
    Else
        mobjCobreBem.ArquivoLicenca = etxArquivoLicenca.Valor
    End If
      
    If Not mobjCobreBem Is Nothing Then
        'Layout Boleto
        If Not mobjCobreBem.LayoutsBoleto Is Nothing Then
            If mobjCobreBem.LayoutsBoleto.Count <> 0 Then
                cboLayoutBoleto.Clear
                For i = 0 To mobjCobreBem.LayoutsBoleto.Count - 1
                    cboLayoutBoleto.AddItem mobjCobreBem.LayoutsBoleto(i)
                Next
                cboLayoutBoleto.SelectItem mobjCobreBem.LayoutsBoleto(0)
            End If
        End If
        
        'Layout Remessa
        If Not mobjCobreBem.LayoutsArquivoRemessa Is Nothing Then
            If mobjCobreBem.LayoutsArquivoRemessa.Count <> 0 Then
                cboLayoutRemessa.Clear
                For i = 0 To mobjCobreBem.LayoutsArquivoRemessa.Count - 1
                    cboLayoutRemessa.AddItem mobjCobreBem.LayoutsArquivoRemessa(i)
                Next
                cboLayoutRemessa.SelectItem mobjCobreBem.LayoutsArquivoRemessa(0)
            End If
        End If
        'Layout Retorno
        If Not mobjCobreBem.LayoutsArquivoRetorno Is Nothing Then
            If mobjCobreBem.LayoutsArquivoRetorno.Count <> 0 Then
                cboLayoutRetorno.Clear
                For i = 0 To mobjCobreBem.LayoutsArquivoRetorno.Count - 1
                    cboLayoutRetorno.AddItem mobjCobreBem.LayoutsArquivoRetorno(i)
                Next
                cboLayoutRetorno.SelectItem mobjCobreBem.LayoutsArquivoRetorno(0)
            End If
        End If
        'Pt. 96180 - Moacir Pfau(08/12/2009)
        'Especie do Documento
        If Not mobjCobreBem.TiposDocumentosCobranca Is Nothing Then
            If mobjCobreBem.TiposDocumentosCobranca.Count <> 0 Then
                'combo.clear
                For i = 0 To mobjCobreBem.TiposDocumentosCobranca.Count - 1
                    cboEspecieDoc.AddItem mobjCobreBem.TiposDocumentosCobranca(i).Codigo
                    If mobjCobreBem.TiposDocumentosCobranca(i).Codigo = "RC" Then
                        blnEspecieDefault = True
                    End If
                Next
                If blnEspecieDefault Then
                    cboEspecieDoc.SelectItem "RC"
                Else
                    cboEspecieDoc.SelectItem mobjCobreBem.TiposDocumentosCobranca(0).Codigo
                End If
            End If
        End If
    End If
        
        
    'Label.
    If mobjCobreBem.MascaraCodigoCedente <> "" Then
        lblCendente.Caption = Mid(mobjCobreBem.CabecalhoCodigoCedente, 1, 21)
        etxCendente.Enabled = True
        lblCendente.Enabled = True
        etxCendente.MaxLength = Len(Trim(mobjCobreBem.MascaraCodigoCedente))
        etxCendente.ToolTipText = "Máscara: " & mobjCobreBem.MascaraCodigoCedente
    Else
        lblCendente.Caption = "Códig&o Cedente"
        etxCendente.Enabled = False
        lblCendente.Enabled = False
        etxCendente.Clear
    End If


    If mobjCobreBem.MascaraOutroDadoConfiguracao1 <> "" Then
        lblOutro1.Caption = Mid(mobjCobreBem.CabecalhoOutroDadoConfiguracao1, 1, 21)
        etxOutro1.Enabled = True
        lblOutro1.Enabled = True
        etxOutro1.MaxLength = Len(Trim(mobjCobreBem.MascaraOutroDadoConfiguracao1))
        etxOutro1.ToolTipText = "Máscara: " & mobjCobreBem.MascaraOutroDadoConfiguracao1
        If Not (mobjCobreBem.NumeroBanco = "341-7") Then
            lblOutro1.FontBold = True
        Else
            lblOutro1.FontBold = False
        End If
    Else
        lblOutro1.Caption = "Outro Dado Conf &1"
        etxOutro1.Enabled = False
        lblOutro1.Enabled = False
        etxOutro1.Clear
    End If

    If mobjCobreBem.MascaraOutroDadoConfiguracao2 <> "" Then
        lblOutro2.Caption = Mid(mobjCobreBem.CabecalhoOutroDadoConfiguracao2, 1, 21)
        etxOutro2.Enabled = True
        lblOutro2.Enabled = True
        etxOutro2.MaxLength = Len(Trim(mobjCobreBem.MascaraOutroDadoConfiguracao2))
        etxOutro2.ToolTipText = "Máscara: " & mobjCobreBem.MascaraOutroDadoConfiguracao2
    Else
        lblOutro2.Caption = "Outro Dado Conf &2"
        etxOutro2.Enabled = False
        lblOutro2.Enabled = False
        etxOutro2.Clear
    End If
    
    etxInicioNN.MaxLength = Len(Trim(mobjCobreBem.MascaraNossoNumero))
    etxFimNN.MaxLength = Len(Trim(mobjCobreBem.MascaraNossoNumero))
    etxProximoNN.MaxLength = Len(Trim(mobjCobreBem.MascaraNossoNumero))
    
End Function

Private Function fAplicaMascara()
    'Label.
    If mobjCobreBem.MascaraCodigoCedente <> "" Then
        mobjCarteira.Codigo_cedente = fColocaMascara(fTiraCaracter(mobjCarteira.Codigo_cedente), mobjCobreBem.MascaraCodigoCedente)
    End If

    If mobjCobreBem.MascaraOutroDadoConfiguracao1 <> "" Then
        mobjCarteira.Outro_dado_configuracao1 = fColocaMascara(fTiraCaracter(mobjCarteira.Outro_dado_configuracao1), mobjCobreBem.MascaraOutroDadoConfiguracao1)
    End If
    
    If mobjCobreBem.MascaraOutroDadoConfiguracao2 <> "" Then
        mobjCarteira.Outro_dado_configuracao2 = fColocaMascara(fTiraCaracter(mobjCarteira.Outro_dado_configuracao2), mobjCobreBem.MascaraOutroDadoConfiguracao2)
    End If
    
    mobjCarteira.Inicio_nosso_numero = fColocaZero(mobjCarteira.Inicio_nosso_numero, mobjCobreBem.MascaraNossoNumero)
    mobjCarteira.Fim_nosso_numero = fColocaZero(mobjCarteira.Fim_nosso_numero, mobjCobreBem.MascaraNossoNumero)
    mobjCarteira.Proximo_nosso_numero = fColocaZero(mobjCarteira.Proximo_nosso_numero, mobjCobreBem.MascaraNossoNumero)
    
    etxCendente.valorTexto = mobjCarteira.Codigo_cedente
    etxInicioNN.valorTexto = mobjCarteira.Inicio_nosso_numero
    etxFimNN.valorTexto = mobjCarteira.Fim_nosso_numero
    etxProximoNN.valorTexto = mobjCarteira.Proximo_nosso_numero
    etxOutro1.valorTexto = mobjCarteira.Outro_dado_configuracao1
    etxOutro2.valorTexto = mobjCarteira.Outro_dado_configuracao2
    
End Function

Private Sub fchamaAjuda()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub fBotaoNovo()
    cmdNovo.Enabled = True
    cmdGravar.Enabled = True
    cmdExcluir.Enabled = False
    cmdCancelar.Enabled = True
    'Conf Especial Remessa
    cmdCEExcluir.Enabled = False
    etxProximoNN.Enabled = True
End Sub

Private Sub fBotaoGravar()
    cmdNovo.Enabled = True
    cmdGravar.Enabled = False
    cmdExcluir.Enabled = True
    cmdCancelar.Enabled = True
    cmdPesquisar.Enabled = True
    etxProximoNN.Enabled = False
End Sub

Public Sub fBotaoPesquisar()
    cmdNovo.Enabled = True
    cmdGravar.Enabled = True
    cmdExcluir.Enabled = True
    cmdCancelar.Enabled = True
    cmdPesquisar.Enabled = True
    etxProximoNN.Enabled = False
End Sub

Private Function fTiraCaracter(ByVal strCampo As String)
    strCampo = Replace(strCampo, ".", "")
    strCampo = Replace(strCampo, ",", "")
    strCampo = Replace(strCampo, "-", "")
    strCampo = Replace(strCampo, "/", "")
    fTiraCaracter = strCampo
End Function

Private Function fColocaZero(ByVal strValor As String, ByVal strMascara As String) As String
    Dim lngQtdeZero             As Long
    Dim lngQtdeMascara          As Long
    Dim lngQtdeValor            As Long

    lngQtdeMascara = Len(fTiraCaracter(strMascara))
    lngQtdeValor = Len(CStr(val(strValor)))
    lngQtdeZero = lngQtdeMascara - lngQtdeValor

    fColocaZero = String(lngQtdeZero, "0") & val(strValor)

End Function

Private Function fColocaMascara(ByVal strValor As String, ByVal strMascara As String) As String
    Dim Index As Integer
    Dim temp As String
    Dim i As Integer
    
    For i = 1 To Len(strMascara)
        If (Mid(strMascara, i, 1) = "#") Or (Mid(strMascara, i, 1) = ".") Or (Mid(strMascara, i, 1) = ",") Or (Mid(strMascara, i, 1) = "-") Or (Mid(strMascara, i, 1) = "/") Then
            temp = temp & Mid(strMascara, i, 1)
            Index = Index + 1
        Else
            If Index = 0 Then
                temp = temp & Mid(strValor, i, 1)
            Else
                temp = temp & Mid(strValor, i - Index, 1)
            End If
        End If
    Next
    fColocaMascara = temp
End Function


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
