VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConhecimentoFretePagar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conhecimento de Frete (A Pagar)"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   13050
   Begin VB.Frame fraBotoes 
      Height          =   7590
      Left            =   11520
      TabIndex        =   61
      Top             =   0
      Width           =   1500
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   135
         TabIndex        =   59
         Top             =   1845
         Width           =   1215
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   135
         TabIndex        =   58
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   135
         TabIndex        =   57
         Top             =   1035
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   135
         TabIndex        =   56
         Top             =   630
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   135
         TabIndex        =   55
         Top             =   225
         Width           =   1215
      End
      Begin MSComctlLib.ImageList imgImagens 
         Left            =   315
         Top             =   2610
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
               Picture         =   "ConhecimentoFretePagar.frx":0000
               Key             =   "selecionado"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ConhecimentoFretePagar.frx":015A
               Key             =   "item"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraControles 
      Height          =   7590
      Left            =   0
      TabIndex        =   60
      Top             =   -45
      Width           =   11520
      Begin TabDlg.SSTab tabConhecimentoPagar 
         Height          =   7365
         Left            =   90
         TabIndex        =   62
         Top             =   180
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   12991
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Notas Fiscais"
         TabPicture(0)   =   "ConhecimentoFretePagar.frx":02B4
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label7"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lstNotaFiscalNotas"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fraFiltros"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtNotaFiscalQuantidade"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtNotaFiscalValorTotal"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Conhecimento"
         TabPicture(1)   =   "ConhecimentoFretePagar.frx":02D0
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "fraConhecimento"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Titulos a Pagar"
         TabPicture(2)   =   "ConhecimentoFretePagar.frx":02EC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraTitulos"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame fraTitulos 
            Height          =   6975
            Left            =   -74970
            TabIndex        =   127
            Top             =   330
            Width           =   11265
            Begin VB.ComboBox cboTipoGlobalTitulo 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1230
               TabIndex        =   37
               Top             =   195
               Width           =   1455
            End
            Begin VB.TextBox txtTituloConhecimento 
               Enabled         =   0   'False
               Height          =   315
               Left            =   4860
               MaxLength       =   9
               TabIndex        =   38
               Top             =   180
               Width           =   1185
            End
            Begin VB.TextBox txtTituloEmpresa 
               Height          =   315
               Left            =   1230
               MaxLength       =   15
               TabIndex        =   39
               Top             =   540
               Width           =   1635
            End
            Begin VB.TextBox txtTituloEmissao 
               Height          =   315
               Left            =   1230
               MaxLength       =   10
               TabIndex        =   40
               Top             =   900
               Width           =   1635
            End
            Begin VB.TextBox txtTituloDescricao 
               Height          =   315
               Left            =   1230
               MaxLength       =   80
               TabIndex        =   41
               Top             =   1260
               Width           =   4965
            End
            Begin VB.TextBox txtTituloControle 
               Height          =   315
               Left            =   8085
               MaxLength       =   18
               TabIndex        =   42
               Top             =   1260
               Width           =   1230
            End
            Begin VB.TextBox txtTituloBanco 
               Height          =   315
               Left            =   1230
               MaxLength       =   9
               TabIndex        =   43
               Top             =   1620
               Width           =   1455
            End
            Begin VB.TextBox txtTituloCarteira 
               Height          =   315
               Left            =   8085
               MaxLength       =   3
               TabIndex        =   44
               Top             =   1620
               Width           =   1230
            End
            Begin VB.TextBox txtTituloConta 
               Height          =   315
               Left            =   1230
               TabIndex        =   45
               Top             =   1980
               Width           =   1455
            End
            Begin VB.TextBox txtTituloCondicaoPagamento 
               Height          =   315
               Left            =   1230
               MaxLength       =   4
               TabIndex        =   47
               Top             =   2340
               Width           =   780
            End
            Begin VB.TextBox txtTituloValor 
               Enabled         =   0   'False
               Height          =   315
               Left            =   1230
               MaxLength       =   25
               TabIndex        =   48
               Top             =   2700
               Width           =   1185
            End
            Begin VB.CommandButton cmdTituloCalcular 
               Caption         =   "&Calcular"
               Height          =   375
               Left            =   9720
               TabIndex        =   49
               Top             =   2640
               Width           =   1215
            End
            Begin VB.Frame fraParcelas 
               Caption         =   "Parcelas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   3615
               Left            =   60
               TabIndex        =   128
               Top             =   3300
               Width           =   11025
               Begin VB.TextBox txtTituloValorTotalParcelas 
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000001&
                  Height          =   315
                  Left            =   8685
                  TabIndex        =   129
                  Top             =   3180
                  Width           =   2220
               End
               Begin VB.TextBox txtTituloParcelaNumero 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1260
                  TabIndex        =   50
                  Top             =   450
                  Width           =   1185
               End
               Begin VB.TextBox txtTituloParcelaVencimento 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1260
                  MaxLength       =   10
                  TabIndex        =   51
                  Top             =   810
                  Width           =   1185
               End
               Begin VB.TextBox txtTituloParcelaValor 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1260
                  MaxLength       =   9
                  TabIndex        =   52
                  Top             =   1170
                  Width           =   1185
               End
               Begin VB.CommandButton cmdTituloParcelaConfirmar 
                  Caption         =   "Confirmar"
                  Height          =   375
                  Left            =   450
                  TabIndex        =   53
                  Top             =   1755
                  Width           =   1215
               End
               Begin VB.CommandButton cmdTituloParcelaCancelar 
                  Caption         =   "Cancelar"
                  Height          =   375
                  Left            =   2115
                  TabIndex        =   54
                  Top             =   1755
                  Width           =   1215
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTituloParcelas 
                  Height          =   2715
                  Left            =   5520
                  TabIndex        =   130
                  Top             =   360
                  Width           =   5325
                  _ExtentX        =   9393
                  _ExtentY        =   4789
                  _Version        =   393216
                  _NumberOfBands  =   1
                  _Band(0).Cols   =   2
               End
               Begin VB.Label labValorTotal 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Valor Total (=):"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   7335
                  TabIndex        =   134
                  Top             =   3225
                  Width           =   1290
               End
               Begin VB.Label Label48 
                  AutoSize        =   -1  'True
                  Caption         =   "Parcela"
                  Enabled         =   0   'False
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   675
                  TabIndex        =   133
                  Top             =   495
                  Width           =   540
               End
               Begin VB.Label Label49 
                  AutoSize        =   -1  'True
                  Caption         =   "Vencimento"
                  Enabled         =   0   'False
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   360
                  TabIndex        =   132
                  Top             =   855
                  Width           =   840
               End
               Begin VB.Label Label50 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor"
                  Enabled         =   0   'False
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   810
                  TabIndex        =   131
                  Top             =   1215
                  Width           =   360
               End
            End
            Begin VB.TextBox txtTituloCentroCusto 
               Height          =   315
               Left            =   8085
               TabIndex        =   46
               Top             =   1980
               Width           =   1230
            End
            Begin VB.Label labTipo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tipo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   780
               TabIndex        =   151
               Top             =   270
               Width           =   390
            End
            Begin VB.Label labConhecimento 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Conhecimento"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3570
               TabIndex        =   150
               Top             =   225
               Width           =   1215
            End
            Begin VB.Label labEmpresa 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Empresa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   435
               TabIndex        =   149
               Top             =   600
               Width           =   735
            End
            Begin VB.Label lblTituloEmpresa 
               AutoSize        =   -1  'True
               Caption         =   "lblTituloEmpresa"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2850
               TabIndex        =   148
               Top             =   630
               Width           =   1155
            End
            Begin VB.Label labEmissao 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Emissão"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   465
               TabIndex        =   147
               Top             =   945
               Width           =   705
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Descrição"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   450
               TabIndex        =   146
               Top             =   1305
               Width           =   720
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "Controle"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7410
               TabIndex        =   145
               Top             =   1350
               Width           =   585
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Banco"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   705
               TabIndex        =   144
               Top             =   1665
               Width           =   465
            End
            Begin VB.Label lblTituloBanco 
               AutoSize        =   -1  'True
               Caption         =   "lblTituloBanco"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2805
               TabIndex        =   143
               Top             =   1710
               Width           =   1005
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Carteira"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7470
               TabIndex        =   142
               Top             =   1665
               Width           =   540
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Conta"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   750
               TabIndex        =   141
               Top             =   2025
               Width           =   420
            End
            Begin VB.Label lblTituloConta 
               AutoSize        =   -1  'True
               Caption         =   "lblTituloConta"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2805
               TabIndex        =   140
               Top             =   2070
               Width           =   960
            End
            Begin VB.Label labCondPagto 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cond. Pagto"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   139
               Top             =   2385
               Width           =   1065
            End
            Begin VB.Label lblTituloCondicaoPagamento 
               AutoSize        =   -1  'True
               Caption         =   "lblTituloCondicaoPagamento"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2115
               TabIndex        =   138
               Top             =   2430
               Width           =   2025
            End
            Begin VB.Label labValor 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Valor"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   720
               TabIndex        =   137
               Top             =   2790
               Width           =   450
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               Caption         =   "C. Custo"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   7440
               TabIndex        =   136
               Top             =   2040
               Width           =   600
            End
            Begin VB.Label lblTituloCentroCusto 
               AutoSize        =   -1  'True
               Caption         =   "lblTituloCentroCusto"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   9435
               TabIndex        =   135
               Top             =   2025
               Width           =   1410
            End
         End
         Begin VB.Frame fraConhecimento 
            Height          =   6975
            Left            =   30
            TabIndex        =   77
            Top             =   330
            Width           =   11265
            Begin VB.TextBox txtSerieCTRC 
               Height          =   315
               Left            =   3480
               MaxLength       =   3
               TabIndex        =   10
               Top             =   555
               Width           =   585
            End
            Begin VB.TextBox txtChaveAcessoEnt 
               Height          =   315
               Left            =   1425
               MaxLength       =   44
               TabIndex        =   22
               Top             =   3450
               Width           =   4845
            End
            Begin VB.TextBox txtConhecimentoTransportadora 
               Height          =   315
               Left            =   1425
               MaxLength       =   4
               TabIndex        =   8
               Top             =   195
               Width           =   825
            End
            Begin VB.TextBox txtConhecimentoNumero 
               Height          =   315
               Left            =   1425
               MaxLength       =   9
               TabIndex        =   9
               Top             =   555
               Width           =   975
            End
            Begin VB.TextBox txtConhecimentoDataEmissao 
               Height          =   315
               Left            =   7815
               MaxLength       =   10
               TabIndex        =   12
               Top             =   555
               Width           =   1095
            End
            Begin VB.TextBox txtConhecimentoSituacao 
               Enabled         =   0   'False
               Height          =   315
               Left            =   10065
               TabIndex        =   104
               Top             =   195
               Width           =   1095
            End
            Begin VB.TextBox txtConhecimentoRemetente 
               Height          =   315
               Left            =   1425
               MaxLength       =   15
               TabIndex        =   13
               Top             =   930
               Width           =   1185
            End
            Begin VB.TextBox txtConhecimentoDestinatario 
               Height          =   315
               Left            =   1425
               MaxLength       =   15
               TabIndex        =   14
               Top             =   1290
               Width           =   1185
            End
            Begin VB.TextBox txtConhecimentoConsignatario 
               Height          =   315
               Left            =   1425
               MaxLength       =   15
               TabIndex        =   15
               Top             =   1650
               Width           =   1185
            End
            Begin VB.TextBox txtConhecimentoRedespacho 
               Height          =   315
               Left            =   1425
               MaxLength       =   4
               TabIndex        =   16
               Top             =   2010
               Width           =   825
            End
            Begin VB.TextBox txtConhecimentoNaturezaOperacao 
               Height          =   315
               Left            =   1425
               MaxLength       =   4
               TabIndex        =   17
               Top             =   2370
               Width           =   825
            End
            Begin VB.TextBox txtConhecimentoNaturezaComplemento 
               Height          =   315
               Left            =   2280
               MaxLength       =   1
               TabIndex        =   18
               Top             =   2370
               Width           =   330
            End
            Begin VB.TextBox txtConhecimentoDistancia 
               Height          =   315
               Left            =   1425
               MaxLength       =   6
               TabIndex        =   19
               Top             =   2730
               Width           =   1185
            End
            Begin VB.TextBox txtConhecimentoObservacao 
               Height          =   315
               Left            =   1425
               MaxLength       =   250
               TabIndex        =   21
               Top             =   3120
               Width           =   8385
            End
            Begin VB.Frame fraConhecimentoValores 
               Caption         =   "Valores"
               ForeColor       =   &H00000000&
               Height          =   3150
               Left            =   60
               TabIndex        =   84
               Top             =   3765
               Width           =   9945
               Begin VB.CheckBox chkConhecimentoAdionarFreteMercadoria 
                  Caption         =   "Acrescentar Frete no valor da mercadoria"
                  ForeColor       =   &H00000000&
                  Height          =   330
                  Left            =   360
                  TabIndex        =   36
                  Top             =   2745
                  Width           =   4110
               End
               Begin VB.TextBox txtConhecimentoValorTotal 
                  BackColor       =   &H80000018&
                  Enabled         =   0   'False
                  ForeColor       =   &H00000000&
                  Height          =   315
                  Left            =   7110
                  TabIndex        =   102
                  Top             =   2760
                  Width           =   2505
               End
               Begin VB.Frame Frame5 
                  Caption         =   "Outros"
                  ForeColor       =   &H00000000&
                  Height          =   1005
                  Left            =   6750
                  TabIndex        =   100
                  Top             =   1710
                  Width           =   2850
                  Begin VB.TextBox txtConhecimentoValorIsentas 
                     Height          =   315
                     Left            =   1440
                     MaxLength       =   9
                     TabIndex        =   35
                     Top             =   360
                     Width           =   1185
                  End
                  Begin VB.Label Label33 
                     AutoSize        =   -1  'True
                     Caption         =   "Valor Isentas"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   405
                     TabIndex        =   101
                     Top             =   450
                     Width           =   915
                  End
               End
               Begin VB.Frame Frame4 
                  Caption         =   "ICMS"
                  ForeColor       =   &H00000000&
                  Height          =   1455
                  Left            =   6750
                  TabIndex        =   96
                  Top             =   225
                  Width           =   2850
                  Begin VB.TextBox txtConhecimentoValorIcms 
                     BackColor       =   &H80000018&
                     Enabled         =   0   'False
                     ForeColor       =   &H80000001&
                     Height          =   315
                     Left            =   1395
                     MaxLength       =   9
                     TabIndex        =   34
                     Top             =   990
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoPorcentagemIcms 
                     Height          =   315
                     Left            =   1395
                     MaxLength       =   9
                     TabIndex        =   33
                     Top             =   630
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoBaseIcms 
                     Height          =   315
                     Left            =   1395
                     MaxLength       =   9
                     TabIndex        =   32
                     Top             =   270
                     Width           =   1185
                  End
                  Begin VB.Label Label32 
                     AutoSize        =   -1  'True
                     Caption         =   "Valor ICMS"
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   360
                     TabIndex        =   99
                     Top             =   1035
                     Width           =   960
                  End
                  Begin VB.Label Label31 
                     AutoSize        =   -1  'True
                     Caption         =   "% ICMS"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   765
                     TabIndex        =   98
                     Top             =   675
                     Width           =   555
                  End
                  Begin VB.Label Label30 
                     AutoSize        =   -1  'True
                     Caption         =   "Base ICMS"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   540
                     TabIndex        =   97
                     Top             =   315
                     Width           =   795
                  End
               End
               Begin VB.Frame Frame3 
                  Caption         =   "(+) Conhecimento"
                  ForeColor       =   &H00000000&
                  Height          =   2490
                  Left            =   3375
                  TabIndex        =   89
                  Top             =   225
                  Width           =   3345
                  Begin VB.TextBox txtConhecimentoEncargosValor 
                     BackColor       =   &H80000018&
                     Enabled         =   0   'False
                     ForeColor       =   &H80000001&
                     Height          =   315
                     Left            =   1620
                     TabIndex        =   31
                     Top             =   2115
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoDesconto 
                     Height          =   315
                     Left            =   1620
                     MaxLength       =   9
                     TabIndex        =   30
                     Top             =   1755
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoAcrescimo 
                     Height          =   315
                     Left            =   1620
                     MaxLength       =   9
                     TabIndex        =   29
                     Top             =   1395
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoOutros 
                     Height          =   315
                     Left            =   1620
                     MaxLength       =   9
                     TabIndex        =   28
                     Top             =   1035
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoSeguro 
                     Height          =   315
                     Left            =   1620
                     MaxLength       =   9
                     TabIndex        =   27
                     Top             =   675
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoPedagio 
                     Height          =   315
                     Left            =   1620
                     MaxLength       =   9
                     TabIndex        =   26
                     Top             =   315
                     Width           =   1185
                  End
                  Begin VB.Label Label29 
                     AutoSize        =   -1  'True
                     Caption         =   "Valor (=)"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   780
                     TabIndex        =   95
                     Top             =   2160
                     Width           =   735
                  End
                  Begin VB.Label Label28 
                     AutoSize        =   -1  'True
                     Caption         =   "Desconto (-)"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   645
                     TabIndex        =   94
                     Top             =   1800
                     Width           =   870
                  End
                  Begin VB.Label Label27 
                     AutoSize        =   -1  'True
                     Caption         =   "Acréscimo (+)"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   570
                     TabIndex        =   93
                     Top             =   1440
                     Width           =   960
                  End
                  Begin VB.Label Label26 
                     AutoSize        =   -1  'True
                     Caption         =   "Outros (+)"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   825
                     TabIndex        =   92
                     Top             =   1080
                     Width           =   690
                  End
                  Begin VB.Label Label25 
                     AutoSize        =   -1  'True
                     Caption         =   "Seguro(+)"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   825
                     TabIndex        =   91
                     Top             =   720
                     Width           =   690
                  End
                  Begin VB.Label Label24 
                     AutoSize        =   -1  'True
                     Caption         =   "Pedágio (+)"
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   690
                     TabIndex        =   90
                     Top             =   360
                     Width           =   810
                  End
               End
               Begin VB.Frame Frame2 
                  Caption         =   "Frete"
                  ForeColor       =   &H00000000&
                  Height          =   2490
                  Left            =   360
                  TabIndex        =   85
                  Top             =   225
                  Width           =   2985
                  Begin VB.TextBox txtConhecimentoValor 
                     BackColor       =   &H80000014&
                     Height          =   315
                     Left            =   1395
                     MaxLength       =   9
                     TabIndex        =   25
                     Top             =   1080
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoTarifa 
                     Height          =   315
                     Left            =   1395
                     MaxLength       =   9
                     TabIndex        =   24
                     Top             =   720
                     Width           =   1185
                  End
                  Begin VB.TextBox txtConhecimentoVolume 
                     Height          =   315
                     Left            =   1395
                     MaxLength       =   9
                     TabIndex        =   23
                     Top             =   360
                     Width           =   1185
                  End
                  Begin VB.Label Label23 
                     AutoSize        =   -1  'True
                     Caption         =   "Valor(=)"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   645
                     TabIndex        =   88
                     Top             =   1125
                     Width           =   675
                  End
                  Begin VB.Label Label22 
                     AutoSize        =   -1  'True
                     Caption         =   "Tarifa (*)"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   555
                     TabIndex        =   87
                     Top             =   765
                     Width           =   765
                  End
                  Begin VB.Label Label21 
                     AutoSize        =   -1  'True
                     Caption         =   "Volume (N.F.)"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   195
                     Left            =   165
                     TabIndex        =   86
                     Top             =   405
                     Width           =   1170
                  End
               End
               Begin VB.Label Label34 
                  AutoSize        =   -1  'True
                  Caption         =   "Valor total do conhecimento"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   4635
                  TabIndex        =   103
                  Top             =   2790
                  Width           =   2400
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Tipo de frete"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1470
               Left            =   8520
               TabIndex        =   79
               Top             =   1200
               Width           =   2655
               Begin VB.OptionButton optConhecimentoTerceiros 
                  Caption         =   "Por conta de Terceiros"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   75
                  TabIndex        =   83
                  Top             =   1080
                  Width           =   2550
               End
               Begin VB.OptionButton optConhecimentoEmitente 
                  Caption         =   "Por conta do Emitente"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   75
                  TabIndex        =   82
                  Top             =   840
                  Width           =   2500
               End
               Begin VB.OptionButton optConhecimentoFob 
                  Caption         =   "Por conta do Destinatário (FOB)"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   75
                  TabIndex        =   81
                  Top             =   600
                  Width           =   2550
               End
               Begin VB.OptionButton optConhecimentoCIF 
                  Caption         =   "Por conta do Remetente (CIF)"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   75
                  TabIndex        =   80
                  Top             =   360
                  Width           =   2505
               End
            End
            Begin VB.TextBox txtConhecimentoDataEntrada 
               Height          =   315
               Left            =   10065
               TabIndex        =   78
               Top             =   555
               Width           =   1095
            End
            Begin VB.ComboBox cboTipoGlobal 
               Height          =   315
               Left            =   4980
               TabIndex        =   11
               Text            =   "cboTipoGlobal"
               Top             =   555
               Width           =   1455
            End
            Begin Fox.EBSText txtConhecimentoOperacaoContabil 
               Height          =   330
               Left            =   5430
               TabIndex        =   20
               Top             =   2730
               Width           =   1035
               _ExtentX        =   265
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
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Série"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3000
               TabIndex        =   153
               Top             =   615
               Width           =   360
            End
            Begin VB.Label lblChaveAcessoEnt 
               AutoSize        =   -1  'True
               Caption         =   "Chave Acesso"
               Height          =   195
               Left            =   300
               TabIndex        =   152
               Top             =   3480
               Width           =   1035
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Transportadora"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   60
               TabIndex        =   126
               Top             =   255
               Width           =   1305
            End
            Begin VB.Label lblConhecimentoTransportadora 
               AutoSize        =   -1  'True
               Caption         =   "lblConhecimentoTransportadora"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2325
               TabIndex        =   125
               Top             =   255
               Width           =   2250
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Conhecimento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   124
               Top             =   615
               Width           =   1215
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Tipo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4530
               TabIndex        =   123
               Top             =   615
               Width           =   390
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Data Emissão"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   6555
               TabIndex        =   122
               Top             =   615
               Width           =   1170
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Situação"
               Enabled         =   0   'False
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   9360
               TabIndex        =   121
               Top             =   255
               Width           =   630
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Remetente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   405
               TabIndex        =   120
               Top             =   975
               Width           =   930
            End
            Begin VB.Label lblConhecimentoRemetente 
               AutoSize        =   -1  'True
               Caption         =   "lblConhecimentoRemetente"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2730
               TabIndex        =   119
               Top             =   1020
               Width           =   1950
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Destinatário"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   300
               TabIndex        =   118
               Top             =   1335
               Width           =   1035
            End
            Begin VB.Label lblConhecimentoDestinatario 
               AutoSize        =   -1  'True
               Caption         =   "lblConhecimentoDestinatario"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2730
               TabIndex        =   117
               Top             =   1335
               Width           =   2010
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Consignatário"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   255
               TabIndex        =   116
               Top             =   1695
               Width           =   1080
            End
            Begin VB.Label lblConhecimentoConsignatario 
               AutoSize        =   -1  'True
               Caption         =   "lblConhecimentoConsignatario"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2730
               TabIndex        =   115
               Top             =   1695
               Width           =   2130
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Redespacho"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   255
               TabIndex        =   114
               Top             =   2055
               Width           =   1080
            End
            Begin VB.Label lblConhecimentoRedespacho 
               AutoSize        =   -1  'True
               Caption         =   "lblConhecimentoRedespacho"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2355
               TabIndex        =   113
               Top             =   2055
               Width           =   2085
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Nat.Operação"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   135
               TabIndex        =   112
               Top             =   2415
               Width           =   1200
            End
            Begin VB.Label lblConhecimentoNaturezaOperacao 
               AutoSize        =   -1  'True
               Caption         =   "lblConhecimentoNaturezaOperacao"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2805
               TabIndex        =   111
               Top             =   2460
               Width           =   2520
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Distância"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   255
               TabIndex        =   110
               Top             =   2790
               Width           =   1080
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Observação"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   255
               TabIndex        =   109
               Top             =   3135
               Width           =   1080
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               Caption         =   "km"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2745
               TabIndex        =   108
               Top             =   2760
               Width           =   210
            End
            Begin VB.Label Label52 
               AutoSize        =   -1  'True
               Caption         =   "Data Entrada"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   9030
               TabIndex        =   107
               Top             =   615
               Width           =   945
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               Caption         =   "Operação Contábil"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3750
               TabIndex        =   106
               Top             =   2790
               Width           =   1590
            End
            Begin VB.Label lblConhecimentoOperacaoContabil 
               AutoSize        =   -1  'True
               Caption         =   "lblConhecimentoOperacaoContabil"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   6555
               TabIndex        =   105
               Top             =   2790
               Width           =   2445
            End
         End
         Begin VB.TextBox txtNotaFiscalValorTotal 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   -65310
            TabIndex        =   74
            Top             =   6930
            Width           =   1410
         End
         Begin VB.TextBox txtNotaFiscalQuantidade 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   -67605
            TabIndex        =   72
            Top             =   6930
            Width           =   1185
         End
         Begin VB.Frame fraFiltros 
            Height          =   1440
            Left            =   -74970
            TabIndex        =   63
            Top             =   330
            Width           =   11115
            Begin VB.ComboBox cboNotaFiscalEntradaSaida 
               Height          =   315
               ItemData        =   "ConhecimentoFretePagar.frx":0308
               Left            =   1410
               List            =   "ConhecimentoFretePagar.frx":0312
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   195
               Width           =   1380
            End
            Begin VB.CommandButton cmdNotaFiscalExecutar 
               Caption         =   "&Executar"
               Height          =   375
               Left            =   9720
               TabIndex        =   7
               Top             =   600
               Width           =   1215
            End
            Begin VB.ComboBox cboNotaFiscalTipo 
               Height          =   315
               Left            =   3510
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   195
               Width           =   1230
            End
            Begin VB.TextBox txtNotaFiscalTransportadora 
               Height          =   315
               Left            =   1395
               MaxLength       =   4
               TabIndex        =   6
               Top             =   1005
               Width           =   870
            End
            Begin VB.TextBox txtNotaFiscalNotaFinal 
               Height          =   315
               Left            =   6705
               MaxLength       =   9
               TabIndex        =   3
               Top             =   195
               Width           =   1095
            End
            Begin VB.TextBox txtNotaFiscalNotaInicial 
               Height          =   315
               Left            =   5355
               MaxLength       =   9
               TabIndex        =   2
               Top             =   195
               Width           =   1050
            End
            Begin VB.TextBox txtNotaFiscalDataEmissaoFinal 
               Height          =   315
               Left            =   2745
               MaxLength       =   10
               TabIndex        =   5
               Top             =   600
               Width           =   1095
            End
            Begin VB.TextBox txtNotaFiscalDataEmissaoInicial 
               Height          =   315
               Left            =   1395
               MaxLength       =   10
               TabIndex        =   4
               Top             =   600
               Width           =   1050
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Entrada/Saída"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   76
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Tipo"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   3090
               TabIndex        =   70
               Top             =   240
               Width           =   315
            End
            Begin VB.Label lblNotaFiscalTransportadora 
               AutoSize        =   -1  'True
               Caption         =   "lblNotaFiscalTransportadora"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2325
               TabIndex        =   69
               Top             =   1050
               Width           =   1980
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Transportadora"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   68
               Top             =   1050
               Width           =   1080
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "a"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   6525
               TabIndex        =   67
               Top             =   240
               Width           =   90
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Nota"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   4905
               TabIndex        =   66
               Top             =   240
               Width           =   345
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "a"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   2535
               TabIndex        =   65
               Top             =   645
               Width           =   90
            End
            Begin VB.Label lblDataEmissao 
               AutoSize        =   -1  'True
               Caption         =   "Data Emissão"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   180
               TabIndex        =   64
               Top             =   645
               Width           =   975
            End
         End
         Begin MSComctlLib.ListView lstNotaFiscalNotas 
            Height          =   4965
            Left            =   -74985
            TabIndex        =   75
            Top             =   1800
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   8758
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "imgImagens"
            SmallIcons      =   "imgImagens"
            ForeColor       =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sel"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nota"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Tipo"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Situação"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Data"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Cliente"
               Object.Width           =   4939
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "valor n.f."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "apelido"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "transportadora"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "cod_transp"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "valor Frete"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   -66345
            TabIndex        =   73
            Top             =   6975
            Width           =   945
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Qtde.NF."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   -68445
            TabIndex        =   71
            Top             =   6975
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "frmConhecimentoFretePagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Objeto de navegação
Private navigator As New cFretePagarNavigator
'Variavel que guarda o indice do item selecionado na lista
Private lngItem As Long
'Objeto que guarda informações do conhecimento
Private objConhecimento As New cFretePagar
'Objeto estado de Origem
Private objUfOrigem As cEstado
'Objeto estado de Destino
Private objUfDestino As cEstado
'Variavel que define se o foco está no valor
Private booValorFoco As Boolean
'Variavel que define se o objeto está em alteração
Private booAlterando As Boolean
'Variavel que indica a condicao de pagamento gerada
Private intCondPag As Integer

Private Const col_nota% = 1
Private Const col_tipo% = 2
Private Const col_situ% = 3
Private Const col_emissao% = 4
Private Const col_razao% = 5
Private Const col_valor% = 6
Private Const COL_APEL% = 7
Private Const col_trans% = 8
Private Const col_codTrans% = 9
Private Const col_vlrTrans% = 10
Private Const strTituloGridParcela$ = "campo=parcela;label=Parcela;tamanho=1000|" & _
                                      "campo=vencimento;label=Vencimento;tamanho=1800|" & _
                                      "campo=valor;label=Valor;tamanho=2000;formato=###,##0.00"
Private MatrizDAO As New cMatrizContabilizacaoDAO
Private matriz As cMatrizContabilizacao
'Ivo Sousa (21/08/2012) - Criado a variavel para possibilitar que o sistema guarde o numero do conhecimento original
Private mlngNrConhecimentoOld As Long
Private bGeraDup As Boolean

Private Sub cboNotaFiscalEntradaSaida_Click()
    'pt. 85573 - Moacir Pfau(12/05/2008)
    If cboNotaFiscalEntradaSaida = "Saída" Then
        chkConhecimentoAdionarFreteMercadoria.value = 0
        chkConhecimentoAdionarFreteMercadoria.Enabled = False
    ElseIf cboNotaFiscalEntradaSaida = "Entrada" Then
        chkConhecimentoAdionarFreteMercadoria.Enabled = True
    End If
End Sub

Private Sub cboNotaFiscalEntradaSaida_GotFocus()
    tabConhecimentoPagar.Tab = 0
End Sub

'Projeto: 1222 - História: #9972 - Ivo Sousa (12/04/2012)
Private Sub cboTipoGlobal_Click()
    cboTipoGlobalTitulo.Text = cboTipoGlobal.Text
End Sub

Private Sub chkConhecimentoAdionarFreteMercadoria_GotFocus()
    tabConhecimentoPagar.Tab = 1
End Sub

Private Sub cmdExcluir_Click()
    Call LibProc(WL_DELETAR)
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
End Sub

Private Sub cmdNovo_Click()
    Call LibProc(WL_NOVO)
End Sub

Private Sub cmdPesquisar_Click()
    frmConsultaConhecimentoFretePagar.Show vbModal
    tabConhecimentoPagar.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set navigator = Nothing
End Sub

Private Sub optConhecimentoCIF_Click()
    If txtConhecimentoOperacaoContabil.Enabled Then
        Set matriz = MatrizDAO.Carregar(cboTipoGlobal.Text)
        If Not matriz Is Nothing Then
            txtConhecimentoOperacaoContabil.valorInteiro = matriz.conhecimentoEntradaCif
        Else
            txtConhecimentoOperacaoContabil.valorInteiro = 0
        End If
        Set matriz = Nothing
    Else
        txtConhecimentoOperacaoContabil.valorInteiro = 0
    End If
End Sub

Private Sub optConhecimentoFob_Click()
    If txtConhecimentoOperacaoContabil.Enabled Then
        Set matriz = MatrizDAO.Carregar(cboTipoGlobal.Text)
        If Not matriz Is Nothing Then
            txtConhecimentoOperacaoContabil.valorInteiro = matriz.conhecimentoEntradaFob
        Else
            txtConhecimentoOperacaoContabil.valorInteiro = 0
        End If
        Set matriz = Nothing
    Else
        txtConhecimentoOperacaoContabil.valorInteiro = 0
    End If
End Sub

Private Sub tabConhecimentoPagar_Click(PreviousTab As Integer)
    Select Case tabConhecimentoPagar.Tab
        Case 0
            If cboNotaFiscalEntradaSaida.Enabled Then
                cboNotaFiscalEntradaSaida.SetFocus
            End If
        Case 1
            If txtConhecimentoTransportadora.Enabled Then
                txtConhecimentoTransportadora.SetFocus
            End If
        Case 2
            If txtTituloEmpresa.Enabled Then
                txtTituloEmpresa.SetFocus
            End If
    End Select
End Sub

Private Sub txtChaveAcessoEnt_KeyPress(KeyAscii As Integer)
    If IsCharAlfa(CByte(KeyAscii)) Or IsPunct(CByte(KeyAscii)) Then
        Sendkeys "{BACKSPACE}"
    End If
End Sub

Private Sub txtChaveAcessoEnt_LostFocus()
    If txtChaveAcessoEnt.Text <> "" And txtChaveAcessoEnt.Text <> "0" Then
        If Not ValidaChaveAcesso Then
            Call MsgBox("A Chave de Acesso informada é invalida.", vbInformation, NomeModulo)
        End If
    End If
End Sub

Private Sub txtConhecimentoAcrescimo_GotFocus()
    With txtConhecimentoAcrescimo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoBaseIcms_GotFocus()
    With txtConhecimentoBaseIcms
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoConsignatario_GotFocus()
    With txtConhecimentoConsignatario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoDataEmissao_GotFocus()
    With txtConhecimentoDataEmissao
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoDataEmissao_LostFocus()
    'pt. 87834 - Moacir Pfau(15/07/2008)
    If txtConhecimentoDataEmissao.Text <> "" Then
        If Not EData(txtConhecimentoDataEmissao.Text) Then
          MsgBox "Data informada inválida."
          txtConhecimentoDataEmissao.SetFocus
          Exit Sub
        End If
    End If
End Sub

Private Sub txtConhecimentoDataEntrada_GotFocus()
    With txtConhecimentoDataEntrada
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoDataEntrada_KeyPress(KeyAscii As Integer)
    Call SetMascara(KeyAscii, txtConhecimentoDataEntrada.SelStart, MASK_DATA)
End Sub

Private Sub txtConhecimentoDataEntrada_LostFocus()
    'pt. 87834 - Moacir Pfau(15/07/2008)
    If txtConhecimentoDataEntrada.Text <> "" Then
        If Not EData(txtConhecimentoDataEntrada.Text) Then
          MsgBox "Data informada inválida."
          txtConhecimentoDataEntrada.SetFocus
          Exit Sub
        End If
    End If
End Sub

Private Sub txtConhecimentoDesconto_GotFocus()
    With txtConhecimentoDesconto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoDestinatario_GotFocus()
    With txtConhecimentoDestinatario
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoDistancia_GotFocus()
    With txtConhecimentoDistancia
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoDistancia_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_decimal, txtConhecimentoDistancia, 4)
End Sub

Private Sub txtConhecimentoNaturezaComplemento_Change()
    If Trim(txtConhecimentoNaturezaOperacao.Text) <> "" Then
        lblConhecimentoNaturezaOperacao.Caption = descricaoNatureza(Trim(txtConhecimentoNaturezaOperacao.Text), Trim(txtConhecimentoNaturezaComplemento.Text))
    Else
        lblConhecimentoNaturezaOperacao.Caption = ""
    End If
End Sub

Private Sub txtConhecimentoNaturezaComplemento_GotFocus()
    With txtConhecimentoNaturezaComplemento
        .Tag = .Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoNaturezaComplemento_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Dim strSql As String
        strSql = "SELECT * FROM [Naturezas de Operação] WHERE "
        strSql = strSql & "[Código] LIKE '%" & Trim(txtConhecimentoNaturezaOperacao.Text) & "%'"
        Call PMultiCampo("Consulta de Natureza de Operacao", strSql, pbCampo, "Código;Complemento", txtConhecimentoNaturezaOperacao, txtConhecimentoNaturezaComplemento)
    End If
End Sub

Private Sub txtConhecimentoNaturezaComplemento_LostFocus()
    txtConhecimentoBaseIcms.Text = valorBaseICMS
    If txtConhecimentoNaturezaComplemento.Text <> txtConhecimentoNaturezaComplemento.Tag Then
        Call txtConhecimentoNaturezaOperacao_LostFocus
    End If
End Sub

Private Sub txtConhecimentoNaturezaOperacao_GotFocus()
    With txtConhecimentoNaturezaOperacao
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoNaturezaOperacao_LostFocus()
    If txtConhecimentoNaturezaOperacao.Text = "" Then
        Exit Sub
    End If
        
    Dim sSql As String
    Dim conn As clsConexao
    Dim rstDup As ADODB.Recordset
    Set conn = New clsConexao
    Set rstDup = New ADODB.Recordset
    sSql = "Select [Gerar Duplicatas] From [Naturezas de Operação] Where Código = '" & txtConhecimentoNaturezaOperacao.Text & "' And Complemento = '" & txtConhecimentoNaturezaComplemento.Text & "'"
    If conn.Query(sSql, rstDup) <> 0 Then
        MsgBox "Ocorreu um erro na consulta da Natureza de Operação."
        GoTo desconecta
    End If
    
    If rstDup.Recordcount = 0 Then
        MsgBox "Natureza de Operação inválida."
        txtConhecimentoNaturezaOperacao.Text = ""
        txtConhecimentoNaturezaComplemento.Text = ""
    Else
        bGeraDup = rstDup.Fields.item("Gerar Duplicatas").value
        labTipo.FontBold = bGeraDup
        labConhecimento.FontBold = bGeraDup
        labEmpresa.FontBold = bGeraDup
        labEmissao.FontBold = bGeraDup
        labCondPagto.FontBold = bGeraDup
        labValor.FontBold = bGeraDup
        fraParcelas.FontBold = bGeraDup
        labValorTotal.FontBold = bGeraDup
    End If
    
desconecta:
    rstDup.Close
    conn.Disconnect
    Set conn = Nothing
End Sub

Private Sub txtConhecimentoNumero_GotFocus()
    With txtConhecimentoNumero
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    txtConhecimentoNumero.Tag = txtConhecimentoNumero.Text
End Sub

Private Sub txtConhecimentoNumero_LostFocus()
    If txtConhecimentoNumero.Text = txtConhecimentoNumero.Tag Then
        Exit Sub
    End If
    
    If Not (IsNumeric(txtConhecimentoNumero.Text) And IsNumeric(txtConhecimentoTransportadora.Text)) Then
        Exit Sub
    End If
        
    Dim dao As New cFretePagarDAO
    Dim colecaoTemp As cColecaoNotaFiscal
    Dim lngTransportadora As Long
    Dim lngConhecimento As Long
    Dim booHabilitaTransp As Boolean
    
    'Projeto: 1222 - História: #9972 - Ivo Sousa (13/04/2012)
    If dao.existir(CLng(txtConhecimentoNumero.Text), CLng(txtConhecimentoTransportadora.Text), cboTipoGlobal.Text) Then
        'pt. 86013 - Ivo Sousa(20/05/2008)
        'Set objConhecimento = dao.Carregar(CLng(txtConhecimentoNumero.Text), CLng(txtConhecimentoTransportadora.Text))
        'Call LimpaCampos
        'Call mostraCamposClasse
        
        'pt. 85573 - Moacir Pfau(12/05/2008)
        If Not chkConhecimentoAdionarFreteMercadoria.Enabled Then
           chkConhecimentoAdionarFreteMercadoria.value = 0
        End If
    Else
        If Not booAlterando Then
            Set colecaoTemp = objConhecimento.notasFiscais
            Set objConhecimento = New cFretePagar
            lngTransportadora = CLng(strToDbl(txtConhecimentoTransportadora.Text))
            booHabilitaTransp = txtConhecimentoTransportadora.Enabled
            lngConhecimento = CLng(strToDbl(txtConhecimentoNumero.Text))
            Call LimpaCampos(False)
            txtConhecimentoTransportadora.Text = lngTransportadora
            txtConhecimentoTransportadora.Enabled = booHabilitaTransp
            txtConhecimentoNumero.Text = lngConhecimento
            objConhecimento.notasFiscais = colecaoTemp
        End If
    End If
End Sub

Private Sub txtConhecimentoOperacaoContabil_Change()
    If txtConhecimentoOperacaoContabil.valorInteiro > 0 Then
        lblConhecimentoOperacaoContabil.Caption = descricaoOperacao(txtConhecimentoOperacaoContabil.valorInteiro)
    Else
        lblConhecimentoOperacaoContabil.Caption = ""
    End If
End Sub

Private Sub txtConhecimentoOperacaoContabil_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        Call PCampo("Operações contábeis", "select * from OperacaoContabil", pbCampo, txtConhecimentoOperacaoContabil, "cd_operacao")
    End If
End Sub

Private Sub txtConhecimentoOutros_GotFocus()
    With txtConhecimentoOutros
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoPedagio_GotFocus()
    With txtConhecimentoPedagio
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoPorcentagemIcms_GotFocus()
    With txtConhecimentoPorcentagemIcms
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoRedespacho_GotFocus()
    With txtConhecimentoRedespacho
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoRemetente_GotFocus()
    With txtConhecimentoRemetente
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoSeguro_GotFocus()
    With txtConhecimentoSeguro
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoTarifa_GotFocus()
    With txtConhecimentoTarifa
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoTarifa_LostFocus()
    txtConhecimentoBaseIcms.Text = valorBaseICMS
End Sub

Private Sub txtConhecimentoTransportadora_GotFocus()
    tabConhecimentoPagar.Tab = 1
    With txtConhecimentoTransportadora
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoValor_GotFocus()
    booValorFoco = True
    With txtConhecimentoValor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoValor_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoValor)
End Sub

Private Sub txtConhecimentoValor_LostFocus()
    booValorFoco = False
End Sub

Private Sub txtConhecimentoValorIsentas_GotFocus()
    With txtConhecimentoValorIsentas
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtConhecimentoVolume_Change()
    Dim dblValor As Double
    
    If Not booValorFoco Then
        dblValor = strToDbl(txtConhecimentoVolume.Text) * strToDbl(txtConhecimentoTarifa.Text)
        txtConhecimentoValor.Text = Format(dblValor, "##,##0.00")
    End If
End Sub

Private Sub txtConhecimentoVolume_GotFocus()
    With txtConhecimentoVolume
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtNotaFiscalDataEmissaoFinal_LostFocus()
    'pt. 87834 - Moacir Pfau(15/07/2008)
    If txtNotaFiscalDataEmissaoFinal.Text <> "" Then
        If Not EData(txtNotaFiscalDataEmissaoFinal.Text) Then
          MsgBox "Data informada inválida."
          txtNotaFiscalDataEmissaoFinal.SetFocus
          Exit Sub
        End If
    End If
End Sub

Private Sub txtNotaFiscalDataEmissaoInicial_LostFocus()
    'pt. 87834 - Moacir Pfau(15/07/2008)
    If txtNotaFiscalDataEmissaoInicial.Text <> "" Then
        If Not EData(txtNotaFiscalDataEmissaoInicial.Text) Then
          MsgBox "Data informada inválida."
          txtNotaFiscalDataEmissaoInicial.SetFocus
          Exit Sub
        End If
    End If
End Sub

Private Sub txtTituloBanco_Change()
    If IsNumeric(txtTituloBanco.Text) Then
        lblTituloBanco.Caption = NomeBanco(strToLng(txtTituloBanco.Text))
    Else
        lblTituloBanco.Caption = ""
    End If
End Sub

Private Sub txtTituloBanco_GotFocus()
    With txtTituloBanco
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTituloBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta de Bancos", "SELECT * FROM Bancos", pbCampo, txtTituloBanco, "Banco")
    End If
End Sub

Private Sub txtTituloBanco_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtTituloCarteira_GotFocus()
    With txtTituloCarteira
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTituloCarteira_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Dim strSql As String
        strSql = "SELECT * FROM Carteiras"
        If IsNumeric(txtTituloBanco.Text) Then
            strSql = strSql & " WHERE Banco = " & txtTituloBanco.Text
        End If
        Call PCampo("Carteiras", strSql, pbCampo, txtTituloCarteira, "Carteira")
    End If
End Sub

Private Sub txtTituloCentroCusto_Change()
    If IsNumeric(txtTituloCentroCusto.Text) Then
        lblTituloCentroCusto.Caption = descricaoCentroCusto(strToLng(txtTituloCentroCusto.Text))
    Else
        lblTituloCentroCusto.Caption = ""
    End If
End Sub

Private Sub txtTituloCentroCusto_GotFocus()
    With txtTituloCentroCusto
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTituloCentroCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Centro de custo", "SELECT * FROM Centros", pbCampo, txtTituloCentroCusto, "Código")
    End If
End Sub

Private Sub txtTituloCentroCusto_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtTituloCondicaoPagamento_Change()
    If txtTituloCondicaoPagamento.Text <> "" Then
        If objConhecimento.Titulo.CondicaoPagamento.Carregar(txtTituloCondicaoPagamento.Text) Then
            lblTituloCondicaoPagamento.Caption = objConhecimento.Titulo.CondicaoPagamento.Descricao
        Else
            lblTituloCondicaoPagamento.Caption = ""
        End If
    Else
            lblTituloCondicaoPagamento.Caption = ""
    End If
    Call preencheGridParcela
    Exit Sub
End Sub

Private Sub txtTituloCondicaoPagamento_GotFocus()
    With txtTituloCondicaoPagamento
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTituloCondicaoPagamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta Condição de Pagamento", "SELECT * FROM [Condições de Pagamento]", pbCampo, txtTituloCondicaoPagamento, "Código")
    End If
End Sub

Private Sub txtTituloCondicaoPagamento_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtTituloCondicaoPagamento_LostFocus()
    If lblTituloCondicaoPagamento.Caption = "" Then
        objConhecimento.Titulo.parcelas = New cColecaoParcela
    End If
End Sub

Private Sub txtTituloConhecimento_GotFocus()
    tabConhecimentoPagar.Tab = 2
End Sub

Private Sub txtTituloConhecimento_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtTituloConta_Change()
    If IsNumeric(txtTituloConta.Text) Then
        lblTituloConta.Caption = descricaoConta(strToLng(txtTituloConta.Text))
    Else
        lblTituloConta.Caption = ""
    End If
End Sub

Private Sub txtTituloConta_GotFocus()
    With txtTituloConta
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTituloConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta de Contas", "SELECT * FROM Contas", pbCampo, txtTituloConta, "Código")
    End If
End Sub

Private Sub txtTituloConta_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtTituloControle_GotFocus()
    With txtTituloControle
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTituloEmissao_KeyPress(KeyAscii As Integer)
    Call SetMascara(KeyAscii, txtTituloEmissao.SelStart, MASK_DATA)
End Sub

Private Sub txtTituloEmissao_LostFocus()
    'pt. 87834 - Moacir Pfau(15/07/2008)
    If txtTituloEmissao.Text <> "" Then
        If Not EData(txtTituloEmissao.Text) Then
          MsgBox "Data informada inválida."
          txtTituloEmissao.SetFocus
          Exit Sub
        End If
    End If
End Sub

Private Sub txtTituloEmpresa_Change()
    If Trim(txtTituloEmpresa.Text) <> "" Then
        lblTituloEmpresa.Caption = dadosEmpresa(txtTituloEmpresa.Text, booFornec:=True)
        Call infoFinanceiras(txtTituloEmpresa.Text)
    Else
        lblTituloEmpresa.Caption = ""
    End If
End Sub

Private Sub txtTituloEmpresa_GotFocus()
    tabConhecimentoPagar.Tab = 2
End Sub

Private Sub txtTituloEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta de empresa", "SELECT * FROM Empresas WHERE Tipo <> 'Cliente'", pbCampo, txtTituloEmpresa, "Apel")
    End If
End Sub

Private Sub txtTituloParcelaNumero_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtTituloParcelaValor_GotFocus()
    With txtTituloParcelaValor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTituloParcelaValor_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtTituloParcelaValor)
End Sub

Private Sub txtTituloParcelaVencimento_KeyPress(KeyAscii As Integer)
    Call SetMascara(KeyAscii, txtTituloParcelaVencimento.SelStart, MASK_DATA)
End Sub

Private Sub txtTituloValor_GotFocus()
    With txtTituloValor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtTituloValor_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtTituloValor)
End Sub

Private Sub cmdNotaFiscalExecutar_Click()
    Screen.MousePointer = vbHourglass
    If validaExecutar Then
        Dim cmd As IDBSelectCommand
        Dim oJoinTransp As CDBJoin
        Dim oJoinEmp As CDBJoin
        Dim rdResult As IDBReader
        Dim strTabela As String
        
        If cboNotaFiscalEntradaSaida.ListIndex = 0 Then
            strTabela = "[Notas Fiscais de Entrada]"
        Else
            strTabela = "[Notas Fiscais de Saída]"
        End If
           
        Aplicacao.Connect
        Set cmd = Aplicacao.CreateSelectCommand
        If cboNotaFiscalEntradaSaida.ListIndex = 0 Then
            cmd.SelectClause = "Situação, [Número], [Tipo de Registro], [Emissão], Fornecedor AS Empresa, [Valor Total], Empresas.[Razão], Transportadora, Transportadoras.[Razão] AS nome_transp, [Valor do Frete]"
        Else
            cmd.SelectClause = "Situação, [Número], [Tipo de Registro], [Emissão], Empresa, [Valor Total], Empresas.[Razão], Transportadora, Transportadoras.[Razão] AS nome_transp, [Valor do Frete]"
        End If
        cmd.Table.TableName = strTabela
        
        'Join com a tabela de Empresas
        Set oJoinEmp = New CDBJoin
        oJoinEmp.init
        oJoinEmp.JoinType = dbJoinTypeLeft
        oJoinEmp.LeftTable.TableName = strTabela
        oJoinEmp.RightTable.TableName = "Empresas"
        If cboNotaFiscalEntradaSaida.ListIndex = 0 Then
            Call oJoinEmp.AddJoinField("Fornecedor", "Apel")
        Else
            Call oJoinEmp.AddJoinField("Empresa", "Apel")
        End If
        Call cmd.AddJoin(oJoinEmp)
        
        'Join com a tabela de transportadoras
        Set oJoinTransp = New CDBJoin
        oJoinTransp.init
        oJoinTransp.JoinType = dbJoinTypeLeft
        oJoinTransp.LeftTable.TableName = strTabela
        oJoinTransp.RightTable.TableName = "Transportadoras"
        Call oJoinTransp.AddJoinField("Transportadora", "Código")
        Call cmd.AddJoin(oJoinTransp)
        
        'Critério para a data de emissão da Nota Fiscal
        If IsDate(txtNotaFiscalDataEmissaoInicial.Text) Then
            Call cmd.Filter.Append("[Emissão]>=@pEmissaoInicial")
            Call cmd.Parameters.add(cmd.CreateParameter("@pEmissaoInicial", txtNotaFiscalDataEmissaoInicial.Text, dbFieldTypeDate))
        End If
        If IsDate(txtNotaFiscalDataEmissaoFinal.Text) Then
            Call cmd.Filter.Append("[Emissão]<=@pEmissaoFinal")
            Call cmd.Parameters.add(cmd.CreateParameter("@pEmissaoFinal", txtNotaFiscalDataEmissaoFinal.Text, dbFieldTypeDate))
        End If
        'Critério para o tipo da Nota Fiscal
        Call cmd.Filter.Append("[Tipo de Registro] = @pTipoRegistro")
        Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", cboNotaFiscalTipo.Text, dbFieldTypeString))
        'Critério para verificar se a nota final foi informada
        If IsNumeric(txtNotaFiscalNotaFinal.Text) Then
            'Critério para verificar se a nota inicial também foi informada
            If IsNumeric(txtNotaFiscalNotaInicial.Text) Then
                If strToLng(txtNotaFiscalNotaFinal.Text) >= strToLng(txtNotaFiscalNotaInicial.Text) Then
                    Call cmd.Filter.Append("[Número] >= @pNumeroInicial")
                    Call cmd.Parameters.add(cmd.CreateParameter("@pNumeroInicial", strToLng(txtNotaFiscalNotaInicial.Text), dbFieldTypeLong))
                End If
                Call cmd.Filter.Append("[Número] <= @pNumeroFinal")
                Call cmd.Parameters.add(cmd.CreateParameter("@pNumeroFinal", strToLng(txtNotaFiscalNotaFinal.Text), dbFieldTypeLong))
            End If
        Else
            'Critério para a nota inicial informada
            If IsNumeric(txtNotaFiscalNotaInicial.Text) Then
                Call cmd.Filter.Append("[Número] >= @pNumeroInicial")
                Call cmd.Parameters.add(cmd.CreateParameter("@pNumeroInicial", strToLng(txtNotaFiscalNotaInicial.Text), dbFieldTypeLong))
            End If
        End If
        If IsNumeric(txtNotaFiscalTransportadora.Text) Then
            If nomeTransportadora(strToLng(txtNotaFiscalTransportadora.Text)) <> "" Then
                Call cmd.Filter.Append("Transportadora = @pTransportadora")
                Call cmd.Parameters.add(cmd.CreateParameter("@pTransportadora", strToLng(txtNotaFiscalTransportadora.Text), dbFieldTypeLong))
            End If
        End If
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
        lstNotaFiscalNotas.ListItems.Clear
        objConhecimento.notasFiscais.Clear
        If Not rdResult.EOF Then
            Call mostraNotasFiscais(rdResult)
        End If
        rdResult.CloseReader
        Set rdResult = Nothing
        Set cmd = Nothing
        Aplicacao.Disconnect
        Call mostraCamposNotasFiscais
    Else
        MsgBox "Informe algum campo do filtro.", vbInformation, Me.Caption
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
End Sub

Private Sub cmdTituloCalcular_Click()
    If validaCalculoParcelas Then
        If objConhecimento.Titulo.CondicaoPagamento.Codigo = 0 Then
            If Not objConhecimento.Titulo.CondicaoPagamento.Carregar(txtTituloCondicaoPagamento.Text) Then
                MsgBox "Informe uma condição de pagamento válida.", vbInformation, Me.Caption
            Else
                intCondPag = objConhecimento.Titulo.CondicaoPagamento.Codigo
            End If
        Else
            intCondPag = objConhecimento.Titulo.CondicaoPagamento.Codigo
        End If
        With objConhecimento.Titulo
            .valor = CDbl(txtTituloValor.Text)
            .Emissao = CDate(txtTituloEmissao.Text)
            .parcelar
        End With
        Call preencheGridParcela
    End If
End Sub

Private Sub cmdTituloParcelaCancelar_Click()
    txtTituloParcelaNumero.Text = ""
    txtTituloParcelaVencimento.Text = ""
    txtTituloParcelaValor.Text = ""
    cmdTituloParcelaConfirmar.Enabled = False
    Label49.Enabled = False
    txtTituloParcelaVencimento.Enabled = False
    Label50.Enabled = False
    txtTituloParcelaValor.Enabled = False
End Sub

Private Sub cmdTituloParcelaConfirmar_Click()
    Dim objParcela As cParcela
    If validaCamposParcela Then
        With objConhecimento.Titulo
            Set objParcela = New cParcela
            objParcela.Parcela = CInt(txtTituloParcelaNumero.Text)
            objParcela.vencimento = CDate(txtTituloParcelaVencimento.Text)
            objParcela.valor = CDbl(txtTituloParcelaValor.Text)
            Call .parcelas.update(objParcela)
        End With
        Call preencheGridParcela
        Call cmdTituloParcelaCancelar_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 7995
    Me.Width = 13140
    CenterForm Me
    Set objConhecimento = New cFretePagar
    Call preencheComboTipos
    Label53.Enabled = Configuracao("Utiliza Integração Contábil", False)
    txtConhecimentoOperacaoContabil.Enabled = Configuracao("Utiliza Integração Contábil", False)
    lblConhecimentoOperacaoContabil.Enabled = Configuracao("Utiliza Integração Contábil", False)
    Call LimpaCampos
    booValorFoco = False
End Sub

Private Sub grdTituloParcelas_DblClick()
    With grdTituloParcelas
        If .Row > 0 And objConhecimento.PermiteAlteracao Then
            txtTituloParcelaNumero.Text = .TextMatrix(.Row, 0)
            txtTituloParcelaVencimento.Text = .TextMatrix(.Row, 1)
            txtTituloParcelaValor.Text = .TextMatrix(.Row, 2)
            Label49.Enabled = True
            txtTituloParcelaVencimento.Enabled = True
            Label50.Enabled = True
            txtTituloParcelaValor.Enabled = True
            cmdTituloParcelaConfirmar.Enabled = True
            cmdTituloParcelaCancelar.Enabled = True
        End If
    End With
End Sub

Private Sub lstNotaFiscalNotas_Click()
    Dim objNota As New cFretePagarNotaFiscal
    If lngItem > 0 Then
        With lstNotaFiscalNotas
            objNota.Nota = .ListItems.item(lngItem).SubItems(col_nota)
            objNota.tipo = .ListItems.item(lngItem).SubItems(col_tipo)
            objNota.situacao = .ListItems.item(lngItem).SubItems(col_situ)
            objNota.Emissao = .ListItems.item(lngItem).SubItems(col_emissao)
            objNota.Empresa = .ListItems.item(lngItem).SubItems(col_razao)
            objNota.Apel = .ListItems.item(lngItem).SubItems(COL_APEL)
            objNota.valor = .ListItems.item(lngItem).SubItems(col_valor)
            objNota.Transportadora = .ListItems.item(lngItem).SubItems(col_codTrans)
            
            If .ListItems(lngItem).SmallIcon = "item" Then
                If objConhecimento.notasFiscais.Transportadora = 0 Or objConhecimento.notasFiscais.Count = 0 Then
                    objConhecimento.notasFiscais.Transportadora = objNota.Transportadora
                    objConhecimento.notasFiscais.registroEntrada = CBool(cboNotaFiscalEntradaSaida.ListIndex = 0)
                    txtConhecimentoTransportadora.Text = objNota.Transportadora
                    If objNota.Transportadora > 0 Then
                        txtConhecimentoTransportadora.Enabled = False
                        Label8.Enabled = False
                    End If
                End If
                If objConhecimento.notasFiscais.Transportadora = objNota.Transportadora Then
                    .ListItems(lngItem).SmallIcon = "selecionado"
                    Call objConhecimento.notasFiscais.add(objNota)
                Else
                    MsgBox "Só é possivel incluir notas da mesma transportadora.", vbInformation, Me.Caption
                End If
            Else
                .ListItems(lngItem).SmallIcon = "item"
                Call objConhecimento.notasFiscais.Remove(objNota)
                If objConhecimento.notasFiscais.Count = 0 Then
                    txtConhecimentoTransportadora.Text = ""
                    txtConhecimentoTransportadora.Enabled = True
                    Label8.Enabled = True
                End If
            End If
        End With
        lngItem = 0
        Call mostraCamposNotasFiscais
    End If
End Sub

Private Sub lstNotaFiscalNotas_ItemClick(ByVal item As MSComctlLib.ListItem)
    lngItem = item.Index
End Sub

Private Sub txtConhecimentoAcrescimo_Change()
    txtConhecimentoEncargosValor.Text = Format(valorComEncargos, "##,##0.00")
End Sub

Private Sub txtConhecimentoAcrescimo_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoAcrescimo)
End Sub

Private Sub txtConhecimentoBaseIcms_Change()
    txtConhecimentoValorIcms.Text = Format(ValorICMS, "##,##0.00")
End Sub

Private Sub txtConhecimentoBaseIcms_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoBaseIcms)
End Sub

Private Sub txtConhecimentoConsignatario_Change()
    If Trim(txtConhecimentoConsignatario.Text) <> "" Then
        lblConhecimentoConsignatario.Caption = dadosEmpresa(txtConhecimentoConsignatario.Text, booFornec:=True)
        txtTituloEmpresa.Text = txtConhecimentoConsignatario.Text
    Else
        lblConhecimentoConsignatario.Caption = ""
    End If
End Sub

Private Sub txtConhecimentoConsignatario_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta de Consinatário", "SELECT * FROM Empresas WHERE Tipo <> 'Cliente'", pbCampo, txtConhecimentoConsignatario, "Apel")
    End If
End Sub

Private Sub txtConhecimentoDataEmissao_KeyPress(KeyAscii As Integer)
    Call SetMascara(KeyAscii, txtConhecimentoDataEmissao.SelStart, MASK_DATA)
End Sub

Private Sub txtConhecimentoDesconto_Change()
    txtConhecimentoEncargosValor.Text = Format(valorComEncargos, "##,##0.00")
End Sub

Private Sub txtConhecimentoDesconto_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoDesconto)
End Sub

Private Sub txtConhecimentoDestinatario_Change()
    If Trim(txtConhecimentoDestinatario.Text) <> "" Then
        If objUfDestino Is Nothing Then
            Set objUfDestino = New cEstado
        End If
        lblConhecimentoDestinatario.Caption = dadosEmpresa(txtConhecimentoDestinatario.Text, objUfDestino)
    Else
        lblConhecimentoDestinatario.Caption = ""
    End If
    Call atualizaICMS
End Sub

Private Sub txtConhecimentoDestinatario_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta de Destinatário", "SELECT * FROM Empresas", pbCampo, txtConhecimentoDestinatario, "Apel")
    End If
End Sub

Private Sub txtConhecimentoEncargosValor_Change()
    txtConhecimentoBaseIcms.Text = Format(txtConhecimentoEncargosValor.Text, "##,##0.00")
    txtConhecimentoValorTotal.Text = Format(txtConhecimentoEncargosValor.Text, "##,##0.00")
End Sub

Private Sub txtConhecimentoNaturezaOperacao_Change()
    If Trim(txtConhecimentoNaturezaOperacao.Text) <> "" Then
        lblConhecimentoNaturezaOperacao.Caption = descricaoNatureza(Trim(txtConhecimentoNaturezaOperacao.Text), Trim(txtConhecimentoNaturezaComplemento.Text))
    Else
        lblConhecimentoNaturezaOperacao.Caption = ""
    End If
End Sub

Private Sub txtConhecimentoNaturezaOperacao_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        strSql = "SELECT * FROM [Naturezas de Operação] WHERE "
        strSql = strSql & "[Código] LIKE '1%' OR "
        strSql = strSql & "[Código] LIKE '2%' OR "
        strSql = strSql & "[Código] LIKE '3%' AND "
        strSql = strSql & "[Tipo de Movimentação] = 'Entrada'"
        Call PMultiCampo("Consulta de Natureza de Operacao", strSql, pbCampo, "Código;Complemento", txtConhecimentoNaturezaOperacao, txtConhecimentoNaturezaComplemento)
    End If
End Sub

Private Sub txtConhecimentoNaturezaOperacao_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtConhecimentoNumero_Change()
    txtTituloConhecimento.Text = txtConhecimentoNumero.Text
End Sub

Private Sub txtConhecimentoNumero_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtConhecimentoOutros_Change()
    txtConhecimentoEncargosValor.Text = Format(valorComEncargos, "##,##0.00")
End Sub

Private Sub txtConhecimentoOutros_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoOutros)
End Sub

Private Sub txtConhecimentoPedagio_Change()
    txtConhecimentoEncargosValor.Text = Format(valorComEncargos, "##,##0.00")
End Sub

Private Sub txtConhecimentoPedagio_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoPedagio)
End Sub

Private Sub txtConhecimentoPorcentagemIcms_Change()
    txtConhecimentoValorIcms.Text = Format(ValorICMS, "##,##0.00")
End Sub

Private Sub txtConhecimentoPorcentagemIcms_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoPorcentagemIcms)
End Sub

Private Sub txtConhecimentoRedespacho_Change()
    If IsNumeric(txtConhecimentoRedespacho.Text) Then
        lblConhecimentoRedespacho.Caption = nomeTransportadora(strToLng(txtConhecimentoRedespacho.Text))
    Else
        lblConhecimentoRedespacho.Caption = ""
    End If
End Sub

Private Sub txtConhecimentoRedespacho_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta de Redespacho", "SELECT [Código], Apel, [Razão] FROM Transportadoras", pbCampo, txtConhecimentoRedespacho, "Código")
    End If
End Sub

Private Sub txtConhecimentoRedespacho_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtConhecimentoRemetente_Change()
    If Trim(txtConhecimentoRemetente.Text) <> "" Then
        If objUfOrigem Is Nothing Then
            Set objUfOrigem = New cEstado
        End If
        lblConhecimentoRemetente.Caption = dadosEmpresa(txtConhecimentoRemetente.Text, objUfOrigem)
    Else
        lblConhecimentoRemetente.Caption = ""
    End If
    Call atualizaICMS
End Sub

Private Sub txtConhecimentoRemetente_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta de Remetente", "SELECT * FROM Empresas", pbCampo, txtConhecimentoRemetente, "Apel")
    End If
End Sub

Private Sub txtConhecimentoSeguro_Change()
    txtConhecimentoEncargosValor.Text = Format(valorComEncargos, "##,##0.00")
End Sub

Private Sub txtConhecimentoSeguro_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoSeguro)
End Sub

Private Sub txtConhecimentoTarifa_Change()
    Dim dblValor As Double
    
    If Not booValorFoco Then
        dblValor = strToDbl(txtConhecimentoVolume.Text) * strToDbl(txtConhecimentoTarifa.Text)
        txtConhecimentoValor.Text = Format(dblValor, "##,##0.00")
    End If
End Sub

Private Sub txtConhecimentoTarifa_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_decimal, txtConhecimentoTarifa, 5)
End Sub

Private Sub txtConhecimentoTransportadora_Change()
    If IsNumeric(txtConhecimentoTransportadora) Then
        lblConhecimentoTransportadora.Caption = nomeTransportadora(CInt(txtConhecimentoTransportadora.Text))
    Else
        lblConhecimentoTransportadora.Caption = ""
    End If
End Sub

Private Sub txtConhecimentoTransportadora_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        'pt. 86426 - Ivo Sousa (01/07/2008) / 'pt. 87361 - Moacir Pfau(02/07/2008) - INCLUSÃO DO CAMPO "[IEst/RG]".
        Call PCampo("Consulta de transportadoras", "SELECT [Código], Apel, [Razão], [CNPJ/CPF], [IEst/RG] FROM Transportadoras", pbCampo, txtConhecimentoTransportadora, "Código")
    End If
End Sub

Private Sub txtConhecimentoTransportadora_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtConhecimentoValor_Change()
    If booValorFoco Then
        If IsNumeric(txtConhecimentoValor.Text) Then
            If strToDbl(txtConhecimentoVolume.Text) > 0 And strToDbl(txtConhecimentoValor.Text) > 0 Then
                txtConhecimentoTarifa.Text = Format(strToDbl(txtConhecimentoValor.Text) / strToDbl(txtConhecimentoVolume.Text), "##,##0.00##")
            Else
                txtConhecimentoTarifa.Text = Format("0", "##,##0.00##")
            End If
        Else
            txtConhecimentoTarifa.Text = Format("0", "##,##0.00##")
        End If
    End If
    txtConhecimentoEncargosValor.Text = Format(valorComEncargos, "##,##0.00")
End Sub

Private Sub txtConhecimentoValorIcms_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoValorIcms)
End Sub

Private Sub txtConhecimentoValorIsentas_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_moeda, txtConhecimentoValorIsentas)
End Sub

Private Sub txtConhecimentoValorTotal_Change()
    txtTituloValor.Text = Format(txtConhecimentoValorTotal.Text, "##,##0.00")
End Sub

Private Sub txtConhecimentoVolume_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_decimal, txtConhecimentoVolume)
End Sub

Private Sub txtNotaFiscalDataEmissaoFinal_KeyPress(KeyAscii As Integer)
    Call SetMascara(KeyAscii, txtNotaFiscalDataEmissaoFinal.SelStart, MASK_DATA)
End Sub

Private Sub txtNotaFiscalDataEmissaoInicial_KeyPress(KeyAscii As Integer)
    Call SetMascara(KeyAscii, txtNotaFiscalDataEmissaoInicial.SelStart, MASK_DATA)
End Sub

Private Sub txtNotaFiscalNotaFinal_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_inteiro, txtNotaFiscalDataEmissaoFinal)
End Sub

Private Sub txtNotaFiscalNotaInicial_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_inteiro, txtNotaFiscalNotaInicial)
End Sub

Public Function LibProc(strFuncao As String, Optional lngFuncao As Long) As Boolean
    Dim dao As New cFretePagarDAO
    Dim facIntegra As New cDAOFactory
On Error GoTo erro_libproc
    
    Select Case strFuncao
        Case WL_SAIR
            Unload Me
            Exit Function
        Case WL_NOVO
            Set objConhecimento = New cFretePagar
            Call LimpaCampos
            Call preencheComboTipos
            tabConhecimentoPagar.Tab = 0
            cboNotaFiscalEntradaSaida.SetFocus
        Case WL_SALVAR
            If ValidaCampos Then
                Aplicacao.Connect
                Aplicacao.BeginTransaction
                Call preencheClasse
                If objConhecimento.notasFiscais.Transportadora = 0 Then
                    objConhecimento.notasFiscais.Transportadora = objConhecimento.codigoTransportadora
                End If
                dao.GeraDup = bGeraDup
                If Not booAlterando Then
                    If Not dao.persistir(objConhecimento, Aplicacao, CDate(txtConhecimentoDataEntrada.Text)) Then
                        MsgBox "Ocorreu um erro ao gravar o conhecimento.", vbInformation, Me.Caption
                        Aplicacao.RollbackTransaction
                    Else
                        If chkConhecimentoAdionarFreteMercadoria.value = vbChecked Then
                            If Not objConhecimento.MovimentaEstoque Then
                                Call MsgBox("Não foi possível movimentar estoque. Tente novamente.", vbInformation, NomeModulo)
                                Aplicacao.RollbackTransaction
                            Else
                                Aplicacao.CommitTransaction
                                Call LibProc(WL_NOVO)
                            End If
                        Else
                            Aplicacao.CommitTransaction
                            Call LibProc(WL_NOVO)
                        End If
                    End If
                Else
                    If dao.Atualizar(objConhecimento, Aplicacao, CDate(txtConhecimentoDataEntrada.Text)) Then
                        Aplicacao.CommitTransaction
                        Call LibProc(WL_NOVO)
                    Else
                        MsgBox "Ocorreu um erro ao alterar o conhecimento.", vbInformation, Me.Caption
                        Aplicacao.RollbackTransaction
                    End If
                End If
                Aplicacao.Disconnect
            End If
        Case WL_DELETAR
            If MsgBox("Confirma a exclusão do conhecimento?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                'Projeto: 1222 - História: #9972 - Ivo Sousa (13/04/2012)
                If dao.existir(objConhecimento.numeroConhecimento, objConhecimento.codigoTransportadora, objConhecimento.TipoRegistro) Then
                    If objConhecimento.PermiteAlteracao Then
                        Aplicacao.Connect
                        Aplicacao.BeginTransaction
                        If objConhecimento.rateiaValorProdutos Then
                            If Not objConhecimento.ApagaMovimento Then
                                MsgBox "Não foi possivel excluir a movimentação.", vbInformation, Me.Caption
                            End If
                        End If
                        If dao.Excluir(objConhecimento, Aplicacao) Then
                            Aplicacao.CommitTransaction
                            Call LimpaCampos
                        Else
                            Aplicacao.RollbackTransaction
                            MsgBox "Ocorreu um erro ao tentar excluir o conhecimento", vbInformation, Me.Caption
                        End If
                        Aplicacao.Disconnect
                    Else
                        MsgBox "O conhecimento possui duplicatas quitadas e não pode ser excluido.", vbInformation, Me.Caption
                    End If
                Else
                    MsgBox "Conhecimento não existente, impossivel excluir.", vbInformation, Me.Caption
                End If
            End If
        Case WL_PRIMEIRO
            navigator.MoveFirst
            Call setConhecimento(navigator.CurrentObject)
        Case WL_ANTERIOR
            navigator.MovePrevious
            If Not navigator.BOF Then
                Call setConhecimento(navigator.CurrentObject)
            End If
        Case WL_PROXIMO
            navigator.MoveNext
            If Not navigator.EOF Then
                Call setConhecimento(navigator.CurrentObject)
            End If
        Case WL_ULTIMO
            navigator.MoveLast
            Call setConhecimento(navigator.CurrentObject)
    End Select
    Exit Function
erro_libproc:
    FinallyConnection Aplicacao, True
    MsgBox "Ocorreu um erro tentar persistir o Conhecimento: " & err.Description, vbCritical, Me.Caption
End Function

Private Sub txtNotaFiscalTransportadora_Change()
    If Trim(txtNotaFiscalTransportadora.Text) <> "" Then
        lblNotaFiscalTransportadora.Caption = nomeTransportadora(strToLng(txtNotaFiscalTransportadora.Text))
    Else
        lblNotaFiscalTransportadora.Caption = ""
    End If
End Sub

Private Sub txtNotaFiscalTransportadora_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Consulta de transportadoras", "SELECT * FROM Transportadoras", pbCampo, txtNotaFiscalTransportadora, "Código")
    End If
End Sub

Private Sub txtNotaFiscalTransportadora_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_inteiro, txtNotaFiscalTransportadora)
End Sub

'***********************************************************************************
'*
'*                          Funções de consulta de informações
'*
'***********************************************************************************
Private Function nomeTransportadora(lngCodigo As Long) As String
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "[Razão]"
    cmd.Table.TableName = "transportadoras"
    Call cmd.Filter.Append("[Código] = @pCodigo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCodigo", lngCodigo, dbFieldTypeLong))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        nomeTransportadora = rdResult.GetString("Razão")
    Else
        nomeTransportadora = ""
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

Private Function dadosEmpresa(strApel As String, Optional ByRef estado As cEstado = Nothing, Optional booFornec As Boolean) As String
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Dim dao As cEstadoDAO
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "[Razão], Estado"
    cmd.Table.TableName = "Empresas"
    Call cmd.Filter.Append("Apel = @pApel")
    Call cmd.Parameters.add(cmd.CreateParameter("@pApel", strApel, dbFieldTypeString))
    
    Call cmd.Filter.Append("[Cliente bloqueado] = @pCliente bloqueado")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCliente bloqueado", False, dbFieldTypeBool))
    

    If booFornec Then
        Call cmd.Filter.Append("Tipo <> 'Cliente'")
    End If
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        dadosEmpresa = rdResult.GetString("Razão")
        If Not estado Is Nothing Then
            Set dao = New cEstadoDAO
            Set estado = dao.Carregar(rdResult.GetString("Estado"))
        End If
    Else
        If Not estado Is Nothing Then
            Set estado = Nothing
        End If
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

Private Function descricaoNatureza(strCodigo As String, strCompNatOperacao As String) As String
'    Dim cmd As IDBSelectCommand
'    Dim rdResult As IDBReader
'
'    Aplicacao.Connect
'    Set cmd = Aplicacao.CreateSelectCommand
'    cmd.Table.TableName = "[Naturezas de Operação]"
'    Call cmd.Filter.Append("[Código] LIKE '1%'", dbLogicOperatorOR)
'    Call cmd.Filter.Append("[Código] = @pCodigo")
'    Call cmd.Filter.Append("[Código] LIKE '2%'", dbLogicOperatorOR)
'    Call cmd.Filter.Append("[Código] = @pCodigo")
'    Call cmd.Filter.Append("[Código] LIKE '3%'", dbLogicOperatorOR)
'    Call cmd.Filter.Append("[Código] = @pCodigo")
'    Call cmd.Parameters.add(cmd.CreateParameter("@pCodigo", strCodigo, dbFieldTypeString, 4))
'
'    Call cmd.Filter.Append("[Complemento] LIKE '" & strCompNatOperacao & "'", dbLogicOperatorOR)
'    Call cmd.Filter.Append("[Complemento] = @pComplemento")
'    Call cmd.Parameters.add(cmd.CreateParameter("@pComplemento", strCompNatOperacao, dbFieldTypeString, 4))
'
'    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
'    If Not rdResult.EOF Then
'        descricaoNatureza = rdResult.GetString("Descrição")
'    Else
'        descricaoNatureza = ""
'    End If
'    rdResult.CloseReader
'    Set rdResult = Nothing
'    Set cmd = Nothing
'    Aplicacao.Disconnect
    Dim objNatur        As New CNaturezasdeOperacao
    descricaoNatureza = ""
    If strCodigo <> "" Then
        If objNatur.CarregarRegistro(strCodigo, strCompNatOperacao) Then
            descricaoNatureza = objNatur.Descricao
        End If
    End If
End Function

Private Function NomeBanco(lngCodigo As Long) As String
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "Bancos"
    cmd.SelectClause = "Nome"
    Call cmd.Filter.Append("Banco = @pBanco")
    Call cmd.Parameters.add(cmd.CreateParameter("@pBanco", lngCodigo, dbFieldTypeLong))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        NomeBanco = rdResult.GetString("Nome")
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

Private Function descricaoCentroCusto(lngNumero As Long) As String
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "Centros"
    cmd.SelectClause = "Descrição"
    Call cmd.Filter.Append("Código = @pCodigo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCodigo", lngNumero, dbFieldTypeLong))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        descricaoCentroCusto = rdResult.GetString("Descrição")
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

Private Function descricaoConta(lngNumero As Long) As String
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "Contas"
    cmd.SelectClause = "Descrição"
    Call cmd.Filter.Append("Código = @pCodigo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCodigo", lngNumero, dbFieldTypeLong))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        descricaoConta = rdResult.GetString("Descrição")
    Else
        descricaoConta = ""
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

Private Sub infoFinanceiras(strEmp As String)
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "Banco, Conta, CondPag"
    cmd.Table.TableName = "Empresas"
    Call cmd.Filter.Append("Apel = @pApel")
    Call cmd.Parameters.add(cmd.CreateParameter("@pApel", strEmp, dbFieldTypeString))
       
    Call cmd.Filter.Append("[Cliente bloqueado] = @pCliente bloqueado")
    Call cmd.Parameters.add(cmd.CreateParameter("@pCliente bloqueado", False, dbFieldTypeBool))
    
    
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        txtTituloBanco.Text = rdResult.GetString("Banco")
        txtTituloConta.Text = rdResult.GetString("Conta")
        txtTituloCondicaoPagamento.Text = rdResult.GetString("CondPag")
    Else
        txtTituloBanco.Text = ""
        txtTituloConta.Text = ""
        txtTituloCondicaoPagamento.Text = ""
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Sub

'***********************************************************************************
'*
'*                      Funções de preenchimentos de campos
'*
'***********************************************************************************
Private Sub preencheComboTipos()
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "Tipo, conhecimento_transp"
    cmd.Table.TableName = "[Tipos Globais]"
    cmd.OrderByClause = "Tipo"
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    cboNotaFiscalTipo.Clear
    cboTipoGlobal.Clear
    cboTipoGlobalTitulo.Clear
    While Not rdResult.EOF
        Call cboNotaFiscalTipo.AddItem(rdResult.GetString("Tipo"))
        'Projeto: 1222 - História: #9972 - Ivo Sousa (09/04/2012)
        If rdResult.GetBoolean("conhecimento_transp") Then
            Call cboTipoGlobal.AddItem(rdResult.GetString("Tipo"))
            Call cboTipoGlobalTitulo.AddItem(rdResult.GetString("Tipo"))
        End If
        rdResult.MoveNext
    Wend
    cboNotaFiscalTipo.ListIndex = 0
    'Projeto: 1222 - História: #9972 - Ivo Sousa (09/04/2012)
    If cboTipoGlobal.ListCount = 0 Then
        Call cboTipoGlobal.AddItem("CTRC")
        Call cboTipoGlobalTitulo.AddItem("CTRC")
    End If
    cboTipoGlobal.ListIndex = 0
    cboTipoGlobalTitulo.ListIndex = 0
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Sub

'***********************************************************************************
'*
'*                      Funções de exibição de registros
'*
'***********************************************************************************
Private Sub mostraNotasFiscais(rd As IDBReader)
    While Not rd.EOF
        With lstNotaFiscalNotas
            .ListItems.add , , " ", , "item"
            .ListItems(.ListItems.Count).SubItems(col_nota) = rd.GetString("Número")
            .ListItems(.ListItems.Count).SubItems(col_tipo) = rd.GetString("Tipo de Registro")
            .ListItems(.ListItems.Count).SubItems(col_situ) = rd.GetString("Situação")
            .ListItems(.ListItems.Count).SubItems(col_emissao) = rd.GetString("Emissão")
            .ListItems(.ListItems.Count).SubItems(COL_APEL) = rd.GetString("Empresa")
            .ListItems(.ListItems.Count).SubItems(col_razao) = rd.GetString("Razão")
            .ListItems(.ListItems.Count).SubItems(col_valor) = Format(rd.GetString("Valor Total"), "##,##0.00")
            .ListItems(.ListItems.Count).SubItems(col_trans) = rd.GetString("nome_transp")
            .ListItems(.ListItems.Count).SubItems(col_codTrans) = rd.GetString("Transportadora")
            .ListItems(.ListItems.Count).SubItems(col_vlrTrans) = Format(rd.GetString("Valor do Frete"), "##,##0.00")
        End With
        rd.MoveNext
    Wend
End Sub

Private Sub preencheGridParcela()
    If Not objConhecimento.Titulo.parcelas Is Nothing Then
       objConhecimento.Titulo.parcelas.MoveFirst
       txtTituloValorTotalParcelas.Text = Format(objConhecimento.Titulo.parcelas.ValorTotal, "###,##0.00")
    Else
        txtTituloValorTotalParcelas.Text = Format("0", "###,##0.00")
    End If
        
    Call CarregaHFlexGrid(grdTituloParcelas, , strTituloGridParcela, , , objConhecimento.Titulo.parcelas)
    grdTituloParcelas.SelectionMode = flexSelectionByRow
End Sub

'***********************************************************************************
'*
'*                      Funções de exibição de conhecimento
'*
'***********************************************************************************
Private Sub mostraCampos()
    Call mostraCamposNotasFiscais
End Sub
Private Sub mostraCamposNotasFiscais()
    With objConhecimento.notasFiscais
        txtNotaFiscalQuantidade.Text = .Count
        txtNotaFiscalValorTotal.Text = Format(.ValorTotal, "##,##0.00##")
    End With
End Sub

'***********************************************************************************
'*
'*                      Funções de limpeza de campos
'*
'***********************************************************************************
Private Sub LimpaCampos(Optional booLimpaNota As Boolean = True)
    Dim dao As New cFretePagarDAO
    If booLimpaNota Then
        Call limpaCamposNotasFiscais
    End If
    Call limpaCamposConhecimento
    Call limpaCamposTitulo
    'pt. 80647 - Ivo Sousa(27/05/2008)
    'txtConhecimentoNumero.Text = dao.lastCodigo
    Set dao = Nothing
    booAlterando = False
    Call bloqueiaCampos
    mlngNrConhecimentoOld = 0
    bGeraDup = True
End Sub

Private Sub limpaCamposNotasFiscais()
    cboNotaFiscalTipo.ListIndex = 0
    cboNotaFiscalEntradaSaida.ListIndex = 0
    txtNotaFiscalDataEmissaoInicial.Text = ""
    txtNotaFiscalDataEmissaoFinal.Text = ""
    txtNotaFiscalNotaInicial.Text = ""
    txtNotaFiscalNotaFinal.Text = ""
    txtNotaFiscalTransportadora.Text = ""
    lblNotaFiscalTransportadora.Caption = ""
    lstNotaFiscalNotas.ListItems.Clear
    txtNotaFiscalQuantidade.Text = "0"
    txtNotaFiscalValorTotal.Text = Format("0", "##,##0.00")
End Sub

Private Sub limpaCamposConhecimento()
    txtConhecimentoTransportadora.Text = ""
    lblConhecimentoTransportadora.Caption = ""
    txtConhecimentoNumero.Text = ""
    txtSerieCTRC.Text = ""
    txtConhecimentoDataEmissao.Text = Date
    txtConhecimentoDataEntrada.Text = Date
    txtConhecimentoSituacao.Text = "Ativo"
    txtConhecimentoRemetente.Text = ""
    lblConhecimentoRemetente.Caption = ""
    txtConhecimentoDestinatario.Text = ""
    lblConhecimentoDestinatario.Caption = ""
    txtConhecimentoConsignatario.Text = ""
    lblConhecimentoConsignatario.Caption = ""
    txtConhecimentoRedespacho.Text = ""
    lblConhecimentoRedespacho.Caption = ""
    txtConhecimentoNaturezaOperacao.Text = ""
    txtConhecimentoNaturezaComplemento.Text = ""
    lblConhecimentoNaturezaOperacao.Caption = ""
    txtConhecimentoDistancia.Text = ""
    txtConhecimentoOperacaoContabil.Clear
    lblConhecimentoOperacaoContabil.Caption = ""
    txtConhecimentoObservacao.Text = ""
    txtConhecimentoVolume.Text = ""
    txtConhecimentoTarifa.Text = ""
    txtConhecimentoPedagio.Text = ""
    txtConhecimentoSeguro.Text = ""
    txtConhecimentoOutros.Text = ""
    txtConhecimentoAcrescimo.Text = ""
    txtConhecimentoDesconto.Text = ""
    txtConhecimentoPorcentagemIcms.Text = ""
    txtConhecimentoValorIsentas.Text = ""
    optConhecimentoCIF.value = True
    txtChaveAcessoEnt.Text = Empty
    chkConhecimentoAdionarFreteMercadoria.value = vbUnchecked
End Sub

Private Sub limpaCamposTitulo()
    txtTituloConhecimento.Text = ""
    txtTituloEmpresa.Text = ""
    lblTituloEmpresa.Caption = ""
    txtTituloEmissao.Text = Date
    txtTituloDescricao.Text = ""
    txtTituloControle.Text = ""
    lblTituloBanco.Caption = ""
    txtTituloBanco.Text = ""
    txtTituloCarteira.Text = ""
    lblTituloConta.Caption = ""
    txtTituloConta.Text = ""
    txtTituloCentroCusto.Text = ""
    lblTituloCentroCusto.Caption = ""
    txtTituloValor.Text = ""
    lblTituloCondicaoPagamento.Caption = ""
    txtTituloCondicaoPagamento.Text = ""
    Call CarregaHFlexGrid(grdTituloParcelas, , strTituloGridParcela, , , Nothing)
    'Limpa Campos das Parcelas
    txtTituloParcelaNumero.Text = ""
    txtTituloParcelaVencimento.Text = ""
    txtTituloParcelaValor.Text = ""
End Sub

Private Function valorComEncargos() As Double
    valorComEncargos = strToDbl(txtConhecimentoValor.Text)
    valorComEncargos = valorComEncargos + strToDbl(txtConhecimentoPedagio.Text)
    valorComEncargos = valorComEncargos + strToDbl(txtConhecimentoSeguro.Text)
    valorComEncargos = valorComEncargos + strToDbl(txtConhecimentoOutros.Text)
    valorComEncargos = valorComEncargos + strToDbl(txtConhecimentoAcrescimo.Text)
    valorComEncargos = valorComEncargos - strToDbl(txtConhecimentoDesconto.Text)
End Function

Private Sub atualizaICMS()
    If Not objUfOrigem Is Nothing And Not objUfDestino Is Nothing Then
        If objUfOrigem.equals(objUfDestino) Then
            txtConhecimentoPorcentagemIcms.Text = Format(objUfOrigem.icmsInterno, "##,##0.00")
            txtConhecimentoValorIcms.Text = ValorICMS
        Else
            txtConhecimentoPorcentagemIcms.Text = Format(objUfOrigem.ICMS, "##,##0.00")
            txtConhecimentoValorIcms.Text = ValorICMS
        End If
    End If
End Sub

Private Function ValorICMS() As Double
    ValorICMS = strToDbl(txtConhecimentoBaseIcms.Text) * (strToDbl(txtConhecimentoPorcentagemIcms.Text) / 100)
End Function

Private Function validaExecutar() As Boolean
    If IsDate(txtNotaFiscalDataEmissaoInicial.Text) Or _
        IsDate(txtNotaFiscalDataEmissaoFinal.Text) Or _
        IsNumeric(txtNotaFiscalNotaInicial.Text) Or _
        IsNumeric(txtNotaFiscalNotaFinal.Text) Or _
        nomeTransportadora(CLng(strToDbl(txtNotaFiscalTransportadora.Text))) <> "" Then
            validaExecutar = True
    Else
        validaExecutar = False
    End If
End Function

Private Function validaCalculoParcelas() As Boolean
    validaCalculoParcelas = False
    If Not IsNumeric(txtTituloValor.Text) Then
        MsgBox "A nota deve conter itens para gerar as parcelas", vbInformation, Me.Caption
    ElseIf txtTituloValor.Text = Format("0", "###,##0.00") Then
        MsgBox "O valor da nota deve ser maior do que ZERO.", vbInformation, Me.Caption
    ElseIf Not IsDate(txtTituloEmissao.Text) Then
        MsgBox "A data da emissão da nota deve ser um data válida.", vbInformation, Me.Caption
        txtTituloEmissao.SetFocus
    ElseIf txtTituloCondicaoPagamento.Text = "" Then
        MsgBox "Para gerar as parcelas é necessário informar a forma de pagamento.", vbInformation, Me.Caption
        txtTituloCondicaoPagamento.SetFocus
    ElseIf objConhecimento.Titulo.CondicaoPagamento Is Nothing Then
        MsgBox "Para gerar as parcelas é necessário informar a forma de pagamento.", vbInformation, Me.Caption
        txtTituloCondicaoPagamento.SetFocus
    ElseIf Not objConhecimento.Titulo.CondicaoPagamento.Existe(CInt(txtTituloCondicaoPagamento.Text)) Then
        MsgBox "Condição de pagamento não cadastrada no sistema.", vbInformation, Me.Caption
        txtTituloCondicaoPagamento.SetFocus
    Else
        validaCalculoParcelas = True
    End If
End Function

Private Function ValidaCampos() As Boolean
    If Not validaCamposNotaFiscal Then
        ValidaCampos = False
    ElseIf Not validaCamposConhecimento Then
        ValidaCampos = False
    ElseIf Not validaCamposTitulo Then
        ValidaCampos = False
    Else
        ValidaCampos = True
    End If
End Function

Private Function validaCamposNotaFiscal() As Boolean
    'Vinicius Elyseu(29/11/2015) - Projeto: #0 - História: #0 - Demanda: #100391
    'validaCamposNotaFiscal = False
    'If strToDbl(txtNotaFiscalQuantidade.Text) = CDbl("0") Then
    '    MsgBox "É nescessário selecionar notas fiscais.", vbInformation, Me.Caption
    '    tabConhecimentoPagar.Tab = 1
    'Else
        validaCamposNotaFiscal = True
    'End If
End Function

Private Function validaCamposConhecimento() As Boolean
    validaCamposConhecimento = False
    If Not IsNumeric(txtConhecimentoNumero.Text) Then
        MsgBox "O campo Conhecimento deve ser preenchido.", vbInformation, Me.Caption
        tabConhecimentoPagar.Tab = 1
        txtConhecimentoNumero.SetFocus
    ElseIf Trim(txtConhecimentoTransportadora.Text) = "" Then
        MsgBox "O campo transportadora deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoTransportadora.SetFocus
    ElseIf Not IsNumeric(txtConhecimentoTransportadora.Text) Then
        MsgBox "O campo Transportadora deve conter um número.", vbInformation, Me.Caption
        txtConhecimentoTransportadora.SetFocus
    ElseIf nomeTransportadora(strToLng(txtConhecimentoTransportadora)) = "" Then
        MsgBox "O campo transportadora deve conter uma transportadora válida.", vbInformation, Me.Caption
        txtConhecimentoTransportadora.SetFocus
    ElseIf Trim(txtConhecimentoDataEmissao.Text) = "" Then
        MsgBox "O campo data de emissão deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoDataEmissao.SetFocus
    ElseIf Not IsDate(txtConhecimentoDataEmissao.Text) Then
        MsgBox "O campo data de emissão deve ser uma data válida.", vbInformation, Me.Caption
        txtConhecimentoDataEmissao.SetFocus
    ElseIf Trim(txtConhecimentoRemetente.Text) = "" Then
        MsgBox "O campo remetente deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoRemetente.SetFocus
    ElseIf dadosEmpresa(txtConhecimentoRemetente.Text) = "" Then
        MsgBox "O campo remetente deve ser uma empresa válida.", vbInformation, Me.Caption
        txtConhecimentoRemetente.SetFocus
    ElseIf Trim(txtConhecimentoDestinatario.Text) = "" Then
        MsgBox "O campo destinatário deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoDestinatario.SetFocus
    ElseIf dadosEmpresa(txtConhecimentoDestinatario.Text) = "" Then
        MsgBox "O campo destinatário deve ser uma empresa válida.", vbInformation, Me.Caption
        txtConhecimentoDestinatario.SetFocus
    ElseIf Trim(txtConhecimentoNaturezaOperacao.Text) = "" Then
        MsgBox "O campo natureza de operação deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoNaturezaOperacao.SetFocus
    ElseIf descricaoNatureza(Trim(txtConhecimentoNaturezaOperacao.Text), Trim(txtConhecimentoNaturezaComplemento.Text)) = "" Then
        MsgBox "O campo natureza de operação deve ser uma natureza válida.", vbInformation, Me.Caption
        txtConhecimentoNaturezaOperacao.SetFocus
    ElseIf Trim(txtConhecimentoVolume.Text) = "" Then
        MsgBox "O campo volume da Nota Fiscal deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoVolume.SetFocus
    ElseIf Not IsNumeric(txtConhecimentoVolume.Text) Then
        MsgBox "O campo volume da Nota Fiscal deve ser um número.", vbInformation, Me.Caption
        txtConhecimentoVolume.SetFocus
    ElseIf strToDbl(txtConhecimentoVolume.Text) = 0 Then
        MsgBox "O campo volume da Nota Fiscal deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoVolume.SetFocus
    ElseIf Trim(txtConhecimentoTarifa.Text) = "" Then
        MsgBox "O campo tarifa deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoTarifa.SetFocus
    ElseIf strToDbl(txtConhecimentoTarifa.Text) = 0 Then
        MsgBox "O campo tarifa deve ser preenchido.", vbInformation, Me.Caption
        txtConhecimentoTarifa.SetFocus
    ElseIf Not ValidaData Then
        txtConhecimentoDataEntrada.SetFocus
    Else
        If txtConhecimentoOperacaoContabil.Enabled Then
            If Len(lblConhecimentoOperacaoContabil.Caption) = 0 Then
                MsgBox "O campo Operação Contábil deve ser preenchido.", vbInformation, Me.Caption
                txtConhecimentoOperacaoContabil.SetFocus
                validaCamposConhecimento = False
            Else
                validaCamposConhecimento = True
            End If
        Else
            validaCamposConhecimento = True
        End If
    End If
End Function

Private Function validaCamposParcela() As Boolean
    validaCamposParcela = False
    If Trim(txtTituloParcelaVencimento.Text) = "" Then
        MsgBox "O campo data de vencimento da parcela deve ser preenchido.", vbInformation, Me.Caption
    ElseIf Not IsDate(txtTituloParcelaVencimento.Text) Then
        MsgBox "O campo data de vencimento da parcela deve ser uma data válida.", vbInformation, Me.Caption
    ElseIf Trim(txtTituloParcelaValor.Text) = "" Then
        MsgBox "O campo valor da parcela deve ser preenchido.", vbInformation, Me.Caption
    ElseIf Not IsNumeric(txtTituloParcelaValor.Text) Then
        MsgBox "O campo valor da parcela deve ser um número.", vbInformation, Me.Caption
    Else
        validaCamposParcela = True
    End If
End Function

Private Function validaCamposTitulo() As Boolean
    If Not bGeraDup Then
        validaCamposTitulo = True
        Exit Function
    End If
    
    validaCamposTitulo = False
    If Trim(txtTituloConhecimento.Text) = "" Then
        MsgBox "O campo numero do conhecimento deve ser preenchido", vbInformation, Me.Caption
    ElseIf Not IsNumeric(txtTituloConhecimento.Text) Then
        MsgBox "O campo número do conhecimento deve ser um número válido.", vbInformation, Me.Caption
    ElseIf Trim(txtTituloEmpresa.Text) = "" Then
        MsgBox "O campo empresa deve ser preenchido.", vbInformation, Me.Caption
        txtTituloEmpresa.SetFocus
    ElseIf dadosEmpresa(txtTituloEmpresa.Text) = "" Then
        MsgBox "O campo empresa deve ser uma empresa válida.", vbInformation, Me.Caption
        txtTituloEmpresa.SetFocus
    ElseIf Trim(txtTituloEmissao.Text) = "" Then
        MsgBox "O campo data de emissão deve ser preenchido.", vbInformation, Me.Caption
        txtTituloEmissao.SetFocus
    ElseIf Not IsDate(txtTituloEmissao.Text) Then
        MsgBox "O campo data de emissão deve ser uma data válida.", vbInformation, Me.Caption
        txtTituloEmissao.SetFocus
    ElseIf Trim(txtTituloCondicaoPagamento.Text) = "" Then
        MsgBox "O campo condição de pagamento deve ser informado.", vbInformation, Me.Caption
        txtTituloCondicaoPagamento.SetFocus
    ElseIf Not IsNumeric(txtTituloValor.Text) Then
        MsgBox "O campo valor do titulo deve ser preenchido.", vbInformation, Me.Caption
    ElseIf objConhecimento.Titulo.parcelas Is Nothing Then
        MsgBox "É obrigatório Calcular a(s) Parcela(s). Natureza de Operação configurada para Gerar Duplicatas..", vbInformation, Me.Caption
        cmdTituloCalcular.SetFocus
    ElseIf objConhecimento.Titulo.CondicaoPagamento.Codigo <> intCondPag Then
        MsgBox "A quantidade de parcelas geradas não coincide com a condição de pagamento.", vbInformation, Me.Caption
    ElseIf Round(objConhecimento.Titulo.parcelas.ValorTotal, 2) <> Round(strToDbl(txtTituloValor), 2) Then
        MsgBox "A soma dos valores das parcelas deve ser o valor do conhecimento.", vbInformation, Me.Caption
    Else
        validaCamposTitulo = True
    End If
End Function

Private Sub preencheClasse()
    Call preencheConhecimentoClasse
    If bGeraDup Then
        Call preencheTituloClasse
    End If
End Sub

Private Sub preencheConhecimentoClasse()
    Dim objFreteDAO As New cFretePagarDAO
    
    With objConhecimento
        .codigoTransportadora = strToLng(txtConhecimentoTransportadora.Text)
        'Projeto: 1222 - História: #9972 - Ivo Sousa (12/04/2012)
        .TipoRegistro = cboTipoGlobal.Text
        'História 15197 - Tarefa 15198 - Ivo Sousa (19/07/2012)
        'If Not booAlterando Then
        '    .numeroConhecimento = objFreteDAO.lastCodigo
        '    txtConhecimentoNumero.Text = .numeroConhecimento
            .numeroConhecimento = strToLng(txtConhecimentoNumero.Text)
            .numeroConhecimentoOld = mlngNrConhecimentoOld
        'Else
            '.numeroConhecimento = strToLng(txtConhecimentoNumero.Text)
        '    If .numeroConhecimento <> strToLng(txtConhecimentoNumero.Text) Then
        '        .numeroConhecimentoOld = strToLng(txtConhecimentoNumero.Text)
        '    End If
        'End If
        .dataEmissao = CDate(txtConhecimentoDataEmissao.Text)
        If IsDate(txtConhecimentoDataEntrada.Text) Then
            .dataEntrada = CDate(txtConhecimentoDataEntrada.Text)
        End If
        .situacao = "A"
        .SerieCTRC = txtSerieCTRC.Text
        .codigoRemetente = txtConhecimentoRemetente.Text
        .codigoDestinatario = txtConhecimentoDestinatario.Text
        .codigoConsignatario = txtConhecimentoConsignatario.Text
        .codigoRedespacho = strToLng(strToDbl(txtConhecimentoRedespacho.Text))
        .codigoCfop = txtConhecimentoNaturezaOperacao.Text
        .codigoCfopVar = txtConhecimentoNaturezaComplemento.Text
        .distancia = strToDbl(txtConhecimentoDistancia.Text)
        .OperacaoContabil = txtConhecimentoOperacaoContabil.valorInteiro
        .Observacao = txtConhecimentoObservacao.Text
        If optConhecimentoCIF.value Then                    'CIF
            .tipoFrete = "C"
        ElseIf optConhecimentoFob.value Then                'FOB
            .tipoFrete = "F"
        'Projeto: 1239 - História: 15187 - Tarefa: Não Planejada - Fernando Paludo 29/06/2012
        ElseIf optConhecimentoEmitente.value Then           'Emitente
            .tipoFrete = "E"
        Else                                                'Terceiro
            .tipoFrete = "T"
        End If
        .volume = strToDbl(txtConhecimentoVolume.Text)
        .porcentagemTarifa = strToDbl(txtConhecimentoTarifa.Text)
        .ValorFrete = strToDbl(txtConhecimentoValor.Text)
        .valorPedagio = strToDbl(txtConhecimentoPedagio.Text)
        .ValorSeguro = strToDbl(txtConhecimentoSeguro.Text)
        .ValorOutros = strToDbl(txtConhecimentoOutros.Text)
        .valorAcrescimo = strToDbl(txtConhecimentoAcrescimo.Text)
        .ValorDesconto = strToDbl(txtConhecimentoDesconto.Text)
        .valorBaseICMS = strToDbl(txtConhecimentoBaseIcms.Text)
        .porcentagemICMS = strToDbl(txtConhecimentoPorcentagemIcms.Text)
        .ValorICMS = strToDbl(txtConhecimentoValorIcms.Text)
        .valorIsentas = strToDbl(txtConhecimentoValorIsentas.Text)
        .valorConhecimento = strToDbl(txtConhecimentoValorTotal.Text)
        .rateiaValorProdutos = (chkConhecimentoAdionarFreteMercadoria.value = vbChecked)
        'pt. 92345 - Ivo Sousa (16/04/2009)
        .CentroCusto = strToLng(txtTituloCentroCusto.Text)
        .ContaFinanceira = strToLng(txtTituloConta.Text)
        .ChaveAcesso = txtChaveAcessoEnt.Text
    End With
End Sub

Private Sub preencheTituloClasse()
    With objConhecimento.Titulo
        .tipo = objConhecimento.TipoRegistro
        .Nota = objConhecimento.numeroConhecimento
        .Empresa = txtTituloEmpresa.Text
        .Emissao = txtTituloEmissao.Text
        .Descricao = txtTituloDescricao.Text
        .Controle = txtTituloControle.Text
        .Banco = strToLng(txtTituloBanco.Text)
        .conta = strToLng(txtTituloConta.Text)
        .CentroCusto = strToLng(strToDbl(txtTituloCentroCusto.Text))
        .OperacaoContabil = txtConhecimentoOperacaoContabil.valorInteiro
        .valor = strToDbl(txtTituloValor.Text)
    End With
End Sub

Public Sub setConhecimento(Frete As cFretePagar)
    Set objConhecimento = Frete
    Call LimpaCampos
    Call mostraCamposClasse
    mlngNrConhecimentoOld = objConhecimento.numeroConhecimento
End Sub

Private Sub mostraCamposClasse()
    Call mostraCamposNotaFiscalClasse
    Call mostraCamposConhecimentoClasse
    Call mostrCamposTituloClasse
    Call bloqueiaCampos
    If Not objConhecimento.notasFiscais.registroEntrada Then
        chkConhecimentoAdionarFreteMercadoria.Enabled = False
    Else
        chkConhecimentoAdionarFreteMercadoria.Enabled = True
    End If
    booAlterando = True
    'Desabilito os campos da nota fiscal na consulta para não permitir alteração
    Label18.Enabled = False
    cboNotaFiscalEntradaSaida.Enabled = False
    Label5.Enabled = False
    cboNotaFiscalTipo.Enabled = False
    Label2.Enabled = False
    txtNotaFiscalNotaInicial.Enabled = False
    Label3.Enabled = False
    txtNotaFiscalNotaFinal.Enabled = False
    cmdNotaFiscalExecutar.Enabled = False
    lblDataEmissao.Enabled = False
    txtNotaFiscalDataEmissaoInicial.Enabled = False
    Label1.Enabled = False
    txtNotaFiscalDataEmissaoFinal.Enabled = False
    Label4.Enabled = False
    txtNotaFiscalTransportadora.Enabled = False
    lstNotaFiscalNotas.Enabled = False
End Sub

Private Sub mostraCamposNotaFiscalClasse()
    With objConhecimento.notasFiscais
        .MoveFirst
        While Not .EOF
            lstNotaFiscalNotas.ListItems.add , , " ", , "selecionado"
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(col_nota) = .CurrentObject.Nota
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(col_tipo) = .CurrentObject.tipo
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(col_situ) = .CurrentObject.situacao
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(col_emissao) = .CurrentObject.Emissao
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(COL_APEL) = .CurrentObject.Apel
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(col_razao) = .CurrentObject.Empresa
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(col_valor) = Format(.CurrentObject.valor, "##,##0.00")
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(col_trans) = .CurrentObject.nomeTransportadora
            lstNotaFiscalNotas.ListItems(lstNotaFiscalNotas.ListItems.Count).SubItems(col_codTrans) = .CurrentObject.Transportadora
            .MoveNext
        Wend
    End With
    Call mostraCamposNotasFiscais
End Sub

Private Sub mostraCamposConhecimentoClasse()
    With objConhecimento
        txtConhecimentoNumero.Text = .numeroConhecimento
        txtSerieCTRC = .SerieCTRC
        cboTipoGlobal.Text = .TipoRegistro
        cboTipoGlobalTitulo.Text = .TipoRegistro
        txtConhecimentoTransportadora.Text = .codigoTransportadora
        txtConhecimentoDataEmissao.Text = .dataEmissao
        If Not IsEmptyDate(.dataEntrada) Then
            txtConhecimentoDataEntrada.Text = .dataEntrada
        End If
        txtConhecimentoSituacao.Text = "A"
        txtConhecimentoRemetente.Text = .codigoRemetente
        txtConhecimentoDestinatario.Text = .codigoDestinatario
        txtConhecimentoConsignatario.Text = .codigoConsignatario
        optConhecimentoCIF.value = .tipoFrete = "C"
        optConhecimentoFob.value = .tipoFrete = "F"
        'Projeto: 1239 - História: 15187 - Tarefa: Não Planejada - Fernando Paludo 29/06/2012
        optConhecimentoEmitente.value = .tipoFrete = "E"
        optConhecimentoTerceiros.value = .tipoFrete = "T"
        txtConhecimentoRedespacho.Text = .codigoRedespacho
        txtConhecimentoNaturezaOperacao.Text = .codigoCfop
        txtConhecimentoNaturezaComplemento.Text = .codigoCfopVar
        Call txtConhecimentoNaturezaOperacao_LostFocus
        txtConhecimentoDistancia.Text = .distancia
        txtConhecimentoOperacaoContabil.valorInteiro = .OperacaoContabil
        txtConhecimentoObservacao.Text = .Observacao
        txtConhecimentoVolume.Text = .volume
        txtConhecimentoTarifa.Text = Format(.porcentagemTarifa, "##,##0.00###")
        txtConhecimentoValor.Text = Format(.ValorFrete, "##,##0.00")
        txtConhecimentoPedagio.Text = Format(.valorPedagio, "##,##0.00")
        txtConhecimentoSeguro.Text = Format(.ValorSeguro, "##,##0.00")
        txtConhecimentoOutros.Text = Format(.ValorOutros, "##,##0.00")
        txtConhecimentoAcrescimo.Text = Format(.valorAcrescimo, "##,##0.00")
        txtConhecimentoDesconto.Text = Format(.ValorDesconto, "##,##0.00")
        txtConhecimentoBaseIcms.Text = Format(.valorBaseICMS, "##,##0.00")
        txtConhecimentoPorcentagemIcms.Text = Format(.porcentagemICMS, "##,##0.00")
        txtConhecimentoValorIcms.Text = Format(.ValorICMS, "##,##0.00")
        txtConhecimentoValorIsentas.Text = Format(.valorIsentas, "##,##0.00")
        If .rateiaValorProdutos Then
            chkConhecimentoAdionarFreteMercadoria.value = vbChecked
        Else
            chkConhecimentoAdionarFreteMercadoria.value = vbUnchecked
        End If
        txtChaveAcessoEnt.Text = Trim(.ChaveAcesso)
    End With
End Sub

Private Sub mostrCamposTituloClasse()
    With objConhecimento.Titulo
        txtTituloEmpresa.Text = .Empresa
        txtTituloEmissao.Text = .Emissao
        txtTituloDescricao.Text = .Descricao
        txtTituloControle.Text = .Controle
        txtTituloBanco.Text = ""
        txtTituloBanco.Text = .Banco
        txtTituloCarteira.Text = ""
        txtTituloCarteira.Text = .Carteira
        txtTituloConta.Text = ""
        txtTituloConta.Text = .conta
        txtTituloCentroCusto.Text = ""
        txtTituloCentroCusto.Text = .CentroCusto
        txtTituloCondicaoPagamento.Text = .CondicaoPagamento.Codigo
        intCondPag = .CondicaoPagamento.Codigo
        txtTituloValor.Text = Format(.valor, "##,##0.00")
        Call preencheGridParcela
    End With
End Sub

Private Sub bloqueiaCampos()
    Call bloqueiaCamposNotaFiscal
    Call bloqueiaCamposConhecimento
    Call bloqueiaCamposTitulo
End Sub

Private Sub bloqueiaCamposNotaFiscal()
    With objConhecimento
        Label18.Enabled = .PermiteAlteracao
        cboNotaFiscalEntradaSaida.Enabled = .PermiteAlteracao
        Label5.Enabled = .PermiteAlteracao
        cboNotaFiscalTipo.Enabled = .PermiteAlteracao
        Label2.Enabled = .PermiteAlteracao
        txtNotaFiscalNotaInicial.Enabled = .PermiteAlteracao
        Label3.Enabled = .PermiteAlteracao
        txtNotaFiscalNotaFinal.Enabled = .PermiteAlteracao
        lblDataEmissao.Enabled = .PermiteAlteracao
        txtNotaFiscalDataEmissaoInicial.Enabled = .PermiteAlteracao
        Label1.Enabled = .PermiteAlteracao
        txtNotaFiscalDataEmissaoFinal.Enabled = .PermiteAlteracao
        Label4.Enabled = .PermiteAlteracao
        txtNotaFiscalTransportadora.Enabled = .PermiteAlteracao
        cmdNotaFiscalExecutar.Enabled = .PermiteAlteracao
        lstNotaFiscalNotas.Enabled = .PermiteAlteracao
    End With
End Sub

Private Sub bloqueiaCamposConhecimento()
    With objConhecimento
        Label8.Enabled = .PermiteAlteracao
        txtConhecimentoTransportadora.Enabled = .PermiteAlteracao
        Label11.Enabled = .PermiteAlteracao
        txtConhecimentoDataEmissao.Enabled = .PermiteAlteracao
        Label13.Enabled = .PermiteAlteracao
        txtConhecimentoRemetente.Enabled = .PermiteAlteracao
        Label14.Enabled = .PermiteAlteracao
        txtConhecimentoDestinatario.Enabled = .PermiteAlteracao
        Label15.Enabled = .PermiteAlteracao
        txtConhecimentoConsignatario.Enabled = .PermiteAlteracao
        Label16.Enabled = .PermiteAlteracao
        txtConhecimentoRedespacho.Enabled = .PermiteAlteracao
        optConhecimentoCIF.Enabled = .PermiteAlteracao
        optConhecimentoFob.Enabled = .PermiteAlteracao
        'Projeto: 1239 - História: 15187 - Tarefa: Não Planejada - Fernando Paludo 29/06/2012
        optConhecimentoEmitente.Enabled = .PermiteAlteracao
        optConhecimentoTerceiros.Enabled = .PermiteAlteracao
        Label17.Enabled = .PermiteAlteracao
        txtConhecimentoNaturezaOperacao.Enabled = .PermiteAlteracao
        txtConhecimentoNaturezaComplemento.Enabled = .PermiteAlteracao
        Label19.Enabled = .PermiteAlteracao
        txtConhecimentoDistancia.Enabled = .PermiteAlteracao
        Label20.Enabled = .PermiteAlteracao
        txtConhecimentoObservacao.Enabled = .PermiteAlteracao
        Label21.Enabled = .PermiteAlteracao
        txtConhecimentoVolume.Enabled = .PermiteAlteracao
        Label22.Enabled = .PermiteAlteracao
        txtConhecimentoTarifa.Enabled = .PermiteAlteracao
        Label23.Enabled = .PermiteAlteracao
        txtConhecimentoValor.Enabled = .PermiteAlteracao
        Label24.Enabled = .PermiteAlteracao
        txtConhecimentoPedagio.Enabled = .PermiteAlteracao
        Label25.Enabled = .PermiteAlteracao
        txtConhecimentoSeguro.Enabled = .PermiteAlteracao
        Label26.Enabled = .PermiteAlteracao
        txtConhecimentoOutros.Enabled = .PermiteAlteracao
        Label27.Enabled = .PermiteAlteracao
        txtConhecimentoAcrescimo.Enabled = .PermiteAlteracao
        Label28.Enabled = .PermiteAlteracao
        txtConhecimentoDesconto.Enabled = .PermiteAlteracao
        Label30.Enabled = .PermiteAlteracao
        txtConhecimentoBaseIcms.Enabled = .PermiteAlteracao
        Label31.Enabled = .PermiteAlteracao
        txtConhecimentoPorcentagemIcms.Enabled = .PermiteAlteracao
        Label33.Enabled = .PermiteAlteracao
        txtConhecimentoValorIsentas.Enabled = .PermiteAlteracao
        chkConhecimentoAdionarFreteMercadoria.Enabled = .PermiteAlteracao
    End With
End Sub

Private Sub bloqueiaCamposTitulo()
    With objConhecimento
        'pt. 86678 - Ivo Sousa(26/05/2008)
        If txtTituloEmpresa.Text <> "" Then
            labEmpresa.Enabled = .PermiteAlteracao
            txtTituloEmpresa.Enabled = .PermiteAlteracao
            labEmissao.Enabled = .PermiteAlteracao
            txtTituloEmissao.Enabled = .PermiteAlteracao
            Label41.Enabled = .PermiteAlteracao
            txtTituloBanco.Enabled = .PermiteAlteracao
            labCondPagto.Enabled = .PermiteAlteracao
            txtTituloCondicaoPagamento.Enabled = .PermiteAlteracao
            cmdTituloCalcular.Enabled = .PermiteAlteracao
            Label49.Enabled = False
            txtTituloParcelaVencimento.Enabled = False
            Label50.Enabled = False
            txtTituloParcelaValor.Enabled = False
            cmdTituloParcelaConfirmar.Enabled = False
            cmdTituloParcelaCancelar.Enabled = False
        Else
            txtTituloValor.Text = txtConhecimentoValor.Text
            txtTituloEmissao.Text = Date
        End If
    End With
End Sub

'Data......: 20/12/2006
'Autor.....: Dulcino Júnior
'Descrição.: Função criada para gerar o registro de integração com o cordilheira do conhecimento de frete.
'Retorno...: [Boolean] Retorna se foi possivel integrar o conhecimento de frete.
Private Function integraRegistro() As Boolean
    Dim facDAO As New cDAOFactory
    Dim blnRet As Boolean
    
On Error GoTo erro_integrando
    If Configuracao("Mostrar integrações contábeis com terceiros", "False") Then
        AppIntegra.Connect
        If facDAO.criarNotaEntradaDAO(AppIntegra).Existe(objConhecimento.TipoRegistro, objConhecimento.numeroConhecimento, objConhecimento.Titulo.Empresa, tp_conhecimento, objConhecimento.codigoTransportadora) Then
            blnRet = facDAO.criarNotaEntradaDAO(AppIntegra).Excluir(objConhecimento.TipoRegistro, objConhecimento.numeroConhecimento, objConhecimento.Titulo.Empresa, CInt(objConhecimento.codigoCfop), tp_conhecimento, objConhecimento.codigoTransportadora)
        Else
            blnRet = True
        End If
        If blnRet Then
            blnRet = blnRet And facDAO.criarNotaEntradaDAO(AppIntegra).Gravar(objConhecimento.Exportar)
        End If
        AppIntegra.Disconnect
    Else
        blnRet = True
    End If
    integraRegistro = blnRet
    Exit Function
erro_integrando:
    Call Throw(err)
End Function

Private Function excluiIntegracao() As Boolean
    Dim facDAO As New cDAOFactory
    Dim blnRet As Boolean
    
On Error GoTo erro_excluindo
    If Configuracao("Mostrar integrações contábeis com terceiros", "False") Then
        AppIntegra.Connect
        With objConhecimento
            If facDAO.criarNotaEntradaDAO(AppIntegra).Existe(.TipoRegistro, .numeroConhecimento, .Titulo.Empresa, tp_conhecimento, .codigoTransportadora) Then
                blnRet = facDAO.criarNotaEntradaDAO(AppIntegra).Excluir(.TipoRegistro, .numeroConhecimento, .Titulo.Empresa, .codigoCfop, tp_conhecimento, .codigoTransportadora)
            Else
                blnRet = True
            End If
        End With
        AppIntegra.Disconnect
    Else
        blnRet = True
    End If
    excluiIntegracao = blnRet
    Exit Function
erro_excluindo:
    Call Throw(err)
    excluiIntegracao = False
End Function

'Data.......: 07/03/2007
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para retornar a descrição da operação contábil.
'Parametros.: [Long] Código da operação contábil a ser retornada a descrição.
'Retorno....: [String] Descrição da Operação Contábil que possui o código
'               informado.
Private Function descricaoOperacao(lngCodigo As Long) As String
    Dim selCmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
On Error GoTo error_handler
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "descricao"
        
        .Table.TableName = "OperacaoContabil"
        
        Call .Filter.Append("cd_operacao = @pCodigoOperacao")
        Call .Parameters.add(.CreateParameter("@pCodigoOperacao", lngCodigo, dbFieldTypeLong))
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    If Not rdResult.EOF Then
        descricaoOperacao = rdResult.GetString("descricao")
    End If
    Set selCmd = Nothing
    rdResult.CloseReader
    Set rdResult = Nothing
    Aplicacao.Disconnect
    
    Exit Function
error_handler:
    FinallyConnection Aplicacao
    descricaoOperacao = ""
End Function

'Data.......: 25/07/2007
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para verificar se o icms deve ser calculado
'               ou não de acordo com a Natureza de Operação.
'Retorno....: [Boolean] Retorna se o icms deve ser tributado ou não.
Private Function IsIcmsTributado() As Boolean
    Dim selCmd   As IDBSelectCommand
    Dim rdResult As IDBReader
    
On Error GoTo error_handler
    If lblConhecimentoNaturezaOperacao.Caption <> "" Then
        Aplicacao.Connect
        Set selCmd = Aplicacao.CreateSelectCommand
        With selCmd
            .SelectClause = "ICMS"
            
            .Table.TableName = "[Naturezas de Operação]"
            Call .Filter.Append("Código = @pCodigo")
            Call .Parameters.add(.CreateParameter("@pCodigo", txtConhecimentoNaturezaOperacao.Text, dbFieldTypeString))
        
            Call .Filter.Append("Complemento = @pComplemento")
            Call .Parameters.add(.CreateParameter("@pComplemento", txtConhecimentoNaturezaComplemento.Text, dbFieldTypeString))
        End With
        Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
        If Not rdResult.EOF Then
            IsIcmsTributado = (rdResult.GetString("ICMS") = "Tributado")
        Else
            IsIcmsTributado = True
        End If
        rdResult.CloseReader
        Set rdResult = Nothing
        Set selCmd = Nothing
        Aplicacao.Disconnect
    Else
        IsIcmsTributado = True
    End If
    Exit Function

error_handler:
    FinallyConnection Aplicacao
    err.Clear
    IsIcmsTributado = True
End Function

'Data.......: 25/07/2007
'Autor......: Dulcino Júnior
'Descrição..: Função que retorna o valor da base de Icms do conhecimento
'               caso o mesmo não seja isento de acordo com a tributação
'               da natureza.
'Retorno....: [Currency] Valor da Base de calculo de Icms do conhecimento.
Private Function valorBaseICMS() As Currency
    If IsIcmsTributado Then
        'Vinicius Elyseu(04/09/2015) - Projeto: #0 - História: #0 - Demanda: 90388
        valorBaseICMS = strToDbl(txtConhecimentoEncargosValor.Text)
        txtConhecimentoValorIsentas.Text = "0"
        Call atualizaICMS
    Else
        valorBaseICMS = "0"
        txtConhecimentoPorcentagemIcms.Text = "0"
        txtConhecimentoValorIsentas.Text = strToLng(txtConhecimentoEncargosValor.Text)
    End If
End Function

'Data.......: 26/05/2008
'Autor......: Ivo Sousa
'Descrição..: Função para validar uma data qualquer para o calendário
'Retorno....: [Boolean]Se a data é valida
Private Function ValidaData() As Boolean
    If txtConhecimentoDataEntrada.Text <> "" Then
        If ValidaDatasDiasUteis(0, , , , , , txtConhecimentoDataEntrada.Text) Then
            ValidaData = True
        End If
    Else
        MsgBox "A data de entrada é de preenchimento obrigatório.", vbOKOnly + vbInformation, NomeModulo
        ValidaData = False
    End If
End Function

Private Sub txtTituloParcelaVencimento_LostFocus()
    'pt. 87834 - Moacir Pfau(15/07/2008)
    If txtTituloParcelaVencimento.Text <> "" Then
        If Not EData(txtTituloParcelaVencimento.Text) Then
          MsgBox "Data informada inválida."
          txtTituloParcelaVencimento.SetFocus
          Exit Sub
        End If
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

Private Function ValidaChaveAcesso() As Boolean
    Dim intCont    As Integer
    Dim intContAux As Integer
    Dim lngSum     As Long
    Dim strChave   As String
    Dim intDV      As Integer
    Dim dblResto   As Double
       
    strChave = txtChaveAcessoEnt.Text
    If UCase(Left(strChave, 3)) = "NFE" Then
        txtChaveAcessoEnt.Text = Right(strChave, Len(strChave) - 3)
        strChave = txtChaveAcessoEnt.Text
    End If
    If Not Len(Trim(strChave)) = 44 Then
        Exit Function
    Else
        If Not IsNumeric(Right(strChave, 1)) Then
            Exit Function
        End If
        intDV = Right(strChave, 1)
        intContAux = 2
        For intCont = 43 To 1 Step -1
            lngSum = lngSum + Mid(strChave, intCont, 1) * intContAux
            intContAux = intContAux + 1
            If intContAux > 9 Then
                intContAux = 2
            End If
        Next
        dblResto = (lngSum Mod 11)
        If (dblResto = 0) Or (dblResto = 1) Then
            ValidaChaveAcesso = (intDV = 0)
        Else
            ValidaChaveAcesso = ((11 - dblResto) = intDV)
        End If
    End If
End Function

