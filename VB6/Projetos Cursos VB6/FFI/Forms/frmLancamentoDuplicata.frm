VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmLancamentoDuplicata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento"
   ClientHeight    =   9765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9765
   ScaleWidth      =   12825
   Begin VB.Frame Frame 
      Height          =   9315
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   11385
      Begin TabDlg.SSTab SSTab 
         Height          =   9105
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   16060
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Dados Gerais"
         TabPicture(0)   =   "frmLancamentoDuplicata.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraLancamentos(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraValores"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "fraSomaValores"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fraControles"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "fraDatas"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "fraBaixas"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "fraUsuario"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Frame(2)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Adicionais"
         TabPicture(1)   =   "frmLancamentoDuplicata.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraLinhaDigitavel"
         Tab(1).Control(1)=   "EBSMemo"
         Tab(1).Control(2)=   "fraMultaJuroDesconto"
         Tab(1).Control(3)=   "fraDadosAdicionais"
         Tab(1).Control(4)=   "fraInformacaoCheque(2)"
         Tab(1).Control(5)=   "fraLancamentos(0)"
         Tab(1).Control(6)=   "Frame1"
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Outros"
         TabPicture(2)   =   "frmLancamentoDuplicata.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraOrigemCheque"
         Tab(2).Control(1)=   "fraDadosBancarios"
         Tab(2).Control(2)=   "fraEnderecoCobranca"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Log"
         TabPicture(3)   =   "frmLancamentoDuplicata.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "grdLog"
         Tab(3).ControlCount=   1
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLog 
            Height          =   8625
            Left            =   -74910
            TabIndex        =   144
            Top             =   390
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   15214
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame Frame 
            Caption         =   "Baixas Parciais"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1875
            Index           =   2
            Left            =   90
            TabIndex        =   140
            Top             =   7200
            Width           =   11055
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBaixasParc 
               Height          =   1605
               Left            =   90
               TabIndex        =   141
               Top             =   210
               Width           =   10905
               _ExtentX        =   19235
               _ExtentY        =   2831
               _Version        =   393216
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Remessa Bancária"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   -74910
            TabIndex        =   137
            Top             =   5940
            Width           =   6255
            Begin Fox.EBSText etxStatusRemessa 
               Height          =   330
               Left            =   705
               TabIndex        =   138
               Tag             =   "Status"
               Top             =   330
               Width           =   1710
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   60
               Enabled         =   0   'False
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
               Caption         =   "Status:"
               Height          =   195
               Left            =   150
               TabIndex        =   139
               Top             =   390
               Width           =   495
            End
         End
         Begin VB.Frame fraEnderecoCobranca 
            Caption         =   "Endereço de Cobrança"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1725
            Left            =   -74910
            TabIndex        =   89
            Top             =   4380
            Width           =   11055
            Begin Fox.EBSText etxCodigoOutros 
               Height          =   330
               Left            =   1305
               TabIndex        =   91
               Tag             =   "Código - Endereço Cobrança"
               Top             =   390
               Width           =   660
               _ExtentX        =   265
               _ExtentY        =   582
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
            Begin Fox.EBSText etxEnderecoOutros 
               Height          =   330
               Left            =   1305
               TabIndex        =   105
               Top             =   810
               Width           =   9615
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   15
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxBairroOutros 
               Height          =   330
               Left            =   1305
               TabIndex        =   107
               Top             =   1230
               Width           =   4275
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   15
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxCidadeOutros 
               Height          =   330
               Left            =   5280
               TabIndex        =   101
               Top             =   390
               Width           =   2790
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   15
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxUFOutros 
               Height          =   330
               Left            =   8790
               TabIndex        =   103
               Top             =   390
               Width           =   480
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   15
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxCEPOutros 
               Height          =   330
               Left            =   3000
               TabIndex        =   99
               Top             =   390
               Width           =   1200
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               TipoTexto       =   0
               MaxLength       =   9
               Enabled         =   0   'False
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
               ExibeDescricao  =   0   'False
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CEP:"
               Height          =   195
               Left            =   2550
               TabIndex        =   131
               Top             =   450
               Width           =   360
            End
            Begin VB.Label lblUFOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF:"
               Height          =   195
               Left            =   8475
               TabIndex        =   102
               Top             =   450
               Width           =   255
            End
            Begin VB.Label lblCidadeOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cidade:"
               Height          =   195
               Left            =   4680
               TabIndex        =   100
               Top             =   450
               Width           =   540
            End
            Begin VB.Label lblCodigoOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Código:"
               Height          =   195
               Left            =   705
               TabIndex        =   90
               Top             =   450
               Width           =   540
            End
            Begin VB.Label lblEnderecoOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Height          =   195
               Left            =   510
               TabIndex        =   104
               Top             =   870
               Width           =   735
            End
            Begin VB.Label lblBairroOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro:"
               Height          =   195
               Left            =   795
               TabIndex        =   106
               Top             =   1290
               Width           =   450
            End
         End
         Begin VB.Frame fraDadosBancarios 
            Caption         =   "Dados Bancários"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1725
            Left            =   -74910
            TabIndex        =   72
            Top             =   2640
            Width           =   11055
            Begin Fox.EBSText etxLinhaDigitavelOutros 
               Height          =   330
               Left            =   1305
               TabIndex        =   76
               Tag             =   "Linha Digitável - Outros"
               Top             =   390
               Width           =   9630
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   80
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxNossoNumero 
               Height          =   330
               Left            =   1290
               TabIndex        =   78
               Tag             =   "Nosso Número"
               Top             =   810
               Width           =   5040
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   100
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxCarteira 
               Height          =   330
               Left            =   1305
               TabIndex        =   80
               Tag             =   "Carteira"
               Top             =   1230
               Width           =   8235
               _ExtentX        =   445823
               _ExtentY        =   582
               MaxLength       =   15
               Enabled         =   0   'False
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "id_carteira"
               TipoCriterio    =   4
               CampoDescricao  =   "desc_carteira"
               TabelaConsulta  =   "FFICarteira"
               TamanhoDescricao=   7000
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
            Begin VB.Label lblCarteira 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Carteira:"
               Height          =   195
               Left            =   660
               TabIndex        =   79
               Top             =   1290
               Width           =   585
            End
            Begin VB.Label lblNossoNumero 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Nosso Número:"
               Height          =   195
               Left            =   150
               TabIndex        =   77
               Top             =   870
               Width           =   1095
            End
            Begin VB.Label lblLinhaDigitavelOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Linha Digitável:"
               Height          =   195
               Left            =   150
               TabIndex        =   74
               Top             =   450
               Width           =   1095
            End
         End
         Begin VB.Frame fraOrigemCheque 
            Caption         =   "Origem do Cheque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   -74910
            TabIndex        =   4
            Top             =   480
            Width           =   11055
            Begin Fox.EBSText etxBancoOutros 
               Height          =   330
               Left            =   1290
               TabIndex        =   69
               Tag             =   "Banco - Outros"
               Top             =   360
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   582
               MaxLength       =   15
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
            End
            Begin Fox.EBSText etxAgenciaOutros 
               Height          =   330
               Left            =   1290
               TabIndex        =   71
               Tag             =   "Agência - Outros"
               Top             =   780
               Width           =   1200
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   10
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
            Begin Fox.EBSText etxContaCorrenteOutros 
               Height          =   330
               Left            =   1290
               TabIndex        =   73
               Tag             =   "Conta Corrente - Outros"
               Top             =   1200
               Width           =   2040
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   20
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
            Begin Fox.EBSText etxCorrentistaOutros 
               Height          =   330
               Left            =   1290
               TabIndex        =   75
               Tag             =   "Correntista - Outros"
               Top             =   1620
               Width           =   5010
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   60
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
            Begin VB.Label lblCorrentistaOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Correntista:"
               Height          =   195
               Left            =   435
               TabIndex        =   8
               Top             =   1680
               Width           =   795
            End
            Begin VB.Label lblContaCorrenteOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Conta Corrente:"
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   1260
               Width           =   1110
            End
            Begin VB.Label lblAgenciaOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Agência:"
               Height          =   195
               Left            =   600
               TabIndex        =   6
               Top             =   840
               Width           =   630
            End
            Begin VB.Label lblBancoOutros 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Banco:"
               Height          =   195
               Left            =   720
               TabIndex        =   5
               Top             =   420
               Width           =   510
            End
         End
         Begin VB.Frame fraLancamentos 
            Caption         =   "Lançamentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2625
            Index           =   0
            Left            =   -68580
            TabIndex        =   83
            Top             =   3270
            Width           =   4725
            Begin ComctlLib.ListView lvwLancamentos 
               Height          =   2205
               Left            =   60
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   300
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   3889
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               _Version        =   327682
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
         End
         Begin VB.Frame fraInformacaoCheque 
            Caption         =   "Informação do Cheque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2805
            Index           =   2
            Left            =   -68580
            TabIndex        =   10
            Top             =   450
            Width           =   4725
            Begin VB.TextBox etxHistorico 
               Height          =   1275
               Left            =   900
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   67
               Tag             =   "Histórico"
               Top             =   870
               Width           =   3735
            End
            Begin VB.CommandButton cmdNominal 
               Caption         =   "..."
               Height          =   375
               Left            =   4350
               TabIndex        =   12
               Top             =   420
               Width           =   255
            End
            Begin Fox.EBSText etxNominal 
               Height          =   330
               Left            =   900
               TabIndex        =   65
               Tag             =   "Nominal"
               Top             =   420
               Width           =   3390
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   60
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
            Begin Fox.EBSText etxTotalInfCheque 
               Height          =   330
               Left            =   885
               TabIndex        =   15
               Tag             =   "Total - Cheque"
               Top             =   2280
               Width           =   3690
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   15
               Enabled         =   0   'False
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
            Begin VB.Label lblTotalInfCheque 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Total:"
               Height          =   195
               Left            =   420
               TabIndex        =   14
               Top             =   2340
               Width           =   405
            End
            Begin ComctlLib.ImageList imgDupl 
               Left            =   210
               Top             =   1200
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               MaskColor       =   12632256
               _Version        =   327682
            End
            Begin VB.Label lblHistorico 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Histórico:"
               Height          =   195
               Left            =   195
               TabIndex        =   13
               Top             =   870
               Width           =   660
            End
            Begin VB.Label lblNominal 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Nominal:"
               Height          =   195
               Left            =   225
               TabIndex        =   11
               Top             =   480
               Width           =   615
            End
         End
         Begin VB.Frame fraDadosAdicionais 
            Caption         =   "Dados Adicionais"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1425
            Left            =   -74910
            TabIndex        =   108
            Top             =   4470
            Width           =   6255
            Begin Fox.EBSText etxCidadeAdicional 
               Height          =   330
               Left            =   720
               TabIndex        =   129
               Tag             =   "Cidade"
               Top             =   330
               Width           =   5370
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   20
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxEstadoAdicional 
               Height          =   330
               Left            =   720
               TabIndex        =   130
               Tag             =   "Estado"
               Top             =   720
               Width           =   630
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   20
               Enabled         =   0   'False
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
            Begin VB.Label lblEstado 
               Caption         =   "Estado:"
               Height          =   285
               Left            =   120
               TabIndex        =   110
               Top             =   780
               Width           =   615
            End
            Begin VB.Label lblCidade 
               Caption         =   "Cidade:"
               Height          =   255
               Left            =   120
               TabIndex        =   109
               Top             =   390
               Width           =   615
            End
         End
         Begin VB.Frame fraMultaJuroDesconto 
            Caption         =   "Multa, Juro e Desconto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1785
            Left            =   -74910
            TabIndex        =   55
            Top             =   2640
            Width           =   6255
            Begin Fox.EBSText etxPercMulta 
               Height          =   330
               Left            =   1905
               TabIndex        =   58
               Tag             =   "Perc. Multa"
               Top             =   360
               Width           =   1320
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   15
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
            Begin Fox.EBSText etxPercMora 
               Height          =   330
               Left            =   1905
               TabIndex        =   59
               Tag             =   "Perc. Mora"
               Top             =   810
               Width           =   1320
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   15
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
            Begin Fox.EBSText etxVlrDescPontualidade 
               Height          =   330
               Left            =   1905
               TabIndex        =   60
               Tag             =   "Vlr. Desc. Pontualidade"
               Top             =   1260
               Width           =   1320
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   15
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
            Begin Fox.EBSText etxVlrMulta 
               Height          =   330
               Left            =   4725
               TabIndex        =   61
               Tag             =   "Vlr. Multa"
               Top             =   360
               Width           =   1320
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   15
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
            Begin Fox.EBSText etxVlrMoraDiaria 
               Height          =   330
               Left            =   4725
               TabIndex        =   63
               Tag             =   "Vlr Mora Diária"
               Top             =   780
               Width           =   1320
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   15
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
            Begin VB.Label lblVlrMoraDiaria 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vlr Mora Diária:"
               Height          =   195
               Index           =   1
               Left            =   3585
               TabIndex        =   64
               Top             =   840
               Width           =   1080
            End
            Begin VB.Label lblVlrMulta 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vlr. Multa:"
               Height          =   195
               Index           =   0
               Left            =   3960
               TabIndex        =   62
               Top             =   420
               Width           =   705
            End
            Begin VB.Label lblVlrDescPontualidade 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vlr. Desc. Pontualidade:"
               Height          =   195
               Left            =   135
               TabIndex        =   70
               Top             =   1320
               Width           =   1710
            End
            Begin VB.Label lblPercMora 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Perc. Mora:"
               Height          =   195
               Left            =   1020
               TabIndex        =   66
               Top             =   870
               Width           =   825
            End
            Begin VB.Label lblPercMulta 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Perc. Multa:"
               Height          =   195
               Index           =   0
               Left            =   990
               TabIndex        =   56
               Top             =   420
               Width           =   855
            End
         End
         Begin VB.Frame EBSMemo 
            Caption         =   "Observação"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1365
            Left            =   -74910
            TabIndex        =   9
            Top             =   1230
            Width           =   6255
            Begin VB.TextBox etxObservacao 
               Height          =   975
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   57
               Tag             =   "Observação"
               Top             =   240
               Width           =   6015
            End
         End
         Begin VB.Frame fraLinhaDigitavel 
            Caption         =   "Linha Digitável"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   -74910
            TabIndex        =   2
            Top             =   450
            Width           =   6255
            Begin Fox.EBSText etxLinhaDigitavel 
               Height          =   330
               Left            =   120
               TabIndex        =   3
               Tag             =   "Linha Digitável"
               Top             =   240
               Width           =   6015
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   80
               Enabled         =   0   'False
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
         End
         Begin VB.Frame fraUsuario 
            Height          =   795
            Left            =   7530
            TabIndex        =   124
            Top             =   6420
            Width           =   3615
            Begin Fox.EBSText etxUsuario 
               Height          =   330
               Left            =   840
               TabIndex        =   126
               Top             =   240
               Width           =   1215
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   80
               Enabled         =   0   'False
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
            Begin Fox.EBSData etxAlteracao 
               Height          =   330
               Left            =   2130
               TabIndex        =   127
               Top             =   240
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   582
               Enabled         =   0   'False
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
            Begin VB.Label lblUsuario 
               AutoSize        =   -1  'True
               Caption         =   "Usuário:"
               Height          =   195
               Left            =   180
               TabIndex        =   125
               Top             =   300
               Width           =   585
            End
         End
         Begin VB.Frame fraBaixas 
            Caption         =   "Baixas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7530
            TabIndex        =   122
            Top             =   5580
            Width           =   3615
            Begin Fox.EBSText etxOpContabilBaixa 
               Height          =   330
               Left            =   1350
               TabIndex        =   54
               Tag             =   "Op. Contábil - Baixas"
               Top             =   330
               Width           =   2220
               _ExtentX        =   436113
               _ExtentY        =   582
               MaxLength       =   10
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "cd_operacao"
               TipoCriterio    =   4
               CampoDescricao  =   "descricao"
               TabelaConsulta  =   "OperacaoContabil"
               TamanhoDescricao=   1500
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
            Begin VB.Label lblOpContabilBaixa 
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
               Left            =   210
               TabIndex        =   123
               Top             =   390
               Width           =   1125
            End
         End
         Begin VB.Frame fraDatas 
            Caption         =   "Datas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   7530
            TabIndex        =   84
            Top             =   3510
            Width           =   3615
            Begin Fox.EBSData etxEmissao 
               Height          =   330
               Left            =   1320
               TabIndex        =   50
               Tag             =   "Emissão"
               Top             =   285
               Width           =   1305
               _ExtentX        =   2302
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
            Begin Fox.EBSData etxVencimento 
               Height          =   330
               Left            =   1320
               TabIndex        =   51
               Tag             =   "Vencimento"
               Top             =   705
               Width           =   1305
               _ExtentX        =   2355
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
            Begin Fox.EBSData etxPagamento 
               Height          =   330
               Left            =   1320
               TabIndex        =   52
               Tag             =   "Pagamento"
               Top             =   1132
               Width           =   1305
               _ExtentX        =   3784
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
            Begin Fox.EBSData etxLiberacao 
               Height          =   330
               Left            =   1320
               TabIndex        =   53
               Tag             =   "Liberação"
               Top             =   1552
               Width           =   1305
               _ExtentX        =   3784
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
            Begin VB.Label lblLiberacaoD 
               Height          =   225
               Left            =   2640
               TabIndex        =   136
               Top             =   1620
               Width           =   915
            End
            Begin VB.Label lblPagamentoD 
               Height          =   225
               Left            =   2640
               TabIndex        =   135
               Top             =   1200
               Width           =   915
            End
            Begin VB.Label lblVencimentoD 
               Height          =   225
               Left            =   2640
               TabIndex        =   134
               Top             =   780
               Width           =   915
            End
            Begin VB.Label lblEmissaoD 
               Height          =   225
               Left            =   2640
               TabIndex        =   133
               Top             =   390
               Width           =   915
            End
            Begin VB.Label lblEmissao 
               Alignment       =   1  'Right Justify
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
               Left            =   510
               TabIndex        =   85
               Top             =   360
               Width           =   765
            End
            Begin VB.Label lblVencimento 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vencimento:"
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
               Left            =   210
               TabIndex        =   86
               Top             =   780
               Width           =   1065
            End
            Begin VB.Label lblPagamento 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Pagamento:"
               Height          =   195
               Left            =   420
               TabIndex        =   87
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label lblLiberacao 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Liberação:"
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
               Left            =   360
               TabIndex        =   88
               Top             =   1620
               Width           =   915
            End
         End
         Begin VB.Frame fraControles 
            Caption         =   "Controles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4575
            Left            =   90
            TabIndex        =   111
            Top             =   2640
            Width           =   7365
            Begin VB.CheckBox ChkConciliado 
               Caption         =   "Conciliado"
               Height          =   195
               Left            =   1200
               TabIndex        =   121
               Tag             =   "Conciliado"
               Top             =   3720
               Width           =   1185
            End
            Begin VB.CommandButton cmdRateio 
               Caption         =   "&Rateio..."
               Height          =   375
               Left            =   3960
               TabIndex        =   39
               Top             =   3240
               Width           =   855
            End
            Begin VB.CommandButton cmdProxCheque 
               Caption         =   "..."
               Height          =   345
               Left            =   2730
               TabIndex        =   119
               Top             =   2820
               Width           =   255
            End
            Begin Fox.EBSText etxFormaPagto 
               Height          =   330
               Left            =   1215
               TabIndex        =   31
               Tag             =   "Forma Pagto"
               Top             =   300
               Width           =   5895
               _ExtentX        =   441748
               _ExtentY        =   582
               MaxLength       =   2
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "Código"
               TipoCriterio    =   4
               CampoDescricao  =   "Nome"
               TabelaConsulta  =   "Formas de Pagamento"
               TamanhoDescricao=   4700
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
            Begin Fox.EBSText etxBanco 
               Height          =   330
               Left            =   1215
               TabIndex        =   32
               Tag             =   "Banco"
               Top             =   720
               Width           =   5895
               _ExtentX        =   441748
               _ExtentY        =   582
               MaxLength       =   9
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "Banco"
               TipoCriterio    =   4
               CampoDescricao  =   "Nome"
               TabelaConsulta  =   "Bancos"
               TamanhoDescricao=   4700
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
            Begin Fox.EBSText etxConta 
               Height          =   330
               Left            =   1215
               TabIndex        =   33
               Tag             =   "Conta"
               Top             =   1140
               Width           =   5895
               _ExtentX        =   441748
               _ExtentY        =   582
               MaxLength       =   9
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "Código"
               TipoCriterio    =   4
               CampoDescricao  =   "Descrição"
               TabelaConsulta  =   "Contas"
               TamanhoDescricao=   4700
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
            Begin Fox.EBSText etxCentroCusto 
               Height          =   330
               Left            =   1215
               TabIndex        =   34
               Tag             =   "C. Custo"
               Top             =   1560
               Width           =   5895
               _ExtentX        =   441748
               _ExtentY        =   582
               MaxLength       =   9
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "Código"
               TipoCriterio    =   4
               CampoDescricao  =   "Descrição"
               TabelaConsulta  =   "Centros"
               TamanhoDescricao=   4700
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
            Begin Fox.EBSText etxOpContabil 
               Height          =   330
               Left            =   1215
               TabIndex        =   35
               Tag             =   "Op. Contábil"
               Top             =   1980
               Width           =   5895
               _ExtentX        =   441748
               _ExtentY        =   582
               MaxLength       =   5
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "cd_operacao"
               TipoCriterio    =   4
               CampoDescricao  =   "descricao"
               TabelaConsulta  =   "OperacaoContabil"
               TamanhoDescricao=   4700
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
               Left            =   1215
               TabIndex        =   36
               Tag             =   "Situação"
               Top             =   2400
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
            Begin Fox.EBSText etxCheque 
               Height          =   330
               Left            =   1215
               TabIndex        =   37
               Tag             =   "Cheque"
               Top             =   2850
               Width           =   1530
               _ExtentX        =   265
               _ExtentY        =   582
               MaxLength       =   6
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
            Begin Fox.EBSText etxControle 
               Height          =   330
               Left            =   1215
               TabIndex        =   38
               Tag             =   "Controle"
               Top             =   3240
               Width           =   2715
               _ExtentX        =   265
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
            Begin VB.Label lblExtrato 
               BackColor       =   &H8000000A&
               Height          =   255
               Left            =   3030
               TabIndex        =   143
               Top             =   3690
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label lblSequencialExtrato 
               BackColor       =   &H8000000A&
               Height          =   255
               Left            =   2460
               TabIndex        =   142
               Top             =   3690
               Visible         =   0   'False
               Width           =   525
            End
            Begin VB.Label lblControle 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Controle:"
               Height          =   195
               Left            =   540
               TabIndex        =   120
               Top             =   3300
               Width           =   630
            End
            Begin VB.Label lblCheque 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cheque:"
               Height          =   195
               Left            =   570
               TabIndex        =   118
               Top             =   2880
               Width           =   600
            End
            Begin VB.Label lblSituacao 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Situação:"
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
               TabIndex        =   117
               Top             =   2445
               Width           =   825
            End
            Begin VB.Label lblOpContabil 
               Alignment       =   1  'Right Justify
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
               Left            =   45
               TabIndex        =   116
               Top             =   2040
               Width           =   1125
            End
            Begin VB.Label lblCentroCusto 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "C. Custo:"
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
               Left            =   375
               TabIndex        =   115
               Top             =   1620
               Width           =   795
            End
            Begin VB.Label lblConta 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Conta:"
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
               Left            =   600
               TabIndex        =   114
               Top             =   1200
               Width           =   570
            End
            Begin VB.Label lblBanco 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Banco:"
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
               Left            =   555
               TabIndex        =   113
               Top             =   780
               Width           =   615
            End
            Begin VB.Label lblFormaPagto 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Forma Pagto:"
               Height          =   195
               Left            =   225
               TabIndex        =   112
               Top             =   360
               Width           =   945
            End
         End
         Begin VB.Frame fraSomaValores 
            Caption         =   "Soma dos Valores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7530
            TabIndex        =   81
            Top             =   2640
            Width           =   3615
            Begin Fox.EBSText etxTotal 
               Height          =   330
               Left            =   1320
               TabIndex        =   49
               Tag             =   "Total Geral"
               Top             =   330
               Width           =   2145
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   16
               Enabled         =   0   'False
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
            Begin VB.Label lblTotal 
               AutoSize        =   -1  'True
               Caption         =   "Total Geral:"
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
               Index           =   0
               Left            =   270
               TabIndex        =   82
               Top             =   390
               Width           =   1020
            End
         End
         Begin VB.Frame fraValores 
            Caption         =   "&Valores"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Left            =   7530
            TabIndex        =   40
            Top             =   450
            Width           =   3615
            Begin Fox.EBSText etxValorOriginal 
               Height          =   330
               Left            =   1320
               TabIndex        =   44
               Tag             =   "Valor Original"
               Top             =   795
               Width           =   2145
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   16
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
            Begin Fox.EBSText etxAbatimento 
               Height          =   330
               Left            =   1320
               TabIndex        =   48
               Tag             =   "Abatimento"
               Top             =   1650
               Width           =   2145
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   16
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
            Begin Fox.EBSText etxAcrescimo 
               Height          =   330
               Left            =   1320
               TabIndex        =   46
               Tag             =   "Acréscimo"
               Top             =   1230
               Width           =   2145
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   2
               CasasDecimais   =   2
               MaxLength       =   16
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
            Begin Fox.EBSText etxMoeda 
               Height          =   330
               Left            =   1320
               TabIndex        =   42
               Tag             =   "Moeda"
               Top             =   390
               Width           =   2145
               _ExtentX        =   435054
               _ExtentY        =   582
               Tipo            =   4
               TipoTexto       =   0
               MaxLength       =   10
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "Moeda"
               CampoDescricao  =   "Descrição"
               TabelaConsulta  =   "Moedas"
               TamanhoDescricao=   900
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
            Begin VB.Label lblAcrescimo 
               AutoSize        =   -1  'True
               Caption         =   "Acréscimo:"
               Height          =   195
               Left            =   510
               TabIndex        =   45
               Top             =   1290
               Width           =   780
            End
            Begin VB.Label lblAbatimento 
               AutoSize        =   -1  'True
               Caption         =   "Abatimento:"
               Height          =   195
               Left            =   450
               TabIndex        =   47
               Top             =   1710
               Width           =   840
            End
            Begin VB.Label lblValorOriginal 
               AutoSize        =   -1  'True
               Caption         =   "Valor Original:"
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
               Left            =   75
               TabIndex        =   43
               Top             =   870
               Width           =   1215
            End
            Begin VB.Label lblMoeda 
               AutoSize        =   -1  'True
               Caption         =   "Moeda:"
               Height          =   195
               Left            =   750
               TabIndex        =   41
               Top             =   450
               Width           =   540
            End
         End
         Begin VB.Frame fraLancamentos 
            Caption         =   "Lançamentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2175
            Index           =   2
            Left            =   90
            TabIndex        =   16
            Top             =   450
            Width           =   7365
            Begin VB.CommandButton cmdEfetLanc 
               Caption         =   "Efetivar Lançamento"
               Height          =   375
               Left            =   5430
               TabIndex        =   20
               Top             =   360
               Width           =   1785
            End
            Begin Fox.EBSText etxEmpresa 
               Height          =   330
               Left            =   1215
               TabIndex        =   28
               Tag             =   "Empresa"
               Top             =   1230
               Width           =   4230
               _ExtentX        =   438229
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   15
               PossuiDescricao =   -1  'True
               CampoCriterio   =   "Apel"
               CampoDescricao  =   "Razão"
               TabelaConsulta  =   "Empresas"
               TamanhoDescricao=   2700
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
            Begin Fox.EBSText etxDescricao 
               Height          =   330
               Left            =   1215
               TabIndex        =   30
               Tag             =   "Descrição"
               Top             =   1650
               Width           =   6015
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   80
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
               Left            =   1215
               TabIndex        =   22
               Tag             =   "Tipo"
               Top             =   810
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
            Begin Fox.EBSText etxCodigo 
               Height          =   330
               Left            =   1215
               TabIndex        =   18
               Tag             =   "Codigo/Nota"
               Top             =   390
               Width           =   1785
               _ExtentX        =   265
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
            Begin Fox.EBSText etxParcela 
               Height          =   330
               Left            =   3990
               TabIndex        =   24
               Tag             =   "Parcela"
               Top             =   810
               Width           =   615
               _ExtentX        =   265
               _ExtentY        =   582
               PermiteNegativo =   -1  'True
               MaxLength       =   80
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
            Begin Fox.EBSText etxNrSequencial 
               Height          =   330
               Left            =   5910
               TabIndex        =   26
               Tag             =   "Nr Sequencial"
               Top             =   810
               Width           =   1305
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   80
               Enabled         =   0   'False
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
            Begin Fox.EBSText etxPagRec 
               Height          =   330
               Left            =   3990
               TabIndex        =   19
               Top             =   390
               Visible         =   0   'False
               Width           =   255
               _ExtentX        =   265
               _ExtentY        =   582
               Tipo            =   4
               MaxLength       =   80
               Enabled         =   0   'False
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
            Begin VB.Label lblTipo 
               Alignment       =   1  'Right Justify
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
               Left            =   720
               TabIndex        =   21
               Top             =   855
               Width           =   450
            End
            Begin VB.Label lblNrSequencial 
               AutoSize        =   -1  'True
               Caption         =   "Nr Sequencial:"
               Height          =   195
               Left            =   4830
               TabIndex        =   25
               Top             =   870
               Width           =   1050
            End
            Begin VB.Label lblParcela 
               AutoSize        =   -1  'True
               Caption         =   "Parcela:"
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
               Left            =   3240
               TabIndex        =   23
               Top             =   870
               Width           =   720
            End
            Begin VB.Label lblCodigo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Código:"
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
               Left            =   510
               TabIndex        =   17
               Top             =   450
               Width           =   660
            End
            Begin VB.Label lblDescricao 
               AutoSize        =   -1  'True
               Caption         =   "Descrição:"
               Height          =   195
               Left            =   405
               TabIndex        =   29
               Top             =   1710
               Width           =   765
            End
            Begin VB.Label lblEmpresa 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Empresa:"
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
               Left            =   375
               TabIndex        =   27
               Top             =   1290
               Width           =   795
            End
         End
      End
   End
   Begin VB.Frame Frame 
      Height          =   9765
      Index           =   1
      Left            =   11430
      TabIndex        =   128
      Top             =   -30
      Width           =   1365
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   93
         Top             =   150
         Width           =   1185
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   92
         Top             =   540
         Width           =   1185
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   94
         Top             =   930
         Width           =   1185
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   96
         Top             =   1710
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   98
         Top             =   2490
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   95
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   97
         Top             =   2100
         Width           =   1185
      End
   End
   Begin VB.Image imgInformativa 
      Height          =   480
      Left            =   30
      Picture         =   "frmLancamentoDuplicata.frx":0070
      Top             =   9285
      Width           =   480
   End
   Begin VB.Label lblInformativa 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Caso o código/nota e parcela estiverem com o valor zero, o sistema irá gerar a sugestão automaticamente."
      Height          =   465
      Left            =   510
      TabIndex        =   132
      Top             =   9300
      Width           =   10890
   End
End
Attribute VB_Name = "frmLancamentoDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private menumPagRec                     As enuPagRec
Private menumLancDup                    As enuLancDup
Private mVoLancDup                      As VoLancamentoDuplicata
Private mblnAlterando                   As Boolean
Private mlngCodigo                      As String
Private mlngParcela                     As Long
Private mstrTipo                        As String
Private mstrEmpresa                     As String

Private mVoCheque                       As VoCheque
Private mblnPrevisao                    As Boolean
Private mblnNovoRegistro_Consulta       As Boolean
Private mobjCacheEmpresa                As clsCacheEmpresa
'Projeto: # - História: # - Desenvolvimento# - João Henrique(02/07/2013)
Private mstrEmpresaAnterior             As String
'------------------------------------------------------
Private Const IDB_TRANSF = 509          'Imagem para o ListView para Cheques em Transferências
Private Const IDB_DUPLS = 510           'Ídem para Duplicatas
Private Const IDB_LANCTOS = 511         'Ídem para Lançamentos
'------------------------------------------------------
Private Enum enumModalidade
    Normal = 0
    Baixa = 1
End Enum
'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
Private mdblTotalOrig                   As Double
'Projeto: 17081 - Sugestão de Melhoria: 23370 - Ueder Budni (17/01/2014)
Private mcolBaixasParc                  As cColLancamentoDuplicata
'PJ 61827 - Vinicius Alexandre Elyseu - 23/01/2015
Public mblnOrigemTelaConciliacao        As Boolean

Private Const strBaixasParcHdr$ = "campo=Parcela;label=Parcela;tamanho=700|" & _
                        "campo=Emissao;label=Dt. Emissão;tamanho=1200|" & _
                        "campo=Vencimento;label=Dt. Vencimento;tamanho=1200|" & _
                        "campo=Pagamento;label=Dt. Pagamento;tamanho=1200|" & _
                        "campo=ValorOriginal;label=Valor Original;tamanho=1300;formato=#0.00|" & _
                        "campo=Abatimento;label=Abatimento;tamanho=1300;formato=#0.00|" & _
                        "campo=Acrescimo;label=Acréscimo;tamanho=1300;formato=#0.00|" & _
                        "campo=ValorTotal;label=Valor Total;tamanho=1300;formato=#0.00"
Private mblnQuerExcluir            As Boolean
Private mstrMsgExisteExtrato       As String

'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
Private Enum COL_LOGS
    CL_USUARIO = 1
    CL_DATA_HOTA = 2
    CL_DESCRICAO = 3
End Enum
Private mobjLogLancDup              As clsLogLancamentosDuplicatas
Private mobjOldStateLancDup         As VoLancamentoDuplicata

Public Property Let PagRec(ByVal Valor As enuPagRec)
    menumPagRec = Valor
End Property

Public Property Let LancDup(ByVal Valor As enuLancDup)
    menumLancDup = Valor
End Property

Public Function CarregarLancamentoDuplicataOutrasRotinas(ByVal Codigo As String, Tipo As String, ByVal Parcela As Long, ByVal Empresa As String, ByVal enumPagRec As enuPagRec, ByVal enumLancDup As enuLancDup) As Boolean
    mlngCodigo = Codigo
    mlngParcela = Parcela
    mstrTipo = Tipo
    mstrEmpresa = Empresa
    menumPagRec = enumPagRec
    menumLancDup = enumLancDup
    Call LibProc(WL_CONSULTA)
End Function

Private Sub chkConciliado_Click()
    If ChkConciliado.value = False Then
        VerificaDesconciliaExtrato
    End If
End Sub
Private Sub VerificaDesconciliaExtrato(Optional blnOrigemBotaoExcluir As Boolean)
    Dim objDAO        As CDuplicata
    Dim objDaoExtrato As DaoExtratoBancario
    Dim intIndex      As Integer
    Dim arrSeqExtrato As Variant
    Dim blnErro       As Boolean
    
    Set objDAO = New CDuplicata
    Set objDaoExtrato = New DaoExtratoBancario
    'Verifica se já não existe alguma conciliação
    If objDAO.TemConciliacaoExtrato(CDblDef(etxCodigo.valorTexto, 0), etxEmpresa.valorTexto, IIf(menumPagRec = 0, "P", "R"), cboTipo.SelectedItem, etxVencimento.Data) Then
        If blnOrigemBotaoExcluir Then
            Call DesconciliaLancamento
        Else
            If MsgBox("Existem extratos vinculados a este lançamento/duplicata." & vbNewLine & "Tem certeza que deseja " & IIf(blnOrigemBotaoExcluir, "excluir", "fazer a desconciliação") & "?", vbYesNo, IIf(blnOrigemBotaoExcluir, NomeModulo, "Desconciliação Lançamento/Duplicata ao Extrato")) = vbNo Then
                ChkConciliado.value = 1
                Exit Sub
            Else
                Call DesconciliaLancamento(blnOrigemBotaoExcluir)
            End If
        End If
    End If
End Sub
Private Sub DesconciliaLancamento(Optional blnOrigemBotaoExcluir As Boolean)
    Dim objDAO        As CDuplicata
    Dim objDaoExtrato As DaoExtratoBancario
    Dim intIndex      As Integer
    Dim arrSeqExtrato As Variant
    Dim blnErro       As Boolean
    
    Set objDAO = New CDuplicata
    Set objDaoExtrato = New DaoExtratoBancario
    If objDAO.ConciliaDuplicLanc(False, "0", IIf(menumLancDup = Lancamento, "Lançamentos", "Duplicatas"), 0, etxPagRec.valorTexto, etxCodigo.valorTexto, etxEmpresa.valorTexto, cboTipo.SelectedItem, etxParcela.valorInteiro, etxLiberacao.Data, False, 0) Then
        arrSeqExtrato = Split(lblSequencialExtrato.Caption, ";")
        For intIndex = 0 To UBound(arrSeqExtrato)
            If Not objDaoExtrato.ConciliaExtrato(False, CStr(etxBanco.valorInteiro), CInt(lblExtrato.Caption), CLng(arrSeqExtrato(intIndex))) Then
               blnErro = True
               Exit For
            End If
        Next
        If Not blnOrigemBotaoExcluir Then
            If blnErro Then
                MsgBox "Problema ao desconciliar o lançamento.", vbInformation, "Desconciliação de Título"
                Exit Sub
            Else
                MsgBox "Desconciliação feita com sucesso.", vbInformation, "Desconciliação de Título"
            End If
        End If
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

Private Sub cmdCancelar_Click()
    Call LibProc(WL_CANCELAR)
End Sub

Private Sub cmdEfetLanc_Click()
    If MsgBox("Confirma efetivação do lançamento de previsão ?", vbYesNo, "Confirmação") = vbYes Then
        mblnPrevisao = False
        If LibProc(WL_SALVAR) Then
            cmdEfetLanc.Enabled = False
        End If
    End If
End Sub

Private Sub cmdExcluir_Click()
    Call LibProc(WL_DELETAR)
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
End Sub

Private Sub cmdNominal_Click()
    etxNominal.valorTexto = GetFieldValue("Razão", "Empresas", "Apel = " & Quote(etxEmpresa.valorTexto, "'"), , NUL)
End Sub

Private Sub cmdNovo_Click()
    Call LibProc(WL_NOVO)
End Sub

Private Sub cmdPesquisar_Click()
    Call LibProc(WL_PESQUISAR)
End Sub

Private Sub cmdProxCheque_Click()
    Dim biz As New BizLancamentoDuplicata
    
    etxCheque.valorInteiro = biz.ProximoCheque(etxBanco.valorInteiro)
    Set biz = Nothing
End Sub

Private Sub cmdRateio_Click()
    Dim colTemp As colRateio
    Set colTemp = mVoLancDup.Col_Rateio
    If LibProc(WL_SALVAR, False) Then
        mVoLancDup.Col_Rateio = colTemp
        Call frmLancamentoDuplicataRateio.ObjetoVo(mVoLancDup)
        Load frmLancamentoDuplicataRateio
        Call frmLancamentoDuplicataRateio.CarregaGrid
        Call mostrarForm(frmLancamentoDuplicataRateio, 2980, True)
    End If
    Set colTemp = Nothing
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
    If mblnOrigemTelaConciliacao Then
       frmConciliacaoTitulosAutomatica.RecarregaGrids (False)
       mblnOrigemTelaConciliacao = False
       
       '04/02/2015 - Vinicius A. Elyseu PJ 61827
       'Desconectando porque tem alguma tela que quando é chamada a partir de outra tela, acaba se perdendo.
       'Não conseguimos localizar aonde - por isso este Disconnect esta aqui.
       If Aplicacao.isConnected Then
          Aplicacao.Disconnect
       End If
    End If
End Sub

Private Sub etxAbatimento_LostFocus()
    Call calculoVlrTotal
End Sub

Private Sub etxAcrescimo_LostFocus()
    Call calculoVlrTotal
End Sub

Private Sub etxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    
    If KeyCode = vbKeyPageDown Then
     'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(23/03/2012)
        strSql = "SELECT Banco, Nome " _
        & "FROM Bancos"
        PCampo "Bancos", strSql, PB_CAMPO, etxBanco, "Banco"
    End If
End Sub

Private Sub etxBancoOutros_GotFocus()
    SSTab.Tab = 2
End Sub

Private Sub etxBancoOutros_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Banco, Nome " _
        & "FROM Bancos"
        PCampo "Bancos", strSql, PB_CAMPO, etxBancoOutros, "Banco"
    End If
End Sub

Private Sub etxCentroCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(26/03/2012)
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Código, Descrição, [Data Limite],[cd_conta_contabil], [cd_centro_crd] " _
        & "FROM Centros"
        PCampo "C.Custo", strSql, PB_CAMPO, etxCentroCusto, "Código"
    End If
End Sub

Private Sub etxCodigoOutros_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strApel     As String
    Dim lngCodigo   As Long
    Dim strTipo     As String

    If Shift = 0 And KeyCode = vbKeyPageDown And Trim(etxEmpresa.valorTexto) <> "" Then
        If PMultiCampo("Selecione o endereço", "SELECT [Endereço],Bairro,CEP,Cidade,Estado,Apel,[Código],Tipo FROM [Empresas Endereços] WHERE Tipo = 'Cobrança' AND Apel = '" & etxEmpresa.valorTexto & "'", pbCampo, "Apel;[Código]", strApel, lngCodigo) Then
            etxCodigoOutros.valorInteiro = lngCodigo
            Call ExibeEndereco(strApel, lngCodigo)
        End If
    End If
End Sub

Private Function ExibeEndereco(strApel As String, lngCodigo As Long) As Boolean
    Dim selCmd   As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Dim emp As New CEmpresas
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "[Endereço], Bairro, CEP, Cidade, Estado, Pessoa, [CNPJ/CPF], [IEst/RG], Fone, Ramal, Código"

        .Table.TableName = "[Empresas Endereços]"

        Call .Filter.Append("Apel = @pApel")
        Call .Parameters.add(.CreateParameter("@pApel", strApel, dbFieldTypeString, 15))

        Call .Filter.Append("[Código] = @pCodigo")
        Call .Parameters.add(.CreateParameter("@pCodigo", lngCodigo, dbFieldTypeLong))

        Call .Filter.Append("Tipo = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", "Cobrança", dbFieldTypeString))
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    With rdResult
        If Not .EOF Then
            etxCEPOutros.valorTexto = .GetString("CEP")
            etxCidadeOutros.valorTexto = .GetString("Cidade")
            etxUFOutros.valorTexto = .GetString("Estado")
            etxEnderecoOutros.valorTexto = .GetString("Endereço")
            etxBairroOutros.valorTexto = .GetString("Bairro")
            ExibeEndereco = True
        Else
            etxCEPOutros.Clear
            etxCidadeOutros.Clear
            etxUFOutros.Clear
            etxEnderecoOutros.Clear
            etxBairroOutros.Clear
            ExibeEndereco = False
        End If
    End With
    rdResult.CloseReader
    Set selCmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
End Function

Private Sub etxCodigoOutros_LostFocus()
    If etxCodigoOutros.valorInteiro > 0 Then
        If Not ExibeEndereco(etxEmpresa.valorTexto, etxCodigoOutros.valorInteiro) Then
            etxCodigoOutros.valorInteiro = 0
        End If
    End If
End Sub

Private Sub etxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(26/03/2012)
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Contas.Código as Conta, Contas.Descrição as [Descrição da Conta], " _
        & "Grupos.Código as Grupo, Grupos.Descrição as [Descrição do Grupo] " _
        & "FROM Grupos INNER JOIN Contas ON Grupos.Código = Contas.Grupo where Contas.Ctaati='S' " _
        & "ORDER BY Grupos.Código,Contas.Código"
        PCampo "Conta", strSql, PB_CAMPO, etxConta, "Conta"
    End If
End Sub

Private Sub etxDescricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 60 Or KeyAscii = 62 Then 'Bloquear caracteres "<" e ">"
        KeyAscii = 0
    End If
End Sub

Private Sub etxEmissao_LostFocus()
    lblEmissaoD.Caption = ""
    If etxEmissao.IsValidDate Then
        lblEmissaoD.Caption = Semana(etxEmissao.Data, raUmaPalavra)
    End If
End Sub

Private Sub etxEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(23/03/2012)
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
    Dim objEmpresa      As New CEmpresas
    
    'Projeto: #8404 - História: #9679 - Desenvolvimento#9868 - João Henrique(02/07/2013)
    If Trim(etxEmpresa.valorTexto) <> "" Then
        Set objEmpresa = mobjCacheEmpresa.GetEmpresa(etxEmpresa.valorTexto)
        If Not objEmpresa Is Nothing Then
            'Projeto: #8404 - História: #9679 - Desenvolvimento#9813 - João Henrique(01/07/2013)
            If mstrEmpresaAnterior <> objEmpresa.Apel Then
                If Trim(objEmpresa.Banco) <> Empty And objEmpresa.Banco <> "0" Then
                    etxBanco.valorInteiro = objEmpresa.Banco
                End If
                If Trim(objEmpresa.conta) <> Empty And objEmpresa.conta <> "0" Then
                    etxConta.valorInteiro = objEmpresa.conta
                End If
            End If
        End If
        
        Call DemonstrarInformacaoAdicional
    End If
    'Projeto: #8404 - História: #9679 - Desenvolvimento#9868 - João Henrique(02/07/2013)
    mstrEmpresaAnterior = etxEmpresa.valorTexto
End Sub

Private Sub etxFormaPagto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(26/03/2012)
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Código, Nome, Tipo , Banco, Conta,[Tipo de Exportação],[Gerar KIF]," _
        & "[per_despesa_financeira]" _
        & "FROM [Formas de Pagamento]"
        PCampo "Forma Pagto", strSql, PB_CAMPO, etxFormaPagto, "Código"
    End If
End Sub

Private Sub etxLiberacao_LostFocus()
    lblLiberacaoD.Caption = ""
    If etxLiberacao.IsValidDate Then
        lblLiberacaoD.Caption = Semana(etxLiberacao.Data, raUmaPalavra)
    End If
End Sub

Private Sub etxMoeda_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(26/03/2012)
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Moeda, Descrição, Cotação, Cadastro, Tipo, Singular, " _
        & "Plural, CSingular, CPlural, Símbolo " _
        & "FROM Moedas"
        PCampo "Moeda", strSql, PB_CAMPO, etxMoeda, "Moeda"
    End If
End Sub

Private Sub etxObservacao_GotFocus()
    SSTab.Tab = 1
End Sub

Private Sub etxOpContabil_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(26/03/2012)
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT [cd_operacao], descricao, situacao, [agrupa_centro_custo] " _
        & "FROM OperacaoContabil"
        PCampo "Op. Contabil", strSql, PB_CAMPO, etxOpContabil, "cd_operacao"
    End If
End Sub

Private Sub etxOpContabilBaixa_GotFocus()
    If etxPagamento.IsValidDate Then
        etxOpContabilBaixa.valorInteiro = fSugestaoMatrizContabilizacao(cboTipo.SelectedItem, Baixa)
    End If
End Sub

Private Sub etxOpContabilBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
     Dim strSql  As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(26/03/2012)
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT [cd_operacao], descricao, situacao, [agrupa_centro_custo] " _
        & "FROM OperacaoContabil"
        PCampo "Op. Contabil Baixa", strSql, PB_CAMPO, etxOpContabilBaixa, "cd_operacao"
    End If
End Sub

Private Sub etxPagamento_LostFocus()
    lblPagamentoD.Caption = ""
    If etxPagamento.IsValidDate Then
        etxLiberacao.Data = SugestaoDataLiberacao(etxPagamento.Data)
        lblLiberacaoD.Caption = Semana(etxLiberacao.Data, raUmaPalavra)
        lblPagamentoD.Caption = Semana(etxPagamento.Data, raUmaPalavra)
    End If
End Sub

Private Sub etxPercMora_LostFocus()
    If etxTotal.valorDecimal > 0 Then
        etxVlrMoraDiaria.valorDecimal = Round(etxTotal.valorDecimal * (etxPercMora.valorDecimal / 100), 2)
    Else
        etxVlrMoraDiaria.valorDecimal = 0
    End If
End Sub

Private Sub etxPercMulta_LostFocus()
    If etxTotal.valorDecimal > 0 Then
        etxVlrMulta.valorDecimal = Round(etxTotal.valorDecimal * (etxPercMulta.valorDecimal / 100), 2)
    Else
        etxVlrMulta.valorDecimal = 0
    End If
End Sub

Private Sub etxValorOriginal_LostFocus()
    Call calculoVlrTotal
End Sub


Private Sub etxVencimento_LostFocus()
    lblVencimentoD.Caption = ""
    'Projeto: #1332 - História: #0 - Desenvolvimento#0 - Moacir Pfau(01/11/2012)
    If etxVencimento.IsValidDate Then
        If Not etxPagamento.IsValidDate Then
            etxLiberacao.Data = SugestaoDataLiberacao(etxVencimento.Data)
            lblLiberacaoD.Caption = Semana(etxLiberacao.Data, raUmaPalavra)
        End If
        lblVencimentoD.Caption = Semana(etxVencimento.Data, raUmaPalavra)
    End If
End Sub

Private Sub etxVlrMoraDiaria_LostFocus()
    If etxTotal.valorDecimal > 0 Then
        etxPercMora.valorDecimal = Round(etxVlrMoraDiaria.valorDecimal * 100 / etxTotal.valorDecimal, 2)
    Else
        etxPercMora.valorDecimal = 0
    End If
End Sub

Private Sub etxVlrMulta_LostFocus()
    If etxTotal.valorDecimal > 0 Then
        etxPercMulta.valorDecimal = Round(etxVlrMulta.valorDecimal * 100 / etxTotal.valorDecimal, 2)
    Else
        etxPercMulta.valorDecimal = 0
    End If
End Sub

Private Sub Form_Load()
        
    #If Not FOXSQL = 1 Then
        'Projeto: 100340 - Desenv.: 142881 - Ueder Budni (16/10/2016)
        SSTab.TabVisible(3) = False
        SSTab.TabsPerRow = 3
    #End If
    
    Aplicacao.Connect
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10574 - João Henrique(23/03/2012)
    Call etxEmpresa.AddConexao(Aplicacao)
    Call etxBanco.AddConexao(Aplicacao)
    Call etxFormaPagto.AddConexao(Aplicacao)
    Call etxConta.AddConexao(Aplicacao)
    Call etxCentroCusto.AddConexao(Aplicacao)
    Call etxOpContabil.AddConexao(Aplicacao)
    Call etxOpContabilBaixa.AddConexao(Aplicacao)
    Call etxMoeda.AddConexao(Aplicacao)
    Call etxBancoOutros.AddConexao(Aplicacao)
    Call etxCarteira.AddConexao(Aplicacao)

    Aplicacao.Disconnect

    Call ConfiguracaoInicial
    Call ConfigureList
    
    Call CarregaGridBaixasParciais
    
    If menumLancDup = Lancamento Then
        Call CarregaSequencialExtrato
    End If
    
    'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
    Call CabecalhoGridLog
    Set mobjLogLancDup = New clsLogLancamentosDuplicatas
    
    
    Set mobjCacheEmpresa = New clsCacheEmpresa
    If menumLancDup = Duplicata Then
        etxCodigo.MaxLength = 9
    End If
    If menumPagRec = Pagamento Then
       etxLinhaDigitavel.Enabled = True
       etxLinhaDigitavel.Locked = False
    End If
    
    If ModGeral.ReadOnly Then
        cmdGravar.Enabled = False
        cmdExcluir.Enabled = False
    End If
    
End Sub

 Public Function LibProc(strFuncao As String, Optional blnMostraMensagem As Boolean = True) As Boolean
    Dim biz                 As BizLancamentoDuplicata
    Dim objBloqProcess      As clsBloqueioProcesso
    Dim blnGravar           As Boolean
    Dim blnExcluir          As Boolean
    Dim lngCodigo           As String
    Dim lngParcela          As Long
    Dim strTipo             As String
    Dim strEmpresa          As String
    Dim strArquivoBloqueio  As String
    Dim blnExcluirRateio    As Boolean
    Dim blnVlBPDevolvido    As Boolean
    Dim objDAO              As CDuplicata
    Dim colLog              As Collection
    
On Error GoTo err

    If ModGeral.ReadOnly And (strFuncao = WL_SALVAR Or strFuncao = WL_DELETAR) Then
        MsgBox "Sistema em modo Somente Leitura!", vbInformation, NomeModulo
        LibProc = False
        Exit Function
    End If
    
    Set biz = New BizLancamentoDuplicata
    Set objBloqProcess = New clsBloqueioProcesso
    
    Select Case strFuncao
        Case WL_NOVO
            Call NovoRegistro
            
        Case WL_SAIR
            Unload Me
            Exit Function

        Case WL_SALVAR
            Call fcarregaClasse
            If fValidaCampos() Then
                'Tratamento para gravacao simultanea.
                strArquivoBloqueio = getRetornaArquivoBloqueio
                frmAguarde.Show: DoEvents
                If fVerificaBloqueiaProcesso(strArquivoBloqueio, objBloqProcess) Then
                    If mblnAlterando Then
                        'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
                        blnGravar = True
                        If mdblTotalOrig - etxTotal.valorDecimal <> 0 Then
                            If mVoLancDup.parc_origem_baixa <> 0 Then
                                blnGravar = biz.AdicionaValorParcOrig(mVoLancDup, mdblTotalOrig - etxTotal.valorDecimal, "Abatimento", menumLancDup)
                            End If
                        End If
                        blnGravar = blnGravar And biz.Atualizar(mVoLancDup, mVoCheque)
                        
                        'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
                        If blnGravar Then
                            Call mobjLogLancDup.InsertDiffObject(mobjOldStateLancDup, mVoLancDup)
                            With mVoLancDup
                                mobjLogLancDup.SetKey .PagRec, .Codigo_Nota, .Empresa, .Tipo, .Parcela, .LancDup
                            End With
                            Set colLog = mobjLogLancDup.CarregarLog
                            Call CarregaGridLog(colLog)
                            Set mobjOldStateLancDup = mVoLancDup
                        End If
                        
                        mdblTotalOrig = etxTotal.valorDecimal
                    Else
                        If fValidaProcesso Then
                            'Vinicius Elyseu(07/06/2016) - Projeto: #100340 SP6
                            If mVoLancDup.Codigo_Nota = "" Then
                                mVoLancDup.Codigo_Nota = 0
                            End If
                            blnGravar = biz.Gravar(mVoLancDup, mVoCheque)
                            
                            'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
                            If blnGravar Then
                                With mVoLancDup
                                    mobjLogLancDup.SetKey .PagRec, .Codigo_Nota, .Empresa, .Tipo, .Parcela, .LancDup
                                End With
                                Call mobjLogLancDup.InsertMsg("Título criado.")
                                Set colLog = mobjLogLancDup.CarregarLog
                                Call CarregaGridLog(colLog)
                                Set mobjOldStateLancDup = mVoLancDup
                            End If
                        End If
                    End If
                    'Projeto: #1203 - História: #10564 - Desenvolvimento#10571 - João Henrique(03/04/2012)
                    If blnGravar Then
                        Call preencherCamposAposGravacao
                        mblnAlterando = True
                    End If
                    Call objBloqProcess.DesbloqueiaProcesso
                    Unload frmAguarde: DoEvents
                End If
                
                If blnGravar Then
                    'Vinicius Elyseu (07/03/2016) - Projeto: #100340 / História: #104582
                    #If FOXSQL = 1 Then
                    If etxPagamento.Data <> "00:00:00" Then
                        If DateDiff("m", etxPagamento.Data, Now()) > 0 Then
                            If MsgBox(IIf(mVoLancDup.LancDup = Lancamento, "Este lançamento", "Esta duplicata") & " tem data anterior a data atual e será necessário fazer o Reprocessamento dos Saldos Bancários. Deseja fazer agora?", vbYesNo, "Alerta para Reprocessamento de Saldo") = vbYes Then
                                frmReprocessaSaldo.Show
                                frmReprocessaSaldo.etxBanco.valorInteiro = etxBanco.valorInteiro
                                frmReprocessaSaldo.etxBancoFinal.valorInteiro = etxBanco.valorInteiro
                            End If
                        End If
                    End If
                    #End If
                    If blnMostraMensagem Then
                        Call MensagemAposGravacao
                        Call HabilitaCampoChave(False)
                        Call AbreTelaMensagemConfirmacaoOK(GravacaoAtualizacao)
                    End If
                Else
                    Unload frmAguarde: DoEvents
                    Call AbreTelaMensagemConfirmacaoErro(GravacaoAtualizacao)
                End If
                LibProc = blnGravar
            End If
        Case WL_DELETAR
            If PermissaoExclusao() Then
                Set objDAO = New CDuplicata
                If objDAO.TemConciliacaoExtrato(CDblDef(etxCodigo.valorTexto, 0), etxEmpresa.valorTexto, IIf(menumPagRec = 0, "P", "R"), cboTipo.SelectedItem, etxVencimento.Data) Then
                    mstrMsgExisteExtrato = "Existem extratos vinculados a este lançamento/duplicata." & vbNewLine
                    mblnQuerExcluir = True
                End If
                
                'Projeto: #1203 - História: # - Problema# - João Henrique(18/04/2012)
                If MsgBox(mstrMsgExisteExtrato & "Tem certeza que deseja excluir este registro?", vbYesNo + vbQuestion, "Excluir") = vbYes Then
                
                    'Demanda 131996 - Davi Brito - 22/07/2016
                    If objDAO.TemRemessa(IIf(mVoLancDup.LancDup = Lancamento, "lançamentos", "duplicatas"), CDbl(etxCodigo.valorTexto), etxEmpresa.valorTexto, IIf(menumPagRec = 0, "P", "R"), cboTipo.SelectedItem, etxVencimento.Data) Then
                        If MsgBox(IIf(mVoLancDup.LancDup = Lancamento, "Este lançamento ", "Esta duplicata") & " gerou uma Remessa Bancária. Tem certeza que deseja excluir este registro?", vbYesNo + vbQuestion, "Excluir") = vbNo Then
                            Exit Function
                        End If
                    End If
                
                    If mblnQuerExcluir Then
                        Call VerificaDesconciliaExtrato(True)
                    End If
                    
                    If biz.ExisteRateioLancamentoDuplicataOrigem(menumPagRec, etxCodigo.valorTexto, cboTipo.SelectedItem, etxParcela.valorInteiro, etxEmpresa.valorTexto, menumLancDup) Then
                        blnExcluirRateio = (MsgBox("O registro lançado gerou rateios, gostaria de excluir os registros vinculados?", vbYesNo) = vbYes)
                    End If
                    'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
                    If biz.BaixaParcial(etxCodigo.valorTexto, etxParcela.valorInteiro, menumLancDup) > 0 Then
                        blnVlBPDevolvido = biz.DevolveValorParaParcOrig(etxCodigo.valorTexto, etxParcela.valorInteiro, etxEmpresa.valorTexto, cboTipo.SelectedItem, menumPagRec, menumLancDup)
                    End If
                    blnExcluir = biz.Excluir(menumPagRec, etxCodigo.valorTexto, cboTipo.SelectedItem, etxParcela.valorInteiro, etxEmpresa.valorTexto, menumLancDup, blnExcluirRateio, etxPagamento.Data)
                 
                    If blnExcluir Then
                        'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
                        With mVoLancDup
                            mobjLogLancDup.SetKey IIf(menumPagRec = Pagamento, "P", "R"), etxCodigo.valorTexto, etxEmpresa.valorTexto, cboTipo.SelectedItem, etxParcela.valorInteiro, menumLancDup
                        End With
                        Call mobjLogLancDup.InsertMsg("Título excluído manualmente na rotina " & Me.Caption & ".")
                        
                    
                        'Vinicius Elyseu (07/03/2016) - Projeto: #100340 / História: #104582
                        #If FOXSQL = 1 Then
                        If etxPagamento.Data <> "00:00:00" Then
                            If DateDiff("m", etxPagamento.Data, Now()) > 0 Then
                                If MsgBox(IIf(mVoLancDup.LancDup = Lancamento, "Este lançamento", "Esta duplicata") & " tem data anterior a data atual e será necessário fazer o Reprocessamento dos Saldos Bancários. Deseja fazer agora?", vbYesNo, "Alerta para Reprocessamento de Saldo") = vbYes Then
                                    frmReprocessaSaldo.Show
                                    frmReprocessaSaldo.etxBanco.valorInteiro = etxBanco.valorInteiro
                                    frmReprocessaSaldo.etxBancoFinal.valorInteiro = etxBanco.valorInteiro
                                    ConfigSys.CarregarRegistro
                                End If
                            End If
                        End If
                        #End If
                        Call NovoRegistro
                        Call AbreTelaMensagemConfirmacaoOK(Exclusao)
                    Else
                        Call AbreTelaMensagemConfirmacaoErro(Exclusao)
                    End If
                    LibProc = blnExcluir
                End If
                LibProc = blnExcluir
            End If
        Case WL_PESQUISAR   'Opção Utilizada ao apertar o botão PESQUISAR.
            If PMultiCampo("Consulta - " & Me.Caption, RetornaInstrucaoPesquisa, pbCampo, RetornaCamposPesquisa, lngCodigo, lngParcela, strTipo, strEmpresa) Then
                Call LibProc(WL_NOVO)
                Set mVoLancDup = biz.Carregar(menumPagRec, lngCodigo, strTipo, lngParcela, strEmpresa, menumLancDup)
                If Not mVoLancDup Is Nothing Then
                    Call PreencheCampos
                    mlngCodigo = mVoLancDup.Codigo_Nota
                    mlngParcela = mVoLancDup.Parcela
                    mstrTipo = mVoLancDup.Tipo
                    mstrEmpresa = mVoLancDup.Empresa
                    Call CarregaSequencialExtrato
                    Call calculoVlrTotal
                    'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
                    mdblTotalOrig = etxTotal.valorDecimal
                    
                    'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
                    Call mobjLogLancDup.SetKey(IIf(menumPagRec = Pagamento, "P", "R"), strToDbl(lngCodigo), mstrEmpresa, strTipo, lngParcela, menumLancDup)
                    Set colLog = mobjLogLancDup.CarregarLog
                    Call CarregaGridLog(colLog)
                    Set mobjOldStateLancDup = mVoLancDup
                    
                    Set mcolBaixasParc = biz.CarregarColBaixasParc(mVoLancDup, menumLancDup)
                End If
                'Projeto: 17081 - Sugestão de Melhoria: 23370 - Ueder Budni (17/01/2014)
                Call CarregaGridBaixasParciais
                fgBaixasParc.Enabled = IIf(mcolBaixasParc Is Nothing, False, True)
            End If
            
        Case WL_CANCELAR
            If mblnAlterando Then
                Set mVoLancDup = biz.Carregar(menumPagRec, mlngCodigo, mstrTipo, mlngParcela, mstrEmpresa, menumLancDup)
                If Not mVoLancDup Is Nothing Then
                    Call PreencheCampos
                    Call calculoVlrTotal
                    'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
                    Set mobjOldStateLancDup = mVoLancDup
                End If
            Else
                NovoRegistro
            End If
        Case WL_CONSULTA 'Quando outras rotinas chamam a rotina, valores são passados por propriedade.
            mblnNovoRegistro_Consulta = True
            Call LibProc(WL_NOVO)
            Set mVoLancDup = biz.Carregar(menumPagRec, mlngCodigo, mstrTipo, mlngParcela, mstrEmpresa, menumLancDup)
            If Not mVoLancDup Is Nothing Then
                Call PreencheCampos
                Call calculoVlrTotal
                'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (16/01/2014)
                mdblTotalOrig = etxTotal.valorDecimal
                 
                Set mcolBaixasParc = biz.CarregarColBaixasParc(mVoLancDup, menumLancDup)
                
                'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
                Call mobjLogLancDup.SetKey(IIf(menumPagRec = Pagamento, "P", "R"), strToDbl(mlngCodigo), mstrEmpresa, mstrTipo, mlngParcela, menumLancDup)
                Set colLog = mobjLogLancDup.CarregarLog
                Call CarregaGridLog(colLog)
                Set mobjOldStateLancDup = mVoLancDup
            End If
            'Projeto: 17081 - Sugestão de Melhoria: 23370 - Ueder Budni (17/01/2014)
            Call CarregaGridBaixasParciais
            fgBaixasParc.Enabled = IIf(mcolBaixasParc Is Nothing, False, True)
    End Select
    Call DesabilitaOuHabilitaEfetivaLancamento(CStr(etxCodigo.valorTexto))
    Set objBloqProcess = Nothing
    Exit Function
err:
    Set objBloqProcess = Nothing
    Unload frmAguarde: DoEvents
    Call AbreTelaMessengerBox(branco, "Problema ao executar a rotina.", Me.Caption, True)
End Function

Private Sub DesabilitaOuHabilitaEfetivaLancamento(strCodigo As String)
    Dim rstResult As Object
    Dim strSql As String
    
    If CStr(strCodigo) <> "0" And Trim(strCodigo) <> "" Then
        If menumLancDup = Lancamento Then
            strSql = "SELECT Pagamento FROM Lançamentos WHERE Código = " & strCodigo & " AND Parcela = " & mVoLancDup.Parcela & " AND PagRec = " & Quote(mVoLancDup.PagRec, "''")
    
            If AbreRecordset(rstResult, strSql) = WL_OK Then
                cmdEfetLanc.Enabled = IIf(GetValue(rstResult, "Pagamento", Null), False, True)
            Else
        
                rstResult = NUL
            End If
        End If
    End If
End Sub

Private Function fcarregaClasse()
    'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (17/01/2014)
    Dim lngParcOrigTmp As Long
    
    If mVoLancDup Is Nothing Then
        lngParcOrigTmp = 0
    Else
        lngParcOrigTmp = mVoLancDup.parc_origem_baixa
    End If
    
    Set mVoLancDup = New VoLancamentoDuplicata
    With mVoLancDup
        .PagRec = IIf(menumPagRec = Pagamento, "P", "R")
        .Codigo_Nota = etxCodigo.valorTexto
        .Parcela = etxParcela.valorInteiro
        .Empresa = etxEmpresa.valorTexto
        .Tipo = cboTipo.SelectedItem
        .Descricao = etxDescricao.valorTexto
        .Emissao = etxEmissao.Data
        .Vencimento = etxVencimento.Data
        .Pagamento = etxPagamento.Data
        .Liberacao = etxLiberacao.Data
        .ValorOriginal = etxValorOriginal.valorDecimal
        .Acrescimo = etxAcrescimo.valorDecimal
        .Abatimento = etxAbatimento.valorDecimal
        .Banco = etxBanco.valorInteiro
        .conta = etxConta.valorInteiro
        .Centro = etxCentroCusto.valorInteiro
        .Cheque = etxCheque.valorInteiro
        .Moeda = etxMoeda.valorTexto
        .Controle = etxControle.valorTexto
        .Obs = etxObservacao.Text
        .usuario = UserName
        .Alteracao = Date
        .LINDIG = etxLinhaDigitavelOutros.valorTexto
        If Trim(.LINDIG) = Empty Then
            .LINDIG = etxLinhaDigitavel.valorTexto
        End If
        .PerMrD = etxPercMora.valorDecimal
        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        .SeqNossoNumero = etxNrSequencial.valorTexto
        .CODFPG = etxFormaPagto.valorInteiro
        .CheBan = etxBancoOutros.valorInteiro
        .CheAge = etxAgenciaOutros.valorTexto
        .CheEmi = etxCorrentistaOutros.valorTexto
        .CheCco = etxContaCorrenteOutros.valorTexto
        .VlrMul = etxVlrMulta.valorDecimal
        .VlrMrD = etxVlrMoraDiaria.valorDecimal
        .PerMul = etxPercMulta.valorDecimal
        .PerMrD = etxPercMora.valorDecimal
        .VlrDsP = etxVlrDescPontualidade.valorDecimal
        .NOSNUM = etxNossoNumero.valorTexto
        .cd_operacao_contabil = etxOpContabil.valorInteiro
        .cd_operacao_baixa = etxOpContabilBaixa.valorInteiro
        .Id_carteira = etxCarteira.valorInteiro
        .Situacao = cboSituacao.SelectedItem
        .LancDup = menumLancDup
        .Conciliado = IIf(ChkConciliado.value = 1, True, False)
        .cd_cobranca = etxCodigoOutros.valorInteiro
        'Projeto: 17081 - Desenv.: 22361 - Ueder Budni (17/01/2014)
        .parc_origem_baixa = lngParcOrigTmp
    End With
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10571 - João Henrique(03/04/2012)
    If etxCheque.valorInteiro > 0 Then
        Set mVoCheque = New VoCheque
        With mVoCheque
            .Banco = etxBanco.valorInteiro
            .Cheque = etxCheque.valorInteiro
            .Nominal = etxNominal.valorTexto
            .Situacao = "Normal"
        End With
    End If
End Function

Private Sub PreencheCampos()
    mblnAlterando = True
    With mVoLancDup
        .PagRec = IIf(menumPagRec = Pagamento, "P", "R")
        etxPagRec.valorTexto = .PagRec
        etxCodigo.valorTexto = .Codigo_Nota
        etxParcela.valorInteiro = .Parcela
        etxEmpresa.valorTexto = .Empresa
        cboTipo.SelectItem .Tipo
        etxDescricao.valorTexto = .Descricao
        etxEmissao.Data = .Emissao
        etxVencimento.Data = .Vencimento
        etxPagamento.Data = .Pagamento
        etxLiberacao.Data = .Liberacao
        etxValorOriginal.valorDecimal = .ValorOriginal
        etxAcrescimo.valorDecimal = .Acrescimo
        etxAbatimento.valorDecimal = .Abatimento
        etxBanco.valorInteiro = .Banco
        etxConta.valorInteiro = .conta
        etxCentroCusto.valorInteiro = .Centro
        etxCheque.valorInteiro = .Cheque
        etxMoeda.valorTexto = .Moeda
        etxControle.valorTexto = .Controle
        etxObservacao.Text = .Obs
        etxUsuario.valorTexto = .usuario
        etxAlteracao.Data = .Alteracao
        etxLinhaDigitavel.valorTexto = .LINDIG
        etxPercMora.valorDecimal = .PerMrD
        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        etxNrSequencial.valorTexto = .SeqNossoNumero
        etxFormaPagto.valorInteiro = .CODFPG
        etxBancoOutros.valorInteiro = .CheBan
        etxAgenciaOutros.valorTexto = .CheAge
        etxCorrentistaOutros.valorTexto = .CheEmi
        etxContaCorrenteOutros.valorTexto = .CheCco
        etxVlrMulta.valorDecimal = .VlrMul
        etxVlrMoraDiaria.valorDecimal = .VlrMrD
        etxPercMulta.valorDecimal = .PerMul
        etxPercMora.valorDecimal = .PerMrD
        etxVlrDescPontualidade.valorDecimal = .VlrDsP
        etxNossoNumero.valorTexto = .NOSNUM
        etxOpContabil.valorInteiro = .cd_operacao_contabil
        etxOpContabilBaixa.valorInteiro = .cd_operacao_baixa
        etxCarteira.valorInteiro = .Id_carteira
        cmdEfetLanc.Enabled = .previsao
        cboSituacao.SelectItem .Situacao
        ChkConciliado.value = IIf(.Conciliado = True, 1, 0)
        etxLinhaDigitavelOutros.valorTexto = .LINDIG
        etxCodigoOutros.valorInteiro = .cd_cobranca
        'Projeto: #7373 - História: #6135 - Desenvolvimento: #7433 - Ivo Sousa(10/05/2013)
        If .Id_carteira > 0 And Not IsEmptyDate(.Pagamento) Then
            etxStatusRemessa.valorTexto = "Liquidado"
        'ElseIf .Id_carteira > 0 Then
        'Vinicius Elyseu(01/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
        ElseIf .Remessa Then
            etxStatusRemessa.valorTexto = "Enviado"
        ElseIf Not IsEmptyDate(.Pagamento) Then
            etxStatusRemessa.valorTexto = "Não enviado - Quitado"
        Else
            etxStatusRemessa.valorTexto = "Não enviado"
        End If
    End With
    Call DemonstrarChequeLista
    Call DemonstrarInformacaoAdicional
    cmdRateio.Enabled = True And ConfigSys.ControlarCentrodeCusto
    Call ExibeEndereco(etxEmpresa.valorTexto, etxCodigoOutros.valorInteiro)
    Call HabilitaCampoChave(False)
    Call ValidacaoAposPesquisa
    Call AtualizaSemana
    cmdEfetLanc.Enabled = mVoLancDup.previsao
End Sub

Private Sub LimparCampos()
    etxPagRec.Clear
    etxCodigo.Clear
    etxParcela.Clear
    etxEmpresa.Clear
    etxDescricao.Clear
    etxEmissao.Data = Format(Now, "DD/MM/YYYY")
    etxVencimento.Data = Format(Now, "DD/MM/YYYY")
    etxPagamento.Clear
    etxLiberacao.Data = SugestaoDataLiberacao(etxEmissao.Data)
    etxValorOriginal.Clear
    etxAcrescimo.Clear
    etxAbatimento.Clear
    etxBanco.Clear
    etxConta.Clear
    etxCentroCusto.Clear
    etxCheque.Clear
    etxMoeda.Clear
    etxControle.Clear
    etxObservacao.Text = ""
    etxUsuario.Clear
    etxAlteracao.Clear
    etxLinhaDigitavel.Clear
    etxVlrMoraDiaria.Clear
    etxPercMora.Clear
    etxNrSequencial.Clear
    etxFormaPagto.Clear
    etxBancoOutros.Clear
    etxAgenciaOutros.Clear
    etxCorrentistaOutros.Clear
    etxContaCorrenteOutros.Clear
    etxVlrMulta.Clear
    etxNossoNumero.Clear
    etxOpContabil.valorInteiro = fSugestaoMatrizContabilizacao(cboTipo.SelectedItem, Normal)
    etxOpContabilBaixa.Clear
    etxCarteira.Clear
    etxPercMulta.Clear
    etxPercMora.Clear
    etxVlrDescPontualidade.Clear
    etxCidadeAdicional.Clear
    etxEstadoAdicional.Clear
    etxCodigoOutros.Clear
    etxCEPOutros.Clear
    etxCidadeOutros.Clear
    etxUFOutros.Clear
    etxEnderecoOutros.Clear
    etxBairroOutros.Clear
    cboTipo.SelectItem GetTipoGlobalDefault
    cboSituacao.SelectItem "Normal"
    ChkConciliado.value = 0
    lblEmissaoD.Caption = ""
    lblVencimentoD.Caption = ""
    lblPagamentoD.Caption = ""
    lblLiberacaoD.Caption = ""
    etxNominal.Clear
    etxLinhaDigitavelOutros.Clear
    Call calculoVlrTotal
End Sub

Private Function fSugestaoMatrizContabilizacao(ByVal strTipo As String, ByVal intModalidade As enumModalidade) As Integer
    Dim MatrizDAO As New cMatrizContabilizacaoDAO
    Dim matriz    As cMatrizContabilizacao
    
    Set matriz = MatrizDAO.Carregar(strTipo)
    
    If Not matriz Is Nothing Then
        If menumLancDup = Lancamento Then
            If intModalidade = Baixa Then
                If menumPagRec = Pagamento Then
                    fSugestaoMatrizContabilizacao = matriz.BaixaLancamentosPagar
                Else
                    fSugestaoMatrizContabilizacao = matriz.baixaLancamentosReceber
                End If
            ElseIf intModalidade = Normal Then
                If menumPagRec = Pagamento Then
                    fSugestaoMatrizContabilizacao = matriz.lancamentosPagar
                Else
                    fSugestaoMatrizContabilizacao = matriz.lancamentosReceber
                End If
            End If
        End If
        
        If menumLancDup = Duplicata Then
            If intModalidade = Baixa Then
                If menumPagRec = Pagamento Then
                    fSugestaoMatrizContabilizacao = matriz.BaixaDuplicatasPagar
                Else
                    fSugestaoMatrizContabilizacao = matriz.BaixaDuplicatasReceber
                End If
            ElseIf intModalidade = Normal Then
                If menumPagRec = Pagamento Then
                    fSugestaoMatrizContabilizacao = matriz.duplicatasPagar
                Else
                    fSugestaoMatrizContabilizacao = matriz.duplicatasReceber
                End If
            End If
        End If
    End If
End Function

Private Sub preencherCamposAposGravacao()
    etxCodigo.valorTexto = mVoLancDup.Codigo_Nota
    etxParcela.valorInteiro = mVoLancDup.Parcela
    Call DemonstrarChequeLista
    cmdRateio.Enabled = True And ConfigSys.ControlarCentrodeCusto
    cmdEfetLanc.Enabled = mVoLancDup.previsao
    mlngCodigo = mVoLancDup.Codigo_Nota
    mstrTipo = mVoLancDup.Tipo
    mlngParcela = mVoLancDup.Parcela
    mstrEmpresa = mVoLancDup.Empresa
    etxUsuario.valorTexto = UserName
    etxAlteracao.Data = Date
End Sub

Private Sub NovoRegistro(Optional ByVal bnlProcesso As Boolean)
    LimparCampos
    etxUsuario.valorTexto = UserName
    etxAlteracao.Data = Date
    etxPagRec.valorTexto = IIf(menumPagRec = Pagamento, "P", "R")
    mblnAlterando = False
    SSTab.Tab = 0
    cmdRateio.Enabled = False
    cmdEfetLanc.Enabled = False
    If Not mblnNovoRegistro_Consulta Then
        mlngCodigo = 0
        mstrTipo = ""
        mlngParcela = 0
        mstrEmpresa = ""
        mblnNovoRegistro_Consulta = False
    End If
    Call HabilitaCampoChave(True)
    etxValorOriginal.Enabled = True
    Call AtualizaSemana
    fgBaixasParc.Clear
    'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
    Set mobjOldStateLancDup = Nothing
    Call CabecalhoGridLog
    Call CarregaGridBaixasParciais
    
End Sub

Private Function fValidaCampos() As Boolean
    Dim objBiz              As New BizLancamentoDuplicata
    Dim col                 As New Collection
    
    Call objBiz.validarCampoObrigatorio(etxEmissao.Data, etxVencimento.Data, etxBanco.valorInteiro, etxConta.valorInteiro, _
                                        etxCentroCusto.valorInteiro, etxValorOriginal.valorDecimal, etxParcela.valorInteiro, _
                                        etxOpContabil.valorInteiro, etxLiberacao.Data, etxEmpresa.valorTexto, etxPagamento.Data, etxOpContabilBaixa.valorInteiro, cboSituacao.SelectedItem, etxCodigo.valorTexto, menumPagRec, menumLancDup, col)
                                                                                
    Call objBiz.validarInformacaoGeral(etxEmissao.Data, etxVencimento.Data, etxCentroCusto.valorInteiro, etxConta.valorInteiro, etxLiberacao.Data, etxOpContabilBaixa.valorInteiro, etxPagamento.Data, Financeiro, _
                                       col)
    
    If etxPagamento.IsValidDate Then
        Call objBiz.validarInformacaoBaixa(etxOpContabilBaixa.valorInteiro, etxPagamento.Data, etxVencimento.Data, etxLiberacao.Data, _
                                    etxCheque.valorTexto, etxEmissao.Data, etxOpContabil.valorInteiro, menumLancDup, col)
    End If

    'Exibe eventuais mensagens para o usuário.
    fValidaCampos = AbreTelaMensagem(col, True)
    DoEvents
    Set objBiz = Nothing
End Function

Private Function fValidaProcesso() As Boolean
    Dim objBiz              As New BizLancamentoDuplicata
    Dim col                 As New Collection
    
    If etxCodigo.valorTexto <> "" And etxParcela.valorInteiro > 0 Then
        Call objBiz.validarLancamentoDuplicataExiste(menumPagRec, etxCodigo.valorTexto, cboTipo.SelectedItem, etxParcela.valorInteiro, _
                                                    etxEmpresa.valorTexto, menumLancDup, col)
    End If
    fValidaProcesso = AbreTelaMensagem(col, True)
    DoEvents
    Set objBiz = Nothing
End Function

Private Sub preencheCombo()
    Call preencheComboTipo
    Call preencheComboSituacao
End Sub

Private Sub preencheComboTipo()
    Dim cmd                 As IDBSelectCommand
    Dim rdResult            As IDBReader
    Dim strDefault          As String
  
    Aplicacao.Connect
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
        cboTipo.AddItem rdResult.GetString("Tipo")
        rdResult.MoveNext
    Wend
    rdResult.CloseReader
    
    cboTipo.SelectItem GetTipoGlobalDefault
    
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Sub

Private Sub preencheComboSituacao()
    Dim strDefault          As String
    Dim i                   As Integer
    Dim ArrSituacao() As String
    
    strDefault = "Normal"
    ArrSituacao = Split("Normal;Descontada;Caução;Parcial;Em Cartório;Protestada;Em Cobrança;Jurídico;Devolvida;Cancelada", ";")
    'Projeto: #1332 - História: #0 - Desenvolvimento#0 - Moacir Pfau(01/11/2012)
    For i = 0 To UBound(ArrSituacao)
        cboSituacao.AddItem ArrSituacao(i)
    Next
    
    cboSituacao.SelectItem strDefault
End Sub

Private Sub preencheMensagem()
    Me.Caption = RetornaTituloFormulario
    If menumLancDup = Lancamento Then
        lblCodigo.Caption = "Código:"
    ElseIf menumLancDup = Duplicata Then
        lblCodigo.Caption = "Nota:"
    End If
End Sub

Private Function RetornaTituloFormulario() As String
    If menumLancDup = Lancamento Then
        If menumPagRec = Recebimento Then
            RetornaTituloFormulario = "Lançamentos a Receber ou Recebidos"
        Else
            RetornaTituloFormulario = "Lançamentos a Pagar ou Pagos"
        End If
    ElseIf menumLancDup = Duplicata Then
        If menumPagRec = Recebimento Then
            RetornaTituloFormulario = "Duplicatas a Receber ou Recebidas"
        Else
            RetornaTituloFormulario = "Duplicatas a Pagar ou Pagas"
        End If
    End If
End Function

Private Function GetTipoGlobalDefault() As String
    Dim strTipoGlobal       As String
    
    If menumLancDup = Lancamento Then
        If menumPagRec = Recebimento Then
            strTipoGlobal = ConfigSys.Retorna_tpGlobal_MatrizLancamento(Lancamentos_Receber_Recebidos)
        Else
            strTipoGlobal = ConfigSys.Retorna_tpGlobal_MatrizLancamento(Lancamentos_Pagar_Pagos)
        End If
    ElseIf menumLancDup = Duplicata Then
        If menumPagRec = Recebimento Then
            strTipoGlobal = ConfigSys.Retorna_tpGlobal_MatrizLancamento(Duplicatas_Receber_Recebidas)
        Else
            strTipoGlobal = ConfigSys.Retorna_tpGlobal_MatrizLancamento(Duplicatas_Pagar_Pagas)
        End If
    End If
    
    If strTipoGlobal = "" Then
        strTipoGlobal = "Fatura"
    End If
    GetTipoGlobalDefault = strTipoGlobal
End Function

Private Function RetornaInstrucaoPesquisa() As String
    Dim strSql          As String
    
    If menumLancDup = Lancamento Then
        If menumPagRec = Recebimento Then
            strSql = " SELECT * FROM [Lançamentos] WHERE [PagRec] = 'R'"
        Else
            strSql = " SELECT * FROM [Lançamentos] WHERE [PagRec] = 'P'"
        End If
    ElseIf menumLancDup = Duplicata Then
        If menumPagRec = Recebimento Then
            strSql = " SELECT * FROM [Duplicatas] WHERE [PagRec] = 'R'"
        Else
            strSql = " SELECT * FROM [Duplicatas] WHERE [PagRec] = 'P'"
        End If
    End If
    RetornaInstrucaoPesquisa = strSql
End Function

Private Function RetornaCamposPesquisa() As String
    Dim strCp          As String
    
    If menumLancDup = Lancamento Then
        strCp = "Código;Parcela;Tipo;''"
    ElseIf menumLancDup = Duplicata Then
        strCp = "Nota;Parcela;Tipo;Empresa"
    End If
    RetornaCamposPesquisa = strCp
End Function

Private Sub ConfiguracaoInicial()
    Call preencheCombo
    Call NovoRegistro
    Call preencheMensagem
    'Configuração cheque quando for apagar
    etxCheque.Enabled = (menumPagRec = Pagamento)
    lblCheque.Enabled = (menumPagRec = Pagamento)
    cmdProxCheque.Enabled = (menumPagRec = Pagamento)
    'Integração Contabil
    etxOpContabil.Enabled = ConfigSys.UtilizaIntegracaoContabil
    lblOpContabil.Enabled = ConfigSys.UtilizaIntegracaoContabil
    etxOpContabilBaixa.Enabled = ConfigSys.UtilizaIntegracaoContabil
    lblOpContabilBaixa.Enabled = ConfigSys.UtilizaIntegracaoContabil
    etxCentroCusto.Enabled = ConfigSys.ControlarCentrodeCusto

    etxNominal.Enabled = (menumPagRec = Pagamento)
    cmdNominal.Enabled = (menumPagRec = Pagamento)
    etxHistorico.Enabled = (menumPagRec = Pagamento)
    lvwLancamentos.Enabled = (menumPagRec = Pagamento)

    cmdEfetLanc.Visible = (menumLancDup = Lancamento)
    SSTab.Tab = 0

    'Projeto: #1332 - História: #0 - Desenvolvimento#0 - Moacir Pfau(01/11/2012)
    lblInformativa.Visible = Not (menumPagRec = Pagamento And menumLancDup = Duplicata)
    imgInformativa.Visible = Not (menumPagRec = Pagamento And menumLancDup = Duplicata)
End Sub

Private Sub calculoVlrTotal()
    etxTotal.valorDecimal = Round((etxValorOriginal.valorDecimal + etxAcrescimo.valorDecimal - etxAbatimento.valorDecimal), 2)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : cwb.atualizacao_fox
' Date      : 05/02/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Form_Unload_Error

    Set mobjCacheEmpresa = Nothing

    'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (23/09/2016)
    Set mobjLogLancDup = Nothing
    
    On Error GoTo 0
    Exit Sub

Form_Unload_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure Form_Unload of Formulário frmLancamentoDuplicata"
End Sub

'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (23/09/2016)
Private Sub grdLog_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim CurrY, CurrRow As Integer
    
    With grdLog
        CurrY = Int(Y / .RowHeight(1))
        CurrRow = CurrY + (.TopRow) - 1
        If Not CurrRow > .Rows - 1 And .MouseCol = 3 Then
            .ToolTipText = .TextMatrix(CurrRow, 3)
        Else
            .ToolTipText = ""
        End If
    End With

End Sub

'======================================================================================
Private Sub lvwLancamentos_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvwLancamentos.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub ConfigureList()
    lvwLancamentos.ColumnHeaders.add 1, , "Número", 975, lvwColumnLeft
    lvwLancamentos.ColumnHeaders.add 2, , "Tipo", 975, lvwColumnLeft
    lvwLancamentos.ColumnHeaders.add 3, , "Empresa", 1440, lvwColumnLeft
    lvwLancamentos.ColumnHeaders.add 4, , "Valor", 1440, lvwColumnRight

    imgDupl.ImageHeight = 16
    imgDupl.ImageWidth = 16
    imgDupl.MaskColor = vbWhite
    imgDupl.UseMaskColor = True
    imgDupl.ListImages.add 1, "transferencia", LoadResBitmap(IDB_TRANSF)
    imgDupl.ListImages.add 2, "duplicata", LoadResBitmap(IDB_DUPLS)
    imgDupl.ListImages.add 3, "lancamento", LoadResBitmap(IDB_LANCTOS)

    lvwLancamentos.SmallIcons = imgDupl
End Sub

Private Sub DemonstrarChequeLista()
    Dim lngBanco As Long
    Dim lngCheque As Long
    Dim strCheque As String
    Dim cValor As Double
    If etxCheque.valorInteiro > 0 Then 'Se há um cheque visível agora
        lngBanco = etxBanco.valorInteiro
        lngCheque = etxCheque.valorInteiro
        strCheque = wsprintf("SELECT * FROM Cheque WHERE " & "Banco = %l AND Cheque = %l", lngBanco, lngCheque)
    
        SetPtrWait Me
        If gTipoDB = Access Then
            wvsprintf strCheque, "SELECT FORMAT(Nota, \'000000\') & ' - ' & " & "FORMAT(Parcela, \'00\') AS Cod, Tipo, Empresa, " & "FORMAT(([Valor Original] + Acréscimo - Abatimento), " & "\'###,###,###,##0.00\') AS Total FROM Duplicatas WHERE PagRec = " & "'P' AND Banco = %l AND Cheque = %l;", lngBanco, lngCheque
        Else
            wvsprintf strCheque, "SELECT (CAST(Nota as varchar(20)) +  ' - ' + " & "CAST(Parcela as varchar(5))) AS Cod, Tipo, Empresa, " & "([Valor Original] + Acréscimo - Abatimento) " & " AS Total FROM Duplicatas WHERE PagRec = " & "'P' AND Banco = %l AND Cheque = %l;", lngBanco, lngCheque
        End If
        Call ListViewAddItem(lvwLancamentos, strCheque, "duplicata")
        If gTipoDB = Access Then
            wvsprintf strCheque, "SELECT FORMAT(Código, \'000000\') AS Cod, Tipo, Empresa, FORMAT(([Valor Original] + Acréscimo - Abatimento), \'###,###,###,##0.00\') AS Total FROM Lançamentos WHERE PagRec = 'P' AND Banco = %l AND Cheque = %l;", lngBanco, lngCheque
        Else
            wvsprintf strCheque, "SELECT Código AS Cod, Tipo, Empresa, ([Valor Original] + Acréscimo - Abatimento) AS Total FROM Lançamentos WHERE PagRec = 'P' AND Banco = %l AND Cheque = %l;", lngBanco, lngCheque
        End If
        Call ListViewAddItem(lvwLancamentos, strCheque, "lancamento")
        If gTipoDB = Access Then
            wvsprintf strCheque, "SELECT FORMAT(T.Código, \'000000\') As Cod, 'Transferência', B.Nome, FORMAT(T.Valor, \'###,###,###,##0.00\') FROM [Transf Bancária] AS T, Bancos As B WHERE B.Banco = T.Origem AND T.Origem = %l AND T.Cheque = %l;", lngBanco, lngCheque
        Else
            wvsprintf strCheque, "SELECT T.Código As Cod, 'Transferência', B.Nome, T.Valor FROM [Transf Bancária] AS T, Bancos As B WHERE B.Banco = T.Origem AND T.Origem = %l AND T.Cheque = %l;", lngBanco, lngCheque
        End If
        Call ListViewAddItem(lvwLancamentos, strCheque, "transferencia")
        'Calculando o valor do cheque para exibição na janela
        cValor = Soma("[Valor Original] + Acréscimo - Abatimento", "Duplicatas", wsprintf("PagRec = 'P' AND Banco = %l AND Cheque = %l", lngBanco, lngCheque), ZERO)
        cValor = cValor + Soma("[Valor Original] + Acréscimo - Abatimento", "Lançamentos", wsprintf("PagRec = 'P' AND Banco = %l AND Cheque = %l", lngBanco, lngCheque), ZERO)
        cValor = cValor + Soma("Valor", "Transf Bancária", wsprintf("Banco = %l AND Cheque = %l", lngBanco, lngCheque), ZERO)
        etxTotalInfCheque.valorDecimal = Format$(cValor, FMOEDA)
        SetPtrDef Me
    Else
        etxTotalInfCheque.valorDecimal = 0#
    End If
End Sub
'======================================================================================

Private Sub DemonstrarInformacaoAdicional()
    Dim objEmpresa      As New CEmpresas
    Set objEmpresa = mobjCacheEmpresa.GetEmpresa(etxEmpresa.valorTexto)
    If Not objEmpresa Is Nothing Then
        etxCidadeAdicional.valorTexto = objEmpresa.Cidade
        etxEstadoAdicional.valorTexto = objEmpresa.Estado
    End If
End Sub

Public Sub setVo(ByVal Valor As VoLancamentoDuplicata, ByVal blnRateio As Boolean)
    If Not mVoLancDup Is Nothing Then
        Set mVoLancDup = Valor
    End If
End Sub

Private Sub MensagemAposGravacao()
    Dim biz As New BizLancamentoDuplicata
    Dim col As New Collection
    
    Call biz.mensagemAposBaixaInformativo(etxPagamento.Data, etxVencimento.Data, etxEmpresa.valorTexto, col)
    
    If Not col Is Nothing Then
        If col.Count > 0 Then
            Call AbreTelaMensagem(col, True, "Mensagem Informativa")
        End If
    End If
    Set biz = Nothing
End Sub

Private Function getRetornaArquivoBloqueio() As String
    Dim strRotina       As String
    
    If menumLancDup = Lancamento Then
        strRotina = "Lancamento"
    ElseIf menumLancDup = Duplicata Then
        strRotina = "Duplicata"
    End If
    getRetornaArquivoBloqueio = strRotina
End Function

Private Function fVerificaBloqueiaProcesso(ByVal strArquivoBloqueio As String, ByVal objBloqProcess As clsBloqueioProcesso) As Boolean
    ' 24/03/2020 - HyperCube: INC-26937 - Yuji F. - Ajustes no bloqueio de processos para os casos de concorrência (vários usuários usando a mesma rotina)
    objBloqProcess.TipoRegistro = strArquivoBloqueio
    If objBloqProcess.VerificaLocked Then
        fVerificaBloqueiaProcesso = False
        Call AbreTelaMessengerBox(Alerta, "Processo de gravação/atualização esta sendo realizado por outro usuário. Favor tentar novamente.", NomeModulo, True)

        Exit Function
    End If

    fVerificaBloqueiaProcesso = objBloqProcess.BloqueiaProcesso
End Function

Private Sub HabilitaCampoChave(ByVal blnHabilita As Boolean)
    etxCodigo.Enabled = blnHabilita
    cboTipo.Enabled = blnHabilita
    etxParcela.Enabled = blnHabilita

    If menumLancDup = Duplicata Then
        etxEmpresa.Enabled = blnHabilita
    End If
End Sub

Private Function PermissaoExclusao() As Boolean
    Dim objBiz              As New BizLancamentoDuplicata
    Dim col                 As New Collection
                                                                                   
    Call objBiz.validarPermissaoExclusao(etxCodigo.valorTexto, etxParcela.valorInteiro, cboTipo.SelectedItem, etxEmpresa.valorTexto, _
                                        menumLancDup, menumPagRec, etxEmissao.Data, Financeiro, col)
    
    'Exibe eventuais mensagens para o usuário.
    PermissaoExclusao = AbreTelaMensagem(col, True)
    DoEvents
    Set objBiz = Nothing
End Function

Private Function SugestaoDataLiberacao(ByVal datData As Date) As Date
    Dim datLiberacao        As Date
    Dim biz                 As New BizBanco
    
    datLiberacao = datData
    If menumPagRec = Recebimento Then
        datLiberacao = DateAdd("d", biz.DiasLiberacao(etxBanco.valorInteiro), datLiberacao)
        If calendario.PermiteLancamento(datLiberacao, , False) <> "A" Then
            datLiberacao = datLiberacao + NumeroDiasUteisNaoUteis(datLiberacao, 0)
        End If
    End If
    SugestaoDataLiberacao = datLiberacao
End Function

Private Sub ValidacaoAposPesquisa()
    'Se a duplicada veio de uma nota fiscal nao podera ser alterado o seu valor.
    etxValorOriginal.Enabled = Not BloqueiaValorOriginal
End Sub

Private Function BloqueiaValorOriginal() As Boolean
    Dim biz As New BizLancamentoDuplicata
    Dim blnBloqueiaValorOriginal            As Boolean
    
    blnBloqueiaValorOriginal = blnBloqueiaValorOriginal Or (biz.PertenceNota(mVoLancDup.Codigo_Nota, mVoLancDup.Tipo, IIf(mVoLancDup.PagRec = "R", enuPagRec.Recebimento, enuPagRec.Pagamento), mVoLancDup.Empresa))
    blnBloqueiaValorOriginal = blnBloqueiaValorOriginal Or (biz.PertencePedido(mVoLancDup.Codigo_Nota, mVoLancDup.Tipo, IIf(mVoLancDup.PagRec = "R", enuPagRec.Recebimento, enuPagRec.Pagamento), mVoLancDup.Empresa))
    BloqueiaValorOriginal = blnBloqueiaValorOriginal
    
    Set biz = Nothing
End Function

'Projeto: #1332 - História: # - Desenvolvimento# - Moacir Pfau(01/11/2012)


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oHelpHtml As New clsHelp
    If KeyCode = vbKeyF1 Then
        oHelpHtml.Origem = 0
        oHelpHtml.hWnd = Me.hWnd
        oHelpHtml.HelpContext = Me.HelpContextID
        Call oHelpHtml.ShowHelp
        Set oHelpHtml = Nothing
    End If
    If KeyCode = vbKeyF3 Then
        Call LibProc(WL_PESQUISAR)
    End If
End Sub

Private Sub AtualizaSemana()
    lblEmissaoD.Caption = ""
    If etxEmissao.IsValidDate Then
        lblEmissaoD.Caption = Semana(etxEmissao.Data, raUmaPalavra)
    End If

    lblVencimentoD.Caption = ""
    If etxVencimento.IsValidDate Then
        lblVencimentoD.Caption = Semana(etxVencimento.Data, raUmaPalavra)
    End If

    lblPagamentoD.Caption = ""
    If etxPagamento.IsValidDate Then
        lblPagamentoD.Caption = Semana(etxPagamento.Data, raUmaPalavra)
    End If

    lblLiberacaoD.Caption = ""
    If etxLiberacao.IsValidDate Then
        lblLiberacaoD.Caption = Semana(etxLiberacao.Data, raUmaPalavra)
    End If
    
End Sub

'Projeto: 17081 - Sugestão de Melhoria: 23370 - Ueder Budni (17/01/2014)
Private Sub CarregaGridBaixasParciais()
    fgBaixasParc.Clear
    If mcolBaixasParc Is Nothing Then
        Call CarregaHFlexGrid(fgBaixasParc, Nothing, strBaixasParcHdr)
    Else
        If mcolBaixasParc.Count = 0 Then
            Call CarregaHFlexGrid(fgBaixasParc, Nothing, strBaixasParcHdr)
        Else
            mcolBaixasParc.MoveFirst
            Call CarregaHFlexGrid(fgBaixasParc, , strBaixasParcHdr, , , mcolBaixasParc)
        End If
    End If
End Sub

Private Sub CarregaSequencialExtrato()
    Dim cDaoDuplicata As CDuplicata
    
    Set cDaoDuplicata = New CDuplicata
    lblExtrato.Caption = cDaoDuplicata.RetornaExtratoConciliado(IIf(menumLancDup = Lancamento, "Lançamentos", "Duplicatas"), etxCodigo.valorTexto, etxEmpresa.valorTexto, etxPagRec.valorTexto, cboTipo.SelectedItem)
    lblSequencialExtrato.Caption = cDaoDuplicata.RetornaSequencialExtratoConciliado(IIf(menumLancDup = Lancamento, "Lançamentos", "Duplicatas"), etxCodigo.valorTexto, etxEmpresa.valorTexto, etxPagRec.valorTexto, cboTipo.SelectedItem)
End Sub

'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
Private Sub CabecalhoGridLog()
    With grdLog
        .Cols = 4
        .FixedCols = 1
        .Rows = 2
        .FixedRows = 1
        
        .Clear
        
        'Coluna Fixa
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 150
        
        'Coluna Usuário
        .TextMatrix(0, CL_USUARIO) = "Usuário"
        .ColWidth(CL_USUARIO) = 1500
        .ColAlignment(CL_USUARIO) = flexAlignRightCenter
        
        'Coluna Data/Hora
        .TextMatrix(0, CL_DATA_HOTA) = "Data/Hora"
        .ColWidth(CL_DATA_HOTA) = 2000
        .ColAlignment(CL_DATA_HOTA) = flexAlignRightCenter
        
        'Coluna da Descrição
        .TextMatrix(0, CL_DESCRICAO) = "Descrição"
        .ColWidth(CL_DESCRICAO) = 7000
        .ColAlignment(CL_DESCRICAO) = flexAlignLeftCenter

    End With
End Sub

'Projeto: 100340 - Desenv.: 142885 - Ueder Budni (20/09/2016)
Private Sub CarregaGridLog(colMensagens As Collection)
    Dim strSplittedMsg()    As String
    Dim strLinha            As String
    Dim i                   As Integer
    
    
    CabecalhoGridLog
    If Not colMensagens Is Nothing Then
        With grdLog
    
            Call CabecalhoGridLog
            
            For i = 1 To colMensagens.Count
                strSplittedMsg = Split(colMensagens(i), ";")
                
                strLinha = Chr(vbKeyTab) & strSplittedMsg(0) & _
                           Chr(vbKeyTab) & strSplittedMsg(1) & _
                           Chr(vbKeyTab) & strSplittedMsg(2)
                Call .AddItem(strLinha)
                                
                DoEvents
            Next
            If .Rows > 2 Then
                If .TextMatrix(1, 1) = "" Then
                    Call .RemoveItem(1)
                End If
            End If
        End With
    End If
End Sub
