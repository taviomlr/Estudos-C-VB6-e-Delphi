VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCadastroCamara 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Câmaras"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   708
   Begin VB.Frame Frame1 
      Height          =   5250
      Left            =   30
      TabIndex        =   35
      Top             =   -60
      Width           =   9120
      Begin Fox.EBSText etxCodigo 
         Height          =   330
         Left            =   990
         TabIndex        =   0
         Top             =   270
         Width           =   750
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   5
         ValorSelecionado=   -1  'True
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
      Begin Fox.EBSCombo ecbStatus 
         Height          =   315
         Left            =   7845
         TabIndex        =   1
         Top             =   270
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         OrigemDados     =   2
         Dados           =   "Ativo;Inativo"
         DadosAssist     =   "A;I"
         DefaultValue    =   "Ativo"
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
         Left            =   990
         TabIndex        =   2
         Top             =   690
         Width           =   7965
         _ExtentX        =   5556
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
      Begin TabDlg.SSTab tabRegistros 
         Height          =   4035
         Left            =   60
         TabIndex        =   3
         Top             =   1140
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   7117
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Tipo de Serviço"
         TabPicture(0)   =   "frmCadastroCamara.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label5(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "etxDescTipoServico"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "etxCodTipoServico"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "grdTipoServico"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdExcluirTipoServico"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdConfTipoServico"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdCancelTipoServico"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "imgOrderCodTipoServ"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "imgOrderDescTipoServ"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Forma Lançamento"
         TabPicture(1)   =   "frmCadastroCamara.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(1)=   "Label7"
         Tab(1).Control(2)=   "etxDescFormaLancamento"
         Tab(1).Control(3)=   "etxCodFormaLancamento"
         Tab(1).Control(4)=   "grdFormaLancamento"
         Tab(1).Control(5)=   "cmdExcluirFormaLancamento"
         Tab(1).Control(6)=   "cmdConfFormaLancamento"
         Tab(1).Control(7)=   "cmdCancelFormaLancamento"
         Tab(1).Control(8)=   "imgOrderDescFormaLanc"
         Tab(1).Control(9)=   "imgOrderCodFormaLanc"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Tipo do Movimento"
         TabPicture(2)   =   "frmCadastroCamara.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label8"
         Tab(2).Control(1)=   "Label10"
         Tab(2).Control(2)=   "etxDescTipoMovimento"
         Tab(2).Control(3)=   "etxCodTipoMovimento"
         Tab(2).Control(4)=   "grdTipoMovimento"
         Tab(2).Control(5)=   "cmdExcluirTipoMovimento"
         Tab(2).Control(6)=   "cmdConfTipoMovimento"
         Tab(2).Control(7)=   "cmdCancelTipoMovimento"
         Tab(2).Control(8)=   "imgOrderDescTipoMovimento"
         Tab(2).Control(9)=   "imgOrderCodTipoMovimento"
         Tab(2).ControlCount=   10
         TabCaption(3)   =   "Código Movimento"
         TabPicture(3)   =   "frmCadastroCamara.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "imgOrderCodMovimento"
         Tab(3).Control(1)=   "imgOrderDescCodMovimento"
         Tab(3).Control(2)=   "cmdCancelCodMovimento"
         Tab(3).Control(3)=   "cmdExcluirCodMovimento"
         Tab(3).Control(4)=   "cmdConfCodMovimento"
         Tab(3).Control(5)=   "grdCodigoMovimento"
         Tab(3).Control(6)=   "etxCodMovimento"
         Tab(3).Control(7)=   "etxDescCodMovimento"
         Tab(3).Control(8)=   "Label5(1)"
         Tab(3).Control(9)=   "Label9"
         Tab(3).ControlCount=   10
         TabCaption(4)   =   "Ocorrência Retorno"
         TabPicture(4)   =   "frmCadastroCamara.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "imgOrderCodOcorrencia"
         Tab(4).Control(1)=   "imgOrderDescOcorrencia"
         Tab(4).Control(2)=   "cmdConfOcorrencia"
         Tab(4).Control(3)=   "cmdExcluirOcorrencia"
         Tab(4).Control(4)=   "cmdCancelaOcorrencia"
         Tab(4).Control(5)=   "grdOcorrencias"
         Tab(4).Control(6)=   "etxCodOcorrencia"
         Tab(4).Control(7)=   "etxDescOcorrencia"
         Tab(4).Control(8)=   "Label11"
         Tab(4).Control(9)=   "Label5(2)"
         Tab(4).ControlCount=   10
         Begin VB.PictureBox imgOrderCodOcorrencia 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   -74700
            Picture         =   "frmCadastroCamara.frx":008C
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   64
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderDescOcorrencia 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   -73590
            Picture         =   "frmCadastroCamara.frx":0327
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   63
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderCodMovimento 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   -74700
            Picture         =   "frmCadastroCamara.frx":05C2
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   62
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderDescCodMovimento 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   -73590
            Picture         =   "frmCadastroCamara.frx":085D
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   61
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderCodTipoMovimento 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   -74700
            Picture         =   "frmCadastroCamara.frx":0AF8
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   60
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderDescTipoMovimento 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   -73590
            Picture         =   "frmCadastroCamara.frx":0D93
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   59
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderCodFormaLanc 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   -74700
            Picture         =   "frmCadastroCamara.frx":102E
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   58
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderDescFormaLanc 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   -73590
            Picture         =   "frmCadastroCamara.frx":12C9
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   57
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderDescTipoServ 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   1410
            Picture         =   "frmCadastroCamara.frx":1564
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   56
            Top             =   1380
            Width           =   270
         End
         Begin VB.PictureBox imgOrderCodTipoServ 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   300
            Picture         =   "frmCadastroCamara.frx":17FF
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   55
            Top             =   1380
            Width           =   270
         End
         Begin VB.CommandButton cmdConfOcorrencia 
            Caption         =   "&Confirmar"
            Height          =   375
            Left            =   -69780
            TabIndex        =   26
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdExcluirOcorrencia 
            Caption         =   "&Excluir"
            Height          =   375
            Left            =   -68550
            TabIndex        =   27
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelaOcorrencia 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   -67320
            TabIndex        =   28
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelTipoServico 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   7680
            TabIndex        =   8
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfTipoServico 
            Caption         =   "&Confirmar"
            Height          =   375
            Left            =   5220
            TabIndex        =   6
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdExcluirTipoServico 
            Caption         =   "&Excluir"
            Height          =   375
            Left            =   6450
            TabIndex        =   7
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelFormaLancamento 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   -67320
            TabIndex        =   13
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfFormaLancamento 
            Caption         =   "&Confirmar"
            Height          =   375
            Left            =   -69780
            TabIndex        =   11
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdExcluirFormaLancamento 
            Caption         =   "&Excluir"
            Height          =   375
            Left            =   -68550
            TabIndex        =   12
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelTipoMovimento 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   -67320
            TabIndex        =   18
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfTipoMovimento 
            Caption         =   "&Confirmar"
            Height          =   375
            Left            =   -69780
            TabIndex        =   16
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdExcluirTipoMovimento 
            Caption         =   "&Excluir"
            Height          =   375
            Left            =   -68550
            TabIndex        =   17
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelCodMovimento 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   -67320
            TabIndex        =   23
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdExcluirCodMovimento 
            Caption         =   "&Excluir"
            Height          =   375
            Left            =   -68550
            TabIndex        =   22
            Top             =   420
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfCodMovimento 
            Caption         =   "&Confirmar"
            Height          =   375
            Left            =   -69780
            TabIndex        =   21
            Top             =   420
            Width           =   1215
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTipoServico 
            Height          =   2655
            Left            =   60
            TabIndex        =   39
            Top             =   1320
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   4683
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin Fox.EBSText etxCodTipoServico 
            Height          =   330
            Left            =   960
            TabIndex        =   4
            Top             =   480
            Width           =   750
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   2
            ValorSelecionado=   -1  'True
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
         Begin Fox.EBSText etxDescTipoServico 
            Height          =   330
            Left            =   960
            TabIndex        =   5
            Top             =   900
            Width           =   7935
            _ExtentX        =   5503
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   100
            ValorSelecionado=   -1  'True
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFormaLancamento 
            Height          =   2655
            Left            =   -74940
            TabIndex        =   42
            Top             =   1320
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   4683
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin Fox.EBSText etxCodFormaLancamento 
            Height          =   330
            Left            =   -74040
            TabIndex        =   9
            Top             =   480
            Width           =   750
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   2
            ValorSelecionado=   -1  'True
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTipoMovimento 
            Height          =   2655
            Left            =   -74940
            TabIndex        =   45
            Top             =   1320
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   4683
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin Fox.EBSText etxCodTipoMovimento 
            Height          =   330
            Left            =   -74040
            TabIndex        =   14
            Top             =   480
            Width           =   750
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   1
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
         Begin Fox.EBSText etxDescTipoMovimento 
            Height          =   330
            Left            =   -74040
            TabIndex        =   15
            Top             =   900
            Width           =   7935
            _ExtentX        =   5503
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCodigoMovimento 
            Height          =   2655
            Left            =   -74940
            TabIndex        =   48
            Top             =   1320
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   4683
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin Fox.EBSText etxCodMovimento 
            Height          =   330
            Left            =   -74040
            TabIndex        =   19
            Top             =   480
            Width           =   750
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   2
            ValorSelecionado=   -1  'True
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
         Begin Fox.EBSText etxDescCodMovimento 
            Height          =   330
            Left            =   -74040
            TabIndex        =   20
            Top             =   900
            Width           =   7935
            _ExtentX        =   5503
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   100
            ValorSelecionado=   -1  'True
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
         Begin Fox.EBSText etxDescFormaLancamento 
            Height          =   330
            Left            =   -74040
            TabIndex        =   10
            Top             =   900
            Width           =   7935
            _ExtentX        =   5503
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   100
            ValorSelecionado=   -1  'True
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdOcorrencias 
            Height          =   2655
            Left            =   -74940
            TabIndex        =   52
            Top             =   1320
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   4683
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin Fox.EBSText etxCodOcorrencia 
            Height          =   330
            Left            =   -74040
            TabIndex        =   24
            Top             =   480
            Width           =   750
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
         Begin Fox.EBSText etxDescOcorrencia 
            Height          =   330
            Left            =   -74040
            TabIndex        =   25
            Top             =   900
            Width           =   7935
            _ExtentX        =   5503
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   -74850
            TabIndex        =   54
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   2
            Left            =   -74625
            TabIndex        =   53
            Top             =   525
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   1
            Left            =   -74625
            TabIndex        =   50
            Top             =   525
            Width           =   495
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   -74850
            TabIndex        =   49
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   -74625
            TabIndex        =   47
            Top             =   525
            Width           =   495
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   -74850
            TabIndex        =   46
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   -74625
            TabIndex        =   44
            Top             =   525
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   -74850
            TabIndex        =   43
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   41
            Top             =   525
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   150
            TabIndex        =   40
            Top             =   930
            Width           =   720
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   405
         TabIndex        =   38
         Top             =   315
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   195
         Left            =   7305
         TabIndex        =   37
         Top             =   315
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   720
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5250
      Left            =   9180
      TabIndex        =   34
      Top             =   -60
      Width           =   1410
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   51
         Top             =   1410
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   29
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   30
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   31
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   32
         Top             =   1815
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   33
         Top             =   2220
         Width           =   1215
      End
      Begin MSComctlLib.ImageList imgClassificacao 
         Left            =   420
         Top             =   4590
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
               Picture         =   "frmCadastroCamara.frx":1A9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadastroCamara.frx":1D45
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCadastroCamara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrCodTipoServ()      As String
Private mstrDescTipoServ()     As String
Private mstrCodFormaLanc()     As String
Private mstrDescFormaLanc()    As String
Private mstrCodTipoMov()       As String
Private mstrDescTipoMov()      As String
Private mstrCodMovimento()     As String
Private mstrDescCodMovimento() As String
Private mstrCodOcorrencia()    As String
Private mstrDescOcorrencia()   As String
Private mblnAlteracao          As Boolean
Private mblnAlteraTipoServ     As Boolean
Private mblnAlteraFormaLanc    As Boolean
Private mblnAlteraTipoMov      As Boolean
Private mblnAlteraCodMov       As Boolean
Private mblnAlteraOcorrencia   As Boolean
Private mintIndexTipoServ      As Integer
Private mintIndexFormaLanc     As Integer
Private mintIndexTipoMov       As Integer
Private mintIndexCodMov        As Integer
Private mintIndexOcorrencia    As Integer
Private Const grdAsc = 1
Private Const grdDesc = 2

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdCancelaOcorrencia_Click()
    Call LimpaCampos(etxCodOcorrencia, etxDescOcorrencia)
    mblnAlteraOcorrencia = False
    mintIndexOcorrencia = 0
End Sub

Private Sub cmdCancelCodMovimento_Click()
    Call LimpaCampos(etxCodMovimento, etxDescCodMovimento)
    mblnAlteraCodMov = False
    mintIndexCodMov = 0
End Sub

Private Sub cmdCancelCodMovimento_LostFocus()
    If Screen.ActiveControl.Name = "etxCodOcorrencia" Then
        tabRegistros.Tab = 4
        etxCodOcorrencia.SetFocus
    End If
End Sub

Private Sub cmdCancelFormaLancamento_Click()
    Call LimpaCampos(etxCodFormaLancamento, etxDescFormaLancamento)
    mblnAlteraFormaLanc = False
    mintIndexFormaLanc = 0
End Sub

Private Sub cmdCancelFormaLancamento_LostFocus()
    If Screen.ActiveControl.Name = "etxCodTipoMovimento" Then
        tabRegistros.Tab = 2
        etxCodTipoMovimento.SetFocus
    End If
End Sub

Private Sub cmdCancelTipoMovimento_Click()
    Call LimpaCampos(etxCodTipoMovimento, etxDescTipoMovimento)
    mblnAlteraTipoMov = False
    mintIndexTipoMov = 0
End Sub

Private Sub cmdCancelTipoMovimento_LostFocus()
    If Screen.ActiveControl.Name = "etxCodMovimento" Then
        tabRegistros.Tab = 3
        etxCodMovimento.SetFocus
    End If
End Sub

Private Sub cmdCancelTipoServico_Click()
    Call LimpaCampos(etxCodTipoServico, etxDescTipoServico)
    mblnAlteraTipoServ = False
    mintIndexTipoServ = 0
End Sub

Private Sub cmdCancelTipoServico_LostFocus()
    If Screen.ActiveControl.Name = "etxCodFormaLancamento" Then
        tabRegistros.Tab = 1
        etxCodFormaLancamento.SetFocus
    End If
End Sub

Private Sub cmdConfCodMovimento_Click()
    If ConfirmaRegistro(grdCodigoMovimento, mblnAlteraCodMov, etxCodMovimento, etxDescCodMovimento, mintIndexCodMov, mstrCodMovimento, mstrDescCodMovimento) Then
        mintIndexCodMov = 0
        mblnAlteraCodMov = False
    End If
End Sub

Private Sub cmdConfFormaLancamento_Click()
    If ConfirmaRegistro(grdFormaLancamento, mblnAlteraFormaLanc, etxCodFormaLancamento, etxDescFormaLancamento, mintIndexFormaLanc, mstrCodFormaLanc, mstrDescFormaLanc) Then
        mintIndexFormaLanc = 0
        mblnAlteraFormaLanc = False
    End If
End Sub

Private Sub cmdConfOcorrencia_Click()
    If ConfirmaRegistro(grdOcorrencias, mblnAlteraOcorrencia, etxCodOcorrencia, etxDescOcorrencia, mintIndexOcorrencia, mstrCodOcorrencia, mstrDescOcorrencia) Then
        mintIndexOcorrencia = 0
        mblnAlteraOcorrencia = False
    End If
End Sub

Private Sub cmdConfTipoServico_Click()
    If ConfirmaRegistro(grdTipoServico, mblnAlteraTipoServ, etxCodTipoServico, etxDescTipoServico, mintIndexTipoServ, mstrCodTipoServ, mstrDescTipoServ) Then
        mintIndexTipoServ = 0
        mblnAlteraTipoServ = False
    End If
End Sub

Private Sub cmdConfTipoMovimento_Click()
    If ConfirmaRegistro(grdTipoMovimento, mblnAlteraTipoMov, etxCodTipoMovimento, etxDescTipoMovimento, mintIndexTipoMov, mstrCodTipoMov, mstrDescTipoMov) Then
        mintIndexTipoMov = 0
        mblnAlteraTipoMov = False
    End If
End Sub

Private Sub cmdExcluir_Click()
    If DeletaRegistro Then
        MsgBox "Registro excluído com sucesso.", vbInformation, NomeModulo
        Call NovoRegistro
        mblnAlteracao = False
    End If
End Sub

Private Sub cmdExcluirCodMovimento_Click()
    If ExcluiRegistro(mblnAlteraCodMov, mintIndexCodMov, mstrCodMovimento, mstrDescCodMovimento, grdCodigoMovimento, etxCodMovimento, etxDescCodMovimento) Then
        mintIndexCodMov = 0
        mblnAlteraCodMov = False
    End If
End Sub

Private Sub cmdExcluirFormaLancamento_Click()
    If ExcluiRegistro(mblnAlteraFormaLanc, mintIndexFormaLanc, mstrCodFormaLanc, mstrDescFormaLanc, grdFormaLancamento, etxCodFormaLancamento, etxDescFormaLancamento) Then
        mintIndexFormaLanc = 0
        mblnAlteraFormaLanc = False
    End If
End Sub

Private Sub cmdExcluirOcorrencia_Click()
    If ExcluiRegistro(mblnAlteraOcorrencia, mintIndexOcorrencia, mstrCodOcorrencia, mstrDescOcorrencia, grdOcorrencias, etxCodOcorrencia, etxCodOcorrencia) Then
        mintIndexOcorrencia = 0
        mblnAlteraOcorrencia = False
    End If
End Sub

Private Sub cmdExcluirTipoMovimento_Click()
    If ExcluiRegistro(mblnAlteraTipoMov, mintIndexTipoMov, mstrCodTipoMov, mstrDescTipoMov, grdTipoMovimento, etxCodTipoMovimento, etxCodTipoMovimento) Then
        mintIndexTipoMov = 0
        mblnAlteraTipoMov = False
    End If
End Sub

Private Sub cmdExcluirTipoServico_Click()
    If ExcluiRegistro(mblnAlteraTipoServ, mintIndexTipoServ, mstrCodTipoServ, mstrDescTipoServ, grdTipoServico, etxCodTipoServico, etxDescTipoServico) Then
        mintIndexTipoServ = 0
        mblnAlteraTipoServ = False
    End If
End Sub

Private Sub cmdGravar_Click()
    If SalvaRegistro Then
        MsgBox "Registro gravado com sucesso.", vbInformation, NomeModulo
        mblnAlteracao = True
        cmdExcluir.Enabled = True
    End If
End Sub

Private Sub cmdNovo_Click()
    Call NovoRegistro
    etxCodigo.SetFocus
End Sub

Private Sub cmdPesquisar_Click()
    Call PCampo("Câmaras", "SELECT cd_camara, desc_camara, status FROM FFICamaras", pbCampo, etxCodigo, "cd_camara")
    Call etxCodigo_LostFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub etxCodFormaLancamento_LostFocus()
    If Screen.ActiveControl.Name = "etxDescTipoServico" Then
        tabRegistros.Tab = 0
        cmdCancelTipoServico.SetFocus
    End If
End Sub

Private Sub etxCodigo_LostFocus()
    If etxCodigo.valorInteiro > 0 Then
        If ExisteRegistro Then
            Call MostraRegistro
            etxCodigo.Enabled = False
            cmdExcluir.Enabled = True
        End If
    End If
End Sub

Private Sub etxCodMovimento_LostFocus()
    If Screen.ActiveControl.Name = "etxDescTipoMovimento" Then
        tabRegistros.Tab = 2
        cmdCancelTipoMovimento.SetFocus
    End If
End Sub

Private Sub etxCodTipoMovimento_LostFocus()
    If Screen.ActiveControl.Name = "etxDescFormaLancamento" Then
        tabRegistros.Tab = 1
        cmdCancelFormaLancamento.SetFocus
    End If
End Sub

Private Sub etxCodOcorrencia_LostFocus()
    If Screen.ActiveControl.Name = "etxDescCodMovimento" Then
        tabRegistros.Tab = 3
        cmdCancelCodMovimento.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ecbStatus.preencher
    Call NovoRegistro
    imgOrderCodTipoServ.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderDescTipoServ.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderCodFormaLanc.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderDescFormaLanc.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderCodTipoMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderDescTipoMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderCodMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderDescCodMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderCodOcorrencia.Picture = imgClassificacao.ListImages(grdAsc).Picture
    imgOrderDescOcorrencia.Picture = imgClassificacao.ListImages(grdAsc).Picture
End Sub

Private Sub grdTipoServico_DblClick()
    With grdTipoServico
        If .TextMatrix(.Row, 1) <> "" Then
            mintIndexTipoServ = .Row - 1
            etxCodTipoServico.valorTexto = .TextMatrix(.Row, 1)
            etxDescTipoServico.valorTexto = .TextMatrix(.Row, 2)
            mblnAlteraTipoServ = True
        End If
    End With
End Sub

Private Sub grdFormaLancamento_DblClick()
    With grdFormaLancamento
        If .TextMatrix(.Row, 1) <> "" Then
            mintIndexFormaLanc = .Row - 1
            etxCodFormaLancamento.valorTexto = .TextMatrix(.Row, 1)
            etxDescFormaLancamento.valorTexto = .TextMatrix(.Row, 2)
            mblnAlteraFormaLanc = True
        End If
    End With
End Sub

Private Sub grdTipoMovimento_DblClick()
    With grdTipoMovimento
        If .TextMatrix(.Row, 1) <> "" Then
            mintIndexTipoMov = .Row - 1
            etxCodTipoMovimento.valorTexto = .TextMatrix(.Row, 1)
            etxDescTipoMovimento.valorTexto = .TextMatrix(.Row, 2)
            mblnAlteraTipoMov = True
        End If
    End With
End Sub

Private Sub grdCodigoMovimento_DblClick()
    With grdCodigoMovimento
        If .TextMatrix(.Row, 1) <> "" Then
            mintIndexCodMov = .Row - 1
            etxCodMovimento.valorTexto = .TextMatrix(.Row, 1)
            etxDescCodMovimento.valorTexto = .TextMatrix(.Row, 2)
            mblnAlteraCodMov = True
        End If
    End With
End Sub

Private Sub grdOcorrencias_DblClick()
    With grdOcorrencias
        If .TextMatrix(.Row, 1) <> "" Then
            mintIndexOcorrencia = .Row - 1
            etxCodOcorrencia.valorTexto = .TextMatrix(.Row, 1)
            etxDescOcorrencia.valorTexto = .TextMatrix(.Row, 2)
            mblnAlteraOcorrencia = True
        End If
    End With
End Sub

Private Sub imgOrderCodOcorrencia_Click()
    If imgOrderCodOcorrencia.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderCodOcorrencia.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(, , , , True)
        Call MostraRegistro("cd_ocorrencia_retorno DESC", False, False, False, False, True)
    Else
        imgOrderCodOcorrencia.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(, , , , True)
        Call MostraRegistro("cd_ocorrencia_retorno ASC", False, False, False, False, True)
    End If
End Sub

Private Sub imgOrderDescOcorrencia_Click()
    If imgOrderDescOcorrencia.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderDescOcorrencia.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(, , , , True)
        Call MostraRegistro("desc_ocorrencia_retorno DESC", False, False, False, False, True)
    Else
        imgOrderDescOcorrencia.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(, , , , True)
        Call MostraRegistro("desc_ocorrencia_retorno ASC", False, False, False, False, True)
    End If
End Sub

Private Sub imgOrderCodMovimento_Click()
    If imgOrderCodMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderCodMovimento.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(, , , True)
        Call MostraRegistro("cd_movimento DESC", False, False, False, True, False)
    Else
        imgOrderCodMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(, , , True)
        Call MostraRegistro("cd_movimento ASC", False, False, False, True, False)
    End If
End Sub

Private Sub imgOrderDescCodMovimento_Click()
    If imgOrderDescCodMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderDescCodMovimento.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(, , , True)
        Call MostraRegistro("desc_cd_movimento DESC", False, False, False, True, False)
    Else
        imgOrderDescCodMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(, , , True)
        Call MostraRegistro("desc_cd_movimento ASC", False, False, False, True, False)
    End If
End Sub

Private Sub imgOrderCodFormaLanc_Click()
    If imgOrderCodFormaLanc.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderCodFormaLanc.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(, True)
        Call MostraRegistro("cd_forma_lancamento DESC", False, True, False, False, False)
    Else
        imgOrderCodFormaLanc.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(, True)
        Call MostraRegistro("cd_forma_lancamento ASC", False, True, False, False, False)
    End If
End Sub

Private Sub imgOrderDescFormaLanc_Click()
    If imgOrderDescFormaLanc.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderDescFormaLanc.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(, True)
        Call MostraRegistro("desc_forma_lancamento DESC", False, True, False, False, False)
    Else
        imgOrderDescFormaLanc.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(, True)
        Call MostraRegistro("desc_forma_lancamento ASC", False, True, False, False, False)
    End If
End Sub

Private Sub imgOrderCodTipoServ_Click()
    If imgOrderCodTipoServ.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderCodTipoServ.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(True)
        Call MostraRegistro("cd_tipo_servico DESC", True, False, False, False, False)
    Else
        imgOrderCodTipoServ.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(True)
        Call MostraRegistro("cd_tipo_servico ASC", True, False, False, False, False)
    End If
End Sub

Private Sub imgOrderDescTipoServ_Click()
    If imgOrderDescTipoServ.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderDescTipoServ.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(True)
        Call MostraRegistro("desc_tipo_servico DESC", True, False, False, False, False)
    Else
        imgOrderDescTipoServ.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(True)
        Call MostraRegistro("desc_tipo_servico ASC", True, False, False, False, False)
    End If
End Sub

Private Sub imgOrderCodTipoMovimento_Click()
    If imgOrderCodTipoMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderCodTipoMovimento.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(, , True)
        Call MostraRegistro("cd_tipo_movimento DESC", False, False, True, False, False)
    Else
        imgOrderCodTipoMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(, , True)
        Call MostraRegistro("cd_tipo_movimento ASC", False, False, True, False, False)
    End If
End Sub

Private Sub imgOrderDescTipoMovimento_Click()
    If imgOrderDescTipoMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture Then
        imgOrderDescTipoMovimento.Picture = imgClassificacao.ListImages(grdDesc).Picture
        Call PreparaGrid(, , True)
        Call MostraRegistro("desc_tipo_movimento DESC", False, False, True, False, False)
    Else
        imgOrderDescTipoMovimento.Picture = imgClassificacao.ListImages(grdAsc).Picture
        Call PreparaGrid(, , True)
        Call MostraRegistro("desc_tipo_movimento ASC", False, False, True, False, False)
    End If
End Sub


Private Sub tabRegistros_LostFocus()
    If Screen.ActiveControl.Name = "etxCodTipoServico" Then
        Select Case tabRegistros.Tab
            Case 0
                etxCodTipoServico.SetFocus
            Case 1
                etxCodFormaLancamento.SetFocus
            Case 2
                etxCodTipoMovimento.SetFocus
            Case 3
                etxCodMovimento.SetFocus
            Case 4
                etxCodOcorrencia.SetFocus
        End Select
    End If
End Sub

'Data.......: 03/10/2008
'Autor......: Ivo Sousa
'Descrição..: Utilizado para realizar a preparação do grid, formatação de cabeçalho e colunas.
Private Sub PreparaGrid(Optional blnTipoServico As Boolean, Optional blnFormaLancamento As Boolean, Optional blnTipoMovimento As Boolean, Optional blnCodMovimento As Boolean, Optional blnOcorrencia As Boolean)
    Dim intIndex As Integer
    
    'Aba Tipo de Serviço
    If blnTipoServico Then
        With grdTipoServico
            .Cols = 3
            .FixedCols = 1
            .Rows = 2
            
            .RowHeight(0) = 320
            .TextMatrix(0, 0) = ""
            .ColWidth(0) = 150
            
            .TextMatrix(0, 1) = "        Código"
            .ColWidth(1) = 1100
            .ColAlignment(1) = flexAlignRightCenter
            
            .TextMatrix(0, 2) = "        Descrição"
            .ColWidth(2) = 7250
            .ColAlignment(2) = flexAlignLeftCenter
            
            For intIndex = 0 To .Cols - 1
                .TextMatrix(1, intIndex) = ""
            Next
        End With
    End If
    
    'Aba Forma de Lançamento
    If blnFormaLancamento Then
        With grdFormaLancamento
            .Cols = 3
            .FixedCols = 1
            .Rows = 2
            
            .RowHeight(0) = 320
            .TextMatrix(0, 0) = ""
            .ColWidth(0) = 150
            
            .TextMatrix(0, 1) = "        Código"
            .ColWidth(1) = 1100
            .ColAlignment(1) = flexAlignRightCenter
            
            .TextMatrix(0, 2) = "        Descrição"
            .ColWidth(2) = 7250
            .ColAlignment(2) = flexAlignLeftCenter
            
            For intIndex = 0 To .Cols - 1
                .TextMatrix(1, intIndex) = ""
            Next
        End With
    End If
    
    'Aba Tipo de Movimento
    If blnTipoMovimento Then
        With grdTipoMovimento
            .Cols = 3
            .FixedCols = 1
            .Rows = 2
            
            .RowHeight(0) = 320
            .TextMatrix(0, 0) = ""
            .ColWidth(0) = 150
            
            .ColAlignment(1) = flexAlignRightCenter
            .TextMatrix(0, 1) = "        Código"
            .ColWidth(1) = 1100
            
            .TextMatrix(0, 2) = "        Descrição"
            .ColWidth(2) = 7250
            .ColAlignment(2) = flexAlignLeftCenter
            
            For intIndex = 0 To .Cols - 1
                .TextMatrix(1, intIndex) = ""
            Next
        End With
    End If
    
    'Aba Código do Movimento
    If blnCodMovimento Then
        With grdCodigoMovimento
            .Cols = 3
            .FixedCols = 1
            .Rows = 2
            
            .RowHeight(0) = 320
            .TextMatrix(0, 0) = ""
            .ColWidth(0) = 150
            
            .TextMatrix(0, 1) = "        Código"
            .ColWidth(1) = 1100
            .ColAlignment(1) = flexAlignRightCenter
            
            .TextMatrix(0, 2) = "        Descrição"
            .ColWidth(2) = 7250
            .ColAlignment(2) = flexAlignLeftCenter
            
            For intIndex = 0 To .Cols - 1
                .TextMatrix(1, intIndex) = ""
            Next
        End With
    End If
    
    'Aba Ocorrência de Retorno
    If blnOcorrencia Then
        With grdOcorrencias
            .Cols = 3
            .FixedCols = 1
            .Rows = 2
            
            .RowHeight(0) = 320
            .TextMatrix(0, 0) = ""
            .ColWidth(0) = 150
            
            .TextMatrix(0, 1) = "        Código"
            .ColWidth(1) = 1100
            .ColAlignment(1) = flexAlignRightCenter
            
            .TextMatrix(0, 2) = "        Descrição"
            .ColWidth(2) = 7250
            .ColAlignment(2) = flexAlignLeftCenter
            
            For intIndex = 0 To .Cols - 1
                .TextMatrix(1, intIndex) = ""
            Next
        End With
    End If
End Sub

'Data.......: 06/10/2008
'Autor......: Ivo Sousa
'Descrição..: Utilizado para Limpar a tela para inserção de um novo registro
Private Sub NovoRegistro()
    Call PreparaGrid(True, True, True, True, True)
    etxCodigo.clear
    etxDescricao.clear
    etxCodTipoServico.clear
    etxDescTipoServico.clear
    etxCodFormaLancamento.clear
    etxDescFormaLancamento.clear
    etxCodTipoMovimento.clear
    etxDescTipoMovimento.clear
    etxCodMovimento.clear
    etxDescCodMovimento.clear
    etxCodOcorrencia.clear
    etxDescOcorrencia.clear
    ecbStatus.SelectItem "Ativo"
    tabRegistros.Tab = 0
    ReDim mstrDescFormaLanc(0)
    ReDim mstrDescCodMovimento(0)
    ReDim mstrDescTipoMov(0)
    ReDim mstrDescTipoServ(0)
    ReDim mstrCodFormaLanc(0)
    ReDim mstrCodMovimento(0)
    ReDim mstrCodTipoMov(0)
    ReDim mstrCodTipoServ(0)
    ReDim mstrCodOcorrencia(0)
    mblnAlteraTipoServ = False
    mblnAlteraOcorrencia = False
    mblnAlteraFormaLanc = False
    mblnAlteraTipoMov = False
    mblnAlteraCodMov = False
    mblnAlteracao = False
    etxCodigo.Enabled = True
    cmdExcluir.Enabled = False
End Sub

'Data.......: 06/10/2008
'Autor......: Ivo Sousa
'Descrição..: Utilizado para atualizar os registros das variaveis de acumulo e atualizar a grid
'Parametros.: [MSHFlexGrid] Grid de destino dos campos
'...........: [Boolean] Se o registro vem de uma alteração ou sera uma inserção
'...........: [EBSText] TextBox da tela referente ao codigo do registro
'...........: [EBSText] TextBox da tela referente a descrição do registro
'...........: [Integer] O index do registro, só para alterações
'...........: [String] Arreio com os códigos onde será inserido o novo registro
'...........: [String] Arreio com as descrições onde será inserido o novo registro
Private Function ConfirmaRegistro(grdFlexGrid As MSHFlexGrid, blnAlteracao As Boolean, txtCodigo As Object, txtDescricao As Object, intIndex As Integer, strCodigo() As String, strDescricao() As String) As Boolean
    Dim intRow As Integer
    
    If ValidaRegistro(txtCodigo, txtDescricao) Then
        If Not blnAlteracao Then
            'Se não for o primeiro registro, é criado próximo indice.
            If strCodigo(0) <> "" Then
                intIndex = UBound(strCodigo) + 1
                ReDim Preserve strCodigo(intIndex)
                ReDim Preserve strDescricao(intIndex)
                intRow = grdFlexGrid.Rows
            Else
                intRow = grdFlexGrid.Rows - 1
            End If
        Else
            intRow = intIndex + 1
        End If
        strCodigo(intIndex) = txtCodigo.valorTexto
        strDescricao(intIndex) = txtDescricao.valorTexto
        With grdFlexGrid
            If .TextMatrix(1, 1) <> "" And Not blnAlteracao Then
                .AddItem ""
            End If
            .TextMatrix(intRow, 1) = txtCodigo.valorTexto
            .TextMatrix(intRow, 2) = txtDescricao.valorTexto
        End With
        Call LimpaCampos(txtCodigo, txtDescricao)
        ConfirmaRegistro = True
    Else
        ConfirmaRegistro = False
    End If
End Function

Private Function ExcluiRegistro(blnAlteracao As Boolean, intIndex As Integer, strCodigo() As String, strDescricao() As String, grdRegistros As MSHFlexGrid, txtCodigo As EBSText, txtDescricao As EBSText) As Boolean
    Dim intCont        As Integer
    Dim intRegistro    As Integer
    Dim intRegistroAnt As Integer
    
    intCont = UBound(strCodigo)
    If Not blnAlteracao Then
        MsgBox "Selecione um registro antes de tentar excluir.", vbInformation, NomeModulo
        Exit Function
    Else
        strCodigo(intIndex) = 0
        strDescricao(intIndex) = ""
        intRegistroAnt = intIndex
        If intIndex < intCont Then
            For intRegistro = intIndex + 1 To intCont
                strCodigo(intRegistroAnt) = strCodigo(intRegistro)
                strDescricao(intRegistroAnt) = strDescricao(intRegistro)
                intRegistroAnt = intRegistro
            Next intRegistro
        End If
        If intCont > 0 Then
            ReDim Preserve strCodigo(intCont - 1)
            ReDim Preserve strDescricao(intCont - 1)
        Else
            ReDim strCodigo(intCont)
            ReDim strDescricao(intCont)
        End If
        With grdRegistros
            If .Rows = 2 Then
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
            Else
                .RemoveItem (intIndex + 1)
            End If
        End With
        Call LimpaCampos(txtCodigo, txtDescricao)
    End If
    ExcluiRegistro = True
End Function

Private Function ValidaRegistro(txtCodigo As Object, txtDescricao As Object)
    If Trim(txtCodigo.valorTexto) = "" Then
        MsgBox "O campo Código é de preenchimento obrigatório.", vbInformation, NomeModulo
        txtCodigo.SetFocus
        Exit Function
    End If
    If Trim(txtDescricao.valorTexto) = "" Then
        MsgBox "O campo Descrição é de preenchimento obrigatório.", vbInformation, NomeModulo
        txtDescricao.SetFocus
        Exit Function
    End If
    ValidaRegistro = True
End Function

Private Sub LimpaCampos(txtCodigo As EBSText, txtDescricao As EBSText)
    txtCodigo.clear
    txtDescricao.clear
End Sub

Private Function SalvaRegistro() As Boolean
    Dim strSql     As String
    Dim strCampos  As String
    Dim strValores As String
    Dim intCont    As Integer
    
    If ValidaCampos Then
        BeginTrans
        If Not mblnAlteracao Then
            strSql = "INSERT INTO FFICamaras (cd_camara, desc_camara, status) "
            strSql = strSql & "VALUES (" & etxCodigo.valorInteiro & ", '" & etxDescricao.valorTexto & "', '" & ecbStatus.SelectedItem & "')"
        Else
            Call ExecuteSQL("DELETE FROM FFICamaraFormaLancamento WHERE cd_camara = " & etxCodigo.valorInteiro)
            Call ExecuteSQL("DELETE FROM FFICamaraCodigoMovimento WHERE cd_camara = " & etxCodigo.valorInteiro)
            Call ExecuteSQL("DELETE FROM FFICamaraTipoMovimento WHERE cd_camara = " & etxCodigo.valorInteiro)
            Call ExecuteSQL("DELETE FROM FFICamaraTipoServico WHERE cd_camara = " & etxCodigo.valorInteiro)
            Call ExecuteSQL("DELETE FROM FFICamaraOcorrenciasRetorno WHERE cd_camara = " & etxCodigo.valorInteiro)
            strSql = "UPDATE FFICamaras SET desc_camara = '" & etxDescricao.valorTexto & "', status = '" & ecbStatus.SelectedItem & "' WHERE cd_camara = " & etxCodigo.valorInteiro
        End If
        If ExecuteSQL(strSql) > 0 Then
            If Trim(mstrCodFormaLanc(0)) <> "" Then
                For intCont = 0 To UBound(mstrCodFormaLanc)
                    strSql = "INSERT INTO FFICamaraFormaLancamento (cd_camara, cd_forma_lancamento, desc_forma_lancamento) "
                    strSql = strSql & " VALUES (" & etxCodigo.valorInteiro & ", '" & mstrCodFormaLanc(intCont) & "', '" & mstrDescFormaLanc(intCont) & "')"
                    Call ExecuteSQL(strSql)
                Next
            End If
            If Trim(mstrCodMovimento(0)) <> "" Then
                For intCont = 0 To UBound(mstrCodMovimento)
                    strSql = "INSERT INTO FFICamaraCodigoMovimento (cd_camara, cd_movimento, desc_cd_movimento) "
                    strSql = strSql & " VALUES (" & etxCodigo.valorInteiro & ", '" & mstrCodMovimento(intCont) & "', '" & mstrDescCodMovimento(intCont) & "')"
                    Call ExecuteSQL(strSql)
                Next
            End If
            If Trim(mstrCodTipoMov(0)) <> "" Then
                For intCont = 0 To UBound(mstrCodTipoMov)
                    strSql = "INSERT INTO FFICamaraTipoMovimento (cd_camara, cd_tipo_movimento, desc_tipo_movimento) "
                    strSql = strSql & " VALUES (" & etxCodigo.valorInteiro & ", '" & mstrCodTipoMov(intCont) & "', '" & mstrDescTipoMov(intCont) & "')"
                    Call ExecuteSQL(strSql)
                Next
            End If
            If Trim(mstrCodTipoServ(0)) <> "" Then
                For intCont = 0 To UBound(mstrCodTipoServ)
                    strSql = "INSERT INTO FFICamaraTipoServico (cd_camara, cd_tipo_servico, desc_tipo_servico) "
                    strSql = strSql & " VALUES (" & etxCodigo.valorInteiro & ", '" & mstrCodTipoServ(intCont) & "', '" & mstrDescTipoServ(intCont) & "')"
                    Call ExecuteSQL(strSql)
                Next
            End If
            If Trim(mstrCodOcorrencia(0)) <> "" Then
                For intCont = 0 To UBound(mstrCodOcorrencia)
                    strSql = "INSERT INTO FFICamaraOcorrenciasRetorno (cd_camara, cd_ocorrencia_retorno, desc_ocorrencia_retorno) "
                    strSql = strSql & " VALUES (" & etxCodigo.valorInteiro & ", '" & mstrCodOcorrencia(intCont) & "', '" & mstrDescOcorrencia(intCont) & "')"
                    Call ExecuteSQL(strSql)
                Next
            End If
        Else
            MsgBox "Não foi possível gravar o registro.", vbInformation, NomeModulo
            Rollback
            Exit Function
        End If
        CommitTrans
        SalvaRegistro = True
    End If
    Exit Function
    
Erro_Gravacao:
    MsgBox "Erro ao gravar o registro:" & err.Description & ".", vbError, NomeModulo
    Rollback
End Function

Private Function ValidaCampos() As Boolean
    If Not etxCodigo.valorInteiro > 0 Then
        MsgBox "Informe o código da câmara antes de salvar.", vbInformation, NomeModulo
        etxCodigo.SetFocus
        Exit Function
    End If
    If Trim(etxDescricao.valorTexto) = "" Then
        MsgBox "Informe a descrição da câmara antes de salvar.", vbInformation, NomeModulo
        etxDescricao.SetFocus
        Exit Function
    End If
    ValidaCampos = True
End Function

Private Function ExisteRegistro() As Boolean
    Dim strSql As String
    Dim rstResult As Object
    
    strSql = "SELECT * FROM FFICamaras WHERE cd_camara = " & etxCodigo.valorInteiro
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        ExisteRegistro = True
    End If
End Function


Private Sub MostraRegistro(Optional ByRef strORDERBY As String, Optional blnTipoServico As Boolean = True, Optional blnFormaLancamento As Boolean = True, Optional blnTipoMovimento As Boolean = True, Optional blnCodMovimento As Boolean = True, Optional blnOcorrencia As Boolean = True)
    Dim strSql             As String
    Dim rstCamara          As Object
    Dim rstTipoServico     As Object
    Dim rstFormaLancamento As Object
    Dim rstCodMovimento    As Object
    Dim rstTipoMovimento   As Object
    Dim rstOcorrencias     As Object
    Dim intCont            As Integer
    
    strSql = "SELECT * FROM FFICamaras WHERE cd_camara = " & etxCodigo.valorInteiro
    
    If AbreRecordset(rstCamara, strSql) = WL_OK Then
        etxCodigo.valorInteiro = rstCamara("cd_camara").Value
        etxDescricao.valorTexto = rstCamara("desc_camara").Value
        ecbStatus.SelectItem rstCamara("status").Value
        mblnAlteracao = True
    End If
    If strORDERBY <> "" Then
        strSql = strSql & " ORDER BY " & strORDERBY
    End If
    'Código do Movimento
    If blnCodMovimento Then
        If AbreRecordset(rstCodMovimento, Replace(strSql, "FFICamaras", "FFICamaraCodigoMovimento")) = WL_OK Then
            intCont = 0
            While Not rstCodMovimento.EOF
                ReDim Preserve mstrCodMovimento(intCont)
                ReDim Preserve mstrDescCodMovimento(intCont)
                mstrCodMovimento(intCont) = rstCodMovimento("cd_movimento").Value
                mstrDescCodMovimento(intCont) = rstCodMovimento("desc_cd_movimento").Value
                With grdCodigoMovimento
                    If intCont > 0 Then
                        .AddItem ""
                    End If
                    .TextMatrix(.Rows - 1, 1) = rstCodMovimento("cd_movimento").Value
                    .TextMatrix(.Rows - 1, 2) = rstCodMovimento("desc_cd_movimento").Value
                End With
                rstCodMovimento.MoveNext
                intCont = intCont + 1
            Wend
        End If
    End If
    
    'Forma de Lançamento
    If blnFormaLancamento Then
        If AbreRecordset(rstFormaLancamento, Replace(strSql, "FFICamaras", "FFICamaraFormaLancamento")) = WL_OK Then
            intCont = 0
            While Not rstFormaLancamento.EOF
                ReDim Preserve mstrCodFormaLanc(intCont)
                ReDim Preserve mstrDescFormaLanc(intCont)
                mstrCodFormaLanc(intCont) = rstFormaLancamento("cd_forma_lancamento").Value
                mstrDescFormaLanc(intCont) = rstFormaLancamento("desc_forma_lancamento").Value
                With grdFormaLancamento
                    If intCont > 0 Then
                        .AddItem ""
                    End If
                    .TextMatrix(.Rows - 1, 1) = rstFormaLancamento("cd_forma_lancamento").Value
                    .TextMatrix(.Rows - 1, 2) = rstFormaLancamento("desc_forma_lancamento").Value
                End With
                rstFormaLancamento.MoveNext
                intCont = intCont + 1
            Wend
        End If
    End If
    
    'Tipo de Movimento
    If blnTipoMovimento Then
        If AbreRecordset(rstTipoMovimento, Replace(strSql, "FFICamaras", "FFICamaraTipoMovimento")) = WL_OK Then
            intCont = 0
            While Not rstTipoMovimento.EOF
                ReDim Preserve mstrCodTipoMov(intCont)
                ReDim Preserve mstrDescTipoMov(intCont)
                mstrCodTipoMov(intCont) = rstTipoMovimento("cd_tipo_movimento").Value
                mstrDescTipoMov(intCont) = rstTipoMovimento("desc_tipo_movimento").Value
                With grdTipoMovimento
                    If intCont > 0 Then
                        .AddItem ""
                    End If
                    .TextMatrix(.Rows - 1, 1) = rstTipoMovimento("cd_tipo_movimento").Value
                    .TextMatrix(.Rows - 1, 2) = rstTipoMovimento("desc_tipo_movimento").Value
                End With
                rstTipoMovimento.MoveNext
                intCont = intCont + 1
            Wend
        End If
    End If
    
    'Tipo de Serviço
    If blnTipoServico Then
        If AbreRecordset(rstTipoServico, Replace(strSql, "FFICamaras", "FFICamaraTipoServico")) = WL_OK Then
            intCont = 0
            While Not rstTipoServico.EOF
                ReDim Preserve mstrCodTipoServ(intCont)
                ReDim Preserve mstrDescTipoServ(intCont)
                mstrCodTipoServ(intCont) = rstTipoServico("cd_tipo_servico").Value
                mstrDescTipoServ(intCont) = rstTipoServico("desc_tipo_servico").Value
                With grdTipoServico
                    If intCont > 0 Then
                        .AddItem ""
                    End If
                    .TextMatrix(.Rows - 1, 1) = rstTipoServico("cd_tipo_servico").Value
                    .TextMatrix(.Rows - 1, 2) = rstTipoServico("desc_tipo_servico").Value
                End With
                rstTipoServico.MoveNext
                intCont = intCont + 1
            Wend
        End If
    End If
    
    'Ocorrências de Retorno
    If blnOcorrencia Then
        If AbreRecordset(rstOcorrencias, Replace(strSql, "FFICamaras", "FFICamaraOcorrenciasRetorno")) = WL_OK Then
            intCont = 0
            While Not rstOcorrencias.EOF
                ReDim Preserve mstrCodOcorrencia(intCont)
                ReDim Preserve mstrDescOcorrencia(intCont)
                mstrCodOcorrencia(intCont) = rstOcorrencias("cd_ocorrencia_retorno").Value
                mstrDescOcorrencia(intCont) = rstOcorrencias("desc_ocorrencia_retorno").Value
                With grdOcorrencias
                    If intCont > 0 Then
                        .AddItem ""
                    End If
                    .TextMatrix(.Rows - 1, 1) = rstOcorrencias("cd_ocorrencia_retorno").Value
                    .TextMatrix(.Rows - 1, 2) = rstOcorrencias("desc_ocorrencia_retorno").Value
                End With
                rstOcorrencias.MoveNext
                intCont = intCont + 1
            Wend
        End If
    End If
End Sub

Private Function DeletaRegistro() As Boolean
    Dim strSql As String

On Error GoTo Erro_Deletando
    If mblnAlteracao Then
        
        BeginTrans
        strSql = "DELETE FROM FFICamaras WHERE cd_camara = " & etxCodigo.valorInteiro
        If ExecuteSQL(strSql) > 0 Then
            Call ExecuteSQL(Replace(strSql, "FFICamaras", "FFICamaraCodigoMovimento"))
            Call ExecuteSQL(Replace(strSql, "FFICamaras", "FFICamaraFormaLancamento"))
            Call ExecuteSQL(Replace(strSql, "FFICamaras", "FFICamaraTipoMovimento"))
            Call ExecuteSQL(Replace(strSql, "FFICamaras", "FFICamaraTipoServico"))
        End If
        CommitTrans
    Else
        MsgBox "Escolha um registro antes de tentar excluír.", vbInformation, NomeModulo
        Exit Function
    End If
    DeletaRegistro = True
    Exit Function
Erro_Deletando:
    MsgBox "Erro excluindo o registro: " & err.Description & ".", vbInformation, NomeModulo
    DeletaRegistro = False
    Rollback
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
