VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmGeracaoTitulosPagar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Títulos Pagar"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   8505
   Tag             =   "FFITituloPagar"
   Begin TabDlg.SSTab tabTitulos 
      Height          =   4050
      Left            =   90
      TabIndex        =   34
      Top             =   1050
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   7144
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nota Fiscal"
      TabPicture(0)   =   "GeracaoTitulosPagar.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "etxValorNota"
      Tab(0).Control(1)=   "etxEmpresa"
      Tab(0).Control(2)=   "ecbTipoRegistro"
      Tab(0).Control(3)=   "etxNumeroNota"
      Tab(0).Control(4)=   "edtDataEmissao"
      Tab(0).Control(5)=   "etxParcela"
      Tab(0).Control(6)=   "Label3"
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(11)=   "Label7"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Duplicata"
      TabPicture(1)   =   "GeracaoTitulosPagar.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraData"
      Tab(1).Control(1)=   "cmdRateio"
      Tab(1).Control(2)=   "etxCodigoBanco"
      Tab(1).Control(3)=   "etxCodigoConta"
      Tab(1).Control(4)=   "etxCentroCusto"
      Tab(1).Control(5)=   "etxMoeda"
      Tab(1).Control(6)=   "etxOperacaoContabil"
      Tab(1).Control(7)=   "etxIntervalo"
      Tab(1).Control(8)=   "Label11"
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(10)=   "Label9"
      Tab(1).Control(11)=   "Label8"
      Tab(1).Control(12)=   "Label13"
      Tab(1).Control(13)=   "Label15"
      Tab(1).Control(14)=   "imgInformativa"
      Tab(1).Control(15)=   "Label14"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Financeiro"
      TabPicture(2)   =   "GeracaoTitulosPagar.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "etxCentroFinan"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "etxParcelaFinan"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "etxNotaFinan"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "etxValorFinan"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "edtVencimento"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "grdTitFin"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdAlterar"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "&Alterar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4935
         TabIndex        =   20
         Top             =   750
         Width           =   1240
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTitFin 
         Height          =   2805
         Left            =   60
         TabIndex        =   36
         Top             =   1185
         Width           =   6690
         _ExtentX        =   11800
         _ExtentY        =   4948
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame fraData 
         Caption         =   "Regra Exceção"
         Height          =   480
         Left            =   -71580
         TabIndex        =   37
         Top             =   480
         Width           =   2580
         Begin VB.OptionButton optAnterior 
            Caption         =   "Antecipar"
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optProximo 
            Caption         =   "Prorrogar"
            Height          =   195
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdRateio 
         Caption         =   "&Rateio"
         Height          =   375
         Left            =   -69555
         TabIndex        =   16
         Top             =   2520
         Width           =   1215
      End
      Begin Fox.EBSText etxValorNota 
         Height          =   330
         Left            =   -73575
         TabIndex        =   5
         Top             =   1845
         Width           =   1335
         _extentx        =   265
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":0054
         tipo            =   1
         casasdecimais   =   2
         tipotexto       =   0
         maxlength       =   18
         tipocriterio    =   6
         alinhamento     =   1
         mascara         =   "##,##0.00"
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSText etxEmpresa 
         Height          =   330
         Left            =   -73575
         TabIndex        =   4
         Top             =   1440
         Width           =   5025
         _extentx        =   439632
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":0080
         tipo            =   4
         tipotexto       =   0
         maxlength       =   15
         possuidescricao =   -1  'True
         campocriterio   =   "Apel"
         campodescricao  =   "Razão"
         tabelaconsulta  =   "Empresas"
         tamanhodescricao=   3500
      End
      Begin Fox.EBSCombo ecbTipoRegistro 
         Height          =   315
         Left            =   -73575
         TabIndex        =   3
         Tag             =   "GerTitPagar"
         Top             =   1035
         Width           =   1860
         _extentx        =   3281
         _extenty        =   556
         dados           =   ""
         dadosassist     =   ""
         font            =   "GeracaoTitulosPagar.frx":00AC
      End
      Begin Fox.EBSText etxNumeroNota 
         Height          =   330
         Left            =   -73575
         TabIndex        =   2
         Top             =   630
         Width           =   735
         _extentx        =   265
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":00D8
         tipotexto       =   0
         maxlength       =   6
         tipocriterio    =   4
         alinhamento     =   1
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSData edtDataEmissao 
         Height          =   330
         Left            =   -73575
         TabIndex        =   6
         Top             =   2250
         Width           =   1275
         _extentx        =   2249
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":0104
      End
      Begin Fox.EBSText etxParcela 
         Height          =   330
         Left            =   -73575
         TabIndex        =   7
         Top             =   2655
         Width           =   585
         _extentx        =   265
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":0130
         tipotexto       =   0
         maxlength       =   3
         tipocriterio    =   4
         alinhamento     =   1
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSText etxCodigoBanco 
         Height          =   330
         Left            =   -73275
         TabIndex        =   11
         Top             =   1020
         Width           =   4755
         _extentx        =   440161
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":015C
         tipotexto       =   0
         maxlength       =   9
         possuidescricao =   -1  'True
         campocriterio   =   "Banco"
         tipocriterio    =   4
         campodescricao  =   "Nome"
         tabelaconsulta  =   "Bancos"
         tamanhodescricao=   3800
         alinhamento     =   1
      End
      Begin Fox.EBSText etxCodigoConta 
         Height          =   330
         Left            =   -73275
         TabIndex        =   12
         Top             =   1425
         Width           =   4785
         _extentx        =   440161
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":0188
         tipotexto       =   0
         maxlength       =   9
         possuidescricao =   -1  'True
         campocriterio   =   "Código"
         tipocriterio    =   4
         campodescricao  =   "Descrição"
         tabelaconsulta  =   "Contas"
         tamanhodescricao=   3800
         alinhamento     =   1
      End
      Begin Fox.EBSText etxCentroCusto 
         Height          =   330
         Left            =   -73275
         TabIndex        =   13
         Top             =   1830
         Width           =   4725
         _extentx        =   440002
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":01B4
         tipotexto       =   0
         maxlength       =   9
         possuidescricao =   -1  'True
         campocriterio   =   "Código"
         tipocriterio    =   4
         campodescricao  =   "Descrição"
         tabelaconsulta  =   "Centros"
         tamanhodescricao=   3700
         alinhamento     =   1
      End
      Begin Fox.EBSText etxMoeda 
         Height          =   330
         Left            =   -73275
         TabIndex        =   15
         Top             =   2640
         Width           =   3615
         _extentx        =   437885
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":01E0
         tipo            =   4
         tipotexto       =   0
         maxlength       =   10
         possuidescricao =   -1  'True
         campocriterio   =   "Moeda"
         campodescricao  =   "Moeda"
         tabelaconsulta  =   "Moedas"
         tamanhodescricao=   2500
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSText etxOperacaoContabil 
         Height          =   330
         Left            =   -73275
         TabIndex        =   14
         Top             =   2220
         Width           =   4755
         _extentx        =   440531
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":020C
         tipotexto       =   0
         maxlength       =   5
         possuidescricao =   -1  'True
         campocriterio   =   "cd_operacao"
         tipocriterio    =   4
         campodescricao  =   "descricao"
         tabelaconsulta  =   "OperacaoContabil"
         tamanhodescricao=   4000
         alinhamento     =   1
      End
      Begin Fox.EBSText etxIntervalo 
         Height          =   330
         Left            =   -73275
         TabIndex        =   8
         Top             =   630
         Width           =   555
         _extentx        =   265
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":0238
         tipotexto       =   0
         maxlength       =   3
         tipocriterio    =   0
         alinhamento     =   1
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSData edtVencimento 
         Height          =   330
         Left            =   1725
         TabIndex        =   19
         Top             =   750
         Width           =   2175
         _extentx        =   208889
         _extenty        =   582
         caption         =   "Vencimento"
         font            =   "GeracaoTitulosPagar.frx":0264
      End
      Begin Fox.EBSText etxValorFinan 
         Height          =   330
         Left            =   3855
         TabIndex        =   17
         Top             =   360
         Width           =   1935
         _extentx        =   92710
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":0290
         tipo            =   1
         casasdecimais   =   2
         tipotexto       =   0
         maxlength       =   18
         caption         =   "Valor"
         tipocriterio    =   6
         alinhamento     =   1
         mascara         =   "##,##0.00"
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSText etxNotaFinan 
         Height          =   330
         Left            =   495
         TabIndex        =   52
         Top             =   390
         Width           =   1275
         _extentx        =   88239
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":02BC
         tipotexto       =   0
         maxlength       =   6
         caption         =   "Nota"
         enabled         =   0   'False
         tipocriterio    =   4
         alinhamento     =   1
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSText etxParcelaFinan 
         Height          =   330
         Left            =   2025
         TabIndex        =   53
         Top             =   390
         Width           =   1590
         _extentx        =   131789
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":02E8
         tipotexto       =   0
         maxlength       =   6
         caption         =   "Parcela"
         enabled         =   0   'False
         tipocriterio    =   4
         alinhamento     =   1
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSText etxCentroFinan 
         Height          =   330
         Left            =   240
         TabIndex        =   18
         Top             =   750
         Width           =   1380
         _extentx        =   154570
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":0314
         tipotexto       =   0
         maxlength       =   9
         caption         =   "C. Custo"
         possuidescricao =   -1  'True
         campocriterio   =   "Código"
         tipocriterio    =   4
         campodescricao  =   "Descrição"
         tabelaconsulta  =   "Centros"
         alinhamento     =   1
         exibedescricao  =   0   'False
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "M&oeda"
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
         Left            =   -74820
         TabIndex        =   51
         Top             =   2685
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Centro C&usto"
         Height          =   195
         Left            =   -74820
         TabIndex        =   50
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Conta &Financeira"
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
         Left            =   -74820
         TabIndex        =   49
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Banco"
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
         Left            =   -74820
         TabIndex        =   48
         Top             =   1065
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Op.&Contabil"
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
         Left            =   -74820
         TabIndex        =   47
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Intervalo (Dias)"
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
         Left            =   -74820
         TabIndex        =   46
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Va&lor Total"
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
         Left            =   -74610
         TabIndex        =   45
         Top             =   1890
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Em&presa"
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
         Left            =   -74610
         TabIndex        =   44
         Top             =   1485
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Tipo"
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
         Left            =   -74610
         TabIndex        =   43
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nú&mero"
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
         Left            =   -74610
         TabIndex        =   42
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Em&issão"
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
         Left            =   -74610
         TabIndex        =   41
         Top             =   2295
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pa&rcelas"
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
         Left            =   -74610
         TabIndex        =   40
         Top             =   2700
         Width           =   945
      End
      Begin VB.Image imgInformativa 
         Height          =   480
         Left            =   -74880
         Picture         =   "GeracaoTitulosPagar.frx":0340
         Top             =   3375
         Width           =   480
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFFF&
         Caption         =   $"GeracaoTitulosPagar.frx":0F82
         Height          =   585
         Left            =   -74940
         TabIndex        =   38
         Top             =   3330
         Width           =   6720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5475
      Left            =   7065
      TabIndex        =   33
      Top             =   -45
      Width           =   1410
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "A&juda"
         Height          =   375
         Left            =   90
         TabIndex        =   28
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   26
         Top             =   2190
         Width           =   1215
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   27
         Top             =   2600
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluirDuplicatas 
         Caption         =   "Exc.Duplicatas"
         Height          =   375
         Left            =   90
         TabIndex        =   25
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   29
         Top             =   3400
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   24
         Top             =   1395
         Width           =   1215
      End
      Begin VB.CommandButton cmdGerarDuplicatas 
         Caption         =   "&Calcular"
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   21
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5475
      Left            =   0
      TabIndex        =   32
      Top             =   -45
      Width           =   7035
      Begin Fox.EBSText etxCodigoTitulo 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   225
         Width           =   1215
         _extentx        =   265
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":106C
         tipotexto       =   0
         maxlength       =   9
         tipocriterio    =   0
         alinhamento     =   1
         exibedescricao  =   0   'False
      End
      Begin Fox.EBSText etxDescricao 
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   630
         Width           =   5775
         _extentx        =   2090
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":1098
         tipo            =   4
         tipotexto       =   0
         maxlength       =   50
         exibedescricao  =   0   'False
      End
      Begin ComctlLib.ProgressBar pgrTitulo 
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   5160
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   1
      End
      Begin Fox.EBSText etxStatus 
         Height          =   330
         Left            =   4350
         TabIndex        =   39
         Top             =   210
         Width           =   1920
         _extentx        =   62918
         _extenty        =   582
         font            =   "GeracaoTitulosPagar.frx":10C4
         tipo            =   4
         tipotexto       =   0
         caption         =   "Situação"
         enabled         =   0   'False
         exibedescricao  =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "De&scrição"
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
         Left            =   105
         TabIndex        =   31
         Top             =   675
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Có&digo"
         Height          =   195
         Left            =   450
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmGeracaoTitulosPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'pt. 85684 - Moacir Pfau(02/07/2008) - GERAÇÃO DE TÍTULOS
'Objeto de navegação
Private navigator          As New cGeracaoTituloPagarNavigator
'Variavel que guarda o indice do item selecionado na lista
Private lngItem            As Long
'OBJETO PRINCIPAL.
Private objGerTitPagar     As New cGeracaoTituloPagar
'OBJETO QUE GUARDA AS DUPLICATAS.
Private objCGerTitPagar    As New cGeracaoDuplicataPagar
'OBJETO QUE GUARDA O RATEIO
Private objRateioTitPagar  As New cGeracaoTituloPagar
Private objFinanPagar      As New cGeracaoDuplicataPagar
'Variavel que define se o foco está no valor
Private booValorFoco       As Boolean
'Variavel que define se o objeto está em alteração
Private booAlterando       As Boolean
Private mobjRateio         As cGeracaoTituloPagar
Private mobjHelp           As New clsHelp
Private blnGerouDuplicatas As Boolean

'CONFIGURAÇÃO DO GRID
Private Const strTituloGrid$ = "campo=P_PagRec;label=PagRec;tamanho=100|" & _
                    "campo=P_Nota;label=Nota;tamanho=800|" & _
                    "campo=P_Empresa;label=Empresa;tamanho=1|" & _
                    "campo=P_Tipo;label=Tipo;tamanho=1|" & _
                    "campo=P_Parcela;label=Parcela;tamanho=900|" & _
                    "campo=P_Descricao;label=Descricao;tamanho=1|" & _
                    "campo=P_Valor_Original;label=Valor;tamanho=1500;formato=###,##0.00|" & _
                    "campo=P_Banco;label=Banco;tamanho=1|" & _
                    "campo=P_Conta;label=Conta;tamanho=1|" & _
                    "campo=P_Centro;label=Centro;tamanho=800|" & _
                    "campo=P_cd_operacao_contabil;label=OpContabil;tamanho=1|" & _
                    "campo=P_Moeda;label=Moeda;tamanho=1|" & _
                    "campo=P_Vencimento;label=Vencimento;tamanho=1200|" & _
                    "campo=P_Pagamento;label=Pagamento;tamanho=1|" & _
                    "campo=P_Emissao;label=Emissao;tamanho=1"

Private Sub cmdAjuda_Click()
    Call LibProc(WL_AJUDA)
End Sub

Private Sub cmdExcluir_Click()
    Call LibProc(WL_DELETAR)
End Sub

Private Sub cmdExcluirDuplicatas_Click()
    fExclusaoDuplicatas
End Sub

Private Sub cmdRateio_LostFocus()
    'pt. 85684 - Ivo Sousa(14/07/2008)
    If Screen.ActiveControl.name = "etxValorFinan" Then
        tabTitulos.Tab = 2
    End If
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: GerarDuplicatas
'Data.......: 02/07/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para GERAR AS DUPLICATAS E LANÇAR NA COLEÇÃO.
'Protocolo..: 85684
'---------------------------------------------------------------------------------------
Private Sub GerarDuplicatas()
    Dim strSql           As String
    Dim rstTab           As Object
    Dim i                As Integer
    Dim j                As Integer
    Dim strMensagem      As String
    Dim CurrentObject    As cGeracaoTituloPagar
    Dim intTotalParc     As Integer
    Dim dtaData          As Date
    Dim GerTitPagar      As New cGeracaoDuplicataPagar
    Dim intParcela       As Integer
    Dim curTotal         As Currency
    Dim intTotalParcelas As Integer
    Dim curPerc             As Currency
    
    Set objCGerTitPagar = Nothing
    Set objCGerTitPagar = New cGeracaoDuplicataPagar
    Call preencheClasse

    strMensagem = ""
    intTotalParc = 1
    If objRateioTitPagar.Rateio.Count = 0 Then
        intTotalParc = objGerTitPagar.nr_parcela
    Else
        intTotalParc = objGerTitPagar.nr_parcela * objRateioTitPagar.Rateio.Count
    End If
    
    'Verifica se existe parcelas geradas.
    For i = 1 To intTotalParc
        strSql = "SELECT * FROM Duplicatas WHERE PagRec='P' AND Nota=" & etxNumeroNota.valorInteiro & " AND Empresa='" & etxEmpresa.valorTexto & "' AND Tipo='" & ecbTipoRegistro.SelectedItem & "' AND Parcela=" & i
        If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
            strMensagem = strMensagem & "-Nota: " & GetValue(rstTab, "Nota") & ", Parcela=" & GetValue(rstTab, "Parcela") & vbCrLf
        End If
        FechaRecordset (rstTab)
    Next
    
    If strMensagem <> "" Then
        MsgBox "Existe duplicatas geradas, não será possivel continuar." & vbCrLf & strMensagem, vbInformation, NomeModulo
        Exit Sub
    End If

    pgrTitulo.Min = 0
    pgrTitulo.value = 0
    pgrTitulo.Max = intTotalParc
    
    'INICIO PARA LANÇAR NA COLEÇÃO.
    j = 1
    intParcela = 1
    If objRateioTitPagar.Rateio.Count > 0 Then
        objRateioTitPagar.Rateio.MoveFirst
        curTotal = 0
        intTotalParcelas = objRateioTitPagar.Rateio.Count * objGerTitPagar.nr_parcela
        'SE EXISTIR O RATEIO.
        While Not objRateioTitPagar.Rateio.EOF
            j = 1
            For j = 1 To objGerTitPagar.nr_parcela
                Set GerTitPagar = New cGeracaoDuplicataPagar
                With GerTitPagar
                    .P_PagRec = "P"
                    .P_Nota = objGerTitPagar.Numero_nota
                    .P_Empresa = objGerTitPagar.Empresa
                    .P_Tipo = objGerTitPagar.Tipo_registro
                    .P_Parcela = intParcela
                    .P_Descricao = objGerTitPagar.Descricao
                    'pt. 85684 - Ivo Sousa(15/07/2008)
                    'Segundo Carlos Dias, o valor original deve ser composto de apenas duas casas decimais
                    curPerc = objRateioTitPagar.Rateio.CurrentObject.R_Percentual
                    .P_Valor_Original = FormatNumber(((objGerTitPagar.Vl_valor_nota / objGerTitPagar.nr_parcela) * curPerc) / 100, 2)
                    curTotal = curTotal + FormatNumber(((objGerTitPagar.Vl_valor_nota / objGerTitPagar.nr_parcela) * objRateioTitPagar.Rateio.CurrentObject.R_Percentual) / 100, 2)
                    .P_Banco = objGerTitPagar.Cd_banco
                    .P_Conta = objRateioTitPagar.Rateio.CurrentObject.R_Cd_conta
                    .P_Centro = objRateioTitPagar.Rateio.CurrentObject.R_Cd_centro_custo
                    .P_cd_operacao_contabil = objGerTitPagar.cd_operacao_contabil
                    .P_Moeda = objGerTitPagar.Cd_moeda
                    i = j
                    'pt. 85684 - Ivo Sousa (14/07/2008)
                    If j = 1 Then
                        dtaData = objGerTitPagar.Dt_data_emissao
                    End If
                    .P_Vencimento = fDataUtil(dtaData, i, objGerTitPagar.Intervalo_vencimento)
                    dtaData = .P_Vencimento
                    .P_Emissao = objGerTitPagar.Dt_data_emissao
                End With
                'pt. 85684 - Ivo Sousa(14/07/2008)
                'Segundo o Carlos Dias, se o valor total der diferente em função de arredondamento
                'colocar a diferença na ultima parcela
                If intTotalParcelas = intParcela Then
                    If curTotal < etxValorNota.valorMoeda Then
                        GerTitPagar.P_Valor_Original = GerTitPagar.P_Valor_Original + FormatNumber(CCur(etxValorNota.valorMoeda) - curTotal, 2)
                    Else
                        GerTitPagar.P_Valor_Original = GerTitPagar.P_Valor_Original - FormatNumber(curTotal - CCur(etxValorNota.valorMoeda), 2)
                    End If
                End If
                Call objCGerTitPagar.parcelas.add(GerTitPagar)
                Set GerTitPagar = Nothing
                intParcela = intParcela + 1
                pgrTitulo.value = pgrTitulo.value + 1
            Next
            objRateioTitPagar.Rateio.MoveNext
        Wend
    Else
        'UM ÚNICO CENTRO DE CUSTO, CADASTRADO NA TELA PRINCIPAL.
        For j = 1 To objGerTitPagar.nr_parcela
            Set GerTitPagar = New cGeracaoDuplicataPagar
            With GerTitPagar
                .P_PagRec = "P"
                .P_Nota = objGerTitPagar.Numero_nota
                .P_Empresa = objGerTitPagar.Empresa
                .P_Tipo = objGerTitPagar.Tipo_registro
                .P_Parcela = j
                .P_Descricao = objGerTitPagar.Descricao
                'pt. 85684 - Ivo Sousa(15/07/2008)
                'Segundo Carlos Dias, o valor original deve ser composto de apenas duas casas decimais
                .P_Valor_Original = FormatNumber((objGerTitPagar.Vl_valor_nota / objGerTitPagar.nr_parcela), 2)
                curTotal = curTotal + FormatNumber((objGerTitPagar.Vl_valor_nota / objGerTitPagar.nr_parcela), 2)
                .P_Banco = objGerTitPagar.Cd_banco
                .P_Conta = objGerTitPagar.Cd_conta
                .P_Centro = objGerTitPagar.Cd_centro_custo
                .P_cd_operacao_contabil = objGerTitPagar.cd_operacao_contabil
                .P_Moeda = objGerTitPagar.Cd_moeda
                i = j
                'pt. 85684 - Ivo Sousa (14/07/2008)
                If j = 1 Then
                    dtaData = objGerTitPagar.Dt_data_emissao
                End If
                .P_Vencimento = fDataUtil(dtaData, i, objGerTitPagar.Intervalo_vencimento)
                dtaData = .P_Vencimento
                .P_Emissao = objGerTitPagar.Dt_data_emissao
            End With
            'pt. 85684 - Ivo Sousa(14/07/2008)
            'Segundo o Carlos Dias, se o valor total der diferente em função de arredondamento
            'colocar a diferença na ultima parcela
            If j = objGerTitPagar.nr_parcela Then
                If curTotal <> etxValorNota.valorMoeda Then
                    GerTitPagar.P_Valor_Original = GerTitPagar.P_Valor_Original + FormatNumber(CCur(etxValorNota.valorMoeda) - curTotal, 2)
                End If
            End If
            Call objCGerTitPagar.parcelas.add(GerTitPagar)
            Set GerTitPagar = Nothing
            pgrTitulo.value = pgrTitulo.value + 1
        Next
    End If
    etxStatus.valorTexto = "Gerado"
    cmdGerarDuplicatas.Enabled = False
    cmdExcluirDuplicatas.Enabled = True
    CarregaGrid
End Sub

Private Sub cmdGerarDuplicatas_Click()
    If ValidaCampos Then
        Call GerarDuplicatas
        'pt. 85684 - Ivo Sousa(14/07/2008)
        MsgBox "As duplicatas ainda não foram geradas. Para gera-las clique no botão 'Gravar'.", vbInformation, NomeModulo
        Call DesabilitaCampos
        blnGerouDuplicatas = False
    End If
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
End Sub

Private Sub cmdNovo_Click()
    Call LibProc(WL_NOVO)
    etxCodigoTitulo.SetFocus
End Sub

Private Sub cmdPesquisar_Click()
    Call LibProc(WL_PESQUISAR)
End Sub

Private Sub cmdRateio_Click()
    frmGeracaoTituloRateioPagar.CarregaObj objGerTitPagar
    frmGeracaoTituloRateioPagar.CarregaCol objRateioTitPagar
    
    'pt. 85684 - Ivo Sousa(15/07/2008)
    frmGeracaoTituloRateioPagar.GerouDuplicatas = (grdTitFin.TextMatrix(1, 1) <> "")
    frmGeracaoTituloRateioPagar.Show vbModal
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
End Sub

Private Sub cmdCancelar_Click()
    'pt. 85684 - Ivo Sousa(14/07/2008)
    If Not booAlterando Then
        Call LibProc(WL_NOVO)
    End If
End Sub

Private Sub ecbTipoRegistro_Change()
    Dim strSql As String
    Dim rstTab As Object
    
    strSql = "SELECT cd_opercontabil_duplpag From MatrizContabilizacao Where tp_registro='" & ecbTipoRegistro.SelectedItem & "'"
    If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
        If etxOperacaoContabil.valorInteiro = 0 And etxOperacaoContabil.Enabled = True Then
            etxOperacaoContabil.valorInteiro = GetValue(rstTab, "cd_opercontabil_duplpag")
        End If
    End If
    FechaRecordset (rstTab)
End Sub

Private Sub etxCentroCusto_Change()
    If etxCentroCusto.valorInteiro > 0 Then
        cmdRateio.Enabled = False
    Else
        cmdRateio.Enabled = True
    End If
End Sub

Private Sub etxCentroFinan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCentroFinan.valorInteiro <> 0 Then
            etxCentroFinan.valorInteiro = 0
        End If
        PCampo "Centro", "SELECT Código, Descrição FROM Centros", pbCampo, etxCentroFinan, "Código"
    End If
End Sub

Private Sub etxCodigoConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCodigoConta.ValorDescricao = "" Then
            etxCodigoConta.valorInteiro = 0
        End If
        'Demanda 130222 - Davi Brito - 13/07/2016
        PCampo "Contas", "SELECT Código, Grupo, Descrição FROM Contas", pbCampo, etxCodigoConta, "Código"
    End If
End Sub

Private Sub etxCodigoTitulo_LostFocus()
    'If Not booAlterando Then
        'pt. 85684 - Ivo Sousa(14/07/2008)
        If ExisteRegistro Then
            Call LibProc(WL_LOCALIZAR)
        End If
    'End If
End Sub

Private Sub etxEmpresa_LostFocus()
    Dim strSql As String
    Dim rstTab As Object

    If Trim(etxEmpresa.valorTexto) <> "" Then
        strSql = "SELECT Banco, Conta FROM Empresas WHERE Apel = '" & etxEmpresa.valorTexto & "';"
        If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
            If etxCodigoBanco.valorInteiro = 0 Then
                etxCodigoBanco.valorInteiro = GetValue(rstTab, "Banco")
            End If
            If etxCodigoConta.valorInteiro = 0 Then
                etxCodigoConta.valorInteiro = GetValue(rstTab, "Conta")
            End If
        End If
        FechaRecordset (rstTab)
    End If
End Sub


Private Sub etxIntervalo_LostFocus()
    'pt. 85684 - Ivo Sousa(14/07/2008)
    If Screen.ActiveControl.name = "etxParcela" Then
        tabTitulos.Tab = 0
    End If
End Sub

Private Sub etxMoeda_LostFocus()
    'pt. 85684 - Ivo Sousa(14/07/2008)
    If Screen.ActiveControl.name = "etxValorFinan" Then
        tabTitulos.Tab = 2
    End If
End Sub

Private Sub etxParcela_LostFocus()
    'pt. 85684 - Ivo Sousa(14/07/2008)
    If Screen.ActiveControl.name = "etxIntervalo" Then
        tabTitulos.Tab = 1
    End If
End Sub

Private Sub etxValorFinan_LostFocus()
    'pt. 85684 - Ivo Sousa(14/07/2008)
    If Screen.ActiveControl.name = "cmdRateio" Or Screen.ActiveControl.name = "etxMoeda" Then
        tabTitulos.Tab = 1
        If cmdRateio.Enabled Then
            cmdRateio.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    Call etxEmpresa.AddConexao(Aplicacao)
    Call etxOperacaoContabil.AddConexao(Aplicacao)
    Call etxCentroCusto.AddConexao(Aplicacao)
    Call etxCodigoBanco.AddConexao(Aplicacao)
    Call etxMoeda.AddConexao(Aplicacao)
    Call etxCodigoConta.AddConexao(Aplicacao)
'    Call etxNotaFinan.AddConexao(Aplicacao)
'    Call etxParcelaFinan.AddConexao(Aplicacao)
'    Call etxValorFinan.AddConexao(Aplicacao)
    Call etxCentroFinan.AddConexao(Aplicacao)

    Set objGerTitPagar = New cGeracaoTituloPagar
    Set objFinanPagar = New cGeracaoDuplicataPagar
    Call preencheCombo
    booValorFoco = False
    Call LibProc(WL_NOVO)
'    etxCodigoTitulo.SetFocus
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: ecbTipoRegistro_Change
'Data.......: 02/07/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para CARREGAR A COMBO DE TIPOS.
'Protocolo..: 85684
'---------------------------------------------------------------------------------------
Private Sub preencheCombo()
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "Tipo"
    cmd.Table.TableName = "[Tipos Globais]"
    cmd.OrderByClause = "Tipo"
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    ecbTipoRegistro.Clear
    While Not rdResult.EOF
        ecbTipoRegistro.AddItem rdResult.GetString("Tipo")
        rdResult.MoveNext
    Wend
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Sub

Private Sub LimpaCampos(Optional booLimpaNota As Boolean = True)
    Dim dao As New cGeracaoTituloPagarDAO
    Call limpaCamposTitulos
    Set dao = Nothing
    booAlterando = False
End Sub

Private Sub limpaCamposTitulos()
    etxCodigoTitulo.valorInteiro = 0
    etxDescricao.valorTexto = ""
    etxNumeroNota.valorInteiro = 0
    ecbTipoRegistro.SelectItem "Fatura"
    etxEmpresa.valorTexto = ""
    etxValorNota.valorMoeda = 0
    edtDataEmissao.Data = Date
    etxIntervalo.valorInteiro = 30
    etxCodigoBanco.valorInteiro = 0
    etxCentroCusto.valorInteiro = 0
    etxCodigoConta.valorInteiro = 0
    etxOperacaoContabil.valorInteiro = 0
    etxMoeda.valorTexto = GetFieldValue("Moeda", "Sistema", "", , "REAL")
    etxParcela.valorInteiro = 0
    etxStatus.valorTexto = ""
    optAnterior.value = True
    grdTitFin.Clear
    grdTitFin.Rows = 2
    pgrTitulo.value = 0
    cmdGravar.Enabled = True
    Call HabilitaCampos
    'Parcelas
    etxNotaFinan.valorInteiro = 0
    etxParcelaFinan.valorInteiro = 0
    etxValorFinan.valorMoeda = 0
    etxCentroFinan.valorInteiro = 0
    edtVencimento.Clear

End Sub

Private Sub bloqueiaCampos()
    etxOperacaoContabil.Enabled = ConfigSys.UtilizaIntegracaoContabil
    etxCentroCusto.Enabled = ConfigSys.ControlarCentrodeCusto
    cmdRateio.Enabled = ConfigSys.ControlarCentrodeCusto
    'pt. 85684 - Ivo Sousa(14/07/2008)
    etxCentroFinan.Enabled = ConfigSys.ControlarCentrodeCusto
End Sub

Public Function LibProc(strFuncao As String, Optional lngFuncao As Long) As Boolean
    Dim dao As New cGeracaoTituloPagarDAO
    Dim facIntegra As New cDAOFactory
    Dim strSql As String
    Dim rstTab As Object
    
On Error GoTo erro_libproc
    
    Select Case strFuncao
        Case WL_SAIR
            'pt. 85684 - Ivo Sousa(14/07/2008)
            If Not blnGerouDuplicatas Then
                If MsgBox("As Duplicatas ainda não foram geradas. Deseja salvar o registro para gerar as Duplicatas?", vbInformation + vbYesNo, NomeModulo) = vbYes Then
                    Call LibProc(WL_SALVAR)
                End If
            End If
            Unload Me
            Exit Function
        Case WL_NOVO
            Set objGerTitPagar = New cGeracaoTituloPagar
            Call LimpaCampos
            booAlterando = False
            tabTitulos.Tab = 0
            etxStatus.valorTexto = "Ativo"
            cmdExcluirDuplicatas.Enabled = False
            cmdGerarDuplicatas.Enabled = True
            Set navigator = Nothing
            Set objCGerTitPagar = Nothing
            Set objRateioTitPagar = Nothing
            Call CarregaGrid
            frmGeracaoTituloRateioPagar.mbolObj = False
            'pt. 85684 - Ivo Sousa (14/07/2008)
            etxCodigoTitulo.valorInteiro = ProximoNumero("cd_titulo", "FFITituloPagar", "")
            blnGerouDuplicatas = True
            'Protocolo Nr 98463 - Carlos Felippe Vernizze - 01/09/2010
            Call bloqueiaCampos
        Case WL_SALVAR
            If ValidaCampos Then
                Aplicacao.Connect
                Aplicacao.BeginTransaction
                Call preencheClasse
                If Not booAlterando Then
                    If Not dao.persistir(objGerTitPagar, Aplicacao, objRateioTitPagar, objCGerTitPagar) Then
                        MsgBox "Ocorreu um erro ao gravar o título.", vbInformation, Me.Caption
                        Aplicacao.RollbackTransaction
                    Else
                        Aplicacao.CommitTransaction
                        MsgBox "Registro gravado com sucesso.", vbInformation, NomeModulo
                        cmdGravar.Enabled = True
                    End If
                Else
                    If dao.Atualizar(objGerTitPagar, Aplicacao, objRateioTitPagar, objCGerTitPagar) Then
                        Aplicacao.CommitTransaction
                        MsgBox "Registro alterado com sucesso.", vbInformation, NomeModulo
                        cmdGravar.Enabled = True
                    Else
                        MsgBox "Ocorreu um erro ao alterar o título.", vbInformation, Me.Caption
                        Aplicacao.RollbackTransaction
                    End If
                End If
                Aplicacao.Disconnect
                'pt. 85684 - Ivo Sousa(14/07/2008)
                cmdRateio.Enabled = True
                blnGerouDuplicatas = True
                'Call LibProc(WL_NOVO)
                'Protocolo Nr 98463 - Carlos Felippe Vernizze - 01/09/2010
                Call bloqueiaCampos
            End If
        Case WL_DELETAR
            If MsgBox("Confirma a exclusão?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                If dao.existir(objGerTitPagar.Cd_Titulo) Then
                    If validaexclusao Then
                        Aplicacao.Connect
                        Aplicacao.BeginTransaction
                        If dao.Excluir(objGerTitPagar, Aplicacao) Then
                            Aplicacao.CommitTransaction
                            Call LibProc(WL_NOVO)
                        Else
                            Aplicacao.RollbackTransaction
                            MsgBox "Ocorreu erro ao tentar excluir o título", vbInformation, Me.Caption
                        End If
                        Aplicacao.Disconnect
                    End If
                Else
                    MsgBox "Título não existente, impossível excluir.", vbInformation, Me.Caption
                End If
            End If
        Case WL_PRIMEIRO
            'pt. 85684 - Ivo Sousa(14/07/2008)
            If Not navigator.EOF Then
                navigator.MoveFirst
                If Not navigator.EOF Then
                    Call setGerTitPagar(navigator.CurrentObject)
                End If
            End If
        Case WL_ANTERIOR
            'pt. 85684 - Ivo Sousa(14/07/2008)
            If Not navigator.EOF Then
                navigator.MovePrevious
                If Not navigator.EOF Then
                    Call setGerTitPagar(navigator.CurrentObject)
                End If
            End If
        Case WL_PROXIMO
            'pt. 85684 - Ivo Sousa(14/07/2008)
            If Not navigator.EOF Then
                navigator.MoveNext
                If Not navigator.EOF Then
                    Call setGerTitPagar(navigator.CurrentObject)
                End If
            End If
        Case WL_ULTIMO
            'pt. 85684 - Ivo Sousa(14/07/2008)
            If Not navigator.EOF Then
                navigator.MoveLast
                If Not navigator.EOF Then
                    Call setGerTitPagar(navigator.CurrentObject)
                End If
            End If
        Case WL_LOCALIZAR
            Call setGerTitPagar(navigator.FindObject(etxCodigoTitulo.valorInteiro))
            'Protocolo Nr 98463 - Carlos Felippe Vernizze - 01/09/2010
            Call bloqueiaCampos
        Case WL_PESQUISAR
            'pt. 85684 - Ivo Sousa(14/07/2008)
            strSql = "SELECT cd_titulo as Código, numero_nota, tipo_registro, vl_valor_Nota, dt_data_emissao, nr_parcelas FROM FFITituloPagar ORDER BY cd_titulo;"
            Call PRegistro(rstTab, Me, "Título", "FFITituloPagar", strSql, Tag, 736, 1)
            If GetValue(rstTab, "cd_titulo") > 0 Then
                Call setGerTitPagar(navigator.FindObject(GetValue(rstTab, "cd_titulo")))
            End If
            'Protocolo Nr 98463 - Carlos Felippe Vernizze - 01/09/2010
            Call bloqueiaCampos
            'Projeto: #1203 - História: # - Problema# - João Henrique(24/04/2012)
        Case WL_AJUDA
            Dim oHelpHtml As New clsHelp
    
            oHelpHtml.Origem = 0
            oHelpHtml.hWnd = Me.hWnd
            oHelpHtml.HelpContext = Me.HelpContextID
            Call oHelpHtml.ShowHelp
            Set oHelpHtml = Nothing
        
    End Select
    Exit Function
erro_libproc:
    FinallyConnection Aplicacao, True
    MsgBox err.Description, vbCritical, Me.Caption
End Function


'---------------------------------------------------------------------------------------
'Procedure..: validaCampos
'Data.......: 02/07/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para VALIDAR ANTES DE GRAVAR.
'Protocolo..: 85684
'---------------------------------------------------------------------------------------
Private Function ValidaCampos() As Boolean
    Dim strMensagem As String
    strMensagem = ""
    
    If Trim(etxDescricao.valorTexto) = "" Then
        strMensagem = strMensagem & "Preenchimento do campo descrição é obrigatório." & vbCrLf
    End If
    If etxNumeroNota.valorInteiro = 0 Then
        strMensagem = strMensagem & "Preenchimento do campo número é obrigatório." & vbCrLf
    End If
    If Trim(etxEmpresa.valorTexto) = "" Then
        strMensagem = strMensagem & "Preenchimento do campo empresa é obrigatório." & vbCrLf
    End If
    If etxValorNota.valorMoeda = 0 Then
        strMensagem = strMensagem & "Preenchimento do campo valor é obrigatório." & vbCrLf
    End If
    If Not edtDataEmissao.IsValidDate Then
        strMensagem = strMensagem & "Preenchimento do campo data de emissão é obrigatório." & vbCrLf
    End If
    If etxIntervalo.valorInteiro = 0 Then
        strMensagem = strMensagem & "Preenchimento do campo intervalo é obrigatório." & vbCrLf
    End If
    If etxCodigoBanco.valorInteiro = 0 Then
        strMensagem = strMensagem & "Preenchimento do campo banco é obrigatório." & vbCrLf
    End If
    'pt. 85684 - Ivo Sousa(15/07/2008)
    If etxCodigoConta.valorInteiro = 0 And objRateioTitPagar.Rateio.Count = 0 Then
        strMensagem = strMensagem & "Preenchimento do campo conta é obrigatório." & vbCrLf
    End If
    If (etxCentroCusto.valorInteiro = 0 And objRateioTitPagar.Rateio.Count = 0) And etxCentroCusto.Enabled = True Then
        strMensagem = strMensagem & "Preenchimento do campo centro de custo é obrigatório." & vbCrLf
    Else
        If etxCentroCusto.valorInteiro <> 0 Then
            Set objRateioTitPagar = Nothing
            Set objRateioTitPagar = New cGeracaoTituloPagar
        End If
    End If
    If etxOperacaoContabil.valorInteiro = 0 And etxOperacaoContabil.Enabled = True Then
        strMensagem = strMensagem & "Preenchimento do campo operação contabil é obrigatório." & vbCrLf
    End If
    If Trim(etxMoeda.valorTexto) = "" Then
        strMensagem = strMensagem & "Preenchimento do campo moeda é obrigatório." & vbCrLf
    End If
    If etxParcela.valorInteiro = 0 Then
        strMensagem = strMensagem & "Preenchimento do campo parcela é obrigatório." & vbCrLf
    End If
    'pt. 85684 - Ivo Sousa(14/07/2008)
    If Trim(grdTitFin.TextMatrix(1, 1)) <> "" Then
        If Not ValidaParcelas Then
            tabTitulos.Tab = 2
            Exit Function
        End If
    End If
    If strMensagem = "" Then
        ValidaCampos = True
    Else
        MsgBox strMensagem, vbInformation, NomeModulo
    End If
End Function

Private Sub preencheClasse()
    Call preencheTitPagarClasse
End Sub

Private Sub preencheTitPagarClasse()
    With objGerTitPagar
        .Cd_Titulo = etxCodigoTitulo.valorInteiro
        .Descricao = etxDescricao.valorTexto
        .Numero_nota = etxNumeroNota.valorInteiro
        .Tipo_registro = ecbTipoRegistro.SelectedItem
        .Empresa = etxEmpresa.valorTexto
        .Vl_valor_nota = etxValorNota.valorMoeda
        .Dt_data_emissao = edtDataEmissao.Data
        .Intervalo_vencimento = etxIntervalo.valorInteiro
        .Cd_banco = etxCodigoBanco.valorInteiro
        .Cd_centro_custo = etxCentroCusto.valorInteiro
        .Cd_conta = etxCodigoConta.valorInteiro
        .cd_operacao_contabil = etxOperacaoContabil.valorInteiro
        .Cd_moeda = etxMoeda.valorTexto
        .nr_parcela = etxParcela.valorInteiro
        .Status = IIf(etxStatus.valorTexto = "Gerado", "G", IIf(etxStatus.valorTexto = "Ativo", "A", ""))
    End With
End Sub

Public Sub setGerTitPagar(GeracaoTit As cGeracaoTituloPagar)
    Set objGerTitPagar = GeracaoTit
    Set objRateioTitPagar = Nothing
    Call LimpaCampos
    Call mostraCamposClasse
    Call SincronizaGrid
    Call CarregaColecao
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set navigator = Nothing
    Set objRateioTitPagar = Nothing
End Sub

Private Sub mostraCamposClasse()
    Call mostraCamposTituloPagar
End Sub

Private Sub mostraCamposTituloPagar()
    booAlterando = True
    With objGerTitPagar
        etxCodigoTitulo.valorInteiro = .Cd_Titulo
        etxDescricao.valorTexto = .Descricao
        etxNumeroNota.valorInteiro = .Numero_nota
        ecbTipoRegistro.SelectItem .Tipo_registro
        etxEmpresa.valorTexto = .Empresa
        etxValorNota.valorMoeda = .Vl_valor_nota
        edtDataEmissao.Data = .Dt_data_emissao
        etxIntervalo.valorInteiro = .Intervalo_vencimento
        etxCodigoBanco.valorInteiro = .Cd_banco
        etxCentroCusto.valorInteiro = .Cd_centro_custo: etxCentroCusto_Change
        etxCodigoConta.valorInteiro = .Cd_conta
        etxOperacaoContabil.valorInteiro = .cd_operacao_contabil
        etxMoeda.valorTexto = .Cd_moeda
        etxParcela.valorInteiro = .nr_parcela
        etxStatus.valorTexto = IIf(.Status = "G", "Gerado", IIf(.Status = "A", "Ativo", ""))
        If .Status = "A" Then
            cmdGerarDuplicatas.Enabled = True
            cmdExcluirDuplicatas.Enabled = False
            cmdGravar.Enabled = True
        Else
            cmdGerarDuplicatas.Enabled = False
            cmdExcluirDuplicatas.Enabled = True
            cmdGravar.Enabled = False
        End If
        If .Cd_Titulo = 0 And .Descricao = "" Then
            booAlterando = False
            Call LibProc(WL_NOVO)
        End If
    End With
End Sub

Private Sub etxEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxEmpresa.ValorDescricao = "" Then
            etxEmpresa.valorTexto = ""
        End If
        PCampo "Empresa", "SELECT Apel, Razão, [CNPJ/CPF], [IEst/RG], Cidade, Estado FROM Empresas WHERE Tipo <> 'Cliente'", pbCampo, etxEmpresa, "APEL"
    End If
End Sub

Private Sub etxCodigoBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCodigoBanco.ValorDescricao = "" Then
            etxCodigoBanco.valorInteiro = 0
        End If
        PCampo "Bancos", "SELECT Banco, Nome, Agência, Conta FROM Bancos", pbCampo, etxCodigoBanco, "Banco"
    End If
End Sub

Private Sub etxCentroCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCentroCusto.ValorDescricao = "" Then
            etxCentroCusto.valorInteiro = 0
        End If
        PCampo "Centro", "SELECT Código, Descrição FROM Centros", pbCampo, etxCentroCusto, "Código"
    End If
End Sub

Private Sub etxOperacaoContabil_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxOperacaoContabil.ValorDescricao = "" Then
            etxOperacaoContabil.valorInteiro = 0
        End If
        PCampo "Operação Contabil", "Select cd_operacao, descricao from OperacaoContabil", pbCampo, etxOperacaoContabil, "cd_operacao"
    End If
End Sub

Private Sub etxMoeda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxMoeda.ValorDescricao = "" Then
            etxMoeda.valorTexto = ""
        End If
        PCampo "Moeda", "Select Moeda, descrição from Moedas", pbCampo, etxMoeda, "Moeda"
    End If
End Sub

Private Function fDataUtil(dtaData As Date, intParcela As Integer, intDia As Integer) As Date
    Dim dtaNovaData As Date
    
    If (Day(dtaData) <= intDia) Then
        intParcela = intParcela - 1
    Else
        intParcela = intParcela
    End If
    'dtaNovaData = Format(DateSerial(Year(dtaData), Month(dtaData) + intParcela, intDia), "DD/MM/YYYY")
    'pt. 85684 - Ivo Sousa (14/07/2008)
    dtaNovaData = Format(dtaData + intDia, "dd/mm/yyyy")
    If Not ValidaDatasDiasUteis(dtaNovaData, , , , False) Then
        If optProximo Then
            EDiaUtil dtaNovaData, EDU_POSTERIOR
        ElseIf optAnterior Then
            EDiaUtil dtaNovaData, EDU_ANTERIOR
        End If
    End If
    fDataUtil = dtaNovaData
End Function

Private Sub CarregaGrid()
    grdTitFin.Clear
    If objCGerTitPagar.parcelas.Count = 0 Then
        Call CarregaHFlexGrid(grdTitFin, Nothing, strTituloGrid)
    Else
        objCGerTitPagar.parcelas.MoveFirst
        Call CarregaHFlexGrid(grdTitFin, , strTituloGrid, , , objCGerTitPagar.parcelas)
    End If
    grdTitFin.FixedCols = 1
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: validaexclusao
'Data.......: 02/07/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para VALIDA ANTES DE EXCLUIR O REGISTRO
'Protocolo..: 85684
'---------------------------------------------------------------------------------------
Private Function validaexclusao() As Boolean
    Dim strSql       As String
    Dim rstTab       As Object
    Dim i            As Integer
    Dim strMensagem  As String
    Dim CurrentObject As cGeracaoDuplicataPagar
    
    strMensagem = ""
    If objCGerTitPagar.parcelas.Count > 0 Then
        objCGerTitPagar.parcelas.MoveFirst
        While Not objCGerTitPagar.parcelas.EOF
            strSql = "SELECT * FROM Duplicatas WHERE PagRec='P' AND Nota=" & objCGerTitPagar.parcelas.CurrentObject.P_Nota & " AND Empresa='" & objCGerTitPagar.parcelas.CurrentObject.P_Empresa & "' AND Tipo='" & objCGerTitPagar.parcelas.CurrentObject.P_Tipo & "' AND Parcela=" & objCGerTitPagar.parcelas.CurrentObject.P_Parcela
            If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
                strMensagem = strMensagem & "Nota: " & objCGerTitPagar.parcelas.CurrentObject.P_Nota & ", Parcela=" & objCGerTitPagar.parcelas.CurrentObject.P_Parcela & vbCrLf
            End If
            FechaRecordset (rstTab)
            objCGerTitPagar.parcelas.MoveNext
        Wend
    End If
    If strMensagem <> "" Then
        MsgBox "Ainda existem duplicatas geradas. Para excluí-las clique no botão 'Excl.Duplicatas' antes de continuar." & vbCrLf & strMensagem, vbInformation, NomeModulo
        Exit Function
    End If
    validaexclusao = True
End Function

Private Sub fExclusaoDuplicatas()

    Dim strSql       As String
    Dim rstTab       As Object
    Dim i            As Integer
    Dim strMensagem  As String
    Dim CurrentObject As cGeracaoDuplicataPagar
    Dim strMsgLog    As String
    
    strMensagem = ""
    'Verifica
    If objCGerTitPagar.parcelas.Count > 0 Then
        objCGerTitPagar.parcelas.MoveFirst
        While Not objCGerTitPagar.parcelas.EOF
            strSql = "SELECT * FROM Duplicatas WHERE PagRec='P' AND Nota=" & objCGerTitPagar.parcelas.CurrentObject.P_Nota & " AND Empresa='" & objCGerTitPagar.parcelas.CurrentObject.P_Empresa & "' AND Tipo='" & objCGerTitPagar.parcelas.CurrentObject.P_Tipo & "' AND Parcela=" & objCGerTitPagar.parcelas.CurrentObject.P_Parcela
            If (AbreRecordset(rstTab, strSql, dbOpenDynaset) = WL_OK) Then
                If GetValue(rstTab, "Pagamento") <> "" Then
                    strMensagem = strMensagem & "Nota: " & objCGerTitPagar.parcelas.CurrentObject.P_Nota & ", Parcela=" & objCGerTitPagar.parcelas.CurrentObject.P_Parcela & vbCrLf
                End If
            End If
            FechaRecordset (rstTab)
            objCGerTitPagar.parcelas.MoveNext
        Wend
    End If
    If strMensagem <> "" Then
        MsgBox "Não é possível excluir pois existem duplicatas baixadas." & vbCrLf & strMensagem, vbInformation
    Else
        'exclui
        If objCGerTitPagar.parcelas.Count > 0 Then
            objCGerTitPagar.parcelas.MoveFirst
            While Not objCGerTitPagar.parcelas.EOF
                strSql = "SELECT * FROM Duplicatas WHERE PagRec='P' AND isnull(Pagamento) AND Nota=" & objCGerTitPagar.parcelas.CurrentObject.P_Nota & " AND Empresa='" & objCGerTitPagar.parcelas.CurrentObject.P_Empresa & "' AND Tipo='" & objCGerTitPagar.parcelas.CurrentObject.P_Tipo & "' AND Parcela=" & objCGerTitPagar.parcelas.CurrentObject.P_Parcela
                If (AbreRecordset(rstTab, strSql, dbOpenDynaset) = WL_OK) Then
                    If GetValue(rstTab, "Pagamento") = "" Then
                        rstTab.Delete
                        conexao.Execute "DELETE FROM FVFTituloPagarDuplicata WHERE PagRec='P' AND cd_titulo=" & objGerTitPagar.Cd_Titulo & " AND Nota=" & objCGerTitPagar.parcelas.CurrentObject.P_Nota & " AND Empresa='" & objCGerTitPagar.parcelas.CurrentObject.P_Empresa & "' AND tipo_registro='" & objCGerTitPagar.parcelas.CurrentObject.P_Tipo & "' AND Parcela=" & objCGerTitPagar.parcelas.CurrentObject.P_Parcela
                        'Projeto: 100340 - Desenv.: 142890 - Ueder Budni (23/09/2016)
                        With objCGerTitPagar.parcelas.CurrentObject
                            strMsgLog = "Títulos excluído através da rotina de Geração de Títulos a Pagar"
                            Call RegistraLogLancDup(.P_Nota, .P_Empresa, .P_Tipo, .P_Parcela, "P", Duplicata, strMsgLog)
                        End With
                    End If
                End If
                FechaRecordset (rstTab)
                objCGerTitPagar.parcelas.MoveNext
            Wend
        End If
        etxStatus.valorTexto = "Ativo"
        cmdGerarDuplicatas.Enabled = True
        cmdExcluirDuplicatas.Enabled = False
        cmdGravar.Enabled = True
        conexao.Execute "UPDATE FFITituloPagar SET status='A' WHERE cd_titulo=" & objGerTitPagar.Cd_Titulo
    End If
    'pt. 85684 - Ivo Sousa(14/07/2008)
    Call DesabilitaCampos
    Call SincronizaGrid
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: SincronizaGrid
'Data.......: 02/07/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para ATUALIZAR O GRID NA MUDANÇA DE REGISTRO.
'Protocolo..: 85684
'---------------------------------------------------------------------------------------
Private Sub SincronizaGrid()
    Dim strSql       As String
    Dim rstTab       As Object
    Dim i            As Integer
    Dim j            As Integer
    Dim strMensagem  As String
    Dim GerTitPagar As New cGeracaoDuplicataPagar
    
    grdTitFin.Clear
    
    If objCGerTitPagar Is Nothing Then
        Set objCGerTitPagar = New cGeracaoDuplicataPagar
    Else
        Set objCGerTitPagar = Nothing
        Set objCGerTitPagar = New cGeracaoDuplicataPagar
    End If
    
    strMensagem = ""
    'Verifica se existe parcelas geradas.
    strSql = ""
    strSql = strSql & "SELECT DISTINCT Duplicatas.PagRec , Duplicatas.Nota , Duplicatas.Empresa, Duplicatas.Tipo, Duplicatas.Parcela, Duplicatas.Descrição, Duplicatas.[Valor Original], Duplicatas.Banco, Duplicatas.Conta , Duplicatas.Centro, Duplicatas.cd_operacao_contabil, Duplicatas.Moeda, Duplicatas.Vencimento, Duplicatas.Emissão, Duplicatas.Pagamento "
    strSql = strSql & "FROM Duplicatas INNER JOIN FVFTituloPagarDuplicata "
    strSql = strSql & "ON (Duplicatas.Parcela = FVFTituloPagarDuplicata.Parcela) AND (Duplicatas.Tipo = FVFTituloPagarDuplicata.tipo_registro) AND (Duplicatas.Empresa = FVFTituloPagarDuplicata.empresa) AND (Duplicatas.Nota = FVFTituloPagarDuplicata.nota) AND (Duplicatas.PagRec = FVFTituloPagarDuplicata.pagRec) "
    'strSql = strSql & "WHERE Duplicatas.PagRec='P' AND Duplicatas.Nota=" & objGerTitPagar.Numero_nota & " AND Duplicatas.Empresa='" & objGerTitPagar.Empresa & "' AND Duplicatas.Tipo='" & objGerTitPagar.Tipo_registro & "' ORDER BY Duplicatas.Parcela"
    strSql = strSql & "WHERE FVFTituloPagarDuplicata.Cd_Titulo=" & objGerTitPagar.Cd_Titulo
    If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
        rstTab.MoveFirst
        While Not rstTab.EOF
            Set GerTitPagar = New cGeracaoDuplicataPagar
            With GerTitPagar
                .P_PagRec = GetValue(rstTab, "PagRec")
                .P_Nota = GetValue(rstTab, "Nota")
                .P_Empresa = GetValue(rstTab, "Empresa")
                .P_Tipo = GetValue(rstTab, "Tipo")
                .P_Parcela = GetValue(rstTab, "Parcela")
                .P_Descricao = GetValue(rstTab, "Descrição")
                .P_Valor_Original = GetValue(rstTab, "Valor Original")
                .P_Banco = GetValue(rstTab, "Banco")
                .P_Conta = GetValue(rstTab, "Conta")
                .P_Centro = GetValue(rstTab, "Centro")
                .P_cd_operacao_contabil = GetValue(rstTab, "cd_operacao_contabil")
                .P_Moeda = GetValue(rstTab, "Moeda")
                .P_Vencimento = GetValue(rstTab, "Vencimento")
                .P_Emissao = GetValue(rstTab, "Emissão")
                .P_Pagamento = GetValue(rstTab, "Pagamento")
            End With
            Call objCGerTitPagar.parcelas.add(GerTitPagar)
            Set GerTitPagar = Nothing
            rstTab.MoveNext
            cmdExcluirDuplicatas.Enabled = True
        Wend
        'pt. 85684 - Ivo Sousa(14/07/2008)
        Call DesabilitaCampos
    Else
        Call HabilitaCampos
    End If
    'Debug.Print "Nota: " & objGerTitPagar.Numero_nota & " - Parcelas: " & objCGerTitPagar.parcelas.Count
    Call CarregaGrid
    FechaRecordset (rstTab)
End Sub

'---------------------------------------------------------------------------------------
'Procedure..: CarregaColecao
'Data.......: 02/07/2008
'Autor......: MOACIR PFAU
'Descrição..: Utilizado para CARREGAR AS INFORMAÇÕES DE RATEIO NA COLEÇÃO.
'Protocolo..: 85684
'---------------------------------------------------------------------------------------
Private Sub CarregaColecao()
    Dim strSql       As String
    Dim rstTab       As Object
    Dim i            As Integer
    Dim GerTitPagar As New cGeracaoTituloPagar
    
    Set objRateioTitPagar = Nothing
    Set objRateioTitPagar = New cGeracaoTituloPagar
    'Verifica se existe parcelas geradas.
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & "FROM FFITituloPagarRateio "
    strSql = strSql & "WHERE cd_titulo=" & objGerTitPagar.Cd_Titulo
    If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
        rstTab.MoveFirst
        While Not rstTab.EOF
            Set GerTitPagar = New cGeracaoTituloPagar
            With GerTitPagar
                .R_Cd_titulo = objGerTitPagar.Cd_Titulo
                .R_Cd_centro_custo = GetValue(rstTab, "Cd_Centro_Custo")
                .R_Cd_conta = GetValue(rstTab, "Cd_Conta_Financeira")
                .R_Percentual = GetValue(rstTab, "pr_percentual")
            End With
            Call objRateioTitPagar.Rateio.add(GerTitPagar)
            Set GerTitPagar = Nothing
            rstTab.MoveNext
        Wend
    End If
    FechaRecordset (rstTab)
End Sub

Public Sub CarregaColRateio(mobjRateio As Object)
    Set objRateioTitPagar = mobjRateio
End Sub

'pt. 85684 - Ivo Sousa (14/07/2008)
Private Sub grdTitFin_Click()
    grdTitFin.SetFocus
End Sub

Private Sub grdTitFin_DblClick()
    mostraCamposClasseFinanc
End Sub

Private Sub mostraCamposClasseFinanc()
    'pt. 85684 - Ivo Sousa (14/07/2008)
    If Trim(grdTitFin.TextMatrix(grdTitFin.Row, 1)) = "" Then
        Exit Sub
    End If
    Call carregaCamposFinancPagar
    If objFinanPagar.P_Pagamento <> "" Then
        MsgBox "Duplicata baixada. Não pode ser alterada.", vbCritical + vbInformation
    Else
        Call mostraCamposFinancPagar
        cmdAlterar.Enabled = True
    End If
End Sub

Private Sub carregaCamposFinancPagar()
    With objFinanPagar
        .P_PagRec = grdTitFin.TextMatrix(grdTitFin.Row, 0)
        .P_Nota = grdTitFin.TextMatrix(grdTitFin.Row, 1)
        .P_Empresa = grdTitFin.TextMatrix(grdTitFin.Row, 2)
        .P_Tipo = grdTitFin.TextMatrix(grdTitFin.Row, 3)
        .P_Parcela = grdTitFin.TextMatrix(grdTitFin.Row, 4)
        .P_Descricao = grdTitFin.TextMatrix(grdTitFin.Row, 5)
        .P_Valor_Original = grdTitFin.TextMatrix(grdTitFin.Row, 6)
        .P_Banco = grdTitFin.TextMatrix(grdTitFin.Row, 7)
        .P_Conta = grdTitFin.TextMatrix(grdTitFin.Row, 8)
        .P_Centro = grdTitFin.TextMatrix(grdTitFin.Row, 9)
        .P_cd_operacao_contabil = grdTitFin.TextMatrix(grdTitFin.Row, 10)
        .P_Moeda = grdTitFin.TextMatrix(grdTitFin.Row, 11)
        .P_Vencimento = grdTitFin.TextMatrix(grdTitFin.Row, 12)
        .P_Pagamento = grdTitFin.TextMatrix(grdTitFin.Row, 13)
        .P_Emissao = grdTitFin.TextMatrix(grdTitFin.Row, 14)
    End With
    'booAlterando = True
End Sub

Private Sub mostraCamposFinancPagar()
    With objFinanPagar
        etxNotaFinan.valorInteiro = .P_Nota
        etxParcelaFinan.valorInteiro = .P_Parcela
        etxValorFinan.valorMoeda = .P_Valor_Original
        etxCentroFinan.valorInteiro = .P_Centro
        edtVencimento.Data = .P_Vencimento
    End With
End Sub

Private Sub preencheClasseFinanc()
    Call preencheFinancPagarClasse
End Sub

Private Sub preencheFinancPagarClasse()
    With objFinanPagar
        .P_Nota = etxNotaFinan.valorInteiro
        .P_Parcela = etxParcelaFinan.valorInteiro
        .P_Valor_Original = etxValorFinan.valorMoeda
        .P_Centro = etxCentroFinan.valorInteiro
        .P_Vencimento = edtVencimento.Data
    End With
End Sub

Private Sub cmdAlterar_Click()
    'validação
    If ValidaCamposFinan Then
        preencheClasseFinanc
        Call objCGerTitPagar.parcelas.update(objFinanPagar)
        CarregaGrid
        etxNotaFinan.valorInteiro = 0
        etxParcelaFinan.valorInteiro = 0
        etxValorFinan.valorMoeda = 0
        etxCentroFinan.valorInteiro = 0
        edtVencimento.Clear
        cmdGravar.Enabled = True
        cmdAlterar.Enabled = False
        etxDescricao.SetFocus
    End If
End Sub
Private Function ValidaCamposFinan() As Boolean
    If etxValorFinan.valorMoeda = 0 Then
        MsgBox "Preenchimento do campo valor é obrigatório.", vbInformation, NomeModulo
        etxValorFinan.SetFocus
        Exit Function
    End If
    If etxCentroFinan.Enabled Then
        If etxCentroFinan.valorInteiro = 0 Then
            MsgBox "Preenchimento do campo centro de custo é obrigatório.", vbInformation, NomeModulo
            etxCentroFinan.SetFocus
            Exit Function
        End If
    End If
    If Not edtVencimento.IsValidDate Then
        MsgBox "Preenchimento do campo data de vencimento é obrigatório.", vbInformation, NomeModulo
        edtVencimento.SetFocus
        Exit Function
    End If
    
    'pt. 85684 - Ivo Sousa(14/07/2008)
    If edtVencimento.Data < edtDataEmissao.Data Then
        MsgBox "A data de vencimento não pode ser menor que a data de emissão.", vbInformation, NomeModulo
        edtVencimento.SetFocus
        Exit Function
    End If

    'pt. 85684 - Ivo Sousa(14/07/2008)
    Call ValidaParcelas(True)
    ValidaCamposFinan = True
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
    If KeyCode = 114 Then
        Call LibProc(WL_PESQUISAR)
    End If
End Sub
'pt.........: 85684
'Data.......: 14/07/2008
'Autor......: Ivo Sousa
'Descrição..: Função utilizada para a verficar se um determinado registro existe
'Retorno....: [Boolean] Se o registro existe
Private Function ExisteRegistro() As Boolean
    If GetFieldValue("cd_titulo", "FFITituloPagar", "cd_titulo = " & etxCodigoTitulo.valorInteiro, , 0) > 0 Then
        ExisteRegistro = True
    End If
End Function

'pt.........: 85684
'Data.......: 14/07/2008
'Autor......: Ivo Sousa
'Descrição..: Função utilizada para a verficação se o valor total das parcelas
'             é igual ao valor da nota
'Parametros.: [Boolean] Se pode ser exibida a diferença.
'Retorno....: [Boolean] Se validou a parcela que esta sendo gerada.
Private Function ValidaParcelas(Optional blnMostraDiferenca As Boolean) As Boolean
    Dim intCont            As Integer
    Dim curTotalDuplicatas As Currency
    Dim curDiferenca       As Currency
    
    intCont = 1
    With grdTitFin
        While intCont <= .Rows - 1
            curTotalDuplicatas = curTotalDuplicatas + .TextMatrix(intCont, 6)
            intCont = intCont + 1
        Wend
        If blnMostraDiferenca Then
            curTotalDuplicatas = (curTotalDuplicatas - grdTitFin.TextMatrix(grdTitFin.Row, 6)) + etxValorFinan.valorMoeda
        End If
    End With
    If Not curTotalDuplicatas = CCur(etxValorNota.valorMoeda) Then
        If Not blnMostraDiferenca Then
            MsgBox "O valor total das duplicatas não pode ser diferente do valor da Nota.", vbInformation, NomeModulo
            etxValorFinan.SetFocus
            Exit Function
        Else
            curDiferenca = etxValorNota.valorMoeda - curTotalDuplicatas
            MsgBox "O saldo restante para o valor da Nota Fiscal é de R$ " & curDiferenca & ".", vbInformation, NomeModulo
        End If
    End If
    ValidaParcelas = True
End Function

'pt.........: 85684
'Data.......: 15/07/2008
'Autor......: Ivo Sousa
'Descrição..: Função utilizada para desabilitar os campos da tela
Private Sub DesabilitaCampos()
    etxNumeroNota.Enabled = False
    ecbTipoRegistro.Enabled = False
    etxEmpresa.Enabled = False
    etxValorNota.Enabled = False
    edtDataEmissao.Enabled = False
    etxIntervalo.Enabled = False
    etxCodigoBanco.Enabled = False
    etxCentroCusto.Enabled = False
    etxCodigoConta.Enabled = False
    etxOperacaoContabil.Enabled = False
    etxMoeda.Enabled = False
    etxParcela.Enabled = False
    fraData.Enabled = False
End Sub

'pt.........: 85684
'Data.......: 15/07/2008
'Autor......: Ivo Sousa
'Descrição..: Função utilizada para Habilitar os campos da tela
Private Sub HabilitaCampos()
    etxNumeroNota.Enabled = True
    ecbTipoRegistro.Enabled = True
    etxEmpresa.Enabled = True
    etxValorNota.Enabled = True
    edtDataEmissao.Enabled = True
    etxIntervalo.Enabled = True
    etxCodigoBanco.Enabled = True
    etxCentroCusto.Enabled = True
    etxCodigoConta.Enabled = True
    etxOperacaoContabil.Enabled = True
    etxMoeda.Enabled = True
    etxParcela.Enabled = True
    fraData.Enabled = True
End Sub

'Projeto: 100340 - Desenv.: 142890 - Ueder Budni (23/09/2016)
Private Sub RegistraLogLancDup(dblNumero As Double, strEmpresa As String, strTipo As String, lngParcela As Long, strPagRec As String, enuTabela As enuLancDup, strMsg As String)
    Dim objLogLancDup   As New clsLogLancamentosDuplicatas

On Error GoTo erro
    Call objLogLancDup.SetKey(strPagRec, dblNumero, strEmpresa, strTipo, lngParcela, enuTabela)
    Call objLogLancDup.InsertMsg(strMsg)
erro:
    Set objLogLancDup = Nothing
End Sub

