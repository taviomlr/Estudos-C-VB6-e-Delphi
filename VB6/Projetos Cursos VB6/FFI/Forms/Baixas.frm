VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmBaixas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixas"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   990
   ClientWidth     =   12300
   Icon            =   "Baixas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   12300
   Begin VB.Frame Frame1 
      Height          =   8655
      Left            =   10905
      TabIndex        =   75
      Top             =   -45
      Width           =   1365
      Begin VB.ComboBox cboBaixas 
         Height          =   315
         Index           =   1
         ItemData        =   "Baixas.frx":0682
         Left            =   45
         List            =   "Baixas.frx":0695
         TabIndex        =   84
         Top             =   5940
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtBaixas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Height          =   285
         Index           =   7
         Left            =   45
         MultiLine       =   -1  'True
         TabIndex        =   83
         Top             =   6975
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.TextBox txtBaixas 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
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
         Height          =   285
         Index           =   6
         Left            =   45
         MultiLine       =   -1  'True
         TabIndex        =   82
         Top             =   6660
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.CommandButton cmdBaixas 
         Caption         =   "&Visualizar"
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   38
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdBaixas 
         Caption         =   "&Editar"
         Height          =   375
         Index           =   1
         Left            =   75
         TabIndex        =   39
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton cmdBaixaLote 
         Caption         =   "&Baixar Lote"
         Enabled         =   0   'False
         Height          =   375
         Left            =   75
         TabIndex        =   40
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair2 
         Caption         =   "Sair"
         Height          =   375
         Left            =   75
         TabIndex        =   41
         Top             =   1410
         Width           =   1215
      End
      Begin ComctlLib.ListView lvwBaixas2 
         Height          =   405
         Left            =   45
         TabIndex        =   81
         Top             =   6255
         Visible         =   0   'False
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   714
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327682
         SmallIcons      =   "imgBaixas"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Número"
            Object.Width           =   1693
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Parcela/Tipo"
            Object.Width           =   1693
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Descrição"
            Object.Width           =   5080
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Empresa"
            Object.Width           =   1773
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Vencimento"
            Object.Width           =   1693
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Valor"
            Object.Width           =   1773
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Emissão"
            Object.Width           =   1693
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Controle"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList imgImagens 
         Left            =   45
         Top             =   7290
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
               Picture         =   "Baixas.frx":06CE
               Key             =   "Checked"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Baixas.frx":0828
               Key             =   "Unchecked"
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList imgBaixas 
         Left            =   30
         Top             =   7890
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
               Picture         =   "Baixas.frx":0982
               Key             =   "Checked"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "Baixas.frx":0C9C
               Key             =   "Unchecked"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8625
      Left            =   30
      TabIndex        =   42
      Top             =   -60
      Width           =   10875
      Begin VB.Frame Frame3 
         Caption         =   "Dados de Duplicata / Lançamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1515
         Left            =   60
         TabIndex        =   104
         Top             =   7050
         Width           =   7240
         Begin VB.CommandButton cmdComfirmar 
            Caption         =   "Confirmar"
            Height          =   375
            Left            =   5860
            TabIndex        =   37
            Top             =   520
            Width           =   1215
         End
         Begin Fox.EBSText txtVlAcrescimo 
            Height          =   330
            Left            =   1200
            TabIndex        =   35
            Top             =   720
            Width           =   1670
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
         Begin Fox.EBSText txtVlAbatimento 
            Height          =   330
            Left            =   4080
            TabIndex        =   36
            Top             =   360
            Width           =   1665
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
         Begin VB.Label lblVlTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4080
            TabIndex        =   111
            Top             =   720
            Width           =   1635
         End
         Begin VB.Label lblValorOriginal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1200
            TabIndex        =   110
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Vl. Abatimento"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   2880
            TabIndex        =   109
            Top             =   390
            Width           =   1125
         End
         Begin VB.Label lblDescricao 
            Caption         =   "Pressione a tecla 'P' para efetuar uma baixa parcial."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   1140
            Width           =   3885
         End
         Begin VB.Label lblValor 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   675
            TabIndex        =   107
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Vl. Acréscimo"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Vl. Total"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   3195
            TabIndex        =   105
            Top             =   780
            Width           =   795
         End
      End
      Begin VB.TextBox txtEmpresaUsuaria 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1395
         TabIndex        =   0
         Text            =   "txtEmpresaUsuaria"
         Top             =   180
         Width           =   1500
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Espéc&ie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1005
         Index           =   1
         Left            =   9360
         TabIndex        =   43
         Top             =   2625
         Visible         =   0   'False
         Width           =   1450
         Begin VB.OptionButton optBaixas 
            Caption         =   "À Pagar"
            ForeColor       =   &H80000006&
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   33
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton optBaixas 
            Caption         =   "À Receber"
            ForeColor       =   &H80000006&
            Height          =   255
            Index           =   1
            Left            =   60
            TabIndex        =   34
            Top             =   570
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Selecionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1005
         Index           =   5
         Left            =   6870
         TabIndex        =   63
         Top             =   2625
         Width           =   2460
         Begin VB.CommandButton cmdNenhum 
            Caption         =   "Nenhum"
            Height          =   315
            Left            =   140
            TabIndex        =   32
            Top             =   630
            Width           =   2175
         End
         Begin VB.CommandButton cmdTodos 
            Caption         =   "Todos"
            Height          =   345
            Left            =   140
            TabIndex        =   31
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Tipos de Registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1005
         Index           =   4
         Left            =   60
         TabIndex        =   62
         Top             =   2625
         Width           =   3000
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optDup 
            Caption         =   "Duplicatas"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1395
         End
         Begin VB.OptionButton optLanc 
            Caption         =   "Lançamentos"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1395
         End
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Ordem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1005
         Index           =   3
         Left            =   3105
         TabIndex        =   61
         Top             =   2625
         Width           =   3720
         Begin VB.OptionButton optNotaCOd 
            Caption         =   "Nota/Código"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   210
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optEmpresa 
            Caption         =   "Empresa"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   450
            Width           =   1335
         End
         Begin VB.OptionButton optControle 
            Caption         =   "Controle"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   690
            Width           =   1335
         End
         Begin VB.OptionButton optEmissao 
            Caption         =   "Emissão"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   1920
            TabIndex        =   29
            Top             =   210
            Width           =   915
         End
         Begin VB.OptionButton optVenc 
            Caption         =   "Vencimento"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   1920
            TabIndex        =   30
            Top             =   450
            Width           =   1155
         End
      End
      Begin MSComctlLib.ListView lvwBaixas 
         Height          =   3285
         Left            =   30
         TabIndex        =   79
         Top             =   3660
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   5794
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgImagens"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Número"
            Object.Width           =   1693
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Origem"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Parc/Tipo"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descrição"
            Object.Width           =   3683
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Empresa"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Banco"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Conta"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "C.C"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Vencimento"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Valor"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Emissão"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Controle"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame fraDesc 
         Caption         =   "Total de Registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1515
         Left            =   7350
         TabIndex        =   57
         Top             =   7050
         Width           =   3465
         Begin VB.Label txtTotalSelecionados 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1665
            TabIndex        =   73
            Top             =   1050
            Width           =   1635
         End
         Begin VB.Label txtQtlSelecionados 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   72
            Top             =   690
            Width           =   1155
         End
         Begin VB.Label txtTotalListados 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1680
            TabIndex        =   71
            Top             =   330
            Width           =   1155
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   315
            TabIndex        =   66
            Top             =   1080
            Width           =   1275
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Qt. Selecionados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   1485
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Qtd. Listados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   435
            TabIndex        =   64
            Top             =   360
            Width           =   1185
         End
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Baixas em Lote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   2115
         Index           =   2
         Left            =   60
         TabIndex        =   67
         Top             =   510
         Visible         =   0   'False
         Width           =   10755
         Begin VB.CheckBox chkConciliado 
            Caption         =   "Conciliado"
            Height          =   255
            Left            =   7050
            TabIndex        =   92
            Top             =   180
            Width           =   1035
         End
         Begin VB.CommandButton cmdBaixar 
            Caption         =   "&Baixar em Lote"
            Height          =   375
            Left            =   9460
            TabIndex        =   100
            Top             =   210
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   9460
            TabIndex        =   101
            Top             =   630
            Width           =   1215
         End
         Begin VB.CommandButton cmdSair 
            Caption         =   "&Sair"
            Height          =   375
            Left            =   9460
            TabIndex        =   103
            Top             =   1050
            Width           =   1215
         End
         Begin VB.CommandButton cmdProximoCheque 
            Caption         =   "..."
            Height          =   345
            Left            =   6030
            TabIndex        =   99
            ToolTipText     =   "Trazer Próximo Número do Cheque"
            Top             =   1170
            Width           =   255
         End
         Begin Fox.EBSText etxOpContabilDupl 
            Height          =   330
            Left            =   4800
            TabIndex        =   97
            Top             =   840
            Width           =   4410
            _ExtentX        =   439103
            _ExtentY        =   582
            TipoTexto       =   0
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "cd_operacao"
            TipoCriterio    =   4
            CampoDescricao  =   "descricao"
            TabelaConsulta  =   "OperacaoContabil"
            TamanhoDescricao=   3200
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
         Begin Fox.EBSText etxBancoBaixa 
            Height          =   330
            Left            =   1575
            TabIndex        =   91
            Top             =   150
            Width           =   5115
            _ExtentX        =   440346
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Banco"
            TipoCriterio    =   4
            CampoDescricao  =   "Nome"
            TabelaConsulta  =   "Bancos"
            TamanhoDescricao=   3900
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
         Begin Fox.EBSText etxControleBaixa 
            Height          =   330
            Left            =   1575
            TabIndex        =   95
            Top             =   1185
            Width           =   1680
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   18
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
         Begin Fox.EBSText etxChequeBaixa 
            Height          =   330
            Left            =   4800
            TabIndex        =   98
            Top             =   1185
            Width           =   1230
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   6
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
         Begin Fox.EBSData edtDataLiberacao 
            Height          =   330
            Left            =   1575
            TabIndex        =   94
            Top             =   840
            Width           =   1230
            _ExtentX        =   2170
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
         Begin Fox.EBSText etxOpContabilLanc 
            Height          =   330
            Left            =   4800
            TabIndex        =   96
            Top             =   495
            Width           =   4410
            _ExtentX        =   439103
            _ExtentY        =   582
            TipoTexto       =   0
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "cd_operacao"
            TipoCriterio    =   4
            CampoDescricao  =   "descricao"
            TabelaConsulta  =   "OperacaoContabil"
            TamanhoDescricao=   3200
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
         Begin Fox.EBSData edtDataPagamento 
            Height          =   330
            Left            =   1575
            TabIndex        =   93
            Top             =   495
            Width           =   1230
            _ExtentX        =   2170
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
         Begin VB.Line Line 
            BorderColor     =   &H80000010&
            X1              =   9375
            X2              =   9375
            Y1              =   90
            Y2              =   2085
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data de Pagamento"
            ForeColor       =   &H80000006&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   114
            Top             =   570
            Width           =   1440
         End
         Begin VB.Image Image 
            Height          =   480
            Index           =   2
            Left            =   120
            Picture         =   "Baixas.frx":0FB6
            Top             =   1530
            Width           =   480
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0FFFF&
            Caption         =   $"Baixas.frx":1BF8
            Height          =   480
            Left            =   600
            TabIndex        =   113
            Top             =   1530
            Width           =   8775
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Op. Contábil Lançamentos"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   2880
            TabIndex        =   112
            Top             =   570
            Width           =   1875
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Data de Liberação"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   60
            TabIndex        =   102
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Op. Contábil Duplicatas"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   2880
            TabIndex        =   78
            Top             =   900
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   70
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Controle"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   3
            Left            =   900
            TabIndex        =   69
            Top             =   1245
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cheque"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   1
            Left            =   4170
            TabIndex        =   68
            Top             =   1260
            Width           =   555
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   9375
            X2              =   9375
            Y1              =   105
            Y2              =   2075
         End
      End
      Begin VB.Frame fraBaixas 
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
         ForeColor       =   &H00000000&
         Height          =   2130
         Index           =   0
         Left            =   60
         TabIndex        =   44
         Top             =   490
         Width           =   10755
         Begin VB.ComboBox cboBaixas 
            Height          =   315
            Index           =   0
            ItemData        =   "Baixas.frx":1C8C
            Left            =   8745
            List            =   "Baixas.frx":1C8E
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   555
            Width           =   1830
         End
         Begin Fox.EBSText etxBancoInicial 
            Height          =   330
            Left            =   1080
            TabIndex        =   7
            Top             =   1260
            Width           =   1230
            _ExtentX        =   1296
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
         Begin Fox.EBSData edtDataLiberacaoInicial 
            Height          =   330
            Left            =   1080
            TabIndex        =   1
            Top             =   180
            Width           =   1230
            _ExtentX        =   2170
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
         Begin Fox.EBSData edtDataLiberacaoFinal 
            Height          =   330
            Left            =   2610
            TabIndex        =   2
            Top             =   180
            Width           =   1230
            _ExtentX        =   2170
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
         Begin Fox.EBSData edtDataVencimentoInicial 
            Height          =   330
            Left            =   1080
            TabIndex        =   3
            Top             =   540
            Width           =   1230
            _ExtentX        =   2170
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
         Begin Fox.EBSData edtDataVencimentoFinal 
            Height          =   330
            Left            =   2610
            TabIndex        =   4
            Top             =   540
            Width           =   1230
            _ExtentX        =   2170
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
         Begin Fox.EBSData edtDataEmissaoInicial 
            Height          =   330
            Left            =   1080
            TabIndex        =   5
            Top             =   900
            Width           =   1230
            _ExtentX        =   2170
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
         Begin Fox.EBSData edtDataEmissaoFinal 
            Height          =   330
            Left            =   2610
            TabIndex        =   6
            Top             =   900
            Width           =   1230
            _ExtentX        =   2170
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
         Begin Fox.EBSText etxBancoFinal 
            Height          =   330
            Left            =   2610
            TabIndex        =   8
            Top             =   1260
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
            Left            =   1080
            TabIndex        =   9
            Top             =   1620
            Width           =   1230
            _ExtentX        =   2011
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
            Left            =   2610
            TabIndex        =   10
            Top             =   1620
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
         Begin Fox.EBSText etxCentroCustoFinal 
            Height          =   330
            Left            =   6480
            TabIndex        =   12
            Top             =   180
            Width           =   1230
            _ExtentX        =   2090
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
         Begin Fox.EBSText etxValorOriginalInicial 
            Height          =   330
            Left            =   5040
            TabIndex        =   13
            Top             =   540
            Width           =   1230
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
            MaxLength       =   18
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
         Begin Fox.EBSText etxCentroCustoInicial 
            Height          =   330
            Left            =   5040
            TabIndex        =   11
            Top             =   180
            Width           =   1230
            _ExtentX        =   2170
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
         Begin Fox.EBSText etxValorOriginalFinal 
            Height          =   330
            Left            =   6480
            TabIndex        =   14
            Top             =   540
            Width           =   1230
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            TipoTexto       =   0
            MaxLength       =   18
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
         Begin Fox.EBSText etxNumero 
            Height          =   330
            Left            =   5040
            TabIndex        =   15
            Top             =   900
            Width           =   1230
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
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
            Left            =   7155
            TabIndex        =   16
            Top             =   900
            Width           =   555
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   3
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
         Begin Fox.EBSText etxCidade 
            Height          =   330
            Left            =   5040
            TabIndex        =   17
            Top             =   1260
            Width           =   2670
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
         Begin Fox.EBSText etxEmpresas 
            Height          =   330
            Left            =   5040
            TabIndex        =   18
            Top             =   1620
            Width           =   5280
            _ExtentX        =   439605
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Apel"
            CampoDescricao  =   "Razão"
            TabelaConsulta  =   "Empresas"
            TamanhoDescricao=   3480
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
         Begin Fox.EBSText etxEstado 
            Height          =   330
            Left            =   8745
            TabIndex        =   21
            Top             =   900
            Width           =   1935
            _ExtentX        =   435848
            _ExtentY        =   582
            Tipo            =   4
            MaxLength       =   2
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Sigla"
            CampoDescricao  =   "Nome"
            TabelaConsulta  =   "Estados"
            TamanhoDescricao=   1350
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
            Left            =   8745
            TabIndex        =   22
            Top             =   1260
            Width           =   1815
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   18
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
            Left            =   8730
            TabIndex        =   19
            Top             =   180
            Width           =   1860
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   40
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "a"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   6315
            TabIndex        =   90
            Top             =   270
            Width           =   90
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "C.Custo"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   4365
            TabIndex        =   89
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "a"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   2415
            TabIndex        =   88
            Top             =   1665
            Width           =   90
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cont&a"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   540
            TabIndex        =   87
            Top             =   1665
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "a"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   2415
            TabIndex        =   86
            Top             =   1305
            Width           =   90
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "&Banco"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   495
            TabIndex        =   85
            Top             =   1305
            Width           =   465
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "a"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   16
            Left            =   2415
            TabIndex        =   77
            Top             =   960
            Width           =   90
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "&Emissão"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   5
            Left            =   405
            TabIndex        =   76
            Top             =   960
            Width           =   585
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "a"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   15
            Left            =   2415
            TabIndex        =   60
            Top             =   270
            Width           =   90
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "&Liberação"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   13
            Left            =   285
            TabIndex        =   58
            Top             =   270
            Width           =   705
         End
         Begin VB.Label lblBaixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Nota/Código"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   0
            Left            =   4005
            TabIndex        =   56
            Top             =   960
            Width           =   915
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "&Parcela"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   1
            Left            =   6510
            TabIndex        =   55
            Top             =   960
            Width           =   540
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "&Tipo"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   2
            Left            =   8250
            TabIndex        =   54
            Top             =   615
            Width           =   315
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "&Controle"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   3
            Left            =   7995
            TabIndex        =   53
            Top             =   1305
            Width           =   585
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "E&mpresa"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   4
            Left            =   4290
            TabIndex        =   52
            Top             =   1665
            Width           =   615
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "Vencime&nto"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   51
            Top             =   615
            Width           =   840
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "a"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   7
            Left            =   2415
            TabIndex        =   50
            Top             =   615
            Width           =   90
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "Ci&dade"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   8
            Left            =   4410
            TabIndex        =   49
            Top             =   1305
            Width           =   495
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "Es&tado"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   9
            Left            =   8085
            TabIndex        =   48
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "a"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   10
            Left            =   6315
            TabIndex        =   47
            Top             =   615
            Width           =   90
         End
         Begin VB.Label lblBaixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "V&l Original"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   11
            Left            =   4200
            TabIndex        =   46
            Top             =   615
            Width           =   705
         End
         Begin VB.Label lblBaixas 
            AutoSize        =   -1  'True
            Caption         =   "Nosso Nr"
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   12
            Left            =   7935
            TabIndex        =   45
            Top             =   270
            Width           =   660
         End
      End
      Begin VB.Label lblEmpresaUsuaria 
         AutoSize        =   -1  'True
         Caption         =   "lblEmpresaUsuaria"
         Height          =   195
         Left            =   2970
         TabIndex        =   80
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Empresa Usuária"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   105
         TabIndex        =   74
         Top             =   210
         Width           =   1200
      End
   End
   Begin VB.Label lblBaixas 
      AutoSize        =   -1  'True
      Caption         =   "a"
      ForeColor       =   &H80000002&
      Height          =   195
      Index           =   14
      Left            =   6120
      TabIndex        =   59
      Top             =   1920
      Width           =   90
   End
End
Attribute VB_Name = "frmBaixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'// Menu especial da janela de baixas
Private Const IDM_BAIXAS = 31000
Private Const IDM_BX_NOVO = 31001
Private Const IDM_BX_VIEW = 31002
Private Const IDM_BX_EDITAR = 31003
Private Const IDM_BX_PARCIAL = 31004
Private Const IDM_BX_FECHAR = 31010
Private Const IDM_BX_EMPRESAS = 31011
Private Const IDM_BX_NOTAS = 31012
Private Const BX_TOTAL = 1          'Baixa total
Private Const BX_PARCIAL = 2        'Baixa parcial
Private Const DL_MARCADO = 1        'Índice do ícone de lançamento marcado no ImageList
Private Const DL_DESMARCADO = 2     'Índice do ícone de lançamento desmarcado no ImageList
Private mstrDados         As String         'Instrução Select
Private mrstDados         As Object
Private mlngItem          As Long           'Ítem selecionado da lista
Private mlngBancos        As Long
Private mintDiasLiberacao As Integer
Private mstrPagRec        As String

'Protocolo Nr 89509 - Carlos Felippe Vernizze - 23/09/2010
Private Sub ConfiguraGrid(intSetRegistro As Integer)
    Dim strTabela As String
    
    If RegistrosSelecionados(intSetRegistro) = 1 Then
        If lvwBaixas.ListItems(intSetRegistro).SmallIcon = DL_MARCADO Then
            strTabela = TabelaRegistro(intSetRegistro)
            txtVlAcrescimo.valorMoeda = GetFieldValue("Acréscimo", strTabela, MontaClausula(intSetRegistro))
            txtVlAbatimento.valorMoeda = GetFieldValue("Abatimento", strTabela, MontaClausula(intSetRegistro))
            lblValorOriginal.Caption = FormatCurrency(lvwBaixas.ListItems(intSetRegistro).SubItems(9) - txtVlAcrescimo.valorMoeda + txtVlAbatimento.valorMoeda)
            lblVlTotal.Caption = lvwBaixas.ListItems(intSetRegistro).SubItems(9)
            cmdComfirmar.Enabled = True
            cmdBaixas(1).Enabled = True
        End If
    Else
        Call LimpaAdicionais
        cmdComfirmar.Enabled = False
        cmdBaixas(1).Enabled = False
    End If
End Sub

Private Sub cboBaixas_Click(Index As Integer)
    If Index = 0 Then
        Call SugereOperacaoContabil
    End If
End Sub

Private Sub cboBaixas_GotFocus(Index As Integer)
    Selecione cboBaixas(Index)
    DescStatus cboBaixas(Index).TabIndex
End Sub

Private Sub cmdBaixaLote_Click()
    Dim lngItem As Long
    Dim bTem    As Boolean
    
    For lngItem = 1 To lvwBaixas.ListItems.Count
        If lvwBaixas.ListItems(lngItem).SmallIcon = DL_MARCADO Then
            bTem = True
        End If
    Next
    If Not bTem Then
        MsgFunc "É necessário marcar no mínimo um registro!"
    Else
        cmdBaixas(0).Enabled = False
        cmdBaixas(1).Enabled = False
        cmdBaixaLote.Enabled = False
        fraBaixas(0).Visible = False
        fraBaixas(2).Top = 495
        fraBaixas(2).Visible = True
        fraBaixas(5).Visible = False
        edtDataPagamento.Data = Date
        edtDataLiberacao.Data = Date
        etxBancoBaixa.SetFocus
        If optBaixas(1).value = True Then
            Label1(1).Visible = False
            etxChequeBaixa.Visible = False
            cmdProximoCheque.Visible = False
        Else
            Label1(1).Visible = True
            etxChequeBaixa.Visible = True
            cmdProximoCheque.Visible = True
        End If
    End If
End Sub

Private Sub cmdBaixar_Click()
    Dim Index           As Long
    Dim lngItenMarcado  As Long
    Dim strCampos       As String
    Dim SQL             As String
    Dim SqlCheque       As String
    Dim bEDuplicata     As Boolean
    'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
    Dim dblCodigo       As Double
    Dim strEmpresa      As String
    Dim bytParcela      As String
    Dim strTipo         As String
    Dim strPagRec       As String
    Dim bBaixou         As Boolean
    Dim Conciliado      As Boolean
    Dim lngOperContabil As Long
    Dim intCont         As Integer
    Dim strParcela()    As String
    Dim intContPar      As Integer
    Dim rst             As Object
    'Projeto: 100340 - Desenv.: 142890 - Ueder Budni (11/10/2016)
    Dim bizTitulo       As New BizLancamentoDuplicata
    Dim voTitulo        As New VoLancamentoDuplicata

    
    If ValidaBaixaLote Then
        strCampos = "Pagamento= " & InverteData(edtDataPagamento.Data, True)
        strCampos = strCampos & ", Liberação = " & InverteData(edtDataLiberacao.Data, True)
        If etxBancoBaixa.valorInteiro > 0 Then
            strCampos = strCampos + ", Banco= " & etxBancoBaixa.valorInteiro
        End If
        'pt. 87029 - Moacir Pfau(21/05/2008)
        If Trim(etxControleBaixa.valorTexto) <> "" Then
            strCampos = strCampos + ", Controle =" & Quote(etxControleBaixa.valorTexto, "''")
        End If
        If etxChequeBaixa.valorInteiro > 0 Then
            strCampos = strCampos + ", Cheque =" & etxChequeBaixa.valorInteiro
        End If
        If etxBancoBaixa.valorInteiro > 0 And etxChequeBaixa.valorInteiro > 0 Then
            If AbreRecordset(rst, "SELECT * FROM CHEQUE WHERE BANCO = " & etxBancoBaixa.valorInteiro & " AND CHEQUE = " & etxChequeBaixa.valorInteiro) = WL_NORECORD Then
                SqlCheque = "INSERT INTO Cheque (Banco, Cheque) VALUES (" & etxBancoBaixa.valorInteiro & "," & etxChequeBaixa.valorInteiro & ")"
            End If
        End If
        strPagRec = IIf(optBaixas(1).value, "R", "P")
        Call ValidaTitulosAtraso
        For lngItenMarcado = lvwBaixas.ListItems.Count To 1 Step -1
            strTipo = ""
            If lvwBaixas.ListItems(lngItenMarcado).SmallIcon = DL_MARCADO Then
                'Buscando a chave do registro para baixa-lo
                'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                dblCodigo = CDblDef(lvwBaixas.ListItems(lngItenMarcado).Text)
                'pt. 00000 - Ivo Sousa (30/03/2010)
                'Alteração para baixar duplicatas de baixas parciais
                strParcela = Split(lvwBaixas.ListItems(lngItenMarcado).SubItems(2), "-")
                If Trim(strParcela(0)) = "" Then
                    bytParcela = "-" & strParcela(1)
                    For intContPar = 2 To UBound(strParcela)
                        strTipo = strTipo & strParcela(intContPar) & "-"
                    Next
                    strTipo = Left(strTipo, Len(strTipo) - 1)
                Else
                    bytParcela = strParcela(0)
                    For intContPar = 1 To UBound(strParcela)
                        strTipo = strTipo & strParcela(intContPar) & "-"
                    Next
                    strTipo = Left(strTipo, Len(strTipo) - 1)
                End If
                'pt. 79903 - Ivo Sousa(07/05/2008)
                If lvwBaixas.ListItems(lngItenMarcado).SubItems(1) = "Dupl" Then
                    bEDuplicata = True
                    If lngOperContabil <> etxOpContabilDupl.valorInteiro Then
                        If intCont = 0 Then
                            strCampos = strCampos + ", cd_operacao_baixa = " & etxOpContabilDupl.valorInteiro
                        Else
                            strCampos = Replace(strCampos, "cd_operacao_baixa = " & lngOperContabil, "cd_operacao_baixa = " & etxOpContabilDupl.valorInteiro)
                        End If
                        lngOperContabil = etxOpContabilDupl.valorInteiro
                    ElseIf intCont = 0 Then
                        strCampos = strCampos + ", cd_operacao_baixa = " & etxOpContabilDupl.valorInteiro
                    End If
                Else
                    bEDuplicata = False
                    If lngOperContabil <> etxOpContabilLanc.valorInteiro Then
                        If intCont = 0 Then
                            strCampos = strCampos + ", cd_operacao_baixa = " & etxOpContabilLanc.valorInteiro
                        Else
                            strCampos = Replace(strCampos, "cd_operacao_baixa = " & lngOperContabil, "cd_operacao_baixa = " & etxOpContabilLanc.valorInteiro)
                        End If
                        lngOperContabil = etxOpContabilLanc.valorInteiro
                    ElseIf intCont = 0 Then
                        strCampos = strCampos + ", cd_operacao_baixa = " & etxOpContabilLanc.valorInteiro
                    End If
                End If
                
                intCont = intCont + 1
                If InStr(1, strCampos, "Usuário", vbTextCompare) = 0 Then
                    strCampos = strCampos & ", Usuário = " & Quote(UserName, "''")
                End If
                
                If chkConciliado Then
                    Conciliado = True
                End If
                strEmpresa = lvwBaixas.ListItems(lngItenMarcado).SubItems(4)
                'Projeto: 100340 - Desenv.: 142888 - Ueder Budni (11/10/2016)
                Set voTitulo = bizTitulo.Carregar(IIf(strPagRec = "R", Recebimento, Pagamento), CStr(dblCodigo), strTipo, CLng(bytParcela), strEmpresa, IIf(bEDuplicata = True, Duplicata, Lancamento))
                'Quando for duplicata
                If bEDuplicata Then
                    'pt. 79700 - Ivo Sousa (12/11/2007)
                    If Conciliado Then
                        'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                        bBaixou = (ExecuteSQL("UPDATE Duplicatas SET " & strCampos & ",Conciliado = True WHERE Pagrec =" & Quote(strPagRec, "''") & " AND Nota = " & str(dblCodigo) & " AND Tipo=" & Quote(strTipo, "''") & " AND Empresa =" & Quote(strEmpresa, "''") & " AND Parcela=" & str(bytParcela)))
                    Else
                        'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                        bBaixou = (ExecuteSQL("UPDATE Duplicatas SET " & strCampos & " WHERE Pagrec =" & Quote(strPagRec, "''") & " AND Nota = " & str(dblCodigo) & " AND Tipo=" & Quote(strTipo, "''") & " AND Empresa =" & Quote(strEmpresa, "''") & " AND Parcela=" & str(bytParcela)))
                    End If
                    'Projeto: 100340 - Desenv.: 142888 - Ueder Budni (11/10/2016)
                    If bBaixou Then
                        Call RegistraLogLancDupBaixa(dblCodigo, strEmpresa, strTipo, CLng(bytParcela), strPagRec, Duplicata, voTitulo)
                    End If
                Else
                    If Conciliado Then
                        'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                        bBaixou = (ExecuteSQL("UPDATE Lançamentos SET " & strCampos & ",Conciliado = True WHERE pagrec=" & Quote(strPagRec, "''") & " AND Código=" & str(dblCodigo) & " AND Parcela=" & str(bytParcela)))
                    Else
                        'pt. 84204 - Dulcino Júnior (07/11/2007)
                        'A parcela deve ser utilizada para baixar os lançamentos, já que ela faz parte da identificação.
                        'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                        bBaixou = (ExecuteSQL("UPDATE Lançamentos SET " & strCampos & " WHERE pagrec=" & Quote(strPagRec, "''") & " AND Código=" & str(dblCodigo) & " AND Parcela=" & str(bytParcela)))
                    End If
                    'Projeto: 100340 - Desenv.: 142888 - Ueder Budni (11/10/2016)
                    If bBaixou Then
                        Call RegistraLogLancDupBaixa(dblCodigo, strEmpresa, strTipo, CLng(bytParcela), strPagRec, Lancamento, voTitulo)
                    End If
                End If
                If bBaixou Then
                  If IsValid(SqlCheque) Then ExecuteSQL SqlCheque
                  'pt. 88289 - Dulcino Júnior (15/10/2008)
                  Call BaixaRateio(lngItenMarcado)
                End If
            End If
        Next
        Call cmdBaixas_Click(0)
        etxBancoBaixa.Clear
        etxControleBaixa.Clear
        MsgFunc "Duplicata(s) baixada(s) com sucesso."
        cmdBaixas(0).Enabled = True
        fraBaixas(0).Visible = True
        fraBaixas(2).Top = 495
        fraBaixas(2).Visible = False
        fraBaixas(5).Visible = True
        'Projeto: 100340 - Desenv.: 142890 - Ueder Budni (11/10/2016)
        Set bizTitulo = Nothing
        Set voTitulo = Nothing
    End If
End Sub

Private Sub cmdBaixas_Click(Index As Integer)
    
    Select Case Index
        Case 0 'Botão Visualizar
            'Seleciona os dados filtrados pelo usuário
            If ValidaCampos Then
                SeleDocumentos
                'Se houver algum registro
                If (lvwBaixas.ListItems.Count) Then
                    lvwBaixas.SetFocus
                End If
                If optBaixas(0).value Then
                    mstrPagRec = "P"
                Else
                    mstrPagRec = "R"
                End If
                Call LimpaAdicionais
            End If
        Case 1 'Botão Editar...
            Call EditaBaixa
    End Select
End Sub

Private Sub cmdCancelar_Click()
    fraBaixas(2).Visible = False
    cmdBaixas(0).Enabled = True
    cmdBaixas(1).Enabled = True
    cmdBaixaLote.Enabled = True
    fraBaixas(0).Visible = True
    fraBaixas(5).Visible = True
End Sub

'pt. 84737 - Ivo Sousa(06/05/2008)
Private Sub cmdComfirmar_Click()
    Dim intSetRegistro As Integer
    Dim strSql         As String
    Dim strParcela     As String
    Dim strOrigem      As String
    Dim strTabela      As String
    
    If RegistrosSelecionados(intSetRegistro) = 1 Then
        strTabela = TabelaRegistro(intSetRegistro)
        If strTabela = "Duplicatas" Then
            strOrigem = "da Duplicata"
        Else
            strOrigem = "do Lançamento"
        End If
        strSql = "UPDATE " & strTabela & " SET Acréscimo = " & Replace(StrToCur(txtVlAcrescimo.valorMoeda), ",", ".") & ", Abatimento = " & Replace(StrToCur(txtVlAbatimento.valorMoeda), ",", ".") & " WHERE" & MontaClausula(intSetRegistro, strParcela)
        If MsgBox("Confirma a alteração " & strOrigem & " número " & lvwBaixas.ListItems(intSetRegistro) & " parcela " & strParcela & " ?", vbQuestion + vbYesNo) = vbYes Then
            If ExecuteSQL(strSql) Then
                MsgBox "Registro atualizado com sucesso.", vbInformation + vbOKOnly, NomeModulo
                'Se o valor que esta na grid for diferente do que o confirmado, altera na grid
                If lvwBaixas.ListItems(intSetRegistro).ListSubItems(9) <> lblVlTotal.Caption Then
                    lvwBaixas.ListItems(intSetRegistro).ListSubItems(9) = lblVlTotal.Caption
                    'Soma os registros que estão na grid novamente
                    Call SomaRegistros
                End If
                SeleDocumentos
            End If
        End If
    End If
End Sub

Private Sub cmdNenhum_Click()
    Dim i As Long

    For i = 1 To Me.lvwBaixas.ListItems.Count
        XMarkRules i, DL_DESMARCADO
    Next
    Call LimpaAdicionais
    Call SomaRegistros
End Sub

Private Sub cmdSair_Click()
    fraBaixas(2).Visible = False
    cmdBaixas(0).Enabled = True
    cmdBaixas(1).Enabled = True
    cmdBaixaLote.Enabled = True
    fraBaixas(0).Visible = True
    fraBaixas(5).Visible = True
End Sub

'Incluido por Edilberto Conforme protocolo 71707
Private Sub cmdProximoCheque_Click()
    Dim rstProximoCheque     As Object
  
    If AbreRecordset(rstProximoCheque, "Select * from Cheque " & _
       "WHERE Banco = " & etxBancoBaixa.valorInteiro & " AND Situação = 'Normal' " & _
       "AND (Cheque not in (Select Cheque from Duplicatas where Banco = Cheque.Banco) " & _
       "AND Cheque not in (Select Cheque from Lançamentos where Banco = Cheque.Banco) " & _
       "AND Cheque not in (Select Cheque from [Transf Bancária] where Banco = Cheque.Banco)) " & _
       "ORDER BY Cheque ASC", dbOpenSnapshot) = WL_OK Then
        etxChequeBaixa.valorInteiro = GetValue(rstProximoCheque, "Cheque", ZERO)
    Else
        etxChequeBaixa.valorInteiro = ProximoNumero("Cheque", "Cheque", "Banco = " & etxBancoBaixa.valorInteiro)
    End If
    Call FechaRecordset(rstProximoCheque)
End Sub

Private Sub cmdSair2_Click()
    Call LibProc(WL_SAIR)
End Sub

Private Sub cmdTodos_Click()
    Dim i As Long

    For i = 1 To Me.lvwBaixas.ListItems.Count
        If ValidaSelecao(i, False, False) Then
            XMarkRules i, DL_MARCADO
        End If
    Next
    SomaRegistros
    Call LimpaAdicionais
End Sub

Private Sub edtDataPagamento_LostFocus()
    If Not IsEmptyDate(edtDataPagamento.Data) Then
        If mintDiasLiberacao > 0 Then
            edtDataLiberacao.Data = DateAdd("d", mintDiasLiberacao, edtDataPagamento.Data)
        Else
            edtDataLiberacao.Data = edtDataPagamento.Data
        End If
    End If
End Sub

Private Sub etxBancoBaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxBancoBaixa.ValorDescricao = "" Then
            etxBancoBaixa.valorInteiro = 0
        End If
        Call PCampo("Bancos", "Select * From Bancos", pbCampo, etxBancoBaixa, 0)
    End If
End Sub

'Data: 27/03/2008.
'Conforme reunião de corredor com a consultoria(Carlos Dias), o banco que deve ser utilizado
'como padrão para sugestão da data de liberação no caso das duplicatas ou lançamentos a receber, é o banco informado
'na baixa, sendo assim, a consulta deve ser feita ao sair do campo banco.
Private Sub etxBancoBaixa_LostFocus()
    Dim strSql   As String
    Dim rstBanco As Object

    'pt. 86113 - Dulcino Júnior (27/03/2008)
    If optBaixas(1).value Then
        If etxBancoBaixa.ValorDescricao <> "" Then
            strSql = "SELECT [Dias para Liberação] FROM Bancos WHERE Banco=" & etxBancoBaixa.valorInteiro
            If AbreRecordset(rstBanco, strSql) = WL_OK Then
                mintDiasLiberacao = rstBanco.Fields("Dias para Liberação").value
                edtDataLiberacao.Data = DateAdd("d", mintDiasLiberacao, edtDataPagamento.Data)
            End If
            Call FechaRecordset(rstBanco)
        End If
    End If
End Sub

Private Sub etxEmpresas_GotFocus()
    DescStatus etxEmpresas.TabIndex
End Sub

Private Sub etxEmpresas_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxEmpresas.ValorDescricao = "" Then
            etxEmpresas.valorTexto = ""
        End If
        Call ConsultaEmpresas
    End If
End Sub

Private Sub etxEstado_GotFocus()
    DescStatus etxEstado.TabIndex
End Sub

Private Sub etxEstado_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxEstado.ValorDescricao = "" Then
            etxEstado.valorTexto = ""
        End If
        Call PCampo("Estados", "Estados", pbCampo, etxEstado, "Sigla")
    End If
End Sub

Private Sub etxNossoNumero_GotFocus()
    DescStatus etxNossoNumero.TabIndex
End Sub

Private Sub etxNumero_GotFocus()
    DescStatus etxNumero.TabIndex
End Sub

Private Sub etxNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call ConsultaNotas
    End If
End Sub

Private Sub etxParcela_GotFocus()
    DescStatus etxParcela.TabIndex
End Sub

Private Sub etxValorOriginalInicial_GotFocus()
    DescStatus etxValorOriginalInicial.TabIndex
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

Private Sub Form_Load()
    Dim strOpcoes As String

    Call etxBancoInicial.AddConexao(Aplicacao)
    Call etxBancoFinal.AddConexao(Aplicacao)
    Call etxContaInicial.AddConexao(Aplicacao)
    Call etxContaFinal.AddConexao(Aplicacao)
    Call etxCentroCustoInicial.AddConexao(Aplicacao)
    Call etxCentroCustoFinal.AddConexao(Aplicacao)
    Call etxEmpresas.AddConexao(Aplicacao)
    Call etxEstado.AddConexao(Aplicacao)
    Call etxBancoBaixa.AddConexao(Aplicacao)
    Call etxOpContabilDupl.AddConexao(Aplicacao)
    Call etxOpContabilLanc.AddConexao(Aplicacao)
    cmdBaixaLote.Enabled = False
    
    'Preenchendo as caixas de combinação com os tipos de duplicatas existentes.
    strOpcoes = "SELECT Texto FROM Opções WHERE Rotina = 'Dupl. a Pagar';"
    ComboAddItem cboBaixas(0), strOpcoes, "Texto"
    
    'pt. 84490 - Dulcino Júnior (29/11/2007)
    'Correção da sugestão de Operação contábil
    etxOpContabilDupl.Clear
    etxOpContabilLanc.Clear
    cboBaixas(0).AddItem "Todos"
    cboBaixas(0).Text = "Todos"
    
    cboBaixas(1).ListIndex = 0
    
    'pt. 81189 - Dulcino Júnior
    'Integração Contábil
    Label6.Enabled = Configuracao("Utiliza Integração Contábil", False)
    Label14.Enabled = Configuracao("Utiliza Integração Contábil", False)
    etxOpContabilDupl.Enabled = Configuracao("Utiliza Integração Contábil", False)
    etxOpContabilLanc.Enabled = Configuracao("Utiliza Integração Contábil", False)
    Call LimpaCampos
    If mstrPagRecBaixas = "P" Then
        optBaixas(0).value = True
        frmBaixas.Caption = "Baixas - Contas à Pagar"
        optBaixas_GotFocus (0)
    Else
        optBaixas(1).value = True
        frmBaixas.Caption = "Baixas - Contas à Receber"
        optBaixas_GotFocus (1)
    End If
    cmdComfirmar.Enabled = False
    cmdBaixas(1).Enabled = False
    lblVlTotal.Caption = FormatCurrency(0)
    lblValorOriginal.Caption = FormatCurrency(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBaixas = Nothing
End Sub

Private Sub lvwBaixas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwBaixas.Sorted = True
    lvwBaixas.SortKey = (ColumnHeader.Index - 1)
    lvwBaixas.Sorted = False
End Sub

Private Sub lvwBaixas_DblClick()
    Dim intSetRegistro As Integer
    Dim strTabela      As String
    
On Error GoTo Error_Handler
    If ValidaSelecao(mlngItem, True, False) Then
        DoEvents
        XMark mlngItem
        
        'Protocolo Nr 89509 - Carlos Felippe Vernizze - 23/09/2010
        Call ConfiguraGrid(intSetRegistro)
        
    End If
    SomaRegistros
    Exit Sub
Error_Handler:
    err.Clear
End Sub

Private Sub lvwBaixas_GotFocus()
    'Mensagem descritiva da barra de status
    MsgBar LoadResString(150)
End Sub

Private Sub lvwBaixas_ItemClick(ByVal item As MSComctlLib.ListItem)
    mlngItem = item.Index
End Sub

Private Sub lvwBaixas_KeyPress(KeyAscii As Integer)
    Dim intSetRegistro As Integer
    
    If KeyAscii = vbKeySpace Then
        'Marcar e somar os itens
        DoEvents
        'Protocolo Nr 89509 - Carlos Felippe Vernizze - 23/09/2010
        If lvwBaixas.ListItems.Count > 0 Then
            If ValidaSelecao(mlngItem, True, False) Then
                XMark mlngItem
            End If
        End If
        SomaRegistros
        
        'Protocolo Nr 89509 - Carlos Felippe Vernizze - 23/09/2010
        Call ConfiguraGrid(intSetRegistro)
        
    ElseIf ((KeyAscii = Asc("P")) Or (KeyAscii = Asc("p"))) Then
        If RegistrosSelecionados(intSetRegistro) = 0 Then
            MsgBox "Selecione um item!", vbInformation, NomeModulo
        Else
            If (lvwBaixas.ListItems(intSetRegistro).SmallIcon) Then
                If ValidaSelecao(mlngItem, True, True) Then
                    EditaBaixaParcial
                End If
            End If
        End If
    End If
End Sub

Private Sub lvwBaixas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngItens As Long

    'Verifica se o usuário está clicando sobre um ítem da lista
    If Button = vbRightButton Then
        For lngItens = 1 To lvwBaixas.ListItems.Count
            If (X > lvwBaixas.ListItems(lngItens).Left) And (X < lvwBaixas.ListItems(lngItens).Width) Then
                If (Y > lvwBaixas.ListItems(lngItens).Top) And (Y < (lvwBaixas.ListItems(lngItens).Height + lvwBaixas.ListItems(lngItens).Top)) Then
                    lvwBaixas.ListItems(lngItens).Selected = True
                    mlngItem = lngItens
                    lvwBaixas.Refresh
                    Exit For
                End If
            End If
        Next
        If mlngItem = 0 Then
            mlngItem = 1
        End If
    End If
End Sub

'SUB.......: ConsultaNotas
'Objetivo..: Abre a jenala de pesquisa para Duplicatas ou Lançamentos.
Private Sub ConsultaNotas()
    Dim strCodigo     As String
    Dim strVarExpDupl As String
    Dim strVarExpLanc As String
    
    'Select geral
    strCodigo = "SELECT Nota, Código, Tipo, Parcela, Empresa, Descrição, " & _
                "[Valor Original], Acréscimo, Abatimento, Controle, " & _
                "Situação FROM <Tabela> WHERE PagRec = '" & _
                IIf(optBaixas(0).value, "P", "R") & "' AND (Pagamento IS NULL);"

    If optDup.value Then
        DeleteStr strCodigo, ", Código"
        'Define o nome da tabela
        InsereStr strCodigo, "Duplicatas", DeleteStr(strCodigo, "<Tabela>")
        Call PCampo("Duplicatas", strCodigo, pbCampo, etxNumero, "Nota")
    ElseIf optLanc.value Then
        DeleteStr strCodigo, "Nota, "
        DeleteStr strCodigo, "Parcela, "
        'Define o nome da tabela
        InsereStr strCodigo, "Lançamentos", DeleteStr(strCodigo, "<Tabela>")
        Call PCampo("Lançamentos", strCodigo, pbCampo, etxNumero, "Código")
    Else
        'pt. 79903 - Ivo Sousa(08/05/2008)
        Call ResolveExpDuplLanc(strVarExpDupl, strVarExpLanc, True)
        strCodigo = Replace("(" & strVarExpDupl & ") UNION (" & strVarExpLanc & ") ORDER BY " & getOrderBy, "cod_id", "Numero")
        Call PCampo("Lançamentos/Duplicatas", strCodigo, pbCampo, etxNumero, 1)
    End If
End Sub

'SUB.......: ConsultaEmpresas
'Objetivo..: Abre a janela de pesquisa para empresas.
Private Sub ConsultaEmpresas()
    Dim strExpEmpresa As String

    strExpEmpresa = "SELECT Apel, Razão, Tipo, Pessoa, [CNPJ/CPF], [IEst/RG], " & _
                    "Ramo, Endereço, Bairro, CEP, Cidade, Estado, Região, País, " & _
                    "Fone1, Ramal1, Contato, Dpto FROM Empresas WHERE "
    'Verifica se deve separar as empresas por tipo
    If LerArquivoASCII("KinSys", "Separar Empresa por Tipo", gstrTempSys) = "S" Then
        If optBaixas(0).value Then
            AppendStr strExpEmpresa, "Tipo <> 'Fornecedor'"
        Else
            AppendStr strExpEmpresa, "Tipo <> 'Cliente'"
        End If
    Else
        InsereStr strExpEmpresa, "", DeleteStr(strExpEmpresa, " WHERE ")
    End If
    Call PCampo("Empresas Ativas", strExpEmpresa, pbCampo, etxEmpresas, 0)
End Sub

Private Sub optBaixas_Click(Index As Integer)
    optBaixas_GotFocus (Index)
End Sub

Private Sub optBaixas_GotFocus(Index As Integer)
    Dim strOpcoes As String
  
    DescStatus optBaixas(Index).TabIndex
    strOpcoes = "SELECT Texto FROM Opções WHERE Rotina = '" & IIf(Left$(optBaixas(Index).Caption, 11) = "Lançamentos", OPT_LANCAMENTOS, OPT_DUPLICATAS) & "'"
    cboBaixas(0).Clear
    ComboAddItem cboBaixas(0), strOpcoes, "Texto"
    cboBaixas(0).AddItem "Todos"
    cboBaixas(0).Text = "Todos"
    If Index = 1 Then
        Label1(1).Visible = False
        etxChequeBaixa.Visible = False
        cmdProximoCheque.Visible = False
    Else
        Label1(1).Visible = True
        etxChequeBaixa.Visible = True
        cmdProximoCheque.Visible = True
    End If
End Sub

Private Sub optControle_Click()
    cboBaixas(1).Text = "Controle"
End Sub

Private Sub optDup_Click()
    cboBaixas(1).List(0) = "Nota"
    optBaixas_GotFocus (0)
    Me.txtTotalSelecionados.Caption = ""
    Me.txtQtlSelecionados.Caption = ""
    Me.txtTotalListados.Caption = ""
    lvwBaixas.ListItems.Clear
    cmdBaixaLote.Enabled = False
End Sub

Private Sub optEmissao_Click()
    cboBaixas(1).Text = "Emissão"
End Sub

Private Sub optEmpresa_Click()
    cboBaixas(1).Text = "Empresa"
End Sub

Private Sub optLanc_Click()
    cboBaixas(1).List(0) = "Código"
    optBaixas_GotFocus (0)
    Me.txtTotalSelecionados.Caption = ""
    Me.txtQtlSelecionados.Caption = ""
    Me.txtTotalListados.Caption = ""
    lvwBaixas.ListItems.Clear
    cmdBaixaLote.Enabled = False
End Sub

Private Sub optNotaCOd_Click()
    If optDup.value Then
        cboBaixas(1).Text = "Nota"
    ElseIf optLanc.value Then
        cboBaixas(1).Text = "Código"
    Else
        cboBaixas(1).Text = "Código/Nota"
    End If
End Sub

Private Sub optTodos_Click()
    cboBaixas(1).List(0) = "Código/Nota"
    optBaixas_GotFocus (0)
    Me.txtTotalSelecionados.Caption = ""
    Me.txtQtlSelecionados.Caption = ""
    Me.txtTotalListados.Caption = ""
    lvwBaixas.ListItems.Clear
    cmdBaixaLote.Enabled = False
End Sub

Private Sub optVenc_Click()
    cboBaixas(1).Text = "Vencimento"
End Sub

Private Sub txtBaixas_GotFocus(Index As Integer)
    Selecione txtBaixas(Index)
    DescStatus txtBaixas(Index).TabIndex
End Sub

'SUB.......: DescStatus
'Objetivo..: Escreve descrições para os campos do formulário na barra de status.
'Argumento.: [iTabIndex]: TabIndex do controle que recebe o foco.
Private Sub DescStatus(iTabIndex As Integer)
    Select Case iTabIndex
        Case 1 'Campo de Liberação Inicial
            MsgBar "Data de Liberação inicial"
        Case 2 'Campo de Liberação Final
            MsgBar "Data de Liberação final"
        Case 3 'Campo de Vencimento Inicial
            MsgBar "Data de Vencimento inicial"
        Case 4 'Campo de Vencimento Final
            MsgBar "Data de Vencimento final"
        Case 5 'Campo de Emissão Inicial
            MsgBar "Data de Emissão inicial"
        Case 6 'Campo de Emissão Final
            MsgBar "Data de Emissão final"
        Case 7 'Campo de Banco inicial
            MsgBar "Código do banco inicial"
        Case 8 'Campo de Banco final
            MsgBar "Código do banco final"
        Case 9 'Campo de Conta inicial
            MsgBar "Código da conta inicial"
        Case 10 'Campo de Conta final
            MsgBar "Código da conta final"
        Case 11 'Campo de Centro de Custo inicial
            MsgBar "Código do centro de custo inicial"
        Case 12 'Campo de Centro de Custo final
            MsgBar "Código do centro de custo final"
        Case 13 'Campo de valor Original inicial
            MsgBar "Valor Originial inicial"
        Case 14 'Campo de valor Original final
            MsgBar "Valor Originial final"
        Case 15 'Campo Nota ou Código
            If optDup.value = True Then
                MsgBar "Número da Nota" & ResolveResString(75, resUM, "Duplicatas")
            Else
                MsgBar "Código do Lançamento" & ResolveResString(75, resUM, "Lançamentos")
            End If
        Case 16 'Campo Parcela
            MsgBar "Número da Parcela"
        Case 17 'Campo Cidade
            MsgBar ""
        Case 18 'Campo Empresa
            MsgBar "Nome Fantasia da Empresa" & ResolveResString(75, resUM, "Empresas Ativas")
        Case 19 'Campo Nosso número
            MsgBar "Nosso número"
        Case 20 'Campo Tipo
            MsgBar "Tipo do documento"
        Case 21 'Campo estado
            MsgBar ""
        Case 22 'Campo Controle
            MsgBar "Código de controle do documento"
'        Case 13 'Campo Ordem
'            MsgBar "Ordem para a apresentação dos dados"
'        Case 19, 20 'Botões de Opção: A pagar ou a Receber
'            MsgBar "Espécie de documento que será apresentada"
    End Select
End Sub

'SUB.......: SeleDocumentos
'Objetivo..: Inicia a expressão de consulta para preencher o controle ListView
'            com as Duplicatas ou Lançamentos escolhidos pelo usuário.
Private Sub SeleDocumentos()
    Dim strSeleDocs As String 'Para a expressão Select
    Dim lngIndice   As Long   'Índice do registro no controle
    Dim bolDupl     As Boolean
    Dim strWhere    As String
    
    'Verificando as caixas combo
    If Len(cboBaixas(1).Text) > 0 Then
        If IndexOf(cboBaixas(1).Text, cboBaixas(1)) = NENHUM Then
            MsgBox ResolveResString(12, resUM, "Ordem"), vbInformation, NomeModulo
            Exit Sub
        End If
    End If

    SetPtr vbHourglass
    SimpleMsgBar LoadResString(13) & LoadResString(14)
    'Resolve a expressão de consulta
    'pt. 79903 - Ivo Sousa(07/05/2008)
    Call ResolveExp(strSeleDocs)
    lvwBaixas.ListItems.Clear
    mlngItem = 0 'Nenhum item selecionado
    If AbreRecordset(mrstDados, strSeleDocs) = WL_OK Then
        bolDupl = (optDup.value)
        Do Until mrstDados.EOF
            Inc lngIndice
            lvwBaixas.ListItems.add lngIndice, , StrZero(GetValue(mrstDados, 1), 6)
            
            'pt. 79903 - Ivo Sousa(07/05/2008)
            If GetValue(mrstDados, "Origem") <> "" Then
                lvwBaixas.ListItems(lngIndice).SubItems(1) = GetValue(mrstDados, "Origem")
            Else
                If optDup.value Then
                    lvwBaixas.ListItems(lngIndice).SubItems(1) = "Dupl"
                Else
                    lvwBaixas.ListItems(lngIndice).SubItems(1) = "Lanc"
                End If
            End If
            lvwBaixas.ListItems(lngIndice).SubItems(2) = StrZero(GetValue(mrstDados, "Parcela"), 2) & "-" & GetValue(mrstDados, "Tipo")
            lvwBaixas.ListItems(lngIndice).SubItems(3) = GetValue(mrstDados, "Descrição")
            lvwBaixas.ListItems(lngIndice).SubItems(4) = GetValue(mrstDados, "Empresa", NUL)
            lvwBaixas.ListItems(lngIndice).SubItems(5) = GetValue(mrstDados, "Banco", NUL)
            lvwBaixas.ListItems(lngIndice).SubItems(6) = GetValue(mrstDados, "Conta", NUL)
            lvwBaixas.ListItems(lngIndice).SubItems(7) = GetValue(mrstDados, "Centro", NUL)
            lvwBaixas.ListItems(lngIndice).SubItems(8) = GetValue(mrstDados, "Vencimento", NUL)
            lvwBaixas.ListItems(lngIndice).SubItems(9) = Format$(Kif_Valor(mrstDados), FCURRENCY)
            lvwBaixas.ListItems(lngIndice).SubItems(10) = GetValue(mrstDados, "Emissão", NUL)
            lvwBaixas.ListItems(lngIndice).SubItems(11) = GetValue(mrstDados, "Controle")
            
            'pt. 79903 - Ivo Sousa(07/05/2008)
            If optTodos.value Then
                Select Case GetValue(mrstDados, "Origem")
                    Case "Dupl"
                        Call ExecuteSQL("UPDATE Duplicatas SET Marcação = False WHERE PagRec = '" & GetValue(mrstDados, "PagRec") & "' AND Nota = " & GetValue(mrstDados, "cod_id") & " AND Parcela = " & GetValue(mrstDados, "Parcela"))
                    Case "Lanc"
                        Call ExecuteSQL("UPDATE Lançamentos SET Marcação = False WHERE PagRec = '" & GetValue(mrstDados, "PagRec") & "' AND Código = " & GetValue(mrstDados, "cod_id") & " AND Parcela = " & GetValue(mrstDados, "Parcela"))
                End Select
            Else
                'pt. 77398 - Alisson Ricardo
                'Foi pedido para listar as dupl/lanc desmarcadas ,e não alteramos o campo marcação da tabela
                'Pt. 95368 - Moacir Pfau(03/11/2009)
                'If gTipoDB = Access Then
                '    mrstDados.Edit
                'End If
                mrstDados("Marcação") = False
                mrstDados.update
            End If
'            If (GetValue(mrstDados, "Marcação")) Then
'                lvwBaixas.ListItems(lngIndice).SmallIcon = DL_MARCADO
'            Else
            lvwBaixas.ListItems(lngIndice).SmallIcon = DL_DESMARCADO
'            End If
            mrstDados.MoveNext
        Loop
        lvwBaixas.ListItems(1).Selected = True
        mrstDados.MoveFirst
    End If
    SomaRegistros
    mstrDados = strSeleDocs 'Guarda a instrução para ser utilizada posteriormente
    'Projeto: #218 - História: #268 - Desenvolvimento#652 - Moacir Pfau(24/09/2012)
    'FechaRecordset mrstDados
    Call LimpaAdicionais
    SetPtr vbDefault
End Sub

Private Sub SomaRegistros()
    Dim curMarcado                   As Currency
    Dim curDesmarcado                As Currency
    Dim curTotal                     As Currency
    Dim intQtdMarcada                As Integer
    Dim intQtdDesmarcada             As Integer
    Dim intQtdTotal                  As Integer
    Dim nCont                        As Integer
    Dim QtdeTotalTitulosListados     As Integer
  
    curMarcado = 0
    curDesmarcado = 0
    intQtdMarcada = 0
    intQtdDesmarcada = 0
    intQtdTotal = 0
    
    For nCont = 1 To lvwBaixas.ListItems.Count
        If (lvwBaixas.ListItems(nCont).SmallIcon = DL_MARCADO) Then
            curMarcado = curMarcado + CCurDef(lvwBaixas.ListItems(nCont).SubItems(9)) 'pt. 86458 - Moacir Pfau(09/04/2008)
            intQtdMarcada = intQtdMarcada + 1
        Else
            curDesmarcado = curDesmarcado + CCurDef(lvwBaixas.ListItems(nCont).SubItems(9))
            intQtdDesmarcada = intQtdDesmarcada + 1
        End If
        QtdeTotalTitulosListados = QtdeTotalTitulosListados + 1
    Next
    If curMarcado <> 0 Then
        curTotal = curMarcado
        intQtdTotal = intQtdMarcada
    Else
        curTotal = curDesmarcado
        intQtdTotal = intQtdDesmarcada
    End If
    txtBaixas(6).Text = wsprintf("Total: %C", curTotal)
    txtBaixas(7).Text = wsprintf("%l %s", intQtdTotal, IIf((optDup.value = True), "Duplicata(s)", "Lançamentos(s)"))
    'pt. 77398
    Me.txtTotalSelecionados.Caption = FormatCurrency(curMarcado)
    Me.txtQtlSelecionados.Caption = intQtdMarcada
    Me.txtTotalListados.Caption = QtdeTotalTitulosListados
End Sub

'SUB.......: ResolveExp
'Objetivo..: Resolve a expressão final de consulta.
'Argumento.: [strVarExpLanc]: Variável que receberá a expressão.
Private Sub ResolveExp(strVarExp As String)
    Dim dblValorIni    As Double
    Dim dblValorFin    As Double
    Dim strVarExpLanc  As String
    Dim strVarExpDupl  As String
    
    'Verificando qual o tipo da baixa, iniciando a expressão
    If optDup.value Then
        Call ResolveExpDupl(strVarExpDupl)
    ElseIf optLanc.value Then
        Call ResolveExpLancto(strVarExpLanc)
    Else
        Call ResolveExpDuplLanc(strVarExpLanc, strVarExpDupl)
    End If
    
    'Tipo da duplicata
    If Not CompStr(cboBaixas(0).Text, "Todos") And Len(cboBaixas(0).Text) > 0 Then
        Concat strVarExpLanc, " AND Tipo = '", cboBaixas(0).Text, "'"
        Concat strVarExpDupl, " AND Tipo = '", cboBaixas(0).Text, "'"
    End If
    
    'Empresa
    If etxEmpresas.valorTexto <> "" Then
        Concat strVarExpLanc, " AND Empresa = '", etxEmpresas.valorTexto, "'"
        Concat strVarExpDupl, " AND Empresa = '", etxEmpresas.valorTexto, "'"
    End If
    
    'Controle
    If etxControle.valorTexto <> "" Then
        Concat strVarExpLanc, " AND Controle = '", etxControle.valorTexto, "'"
        Concat strVarExpDupl, " AND Controle = '", etxControle.valorTexto, "'"
    End If

    'Filtrando Vencimento
    If Not IsEmptyDate(edtDataVencimentoInicial.Data) And Not IsEmptyDate(edtDataVencimentoFinal.Data) Then
        Concat strVarExpLanc, " AND Vencimento BETWEEN ", InverteData(edtDataVencimentoInicial.Data, True), " AND ", InverteData(edtDataVencimentoFinal.Data, True)
        Concat strVarExpDupl, " AND Vencimento BETWEEN ", InverteData(edtDataVencimentoInicial.Data, True), " AND ", InverteData(edtDataVencimentoFinal.Data, True)
    ElseIf Not IsEmptyDate(edtDataVencimentoInicial.Data) Then
        Concat strVarExpLanc, " AND Vencimento >= ", InverteData(edtDataVencimentoInicial.Data, True)
        Concat strVarExpDupl, " AND Vencimento >= ", InverteData(edtDataVencimentoInicial.Data, True)
    ElseIf Not IsEmptyDate(edtDataVencimentoFinal.Data) Then
        Concat strVarExpLanc, " AND Vencimento <= ", InverteData(edtDataVencimentoFinal.Data, True)
        Concat strVarExpDupl, " AND Vencimento <= ", InverteData(edtDataVencimentoFinal.Data, True)
    End If
  
    'Filtrando Liberação
    If Not IsEmptyDate(edtDataLiberacaoInicial.Data) And Not IsEmptyDate(edtDataLiberacaoFinal.Data) Then
        Concat strVarExpLanc, " AND Liberação BETWEEN ", InverteData(edtDataLiberacaoInicial.Data, True), " AND ", InverteData(edtDataLiberacaoFinal.Data, True)
        Concat strVarExpDupl, " AND Liberação BETWEEN ", InverteData(edtDataLiberacaoInicial.Data, True), " AND ", InverteData(edtDataLiberacaoFinal.Data, True)
    ElseIf Not IsEmptyDate(edtDataLiberacaoInicial.Data) Then
        Concat strVarExpLanc, " AND Liberação >= ", InverteData(edtDataLiberacaoInicial.Data, True)
        Concat strVarExpDupl, " AND Liberação >= ", InverteData(edtDataLiberacaoInicial.Data, True)
    ElseIf Not IsEmptyDate(edtDataLiberacaoFinal.Data) Then
        Concat strVarExpLanc, " AND Liberação <= ", InverteData(edtDataLiberacaoFinal.Data, True)
        Concat strVarExpDupl, " AND Liberação <= ", InverteData(edtDataLiberacaoFinal.Data, True)
    End If

    'Filtrando Emissão
    If Not IsEmptyDate(edtDataEmissaoInicial.Data) And Not IsEmptyDate(edtDataEmissaoFinal.Data) Then
        Concat strVarExpLanc, " AND Emissão BETWEEN ", InverteData(edtDataEmissaoInicial.Data, True), " AND ", InverteData(edtDataEmissaoFinal.Data, True)
        Concat strVarExpDupl, " AND Emissão BETWEEN ", InverteData(edtDataEmissaoInicial.Data, True), " AND ", InverteData(edtDataEmissaoFinal.Data, True)
    ElseIf Not IsEmptyDate(edtDataEmissaoInicial.Data) Then
        Concat strVarExpLanc, " AND Emissão >= ", InverteData(edtDataEmissaoInicial.Data, True)
        Concat strVarExpDupl, " AND Emissão >= ", InverteData(edtDataEmissaoInicial.Data, True)
    ElseIf Not IsEmptyDate(edtDataEmissaoFinal.Data) Then
        Concat strVarExpLanc, " AND Emissão <= ", InverteData(edtDataEmissaoFinal.Data, True)
        Concat strVarExpDupl, " AND Emissão <= ", InverteData(edtDataEmissaoFinal.Data, True)
    End If

    'Filtrando entre pagos e recebidos
    If optBaixas(0).value Then
        AppendStr strVarExpLanc, " AND PagRec = 'P'"
        AppendStr strVarExpDupl, " AND PagRec = 'P'"
    Else
        AppendStr strVarExpLanc, " AND PagRec = 'R'"
        AppendStr strVarExpDupl, " AND PagRec = 'R'"
    End If
    
    'pt. 86113 - Dulcino Júnior(25/03/2008)
    If etxBancoInicial.valorInteiro > 0 Then
        AppendStr strVarExpLanc, " AND Banco >=" & etxBancoInicial.valorInteiro
        AppendStr strVarExpDupl, " AND Banco >=" & etxBancoInicial.valorInteiro
    End If
    If etxBancoFinal.valorInteiro > 0 Then
        AppendStr strVarExpLanc, " AND Banco <=" & etxBancoFinal.valorInteiro
        AppendStr strVarExpDupl, " AND Banco <=" & etxBancoFinal.valorInteiro
    End If
    
    'pt. 86113 - Dulcino Júnior(25/03/2008)
    If etxContaInicial.valorInteiro > 0 Then
        AppendStr strVarExpLanc, " AND Conta >=" & etxContaInicial.valorInteiro
        AppendStr strVarExpDupl, " AND Conta >=" & etxContaInicial.valorInteiro
    End If
    If etxContaFinal.valorInteiro > 0 Then
        AppendStr strVarExpLanc, " AND Conta <=" & etxContaFinal.valorInteiro
        AppendStr strVarExpDupl, " AND Conta <=" & etxContaFinal.valorInteiro
    End If
    
    'pt. 86113 - Dulcino Júnior(25/03/2008)
    If etxCentroCustoInicial.valorInteiro > 0 Then
        AppendStr strVarExpLanc, " AND Centro >=" & etxCentroCustoInicial.valorInteiro
        AppendStr strVarExpDupl, " AND Centro >=" & etxCentroCustoInicial.valorInteiro
    End If
    If etxCentroCustoFinal.valorInteiro > 0 Then
        AppendStr strVarExpLanc, " AND Centro <=" & etxCentroCustoFinal.valorInteiro
        AppendStr strVarExpDupl, " AND Centro <=" & etxCentroCustoFinal.valorInteiro
    End If
    
    'Especificando apenas os registros não pagos
    AppendStr strVarExpLanc, " AND (Pagamento IS NULL)"
    AppendStr strVarExpDupl, " AND (Pagamento IS NULL)"

    'Especificando Cidade e Estado
    If etxCidade.valorTexto <> "" Then
        Concat strVarExpLanc, " AND (Select Cidade from Empresas where Apel = Empresa) = " & Quote(etxCidade.valorTexto, "'")
        Concat strVarExpDupl, " AND (Select Cidade from Empresas where Apel = Empresa) = " & Quote(etxCidade.valorTexto, "'")
    End If
    If etxEstado.valorTexto <> "" Then
        Concat strVarExpLanc, " AND (Select Estado from Empresas where Apel = Empresa) = " & Quote(etxEmpresas.valorTexto, "'")
        Concat strVarExpDupl, " AND (Select Estado from Empresas where Apel = Empresa) = " & Quote(etxEmpresas.valorTexto, "'")
    End If
  
    'Filtro Valor Original
    If (etxValorOriginalInicial.valorMoeda > 0) And (etxValorOriginalFinal.valorMoeda > 0) Then
        Concat strVarExpLanc, " AND [Valor Original] BETWEEN ", Replace(etxValorOriginalInicial.valorMoeda, ",", "."), " AND ", Replace(etxValorOriginalFinal.valorMoeda, ",", ".")
        Concat strVarExpDupl, " AND [Valor Original] BETWEEN ", Replace(etxValorOriginalInicial.valorMoeda, ",", "."), " AND ", Replace(etxValorOriginalFinal.valorMoeda, ",", ".")
    ElseIf etxValorOriginalInicial.valorMoeda > 0 Then
        Concat strVarExpLanc, " AND [Valor Original] >= ", Replace(etxValorOriginalInicial.valorMoeda, ",", ".")
        Concat strVarExpDupl, " AND [Valor Original] >= ", Replace(etxValorOriginalInicial.valorMoeda, ",", ".")
    ElseIf etxValorOriginalFinal.valorMoeda > 0 Then
        Concat strVarExpLanc, " AND [Valor Original] <= ", Replace(etxValorOriginalFinal.valorMoeda, ",", ".")
        Concat strVarExpDupl, " AND [Valor Original] <= ", Replace(etxValorOriginalFinal.valorMoeda, ",", ".")
    End If
  
    'filtrando o nosso numero
    If etxNossoNumero.valorInteiro > 0 Then
        'Projeto: #4350 - História: #4336 - Desenvolvimento: #5286 - Ivo Sousa(26/02/2013)
        Concat strVarExpLanc, " AND SeqNossoNumero = '" & etxNossoNumero.valorTexto & "'"
        Concat strVarExpDupl, " AND SeqNossoNumero = '" & etxNossoNumero.valorTexto & "'"
    End If
    If optDup.value Then
        Concat strVarExpDupl, " ORDER BY ", getOrderBy, ";"
        strVarExp = strVarExpDupl
    ElseIf optLanc.value Then
        'pt. 80029
        'Receber expressao Order By conforme o OptioButton Selecionado
        Concat strVarExpLanc, " ORDER BY ", getOrderBy, ";"
        strVarExp = strVarExpLanc
    Else
        strVarExp = "(" & strVarExpDupl & ") UNION (" & strVarExpLanc & ") ORDER BY " & getOrderBy
    End If
End Sub

'FUNCTION..: ResolveExpDupl
'Objetivo..: Cria a expressão que seleciona os dados de duplicatas para as baixas.
'Argumento.: [strRetorno]: Variável que irá receber a expressão.
Private Sub ResolveExpDupl(strRetorno As String, Optional blnEditar As Boolean = False)
    
    'pt. 87216 - Ivo Sousa(03/06/2008)
    If blnEditar Then
        strRetorno = "SELECT * FROM Duplicatas WHERE "
    Else
        strRetorno = "SELECT 'Dupl' AS Origem, Nota as cod_id, Parcela, Tipo, Empresa, Descrição, Centro, Emissão, Vencimento, " & _
                     "Pagamento, Liberação, [Valor Original], Acréscimo, Abatimento, Banco, Conta, Moeda, Marcação, PagRec, Situação, VlrMul, VlrMrd, PerMul, VlrDsP , Controle " & _
                     "FROM Duplicatas WHERE "
    End If
    'Verifica se o usuário escolheu um número de nota ou lançamento específico.
    If CDblDef(etxNumero.valorTexto, 0) > 0 Then
        Concat strRetorno, "Nota = ", etxNumero.valorTexto
    Else
        AppendStr strRetorno, "Nota > 0"
    End If
    Concat strRetorno, " AND Situação <> 'Cancelada'"
    If etxParcela.valorInteiro > 0 Then   'Se o usuário escolheu uma parcela
        Concat strRetorno, " AND Parcela = ", etxParcela.valorInteiro
    End If
End Sub

'FUNCTION..: ResolveExpLancto
'Objetivo..: Resove a expressão de consulta quando o usuário deseja ver os lançamentos
'Argumento.: [strResult]: Variável string que será retornada.
Private Sub ResolveExpLancto(strResult As String, Optional blnEditar As Boolean = False)
    
    'pt. 87216 - Ivo Sousa(03/06/2008)
    If blnEditar Then
        strResult = "SELECT * FROM Lançamentos WHERE "
    Else
        strResult = "SELECT 'Lanc' AS Origem, Código as cod_id, Parcela, Tipo, Empresa, Descrição, Centro, Emissão, Vencimento, " & _
                    "Pagamento, Liberação, [Valor Original], Acréscimo, Abatimento, Banco, Conta, Moeda, Marcação, PagRec, Situação, VlrMul, VlrMrd, PerMul, VlrDsP , Controle " & _
                    "FROM Lançamentos WHERE "
    End If
    'Verifica se o usuário escolheu um número de nota ou lançamento específico.
    If CDblDef(etxNumero.valorTexto, 0) > 0 Then
        Concat strResult, "Código = ", etxNumero.valorTexto
    Else
        AppendStr strResult, "Código > 0"
    End If
    Concat strResult, " AND Situação <> 'Cancelada'"
    
    'Se o usuário escolheu uma parcela
    If etxParcela.valorInteiro Then
        Concat strResult, " AND Parcela = ", etxParcela.valorInteiro
    End If
End Sub

'Date.......: 07/05/2008
'Author.....: Ivo Sousa
'Descrição..: Resove a expressão de consulta quando o usuário deseja ver os lançamentos e as duplicatas
'Parametros.: [String] Retorno da expressão
Private Sub ResolveExpDuplLanc(strResultLanc As String, strResultDupl As String, Optional blnConsulta As Boolean)
    
    strResultLanc = "SELECT 'Lanc' AS Origem, Código as cod_id, Parcela, Lançamentos.Tipo, Empresa, Descrição, Centro, Emissão, Vencimento, " & _
             "Pagamento, Liberação, [Valor Original], Acréscimo, Abatimento, Banco, Conta, Moeda, Marcação, PagRec, Situação, VlrMul, VlrMrd, PerMul, VlrDsP , Controle " & _
             "FROM Lançamentos WHERE "
    strResultDupl = "SELECT 'Dupl' AS Origem, Nota as cod_id, Parcela, Duplicatas.Tipo, Empresa, Descrição, Centro, Emissão, Vencimento, " & _
             "Pagamento, Liberação, [Valor Original], Acréscimo, Abatimento, Banco, Conta, Moeda, Marcação, PagRec, Situação, VlrMul, VlrMrd, PerMul, VlrDsP , Controle " & _
             "FROM Duplicatas WHERE "

    'Verifica se o usuário escolheu um número de nota ou lançamento específico.
    If CDblDef(etxNumero.valorTexto, 0) > 0 And Not blnConsulta Then
        Concat strResultLanc, "Lançamentos.Código = ", etxNumero.valorTexto
        Concat strResultDupl, "Duplicatas.Nota = ", etxNumero.valorTexto
    Else
        AppendStr strResultLanc, "Lançamentos.Código > 0"
        AppendStr strResultDupl, "Duplicatas.Nota > 0"
    End If
    Concat strResultLanc, " AND Situação <> 'Cancelada'"
    Concat strResultDupl, " AND Situação <> 'Cancelada'"
    
    'Se o usuário escolheu uma parcela
    If etxParcela.valorInteiro And Not blnConsulta Then
        Concat strResultLanc, " AND Parcela = ", etxParcela.valorInteiro
        Concat strResultDupl, " AND Parcela = ", etxParcela.valorInteiro
    End If
End Sub
'SUB.......: EditaLancto
'Objetivo..: Abre a janela de Duplicata ou Lançamento para que o usuário possa
'            alterar os dados atuais.
Private Sub EditaBaixa()
    Dim strTabela      As String
    Dim strDupls       As String  'Instrução de abertura para a tabela
    Dim fDupl          As Form
    Dim lWnd           As Long    'Manipulador da janela de Duplicatas
    Dim strParcela     As String
    Dim strTipo        As String
    Dim intSetRegistro As Integer
    Dim strOrigem      As String
    Dim blnEscreve     As Boolean
    'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
    Dim dblCodigo      As Double
    Dim lngParcela     As Long
    Dim strEmpresa     As String
    Dim enumPagRec     As enuPagRec
    Dim enumLancDup    As enuLancDup
    
    
    
'Verifica se há algum lançamento para ser editado
If (lvwBaixas.ListItems.Count) Then
    'Carrega a janela de duplicatas e configura sua abertura
    If RegistrosSelecionados(intSetRegistro) = 1 Then
        strOrigem = lvwBaixas.ListItems(intSetRegistro).SubItems(1)
        strTabela = TabelaRegistro(intSetRegistro)
    End If
    
    Me.Hide     'Oculto esta janela enquanto o usuário estiver editando a baixa

'A consulta, aqui em Baixas, é ordenada por Nota ou Código, conforme a opção do
'usuário. Para abrir a tabela eu preciso retirar esta instrução de ordem pois
'a instrução Select deve ser livremente alterada na janela do cadastro.

    If strOrigem = "Dupl" Then
        Call ResolveExpDupl(strDupls, True)
    Else
        Call ResolveExpLancto(strDupls, True)
        'strDupls = RTrim$(Left$(mstrDados, (InStr(1, mstrDados, "ORDER BY", vbTextCompare) - 1)))
    End If

    'Se o usuário tem um registro específica selecionado, tenho que trazê-lo
    If IsValid(lvwBaixas.SelectedItem.Text) Then
        'Projeto: #1203 - História: #10582 - Desenvolvimento#12134 - João Henrique(18/04/2012)
        Concat strDupls, " AND ", IIf(strTabela = "Duplicatas", "Nota", "Código"), " = ", lvwBaixas.SelectedItem.Text
        If strParcela = Left(lvwBaixas.ListItems(lvwBaixas.SelectedItem.Index).SubItems(2), 3) Then
            If Right(strParcela, 1) = "-" Then
                strParcela = Left(strParcela, 2)
            End If
        Else
            strParcela = CLngDef(Left(lvwBaixas.ListItems(lvwBaixas.SelectedItem.Index).SubItems(2), 3))
            If Right(strParcela, 1) = "-" Then
                strParcela = Left(strParcela, 2)
            End If
            'pt. 86401 - Dulcino Júnior (07/04/2008)
            If Left(strParcela, 1) = "-" Then
                strParcela = Right(strParcela, Len(strParcela) - 1)
            End If
        End If
        'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
        dblCodigo = lvwBaixas.SelectedItem.Text
        strTipo = Mid(lvwBaixas.ListItems(lvwBaixas.SelectedItem.Index).SubItems(2), 4, 20)
        lngParcela = strParcela
        strEmpresa = lvwBaixas.ListItems(lvwBaixas.SelectedItem.Index).SubItems(4)
        'Projeto: #1203 - História: #10582 - Desenvolvimento#12134 - João Henrique(18/04/2012)
        If strOrigem = "Dupl" Then
            enumLancDup = Duplicata
        Else
            enumLancDup = Lancamento
        End If
        If mstrPagRec = "R" Then
            enumPagRec = Recebimento
        Else
            enumPagRec = Pagamento
        End If
        frmLancamentoDuplicata.LancDup = enumLancDup
        frmLancamentoDuplicata.PagRec = enumPagRec
        blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, 2061, frmLancamentoDuplicata.name, "Lançamentos a Pagar ou Pagos")
        Call mostrarForm(frmLancamentoDuplicata, 2061)
        'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
        Call frmLancamentoDuplicata.CarregarLancamentoDuplicataOutrasRotinas(dblCodigo, strTipo, lngParcela, strEmpresa, enumPagRec, enumLancDup)
    End If
End If

'Aguarda até que o cadastro seja fechado
lWnd = frmLancamentoDuplicata.hWnd

WaitWindowClose lWnd
'Problema: As vêzes a janela é fechada corretamente, as vêzes não. O que acontece?
'Talvez a referência a janela na veriável fDupl não seja completamente terminada
'na instrução abaixo. Isto torna a janela ainda carregada no sistema. Impedindo que
'o programa termine corretamente. Para evitar este problema, forço o fechamento da
'janela aqui.
    On Error Resume Next
    Unload fDupl
    If err.Number <> 0 Then
        err.Clear
    End If
    Set fDupl = Nothing
    'Recarrega todo o recordset para atualizar as alterações
    SeleDocumentos
    Me.Show
End Sub

'SUB.......: EditaBaixaParcial
'Objetivo..: Configura a janela de Duplicatas/Lançamentos para baixas parcial
Private Sub EditaBaixaParcial()
    Dim lngCodBaixa     As Double 'Código do lançamento/duplicata gerado(a)
    Dim Parcela         As Byte
    Dim strTabela       As String 'Tabela atual
    Dim fLancto         As frmLancamentoDuplicata
    Dim sLancto         As String 'Instrução de seleção
    Dim lWnd            As Long   'Manipulador da janela de Duplicatas
    Dim intSetRegistro  As Integer
    Dim dblAbatimento   As Double
    Dim rs              As New ADODB.Recordset
    'Projeto: 100340 - Problema.: 146186 - Ueder Budni (14/10/2016)
    Dim objLogLancDup       As New clsLogLancamentosDuplicatas
    Dim strValorInformado   As String
    Dim intParcelaSel   As Integer
    Dim strEmpresa      As String
    Dim strTipo         As String
    
    'Seleciona a tabela
    Call RegistrosSelecionados(intSetRegistro)
    strTabela = TabelaRegistro(intSetRegistro)
    'Verifica se há registros a alterar
    If (lvwBaixas.ListItems.Count) Then
        If (Not MovConferido(Format$(Date, FDATA), "KIF")) Then 'Apenas se o movimento do mês ainda não estiver conferido
            lngCodBaixa = CriaBaixaParcial(strValorInformado)
            'Se foi possível cria o Lançamento
            If (lngCodBaixa > 0) Then
                sLancto = "SELECT * FROM " & strTabela & " WHERE PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND "
                If strTabela = "Lançamentos" Then
                    Concat sLancto, "Código = ", CStr(lngCodBaixa), " AND Parcela = (SELECT MAX(Parcela) FROM Lançamentos WHERE PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND Código = ", CStr(lngCodBaixa), ");"
                Else
                    Concat sLancto, "Nota = ", CStr(lngCodBaixa), " AND Parcela = (SELECT MAX(Parcela) FROM Duplicatas WHERE PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND Nota = ", CStr(lngCodBaixa), ");"
                End If

                Call AbreRecordset(rs, sLancto)
                Set fLancto = New frmLancamentoDuplicata
                Call fLancto.CarregarLancamentoDuplicataOutrasRotinas(lngCodBaixa, rs![Tipo], rs![Parcela], rs![Empresa], IIf(optBaixas(0).value = True, enuPagRec.Pagamento, enuPagRec.Recebimento), IIf(strTabela = "Duplicatas", Duplicata, Lancamento))
                Load fLancto
                
                'Aguarda até que o usuário feche a janela
                lWnd = fLancto.hWnd
                WaitWindowClose lWnd
                'Força o fechamento da janela
                On Error Resume Next
                Unload fLancto
                If err.Number <> 0 Then
                    err.Clear
                End If
                Set fLancto = Nothing
                mrstDados("Situação").value = GetResOptions(1000, 4)        '// Parcial
                
                'Projeto: 100340 - Desenv.: 143991 - Ueder Budni (28/09/2016)
                mrstDados.Requery
                intParcelaSel = Left(lvwBaixas.ListItems(intSetRegistro).SubItems(2), 2)
                strTipo = Right(lvwBaixas.ListItems(intSetRegistro).SubItems(2), Len(lvwBaixas.ListItems(intSetRegistro).SubItems(2)) - 3)
                strEmpresa = lvwBaixas.ListItems(intSetRegistro).SubItems(4)
                
                If strTabela = "Lançamentos" Then
                    'pt. 87216 - Ivo Sousa (03/06/2008)
                    'O insert tem que ser feito direto no banco e não atravéz de update na
                    'recordset em função do UNION que foi feito na tabela.
                    'dblAbatimento = GetValue(mrstDados, "Abatimento") +
                    dblAbatimento = GetFieldValue("Abatimento", strTabela, "PagRec = '" & mrstDados.Fields("PagRec").value & "' AND Código = " & CStr(lngCodBaixa) & " AND Parcela = " & intParcelaSel & " AND Empresa = '" & strEmpresa & "'" & " AND Tipo = '" & strTipo & "'") + _
                                            Soma("([Valor Original] + Acréscimo - Abatimento)", _
                                            "Lançamentos", "PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND Código = " & _
                                            CStr(lngCodBaixa) & " AND Parcela = (SELECT MAX(Parcela) FROM Lançamentos WHERE PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND Código = " & CStr(lngCodBaixa) & ")")
                    If dblAbatimento > 0 Then
                        ExecuteSQL ("UPDATE " & strTabela & " SET Abatimento = " & Replace(dblAbatimento, ",", ".") & " WHERE PagRec = '" & mrstDados.Fields("PagRec").value & "' AND Código = " & CStr(lngCodBaixa) & " AND Parcela = " & intParcelaSel)
                        'Projeto: 100340 - Problema.: 146186 - Ueder Budni (14/10/2016)
                        With objLogLancDup
                            Call .SetKey(mrstDados.Fields("PagRec").value, CDbl(lngCodBaixa), strEmpresa, strTipo, CLng(intParcelaSel), Lancamento)
                            Call .InsertMsg("Realiza baixa parcial do título no valor de R$" & Format(CDbl(strValorInformado), "##,##0.00") & ".")
                        End With
                    End If
                Else
                    'pt. 87216 - Ivo Sousa (03/06/2008)
                    'dblAbatimento = GetValue(mrstDados, "Abatimento") +
                    dblAbatimento = GetFieldValue("Abatimento", strTabela, "PagRec = '" & mrstDados.Fields("PagRec").value & "' AND Nota = " & CStr(lngCodBaixa) & " AND Parcela = " & intParcelaSel & " AND Empresa = '" & strEmpresa & "'" & " AND Tipo = '" & strTipo & "'") + _
                                            Soma("([Valor Original] + Acréscimo - Abatimento)", _
                                            "Duplicatas", "PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND Nota = " & _
                                            CStr(lngCodBaixa) & " AND Parcela = (SELECT MAX(Parcela) FROM Duplicatas WHERE PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND Nota = " & CStr(lngCodBaixa) & ")")
                    If dblAbatimento > 0 Then
                        ExecuteSQL ("UPDATE " & strTabela & " SET Abatimento = " & Replace(dblAbatimento, ",", ".") & " WHERE PagRec = '" & mrstDados.Fields("PagRec").value & "' AND Nota = " & CStr(lngCodBaixa) & " AND Parcela = " & intParcelaSel & " AND Empresa = '" & strEmpresa & "' AND Tipo = '" & strTipo & "'")
                        'Projeto: 100340 - Problema.: 146186 - Ueder Budni (14/10/2016)
                        With objLogLancDup
                            Call .SetKey(mrstDados.Fields("PagRec").value, CStr(lngCodBaixa), strEmpresa, strTipo, CLng(intParcelaSel), Duplicata)
                            Call .InsertMsg("Realiza baixa parcial do título no valor de R$" & Format(CDbl(strValorInformado), "##,##0.00") & ".")
                        End With
                    End If
                End If
            End If
            'Recarrega todos os registros do Recordset
            SeleDocumentos
            Me.Show
        End If
    End If
End Sub

'SUB.......: XMark
'Objetivo..: Marca com um X o ítem selecionado pelo usuário quando esta não está
'            marcado, ou desmarca quando este estiver marcado.
'Argumento.: [lngIndice]: Índico do ítem que deve ser marcado ou desmarcado.
Private Sub XMark(lngIndice As Long)
    If lngIndice > 0 Then
        If (lvwBaixas.ListItems(lngIndice).SmallIcon = DL_MARCADO) Then
            lvwBaixas.ListItems(lngIndice).SmallIcon = DL_DESMARCADO
        Else
            lvwBaixas.ListItems(lngIndice).SmallIcon = DL_MARCADO
        End If
        'Alterando o registro no Recordset
        On Error Resume Next
        'O SQL server não possui índice de registros começando do zero e sim do 1
        'Pt. 95368 - Moacir Pfau(19/10/2009)
        'mrstDados.AbsolutePosition = IIf((lngIndice - 1) = 0, 1, (lngIndice - 1))
        Else
            If mrstDados.Supports(adBookmark) Then
                mrstDados.AbsolutePosition = (lngIndice)
            End If
        End If
        If err.Number > 0 Then
            DAOErros LoadResString(17)
            Exit Sub
            Resume
        Else
            
            If gTipoDB = Access Then
                mrstDados("Marcação").value = (lvwBaixas.ListItems(lngIndice).SmallIcon = DL_MARCADO)
                mrstDados.update
            End If
        End If
End Sub

Private Sub XMarkRules(lngIndice As Long, DL As Integer)
    If lngIndice > 0 Then
        lvwBaixas.ListItems(lngIndice).SmallIcon = DL
        'Alterando o registro no Recordset
        On Error Resume Next
        'O SQL server não possui índice de registros começando do zero e sim do 1
        'If gTipoDB = Access Then
            
            'mrstDados.AbsolutePosition = (lngIndice - 1)
        'Else
            'Pt. 95368 - Moacir Pfau(16/11/2009)
            If mrstDados.Supports(adBookmark) Then
                mrstDados.AbsolutePosition = (lngIndice)
            End If
        'End If
        If err.Number > 0 Then
            DAOErros LoadResString(17)
            Exit Sub
            Resume
        Else
            If TypeOf mrstDados Is dao.Recordset Then mrstDados.Edit
            mrstDados("Marcação").value = (lvwBaixas.ListItems(lngIndice).SmallIcon = DL_MARCADO)
            mrstDados.update
        End If
    End If
End Sub

'FUNCTION..: CriaBaixaParcial
'Objetivo..: Preenche o cadastro de Lançamentos com os dados da Baixa Parcial.
'Retorna...: O código do lançamento gerado.
Private Function CriaBaixaParcial(Optional ByRef strValorInformado As String) As Double
    Dim strValor       As String
    Dim strLancDupli   As String       'Gera o Lançamento
    Dim strTabela      As String       'Tabela atual
    'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
    Dim dblCodigo      As Double
    Dim strParcela     As String
    Dim intSetRegistro As Integer
    Dim strOrigem      As String
    Dim intNrDup       As Integer
    Dim intParcOrigem  As Integer
    'Projeto: 100340 - Desenv.: 146186 - Ueder Budni (14/10/2016)
    Dim objLogLancDup   As New clsLogLancamentosDuplicatas
    
    intNrDup = 0
    
    'Seleciona a tabela
    intNrDup = RegistrosSelecionados(intSetRegistro)
    'Moacir Pfau(08/01/2009)
    If intNrDup > 1 Then
        MsgBox "Para realizar a baixa parcial só poderá ser selecionado um único título.", vbInformation, NomeModulo
        Exit Function
    End If
    
    strOrigem = lvwBaixas.ListItems(intSetRegistro).SubItems(1)
    If optTodos.value Then
        If strOrigem = "Dupl" Then
            strTabela = "Duplicatas"
            cboBaixas(1).List(0) = "Código/Nota"
            optBaixas_GotFocus (0)
        Else
            strTabela = "Lançamentos"
            cboBaixas(1).List(0) = "Código/Nota"
            optBaixas_GotFocus (0)
        End If
    Else
        If optDup.value Then
            strTabela = "Duplicatas"
            cboBaixas(1).List(0) = "Nota"
            optBaixas_GotFocus (1)
            strOrigem = "Dupl"
        Else
            strTabela = "Lançamentos"
            cboBaixas(1).List(0) = "Código"
            optBaixas_GotFocus (0)
            strOrigem = "Lanc"
        End If
    End If
    
    'Uma InputBox simples para obter o valor desejado para a baixa
    strValor = InputBox(LoadResString(148), MsgBoxCaption, Format$(ZERO, FMOEDA))
    strValorInformado = strValor
    'Se o valor retornado for maior que zero e se for menor que o valor
    'atual do registro
    If (CMoeda(strValor) > 0) Then
        mrstDados.MoveFirst
        mrstDados.Move (intSetRegistro - 1)
        If (CMoeda(strValor) < Kif_Valor(mrstDados)) Then
            'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
            dblCodigo = GetValue(mrstDados, "cod_id", ZERO)
            intParcOrigem = GetValue(mrstDados, "Parcela", ZERO)
            
            'Criando um Lançamento com o valor da baixa da nota
            strLancDupli = "INSERT INTO " & strTabela & "(PagRec, " & _
                    IIf(strTabela = "Lançamentos", "Código, Parcela, ", "Nota, Parcela, ") & _
                    "Empresa, Tipo, Descrição, " & _
                    "Emissão, Vencimento, Pagamento, Liberação, [Valor Original], " & _
                    "Acréscimo, Abatimento, Banco, Conta, Centro, Cheque, Moeda, " & _
                    "[Valor da Moeda], Controle, Marcação, Obs, Borderô, parc_origem_baixa) VALUES ( " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " , "
            'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
            AppendStr strLancDupli, CStr(dblCodigo) 'Código/Nota
            If strTabela = "Duplicatas" Then
                strParcela = ProximoNumero("Parcela", "Duplicatas", _
                                        "PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND Nota = " & GetValue(mrstDados, "cod_id", ZERO))  'Parcela
            Else
                strParcela = ProximoNumero("Parcela", "Lançamentos", _
                                        "PagRec = " & Quote(IIf(optBaixas(0).value = True, "P", "R"), "''") & " AND Código = " & GetValue(mrstDados, "cod_id", ZERO))  'Parcela
            End If
            AppendStr strLancDupli, ", " & strParcela
            AppendStr strLancDupli, ", '" & GetValue(mrstDados, "Empresa") & "'"
            AppendStr strLancDupli, ", '" & GetValue(mrstDados, "Tipo") & "'"
            AppendStr strLancDupli, ", '" & GetValue(mrstDados, "Descrição") & "'"
            'pt. 86113 - Dulcino Júnior (07/04/2008)
            'Conforme conversa com a consultoria (Carlos Dias - 07/04/2008) as data devem ser sempre
            'as datas atuais, para a baixa parcial gerada.
            AppendStr strLancDupli, ", " & InverteData(Date, True)
            AppendStr strLancDupli, ", " & InverteData(Date, True)
            AppendStr strLancDupli, ", " & InverteData(Date, True) 'Pagamento
            AppendStr strLancDupli, ", " & InverteData(Date, True) 'Liberação
            AppendStr strLancDupli, ", " & ValStr(CMoeda(strValor)) 'Valor original
            AppendStr strLancDupli, ", 0" 'Acréscimo
            AppendStr strLancDupli, ", 0" 'Abatimento
            AppendStr strLancDupli, ", " & GetValue(mrstDados, "Banco")
            AppendStr strLancDupli, ", " & GetValue(mrstDados, "Conta")
            AppendStr strLancDupli, ", " & GetValue(mrstDados, "Centro")
            AppendStr strLancDupli, ", 0" 'Cheque
            AppendStr strLancDupli, ", '" & GetValue(mrstDados, "Moeda") & "'"
            AppendStr strLancDupli, ", 0" 'Valor da Moeda
            AppendStr strLancDupli, ", '" & GetValue(mrstDados, "Controle") & "'"
            AppendStr strLancDupli, ", 0" 'Marcação
            
            'pt. 79903 - Ivo Sousa(07/05/2008)
            'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
            AppendStr strLancDupli, ", '" & GetObeservacao(dblCodigo, CInt(strParcela), GetValue(mrstDados, "Tipo"), strOrigem, GetValue(mrstDados, "PagRec")) & "'"  'Observação
            AppendStr strLancDupli, ", 0" 'Borderô
            AppendStr strLancDupli, ", " & intParcOrigem & ");" 'Parcela de Origem
            'Cria o Lançamento
            If (ExecuteSQL(strLancDupli)) Then
                'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                CriaBaixaParcial = dblCodigo     'Retorna o código criado
                
                'Projeto: 100340 - Desenv.: 146186 - Ueder Budni (14/10/2016)
                With objLogLancDup
                    Call .SetKey(IIf(optBaixas(0).value = True, "P", "R"), dblCodigo, GetValue(mrstDados, "Empresa"), GetValue(mrstDados, "Tipo"), CLng(strParcela), IIf(strTabela = "Lançamentos", Lancamento, Duplicata))
                    Call .InsertMsg("Título criado como uma baixa parcial proveniente da parcela " & intParcOrigem & ".")
                End With
            End If
        Else
            MsgBox LoadResString(151), vbInformation, NomeModulo
        End If
    End If
    'Projeto: 100340 - Desenv.: 146186 - Ueder Budni (14/10/2016)
    Set objLogLancDup = Nothing
End Function

'FUNCTION..: LibProc
'Objetivo..: Nenhum. Esta função só existe para não gerar erro se o usuário
'            clicar em algum dos botões da barra de ferramentas.
'Argumento.: [strButton]: Propriedade Key do botão clicado.
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
    Select Case (sFuncao)
        Case WL_MENUCLICK
            Select Case (lFuncao)
                Case IDM_BX_NOVO
                    Call LimpaCampos
                Case IDM_BX_VIEW
                    SeleDocumentos
                Case IDM_BX_EDITAR
                    Call EditaBaixa
                Case IDM_BX_PARCIAL
                    EditaBaixaParcial
                Case IDM_BX_FECHAR
                    Unload Me
                    LibProc = True
                    Exit Function
                Case IDM_BX_EMPRESAS
                    ConsultaEmpresas
                Case IDM_BX_NOTAS
                    ConsultaNotas
                Case Else
                    Exit Function
            End Select
            LibProc = True
        Case WL_NOVO
            Call LimpaCampos
        Case WL_SAIR
            Unload Me
            Exit Function
    End Select
End Function

Public Sub Baixar()
    Dim Index           As Long
    Dim lngItenMarcado  As Long
    Dim strCampos       As String
    Dim SQL             As String
    Dim SqlCheque       As String
    Dim bEDuplicata As Boolean
    'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
    Dim dblCodigo   As Double
    Dim strEmpresa  As String
    Dim bytParcela  As Byte
    Dim strTipo     As String
    Dim strPagRec   As String
    Dim bBaixou     As Boolean
    
    SQL = "Select Pagamento,Banco,Controle From Duplicatas "
    If Not edtDataPagamento.IsValidDate Then
        MsgFunc "É necessário preencher a data de pagamento."
        edtDataPagamento.SetFocus
        Exit Sub
    End If
    
    strCampos = "Pagamento= " & InverteData(edtDataPagamento.Data, True)
    strCampos = strCampos & ", Liberação = " & InverteData(edtDataLiberacao.Data, True)
    
    If etxBancoBaixa.valorInteiro > 0 Then
        strCampos = strCampos + ", Banco= " & etxBancoBaixa.valorInteiro
    End If
    If etxControleBaixa.valorInteiro > 0 Then
        strCampos = strCampos + ", Controle =" & Quote(etxControleBaixa.valorInteiro, "''")
    End If
    If etxChequeBaixa.valorInteiro > 0 Then
        strCampos = strCampos + ", Cheque =" & etxChequeBaixa.valorInteiro
    End If
    
    If etxBancoBaixa.valorInteiro > 0 And etxChequeBaixa.valorInteiro > 0 Then
      SqlCheque = "INSERT INTO Cheque (Banco, Cheque) VALUES (" & etxBancoBaixa.valorInteiro & "," & etxChequeBaixa.valorInteiro & ")"
    End If
    
    strPagRec = IIf(optBaixas(1).value, "R", "P")
    
    For lngItenMarcado = lvwBaixas.ListItems.Count To 1 Step -1
        If lvwBaixas.ListItems(lngItenMarcado).SmallIcon = DL_MARCADO Then
            'pt. 79903 - Ivo Sousa(07/05/2008)
            If lvwBaixas.ListItems(lngItenMarcado).SubItems(1) = "Dupl" Then
                bEDuplicata = True
            Else
                bEDuplicata = False
            End If

            'Buscando a chave do registro para baixa-lo
            'Quando for duplicata
            'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
            dblCodigo = CLngDef(lvwBaixas.ListItems(lngItenMarcado).Text)
            
            If bEDuplicata Then
                bytParcela = CByteDef(Left(lvwBaixas.ListItems(lngItenMarcado).SubItems(2), 2))
                strTipo = Mid$(lvwBaixas.ListItems(lngItenMarcado).SubItems(2), 4)
                strEmpresa = lvwBaixas.ListItems(lngItenMarcado).SubItems(4)
                'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                bBaixou = (ExecuteSQL("UPDATE Duplicatas SET " & strCampos & " WHERE Pagrec =" & Quote(strPagRec, "''") & " AND Nota = " & str(dblCodigo) & " AND Tipo=" & Quote(strTipo, "''") & " AND Empresa =" & Quote(strEmpresa, "''") & " AND Parcela=" & str(bytParcela)))
            Else
                bytParcela = CByteDef(Left(lvwBaixas.ListItems(lngItenMarcado).SubItems(2), 2))
                strTipo = Mid$(lvwBaixas.ListItems(lngItenMarcado).SubItems(2), 4)
                'pt. 84204 - Dulcino Júnior (08/11/2007)
                'A parcela deve fazer parte do filtro para identificação do lançamento.
                'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                bBaixou = (ExecuteSQL("UPDATE Lançamentos SET " & strCampos & " WHERE pagrec=" & Quote(strPagRec, "''") & " AND Código=" & str(dblCodigo) & " AND Parcela=" & str(bytParcela)))
            End If
            If bBaixou Then
              If IsValid(SqlCheque) Then ExecuteSQL SqlCheque
            End If
        End If
    Next
    If (lvwBaixas.ListItems.Count) Then     'Se houver algum registro
      lvwBaixas.SetFocus
      cmdBaixaLote.Enabled = True   'Enquanto não houver registro desabilito o botão.
    Else
      cmdBaixaLote.Enabled = False
    End If
    MsgFunc "Duplicata(s) baixada(s) com sucesso."
End Sub

Private Sub etxBancoFinal_GotFocus()
    DescStatus etxBancoFinal.TabIndex
End Sub

Private Sub etxBancoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxBancoFinal.ValorDescricao = "" Then
            etxBancoFinal.valorInteiro = 0
        End If
        Call PCampo("Consulta de Bancos", "SELECT Banco, Nome, Agência, Conta, [Nome Conta] FROM Bancos", pbCampo, etxBancoFinal, "Banco")
    End If
End Sub

Private Sub etxBancoInicial_GotFocus()
    DescStatus etxBancoInicial.TabIndex
End Sub

Private Sub etxBancoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxBancoInicial.ValorDescricao = "" Then
            etxBancoInicial.valorInteiro = 0
        End If
        Call PCampo("Consulta de Bancos", "SELECT Banco, Nome, Agência, Conta, [Nome Conta] FROM Bancos", pbCampo, etxBancoInicial, "Banco")
    End If
End Sub

Private Sub etxCentroCustoFinal_GotFocus()
    DescStatus etxCentroCustoFinal.TabIndex
End Sub

Private Sub etxCentroCustoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxCentroCustoFinal.ValorDescricao = "" Then
            etxCentroCustoFinal.valorInteiro = 0
        End If
        Call PCampo("Consulta de Centro de custos", "SELECT Código, Descrição FROM Centros", pbCampo, etxCentroCustoFinal, "Código")
    End If
End Sub

Private Sub etxCentroCustoInicial_GotFocus()
    DescStatus etxCentroCustoInicial.TabIndex
End Sub

Private Sub etxCentroCustoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxCentroCustoInicial.ValorDescricao = "" Then
            etxCentroCustoInicial.valorInteiro = 0
        End If
        Call PCampo("Consulta de Centro de custos", "SELECT Código, Descrição FROM Centros", pbCampo, etxCentroCustoInicial, "Código")
    End If
End Sub

Private Sub etxContaFinal_GotFocus()
    DescStatus etxContaFinal.TabIndex
End Sub

Private Sub etxContaFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxContaFinal.ValorDescricao = "" Then
            etxContaFinal.valorInteiro = 0
        End If
        Call PCampo("Consulta de contas", "SELECT Código, Grupo, Descrição FROM Contas WHERE Ctaati='S'", pbCampo, etxContaFinal, "Código")
    End If
End Sub

Private Sub etxContaInicial_GotFocus()
    DescStatus etxContaInicial.TabIndex
End Sub

Private Sub etxContaInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If etxContaInicial.ValorDescricao = "" Then
            etxContaInicial.valorInteiro = 0
        End If
        Call PCampo("Consulta de contas", "SELECT Código, Grupo, Descrição FROM Contas WHERE Ctaati='S'", pbCampo, etxContaInicial, "Código")
    End If
End Sub

Private Sub edtDataEmissaoFinal_GotFocus()
    DescStatus edtDataEmissaoFinal.TabIndex
End Sub

Private Sub edtDataEmissaoInicial_GotFocus()
    DescStatus edtDataEmissaoInicial.TabIndex
End Sub

Private Sub edtDataLiberacaoFinal_GotFocus()
    DescStatus edtDataLiberacaoFinal.TabIndex
End Sub

Private Sub edtDataLiberacaoInicial_GotFocus()
    DescStatus edtDataLiberacaoInicial.TabIndex
End Sub

Private Sub edtDataVencimentoFinal_GotFocus()
    DescStatus edtDataVencimentoFinal.TabIndex
End Sub

Private Sub edtDataVencimentoInicial_GotFocus()
    DescStatus edtDataVencimentoInicial.TabIndex
End Sub

Private Sub etxOpContabilDupl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Operações Contábeis", "OperacaoContabil", pbCampo, etxOpContabilDupl, "cd_operacao")
    End If
End Sub

Private Sub etxOpContabilLanc_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call PCampo("Operações Contábeis", "OperacaoContabil", pbCampo, etxOpContabilLanc, "cd_operacao")
    End If
End Sub

Private Sub txtQtlSelecionados_Change()
On Error GoTo err_change
    If CInt(txtQtlSelecionados.Caption) > 0 Then
        cmdBaixaLote.Enabled = True
    Else
        cmdBaixaLote.Enabled = False
    End If
    Exit Sub
err_change:
    err.Clear
    Resume Next
End Sub

Private Function getOrderBy() As String
    If optNotaCod.value Then
        If optDup.value Then
            cboBaixas(1).Text = "Nota"
            getOrderBy = "Nota"
        ElseIf optLanc.value Then
            cboBaixas(1).Text = "Código"
            getOrderBy = "Código"
        Else
            cboBaixas(1).Text = "Código/Nota"
            getOrderBy = "cod_id"
        End If
    End If
    If optEmpresa.value Then
        cboBaixas(1).Text = "Empresa"
        getOrderBy = "Empresa"
    End If
    If optControle.value Then
        cboBaixas(1).Text = "Controle"
        getOrderBy = "Controle"
    End If
    If optEmissao.value Then
        cboBaixas(1).Text = "Emissão"
        getOrderBy = "Emissão"
    End If
    If optVenc.value Then
        cboBaixas(1).Text = "Vencimento"
        getOrderBy = "Vencimento"
    End If
End Function

'Data.......: 30/05/2007
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para sugerir a operação contábil
'               de baixa de acordo com os critérios solicitados no pt
'               82037, conforme combinado com a consultoria.
Private Sub SugereOperacaoContabil()
    Dim objDAO          As cMatrizContabilizacaoDAO
    Dim objMatriz       As cMatrizContabilizacao
    Dim lngOperacaoDupl As Long
    Dim lngOperacaoLanc As Long

    Set objDAO = New cMatrizContabilizacaoDAO
    'Se existir algum tipo gloabl selecionado
    If cboBaixas(0).Text <> "" Then
        'Caso exista algum tipo global selecionado, deve ser sugerida a operação por ele.
        If cboBaixas(0).Text <> "Todos" Then
            Set objMatriz = objDAO.Carregar(cboBaixas(0).Text)
        Else
            Set objMatriz = objDAO.Carregar("Fatura")
        End If
        'Caso exista algum operação configurada na matriz para o tipo global.
        If Not objMatriz Is Nothing Then
            If optDup.value Then
                'Se for duplicata a Pagar
                If optBaixas(0) Then
                    lngOperacaoDupl = objMatriz.BaixaDuplicatasPagar
                Else
                    lngOperacaoDupl = objMatriz.BaixaDuplicatasReceber
                End If
            ElseIf optLanc.value Then
                'Se for lançamento a Pagar
                If optBaixas(0) Then
                    lngOperacaoLanc = objMatriz.BaixaLancamentosPagar
                Else
                    lngOperacaoLanc = objMatriz.baixaLancamentosReceber
                End If
            Else
                'pt. 79903 - Ivo Sousa(08/05/2008)
                If optBaixas(0) Then
                    lngOperacaoDupl = objMatriz.BaixaDuplicatasPagar
                    lngOperacaoLanc = objMatriz.BaixaLancamentosPagar
                Else
                    lngOperacaoDupl = objMatriz.BaixaDuplicatasReceber
                    lngOperacaoLanc = objMatriz.baixaLancamentosReceber
                End If
            End If
        Else
            lngOperacaoDupl = 0
            lngOperacaoLanc = 0
        End If
    End If
    etxOpContabilDupl.valorInteiro = lngOperacaoDupl
    etxOpContabilLanc.valorInteiro = lngOperacaoLanc
    Set objMatriz = Nothing
    Set objDAO = Nothing
End Sub

Private Sub LimpaCampos()
    txtEmpresaUsuaria.Text = DonaSistema
    lblEmpresaUsuaria.Caption = NomeDonaSistema
    edtDataLiberacaoInicial.Clear
    edtDataLiberacaoFinal.Clear
    edtDataVencimentoInicial.Clear
    edtDataVencimentoFinal.Clear
    edtDataEmissaoInicial.Clear
    edtDataEmissaoFinal.Clear
    etxBancoInicial.Clear
    etxBancoFinal.Clear
    etxContaInicial.Clear
    etxContaFinal.Clear
    etxCentroCustoInicial.Clear
    etxCentroCustoFinal.Clear
    etxValorOriginalInicial.Clear
    etxValorOriginalFinal.Clear
    etxEmpresas.Clear
    lvwBaixas.ListItems.Clear
End Sub

'Data.......: 25/03/2008
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para validar os campos informados na tela, não
'               permitindo que seja utilizada informação inválida na consulta.
'Retorno....: [Boolean] Retorna se a consulta pode ser executada ou não.
Private Function ValidaCampos() As Boolean
    ValidaCampos = True
    If Not IsEmptyDate(edtDataLiberacaoInicial.Data) And Not IsEmptyDate(edtDataLiberacaoFinal.Data) Then
        If edtDataLiberacaoInicial.Data > edtDataLiberacaoFinal.Data Then
            MsgBox "O data de liberação inicial deve ser maior do que a data de liberação final.", vbInformation, NomeModulo
            edtDataLiberacaoInicial.SetFocus
            ValidaCampos = False
        End If
    End If
    If ValidaCampos Then
        If Not IsEmptyDate(edtDataVencimentoInicial.Data) And Not IsEmptyDate(edtDataVencimentoFinal.Data) Then
            If edtDataVencimentoInicial.Data > edtDataVencimentoFinal.Data Then
                MsgBox "O data de vencimento inicial deve ser menor do que a data de vencimento final.", vbInformation, NomeModulo
                edtDataVencimentoInicial.SetFocus
                ValidaCampos = False
            End If
        End If
    End If
    If ValidaCampos Then
        If Not IsEmptyDate(edtDataEmissaoInicial.Data) And Not IsEmptyDate(edtDataEmissaoFinal.Data) Then
            If edtDataEmissaoInicial.Data > edtDataEmissaoFinal.Data Then
                MsgBox "O data de emissão inicial deve ser menor do que a data de emissão final.", vbInformation, NomeModulo
                edtDataEmissaoInicial.SetFocus
                ValidaCampos = False
            End If
        End If
    End If
    If ValidaCampos Then
        If etxBancoInicial.valorInteiro > 0 And etxBancoFinal.valorInteiro > 0 Then
            If etxBancoInicial.valorInteiro > etxBancoFinal.valorInteiro Then
                MsgBox "O banco inicial deve ser menor do que o banco final.", vbInformation, NomeModulo
                etxBancoInicial.SetFocus
                ValidaCampos = False
            End If
        End If
    End If
    If ValidaCampos Then
        If etxContaInicial.valorInteiro > 0 And etxContaFinal.valorInteiro > 0 Then
            If etxContaInicial.valorInteiro > etxContaFinal.valorInteiro Then
                MsgBox "A conta inicial deve ser menor do que a conta final.", vbInformation, NomeModulo
                etxContaInicial.SetFocus
                ValidaCampos = False
            End If
        End If
    End If
    If ValidaCampos Then
        If etxCentroCustoInicial.valorInteiro > 0 And etxCentroCustoFinal.valorInteiro > 0 Then
            If etxCentroCustoInicial.valorInteiro > etxCentroCustoFinal.valorInteiro Then
                MsgBox "O centro de custo inicial deve ser menor do que o centro de custo final.", vbInformation, NomeModulo
                etxCentroCustoInicial.SetFocus
                ValidaCampos = False
            End If
        End If
    End If
    If ValidaCampos Then
        If etxValorOriginalInicial.valorMoeda > 0 And etxValorOriginalFinal.valorMoeda > 0 Then
            If etxValorOriginalInicial.valorMoeda > etxValorOriginalFinal.valorMoeda Then
                MsgBox "O valor original inicial deve ser menor do que o valor original final.", vbInformation, NomeModulo
                etxValorOriginalInicial.SetFocus
                ValidaCampos = False
            End If
        End If
    End If
End Function

'Data.......: 26/03/2008
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para validar os campos de baixa.
'Retorno....: [Boolean] Retorna se os campos obrigatórios da frame de baixa
'               foram preenchidos.
Private Function ValidaBaixaLote() As Boolean
    Dim strRetorno  As String
    Dim strMensagem As String
    
    ValidaBaixaLote = True
    If etxBancoBaixa.valorInteiro = 0 Then
        MsgBox "O banco deve ser preenchido.", vbInformation, NomeModulo
        ValidaBaixaLote = False
        etxBancoBaixa.SetFocus
    End If
    
    If ValidaBaixaLote Then
        If Not edtDataPagamento.IsValidDate Then
            MsgBox "A data de pagamento deve ser preenchida.", vbInformation, NomeModulo
            ValidaBaixaLote = False
            edtDataPagamento.SetFocus
        End If
    End If
    
    If ValidaBaixaLote Then
        strRetorno = calendario.PermiteLancamento(edtDataPagamento.Data)
        Select Case strRetorno
            Case "X"
                strMensagem = "A data de pagamento informada está bloqueada." & vbNewLine & "Informe outra data para realizar o lançamento."
            Case "F"
                strMensagem = "A data de pagamento informada é um 'Feriado'" & vbNewLine & "Confirma o lançamento?"
            Case "S"
                strMensagem = "A data de pagamento informada é um 'Sábado'" & vbNewLine & "Confirma o lançamento?"
            Case "D"
                strMensagem = "A data de pagamento informada é um 'Domingo'" & vbNewLine & "Confirma o lançamento?"
            Case Else
                strMensagem = ""
        End Select
        If strMensagem <> "" Then
            If Right(strMensagem, 1) = "?" Then
                If MsgBox(strMensagem, vbQuestion + vbYesNo, NomeModulo) = vbNo Then
                    ValidaBaixaLote = False
                End If
            Else
                Call MsgBox(strMensagem, vbInformation, NomeModulo)
                ValidaBaixaLote = False
                edtDataPagamento.SetFocus
            End If
        End If
        strMensagem = ""
    End If
    
    If ValidaBaixaLote Then
        If Not edtDataLiberacao.IsValidDate Then
            MsgBox "A data de liberação deve ser preenchida.", vbInformation, NomeModulo
            ValidaBaixaLote = False
            edtDataLiberacao.SetFocus
        End If
    End If
    
    If ValidaBaixaLote Then
        strRetorno = calendario.PermiteLancamento(edtDataLiberacao.Data)
        Select Case strRetorno
            Case "X"
                strMensagem = "A data de liberação informada está bloqueada." & vbNewLine & "Informe outra data para realizar o lançamento."
            Case "F"
                strMensagem = "A data de liberação informada é um 'Feriado'" & vbNewLine & "Confirma o lançamento?"
            Case "S"
                strMensagem = "A data de liberação informada é um 'Sábado'" & vbNewLine & "Confirma o lançamento?"
            Case "D"
                strMensagem = "A data de liberação informada é um 'Domingo'" & vbNewLine & "Confirma o lançamento?"
            Case Else
                strMensagem = ""
        End Select
        If strMensagem <> "" Then
            If Right(strMensagem, 1) = "?" Then
                If MsgBox(strMensagem, vbQuestion + vbYesNo, NomeModulo) = vbNo Then
                    ValidaBaixaLote = False
                End If
            Else
                Call MsgBox(strMensagem, vbInformation, NomeModulo)
                ValidaBaixaLote = False
                edtDataLiberacao.SetFocus
            End If
        End If
    End If
    
    'pt. 93779 - Ivo Sousa (14/07/2009)
    If ValidaBaixaLote Then
        ValidaBaixaLote = ValidaOpContabil
    End If
End Function

'Data.......: 27/03/2008
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para validar a seleção do item, de acordo com
'               as regras de validação de baixa, Centro de custo, Conta e Banco.
'Parametros.: [Long] Número da linha que está sendo verificada.
'             [Boolean] Retorna se deve ser exibida a mensagem de validação.
'Retorno....: [Boolean] Retorna se a linha pode ser selecionada.
Private Function ValidaSelecao(lngItem As Long, blnMsg As Boolean, blnBaixaParcial As Boolean) As Boolean
    ValidaSelecao = True
    'Protocolo Nr 89509 - Carlos Felippe Vernizze - 23/09/2010
    If lngItem = 0 Then
        lngItem = 1
    End If
    If blnBaixaParcial Then
        If strToLng(lvwBaixas.ListItems(lngItem).SubItems(5)) = 0 Then
            If blnMsg Then
                MsgBox "Para Utilizar a baixa parcial é necessário preencher o banco.", vbInformation, NomeModulo
            End If
            ValidaSelecao = False
        End If
    End If
    If ValidaSelecao Then
        If strToLng(lvwBaixas.ListItems(lngItem).SubItems(6)) = 0 Then
            If blnMsg Then
                If MsgBox("Para baixar esse título é necessário que o campo " & Chr(34) & "Conta" & Chr(34) & " esteja preenchido. Deseja editar o título?", vbInformation + vbYesNo, NomeModulo) = vbYes Then
                    lvwBaixas.ListItems(lngItem).SmallIcon = DL_MARCADO
                    Call EditaBaixa
                End If
            End If
            ValidaSelecao = False
        End If
    End If
    If ConfigSys.ControlarCentrodeCusto Then
        If ValidaSelecao Then
            If strToLng(lvwBaixas.ListItems(lngItem).SubItems(7)) = 0 Then
                If blnMsg Then
                    MsgBox "Para selecionar esse título é necessário que o campo centro de custo esteja preenchido.", vbInformation, NomeModulo
                End If
                ValidaSelecao = False
            End If
        End If
    End If
End Function

'Date.......: 05/05/2008
'Author.....: Ivo Sousa
'Descrição..: Função utilizada para verificar quantos regitrsos estão selecionados na Grid.
'Retorno....: [Integer] Quantidade de Registros
Private Function RegistrosSelecionados(intSetRegistro As Integer) As Integer
    Dim i       As Integer
    Dim intCont As Integer
    
    For i = 1 To lvwBaixas.ListItems.Count
        If lvwBaixas.ListItems(i).SmallIcon = DL_MARCADO Then
            intSetRegistro = i
            intCont = intCont + 1
        End If
    Next
    RegistrosSelecionados = intCont
End Function

'pt. 84737 - Ivo Sousa(06/05/2008)
Private Sub txtVlAcrescimo_LostFocus()
    lblVlTotal.Caption = FormatCurrency(StrToCur(lblValorOriginal.Caption) + StrToCur(txtVlAcrescimo.valorMoeda) - StrToCur(txtVlAbatimento.valorMoeda))
End Sub

'pt. 84737 - Ivo Sousa(06/05/2008)
Private Sub txtVlAbatimento_LostFocus()
    lblVlTotal.Caption = FormatCurrency(StrToCur(lblValorOriginal.Caption) + StrToCur(txtVlAcrescimo.valorMoeda) - StrToCur(txtVlAbatimento.valorMoeda))
End Sub

'Date.......: 06/05/2008
'Author.....: Ivo Sousa
'Descrição..: Função utilizada para montar a clausula de consulta nas tabelas.
'             Utilizada no GetFieldValue que estão no Código.
'Parametros.: [Integer] A linha que esta selecionada.
'             [String] Retorno da parcela do lançamento ou da dupliacata.
'Retorno....: A clausula para SQL na tela.
Private Function MontaClausula(intSetRegistro As Integer, Optional strParcela As String) As String
    Dim strTabela   As String
    Dim strTipo     As String
    Dim strPagRec   As String
    Dim strCampoCod As String
    Dim strEmpresa  As String
    
    strParcela = lvwBaixas.ListItems(intSetRegistro).ListSubItems(2)
    strEmpresa = lvwBaixas.ListItems(intSetRegistro).ListSubItems(4)
    
    'pt. 79903 - Ivo Sousa(08/05/2008)
    If lvwBaixas.ListItems(intSetRegistro).ListSubItems(1) = "Dupl" Then
        strCampoCod = "Nota"
    Else
        strCampoCod = "Código"
    End If
    'If Len(strParcela) = (Len(Replace(strParcela, "-", "")) + 1) Then
    If Left(strParcela, 1) = "-" Then
        strTipo = Mid(strParcela, 5, Len(strParcela))
        strParcela = Mid(strParcela, 1, 3)
    Else
        strTipo = Mid(strParcela, 4, Len(strParcela))
        strParcela = Mid(strParcela, 1, 2)
    End If
    MontaClausula = " PagRec = '" & mstrPagRec & "' AND " & strCampoCod & " = " & lvwBaixas.ListItems(intSetRegistro) & _
                    " AND Parcela = " & strParcela & " AND Tipo = '" & strTipo & "' AND Empresa = '" & strEmpresa & "'"
End Function


'Date.......: 08/05/2008
'Author.....: Ivo Sousa
'Descrição..: Busca as observações dos registros consultados
'Parametros.: [Long]Numero do documento
'             [Integer]Parcela do documento
'             [String]Tipo do documento
'             [String]A Origem do Documento(Duplicata ou Lançamento)
'             [String]PagRec do Documento
'Retorno....: [Boolean]
Private Function GetObeservacao(lngCod As Double, intParcela As Integer, strTipo As String, strOrigem As String, strPagRec As String) As String
    Select Case strOrigem
        Case "Dupl"
            GetObeservacao = GetFieldValue("Obs", "Duplicatas", "Nota = " & lngCod & " AND Parcela = " & intParcela & " AND Tipo = '" & strTipo & "' AND PagRec = '" & strPagRec & "'")
        Case "Lanc"
            GetObeservacao = GetFieldValue("Obs", "Lançamentos", "Código = " & lngCod & " AND Parcela = " & intParcela & " AND Tipo = '" & strTipo & "' AND PagRec = '" & strPagRec & "'")
    End Select
End Function

'Date.......: 08/05/2008
'Author.....: Ivo Sousa
'Descrição..: Limpar os controles referentes ao titulo selecionado
Private Sub LimpaAdicionais()
    txtVlAcrescimo.valorMoeda = 0
    txtVlAbatimento.valorMoeda = 0
    lblVlTotal.Caption = FormatCurrency(0)
    lblValorOriginal.Caption = FormatCurrency(0)
    cmdComfirmar.Enabled = False
    cmdBaixas(1).Enabled = False
End Sub

'Date.......: 08/05/2008
'Author.....: Ivo Sousa
'Descrição..: Verifica qual tabela vai ser consultada
'Parametros.: [Integer]Registro Selecionado
'Retorno....: O nome da tabela
Private Function TabelaRegistro(intSetRegistro As Integer) As String
    If lvwBaixas.ListItems(intSetRegistro).SubItems(1) = "Dupl" Then
        TabelaRegistro = "Duplicatas"
    Else
        TabelaRegistro = "Lançamentos"
    End If
End Function

'Date.......: 25/09/2008
'Author.....: Ivo Sousa
'Descrição..: Validar títulos que estão em atraso
'Parametros.: [Integer]Registro Selecionado
'Retorno....: O nome da tabela
Private Sub ValidaTitulosAtraso()
    Dim strTabela      As String
    Dim lngItemMarcado As Long
    Dim intCont        As Integer
    Dim intParcela     As Integer
    'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
    Dim dblCodigo      As Double
    Dim strOrigem      As String
    Dim strTipo        As String
    Dim strPagRec      As String
    Dim curTaxaJuros   As Currency
    
    intCont = 0
    curTaxaJuros = GetFieldValue("Mora", "Bancos", "Banco = " & etxBancoBaixa.valorInteiro, , 0)
    strPagRec = IIf(optBaixas(1).value, "R", "P")
    For lngItemMarcado = 1 To lvwBaixas.ListItems.Count
        If lvwBaixas.ListItems(lngItemMarcado).SmallIcon = DL_MARCADO Then
            'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
            dblCodigo = CDblDef(lvwBaixas.ListItems(lngItemMarcado).Text)
            If Left(lvwBaixas.ListItems(lngItemMarcado).SubItems(2), 1) = "-" Then
                intParcela = CInt(Left(lvwBaixas.ListItems(lngItemMarcado).SubItems(2), 3))
                strTipo = Mid$(lvwBaixas.ListItems(lngItemMarcado).SubItems(2), 5)
            Else
                intParcela = CByteDef(Left(lvwBaixas.ListItems(lngItemMarcado).SubItems(2), 2))
                strTipo = Mid$(lvwBaixas.ListItems(lngItemMarcado).SubItems(2), 4)
            End If
            strOrigem = lvwBaixas.ListItems(lngItemMarcado).SubItems(1)
            Select Case strOrigem
                Case "Dupl"
                    strTabela = "Duplicatas"
                Case Else
                    strTabela = "Lançamentos"
            End Select
            If GetFieldValue("Acréscimo", strTabela, MontaClausula(CInt(lngItemMarcado), CStr(intParcela)), , 0) = 0 Then
                If CDate(lvwBaixas.ListItems(lngItemMarcado).SubItems(8)) < edtDataPagamento.Data Then
                    Load frmJurosTitulo
                    'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
                    frmJurosTitulo.documento(intCont) = dblCodigo
                    frmJurosTitulo.Parcela(intCont) = intParcela
                    frmJurosTitulo.Banco = etxBancoBaixa.valorInteiro
                    frmJurosTitulo.PagRec(intCont) = strPagRec
                    frmJurosTitulo.Pagamento = edtDataPagamento.Data
                    frmJurosTitulo.Origem(intCont) = strOrigem
                    frmJurosTitulo.TaxaJuros(intCont) = curTaxaJuros
                    frmJurosTitulo.ValorTitulo(intCont) = lvwBaixas.ListItems(lngItemMarcado).ListSubItems(9)
                    frmJurosTitulo.Vencimento(intCont) = lvwBaixas.ListItems(lngItemMarcado).SubItems(8)
                    intCont = intCont + 1
                End If
            End If
        End If
    Next
    If intCont > 0 Then
        If MsgBox("Há título(s) em atraso no lote selecionado. Deseja informar os Juros?", vbQuestion + vbYesNo, NomeModulo) = vbYes Then
            Call frmJurosTitulo.carregaRegistro
            Call mostrarForm(frmJurosTitulo, 2848, True)
        Else
            Unload frmJurosTitulo
        End If
    End If
End Sub

'Data.......: 15/10/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para baixar as informações de rateio da tabela relacional
'Parametros.: [Long] Número da linha que está sendo analisada.
Private Sub BaixaRateio(lngItem As Long)
    Dim strSql           As String
    Dim strTipoParcela() As String
    
    With lvwBaixas.ListItems(lngItem)
        strSql = "WHERE pag_rec_destino='" & IIf(optBaixas(1).value, "R", "P") & "' AND "
        If .SubItems(1) = "Dupl" Then
            strSql = strSql & "nr_nota_destino=" & .Text & " AND "
            strSql = strSql & "cd_empresa_destino='" & .ListSubItems(4).Text & "' AND "
            strTipoParcela = Split(.SubItems(2), "-")
            If UBound(strTipoParcela) > 0 Then
                'pt. 00000 - Ivo Sousa (30/03/2010)
                'Alteração para baixar duplicatas de baixas parciais
                If strTipoParcela(0) = "" Then
                    strSql = strSql & "tp_registro_destino='" & strTipoParcela(2) & "' AND "
                    strSql = strSql & "nr_parcela_destino= -" & strTipoParcela(1)
                Else
                    strSql = strSql & "tp_registro_destino='" & strTipoParcela(1) & "' AND "
                    strSql = strSql & "nr_parcela_destino=" & strTipoParcela(0)
                End If
            End If
        Else
            strSql = strSql & "cd_lancamento_destino=" & .Text & " AND "
            strSql = strSql & "nr_parcela_destino=" & Left(.SubItems(2), 2)
        End If
        strSql = "SET dt_pagamento=" & InverteData(edtDataPagamento.Data, True) & " " & strSql
        If .SubItems(1) = "Dupl" Then
            strSql = "UPDATE FFIRateioDuplicata " & strSql
        Else
            strSql = "UPDATE FFIRateioLancamento " & strSql
        End If
        Call ExecuteSQL(strSql)
    End With
End Sub

'Pt.........: 93779
'Data.......: 14/07/2009
'Autor......: Ivo Sousa
'Descrição..: Verifica se as Operações Contabeis estão devidamente preenchidas
Private Function ValidaOpContabil() As Boolean
    Dim blnBaixaLanc As Boolean
    Dim blnBaixaDupl As Boolean
    Dim lngItem      As Long
    
    ValidaOpContabil = False
    For lngItem = 1 To lvwBaixas.ListItems.Count
        If lvwBaixas.ListItems(lngItem).SmallIcon = DL_MARCADO Then
            If Not blnBaixaDupl Then
                blnBaixaDupl = (lvwBaixas.ListItems(lngItem).SubItems(1) = "Dupl")
            End If
            If Not blnBaixaLanc Then
                blnBaixaLanc = (lvwBaixas.ListItems(lngItem).SubItems(1) = "Lanc")
            End If
        End If
        If blnBaixaLanc And blnBaixaDupl Then
            Exit For
        End If
    Next
    If blnBaixaLanc Or blnBaixaDupl Then
        If blnBaixaLanc Then
            If etxOpContabilLanc.Enabled And etxOpContabilLanc.ValorDescricao = "" Then
                MsgBox "É necessário informar a Operação Contábil Lançamento.", vbInformation, NomeModulo
                Exit Function
            End If
        End If
        If blnBaixaDupl Then
            If etxOpContabilDupl.Enabled And etxOpContabilDupl.ValorDescricao = "" Then
                MsgBox "É necessário informar a Operação Contábil Duplicatas.", vbInformation, NomeModulo
                Exit Function
            End If
        End If
    Else
        ValidaOpContabil = False
    End If
    ValidaOpContabil = True
End Function

'Projeto: 100340 - Desenv.: 142890 - Ueder Budni (23/09/2016)
Private Sub RegistraLogLancDupBaixa(dblNumero As Double, strEmpresa As String, strTipo As String, lngParcela As Long, strPagRec As String, enuTabela As enuLancDup, voOldStateObj As VoLancamentoDuplicata)
    Dim objLogLancDup   As New clsLogLancamentosDuplicatas
    Dim strStdMsg       As String

On Error GoTo erro

    If Not voOldStateObj Is Nothing Then
        With voOldStateObj
            strStdMsg = "Alterado campo {0} de '{1}' para '{2}' através da rotina de Baixas."
            
            Call objLogLancDup.SetKey(strPagRec, dblNumero, strEmpresa, strTipo, lngParcela, enuTabela)
                
            If Trim(etxControleBaixa.valorTexto) <> "" And Trim(etxControleBaixa.valorTexto) <> .Controle Then
                Call objLogLancDup.InsertCustomMsg(strStdMsg, "Controle", .Controle, etxControleBaixa.valorTexto)
            End If
            
            If etxChequeBaixa.valorInteiro > 0 And etxChequeBaixa.valorInteiro <> .Cheque Then
                Call objLogLancDup.InsertCustomMsg(strStdMsg, "Cheque", .Cheque, etxChequeBaixa.valorInteiro)
            End If
            
            If etxBancoBaixa.valorInteiro > 0 And etxBancoBaixa.valorInteiro <> .Banco Then
                Call objLogLancDup.InsertCustomMsg(strStdMsg, "Banco", .Banco, etxBancoBaixa.valorInteiro)
            End If
            
            If .Conciliado <> CBool(chkConciliado.value) Then
                Call objLogLancDup.InsertCustomMsg("Campo {0} foi {1} através da rotina de Baixas.", "Conciliado", IIf(chkConciliado.value = vbChecked, "marcado", "desmarcado"))
            End If
            
            If .Pagamento <> edtDataPagamento.Data Then
                Call objLogLancDup.InsertCustomMsg(strStdMsg, "Pagamento", .Pagamento, Format(edtDataPagamento.Data, "DD/MM/YYYY"))
            End If
            
            If .Liberacao <> edtDataLiberacao.Data Then
                Call objLogLancDup.InsertCustomMsg(strStdMsg, "Liberação", .Liberacao, Format(edtDataLiberacao.Data, "DD/MM/YYYY"))
            End If
            
            If enuTabela = Duplicata Then
                If .cd_operacao_baixa <> etxOpContabilDupl.valorInteiro Then
                    Call objLogLancDup.InsertCustomMsg(strStdMsg, "Op. Contábil - Baixa", .cd_operacao_baixa, etxOpContabilDupl.valorInteiro)
                End If
            Else
                If .cd_operacao_baixa <> etxOpContabilLanc.valorInteiro Then
                    Call objLogLancDup.InsertCustomMsg(strStdMsg, "Op. Contábil - Baixa", .cd_operacao_baixa, etxOpContabilLanc.valorInteiro)
                End If
            End If
        End With
    End If
    
erro:
    Set objLogLancDup = Nothing
End Sub
