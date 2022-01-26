VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{B9A5F0AF-8C1F-4070-9B08-47ED989F52B3}#1.0#0"; "ReportXWizard.ocx"
Begin VB.MDIForm fMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "FOX - Financeiro"
   ClientHeight    =   6885
   ClientLeft      =   4590
   ClientTop       =   4545
   ClientWidth     =   14040
   HelpContextID   =   6
   Icon            =   "frmFinanceiro.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrPerfil 
      Interval        =   60000
      Left            =   7410
      Top             =   2130
   End
   Begin VB.Timer tmrMenu 
      Interval        =   1000
      Left            =   7410
      Top             =   1620
   End
   Begin VB.Timer tmrAtivacao 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   7410
      Top             =   1125
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   7395
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":07DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":0D78
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":1312
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":18AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":1E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":219A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":26EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":2B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":2C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":31AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceiro.frx":35FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ReportX_Wizard.ReportXW RXWLIB 
      Left            =   8055
      Top             =   435
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin ACTIVESKINLibCtl.Skin SKN 
      Left            =   8715
      OleObjectBlob   =   "frmFinanceiro.frx":3A52
      Top             =   435
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sair"
            Object.ToolTipText     =   "Fechar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "novo"
            Object.ToolTipText     =   "Novo Registro"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deletar"
            Object.ToolTipText     =   "Excluir Registro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pesquisar"
            Object.ToolTipText     =   "Pesquisar Registro"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "localizar"
            Object.ToolTipText     =   "Localizar Registro"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
            Object.ToolTipText     =   "Editar Registro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "primeiro"
            Object.ToolTipText     =   "Primeiro Registro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "anterior"
            Object.ToolTipText     =   "Registro Anterior"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "proximo"
            Object.ToolTipText     =   "Próximo Registro"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ultimo"
            Object.ToolTipText     =   "Último Registro"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   11475
         MouseIcon       =   "frmFinanceiro.frx":3CAC
         MousePointer    =   99  'Custom
         Picture         =   "frmFinanceiro.frx":3FB6
         ScaleHeight     =   15.75
         ScaleMode       =   2  'Point
         ScaleWidth      =   87.75
         TabIndex        =   2
         Top             =   8
         Width           =   1760
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6570
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12966
            MinWidth        =   12966
            Text            =   "Hint"
            TextSave        =   "Hint"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5645
            MinWidth        =   5645
            Text            =   "Nome da Empresa"
            TextSave        =   "Nome da Empresa"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Usuário do Sistema"
            TextSave        =   "Usuário do Sistema"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1589
            MinWidth        =   1589
            Text            =   "Versão"
            TextSave        =   "Versão"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "EBS Sistemas"
            TextSave        =   "EBS Sistemas"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Fox.EBSSBCenter EBSSBCenter 
      Align           =   3  'Align Left
      Height          =   6210
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   10954
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      LarguraMinima   =   250
      Begin MSComctlLib.ImageList imlMenu 
         Left            =   2220
         Top             =   3570
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFinanceiro.frx":A363
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFinanceiro.frx":A67B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFinanceiro.frx":A8BA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin Fox.EBSSideTab eST 
         Height          =   345
         Index           =   5
         Left            =   90
         TabIndex        =   14
         Top             =   1860
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   609
         Caption         =   "Ajuda"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         ForeColor       =   8421504
         MaxHeight       =   2000
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   12632256
         Begin MSComctlLib.TreeView tvwAjuda 
            Height          =   1155
            HelpContextID   =   2425
            Left            =   60
            TabIndex        =   15
            Top             =   420
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   2037
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imlMenu"
            Appearance      =   0
         End
      End
      Begin Fox.EBSSideTab eST 
         Height          =   345
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   1530
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   609
         Caption         =   "Utilitários"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         ForeColor       =   8421504
         MaxHeight       =   2000
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   12632256
         Begin MSComctlLib.TreeView tvwUtilitarios 
            Height          =   1215
            HelpContextID   =   2102
            Left            =   60
            TabIndex        =   13
            Top             =   390
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   2143
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imlMenu"
            Appearance      =   0
         End
      End
      Begin Fox.EBSSideTab eST 
         Height          =   345
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   1200
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   609
         Caption         =   "Relatórios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         ForeColor       =   8421504
         MaxHeight       =   10000
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   12632256
         Begin MSComctlLib.TreeView tvwRelatorios 
            Height          =   3285
            HelpContextID   =   2084
            Left            =   30
            TabIndex        =   11
            Top             =   360
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   5794
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imlMenu"
            Appearance      =   0
         End
      End
      Begin Fox.EBSSideTab eST 
         Height          =   345
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   870
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   609
         Caption         =   "Consultas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         ForeColor       =   8421504
         MaxHeight       =   2000
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   12632256
         Begin MSComctlLib.TreeView tvwConsultas 
            Height          =   1215
            HelpContextID   =   2080
            Left            =   30
            TabIndex        =   9
            Top             =   360
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   2143
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imlMenu"
            Appearance      =   0
         End
      End
      Begin Fox.EBSSideTab eST 
         Height          =   345
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   540
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   609
         Caption         =   "Módulos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         ForeColor       =   8421504
         MaxHeight       =   2000
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   12632256
         Begin MSComctlLib.TreeView tvwModulos 
            Height          =   1395
            HelpContextID   =   2055
            Left            =   60
            TabIndex        =   7
            Top             =   390
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   2461
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imlMenu"
            Appearance      =   0
         End
      End
      Begin Fox.EBSSideTab eST 
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   609
         Caption         =   "Cadastros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8,25
         ForeColor       =   8421504
         MaxHeight       =   2000
         BorderColor     =   12957347
         ContractedForeColor=   8421504
         ExpandedForeColor=   12632256
         Begin MSComctlLib.TreeView tvwCadastro 
            Height          =   1215
            HelpContextID   =   2030
            Left            =   60
            TabIndex        =   5
            Top             =   420
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   2143
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            ImageList       =   "imlMenu"
            Appearance      =   0
         End
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastros"
      HelpContextID   =   2030
      Begin VB.Menu mnuCadGeral 
         Caption         =   "&Geral"
         HelpContextID   =   2031
         Begin VB.Menu mnuCadGerEmpresa 
            Caption         =   "&Empresas"
            HelpContextID   =   2032
         End
         Begin VB.Menu mnuCadGerEmpPotencial 
            Caption         =   "Empresas &Potenciais"
            HelpContextID   =   2033
         End
         Begin VB.Menu mnuCadGerTraco1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCadGerRamo 
            Caption         =   "&Ramos de Atividade"
            HelpContextID   =   2034
         End
         Begin VB.Menu mnuCadMunicipios 
            Caption         =   "&Municípios"
            HelpContextID   =   2862
         End
         Begin VB.Menu mnuCadGerEstado 
            Caption         =   "E&stados"
            HelpContextID   =   2035
         End
         Begin VB.Menu mnuCadGerRegiao 
            Caption         =   "Re&giões"
            HelpContextID   =   2036
         End
         Begin VB.Menu mnuCadGerPais 
            Caption         =   "&Países"
            HelpContextID   =   2037
         End
         Begin VB.Menu mnuCadProcedencias 
            Caption         =   "Procedências"
            HelpContextID   =   2688
         End
      End
      Begin VB.Menu mnuCadCentroCusto 
         Caption         =   "&Centros de Custos"
         HelpContextID   =   2038
      End
      Begin VB.Menu mnuCadGrupoConta 
         Caption         =   "&Grupos de Contas"
         HelpContextID   =   2040
      End
      Begin VB.Menu mnuCadHistBancario 
         Caption         =   "&Históricos Bancários"
         HelpContextID   =   3015
      End
      Begin VB.Menu mnuCadContas 
         Caption         =   "C&ontas"
         HelpContextID   =   2039
      End
      Begin VB.Menu mnuCadBancoCaixa 
         Caption         =   "Banco/Caixa"
         HelpContextID   =   2042
      End
      Begin VB.Menu mnuCarteira 
         Caption         =   "Carteira"
         HelpContextID   =   2926
      End
      Begin VB.Menu mnuCamposEspeciais 
         Caption         =   "Campos Especiais"
         HelpContextID   =   3021
      End
      Begin VB.Menu mnuCadCamara 
         Caption         =   "C&âmara"
         HelpContextID   =   2853
      End
      Begin VB.Menu mnuCadForPagamento 
         Caption         =   "&Formas de Pagamento"
         HelpContextID   =   2801
      End
      Begin VB.Menu mnuOpFinanceira 
         Caption         =   "Operação Financeira"
         HelpContextID   =   2797
      End
      Begin VB.Menu mnuCadIndice 
         Caption         =   "&Índices"
         HelpContextID   =   2043
         Begin VB.Menu mnuCadInTaxasBancaria 
            Caption         =   "&Taxas Bancárias"
            HelpContextID   =   2044
         End
         Begin VB.Menu mnuCadInDespFinanceira 
            Caption         =   "&Despesas Financeiras"
            HelpContextID   =   2045
         End
         Begin VB.Menu mnuCadIndTraco 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCadIndMoeda 
            Caption         =   "&Moedas"
            HelpContextID   =   2046
         End
         Begin VB.Menu mnuCadIndCotacao 
            Caption         =   "&Cotações"
            HelpContextID   =   2047
         End
      End
      Begin VB.Menu mnuCadGenerico 
         Caption         =   "&Genéricos"
         HelpContextID   =   2048
         Begin VB.Menu mnuCadGenProjeto 
            Caption         =   "&Projetos"
            HelpContextID   =   2049
         End
         Begin VB.Menu mnuCadGenTipoGlobal 
            Caption         =   "&Tipos Globais"
            HelpContextID   =   2050
         End
         Begin VB.Menu mnuCadGenFeriado 
            Caption         =   "&Feriados"
            HelpContextID   =   2051
         End
         Begin VB.Menu mnuCadGenObservacao 
            Caption         =   "&Observações"
            HelpContextID   =   2052
         End
      End
      Begin VB.Menu mnuCadTraco1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadConfGeral 
         Caption         =   "Configuração Gera&l"
         HelpContextID   =   2053
      End
      Begin VB.Menu mnuCadTraco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCadSair 
         Caption         =   "Sair"
         HelpContextID   =   2423
      End
   End
   Begin VB.Menu mnuModulo 
      Caption         =   "&Módulos"
      HelpContextID   =   2055
      Begin VB.Menu mnuModReceber 
         Caption         =   "Contas a &Receber"
         HelpContextID   =   2056
         Begin VB.Menu mnuModRecLancamentos 
            Caption         =   "&Lançamentos a Receber ou Recebidos"
            HelpContextID   =   2057
         End
         Begin VB.Menu mnuModRecDuplicatas 
            Caption         =   "&Duplicatas a Receber ou Recebidas"
            HelpContextID   =   2058
         End
         Begin VB.Menu mnuModRecBoleto 
            Caption         =   "&Processar Boleto Bancário"
            HelpContextID   =   2059
         End
         Begin VB.Menu mnuCOntasFixasReceber 
            Caption         =   "&Contas Fixas / Previsão"
            HelpContextID   =   2725
         End
         Begin VB.Menu mnuModGertitreceber 
            Caption         =   "&Geração de Titulos a Receber"
            HelpContextID   =   2812
         End
         Begin VB.Menu mnuModRecTraco1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuModRecBaixas 
            Caption         =   "&Baixas"
            HelpContextID   =   2076
         End
         Begin VB.Menu mnuAlteracaoBancoTitulosReceber 
            Caption         =   "&Alteração de Banco em Títulos à Receber"
            HelpContextID   =   2854
         End
      End
      Begin VB.Menu mnuModPagar 
         Caption         =   "Contas a &Pagar"
         HelpContextID   =   2060
         Begin VB.Menu mnuModPagLancamentos 
            Caption         =   "&Lançamentos a Pagar ou Pagos"
            HelpContextID   =   2061
         End
         Begin VB.Menu mnuModPagDuplicatas 
            Caption         =   "&Duplicatas a Pagar ou Pagas"
            HelpContextID   =   2062
         End
         Begin VB.Menu mnuModPagContas 
            Caption         =   "&Contas Fixas / Previsão"
            HelpContextID   =   2063
         End
         Begin VB.Menu mnuModGertitpagar 
            Caption         =   "&Geração de Titulos a Pagar"
            HelpContextID   =   2810
         End
         Begin VB.Menu mnuModPagTraco1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuModPagBaixas 
            Caption         =   "&Baixas"
            HelpContextID   =   2707
         End
         Begin VB.Menu mnuAlteracaoBancoTitulosPagar 
            Caption         =   "&Alteração de Banco em Títulos à Pagar"
            HelpContextID   =   2855
         End
      End
      Begin VB.Menu mnuModBancos 
         Caption         =   "&Bancos"
         HelpContextID   =   2064
         Begin VB.Menu mnuModBanAplicacoes 
            Caption         =   "&Aplicações Financeiras"
            HelpContextID   =   2065
         End
         Begin VB.Menu mnuModBanMovEntrada 
            Caption         =   "Movimentação de &Entrada"
            HelpContextID   =   2066
         End
         Begin VB.Menu mnuModBanMovSaida 
            Caption         =   "Movimentação de &Saída"
            HelpContextID   =   2067
         End
         Begin VB.Menu mnuModBanTranBancaria 
            Caption         =   "&Transferências Bancárias"
            HelpContextID   =   2068
         End
         Begin VB.Menu mnuModBanSaldoBanco 
            Caption         =   "&Saldos bancários"
            HelpContextID   =   2069
         End
         Begin VB.Menu mnuModBanConcBancaria 
            Caption         =   "C&onciliação Bancária"
            HelpContextID   =   2806
         End
         Begin VB.Menu mnuModBanConcBancariaAut 
            Caption         =   "Co&nciliação Bancária Automática"
            HelpContextID   =   3016
         End
         Begin VB.Menu mnuImpDigExtratoBacario 
            Caption         =   "&Importar/Digitar Extrato Bancário"
            HelpContextID   =   3014
         End
         Begin VB.Menu mnuModBanCadCheque 
            Caption         =   "Cadastro de &Cheques"
            HelpContextID   =   2070
         End
         Begin VB.Menu mnuModBanEditaCheque 
            Caption         =   "&Editor de Cheques"
            HelpContextID   =   2072
         End
         Begin VB.Menu mnuReprocessaSaldoBanc 
            Caption         =   "&Reprocessamento de Saldo Bancário"
            HelpContextID   =   3018
         End
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "Caixa"
         HelpContextID   =   2074
         Begin VB.Menu mnuModCaiLibera 
            Caption         =   "&Controle de Liberações"
            HelpContextID   =   2075
         End
         Begin VB.Menu mnuModCaiTraco1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuModCaiDesconto 
            Caption         =   "&Desconto por Pontualidade"
            HelpContextID   =   2077
         End
         Begin VB.Menu mnuSeparador 
            Caption         =   "-"
         End
         Begin VB.Menu mnuContaCorrente 
            Caption         =   "C&onta Corrente"
            HelpContextID   =   2798
         End
      End
      Begin VB.Menu mnuModTraco1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModGerConDupLancamento 
         Caption         =   "&Geração de Controle de Duplicatas e Lançamentos"
         HelpContextID   =   2078
      End
      Begin VB.Menu mnuModMovConferido 
         Caption         =   "&Movimento Conferido"
         HelpContextID   =   2079
      End
      Begin VB.Menu menuModulosTraco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModuloProcessosCNAB 
         Caption         =   "Sistema Cobrança Bancária (CNAB) - Antigo"
         HelpContextID   =   2842
         Begin VB.Menu mnuModulosProcessosCNABPagamento 
            Caption         =   "Pagamento"
            HelpContextID   =   2843
            Begin VB.Menu mnuModuloProcessosCNABDadosFavorecido 
               Caption         =   "Dados Favorecidos"
               HelpContextID   =   2845
            End
            Begin VB.Menu mnuProcessosCNABPagamentoRemessa 
               Caption         =   "Remessa Pagamento"
               HelpContextID   =   2846
            End
            Begin VB.Menu mnuTraco3 
               Caption         =   "-"
            End
            Begin VB.Menu mnuRetornoPagamento 
               Caption         =   "Retorno Pagamento"
               HelpContextID   =   2864
            End
         End
         Begin VB.Menu mnuModuloProcessosCNABRecebimento 
            Caption         =   "Recebimento"
            HelpContextID   =   2844
            Begin VB.Menu MnuKINComunicacaoCadRemessas 
               Caption         =   "&Manutenção de Layouts"
               HelpContextID   =   2693
            End
            Begin VB.Menu mnuKINEnvioCobrancas 
               Caption         =   "&Envio de Cobranças"
               HelpContextID   =   2694
            End
            Begin VB.Menu mnuKINEnvioPagamento 
               Caption         =   "Envio de &Pagamentos"
               Enabled         =   0   'False
               HelpContextID   =   2695
               Visible         =   0   'False
            End
            Begin VB.Menu mnuKINComunicacaoRemessaImpExp 
               Caption         =   "Exportação e Importação de Layouts"
               HelpContextID   =   2696
            End
            Begin VB.Menu mnuComunicacoesRetornoBancario 
               Caption         =   "Retorno Bancário"
               HelpContextID   =   2691
            End
         End
      End
      Begin VB.Menu mnuProcessosCobrebem 
         Caption         =   "Sistema Cobrança Bancária - Novo"
         HelpContextID   =   2920
         Begin VB.Menu mnuRecebimento 
            Caption         =   "Recebimento"
            HelpContextID   =   2921
            Begin VB.Menu mnuEmissaoBoleto 
               Caption         =   "Emissão de Boleto"
               HelpContextID   =   2922
            End
            Begin VB.Menu mnuEmissaoRemessa 
               Caption         =   "Emissão de Remessa"
               HelpContextID   =   2923
            End
            Begin VB.Menu mnuConfirmacaoRetorno 
               Caption         =   "Confirmação de Retorno"
               HelpContextID   =   2924
            End
         End
      End
      Begin VB.Menu mnuModuloTraco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIntegracao 
         Caption         =   "Integração"
         HelpContextID   =   2642
         Begin VB.Menu mnuIntConfig 
            Caption         =   "Configuração"
            HelpContextID   =   2643
            Begin VB.Menu mnuIntApropImpostos 
               Caption         =   "Apropriação de Impostos"
               HelpContextID   =   2644
            End
            Begin VB.Menu mnuOpContabeis 
               Caption         =   "Operações Contabeis"
               HelpContextID   =   2645
            End
            Begin VB.Menu mnuMatrixContab 
               Caption         =   "Matriz de Contabilização"
               HelpContextID   =   2646
            End
         End
         Begin VB.Menu mnuGeracao 
            Caption         =   "Geração"
            HelpContextID   =   2650
            Begin VB.Menu mnuIntContabil 
               Caption         =   "Integração Contábil"
               HelpContextID   =   2651
            End
            Begin VB.Menu mnuIntFiscal 
               Caption         =   "Integração Fiscal"
               HelpContextID   =   2718
            End
         End
         Begin VB.Menu mnuIntegracaoSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAtualizacaoOperacoes 
            Caption         =   "Alteração de Operações Contábeis"
            HelpContextID   =   2746
         End
      End
   End
   Begin VB.Menu mnuConsulta 
      Caption         =   "&Consultas"
      HelpContextID   =   2080
      Begin VB.Menu mnuConLanDuplicata 
         Caption         =   "&Lançamentos e Duplicatas"
         HelpContextID   =   2081
      End
      Begin VB.Menu mnuConSaldos 
         Caption         =   "&Saldos"
         HelpContextID   =   2082
      End
      Begin VB.Menu mnuConTitAtraso 
         Caption         =   "&Títulos em Atraso"
         HelpContextID   =   2083
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      HelpContextID   =   2084
      Begin VB.Menu mnuRelDupLancamento 
         Caption         =   "&Duplicatas e Lançamentos"
         HelpContextID   =   2085
      End
      Begin VB.Menu mnuRelTitRecAtrSintetico 
         Caption         =   "Duplicatas e Lançamentos em Atraso (&Sintético)"
         HelpContextID   =   2086
      End
      Begin VB.Menu mnuRelTitRecAtrAnalitico 
         Caption         =   "Duplicatas e Lançamentos em Atraso (&Analítico)"
         HelpContextID   =   2087
      End
      Begin VB.Menu mnuRelBolPreImpresso 
         Caption         =   "&Boleto Pré-Impresso"
         HelpContextID   =   2088
      End
      Begin VB.Menu mnuRelBordero 
         Caption         =   "B&orderô"
         HelpContextID   =   2089
      End
      Begin VB.Menu mnuRelRegDuplicata 
         Caption         =   "&Registro de Duplicatas"
         HelpContextID   =   2090
      End
      Begin VB.Menu mnuRelTabelas 
         Caption         =   "&Tabelas"
         HelpContextID   =   2091
      End
      Begin VB.Menu mnuRelTraco1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRelMovCaixa 
         Caption         =   "&Movimento de Caixa"
         HelpContextID   =   2092
      End
      Begin VB.Menu mnuRelFluCaiGeral 
         Caption         =   "Fluxo de Caixa &Geral"
         HelpContextID   =   2093
      End
      Begin VB.Menu mnuRelFluCaiConGrupo 
         Caption         =   "Fluxo de Caixa por &Conta e Grupo"
         HelpContextID   =   2094
      End
      Begin VB.Menu mnuRelConFinanceiro 
         Caption         =   "Controle &Financeiro"
         HelpContextID   =   2095
      End
      Begin VB.Menu mnuRelRazAuxiliar 
         Caption         =   "R&azão Auxiliar"
         HelpContextID   =   2096
      End
      Begin VB.Menu mnuRelCheques 
         Caption         =   "Che&ques"
         HelpContextID   =   2097
      End
      Begin VB.Menu mnuRelRecibo 
         Caption         =   "R&ecibos"
         HelpContextID   =   2098
      End
      Begin VB.Menu mnuRelTraco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRelExtBancario 
         Caption         =   "E&xtrato Bancário"
         HelpContextID   =   2099
      End
      Begin VB.Menu mnuRelAplFinanceiras 
         Caption         =   "A&plicações Financeiras"
         HelpContextID   =   2100
      End
      Begin VB.Menu mnuRelTranBancaria 
         Caption         =   "&Transferências Bancárias"
         HelpContextID   =   2101
      End
      Begin VB.Menu mnuRelEmpresas 
         Caption         =   "&Empresas"
         HelpContextID   =   2689
      End
      Begin VB.Menu mnuRelEtoquetas 
         Caption         =   "Etiquetas"
         HelpContextID   =   2690
      End
   End
   Begin VB.Menu mnuUtilitario 
      Caption         =   "&Utilitários"
      HelpContextID   =   2102
      Begin VB.Menu mnuUtiRelatorio 
         Caption         =   "&Relatórios ERC"
         HelpContextID   =   2103
      End
      Begin VB.Menu mnuUtiImpressora 
         Caption         =   "&Impressoras"
         HelpContextID   =   2104
      End
      Begin VB.Menu mnuUtiCalculadora 
         Caption         =   "&Calculadora"
         HelpContextID   =   2105
      End
      Begin VB.Menu mnuUtiConstrutor 
         Caption         =   "C&onstrutor de Consultas"
         HelpContextID   =   2106
      End
      Begin VB.Menu mnuTraco1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRotinasEspecificas 
         Caption         =   "Rotinas Específicas"
         Begin VB.Menu mnuGeracaoIntegracaoBalanSet 
            Caption         =   "Integração Contábil (Balan-Set)"
         End
      End
      Begin VB.Menu mnuComunicacao 
         Caption         =   "C&omunicação"
         HelpContextID   =   2472
         Begin VB.Menu mnuImportaExportaTab 
            Caption         =   "Importa/Exporta Tabelas"
            HelpContextID   =   2474
         End
         Begin VB.Menu mnuImportaDuplicata 
            Caption         =   "Importação Duplicatas"
            HelpContextID   =   2876
         End
      End
      Begin VB.Menu mnuTraco2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCOnfiguracoes 
         Caption         =   "Con&figurações"
         HelpContextID   =   2478
         Begin VB.Menu mnuConfigRelERC 
            Caption         =   "Relatórios ERC"
            HelpContextID   =   2479
         End
         Begin VB.Menu mnuConfigServidorEmail 
            Caption         =   "Servidor de E-mail"
            HelpContextID   =   2480
         End
         Begin VB.Menu mnuSistema 
            Caption         =   "Sistema"
            HelpContextID   =   2481
         End
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      HelpContextID   =   2425
      Begin VB.Menu mnuAjuAjuda 
         Caption         =   "&Ajuda                    F1"
         HelpContextID   =   2426
      End
      Begin VB.Menu mnuAjuAcessoSite 
         Caption         =   "Suporte &Técnico"
         HelpContextID   =   2428
         Begin VB.Menu mnuSuporteOnLine 
            Caption         =   "Suporte on-line"
            HelpContextID   =   2429
         End
         Begin VB.Menu mnuAjudaSuporteRemoto 
            Caption         =   "Suporte &Remoto - Ammyy"
            HelpContextID   =   2832
         End
      End
      Begin VB.Menu mnuAjuTraco1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjuSobre 
         Caption         =   "&Sobre"
         HelpContextID   =   2425
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'pt. 82050 - Dulcino Júnior
Private mintContador As Integer
Private blnEscreve As Boolean

Private Sub MDIForm_Load()
    Dim obj As Object
    
    'Ivo Sousa (08/07/2013) - Deleta o config do temporaio do usuário, caso ele já exista
    If ArquivoExiste(App.Path & "\..\Configurações\Config_" & Replace(AlteraCaracterAcentuado(UserWindowsID), ".", "") & ".ini") Then
        Set objFso = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFso.GetFile(App.Path & "\..\Configurações\Config_" & Replace(AlteraCaracterAcentuado(UserWindowsID), ".", "") & ".ini")
        Call objFile.Delete(True)
    End If
    
    Call InicializaModulo(Me)   'Inicializar a conexao do Modulo.
    
    'Ivo Sousa (08/07/2013) - Copia o Config da pasta local para a pasta temporária do usuário
    If ArquivoExiste(App.Path & "\..\Configurações\Config.ini") Then
        Call FileCopy(App.Path & "\..\Configurações\Config.ini", App.Path & "\..\Configurações\Config_" & Replace(AlteraCaracterAcentuado(UserWindowsID), ".", "") & ".ini")
    End If
    
    Set InstanciaMenu = New clsMenuFinanceiro
        
    Call ConfiguraMainForm 'Configura mensagens na StatusBar e etc...
    Call addPermissaoMenus(Me, retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO)
    'Protocolo Nr 96268  - Carlos Felippe Vernizze - 24/09/2010
    'Função para carregamento dos menus dentro dos TreeView.
    For Each obj In Me.Controls
        If TypeOf obj Is TreeView Then
            Aplicacao.Connect
            Call CarregaMenu(obj)
            Aplicacao.Disconnect
        End If
    Next
    mnuCadastro.Visible = False
    mnuModulo.Visible = False
    mnuConsulta.Visible = False
    mnuRelatorios.Visible = False
    mnuUtilitario.Visible = False
    mnuAjuda.Visible = False
    
    frmSideBar.Show
    
    mintContador = 30
    tmrAtivacao_Timer
    
    #If DESENV = 0 Then
    ModGeral.configuraAmbientePeloReadOnly
    #End If
    
    'Projeto: 74365 - Ueder Budni (22/04/2015)
    Set objSageAnalytics = New clsSageAnalytics
    
    Call AbrirAlertaContasVencidas
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Projeto: 74365 - Ueder Budni (22/04/2015)
    Set objSageAnalytics = Nothing
    
    Unload frmAlerta
End Sub

Private Sub mnuAjuAjuda_Click()
    Dim oHelpHtml As New clsHelp
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub mnuAjudaSuporteRemoto_Click()
     If Dir(CaminhoPasta(pastaProgramas) & "\TeamViewer.exe", vbArchive) <> "" Then
        Call Shell(CaminhoPasta(pastaProgramas) & "\TeamViewer.exe", vbNormalFocus)
    Else
        If MsgBox("A aplicação de Suporte Remoto Team Viewer não foi encontrada. Deseja baixá-la?", vbQuestion + vbYesNo, NomeModulo) = vbYes Then
            If DownloadFile(LINK_TEAMVIEWER, CaminhoPasta(pastaProgramas) & "\TeamViewer.exe") Then
                MsgBox "Arquivo baixado com sucesso!" & Chr(13) & "Clique na rotina novamente para abrir o Team Viewer.", vbInformation, NomeModulo
            Else
                MsgBox "Erro ao baixar o arquivo.", vbCritical, NomeModulo
            End If
        End If
    End If
End Sub

Private Sub mnuAjuSobre_Click()
    frmFoxSobre.Show vbModal
End Sub

Private Sub mnuAlteracaoBancoTitulosPagar_Click()
    mstrPagRecAlteraBanco = "P"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuAlteracaoBancoTitulosPagar.HelpContextID, frmAltercaoBancoTitulos.name, "Alteração do Banco em Títulos à Pagar")
    Call mostrarForm(frmAltercaoBancoTitulos, mnuAlteracaoBancoTitulosPagar.HelpContextID)
End Sub

Private Sub mnuAlteracaoBancoTitulosReceber_Click()
    mstrPagRecAlteraBanco = "R"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuAlteracaoBancoTitulosReceber.HelpContextID, frmAltercaoBancoTitulos.name, "Alteração do Banco em Títulos à Receber")
    Call mostrarForm(frmAltercaoBancoTitulos, mnuAlteracaoBancoTitulosReceber.HelpContextID)
End Sub

Private Sub mnuCadBanco_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadBanco.HelpContextID, frmBancos.name, "Cadastro de Bancos")
    Call mostrarForm(frmBancos, mnuCadBanco.HelpContextID)
End Sub

Private Sub mnuAtualizacaoOperacoes_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuAtualizacaoOperacoes.HelpContextID, frmAlteracaoOperacaoContabil.name, "Atualização de Operações")
    Call mostrarForm(frmAlteracaoOperacaoContabil, mnuAtualizacaoOperacoes.HelpContextID)
End Sub

Private Sub mnuCadBancoCaixa_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadBancoCaixa.HelpContextID, frmBancos.name, "Cadastro de Banco/Caixa")
    Call mostrarForm(frmBancos, mnuCadBancoCaixa.HelpContextID)
End Sub

Private Sub mnuCadCamara_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadCamara.HelpContextID, frmCadastroCamara.name, "Cadastro de Câmara")
    Call mostrarForm(frmCadastroCamara, mnuCadCamara.HelpContextID)
End Sub

Private Sub mnuCadCentroCusto_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadCentroCusto.HelpContextID, frmCusto.name, "Cadastro de Centros de Custos")
    Call mostrarForm(frmCusto, mnuCadCentroCusto.HelpContextID)
End Sub

Private Sub mnuCadConfFinanceiro_Click()
    FrmConfCad.Configura "Duplicatas"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadConfFinanceiro.HelpContextID, FrmConfCad.name, "Configurações Financeiras")
    Call mostrarForm(FrmConfCad, mnuCadConfFinanceiro.HelpContextID, True)
End Sub

Private Sub mnuCadConfGeral_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadConfGeral.HelpContextID, FrmConfCad.name, "Configurações Gerais")
    FrmConfCad.Configura "TODAS"
    Call mostrarForm(FrmConfCad, mnuCadConfGeral.HelpContextID, False)
End Sub

Private Sub mnuCadContas_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadContas.HelpContextID, frmContas.name, "Cadastro de Contas")
    Call mostrarForm(frmContas, mnuCadContas.HelpContextID)
End Sub

Private Sub mnuCadForPagamento_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloVar, mnuCadForPagamento.HelpContextID, FormaPagamento.name, "Cadastro de Formas de Pagamento")
    Call mostrarForm(FormaPagamento, mnuCadForPagamento.HelpContextID)
End Sub

Private Sub mnuCadGenFeriado_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGenFeriado.HelpContextID, frmFeriados.name, "Feriados")
    Call mostrarForm(frmFeriados, mnuCadGenFeriado.HelpContextID)
End Sub

Private Sub mnuCadGenObservacao_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGenObservacao.HelpContextID, frmObservacoes.name, "Observações")
    Call mostrarForm(frmObservacoes, mnuCadGenObservacao.HelpContextID)
End Sub

Private Sub mnuCadGenProjeto_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGenProjeto.HelpContextID, frmProjeto.name, "Projetos")
    Call mostrarForm(frmProjeto, mnuCadGenProjeto.HelpContextID)
End Sub

Private Sub mnuCadGenTipoGlobal_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGenTipoGlobal.HelpContextID, frmTiposGlobais.name, "Tipos Globais")
    Call mostrarForm(frmTiposGlobais, mnuCadGenTipoGlobal.HelpContextID)
    frmTiposGlobais.ZOrder
End Sub

Private Sub mnuCadGerEmpresa_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGerEmpresa.HelpContextID, frmEmpresas.name, "Cadastro de Empresas")
    Call mostrarForm(frmEmpresas, mnuCadGerEmpresa.HelpContextID)
End Sub

Private Sub mnuCadGerEmpPotencial_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGerEmpPotencial.HelpContextID, frmEmpresasPotenciais.name, "Cadastro de Empresas Potenciais")
    Call mostrarForm(frmEmpresasPotenciais, mnuCadGerEmpPotencial.HelpContextID)
    frmEmpresasPotenciais.ZOrder
End Sub

Private Sub mnuCadGerEstado_Click()
    fblnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGerEstado.HelpContextID, frmEstados.name, "Cadastro de Estados")
    Call mostrarForm(frmEstados, mnuCadGerEstado.HelpContextID)
End Sub

Private Sub mnuCadGerPais_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGerPais.HelpContextID, frmPaises.name, "Cadastro de Países")
    Call mostrarForm(frmPaises, mnuCadGerPais.HelpContextID)
End Sub

Private Sub mnuCadGerRamo_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGerRamo.HelpContextID, frmRamos.name, "Cadastro de Ramos de Atividade")
    Call mostrarForm(frmRamos, mnuCadGerRamo.HelpContextID)
End Sub

Private Sub mnuCadGerRegiao_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGerRegiao.HelpContextID, frmRegioes.name, "Cadastro de Regiões")
    Call mostrarForm(frmRegioes, mnuCadGerRegiao.HelpContextID)
End Sub

Private Sub mnuCadGrupoConta_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadGrupoConta.HelpContextID, frmGrupos.name, "Cadastro de Grupos de Contas")
    Call mostrarForm(frmGrupos, mnuCadGrupoConta.HelpContextID)
End Sub

Private Sub mnuCadHistBancario_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadHistBancario.HelpContextID, frmCadHistBancario.name, "Cadastro de Históricos Bancários")
    Call mostrarForm(frmCadHistBancario, mnuCadHistBancario.HelpContextID)
End Sub

Private Sub mnuCadIndCotacao_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadIndCotacao.HelpContextID, fCotacoes.name, "Cadastro de Cotações")
    Call mostrarForm(fCotacoes, mnuCadIndCotacao.HelpContextID)
End Sub

Private Sub mnuCadInDespFinanceira_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadInDespFinanceira.HelpContextID, frmDespesas.name, "Despesas Financeiras")
    frmDespesas.tipo = "D"      '// D == Despesas Financeiras
    Call mostrarForm(frmDespesas, mnuCadInDespFinanceira.HelpContextID)
End Sub

Private Sub mnuCadIndMoeda_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadIndMoeda.HelpContextID, fMoedas.name, "Cadastro de Moedas")
    Call mostrarForm(fMoedas, mnuCadIndMoeda.HelpContextID)
End Sub

Private Sub mnuCadInTaxasBancaria_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadInTaxasBancaria.HelpContextID, frmDespesas.name, "Cadastro de Taxas Bancárias")
    frmDespesas.tipo = "T"      '// T == Taxas Bancárias
    Call mostrarForm(frmDespesas, mnuCadInTaxasBancaria.HelpContextID)
End Sub

Private Sub mnuCadMunicipios_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadMunicipios.HelpContextID, frmCadMunicipio.name, "Cadastro de Municípios")
    Call mostrarForm(frmCadMunicipio, mnuCadMunicipios.HelpContextID)
End Sub

Private Sub mnuCadProcedencias_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloCompras, mnuCadProcedencias.HelpContextID, frmProcedencias.name, "Cadastro de Procedências")
    Call mostrarForm(frmProcedencias, mnuCadProcedencias.HelpContextID)
    frmProcedencias.ZOrder
End Sub

Private Sub mnuCadSair_Click()
    End
End Sub

Private Sub mnuCarteira_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCarteira.HelpContextID, frmCarteira.name, "Carteira")
    Call mostrarForm(frmCarteira, mnuCarteira.HelpContextID, False)
End Sub

Private Sub mnuComunicacoesRetornoBancario_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuComunicacoesRetornoBancario.HelpContextID, frmRetornoBancario.name, "Retorno Bancário")
    Call mostrarForm(frmRetornoBancario, mnuComunicacoesRetornoBancario.HelpContextID)
End Sub

Private Sub mnuConfigRelERC_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuConfigRelERC.HelpContextID, frmConfigFRE.name, "Configurações Relatórios ERC")
    Call mostrarForm(frmConfigFRE, mnuConfigRelERC.HelpContextID, True)
End Sub

Private Sub mnuConfigServidorEmail_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuConfigServidorEmail.HelpContextID, frmConfEmail.name, "Configurações de Servidor de Email")
    Call mostrarForm(frmConfEmail, mnuConfigServidorEmail.HelpContextID, False)
End Sub

Private Sub mnuConfirmacaoRetorno_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuConfirmacaoRetorno.HelpContextID, frmRetorno.name, "Carregar Retorno")
    Call mostrarForm(frmRetorno, mnuConfirmacaoRetorno.HelpContextID, False)
End Sub

Private Sub mnuConLanDuplicata_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuConLanDuplicata.HelpContextID, frmConsultaKIF.name, "Consulta de Lançamentos e Duplicatas")
    Call mostrarForm(frmConsultaKIF, mnuConLanDuplicata.HelpContextID)
End Sub

Private Sub mnuConSaldos_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuConSaldos.HelpContextID, fdConsultasKIF.name, "Consulta de Saldos")
    Call mostrarForm(fdConsultasKIF, mnuConSaldos.HelpContextID)
End Sub

Private Sub mnuContaCorrente_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuContaCorrente.HelpContextID, frmLancamentoContaCorrente.name, "Lançamento de Conta Corrente")
    Call mostrarForm(frmLancamentoContaCorrente, mnuContaCorrente.HelpContextID)
End Sub

Private Sub mnuCOntasFixasReceber_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCOntasFixasReceber.HelpContextID, frmContasFixas.name, "Cadastro de Contas Fixas/Previsão - Contas a Receber")
    frmContasFixas.Caption = "Cadastro de Contas Fixas - Contas a Receber"
    frmContasFixas.cmdSair.Enabled = False
    frmContasFixas.Configure 1 ' A receber = 1
    frmContasFixas.cmdSair.Enabled = True
    Call mostrarForm(frmContasFixas, mnuCOntasFixasReceber.HelpContextID)
End Sub

Private Sub mnuConTitAtraso_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuConTitAtraso.HelpContextID, frmConsultaAtrasos.name, "Consulta de Títulos em Atraso")
    Call mostrarForm(frmConsultaAtrasos, mnuConTitAtraso.HelpContextID)
End Sub

Private Sub mnuExportaModelos_Click()
    If (MainImportacao) Then
            frmExpModelo.Show vbModal
    End If
End Sub

Private Sub mnuEmissaoBoleto_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuEmissaoBoleto.HelpContextID, frmBoleto.name, "Emissão de Boleto")
    Call mostrarForm(frmBoleto, mnuEmissaoBoleto.HelpContextID, False)
End Sub

Private Sub mnuEmissaoRemessa_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuEmissaoRemessa.HelpContextID, frmRemessa.name, "Geração de Remessa")
    Call mostrarForm(frmRemessa, mnuEmissaoRemessa.HelpContextID, False)
End Sub

Private Sub mnuGeracaoIntegracaoBalanSet_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuGeracaoIntegracaoBalanSet.HelpContextID, frmGeracaoFinanceiroBalanset.name, "Geração de integração fiscal BalanSet")
    Call mostrarForm(frmGeracaoFinanceiroBalanset, mnuGeracaoIntegracaoBalanSet.HelpContextID)
End Sub

'Projeto: 61827 - Desenv.: 62690 - Ueder Budni (12/01/2015)
Private Sub mnuImpDigExtratoBacario_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuImpDigExtratoBacario.HelpContextID, frmImpDigExtratoBancario.name, "Importar/Digitar Extrato Bancário")
    Call mostrarForm(frmImpDigExtratoBancario, mnuImpDigExtratoBacario.HelpContextID, False)
End Sub

Private Sub mnuImportaDuplicata_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuImportaDuplicata.HelpContextID, frmImportarDuplicata.name, "Importar Duplicatas")
    Call mostrarForm(frmImportarDuplicata, mnuImportaDuplicata.HelpContextID)
End Sub

Private Sub mnuImportaExportaTab_Click()
    Dim sMDB As String
    
    'Salvo o caminho do banco de dados
    If gTipoDB = Access Then
        sMDB = GlobalDataBase.name
    Else
        sMDB = GlobalDataBase
    End If
    frmImpExpTabelas.strPathBancoAtual = sMDB
    frmImpExpTabelas.Icon = Me.Icon
    frmImpExpTabelas.Show
End Sub

Private Sub mnuImportaModelo_Click()
    If (MainImportacao) Then
        frmImpModelo.Show vbModal
    End If
End Sub

Private Sub mnuIntApropImpostos_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuIntApropImpostos.HelpContextID, frmApropriacaoImpostos.name, "Apropriação de Impostos")
    Call mostrarForm(frmApropriacaoImpostos, mnuIntApropImpostos.HelpContextID)
End Sub

Private Sub mnuIntContabil_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuIntContabil.HelpContextID, frmGeracaoContabil.name)
    Call mostrarForm(frmGeracaoContabil, mnuIntContabil.HelpContextID)
    frmGeracaoContabil.ZOrder
End Sub

Private Sub mnuIntFiscal_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuIntFiscal.HelpContextID, frmGeracaoArqIntFiscal.name, "Geração de Arquivos para a Integração Fiscal")
    Call mostrarForm(frmGeracaoArqIntFiscal, mnuIntFiscal.HelpContextID)
End Sub

Private Sub MnuKINComunicacaoCadRemessas_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, MnuKINComunicacaoCadRemessas.HelpContextID, frmRemessaBancaria.name, "Cadastro de Remessas Bancárias")
    Call mostrarForm(frmRemessaBancaria, MnuKINComunicacaoCadRemessas.HelpContextID)
    frmRemessaBancaria.ZOrder
End Sub

Private Sub mnuKINComunicacaoRemessaImpExp_Click()
'    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuKINComunicacaoRemessaImpExp.HelpContextID, fexpimpRemessaBancaria.name, "Exportação e Importação de Layout")
'    Call mostrarForm(fexpimpRemessaBancaria, mnuKINComunicacaoRemessaImpExp.HelpContextID)
'    fexpimpRemessaBancaria.ZOrder
End Sub

Private Sub mnuKINEnvioCobrancas_Click()
    frmCobranca.sPagRec = "R"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuKINEnvioCobrancas.HelpContextID, frmCobranca.name, "Envio de Cobranças")
    Call mostrarForm(frmCobranca, mnuKINEnvioCobrancas.HelpContextID)
    frmCobranca.ZOrder
End Sub

Private Sub mnuKINEnvioPagamento_Click()
    frmCobranca.sPagRec = "P"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuKINEnvioPagamento.HelpContextID, frmCobranca.name, "Envio de Pagamento")
    Call mostrarForm(frmCobranca, mnuKINEnvioPagamento.HelpContextID)
    frmCobranca.ZOrder
End Sub

Private Sub mnuMatrixContab_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuMatrixContab.HelpContextID, frmCadastroMatrizContabil.name, "Matriz de Contabilização")
    Call mostrarForm(frmCadastroMatrizContabil, mnuMatrixContab.HelpContextID)
    frmCadastroMatrizContabil.ZOrder
End Sub

Private Sub mnuModBanAplicacoes_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanAplicacoes.HelpContextID, frmAplicacao.name, "Cadastro de Aplicações Financeiras")
    Call mostrarForm(frmAplicacao, mnuModBanAplicacoes.HelpContextID)
End Sub

Private Sub mnuModBanCadCheque_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanCadCheque.HelpContextID, frmCheque.name, "Cadastro de Cheques")
    Call mostrarForm(frmCheque, mnuModBanCadCheque.HelpContextID)
End Sub

Private Sub mnuModBanConcBancaria_Click()
    'pt. 82528 - Moacir Pfau(06/05/2008)
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanConcBancaria.HelpContextID, frmConciliacaoTitulos.name, "Conciliação Bancária")
    Call mostrarForm(frmConciliacaoTitulos, mnuModBanConcBancaria.HelpContextID)
End Sub
Private Sub mnuModBanConcBancariaAut_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanConcBancaria.HelpContextID, frmConciliacaoTitulosAutomatica.name, "Conciliação Bancária Automática")
    Call mostrarForm(frmConciliacaoTitulosAutomatica, mnuModBanConcBancariaAut.HelpContextID)
End Sub
Private Sub mnuModBanEditaCheque_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanEditaCheque.HelpContextID, frmEditChq.name, "Editor de Cheques")
    Call mostrarForm(frmEditChq, mnuModBanEditaCheque.HelpContextID)
End Sub

Private Sub mnuModBanKCN_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanKCN.HelpContextID, fcalcExportaKCN.name, "Exportação e Envio de Movimentação")
    Call mostrarForm(fcalcExportaKCN, mnuModBanKCN.HelpContextID)
End Sub

Private Sub mnuModBanMovEntrada_Click()
    frmMovBancario.gstrPagRec = "R"
    frmMovBancario.Caption = frmMovBancario.Caption & " de Entrada"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanMovEntrada.HelpContextID, frmMovBancario.name, "Movimento Bancário de Entrada")
    Call mostrarForm(frmMovBancario, mnuModBanMovEntrada.HelpContextID)
End Sub

Private Sub mnuModBanMovSaida_Click()
    frmMovBancario.gstrPagRec = "P"
    frmMovBancario.Caption = frmMovBancario.Caption & " de Saída"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanMovSaida.HelpContextID, frmMovBancario.name, "Movimento Bancário de Entrada Saída")
    Call mostrarForm(frmMovBancario, mnuModBanMovSaida.HelpContextID)
End Sub

Private Sub mnuModBanSaldoBanco_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanSaldoBanco.HelpContextID, frmSldBancos.name, "Saldos Bancários")
    Call mostrarForm(frmSldBancos, mnuModBanSaldoBanco.HelpContextID)
End Sub

Private Sub mnuModBanTranBancaria_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModBanTranBancaria.HelpContextID, frmTransfBanco.name, "Transferência Bancária")
    Call mostrarForm(frmTransfBanco, mnuModBanTranBancaria.HelpContextID)
    frmTransfBanco.ZOrder
End Sub

Private Sub mnuModCaiDesconto_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModCaiDesconto.HelpContextID, frmDescPorPontualidade.name, "Desconto por Pontualidade")
    Call mostrarForm(frmDescPorPontualidade, mnuModCaiDesconto.HelpContextID)
End Sub

Private Sub mnuModCaiLibera_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModCaiLibera.HelpContextID, frmLiberacoes.name, "Controle de Liberações")
    Call mostrarForm(frmLiberacoes, mnuModCaiLibera.HelpContextID)
End Sub

Private Sub mnuModGerConDupLancamento_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModGerConDupLancamento.HelpContextID, fcalcDuplLanc.name, "Geração de Controle de Duplicatas e Lançamentos")
    Call mostrarForm(fcalcDuplLanc, mnuModGerConDupLancamento.HelpContextID)
End Sub

Private Sub mnuModGertitpagar_Click()
    'frmGeracaoTitulosPagar.Show
    Dim bnlEscreve As Boolean
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModGertitpagar.HelpContextID, frmGeracaoTitulosPagar.name, "Geração de Titulos a Pagar")
    Call mostrarForm(frmGeracaoTitulosPagar, mnuModGertitpagar.HelpContextID)
End Sub

Private Sub mnuModGertitreceber_Click()
    Dim bnlEscreve As Boolean
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModGertitreceber.HelpContextID, frmGeracaoTitulosReceber.name, "Geração de Titulos a Receber")
    Call mostrarForm(frmGeracaoTitulosReceber, mnuModGertitreceber.HelpContextID)
End Sub

Private Sub mnuModMovConferido_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModMovConferido.HelpContextID, frmConferido.name, "Movimento Conferido")
    Call mostrarForm(frmConferido, mnuModMovConferido.HelpContextID)
End Sub

Private Sub mnuModPagBaixas_Click()
    mstrPagRecBaixas = "P"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModPagBaixas.HelpContextID, frmBaixas.name, "Baixas Financeiras")
    Call mostrarForm(frmBaixas, mnuModPagBaixas.HelpContextID)
End Sub

Private Sub mnuModPagContas_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModPagContas.HelpContextID, frmContasFixas.name, "Cadastro de Contas Fixas/Previsão - Contas a Pagar")
    frmContasFixas.Caption = "Cadastro de Contas Fixas - Contas a Pagar"
    frmContasFixas.Configure 0 ' A pagar = 0
    Call mostrarForm(frmContasFixas, mnuModPagContas.HelpContextID)
End Sub

Private Sub mnuModPagDuplicatas_Click()
    Load frmDuplicatas
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModPagDuplicatas.HelpContextID, frmDuplicatas.name, "Cadastro de Duplicatas a Pagar")
    frmDuplicatas.Configure "Duplicatas", "P"
    frmDuplicatas.Icon = Me.Icon
    frmDuplicatas.HelpContextID = mnuModPagDuplicatas.HelpContextID
End Sub

Private Sub mnuModPagLancamentos_Click()
    frmLancamentos.HelpContextID = mnuModPagLancamentos.HelpContextID
    Load frmLancamentos
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModRecDuplicatas.HelpContextID, frmLancamentos.name, "Cadastro de Lançamentos a Pagar")
    frmLancamentos.Configure "Lançamentos", "P"
    frmLancamentos.Icon = Me.Icon
End Sub

Private Sub mnuModRecBaixas_Click()
    mstrPagRecBaixas = "R"
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModRecBaixas.HelpContextID, frmBaixas.name, "Baixas Financeiras")
    Call mostrarForm(frmBaixas, mnuModRecBaixas.HelpContextID)
End Sub

Private Sub mnuModRecBoleto_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModRecBoleto.HelpContextID, frmBoletos.name, "Processar Boleto Bancário")
    Call mostrarForm(frmBoletos, mnuModRecBoleto.HelpContextID)
    frmBoletos.ZOrder
End Sub
'Private Sub mnuModRecDuplicatas_Click()
'    frmDuplicatas.HelpContextID = mnuModRecDuplicatas.HelpContextID
'    'Load frmDuplicatas
'    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModRecDuplicatas.HelpContextID, frmDuplicatas.name, "Cadastro de Duplicatas a Receber")
'    Call mostrarForm(frmDuplicatas, mnuModRecDuplicatas.HelpContextID)
'    frmDuplicatas.Configure "Duplicatas", "R"
'    frmDuplicatas.Icon = Me.Icon
'End Sub




'Private Sub mnuModRecLancamentos_Click()
'    frmLancamentos.HelpContextID = mnuModRecLancamentos.HelpContextID
'    'Load frmLancamentos
'    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCadInTaxasBancaria.HelpContextID, frmLancamentos.name, "Cadastro de Lançamentos a Receber")
'    Call mostrarForm(frmLancamentos, mnuModRecLancamentos.HelpContextID)
'    frmLancamentos.Configure "Lançamentos", "R"
'    frmLancamentos.Icon = Me.Icon
'    frmLancamentos.cboDuplicatas(3).Text = "Fatura"
'End Sub

Private Sub mnuModuloProcessosCNABDadosFavorecido_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuModuloProcessosCNABDadosFavorecido.HelpContextID, frmDadosFavorecido.name, "Dados Favorecidos")
    Call mostrarForm(frmDadosFavorecido, mnuModuloProcessosCNAB.HelpContextID, False)
End Sub

Private Sub mnuOpContabeis_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuOpContabeis.HelpContextID, frmOperacoesContabeis.name, "Operações Contábeis")
    Call mostrarForm(frmOperacoesContabeis, mnuOpContabeis.HelpContextID)
End Sub

Private Sub mnuOpFinanceira_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuOpFinanceira.HelpContextID, frmOperacaoFinanceira.name, "Operação Financeira")
    Call mostrarForm(frmOperacaoFinanceira, mnuOpFinanceira.HelpContextID)
End Sub

Private Sub mnuProcessosCNABPagamentoRemessa_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuProcessosCNABPagamentoRemessa.HelpContextID, frmPagamentoDigitalFornecedores.name, "Remessa Pagamento")
    Call mostrarForm(frmPagamentoDigitalFornecedores, mnuProcessosCNABPagamentoRemessa.HelpContextID, False)
End Sub

Private Sub mnuRelAplFinanceiras_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelAplFinanceiras.HelpContextID, frptAplicacoes.name, "Aplicações Financeiras")
    Call mostrarForm(frptAplicacoes, mnuRelAplFinanceiras.HelpContextID)
End Sub

Private Sub mnuRelBolPreImpresso_Click()
    MsgBox "Esta versão do sistema não possui este relatório implementado"
End Sub

Private Sub mnuRelBordero_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelBordero.HelpContextID, frmBordero.name, "Borderôs")
    Call mostrarForm(frmBordero, mnuRelBordero.HelpContextID)
End Sub

Private Sub mnuRelCheques_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelCheques.HelpContextID, frptCheque.name, "Relatório de Cheques")
    Call mostrarForm(frptCheque, mnuRelCheques.HelpContextID)
End Sub

Private Sub mnuRelConFinanceiro_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelConFinanceiro.HelpContextID, frptCtrlFinanc.name, "Relatório de Controle Financeiro")
    Call mostrarForm(frptCtrlFinanc, mnuRelConFinanceiro.HelpContextID)
End Sub

Private Sub mnuRelDupLancamento_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelDupLancamento.HelpContextID, frptContasDupls1.name, "Relatório de Duplicatas e Lançamentos")
    frptContasDupls1.HelpContextID = mnuRelDupLancamento.HelpContextID
    Call mostrarForm(frptContasDupls1, mnuRelDupLancamento.HelpContextID)
End Sub

Private Sub mnuRelEmpresas_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloCompras, mnuRelEmpresas.HelpContextID, frptEmpresas.name, "Relatório de Empresas")
    Call mostrarForm(frptEmpresas, mnuRelEmpresas.HelpContextID)
    frptEmpresas.ZOrder
End Sub

Private Sub mnuRelEtoquetas_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelEtoquetas.HelpContextID, frmEtiqueta.name, frmEtiqueta.Caption)
    Call mostrarForm(frmEtiqueta, mnuRelEtoquetas.HelpContextID)
End Sub

Private Sub mnuRelExtBancario_Click()
    Load frptFluxo
    frptFluxo.tipo = 1                '1 = Extrato Bancário
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelExtBancario.HelpContextID, frptFluxo.name, "Extrato Bancário")
    Call mostrarForm(frptFluxo, mnuRelExtBancario.HelpContextID)
End Sub

Private Sub mnuRelFluCaiConGrupo_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelFluCaiConGrupo.HelpContextID, frptFluxoConta.name, "Fluxo de Caixa por Conta e Grupo")
    Call mostrarForm(frptFluxoConta, mnuRelFluCaiConGrupo.HelpContextID)
End Sub

Private Sub mnuCamposEspeciais_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuCamposEspeciais.HelpContextID, frmCamposEspeciais.name, "Campos Especiais")
    Call mostrarForm(frmCamposEspeciais, mnuCamposEspeciais.HelpContextID)
End Sub


Private Sub mnuRelFluCaiGeral_Click()
    Load frptFluxoCaixa
    frptFluxoCaixa.tipo = 0                '0 = Fluxo de Caixa
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelFluCaiGeral.HelpContextID, frptFluxoCaixa.name, "Fluxo de Caixa")
    Call mostrarForm(frptFluxoCaixa, mnuRelFluCaiGeral.HelpContextID)
End Sub

Private Sub mnuRelMovCaixa_Click()
    Load frptFluxo
    frptFluxo.tipo = 2                '2 = Movimento de Caixa
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelMovCaixa.HelpContextID, frptFluxo.name, "Movimento de Caixa")
    Call mostrarForm(frptFluxo, mnuRelMovCaixa.HelpContextID)
End Sub

Private Sub mnuRelRazAuxiliar_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelRazAuxiliar.HelpContextID, frptRazao.name, "Razão Auxiliar")
    Call mostrarForm(frptRazao, mnuRelRazAuxiliar.HelpContextID)
End Sub

Private Sub mnuRelRecibo_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelRecibo.HelpContextID, frptRecibos.name, "Recibos")
    Call mostrarForm(frptRecibos, mnuRelRecibo.HelpContextID)
End Sub

Private Sub mnuRelRegDuplicata_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelRegDuplicata.HelpContextID, frptRegistrodeDuplicatas.name, "Registro de Duplicatas")
    Call mostrarForm(frptRegistrodeDuplicatas, mnuRelRegDuplicata.HelpContextID)
    frptRegistrodeDuplicatas.ZOrder
End Sub

Private Sub mnuRelTabelas_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelTabelas.HelpContextID, frmTabelas.name, "Tabelas")
    Call mostrarForm(frmTabelas, mnuRelTabelas.HelpContextID)
End Sub

'Projeto: #218 - História: # - Problema# - João Henrique(05/10/2012)
'Private Sub mnuRelTitRecAtrAnalitico_Click()
'    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelTitRecAtrAnalitico.HelpContextID, frptDuplLancAtrasoNovo.name, "Relatório de Títulos a Receber em Atraso - Analitico")
'    Call mostrarForm(frptDuplLancAtrasoNovo, mnuRelTitRecAtrAnalitico.HelpContextID)
'End Sub

Private Sub mnuRelTitRecAtrSintetico_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelTitRecAtrSintetico.HelpContextID, frptDuplLancAtraso.name, "Relatório de Títulos a Receber em Atraso - Sintético")
    Call mostrarForm(frptDuplLancAtraso, mnuRelTitRecAtrSintetico.HelpContextID)
End Sub

Private Sub mnuRelTranBancaria_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuRelTranBancaria.HelpContextID, frptTransfBanco.name, "Transferências Bancárias")
    Call mostrarForm(frptTransfBanco, mnuRelTranBancaria.HelpContextID)
    frptTransfBanco.ZOrder
End Sub

Private Sub mnuReprocessaSaldoBanc_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuSistema.HelpContextID, frmReprocessaSaldo.name, "Reprocessamento de Saldos Bancários")
    Call mostrarForm(frmReprocessaSaldo, mnuReprocessaSaldoBanc.HelpContextID, False)
End Sub

Private Sub mnuRetornoPagamento_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuSistema.HelpContextID, frmRetornoArquivoPagamento.name, "Retorno Pagamento")
    Call mostrarForm(frmRetornoArquivoPagamento, mnuRetornoPagamento.HelpContextID, True)
End Sub

Private Sub mnuSistema_Click()
    blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, mnuSistema.HelpContextID, frmKinSys.name, "Configurações do Sistema")
    Call mostrarForm(frmKinSys, mnuSistema.HelpContextID, True)
End Sub

Private Sub mnuSuporteOnLine_Click()
    Call executaTalky
End Sub

Private Sub mnuUtiCalculadora_Click()
    WinCalc
End Sub

Private Sub mnuUtiConstrutor_Click()
    ShowQMaker
End Sub

Private Sub mnuUtiImpressora_Click()
    ShowPrinterSetup
End Sub

Private Sub mnuUtiRelatorio_Click()
    ShowFRE
End Sub

Private Sub Picture1_Click()
    Call executaTalky
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    ToolbarClick Button.Key, ActiveForm
End Sub

Private Sub tmrAtivacao_Timer()
    tmrAtivacao.Enabled = False
    
    If mintContador >= 30 Then
        mintContador = 0
        Call ModGeral.verificaReadOnly(ModFuncoes.IdentificacaoModulo)
    Else
        mintContador = mintContador + 1
    End If
    
    tmrAtivacao.Enabled = True
End Sub

'Protocolo Nr 96268  - Carlos Felippe Vernizze - 24/09/2010
Private Sub tmrMenu_Timer()
    If Forms.Count = 2 And Not EBSSBCenter.Expanded Then
        frmSideBar.Hide
        EBSSBCenter.Expanded = True
        fMain.stbMain.Panels(1).Text = "FOX - Sistema de Gestão - Financeiro"
        frmSideBar.Resize
        frmSideBar.Show
    End If
End Sub

'Comandos e controles para o menu lateral.
Private Sub eST_BeforeExpand(Index As Integer)
    'AJUSTAR ITEM.
        'Passa a propriedade maxima para abertura.
        eST(0).MaxHeight = EBSSBCenter.Height - 2550
        eST(1).MaxHeight = EBSSBCenter.Height - 2550
        eST(2).MaxHeight = EBSSBCenter.Height - 2550
        eST(3).MaxHeight = EBSSBCenter.Height - 2550
        eST(4).MaxHeight = EBSSBCenter.Height - 2550
        eST(5).MaxHeight = EBSSBCenter.Height - 2550
    
        Select Case Index
            Case 0
                tvwCadastro.Move 150, tvwCadastro.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 1
                tvwModulos.Move 150, tvwModulos.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 2
                tvwConsultas.Move 150, tvwConsultas.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 3
                tvwRelatorios.Move 150, tvwRelatorios.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 4
                tvwUtilitarios.Move 150, tvwUtilitarios.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 5
                tvwAjuda.Move 150, tvwAjuda.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
        End Select
End Sub



Private Sub eST_Resize(Index As Integer)
    
    Dim i As Integer
    
    For i = 1 To eST.UBound
        eST(i).Move eST(i).Left, (eST(i - 1).Top + eST(i - 1).Height) - 15
    Next
    
    If eST(Index).Expanded = True Then
        Select Case Index
            Case 0
                tvwCadastro.Move 150, tvwCadastro.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 1
                tvwModulos.Move 150, tvwModulos.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 2
                tvwConsultas.Move 150, tvwConsultas.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 3
                tvwRelatorios.Move 150, tvwRelatorios.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 4
                tvwUtilitarios.Move 150, tvwUtilitarios.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
            Case 5
                tvwAjuda.Move 150, tvwAjuda.Top, eST(Index).Width - 300, eST(Index).MaxHeight - 100
        End Select
    End If
End Sub

Private Sub eST_CompleteExpand(Index As Integer)
    Dim i As Integer
    
    For i = 0 To eST.UBound
        If Index <> i Then
            If eST(i).AutoContract = True Then
                eST(i).Expanded = False
            End If
        End If
    Next
End Sub

Private Sub ResizeEBSSBCenter(ByVal NewWidth As Integer)
    Dim i As Integer
    For i = 0 To eST.UBound
        eST(i).Left = 60
        eST(i).Width = NewWidth - 150
    Next
End Sub

Private Sub EBSSBCenter_BeforeResize(ByVal NewWidth As Integer)
    ResizeEBSSBCenter NewWidth
End Sub

Private Sub CarregaMenu(objTview As Control)
    Dim strSql          As String
    Dim rsMenu          As IDBReader
    Dim rsMenu1         As IDBReader
    Dim lngNode         As Long
    Dim cmd             As IDBSelectCommand
    
    objTview.Nodes.Clear
    lngNode = 0
     
    Set cmd = Aplicacao.CreateSelectCommand
    With cmd
        .Table.TableName = "[FGSMenuModulo]"
        Call cmd.Filter.Append("[rotina_vinculada] = @pRotina_vinculada")
        Call cmd.Parameters.add(cmd.CreateParameter("@pRotina_vinculada", objTview.HelpContextID, dbFieldTypeLong))
        
        'Call cmd.Filter.Append("[id_form] NOT IN ( 2441,2315 )")
        
        'Projeto: 218 - História: 268 - Tarefa: 402 - Fernando Paludo 05/10/2012
        #If FOXSQL Then
            'Retirado os relatórios que utilizam o ReportXWizard que serão substituidos por reltórios ERC
            Call .Filter.Append("[id_form] NOT IN (2424, 2473, 2088, 2690, 2107)")
        #End If
        
        'Projeto: 268 - História: 154 - Tarefa: 987 - Fernando Paludo 05/10/2012
        'Retirado o menu do relatório de Duplicatas e Lançamentos temporáriamente até correção do relatório
        Call .Filter.Append("[id_form] NOT IN (2085, 2087)")
        
        If ArquivoExiste(CaminhoPasta(pastaConfiguracoes) & "ebsserver.ini") Then
            Call .Filter.Append("[id_form] NOT IN (2474)")
        End If
        
        .OrderByClause = "ordem_menu"
    End With
    Set rsMenu = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
       
    While Not rsMenu.EOF
        Call INSERIR_P(rsMenu.GetInteger("id_form"), RTrim(rsMenu.GetString("Descricao")), objTview)
        lngNode = lngNode + 1
        If rsMenu.GetBoolean("possui_filho") Then
            Call INSERIR_F(rsMenu.GetInteger("id_form"), objTview)
        End If
        rsMenu.MoveNext
    Wend
End Sub

Private Function INSERIR_F(IDMENUPAI As String, objTview As Control)
    Dim rsMenu1             As IDBReader
    Dim strSql              As String
    Dim intPermissao        As Integer
    Dim cmd                 As IDBSelectCommand
    
    Set cmd = Aplicacao.CreateSelectCommand
    With cmd
        .Table.TableName = "[FGSMenuModulo]"
        Call cmd.Filter.Append("[rotina_vinculada] = @pRotina_vinculada")
        Call cmd.Parameters.add(cmd.CreateParameter("@pRotina_vinculada", IDMENUPAI, dbFieldTypeLong))

        'Projeto: 218 - História: 268 - Tarefa: 402 - Fernando Paludo 05/10/2012
        #If FOXSQL Then
            'Retirado os relatórios que utilizam o ReportXWizard que serão substituidos por reltórios ERC
            Call .Filter.Append("[id_form] NOT IN (2424, 2473, 2088, 2690, 2107)")
        #End If

        'Projeto: 268 - História: 154 - Tarefa: 987 - Fernando Paludo 05/10/2012
        'Retirado o menu do relatório de Duplicatas e Lançamentos temporáriamente até correção do relatório
        Call .Filter.Append("[id_form] NOT IN (2085, 2087)")
        
        If ArquivoExiste(CaminhoPasta(pastaConfiguracoes) & "ebsserver.ini") Then
            Call .Filter.Append("[id_form] NOT IN (2474)")
        End If

        .OrderByClause = "ordem_menu"
    End With
    Set rsMenu1 = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    
    While Not rsMenu1.EOF
        intPermissao = fPermissaoAcesso(rsMenu1.GetInteger("id_form"))
        objTview.Nodes.add "X" & IDMENUPAI, tvwChild, "X" & rsMenu1.GetInteger("id_form"), RTrim(rsMenu1.GetString("descricao")), intPermissao
        objTview.Nodes("X" & IDMENUPAI).Expanded = False
        objTview.Nodes("X" & IDMENUPAI).ExpandedImage = intPermissao
        If rsMenu1.GetBoolean("possui_filho") Then
            Call INSERIR_F(rsMenu1.GetInteger("id_form"), objTview)
        End If
        rsMenu1.MoveNext
    Wend
End Function

Private Function INSERIR_P(Indice As String, Nome As String, objTview As Control)
    Dim intPermissao        As Integer

    intPermissao = fPermissaoAcesso(CInt(Indice))
    objTview.Nodes.add , , "X" & Indice, Nome, intPermissao
    objTview.Nodes("X" & Indice).ExpandedImage = intPermissao
End Function

Private Function fPermissaoAcesso(intIdForm As Integer) As Integer
    'Imagem 2 possui permissao, imagem 1 não possui permissão.
    If verificaPermissaoUsuario(retornaGrupoUsuario(usuarioParametro(Command())), ID_MODULO_FINANCEIRO, intIdForm) Then
        fPermissaoAcesso = 2
    Else
        fPermissaoAcesso = 1
    End If
End Function

Private Function fCarregaFomularioMenu(strIdForm As String)
    Dim blnContrair         As Boolean
    Dim frmDuplicata        As Form
    Dim frmLancamento       As Form

    blnContrair = True
    Set frmFormSelecionado = Nothing
    Select Case strIdForm
        Case "2032"
            mnuCadGerEmpresa_Click
        Case "2033"
            mnuCadGerEmpPotencial_Click
        Case "2034"
            mnuCadGerRamo_Click
        Case "2862"
            mnuCadMunicipios_Click
        Case "2035"
            mnuCadGerEstado_Click
        Case "2036"
            mnuCadGerRegiao_Click
        Case "2037"
            mnuCadGerPais_Click
        Case "2688"
            mnuCadProcedencias_Click
        Case "2038"
            mnuCadCentroCusto_Click
        Case "2040"
            mnuCadGrupoConta_Click
        Case "2039"
            mnuCadContas_Click
        Case "2042"
            mnuCadBancoCaixa_Click
        Case "2926"
            mnuCarteira_Click
        Case "2853"
            mnuCadCamara_Click
        Case "2801"
            mnuCadForPagamento_Click
        Case "2797"
            mnuOpFinanceira_Click
        Case "2044"
            mnuCadInTaxasBancaria_Click
        Case "2045"
            mnuCadInDespFinanceira_Click
        Case "2046"
            mnuCadIndMoeda_Click
        Case "2047"
            mnuCadIndCotacao_Click
        Case "2049"
            mnuCadGenProjeto_Click
        Case "2050"
            mnuCadGenTipoGlobal_Click
        Case "2051"
            mnuCadGenFeriado_Click
        Case "2052"
            mnuCadGenObservacao_Click
        Case "2053"
            mnuCadConfGeral_Click
        Case "2423"
            mnuCadSair_Click
        Case "2057"
            If Not verificaFormCriado(frmLancamento, strIdForm) Then
                Set frmLancamento = New frmLancamentoDuplicata
            End If
            frmLancamento.LancDup = Lancamento
            frmLancamento.PagRec = Recebimento
            blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, 2057, frmLancamento.name, "Lançamentos a Receber ou Recebidos")
            Call mostrarForm(frmLancamento, 2057)
        Case "2058"
           If Not verificaFormCriado(frmDuplicata, strIdForm) Then
                Set frmDuplicata = New frmLancamentoDuplicata
            End If
            frmDuplicata.LancDup = Duplicata
            frmDuplicata.PagRec = Recebimento
            blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, 2058, frmDuplicata.name, "Duplicatas a Receber ou Recebidas")
            Call mostrarForm(frmDuplicata, 2058)
        Case "2059"
            mnuModRecBoleto_Click
        Case "2725"
            mnuCOntasFixasReceber_Click
        Case "2812"
            mnuModGertitreceber_Click
        Case "2076"
            mnuModRecBaixas_Click
        Case "2854"
            mnuAlteracaoBancoTitulosReceber_Click
        Case "2061"
            If Not verificaFormCriado(frmLancamento, strIdForm) Then
                Set frmLancamento = New frmLancamentoDuplicata
            End If
            frmLancamento.LancDup = Lancamento
            frmLancamento.PagRec = Pagamento
            blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, 2061, frmLancamento.name, "Lançamentos a Pagar ou Pagos")
            Call mostrarForm(frmLancamento, 2061)
        Case "2062"
           If Not verificaFormCriado(frmDuplicata, strIdForm) Then
                Set frmDuplicata = New frmLancamentoDuplicata
            End If
            frmDuplicata.LancDup = Duplicata
            frmDuplicata.PagRec = Pagamento
            blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, 2062, frmDuplicata.name, "Duplicatas a Pagar ou Pagas")
            Call mostrarForm(frmDuplicata, 2062)
        Case "2063"
            mnuModPagContas_Click
        Case "2810"
            mnuModGertitpagar_Click
        Case "2855"
            mnuAlteracaoBancoTitulosPagar_Click
        Case "2065"
            mnuModBanAplicacoes_Click
        Case "2066"
            mnuModBanMovEntrada_Click
        Case "2067"
            mnuModBanMovSaida_Click
        Case "2068"
            mnuModBanTranBancaria_Click
        Case "2069"
            mnuModBanSaldoBanco_Click
        Case "2806"
            mnuModBanConcBancaria_Click
        Case "2070"
            mnuModBanCadCheque_Click
        Case "2072"
            mnuModBanEditaCheque_Click
        Case "2075"
            mnuModCaiLibera_Click
        Case "2077"
            mnuModCaiDesconto_Click
        Case "2798"
            mnuContaCorrente_Click
        Case "2078"
            mnuModGerConDupLancamento_Click
        Case "2079"
            mnuModMovConferido_Click
        Case "2845"
            mnuModuloProcessosCNABDadosFavorecido_Click
        Case "2846"
            mnuProcessosCNABPagamentoRemessa_Click
        Case "2864"
            mnuRetornoPagamento_Click
        Case "2693"
            MnuKINComunicacaoCadRemessas_Click
        Case "2694"
            mnuKINEnvioCobrancas_Click
        Case "2695"
            mnuKINEnvioPagamento_Click
        Case "2696"
            mnuKINComunicacaoRemessaImpExp_Click
        Case "2691"
            mnuComunicacoesRetornoBancario_Click
        Case "2922"
            mnuEmissaoBoleto_Click
        Case "2923"
            mnuEmissaoRemessa_Click
        Case "2924"
            mnuConfirmacaoRetorno_Click
        Case "2644"
            mnuIntApropImpostos_Click
        Case "2645"
            mnuOpContabeis_Click
        Case "2646"
            mnuMatrixContab_Click
        Case "2651"
            mnuIntContabil_Click
        Case "2718"
            mnuIntFiscal_Click
        Case "2746"
            mnuAtualizacaoOperacoes_Click
        Case "2081"
            mnuConLanDuplicata_Click
        Case "2082"
            mnuConSaldos_Click
        Case "2083"
            mnuConTitAtraso_Click
        Case "2085"
            mnuRelDupLancamento_Click
        Case "2086"
            mnuRelTitRecAtrSintetico_Click
        Case "2088"
            mnuRelBolPreImpresso_Click
        Case "2089"
            mnuRelBordero_Click
        Case "2090"
            mnuRelRegDuplicata_Click
        Case "2091"
            mnuRelTabelas_Click
        Case "2092"
            mnuRelMovCaixa_Click
        Case "2093"
            mnuRelFluCaiGeral_Click
        Case "2094"
            mnuRelFluCaiConGrupo_Click
        Case "2095"
            mnuRelConFinanceiro_Click
        Case "2096"
            mnuRelRazAuxiliar_Click
        Case "2097"
            mnuRelCheques_Click
        Case "2098"
            mnuRelRecibo_Click
        Case "2099"
            mnuRelExtBancario_Click
        Case "2100"
            mnuRelAplFinanceiras_Click
        Case "2101"
            mnuRelTranBancaria_Click
        Case "2689"
            mnuRelEmpresas_Click
        Case "2690"
            mnuRelEtoquetas_Click
        Case "2103"
            mnuUtiRelatorio_Click
            blnContrair = False
        Case "2104"
            mnuUtiImpressora_Click
        Case "2105"
            mnuUtiCalculadora_Click
            blnContrair = False
        Case "2106"
            mnuUtiConstrutor_Click
        Case "2474"
            mnuImportaExportaTab_Click
        Case "2876"
            mnuImportaDuplicata_Click
        Case "2479"
            mnuConfigRelERC_Click
        Case "2480"
            mnuConfigServidorEmail_Click
        Case "2481"
            mnuSistema_Click
        Case "2426"
           mnuAjuAjuda_Click
        Case "2707"
            mnuModPagBaixas_Click
        Case "2429"
           mnuSuporteOnLine_Click
           blnContrair = False
        Case "2832"
           mnuAjudaSuporteRemoto_Click
           blnContrair = False
        Case "2431"
           mnuAjuSobre_Click
        'Projeto: #218 - História: # - Problema# - João Henrique(05/10/2012)
        Case "2087"
           'mnuRelTitRecAtrAnalitico_Click
        Case "2743"
            mnuGeracaoIntegracaoBalanSet_Click
        Case "3007"
            blnEscreve = escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, frmReajusteDupLan.HelpContextID, frmReajusteDupLan.name, "Reajuste de Duplicatas a Receber")
            Call mostrarForm(frmReajusteDupLan, frmReajusteDupLan.HelpContextID)
        'Projeto: 61827 - Desenv.: 62690 - Ueder Budni (12/01/2015)
        Case "3014"
            mnuImpDigExtratoBacario_Click
        'Projeto: 61827 - Desenv.: 62687 - Ueder Budni (13/01/2015)
        Case "3015"
            mnuCadHistBancario_Click
        Case "3016"
            mnuModBanConcBancariaAut_Click
            blnContrair = False
        'Vinicius Elyseu(02/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
        Case "3018"
            mnuReprocessaSaldoBanc_Click
            'blnContrair = True
        Case "3018"
            mnuReprocessaSaldoBanc_Click
            
        'Davi Brito(09/05/2016) - #120997
        Case "3021"
            mnuCamposEspeciais_Click
            
        Case Else
            blnContrair = False
    End Select
    If blnContrair Then
        frmSideBar.Hide
        EBSSBCenter.Expanded = False
        frmSideBar.Resize
        frmSideBar.Show
        If (Not IsNothing(frmFormSelecionado)) Then
            frmFormSelecionado.ZOrder
        End If
    End If
End Function

Private Sub tmrPerfil_Timer()
    'pt.Perfil Fernando Paludo(01/08/2011)
    Dim objBizPerfil        As bizPerfil
    Dim col                 As Collection
    
    mIntTimer = mIntTimer + 1
    If mIntTimer = 30 Then
        mIntTimer = 0
        
        'pt.Perfil Fernando Paludo(01/08/2011)
        Set objBizPerfil = New bizPerfil
        Set col = New Collection
        'Verifica pedido/notas emitidos
        
        objBizPerfil.validarPerfil col
        Call EnviaMensagem_Perfil(col)
        
        Set objBizPerfil = Nothing
        Set col = Nothing
    End If
End Sub

Private Sub tvwUtilitarios_DblClick()
    'Imagem 2 possui permissao, imagem 1 não possui permissão.
    If tvwUtilitarios.SelectedItem.Image = 2 Then
        fCarregaFomularioMenu (Replace(tvwUtilitarios.SelectedItem.Key, "X", ""))
    End If
End Sub

Private Sub tvwCadastro_DblClick()
    'Imagem 2 possui permissao, imagem 1 não possui permissão.
    If tvwCadastro.SelectedItem.Image = 2 Then
        fCarregaFomularioMenu (Replace(tvwCadastro.SelectedItem.Key, "X", ""))
    End If
    'Debug.Print tvwCadastro.SelectedItem.Key & "   " & tvwCadastro.SelectedItem.Text
End Sub

Private Sub tvwConsultas_DblClick()
    'Imagem 2 possui permissao, imagem 1 não possui permissão.
    If tvwConsultas.SelectedItem.Image = 2 Then
        fCarregaFomularioMenu (Replace(tvwConsultas.SelectedItem.Key, "X", ""))
    End If
End Sub

Private Sub tvwAjuda_DblClick()
    'Imagem 2 possui permissao, imagem 1 não possui permissão.
    If tvwAjuda.SelectedItem.Image = 2 Then
        fCarregaFomularioMenu (Replace(tvwAjuda.SelectedItem.Key, "X", ""))
    End If
End Sub

Private Sub tvwRelatorios_DblClick()
    'Imagem 2 possui permissao, imagem 1 não possui permissão.
    If tvwRelatorios.SelectedItem.Image = 2 Then
        fCarregaFomularioMenu (Replace(tvwRelatorios.SelectedItem.Key, "X", ""))
    End If
End Sub

Private Sub tvwModulos_DblClick()
    'Imagem 2 possui permissao, imagem 1 não possui permissão.
    If tvwModulos.SelectedItem.Image = 2 Then
        fCarregaFomularioMenu (Replace(tvwModulos.SelectedItem.Key, "X", ""))
'        Debug.Print tvwModulos.SelectedItem.Key & "   " & tvwModulos.SelectedItem.Text
    End If
End Sub

Public Sub CarregarFormulario(ByVal strIdForm As String)
    Call fCarregaFomularioMenu(strIdForm)
End Sub

'pt.101230 - Fernando Paludo(01/12/2010)
Private Function verificaFormCriado(ByRef CurrrentForm As Form, ByVal strChave As String) As Boolean
    Dim blnEstaCriado As Boolean
    Dim frmTemp As Form
    
    blnEstaCriado = False
    For Each frmTemp In Forms
        If frmTemp.HelpContextID = strChave Then
            Set CurrrentForm = frmTemp
            blnEstaCriado = True: Exit For
        End If
    Next
    
    verificaFormCriado = blnEstaCriado

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
'Vinicius Elyseu(23/05/2016) - Projeto: #120807
Public Sub ChamarMovManual()
'Não chama Mov Manual porque não tem chamada para cadastro de produtos.
End Sub
