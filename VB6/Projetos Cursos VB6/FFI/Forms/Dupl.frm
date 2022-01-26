VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmDuplicatas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicatas"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   780
   ClientWidth     =   11280
   Icon            =   "Dupl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11280
   Tag             =   "Duplicatas"
   Begin TabDlg.SSTab SSTab1 
      Height          =   6330
      Left            =   0
      TabIndex        =   49
      Top             =   30
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   11165
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados Gerais"
      TabPicture(0)   =   "Dupl.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDuplicatas(30)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraRateio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraDuplicatas(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDuplicatas(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDuplicatas(26)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FraDuplicatas(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FraDuplicatas(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FraDuplicatas(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Adicionais"
      TabPicture(1)   =   "Dupl.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraDuplicatas(7)"
      Tab(1).Control(1)=   "FraDuplicatas(5)"
      Tab(1).Control(2)=   "FraDuplicatas(4)"
      Tab(1).Control(3)=   "FraDuplicatas(6)"
      Tab(1).Control(4)=   "FraDuplicatas(10)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Outros"
      TabPicture(2)   =   "Dupl.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblDuplicatas(44)"
      Tab(2).Control(1)=   "lblDuplicatas(45)"
      Tab(2).Control(2)=   "lblDuplicatas(46)"
      Tab(2).Control(3)=   "lblDuplDesc(16)"
      Tab(2).Control(4)=   "FraDuplicatas(8)"
      Tab(2).Control(5)=   "txtDuplicatas(39)"
      Tab(2).Control(6)=   "txtDuplicatas(42)"
      Tab(2).Control(7)=   "txtDuplicatas(43)"
      Tab(2).Control(8)=   "Frame"
      Tab(2).ControlCount=   9
      Begin VB.Frame Frame 
         Caption         =   "Endere�o de Cobran�a"
         Height          =   1635
         Left            =   -74940
         TabIndex        =   145
         Top             =   3600
         Width           =   11145
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "cd_cobranca"
            Height          =   315
            Index           =   44
            Left            =   1230
            MaxLength       =   2
            TabIndex        =   40
            Tag             =   "Duplicatas"
            Top             =   360
            Width           =   705
         End
         Begin Fox.EBSText etxCobrancaEndereco 
            Height          =   330
            Left            =   1230
            TabIndex        =   146
            Top             =   720
            Width           =   9840
            _extentx        =   9022
            _extenty        =   582
            tipo            =   4
            tipotexto       =   0
            maxlength       =   70
            locked          =   -1  'True
            font            =   "Dupl.frx":0060
            exibedescricao  =   0   'False
         End
         Begin Fox.EBSText etxCobrancaBairro 
            Height          =   330
            Left            =   735
            TabIndex        =   147
            Top             =   1080
            Width           =   4320
            _extentx        =   314801
            _extenty        =   582
            tipo            =   4
            tipotexto       =   0
            maxlength       =   20
            caption         =   "Bairro"
            locked          =   -1  'True
            font            =   "Dupl.frx":008C
            exibedescricao  =   0   'False
         End
         Begin Fox.EBSText etxCobrancaCep 
            Height          =   330
            Left            =   2715
            TabIndex        =   151
            Top             =   360
            Width           =   1620
            _extentx        =   68898
            _extenty        =   582
            tipo            =   4
            tipotexto       =   0
            maxlength       =   9
            caption         =   "CEP"
            locked          =   -1  'True
            font            =   "Dupl.frx":00B8
            exibedescricao  =   0   'False
         End
         Begin Fox.EBSText etxCobrancaCidade 
            Height          =   330
            Left            =   5250
            TabIndex        =   152
            Top             =   360
            Width           =   2835
            _extentx        =   265
            _extenty        =   582
            tipo            =   4
            tipotexto       =   0
            maxlength       =   30
            locked          =   -1  'True
            font            =   "Dupl.frx":00E4
            exibedescricao  =   0   'False
         End
         Begin Fox.EBSText etxCobrancaEstado 
            Height          =   330
            Left            =   8490
            TabIndex        =   153
            Top             =   360
            Width           =   525
            _extentx        =   265
            _extenty        =   582
            tipo            =   4
            tipotexto       =   0
            maxlength       =   30
            locked          =   -1  'True
            font            =   "Dupl.frx":0110
            exibedescricao  =   0   'False
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   8220
            TabIndex        =   154
            Top             =   420
            Width           =   210
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Endere�o"
            Height          =   195
            Left            =   450
            TabIndex        =   150
            Top             =   750
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Left            =   4695
            TabIndex        =   149
            Top             =   405
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "C�digo"
            Height          =   195
            Left            =   645
            TabIndex        =   148
            Top             =   420
            Width           =   495
         End
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "id_carteira"
         Height          =   315
         Index           =   43
         Left            =   -73710
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Duplicatas"
         Top             =   3060
         Width           =   1245
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "NOSNUM"
         Height          =   315
         Index           =   42
         Left            =   -73710
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   38
         Tag             =   "Duplicatas"
         Top             =   2670
         Width           =   5085
      End
      Begin VB.Frame Frame1 
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
         ForeColor       =   &H00000000&
         Height          =   960
         Left            =   7605
         TabIndex        =   138
         Top             =   4815
         Width           =   3615
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "cd_operacao_baixa"
            Height          =   330
            Index           =   41
            Left            =   1080
            TabIndex        =   22
            Tag             =   "Duplicatas"
            Top             =   315
            Width           =   735
         End
         Begin VB.Label lblDuplDesc 
            AutoSize        =   -1  'True
            Caption         =   "lblDuplDesc(15)"
            Height          =   195
            Index           =   15
            Left            =   1890
            TabIndex        =   140
            Top             =   405
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Op. Cont�bil:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   135
            TabIndex        =   139
            Top             =   405
            Width           =   915
         End
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "LINDIG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   39
         Left            =   -73710
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   37
         Tag             =   "Duplicatas"
         Top             =   2190
         Width           =   9855
      End
      Begin VB.Frame FraDuplicatas 
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
         ForeColor       =   &H00000000&
         Height          =   1785
         Index           =   10
         Left            =   -74940
         TabIndex        =   109
         Top             =   2520
         Width           =   6255
         Begin VB.TextBox txtPercMora 
            DataField       =   "PerMrd"
            Height          =   315
            Left            =   1950
            MaxLength       =   9
            TabIndex        =   26
            Tag             =   "Duplicatas"
            Top             =   870
            Width           =   1365
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "VlrMrd"
            Height          =   315
            Index           =   38
            Left            =   4590
            MaxLength       =   9
            TabIndex        =   29
            Tag             =   "Duplicatas"
            Top             =   870
            Width           =   1365
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "PerMul"
            Height          =   315
            Index           =   37
            Left            =   1950
            MaxLength       =   9
            TabIndex        =   25
            Tag             =   "Duplicatas"
            Top             =   510
            Width           =   1365
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "VlrMul"
            Height          =   315
            Index           =   36
            Left            =   4590
            MaxLength       =   9
            TabIndex        =   28
            Tag             =   "Duplicatas"
            Top             =   510
            Width           =   1365
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "VlrDsP"
            Height          =   315
            Index           =   31
            Left            =   1950
            MaxLength       =   9
            TabIndex        =   27
            Tag             =   "Duplicatas"
            Top             =   1260
            Width           =   1365
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Perc. Mora:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   43
            Left            =   1050
            TabIndex        =   114
            Top             =   900
            Width           =   825
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vlr. Mora Di�ria:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   42
            Left            =   3375
            TabIndex        =   113
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Perc. Multa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   41
            Left            =   1020
            TabIndex        =   112
            Top             =   540
            Width           =   855
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vlr. Multa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   40
            Left            =   3795
            TabIndex        =   111
            Top             =   540
            Width           =   705
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vlr. Desc. Pontualidade:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   120
            TabIndex        =   110
            Top             =   1290
            Width           =   1710
         End
      End
      Begin VB.Frame FraDuplicatas 
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
         ForeColor       =   &H00000000&
         Height          =   1725
         Index           =   8
         Left            =   -74940
         TabIndex        =   104
         Top             =   390
         Width           =   11175
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CheEmi"
            Height          =   315
            Index           =   35
            Left            =   1230
            MaxLength       =   60
            TabIndex        =   36
            Tag             =   "Duplicatas"
            Top             =   1290
            Width           =   5055
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CheCco"
            Height          =   315
            Index           =   34
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   35
            Tag             =   "Duplicatas"
            Top             =   960
            Width           =   2085
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CheAge"
            Height          =   315
            Index           =   33
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   34
            Tag             =   "Duplicatas"
            Top             =   570
            Width           =   1245
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CheBan"
            Height          =   315
            Index           =   32
            Left            =   1230
            MaxLength       =   9
            TabIndex        =   33
            Tag             =   "Duplicatas"
            Top             =   210
            Width           =   1245
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Correntista:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   39
            Left            =   345
            TabIndex        =   108
            Top             =   1350
            Width           =   795
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta Corrente:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   90
            TabIndex        =   107
            Top             =   990
            Width           =   1110
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ag�ncia:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   540
            TabIndex        =   106
            Top             =   630
            Width           =   630
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   660
            TabIndex        =   105
            Top             =   240
            Width           =   510
         End
      End
      Begin VB.Frame FraDuplicatas 
         Caption         =   "Linha Digit�vel"
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
         Height          =   735
         Index           =   6
         Left            =   -74940
         TabIndex        =   103
         Top             =   390
         Width           =   6255
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "LINDIG"
            Height          =   315
            Index           =   22
            Left            =   120
            TabIndex        =   23
            Tag             =   "Duplicatas"
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.Frame FraDuplicatas 
         Caption         =   "Observa��es"
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
         Height          =   1335
         Index           =   4
         Left            =   -74940
         TabIndex        =   102
         Top             =   1140
         Width           =   6255
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Obs"
            Height          =   975
            Index           =   23
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Tag             =   "Duplicatas"
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.Frame FraDuplicatas 
         Caption         =   "&Informa��es do Cheque"
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
         Height          =   4965
         Index           =   5
         Left            =   -68580
         TabIndex        =   95
         Top             =   390
         Width           =   4785
         Begin VB.TextBox txtCheque 
            DataField       =   "Nominal"
            Height          =   315
            Index           =   0
            Left            =   960
            MaxLength       =   60
            TabIndex        =   30
            Tag             =   "Cheques"
            Top             =   360
            Width           =   3435
         End
         Begin VB.TextBox txtCheque 
            DataField       =   "Hist�rico"
            Height          =   1275
            Index           =   1
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Tag             =   "Cheques"
            Top             =   720
            Width           =   3735
         End
         Begin VB.CommandButton cmdNominalRazaoSocial 
            Caption         =   "..."
            Height          =   300
            Left            =   4440
            TabIndex        =   96
            ToolTipText     =   "Cheque Nominal a Empresa do Lan�amento/Duplicata"
            Top             =   360
            Width           =   255
         End
         Begin ComctlLib.ListView lvwLancamentos 
            Height          =   2205
            Left            =   120
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   2640
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
         Begin VB.Label lblDuplicatas 
            AutoSize        =   -1  'True
            Caption         =   "Lan�amentos"
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
            Index           =   26
            Left            =   240
            TabIndex        =   97
            Top             =   2400
            Width           =   1140
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nominal:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   300
            TabIndex        =   101
            Top             =   390
            Width           =   615
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Hist�rico:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   100
            Top             =   720
            Width           =   660
         End
         Begin VB.Line hline 
            BorderColor     =   &H80000014&
            Index           =   2
            X1              =   120
            X2              =   4680
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line hline 
            BorderColor     =   &H80000010&
            Index           =   3
            X1              =   120
            X2              =   4680
            Y1              =   2505
            Y2              =   2505
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   510
            TabIndex        =   99
            Top             =   2070
            Width           =   405
         End
         Begin VB.Label lblDuplDesc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDuplDesc(9)"
            Height          =   315
            Index           =   9
            Left            =   960
            TabIndex        =   98
            Top             =   2040
            Width           =   3735
         End
         Begin ComctlLib.ImageList imgDupl 
            Left            =   120
            Top             =   1080
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
      End
      Begin VB.Frame FraDuplicatas 
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
         ForeColor       =   &H00000000&
         Height          =   1005
         Index           =   7
         Left            =   -74940
         TabIndex        =   91
         Top             =   4350
         Width           =   6255
         Begin VB.Label lblDadosAdcionais 
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   94
            Tag             =   "Descri��o"
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   5955
         End
         Begin VB.Label lblDadosAdcionais 
            Caption         =   "Estado:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   93
            Tag             =   "Descri��o"
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   5955
         End
         Begin VB.Label lblDadosAdcionais 
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   92
            Tag             =   "Descri��o"
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   5955
         End
      End
      Begin VB.Frame FraDuplicatas 
         Caption         =   "Da&tas"
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
         Height          =   1815
         Index           =   3
         Left            =   7620
         TabIndex        =   67
         Top             =   3000
         Width           =   3615
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Emiss�o"
            Height          =   315
            Index           =   6
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "Duplicatas"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Vencimento"
            Height          =   315
            Index           =   7
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   19
            Tag             =   "Duplicatas"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Pagamento"
            Height          =   315
            Index           =   8
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   20
            Tag             =   "Duplicatas"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Libera��o"
            Height          =   315
            Index           =   9
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   21
            Tag             =   "Duplicatas"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Emiss�o:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   420
            TabIndex        =   75
            Top             =   270
            Width           =   630
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   180
            TabIndex        =   74
            Top             =   630
            Width           =   885
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Pagamento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   210
            TabIndex        =   73
            Top             =   990
            Width           =   855
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Libera��o:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   300
            TabIndex        =   72
            Top             =   1350
            Width           =   750
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(5)"
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   71
            Tag             =   "Descri��o"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(6)"
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   70
            Tag             =   "Descri��o"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(7)"
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   69
            Tag             =   "Descri��o"
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(8)"
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   68
            Tag             =   "Descri��o"
            Top             =   1320
            Width           =   1095
         End
      End
      Begin VB.Frame FraDuplicatas 
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
         ForeColor       =   &H00000000&
         Height          =   2535
         Index           =   2
         Left            =   7620
         TabIndex        =   59
         Top             =   390
         Width           =   3615
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Moeda"
            Height          =   315
            Index           =   17
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Duplicatas"
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor Original"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   1200
            MaxLength       =   18
            TabIndex        =   15
            Tag             =   "Duplicatas"
            Top             =   660
            Width           =   2295
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Acr�scimo"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   1200
            MaxLength       =   18
            TabIndex        =   16
            Tag             =   "Duplicatas"
            Top             =   1050
            Width           =   2295
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Abatimento"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Index           =   12
            Left            =   1200
            MaxLength       =   18
            TabIndex        =   17
            Tag             =   "Duplicatas"
            Top             =   1380
            Width           =   2295
         End
         Begin VB.Label lblDuplicatas 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   240
            TabIndex        =   60
            Top             =   1830
            Width           =   1530
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Moeda:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   570
            TabIndex        =   66
            Top             =   330
            Width           =   540
         End
         Begin VB.Line hline 
            BorderColor     =   &H80000014&
            Index           =   1
            X1              =   120
            X2              =   3480
            Y1              =   1950
            Y2              =   1950
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Original:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   150
            TabIndex        =   65
            Top             =   690
            Width           =   975
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Acr�scimo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   64
            Top             =   1050
            Width           =   780
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Abatimento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   300
            TabIndex        =   63
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   720
            TabIndex        =   62
            Top             =   2160
            Width           =   405
         End
         Begin VB.Label lblDuplDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   1200
            TabIndex        =   61
            Tag             =   "Descri��o"
            Top             =   2100
            Width           =   2295
         End
         Begin VB.Line hline 
            BorderColor     =   &H80000010&
            BorderWidth     =   2
            Index           =   0
            X1              =   120
            X2              =   3480
            Y1              =   1950
            Y2              =   1950
         End
      End
      Begin VB.Frame FraDuplicatas 
         Caption         =   "Duplicatas"
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
         Height          =   1815
         Index           =   0
         Left            =   60
         TabIndex        =   50
         Top             =   390
         Width           =   7455
         Begin VB.TextBox txtDuplicatas 
            Alignment       =   1  'Right Justify
            DataField       =   "SeqNossoNumero"
            Enabled         =   0   'False
            Height          =   315
            Index           =   28
            Left            =   6060
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "Duplicatas"
            Top             =   630
            Width           =   1275
         End
         Begin VB.CommandButton btnEfetiva 
            Caption         =   "Efetivar Lancto."
            Height          =   405
            Left            =   5820
            TabIndex        =   51
            Top             =   180
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "PagRec"
            Height          =   315
            Index           =   0
            Left            =   4230
            MaxLength       =   1
            TabIndex        =   41
            Tag             =   "Duplicatas"
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Empresa"
            Height          =   315
            Index           =   2
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   4
            Tag             =   "Duplicatas"
            Top             =   960
            Width           =   1575
         End
         Begin VB.ComboBox cboDuplicatas 
            DataField       =   "Tipo"
            Height          =   315
            Index           =   3
            ItemData        =   "Dupl.frx":013C
            Left            =   1080
            List            =   "Dupl.frx":013E
            TabIndex        =   1
            Tag             =   "Duplicatas"
            Text            =   "cboDuplicatas"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Parcela"
            Height          =   315
            Index           =   4
            Left            =   4230
            MaxLength       =   3
            TabIndex        =   2
            Tag             =   "Duplicatas"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Descri��o"
            Height          =   315
            Index           =   5
            Left            =   1080
            MaxLength       =   80
            TabIndex        =   5
            Tag             =   "Duplicatas"
            Top             =   1320
            Width           =   6255
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Nota"
            Height          =   315
            Index           =   1
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   0
            Tag             =   "Duplicatas"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nr Sequencial:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   4980
            TabIndex        =   58
            Top             =   660
            Width           =   1050
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   57
            Top             =   990
            Width           =   660
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   660
            TabIndex        =   56
            Top             =   630
            Width           =   360
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Parcela:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   3570
            TabIndex        =   55
            Top             =   660
            Width           =   585
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Descri��o:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   54
            Top             =   1350
            Width           =   765
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nota:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   630
            TabIndex        =   53
            Top             =   270
            Width           =   390
         End
         Begin VB.Label lblDuplDesc 
            AutoSize        =   -1  'True
            Caption         =   "lblDuplDesc(0)"
            Height          =   195
            Index           =   0
            Left            =   2760
            TabIndex        =   52
            Tag             =   "Descri��o"
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1035
         End
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "Usu�rio"
         Enabled         =   0   'False
         Height          =   315
         Index           =   26
         Left            =   8550
         MaxLength       =   18
         TabIndex        =   47
         Tag             =   "Duplicatas"
         Top             =   5835
         Width           =   1215
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "Altera��o"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   9990
         MaxLength       =   18
         TabIndex        =   48
         Tag             =   "Duplicatas"
         Top             =   5835
         Width           =   1215
      End
      Begin VB.Frame FraDuplicatas 
         Caption         =   "Co&ntroles"
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
         Height          =   3945
         Index           =   1
         Left            =   60
         TabIndex        =   76
         Top             =   2220
         Width           =   7455
         Begin VB.ComboBox cboDuplicatas 
            DataField       =   "Situa��o"
            Height          =   315
            Index           =   20
            ItemData        =   "Dupl.frx":0140
            Left            =   1170
            List            =   "Dupl.frx":0142
            Style           =   2  'Dropdown List
            TabIndex        =   157
            Tag             =   "Duplicatas"
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CheckBox chkRateio 
            Caption         =   "Identifica se a duplicata faz parte do rateio\Campo Invis�vel"
            DataField       =   "proveniente_rateio"
            Height          =   195
            Left            =   120
            TabIndex        =   141
            Tag             =   "Duplicatas"
            Top             =   3660
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "cd_operacao_contabil"
            Height          =   330
            Index           =   40
            Left            =   1170
            TabIndex        =   10
            Tag             =   "Duplicatas"
            Top             =   1665
            Width           =   1230
         End
         Begin VB.TextBox txtChequeCheque 
            DataField       =   "Cheque"
            Height          =   285
            Left            =   5940
            TabIndex        =   45
            Tag             =   "Cheques"
            Top             =   3150
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtBancoCheque 
            DataField       =   "Banco"
            Height          =   285
            Left            =   6705
            TabIndex        =   46
            Tag             =   "Cheques"
            Top             =   3150
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Carteira"
            Height          =   315
            Index           =   27
            Left            =   3960
            MaxLength       =   3
            TabIndex        =   44
            Tag             =   "Duplicatas"
            Top             =   3135
            Width           =   1095
         End
         Begin VB.CheckBox chkConciliado 
            Caption         =   "Conciliado"
            DataField       =   "Conciliado"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1170
            TabIndex        =   13
            Tag             =   "Duplicatas"
            Top             =   3165
            Width           =   1035
         End
         Begin VB.CommandButton cmdProximoCheque 
            Caption         =   "..."
            Height          =   315
            Left            =   2730
            TabIndex        =   42
            ToolTipText     =   "Trazer Pr�ximo N�mero do Cheque"
            Top             =   2415
            Width           =   255
         End
         Begin VB.CommandButton cmdAbreRateio 
            Caption         =   "&Rateio..."
            Height          =   255
            Left            =   4050
            TabIndex        =   43
            Top             =   2775
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Banco"
            Height          =   315
            Index           =   13
            Left            =   1170
            MaxLength       =   9
            TabIndex        =   7
            Tag             =   "Duplicatas"
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Conta"
            Height          =   315
            Index           =   14
            Left            =   1170
            MaxLength       =   9
            TabIndex        =   8
            Tag             =   "Duplicatas"
            Top             =   930
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Centro"
            Height          =   315
            Index           =   15
            Left            =   1170
            MaxLength       =   9
            TabIndex        =   9
            Tag             =   "Duplicatas"
            Top             =   1290
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Cheque"
            Height          =   315
            Index           =   16
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   11
            Tag             =   "Duplicatas"
            Top             =   2415
            Width           =   1575
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Controle"
            Height          =   315
            Index           =   19
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   12
            Tag             =   "Duplicatas"
            Top             =   2775
            Width           =   2775
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CODFPG"
            Height          =   315
            Index           =   18
            Left            =   1170
            MaxLength       =   9
            TabIndex        =   6
            Tag             =   "Duplicatas"
            Top             =   210
            Width           =   1215
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(14)"
            Height          =   195
            Index           =   14
            Left            =   2475
            TabIndex        =   137
            Tag             =   "Descri��o"
            Top             =   1710
            Width           =   4875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Op. Cont�bil:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   136
            Top             =   1710
            Width           =   915
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Forma Pagto.:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   150
            TabIndex        =   90
            Top             =   270
            Width           =   990
         End
         Begin VB.Label lblDuplicatas 
            AutoSize        =   -1  'True
            Caption         =   "Carteira:"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   31
            Left            =   3210
            TabIndex        =   88
            Top             =   2790
            Width           =   585
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(12)"
            Height          =   195
            Index           =   12
            Left            =   3030
            TabIndex        =   87
            Tag             =   "Descri��o"
            Top             =   2085
            Width           =   4275
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   630
            TabIndex        =   86
            Top             =   600
            Width           =   510
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   660
            TabIndex        =   85
            Top             =   960
            Width           =   465
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C. Custo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   84
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cheque:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   83
            Top             =   2445
            Width           =   900
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Controle:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   510
            TabIndex        =   82
            Top             =   2805
            Width           =   630
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Situa��o:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   450
            TabIndex        =   81
            Top             =   2085
            Width           =   675
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(1)"
            Height          =   195
            Index           =   1
            Left            =   2490
            TabIndex        =   80
            Tag             =   "Descri��o"
            Top             =   570
            UseMnemonic     =   0   'False
            Width           =   4875
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(2)"
            Height          =   195
            Index           =   2
            Left            =   2490
            TabIndex        =   79
            Tag             =   "Descri��o"
            Top             =   930
            Width           =   4875
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(3)"
            Height          =   195
            Index           =   3
            Left            =   2490
            TabIndex        =   78
            Tag             =   "Descri��o"
            Top             =   1290
            Width           =   4875
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(13)"
            Height          =   255
            Index           =   13
            Left            =   2490
            TabIndex        =   77
            Top             =   270
            Width           =   4875
         End
      End
      Begin VB.Frame FraRateio 
         Caption         =   "Rateio"
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
         Height          =   3945
         Left            =   60
         TabIndex        =   115
         Top             =   2220
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox txtDuplicatas 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   45
            Left            =   5505
            MaxLength       =   18
            TabIndex        =   155
            Top             =   1380
            Width           =   1845
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   25
            Left            =   1200
            TabIndex        =   117
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdicionar 
            Caption         =   "&Adicionar..."
            Height          =   375
            Left            =   120
            TabIndex        =   122
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   3270
            TabIndex        =   125
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdRateio 
            Caption         =   "&Ratear..."
            Height          =   375
            Left            =   2220
            TabIndex        =   124
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Excluir..."
            Height          =   375
            Left            =   1170
            TabIndex        =   123
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   21
            Left            =   5520
            MaxLength       =   18
            TabIndex        =   119
            Top             =   210
            Width           =   1845
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Centro"
            Height          =   315
            Index           =   20
            Left            =   1200
            MaxLength       =   9
            TabIndex        =   116
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   24
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   118
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   29
            Left            =   5520
            MaxLength       =   18
            TabIndex        =   120
            Top             =   600
            Width           =   1845
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   30
            Left            =   5520
            MaxLength       =   18
            TabIndex        =   121
            Top             =   990
            Width           =   1845
         End
         Begin ComctlLib.ListView lvwRateio 
            Height          =   1215
            Left            =   120
            TabIndex        =   129
            Top             =   1830
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2143
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            Icons           =   "imgRateio"
            SmallIcons      =   "imgRateio"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   "custo"
               Object.Tag             =   ""
               Text            =   "C.Custo"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   "DescCusto"
               Object.Tag             =   ""
               Text            =   "Descri��o"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   "porcentagem"
               Object.Tag             =   ""
               Text            =   "Porcent."
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   "valor"
               Object.Tag             =   ""
               Text            =   "Valor"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   "acrescimo"
               Object.Tag             =   ""
               Text            =   "Acr�scimo"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   "abatimento"
               Object.Tag             =   ""
               Text            =   "Abatimento"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   "conta"
               Object.Tag             =   ""
               Text            =   "Conta"
               Object.Width           =   1323
            EndProperty
         End
         Begin VB.Label lblsaldoRateio 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Restante:"
            Height          =   195
            Left            =   4320
            TabIndex        =   156
            Top             =   1410
            Width           =   1140
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(11)"
            Height          =   195
            Index           =   11
            Left            =   2520
            TabIndex        =   134
            Tag             =   "Descri��o"
            Top             =   630
            Width           =   1995
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta Financ.:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   133
            Top             =   630
            Width           =   1035
         End
         Begin ComctlLib.ImageList imgRateio 
            Left            =   0
            Top             =   3120
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
                  Picture         =   "Dupl.frx":0144
                  Key             =   "Checked"
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Dupl.frx":01A2
                  Key             =   "Unchecked"
               EndProperty
            EndProperty
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   5070
            TabIndex        =   132
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Porcentagem:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   150
            TabIndex        =   131
            Top             =   990
            Width           =   990
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C. Custo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   510
            TabIndex        =   130
            Top             =   270
            Width           =   645
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(10)"
            Height          =   195
            Index           =   10
            Left            =   2520
            TabIndex        =   128
            Tag             =   "Descri��o"
            Top             =   270
            Width           =   1995
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Acr�scimo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   4695
            TabIndex        =   127
            Top             =   630
            Width           =   780
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Abatimento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   4650
            TabIndex        =   126
            Top             =   1020
            Width           =   840
         End
      End
      Begin VB.Label lblDuplDesc 
         Caption         =   "..."
         Height          =   195
         Index           =   16
         Left            =   -72390
         TabIndex        =   144
         Tag             =   "Descri��o"
         Top             =   3120
         Width           =   4875
      End
      Begin VB.Label lblDuplicatas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carteira:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   46
         Left            =   -74355
         TabIndex        =   143
         Top             =   3120
         Width           =   585
      End
      Begin VB.Label lblDuplicatas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nosso Numero:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   45
         Left            =   -74865
         TabIndex        =   142
         Top             =   2730
         Width           =   1095
      End
      Begin VB.Label lblDuplicatas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linha Digit�vel:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   44
         Left            =   -74880
         TabIndex        =   135
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblDuplicatas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Usu�rio:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   7920
         TabIndex        =   89
         Top             =   5865
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IDI_LANCTO = 500          '�cone (no arquivo de recursos) para Lan�amentos
Private Const IDI_DUPL = 501            '�cone (no arquivo de recursos) para Duplicatas
Private Const TAG_CHEQUE$ = "Cheques"   'Tag dos campos de informa��es do cheque
Private Const DL_MARCADO = 1        '�ndice do �cone de lan�amento marcado no ImageList
Private Const DL_DESMARCADO = 2     '�ndice do �cone de lan�amento desmarcado no ImageList
Private Const IDB_TRANSF = 509          'Imagem para o ListView para Cheques em Transfer�ncias
Private Const IDB_DUPLS = 510           '�dem para Duplicatas
Private Const IDB_LANCTOS = 511         '�dem para Lan�amentos
'Identificadores dos �tens do menu Ferramentas
Private Const IDM_DUPLWNDCALC& = 32010
Private Const IDM_DUPLCHQINFO& = 32011
Private Const IDM_DUPLOBSFIN& = 32012
Private Const IDM_DUPLADDHIST& = 32013
'Valores poss�veis da vari�vel mintBaixa
Private Const CDT_NORMAL = 0            'Abertura normal do cadastro
Private Const CDT_BXTOTAL = 1           'Baixa total de uma Duplicata ou Lan�amento
Private Const CDT_BXPARCIAL = 2         'Baixa parcial de uma Duplicata ou Lan�amento

Private mstrTabela      As String               'Nome da Tabela que est� aberta
Private mstrPagRec      As String               'Tipo a pagar ou a receber
Private mintBaixa       As Integer              'Tipo da Baixa se for uma baixa
Private mlngItem        As Long                 '�tem selecionado da lista de rateio
Private bHabRateio      As Boolean              'Indica se o bot�o de habilitar rateio vai ficar vis�vel ou n�o
Private mstrPesquisa    As String               'Instru��o SQL utilizada na fun��o de Pesquisa
Private mstrDuplicatas  As String               'Instru��o SQL utilizada na abertura da Tabela
Private mrstDuplicatas  As Object               'Abre a tabela
Private mlngDuplicatas  As Long                 'Controla as altera��es do usu�rio
Private mrstCheques     As Object               'Abre a tabela de cheques
Private mlngCheques     As Long                 'Controla as altera��es do usu�rio em cheques
Private SeqLancamentos  As Boolean              'Configura��o para sugerir seq��ncia de Lan�amentos
'pt. 83525 - Dulcino J�nior (27/09/2007)
Private mblnAlteraValor As Boolean              'Flag utilizado para n�o considerar a sugest�o da Opera��o Cont�bil como altera��o.
'Dulcino J�nior (28/10/2007)
Private lngOperacao     As Long
Private mstrOrigem      As String
Private mstrDelete      As String
Private mstrRateio      As String
Private mlngCodigo      As Long
Private mlngPARCELA     As Long
Private mstrEmpresa     As String
Private mstrTipoRegistro As String

'FUNCTION..: LibProc
'Objetivo..: Fun��o de retorno de chamada da Lib.
'Argumentos: [sFuncao]: Fun��o que deve ser executada.
'            [lFuncao]: Par�metro adicional, varia conforme a fun��o.
'Retorna...: True se executar a fun��o com sucesso, False, se n�o.
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
    Dim sTmp            As String
    Dim nBanco          As Long              '// C�digo do Banco
    Dim nCheque         As Long              '// N�mero do Cheque
    Dim strProcura      As String
    Dim rstBanco        As Object
    Dim intParcOrigem   As Integer
    Dim objBizPerfil    As bizPerfil
    Dim col             As Collection
    Dim blnPerfil       As Boolean
    
    
    If cmdAbreRateio.Visible Then
        cmdAbreRateio.Enabled = True
    End If
    
    btnEfetiva.Enabled = False
    btnEfetiva.Visible = False
    
    Select Case (sFuncao)
        Case WL_NOVO
            If (mintBaixa = CDT_NORMAL) Then    '// Somente se for abertura normal
                If (LimpaControles(mrstDuplicatas, Me, Tag, mlngDuplicatas) = WL_OK) Then
                    'pt. 86140 - Moacir Pfau(10/04/2008)
                    Call CarregaPadrao
                    Call ChequeInfo(WL_NOVO)
                    txtDuplicatas(10).Enabled = True
                    NovoRegistro True
                    LibProc = True
                    lblDadosAdcionais(0).Caption = NUL
                    lblDadosAdcionais(1).Caption = NUL
                    lblDadosAdcionais(2).Caption = NUL
                    lngOperacao = 0
                End If
            Else
                MsgFunc ResolveResString(240, resUM, "acrescentar um novo registro")
                LibProc = True
            End If
            'pt. 86140 - Moacir Pfau(17/04/2008)
            cmdCancelar_Click
        Case WL_SETFOCUS
            SSTab1.Tab = 0
            txtDuplicatas(1).SetFocus
        Case WL_DELETAR
            sTmp = GetValue(mrstDuplicatas, "Libera��o", NUL)
            If sTmp <> "" Then
                If Not ValidaDatasDiasUteis(0, 0, CDate(sTmp), True) Then
                    Exit Function
                End If
            End If
        
            'Grava o c�digo do Banco e Cheque atual para a rotina ChequeInfo
            nBanco = GetValue(mrstDuplicatas, "Banco", ZERO)
            nCheque = GetValue(mrstDuplicatas, "Cheque", ZERO)
        
            'pt. 81604 - Dulcino J�nior
            If Not PermiteExclusao(intParcOrigem) Then
                Exit Function
            End If
            
            'pt. 85684 - Moacir Pfau(01/07/2008)
            If Not fValidaExclusao Then
                MsgBox "Duplicata gerado pela tela de t�tulo, n�o pode ser excluida pela tela de duplicatas.", vbInformation
                Exit Function
            End If
            
            'pt. 82831 - Ivo Sousa (23/02/2009)
            BeginTrans
            If intParcOrigem > 0 Then
                Call ExecuteSQL("UPDATE Duplicatas SET Abatimento=0 WHERE PagRec='" & mstrPagRec & "' AND Nota=" & txtDuplicatas(1).Text & " AND Parcela=" & intParcOrigem & " AND Empresa='" & txtDuplicatas(2).Text & "' AND Tipo='" & cboDuplicatas(3).Text & "'")
                intParcOrigem = 0
            End If
            If DeletaRegistro(mrstDuplicatas, Me, Tag, mlngDuplicatas) = WL_OK Then
                CommitTrans
                Call ChequeInfo(WL_DELETAR, nBanco, nCheque): LibProc = True
                If mstrOrigem <> "" Then
                    Call ExecuteSQL(mstrOrigem)
                    mstrOrigem = ""
                    Call ExecuteSQL(mstrDelete)
                    mstrDelete = ""
                    If mstrRateio <> "" Then
                        If Recordcount(mstrRateio) = 0 Then
                            Call ExecuteSQL("UPDATE Duplicatas SET proveniente_rateio=False WHERE PagRec='" & mstrPagRec & "' AND Nota=" & mlngCodigo & _
                                        " AND Empresa='" & mstrEmpresa & "' AND Tipo='" & mstrTipoRegistro & "'")
                        End If
                    End If
                End If
                If mintBaixa <> CDT_NORMAL Then
                    mrstDuplicatas.Requery
                    If EstaVazio(mrstDuplicatas) Then
                        LibProc WL_SAIR
                        Exit Function
                    End If
                End If
            End If
        
        Case WL_LOCALIZAR
            If (mintBaixa = CDT_NORMAL) Then    '// A janela Localizar s� � habilitada em modo normal
                If (localizar(mrstDuplicatas, Me, mstrTabela, Tag, mlngDuplicatas) = WL_OK) Then
                    Call ChequeInfo(WL_LOCALIZAR): LibProc = True
                    txtDuplicatas(10).Enabled = permiteAlterarValor
                End If
            Else
                MsgFunc ResolveResString(240, resUM, "localizar")
            End If
        
        Case WL_PESQUISAR
            Dim strSql      As String
            strSql = NUL
            If (mintBaixa = CDT_NORMAL) Then
                If Configuracao("Visualizar somente Movimenta��es n�o Conferidas", False) Then
                    strSql = " and Libera��o >= " & InverteData(DateAdd("M", 1, MaxValue("M�s Conferido", "Mov Conferido", "KIF = True")), True)
                End If
                If (PRegistro(mrstDuplicatas, Me, Caption, mstrDuplicatas & strSql, _
                                mstrPesquisa & strSql, Tag, mlngDuplicatas, PB_REGISTRO) = WL_OK) Then
                    Call ChequeInfo(WL_PESQUISAR): LibProc = True
                    txtDuplicatas(10).Enabled = permiteAlterarValor
                End If
            Else
                If (FindFirst(mrstDuplicatas, Me, Tag, mstrPesquisa, mlngDuplicatas) = WL_OK) Then
                    Call ChequeInfo(WL_PESQUISAR): LibProc = True
                    txtDuplicatas(10).Enabled = permiteAlterarValor
                End If
            End If
        
        Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
            DoEvents
            If (WL_OK = MoveRecordset(mrstDuplicatas, Me, Tag, mlngDuplicatas, lFuncao)) Then
                Call ChequeInfo(sFuncao): LibProc = True
                txtDuplicatas(10).Enabled = permiteAlterarValor
            End If
        
        Case WL_NAVEGAR
            If (Browse(mrstDuplicatas, Me, Tag, mlngDuplicatas, mstrDuplicatas) = WL_OK) Then
                Call ChequeInfo(WL_NAVEGAR): LibProc = True
                txtDuplicatas(10).Enabled = permiteAlterarValor
            End If
        
        Case WL_SALVAR
            If DuplVerifique Then
                If mstrPagRec = "R" Then
                    If EAdicao(mlngDuplicatas) Then
                        Set objBizPerfil = New bizPerfil
                        Set col = New Collection
                        blnPerfil = objBizPerfil.validarTelaPerfil(col, NumeroDeTitulosNoContasAReceber)
                        Call EnviaMensagem_Perfil(col)
                        If Not blnPerfil Then
                            Exit Function
                        End If
                        Set objBizPerfil = Nothing
                    End If
                End If
                
                'No caso de estar configurado para utilizar Op. Cont�bil.
                If txtDuplicatas(41).Enabled Then
                    'pt. 81902 - Dulcino J�nior
                    If mstrTabela = "Lan�amentos" Then
                        If Not validaIntegracaoLancamentos Then
                            Exit Function
                        End If
                    Else
                        If Not validaIntegracaoDuplicatas Then
                            Exit Function
                        End If
                    End If
                End If
                Dim ObsFin As String
                ObsFin = GetFieldValue("[Obs Financeira]", "Empresas", "Apel = '" & txtDuplicatas(2).Text & "'")
                If Len(ObsFin) > 0 Then
                    MsgBox ObsFin, vbInformation, "Observa��es Financeiras da Empresa"
                End If
                nBanco = GetValue(mrstDuplicatas, "Banco", ZERO)
                nCheque = GetValue(mrstDuplicatas, "Cheque", ZERO)
                txtDuplicatas(26).Text = UserName
                txtDuplicatas(3).Text = Date
                Dim bGeraComissao As Boolean
                bGeraComissao = (IsNull(mrstDuplicatas("Pagamento")) And txtDuplicatas(8).Text <> "")
        
                'pt. 86132 - Ivo Sousa (26/03/2008)
                'Valida��o de Datas
                If ValidaDatas Then
                    If (SalvaRegistro(mrstDuplicatas, Me, Tag, mlngDuplicatas) = WL_OK) Then
                        ExibeSoma
                        'pt: 74271 - Dulcino J�nior
                        'Erro ao alterar uma duplicata que n�o tem cheque
                        If strToDbl(txtDuplicatas(16).Text) > 0 Then
                            txtBancoCheque.Text = txtDuplicatas(13).Text
                            txtChequeCheque.Text = txtDuplicatas(16).Text
                        End If
                        nBanco = GetValue(mrstDuplicatas, "Banco", ZERO)
                        nCheque = GetValue(mrstDuplicatas, "Cheque", ZERO)
                        Call ChequeInfo(WL_SALVAR, nBanco, nCheque): LibProc = True
                        'pt. 88289 - Dulcino J�nior (15/10/2008)
                        If chkRateio.value = vbChecked Then
                            If txtDuplicatas(8).Text <> "" Then
                                Call ExecuteSQL("UPDATE FFIRateioDuplicata SET dt_pagamento=" & InverteData(txtDuplicatas(8).Text, True) & _
                                                " WHERE pag_rec_destino='" & mstrPagRec & "' AND nr_nota_destino=" & txtDuplicatas(1).Text & _
                                                " AND cd_empresa_destino='" & txtDuplicatas(2).Text & "' AND tp_registro_destino='" & _
                                                cboDuplicatas(3).Text & "' AND nr_parcela_destino=" & txtDuplicatas(4).Text)
                            Else
                                Call ExecuteSQL("UPDATE FFIRateioDuplicata SET dt_pagamento=NULL" & _
                                                " WHERE pag_rec_destino='" & mstrPagRec & "' AND nr_nota_destino=" & txtDuplicatas(1).Text & _
                                                " AND cd_empresa_destino='" & txtDuplicatas(2).Text & "' AND tp_registro_destino='" & _
                                                cboDuplicatas(3).Text & "' AND nr_parcela_destino=" & txtDuplicatas(4).Text)
                            
                            End If
                        End If
                    End If
                End If
                'Gera��o da comiss�o
                If Configuracao("TipGcm") = "A" Then
                    If bGeraComissao Then
                        Dim oCom As New CCOMISSAO
                        Call oCom.GeraComissaoDuplicata(GBL_NFS, GetValue(mrstDuplicatas, "Nota"), GetValue(mrstDuplicatas, "Empresa"), GetValue(mrstDuplicatas, "Tipo"), GetValue(mrstDuplicatas, "Parcela"))
                        Set oCom = Nothing
                    End If
                End If
            End If
            
        Case WL_CANCELAR
            If (CancelaEdicao(mrstDuplicatas, Me, Tag, mlngDuplicatas) = WL_OK) Then
                Call ChequeInfo(WL_CANCELAR): LibProc = True
            End If
        
        Case WL_EXIBIR
            If Not EAddNew(mlngDuplicatas) Then
                txtDuplicatas(40).Text = GetValue(mrstDuplicatas, "cd_operacao_contabil", "0")
            Else
                txtDuplicatas(40).Text = lngOperacao
            End If
            If (mintBaixa = CDT_NORMAL) Then
                sTmp = mstrDuplicatas   '// Termina e completa a instru��o conforme a tabela
                If (mstrTabela = "Duplicatas") Then
                    Concat sTmp, " AND Nota = {Nota} AND Parcela = {Parcela} AND Tipo = '{Tipo}' AND Empresa = '{Empresa}';"
                Else
                    'pt. 83992 e 83998 - Dulcino J�nior (19/10/2007)
                    Concat sTmp, " AND C�digo = {C�digo} AND Parcela = {Parcela} AND Tipo = '{Tipo}';"
                End If
                    
                If (RetornaRegs(mrstDuplicatas, Me, Tag, sTmp, mlngDuplicatas) = WL_OK) Then
                    Call ChequeInfo(WL_EXIBIR): LibProc = True
                    txtDuplicatas(10).Enabled = permiteAlterarValor
                ElseIf (UltimoRetorno() = WL_ADDNEW) Then
                    Call NovoRegistro(False)
                    Call ChequeInfo(WL_NOVO)
                    LibProc = True
                    
                    lblDadosAdcionais(0).Caption = NUL
                    lblDadosAdcionais(1).Caption = NUL
                    lblDadosAdcionais(2).Caption = NUL
                    
                    If cmdAbreRateio.Visible Then
                        cmdAbreRateio.Enabled = True
                    End If
                End If
            End If
            'pt. 79561 - Moacir Pfau(04/04/2008)
            If EAdicao(mlngDuplicatas) Or EAddNew(mlngDuplicatas) Then
                'If (strToLng(txtDuplicatas(13).Text) = 0 And strToLng(txtDuplicatas(14).Text) = 0) And txtDuplicatas(2).Text <> "" Then
                If txtDuplicatas(2).Text <> "" Then
                    strProcura = "SELECT Banco, Conta FROM Empresas WHERE Apel = '" & txtDuplicatas(2).Text & "';"
                    If AbreRecordset(rstBanco, strProcura) Then
                        txtDuplicatas(13).Text = strToLng(GetValue(rstBanco, "Banco"))
                        txtDuplicatas(14).Text = strToLng(GetValue(rstBanco, "Conta"))
                    End If
                    FechaRecordset (rstBanco)
                End If
            End If
        Case WL_FILTRAR
            If (mintBaixa = CDT_NORMAL) Then    '// S� filtra se for abertura normal
                If (Filtrar(mrstDuplicatas, Me, Tag, mstrDuplicatas, mlngDuplicatas) = WL_OK) Then
                    Call ChequeInfo(WL_FILTRAR): LibProc = True
                    txtDuplicatas(10).Enabled = permiteAlterarValor
                End If
            Else
                MsgFunc ResolveResString(240, resUM, "filtrar")
            End If
        
        ' Registro Duplicado
        Case WL_DUPLICADO
            If (mintBaixa = CDT_NORMAL) Then    '// S� resolve se for abertura normal
                If mstrTabela = "Lan�amentos" Then
                    ResolveDuplicacao Me, txtDuplicatas(1), "Lan�amentos", "PagRec = " & Quote(mstrPagRec, "''")
                Else
                    If CompStr(mstrPagRec, "P") Then
                        ResolveDuplicacao Me, txtDuplicatas(1), "Duplicatas", "PagRec = 'P'"
                    Else
                        ResolveDuplicacao Me, txtDuplicatas(1), "Duplicatas", "PagRec = 'R'"
                    End If
                End If
            End If
        
        Case WL_SAIR
            Unload Me
            Exit Function
          
        Case "Empresas"
            If (KeybAcesso(LoadResString(2037))) Then
                frmEmpresas.Show
                CallChange frmEmpresas.hWnd, txtDuplicatas(2).hWnd
            End If
        
        Case "Bancos"
            If (KeybAcesso(LoadResString(2003))) Then
                frmBancos.Show
                CallChange frmBancos.hWnd, txtDuplicatas(13).hWnd
            End If
        
        Case "Contas"
            If (KeybAcesso(LoadResString(2007))) Then
                frmContas.Show
                CallChange frmContas.hWnd, txtDuplicatas(14).hWnd
            End If
        
        Case "Custos"
            If (KeybAcesso(LoadResString(2029))) Then
                frmCusto.Show
                CallChange frmCusto.hWnd, txtDuplicatas(15).hWnd
            End If
        
        Case "Moedas"
            If (KeybAcesso(LoadResString(2033))) Then
                fMoedas.Show
                CallChange fMoedas.hWnd, txtDuplicatas(17).hWnd
            End If
          
        Case "Configura��o"
            If KeybAcesso(LoadResString(2106)) Then
                FrmConfCad.Configura "Duplicatas"
                FrmConfCad.Show vbModal
            End If
          
        
        'Atualizar Valor
        Case IDM_DUPLWNDCALC
            Call CalcValor
        
        'Informa��es do Cheque
        Case IDM_DUPLCHQINFO
            Call ChequeInfo("updt")
        
        'Observa��es Financeiras
        Case IDM_DUPLOBSFIN
            If (IsValid(txtDuplicatas(2).Text)) Then
                Call fMemo("Observa��es Financeiras", "Empresas", "[Obs Financeira]", wsprintf("Apel = '%s'", txtDuplicatas(2).Text))
            End If
        
        'Hist�rico do Cheque
        Case IDM_DUPLADDHIST
            If (Len(txtDuplicatas(23).Text)) Then
                txtCheque(1).Text = wsprintf("%s\n%s", txtCheque(1).Text, txtDuplicatas(23).Text)
            End If
    End Select
    
    If (UltimoRetorno() = WL_OK And sFuncao <> WL_NOVO) Then
        If cmdAbreRateio.Visible Then
            cmdAbreRateio.Enabled = True
        End If
    Else
        If UltimoRetorno() <> WL_ADDNEW And UltimoRetorno() <> 0 Then
            If cmdAbreRateio.Visible Then
                cmdAbreRateio.Enabled = True
            End If
        Else
            If cmdAbreRateio.Visible Then
                cmdAbreRateio.Enabled = True
            End If
        End If
    End If
    
    If (Not IsValid(GetValue(mrstDuplicatas, "Pagamento", NUL))) And Not EAdicao(mlngDuplicatas) Then
        If GetValue(mrstDuplicatas, "Vencimento") >= Date Then
            lblDuplDesc(12).Caption = "A Vencer"
        Else
            lblDuplDesc(12).Caption = "Vencida"
        End If
    ElseIf Not EAdicao(mlngDuplicatas) Then
        lblDuplDesc(12).Caption = "Baixada"
    Else
        lblDuplDesc(12).Caption = NUL
    End If
    
    If IsValid(txtDuplicatas(2).Text) Then
        MsgBar IIf(IsValid(txtDuplicatas(1).Text), IIf(mstrTabela = "Duplicatas", " Nota: " & txtDuplicatas(1).Text, " Lan�amento: " & txtDuplicatas(1).Text), " ") _
             & IIf(IsValid(cboDuplicatas(3).Text), " - Tipo de Registro:" & cboDuplicatas(3).Text, " ") _
             & IIf(IsValid(txtDuplicatas(2).Text), " - Empresa: " & txtDuplicatas(2).Text, " ") _
             & IIf(mstrTabela = "Duplicatas", IIf(IsValid(txtDuplicatas(4).Text), " - Parcela: " & txtDuplicatas(4).Text, " "), " ")
    End If
  
    If LibProc = True Then
        lblDadosAdcionais(0).Caption = "Cidade: " & GetFieldValue("Cidade", "Empresas", "Apel = " & Quote(GetValue(mrstDuplicatas, "Empresa", NUL), "'"), , NUL)
        lblDadosAdcionais(1).Caption = "Estado: " & GetFieldValue("Estado", "Empresas", "Apel = " & Quote(GetValue(mrstDuplicatas, "Empresa", NUL), "'"), , NUL)
        If GetValue(mrstDuplicatas, "PagRec") = "R" Then
            Dim Vendedor    As Long
            Vendedor = GetFieldValue("Vendedor01", Quote(GBL_ITENS & GBL_NFS, "[]"), _
                        "N�mero = " & GetValue(mrstDuplicatas, "Nota", ZERO) & _
                        " AND [Tipo de Registro] = " & Quote(GetValue(mrstDuplicatas, "Tipo", NUL), "'"), , ZERO)
            If Vendedor > 0 Then
                lblDadosAdcionais(2).Caption = "Vendedor: " & Format(Vendedor, "000000") & " - " & GetFieldValue("Nome", "Vendedores", "C�digo = " & Vendedor, , NUL)
            End If
        Else
          lblDadosAdcionais(2).Caption = NUL
        End If
    End If
  
    'valida previs�o em lan�amentos.
    If mstrTabela = "Lan�amentos" Then
        If (Not mrstDuplicatas.EOF) Then
            If (mrstDuplicatas.AbsolutePosition > 0) Then
                If (mrstDuplicatas("Previsao") = True) Then
                    btnEfetiva.Enabled = True
                    btnEfetiva.Visible = True
                End If
            End If
        End If
    End If
End Function

Private Sub btnEfetiva_Click()
    Dim strUpdate As String

    If MsgBox("Confirma efetiva��o do lan�amento de previs�o ?", vbYesNo, "Confirma��o") = vbYes Then
        btnEfetiva.Enabled = False
        btnEfetiva.Visible = False
        'Pt. 95368 - Moacir Pfau(23/11/2009)
        'mrstDuplicatas.Edit
        mrstDuplicatas("Previsao").value = False
        mrstDuplicatas.update
    End If
End Sub

Private Sub cboDuplicatas_Click(Index As Integer)
    If (Index = 3) Then
        If (Not ControlaChave(CBCLICK, ZERO, cboDuplicatas(3), mlngDuplicatas)) Then
            cboDuplicatas(3).Text = GetValue(mrstDuplicatas, "Tipo")
        End If
    ElseIf (Index > 3) Then
        AlteraValor mlngDuplicatas
    End If
End Sub

Private Sub cboDuplicatas_DropDown(Index As Integer)
    If ((mstrTabela = "Duplicatas") And (Index = 3)) Then   'Campo Tipo
        ControlaChave CBDROPDOWN, 0, cboDuplicatas(3), mlngDuplicatas
    End If
End Sub

Private Sub cboDuplicatas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (mstrTabela = "Duplicatas") Then   'Tipo n�o entra na chave quando for lan�amentos
        If (Index = 3) Then
            ControlaChave KeyCode, Shift, cboDuplicatas(3), mlngDuplicatas
        End If
    End If
End Sub

Private Sub cboDuplicatas_LostFocus(Index As Integer)
    Dim MatrizDAO               As cMatrizContabilizacaoDAO
    Dim matriz                  As cMatrizContabilizacao

    Dim lngOpContabilbaixa      As Long

    If (mstrTabela = "Duplicatas") Then
        If (Index = 3) Then
            LibProc WL_EXIBIR
        End If
    End If
    'Carrega a opera��o cont�bil padr�o para as duplicatas e lancamentos
    If (Index = 3) Then
        If GetValue(mrstDuplicatas, "Tipo", "") <> cboDuplicatas(3).Text Then
            lngOperacao = 0
        Else
            If lngOperacao = 0 Then
                lngOperacao = strToLng(txtDuplicatas(40).Text)
            End If
        End If
        If cboDuplicatas(Index).Text <> "" And lngOperacao = 0 Then
            mblnAlteraValor = False
            Set MatrizDAO = New cMatrizContabilizacaoDAO
            Set matriz = MatrizDAO.Carregar(cboDuplicatas(3).Text)
            Set MatrizDAO = Nothing
            If Not matriz Is Nothing Then
                If mstrPagRec = "P" Then
                    If mstrTabela = "Duplicatas" Then
                        lngOperacao = matriz.duplicatasPagar
                        'pt. 89501 - Moacir Pfau(19/01/2009)
                        lngOpContabilbaixa = matriz.BaixaDuplicatasPagar
                    Else
                        lngOperacao = matriz.lancamentosPagar
                        'pt. 89501 - Moacir Pfau(19/01/2009)
                        lngOpContabilbaixa = matriz.BaixaLancamentosPagar
                    End If
                Else
                    If mstrTabela = "Duplicatas" Then
                        lngOperacao = matriz.duplicatasReceber
                        'pt. 89501 - Moacir Pfau(19/01/2009)
                        lngOpContabilbaixa = matriz.BaixaDuplicatasReceber
                    Else
                        lngOperacao = matriz.lancamentosReceber
                        'pt. 89501 - Moacir Pfau(19/01/2009)
                        lngOpContabilbaixa = matriz.baixaLancamentosReceber
                    End If
                End If
            Else
                lngOperacao = 0
            End If
            If Not EEdicao(mlngDuplicatas) Then
                txtDuplicatas(40).Text = lngOperacao
                txtDuplicatas(41).Text = lngOpContabilbaixa
            Else
                If txtDuplicatas(40).Text = "" Or txtDuplicatas(40).Text = "0" Then
                    txtDuplicatas(40).Text = lngOperacao
                    txtDuplicatas(41).Text = lngOpContabilbaixa
                Else
                    lngOperacao = strToLng(txtDuplicatas(40).Text)
                End If
            End If
            Set matriz = Nothing
            mblnAlteraValor = True
        Else
            lngOperacao = strToLng(txtDuplicatas(40).Text)
        End If
    End If
End Sub

Private Sub chkConciliado_Click()
    AlteraValor mlngDuplicatas
End Sub

Private Sub cmdAbreRateio_Click()
    FraRateio.Visible = True
    FraDuplicatas(1).Visible = False
    txtDuplicatas(20).SetFocus
    'pt. 86140 - Moacir Pfau(07/04/2008)
    lvwRateio.ListItems.Clear    'limpar list.
    txtDuplicatas(20).Text = ""  'limpa campos.
    txtDuplicatas(25).Text = ""
    txtDuplicatas(24).Text = ""
    txtDuplicatas(21).Text = ""
    txtDuplicatas(29).Text = ""
    txtDuplicatas(30).Text = ""
    Call TotalizaValorRateio
End Sub

Private Sub cmdAdicionar_Click()
  
    If Not IsValid(txtDuplicatas(20).Text) Then
        MsgFunc "Preencha o Centro de custo!"
        txtDuplicatas(20).SetFocus
        Exit Sub
    End If

    If Not IsValid(txtDuplicatas(24).Text) And Not IsValid(txtDuplicatas(21).Text) Then
        MsgFunc "Preencha o campo Valor ou Porcentagem"
        txtDuplicatas(21).SetFocus
        Exit Sub
    End If
  
    If txtDuplicatas(24).Text <> Empty And txtDuplicatas(21).Text <> Empty Then
        MsgFunc "Apenas um dos campos deve ser preenchido. Valor ou Porcentagem."
        txtDuplicatas(24).SetFocus
        Exit Sub
    End If
    
    If Not IsValid(lblDuplDesc(10).Caption) Then
        MsgFunc "Centro de custo n�o cadastrado!"
        txtDuplicatas(20).SetFocus
        Exit Sub
    End If
    
    If IsValid(txtDuplicatas(21).Text) And IsValid(txtDuplicatas(24).Text) Then
        MsgFunc "Informe apenas um valor!"
        txtDuplicatas(21).SetFocus
        Exit Sub
    End If
  
    If IsValid(txtDuplicatas(29).Text) And IsValid(txtDuplicatas(24).Text) Then
        MsgFunc "Acr�scimo e porcentagem n�o devem ser informados ao mesmo tempo!"
        txtDuplicatas(29).SetFocus
        Exit Sub
    End If
     
    If IsValid(txtDuplicatas(30).Text) And IsValid(txtDuplicatas(24).Text) Then
        MsgFunc "Abatimento e porcentagem n�o devem ser informados ao mesmo tempo!"
        txtDuplicatas(30).SetFocus
        Exit Sub
    End If
     
    If Len(txtDuplicatas(25).Text) = 0 Or Len(lblDuplDesc(11).Caption) = 0 Then
        MsgFunc "Conta n�o cadastrada!"
        txtDuplicatas(25).SetFocus
        Exit Sub
    End If
     
     'Verifica se conta est� ativa
    If GetFieldValue("Ctaati", "Contas", " [C�digo]=" & txtDuplicatas(25).Text) = "N" Then
        MsgBox "Conta " & txtDuplicatas(25).Text & " n�o est� ativa", vbCritical, MsgBoxCaption
        txtDuplicatas(25).SetFocus
        Exit Sub
    End If
  
    Dim bUsaPorc  As Boolean
    bUsaPorc = UsaPorcentagemnoRateio
    If lvwRateio.ListItems.Count > 0 Then
          ' Verificando se usa porcentagem ou valor
          If (bUsaPorc And IsValid(txtDuplicatas(21).Text)) Or ((Not bUsaPorc) And IsValid(txtDuplicatas(24).Text)) Then
              MsgFunc "S� � poss�vel utilizar uma forma de rateio de cada vez!"
              Exit Sub
          End If
    End If
  
    If IsValid(txtDuplicatas(24).Text) Then
        Dim dblTotPorcentagem  As Double
        dblTotPorcentagem = SomaPorcentagens()
        If CSng((dblTotPorcentagem + CSngDef(txtDuplicatas(24).Text))) <= CSng(100) Then
            lvwRateio.ListItems.add , , txtDuplicatas(20).Text, , DL_MARCADO
            lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(1) = lblDuplDesc(10).Caption
            lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(2) = Format(txtDuplicatas(24).Text, F4CASAS) & "%"
        Else
            MsgBox "Total de Porcentagens  � maior que 100%", vbCritical, "Rateio"
            Exit Sub
        End If
    Else
        Dim curTotRateio       As Currency
        Dim curTotRateioAcres  As Currency
        Dim curTotRateioAbat   As Currency
        curTotRateio = SomaValores()
        curTotRateioAcres = SomaValoresAcres()
        curTotRateioAbat = SomaValoresAbat()
        If CCur((curTotRateio + CCurDef(txtDuplicatas(21).Text))) <= CCur(txtDuplicatas(10).Text) Then
            If CCur((curTotRateioAcres + CCurDef(txtDuplicatas(29).Text))) <= CCur(txtDuplicatas(11).Text) Then
                If CCur((curTotRateioAbat + CCurDef(txtDuplicatas(30).Text))) <= CCur(txtDuplicatas(12).Text) Then
                    lvwRateio.ListItems.add , , txtDuplicatas(20).Text, , DL_MARCADO
                    lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(1) = lblDuplDesc(10).Caption
                    lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(3) = Format(txtDuplicatas(21).Text, FMOEDA)
                    lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(4) = Format(txtDuplicatas(29).Text, FMOEDA)
                    lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(5) = Format(txtDuplicatas(30).Text, FMOEDA)
                Else
                    MsgBox "A soma dos abatimentos � maior que " & txtDuplicatas(12).Text, vbCritical, "Rateio"
                    Exit Sub
                End If
            Else
                MsgBox "A soma dos acr�scimos � maior que " & txtDuplicatas(11).Text, vbCritical, "Rateio"
                Exit Sub
            End If
        Else
            MsgBox "A soma dos valores originais � maior que " & txtDuplicatas(10).Text, vbCritical, "Rateio"
            Exit Sub
        End If
    End If
  
    lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(6) = txtDuplicatas(25).Text
    txtDuplicatas(20).SetFocus
    'Pt. 114146 - Moacir Pfau(29/02/2012)
    Call TotalizaValorRateio
End Sub

Private Function SomaValores() As Currency
    Dim curTotal         As Currency
    Dim nCont            As Integer
  
    curTotal = 0
    For nCont = 1 To lvwRateio.ListItems.Count
        curTotal = curTotal + CCurDef(lvwRateio.ListItems(nCont).SubItems(3))
    Next
    SomaValores = curTotal
  
End Function

Private Function SomaValoresAcres() As Currency
    Dim curTotal         As Currency
    Dim nCont            As Integer
    
    curTotal = 0
    For nCont = 1 To lvwRateio.ListItems.Count
        curTotal = curTotal + CCurDef(lvwRateio.ListItems(nCont).SubItems(4))
    Next
    SomaValoresAcres = curTotal
End Function

Private Function SomaValoresAbat() As Currency
    Dim curTotal         As Currency
    Dim nCont            As Integer
    
    curTotal = 0
    For nCont = 1 To lvwRateio.ListItems.Count
        curTotal = curTotal + CCurDef(lvwRateio.ListItems(nCont).SubItems(5))
    Next
    SomaValoresAbat = curTotal
  
End Function

Private Function SomaPorcentagens() As Single
    Dim dblTotal         As Single
    Dim nCont            As Integer
    
    dblTotal = 0
    For nCont = 1 To lvwRateio.ListItems.Count
        dblTotal = dblTotal + CDblDef(Replace(lvwRateio.ListItems(nCont).SubItems(2), "%", ""))
    Next
    SomaPorcentagens = dblTotal
End Function

Private Function UsaPorcentagemnoRateio(Optional bConfere As Boolean = False) As Boolean
    Dim lngItens  As Long
    Dim dblRateio As Single
    Dim strTemp   As String
  
    For lngItens = 1 To lvwRateio.ListItems.Count
       If IsValid(lvwRateio.ListItems(lngItens).SubItems(2)) Then
          If bConfere Then
              strTemp = lvwRateio.ListItems(lngItens).SubItems(2)
              MidAll strTemp, "%", ""
              dblRateio = dblRateio + VBA.Round(CSngDef(strTemp), 4)
          Else
              UsaPorcentagemnoRateio = True
              Exit Function
          End If
       End If
    Next lngItens
  
    'Verificando se a porcentagem est� correta ou se usa %
    If bConfere Then
        If CSng(dblRateio) = CSng(100) Then
            UsaPorcentagemnoRateio = True
        Else
            MsgFunc "O rateio do valor tem que ser total, ou seja, 100%"
            UsaPorcentagemnoRateio = False
        End If
    Else
        UsaPorcentagemnoRateio = False
    End If
End Function

Private Sub cmdCancelar_Click()
    lvwRateio.ListItems.Clear
    FraRateio.Visible = False
    FraDuplicatas(1).Visible = True
    'Pt. 114146 - Moacir Pfau(29/02/2012)
    Call TotalizaValorRateio
End Sub

Private Sub cmdExcluir_Click()
    Dim lngItens As Long
    Dim Tem      As Boolean
    
    For lngItens = lvwRateio.ListItems.Count To 1 Step -1
        If lvwRateio.ListItems(lngItens).SmallIcon = DL_DESMARCADO Then
            lvwRateio.ListItems.Remove (lngItens)
             Tem = True
        End If
    Next
    If Not Tem Then
        MsgFunc "Para excluir um ou mais itens marque o(s) mesmo(s) com o X!"
    End If
    'Pt. 114146 - Moacir Pfau(29/02/2012)
    Call TotalizaValorRateio
End Sub

Private Sub cmdNominalRazaoSocial_Click()
    txtCheque(0).Text = GetFieldValue("Raz�o", "Empresas", "Apel = " & Quote(txtDuplicatas(2).Text, "'"), , NUL)
End Sub

Private Sub cmdProximoCheque_Click()
    Dim rstProximoCheque     As Object
    
    If AbreRecordset(rstProximoCheque, "Select * from Cheque " & _
          "WHERE Banco = " & CLngDef(txtDuplicatas(13).Text) & " AND Situa��o = 'Normal' " & _
          "AND (Cheque not in (Select Cheque from Duplicatas where Banco = Cheque.Banco) " & _
          "AND Cheque not in (Select Cheque from Lan�amentos where Banco = Cheque.Banco) " & _
          "AND Cheque not in (Select Cheque from [Transf Banc�ria] where Banco = Cheque.Banco)) " & _
          "ORDER BY Cheque ASC", dbOpenSnapshot) = WL_OK Then
        txtDuplicatas(16).Text = GetValue(rstProximoCheque, "Cheque", ZERO)
    Else
        txtDuplicatas(16).Text = ProximoNumero("Cheque", "Cheque", "Banco = " & CLngDef(txtDuplicatas(13).Text))
    End If
    FechaRecordset rstProximoCheque
End Sub

Private Sub cmdRateio_Click()
    Dim lngItens     As Long
    Dim strTemp      As String
    Dim vrTemp       As Currency
    Dim vrTempAcres  As Currency
    Dim vrTempAbat   As Currency
    Dim VrTotal      As Currency 'valor total rateado
    Dim VrTotalAcres As Currency 'valor total do acr�scimo rateado
    Dim VrTotalAbat  As Currency 'valor total do abatimento rateado
    Dim vrDuplicata  As Currency 'valor da duplicata
    Dim vrAcrescimo  As Currency 'valor do acr�scimo
    Dim vrAbatimento As Currency 'valor do abatimento
    Dim bUsaPorc     As Boolean
    Dim strSql       As String
    Dim lngCodigo    As Long
    Dim bGerouTodas  As Boolean
    Dim strUpdate    As String
    Dim lngParcela   As Long
  
On Error GoTo Error_Handler
    'Verificando a existencia de itens
    If lvwRateio.ListItems.Count = 0 Then
        MsgFunc "Informe o rateio!"
        txtDuplicatas(20).SetFocus
        Exit Sub
    End If
  
    'Realizando as conferencias antes do rateio final
    If Not IsValid(txtDuplicatas(10).Text) Then  'And (UsaPorcentagemnoRateio)
        MsgFunc "Preencha o campo de valor original para que o mesmo seja rateado!"
        txtDuplicatas(10).SetFocus
        Exit Sub
    End If
       
    'Acr�scimo
    If Not IsValid(txtDuplicatas(11).Text) Then
        txtDuplicatas(11).Text = "0"
    End If
  
    'Abatimento
    If Not IsValid(txtDuplicatas(12).Text) Then
        txtDuplicatas(12).Text = "0"
    End If
  
    vrDuplicata = CCurDef(txtDuplicatas(10).Text)
    vrAcrescimo = CCurDef(txtDuplicatas(11).Text)
    vrAbatimento = CCurDef(txtDuplicatas(12).Text)
  
    'Checando Empresa
    If Not IsValid(txtDuplicatas(2).Text) Then
        MsgFunc "O campo de Empresa n�o pode ficar em branco"
        txtDuplicatas(2).SetFocus
        Exit Sub
    End If
  
    'Checando datas
    If Not IsValid(txtDuplicatas(6).Text) Or Not IsValid(txtDuplicatas(7).Text) Or Not IsValid(txtDuplicatas(9).Text) Then
        MsgFunc "Os campos de data de Emiss�o,Libera��o e Vencimento s�o obrigat�rios!"
        Exit Sub
    End If
  
    vrTemp = ZERO
    'valido se o rateio bater� com o valor da duplicata, para o rateio de valores
    If Not UsaPorcentagemnoRateio Then
        vrTemp = SomaValores()
        If vrTemp <> vrDuplicata Then
            MsgFunc "Valor Original � diferente que o valor do rateio"
            Exit Sub
        End If
        vrTempAcres = SomaValoresAcres()
        If vrTempAcres <> vrAcrescimo Then
            MsgFunc "Valor de Acr�scimo � diferente do valor do rateio"
            Exit Sub
        End If
        vrTempAbat = SomaValoresAbat()
        If vrTempAbat <> vrAbatimento Then
            MsgFunc "Valor de Abatimento � diferente que o valor do rateio"
            Exit Sub
        End If
    Else
        'valido se a porcentagem informada atinge 100%, para o rateio de porcentagem
        If Not UsaPorcentagemnoRateio(True) Then Exit Sub
    End If
    
    'inicializo o totalizador
    VrTotal = 0
    VrTotalAcres = 0
    VrTotalAbat = 0
    
    'Conferindo se a porcentagem esta correta
    BeginTrans
    For lngItens = 1 To lvwRateio.ListItems.Count
        If UsaPorcentagemnoRateio Then
            strTemp = lvwRateio.ListItems(lngItens).SubItems(2)
            MidAll strTemp, "%", ""
            vrTemp = CSngDef(strTemp) * vrDuplicata / 100
            'arredondo o valor para 2 decimais
            vrTemp = Round(vrTemp, 2)
            vrTempAcres = CSngDef(strTemp) * vrAcrescimo / 100
            'arredondo o valor para 2 decimais
            vrTempAcres = Round(vrTempAcres, 2)
            vrTempAbat = CSngDef(strTemp) * vrAbatimento / 100
            'arredondo o valor para 2 decimais
            vrTempAbat = Round(vrTempAbat, 2)
            'acumulo o valor total para no fim atribuir a
            'diferenca, devido aos arredondamentos, no ultimo lan�amento
            VrTotal = VrTotal + vrTemp
            VrTotalAcres = VrTotalAcres + vrTempAcres
            VrTotalAbat = VrTotalAbat + vrTempAbat
            'se for o ultimo item do loop
            'acerto o arredondamento
            If lngItens = lvwRateio.ListItems.Count Then
                vrTemp = vrTemp + (vrDuplicata - VrTotal)
                vrTempAcres = vrTempAcres + (vrAcrescimo - VrTotalAcres)
                vrTempAbat = vrTempAbat + (vrAbatimento - VrTotalAbat)
            End If
        Else
            vrTemp = CCurDef(lvwRateio.ListItems(lngItens).SubItems(3))
            vrTempAcres = CCurDef(lvwRateio.ListItems(lngItens).SubItems(4))
            vrTempAbat = CCurDef(lvwRateio.ListItems(lngItens).SubItems(5))
        End If
     
        'Montando a SQL para inserir os registros rateados
        lngCodigo = CLngDef(txtDuplicatas(1).Text)
        
        strSql = "INSERT INTO Duplicatas(PagRec, Nota, Parcela, Empresa, Tipo, Descri��o, Emiss�o, Vencimento, Pagamento, Libera��o, [Valor Original], " & _
                  "Acr�scimo, Abatimento, Banco, Conta, Centro, Cheque, Moeda, " & IIf(ChkConciliado, "Conciliado,", "") & _
                  "[Valor da Moeda], Controle, Marca��o, Obs, Border�, cd_operacao_contabil, Usu�rio, proveniente_rateio) VALUES ('" & mstrPagRec & "', "
      
        If lngItens = 1 And EEdicao(mlngDuplicatas) Then
            'Resolvendo o UPDATE:
            strUpdate = "UPDATE Duplicatas SET [Valor Original]=" & Replace(vrTemp, ",", ".") & " ," & "[Acr�scimo]=" & Replace(vrTempAcres, ",", ".") & " ," & _
                    "[Abatimento]=" & Replace(vrTempAbat, ",", ".") & " ," & "Conta =" & CLngDef(lvwRateio.ListItems(lngItens).SubItems(6)) & " ," & _
                              "Centro =" & CLngDef(lvwRateio.ListItems(lngItens).Text) & " "
            strUpdate = strUpdate & "WHERE PagRec = " & Quote(mstrPagRec, "'") & " AND Nota = " & GetValue(mrstDuplicatas, "Nota", ZERO) & _
                    " AND Empresa = " & Quote(GetValue(mrstDuplicatas, "Empresa", NUL), "'") & " AND Tipo = " & _
                    Quote(GetValue(mrstDuplicatas, "Tipo", ZERO), "'") & " AND Parcela = " & GetValue(mrstDuplicatas, "Parcela", NUL)
            Call ExecuteSQL(strUpdate)
        Else
            AppendStr strSql, CStr(lngCodigo)                                 'C�digo/Nota
            lngParcela = ProximoNumero("Parcela", "Duplicatas", "PagRec = " & Quote(mstrPagRec, "''") & " AND Tipo = " & Quote(cboDuplicatas(3).Text, "''") & " AND Empresa= " & Quote(txtDuplicatas(2).Text, "''") & " AND Nota = " & lngCodigo)
            AppendStr strSql, ", " & lngParcela ' Parcela
            AppendStr strSql, ", " & Quote(txtDuplicatas(2).Text, "''")          'Empresa
            AppendStr strSql, ", " & Quote(cboDuplicatas(3).Text, "''")          'Tipo
            AppendStr strSql, ", " & Quote(txtDuplicatas(5).Text, "''")          'Descri��o
            AppendStr strSql, ", " & InverteData(txtDuplicatas(6).Text, True)    'Emiss�o
            AppendStr strSql, ", " & InverteData(txtDuplicatas(7).Text, True)    'Vencimento
            AppendStr strSql, ", " & IIf(IsValid(txtDuplicatas(8).Text), InverteData(txtDuplicatas(8).Text, True), "Null") 'Pagamento
            AppendStr strSql, ", " & InverteData(txtDuplicatas(9).Text, True)    'Libera��o
            AppendStr strSql, ", " & Replace(vrTemp, ",", ".")                      'Valor original
            AppendStr strSql, ", " & Replace(vrTempAcres, ",", ".")                 'Acr�scimo
            AppendStr strSql, ", " & Replace(vrTempAbat, ",", ".")
            AppendStr strSql, ", " & CLngDef(txtDuplicatas(13).Text)             'Banco
            If IsValid(lvwRateio.ListItems(lngItens).SubItems(6)) Then
                AppendStr strSql, ", " & CLngDef(lvwRateio.ListItems(lngItens).SubItems(6)) 'Conta
            Else
                AppendStr strSql, ", " & CLngDef(txtDuplicatas(14).Text)          'Conta
            End If
            AppendStr strSql, ", " & CLngDef(lvwRateio.ListItems(lngItens).Text) 'Centro
            AppendStr strSql, ", " & CLngDef(txtDuplicatas(16).Text)             'Cheque
            AppendStr strSql, ", " & Quote(txtDuplicatas(17).Text, "''")         'Moeda
            'Verifica se esta o flag de concilia��o esta ativo
            If ChkConciliado Then
                AppendStr strSql, ", True "                                        'Conciliado
            End If
            'o campo valor da moeda foi retirado do formul�rio.
            AppendStr strSql, ", " & ValStr(CMoeda(""))                          'Valor da Moeda
            AppendStr strSql, ", " & Quote(txtDuplicatas(19).Text, "''")         'Controle
            AppendStr strSql, ", 0"                                              'Marca��o
            AppendStr strSql, ", " & Quote(txtDuplicatas(23).Text, "''")         'Observa��o
            AppendStr strSql, ", 0"                                              'Border�
            'pt. 87144 - Moacir Pfau(08/07/2008)
            AppendStr strSql, ", " & CLngDef(txtDuplicatas(40).Text)
            AppendStr strSql, ", " & Quote(txtDuplicatas(26).Text, "''")
            AppendStr strSql, ", True)" 'pt. 88289 - Dulcino J�nior(10/10/2008)
            'pt. 88289 - Dulcino J�nior(10/10/2008)
            If ExecuteSQL(strSql) > 0 Then
                strSql = "INSERT INTO FFIRateioDuplicata(pag_rec_origem, nr_nota_origem, cd_empresa_origem, tp_registro_origem, "
                strSql = strSql & "nr_parcela_origem, pag_rec_destino, nr_nota_destino, cd_empresa_destino, tp_registro_destino, "
                strSql = strSql & "nr_parcela_destino, cd_centro, cd_conta, vl_valor) "
                strSql = strSql & "VALUES('" & mstrPagRec & "', " & lngCodigo & ", '" & txtDuplicatas(2).Text & "', '" & cboDuplicatas(3).Text & "', "
                strSql = strSql & txtDuplicatas(4).Text & ", '" & mstrPagRec & "', " & lngCodigo & ", '" & txtDuplicatas(2).Text & "', "
                strSql = strSql & "'" & cboDuplicatas(3).Text & "', " & lngParcela & ", " & CLngDef(lvwRateio.ListItems(lngItens).Text)
                strSql = strSql & ", " & CLngDef(lvwRateio.ListItems(lngItens).SubItems(6)) & ", " & Replace(vrTemp, ",", ".") & ")"
                If ExecuteSQL(strSql) = 0 Then
                    GoTo Error_Handler
                End If
            Else
                GoTo Error_Handler
            End If
        End If
    Next
    CommitTrans
    MsgFunc "Rateio conclu�do!"
    FraRateio.Visible = False
    FraDuplicatas(1).Visible = True
    cmdAbreRateio.Enabled = False
    Call DefEditNone(mlngDuplicatas)
    'pt. 86140 - Moacir Pfau(07/04/2008)
    mlngDuplicatas = 735
    Call txtDuplicatas_LostFocus(2)
    Exit Sub

Error_Handler:
    MsgBox "N�o foi possivel concluir o rateio.", vbInformation, NomeModulo
    Rollback
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
    Call GetKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub Form_Load()
    Dim intLabels As Integer
  
    SSTab1.Tab = 0
    'Protocolo 73916: Verifica se o usu�rio pode baixar duplicatas/lan�amentos
    If GetAcesso(LoadResString(2002)) <> SEM_ACESSO Then
        txtDuplicatas(8).Enabled = True
        lblDuplicatas(16).ForeColor = &H0&         'Ativado
    Else
        txtDuplicatas(8).Enabled = False
        lblDuplicatas(16).ForeColor = &H0         'Desativado
    End If
    ConfigCampos Me, Tag, Tag
    cmdAbreRateio.Visible = CentrodeCusto(MFinanceiro)
    
    'Retirando os Captions dos Labels de Descri��o que coloquei em
    'design time.
    For intLabels = 0 To 13
        lblDuplDesc(intLabels).Caption = NUL
    Next
    
    'Prefer� configurar o controle ListView no c�digo para facilitar
    'futuras altera��es
    lvwLancamentos.ColumnHeaders.add 1, , "N�mero", 975, lvwColumnLeft
    lvwLancamentos.ColumnHeaders.add 2, , "Tipo", 975, lvwColumnLeft
    lvwLancamentos.ColumnHeaders.add 3, , "Empresa", 1440, lvwColumnLeft
    lvwLancamentos.ColumnHeaders.add 4, , "Valor", 1440, lvwColumnRight
    
    'Carregando as imagens no controle ImageList a partir do arquivo
    'de recursos
    imgDupl.ImageHeight = 16
    imgDupl.ImageWidth = 16
    imgDupl.MaskColor = vbWhite
    imgDupl.UseMaskColor = True
    imgDupl.ListImages.add 1, "transferencia", LoadResBitmap(IDB_TRANSF)
    imgDupl.ListImages.add 2, "duplicata", LoadResBitmap(IDB_DUPLS)
    imgDupl.ListImages.add 3, "lancamento", LoadResBitmap(IDB_LANCTOS)
    
    'Define o ImageList a ser usado com o ListView
    lvwLancamentos.SmallIcons = imgDupl
    'PT. 81189 - Dulcino J�nior
    'Integra��o Cont�bil
    Label1.Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    txtDuplicatas(40).Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    lblDuplDesc(14).Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    Label2.Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    txtDuplicatas(41).Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    lblDuplDesc(15).Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    mblnAlteraValor = True
    'Pt. 94752 - Moacir Pfau(22/10/2009)
    txtDuplicatas(44).MaxLength = 3
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If (mintBaixa <> CDT_NORMAL) Then
        If (UnloadMode > vbFormCode) Then
            ' Quando esta janela � aberta pela janela de baixas, ela fica oculta,
            ' aguardando que a janela de Duplicatas seja fechada para retornar. Se o
            ' usu�rio fechar o Sistema antes de retornar para a janela de baixas ocorre
            ' um erro n�o-intersept�vel de exce��o do Windows. Aqui, ent�o, eu obrigo
            ' o usu�rio a sair da janela de baixas antes de sair do Sistema.
            '
            MsgFunc LoadResString(245)
            Cancel = True
            Exit Sub
        End If
    End If

    If (Not UnloadForm(mrstDuplicatas, Me, Tag, mlngDuplicatas)) Then
        ' Verifica se h� alguma altera��o nos campos de cheque
        If (EstaEditando(mlngCheques) And IsVisibleRecord(mlngCheques)) Then
            'Pt. 95368 - Moacir Pfau(17/11/2009)
            'If gTipoDB = Access Then mrstCheques.Edit
            mrstCheques("Nominal").value = txtCheque(0).Text
            mrstCheques("Hist�rico").value = txtCheque(1).Text
            mrstCheques.update
        End If
        FechaRecordset mrstCheques
    Else
        Cancel = True
    End If
    lblDuplDesc(14).Caption = ""
    lblDuplDesc(15).Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SavePosForm Me
    Set frmDuplicatas = Nothing
End Sub

Private Sub lvwLancamentos_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvwLancamentos.SortKey = ColumnHeader.Index - 1
End Sub

'SUB.......: Configure
'Objetivo..: Configura o Cadastro antes da abertura.
'Argumentos: [strTabela]: Nome da Tabela que ser� aberta: Duplicatas ou Lan�amentos
'            [strPagRec]: Tipo da Tabela, A Pagar ou A Receber.
'            [intBaixa] : Opcional. Se estiver abrindo da janela de baixas procede
'                         configura��es extras do cadastro.
'            [strBaixa] : Opcional. Instru��o Select para abertura do cadastro quando
'                         este � chamado da janela de baixas.
Public Sub Configure(strTabela As String, strPagRec As String, Optional intBaixa As Integer, Optional strBaixa As String)
    Dim strOptCombo As String           'Instru��o para as op��es do campo Tipo.

    SetPtr vbHourglass
    'Configurando as instru��es de abertura da tabela conforme o nome
    btnEfetiva.Visible = False
    'Instru��o utilizada na fun��o de Pesquisa
    mstrPesquisa = "SELECT Nota, Empresa, Tipo, Parcela, Descri��o, Emiss�o, Vencimento, Pagamento, Libera��o, [Valor Original], Acr�scimo, Abatimento, " & _
                   "Banco, Conta, Centro, Cheque, Moeda, [Valor da Moeda], Controle, Situa��o, Comiss�o, SeqNossoNumero FROM Duplicatas WHERE PagRec = '" & strPagRec & "'"
    'Instru��o utilizada na abertura do Cadastro
    mstrDuplicatas = "SELECT * FROM Duplicatas WHERE PagRec = '" & strPagRec & "'"
    
    'Lista de op��es do campo Tipo (ComboBox)
    strOptCombo = "SELECT Texto FROM Op��es WHERE Rotina = '" & OPT_DUPLICATAS & "';"
    If (intBaixa <> CDT_NORMAL) Then            'Se estiver sendo carregado da janela de baixas
        mstrDuplicatas = strBaixa                 'Instru��o para a abertura do Recordset
        'Completa a instru��o de pesquisa com as mesmas compara��es utilizadas para abrir a
        'tabela. Em baixas o usu�rio n�o pode pesquisar o Banco de Dados exceto nas Duplicatas
        'ou Lan�amentos abertos por ele em Baixas.
        MidStr mstrPesquisa, " PagRec = '" & strPagRec & "'", ExtractStr(strBaixa, "WHERE", NUL)
    End If

    'Campo Situa��o, vis�vel apenas quando a tabela for a Receber
    cmdNominalRazaoSocial.Visible = (strPagRec = "P")
    LoadResOptions 1000, cboDuplicatas(20)      'Carrega a lista de op��es do campo

    'Carregando as op��es do Campo Tipo
    ComboAddItem cboDuplicatas(3), strOptCombo, "Texto"

    'O campo Cheque permanece vis�vel apenas quando o tipo for a pagar
    lblDuplicatas(8).Visible = (strPagRec = "P")      'Label do campo Cheque
    txtDuplicatas(16).Visible = (strPagRec = "P")     'Campo Cheque
    cmdProximoCheque.Visible = (strPagRec = "P")
    'Oculta o campo Centro de Custo se o usu�rio desejar
    If Configuracao("Utiliza Pagamento a Fornecedores", True) And strPagRec = "P" Then
        FraDuplicatas(6).Visible = True
        FraDuplicatas(4).Top = 1110
    Else
        FraDuplicatas(6).Visible = False
        FraDuplicatas(4).Top = 360
    End If

    If Not CentrodeCusto(MFinanceiro) Then
        lblDuplicatas(7).Visible = False                'Label do campo Centro
        txtDuplicatas(15).Visible = False               'Campo Centro
        lblDuplDesc(3).Visible = False                  'Descri��o do campo Centro
    End If

    'pt. 81828 - Dulcino J�nior
    Call CenterForm(Me)

    'A parte de informa��es do cheque deve estar vis�vel somente quando A Pagar
    If strPagRec = "R" Then
        FraDuplicatas(5).Visible = False                'Oculta o Frame de informa��es do cheque
    End If
    If strPagRec = "P" Then
        txtDuplicatas(27).Visible = False
        lblDuplicatas(31).Visible = False
    End If
  
    'Termina de completar o caption do formul�rio conforme o tipo
    If strPagRec = "P" Then
        Caption = Caption & " a Pagar ou Pagas"
    Else
        Caption = Caption & " a Receber ou Recebidas"
    End If
  
    'Configura as vari�veis de controle
    mstrTabela = strTabela
    mstrPagRec = strPagRec
    mintBaixa = intBaixa

    'Abrindo o Cadastro
    AbreRecordset mrstDuplicatas, mstrDuplicatas, dbOpenDynaset     'Abre a tabela de duplicatas
    AbreRecordset mrstCheques, "Cheque", , , dbOpenDynaset  'Abre a tabela de cheques
    If intBaixa = CDT_NORMAL Then                         'Como padr�o abre com um novo registro
        Me.LibProc WL_NOVO
    End If
    DefineAcesso mlngDuplicatas, Acesso
    mlngCheques = WL_USERADDNEW                         'Define a vari�vel para os campos do cheque
    DefineAcesso mlngCheques, AC_CADASTRAR Or AC_EDITAR 'Define o acesso aos campos do cheque
    
    'Se a abertura do cadastro for atrav�s das baixas reconfiguro o acesso do usu�rio
    If (intBaixa > CDT_NORMAL) Then
        DeleteFlag AC_CADASTRAR, mlngDuplicatas    'N�o � permitido adicionar duplicatas em baixas
        If (CompStr(strTabela, "Duplicatas") And CompStr(strPagRec, "R")) Then
            txtDuplicatas(10).Enabled = False        'N�o � permitido alterar o valor original quando em baixas de Duplicatas a Receber
        End If
        LibProc WL_PRIMEIRO, MC_MOVEFIRST          'Posiciona no primeiro registro
    End If
    SetPtr vbDefault
End Sub

'FUNCTION..: DuplVerifique
'Objetivo..: Faz as verifica��es padr�o do cadastro
'Retorna...: True se for poss�vel salvar, False se n�o.
Private Function DuplVerifique() As Boolean
    Dim strOptions As String
    Dim strData    As String

    SetPtr vbHourglass
    ' Verificando as datas do cadastro Emiss�o (Verifica se a data de emiss�o � uma data v�lida)
    If Not EData(txtDuplicatas(6).Text) Then
        MsgFunc ResolveResString(26, resUM, txtDuplicatas(6).Text), vbInformation
        GoTo DuplVerifique_Erro
    End If
    ' Vencimento
    If Not EData(txtDuplicatas(7).Text) Then
        MsgFunc ResolveResString(26, resUM, txtDuplicatas(7).Text), vbInformation
        GoTo DuplVerifique_Erro
    Else
        ' Verifica se a data de Vencimento n�o � menor que a data de Emiss�o
        If DateDiff("d", txtDuplicatas(6).Text, txtDuplicatas(7).Text) < 0 Then
            MsgFunc ResolveResString(139, resUM, "de Vencimento", resDOIS, "de Emiss�o"), vbInformation
            GoTo DuplVerifique_Erro
        End If
    End If
    strData = CDateDef(txtDuplicatas(9).Text)
    If CLngDef(txtDuplicatas(15).Text) > 0 And Len(strData) Then
        ' Verifica se a data de libera��o est� dentro da data limite do centro de custo
        If DataLimiteCentroCusto(CLngDef(txtDuplicatas(15).Text), strData) Then
            GoTo DuplVerifique_Erro
        End If
    End If
       
    ' Pagamento
    If txtDuplicatas(8).Text <> "" Then         'Se o usu�rio indicou o pagamento
        If IsDate(txtDuplicatas(8).Text) Then
            ' Se a data de Pagamento n�o � anterior a emiss�o
            If (DateDiff("d", txtDuplicatas(6).Text, txtDuplicatas(8).Text) < 0) Then
                MsgFunc ResolveResString(139, resUM, "de Pagamento", resDOIS, "de Emiss�o"), vbInformation
                GoTo DuplVerifique_Erro
            End If
        Else
            MsgBox "Informe uma data de pagamento v�lida.", vbInformation, NomeModulo
            txtDuplicatas(8).SetFocus
            GoTo DuplVerifique_Erro
        End If
    Else    'Se o campo Cheque estiver preenchido n�o deixa Pagamento passar em Branco
        If mstrPagRec = "P" And IsValid(txtDuplicatas(16).Text) Then
            MsgBox ResolveResString(23, resUM, "Pagamento"), vbInformation, MsgBoxCaption
            GoTo DuplVerifique_Erro
        End If
    End If
    
    'Exibe a mensagem caso a Data de Pagamento seja posteior a Data de Vencimento
    If EData(txtDuplicatas(8).Text) And EData(txtDuplicatas(7).Text) Then
        If CDateDef(txtDuplicatas(8).Text) > CDateDef(txtDuplicatas(7).Text) Then
            If Not (CDateDef(txtDuplicatas(11).Text) > 0 Or CDateDef(txtDuplicatas(12).Text) > 0) Then
                MsgFunc "A Data de Pagamento informada est� em atraso h� " & _
                DateDiff("d", CDateDef(txtDuplicatas(7).Text), CDateDef(txtDuplicatas(8).Text)) & " dia(s)." & _
                vbCrLf & "Informe 'Acr�scimo' ou 'Multa' se necess�rio."
            End If
        End If
    End If
  
    If Not IsValid(GetValue(mrstDuplicatas, "Pagamento", NUL)) And IsValid(txtDuplicatas(8).Text) Then
        If Not IsValid(txtDuplicatas(13).Text) Then
            MsgFunc "O Campo Banco dever� ser preenchido"
            GoTo DuplVerifique_Erro
        End If
    End If
  
    'Libera��o
    If EData(txtDuplicatas(9).Text) Then    'Se for uma data v�lida
        If EData(txtDuplicatas(8).Text) Then    'Se o usu�rio preencheu o campo Pagamento
            If DateDiff("d", txtDuplicatas(8).Text, txtDuplicatas(9).Text) < 0 Then
                'A data de Libera��o n�o pode ser menor que a data de Pagamento
                MsgFunc ResolveResString(139, resUM, "de Libera��o", resDOIS, "de Pagamento"), vbInformation
                GoTo DuplVerifique_Erro
            End If
        End If
    End If
  
    'Verificando se o Banco indicado existe no cadastro de Bancos
    If IsValid(txtDuplicatas(13).Text) Then
        If Len(lblDuplDesc(1).Caption) = 0 Then
            If MsgBox(ResolveResString(35, resUM, txtDuplicatas(13).Text, resDOIS, "Bancos"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                LibProc "Bancos"
            End If
            GoTo DuplVerifique_Erro
        End If
    Else
        'Se n�o h� n�mero de banco o usu�rio n�o pode especificar um
        'n�mero de cheque.
        If mstrPagRec = "P" And IsValid(txtDuplicatas(16).Text) Then
            MsgFunc LoadResString(249)
            GoTo DuplVerifique_Erro
        End If
    End If
  
    'BANCO - Verificando se Carteira existe no Cadastro  de Carteiras
    If IsValid(txtDuplicatas(13).Text) Then
        If IsValid(txtDuplicatas(27).Text) Then
            If Recordcount("SELECT Carteira From Carteiras WHERE Banco=" & CLngDef(txtDuplicatas(13).Text) & " AND Carteira=" & Quote(txtDuplicatas(27).Text, "'")) = 0 Then
                MsgFunc " Carteira n�o cadastrada no Banco " & txtDuplicatas(13).Text
                GoTo DuplVerifique_Erro
            End If
        End If
    Else
        MsgBox "O campo 'Banco' n�o pode ser zero", vbCritical, MsgBoxCaption
        GoTo DuplVerifique_Erro
    End If
  
    'CONTA - Verificando se a Conta indicada existe no cadastro de Contas Cont�beis
    If IsValid(txtDuplicatas(14).Text) Then
        If Len(lblDuplDesc(2).Caption) = 0 Then
            If MsgBox(ResolveResString(35, resUM, txtDuplicatas(14).Text, resDOIS, "Contas"), _
                vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                LibProc "Contas"
            End If
            GoTo DuplVerifique_Erro
        End If
    Else
        MsgBox "O campo 'Conta' n�o pode ser zero", vbCritical, MsgBoxCaption
        GoTo DuplVerifique_Erro
    End If
  
    'Verificar se a conta est� ativa ou nao
    If GetFieldValue("Ctaati", "Contas", " [C�digo]=" & txtDuplicatas(14).Text) = "N" Then
        MsgBox "Conta " & txtDuplicatas(14).Text & " n�o est� ativa", vbCritical, MsgBoxCaption
        txtDuplicatas(14).SetFocus
        GoTo DuplVerifique_Erro
    End If

    'Verificando se o C�digo de Centro de Custo existe no Cadastro
    If txtDuplicatas(15).Visible Then
        If (IsValid(txtDuplicatas(15).Text)) Then
            If (Len(lblDuplDesc(3).Caption) = 0) Then
                If MsgBox(ResolveResString(35, resUM, txtDuplicatas(15).Text, resDOIS, "Centros de Custo"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                    LibProc "Custos"
                End If
                GoTo DuplVerifique_Erro
            End If
        Else
            MsgFunc ResolveResString(IDS_COMPLETECAMPO, resUM, "Centro de Custo")
            GoTo DuplVerifique_Erro
        End If
    End If

    'Verificando se a Moeda indicada existe no cadastro de Moedas e �ndices
    If Len(txtDuplicatas(17).Text) > 0 Then
        If ConfereDuplicidade("Moeda", "Moedas", "Moeda = '" & txtDuplicatas(17).Text & "'") = 0 Then
            If MsgBox(ResolveResString(35, resUM, txtDuplicatas(17).Text, resDOIS, "Moedas & �ndices"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                LibProc "Moedas"
            End If
            GoTo DuplVerifique_Erro
        End If
    End If
  
    If IsValid(txtDuplicatas(2).Text) Then
        If Recordcount("SELECT Raz�o, Apel FROM Empresas WHERE Apel = '" & txtDuplicatas(2).Text & "'") = 0 Then
            If (MsgBox(ResolveResString(35, "|1", txtDuplicatas(2).Text, "|2", "Empresas"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes) Then
                LibProc "Empresas"
            End If
            GoTo DuplVerifique_Erro
        End If
    End If
  
    'Verifica se n�o h� datas diferentes para o cheque cadastrado agora
    If mstrPagRec = "P" Then
        If Not ConfDataCheque(txtDuplicatas(13).Text, txtDuplicatas(16).Text, txtDuplicatas(8).Text, mlngDuplicatas) Then
            GoTo DuplVerifique_Erro
        End If
    End If
    
    'Protocolo 78787 -Alisson
    ' Verificando se valor original maior que zero
    If IsNumeric(txtDuplicatas(10).Text) Then
        If txtDuplicatas(10) = 0 Then
            MsgBox "Valor Original deve ter um valor maior que zero.", vbInformation
            GoTo DuplVerifique_Erro
        End If
    Else
        MsgBox "Valor Original deve ter um valor maior que zero.", vbInformation
        GoTo DuplVerifique_Erro
    End If
    
    If Len(cboDuplicatas(3).Text) > 0 Then
        ' Verificando se o tipo da duplicata digitado � um novo tipo
        If Recordcount("SELECT TIPO FROM [TIPOS GLOBAIS] WHERE TIPO = '" & cboDuplicatas(3).Text & "'") = 0 Then
            MsgBox "Tipo global informado n�o cadastrado!", vbInformation
            cboDuplicatas(3).SetFocus
            GoTo DuplVerifique_Erro
        End If
    End If
  
    'Verifica��o do campo Forma de Pagamento
    If IsNumeric(txtDuplicatas(18).Text) Then
        If CInt(txtDuplicatas(18).Text) > 0 Then
            If Len(lblDuplDesc(13).Caption) = 0 Then
                MsgBox "Forma de pagamento n�o encontrada.", vbInformation, "Valida��o de Campos"
                txtDuplicatas(18).SetFocus
                GoTo DuplVerifique_Erro
            End If
        End If
    End If
  
    If IsDate(txtDuplicatas(9).Text) Then
        'Se a data de pagamento estiver informada
        If IsDate(txtDuplicatas(8).Text) Then
            If CDate(txtDuplicatas(9).Text) < CDate(txtDuplicatas(8).Text) Then
                MsgBox "A data de libera��o deve ser maior do que a data de pagamento.", vbInformation
                txtDuplicatas(9).SetFocus
                GoTo DuplVerifique_Erro
            End If
        Else
            If CDate(txtDuplicatas(9).Text) < CDate(txtDuplicatas(7).Text) Then
                MsgBox "A data de libera��o deve ser maior do que a data de vencimento do documento.", vbInformation
                txtDuplicatas(9).SetFocus
                GoTo DuplVerifique_Erro
            End If
        End If
    End If
    
    'pt. 89506 - Dulcino J�nior (29/10/2008)
    If strToLng(txtDuplicatas(4).Text) = 0 Then
        Call MsgBox("O campo parcela deve ser preenchido.", vbInformation, NomeModulo)
        txtDuplicatas(4).SetFocus
        GoTo DuplVerifique_Erro
    End If
    
    'pt. 86728 - Moacir Pfau(09/06/2008)
    DuplVerifique = fEmpresaBloqueada(txtDuplicatas(2).Text, CDate(txtDuplicatas(6).Text))
    If Not DuplVerifique Then
       GoTo DuplVerifique_Erro
    End If
    
    DuplVerifique = True
DuplVerifique_Erro:
    SetPtr vbDefault
End Function

'SUB.......: NovoRegistro
'Objetivo..: Configura alguns controles como adi��o de registro.
'Argumento.: [blnProcChave]: Quando a rotina deve procurar a nova chave.
Private Sub NovoRegistro(blnProcChave As Boolean)

    If blnProcChave Then
        If CompStr(mstrTabela, "Duplicatas") Then
            If mstrPagRec = "R" Then
                txtDuplicatas(1).Text = ProximoNumero("Nota", "Duplicatas", "Tipo = '" & _
                                                  cboDuplicatas(3).Text & "' AND " & _
                                                  "PagRec = '" & mstrPagRec & "'")
            End If
        Else
            txtDuplicatas(1).Text = ProximoNumero("C�digo", "Lan�amentos", _
                                                IIf(SeqLancamentos, NUL, "PagRec = '" & mstrPagRec & "'"))
          
        End If
    End If

    txtDuplicatas(26).Text = UserName
    txtDuplicatas(3).Text = Date
    
    txtDuplicatas(0).Text = mstrPagRec
    
    Call Selecione(txtDuplicatas(1))
    DoEvents
    DefAddNew mlngDuplicatas

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If IsValid(txtDuplicatas(2).Text) Then
        MsgBar IIf(IsValid(txtDuplicatas(1).Text), IIf(mstrTabela = "Duplicatas", " Nota: " & txtDuplicatas(1).Text, " Lan�amento: " & txtDuplicatas(1).Text), " ") _
            & IIf(IsValid(cboDuplicatas(3).Text), " - Tipo de Registro:" & cboDuplicatas(3).Text, " ") _
            & IIf(IsValid(txtDuplicatas(2).Text), " - Empresa: " & txtDuplicatas(2).Text, " ") _
            & IIf(mstrTabela = "Duplicatas", IIf(IsValid(txtDuplicatas(4).Text), " - Parcela: " & txtDuplicatas(4).Text, " "), " ")
    End If
    'Pt. 94752 - Moacir Pfau(21/10/2009)
    Call fPreenche_CodCobranca
  
End Sub

Private Sub txtCheque_Change(Index As Integer)
    AlteraValor mlngCheques
End Sub

Private Sub txtDuplicatas_Change(Index As Integer)
    Dim strProcura As String
    
    Select Case Index
        Case 2 ' Campo Empresa
            strProcura = "SELECT Raz�o, Apel FROM Empresas WHERE Apel = '" & txtDuplicatas(2).Text & "';"
            GetAssocValue strProcura, lblDuplDesc(0)
            If (mstrTabela = "Lan�amentos") Then
                'Empresa n�o faz parte da chave em Lan�amentos
                AlteraValor mlngDuplicatas
            End If
            'Pt. 94752 - Moacir Pfau(21/10/2009)
            Call fPreenche_CodCobranca
        Case 6 To 9 ' Campos: Data de Emiss�o, Vencimento, Pagamento, Libera��o
            lblDuplDesc(Index - 1).Caption = Semana(txtDuplicatas(Index).Text, raUmaPalavra)
        Case 10 To 12 ' Valores atualizam Total
            ExibeSoma
        Case 13 ' Campo Banco
            strProcura = "SELECT Nome FROM Bancos WHERE Banco = " & txtDuplicatas(13).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(1)
            txtBancoCheque.Text = txtDuplicatas(13).Text
        Case 43 ' Campo Carteira
            strProcura = "SELECT desc_carteira FROM FFICarteira WHERE id_carteira = " & txtDuplicatas(43).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(16)
        Case 14, 25 ' Campo Conta
            strProcura = "SELECT Descri��o FROM Contas WHERE C�digo = " & txtDuplicatas(Index).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(IIf(Index = 14, 2, 11))
        Case 15, 20 ' Campo Centro de Custo
            strProcura = "SELECT Descri��o FROM Centros WHERE C�digo = " & txtDuplicatas(Index).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(IIf(Index = 20, 10, 3))
        Case 16 ' Campo n�mero do cheque
            txtChequeCheque.Text = txtDuplicatas(16).Text
        Case 17 ' Campo Moeda
            strProcura = "SELECT Descri��o, Moeda FROM Moedas WHERE Moeda = '" & txtDuplicatas(17).Text & "';"
            GetAssocValue strProcura, Nothing, txtDuplicatas(17)
        Case 18 'Forma de Pagamento
            strProcura = "SELECT Nome FROM [Formas de Pagamento] WHERE C�digo = " & txtDuplicatas(18).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(13)
        Case 40 'Opera��o cont�bil
            If Len(txtDuplicatas(Index).Text) > 0 Then
                lblDuplDesc(14).Caption = GetFieldValue("descricao", "OperacaoContabil", "cd_operacao = " & txtDuplicatas(Index).Text)
            Else
                lblDuplDesc(14).Caption = vbNullString
            End If
        Case 41 'Operacao Cont�bil de Baixa
            If Len(txtDuplicatas(Index).Text) > 0 Then
                lblDuplDesc(15).Caption = GetFieldValue("descricao", "OperacaoContabil", "cd_operacao = " & txtDuplicatas(Index).Text)
            Else
                lblDuplDesc(15).Caption = vbNullString
            End If
    End Select
    
    If mstrTabela = "Lan�amentos" Then
        If Len(txtDuplicatas(1).Text) = 0 Or Len(txtDuplicatas(4).Text) = 0 Then Exit Sub
    End If
    If (Index > 4) Then
        If Index <> 40 And Index <> 45 Then
            If Index = 40 Or Index = 45 Then
                Debug.Assert False
            End If
            AlteraValor mlngDuplicatas
        Else
            'pt. 83525 - Dulcino J�nior (27/09/2007)
            If mblnAlteraValor Then
                AlteraValor mlngDuplicatas
            End If
        End If
    End If
End Sub

Private Sub txtDuplicatas_GotFocus(Index As Integer)
    Dim strMensagem As String
    Dim strEdidado  As String
On Error GoTo TrapErro

    Select Case txtDuplicatas(Index).DataField
        Case "Empresa"
            strMensagem = ResolveResString(75, resUM, "Empresas")
        Case "Banco"
            strMensagem = ResolveResString(75, resUM, "Bancos")
        Case "Conta"
            strMensagem = ResolveResString(75, resUM, "Contas")
        Case "Centro"
            strMensagem = ResolveResString(75, resUM, "Centro de Custo")
        Case "Cheque"
            strMensagem = ResolveResString(75, resUM, "Cheques")
        Case "Moeda"
            strMensagem = ResolveResString(75, resUM, "Moedas e �ndices")
        Case "Obs"
            'Posiciona no segundo tab
            SSTab1.Tab = 1
        Case Else  ' Qualquer outro campo
            strMensagem = NUL
    End Select
    Selecione txtDuplicatas(Index)
    If IsValid(txtDuplicatas(2).Text) Then
        MsgBar IIf(IsValid(txtDuplicatas(1).Text), IIf(mstrTabela = "Duplicatas", " Nota: " & txtDuplicatas(1).Text, " Lan�amento: " & txtDuplicatas(1).Text), " ") _
        & IIf(IsValid(cboDuplicatas(3).Text), " - Tipo de Registro:" & cboDuplicatas(3).Text, " ") _
        & IIf(IsValid(txtDuplicatas(2).Text), " - Empresa: " & txtDuplicatas(2).Text, " ") _
        & IIf(mstrTabela = "Duplicatas", IIf(IsValid(txtDuplicatas(4).Text), " - Parcela: " & txtDuplicatas(4).Text, " "), " ") _
        & "  -  " & DescCampo(mrstDuplicatas, txtDuplicatas(Index).DataField) & strMensagem
    Else
        MsgBar DescCampo(mrstDuplicatas, txtDuplicatas(Index).DataField) & strMensagem
    End If
    
    'Autor: Edilberto
    'Data : 24/09/2007
    'PT   : 83663
    If txtDuplicatas(39).Text <> "" Then
        If Index = 7 Or Index = 10 Or Index = 13 Then
            strEdidado = txtDuplicatas(Index).Text
            If MsgBox("Esse t�tulo j� possui boleto processado. Se continuar a linha digit�vel do boleto ser� zerada e n�o ser� poss�vel processar o Retorno banc�rio. Continuar?", vbYesNo, "Confirma��o") = vbYes Then
                strEdidado = ""
                ExecuteSQL ("UPDATE [Duplicatas] SET LINDIG = '" & strEdidado & "' WHERE Nota = " & txtDuplicatas(1).Text)
                ExecuteSQL ("UPDATE [Duplicatas] SET CodBar = '" & strEdidado & "' WHERE Nota = " & txtDuplicatas(1).Text)
                ExecuteSQL ("UPDATE [Duplicatas] SET NOSNUM = '" & strEdidado & "' WHERE Nota = " & txtDuplicatas(1).Text)
                ExecuteSQL ("UPDATE [Duplicatas] SET AGECCE = '" & strEdidado & "' WHERE Nota = " & txtDuplicatas(1).Text)
                txtDuplicatas(39).Text = ""
                'Pt. 95158 - Moacir Pfau(27/10/2009)
                txtDuplicatas(42).Text = ""
            Else
                txtDuplicatas(Index).Text = strEdidado
            End If
        End If
    End If
    
    'pt. 86868 - Moacir Pfau(13/05/2008)
    If Index = 4 Then
        If txtDuplicatas(4).Text = "" Or txtDuplicatas(4).Text = 0 Then
            txtDuplicatas(4).Text = 1
        End If
    End If
    
    Exit Sub
  
TrapErro:
  If err > 0 Then
    If err = 3270 Then
      err = 0
    Else
      DAOErros vbNullString
    End If
  End If
End Sub

Private Sub txtDuplicatas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strPCampo As String
    
    If mstrTabela = "Duplicatas" Then
        If Index > 0 And Index < 5 Then
            ControlaChave KeyCode, Shift, txtDuplicatas(Index), mlngDuplicatas
        End If
    Else
        If Index = 1 Then                 'Em Lan�amentos apenas C�digo � chave
            ControlaChave KeyCode, Shift, txtDuplicatas(1), mlngDuplicatas
        End If
    End If
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Select Case Index
            Case 2 ' Empresa
                strPCampo = "SELECT Apel, Raz�o, Pessoa, Tipo, [CNPJ/CPF], [IEst/RG], CCM, " _
                & "Ramo, Endere�o, Bairro, CEP, Cidade, Estado, " _
                & "Regi�o, Pa�s, Fone1, Ramal1, Contato, Dpto, Vendedor " _
                & "FROM Empresas"
                ' Verifica a configura��o para separar as empresas por tipo
                If mstrPagRec = "P" Then
                    AppendStr strPCampo, " WHERE Tipo <> '" & GetResOptions(1003, 2) & "';" 'Cliente
                Else
                    AppendStr strPCampo, " WHERE Tipo <> '" & GetResOptions(1003, 1) & "';" 'Fornecedor
                End If
                PCampo "Empresas", strPCampo, PB_CAMPO, txtDuplicatas(2), 0
            Case 13 ' Banco
                PCampo "Bancos", "Bancos", PB_CAMPO, txtDuplicatas(13), 0
            Case 14, 25 ' Conta
                'pt. 83864 - Dulcino J�nior (11/10/2007)
                PCampo "Contas", "SELECT Contas.C�digo as Conta, Contas.Descri��o as [Descri��o da Conta], Grupos.C�digo as Grupo, Grupos.Descri��o as [Descri��o do Grupo] " & _
                       " FROM Grupos INNER JOIN Contas ON Grupos.C�digo = Contas.Grupo where Contas.Ctaati='S' " & _
                       " ORDER BY Grupos.C�digo,Contas.C�digo", PB_CAMPO, txtDuplicatas(Index), 0
            Case 15, 20 ' Centro de Custo
                PCampo "Centro de Custo", "Centros", PB_CAMPO, txtDuplicatas(Index), 0
            Case 16 ' Campo Cheque
                If IsValid(txtDuplicatas(13).Text) Then
                    PCampo "Cheque", "SELECT * FROM Cheque WHERE Banco = " & txtDuplicatas(13).Text & ";", _
                    PB_CAMPO, txtDuplicatas(16), 1
                Else
                    PCampo "Cheque", "Cheque", PB_CAMPO, txtDuplicatas(16), 1
                End If
            Case 17 ' Moeda
                PCampo "Moedas e �ndices", "Moedas", PB_CAMPO, txtDuplicatas(17), 0
            Case 18 'Forma de Pagamento
                PCampo "Formas de Pagamento", "SELECT * FROM [Formas de Pagamento]", PB_CAMPO, txtDuplicatas(18), "C�digo"
            Case 27 'Carteira
                If IsValid(txtDuplicatas(13).Text) Then
                    PCampo "Carteiras", "Select Carteira from Carteiras WHERE Banco=" & CLngDef(txtDuplicatas(13).Text), PB_CAMPO, txtDuplicatas(27), 0
                End If
            Case 40 'Opera��o Cont�bil
                PCampo "Opera��es Contabeis", "OperacaoContabil", pbCampo, txtDuplicatas(40), "cd_operacao"
            Case 41 'Opera��o Cont�bil Baixa
                PCampo "Opera��es Contabeis", "OperacaoContabil", pbCampo, txtDuplicatas(41), "cd_operacao"
            Case 44
                'Pt. 94752 - Moacir Pfau(21/10/2009)
                Call fLocaliza_CodCobranca
        End Select
    End If
End Sub

'Pt. 94752 - Moacir Pfau(21/10/2009)
Private Sub fLocaliza_CodCobranca()
    Dim strApel                 As String
    Dim lngCodigo               As Long
    Dim strTipo                 As String
    Dim strEndereco             As String
    Dim strBairro               As String
    Dim strCep                  As String
    Dim strCidade               As String
    Dim strEstado               As String
    
    If CStr(txtDuplicatas(2).Text) <> "" Then
        If PMultiCampo("Selecione o endere�o", "SELECT [Endere�o],Bairro,CEP,Cidade,Estado,Apel,[C�digo],Tipo FROM [Empresas Endere�os] WHERE Tipo = 'Cobran�a' AND Apel = '" & txtDuplicatas(2).Text & "'", pbCampo, "Apel;C�digo;Tipo;Endere�o;Bairro;CEP;Cidade;Estado", strApel, lngCodigo, strTipo, strEndereco, strBairro, strCep, strCidade, strEstado) Then
            etxCobrancaCep.valorTexto = strCep
            etxCobrancaCidade.valorTexto = strCidade
            etxCobrancaEstado.valorTexto = strEstado
            etxCobrancaEndereco.valorTexto = strEndereco
            etxCobrancaBairro.valorTexto = strBairro
            txtDuplicatas(44).Text = lngCodigo
        End If
    End If
End Sub

'Pt. 94752 - Moacir Pfau(21/10/2009)
Private Sub fPreenche_CodCobranca()
    Dim strSql                  As String
    Dim rstTab                  As Object
    
    etxCobrancaCep.Clear: etxCobrancaCidade.Clear: etxCobrancaEstado.Clear: etxCobrancaEndereco.Clear: etxCobrancaBairro.Clear
    If CStr(txtDuplicatas(2).Text) <> "" And val(txtDuplicatas(44).Text) > 0 Then
        strSql = "SELECT [Endere�o],Bairro,CEP,Cidade,Estado,Apel,[C�digo],Tipo FROM [Empresas Endere�os] WHERE Tipo = 'Cobran�a' AND Apel = '" & txtDuplicatas(2).Text & "'"
        'Pt. 95368 - Moacir Pfau(03/11/2009)
        If (AbreRecordset(rstTab, strSql, dbOpenDynaset) = WL_OK) Then
            etxCobrancaCep.valorTexto = GetValue(rstTab, "Cep")
            etxCobrancaCidade.valorTexto = GetValue(rstTab, "Cidade")
            etxCobrancaEstado.valorTexto = GetValue(rstTab, "Estado")
            etxCobrancaEndereco.valorTexto = GetValue(rstTab, "Endere�o")
            etxCobrancaBairro.valorTexto = GetValue(rstTab, "Bairro")
        End If
        FechaRecordset (rstTab)
    End If
End Sub

Private Sub txtDuplicatas_KeyPress(Index As Integer, KeyAscii As Integer)
     Select Case Index
        Case 1 ' Campo Nota
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, InputMask(mrstDuplicatas, 1)
        Case 2 ' Campo Empresa
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, MaskEmpresa
        Case 13, 32 ' Campo Banco
            If Index = 13 Then
                SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask("Bancos", "Banco")
            ElseIf Index = 32 Then
                SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask(mstrTabela, "CheBan")
            End If
        Case 14, 3 ' Campo Conta
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask("Contas", "C�digo")
        Case 15, 20 ' Campo Centro de Custo
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask("Centros", "C�digo")
        Case 16 ' Campo Cheque
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask("Cheque", "Cheque")
        Case 4 ' Campo Parcela
            SetMascara KeyAscii, txtDuplicatas(4).SelStart, "###"
        Case 6 To 9 ' Campos Emiss�o, Vencimento, Pagamento e Libera��o
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, MASK_DATE4
        Case 10 To 12, 18, 21, 24, 29, 30 ' Campos Valor Original, Acr�scimo, Abatimento, Valor em Moeda
            If Index <= 12 Then
                ValidaNaoAceitaPonto KeyAscii
            End If
            DMoeda KeyAscii
        Case 22, 31, 36, 37, 38   ' 31 Valor do desconto por pontualide, 36 Valor Multa, 37 Percentual de Multa, 38 Valor Juros de Mora Di�rio
            DValor KeyAscii
        Case 40, 41 'Campo Opera��o cont�bil
            'Pt. 00000 - Moacir Pfau(13/04/2009)
            If KeyAscii <> 8 Then
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    KeyAscii = 0
                End If
            End If
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask(mstrTabela, "cd_operacao_contabil")
        'Pt. 94752 - Moacir Pfau(22/10/2009)
        Case 44
            If KeyAscii <> 8 Then
                If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
                    KeyAscii = 0
                End If
            End If
    End Select
End Sub

Private Sub txtDuplicatas_LostFocus(Index As Integer)
    Dim datLiberacao                    As Date
    'pt.81487 Ivo Sousa (25/10/07)
    Dim strProcura                      As String
    'Pt. 88817 - Moacir Pfau(06/11/2009)
    Dim dblPerDespesaFinanceira         As Double
    
    If Index = 2 Then
        txtDuplicatas(Index).Text = FormataEmpresa(txtDuplicatas(Index).Text)
        strProcura = "SELECT Raz�o, Apel FROM Empresas WHERE Apel = '" & txtDuplicatas(2).Text & "';"
        GetAssocValue strProcura, lblDuplDesc(0), txtDuplicatas(2)
        LibProc WL_EXIBIR
    End If
        'Pt. 95023 - Moacir Pfau(21/09/2009)
        If Index < 5 Then
            If lngOperacao = 0 Then
                lngOperacao = val(txtDuplicatas(40).Text)
            End If
            LibProc WL_EXIBIR
        End If
    
    Select Case Index
        
        'Percentual de Multa
        Case 37
            If CDbl(lblDuplDesc(4).Caption) > 0 And txtDuplicatas(Index).Text <> "" Then
                txtDuplicatas(36).Text = FormatNumber(CDbl(lblDuplDesc(4).Caption) * (CDbl(txtDuplicatas(37).Text) / 100), 2)
            Else
                txtDuplicatas(36).Text = "0,00"
            End If
            If txtDuplicatas(Index).Text <> "" Then
                txtDuplicatas(Index).Text = FormatNumber(txtDuplicatas(Index).Text, 2)
            Else
                txtDuplicatas(Index).Text = "0,00"
            End If
            
        'Valor da Multa
        Case 36
            If CDbl(lblDuplDesc(4).Caption) > 0 And txtDuplicatas(Index).Text <> "" Then
                txtDuplicatas(37).Text = FormatNumber(CDbl(txtDuplicatas(36).Text) * 100 / CDbl(lblDuplDesc(4).Caption), 2)
            Else
                txtDuplicatas(36).Text = "0,00"
            End If
            If txtDuplicatas(Index).Text <> "" Then
                txtDuplicatas(Index).Text = FormatNumber(txtDuplicatas(Index).Text, 2)
            Else
                txtDuplicatas(Index).Text = "0,00"
            End If
            
        Case 38
            If CDbl(lblDuplDesc(4).Caption) > 0 And txtDuplicatas(Index).Text <> "" Then
                txtPercMora.Text = FormatNumber(CDbl(txtDuplicatas(38).Text) * 100 / CDbl(lblDuplDesc(4).Caption), 2)
            Else
                txtPercMora.Text = "0,00"
            End If
            If txtDuplicatas(Index).Text <> "" Then
                txtDuplicatas(Index).Text = FormatNumber(txtDuplicatas(Index).Text, 2)
            Else
                txtDuplicatas(Index).Text = "0,00"
            End If
            
        Case 7
            If IsDate(txtDuplicatas(Index).Text) Then
                datLiberacao = CDate(txtDuplicatas(Index).Text)
                'pt. 88523 - Ivo Sousa (24/09/2008)
                If UCase(mstrPagRec) = "R" Then
                    datLiberacao = DateAdd("d", DiasLiberacao, datLiberacao)
                    If calendario.PermiteLancamento(datLiberacao, , False) = "A" Then
                        datLiberacao = datLiberacao + NumeroDiasUteisNaoUteis(datLiberacao, 0)
                    End If
                End If
                txtDuplicatas(9).Text = datLiberacao
            End If
        Case 8
            If IsDate(txtDuplicatas(Index).Text) Then
                datLiberacao = CDate(txtDuplicatas(Index).Text)
                'pt. 88523 - Ivo Sousa (24/09/2008)
                If UCase(mstrPagRec) = "R" Then
                    datLiberacao = DateAdd("d", DiasLiberacao, datLiberacao)
                    If calendario.PermiteLancamento(datLiberacao, , False) = "A" Then
                        datLiberacao = datLiberacao + NumeroDiasUteisNaoUteis(datLiberacao, 0)
                    End If
                End If
                txtDuplicatas(9).Text = datLiberacao
            End If
            Call SugestaoOperacaoContabilBaixa   'Opera��o cont�bil de baixa
        Case 10, 18
            'Pt. 88817 - Moacir Pfau(06/11/2009)
                dblPerDespesaFinanceira = 0
            If IsNumeric(txtDuplicatas(11).Text) Then
                If val(txtDuplicatas(18).Text) > 0 And txtDuplicatas(11).Text = 0 Then
                    dblPerDespesaFinanceira = GetFieldValue("per_despesa_financeira", "[Formas de Pagamento]", "C�digo=" & txtDuplicatas(18).Text)
                    txtDuplicatas(11).Text = Format(txtDuplicatas(11).Text + (txtDuplicatas(10).Text * dblPerDespesaFinanceira / 100), "#,#0.#0")
                End If
            End If
        Case 31
            If txtDuplicatas(Index).Text <> "" Then
                txtDuplicatas(Index).Text = FormatNumber(txtDuplicatas(Index).Text, 2)
            Else
                txtDuplicatas(Index).Text = "0,00"
            End If
    End Select
End Sub

'SUB: ExibeSoma
'Soma o valor original com os Acr�scimos e diminui os Abatimentos.
'Exibe o resultado no label do formul�rio.
Private Sub ExibeSoma()
    Dim curResult As Currency
  
    If Len(txtDuplicatas(10).Text) Then
        If IsNumeric(txtDuplicatas(10).Text) Then
            curResult = CCurDef(txtDuplicatas(10).Text, ZERO)
        End If
    End If
    If Len(txtDuplicatas(11).Text) Then
        If IsNumeric(txtDuplicatas(11).Text) Then
            curResult = curResult + CCurDef(txtDuplicatas(11).Text, ZERO)
        End If
    End If
    If Len(txtDuplicatas(12).Text) Then
        If IsNumeric(txtDuplicatas(12).Text) Then
            curResult = curResult - CCurDef(txtDuplicatas(12).Text, ZERO)
        End If
    End If
    lblDuplDesc(4).Caption = FormatNumber(curResult, 2)
End Sub

'SUB.......: CalcValor
'Objetivo..: Exibe a janela de c�lculo do valor da duplicata.
Private Sub CalcValor()
Dim cVlrOriginal As Currency
Dim cAumento As Currency

  ' Verifica se o usu�rio j� preencheu o Valor Original

  If (Not IsValid(txtDuplicatas(10).Text)) Then Exit Sub

  ' Verifica se a data de pagamento foi preenchida e se � diferente de zero

  If (IsEmptyDate(txtDuplicatas(8).Text)) Then
    MsgFunc ResolveResString(26, resUM, txtDuplicatas(8).Text)
  Else
    ' Verifica se a data de vencimento foi preenchida

    If (IsEmptyDate(txtDuplicatas(7).Text)) Then
      MsgFunc ResolveResString(26, resUM, txtDuplicatas(7).Text)
    Else
      ' Verifica se a data de pagamento � posterior a data de vencimento

      If (DateDiff(DD_DIA, txtDuplicatas(7).Text, txtDuplicatas(8).Text) > ZERO) Then

        ' Chama a fun��o que exibe a janela de c�lculo e aguarda

        cVlrOriginal = CMoeda(txtDuplicatas(10).Text)
        cAumento = CMoeda(txtDuplicatas(11).Text)
'        If (CValorFinal(cVlrOriginal, cAumento, _
'                        CDate(txtDuplicatas(7).Text), _
'                        CDate(txtDuplicatas(8).Text))) Then
'
'          ' Retornando o valor j� calculado
'
'          lblDuplDesc(4).Caption = Format$(cVlrOriginal, FMOEDA)
'          txtDuplicatas(11).Text = Format$(cAumento, FMOEDA)
'        End If

      End If
    End If
  End If

End Sub

'SUB.......: ChequeInfo
'Objetivo..: Exibe informa��es do cheque para o usu�rio
'Argumentos: [sFuncao]: O mesmo argumento sFuncao da fun��o LibProc
'            [nBco   ]: Opcional. C�digo do Banco.
'            [nChq   ]: Opcional. N�mero do Cheque.
Private Sub ChequeInfo(sFuncao As String, Optional nBco As Long, Optional nChq As Long)
Dim strCheque     As String
Dim lngCheque     As Long
Dim lngBanco      As Long
Dim cValor        As Currency

  If (mstrPagRec = "P") Then

    '// Somente se for pagamento

    Select Case (sFuncao)
      Case WL_NOVO: Call LimpaControles(mrstCheques, Me, TAG_CHEQUE, mlngCheques, True)
      Case WL_SALVAR
        If ((CBool(nBco)) And (CBool(nChq))) Then
          If ((nBco <> GetValue(mrstDuplicatas, "Banco", 0)) Or (nChq <> GetValue(mrstDuplicatas, "Cheque", 0))) Then
            If (ExisteCheque(nBco, nChq) = ZERO) Then
              DeleteAll "Cheque", wsprintf("Banco = %l AND Cheque = %l", nBco, nChq)
            End If
          Else
            'Caso contr�rio apenas chama a fun��o salva registro para
            'gravar eventuais altera��es nos campos Nominal e Hist�rico
            Call SalvaRegistro(mrstCheques, Me, TAG_CHEQUE, mlngCheques)
          End If
        End If
        'Verifica se o cheque atual existe na tabela de Cheques, se n�o
        'existir acrescenta-o.

        nBco = GetValue(mrstDuplicatas, "Banco", ZERO)
        nChq = GetValue(mrstDuplicatas, "Cheque", ZERO)

        If ((CBool(nBco)) And (CBool(nChq))) Then
          strCheque = wsprintf("FROM Cheque WHERE Banco = %l AND Cheque = %l", nBco, nChq)

          If (Recordcount(strCheque) = 0) Then
            strCheque = "INSERT INTO Cheque (Banco, Cheque, Nominal, Hist�rico) " & _
                        wsprintf("VALUES (%l, %l, \'%s\', \'%s\');", nBco, nChq, _
                                 txtCheque(0).Text, _
                                 txtCheque(1).Text)
            Call ExecuteSQL(strCheque)
          End If
        End If
        Call ChequeInfo("ExibeRegistro")

      Case WL_CANCELAR: Call CancelaEdicao(mrstCheques, Me, TAG_CHEQUE, mlngCheques)

      Case WL_DELETAR
        If (CBool(nBco) And CBool(nChq)) Then
          If (ExisteCheque(nBco, nChq) = ZERO) Then
            DeleteAll "Cheque", wsprintf("Banco = %l AND Cheque = %l", nBco, nChq)
          End If
        End If
        Call ChequeInfo("ExibeRegistro")

      Case Else
        Call SalvaRegistro(mrstCheques, Me, TAG_CHEQUE, mlngCheques)
        lngBanco = CLngDef(txtDuplicatas(13).Text)
        lngCheque = CLngDef(txtDuplicatas(16).Text)
        strCheque = wsprintf("SELECT * FROM Cheque WHERE " & _
                             "Banco = %l AND Cheque = %l", _
                             lngBanco, lngCheque)

        If (AbreRecordset(mrstCheques, strCheque) = WL_OK) Then
          Call ExibeRegistro(mrstCheques, Me, TAG_CHEQUE, mlngCheques)
        Else
          Call LimpaControles(mrstCheques, Me, TAG_CHEQUE, mlngCheques, True)
        End If

        If (sFuncao <> WL_SAIR) Then
          lvwLancamentos.ListItems.Clear    '// Limpa o conte�do atual do ListView

          If (IsVisibleRecord(mlngCheques)) Then '// Se h� um cheque vis�vel agora
            SetPtrWait Me
            
            If gTipoDB = Access Then
              wvsprintf strCheque, _
                        "SELECT FORMAT(Nota, \'000000\') & ' - ' & " & _
                        "FORMAT(Parcela, \'00\') AS Cod, Tipo, Empresa, " & _
                        "FORMAT(([Valor Original] + Acr�scimo - Abatimento), " & _
                        "\'###,###,###,##0.00\') AS Total FROM Duplicatas WHERE PagRec = " & _
                        "'P' AND Banco = %l AND Cheque = %l;", lngBanco, lngCheque
            Else
              wvsprintf strCheque, _
                        "SELECT (Nota +  ' - ' & " & _
                        "Parcela) AS Cod, Tipo, Empresa, " & _
                        "([Valor Original] + Acr�scimo - Abatimento) " & _
                        " AS Total FROM Duplicatas WHERE PagRec = " & _
                        "'P' AND Banco = %l AND Cheque = %l;", lngBanco, lngCheque
            End If

            Call ListViewAddItem(lvwLancamentos, strCheque, "duplicata")

            If gTipoDB = Access Then
              wvsprintf strCheque, _
                        "SELECT FORMAT(C�digo, \'000000\') AS Cod, " & _
                        "Tipo, Empresa, FORMAT(([Valor Original] + " & _
                        "Acr�scimo - Abatimento), \'###,###,###,##0.00\') AS Total " & _
                        "FROM Lan�amentos WHERE PagRec = 'P' AND " & _
                        "Banco = %l AND Cheque = %l;", lngBanco, lngCheque
            Else
              wvsprintf strCheque, _
                        "SELECT C�digo AS Cod, " & _
                        "Tipo, Empresa, ([Valor Original] + " & _
                        "Acr�scimo - Abatimento) AS Total " & _
                        "FROM Lan�amentos WHERE PagRec = 'P' AND " & _
                        "Banco = %l AND Cheque = %l;", lngBanco, lngCheque
            End If


            Call ListViewAddItem(lvwLancamentos, strCheque, "lancamento")

            If gTipoDB = Access Then
            
              wvsprintf strCheque, _
                        "SELECT FORMAT(T.C�digo, \'000000\') As Cod, " & _
                        "'Transfer�ncia', B.Nome, FORMAT(T.Valor, \'###,###,###,##0.00\') " & _
                        "FROM [Transf Banc�ria] AS T, Bancos As B WHERE " & _
                        "B.Banco = T.Origem AND T.Origem = %l AND T.Cheque = %l;", _
                        lngBanco, lngCheque
            Else
              wvsprintf strCheque, _
                        "SELECT T.C�digo As Cod, " & _
                        "'Transfer�ncia', B.Nome, T.Valor " & _
                        "FROM [Transf Banc�ria] AS T, Bancos As B WHERE " & _
                        "B.Banco = T.Origem AND T.Origem = %l AND T.Cheque = %l;", _
                        lngBanco, lngCheque
            End If

            Call ListViewAddItem(lvwLancamentos, strCheque, "transferencia")

            '// Calculando o valor do cheque para exibi��o na janela

            cValor = Soma("[Valor Original] + Acr�scimo - Abatimento", _
                          "Duplicatas", wsprintf("PagRec = 'P' AND Banco = %l AND Cheque = %l", _
                                                 lngBanco, lngCheque), ZERO)

            cValor = cValor + Soma("[Valor Original] + Acr�scimo - Abatimento", _
                                   "Lan�amentos", wsprintf("PagRec = 'P' AND Banco = %l AND Cheque = %l", _
                                                           lngBanco, lngCheque), ZERO)

            cValor = cValor + Soma("Valor", "Transf Banc�ria", _
                                   wsprintf("Banco = %l AND Cheque = %l", _
                                            lngBanco, lngCheque), ZERO)

            lblDuplDesc(9).Caption = Format$(cValor, FMOEDA)
            SetPtrDef Me
          Else
            lblDuplDesc(9).Caption = NUL
          End If
        End If
    End Select
  End If
End Sub

Private Sub lvwRateio_DblClick()
    DoEvents
    XMark mlngItem
End Sub

Private Sub lvwRateio_ItemClick(ByVal item As ComctlLib.ListItem)
    mlngItem = item.Index
End Sub

Private Sub lvwRateio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        XMark mlngItem
    End If
End Sub

'SUB.......: XMark
'Objetivo..: Marca com um X o �tem selecionado pelo usu�rio quando esta n�o est�
'            marcado, ou desmarca quando este estiver marcado.
'Argumento.: [lngIndice]: �ndico do �tem que deve ser marcado ou desmarcado.
Private Sub XMark(lngIndice As Long)
    If lngIndice > 0 Then
        If (lvwRateio.ListItems(lngIndice).SmallIcon = DL_MARCADO) Then
            lvwRateio.ListItems(lngIndice).SmallIcon = DL_DESMARCADO
        Else
            lvwRateio.ListItems(lngIndice).SmallIcon = DL_MARCADO
        End If
    End If
End Sub

Private Sub txtPercMora_GotFocus()
    Selecione txtPercMora
End Sub

Private Sub txtPercMora_LostFocus()
    If val(lblDuplDesc(4).Caption) > 0 And txtPercMora.Text <> "" Then
        txtDuplicatas(38).Text = FormatNumber(CDbl(lblDuplDesc(4).Caption) * CDbl(txtPercMora.Text / 100), 2)
    Else
        txtDuplicatas(38).Text = "0,00"
    End If
    If txtPercMora.Text <> "" Then
        txtPercMora.Text = FormatNumber(txtPercMora.Text, 2)
    Else
        txtPercMora.Text = "0,00"
    End If
End Sub

'Autor: Dulcino J�nior
'Data: 10/11/2006
'Fun��o utilizada para verificar se a duplicata pode ou n�o ser alterada por esta tela
'caso a duplicata sejam oriunda de Notas Fiscais ou Pedios(no caso de sinais de neg�cio)
'o sistema n�o deve permitir a altera��o do valor
Private Function permiteAlterarValor() As Boolean
    If UCase(mstrTabela) <> UCase("Lan�amentos") Then
        If UCase(mstrPagRec) = "P" Then
            'Caso a duplicata a pagar esteja relacionado com nota de entrada ou pedido de compra n�o alterar
            permiteAlterarValor = Not (permiteAlterarValorNFE Xor permiteAlterarValorPDC)
        Else
            'Caso a duplicata a pagar esteja relacionado com nota de sa�da ou pedido de venda n�o alterar
            permiteAlterarValor = Not (permiteAlterarValorNFS Xor permiteAlterarValorPDV)
        End If
    Else
        permiteAlterarValor = True
    End If
End Function

'Autor: Dulcino J�nior
'Data: 10/11/2006
'Fun��o utilizada para verificar se a duplicata esta ligada a uma nota fiscal de entrada
'caso esteja n�o ser� permitido a altera��o do valor da duplicata.
Private Function permiteAlterarValorNFE() As Boolean
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "[Notas Fiscais de Entrada]"
    Call cmd.Filter.Append("[Tipo de Registro] = @pTipoRegistro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", GetValue(mrstDuplicatas, "Tipo"), dbFieldTypeString, 20))
    Call cmd.Filter.Append("[N�mero] = @pNumero")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumero", GetValue(mrstDuplicatas, "Nota"), dbFieldTypeLong))
    Call cmd.Filter.Append("Fornecedor = @pFornec")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFornec", GetValue(mrstDuplicatas, "Empresa"), dbFieldTypeString, 15))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    Call cmd.Filter.Append("Empresa = @pEmpresa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", Left(DonaSistema, 15), dbFieldTypeString, 15))
    permiteAlterarValorNFE = rdResult.EOF
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

'Autor: Dulcino J�nior
'Data: 10/11/2006
'Fun��o utilizada para verficar se a duplicata esta ligada a uma nota fiscal de sa�da
'caso esteja n�o ser� permitido a altera��o do valor da duplicata.
Private Function permiteAlterarValorNFS() As Boolean
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "[Notas Fiscais de Sa�da]"
    Call cmd.Filter.Append("[Tipo de Registro] = @pTipoRegistro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", GetValue(mrstDuplicatas, "Tipo"), dbFieldTypeString, 20))
    Call cmd.Filter.Append("[N�mero] = @pNumero")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumero", GetValue(mrstDuplicatas, "Nota"), dbFieldTypeDouble))
    Call cmd.Filter.Append("Fornecedor = @pFornec")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFornec", Left(DonaSistema, 15), dbFieldTypeString, 15))
    Call cmd.Filter.Append("Empresa = @pEmpresa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", GetValue(mrstDuplicatas, "Empresa"), dbFieldTypeString, 15))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    permiteAlterarValorNFS = rdResult.EOF
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

'Autor: Dulcino J�nior
'Data: 10/11/2006
'Fun��o utilizada para verificar se a duplicata esta ligada a um pedido de venda
'caso esteja n�o ser� permitido a altera��o do valor da duplicata.
Private Function permiteAlterarValorPDV() As Boolean
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "[Pedidos de Venda]"
    Call cmd.Filter.Append("[Tipo de Registro] = @pTipoRegistro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", GetValue(mrstDuplicatas, "Tipo"), dbFieldTypeString, 20))
    Call cmd.Filter.Append("[N�mero] = @pNumero")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumero", GetValue(mrstDuplicatas, "Nota"), dbFieldTypeDouble))
    Call cmd.Filter.Append("Fornecedor = @pFornec")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFornec", Left(DonaSistema, 15), dbFieldTypeString, 15))
    Call cmd.Filter.Append("Empresa = @pEmpresa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", GetValue(mrstDuplicatas, "Empresa"), dbFieldTypeString, 15))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    permiteAlterarValorPDV = rdResult.EOF
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

'Autor: Dulcino J�nior
'Data: 10/11/2006
'Fun��o utilizada para verficar se a duplicata esta ligada a um pedido de compra
'caso esteja n�o ser� permitido a altera��o do valor da duplicata
Private Function permiteAlterarValorPDC() As Boolean
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "[Pedidos de Compra]"
    Call cmd.Filter.Append("[Tipo de Registro] = @pTipoRegistro")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipoRegistro", GetValue(mrstDuplicatas, "Tipo"), dbFieldTypeString, 20))
    Call cmd.Filter.Append("[N�mero] = @pNumero")
    Call cmd.Parameters.add(cmd.CreateParameter("@pNumero", GetValue(mrstDuplicatas, "Nota"), dbFieldTypeLong))
    Call cmd.Filter.Append("Fornecedor = @pFornec")
    Call cmd.Parameters.add(cmd.CreateParameter("@pFornec", GetValue(mrstDuplicatas, "Empresa"), dbFieldTypeString, 15))
    Call cmd.Filter.Append("Empresa = @pEmpresa")
    Call cmd.Parameters.add(cmd.CreateParameter("@pEmpresa", Left(DonaSistema, 15), dbFieldTypeString, 15))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    permiteAlterarValorPDC = rdResult.EOF
    rdResult.CloseReader
    Set rdResult = Nothing
    Set cmd = Nothing
    Aplicacao.Disconnect
End Function

'Data.......: 18/05/2007
'Autor......: Dulcino J�nior
'Descri��o..: Fun��o utilizada para a verfica��o do preenchimento dos campos
'               de opera��o cont�bil dos lan�amentos. ref pt 81902
'Retorno....: [Boolean] Retorna se o registro pode ou n�o ser gravado.
Private Function validaIntegracaoLancamentos() As Boolean
    validaIntegracaoLancamentos = True
    If Not IsEmptyDate(txtDuplicatas(8).Text) Then
        If Trim(txtDuplicatas(41).Text) = "0" Or Trim(txtDuplicatas(41).Text) = "" Then
            MsgBox "Para lan�amentos quitados � necess�rio informar a Opera��o de Baixa!", vbInformation, NomeModulo
            txtDuplicatas(41).SetFocus
            validaIntegracaoLancamentos = False
        Else
            If Not IsEmptyDate(txtDuplicatas(6).Text) Then
                If CDate(txtDuplicatas(6).Text) <> CDate(txtDuplicatas(8).Text) Then
                    If Trim(txtDuplicatas(40).Text) = "0" Or Trim(txtDuplicatas(40).Text) = "" Then
                        MsgBox "Para lan�amentos � necess�rio informar a Opera��o Cont�bil", vbInformation, NomeModulo
                        txtDuplicatas(40).SetFocus
                        validaIntegracaoLancamentos = False
                    End If
                Else
                    If Trim(txtDuplicatas(40).Text) <> "0" And Trim(txtDuplicatas(40).Text) <> "" Then
                        MsgBox "Para movimentos banc�rios a opera��o de emiss�o n�o pode ser informada.", vbInformation, NomeModulo
                        txtDuplicatas(40).Text = "0"
                        txtDuplicatas(40).SetFocus
                        validaIntegracaoLancamentos = False
                    End If
                End If
            End If
        End If
    Else
        If Trim(txtDuplicatas(40).Text) = "0" Or Trim(txtDuplicatas(40).Text) = "" Then
            'pt. 82355 Ivo Sousa (24/10/07).
            MsgBox "Para Lan�amentos � necess�rio informar a Opera��o Cont�bil", vbInformation, NomeModulo
            txtDuplicatas(40).SetFocus
            validaIntegracaoLancamentos = False
        Else
            If txtDuplicatas(41).Text <> "" And txtDuplicatas(41).Text <> "0" Then
                MsgBox "Para lan�amentos com opera��o cont�bil de baixa � necess�rio informar a data de pagamento.", vbInformation, NomeModulo
                txtDuplicatas(8).SetFocus
                validaIntegracaoLancamentos = False
            End If
        End If
    End If
End Function

'Data.......: 18/05/2007
'Autor......: Dulcino J�nior
'Descri��o..: Fun��o utilizada para a verfica��o do preenchimento dos campos
'               de opera��o cont�bil dos lan�amentos. ref pt 81902
'Retorno....: [Boolean] Retorna se o registro pode ou n�o ser gravado.
Private Function validaIntegracaoDuplicatas() As Boolean
    validaIntegracaoDuplicatas = True
    If Not IsEmptyDate(txtDuplicatas(8).Text) Then
        If Trim(txtDuplicatas(41).Text) = "0" Or Trim(txtDuplicatas(41).Text) = "" Then
            MsgBox "Para Duplicatas quitadas � necess�rio informar a Opera��o de Baixa!", vbInformation, "Valida��o de Campos"
            txtDuplicatas(41).SetFocus
            validaIntegracaoDuplicatas = False
        End If
    Else
        If txtDuplicatas(41).Text <> "" And txtDuplicatas(41).Text <> "0" Then
            'pt. 82355 Ivo Sousa (24/10/07)
            MsgBox "Para Duplicatas com opera��o cont�bil de baixa � necess�rio informar a data de pagamento.", vbInformation, NomeModulo
            txtDuplicatas(8).SetFocus
            validaIntegracaoDuplicatas = False
        End If
    End If
    If Not IsEmptyDate(txtDuplicatas(6).Text) Then
        If Trim(txtDuplicatas(40).Text) = "0" Or Trim(txtDuplicatas(40).Text) = "" Then
            MsgBox "Para Duplicatas � necess�rio informar a Opera��o Cont�bil", vbInformation, "Valida��o de Campos"
            txtDuplicatas(40).SetFocus
            validaIntegracaoDuplicatas = False
        End If
    End If
End Function

'Data.......: 23/05/2007
'Autor......: Dulcino J�nior
'Descri��o..: Fun��o utilizada para verificar se existem notas fiscais vinculadas
'               a duplicata, caso exista a mesma n�o pode ser excluida.
'Retorno....: [Boolean] Retorna se a duplicata pode ou n�o ser excluida.
Private Function PermiteExclusao(ByRef intParcOrigem As Integer) As Boolean
    
    PermiteExclusao = True
    'pt. 82831 - Ivo Sousa (23/02/2009)
    If Not BaixaParcial(intParcOrigem) Then
        PermiteExclusao = PermiteExclusao And (Not PertenceNota)
        PermiteExclusao = PermiteExclusao And ((Not PertencePedido))
        If Not PermiteExclusao Then
            MsgBox "N�o foi possivel excluir a duplicata pois a mesma pertence a uma Nota Fiscal ou Pedido.", vbInformation, "Valida��o de Campos"
        End If
    End If
    PermiteExclusao = PermiteExclusao And ValidaRateio
    
    'pt. 88289 - Ivo Sousa (19/12/2008)
    If PermiteExclusao And GerouPagFor Then
        MsgBox "N�o foi poss�vel excluir a duplicata pois a mesma j� foi enviada para o Banco.", vbInformation, "Valida��o de Campos"
        PermiteExclusao = False
    End If
End Function

'Data.......: 23/05/2007
'Autor......: Dulcino J�nior
'Descri��o..: Fun��o respons�vel por verificar se existe nota fiscal para essa
'               duplicata.
'Retorno....: [Boolean] Retorna se a duplicata possui nota fiscal vinculada.
Private Function PertenceNota() As Boolean
    Dim selCmd    As IDBSelectCommand
    Dim rdResult  As IDBReader
    Dim strTabela As String
    
On Error GoTo Error_Handler
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "N�mero"
        
        If mstrPagRec = "P" Then
            strTabela = "[Notas Fiscais de Entrada]"
        Else
            strTabela = "[Notas Fiscais de Sa�da]"
        End If
        
        .Table.TableName = strTabela
        
        Call .Filter.Append("N�mero = @pNumero")
        Call .Parameters.add(.CreateParameter("@pNumero", txtDuplicatas(1).Text, dbFieldTypeLong))
        
        Call .Filter.Append("[Tipo de Registro] = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", cboDuplicatas(3).Text, dbFieldTypeString, 30))
        
        If mstrPagRec = "P" Then
            Call .Filter.Append("Fornecedor = @pFornecedor")
            Call .Parameters.add(.CreateParameter("@pFornecedor", txtDuplicatas(2).Text, dbFieldTypeString, 15))
        Else
            Call .Filter.Append("Empresa = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", txtDuplicatas(2).Text, dbFieldTypeString, 15))
        End If
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    PertenceNota = Not rdResult.EOF
    rdResult.CloseReader
    Set selCmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
    
    Exit Function
Error_Handler:
    FinallyConnection Aplicacao
    err.Clear
    PertenceNota = False
End Function

'Data.......: 23/05/2007
'Autor......: Dulcino J�nior
'Descri��o..: Fun��o respons�vel por verificar se existe pedido para essa
'               duplicata.
'Retorno....: [Boolean] Retorna se a duplicata possui pedido vinculada.
Private Function PertencePedido() As Boolean
    Dim selCmd    As IDBSelectCommand
    Dim rdResult  As IDBReader
    Dim strTabela As String
    
On Error GoTo Error_Handler
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    With selCmd
        .SelectClause = "N�mero"
        
        If mstrPagRec = "P" Then
            strTabela = "[Pedidos de Compra]"
        Else
            strTabela = "[Pedidos de Venda]"
        End If
        
        .Table.TableName = strTabela
        
        Call .Filter.Append("N�mero = @pNumero")
        Call .Parameters.add(.CreateParameter("@pNumero", txtDuplicatas(1).Text, dbFieldTypeLong))
        
        Call .Filter.Append("[Tipo de Registro] = @pTipo")
        Call .Parameters.add(.CreateParameter("@pTipo", cboDuplicatas(3).Text, dbFieldTypeString, 30))
        
        If mstrPagRec = "P" Then
            Call .Filter.Append("Fornecedor = @pFornecedor")
            Call .Parameters.add(.CreateParameter("@pFornecedor", txtDuplicatas(2).Text, dbFieldTypeString, 15))
        Else
            Call .Filter.Append("Empresa = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", txtDuplicatas(2).Text, dbFieldTypeString, 15))
        End If
    End With
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    PertencePedido = Not rdResult.EOF
    rdResult.CloseReader
    Set selCmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
    
    Exit Function
Error_Handler:
    FinallyConnection Aplicacao
    err.Clear
    PertencePedido = False
End Function

'Data.......: 30/05/2007
'Autor......: Dulcino J�nior
'Descri��o..: Procedimento utilizado para sugerir a opera��o cont�bil
'               de acordo com o tipo global da duplicata ou lan�amento.
Public Sub SugestaoOperacaoContabilBaixa()
    Dim DAOMatriz   As cMatrizContabilizacaoDAO
    Dim matriz      As cMatrizContabilizacao
    Dim lngOperacao As Long
    
    If IsDate(txtDuplicatas(8).Text) Then
        Set DAOMatriz = New cMatrizContabilizacaoDAO
        Set matriz = DAOMatriz.Carregar(cboDuplicatas(3).Text)
        If Not matriz Is Nothing Then
            If mstrTabela = "Lan�amentos" Then
                If mstrPagRec = "P" Then
                    lngOperacao = matriz.BaixaLancamentosPagar
                Else
                    lngOperacao = matriz.baixaLancamentosReceber
                End If
            Else
                If mstrPagRec = "P" Then
                    lngOperacao = matriz.BaixaDuplicatasPagar
                Else
                    lngOperacao = matriz.BaixaDuplicatasReceber
                End If
            End If
        Else
            lngOperacao = 0
        End If
        Set matriz = Nothing
        Set DAOMatriz = Nothing
    Else
        lngOperacao = 0
    End If
    txtDuplicatas(41).Text = lngOperacao
End Sub

'Data.......: 18/04/2007
'Autor......: Dulcino J�nior
'Descri��o..: Fun��o utilizada para retornar a quantidade de dias
'               que o banco possui para a libera��o da duplicata.
'Retorno....: [Integer] N�mero de dias para libera��o do pagamento.
Private Function DiasLiberacao() As Double
    Dim selCmd   As IDBSelectCommand
    Dim rdResult As IDBReader
    
    If IsNumeric(txtDuplicatas(13).Text) Then
        If CLng(txtDuplicatas(13).Text) > 0 Then
            Aplicacao.Connect
            Set selCmd = Aplicacao.CreateSelectCommand
            With selCmd
                .SelectClause = "[Dias para Libera��o]"
                
                .Table.TableName = "Bancos"
                
                Call .Filter.Append("Banco = @pNumero")
                Call .Parameters.add(.CreateParameter("@pNumero", CLng(txtDuplicatas(13).Text), dbFieldTypeLong))
            End With
            Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
            If Not rdResult.EOF Then
                DiasLiberacao = rdResult.GetDouble("Dias para Libera��o")
            End If
            rdResult.CloseReader
            Set rdResult = Nothing
            Set selCmd = Nothing
            Aplicacao.Disconnect
        Else
            DiasLiberacao = 0
        End If
    Else
        DiasLiberacao = 0
    End If
End Function

'Data.......: 26/03/2008
'Autor......: Ivo Sousa(pt. 86132)
'Descri��o..: Fun��o utilizada para Valida��o de datas como feriados, domingos
'             sabados ou se o periodo esta bloqueado.
'Retorno....: [boolean] Se a data � valida
Private Function ValidaDatas() As Boolean
    Dim intIndexTXT As Integer
    Dim intIndexLBL As Integer
    Dim strSinal As String
    
    intIndexTXT = 6
    intIndexLBL = 14
    While intIndexTXT <= 9
        If txtDuplicatas(intIndexTXT).Text <> "" Then
            strSinal = calendario.PermiteLancamento(txtDuplicatas(intIndexTXT).Text)
            If strSinal = "X" Then
                MsgBox "O movimento esta bloqueado para a data Informada no campo " & Replace(lblDuplicatas(intIndexLBL).Caption, ":", ""), vbOKOnly, NomeModulo
                txtDuplicatas(intIndexTXT).SetFocus
                ValidaDatas = False
                Exit Function
            ElseIf strSinal = "A" Then
                ValidaDatas = True
            ElseIf strSinal = "F" Or strSinal = "S" Or strSinal = "D" Then
                If MsgBox("A data Informada no campo " & Replace(lblDuplicatas(intIndexLBL).Caption, ":", "") & " n�o � um dia �til." & vbNewLine & _
                "Deseja salvar a duplicata assim mesmo?", vbYesNo + vbInformation, NomeModulo) = vbYes Then
                    ValidaDatas = True
                Else
                    txtDuplicatas(intIndexTXT).SetFocus
                    ValidaDatas = False
                    Exit Function
                End If
            End If
        End If
        intIndexTXT = intIndexTXT + 1
        intIndexLBL = intIndexLBL + 1
    Wend
End Function

'Procedure..: CarregaPadrao
'Data.......: 10/04/2008
'Autor......: MOACIR PFAU
'Data.......: 10/04/2008
'Descri��o..: Utilizado para carregar os campos na tela, metodo Paliativo.
'Protocolo..: 86140
Private Function CarregaPadrao()
    cboDuplicatas.item(3).Text = "Fatura"               'Tipo
    txtDuplicatas(4).Text = "0"                         'Parcela
    txtDuplicatas(18).Text = "0"                        'Forma Pagto
    txtDuplicatas(13).Text = "0"                        'Banco
    txtDuplicatas(14).Text = "0"                        'Conta
    txtDuplicatas(40).Text = "0"                        'Op. Cantabil
    cboDuplicatas.item(20).Text = "Normal"              'Situa��o
    txtDuplicatas(10).Text = "0"                        'Valor Original
    txtDuplicatas(11).Text = "0"                        'Acr�scimo
    txtDuplicatas(12).Text = "0"                        'Abatimento
    txtDuplicatas(6).Text = Format(Date, "DD/MM/YYYY")  'Emiss�o
    txtDuplicatas(7).Text = Format(Date, "DD/MM/YYYY")  'Vencimento
    txtDuplicatas(9).Text = Format(Date, "DD/MM/YYYY")  'Libera��o
    txtDuplicatas(41).Text = "0"                        'Libera��o
    lblDuplDesc(4).Caption = "0,00"                     'Total
End Function

'pt. 85684 - Moacir Pfau(01/07/2008)
'Verifica se a duplicata foi gerada dentro da tela de gera��o de t�tulos. se sim, n�o pode ser excluida por esta tela.
Private Function fValidaExclusao() As Boolean
    Dim strSql As String
    Dim rstTab As Object
    
    strSql = ""
    If txtDuplicatas(0).Text = "P" Then
        strSql = "SELECT [cd_titulo], [nota], [tipo_registro], [empresa], [pagRec], [Parcela] "
        strSql = strSql & "FROM FVFTituloPagarDuplicata "
        strSql = strSql & "WHERE [nota]=" & txtDuplicatas(1).Text & " AND [tipo_registro]='" & cboDuplicatas(3).Text & "' AND [empresa]='" & txtDuplicatas(2).Text & "' AND [pagRec]='P' AND [Parcela]=" & txtDuplicatas(4).Text
    ElseIf txtDuplicatas(0).Text = "R" Then
        strSql = "SELECT [cd_titulo], [nota], [tipo_registro], [empresa], [pagRec], [Parcela] "
        strSql = strSql & "FROM FVFTituloReceberDuplicata "
        strSql = strSql & "WHERE [nota]=" & txtDuplicatas(1).Text & " AND [tipo_registro]='" & cboDuplicatas(3).Text & "' AND [empresa]='" & txtDuplicatas(2).Text & "' AND [pagRec]='R' AND [Parcela]=" & txtDuplicatas(4).Text
    End If
    If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
        fValidaExclusao = False
        Exit Function
    End If
    fValidaExclusao = True
End Function

'Data.......: 23/05/2007
'Autor......: Dulcino J�nior
'Descri��o..: A fun��o verifica a exist�ncia de rateio para o lan�amento que est� carregado na tela,
'               conforme os registros da tabela FFIRateioLancamento, se existe registro nessa tabela
'               ser� buscado o registro que originou o rateio e fazer a verifica��o se o mesmo est�
'               quitado, caso n�o esteja, ser� permitida a exclus�o do mesmo, e o valor referente ao
'               titulo excluido ser� retornado para o titulo que originou o rateio, do contr�rio o
'               sistema vai avisar ao usu�rio qual o titulo que originou o rateio e dizer que o
'               mesmo est� quitado.
'Retorno....: [Boolean] Retorna se a duplicata pode ou n�o ser excluida.
Private Function ValidaRateio() As Boolean
    Dim strSql     As String
    Dim rstResult  As Object

    ValidaRateio = True
    If chkRateio.value = vbChecked Then
        strSql = "SELECT pag_rec_origem, nr_nota_origem, cd_empresa_origem, nr_parcela_origem, tp_registro_origem FROM FFIRateioDuplicata"
        strSql = strSql & " WHERE pag_rec_destino='" & mstrPagRec & "' AND nr_nota_destino=" & txtDuplicatas(1).Text
        strSql = strSql & " AND nr_parcela_destino=" & txtDuplicatas(4).Text & " AND cd_empresa_destino='" & txtDuplicatas(2).Text & "'"
        strSql = strSql & " AND tp_registro_destino='" & cboDuplicatas(3).Text & "'"
        If AbreRecordset(rstResult, strSql) = WL_OK Then
            strSql = "SELECT Pagamento FROM Duplicatas WHERE PagRec='" & rstResult.Fields("pag_rec_origem").value
            strSql = strSql & "' AND Nota=" & rstResult.Fields("nr_nota_origem").value & " AND Parcela="
            strSql = strSql & rstResult.Fields("nr_parcela_origem").value & " AND Empresa='" & rstResult.Fields("cd_empresa_origem").value & "'"
            strSql = strSql & " AND Tipo='" & rstResult.Fields("tp_registro_origem").value & "'"
            mlngCodigo = rstResult.Fields("nr_nota_origem").value
            mlngPARCELA = rstResult.Fields("nr_parcela_origem").value
            mstrEmpresa = rstResult.Fields("cd_empresa_origem").value
            mstrTipoRegistro = rstResult.Fields("tp_registro_origem").value
        Else
            strSql = ""
        End If
        Call FechaRecordset(rstResult)
        If strSql <> "" Then
            If AbreRecordset(rstResult, strSql) = WL_OK Then
                If IsEmptyDate(rstResult.Fields("Pagamento").value) Then
                    mstrOrigem = "UPDATE Duplicatas SET [Valor Original]=[Valor Original]+" & Replace(txtDuplicatas(10).Text, ",", ".")
                    mstrOrigem = mstrOrigem & " WHERE PagRec='" & mstrPagRec & "' AND Nota=" & mlngCodigo & " AND "
                    mstrOrigem = mstrOrigem & "Parcela=" & mlngPARCELA & " AND Empresa='" & mstrEmpresa & "' AND Tipo='" & mstrTipoRegistro & "'"
                    
                    mstrDelete = "DELETE FROM FFIRateioDuplicata WHERE pag_rec_destino='" & mstrPagRec & "' AND "
                    mstrDelete = mstrDelete & "nr_nota_destino=" & txtDuplicatas(1).Text & " AND "
                    mstrDelete = mstrDelete & "nr_parcela_destino=" & txtDuplicatas(4).Text & " AND cd_empresa_destino='"
                    mstrDelete = mstrDelete & mstrEmpresa & "' AND tp_registro_destino='" & cboDuplicatas(3).Text & "'"
                    
                    mstrRateio = "SELECT nr_nota_destino FROM FFIRateioDuplicata WHERE pag_rec_origem='" & mstrPagRec & "' AND nr_nota_origem="
                    mstrRateio = mstrRateio & mlngCodigo & " AND nr_parcela_origem=" & mlngPARCELA & " AND cd_empresa_origem='" & mstrEmpresa & "' AND "
                    mstrRateio = mstrRateio & "tp_registro_origem='" & mstrTipoRegistro & "'"
                    ValidaRateio = True
                Else
                    MsgBox "N�o � possivel excluir a parcela por que a parcela de origem do rateio j� est� quitada.", vbInformation, NomeModulo
                    mstrOrigem = ""
                    mstrDelete = ""
                    ValidaRateio = False
                End If
            End If
            Call FechaRecordset(rstResult)
        Else
            mstrOrigem = ""
            mstrDelete = ""
            ValidaRateio = True
        End If
    End If
End Function

'Data.......: 19/12/2008
'Autor......: Ivo Sousa
'Descri��o..: Utilizado para validar se o documento j� foi enviado para o banco.
'Protocolo..: 88289
Private Function GerouPagFor() As Boolean
    If GetFieldValue("cd_arquivoPagamento", "FFIItemPagamento", "tp_documento = 'Dup' AND tipo_registro = '" & cboDuplicatas(3).Text & "' AND nr_documento = " & txtDuplicatas(1).Text & " AND nr_parcela = " & txtDuplicatas(4).Text & " AND cd_empresa = '" & txtDuplicatas(2).Text & "'", , 0) > 0 Then
        GerouPagFor = True
    Else
        GerouPagFor = False
    End If
End Function

'Data.......: 23/02/2009
'Autor......: Ivo Sousa
'Descri��o..: Utilizado para validar se o documento � uma baixa parcial
'Protocolo..: 88289
Private Function BaixaParcial(ByRef intParcOrigem As Integer) As Boolean
    intParcOrigem = GetFieldValue("parc_origem_baixa", "Duplicatas", "PagRec='" & mstrPagRec & "' AND Nota=" & txtDuplicatas(1).Text & " AND Parcela=" & txtDuplicatas(4).Text & " AND Empresa='" & txtDuplicatas(2).Text & "' AND Tipo='" & cboDuplicatas(3).Text & "'")
    If intParcOrigem > 0 Then
        BaixaParcial = True
    Else
        BaixaParcial = False
    End If
End Function

'Pt. 114146 - Moacir Pfau(29/02/2012)
Private Sub TotalizaValorRateio()
    If Not EAddNew(mlngDuplicatas) Then
        txtDuplicatas(45).Text = Format(txtDuplicatas(10).Text - SomaValores(), FMOEDA)
    End If
End Sub

