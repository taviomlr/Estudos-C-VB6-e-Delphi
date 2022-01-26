VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmLancamentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamentos"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   780
   ClientWidth     =   11265
   Icon            =   "Lancamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   11265
   Tag             =   "Duplicatas"
   Begin TabDlg.SSTab SSTab1 
      Height          =   6330
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   11165
      _Version        =   393216
      Tab             =   2
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
      TabPicture(0)   =   "Lancamento.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDuplicatas(30)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraRateio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraDuplicatas(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FraDuplicatas(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FraDuplicatas(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FraDuplicatas(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDuplicatas(26)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDuplicatas(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Adicionais"
      TabPicture(1)   =   "Lancamento.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraDuplicatas(10)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FraDuplicatas(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FraDuplicatas(4)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "FraDuplicatas(5)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "FraDuplicatas(7)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Outros"
      TabPicture(2)   =   "Lancamento.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblDuplicatas(44)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblDuplDesc(16)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblDuplicatas(46)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblDuplicatas(45)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtDuplicatas(39)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "FraDuplicatas(8)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtDuplicatas(43)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtDuplicatas(42)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.Frame Frame 
         Caption         =   "Endereço de Cobrança"
         Height          =   1635
         Left            =   60
         TabIndex        =   146
         Top             =   3600
         Width           =   11145
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "cd_cobranca"
            Height          =   315
            Index           =   44
            Left            =   1230
            MaxLength       =   3
            TabIndex        =   41
            Tag             =   "Duplicatas"
            Top             =   360
            Width           =   705
         End
         Begin Fox.EBSText etxCobrancaEndereco 
            Height          =   330
            Left            =   1230
            TabIndex        =   147
            Top             =   720
            Width           =   9840
            _extentx        =   10610
            _extenty        =   582
            font            =   "Lancamento.frx":0060
            tipo            =   4
            tipotexto       =   0
            maxlength       =   70
            locked          =   -1
            exibedescricao  =   0
         End
         Begin Fox.EBSText etxCobrancaBairro 
            Height          =   330
            Left            =   735
            TabIndex        =   148
            Top             =   1080
            Width           =   4320
            _extentx        =   298926
            _extenty        =   582
            font            =   "Lancamento.frx":008C
            tipo            =   4
            tipotexto       =   0
            maxlength       =   20
            caption         =   "Bairro"
            locked          =   -1
            exibedescricao  =   0
         End
         Begin Fox.EBSText etxCobrancaCep 
            Height          =   330
            Left            =   2715
            TabIndex        =   149
            Top             =   360
            Width           =   1620
            _extentx        =   56198
            _extenty        =   582
            font            =   "Lancamento.frx":00B8
            tipo            =   4
            tipotexto       =   0
            maxlength       =   9
            caption         =   "CEP"
            locked          =   -1
            exibedescricao  =   0
         End
         Begin Fox.EBSText etxCobrancaCidade 
            Height          =   330
            Left            =   5250
            TabIndex        =   150
            Top             =   360
            Width           =   2835
            _extentx        =   265
            _extenty        =   582
            font            =   "Lancamento.frx":00E4
            tipo            =   4
            tipotexto       =   0
            maxlength       =   30
            locked          =   -1
            exibedescricao  =   0
         End
         Begin Fox.EBSText etxCobrancaEstado 
            Height          =   330
            Left            =   8490
            TabIndex        =   151
            Top             =   360
            Width           =   525
            _extentx        =   265
            _extenty        =   582
            font            =   "Lancamento.frx":0110
            tipo            =   4
            tipotexto       =   0
            maxlength       =   30
            locked          =   -1
            exibedescricao  =   0
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   645
            TabIndex        =   155
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Left            =   4695
            TabIndex        =   154
            Top             =   405
            Width           =   495
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Endereço"
            Height          =   195
            Left            =   450
            TabIndex        =   153
            Top             =   750
            Width           =   690
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   8220
            TabIndex        =   152
            Top             =   420
            Width           =   210
         End
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "NOSNUM"
         Height          =   315
         Index           =   42
         Left            =   1275
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Duplicatas"
         Top             =   2640
         Width           =   5085
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "id_carteira"
         Height          =   315
         Index           =   43
         Left            =   1275
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   40
         Tag             =   "Duplicatas"
         Top             =   3030
         Width           =   1245
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "Alteração"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   -65010
         MaxLength       =   18
         TabIndex        =   139
         Tag             =   "Duplicatas"
         Top             =   5835
         Width           =   1215
      End
      Begin VB.TextBox txtDuplicatas 
         DataField       =   "Usuário"
         Enabled         =   0   'False
         Height          =   315
         Index           =   26
         Left            =   -66450
         MaxLength       =   18
         TabIndex        =   138
         Tag             =   "Duplicatas"
         Top             =   5835
         Width           =   1215
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
         Left            =   -74940
         TabIndex        =   128
         Top             =   390
         Width           =   7455
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
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Descrição"
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
         Begin VB.ComboBox cboDuplicatas 
            DataField       =   "Tipo"
            Height          =   315
            Index           =   3
            ItemData        =   "Lancamento.frx":013C
            Left            =   1080
            List            =   "Lancamento.frx":013E
            TabIndex        =   1
            Tag             =   "Duplicatas"
            Text            =   "cboDuplicatas"
            Top             =   600
            Width           =   1815
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
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "PagRec"
            Height          =   315
            Index           =   0
            Left            =   4230
            MaxLength       =   1
            TabIndex        =   130
            Tag             =   "Duplicatas"
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton btnEfetiva 
            Caption         =   "Efetivar Lancto."
            Height          =   405
            Left            =   5820
            TabIndex        =   129
            Top             =   180
            Width           =   1575
         End
         Begin VB.TextBox txtDuplicatas 
            Alignment       =   1  'Right Justify
            DataField       =   "SeqNossoNumero"
            Enabled         =   0   'False
            Height          =   315
            Index           =   28
            Left            =   6000
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "Duplicatas"
            Top             =   630
            Width           =   1305
         End
         Begin VB.Label lblDuplDesc 
            AutoSize        =   -1  'True
            Caption         =   "lblDuplDesc(0)"
            Height          =   195
            Index           =   0
            Left            =   2760
            TabIndex        =   137
            Tag             =   "Descrição"
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nota:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   630
            TabIndex        =   136
            Top             =   270
            Width           =   390
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   270
            TabIndex        =   135
            Top             =   1350
            Width           =   765
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Parcela:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   3570
            TabIndex        =   134
            Top             =   660
            Width           =   585
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   660
            TabIndex        =   133
            Top             =   630
            Width           =   360
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   132
            Top             =   990
            Width           =   660
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nr Sequencial:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   4920
            TabIndex        =   131
            Top             =   660
            Width           =   1050
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
         Left            =   -67380
         TabIndex        =   120
         Top             =   390
         Width           =   3615
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
            TabIndex        =   18
            Tag             =   "Duplicatas"
            Top             =   1380
            Width           =   2295
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Acréscimo"
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
            TabIndex        =   17
            Tag             =   "Duplicatas"
            Top             =   1050
            Width           =   2295
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
            TabIndex        =   16
            Tag             =   "Duplicatas"
            Top             =   660
            Width           =   2295
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Moeda"
            Height          =   315
            Index           =   17
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Duplicatas"
            Top             =   300
            Width           =   1095
         End
         Begin VB.Line hline 
            BorderColor     =   &H80000010&
            BorderWidth     =   2
            Index           =   0
            X1              =   120
            X2              =   3480
            Y1              =   1770
            Y2              =   1770
         End
         Begin VB.Label lblDuplDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   4
            Left            =   1200
            TabIndex        =   127
            Tag             =   "Descrição"
            Top             =   2100
            Width           =   2295
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   720
            TabIndex        =   126
            Top             =   2160
            Width           =   405
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Abatimento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   300
            TabIndex        =   125
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Acréscimo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   360
            TabIndex        =   124
            Top             =   1050
            Width           =   780
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor Original:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   150
            TabIndex        =   123
            Top             =   690
            Width           =   975
         End
         Begin VB.Line hline 
            BorderColor     =   &H80000014&
            Index           =   1
            X1              =   120
            X2              =   3480
            Y1              =   1770
            Y2              =   1770
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Moeda:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   570
            TabIndex        =   122
            Top             =   330
            Width           =   540
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
            TabIndex        =   121
            Top             =   1830
            Width           =   1530
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
         Left            =   -67380
         TabIndex        =   111
         Top             =   3000
         Width           =   3615
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Liberação"
            Height          =   315
            Index           =   9
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   22
            Tag             =   "Duplicatas"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Pagamento"
            Height          =   315
            Index           =   8
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   21
            Tag             =   "Duplicatas"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Vencimento"
            Height          =   315
            Index           =   7
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   20
            Tag             =   "Duplicatas"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Emissão"
            Height          =   315
            Index           =   6
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   19
            Tag             =   "Duplicatas"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(8)"
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   119
            Tag             =   "Descrição"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(7)"
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   118
            Tag             =   "Descrição"
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(6)"
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   117
            Tag             =   "Descrição"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(5)"
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   116
            Tag             =   "Descrição"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Liberação:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   300
            TabIndex        =   115
            Top             =   1350
            Width           =   750
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Pagamento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   210
            TabIndex        =   114
            Top             =   990
            Width           =   855
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   180
            TabIndex        =   113
            Top             =   630
            Width           =   885
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   420
            TabIndex        =   112
            Top             =   270
            Width           =   630
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
         TabIndex        =   107
         Top             =   4350
         Width           =   6255
         Begin VB.Label lblDadosAdcionais 
            Caption         =   "Vendedor:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   110
            Tag             =   "Descrição"
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   5955
         End
         Begin VB.Label lblDadosAdcionais 
            Caption         =   "Estado:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   109
            Tag             =   "Descrição"
            Top             =   480
            UseMnemonic     =   0   'False
            Width           =   5955
         End
         Begin VB.Label lblDadosAdcionais 
            Caption         =   "Cidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   108
            Tag             =   "Descrição"
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   5955
         End
      End
      Begin VB.Frame FraDuplicatas 
         Caption         =   "&Informações do Cheque"
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
         TabIndex        =   100
         Top             =   390
         Width           =   4785
         Begin VB.CommandButton cmdNominalRazaoSocial 
            Caption         =   "..."
            Height          =   300
            Left            =   4440
            TabIndex        =   101
            ToolTipText     =   "Cheque Nominal a Empresa do Lançamento/Duplicata"
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtCheque 
            DataField       =   "Histórico"
            Height          =   1275
            Index           =   1
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Tag             =   "Cheques"
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtCheque 
            DataField       =   "Nominal"
            Height          =   315
            Index           =   0
            Left            =   960
            MaxLength       =   60
            TabIndex        =   31
            Tag             =   "Cheques"
            Top             =   360
            Width           =   3435
         End
         Begin ComctlLib.ListView lvwLancamentos 
            Height          =   2205
            Left            =   120
            TabIndex        =   33
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
         Begin ComctlLib.ImageList imgDupl 
            Left            =   120
            Top             =   1080
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
         Begin VB.Label lblDuplDesc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblDuplDesc(9)"
            Height          =   315
            Index           =   9
            Left            =   960
            TabIndex        =   106
            Top             =   2040
            Width           =   3735
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   510
            TabIndex        =   105
            Top             =   2070
            Width           =   405
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Histórico:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   104
            Top             =   720
            Width           =   660
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nominal:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   300
            TabIndex        =   103
            Top             =   390
            Width           =   615
         End
         Begin VB.Label lblDuplicatas 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   240
            TabIndex        =   102
            Top             =   2400
            Width           =   1140
         End
      End
      Begin VB.Frame FraDuplicatas 
         Caption         =   "Observações"
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
         TabIndex        =   99
         Top             =   1140
         Width           =   6255
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Obs"
            Height          =   975
            Index           =   23
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Tag             =   "Duplicatas"
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.Frame FraDuplicatas 
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
         ForeColor       =   &H00000000&
         Height          =   735
         Index           =   6
         Left            =   -74940
         TabIndex        =   98
         Top             =   390
         Width           =   6255
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "LINDIG"
            Height          =   315
            Index           =   22
            Left            =   120
            TabIndex        =   24
            Tag             =   "Duplicatas"
            Top             =   240
            Width           =   6015
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
         Left            =   60
         TabIndex        =   93
         Top             =   390
         Width           =   11175
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CheBan"
            Height          =   315
            Index           =   32
            Left            =   1230
            MaxLength       =   9
            TabIndex        =   34
            Tag             =   "Duplicatas"
            Top             =   210
            Width           =   1245
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CheAge"
            Height          =   315
            Index           =   33
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   35
            Tag             =   "Duplicatas"
            Top             =   570
            Width           =   1245
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CheCco"
            Height          =   315
            Index           =   34
            Left            =   1230
            MaxLength       =   20
            TabIndex        =   36
            Tag             =   "Duplicatas"
            Top             =   960
            Width           =   2085
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "CheEmi"
            Height          =   315
            Index           =   35
            Left            =   1230
            MaxLength       =   60
            TabIndex        =   37
            Tag             =   "Duplicatas"
            Top             =   1290
            Width           =   5055
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   660
            TabIndex        =   97
            Top             =   240
            Width           =   510
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Agência:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   540
            TabIndex        =   96
            Top             =   630
            Width           =   630
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta Corrente:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   90
            TabIndex        =   95
            Top             =   990
            Width           =   1110
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Correntista:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   39
            Left            =   345
            TabIndex        =   94
            Top             =   1350
            Width           =   795
         End
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
         TabIndex        =   87
         Top             =   2520
         Width           =   6255
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "VlrDsP"
            Height          =   315
            Index           =   31
            Left            =   1950
            MaxLength       =   9
            TabIndex        =   28
            Tag             =   "Duplicatas"
            Top             =   1260
            Width           =   1365
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "VlrMul"
            Height          =   315
            Index           =   36
            Left            =   4590
            MaxLength       =   9
            TabIndex        =   29
            Tag             =   "Duplicatas"
            Top             =   510
            Width           =   1365
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "PerMul"
            Height          =   315
            Index           =   37
            Left            =   1950
            MaxLength       =   9
            TabIndex        =   26
            Tag             =   "Duplicatas"
            Top             =   510
            Width           =   1365
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "VlrMrd"
            Height          =   315
            Index           =   38
            Left            =   4590
            MaxLength       =   9
            TabIndex        =   30
            Tag             =   "Duplicatas"
            Top             =   870
            Width           =   1365
         End
         Begin VB.TextBox txtPercMora 
            DataField       =   "PerMrd"
            Height          =   315
            Left            =   1950
            MaxLength       =   9
            TabIndex        =   27
            Tag             =   "Duplicatas"
            Top             =   870
            Width           =   1365
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vlr. Desc. Pontualidade:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   120
            TabIndex        =   92
            Top             =   1290
            Width           =   1710
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vlr. Multa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   40
            Left            =   3795
            TabIndex        =   91
            Top             =   540
            Width           =   705
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Perc. Multa:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   41
            Left            =   1020
            TabIndex        =   90
            Top             =   540
            Width           =   855
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vlr. Mora Diária:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   42
            Left            =   3375
            TabIndex        =   89
            Top             =   900
            Width           =   1125
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Perc. Mora:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   43
            Left            =   1050
            TabIndex        =   88
            Top             =   900
            Width           =   825
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   38
         Tag             =   "Duplicatas"
         Top             =   2190
         Width           =   9855
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
         Left            =   -67395
         TabIndex        =   84
         Top             =   4815
         Width           =   3615
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "cd_operacao_baixa"
            Height          =   330
            Index           =   41
            Left            =   1080
            TabIndex        =   23
            Tag             =   "Duplicatas"
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Op. Contábil:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   135
            TabIndex        =   86
            Top             =   405
            Width           =   915
         End
         Begin VB.Label lblDuplDesc 
            AutoSize        =   -1  'True
            Caption         =   "lblDuplDesc(15)"
            Height          =   195
            Index           =   15
            Left            =   1890
            TabIndex        =   85
            Top             =   405
            Width           =   1125
         End
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
         Left            =   -74940
         TabIndex        =   63
         Top             =   2220
         Width           =   7455
         Begin VB.CheckBox chkRateio 
            Caption         =   "Identifica se o lançamento faz parte do rateio\Campo Invisível"
            DataField       =   "proveniente_rateio"
            Height          =   195
            Left            =   60
            TabIndex        =   142
            Tag             =   "Duplicatas"
            Top             =   3660
            Visible         =   0   'False
            Width           =   4695
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
         Begin VB.ComboBox cboDuplicatas 
            DataField       =   "Situação"
            Height          =   315
            Index           =   20
            ItemData        =   "Lancamento.frx":0140
            Left            =   1170
            List            =   "Lancamento.frx":0142
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Tag             =   "Duplicatas"
            Top             =   2055
            Width           =   1815
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Controle"
            Height          =   315
            Index           =   19
            Left            =   1170
            MaxLength       =   15
            TabIndex        =   13
            Tag             =   "Duplicatas"
            Top             =   2775
            Width           =   2775
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Cheque"
            Height          =   315
            Index           =   16
            Left            =   1170
            MaxLength       =   6
            TabIndex        =   12
            Tag             =   "Duplicatas"
            Top             =   2415
            Width           =   1575
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
         Begin VB.CommandButton cmdAbreRateio 
            Caption         =   "&Rateio..."
            Height          =   255
            Left            =   4050
            TabIndex        =   68
            Top             =   2775
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdProximoCheque 
            Caption         =   "..."
            Height          =   315
            Left            =   2730
            TabIndex        =   67
            ToolTipText     =   "Trazer Próximo Número do Cheque"
            Top             =   2415
            Width           =   255
         End
         Begin VB.CheckBox chkConciliado 
            Caption         =   "Conciliado"
            DataField       =   "Conciliado"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1170
            TabIndex        =   14
            Tag             =   "Duplicatas"
            Top             =   3165
            Width           =   1035
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Carteira"
            Height          =   315
            Index           =   27
            Left            =   3960
            MaxLength       =   3
            TabIndex        =   66
            Tag             =   "Duplicatas"
            Top             =   3135
            Width           =   1095
         End
         Begin VB.TextBox txtBancoCheque 
            DataField       =   "Banco"
            Height          =   285
            Left            =   6705
            TabIndex        =   65
            Tag             =   "Cheques"
            Top             =   3150
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtChequeCheque 
            DataField       =   "Cheque"
            Height          =   285
            Left            =   5940
            TabIndex        =   64
            Tag             =   "Cheques"
            Top             =   3150
            Visible         =   0   'False
            Width           =   735
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
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(13)"
            Height          =   255
            Index           =   13
            Left            =   2490
            TabIndex        =   83
            Top             =   270
            Width           =   4875
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(3)"
            Height          =   195
            Index           =   3
            Left            =   2490
            TabIndex        =   82
            Tag             =   "Descrição"
            Top             =   1290
            Width           =   4875
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(2)"
            Height          =   195
            Index           =   2
            Left            =   2490
            TabIndex        =   81
            Tag             =   "Descrição"
            Top             =   930
            Width           =   4875
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(1)"
            Height          =   195
            Index           =   1
            Left            =   2490
            TabIndex        =   80
            Tag             =   "Descrição"
            Top             =   570
            UseMnemonic     =   0   'False
            Width           =   4875
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Situação:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   450
            TabIndex        =   79
            Top             =   2085
            Width           =   675
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Controle:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   510
            TabIndex        =   78
            Top             =   2805
            Width           =   630
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Cheque:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   77
            Top             =   2445
            Width           =   900
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C. Custo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   480
            TabIndex        =   76
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   660
            TabIndex        =   75
            Top             =   960
            Width           =   465
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Banco:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   630
            TabIndex        =   74
            Top             =   600
            Width           =   510
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(12)"
            Height          =   195
            Index           =   12
            Left            =   3030
            TabIndex        =   73
            Tag             =   "Descrição"
            Top             =   2085
            Width           =   4275
         End
         Begin VB.Label lblDuplicatas 
            AutoSize        =   -1  'True
            Caption         =   "Carteira:"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   31
            Left            =   3210
            TabIndex        =   72
            Top             =   2790
            Width           =   585
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Forma Pagto.:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   150
            TabIndex        =   71
            Top             =   270
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Op. Contábil:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   70
            Top             =   1710
            Width           =   915
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(14)"
            Height          =   195
            Index           =   14
            Left            =   2475
            TabIndex        =   69
            Tag             =   "Descrição"
            Top             =   1710
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
         Left            =   -74940
         TabIndex        =   43
         Top             =   2220
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox txtDuplicatas 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Index           =   45
            Left            =   5520
            MaxLength       =   18
            TabIndex        =   157
            Top             =   1380
            Width           =   1845
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   30
            Left            =   5520
            MaxLength       =   18
            TabIndex        =   51
            Top             =   990
            Width           =   1845
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   29
            Left            =   5520
            MaxLength       =   18
            TabIndex        =   49
            Top             =   600
            Width           =   1845
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   24
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   46
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Centro"
            Height          =   315
            Index           =   20
            Left            =   1200
            MaxLength       =   9
            TabIndex        =   44
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   21
            Left            =   5520
            MaxLength       =   18
            TabIndex        =   47
            Top             =   210
            Width           =   1845
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Excluir..."
            Height          =   375
            Left            =   1170
            TabIndex        =   55
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdRateio 
            Caption         =   "&Ratear..."
            Height          =   375
            Left            =   2220
            TabIndex        =   57
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   3270
            TabIndex        =   59
            Top             =   1320
            Width           =   975
         End
         Begin VB.CommandButton cmdAdicionar 
            Caption         =   "&Adicionar..."
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtDuplicatas 
            DataField       =   "Valor da Moeda"
            Height          =   315
            Index           =   25
            Left            =   1200
            TabIndex        =   45
            Top             =   600
            Width           =   1215
         End
         Begin ComctlLib.ListView lvwRateio 
            Height          =   1215
            Left            =   120
            TabIndex        =   48
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
               Text            =   "Descrição"
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
               Text            =   "Acréscimo"
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
            Left            =   4335
            TabIndex        =   156
            Top             =   1410
            Width           =   1140
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Abatimento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   4635
            TabIndex        =   62
            Top             =   1020
            Width           =   840
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Acréscimo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   4695
            TabIndex        =   61
            Top             =   630
            Width           =   780
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(10)"
            Height          =   195
            Index           =   10
            Left            =   2520
            TabIndex        =   60
            Tag             =   "Descrição"
            Top             =   270
            Width           =   1995
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C. Custo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   510
            TabIndex        =   58
            Top             =   270
            Width           =   645
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Porcentagem:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   150
            TabIndex        =   56
            Top             =   990
            Width           =   990
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   5070
            TabIndex        =   54
            Top             =   240
            Width           =   405
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
                  Picture         =   "Lancamento.frx":0144
                  Key             =   "Checked"
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Lancamento.frx":045E
                  Key             =   "Unchecked"
               EndProperty
            EndProperty
         End
         Begin VB.Label lblDuplicatas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta Financ.:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   52
            Top             =   630
            Width           =   1035
         End
         Begin VB.Label lblDuplDesc 
            Caption         =   "lblDuplDesc(11)"
            Height          =   195
            Index           =   11
            Left            =   2520
            TabIndex        =   50
            Tag             =   "Descrição"
            Top             =   630
            Width           =   1995
         End
      End
      Begin VB.Label lblDuplicatas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nosso Numero:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   45
         Left            =   120
         TabIndex        =   145
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label lblDuplicatas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Carteira:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   46
         Left            =   630
         TabIndex        =   144
         Top             =   3090
         Width           =   585
      End
      Begin VB.Label lblDuplDesc 
         Caption         =   "..."
         Height          =   195
         Index           =   16
         Left            =   2595
         TabIndex        =   143
         Tag             =   "Descrição"
         Top             =   3090
         Width           =   4875
      End
      Begin VB.Label lblDuplicatas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Usuário:"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   -67080
         TabIndex        =   141
         Top             =   5865
         Width           =   585
      End
      Begin VB.Label lblDuplicatas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Linha Digitável:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   44
         Left            =   120
         TabIndex        =   140
         Top             =   2280
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLancamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const IDI_LANCTO = 500          'Ícone (no arquivo de recursos) para Lançamentos
Private Const IDI_DUPL = 501            'Ícone (no arquivo de recursos) para Duplicatas
Private Const TAG_CHEQUE$ = "Cheques"   'Tag dos campos de informações do cheque
Private Const DL_MARCADO = 1        'Índice do ícone de lançamento marcado no ImageList
Private Const DL_DESMARCADO = 2     'Índice do ícone de lançamento desmarcado no ImageList
Private Const IDB_TRANSF = 509          'Imagem para o ListView para Cheques em Transferências
Private Const IDB_DUPLS = 510           'Ídem para Duplicatas
Private Const IDB_LANCTOS = 511         'Ídem para Lançamentos
'Identificadores dos ítens do menu Ferramentas
Private Const IDM_DUPLWNDCALC& = 32010
Private Const IDM_DUPLCHQINFO& = 32011
Private Const IDM_DUPLOBSFIN& = 32012
Private Const IDM_DUPLADDHIST& = 32013
'Valores possíveis da variável mintBaixa
Private Const CDT_NORMAL = 0            'Abertura normal do cadastro
Private Const CDT_BXTOTAL = 1           'Baixa total de uma Duplicata ou Lançamento
Private Const CDT_BXPARCIAL = 2         'Baixa parcial de uma Duplicata ou Lançamento
Private mstrTabela      As String               'Nome da Tabela que está aberta
Private mstrPagRec      As String               'Tipo a pagar ou a receber
Private mintBaixa       As Integer              'Tipo da Baixa se for uma baixa
Private mlngItem        As Long                 'Ítem selecionado da lista de rateio
Private bHabRateio      As Boolean              'Indica se o botão de habilitar rateio vai ficar visível ou não
Private mstrPesquisa    As String               'Instrução SQL utilizada na função de Pesquisa
Private mstrDuplicatas  As String               'Instrução SQL utilizada na abertura da Tabela
Private mrstDuplicatas  As Object               'Abre a tabela
Private mlngDuplicatas  As Long                 'Controla as alterações do usuário
Private mrstCheques     As Object               'Abre a tabela de cheques
Private mlngCheques     As Long                 'Controla as alterações do usuário em cheques
Private SeqLancamentos  As Boolean              'Configuração para sugerir seqüência de Lançamentos
'pt. 83525 - Dulcino Júnior (27/09/2007)
Private mblnAlteraValor As Boolean              'Flag utilizado para não considerar a sugestão da Operação Contábil como alteração.
'Dulcino Júnior (28/10/2007)
Private lngOperacao     As Long
Private mstrOrigem      As String
Private mstrDelete      As String
Private mstrRateio      As String
Private mlngCodigo      As Long
Private mlngPARCELA     As Long

'FUNCTION..: LibProc
'Objetivo..: Função de retorno de chamada da Lib.
'Argumentos: [sFuncao]: Função que deve ser executada.
'            [lFuncao]: Parâmetro adicional, varia conforme a função.
'Retorna...: True se executar a função com sucesso, False, se não.
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
    Dim sTmp          As String
    Dim nBanco        As Long 'Código do Banco
    Dim nCheque       As Long 'Número do Cheque
    Dim blnBtLanc     As Boolean
    Dim strSql        As String
    Dim ObsFin        As String
    Dim bGeraComissao As Boolean
    Dim oCom          As New CCOMISSAO
    Dim Vendedor      As Long
    
    If cmdAbreRateio.Visible Then
        cmdAbreRateio.Enabled = True
    End If
    'pt. 85297 - Moacir Pfau(26/05/2008)
    'Tratamento botao "Efetivar lancto", pelo motivo do WL_EXIBIR.
    blnBtLanc = True

    Select Case sFuncao
        Case WL_NOVO
            If mintBaixa = CDT_NORMAL Then 'Somente se for abertura normal
                If LimpaControles(mrstDuplicatas, Me, Tag, mlngDuplicatas) = WL_OK Then
                    Call CarregaPadrao  'pt. 86140 - Moacir Pfau(10/04/2008)
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
            btnEfetiva.Enabled = False
      
        Case WL_SETFOCUS
            SSTab1.Tab = 0
            txtDuplicatas(1).SetFocus
      
        Case WL_DELETAR
            sTmp = GetValue(mrstDuplicatas, "Liberação", NUL)
            If sTmp <> "" Then
                If Not ValidaDatasDiasUteis(0, 0, CDate(sTmp), True) Then
                    Exit Function
                End If
            End If
            'Grava o código do Banco e Cheque atual para a rotina ChequeInfo
            nBanco = GetValue(mrstDuplicatas, "Banco", ZERO)
            nCheque = GetValue(mrstDuplicatas, "Cheque", ZERO)
            'pt. 81604 - Dulcino Júnior
            If Not PermiteExclusao Then
                Exit Function
            End If
            If DeletaRegistro(mrstDuplicatas, Me, Tag, mlngDuplicatas) = WL_OK Then
                Call ChequeInfo(WL_DELETAR, nBanco, nCheque): LibProc = True
                If mstrOrigem <> "" Then
                    Call ExecuteSQL(mstrOrigem)
                    mstrOrigem = ""
                    Call ExecuteSQL(mstrDelete)
                    mstrDelete = ""
                    If mstrRateio <> "" Then
                        If Recordcount(mstrRateio) = 0 Then
                            Call ExecuteSQL("UPDATE Lançamentos SET proveniente_rateio=False WHERE PagRec='" & mstrPagRec & _
                                "' AND Código=" & mlngCodigo & " AND Parcela= " & mlngPARCELA)
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
            If mintBaixa = CDT_NORMAL Then 'A janela Localizar só é habilitada em modo normal
                If localizar(mrstDuplicatas, Me, mstrTabela, Tag, mlngDuplicatas) = WL_OK Then
                    Call ChequeInfo(WL_LOCALIZAR): LibProc = True
                End If
            Else
                MsgFunc ResolveResString(240, resUM, "localizar")
            End If
      
        Case WL_PESQUISAR
            strSql = NUL
            If mintBaixa = CDT_NORMAL Then
                If Configuracao("Visualizar somente Movimentações não Conferidas", False) Then
                    strSql = " and Liberação >= " & InverteData(DateAdd("M", 1, MaxValue("Mês Conferido", "Mov Conferido", "KIF = True")), True)
                End If
                If PRegistro(mrstDuplicatas, Me, Caption, mstrDuplicatas & strSql, mstrPesquisa & strSql, Tag, mlngDuplicatas, PB_REGISTRO) = WL_OK Then
                    Call ChequeInfo(WL_PESQUISAR): LibProc = True
                End If
            Else
                If (FindFirst(mrstDuplicatas, Me, Tag, mstrPesquisa, mlngDuplicatas) = WL_OK) Then
                    Call ChequeInfo(WL_PESQUISAR): LibProc = True
                End If
            End If
      
        Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
            DoEvents
            If WL_OK = MoveRecordset(mrstDuplicatas, Me, Tag, mlngDuplicatas, lFuncao) Then
                Call ChequeInfo(sFuncao): LibProc = True
            End If
      
        Case WL_NAVEGAR
            If Browse(mrstDuplicatas, Me, Tag, mlngDuplicatas, mstrDuplicatas) = WL_OK Then
                Call ChequeInfo(WL_NAVEGAR): LibProc = True
            End If
      
        Case WL_SALVAR
            If DuplVerifique() Then
                'No caso de estar configurado para utilizar Op. Contábil.
                If txtDuplicatas(41).Enabled Then
                    'pt. 81902 - Dulcino Júnior
                    If mstrTabela = "Lançamentos" Then
                        If Not validaIntegracaoLancamentos Then
                            Exit Function
                        End If
                    Else
                        If Not validaIntegracaoDuplicatas Then
                            Exit Function
                        End If
                    End If
                End If
                ObsFin = GetFieldValue("[Obs Financeira]", "Empresas", "Apel = '" & txtDuplicatas(2).Text & "'")
                If Len(ObsFin) > 0 Then
                    MsgBox ObsFin, vbInformation, "Observações Financeiras da Empresa"
                End If
                nBanco = GetValue(mrstDuplicatas, "Banco", ZERO)
                nCheque = GetValue(mrstDuplicatas, "Cheque", ZERO)
                txtDuplicatas(26).Text = UserName
                txtDuplicatas(3).Text = Date
                bGeraComissao = (IsNull(mrstDuplicatas("Pagamento")) And txtDuplicatas(8).Text <> "")
                'pt. 86132 - Ivo Sousa (25/03/2008)
                'Validação de Datas
                If ValidaDatas Then
                    If SalvaRegistro(mrstDuplicatas, Me, Tag, mlngDuplicatas) = WL_OK Then
                        ExibeSoma
                        'pt: 74271 - Dulcino Júnior
                        'Erro ao alterar uma duplicata que não tem cheque
                        If strToDbl(txtDuplicatas(16).Text) > 0 Then
                            txtBancoCheque.Text = txtDuplicatas(13).Text
                            txtChequeCheque.Text = txtDuplicatas(16).Text
                        End If
                        nBanco = GetValue(mrstDuplicatas, "Banco", ZERO)
                        nCheque = GetValue(mrstDuplicatas, "Cheque", ZERO)
                        Call ChequeInfo(WL_SALVAR, nBanco, nCheque): LibProc = True
                        'pt. 88289 - Dulcino Júnior (15/10/2008)
                        If chkRateio.value = vbChecked Then
                            If txtDuplicatas(8).Text <> "" Then
                                Call ExecuteSQL("UPDATE FFIRateioLancamento SET dt_pagamento=" & InverteData(txtDuplicatas(8).Text, True) & _
                                                " WHERE pag_rec_destino='" & mstrPagRec & "' AND cd_lancamento_destino=" & txtDuplicatas(1).Text & _
                                                "AND nr_parcela_destino=" & txtDuplicatas(4).Text)
                            Else
                                Call ExecuteSQL("UPDATE FFIRateioLancamento SET dt_pagamento=NULL" & _
                                                " WHERE pag_rec_destino='" & mstrPagRec & "' AND cd_lancamento_destino=" & txtDuplicatas(1).Text & _
                                                "AND nr_parcela_destino=" & txtDuplicatas(4).Text)
                            End If
                        End If
                    End If
                End If
                'Geração da comissão
                If Configuracao("TipGcm") = "A" Then
                    If bGeraComissao Then
                        Call oCom.GeraComissaoDuplicata(GBL_NFS, GetValue(mrstDuplicatas, "Nota"), GetValue(mrstDuplicatas, "Empresa"), GetValue(mrstDuplicatas, "Tipo"), GetValue(mrstDuplicatas, "Parcela"))
                        Set oCom = Nothing
                    End If
                End If
            End If
      
        Case WL_CANCELAR
            If CancelaEdicao(mrstDuplicatas, Me, Tag, mlngDuplicatas) = WL_OK Then
                Call ChequeInfo(WL_CANCELAR): LibProc = True
            End If
      
        Case WL_EXIBIR
            If mintBaixa = CDT_NORMAL Then
                sTmp = mstrDuplicatas   '// Termina e completa a instrução conforme a tabela
                If mstrTabela = "Duplicatas" Then
                    Concat sTmp, " AND Nota = {Nota} AND Parcela = {Parcela} AND Tipo = '{Tipo}' AND Empresa = '{Empresa}';"
                Else
                    'pt. 83992 e 83998 - Dulcino Júnior (19/10/2007)
                    Concat sTmp, " AND Código = {Código} AND Parcela = {Parcela} AND Tipo = '{Tipo}';"
                End If
                If RetornaRegs(mrstDuplicatas, Me, Tag, sTmp, mlngDuplicatas) = WL_OK Then
                    Call ChequeInfo(WL_EXIBIR): LibProc = True
                ElseIf UltimoRetorno() = WL_ADDNEW Then
                    Call NovoRegistro(False)
                    Call ChequeInfo(WL_NOVO)
                    LibProc = True
                    lblDadosAdcionais(0).Caption = NUL
                    lblDadosAdcionais(1).Caption = NUL
                    lblDadosAdcionais(2).Caption = NUL
                    If cmdAbreRateio.Visible Then
                        cmdAbreRateio.Enabled = True
                    End If
                    blnBtLanc = False
                End If
            End If
      
        Case WL_FILTRAR
            If mintBaixa = CDT_NORMAL Then    '// Só filtra se for abertura normal
                If Filtrar(mrstDuplicatas, Me, Tag, mstrDuplicatas, mlngDuplicatas) = WL_OK Then
                    Call ChequeInfo(WL_FILTRAR): LibProc = True
                End If
            Else
                MsgFunc ResolveResString(240, resUM, "filtrar")
            End If
        
        Case WL_DUPLICADO 'Registro Duplicado
            If mintBaixa = CDT_NORMAL Then    '// Só resolve se for abertura normal
                If mstrTabela = "Lançamentos" Then
                    ResolveDuplicacao Me, txtDuplicatas(1), "Lançamentos", "PagRec = " & Quote(mstrPagRec, "''")
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
            If KeybAcesso(LoadResString(2037)) Then
                frmEmpresas.Show
                CallChange frmEmpresas.hWnd, txtDuplicatas(2).hWnd
            End If
      
        Case "Bancos"
            If KeybAcesso(LoadResString(2003)) Then
                frmBancos.Show
                CallChange frmBancos.hWnd, txtDuplicatas(13).hWnd
            End If
      
        Case "Contas"
            If KeybAcesso(LoadResString(2007)) Then
                frmContas.Show
                CallChange frmContas.hWnd, txtDuplicatas(14).hWnd
            End If
      
        Case "Custos"
            If KeybAcesso(LoadResString(2029)) Then
                frmCusto.Show
                CallChange frmCusto.hWnd, txtDuplicatas(15).hWnd
            End If
      
        Case "Moedas"
            If KeybAcesso(LoadResString(2033)) Then
                fMoedas.Show
                CallChange fMoedas.hWnd, txtDuplicatas(17).hWnd
            End If
      
        Case "Configuração"
            If KeybAcesso(LoadResString(2106)) Then
                FrmConfCad.Configura "Duplicatas"
                FrmConfCad.Show vbModal
            End If
      
        Case IDM_DUPLWNDCALC 'Atualizar Valor
            CalcValor
        
        Case IDM_DUPLCHQINFO 'Informações do Cheque
            Call ChequeInfo("updt")
      
        Case IDM_DUPLOBSFIN 'Observações Financeiras
            If IsValid(txtDuplicatas(2).Text) Then
                Call fMemo("Observações Financeiras", "Empresas", "[Obs Financeira]", wsprintf("Apel = '%s'", txtDuplicatas(2).Text))
            End If
      
        Case IDM_DUPLADDHIST 'Histórico do Cheque
            If Len(txtDuplicatas(23).Text) Then
                txtCheque(1).Text = wsprintf("%s\n%s", txtCheque(1).Text, txtDuplicatas(23).Text)
            End If
    End Select
    If UltimoRetorno = WL_OK And sFuncao <> WL_NOVO Then
        If cmdAbreRateio.Visible Then
            cmdAbreRateio.Enabled = True
        End If
    Else
        If UltimoRetorno <> WL_ADDNEW And UltimoRetorno <> 0 Then
            If cmdAbreRateio.Visible Then
                cmdAbreRateio.Enabled = True
            End If
        Else
            If cmdAbreRateio.Visible Then
                cmdAbreRateio.Enabled = True
            End If
        End If
    End If
    If Not IsValid(GetValue(mrstDuplicatas, "Pagamento", NUL)) And Not EAdicao(mlngDuplicatas) Then
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
      MsgBar IIf(IsValid(txtDuplicatas(1).Text), IIf(mstrTabela = "Duplicatas", " Nota: " & txtDuplicatas(1).Text, " Lançamento: " & txtDuplicatas(1).Text), " ") _
        & IIf(IsValid(cboDuplicatas(3).Text), " - Tipo de Registro:" & cboDuplicatas(3).Text, " ") _
        & IIf(IsValid(txtDuplicatas(2).Text), " - Empresa: " & txtDuplicatas(2).Text, " ") _
        & IIf(mstrTabela = "Duplicatas", IIf(IsValid(txtDuplicatas(4).Text), " - Parcela: " & txtDuplicatas(4).Text, " "), " ")
    End If
    If LibProc = True Then
        lblDadosAdcionais(0).Caption = "Cidade: " & GetFieldValue("Cidade", "Empresas", "Apel = " & Quote(GetValue(mrstDuplicatas, "Empresa", NUL), "'"), , NUL)
        lblDadosAdcionais(1).Caption = "Estado: " & GetFieldValue("Estado", "Empresas", "Apel = " & Quote(GetValue(mrstDuplicatas, "Empresa", NUL), "'"), , NUL)
'        If GetValue(mrstDuplicatas, "PagRec") = "R" Then
'            Vendedor = GetFieldValue("Vendedor01", Quote(GBL_ITENS & GBL_NFS, "[]"), "Número = " & GetValue(mrstDuplicatas, "Nota", ZERO) & " AND [Tipo de Registro] = " & Quote(GetValue(mrstDuplicatas, "Tipo", NUL), "'"), , ZERO)
'            If Vendedor > 0 Then
'                lblDadosAdcionais(2).Caption = "Vendedor: " & Format(Vendedor, "000000") & " - " & GetFieldValue("Nome", "Vendedores", "Código = " & Vendedor, , NUL)
'            End If
'        Else
'            lblDadosAdcionais(2).Caption = NUL
'        End If
    End If
    'pt. 85297 - Moacir Pfau(23/05/2008)
    'valida previsão em lançamentos.
    If mstrTabela = "Lançamentos" Then
        If Not mrstDuplicatas.EOF Then
            If sFuncao <> "novo" And sFuncao <> "excluido" And blnBtLanc Then
                If GetValue(mrstDuplicatas, "Previsao", False) = True Then
                    btnEfetiva.Enabled = True
                Else
                    btnEfetiva.Enabled = False
                End If
            Else
                btnEfetiva.Enabled = False
            End If
        End If
    End If
End Function

Private Sub btnEfetiva_Click()
    Dim strUpdate As String
    
    If MsgBox("Confirma efetivação do lançamento de previsão ?", vbYesNo, "Confirmação") = vbYes Then
        btnEfetiva.Enabled = False
        'Pt. 95368 - Moacir Pfau(12/11/2009)
        'mrstDuplicatas.Edit
        mrstDuplicatas("Previsao").value = False
        mrstDuplicatas.update
    End If
End Sub

Private Sub cboDuplicatas_Change(Index As Integer)
    If mstrTabela = "Lançamentos" Then
        If Index <> 3 Then
            AlteraValor mlngDuplicatas
        End If
    End If
End Sub

Private Sub cboDuplicatas_Click(Index As Integer)
    If mstrTabela = "Lançamentos" Then
        'pt. 85297 - Moacir Pfau(26/05/2008)
        If Not EAddNew(mlngDuplicatas) Then
            AlteraValor mlngDuplicatas
        End If
    Else
        If Index = 3 Then
            If Not ControlaChave(CBCLICK, ZERO, cboDuplicatas(3), mlngDuplicatas) Then
                cboDuplicatas(3).Text = GetValue(mrstDuplicatas, "Tipo")
            End If
        ElseIf Index > 3 Then
            AlteraValor mlngDuplicatas
        End If
    End If
End Sub

Private Sub cboDuplicatas_DropDown(Index As Integer)
    If mstrTabela = "Duplicatas" And Index = 3 Then   'Campo Tipo
        ControlaChave CBDROPDOWN, 0, cboDuplicatas(3), mlngDuplicatas
    End If
End Sub

Private Sub cboDuplicatas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If mstrTabela = "Duplicatas" Then   'Tipo não entra na chave quando for lançamentos
        If Index = 3 Then
            ControlaChave KeyCode, Shift, cboDuplicatas(3), mlngDuplicatas
        End If
    End If
End Sub

Private Sub cboDuplicatas_LostFocus(Index As Integer)
    Dim MatrizDAO               As cMatrizContabilizacaoDAO
    Dim matriz                  As cMatrizContabilizacao
    Dim lngOpContabilbaixa      As Long

    If mstrTabela = "Duplicatas" Then
        If Index = 3 Then
            LibProc WL_EXIBIR
        End If
    End If
    'Carrega a operação contábil padrão para as duplicatas e lancamentos
    If Index = 3 Then
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
    'Pt. 114146 - Moacir Pfau(29/02/2012)
    Call TotalizaValorRateio
End Sub

Private Sub cmdAdicionar_Click()
    Dim bUsaPorc           As Boolean
    Dim dblTotPorcentagem  As Double
    Dim curTotRateio       As Currency
    Dim curTotRateioAcres  As Currency
    Dim curTotRateioAbat   As Currency


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
        MsgFunc "Centro de custo não cadastrado!"
        txtDuplicatas(20).SetFocus
        Exit Sub
    End If
    If IsValid(txtDuplicatas(21).Text) And IsValid(txtDuplicatas(24).Text) Then
        MsgFunc "Informe apenas um valor!"
        txtDuplicatas(21).SetFocus
        Exit Sub
    End If
    If IsValid(txtDuplicatas(29).Text) And IsValid(txtDuplicatas(24).Text) Then
        MsgFunc "Acréscimo e porcentagem não devem ser informados ao mesmo tempo!"
        txtDuplicatas(29).SetFocus
        Exit Sub
    End If
    If IsValid(txtDuplicatas(30).Text) And IsValid(txtDuplicatas(24).Text) Then
        MsgFunc "Abatimento e porcentagem não devem ser informados ao mesmo tempo!"
        txtDuplicatas(30).SetFocus
        Exit Sub
    End If
    'pt. 86140 - Moacir Pfau(07/04/2008)
    If Len(txtDuplicatas(25).Text) = 0 Or Len(lblDuplDesc(11).Caption) = 0 Then
        MsgFunc "Conta não cadastrada!"
        txtDuplicatas(25).SetFocus
        Exit Sub
    End If
    'Verifica se conta está ativa
    If GetFieldValue("Ctaati", "Contas", " [Código]=" & txtDuplicatas(25).Text) = "N" Then
        MsgBox "Conta " & txtDuplicatas(25).Text & " não está ativa", vbCritical, MsgBoxCaption
        txtDuplicatas(25).SetFocus
        Exit Sub
    End If
    bUsaPorc = UsaPorcentagemnoRateio
    If lvwRateio.ListItems.Count > 0 Then
        'Verificando se usa porcentagem ou valor
        If bUsaPorc And IsValid(txtDuplicatas(21).Text) Or Not bUsaPorc And IsValid(txtDuplicatas(24).Text) Then
            MsgFunc "Só é possível utilizar uma forma de rateio de cada vez!"
            Exit Sub
        End If
    End If
  
    If IsValid(txtDuplicatas(24).Text) Then
        dblTotPorcentagem = SomaPorcentagens
        If CSng((dblTotPorcentagem + CSngDef(txtDuplicatas(24).Text))) <= CSng(100) Then
            lvwRateio.ListItems.add , , txtDuplicatas(20).Text, , DL_MARCADO
            lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(1) = lblDuplDesc(10).Caption
            lvwRateio.ListItems(lvwRateio.ListItems.Count).SubItems(2) = Format(txtDuplicatas(24).Text, F4CASAS) & "%"
        Else
            MsgBox "Total de Porcentagens  é maior que 100%", vbCritical, "Rateio"
            Exit Sub
        End If
    Else
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
                    MsgBox "A soma dos abatimentos é maior que " & txtDuplicatas(12).Text, vbCritical, "Rateio"
                    Exit Sub
                End If
            Else
                MsgBox "A soma dos acréscimos é maior que " & txtDuplicatas(11).Text, vbCritical, "Rateio"
                Exit Sub
            End If
        Else
            MsgBox "A soma dos valores originais é maior que " & txtDuplicatas(10).Text, vbCritical, "Rateio"
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
    'Verificando se a porcentagem está correta ou se usa %
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
    txtCheque(0).Text = GetFieldValue("Razão", "Empresas", "Apel = " & Quote(txtDuplicatas(2).Text, "'"), , NUL)
End Sub

Private Sub cmdProximoCheque_Click()
    Dim rstProximoCheque As Object
    Dim strSql           As String
  
    strSql = "SELECT * FROM Cheque WHERE Banco = " & CLngDef(txtDuplicatas(13).Text) & " AND Situação = 'Normal' "
    strSql = strSql & "AND (Cheque not in (SELECT Cheque FROM Duplicatas WHERE Banco = Cheque.Banco) AND Cheque NOT IN "
    strSql = strSql & "(SELECT Cheque FROM Lançamentos WHERE Banco = Cheque.Banco) AND Cheque NOT IN "
    strSql = strSql & "(SELECT Cheque FROM [Transf Bancária] WHERE Banco = Cheque.Banco)) ORDER BY Cheque ASC"
    If AbreRecordset(rstProximoCheque, strSql, dbOpenSnapshot) = WL_OK Then
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
    Dim VrTotalAcres As Currency 'valor total do acréscimo rateado
    Dim VrTotalAbat  As Currency 'valor total do abatimento rateado
    Dim vrDuplicata  As Currency 'valor da duplicata
    Dim vrAcrescimo  As Currency 'valor do acréscimo
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
    If Not IsValid(txtDuplicatas(10).Text) Then
        MsgFunc "Preencha o campo de valor original para que o mesmo seja rateado!"
        txtDuplicatas(10).SetFocus
        Exit Sub
    End If
       
    'Acréscimo
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
        MsgFunc "O campo de Empresa não pode ficar em branco"
        txtDuplicatas(2).SetFocus
        Exit Sub
    End If
    'Checando datas
    If Not IsValid(txtDuplicatas(6).Text) Or Not IsValid(txtDuplicatas(7).Text) Or Not IsValid(txtDuplicatas(9).Text) Then
        MsgFunc "Os campos de data de Emissão,Liberação e Vencimento são obrigatórios!"
        Exit Sub
    End If
    vrTemp = ZERO
    'valido se o rateio baterá com o valor da duplicata, para o rateio de valores
    If Not UsaPorcentagemnoRateio Then
        vrTemp = SomaValores()
        If vrTemp <> vrDuplicata Then
            MsgFunc "Valor Original é diferente que o valor do rateio"
            Exit Sub
        End If
        vrTempAcres = SomaValoresAcres()
        If vrTempAcres <> vrAcrescimo Then
            MsgFunc "Valor de Acréscimo é diferente do valor do rateio"
            Exit Sub
        End If
        vrTempAbat = SomaValoresAbat()
        If vrTempAbat <> vrAbatimento Then
            MsgFunc "Valor de Abatimento é diferente que o valor do rateio"
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
            'diferenca, devido aos arredondamentos, no ultimo lançamento
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
        lngParcela = ProximoNumero("Parcela", "Lançamentos", "PagRec='" & mstrPagRec & "' AND Código=" & lngCodigo & "AND Tipo='" & cboDuplicatas(3).Text & "'")
        strSql = "INSERT INTO Lançamentos(PagRec, Código, Parcela, Empresa, Tipo, Descrição, "
        strSql = strSql & "Emissão, Vencimento, Pagamento, Liberação, [Valor Original], Acréscimo, "
        strSql = strSql & "Abatimento, Banco, Conta, Centro, Cheque, Moeda, " & IIf(ChkConciliado, "Conciliado,", "")
        strSql = strSql & "[Valor da Moeda], Controle, Marcação, Obs, Borderô, cd_operacao_contabil, Usuário, proveniente_rateio) VALUES ('" & mstrPagRec & "', "
        If lngItens = 1 And EEdicao(mlngDuplicatas) Then
            strUpdate = "UPDATE " & mstrTabela & " SET [Valor Original]=" & Replace(vrTemp, ",", ".") & " ,"
            strUpdate = strUpdate & "[Acréscimo]=" & Replace(vrTempAcres, ",", ".") & " ,"
            strUpdate = strUpdate & "[Abatimento]=" & Replace(vrTempAbat, ",", ".") & " ,"
            strUpdate = strUpdate & "Conta =" & CLngDef(lvwRateio.ListItems(lngItens).SubItems(6)) & " ,"
            strUpdate = strUpdate & "Centro =" & CLngDef(lvwRateio.ListItems(lngItens).Text) & ", "
            strUpdate = strUpdate & "proveniente_rateio=True "
            strUpdate = strUpdate & "WHERE PagRec = " & Quote(mstrPagRec, "'") & " AND Código = " & CLngDef(txtDuplicatas(1).Text)
            Call ExecuteSQL(strUpdate)
        Else
            AppendStr strSql, CStr(lngCodigo) ' Codigo
            AppendStr strSql, ", " & lngParcela 'Parcela
            AppendStr strSql, ", " & Quote(txtDuplicatas(2).Text, "''")          'Empresa
            AppendStr strSql, ", " & Quote(cboDuplicatas(3).Text, "''")          'Tipo
            AppendStr strSql, ", " & Quote(txtDuplicatas(5).Text, "''")          'Descrição
            AppendStr strSql, ", " & InverteData(txtDuplicatas(6).Text, True)    'Emissão
            AppendStr strSql, ", " & InverteData(txtDuplicatas(7).Text, True)    'Vencimento
            AppendStr strSql, ", " & IIf(IsValid(txtDuplicatas(8).Text), InverteData(txtDuplicatas(8).Text, True), "Null") 'Pagamento
            AppendStr strSql, ", " & InverteData(txtDuplicatas(9).Text, True)    'Liberação
            AppendStr strSql, ", " & Replace(vrTemp, ",", ".")                      'Valor original
            AppendStr strSql, ", " & Replace(vrTempAcres, ",", ".")                 'Acréscimo
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
            'Verifica se esta o flag de conciliação esta ativo
            If ChkConciliado Then
                AppendStr strSql, ", True "                                        'Conciliado
            End If
            'o campo valor da moeda foi retirado do formulário.
            AppendStr strSql, ", " & ValStr(CMoeda(""))                          'Valor da Moeda
            AppendStr strSql, ", " & Quote(txtDuplicatas(19).Text, "''")         'Controle
            AppendStr strSql, ", 0"                                              'Marcação
            AppendStr strSql, ", " & Quote(txtDuplicatas(23).Text, "''")         'Observação
            AppendStr strSql, ", 0"
            'pt. 87144 - Moacir Pfau(08/07/2008)
            AppendStr strSql, ", " & CLngDef(txtDuplicatas(40).Text)
            AppendStr strSql, ", " & Quote(txtDuplicatas(26).Text, "''")
            AppendStr strSql, ", True)" 'pt. 88289 - Dulcino Júnior(10/10/2008)
            'pt. 88289 - Dulcino Júnior(10/10/2008)
            If ExecuteSQL(strSql) > 0 Then
                strSql = "INSERT INTO FFIRateioLancamento(pag_rec_origem, cd_lancamento_origem, nr_parcela_origem, "
                strSql = strSql & "pag_rec_destino, cd_lancamento_destino, nr_parcela_destino, cd_centro, cd_conta, "
                strSql = strSql & "vl_valor) VALUES('" & mstrPagRec & "', " & lngCodigo & ", " & txtDuplicatas(4).Text
                strSql = strSql & ", '" & mstrPagRec & "', " & lngCodigo & ", " & lngParcela & ", " & CLngDef(lvwRateio.ListItems(lngItens).Text)
                strSql = strSql & ", " & CLngDef(lvwRateio.ListItems(lngItens).SubItems(6)) & ", " & Replace(vrTemp, ",", ".") & ")"
                If ExecuteSQL(strSql) = 0 Then
                    GoTo Error_Handler
                End If
            End If
        End If
    Next
    CommitTrans
    MsgFunc "Rateio concluído!"
    FraRateio.Visible = False
    FraDuplicatas(1).Visible = True
    cmdAbreRateio.Enabled = False
    DefEditNone mlngDuplicatas
    'PT.  - Carlos Felippe Vernizze - 17/05/2011
    mrstDuplicatas.Requery
    'pt. 86140 - Moacir Pfau(07/04/2008)
    Call txtDuplicatas_LostFocus(4)
    Exit Sub
    
Error_Handler:
    MsgBox "O rateio não foi concluido!"
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
    'Protocolo 73916: Verifica se o usuário pode baixar duplicatas/lançamentos
    If GetAcesso(LoadResString(2002)) <> SEM_ACESSO Then
        txtDuplicatas(8).Enabled = True
        lblDuplicatas(16).ForeColor = &H0&         'Ativado
    Else
        txtDuplicatas(8).Enabled = False
        lblDuplicatas(16).ForeColor = &H0&         'Desativado
    End If
    Call CarregaTipos
    ConfigCampos Me, Tag, Tag
    cmdAbreRateio.Visible = CentrodeCusto(MFinanceiro)

    'Retirando os Captions dos Labels de Descrição que coloquei em
    'design time.
    For intLabels = 0 To 13
        lblDuplDesc(intLabels).Caption = NUL
    Next

    Call ConfigureList
    'pt. 81189 - Dulcino Júnior
    'Integração Contábil
    Label1.Enabled = ConfigSys.UtilizaIntegracaoContabil
    txtDuplicatas(40).Enabled = ConfigSys.UtilizaIntegracaoContabil
    lblDuplDesc(14).Enabled = ConfigSys.UtilizaIntegracaoContabil
    Label2.Enabled = ConfigSys.UtilizaIntegracaoContabil
    txtDuplicatas(41).Enabled = ConfigSys.UtilizaIntegracaoContabil
    lblDuplDesc(15).Enabled = ConfigSys.UtilizaIntegracaoContabil
    mblnAlteraValor = True
    'Pt. 94752 - Moacir Pfau(22/10/2009)
    txtDuplicatas(44).MaxLength = 3
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (mintBaixa <> CDT_NORMAL) Then
        If (UnloadMode > vbFormCode) Then
            MsgFunc LoadResString(245)
            Cancel = True
            Exit Sub
        End If
    End If

    If Not UnloadForm(mrstDuplicatas, Me, Tag, mlngDuplicatas) Then
        'Verifica se há alguma alteração nos campos de cheque
        If EstaEditando(mlngCheques) And IsVisibleRecord(mlngCheques) Then
            'Pt. 95368 - Moacir Pfau(17/11/2009)
            'If gTipoDB = Access Then mrstCheques.Edit
            mrstCheques("Nominal").value = txtCheque(0).Text
            mrstCheques("Histórico").value = txtCheque(1).Text
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
    Set frmDuplicatas = Nothing
    mlngDuplicatas = 0
End Sub

Private Sub lvwLancamentos_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    lvwLancamentos.SortKey = ColumnHeader.Index - 1
End Sub

'SUB.......: Configure
'Objetivo..: Configura o Cadastro antes da abertura.
'Argumentos: [strTabela]: Nome da Tabela que será aberta: Duplicatas ou Lançamentos
'            [strPagRec]: Tipo da Tabela, A Pagar ou A Receber.
'            [intBaixa] : Opcional. Se estiver abrindo da janela de baixas procede
'                         configurações extras do cadastro.
'            [strBaixa] : Opcional. Instrução Select para abertura do cadastro quando
'                         este é chamado da janela de baixas.
Public Sub Configure(strTabela As String, strPagRec As String, Optional intBaixa As Integer, Optional strBaixa As String)
    Dim strOptCombo As String

    SetPtr vbHourglass
    'Configurando as instruções de abertura da tabela conforme o nome
    btnEfetiva.Enabled = False
  
    'Instrução utilizada na função de Pesquisa
    mstrPesquisa = "SELECT Código, Parcela, Empresa, Tipo, Descrição, Emissão, Vencimento, Pagamento, Liberação, [Valor Original], Acréscimo, Abatimento, " & _
    "Banco, Conta, Centro, Cheque, Moeda, [Valor da Moeda], Controle, Previsao, SeqNossoNumero, CODFPG FROM Lançamentos WHERE PagRec = '" & strPagRec & "'"
    'Instrução utilizada na abertura do Cadastro
    mstrDuplicatas = "SELECT * FROM Lançamentos WHERE PagRec = '" & strPagRec & "'"
'    'Lista de opções do campo Tipo (ComboBox)
'    strOptCombo = "SELECT Texto FROM Opções WHERE Rotina = '" & OPT_LANCAMENTOS & "';"
    lblDuplicatas(4).Caption = "Código:"      'Label do campo chave da tabela
    txtDuplicatas(1).DataField = "Código"     'Altera o DataField do campo chave
    FraDuplicatas(0).Caption = "Lançamentos"  'Título do Frame
    Set Me.Icon = LoadResIcon(IDI_LANCTO)     'Carrega o ícone para Lançamentos
    Caption = "Lançamentos"                   'Caption do formulário para a tabela de Lançamentos
    SeqLancamentos = Configuracao("Seqüenciar Lançamentos de Entrada e de Saída", False)

    If (intBaixa <> CDT_NORMAL) Then            'Se estiver sendo carregado da janela de baixas
        mstrDuplicatas = strBaixa                 'Instrução para a abertura do Recordset
        'Completa a instrução de pesquisa com as mesmas comparações utilizadas para abrir a
        'tabela. Em baixas o usuário não pode pesquisar o Banco de Dados exceto nas Duplicatas
        'ou Lançamentos abertos por ele em Baixas.
        MidStr mstrPesquisa, " PagRec = '" & strPagRec & "'", ExtractStr(strBaixa, "WHERE", NUL)
    End If
    'Campo Situação, visível apenas quando a tabela for a Receber
    cmdNominalRazaoSocial.Visible = (strPagRec = "P")
    LoadResOptions 1000, cboDuplicatas(20)      'Carrega a lista de opções do campo
'
'    'Carregando as opções do Campo Tipo
'    ComboAddItem cboDuplicatas(3), strOptCombo, "Texto"
    'O campo Cheque permanece visível apenas quando o tipo for a pagar
    lblDuplicatas(8).Visible = (strPagRec = "P")      'Label do campo Cheque
    txtDuplicatas(16).Visible = (strPagRec = "P")     'Campo Cheque
    cmdProximoCheque.Visible = (strPagRec = "P")

    'Oculta o campo Centro de Custo se o usuário desejar
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
        lblDuplDesc(3).Visible = False                  'Descrição do campo Centro
    End If

    'Exibo a janela antes de terminar a configuração porque ao iniciar com
    'pt. 81828 - Dulcino Júnior
    Call CenterForm(Me)

    'A parte de informações do cheque deve estar visível somente quando A Pagar
    If strPagRec = "R" Then
        FraDuplicatas(5).Visible = False 'Oculta o Frame de informações do cheque
    End If
    If strPagRec = "P" Then
        txtDuplicatas(27).Visible = False
        lblDuplicatas(31).Visible = False
    End If
    'Termina de completar o caption do formulário conforme o tipo
    If strPagRec = "P" Then
        Caption = Caption & " a Pagar ou Pagos"
    Else
        Caption = Caption & " a Receber ou Recebidos"
    End If
    'Configura as variáveis de controle
    mstrTabela = strTabela
    mstrPagRec = strPagRec
    mintBaixa = intBaixa
    'Abrindo o Cadastro
    AbreRecordset mrstDuplicatas, mstrDuplicatas            'Abre a tabela de duplicatas
    AbreRecordset mrstCheques, "Cheque", dbOpenDynaset 'Abre a tabela de cheques
    If (intBaixa = CDT_NORMAL) Then                         'Como padrão abre com um novo registro
        Me.LibProc WL_NOVO
    End If
    DefineAcesso mlngDuplicatas, Acesso
    mlngCheques = WL_USERADDNEW                         'Define a variável para os campos do cheque
    DefineAcesso mlngCheques, AC_CADASTRAR Or AC_EDITAR 'Define o acesso aos campos do cheque
    'Se a abertura do cadastro for através das baixas reconfiguro o acesso do usuário
    If intBaixa > CDT_NORMAL Then
        DeleteFlag AC_CADASTRAR, mlngDuplicatas    'Não é permitido adicionar duplicatas em baixas
        If (CompStr(strTabela, "Duplicatas") And CompStr(strPagRec, "R")) Then
            txtDuplicatas(10).Enabled = False        'Não é permitido alterar o valor original quando em baixas de Duplicatas a Receber
        End If
        LibProc WL_PRIMEIRO, MC_MOVEFIRST          'Posiciona no primeiro registro
    End If
    SetPtr vbDefault
End Sub

'FUNCTION..: DuplVerifique
'Objetivo..: Faz as verificações padrão do cadastro
'Retorna...: True se for possível salvar, False se não.
Private Function DuplVerifique() As Boolean
    Dim strOptions As String
    Dim strData    As String

    SetPtr vbHourglass
    ' Verificando as datas do cadastro Emissão (Verifica se a data de emissão é uma data válida)
    If Not EData(txtDuplicatas(6).Text) Then
        MsgFunc ResolveResString(26, resUM, txtDuplicatas(6).Text), vbInformation
        GoTo DuplVerifique_Erro
    End If
    ' Vencimento
    If Not EData(txtDuplicatas(7).Text) Then
        MsgFunc ResolveResString(26, resUM, txtDuplicatas(7).Text), vbInformation
        GoTo DuplVerifique_Erro
    Else
        ' Verifica se a data de Vencimento não é menor que a data de Emissão
        If DateDiff("d", txtDuplicatas(6).Text, txtDuplicatas(7).Text) < 0 Then
            MsgFunc ResolveResString(139, resUM, "de Vencimento", resDOIS, "de Emissão"), vbInformation
            GoTo DuplVerifique_Erro
        End If
    End If
    strData = CDateDef(txtDuplicatas(9).Text)
    If (CLngDef(txtDuplicatas(15).Text) > 0) And Len(strData) Then
        ' Verifica se a data de liberação está dentro da data limite do centro de custo
        If DataLimiteCentroCusto(CLngDef(txtDuplicatas(15).Text), strData) Then
            GoTo DuplVerifique_Erro
        End If
    End If
       
    'Pagamento
    'pt. 86132 - Ivo Sousa (26/03/2008)
    'Retirado o tratamento de validação de datas(Dias Uteis)
    If txtDuplicatas(8).Text <> "" Then         'Se o usuário indicou o pagamento
        If IsDate(txtDuplicatas(8).Text) Then
            'Se a data de Pagamento não é anterior a emissão
            If (DateDiff("d", txtDuplicatas(6).Text, txtDuplicatas(8).Text) < 0) Then
                MsgFunc ResolveResString(139, resUM, "de Pagamento", resDOIS, "de Emissão"), vbInformation
                GoTo DuplVerifique_Erro
            End If
        Else
            MsgBox "Informe uma data de pagamento válida.", vbInformation, NomeModulo
            txtDuplicatas(8).SetFocus
            GoTo DuplVerifique_Erro
        End If
    Else    'Se o campo Cheque estiver preenchido não deixa Pagamento passar em Branco
        If mstrPagRec = "P" And IsValid(txtDuplicatas(16).Text) Then
            MsgBox ResolveResString(23, resUM, "Pagamento"), vbInformation, MsgBoxCaption
            GoTo DuplVerifique_Erro
        End If
    End If
    'Exibe a mensagem caso a Data de Pagamento seja posteior a Data de Vencimento
    If EData(txtDuplicatas(8).Text) And EData(txtDuplicatas(7).Text) Then
        If CDateDef(txtDuplicatas(8).Text) > CDateDef(txtDuplicatas(7).Text) Then
            If Not CDateDef(txtDuplicatas(11).Text) > 0 Or CDateDef(txtDuplicatas(12).Text) > 0 Then
                MsgFunc "A Data de Pagamento informada está em atraso há " & _
                DateDiff("d", CDateDef(txtDuplicatas(7).Text), CDateDef(txtDuplicatas(8).Text)) & " dia(s)." & _
                vbCrLf & "Informe 'Acréscimo' ou 'Multa' se necessário."
            End If
        End If
    End If
  
    If Not IsValid(GetValue(mrstDuplicatas, "Pagamento", NUL)) And IsValid(txtDuplicatas(8).Text) Then
        If Not IsValid(txtDuplicatas(13).Text) Then
            MsgFunc "O Campo Banco deverá ser preenchido"
            GoTo DuplVerifique_Erro
        End If
    End If
  
    'Liberação
    'pt. 86132 - Ivo Sousa (26/03/2008)
    'Retirado o tratamento de validação de datas(Dias Uteis)
    If EData(txtDuplicatas(9).Text) Then    'Se for uma data válida
        If EData(txtDuplicatas(8).Text) Then    'Se o usuário preencheu o campo Pagamento
            If DateDiff("d", txtDuplicatas(8).Text, txtDuplicatas(9).Text) < 0 Then
                'A data de Liberação não pode ser menor que a data de Pagamento
                MsgFunc ResolveResString(139, resUM, "de Liberação", resDOIS, "de Pagamento"), vbInformation
                GoTo DuplVerifique_Erro
            End If
        End If
    End If
  
    'Verificando se o Banco indicado existe no cadastro de Bancos
    If IsValid(txtDuplicatas(13).Text) Then
        If Len(lblDuplDesc(1).Caption) = 0 Then
            If MsgBox(ResolveResString(35, resUM, txtDuplicatas(13).Text, resDOIS, "Bancos"), _
                vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                LibProc "Bancos"
            End If
            GoTo DuplVerifique_Erro
        End If
    Else
        'Se não há número de banco o usuário não pode especificar um
        'número de cheque.
        If mstrPagRec = "P" And IsValid(txtDuplicatas(16).Text) Then
            MsgFunc LoadResString(249)
            GoTo DuplVerifique_Erro
        End If
    End If
  
    'BANCO - Verificando se Carteira existe no Cadastro  de Carteiras
    If IsValid(txtDuplicatas(13).Text) Then
        If IsValid(txtDuplicatas(27).Text) Then
            If Recordcount("SELECT Carteira From Carteiras WHERE Banco=" & CLngDef(txtDuplicatas(13).Text) & " AND Carteira=" & Quote(txtDuplicatas(27).Text, "'")) = 0 Then
                MsgFunc " Carteira não cadastrada no Banco " & txtDuplicatas(13).Text
                GoTo DuplVerifique_Erro
            End If
        End If
    Else
        MsgBox "O campo 'Banco' não pode ser zero", vbCritical, MsgBoxCaption
        GoTo DuplVerifique_Erro
    End If
  
    'CONTA - Verificando se a Conta indicada existe no cadastro de Contas Contábeis
    If IsValid(txtDuplicatas(14).Text) Then
        If Len(lblDuplDesc(2).Caption) = 0 Then
            If MsgBox(ResolveResString(35, resUM, txtDuplicatas(14).Text, resDOIS, "Contas"), _
                vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                LibProc "Contas"
            End If
            GoTo DuplVerifique_Erro
        End If
    Else
        MsgBox "O campo 'Conta' não pode ser zero", vbCritical, MsgBoxCaption
        GoTo DuplVerifique_Erro
    End If
  
    'Verificar se a conta está ativa ou nao
    If GetFieldValue("Ctaati", "Contas", " [Código]=" & txtDuplicatas(14).Text) = "N" Then
        MsgBox "Conta " & txtDuplicatas(14).Text & " não está ativa", vbCritical, MsgBoxCaption
        txtDuplicatas(14).SetFocus
        GoTo DuplVerifique_Erro
    End If
  
    ' Verificando se o Código de Centro de Custo existe no Cadastro
    If txtDuplicatas(15).Visible Then
        If IsValid(txtDuplicatas(15).Text) Then
            If Len(lblDuplDesc(3).Caption) = 0 Then
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
  
    'Verificando se a Moeda indicada existe no cadastro de Moedas e Índices
    If Len(txtDuplicatas(17).Text) > 0 Then
        If ConfereDuplicidade("Moeda", "Moedas", "Moeda = '" & txtDuplicatas(17).Text & "'") = 0 Then
            If MsgBox(ResolveResString(35, resUM, txtDuplicatas(17).Text, resDOIS, "Moedas & Índices"), _
                vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                LibProc "Moedas"
            End If
            GoTo DuplVerifique_Erro
        End If
    End If
  
    'Protocolo 81289
    'Antigamente validava-se EMPRESA somente para as duplicatas
    'Devemos validar para lançamentos também.
    If mstrTabela = "Duplicatas" Or mstrTabela = "Lançamentos" Then
        If IsValid(txtDuplicatas(2).Text) Then
            If Recordcount("SELECT Razão, Apel FROM Empresas WHERE Apel = '" & txtDuplicatas(2).Text & "'") = 0 Then
                If MsgBox(ResolveResString(35, "|1", txtDuplicatas(2).Text, "|2", "Empresas"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                    LibProc "Empresas"
                End If
                GoTo DuplVerifique_Erro
            End If
        End If
    End If
  
    ' Verifica se não há datas diferentes para o cheque cadastrado agora
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
        'Verificando se o tipo da duplicata digitado é um novo tipo
        If Recordcount("SELECT TIPO FROM [TIPOS GLOBAIS] WHERE TIPO = '" & cboDuplicatas(3).Text & "'") = 0 Then
            MsgBox "Tipo global informado não cadastrado!", vbInformation
            cboDuplicatas(3).SetFocus
            GoTo DuplVerifique_Erro
        End If
    End If
  
    'Verificação do campo Forma de Pagamento
    If IsNumeric(txtDuplicatas(18).Text) Then
        If CInt(txtDuplicatas(18).Text) > 0 Then
            If Len(lblDuplDesc(13).Caption) = 0 Then
                MsgBox "Forma de pagamento não encontrada.", vbInformation, "Validação de Campos"
                txtDuplicatas(18).SetFocus
                GoTo DuplVerifique_Erro
            End If
        End If
    End If
  
    If IsDate(txtDuplicatas(9).Text) Then
        'Se a data de pagamento estiver informada
        If IsDate(txtDuplicatas(8).Text) Then
            If CDate(txtDuplicatas(9).Text) < CDate(txtDuplicatas(8).Text) Then
                MsgBox "A data de liberação deve ser maior do que a data de pagamento.", vbInformation
                txtDuplicatas(9).SetFocus
                GoTo DuplVerifique_Erro
            End If
        Else
            If CDate(txtDuplicatas(9).Text) < CDate(txtDuplicatas(7).Text) Then
                MsgBox "A data de liberação deve ser maior do que a data de vencimento do documento.", vbInformation
                txtDuplicatas(9).SetFocus
                GoTo DuplVerifique_Erro
            End If
        End If
    End If
    
    'pt. 89506 - Dulcino Júnior (29/10/2008)
    If strToLng(txtDuplicatas(4).Text) = 0 Then
        MsgBox "O campo parcela deve ser preenchido.", vbInformation, NomeModulo
        txtDuplicatas(4).SetFocus
        GoTo DuplVerifique_Erro
    End If
    
    'pt. 86728 - Moacir Pfau(09/06/2008)
    If Not fEmpresaBloqueada(txtDuplicatas(2).Text, CDate(txtDuplicatas(6).Text)) Then
       GoTo DuplVerifique_Erro
    End If
    
    'pt. 98848 - Ivo Sousa(24/11/2010)
    If EAdicao(mlngDuplicatas) Then
        If ExisteLancamento(mstrPagRec, txtDuplicatas(1).Text, txtDuplicatas(4).Text) Then
            If MsgBox("O Lançamento número " & txtDuplicatas(1).Text & " Parcela " & txtDuplicatas(4).Text & " já existe. Deseja que o sistema sugira o próximo número?", vbYesNo + vbInformation, NomeModulo) = vbYes Then
                txtDuplicatas(1).Text = ProximoNumero("Código", "Lançamentos", IIf(SeqLancamentos, NUL, "PagRec = '" & mstrPagRec & "'"))
                txtDuplicatas(4).Text = "1"
            Else
                GoTo DuplVerifique_Erro
            End If
        End If
    End If
    DuplVerifique = True

DuplVerifique_Erro:
    SetPtr vbDefault
End Function

'SUB.......: NovoRegistro
'Objetivo..: Configura alguns controles como adição de registro.
'Argumento.: [blnProcChave]: Quando a rotina deve procurar a nova chave.
Private Sub NovoRegistro(blnProcChave As Boolean)
    If blnProcChave Then
        If CompStr(mstrTabela, "Duplicatas") Then
            If mstrPagRec = "R" Then
                txtDuplicatas(1).Text = ProximoNumero("Nota", "Duplicatas", "Tipo = '" & cboDuplicatas(3).Text & "' AND PagRec = '" & mstrPagRec & "'")
            End If
        Else
            txtDuplicatas(1).Text = ProximoNumero("Código", "Lançamentos", IIf(SeqLancamentos, NUL, "PagRec = '" & mstrPagRec & "'"))
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
        MsgBar IIf(IsValid(txtDuplicatas(1).Text), IIf(mstrTabela = "Duplicatas", " Nota: " & txtDuplicatas(1).Text, " Lançamento: " & txtDuplicatas(1).Text), " ") _
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
            strProcura = "SELECT Razão, Apel FROM Empresas WHERE Apel = '" & txtDuplicatas(2).Text & "';"
            GetAssocValue strProcura, lblDuplDesc(0)
            If (mstrTabela = "Lançamentos") Then
                'Empresa não faz parte da chave em Lançamentos
                AlteraValor mlngDuplicatas
            End If
            'Pt. 94752 - Moacir Pfau(21/10/2009)
            Call fPreenche_CodCobranca
        Case 6 To 9 ' Campos: Data de Emissão, Vencimento, Pagamento, Liberação
            lblDuplDesc(Index - 1).Caption = Semana(txtDuplicatas(Index).Text, raUmaPalavra)
        
        Case 10 To 12 ' Valores atualizam Total
            ExibeSoma
        
        Case 13 ' Campo Banco
            strProcura = "SELECT Nome FROM Bancos WHERE Banco = " & txtDuplicatas(13).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(1)
            txtBancoCheque.Text = txtDuplicatas(13).Text
        
        Case 14, 25 ' Campo Conta
            strProcura = "SELECT Descrição FROM Contas WHERE Código = " & txtDuplicatas(Index).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(IIf(Index = 14, 2, 11))
        
        Case 15, 20 ' Campo Centro de Custo
            strProcura = "SELECT Descrição FROM Centros WHERE Código = " & txtDuplicatas(Index).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(IIf(Index = 20, 10, 3))
        
        Case 16 ' Campo número do cheque
            txtChequeCheque.Text = txtDuplicatas(16).Text
        
        Case 17 ' Campo Moeda
            strProcura = "SELECT Descrição, Moeda FROM Moedas WHERE Moeda = '" & txtDuplicatas(17).Text & "';"
            GetAssocValue strProcura, Nothing, txtDuplicatas(17)
        
        Case 18 'Forma de Pagamento
            strProcura = "SELECT Nome FROM [Formas de Pagamento] WHERE Código = " & txtDuplicatas(18).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(13)
        
        Case 40 'Operação contábil
            If Len(txtDuplicatas(Index).Text) > 0 Then
                lblDuplDesc(14).Caption = GetFieldValue("descricao", "OperacaoContabil", "cd_operacao = " & txtDuplicatas(Index).Text)
            Else
                lblDuplDesc(14).Caption = vbNullString
            End If
        
        Case 41 'Operacao Contábil de Baixa
            If Len(txtDuplicatas(Index).Text) > 0 Then
                lblDuplDesc(15).Caption = GetFieldValue("descricao", "OperacaoContabil", "cd_operacao = " & txtDuplicatas(Index).Text)
            Else
                lblDuplDesc(15).Caption = vbNullString
            End If
        Case 43 ' Campo Carteira
            strProcura = "SELECT desc_carteira FROM FFICarteira WHERE id_carteira = " & txtDuplicatas(43).Text & ";"
            GetAssocValue strProcura, lblDuplDesc(16)
    End Select
    If mstrTabela = "Lançamentos" Then
        If Len(txtDuplicatas(1).Text) = 0 Or Len(txtDuplicatas(4).Text) = 0 Then Exit Sub
    End If
    If Index > 4 Then
        If Index <> 40 And Index <> 45 Then
            AlteraValor mlngDuplicatas
        Else
            'pt. 83525 - Dulcino Júnior (27/09/2007)
            If mblnAlteraValor Then
                AlteraValor mlngDuplicatas
            End If
        End If
    End If
End Sub

'Pt. 94752 - Moacir Pfau(21/10/2009)
Private Sub fPreenche_CodCobranca()
    Dim strSql                  As String
    Dim rstTab                  As Object
    
    etxCobrancaCep.Clear: etxCobrancaCidade.Clear: etxCobrancaEstado.Clear: etxCobrancaEndereco.Clear: etxCobrancaBairro.Clear
    If CStr(txtDuplicatas(2).Text) <> "" And val(txtDuplicatas(44).Text) > 0 Then
        strSql = "SELECT [Endereço],Bairro,CEP,Cidade,Estado,Apel,[Código],Tipo FROM [Empresas Endereços] WHERE Tipo = 'Cobrança' AND Apel = '" & txtDuplicatas(2).Text & "'"
        If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
            etxCobrancaCep.valorTexto = GetValue(rstTab, "Cep")
            etxCobrancaCidade.valorTexto = GetValue(rstTab, "Cidade")
            etxCobrancaEstado.valorTexto = GetValue(rstTab, "Estado")
            etxCobrancaEndereco.valorTexto = GetValue(rstTab, "Endereço")
            etxCobrancaBairro.valorTexto = GetValue(rstTab, "Bairro")
        End If
        FechaRecordset (rstTab)
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
        If PMultiCampo("Selecione o endereço", "SELECT [Endereço],Bairro,CEP,Cidade,Estado,Apel,[Código],Tipo FROM [Empresas Endereços] WHERE Tipo = 'Cobrança' AND Apel = '" & txtDuplicatas(2).Text & "'", pbCampo, "Apel;Código;Tipo;Endereço;Bairro;CEP;Cidade;Estado", strApel, lngCodigo, strTipo, strEndereco, strBairro, strCep, strCidade, strEstado) Then
            etxCobrancaCep.valorTexto = strCep
            etxCobrancaCidade.valorTexto = strCidade
            etxCobrancaEstado.valorTexto = strEstado
            etxCobrancaEndereco.valorTexto = strEndereco
            etxCobrancaBairro.valorTexto = strBairro
            txtDuplicatas(44).Text = lngCodigo
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
            strMensagem = ResolveResString(75, resUM, "Moedas e Índices")
        
        Case "Obs"
            'Posiciona no segundo tab
            SSTab1.Tab = 1
        
        Case "CheBan"
            SSTab1.Tab = 2
        
        Case Else  ' Qualquer outro campo
            strMensagem = NUL
    End Select
    Selecione txtDuplicatas(Index)
    If IsValid(txtDuplicatas(2).Text) Then
        MsgBar IIf(IsValid(txtDuplicatas(1).Text), IIf(mstrTabela = "Duplicatas", " Nota: " & txtDuplicatas(1).Text, " Lançamento: " & txtDuplicatas(1).Text), " ") _
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
            If MsgBox("Esse título já possui boleto processado. Se continuar a linha digitável do boleto será zerada e não será possível processar o Retorno bancário. Continuar?", vbYesNo, "Confirmação") = vbYes Then
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
        If Index = 1 Then 'Em Lançamentos apenas Código é chave
            ControlaChave KeyCode, Shift, txtDuplicatas(1), mlngDuplicatas
        End If
    End If
    If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
        Select Case Index
            Case 2 ' Empresa
                
                strPCampo = "SELECT Apel, Razão, Pessoa, Tipo, [CNPJ/CPF], [IEst/RG], CCM, " _
                & "Ramo, Endereço, Bairro, CEP, Cidade, Estado, " _
                & "Região, País, Fone1, Ramal1, Contato, Dpto " _
                & "FROM Empresas"
                ' Verifica a configuração para separar as empresas por tipo
                If (mstrPagRec = "P") Then
                    AppendStr strPCampo, " WHERE Tipo <> '" & GetResOptions(1003, 2) & "';" 'Cliente
                Else
                    AppendStr strPCampo, " WHERE Tipo <> '" & GetResOptions(1003, 1) & "';" 'Fornecedor
                End If
                PCampo "Empresas", strPCampo, PB_CAMPO, txtDuplicatas(2), 0
            
            Case 13 ' Banco
                PCampo "Bancos", "Bancos", PB_CAMPO, txtDuplicatas(13), 0
            
            Case 14, 25 ' Conta
                'pt. 83864 - Dulcino Júnior (11/10/2007)
                PCampo "Contas", "SELECT Contas.Código as Conta, Contas.Descrição as [Descrição da Conta], Grupos.Código as Grupo, Grupos.Descrição as [Descrição do Grupo] " & _
                       " FROM Grupos INNER JOIN Contas ON Grupos.Código = Contas.Grupo where Contas.Ctaati='S' " & _
                       " ORDER BY Grupos.Código,Contas.Código", PB_CAMPO, txtDuplicatas(Index), 0
            
            Case 15, 20 ' Centro de Custo
                PCampo "Centro de Custo", "Centros", PB_CAMPO, txtDuplicatas(Index), 0
            
            Case 16 ' Campo Cheque
                If IsValid(txtDuplicatas(13).Text) Then
                    PCampo "Cheques", "SELECT * FROM Cheque WHERE Banco = " & txtDuplicatas(13).Text & ";", _
                    PB_CAMPO, txtDuplicatas(16), 1
                Else
                    PCampo "Cheques", "Cheque", PB_CAMPO, txtDuplicatas(16), 1
                End If
            
            Case 17 ' Moeda
                PCampo "Moedas e Índices", "Moedas", PB_CAMPO, txtDuplicatas(17), 0
            
            Case 18 'Forma de Pagamento
                PCampo "Formas de Pagamento", "SELECT * FROM [Formas de Pagamento]", PB_CAMPO, txtDuplicatas(18), "Código"
            
            Case 27 'Carteira
                If IsValid(txtDuplicatas(13).Text) Then
                    PCampo "Carteiras", "Select Carteira from Carteiras WHERE Banco=" & CLngDef(txtDuplicatas(13).Text), PB_CAMPO, txtDuplicatas(27), 0
                End If
            
            Case 40 'Operação Contábil
                PCampo "Operações Contabeis", "OperacaoContabil", pbCampo, txtDuplicatas(40), "cd_operacao"
            
            Case 41 'Operação Contábil Baixa
                PCampo "Operações Contabeis", "OperacaoContabil", pbCampo, txtDuplicatas(41), "cd_operacao"
            Case 44
                'Pt. 94752 - Moacir Pfau(21/10/2009)
                Call fLocaliza_CodCobranca
        End Select
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
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask("Contas", "Código")
        
        Case 15, 20 ' Campo Centro de Custo
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask("Centros", "Código")
        
        Case 16 ' Campo Cheque
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, fMask("Cheques", "Cheque")
        
        Case 4 ' Campo Parcela
            SetMascara KeyAscii, txtDuplicatas(4).SelStart, "###"
        
        Case 6 To 9 ' Campos Emissão, Vencimento, Pagamento e Liberação
            SetMascara KeyAscii, txtDuplicatas(Index).SelStart, MASK_DATE4
        
        Case 10 To 12, 18, 21, 24, 29, 30 ' Campos Valor Original, Acréscimo, Abatimento, Valor em Moeda
            If Index <= 12 Then
                ValidaNaoAceitaPonto KeyAscii
            End If
            DMoeda KeyAscii
        
        Case 22, 31, 36, 37, 38   ' 31 Valor do desconto por pontualide, 36 Valor Multa, 37 Percentual de Multa, 38 Valor Juros de Mora Diário
            DValor KeyAscii
        
        Case 40, 41 'Campo Operação contábil
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
    Dim datLiberacao                As Date
    'pt.81487 Ivo Sousa (25/10/07)
    Dim strProcura                  As String
    Dim rstOpContabil               As Object
    Dim rstBanco                    As Object
    'Pt. 88817 - Moacir Pfau(06/11/2009)
    Dim dblPerDespesaFinanceira     As Double
 
    If Index = 2 Then
        txtDuplicatas(Index).Text = FormataEmpresa(txtDuplicatas(Index).Text)
        strProcura = "SELECT Razão, Apel FROM Empresas WHERE Apel = '" & txtDuplicatas(2).Text & "';"
        GetAssocValue strProcura, lblDuplDesc(0), txtDuplicatas(2)
        'pt. 79561 - Moacir Pfau(04/04/2008)
        If EAdicao(mlngDuplicatas) Then
            If strToLng(txtDuplicatas(13).Text) = 0 And strToLng(txtDuplicatas(14).Text) = 0 Then 'pt. 87031 - Moacir Pfau(21/05/2008)
                strProcura = "SELECT Banco, Conta FROM Empresas WHERE Apel = '" & txtDuplicatas(2).Text & "';"
                AbreRecordset rstBanco, strProcura
                txtDuplicatas(13).Text = strToLng(GetValue(rstBanco, "Banco"))
                txtDuplicatas(14).Text = strToLng(GetValue(rstBanco, "Conta"))
                FechaRecordset (rstBanco)
            End If
        End If
    End If

    If mstrTabela = "Duplicatas" Then
        If Index < 5 Then
            'Pt. 95023 - Moacir Pfau(21/09/2009)
            If lngOperacao = 0 Then
                lngOperacao = val(txtDuplicatas(40).Text)
            End If
            LibProc WL_EXIBIR
            txtDuplicatas(40).Text = lngOperacao
        End If
    Else                                  'No caso de Lançamentos
        If Index = 4 Then
            If Trim(txtDuplicatas(1).Text) <> "" Then
                'Não remover em hipotese nenhuma
                'Dulcino Junior e Edilberto - 06/10/2006
                If mlngDuplicatas = 736 Then
                    mlngDuplicatas = 737
                End If
                If lngOperacao = 0 Then
                    lngOperacao = txtDuplicatas(40).Text
                End If
                LibProc WL_EXIBIR
                If Not EAddNew(mlngDuplicatas) Then
                    txtDuplicatas(40).Text = GetValue(mrstDuplicatas, "cd_operacao_contabil", "0")
                Else
                    txtDuplicatas(40).Text = lngOperacao
                End If
            End If
        End If
    End If

    Select Case Index
        'Percentual de Multa
        Case 37
            If val(lblDuplDesc(4).Caption) > 0 Then
                txtDuplicatas(36).Text = Round(lblDuplDesc(4).Caption * (txtDuplicatas(37).Text / 100), 2)
            Else
                txtDuplicatas(36).Text = 0
            End If
        'Valor da Multa
        Case 36
            If val(lblDuplDesc(4).Caption) > 0 Then
                txtDuplicatas(37).Text = Round(txtDuplicatas(36).Text * 100 / lblDuplDesc(4).Caption, 2)
            Else
                txtDuplicatas(37).Text = 0
            End If
           
        Case 38
            If val(lblDuplDesc(4).Caption) > 0 Then
                txtPercMora.Text = Round(txtDuplicatas(38).Text * 100 / lblDuplDesc(4).Caption, 2)
            Else
                txtPercMora.Text = 0
            End If
           
        Case 7 'Data de Vencimento
            If IsDate(txtDuplicatas(Index).Text) Then
                If UCase(mstrPagRec) = "R" Then
                    datLiberacao = CDate(txtDuplicatas(Index).Text)
                    'pt. 88523 - Ivo Sousa (24/09/2008)
                    datLiberacao = DateAdd("d", DiasLiberacao, datLiberacao)
                    If calendario.PermiteLancamento(datLiberacao, , False) <> "A" Then
                        datLiberacao = datLiberacao + NumeroDiasUteisNaoUteis(datLiberacao, 0)
                    End If
                Else
                    datLiberacao = CDate(txtDuplicatas(Index).Text)
                End If
                txtDuplicatas(9).Text = datLiberacao
            End If
        Case 8 'Data de Pagamento
            If IsDate(txtDuplicatas(Index).Text) Then
                datLiberacao = CDate(txtDuplicatas(Index).Text)
                'pt. 88523 - Ivo Sousa (24/09/2008)
                If UCase(mstrPagRec) = "R" Then
                    datLiberacao = DateAdd("d", DiasLiberacao, datLiberacao)
                    If calendario.PermiteLancamento(datLiberacao, , False) <> "A" Then
                        datLiberacao = datLiberacao + NumeroDiasUteisNaoUteis(datLiberacao, 0)
                    End If
                End If
                txtDuplicatas(9).Text = datLiberacao
            End If
            Call SugestaoOperacaoContabilBaixa   'Operação contábil de baixa
        Case 39 'pt. 87031 - Moacir Pfau(21/05/2008)
            SSTab1.Tab = 0
        Case 10, 18
            'Pt. 88817 - Moacir Pfau(06/11/2009)
            dblPerDespesaFinanceira = 0
            If IsNumeric(txtDuplicatas(11).Text) Then
                If val(txtDuplicatas(18).Text) > 0 And txtDuplicatas(11).Text = 0 Then
                    dblPerDespesaFinanceira = GetFieldValue("per_despesa_financeira", "[Formas de Pagamento]", "Código=" & txtDuplicatas(18).Text)
                    txtDuplicatas(11).Text = Format(txtDuplicatas(11).Text + (txtDuplicatas(10).Text * dblPerDespesaFinanceira / 100), "#,#0.#0")
                End If
            End If
    End Select
End Sub

'SUB: ExibeSoma
'Soma o valor original com os Acréscimos e diminui os Abatimentos.
'Exibe o resultado no label do formulário.
Private Sub ExibeSoma()
    Dim curResult As Currency
  
    'Protocolo 76736
    'Alteração alternativa visando solucionar um problema que ocorre
    'apenas no cliente LaserFlex (via terminal server).
    'Digitava 900 no valor original, aparecia 90000,00 no Total
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
'Objetivo..: Exibe a janela de cálculo do valor da duplicata.
Private Sub CalcValor()
    Dim cVlrOriginal As Currency
    Dim cAumento As Currency

    'Verifica se o usuário já preencheu o Valor Original
    If Not IsValid(txtDuplicatas(10).Text) Then Exit Sub
    'Verifica se a data de pagamento foi preenchida e se é diferente de zero
    If IsEmptyDate(txtDuplicatas(8).Text) Then
        MsgFunc ResolveResString(26, resUM, txtDuplicatas(8).Text)
    Else
        'Verifica se a data de vencimento foi preenchida
        If IsEmptyDate(txtDuplicatas(7).Text) Then
            MsgFunc ResolveResString(26, resUM, txtDuplicatas(7).Text)
        Else
            'Verifica se a data de pagamento é posterior a data de vencimento
            If DateDiff(DD_DIA, txtDuplicatas(7).Text, txtDuplicatas(8).Text) > ZERO Then
                'Chama a função que exibe a janela de cálculo e aguarda
                cVlrOriginal = CMoeda(txtDuplicatas(10).Text)
                cAumento = CMoeda(txtDuplicatas(11).Text)
'                If CValorFinal(cVlrOriginal, cAumento, CDate(txtDuplicatas(7).Text), CDate(txtDuplicatas(8).Text)) Then
'                    'Retornando o valor já calculado
'                    lblDuplDesc(4).Caption = Format$(cVlrOriginal, FMOEDA)
'                    txtDuplicatas(11).Text = Format$(cAumento, FMOEDA)
'                End If
            End If
        End If
    End If
End Sub

'SUB.......: ChequeInfo
'Objetivo..: Exibe informações do cheque para o usuário
'Argumentos: [sFuncao]: O mesmo argumento sFuncao da função LibProc
'            [nBco   ]: Opcional. Código do Banco.
'            [nChq   ]: Opcional. Número do Cheque.
Private Sub ChequeInfo(sFuncao As String, Optional nBco As Long, Optional nChq As Long)
    Dim strCheque     As String
    Dim lngCheque     As Long
    Dim lngBanco      As Long
    Dim cValor        As Currency

    If mstrPagRec = "P" Then
        'Somente se for pagamento
        Select Case sFuncao
            Case WL_NOVO
                Call LimpaControles(mrstCheques, Me, TAG_CHEQUE, mlngCheques, True)
            
            Case WL_SALVAR
                If CBool(nBco) And CBool(nChq) Then
                    'Verifica se o cheque anterior é diferente do atual
                    'Se for, verifico se existem registro com o cheque
                    'anterior, caso não exista mais nenhum o cheque deve ser
                    'excluído da tabela de cheques.
                    If nBco <> GetValue(mrstDuplicatas, "Banco", 0) Or nChq <> GetValue(mrstDuplicatas, "Cheque", 0) Then
                        If (ExisteCheque(nBco, nChq) = ZERO) Then
                            DeleteAll "Cheque", wsprintf("Banco = %l AND Cheque = %l", nBco, nChq)
                        End If
                    Else
                        'Caso contrário apenas chama a função salva registro para
                        'gravar eventuais alterações nos campos Nominal e Histórico
                        Call SalvaRegistro(mrstCheques, Me, TAG_CHEQUE, mlngCheques)
                    End If
                End If
                'Verifica se o cheque atual existe na tabela de Cheques, se não
                'existir acrescenta-o.
                nBco = GetValue(mrstDuplicatas, "Banco", ZERO)
                nChq = GetValue(mrstDuplicatas, "Cheque", ZERO)
                If CBool(nBco) And CBool(nChq) Then
                    strCheque = wsprintf("FROM Cheque WHERE Banco = %l AND Cheque = %l", nBco, nChq)
                    If Recordcount(strCheque) = 0 Then
                        strCheque = "INSERT INTO Cheque (Banco, Cheque, Nominal, Histórico) " & wsprintf("VALUES (%l, %l, \'%s\', \'%s\');", nBco, nChq, txtCheque(0).Text, txtCheque(1).Text)
                        Call ExecuteSQL(strCheque)
                    End If
                End If
                Call ChequeInfo("ExibeRegistro")
            
            Case WL_CANCELAR
                Call CancelaEdicao(mrstCheques, Me, TAG_CHEQUE, mlngCheques)
            
            Case WL_DELETAR
                If CBool(nBco) And CBool(nChq) Then
                    If ExisteCheque(nBco, nChq) = ZERO Then
                        DeleteAll "Cheque", wsprintf("Banco = %l AND Cheque = %l", nBco, nChq)
                    End If
                End If
                Call ChequeInfo("ExibeRegistro")
            
            Case Else
                Call SalvaRegistro(mrstCheques, Me, TAG_CHEQUE, mlngCheques)
                lngBanco = CLngDef(txtDuplicatas(13).Text)
                lngCheque = CLngDef(txtDuplicatas(16).Text)
                strCheque = wsprintf("SELECT * FROM Cheque WHERE " & "Banco = %l AND Cheque = %l", lngBanco, lngCheque)
                If AbreRecordset(mrstCheques, strCheque) = WL_OK Then
                    Call ExibeRegistro(mrstCheques, Me, TAG_CHEQUE, mlngCheques)
                Else
                    Call LimpaControles(mrstCheques, Me, TAG_CHEQUE, mlngCheques, True)
                End If
                If sFuncao <> WL_SAIR Then
                    lvwLancamentos.ListItems.Clear    'Limpa o conteúdo atual do ListView
                    If (IsVisibleRecord(mlngCheques)) Then 'Se há um cheque visível agora
                        SetPtrWait Me
                        If gTipoDB = Access Then
                            wvsprintf strCheque, "SELECT FORMAT(Nota, \'000000\') & ' - ' & " & "FORMAT(Parcela, \'00\') AS Cod, Tipo, Empresa, " & "FORMAT(([Valor Original] + Acréscimo - Abatimento), " & "\'###,###,###,##0.00\') AS Total FROM Duplicatas WHERE PagRec = " & "'P' AND Banco = %l AND Cheque = %l;", lngBanco, lngCheque
                        Else
                            wvsprintf strCheque, "SELECT (Nota +  ' - ' & " & "Parcela) AS Cod, Tipo, Empresa, " & "([Valor Original] + Acréscimo - Abatimento) " & " AS Total FROM Duplicatas WHERE PagRec = " & "'P' AND Banco = %l AND Cheque = %l;", lngBanco, lngCheque
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
'Objetivo..: Marca com um X o ítem selecionado pelo usuário quando esta não está
'            marcado, ou desmarca quando este estiver marcado.
'Argumento.: [lngIndice]: Índico do ítem que deve ser marcado ou desmarcado.
Private Sub XMark(lngIndice As Long)
    If lngIndice > 0 Then
        If lvwRateio.ListItems(lngIndice).SmallIcon = DL_MARCADO Then
            lvwRateio.ListItems(lngIndice).SmallIcon = DL_DESMARCADO
        Else
            lvwRateio.ListItems(lngIndice).SmallIcon = DL_MARCADO
        End If
    End If
End Sub

Private Sub txtPercMora_LostFocus()
    If val(lblDuplDesc(4).Caption) > 0 Then
        txtDuplicatas(38).Text = Round(lblDuplDesc(4).Caption * (txtPercMora.Text / 100), 2)
    Else
        txtDuplicatas(38).Text = 0
    End If
End Sub

'Data.......: 18/05/2007
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para a verficação do preenchimento dos campos
'               de operação contábil dos lançamentos. ref pt 81902
'Retorno....: [Boolean] Retorna se o registro pode ou não ser gravado.
Private Function validaIntegracaoLancamentos() As Boolean
    validaIntegracaoLancamentos = True
    If Not IsEmptyDate(txtDuplicatas(8).Text) Then
        If Trim(txtDuplicatas(41).Text) = "0" Or Trim(txtDuplicatas(41).Text) = "" Then
            MsgBox "Para lançamentos quitados é necessário informar a Operação de Baixa!", vbInformation, NomeModulo
            txtDuplicatas(41).SetFocus
            validaIntegracaoLancamentos = False
        Else
            If Not IsEmptyDate(txtDuplicatas(6).Text) Then
                If CDate(txtDuplicatas(6).Text) <> CDate(txtDuplicatas(8).Text) Then
                    If Trim(txtDuplicatas(40).Text) = "0" Or Trim(txtDuplicatas(40).Text) = "" Then
                        MsgBox "Para lançamentos é necessário informar a Operação Contábil", vbInformation, NomeModulo
                        txtDuplicatas(40).SetFocus
                        validaIntegracaoLancamentos = False
                    End If
                Else
                    If Trim(txtDuplicatas(40).Text) <> "0" And Trim(txtDuplicatas(40).Text) <> "" Then
                        MsgBox "Para movimentos bancários a operação de emissão não pode ser informada.", vbInformation, NomeModulo
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
            MsgBox "Para Lançamentos é necessário informar a Operação Contábil", vbInformation, NomeModulo
            txtDuplicatas(40).SetFocus
            validaIntegracaoLancamentos = False
        Else
            If txtDuplicatas(41).Text <> "" And txtDuplicatas(41).Text <> "0" Then
                MsgBox "Para lançamentos com operação contábil de baixa é necessário informar a data de pagamento.", vbInformation, NomeModulo
                txtDuplicatas(8).SetFocus
                validaIntegracaoLancamentos = False
            End If
        End If
    End If
End Function

'Data.......: 18/05/2007
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para a verficação do preenchimento dos campos
'               de operação contábil dos lançamentos. ref pt 81902
'Retorno....: [Boolean] Retorna se o registro pode ou não ser gravado.
Private Function validaIntegracaoDuplicatas() As Boolean
    validaIntegracaoDuplicatas = True
    If Not IsEmptyDate(txtDuplicatas(8).Text) Then
        If Trim(txtDuplicatas(41).Text) = "0" Or Trim(txtDuplicatas(41).Text) = "" Then
            MsgBox "Para Duplicatas quitadas é necessário informar a Operação de Baixa!", vbInformation, "Validação de Campos"
            txtDuplicatas(41).SetFocus
            validaIntegracaoDuplicatas = False
        End If
    Else
        If txtDuplicatas(41).Text <> "" And txtDuplicatas(41).Text <> "0" Then
            'pt. 82355 Ivo Sousa (24/10/07)
            MsgBox "Para Duplicatas com operação contábil de baixa é necessário informar a data de pagamento.", vbInformation, NomeModulo
            txtDuplicatas(8).SetFocus
            validaIntegracaoDuplicatas = False
        End If
    End If
    If Not IsEmptyDate(txtDuplicatas(6).Text) Then
        If Trim(txtDuplicatas(40).Text) = "0" Or Trim(txtDuplicatas(40).Text) = "" Then
            MsgBox "Para Duplicatas é necessário informar a Operação Contábil", vbInformation, "Validação de Campos"
            txtDuplicatas(40).SetFocus
            validaIntegracaoDuplicatas = False
        End If
    End If
End Function

'Data.......: 23/05/2007
'Autor......: Dulcino Júnior
'Descrição..: A função verifica a existência de rateio para o lançamento que está carregado na tela,
'               conforme os registros da tabela FFIRateioLancamento, se existe registro nessa tabela
'               será buscado o registro que originou o rateio e fazer a verificação se o mesmo está
'               quitado, caso não esteja, será permitida a exclusão do mesmo, e o valor referente ao
'               titulo excluido será retornado para o titulo que originou o rateio, do contrário o
'               sistema vai avisar ao usuário qual o titulo que originou o rateio e dizer que o
'               mesmo está quitado.
'Retorno....: [Boolean] Retorna se a duplicata pode ou não ser excluida.
Private Function PermiteExclusao() As Boolean
    Dim strSql     As String
    Dim rstResult  As Object

    PermiteExclusao = True
    If chkRateio.value = vbChecked Then
        strSql = "SELECT pag_rec_origem, cd_lancamento_origem, nr_parcela_origem FROM FFIRateioLancamento"
        strSql = strSql & " WHERE pag_rec_destino='" & mstrPagRec & "' AND cd_lancamento_destino=" & txtDuplicatas(1).Text
        strSql = strSql & " AND nr_parcela_destino=" & txtDuplicatas(4).Text
        If AbreRecordset(rstResult, strSql) = WL_OK Then
            strSql = "SELECT Pagamento FROM Lançamentos WHERE PagRec='" & rstResult.Fields("pag_rec_origem").value
            strSql = strSql & "' AND Código=" & rstResult.Fields("cd_lancamento_origem").value & " AND Parcela="
            strSql = strSql & rstResult.Fields("nr_parcela_origem").value
            mlngCodigo = rstResult.Fields("cd_lancamento_origem").value
            mlngPARCELA = rstResult.Fields("nr_parcela_origem").value
        Else
            strSql = ""
        End If
        Call FechaRecordset(rstResult)
        If strSql <> "" Then
            If AbreRecordset(rstResult, strSql) = WL_OK Then
                If IsEmptyDate(rstResult.Fields("Pagamento").value) Then
                    mstrOrigem = "UPDATE Lançamentos SET [Valor Original]=[Valor Original]+" & Replace(txtDuplicatas(10).Text, ",", ".")
                    mstrOrigem = mstrOrigem & " WHERE PagRec='" & mstrPagRec & "' AND Código=" & mlngCodigo & " AND "
                    mstrOrigem = mstrOrigem & "Parcela=" & mlngPARCELA
                    
                    mstrDelete = "DELETE FROM FFIRateioLancamento WHERE pag_rec_destino='" & mstrPagRec & "' AND "
                    mstrDelete = mstrDelete & "cd_lancamento_destino=" & txtDuplicatas(1).Text & " AND "
                    mstrDelete = mstrDelete & "nr_parcela_destino=" & txtDuplicatas(4).Text
                    
                    mstrRateio = "SELECT cd_lancamento_destino FROM FFIRateioLancamento WHERE cd_lancamento_origem=" & mlngCodigo
                    mstrRateio = mstrRateio & " AND nr_parcela_origem=" & mlngPARCELA
                    PermiteExclusao = True
                Else
                    MsgBox "Não é possivel excluir a parcela por que a parcela de origem do rateio já está quitada.", vbInformation, NomeModulo
                    mstrOrigem = ""
                    mstrDelete = ""
                    PermiteExclusao = False
                End If
            End If
            Call FechaRecordset(rstResult)
        Else
            mstrOrigem = ""
            mstrDelete = ""
            PermiteExclusao = True
        End If
    End If
End Function

'Data.......: 30/05/2007
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para sugerir a operação contábil
'               de acordo com o tipo global da duplicata ou lançamento.
Public Sub SugestaoOperacaoContabilBaixa()
    Dim DAOMatriz   As cMatrizContabilizacaoDAO
    Dim matriz      As cMatrizContabilizacao
    Dim lngOperacao As Long
    
    If IsDate(txtDuplicatas(8).Text) Then
        Set DAOMatriz = New cMatrizContabilizacaoDAO
        Set matriz = DAOMatriz.Carregar(cboDuplicatas(3).Text)
        If Not matriz Is Nothing Then
            If mstrTabela = "Lançamentos" Then
                If mstrPagRec = "P" Then
                    lngOperacao = matriz.BaixaLancamentosPagar
                Else
                    lngOperacao = matriz.baixaLancamentosReceber
                End If
            Else
                If mstrPagRec = "P" Then
                    lngOperacao = matriz.BaixaDuplicatasPagar
                Else
                    lngOperacao = matriz.baixaLancamentosReceber
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
'Autor......: Dulcino Júnior
'Descrição..: Função utilizada para retornar a quantidade de dias
'               que o banco possui para a liberação da duplicata.
'Retorno....: [Integer] Número de dias para liberação do pagamento.
Private Function DiasLiberacao() As Double
    Dim selCmd   As IDBSelectCommand
    Dim rdResult As IDBReader
    
    If IsNumeric(txtDuplicatas(13).Text) Then
        If CLng(txtDuplicatas(13).Text) > 0 Then
            Aplicacao.Connect
            Set selCmd = Aplicacao.CreateSelectCommand
            With selCmd
                .SelectClause = "[Dias para Liberação]"
                
                .Table.TableName = "Bancos"
                
                Call .Filter.Append("Banco = @pNumero")
                Call .Parameters.add(.CreateParameter("@pNumero", CLng(txtDuplicatas(13).Text), dbFieldTypeLong))
            End With
            Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
            If Not rdResult.EOF Then
                DiasLiberacao = rdResult.GetDouble("Dias para Liberação")
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
'Descrição..: Função utilizada para Validação de datas como feriados, domingos
'             sabados ou se o periodo esta bloqueado.
'Retorno....: [boolean] Se a data é valida
Private Function ValidaDatas() As Boolean
    Dim intIndexTXT As Integer
    Dim intIndexLBL As Integer
    Dim strSinal    As String
    
    intIndexTXT = 6
    intIndexLBL = 14
    While intIndexTXT <= 9
        If txtDuplicatas(intIndexTXT).Text <> "" Then
            strSinal = calendario.PermiteLancamento(txtDuplicatas(intIndexTXT).Text)
            If strSinal = "X" Then
                MsgBox "O movimento esta bloqueado para a data Informada no campo " & Replace(lblDuplicatas(intIndexLBL).Caption, ":", ""), vbOKOnly + vbInformation, NomeModulo
                txtDuplicatas(intIndexTXT).SetFocus
                ValidaDatas = False
                Exit Function
            ElseIf strSinal = "A" Then
                ValidaDatas = True
            ElseIf strSinal = "F" Or strSinal = "S" Or strSinal = "D" Then
                If MsgBox("A data Informada no campo " & Replace(lblDuplicatas(intIndexLBL).Caption, ":", "") & " não é um dia útil." & vbNewLine & _
                "Deseja salvar o lançamento assim mesmo?", vbYesNo + vbInformation, NomeModulo) = vbYes Then
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
'Descrição..: Utilizado para carregar os campos na tela, metodo Paliativo.
'Protocolo..: 86140
Private Function CarregaPadrao()
    cboDuplicatas.item(3).Text = "Fatura"               'Tipo
    txtDuplicatas(4).Text = "0"                         'Parcela
    txtDuplicatas(18).Text = "0"                        'Forma Pagto
    txtDuplicatas(13).Text = "0"                        'Banco
    txtDuplicatas(14).Text = "0"                        'Conta
    txtDuplicatas(40).Text = "0"                        'Op. Cantabil
    cboDuplicatas.item(20).Text = "Normal"              'Situação
    txtDuplicatas(10).Text = "0"                        'Valor Original
    txtDuplicatas(11).Text = "0"                        'Acréscimo
    txtDuplicatas(12).Text = "0"                        'Abatimento
    txtDuplicatas(6).Text = Format(Date, "DD/MM/YYYY")  'Emissão
    txtDuplicatas(7).Text = Format(Date, "DD/MM/YYYY")  'Vencimento
    txtDuplicatas(9).Text = Format(Date, "DD/MM/YYYY")  'Liberação
    txtDuplicatas(41).Text = "0"                        'Liberação
    lblDuplDesc(4).Caption = "0,00"                     'Total
End Function

'Data.......: 10/10/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para carregar as informações da tabela tipos globais
'               para o campo de seleção da tela.
Private Sub CarregaTipos()
    Dim strSql    As String
    Dim rstResult As Object
    
    'Dulcino Júnior
    'Conforme orientação do Carlos dias devem ser exibidos todos os registros da tabela Tipos Globais
    strSql = "SELECT Tipo FROM [Tipos Globais] ORDER BY Tipo"  'WHERE Empresa='" & EmpresaUsuaria.Fantasia & "'"
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        rstResult.MoveFirst
        While Not rstResult.EOF
            cboDuplicatas(3).AddItem rstResult.Fields("Tipo").value
            rstResult.MoveNext
        Wend
    End If
    Call FechaRecordset(rstResult)
End Sub

'Data.......: 10/10/2008
'Autor......: Dulcino Júnior
'Descrição..: Procedimento utilizado para parametrizar a List view de acordo com
'               as informações necessárias para exibição na tela.
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

'Pt. 114146 - Moacir Pfau(29/02/2012)
Private Sub TotalizaValorRateio()
    txtDuplicatas(45).Text = Format(txtDuplicatas(10).Text - SomaValores(), FMOEDA)
End Sub
