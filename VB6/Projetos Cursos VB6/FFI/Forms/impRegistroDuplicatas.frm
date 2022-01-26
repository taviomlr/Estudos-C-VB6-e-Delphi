VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form fimpRegistroDuplicatas 
   AutoRedraw      =   -1  'True
   Caption         =   "Relatório Registro de Duplicatas"
   ClientHeight    =   6405
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "impRegistroDuplicatas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   112.977
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   209.55
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection4 
      Align           =   1  'Align Top
      Height          =   585
      Left            =   0
      Top             =   4365
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1032
      Tipo            =   6
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   23
         Left            =   4740
         TabIndex        =   53
         Top             =   150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   370
         Campo           =   "TotalGeralNF"
         Formato         =   "Standard"
         Caption         =   "ReportField1"
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   24
         Left            =   9630
         TabIndex        =   54
         Top             =   180
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   370
         Campo           =   "TotalGeralDUP"
         Formato         =   "Standard"
         Caption         =   "ReportField1"
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Geral..............................................................................:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   20
         Left            =   300
         TabIndex        =   52
         Top             =   180
         Width           =   4335
      End
   End
   Begin ReportX.ReportSection rpsDetalhe 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      Top             =   3510
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   10
         Left            =   300
         TabIndex        =   38
         Top             =   30
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   370
         Campo           =   "Apel"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   11
         Left            =   3420
         TabIndex        =   39
         Top             =   30
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   370
         Campo           =   "Número"
         Formato         =   "000000"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   12
         Left            =   4260
         TabIndex        =   40
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "Emissão"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   13
         Left            =   5310
         TabIndex        =   41
         Top             =   30
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   370
         Campo           =   "Valor Total"
         Formato         =   "Standard"
         Caption         =   "ReportField1"
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   14
         Left            =   6480
         TabIndex        =   42
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "Duplicata"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   15
         Left            =   7380
         TabIndex        =   43
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "Emissão"
         Caption         =   "ReportField1"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   16
         Left            =   8280
         TabIndex        =   44
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "Vencimento"
         Caption         =   "ReportField1"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   17
         Left            =   9240
         TabIndex        =   45
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   370
         Campo           =   "Pagamento"
         Caption         =   "ReportField1"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   18
         Left            =   10200
         TabIndex        =   46
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   370
         Campo           =   "Valor Original"
         Formato         =   "Standard"
         Caption         =   "ReportField1"
         TipoCampo       =   1
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   19
         Left            =   300
         TabIndex        =   47
         Top             =   270
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   370
         Campo           =   "Cnpj/Cpf"
         Caption         =   "ReportField1"
         MostrarSeRepetir=   0   'False
         Mostrar         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   20
         Left            =   2010
         TabIndex        =   48
         Top             =   270
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   370
         Campo           =   "Cidade"
         Caption         =   "ReportField1"
         MostrarSeRepetir=   0   'False
         Mostrar         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
   End
   Begin ReportX.ReportSection rpsRodGrupo 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      Top             =   3795
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1005
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   21
         Left            =   5070
         TabIndex        =   50
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   370
         Campo           =   "TotalNF"
         Formato         =   "Standard"
         Caption         =   "ReportField1"
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   22
         Left            =   10200
         TabIndex        =   51
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   370
         Campo           =   "TotalDUP"
         Formato         =   "Standard"
         Caption         =   "ReportField1"
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Line linBotton 
         BorderStyle     =   3  'Dot
         X1              =   240
         X2              =   11520
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   300
         TabIndex        =   49
         Top             =   240
         Width           =   345
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      Top             =   3225
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      Tipo            =   3
      Ordem           =   1
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da Duplicata"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   18
         Left            =   10170
         TabIndex        =   37
         Top             =   60
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagamento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   17
         Left            =   9270
         TabIndex        =   36
         Top             =   60
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   16
         Left            =   8250
         TabIndex        =   35
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   15
         Left            =   7470
         TabIndex        =   34
         Top             =   60
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Duplicata"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   14
         Left            =   6510
         TabIndex        =   33
         Top             =   60
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da Nota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   13
         Left            =   5280
         TabIndex        =   32
         Top             =   60
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emissão"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   4350
         TabIndex        =   31
         Top             =   60
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   11
         Left            =   3360
         TabIndex        =   30
         Top             =   60
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   10
         Left            =   300
         TabIndex        =   29
         Top             =   60
         Width           =   630
      End
   End
   Begin ReportX.ReportMain REL 
      Height          =   480
      Left            =   570
      TabIndex        =   0
      Top             =   5820
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
   Begin ReportX.ReportSection Titulo 
      Align           =   1  'Align Top
      Height          =   3225
      Left            =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   5689
      Tipo            =   2
      Begin ReportX.ReportField Campo 
         Height          =   270
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
         Campo           =   "Data"
         Formato         =   "Short Date"
         Caption         =   "Data"
         TipoCampo       =   2
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Campo 
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   423
         Campo           =   "Hora"
         Formato         =   "hh:nn"
         Caption         =   "Hora"
         TipoCampo       =   3
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Campo 
         Height          =   225
         Index           =   2
         Left            =   9720
         TabIndex        =   3
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   397
         Caption         =   "Página:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Campo 
         Height          =   270
         Index           =   6
         Left            =   2280
         TabIndex        =   4
         Top             =   600
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   476
         Campo           =   "Sistema"
         Caption         =   "Sistema"
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Campo 
         Height          =   225
         Index           =   3
         Left            =   10440
         TabIndex        =   5
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         Campo           =   "Página"
         Caption         =   "Página"
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Campo 
         Height          =   240
         Index           =   4
         Left            =   9720
         TabIndex        =   6
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   423
         Campo           =   "User"
         Caption         =   "User"
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Campo 
         Height          =   285
         Index           =   5
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   503
         Campo           =   "Empresa"
         Caption         =   "Empresa"
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField Campo 
         Height          =   450
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   794
         Linhas          =   2
         Campo           =   "Title"
         Caption         =   "Titulo"
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   0
         Left            =   1350
         TabIndex        =   9
         Top             =   1710
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   370
         Campo           =   "EMPRESAsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   1
         Left            =   1350
         TabIndex        =   15
         Top             =   1980
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   370
         Campo           =   "ENDEREÇOsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   2
         Left            =   6300
         TabIndex        =   18
         Top             =   1680
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   370
         Campo           =   "CNPJsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   3
         Left            =   9600
         TabIndex        =   19
         Top             =   1710
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   370
         Campo           =   "IEsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   4
         Left            =   1350
         TabIndex        =   20
         Top             =   2250
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   370
         Campo           =   "BAIRROsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   5
         Left            =   1350
         TabIndex        =   21
         Top             =   2520
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   370
         Campo           =   "CIDADEsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   6
         Left            =   1350
         TabIndex        =   22
         Top             =   2790
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   370
         Campo           =   "SITEsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   7
         Left            =   6300
         TabIndex        =   26
         Top             =   2250
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   370
         Campo           =   "CEPsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   8
         Left            =   6300
         TabIndex        =   27
         Top             =   2520
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   370
         Campo           =   "UFsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rpfREL 
         Height          =   210
         Index           =   9
         Left            =   6300
         TabIndex        =   28
         Top             =   2790
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   370
         Campo           =   "EMAILsys"
         Caption         =   "ReportField1"
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   240
         X2              =   11520
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   5580
         TabIndex        =   25
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF......:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   5580
         TabIndex        =   24
         Top             =   2490
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEP...:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   5580
         TabIndex        =   23
         Top             =   2220
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insc.Estadual..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   8160
         TabIndex        =   17
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   5580
         TabIndex        =   16
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Site........:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   300
         TabIndex        =   14
         Top             =   2760
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade....:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   300
         TabIndex        =   13
         Top             =   2490
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro......:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   300
         TabIndex        =   12
         Top             =   2220
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   300
         TabIndex        =   11
         Top             =   1950
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   1680
         Width           =   960
      End
      Begin VB.Shape shpHeader 
         Height          =   1365
         Index           =   0
         Left            =   240
         Top             =   120
         Width           =   11295
      End
   End
End
Attribute VB_Name = "fimpRegistroDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bRelCompleto    As Boolean
Dim bDiretoImp      As Boolean
Dim nTotNF          As Currency
Dim nTotDUP         As Currency
Dim nTotGeralNF     As Currency
Dim nTotGeralDUP    As Currency
Dim sTitulo         As String
Dim NF              As String
Dim Apel            As String

Sub Config(rst As Object, sTipo As String, Destino As Long, sPeriodo As String, bTotalizaEmpresa As Boolean)

    Set REL.Recordset = rst
    
    ' Verifico se é pra mandar direto pra impressora
    bDiretoImp = IIf(Destino = 0, False, True)
    
    ' Se o relatório for completo exibo alguns controles a mais e altero o tamanho
    ' da seção detalhe
    bRelCompleto = IIf(sTipo = "Completo", True, False)
    
    ' Título do relatório
    sTitulo = "Registro de Duplicatas Ref. LEI Nº 5474/68" & vbCrLf & " Período : " & sPeriodo
    
    ' Se não for totalizar por empresa, reconfiguro o rodapé de grupo
    If bTotalizaEmpresa = False Then
        linBotton.Y1 = 180
        linBotton.Y2 = 180
        lblTotal.Visible = False
        rpfREL(21).Mostrar = False
        rpfREL(22).Mostrar = False
        rpsRodGrupo.Height = 3
    End If
    
    ' se o relatório for completo, exibo alguns campos a mais
    If bRelCompleto Then
        rpfREL(19).Mostrar = True
        rpfREL(20).Mostrar = True
        rpsDetalhe.Height = 10
    Else
        rpfREL(19).Mostrar = False
        rpfREL(20).Mostrar = False
        rpsDetalhe.Height = 5
    End If
    
    REL.Ativar
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fimpRegistroDuplicatas = Nothing
End Sub

Private Sub REL_Erro(ByVal Numero As Long)
    RpxMsgErro Numero
End Sub

Private Sub REL_FormulaCampo(ByVal Campo As String, Valor As Variant)
    
    Valor = FormulasHeader(Campo, REL, sTitulo)
    
    Select Case Campo
        Case "EMPRESAsys": Valor = GetFieldValue("empresa", "kinsys", "")
        Case "CNPJsys": Valor = GetFieldValue("cnpj", "kinsys", "")
        Case "IEsys": Valor = GetFieldValue("[Inscrição estadual]", "kinsys", "")
        Case "ENDEREÇOsys": Valor = GetFieldValue("endereço", "kinsys", "")
        Case "BAIRROsys": Valor = GetFieldValue("bairro", "kinsys", "")
        Case "CEPsys": Valor = GetFieldValue("Cep", "kinsys", "")
        Case "CIDADEsys": Valor = GetFieldValue("Cidade", "kinsys", "")
        Case "UFsys": Valor = GetFieldValue("Estado", "kinsys", "")
        Case "SITEsys": Valor = GetFieldValue("Site", "kinsys", "")
        Case "EMAILsys": Valor = GetFieldValue("[e-mail]", "kinsys", "")
        Case "Duplicata": Valor = Format(CStr(GetValue(REL.Recordset, "Número", ZERO)), "000000") & _
                        "-" & Left(GetValue(REL.Recordset, "Tipo", NUL), 1) & " " & _
                        GetValue(REL.Recordset, "Parcela", ZERO)
        Case "TotalNF": Valor = nTotNF
        Case "TotalDUP": Valor = nTotDUP
        Case "TotalGeralNF": Valor = nTotGeralNF
        Case "TotalGeralDUP": Valor = nTotGeralDUP
        Case "Apel", "Número", "Emissão", "Valor Total"
            Valor = GetValue(REL.Recordset, Campo, NUL)
            
    End Select
End Sub

Private Sub REL_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
    Valor = GetValue(REL.Recordset, "Apel", NUL)
End Sub

Private Sub REL_ImprimiuRegistro(Cancelar As Boolean)

    ' Só incremento os totalizadores se for uma NF diferente da última
    ' totalizadores do grupo
    If NF <> REL.Recordset("Número") Or Apel <> REL.Recordset("Apel") Then
        nTotNF = nTotNF + REL.Recordset("Valor Total")          'Valor da NF
        
        ' pego a NF e o Apel atual
        NF = REL.Recordset("Número")
        Apel = REL.Recordset("Apel")
        
        ' totalizadores do relatório
        'Projeto: 44895 - Desenv.: 47104 - Ueder Budni (22/08/2014)
        nTotGeralNF = nTotGeralNF + REL.Recordset("Valor Total")
        
    End If
    
    nTotDUP = nTotDUP + REL.Recordset("Valor Original")     'Valor da Duplicata
    
    nTotGeralDUP = nTotGeralDUP + REL.Recordset("Valor Original")
End Sub

Private Sub REL_IniciarGrupo(ByVal Ordem As Byte)
    ' Zero os totalizadores de grupo
    nTotNF = 0
    nTotDUP = 0
End Sub

Private Sub REL_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    
    ' Verifico se é pra mandar direto para a impressora
    If bDiretoImp Then
        If Impressora Then
            Cancelar = False
        Else
            Cancelar = True
        End If
    End If
    
    ' Zero os totalizadores do relatório
    nTotGeralNF = 0
    nTotGeralDUP = 0

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
