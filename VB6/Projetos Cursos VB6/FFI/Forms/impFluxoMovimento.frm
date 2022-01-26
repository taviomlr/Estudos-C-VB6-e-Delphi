VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form fimpFluxoMovimento 
   Caption         =   "Movimento de Caixa"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "impFluxoMovimento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "fimpFluxoMovimento"
   ScaleHeight     =   116.681
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   180.181
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection RodapeSaldo 
      Align           =   1  'Align Top
      Height          =   1335
      Left            =   0
      Top             =   2955
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2355
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   24
         Left            =   4440
         TabIndex        =   0
         Top             =   120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   370
         Caption         =   "A Transportar Totais do Dia:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   25
         Left            =   6960
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Total Entrada"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   26
         Left            =   8400
         TabIndex        =   2
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Total Saída"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   62
         Left            =   4440
         TabIndex        =   3
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   370
         Caption         =   "Saldo Atual:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   63
         Left            =   8400
         TabIndex        =   4
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Total Saldo"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   7
         Left            =   9840
         TabIndex        =   39
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Saldo Anterior"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   11
         Left            =   4440
         TabIndex        =   40
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   370
         Caption         =   "Saldo Anterior:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   12
         Left            =   6960
         TabIndex        =   41
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Total Saldo Anterior"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   13
         Left            =   4440
         TabIndex        =   42
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   370
         Caption         =   "Totais:"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   14
         Left            =   6960
         TabIndex        =   43
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Total Entrada + Saldo Anterior"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   18
         Left            =   8400
         TabIndex        =   44
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Total Saída + Total Saldo"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Line lneFluxo 
         BorderStyle     =   4  'Dash-Dot
         Index           =   0
         X1              =   120
         X2              =   11400
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line lneFluxo 
         BorderStyle     =   4  'Dash-Dot
         Index           =   3
         X1              =   120
         X2              =   11400
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection QuebraSaldo 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      Top             =   1575
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1296
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   370
         Caption         =   "Banco:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "Banco"
         Caption         =   ""
         Formula         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   370
         Campo           =   "Nome"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   3
         Left            =   7920
         TabIndex        =   8
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   370
         Caption         =   "Saldo Anterior:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   4
         Left            =   9480
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Campo           =   "Saldo Anterior"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Shape shpFluxo 
         Height          =   495
         Left            =   120
         Top             =   120
         Width           =   11175
      End
   End
   Begin ReportX.ReportSection RodapeCabecalho 
      Align           =   1  'Align Top
      Height          =   150
      Left            =   0
      Top             =   2805
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   265
      Tipo            =   5
      Ordem           =   2
   End
   Begin ReportX.ReportSection Detalhe 
      Align           =   1  'Align Top
      Height          =   225
      Left            =   0
      Top             =   2580
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   397
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   28
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   370
         Campo           =   "Duplicata"
         Caption         =   ""
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   29
         Left            =   1560
         TabIndex        =   11
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   370
         Campo           =   "Tipo"
         Caption         =   ""
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   30
         Left            =   3120
         TabIndex        =   12
         Top             =   0
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   370
         Campo           =   "Descrição"
         Caption         =   ""
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   35
         Left            =   2310
         TabIndex        =   13
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   370
         Campo           =   "Pagamento"
         Caption         =   ""
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   36
         Left            =   6960
         TabIndex        =   14
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Entrada"
         Formato         =   "Standard"
         Caption         =   ""
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   37
         Left            =   8400
         TabIndex        =   15
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Saída"
         Formato         =   "Standard"
         Caption         =   ""
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   38
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Caption         =   "Descrição:"
         Mostrar         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   39
         Left            =   2520
         TabIndex        =   17
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   370
         Campo           =   "Descrição"
         Caption         =   ""
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   40
         Left            =   6480
         TabIndex        =   18
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Caption         =   "Controle:"
         Mostrar         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   41
         Left            =   7560
         TabIndex        =   19
         Top             =   480
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   370
         Campo           =   "Controle"
         Caption         =   ""
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   42
         Left            =   1500
         TabIndex        =   20
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Caption         =   "Razão:"
         Mostrar         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   43
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   370
         Campo           =   "Razão"
         Caption         =   ""
         Formula         =   -1  'True
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   6
         Left            =   9840
         TabIndex        =   38
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Saldo"
         Formato         =   "Standard"
         Caption         =   ""
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
   End
   Begin ReportX.ReportSection QuebraCabecalho 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   2310
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   476
      Tipo            =   3
      Ordem           =   2
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   370
         Caption         =   "Código"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   9
         Left            =   1560
         TabIndex        =   23
         Top             =   0
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   370
         Caption         =   "Tipo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   10
         Left            =   3120
         TabIndex        =   24
         Top             =   0
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   370
         Caption         =   "Descrição"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   15
         Left            =   2310
         TabIndex        =   25
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   370
         Caption         =   "Pagto."
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   16
         Left            =   6960
         TabIndex        =   26
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Caption         =   "Entrada"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   17
         Left            =   8400
         TabIndex        =   27
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Caption         =   "Saída"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   5
         Left            =   9840
         TabIndex        =   37
         Top             =   0
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Caption         =   "Saldo"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
   End
   Begin ReportX.ReportMain Imprimir 
      Height          =   480
      Left            =   360
      TabIndex        =   28
      Top             =   5640
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   5
      Regua           =   -1  'True
      MargemEsquerda  =   5
      MargemDireita   =   5
      MargemSuperior  =   5
      MargemInferior  =   5
      Titulo          =   ""
      HTMLUnico       =   -1  'True
   End
   Begin ReportX.ReportSection Titulo 
      Align           =   1  'Align Top
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   2778
      Tipo            =   2
      Begin ReportX.ReportField Campo 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   29
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
         Left            =   480
         TabIndex        =   30
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
         TabIndex        =   31
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
         Left            =   2160
         TabIndex        =   32
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
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
         TabIndex        =   33
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
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
         TabIndex        =   34
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   2160
         TabIndex        =   35
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
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
         Height          =   240
         Index           =   7
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   423
         Campo           =   "Title"
         Caption         =   "Titulo"
         Formula         =   -1  'True
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin VB.Shape shpHeader 
         Height          =   1215
         Left            =   120
         Top             =   120
         Width           =   11175
      End
   End
End
Attribute VB_Name = "fimpFluxoMovimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private QuebraBanco          As Boolean
Private SaldoInicial         As Double
Private TotalEntradaBanco    As Double
Private TotalSaidaBanco      As Double
Private SaldoFinal           As Double
Private TabelaSaldo          As String
Private TituloRel            As String
Private Banco                As Long
Private dblSaldoAnterior     As Double

Public Sub Config(rstDados As Object, strTitulo As String, blnQuebraBanco As Boolean, strTabelaSaldos As String)

On Error GoTo err_Handler
    QuebraBanco = blnQuebraBanco
    TabelaSaldo = strTabelaSaldos
    
    If Not QuebraBanco Then
        cmpFluxo(0).Mostrar = False
        cmpFluxo(1).Mostrar = False
        cmpFluxo(2).Mostrar = False
        cmpFluxo(24).Caption = "Total:"
        cmpFluxo(62).Caption = "Saldo:"
    End If
    
    TituloRel = strTitulo
    Set Imprimir.Recordset = rstDados
    Imprimir.Ativar
    Unload Me
    Exit Sub
    
err_Handler:
    MsgBox "Erro ao configurar o relatório: " & err.Description, vbInformation, NomeModulo
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set fimpFluxoMovimento = Nothing
End Sub

Private Sub Imprimir_Erro(ByVal Numero As Long)
  RpxMsgErro Numero
End Sub

Private Sub Imprimir_FormulaCampo(ByVal Campo As String, Valor As Variant)
  Valor = FormulasHeader(Campo, Imprimir, TituloRel)

  Select Case Campo
    Case "Duplicata"
      If GetValue(Imprimir.Recordset, Campo, ZERO) <> 0 Then
        Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000")
      End If
    
    Case "Descrição"
      If GetValue(Imprimir.Recordset, Campo, ZERO) <> 0 Then
        Valor = GetValue(Imprimir.Recordset, Campo, ZERO)
      Else
        'pt. 85283 - Moacir Pfau(26/01/2009)
        'Valor = "Saldo Anterior"
      End If
    
    Case "Banco"
      Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000000")
      Banco = GetValue(Imprimir.Recordset, Campo, ZERO)
  
    Case "Pagamento"
      Valor = Format(GetValue(Imprimir.Recordset, Campo), "dd/mm/yy")
  
    Case "Total Entrada"
      Valor = TotalEntradaBanco

    Case "Total Saída"
      Valor = TotalSaidaBanco
      
    Case "Total Saldo"
      Valor = SaldoInicial + TotalEntradaBanco - TotalSaidaBanco

    Case "Total Entrada + Saldo Anterior", "Total Saída + Total Saldo"
      Valor = TotalEntradaBanco + SaldoInicial

    Case "Saldo Anterior", "Total Saldo Anterior"
      If QuebraBanco Then
        Valor = dblSaldoAnterior
      Else
        Valor = GetFieldValue("Valor", TabelaSaldo, "Tipo = false", , ZERO)
      End If
      SaldoInicial = Valor
      SaldoFinal = SaldoInicial
    
    Case "Saldo"
      ' 01/04/2020 - HyperCube: INC-14012 - Yuji F. - Ajuste para não alterar a variável global
      Valor = SaldoFinal + (GetValue(Imprimir.Recordset, "Entrada", ZERO) - GetValue(Imprimir.Recordset, "Saída", ZERO))
      
  End Select
End Sub

Private Sub Imprimir_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
  If Ordem = 1 Then
    If QuebraBanco Then
      Valor = GetValue(Imprimir.Recordset, "Banco", ZERO)
    End If
  End If
End Sub

Private Sub Imprimir_ImprimiuRegistro(Cancelar As Boolean)
  TotalEntradaBanco = TotalEntradaBanco + GetValue(Imprimir.Recordset, "Entrada", ZERO)
  TotalSaidaBanco = TotalSaidaBanco + GetValue(Imprimir.Recordset, "Saída", ZERO)
  
  ' 01/04/2020 - HyperCube: INC-14012 - Yuji F. - Alterando variável global após a impressão do registro
  SaldoFinal = SaldoFinal + (GetValue(Imprimir.Recordset, "Entrada", ZERO) - GetValue(Imprimir.Recordset, "Saída", ZERO))
End Sub

Private Sub Imprimir_IniciarGrupo(ByVal Ordem As Byte)
  If Ordem = 1 Then
    TotalEntradaBanco = 0
    TotalSaidaBanco = 0
    dblSaldoAnterior = GetFieldValue("Valor", TabelaSaldo, "Banco = " & GetValue(Imprimir.Recordset, "Banco", ZERO) & " and Tipo = false", , ZERO)
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
