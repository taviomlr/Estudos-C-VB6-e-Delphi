VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.3#0"; "ReportX.ocx"
Begin VB.Form fimpFluxoSintetico 
   KeyPreview      =   -1  'True
   Caption         =   "Fluxo de Caixa Sintético"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   ScaleHeight     =   151.606
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   220.398
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection RodapeMes 
      Align           =   1  'Align Top
      Height          =   585
      Left            =   0
      Top             =   2820
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1032
      Tipo            =   5
      Ordem           =   3
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   23
         Left            =   9150
         TabIndex        =   34
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Campo           =   "Total Saldo Mês"
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   24
         Left            =   3180
         TabIndex        =   35
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   370
         Caption         =   "Saldo mensal:"
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
         Left            =   5520
         TabIndex        =   36
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Campo           =   "Total Entrada Mês"
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   26
         Left            =   7320
         TabIndex        =   37
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Campo           =   "Total Saída Mês"
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
   Begin ReportX.ReportSection CebecalhoMes 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      Top             =   4395
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   -26
      Tipo            =   3
      Ordem           =   3
   End
   Begin ReportX.ReportSection Rodape 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   3900
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   873
      Tipo            =   7
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   19
         Left            =   5520
         TabIndex        =   30
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Index           =   20
         Left            =   7320
         TabIndex        =   31
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Index           =   21
         Left            =   9120
         TabIndex        =   32
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Index           =   22
         Left            =   3840
         TabIndex        =   33
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "TOTAL GERAL:"
         TipoCampo       =   1
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
   Begin ReportX.ReportSection RodapeSaldo 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   3405
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   873
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   15
         Left            =   5520
         TabIndex        =   26
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Campo           =   "Total Entrada Banco"
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   16
         Left            =   7320
         TabIndex        =   27
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Campo           =   "Total Saída Banco"
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   17
         Left            =   9120
         TabIndex        =   28
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Campo           =   "Total Saldo Banco"
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
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   18
         Left            =   3180
         TabIndex        =   29
         Top             =   120
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   370
         Formato         =   "Standard"
         Caption         =   "Total do Banco:"
         TipoCampo       =   1
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
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   120
         Top             =   120
         Width           =   11055
      End
   End
   Begin ReportX.ReportSection RodapeCabecalho 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      Top             =   2805
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   26
      Tipo            =   5
      Ordem           =   2
   End
   Begin ReportX.ReportSection QuebraSaldo 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      Top             =   1575
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1296
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   0
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
         TabIndex        =   1
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
         TabIndex        =   2
         Top             =   240
         Width           =   5715
         _ExtentX        =   10081
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
         Left            =   8100
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         TabIndex        =   4
         Top             =   240
         Width           =   1605
         _ExtentX        =   2831
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
      Begin VB.Shape Shape2 
         Height          =   495
         Left            =   120
         Top             =   120
         Width           =   11055
      End
   End
   Begin ReportX.ReportSection Detalhe 
      Align           =   1  'Align Top
      Height          =   225
      Left            =   0
      Top             =   2580
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   397
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   36
         Left            =   5520
         TabIndex        =   5
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
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
         Left            =   7320
         TabIndex        =   6
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
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
         Index           =   5
         Left            =   9120
         TabIndex        =   22
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Campo           =   "Saldo"
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
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   0
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   13
         Left            =   1320
         TabIndex        =   24
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   370
         Campo           =   "Nome"
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
         Index           =   14
         Left            =   4320
         TabIndex        =   25
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "Data"
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
   End
   Begin ReportX.ReportSection QuebraCabecalho 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   2310
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   476
      Tipo            =   3
      Ordem           =   2
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Caption         =   "Banco"
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
         Index           =   8
         Left            =   1320
         TabIndex        =   8
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   370
         Caption         =   "Nome"
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
         Left            =   4320
         TabIndex        =   9
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Caption         =   "Data"
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
         Left            =   5520
         TabIndex        =   10
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Caption         =   "Total de Entradas"
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
         Left            =   7320
         TabIndex        =   11
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Caption         =   "Total de Saídas"
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
         Left            =   9120
         TabIndex        =   12
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   370
         Caption         =   "Saldo Final"
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
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
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
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2778
      Tipo            =   2
      Begin ReportX.ReportField Campo 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   476
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
         TabIndex        =   15
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
         Left            =   9510
         TabIndex        =   16
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
         Left            =   1950
         TabIndex        =   17
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
         Left            =   10230
         TabIndex        =   18
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   9510
         TabIndex        =   19
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
         Left            =   1920
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   960
         Width           =   10845
         _ExtentX        =   19129
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
         Width           =   11055
      End
   End
End
Attribute VB_Name = "fimpFluxoSintetico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private QuebraBanco          As Boolean
Private SaldoInicial         As Double
Private EntradaBanco         As Double
Private SaidaBanco           As Double
Private SaldoBanco           As Double
Private EntradaMes           As Double
Private SaidaMes             As Double
Private SaldoMes             As Double
Private Entrada              As Double
Private Saida                As Double
Private Saldo                As Double
Private TabelaSaldo          As String
Private TituloRel            As String
Public Sub Config(rstDados As Object, strTitulo As String, blnQuebraBanco As Boolean, strTabelaSaldos As String)

  QuebraBanco = blnQuebraBanco
  TabelaSaldo = strTabelaSaldos
  
  If Not QuebraBanco Then
    cmpFluxo(0).Mostrar = False
    cmpFluxo(1).Mostrar = False
    cmpFluxo(2).Mostrar = False
    cmpFluxo(18).Caption = "Total do período:"
    Rodape.Mostrar = False
  Else
    cmpFluxo(7).Mostrar = False
    cmpFluxo(8).Mostrar = False
    cmpFluxo(6).Mostrar = False
    cmpFluxo(13).Mostrar = False
    RodapeMes.Mostrar = False
  End If
   
  Campo(0).Caption = Date
   
  Saldo = 0
  TituloRel = strTitulo
    
  Set Imprimir.Recordset = rstDados
   
  Imprimir.Ativar
  
  Unload Me
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set fimpFluxoSintetico = Nothing
End Sub
Private Sub Imprimir_Erro(ByVal Numero As Long)
  RpxMsgErro Numero
End Sub
Private Sub Imprimir_FormulaCampo(ByVal Campo As String, Valor As Variant)
  Valor = FormulasHeader(Campo, Imprimir, TituloRel)

  Select Case Campo
    Case "Banco"
      Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000000")
  
    Case "Saldo Anterior"
      If QuebraBanco Then
        Valor = GetFieldValue("Valor", TabelaSaldo, "Banco = " & GetValue(Imprimir.Recordset, "Banco", ZERO), , ZERO)
      Else
        Valor = GetFieldValue("Valor", TabelaSaldo, NUL, , ZERO)
      End If
      'Autor:Leandro Mesquita
      'Data: 26/06/2007
      'pt. 82286
      'SaldoInicial = Valor
      'Saldo = Saldo + SaldoInicial
      
    Case "Total Entrada Banco"
      Valor = EntradaBanco
      
    Case "Total Saída Banco"
      Valor = SaidaBanco
      
    Case "Total Saldo Banco"
      Valor = SaldoBanco
          
    Case "Total Entrada Mês"
      Valor = EntradaMes
    
    Case "Total Saída Mês"
      Valor = SaidaMes
          
    Case "Total Saldo Mês"
      Valor = SaldoMes
          
    Case "Total Entrada"
      Valor = Entrada
      
    Case "Total Saída"
      Valor = Saida
      
    Case "Total Saldo"
      Valor = Saldo
  End Select
End Sub
Private Sub Imprimir_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
  If Ordem = 1 Then
    If QuebraBanco Then
      Valor = GetValue(Imprimir.Recordset, "Banco", ZERO)
    End If
  End If
  If Ordem = 3 Then
    If Not QuebraBanco Then
       Valor = GetValue(Imprimir.Recordset, "Mes", ZERO)
    End If
  End If

End Sub

Private Sub Imprimir_ImprimiuRegistro(Cancelar As Boolean)
  EntradaBanco = EntradaBanco + GetValue(Imprimir.Recordset, "Entrada", ZERO)
  SaidaBanco = SaidaBanco + GetValue(Imprimir.Recordset, "Saída", ZERO)
  SaldoBanco = GetValue(Imprimir.Recordset, "Saldo", ZERO)
  
  EntradaMes = EntradaMes + GetValue(Imprimir.Recordset, "Entrada", ZERO)
  SaidaMes = SaidaMes + GetValue(Imprimir.Recordset, "Saída", ZERO)
  SaldoMes = SaldoBanco
  
  Entrada = Entrada + GetValue(Imprimir.Recordset, "Entrada", ZERO)
  Saida = Saida + GetValue(Imprimir.Recordset, "Saída", ZERO)
    'Autor:Leandro Mesquita
    'Data: 26/06/2007
    'pt. 82286
  Saldo = Saldo + SaldoBanco
End Sub

Private Sub Imprimir_IniciarGrupo(ByVal Ordem As Byte)
  If Ordem = 1 Then
    EntradaBanco = 0
    SaidaBanco = 0
    SaldoBanco = 0
  End If
  If Ordem = 3 And Not QuebraBanco Then
    cmpFluxo(24).Caption = "Total " & LCase$(GetValue(Imprimir.Recordset, "Mes", ZERO)) & ":"
    EntradaMes = 0
    SaidaMes = 0
    SaldoMes = 0
  End If
End Sub

Private Sub Imprimir_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
   Saldo = 0
   Saida = 0
   Entrada = 0
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
