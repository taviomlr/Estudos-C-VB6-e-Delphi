VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.3#0"; "ReportX.ocx"
Begin VB.Form fimpFluxoSinteticoConta 
   KeyPreview      =   -1  'True
   Caption         =   "Fluxo de Caixa Sintético por Conta"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   120.386
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   254.529
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Imprimir 
      Height          =   480
      Left            =   240
      TabIndex        =   0
      Top             =   6000
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Titulo          =   ""
   End
   Begin ReportX.ReportSection RodapeSaldo 
      Align           =   1  'Align Top
      Height          =   2340
      Left            =   0
      Top             =   4245
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   4128
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   24
         Left            =   5760
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Caption         =   "Total do Banco:"
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
         Left            =   8640
         TabIndex        =   2
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   9960
         TabIndex        =   3
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   44
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   "Resumo"
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
         Index           =   45
         Left            =   7440
         TabIndex        =   5
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Caption         =   "Entradas"
         Alignment       =   1
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
         Index           =   46
         Left            =   8760
         TabIndex        =   6
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Caption         =   "Saídas"
         Alignment       =   1
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
         Index           =   47
         Left            =   10080
         TabIndex        =   7
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Caption         =   "Total"
         Alignment       =   1
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
         Index           =   48
         Left            =   5880
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   370
         Caption         =   "Aplicações:"
         Alignment       =   1
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
         Index           =   49
         Left            =   5880
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   370
         Caption         =   "Transferências:"
         Alignment       =   1
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
         Index           =   50
         Left            =   5880
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   370
         Caption         =   "Movimentação:"
         Alignment       =   1
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
         Index           =   51
         Left            =   5880
         TabIndex        =   11
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   370
         Caption         =   "Saldo:"
         Alignment       =   1
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
         Index           =   52
         Left            =   7440
         TabIndex        =   12
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Entrada Aplicações"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   53
         Left            =   8760
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Saída Aplicações"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   54
         Left            =   7440
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Entrada Transferência"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   55
         Left            =   8760
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Saída Transferência"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   56
         Left            =   7440
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Entrada Movimentação"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   57
         Left            =   8760
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Saída Movimentação"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   58
         Left            =   10080
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Total Movimentação"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   59
         Left            =   10080
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Saldo"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   60
         Left            =   10080
         TabIndex        =   20
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Total Aplicações"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   61
         Left            =   10080
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Total Transferência"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
         Index           =   62
         Left            =   5760
         TabIndex        =   22
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Caption         =   "Saldo do Banco:"
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
         Left            =   8640
         TabIndex        =   23
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   64
         Left            =   5880
         TabIndex        =   24
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   370
         Caption         =   "Limite de crédito:"
         Alignment       =   1
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
         Index           =   65
         Left            =   7440
         TabIndex        =   25
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   370
         Campo           =   "Limite"
         Formato         =   "Standard"
         Caption         =   ""
         TipoCampo       =   1
         Formula         =   -1  'True
         Alignment       =   1
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
      Begin VB.Line lneFluxo 
         BorderStyle     =   4  'Dash-Dot
         Index           =   0
         X1              =   120
         X2              =   11280
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line lneFluxo 
         BorderStyle     =   4  'Dash-Dot
         Index           =   1
         Visible         =   0   'False
         X1              =   120
         X2              =   11400
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line lneFluxo 
         BorderStyle     =   4  'Dash-Dot
         Index           =   3
         X1              =   120
         X2              =   11280
         Y1              =   0
         Y2              =   0
      End
   End
   Begin ReportX.ReportSection RodapeData 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      Top             =   3555
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   1217
      Tipo            =   5
      Ordem           =   2
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   18
         Left            =   6000
         TabIndex        =   26
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Caption         =   "Total do Dia:"
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
         Index           =   19
         Left            =   8640
         TabIndex        =   27
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Total Entrada Dia"
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
         Left            =   9960
         TabIndex        =   28
         Top             =   120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Total Saída Dia"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Caption         =   "Saldo do Dia:"
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
         Left            =   7200
         TabIndex        =   30
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   370
         Campo           =   "Data"
         Caption         =   ""
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
         Index           =   23
         Left            =   8640
         TabIndex        =   31
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "Saldo do Dia"
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
   End
   Begin ReportX.ReportSection QuebraData 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   2415
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   873
      Tipo            =   3
      Ordem           =   2
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   370
         Caption         =   "Movimentação do Dia "
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
         Index           =   6
         Left            =   2040
         TabIndex        =   33
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Campo           =   "Data"
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
      Begin VB.Line lneFluxo 
         BorderStyle     =   4  'Dash-Dot
         Index           =   2
         X1              =   3120
         X2              =   11280
         Y1              =   240
         Y2              =   240
      End
   End
   Begin ReportX.ReportSection QuebraSaldo 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      Top             =   1575
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   1482
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   34
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
         TabIndex        =   35
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
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
      Top             =   3405
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   265
      Tipo            =   5
      Ordem           =   3
   End
   Begin ReportX.ReportSection Detalhe 
      Align           =   1  'Align Top
      Height          =   225
      Left            =   0
      Top             =   3180
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   397
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   32
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   370
         Campo           =   "Conta"
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
         Index           =   36
         Left            =   8640
         TabIndex        =   40
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "CEntrada"
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
         Left            =   9960
         TabIndex        =   41
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   370
         Campo           =   "CSaída"
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
         TabIndex        =   42
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         TabIndex        =   45
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
         Left            =   1440
         TabIndex        =   46
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
         TabIndex        =   47
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
   End
   Begin ReportX.ReportSection QuebraCabecalho 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      Top             =   2910
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   476
      Tipo            =   3
      Ordem           =   3
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   12
         Left            =   120
         TabIndex        =   48
         Top             =   0
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   370
         Caption         =   "Conta"
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
         Left            =   8640
         TabIndex        =   49
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   9960
         TabIndex        =   50
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
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
   End
   Begin ReportX.ReportMain ReportMain1 
      Height          =   480
      Left            =   720
      TabIndex        =   51
      Top             =   7560
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
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   2778
      Tipo            =   2
      Begin ReportX.ReportField Campo 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   52
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
         TabIndex        =   53
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
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   57
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
         TabIndex        =   58
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
         TabIndex        =   59
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
Attribute VB_Name = "fimpFluxoSinteticoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private QuebraDia            As Boolean
Private QuebraBanco          As Boolean
Private SaldoInicial         As Double
Private TotalEntradaDia      As Double
Private TotalSaidaDia        As Double
Private TotalEntradaBanco    As Double
Private TotalSaidaBanco      As Double
Private TabelaSaldo          As String
Private TituloRel            As String
Private banco                As Long
Private Limite               As Long

Public Sub Config(rstDados As Object, strTitulo As String, blnQuebraDia As Boolean, blnQuebraBanco As Boolean, strTabelaSaldos As String, blnImpDescControle As Boolean, blnImpRazao As Boolean, blnImpResumo As Boolean)

  TabelaSaldo = strTabelaSaldos
  
 ' If Not QuebraDia Then
 '   QuebraData.Mostrar = False
 '   RodapeData.Mostrar = False
 ' End If
  QuebraBanco = blnQuebraBanco
  QuebraDia = blnQuebraDia
  
  If Not QuebraBanco Then
    cmpFluxo(0).Mostrar = False
    cmpFluxo(1).Mostrar = False
    cmpFluxo(2).Mostrar = False
    cmpFluxo(24).Caption = "Total:"
    cmpFluxo(62).Caption = "Saldo:"
  End If
  
 ' If blnImpRazao Then
 '   cmpFluxo(42).Mostrar = True
 '   cmpFluxo(43).Mostrar = True
 '   Detalhe.Height = 8.202
 ' End If
  
' ' If blnImpDescControle Then
' '   cmpFluxo(38).Mostrar = True
' '   cmpFluxo(39).Mostrar = True
'  '  cmpFluxo(40).Mostrar = True
'  '  cmpFluxo(41).Mostrar = True
'
'    cmpFluxo(10).Mostrar = False
'    cmpFluxo(11).Mostrar = False
'    cmpFluxo(30).Mostrar = False
'    cmpFluxo(31).Mostrar = False
'
'    cmpFluxo(12).Left = cmpFluxo(12).Left - 3000
'    cmpFluxo(13).Left = cmpFluxo(13).Left - 3000
'    cmpFluxo(14).Left = cmpFluxo(14).Left - 3000
'    cmpFluxo(15).Left = cmpFluxo(15).Left - 3000
'    cmpFluxo(32).Left = cmpFluxo(32).Left - 3000
'    cmpFluxo(33).Left = cmpFluxo(33).Left - 3000
'    cmpFluxo(34).Left = cmpFluxo(34).Left - 3000
'    cmpFluxo(35).Left = cmpFluxo(35).Left - 3000
'
'    If blnImpRazao Then
'      Detalhe.Height = 12.435
'    Else
'      Detalhe.Height = 8.202
'      cmpFluxo(38).Top = 240
'      cmpFluxo(39).Top = 240
'      cmpFluxo(40).Top = 240
'      cmpFluxo(41).Top = 240
'    End If
'  End If
'
  If blnImpResumo Then
    Dim X As Integer
    For X = 44 To 65
      cmpFluxo(X).Mostrar = True
    Next X
    lneFluxo(1).Visible = True
    RodapeSaldo.Height = 38.365
  End If
  
  TituloRel = strTitulo
  'Campo(8).Caption = strBancosSelecionados
  
  Set Imprimir.Recordset = rstDados
   
  Imprimir.Ativar
  
  Unload Me
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set fimpFluxoSinteticoConta = Nothing
End Sub
Private Sub Imprimir_Erro(ByVal numero As Long)
  RpxMsgErro numero
End Sub
Private Sub Imprimir_FormulaCampo(ByVal Campo As String, Valor As Variant)
  Valor = FormulasHeader(Campo, Imprimir, TituloRel)

  Select Case Campo
    Case "Razão"
      Valor = GetFieldValue("Razão", "Empresas", "Apel = " & Quote(GetValue(Imprimir.Recordset, "Empresa", NUL), "'"), , NUL)
    
    Case "Duplicata"
      Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000")
    
    Case "Cheque"
      Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000")
    
    Case "Conta"
      Valor = Format(GetValue(Imprimir.Recordset, "Conta", ZERO), "000000000") & " - " & GetValue(Imprimir.Recordset, "DescConta", NUL)
      
    Case "Banco"
      Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000000")
      banco = GetValue(Imprimir.Recordset, Campo, ZERO)
  
    Case "Vencimento", "Pagamento"
      Valor = Format(GetValue(Imprimir.Recordset, Campo), "dd/mm/yy")
  
    Case "Total Entrada Dia"
      Valor = TotalEntradaDia

    Case "Total Saída Dia"
      Valor = TotalSaidaDia

    Case "Saldo do Dia"
      Valor = SaldoInicial + TotalEntradaBanco - TotalSaidaBanco

    Case "Total Entrada Banco"
      Valor = TotalEntradaBanco

    Case "Total Saída Banco"
      Valor = TotalSaidaBanco
      
    Case "Total Saldo Banco"
      Valor = SaldoInicial + TotalEntradaBanco - TotalSaidaBanco

    Case "Saldo Anterior"
      If QuebraBanco Then
        Valor = GetFieldValue("Valor", TabelaSaldo, "Banco = " & GetValue(Imprimir.Recordset, "Banco", ZERO) & " and Tipo = false", , ZERO)
      Else
        Valor = GetFieldValue("Valor", TabelaSaldo, "Tipo = false", , ZERO)
      End If
      SaldoInicial = Valor
      
    Case "Entrada Aplicações"
      Valor = Soma("CEntrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 2 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
    
    Case "Saída Aplicações"
      Valor = Soma("CSaída", NomeTabeladoRST(Imprimir.Recordset), "Type = 2 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
    
    Case "Total Aplicações"
      Valor = Soma("CEntrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 2 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO) - Soma("CSaída", NomeTabeladoRST(Imprimir.Recordset), "Type = 2 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
    
    Case "Entrada Transferência"
      Valor = Soma("CEntrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 1 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
    
    Case "Saída Transferência"
      Valor = Soma("CSaída", NomeTabeladoRST(Imprimir.Recordset), "Type = 1 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
    
    Case "Total Transferência"
      Valor = Soma("CEntrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 1 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO) - Soma("CSaída", NomeTabeladoRST(Imprimir.Recordset), "Type = 1 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
      
    Case "Entrada Movimentação"
      Valor = Soma("CEntrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 3 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
    
    Case "Saída Movimentação"
      Valor = Soma("CSaída", NomeTabeladoRST(Imprimir.Recordset), "Type = 3 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
    
    Case "Total Movimentação"
      Valor = Soma("CEntrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 3 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO) - Soma("CSaída", NomeTabeladoRST(Imprimir.Recordset), "Type = 3 " & IIf(QuebraBanco, " and Banco = " & banco, ""), ZERO)
    
    Case "Saldo"
      Valor = SaldoInicial + (TotalEntradaBanco - TotalSaidaBanco)
    
    'Sérgio 08/09/2004
    'Rotina criada para a visualização do campo limite de crédito do cadastro de bancos.
    Case "Limite"
      Imprimir.Recordset.MovePrevious
      Valor = GetFieldValue("[Limite de crédito]", "Bancos", "Banco = " & GetValue(Imprimir.Recordset, "Banco", ZERO), , ZERO)
      Limite = Valor
      Imprimir.Recordset.MoveNext
  End Select
  
End Sub
Private Sub Imprimir_FormulaGrupo(ByVal Ordem As Byte, Valor As Variant)
  If Ordem = 1 Then
    If QuebraBanco Then
      Valor = GetValue(Imprimir.Recordset, "Banco", ZERO)
    End If
  ElseIf Ordem = 2 Then
    If QuebraDia Then
      Valor = GetValue(Imprimir.Recordset, "Data", ZERO)
    End If
  End If
End Sub
Private Sub Imprimir_ImprimiuRegistro(Cancelar As Boolean)
  TotalEntradaDia = TotalEntradaDia + GetValue(Imprimir.Recordset, "CEntrada", ZERO)
  TotalSaidaDia = TotalSaidaDia + GetValue(Imprimir.Recordset, "CSaída", ZERO)
  
  TotalEntradaBanco = TotalEntradaBanco + GetValue(Imprimir.Recordset, "CEntrada", ZERO)
  TotalSaidaBanco = TotalSaidaBanco + GetValue(Imprimir.Recordset, "CSaída", ZERO)
End Sub
Private Sub Imprimir_IniciarGrupo(ByVal Ordem As Byte)
  If Ordem = 1 Then
    TotalEntradaBanco = 0
    TotalSaidaBanco = 0
  ElseIf Ordem = 2 Then
    TotalEntradaDia = 0
    TotalSaidaDia = 0
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
