VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form fimpFluxoAnalitico 
   Caption         =   "Fluxo de Caixa Analítico"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12225
   Icon            =   "impFluxoAnalitico.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   131.763
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   215.636
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection RodapeSaldo 
      Align           =   1  'Align Top
      Height          =   2340
      Left            =   0
      Top             =   4245
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   4128
      Tipo            =   5
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   24
         Left            =   6480
         TabIndex        =   33
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
         Left            =   9000
         TabIndex        =   34
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   10320
         TabIndex        =   35
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
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
         TabIndex        =   53
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
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   57
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
         TabIndex        =   58
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
         TabIndex        =   59
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
         TabIndex        =   60
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
         TabIndex        =   61
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
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         TabIndex        =   67
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
         TabIndex        =   68
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
         TabIndex        =   69
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
         TabIndex        =   70
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
         Left            =   6480
         TabIndex        =   71
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
         Left            =   8880
         TabIndex        =   72
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
         TabIndex        =   73
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
         TabIndex        =   74
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
         Index           =   3
         X1              =   120
         X2              =   11280
         Y1              =   0
         Y2              =   0
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
         Index           =   0
         X1              =   120
         X2              =   11280
         Y1              =   600
         Y2              =   600
      End
   End
   Begin ReportX.ReportSection RodapeData 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      Top             =   3555
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   1217
      Tipo            =   5
      Ordem           =   2
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   18
         Left            =   6480
         TabIndex        =   27
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
         Left            =   9000
         TabIndex        =   28
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   10320
         TabIndex        =   29
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
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
         Left            =   6480
         TabIndex        =   30
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
         Left            =   7680
         TabIndex        =   31
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Campo           =   "Data"
         Caption         =   ""
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
         Index           =   23
         Left            =   9120
         TabIndex        =   32
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
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   873
      Tipo            =   3
      Ordem           =   2
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   14
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
         TabIndex        =   15
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
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   1482
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
         TabIndex        =   10
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         Left            =   0
         Top             =   120
         Width           =   11415
      End
   End
   Begin ReportX.ReportSection RodapeCabecalho 
      Align           =   1  'Align Top
      Height          =   150
      Left            =   0
      Top             =   3405
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   265
      Tipo            =   5
      Ordem           =   3
   End
   Begin ReportX.ReportSection Detalhe 
      Align           =   1  'Align Top
      Height          =   225
      Left            =   0
      Top             =   3180
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   397
      Ordem           =   1
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   27
         Left            =   60
         TabIndex        =   36
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   370
         Campo           =   "Empresa"
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
         Index           =   28
         Left            =   1350
         TabIndex        =   37
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
         Left            =   3120
         TabIndex        =   38
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   3840
         TabIndex        =   39
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   370
         Campo           =   "Descrição"
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
         Index           =   31
         Left            =   5160
         TabIndex        =   40
         Top             =   0
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   370
         Campo           =   "Controle"
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
         Index           =   32
         Left            =   6120
         TabIndex        =   41
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Campo           =   "Conta"
         Caption         =   ""
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
         Index           =   33
         Left            =   7200
         TabIndex        =   42
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
         Campo           =   "Cheque"
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
         Index           =   34
         Left            =   7800
         TabIndex        =   43
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Campo           =   "Vencimento"
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
         Index           =   35
         Left            =   8640
         TabIndex        =   44
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
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
         Left            =   9480
         TabIndex        =   45
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   10320
         TabIndex        =   46
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         Index           =   66
         Left            =   2760
         TabIndex        =   75
         Top             =   0
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   370
         Campo           =   "Parcela"
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
      Top             =   2910
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   476
      Tipo            =   3
      Ordem           =   3
      Begin ReportX.ReportField cmpFluxo 
         Height          =   210
         Index           =   7
         Left            =   60
         TabIndex        =   16
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   370
         Caption         =   "Empresa"
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
         Left            =   1350
         TabIndex        =   17
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   3120
         TabIndex        =   18
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
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
         Left            =   3840
         TabIndex        =   19
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
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
         Index           =   11
         Left            =   5160
         TabIndex        =   20
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   370
         Caption         =   "Controle"
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
         Left            =   6240
         TabIndex        =   21
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   "Conta"
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
         Left            =   7200
         TabIndex        =   22
         Top             =   0
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   370
         Caption         =   "Cheque"
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
         Left            =   7800
         TabIndex        =   23
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   370
         Caption         =   "Vencto."
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
         Index           =   15
         Left            =   8640
         TabIndex        =   24
         Top             =   0
         Width           =   705
         _ExtentX        =   1244
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
         Left            =   9360
         TabIndex        =   25
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
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
         Left            =   10320
         TabIndex        =   26
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
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
   Begin ReportX.ReportMain Imprimir 
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   7560
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
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   2778
      Tipo            =   2
      Begin ReportX.ReportField Campo 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   1
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
         Left            =   2160
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
         Top             =   690
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         Left            =   0
         Top             =   120
         Width           =   11415
      End
   End
End
Attribute VB_Name = "fimpFluxoAnalitico"
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
Private UltimaData           As String
Private Banco                As Long
Private Limite               As Long
Private PagInicial           As Integer

Public Sub Config(rstDados As Object, strTitulo As String, blnQuebraDia As Boolean, blnQuebraBanco As Boolean, strTabelaSaldos As String, blnImpDescControle As Boolean, blnImpRazao As Boolean, blnImpResumo As Boolean, bPagInicial As Integer)
    Dim X As Integer
    
On Error GoTo err_Handler
    QuebraDia = blnQuebraDia
    QuebraBanco = blnQuebraBanco
    TabelaSaldo = strTabelaSaldos
    PagInicial = bPagInicial
    If Not QuebraDia Then
        QuebraData.Mostrar = False
        RodapeData.Mostrar = False
    End If
    
    If Not QuebraBanco Then
        cmpFluxo(0).Mostrar = False
        cmpFluxo(1).Mostrar = False
        cmpFluxo(2).Mostrar = False
        cmpFluxo(24).Caption = "Total:"
        cmpFluxo(62).Caption = "Saldo:"
    End If
    
    If blnImpRazao Then
        cmpFluxo(42).Mostrar = True
        cmpFluxo(43).Mostrar = True
        Detalhe.Height = 8.202
    End If
    
    If blnImpDescControle Then
        cmpFluxo(38).Mostrar = True
        cmpFluxo(39).Mostrar = True
        cmpFluxo(40).Mostrar = True
        cmpFluxo(41).Mostrar = True
        
        cmpFluxo(10).Mostrar = False
        cmpFluxo(11).Mostrar = False
        cmpFluxo(30).Mostrar = False
        cmpFluxo(31).Mostrar = False
        
        cmpFluxo(12).Left = cmpFluxo(12).Left - 2300
        cmpFluxo(13).Left = cmpFluxo(13).Left - 2300
        cmpFluxo(14).Left = cmpFluxo(14).Left - 2300
        cmpFluxo(15).Left = cmpFluxo(15).Left - 2300
        cmpFluxo(32).Left = cmpFluxo(32).Left - 2300
        cmpFluxo(33).Left = cmpFluxo(33).Left - 2300
        cmpFluxo(34).Left = cmpFluxo(34).Left - 2300
        cmpFluxo(35).Left = cmpFluxo(35).Left - 2300
        
        If blnImpRazao Then
            Detalhe.Height = 12.435
        Else
            Detalhe.Height = 8.202
            cmpFluxo(38).Top = 240
            cmpFluxo(39).Top = 240
            cmpFluxo(40).Top = 240
            cmpFluxo(41).Top = 240
        End If
    End If
    
    If blnImpResumo Then
        For X = 44 To 65
            cmpFluxo(X).Mostrar = True
        Next X
        lneFluxo(1).Visible = True
        RodapeSaldo.Height = 38.365
    End If
    
    Campo(0).Caption = Date
    TituloRel = strTitulo
    'Campo(8).Caption = strBancosSelecionados
    
    Set Imprimir.Recordset = rstDados
    Imprimir.Ativar
    Unload Me
    Exit Sub
    
err_Handler:
    Call MsgBox("Erro ao configurar o relatório: " & err.Description, vbInformation, NomeModulo)
End Sub



Private Sub Form_Unload(Cancel As Integer)
  Set fimpFluxoAnalitico = Nothing
End Sub
Private Sub Imprimir_Erro(ByVal Numero As Long)
  RpxMsgErro Numero
End Sub
Private Sub Imprimir_FormulaCampo(ByVal Campo As String, Valor As Variant)
  Valor = FormulasHeader(Campo, Imprimir, TituloRel)

  Select Case Campo
    Case "Razão"
      Valor = GetFieldValue("Razão", "Empresas", "Apel = " & Quote(GetValue(Imprimir.Recordset, "Empresa", NUL), "'"), , NUL)
    
    Case "Página":  Valor = Valor + PagInicial
    
    Case "Duplicata"
      Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000")
      
    Case "Parcela"
      'Protocolo 73121
      Valor = GetValue(Imprimir.Recordset, Campo, NUL)
    
    Case "Cheque"
      Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000")
      
    Case "Banco"
      Valor = Format(GetValue(Imprimir.Recordset, Campo, ZERO), "000000000")
      Banco = GetValue(Imprimir.Recordset, Campo, ZERO)
  
    Case "Vencimento", "Pagamento"
      Valor = Format(GetValue(Imprimir.Recordset, Campo), "dd/mm/yy")
  
    Case "Data"
      Valor = UltimaData
      
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
      Valor = Soma("Entrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 2 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
    
    Case "Saída Aplicações"
      Valor = Soma("Saída", NomeTabeladoRST(Imprimir.Recordset), "Type = 2 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
    
    Case "Total Aplicações"
      Valor = Soma("Entrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 2 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO) - Soma("Saída", NomeTabeladoRST(Imprimir.Recordset), "Type = 2 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
    
    Case "Entrada Transferência"
      Valor = Soma("Entrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 1 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
    
    Case "Saída Transferência"
      Valor = Soma("Saída", NomeTabeladoRST(Imprimir.Recordset), "Type = 1 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
    
    Case "Total Transferência"
      Valor = Soma("Entrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 1 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO) - Soma("Saída", NomeTabeladoRST(Imprimir.Recordset), "Type = 1 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
      
    Case "Entrada Movimentação"
      Valor = Soma("Entrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 3 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
    
    Case "Saída Movimentação"
      Valor = Soma("Saída", NomeTabeladoRST(Imprimir.Recordset), "Type = 3 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
    
    Case "Total Movimentação"
      Valor = Soma("Entrada", NomeTabeladoRST(Imprimir.Recordset), "Type = 3 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO) - Soma("Saída", NomeTabeladoRST(Imprimir.Recordset), "Type = 3 " & IIf(QuebraBanco, " and Banco = " & Banco, ""), ZERO)
    
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
  TotalEntradaDia = TotalEntradaDia + GetValue(Imprimir.Recordset, "Entrada", ZERO)
  TotalSaidaDia = TotalSaidaDia + GetValue(Imprimir.Recordset, "Saída", ZERO)
  
  TotalEntradaBanco = TotalEntradaBanco + GetValue(Imprimir.Recordset, "Entrada", ZERO)
  TotalSaidaBanco = TotalSaidaBanco + GetValue(Imprimir.Recordset, "Saída", ZERO)
End Sub
Private Sub Imprimir_IniciarGrupo(ByVal Ordem As Byte)
  If Ordem = 1 Then
    UltimaData = GetValue(Imprimir.Recordset, "Data")
    TotalEntradaBanco = 0
    TotalSaidaBanco = 0
  ElseIf Ordem = 2 Then
    UltimaData = GetValue(Imprimir.Recordset, "Data")
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
