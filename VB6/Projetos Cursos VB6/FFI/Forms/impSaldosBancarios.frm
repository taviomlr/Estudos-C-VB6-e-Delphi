VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.3#0"; "reportx.ocx"
Begin VB.Form fimpSaldosBancarios 
   KeyPreview      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   47.89
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   197.379
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection4 
      Align           =   1  'Align Top
      Height          =   30
      Left            =   0
      Top             =   2085
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   53
      Tipo            =   5
      Ordem           =   1
   End
   Begin ReportX.ReportMain Imprimir 
      Height          =   480
      Left            =   120
      TabIndex        =   28
      Top             =   2160
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Titulo          =   ""
   End
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   1830
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   450
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   370
         Campo           =   "Banco"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   1
         Left            =   840
         TabIndex        =   9
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   2
         Left            =   2640
         TabIndex        =   10
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "Inicial"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   3
         Left            =   3840
         TabIndex        =   11
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Campo           =   "Entradas"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   4
         Left            =   4800
         TabIndex        =   12
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Campo           =   "Saidas"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   5
         Left            =   5760
         TabIndex        =   13
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Campo           =   "Aplicações"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   6
         Left            =   6840
         TabIndex        =   14
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Campo           =   "Resgates"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   7
         Left            =   7800
         TabIndex        =   15
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Campo           =   "Taxas"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   8
         Left            =   8880
         TabIndex        =   16
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Campo           =   "Juros"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   9
         Left            =   9840
         TabIndex        =   17
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Campo           =   "Final"
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
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      Top             =   1335
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   873
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   11
         Left            =   840
         TabIndex        =   19
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   12
         Left            =   2640
         TabIndex        =   20
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Caption         =   "Saldo Inicial"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   13
         Left            =   3840
         TabIndex        =   21
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   "Entradas"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   14
         Left            =   4800
         TabIndex        =   22
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   "Saídas"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   15
         Left            =   7800
         TabIndex        =   23
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Caption         =   "Taxas"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   16
         Left            =   6840
         TabIndex        =   24
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   "Resgates"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   17
         Left            =   5760
         TabIndex        =   25
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   370
         Caption         =   "Aplicações"
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   18
         Left            =   9840
         TabIndex        =   26
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
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
      Begin ReportX.ReportField field 
         Height          =   210
         Index           =   19
         Left            =   8880
         TabIndex        =   27
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   370
         Caption         =   "Juros"
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
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11040
         Y1              =   360
         Y2              =   360
      End
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   2355
      Tipo            =   2
      Begin ReportX.ReportField Campo 
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   120
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
         Left            =   360
         TabIndex        =   1
         Top             =   480
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
         Left            =   9360
         TabIndex        =   2
         Top             =   120
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
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
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
         Left            =   10080
         TabIndex        =   4
         Top             =   120
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
         Left            =   9360
         TabIndex        =   5
         Top             =   480
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
         Left            =   1680
         TabIndex        =   6
         Top             =   120
         Width           =   7575
         _ExtentX        =   13361
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
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   10815
         _ExtentX        =   19076
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
         Top             =   0
         Width           =   11055
      End
   End
End
Attribute VB_Name = "fimpSaldosBancarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Config(Dados As Object)
  Set Imprimir.Recordset = Dados
  Imprimir.Ativar
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fimpSaldosBancarios = Nothing
End Sub
Private Sub Imprimir_Erro(ByVal Numero As Long)
  RpxMsgErro Numero
End Sub

Private Sub Imprimir_FormulaCampo(ByVal Campo As String, Valor As Variant)
 Valor = FormulasHeader(Campo, Imprimir, "Saldos Bancários ")
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
