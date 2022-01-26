VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.4#0"; "ReportX.ocx"
Begin VB.Form fimpDuplLancAtraso 
   Caption         =   "Duplicatas/Lançamentos em Atraso"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   50.535
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   195.527
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportSection ReportSection4 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   1845
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   661
      Tipo            =   5
      Ordem           =   1
      Begin Fox.EBSReport EBSReport 
         Height          =   795
         Left            =   2010
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1402
         NomeRelatorio   =   "FOXFCO00041.ERC"
      End
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   21
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   370
         Caption         =   "Total"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   22
         Left            =   6390
         TabIndex        =   22
         Top             =   120
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
         Campo           =   "Total60"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   23
         Left            =   7560
         TabIndex        =   23
         Top             =   120
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
         Campo           =   "Total90"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   24
         Left            =   8730
         TabIndex        =   24
         Top             =   120
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   370
         Campo           =   "Total120"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   25
         Left            =   9870
         TabIndex        =   25
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Campo           =   "TotalMais120"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   13
         Left            =   5130
         TabIndex        =   26
         Top             =   120
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Campo           =   "Total30"
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
   End
   Begin ReportX.ReportSection ReportSection3 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   1590
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   450
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   28
         Left            =   3735
         TabIndex        =   27
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   370
         Campo           =   "Contato"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   14
         Left            =   5130
         TabIndex        =   14
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Campo           =   "Ate30"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   15
         Left            =   6345
         TabIndex        =   15
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Campo           =   "Ate60"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   16
         Left            =   7515
         TabIndex        =   16
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Campo           =   "Ate90"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   17
         Left            =   8685
         TabIndex        =   17
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Campo           =   "Ate120"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   18
         Left            =   9870
         TabIndex        =   18
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Campo           =   "Apos120"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   19
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   370
         Campo           =   "NmEmpresa"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   29
         Left            =   2430
         TabIndex        =   29
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   370
         Campo           =   "Telefone"
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
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   1335
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   450
      Tipo            =   3
      Ordem           =   1
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   2265
         _ExtentX        =   3995
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   8
         Left            =   5130
         TabIndex        =   9
         Top             =   0
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   370
         Caption         =   "30 "
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   9
         Left            =   6345
         TabIndex        =   10
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Caption         =   "60 "
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   10
         Left            =   7515
         TabIndex        =   11
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Caption         =   "90 "
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   11
         Left            =   8685
         TabIndex        =   12
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Caption         =   "120 "
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   12
         Left            =   9870
         TabIndex        =   13
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   370
         Caption         =   "Acima de 120 "
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   26
         Left            =   3735
         TabIndex        =   28
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   370
         Caption         =   "Contato"
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
      Begin ReportX.ReportField Campo 
         Height          =   210
         Index           =   27
         Left            =   2430
         TabIndex        =   30
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   370
         Caption         =   "Telefone"
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
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   2355
      Tipo            =   2
      Begin ReportX.ReportField Campo 
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   1
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
         TabIndex        =   2
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
         Left            =   9495
         TabIndex        =   3
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
         TabIndex        =   4
         Top             =   480
         Width           =   7755
         _ExtentX        =   13679
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
         Left            =   10215
         TabIndex        =   5
         Top             =   120
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
         Left            =   9495
         TabIndex        =   6
         Top             =   480
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
         Left            =   1680
         TabIndex        =   7
         Top             =   120
         Width           =   7755
         _ExtentX        =   13679
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
         Index           =   20
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   10875
         _ExtentX        =   19182
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
         Width           =   11655
      End
   End
   Begin ReportX.ReportMain Imprimir 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   1530
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Titulo          =   ""
   End
End
Attribute VB_Name = "fimpDuplLancAtraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curTot30      As Currency
Dim curTot60      As Currency
Dim curTot120     As Currency
Dim curTot90      As Currency
Dim curTotMais120 As Currency
Private strTelefone As String

Public Sub Config(Dados As Object)
    Set Imprimir.Recordset = Dados
    Imprimir.Ativar
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fimpDuplLancAtraso = Nothing
End Sub
Private Sub Imprimir_Erro(ByVal Numero As Long)
    RpxMsgErro Numero
End Sub

Private Sub Imprimir_FormulaCampo(ByVal Campo As String, Valor As Variant)
 
    Dim strTitulo As String
    
    If frptDuplLancAtraso.cboDuplLancAtraso.Text = "Duplicatas" Then
        strTitulo = "DUPLICATAS EM ATRASO (A RECEBER)"
    ElseIf frptDuplLancAtraso.cboDuplLancAtraso.Text = "Lançamentos" Then
        strTitulo = "LANÇAMENTOS EM ATRASO (A RECEBER)"
    Else
        strTitulo = "DUPLICATAS E LANÇAMENTOS EM ATRASO (A RECEBER)"
    End If
    
    Valor = FormulasHeader(Campo, Imprimir, strTitulo)
    
    Select Case Campo
        
        Case "NmEmpresa"
            Valor = GetValue(Imprimir.Recordset, "razaoEmpresa", NUL)
        
        Case "Total30"
            Valor = Format(curTot30, FMOEDA)
        
        Case "Total60"
            Valor = Format(curTot60, FMOEDA)
        
        Case "Total90"
            Valor = Format(curTot90, FMOEDA)
        
        Case "Total120"
            Valor = Format(curTot120, FMOEDA)
        
        Case "TotalMais120"
            Valor = Format(curTotMais120, FMOEDA)
        
        Case "Contato"
            Valor = GetValue(Imprimir.Recordset, "Contato", "")
            
        Case "Telefone"
            If GetValue(Imprimir.Recordset, "Fone1", "") <> "" Then
                Valor = GetValue(Imprimir.Recordset, "Fone1", "")
            Else
                Valor = GetValue(Imprimir.Recordset, "foneCobranca", "")
            End If
    End Select

End Sub

Private Sub Imprimir_ImprimiuRegistro(Cancelar As Boolean)

    curTot30 = curTot30 + GetValue(Imprimir.Recordset, "Ate30", ZERO)
    curTot60 = curTot60 + GetValue(Imprimir.Recordset, "Ate60", ZERO)
    curTot90 = curTot90 + GetValue(Imprimir.Recordset, "Ate90", ZERO)
    curTot120 = curTot120 + GetValue(Imprimir.Recordset, "Ate120", ZERO)
    curTotMais120 = curTotMais120 + GetValue(Imprimir.Recordset, "Apos120", ZERO)
  
End Sub

Private Sub Imprimir_IniciarRelatorio(ByVal Impressora As Boolean, Cancelar As Boolean)
    curTot30 = 0
    curTot60 = 0
    curTot120 = 0
    curTot90 = 0
    curTotMais120 = 0
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
