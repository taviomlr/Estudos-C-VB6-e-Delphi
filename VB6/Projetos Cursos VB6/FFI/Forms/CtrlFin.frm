VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frptCtrlFinanc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle Financeiro"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "CtrlFin.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCtrlFinanc 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   47
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton cmdCtrlFinanc 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   46
      Top             =   7380
      Width           =   1215
   End
   Begin VB.CommandButton cmdCtrlFinanc 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   45
      Top             =   7380
      Width           =   1215
   End
   Begin VB.Frame fraTab 
      Caption         =   "Controle Financeiro Sint�tico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6795
      Left            =   240
      TabIndex        =   40
      Top             =   360
      Width           =   5895
      Begin VB.Frame fraCentro 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   120
         TabIndex        =   44
         Top             =   4320
         Width           =   5715
         Begin VB.TextBox txtCtrlFinanc 
            Height          =   315
            Index           =   10
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   13
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chkDiscCentroCusto 
            Caption         =   "Discriminar Centro de Custo"
            Height          =   255
            Left            =   1080
            TabIndex        =   49
            Top             =   960
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txtCtrlFinanc 
            Height          =   315
            Index           =   9
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   12
            Top             =   270
            Width           =   855
         End
         Begin VB.Label lblFrame 
            AutoSize        =   -1  'True
            Caption         =   "C�digo dos Centros de Custo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   0
            Width           =   2475
         End
         Begin VB.Line linFrame 
            BorderColor     =   &H80000010&
            Index           =   6
            X1              =   5640
            X2              =   0
            Y1              =   90
            Y2              =   90
         End
         Begin VB.Line linFrame 
            BorderColor     =   &H80000014&
            Index           =   7
            X1              =   5640
            X2              =   0
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblNomes 
            Caption         =   "lblNomes(8)"
            Height          =   195
            Index           =   8
            Left            =   2040
            TabIndex        =   36
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   3585
         End
         Begin VB.Label lblNomes 
            Caption         =   "lblNomes(7)"
            Height          =   195
            Index           =   7
            Left            =   2040
            TabIndex        =   34
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   3585
         End
         Begin VB.Label lblCtrlFinanc 
            AutoSize        =   -1  'True
            Caption         =   "Final:"
            Height          =   195
            Index           =   12
            Left            =   0
            TabIndex        =   35
            Top             =   660
            Width           =   375
         End
         Begin VB.Label lblCtrlFinanc 
            AutoSize        =   -1  'True
            Caption         =   "Ini&cial:"
            Height          =   195
            Index           =   13
            Left            =   0
            TabIndex        =   33
            Top             =   300
            Width           =   450
         End
      End
      Begin VB.Frame fraModelo 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   6120
         Width           =   5655
         Begin VB.TextBox txtCtrlFinanc 
            Height          =   315
            Index           =   11
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   54
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblFrame 
            AutoSize        =   -1  'True
            Caption         =   "Modelo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   56
            Top             =   0
            Width           =   630
         End
         Begin VB.Line linFrame 
            BorderColor     =   &H80000014&
            Index           =   10
            X1              =   5640
            X2              =   0
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label lblNomes 
            Caption         =   "lblNomes(9)"
            Height          =   195
            Index           =   9
            Left            =   2400
            TabIndex        =   57
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   3165
         End
         Begin VB.Line linFrame 
            BorderColor     =   &H80000010&
            Index           =   11
            X1              =   5640
            X2              =   0
            Y1              =   105
            Y2              =   105
         End
         Begin VB.Label lblCtrlFinanc 
            AutoSize        =   -1  'True
            Caption         =   "Mo&delo:"
            Height          =   195
            Index           =   16
            Left            =   0
            TabIndex        =   55
            Top             =   240
            Width           =   570
         End
      End
      Begin VB.ComboBox cboConciliado 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cboTipoData 
         Height          =   315
         ItemData        =   "CtrlFin.frx":0C42
         Left            =   1200
         List            =   "CtrlFin.frx":0C52
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkMostrarSaldoBanco 
         Caption         =   "Mostrar Saldo anterior e atual dos Bancos"
         Height          =   255
         Left            =   2520
         TabIndex        =   50
         Top             =   1350
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   8
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   14
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CheckBox chkSaldoAnterior 
         Caption         =   "Calcular Saldo Anterior?"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "CtrlFin.frx":0C81
         Left            =   4080
         List            =   "CtrlFin.frx":0C8E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboOrigem 
         Height          =   315
         ItemData        =   "CtrlFin.frx":0CAD
         Left            =   1200
         List            =   "CtrlFin.frx":0CBA
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   7
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   9
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   6
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   5
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   7
         Top             =   2100
         Width           =   1215
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   4
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   6
         Top             =   1740
         Width           =   1215
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   3
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   11
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   2
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   10
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   1
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtCtrlFinanc 
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Conciliado:"
         Height          =   195
         Index           =   15
         Left            =   3000
         TabIndex        =   52
         Top             =   600
         Width           =   780
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Filtro por Data:"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Moeda:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   38
         Top             =   5820
         Width           =   540
      End
      Begin VB.Label lblNomes 
         Caption         =   "lblNomes(6)"
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   39
         Top             =   5760
         UseMnemonic     =   0   'False
         Width           =   3165
      End
      Begin VB.Label lblFrame 
         AutoSize        =   -1  'True
         Caption         =   "Moeda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   37
         Top             =   5520
         Width           =   585
      End
      Begin VB.Line linFrame 
         BorderColor     =   &H80000014&
         Index           =   9
         X1              =   5760
         X2              =   120
         Y1              =   5655
         Y2              =   5655
      End
      Begin VB.Line linFrame 
         BorderColor     =   &H80000010&
         Index           =   8
         X1              =   5760
         X2              =   120
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         Height          =   195
         Index           =   11
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Origem:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblNomes 
         Caption         =   "lblNomes(5)"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   27
         Top             =   3000
         UseMnemonic     =   0   'False
         Width           =   3225
      End
      Begin VB.Label lblNomes 
         Caption         =   "lblNomes(4)"
         Height          =   195
         Index           =   4
         Left            =   2520
         TabIndex        =   25
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   3225
      End
      Begin VB.Label lblNomes 
         Caption         =   "lblNomes(3)"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   23
         Top             =   2100
         UseMnemonic     =   0   'False
         Width           =   3225
      End
      Begin VB.Label lblNomes 
         Caption         =   "lblNomes(2)"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   21
         Top             =   1740
         UseMnemonic     =   0   'False
         Width           =   3225
      End
      Begin VB.Label lblNomes 
         Caption         =   "lblNomes(1)"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   31
         Top             =   3960
         UseMnemonic     =   0   'False
         Width           =   3585
      End
      Begin VB.Label lblNomes 
         Caption         =   "lblNomes(0)"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   29
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   3585
      End
      Begin VB.Label lblFrame 
         AutoSize        =   -1  'True
         Caption         =   "C�digo dos Grupos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   43
         Top             =   3360
         Width           =   1620
      End
      Begin VB.Line linFrame 
         BorderColor     =   &H80000010&
         Index           =   5
         X1              =   5760
         X2              =   120
         Y1              =   3465
         Y2              =   3465
      End
      Begin VB.Line linFrame 
         BorderColor     =   &H80000014&
         Index           =   4
         X1              =   5760
         X2              =   120
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "I&nicial:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblFrame 
         AutoSize        =   -1  'True
         Caption         =   "C�digo das Contas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Top             =   2430
         Width           =   1605
      End
      Begin VB.Line linFrame 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   5760
         X2              =   120
         Y1              =   2535
         Y2              =   2535
      End
      Begin VB.Line linFrame 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   5760
         X2              =   120
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   2100
         Width           =   375
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Inicial:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1740
         Width           =   450
      End
      Begin VB.Label lblFrame 
         AutoSize        =   -1  'True
         Caption         =   "C�digo dos Bancos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   41
         Top             =   1530
         Width           =   1650
      End
      Begin VB.Line linFrame 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   5760
         X2              =   120
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Line linFrame 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   5760
         X2              =   120
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Ini&cial:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   3600
         Width           =   450
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   18
         Top             =   960
         Width           =   765
      End
      Begin VB.Label lblCtrlFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Data Inicial:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   840
      End
   End
   Begin ComctlLib.TabStrip tabCtrlFinanc 
      Height          =   7275
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   12832
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sint�tico"
            Key             =   "sintetico"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Anal�tico"
            Key             =   "analitico"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Anual"
            Key             =   "anual"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Or�ado x Realizado"
            Key             =   "or�ado"
            Object.Tag             =   ""
            ImageVarType    =   2
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
End
Attribute VB_Name = "frptCtrlFinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CF_CONTAS$ = "Contas "
Private Const CF_CONTASQUITADAS$ = "Quitadas e em Aberto"
Private Const IDX_INICIO = 0        ' Utilizado para �ndices de matrizes
Private Const IDX_FINAL = 1         ' �dem
Private dblCotacao   As Double      'Vari�vel que verifica se existe cota��o para a Moeda na Data Indicada
Private mbolCancelou As Boolean     ' Verifica se o usu�rio cancelou a impress�o
Private NomeAuxiliar    As String
Private strData         As String
Private UsandoModelo    As Boolean

Private Sub cmdCtrlFinanc_Click(Index As Integer)
    If (Index < 2) Then
        If EData(txtCtrlFinanc(0).Text) And EData(txtCtrlFinanc(1).Text) Then
            cmdCtrlFinanc(0).Enabled = False
            cmdCtrlFinanc(1).Enabled = False
            cmdCtrlFinanc(2).Caption = LoadResString(170)
            'pt. 88454 - Ivo Sousa (18/09/2008)
            'Retirado o codigo a baixo por solicita��o do Dulcino
            'If (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
                'If chkDiscCentroCusto.Value = vbChecked Then chkDiscCentroCusto.Value = Unchecked
            'End If
            UsandoModelo = False
            If IsValid(txtCtrlFinanc(11).Text) Then
                UsandoModelo = True
            End If
            FinancFiltro IIf(Index, wrToPrinter, wrToWindow)
            cmdCtrlFinanc(0).Enabled = True
            cmdCtrlFinanc(1).Enabled = True
            cmdCtrlFinanc(2).Caption = LoadResString(169)
        Else
            MsgBox "Favor Verificar os Campos de Datas Inicial e Final.", vbExclamation, "Data Invalida"
        End If
    Else
        If cmdCtrlFinanc(0).Enabled Then
            Unload Me
        Else
            mbolCancelou = True
            SimpleMsgBar LoadResString(171) & LoadResString(14)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim intContador As Integer
    
    'Configurando a janela na abertura
    For intContador = 0 To 9
        lblNomes(intContador).Caption = NUL
    Next
    'Trazendo valores padr�o para os campos do formul�rio
    'Datas:
    txtCtrlFinanc(0).Text = Format$(Date, FDATA)
    txtCtrlFinanc(1).Text = Format$(Date, FDATA)

    'Bancos:
    txtCtrlFinanc(4).Text = GetFieldValue("MIN(Banco)", "Bancos", NUL, 0)
    txtCtrlFinanc(5).Text = GetFieldValue("MAX(Banco)", "Bancos", NUL, 0)

    'Contas:
    txtCtrlFinanc(6).Text = GetFieldValue("MIN([C�digo])", "Contas", NUL, 0)
    txtCtrlFinanc(7).Text = GetFieldValue("MAX([C�digo])", "Contas", NUL, 0)

    'Grupos:
    txtCtrlFinanc(2).Text = GetFieldValue("MIN([C�digo])", "Grupos", NUL, 0)
    txtCtrlFinanc(3).Text = GetFieldValue("MAX([C�digo])", "Grupos", NUL, 0)
    
    'Centros:
    txtCtrlFinanc(9).Text = GetFieldValue("MIN([C�digo])", "Centros", NUL, 0, 0)
    txtCtrlFinanc(10).Text = GetFieldValue("MAX([C�digo])", "Centros", NUL, 0, 0)
  
    'Origem Padr�o � Ambos
    cboOrigem.ListIndex = 0

    'Tipo padr�o Todas
    cboTipo.ListIndex = 0
  
    'Tipo padr�o = LIBERA��O
    cboTipoData.ListIndex = 1
  
    'Tipo Conciliado
    cboConciliado.AddItem "Todos"
    cboConciliado.AddItem "Sim"
    cboConciliado.AddItem "N�o"
    cboConciliado.Text = "Todos"

    'Verfica se o Centro de Custo deve ser exibido
    fraCentro.Enabled = CentrodeCusto(MFinanceiro)
  
    'Verifica se o Modelo deve ser exibido
    fraModelo.Visible = Configuracao("Utiliza modelos espec�ficos para Rel Controle Financeiro", False)
  
    'Definindo a primeira op��o visivel do formul�rio
    tabCtrlFinanc.Tabs(1).Selected = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frptCtrlFinanc = Nothing
    MsgBar MsgBoxCaption
End Sub

Private Sub tabCtrlFinanc_Click()
    'Apenas altero o Caption do Frame
    If (tabCtrlFinanc.SelectedItem.Key = "sintetico") Then
        fraTab.Caption = LoadResString(166)
        chkDiscCentroCusto.Visible = False
    ElseIf (tabCtrlFinanc.SelectedItem.Key = "analitico") Then
        fraTab.Caption = LoadResString(167)
        chkDiscCentroCusto.Visible = False
    ElseIf (tabCtrlFinanc.SelectedItem.Key = "anual") Then
        fraTab.Caption = LoadResString(168)
        chkDiscCentroCusto.Visible = False
    ElseIf (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
        fraTab.Caption = tabCtrlFinanc.SelectedItem.Caption
        chkDiscCentroCusto.Visible = True
    End If
    chkSaldoAnterior.Visible = (tabCtrlFinanc.SelectedItem.Key = "anual")
End Sub

Private Sub txtCtrlFinanc_Change(Index As Integer)
    Select Case Index
        Case 2, 3 'Grupos
            GetAssocValue "SELECT Descri��o FROM Grupos WHERE [C�digo] = " & txtCtrlFinanc(Index).Text, lblNomes(Index - 2)
        Case 4, 5 'Bancos
            GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtCtrlFinanc(Index).Text, lblNomes(Index - 2)
        Case 6, 7 'Contas
            GetAssocValue "SELECT Descri��o FROM Contas WHERE [C�digo] = " & txtCtrlFinanc(Index).Text, lblNomes(Index - 2)
        Case 8 'Moeda
            GetAssocValue "SELECT Descri��o, Moeda FROM Moedas WHERE Moeda = '" & txtCtrlFinanc(8).Text & "'", lblNomes(6), txtCtrlFinanc(8)
        Case 9, 10 'Centros
            GetAssocValue "SELECT Descri��o FROM Centros WHERE [C�digo] = " & txtCtrlFinanc(Index).Text, lblNomes(Index - 2)
        Case 11 'Modelos
            GetAssocValue "Select Descri��o from Modelos where [C�digo] = " & txtCtrlFinanc(Index).Text, lblNomes(9)
    End Select
End Sub

Private Sub txtCtrlFinanc_GotFocus(Index As Integer)
    Selecione txtCtrlFinanc(Index)
    FinancStatusMsg txtCtrlFinanc(Index).TabIndex
End Sub

Private Sub txtCtrlFinanc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strBcoInicial As String
    Dim strBcoFinal   As String
  
    If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
        Select Case Index
            Case 2, 3 'Grupos
                PCampo "Grupos de Contas", "Grupos", pbCampo, txtCtrlFinanc(Index), 0
            Case 4, 5 'Bancos
                PCampo "Bancos", "Bancos", pbCampo, txtCtrlFinanc(Index), 0
            Case 6, 7 'Contas
                If (IsValid(txtCtrlFinanc(2).Text) Or IsValid(txtCtrlFinanc(3).Text)) Then
                    strBcoInicial = IIf(CLngDef(txtCtrlFinanc(2).Text), txtCtrlFinanc(2).Text, "1")
                    strBcoFinal = IIf(CLngDef(txtCtrlFinanc(3).Text), txtCtrlFinanc(3).Text, "999999999")
                    PCampo "Contas", "SELECT * FROM Contas WHERE Grupo BETWEEN " & strBcoInicial & " AND " & strBcoFinal, pbCampo, txtCtrlFinanc(Index), 0
                Else
                    PCampo "Contas", "Contas", pbCampo, txtCtrlFinanc(Index), 0
                End If
            Case 8
                PCampo "Moedas e �ndices", "Moedas", PB_CAMPO, txtCtrlFinanc(8), "Moeda"
            Case 9, 10 'Centros
                PCampo "Centros", "Centros", pbCampo, txtCtrlFinanc(Index), 0
            Case 11 'Modelos
                PCampo "Modelos", "Modelos", pbCampo, txtCtrlFinanc(Index), "[C�digo]"
        End Select
    End If
End Sub

Private Sub txtCtrlFinanc_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 1 'Campos de data
            SetMascara KeyAscii, txtCtrlFinanc(Index).SelStart, MASK_DATE4
        Case 2 'Campos de C�digo de Grupos, Inicial e Final
            SetMascara KeyAscii, txtCtrlFinanc(2).SelStart, fMask("Grupos", "C�digo")
        Case 3
            SetMascara KeyAscii, txtCtrlFinanc(3).SelStart, fMask("Grupos", "C�digo"), txtCtrlFinanc(2).hWnd
        Case 4 'Campos de C�digo do Banco, Inicial e Final
            SetMascara KeyAscii, txtCtrlFinanc(4).SelStart, fMask("Bancos", "Banco")
        Case 5
            SetMascara KeyAscii, txtCtrlFinanc(5).SelStart, fMask("Bancos", "Banco"), txtCtrlFinanc(4).hWnd
        Case 6 'Campos de C�digo de Contas, Inicial e Final
            SetMascara KeyAscii, txtCtrlFinanc(6).SelStart, fMask("Contas", "C�digo")
        Case 7
            SetMascara KeyAscii, txtCtrlFinanc(7).SelStart, fMask("Contas", "C�digo"), txtCtrlFinanc(6).hWnd
        Case 9 'C�digo do Centro de Custo, Inicial e Final
            SetMascara KeyAscii, txtCtrlFinanc(9).SelStart, fMask("Centros", "C�digo")
        Case 10
            SetMascara KeyAscii, txtCtrlFinanc(10).SelStart, fMask("Centros", "C�digo"), txtCtrlFinanc(9).hWnd
    End Select
End Sub

' SUB.......: FinancStatusMsg
' Objetivo..: Exibe mensagens de ajuda na barra de Status do Sistema
' Argumento.: [intTabIndex]: Valor da propriedade TabIndex do Controle.
Private Sub FinancStatusMsg(intTabIndex As Integer)
    Select Case intTabIndex
        Case 6 'Data Inicial:
            MsgBar ResolveResString(161, resUM, "de " & strData)
        Case 8 'Data Final:
            MsgBar ResolveResString(162, resUM, "de " & strData)
        Case 10 'Tipos de Conta:
            MsgBar LoadResString(163)
        Case 12 'Centro do Custo:
            MsgBar LoadResString(156) & ResolveResString(75, resUM, "Centro de Custo")
        Case 15, 18 'Banco
            MsgBar LoadResString(152) & ResolveResString(75, resUM, "Bancos")
        Case 22, 25 'Conta
            MsgBar LoadResString(164) & ResolveResString(75, resUM, "Contas")
        Case 29, 32 'Grupo
            MsgBar LoadResString(165) & ResolveResString(75, resUM, "Grupos de Conta")
    End Select
End Sub

' SUB.......: FinancFiltro
' Objetivo..: Cria a string de Instru��o Select que ser� utilizada para filtrar
'             os dados de busca e criar o arquivo tempor�rio para impress�o do
'             relat�rio.
' Argumento.: [pdImpressao]: Destino da impress�o.
Private Sub FinancFiltro(pdImpressao As PrintDestinoEnum)
    Dim rstContas     As Object
    Dim rstAux        As Object
    Dim strContas     As String
    Dim dtDatas(1)    As Date        'Data inicial e Final
    Dim lBancos(1)    As Long        'Bancos Inicial e Final
    Dim lContas(1)    As Long        'Contas Inicial e Final
    Dim lGrupos(1)    As Long        'Grupos Inicial e Final
    Dim dtInicial     As Date
    Dim dtFinal       As Date
    Dim lngModelo     As Long
    Dim strSubTitulo  As String        'Sub t�tulo do relat�rio
    Dim strSubTitulo2 As String        'Sub t�tulo do relat�rio
    Dim dtTmp         As Date
    Dim SaldoAnterior As Currency
    
    SetPtr vbHourglass
    SimpleMsgBar LoadResString(13) & LoadResString(14)
    mbolCancelou = False
    If cboTipoData.Text = "Libera��o" Then
        strData = "Libera��o"
    ElseIf cboTipoData.Text = "Emiss�o" Then
        strData = "Emiss�o"
    ElseIf cboTipoData.Text = "Vencimento" Then
        strData = "Vencimento"
    ElseIf cboTipoData.Text = "Pagamento" Then
        strData = "Pagamento"
    End If
    If (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
        If EData(txtCtrlFinanc(0).Text) Then
            txtCtrlFinanc(0).Text = FirstDayS(txtCtrlFinanc(0).Text)
            txtCtrlFinanc(1).Text = LastDayS(txtCtrlFinanc(1).Text)
        End If
    End If
    dtInicial = CDateDef(txtCtrlFinanc(0).Text)
    dtFinal = CDateDef(txtCtrlFinanc(1).Text)
    dblCotacao = TemCotacao(txtCtrlFinanc(8).Text, lblNomes(6).Caption, dtInicial, dtFinal)
    'Verifica se a Moeda Informada � v�lida antes de executar a Convers�o
    If lblNomes(6).Caption = NUL And txtCtrlFinanc(8).Text <> NUL Then
        MsgBox "Informe uma MOEDA v�lida para a Convers�o de Valores", vbOKOnly Or vbExclamation, MsgBoxCaption
        LetFocus txtCtrlFinanc(8).Text
        Selecione txtCtrlFinanc(8)
        mbolCancelou = True
        Exit Sub
    End If
    'Verifica se a Moeda Informada tem Cota��o
    If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) Then
        If dblCotacao = 0 Then
            MsgBox "Informe uma Cota��o v�lida para a Moeda '" & txtCtrlFinanc(8).Text & "' na Data de " & txtCtrlFinanc(0).Text & " At� " & txtCtrlFinanc(1).Text, vbOKOnly Or vbExclamation, MsgBoxCaption
            LetFocus txtCtrlFinanc(8).hWnd
            Selecione txtCtrlFinanc(8)
            mbolCancelou = True
            Exit Sub
        End If
    End If
    If UsandoModelo Then
        lngModelo = CLngDef(txtCtrlFinanc(11).Text)
        strContas = "SELECT [Grupos Auxiliares].[C�digo] as Grupo, [Grupos Auxiliares].Descri��o as DescGrupo, " & _
                 "[Contas Auxiliares].[C�digo] as ContaAuxiliar, [Contas Auxiliares].Descri��o as Descri��o, " & _
                 "[Contas de Contas Auxiliares].[Conta Cont�bil] as [C�digo] " & _
                 "FROM (([Grupos de Modelos] INNER JOIN [Contas Auxiliares] " & _
                 "ON [Grupos de Modelos].Grupo = [Contas Auxiliares].Grupo) " & _
                 "INNER JOIN [Contas de Contas Auxiliares] " & _
                 "ON [Contas Auxiliares].[C�digo] = [Contas de Contas Auxiliares].Conta) " & _
                 "INNER JOIN [Grupos Auxiliares] ON [Grupos de Modelos].Grupo = [Grupos Auxiliares].[C�digo] " & _
                 "WHERE [Grupos de Modelos].Modelo = " & lngModelo
    Else
        strContas = "SELECT * FROM Contas WHERE "
        'Se o usu�rio filtrou por Grupo
        lGrupos(IDX_INICIO) = Min(CLngDef(txtCtrlFinanc(2).Text), CLngDef(txtCtrlFinanc(3).Text))
        lGrupos(IDX_FINAL) = Max(CLngDef(txtCtrlFinanc(2).Text), CLngDef(txtCtrlFinanc(3).Text))
        If ((lGrupos(IDX_INICIO) > 0) And (lGrupos(IDX_FINAL) > 0)) Then
            If (lGrupos(IDX_INICIO) = lGrupos(IDX_FINAL)) Then
                AppendStr strContas, "Grupo = " & CStr(lGrupos(IDX_INICIO))
            Else
                Concat strContas, "(Grupo BETWEEN ", CStr(lGrupos(IDX_INICIO)), " AND ", CStr(lGrupos(IDX_FINAL)), ")"
            End If
        ElseIf (lGrupos(IDX_INICIO) > 0) Then
            AppendStr strContas, "Grupo >= " & CStr(lGrupos(IDX_INICIO))
        ElseIf (lGrupos(IDX_FINAL) > 0) Then
            AppendStr strContas, "Grupo <= " & CStr(lGrupos(IDX_FINAL))
        Else
            AppendStr strContas, "Grupo >= 1"
        End If
        'Se o usu�rio filtrou por conta
        lContas(IDX_INICIO) = Min(CLngDef(txtCtrlFinanc(6).Text), CLngDef(txtCtrlFinanc(7).Text))
        lContas(IDX_FINAL) = Max(CLngDef(txtCtrlFinanc(6).Text), CLngDef(txtCtrlFinanc(7).Text))
        If ((lContas(IDX_INICIO) > 0) And (lContas(IDX_FINAL) > 0)) Then
            If (lContas(IDX_INICIO) = lContas(IDX_FINAL)) Then
                AppendStr strContas, " AND [C�digo] = " & CStr(lContas(IDX_INICIO))
            Else
                Concat strContas, " AND ([C�digo] BETWEEN ", CStr(lContas(IDX_INICIO)), " AND ", CStr(lContas(IDX_FINAL)), ")"
            End If
        ElseIf ((lContas(IDX_INICIO) > 0) And (lContas(IDX_FINAL) = 0)) Then
            AppendStr strContas, " AND [C�digo] >= " & CStr(lContas(IDX_INICIO))
        ElseIf ((lContas(IDX_INICIO) = 0) And (lContas(IDX_FINAL) > 0)) Then
            AppendStr strContas, " AND [C�digo] <= " & CStr(lContas(IDX_FINAL))
        End If
        'Ordenando os dados
        AppendStr strContas, " ORDER BY Grupo, [C�digo];"
    End If
  
    If (WL_OK = AbreRecordset(rstContas, strContas, dbOpenSnapshot)) Then
        'Resolvendo o Filtro de Bancos
        lBancos(IDX_INICIO) = Min(CLngDef(txtCtrlFinanc(4).Text), CLngDef(txtCtrlFinanc(5).Text))
        lBancos(IDX_FINAL) = Max(CLngDef(txtCtrlFinanc(4).Text), CLngDef(txtCtrlFinanc(5).Text))
        'Resolvendo o filtro de Data
        dtDatas(IDX_INICIO) = Empty
        dtDatas(IDX_FINAL) = Empty
        If (IsValid(txtCtrlFinanc(0).Text)) Then
            If (EData(txtCtrlFinanc(0).Text)) Then
                dtDatas(IDX_INICIO) = CDate(txtCtrlFinanc(0).Text)
            Else
                MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, "Data Inicial")
                GoTo FinancFiltro_Erro
            End If
        Else
            dtDatas(IDX_INICIO) = Empty
        End If
        If (IsValid(txtCtrlFinanc(1).Text)) Then
            If (EData(txtCtrlFinanc(1).Text)) Then
                dtDatas(IDX_FINAL) = CDate(txtCtrlFinanc(1).Text)
            Else
                MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, "Data Final")
                GoTo FinancFiltro_Erro
            End If
        Else
            dtDatas(IDX_FINAL) = Date
        End If
        If Len(dtDatas(IDX_INICIO)) > 0 And Len(dtDatas(IDX_FINAL)) > 0 Then
            If CDateDef(dtDatas(IDX_FINAL)) < CDateDef(dtDatas(IDX_INICIO)) Then
                MsgFunc "Data Final menor que Data Inicial"
                GoTo FinancFiltro_Erro
            End If
        End If
        If (Not (IsEmptyDate(dtDatas(IDX_INICIO)))) Then
            If (DateDiff("d", dtDatas(IDX_INICIO), dtDatas(IDX_FINAL)) < ZERO) Then
                dtTmp = dtDatas(IDX_INICIO)
                dtDatas(IDX_INICIO) = dtDatas(IDX_FINAL)
                dtDatas(IDX_FINAL) = dtTmp
            End If
        End If
        If Len(dtDatas(IDX_INICIO)) > 0 And Len(dtDatas(IDX_FINAL)) > 0 Then
            If CDateDef(dtDatas(IDX_FINAL)) < CDateDef(dtDatas(IDX_INICIO)) Then
                MsgFunc "Data Final menor que Data Inicial"
            End If
        End If
        If (tabCtrlFinanc.SelectedItem.Key = "anual") Then
            'Se o relat�rio for anual e o usu�rio n�o indicar a data final, a data final ser� o hoje.
            If (IsEmptyDate(dtDatas(IDX_FINAL))) Then
                dtDatas(IDX_FINAL) = Date
            End If
            'Se a data inicial n�o foi indicada def�no-a como um ano antes da data final
            If (IsEmptyDate(dtDatas(IDX_INICIO))) Then
                dtDatas(IDX_INICIO) = DateAdd("yyyy", -1, dtDatas(IDX_FINAL))
            End If
        End If
        'Resolvendo o sub t�tulo do relat�rio
        'Titulo 1
        strSubTitulo = Choose((cboOrigem.ListIndex + 1), "Duplicatas e Lan�amentos ", "Duplicatas ", "Lan�amentos ")
        AppendStr strSubTitulo, Choose((cboTipo.ListIndex + 1), "", "� Pagar ", "� Receber ")
        If (IsValid(txtCtrlFinanc(0).Text)) Then
            AppendStr strSubTitulo, " de " & txtCtrlFinanc(0).Text
        End If
        If (IsValid(txtCtrlFinanc(1).Text)) Then
            AppendStr strSubTitulo, " at� " & txtCtrlFinanc(1).Text
        End If
        If CentrodeCusto(MFinanceiro) Then
            If (IsValid(txtCtrlFinanc(9).Text)) Or (IsValid(txtCtrlFinanc(10).Text)) Then
                AppendStr strSubTitulo, " -  Centros de Custo"
            End If
            If (IsValid(txtCtrlFinanc(9).Text)) Then
                AppendStr strSubTitulo, " de " & txtCtrlFinanc(9).Text
            End If
            If (IsValid(txtCtrlFinanc(10).Text)) Then
                AppendStr strSubTitulo, " at� " & txtCtrlFinanc(10).Text
            End If
        End If
        AppendStr strSubTitulo, Choose((cboConciliado.ListIndex + 1), " ", " / Conciliados", " / N�o Conciliados")
        AppendStr strSubTitulo, Choose((cboTipoData.ListIndex + 1), " / Por Emiss�o", " / Por Libera��o", " / Por Vencimento", " / Por Pagamento")
        
        'Titulo 2
        strSubTitulo2 = "Banco "
        If (IsValid(txtCtrlFinanc(4).Text)) Then
            AppendStr strSubTitulo2, " de " & txtCtrlFinanc(4).Text
        End If
        If (IsValid(txtCtrlFinanc(5).Text)) Then
            AppendStr strSubTitulo2, " at� " & txtCtrlFinanc(5).Text
        End If
        AppendStr strSubTitulo2, " / Conta "
        
        If (IsValid(txtCtrlFinanc(6).Text)) Then
            AppendStr strSubTitulo2, " de " & txtCtrlFinanc(6).Text
        End If
        If (IsValid(txtCtrlFinanc(7).Text)) Then
            AppendStr strSubTitulo2, " at� " & txtCtrlFinanc(7).Text
        End If
        AppendStr strSubTitulo2, " / Grupo "
        If (IsValid(txtCtrlFinanc(2).Text)) Then
            AppendStr strSubTitulo2, " de " & txtCtrlFinanc(2).Text
        End If
        If (IsValid(txtCtrlFinanc(3).Text)) Then
            AppendStr strSubTitulo2, " at� " & txtCtrlFinanc(3).Text
        End If
        
        'Cria a tabela auxiliar para gravar os dados a serem impressos
        If CriaTabelaTemp(rstAux, dtDatas()) Then
            'Grava os dados na tabela auxiliar e imprime o relat�rio. Como a fun��o
            'que grava os dados � a mais demorada, ela � que verifica se o usu�rio
            'cancelou a impress�o.
            Select Case tabCtrlFinanc.SelectedItem.Key
                
                Case "or�ado"
                    'Se data inicial for Janeiro e data final for Dezembro
                    If (Month(dtDatas(0)) = 1 And Month(dtDatas(1)) = 12) Then
                        If AppendTempOrcado(rstAux, rstContas, lBancos(), dtDatas()) Then
                            RelatorioOrcado pdImpressao, rstAux, strSubTitulo, strSubTitulo2
                        End If
                    Else
                        If AppendTemp(rstAux, rstContas, lBancos(), dtDatas()) Then
                            RelatorioSintetico pdImpressao, rstAux, strSubTitulo, strSubTitulo2
                        End If
                    End If
                   
                Case "sintetico"
                    If AppendTemp(rstAux, rstContas, lBancos(), dtDatas()) Then
                        RelatorioSintetico pdImpressao, rstAux, strSubTitulo, strSubTitulo2
                    End If
                    
                Case "anual"
                    If AppendTempAnual(rstAux, rstContas, SaldoAnterior, lBancos(), dtDatas()) Then
                        RelatorioAnual pdImpressao, rstAux, dtDatas(), SaldoAnterior
                    End If
                    
                Case "analitico"
                    If (AppendTempAnalitico(rstAux, rstContas, lBancos(), dtDatas())) Then
                        RelatorioAnalitico rstAux, pdImpressao, strSubTitulo, strSubTitulo2
                    End If
            End Select
        End If
        If (IsValid(NomeAuxiliar)) And (tabCtrlFinanc.SelectedItem.Key = "or�ado") And ((Month(dtDatas(0)) = 1 And Month(dtDatas(1)) = 12)) Then
            DeleteAux rstAux, NomeAuxiliar
        Else
            DeleteAux rstAux, NUL
        End If
    End If
    SetPtr vbDefault
    Exit Sub
FinancFiltro_Erro:
  FechaRecordset rstContas
  SetPtr vbDefault
  MsgBar Caption
End Sub

'FUNCTION..: AddTransfBancarias
'Objetivo..: Adiciona os dados de transfer�ncia Banc�ria � tabela
'            auxiliar para gera��o do relat�rio
'Argumentos: [lConta]: Conta de sele��o.
'            [dDatas]: Matriz com as datas inicial e final.
'            [lBanco]: Matriz com os c�digos dos Bancos.
'            [bSrc]  : True para Bancos de Origem, False para Destino.
'Retorna...: Uma String contendo os filtros para a tabela.
Private Function AddTransfBancarias(lConta As Double, dDatas() As Date, lBanco() As Long, bSrc As Boolean, Optional CentroCusto As Long) As String
Dim strTransf As String                   'String para a instru��o Select
Dim strBanco  As String

  If (lConta = 0) Then Exit Function
  
  If bSrc Then
    strBanco = "Origem"
  Else
    strBanco = "Destino"
  End If
  
  strTransf = "Conta = " & CStr(lConta)
   
  '
  ' Resolvendo o Filtro de Datas
  '
  If ((Not IsEmpty(dDatas(IDX_INICIO))) And (Not IsEmpty(dDatas(IDX_FINAL)))) Then
    If (DateDiff("d", dDatas(IDX_INICIO), dDatas(IDX_FINAL)) = 0) Then
      If (strTransf <> "") Then
          Concat strTransf, " AND "
      End If
      AppendStr strTransf, " (Data = " & InverteData(dDatas(IDX_INICIO), True) & ")"
    Else
      If (strTransf <> "") Then
        Concat strTransf, " AND "
      End If
      Concat strTransf, " (Data BETWEEN ", InverteData(dDatas(IDX_INICIO), True), _
             " AND ", InverteData(dDatas(IDX_FINAL), True), ")"
    End If
  ElseIf (Not IsEmpty(dDatas(IDX_INICIO))) Then
    If (strTransf <> "") Then
        Concat strTransf, " AND "
    End If
    AppendStr strTransf, " (Data >= " & InverteData(dDatas(IDX_INICIO), True) & ")"
  ElseIf (Not IsEmpty(dDatas(IDX_FINAL))) Then
    If (strTransf <> "") Then
        Concat strTransf, " AND "
    End If
    AppendStr strTransf, " (Data <= " & InverteData(dDatas(IDX_FINAL), True) & ")"
  End If
  '
  ' Resolve o filtro de Bancos
  '
    
  If ((lBanco(IDX_INICIO) > 0) And (lBanco(IDX_FINAL) > 0)) Then
    If (lBanco(IDX_INICIO) = lBanco(IDX_FINAL)) Then
      If (strTransf <> "") Then
        Concat strTransf, " AND "
      End If
      Concat strTransf, strBanco, " = ", CStr(lBanco(IDX_INICIO))
    Else
      If (strTransf <> "") Then
        Concat strTransf, " AND "
      End If
      Concat strTransf, " ( ", strBanco, " BETWEEN ", CStr(lBanco(IDX_INICIO)), " AND ", CStr(lBanco(IDX_FINAL)), ")"
    End If
  ElseIf (lBanco(IDX_INICIO) > 0) Then
    If (strTransf <> "") Then
      Concat strTransf, " AND "
    End If
    Concat strTransf, strBanco, " >= ", CStr(lBanco(IDX_INICIO))
  ElseIf (lBanco(IDX_FINAL) > 0) Then
    If (strTransf <> "") Then
      Concat strTransf, " AND "
    End If
    Concat strTransf, strBanco, " <= ", CStr(lBanco(IDX_FINAL))

  End If
  '
  ' Resolvendo o Centro de Custo
  '

    
  If (tabCtrlFinanc.SelectedItem.Key = "or�ado") And CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
    If (strTransf <> "") Then
      Concat strTransf, " AND "
    End If
    Concat strTransf, " Centro = ", CentroCusto
  Else
    If (CentrodeCusto(MFinanceiro) And (IsValid(txtCtrlFinanc(9).Text) Or IsValid(txtCtrlFinanc(10).Text))) Then
      If (IsValid(txtCtrlFinanc(9).Text) And IsValid(txtCtrlFinanc(10).Text)) Then
        If txtCtrlFinanc(9).Text = txtCtrlFinanc(10).Text Then
          If (strTransf <> "") Then
            Concat strTransf, " AND "
          End If
          Concat strTransf, " Centro = ", txtCtrlFinanc(9).Text
        Else
          If (strTransf <> "") Then
            Concat strTransf, " AND "
          End If
          Concat strTransf, " (Centro BETWEEN ", txtCtrlFinanc(9).Text, " AND ", txtCtrlFinanc(10).Text, ")"
        End If
      ElseIf (IsValid(txtCtrlFinanc(9).Text)) Then
        If (strTransf <> "") Then
          Concat strTransf, " AND "
        End If
        Concat strTransf, " Centro >= ", txtCtrlFinanc(9).Text
      ElseIf (IsValid(txtCtrlFinanc(10).Text)) Then
        If (strTransf <> "") Then
          Concat strTransf, " AND "
        End If
        Concat strTransf, "  Centro <= ", txtCtrlFinanc(10).Text
      End If
    End If
  End If
  
  '
  ' Retorna a instru��o construida
  '
  AddTransfBancarias = strTransf
  
End Function

' FUNCTION..: AddAplicacoes
' Objetivo..: Cria a instru��o de filtro para os dados da tabela de aplica��es.
' Argumentos: [lConta]: C�digo da conta.
'             [lBco]  : Matriz com os c�digos dos Bancos.
'             [dDta]  : Matriz com as datas.
'             [bCred] : True para opera��es de Cr�dito, False para d�bito.
' Retorna...: Uma string contendo o filtro para os dados da tabela.
' ------------------------------------------------------------------------------
Private Function AddAplicacoes(lConta As Double, lBco() As Long, dDta() As Date, bCred As Boolean, Optional CentroCusto As Long) As String
Dim strAplic As String            'Para montar o filtro

  If (lConta = 0) Then Exit Function
  '
  ' Resolve a conta
  '
  strAplic = "Conta = " & CStr(lConta)
  '
  ' Resolve os Bancos, inicial e final
  '
  If ((lBco(IDX_INICIO) > 0) And (lBco(IDX_FINAL) > 0)) Then
    If (lBco(IDX_INICIO) = lBco(IDX_FINAL)) Then
      AppendStr strAplic, " AND Banco = " & CStr(lBco(IDX_INICIO))
    Else
      Concat strAplic, " AND (Banco BETWEEN ", CStr(lBco(IDX_INICIO)), " AND ", CStr(lBco(IDX_FINAL)), ")"
    End If
  ElseIf (lBco(IDX_INICIO) > 0) Then
    AppendStr strAplic, " AND Banco >= " & CStr(lBco(IDX_INICIO))
  ElseIf (lBco(IDX_FINAL) > 0) Then
    AppendStr strAplic, " AND Banco <= ", CStr(lBco(IDX_FINAL))
  End If
  '
  ' Resolve as datas
  '
  If ((Not IsEmpty(dDta(IDX_INICIO))) And (Not IsEmpty(dDta(IDX_FINAL)))) Then
    If (DateDiff("d", dDta(IDX_INICIO), dDta(IDX_FINAL)) = 0) Then
      AppendStr strAplic, " AND Data = " & InverteData(dDta(IDX_INICIO), True)
    Else
      Concat strAplic, " AND (Data BETWEEN ", InverteData(dDta(IDX_INICIO), True), " AND ", _
                       InverteData(dDta(IDX_FINAL), True), ")"
    End If
  ElseIf (Not IsEmpty(dDta(IDX_INICIO))) Then
    AppendStr strAplic, " AND Data >= " & InverteData(dDta(IDX_INICIO), True)
  ElseIf (Not IsEmpty(dDta(IDX_FINAL))) Then
    AppendStr strAplic, " AND Data <= " & InverteData(dDta(IDX_FINAL), True)
  End If
  '
  ' Resolvendo o Centro de Custo
  '
  If (tabCtrlFinanc.SelectedItem.Key = "or�ado") And CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
    Concat strAplic, " AND Centro = ", CentroCusto
  Else
    If (CentrodeCusto(MFinanceiro) And (IsValid(txtCtrlFinanc(9).Text) Or IsValid(txtCtrlFinanc(10).Text))) Then
      If (IsValid(txtCtrlFinanc(9).Text) And IsValid(txtCtrlFinanc(10).Text)) Then
        If txtCtrlFinanc(9).Text = txtCtrlFinanc(10).Text Then
          Concat strAplic, " AND Centro = ", txtCtrlFinanc(9).Text
        Else
          Concat strAplic, " AND (Centro BETWEEN ", txtCtrlFinanc(9).Text, " AND ", txtCtrlFinanc(10).Text, ")"
        End If
      ElseIf (IsValid(txtCtrlFinanc(9).Text)) Then
        Concat strAplic, " AND Centro >= ", txtCtrlFinanc(9).Text
      
      ElseIf (IsValid(txtCtrlFinanc(10).Text)) Then
        Concat strAplic, " AND Centro <= ", txtCtrlFinanc(10).Text
      
      End If
    End If
  End If
  '
  ' Se forem opera��es de Cr�dito o campo Tipo deve conter o valor: Juros/Corre��o.
  ' Todos os outros valores deste campo s�o considerados d�bito.
  '
  If bCred Then
    AppendStr strAplic, " AND Tipo = 'Juros/Corre��o'"
  Else
    AppendStr strAplic, " AND Tipo <> 'Juros/Corre��o'"
  End If
  
  '
  ' Retorna a string de filtro
  '
  AddAplicacoes = strAplic
  
End Function

' FUNCTION..: AddLancDupl
' Objetivo..: Resolve a express�o de filtro para Lan�amentos e Duplicatas
' Argumentos: [lCta]     : C�digo da Conta.
'             [lngBancos]: Matriz com os Bancos.
'             [datDatas] : Matriz com as datas.
'             [bolPagos] : True para Contas Pagas, False para Recebidos.
'             [Realizado]: Se for 0 (Entradas e Sa�das), se for 1 (A Entrar e A Sair) e se for o padr�o = 2 ent�o Tudo
' Retorna...: A string de filtro de dados.
' ---------------------------------------------------------------------------------
Private Function AddLancDupl(lCta As Double, lngBancos() As Long, datDatas() As Date, bolPagos As Boolean, Optional CentroCusto As Long, Optional Realizado As Long = 2) As String
  
  Dim strld As String
  
  If (lCta = 0) Then Exit Function
  '
  strld = "Conta = " & CStr(lCta)
  
  
  '
  ' Resolve as datas
  '
  If ((Not IsEmpty(datDatas(IDX_INICIO))) And (Not IsEmpty(datDatas(IDX_FINAL)))) Then
    If (DateDiff("d", datDatas(IDX_INICIO), datDatas(IDX_FINAL)) = 0) Then
      AppendStr strld, " AND (" & strData & " = " & InverteData(datDatas(IDX_INICIO), True) & ")"
    Else
      Concat strld, " AND (" & strData & " BETWEEN ", InverteData(datDatas(IDX_INICIO), True), _
                    " AND ", InverteData(datDatas(IDX_FINAL), True), ")"
    End If
  ElseIf (Not IsEmpty(datDatas(IDX_INICIO))) Then
    AppendStr strld, " AND (" & strData & " >= " & InverteData(datDatas(IDX_INICIO), True) & ")"
  ElseIf (Not IsEmpty(datDatas(IDX_FINAL))) Then
    AppendStr strld, " AND (" & strData & " <= " & InverteData(datDatas(IDX_FINAL), True) & ")"
  End If
  '
  ' Resolve o Banco
  '
  If ((lngBancos(IDX_INICIO) > 0) And (lngBancos(IDX_FINAL) > 0)) Then
    If (lngBancos(IDX_INICIO) = lngBancos(IDX_FINAL)) Then
      AppendStr strld, " AND Banco = " & CStr(lngBancos(IDX_INICIO))
    Else
      Concat strld, " AND (Banco BETWEEN ", CStr(lngBancos(IDX_INICIO)), _
                    " AND ", CStr(lngBancos(IDX_FINAL)), ")"
    End If
  ElseIf (lngBancos(IDX_INICIO) > 0) Then
    AppendStr strld, " AND Banco >= " & CStr(lngBancos(IDX_INICIO))
  ElseIf (lngBancos(IDX_FINAL) > 0) Then
    AppendStr strld, " AND Banco <= " & CStr(lngBancos(IDX_FINAL))
  End If
  '
  ' Resolvendo as contas Quitadas ou Em Aberto: 0 = Em Aberto; 1 = Quitadas; 2 = Todas
  
   Select Case Realizado   '2 � o valor padr�o ou seja TODAS
     Case 0   'Contas j� Recebidas ou Pagas (Em Aberto)
         AppendStr strld, " AND (Pagamento IS NULL)"
     Case 1   'Contas a Pagar ou a Receber (Quitadas)
         AppendStr strld, " AND (Not (Pagamento IS NULL))"
   End Select
'
  ' Resolvendo o Centro de Custo
  '
  If (tabCtrlFinanc.SelectedItem.Key = "or�ado") And CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
    Concat strld, " AND Centro = ", CentroCusto
  Else
    If (CentrodeCusto(MFinanceiro) And (IsValid(txtCtrlFinanc(9).Text) Or IsValid(txtCtrlFinanc(10).Text))) Then
      If (IsValid(txtCtrlFinanc(9).Text) And IsValid(txtCtrlFinanc(10).Text)) Then
        If txtCtrlFinanc(9).Text = txtCtrlFinanc(10).Text Then
          Concat strld, " AND Centro = ", txtCtrlFinanc(9).Text
        Else
          Concat strld, " AND (Centro BETWEEN ", txtCtrlFinanc(9).Text, " AND ", txtCtrlFinanc(10).Text, ")"
        End If
      ElseIf (IsValid(txtCtrlFinanc(9).Text)) Then
        Concat strld, " AND Centro >= ", txtCtrlFinanc(9).Text
      
      ElseIf (IsValid(txtCtrlFinanc(10).Text)) Then
        Concat strld, " AND Centro <= ", txtCtrlFinanc(10).Text
      
      End If
    End If
  End If

  '
  ' Contas Pagas ou Recebidas
  '
  If bolPagos Then
    AppendStr strld, " AND PagRec = 'P'"
  Else
    AppendStr strld, " AND PagRec = 'R'"
  End If
  
  ' Conciliados
  '
  If cboConciliado.Text = "Sim" Then
    AppendStr strld, " AND Conciliado = True"
  End If
  If cboConciliado.Text = "N�o" Then
    AppendStr strld, " AND Conciliado = False"
  End If
  
  ' Retorna a instru��o de Filtro
  '
  AddLancDupl = strld
  
End Function

'FUNCTION..: CriaTabelaTemp
'Objetivo..: Cria a tabela auxiliar onde ser�o gravados os dados que ser�o
'            impressos.
'Argumento.: [rstTemp]: Vari�vel Recordset que receber� uma refer�ncia a tabela
'                       criada.
'Retorna...: True se puder criar a tabela com sucesso, False se n�o.
Private Function CriaTabelaTemp(rstTemp As Object, dDtAux() As Date) As Boolean
    If (tabCtrlFinanc.SelectedItem.Key = "or�ado") And (Month(dDtAux(0)) = 1 And Month(dDtAux(1)) = 12) Then
        Dim fsControle(29) As FieldStruct
        
        AppendVar fsControle(0), "GrupoC�digo", dbLong
        AppendVar fsControle(1), "GrupoNome", dbText, 60
        AppendVar fsControle(2), "ContaC�digo", dbLong
        AppendVar fsControle(3), "ContaNome", dbText, 40
        AppendVar fsControle(4), "Saldo1", dbCurrency
        AppendVar fsControle(5), "Saldo2", dbCurrency
        AppendVar fsControle(6), "Saldo3", dbCurrency
        AppendVar fsControle(7), "Saldo4", dbCurrency
        AppendVar fsControle(8), "Saldo5", dbCurrency
        AppendVar fsControle(9), "Saldo6", dbCurrency
        AppendVar fsControle(10), "Saldo7", dbCurrency
        AppendVar fsControle(11), "Saldo8", dbCurrency
        AppendVar fsControle(12), "Saldo9", dbCurrency
        AppendVar fsControle(13), "Saldo10", dbCurrency
        AppendVar fsControle(14), "Saldo11", dbCurrency
        AppendVar fsControle(15), "Saldo12", dbCurrency
        AppendVar fsControle(16), "Or�ado1", dbCurrency
        AppendVar fsControle(17), "Or�ado2", dbCurrency
        AppendVar fsControle(18), "Or�ado3", dbCurrency
        AppendVar fsControle(19), "Or�ado4", dbCurrency
        AppendVar fsControle(20), "Or�ado5", dbCurrency
        AppendVar fsControle(21), "Or�ado6", dbCurrency
        AppendVar fsControle(22), "Or�ado7", dbCurrency
        AppendVar fsControle(23), "Or�ado8", dbCurrency
        AppendVar fsControle(24), "Or�ado9", dbCurrency
        AppendVar fsControle(25), "Or�ado10", dbCurrency
        AppendVar fsControle(26), "Or�ado11", dbCurrency
        AppendVar fsControle(27), "Or�ado12", dbCurrency
        AppendVar fsControle(28), "TotalSaldo", dbCurrency
        AppendVar fsControle(29), "TotalOrcado", dbCurrency
        If CrieAux(rstTemp, fsControle()) Then
            CriaTabelaTemp = True
        End If
    Else
        Dim fsControle1(20) As FieldStruct
        
        AppendVar fsControle1(0), "GrupoC�digo", dbLong
        AppendVar fsControle1(1), "GrupoNome", dbText, 60
        'Projeto: 100340 - Desenv.: 100823 - Ueder Budni (18/11/2015)
        AppendVar fsControle1(2), "ContaC�digo", dbLong
        
        AppendVar fsControle1(3), "ContaNome", dbText, 40
        AppendVar fsControle1(4), "Sa�da", dbCurrency
        AppendVar fsControle1(5), "Entrada", dbCurrency
        AppendVar fsControle1(6), "Saldo", dbCurrency
        AppendVar fsControle1(7), "MesAno", dbDate
        AppendVar fsControle1(8), "Parcela", dbInteger
        AppendVar fsControle1(9), "C�digo", dbDouble
        AppendVar fsControle1(10), "Descri��o", dbText, 80
        AppendVar fsControle1(11), "Origem", dbText, 30
        AppendVar fsControle1(12), "Empresa", dbText, 15
        AppendVar fsControle1(13), "Or�ado", dbCurrency
        AppendVar fsControle1(14), "CentroC�digo", dbLong
        AppendVar fsControle1(15), "CentroNome", dbText, 40
        AppendVar fsControle1(16), "Data", dbDate
        AppendVar fsControle1(17), "ADebitar", dbCurrency
        AppendVar fsControle1(18), "ACreditar", dbCurrency
        AppendVar fsControle1(19), "ASaldo", dbCurrency
        AppendVar fsControle1(20), "Percentual", dbCurrency
        If CrieAux(rstTemp, fsControle1()) Then
            CriaTabelaTemp = True
        End If
    End If
End Function

'FUNCTION..: AppendTemp
'Objetivo..: Adiciona os dados obtidos das tabelas de Lan�amentos e Duplicatas
'            na tabela tempor�ria criada para imprimir o relat�rio.
'Argumentos: [rstTemp]: Recordset da tabela auxiliar.
'            [rstSrc] : Recordset com os Grupos e Contas.
'            [lBco]   : Matriz com os bancos escolhidos pelo usu�rio.
'            [dDatas] : Matriz com as datas escolhidas pelo usu�rio.
'Retorna...: True se terminar, False se o usu�rio cancelar
Private Function AppendTemp(rstTemp As Object, rstSrc As Object, lBco() As Long, dDatas() As Date) As Boolean
    Dim curEntrada     As Currency
    Dim curSaida       As Currency
    Dim curAEntrar     As Currency
    Dim curASair       As Currency
    Dim curPercentual  As Currency
    Dim lngConta       As Double
    Dim lngGrupo       As Long
    Dim strGrupo       As String
    Dim strWhere       As String
    Dim strWhere1      As String
    Dim lngContaAux    As Double
    Dim rstCentroCusto As Object
    Dim strCentroCusto As String
    Dim CentroCusto    As Long
    Dim fakedao        As New CGenericRecordset
    
    fakedao.Initialize rstTemp
    
    rstSrc.MoveFirst
    Do
        If lngGrupo <> GetValue(rstSrc, "Grupo") Then
            lngGrupo = GetValue(rstSrc, "Grupo")
            If UsandoModelo Then
                strGrupo = GetValue(rstSrc, "DescGrupo") 'C�digo de Descri��o do Grupo
            Else
                strGrupo = GetFieldValue("Descri��o", "Grupos", "[C�digo] = " & CStr(lngGrupo))  'Verifica descri��o do GRUPO
            End If
            SimpleMsgBar "Calculando Grupo: " & StrZero(lngGrupo, 9) & " - " & strGrupo
        End If
        If mbolCancelou Then Exit Function
        DoEvents                          'Permite ao usu�rio cancelar a gera��o
        If lngContaAux <> GetValue(rstSrc, "ContaAuxiliar", ZERO) Then
            curSaida = 0
            curEntrada = 0
            curASair = 0
            curAEntrar = 0
        End If
        lngConta = GetValue(rstSrc, "C�digo")
        strCentroCusto = NUL
        CentroCusto = ZERO
        If (CentrodeCusto(MFinanceiro) And (IsValid(txtCtrlFinanc(9).Text) Or IsValid(txtCtrlFinanc(10).Text))) Then
            If (IsValid(txtCtrlFinanc(9).Text) And IsValid(txtCtrlFinanc(10).Text)) Then
                If txtCtrlFinanc(9).Text = txtCtrlFinanc(10).Text Then
                    Concat strCentroCusto, " WHERE [C�digo] = ", txtCtrlFinanc(9).Text
                Else
                    Concat strCentroCusto, " WHERE ([C�digo] BETWEEN ", txtCtrlFinanc(9).Text, " AND ", txtCtrlFinanc(10).Text, ")"
                End If
            ElseIf (IsValid(txtCtrlFinanc(9).Text)) Then
                Concat strCentroCusto, " WHERE Centro >= ", txtCtrlFinanc(9).Text
            ElseIf (IsValid(txtCtrlFinanc(10).Text)) Then
                Concat strCentroCusto, " WHERE [C�digo] <= ", txtCtrlFinanc(10).Text
            End If
        End If
        If CentrodeCusto(MFinanceiro) And (tabCtrlFinanc.SelectedItem.Key = "or�ado") And chkDiscCentroCusto.value = vbChecked Then
            AbreRecordset rstCentroCusto, "Select * from [Centros] " & strCentroCusto, dbOpenSnapshot
        End If
        Do
            If CentrodeCusto(MFinanceiro) And (tabCtrlFinanc.SelectedItem.Key = "or�ado") And chkDiscCentroCusto.value = vbChecked Then
                curSaida = 0
                curEntrada = 0
                curASair = 0
                curAEntrar = 0
                If rstCentroCusto.EOF Then Exit Do
            End If
            'Resolve a instru��o de Transfer�ncias com o Banco de Destino
            strWhere = AddTransfBancarias(lngConta, dDatas(), lBco(), False, CentroCusto)
            If (Len(strWhere)) Then
                If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
                    curEntrada = curEntrada + Soma("Valor", "[Transf Banc�ria]", strWhere)
                Else
                    'Protocolo 74572: Transfer�ncia sempre em reais
                    curEntrada = curEntrada + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = [Transf Banc�ria].Data)", "[Transf Banc�ria]", strWhere)
                End If
            End If
            If mbolCancelou Then Exit Function
            DoEvents
            'Resolve a instru��o de Transfer�ncias com o Banco de Origem
            strWhere = AddTransfBancarias(lngConta, dDatas(), lBco(), True)
            If (Len(strWhere)) Then
                'Protocolo 74572:
                If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
                    curSaida = curSaida - Soma("Valor", "[Transf Banc�ria]", strWhere)
                Else
                    curSaida = curSaida - Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = [Transf Banc�ria].Data)", "[Transf Banc�ria]", strWhere)
                End If
            End If
            If mbolCancelou Then Exit Function
            DoEvents
            'Resolve a instru��o de Aplica��es para opera��es de cr�dito
            strWhere = AddAplicacoes(lngConta, lBco(), dDatas(), True, CentroCusto)
            If (Len(strWhere)) Then
                If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
                    curEntrada = curEntrada + Soma("Valor", "Aplica��es", strWhere)
                Else
                    'Protocolo 74572: Aplica��es sempre em reais
                    curEntrada = curEntrada + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Aplica��es.Data)", "Aplica��es", strWhere)
                End If
            End If
            If mbolCancelou Then Exit Function
            DoEvents
            'Resolve a instru��o de Aplica��es para opera��es de D�bito
            strWhere = AddAplicacoes(lngConta, lBco(), dDatas(), False, CentroCusto)
            If (Len(strWhere)) Then
                If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
                    curSaida = curSaida - Soma("Valor", "Aplica��es", strWhere)
                Else
                    'Protocolo 74572:
                    curSaida = curSaida - Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Aplica��es.Data)", "Aplica��es", strWhere)
                End If
            End If
            If mbolCancelou Then Exit Function
            DoEvents
            'Resolve a instru��o para Duplicatas Recebidas ou A Receber
            If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Duplicatas")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Receber")) Then
                strWhere = AddLancDupl(lngConta, lBco(), dDatas(), False, CentroCusto, 0) '0 = N�o Realizado
                strWhere1 = AddLancDupl(lngConta, lBco(), dDatas(), False, CentroCusto, 1) '1 = Realizado
                If (Len(strWhere)) And (Len(strWhere1)) Then
                    'Protocolo 74572:
                    curAEntrar = curAEntrar + SomarMoedas("Duplicatas", strWhere, txtCtrlFinanc(8).Text)
                    curEntrada = curEntrada + SomarMoedas("Duplicatas", strWhere1, txtCtrlFinanc(8).Text)
                End If
                If mbolCancelou Then Exit Function
                DoEvents
            End If
            'Resolve a instru��o para Duplicatas Pagas ou A Pagar
            If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Duplicatas")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Pagar")) Then
                strWhere = AddLancDupl(lngConta, lBco(), dDatas(), True, CentroCusto, 0) '0 = N�o Realizado
                strWhere1 = AddLancDupl(lngConta, lBco(), dDatas(), True, CentroCusto, 1) '1 = Realizado
                If (Len(strWhere)) And (Len(strWhere1)) Then
                    'Protocolo 74572:
                    curASair = curASair - SomarMoedas("Duplicatas", strWhere, txtCtrlFinanc(8).Text)
                    curSaida = curSaida - SomarMoedas("Duplicatas", strWhere1, txtCtrlFinanc(8).Text)
                End If
                If mbolCancelou Then Exit Function
                DoEvents
            End If
            'Resolve a instru��o para Lan�amentos Recebidos ou A Receber
            If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Lan�amentos")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Receber")) Then
                strWhere = AddLancDupl(lngConta, lBco(), dDatas(), False, CentroCusto, 0) 'A Realizar
                strWhere1 = AddLancDupl(lngConta, lBco(), dDatas(), False, CentroCusto, 1) 'Realizado
                If (Len(strWhere)) And (Len(strWhere1)) Then
                    'Protocolo 74572:
                    curAEntrar = curAEntrar + SomarMoedas("Lan�amentos", strWhere, txtCtrlFinanc(8).Text)
                    curEntrada = curEntrada + SomarMoedas("Lan�amentos", strWhere1, txtCtrlFinanc(8).Text)
                End If
                If mbolCancelou Then Exit Function
                DoEvents
            End If
            'Resolve a instru��o para Lan�amentos Pagos ou A Pagar
            If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Lan�amentos")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Pagar")) Then
                strWhere = AddLancDupl(lngConta, lBco(), dDatas(), True, CentroCusto, 0) 'A Realizar
                strWhere1 = AddLancDupl(lngConta, lBco(), dDatas(), True, CentroCusto, 1) 'Realizado
                If (Len(strWhere)) And (Len(strWhere1)) Then
                    'Protocolo 74572:
                    curASair = curASair - SomarMoedas("Lan�amentos", strWhere, txtCtrlFinanc(8).Text)
                    curSaida = curSaida - SomarMoedas("Lan�amentos", strWhere1, txtCtrlFinanc(8).Text)
                End If
                If mbolCancelou Then Exit Function
                DoEvents
            End If
            'Grava os dados na tabela tempor�ria
            If lngConta <> 0 Then
                If tabCtrlFinanc.SelectedItem.Key = "or�ado" Or ((curEntrada <> 0) Or (curSaida <> 0)) Or ((curAEntrar <> 0) Or (curASair <> 0)) Then
                    'GRAVAR DADOS DO OR�ADO MESMO QUE N�O HAJA LAN�AMENTOS OU DUPLICATAS
                    If (Not (chkDiscCentroCusto.value = vbChecked) And (Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO) <> 0 Or _
                        curEntrada <> 0 Or curSaida <> 0 Or curAEntrar <> 0 Or curASair <> 0)) Or ((chkDiscCentroCusto.value = vbChecked) And _
                        (Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND Centro = " & CentroCusto & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO) <> 0 Or _
                        curEntrada <> 0 Or curSaida <> 0 Or curAEntrar <> 0 Or curASair <> 0)) Then
                        If UsandoModelo Then
                            lngContaAux = GetValue(rstSrc, "ContaAuxiliar", ZERO)
                        Else
                            lngContaAux = lngConta
                        End If
                        'Pesquisa o primeiro registro de grupo, conta e centro de custo informados
                        
                        'rstTemp.FindFirst "GrupoC�digo=" & lngGrupo & " AND ContaC�digo=" & lngContaAux & " AND CentroC�digo = " & CentroCusto
                        fakedao.FindFirst "[GrupoC�digo]=" & lngGrupo & " AND [ContaC�digo]=" & lngContaAux & " AND [CentroC�digo] = " & CentroCusto
                        
                        'Se centro de custo financeiro, se for relat�rio Or�ado x Realizado e discriminar CCusto = True
                        If fakedao.NoMatch Then
                            fakedao.AddNew
                        Else
                            fakedao.Edit
                        End If
                        rstTemp("GrupoC�digo").value = lngGrupo
                        rstTemp("GrupoNome").value = strGrupo
                        rstTemp("ContaC�digo").value = lngContaAux
                        rstTemp("ContaNome").value = rstSrc("Descri��o").value
                        rstTemp("Sa�da").value = curSaida
                        rstTemp("Entrada").value = curEntrada
                        rstTemp("ADebitar").value = curASair
                        rstTemp("ACreditar").value = curAEntrar
                        rstTemp("Saldo").value = (curEntrada + curSaida)   'Saida � sempre negativo por isso sinal +
                        
                        'pt. 88454 - Ivo Sousa (17/09/2008)
                        curPercentual = (curEntrada + curSaida)
                        If tabCtrlFinanc.SelectedItem.Key = "or�ado" Then
                            rstTemp("ASaldo").value = (curAEntrar + curASair) 'Saida � sempre negativo por isso sinal +
                            curPercentual = curPercentual + (curAEntrar + curASair)
                        End If
                        rstTemp("CentroC�digo").value = CentroCusto
                        rstTemp("CentroNome").value = GetFieldValue("Descri��o", "Centros", "[C�digo] = " & CentroCusto, , NUL)
                        If UsandoModelo Then
                            If CentrodeCusto(MFinanceiro) And (tabCtrlFinanc.SelectedItem.Key = "or�ado") And chkDiscCentroCusto.value = vbChecked Then
                                rstTemp("Or�ado").value = Soma("Valor", "[Or�amentos de Contas]", "Conta in (Select [Conta Cont�bil] from [Contas de Contas Auxiliares] where Conta = " & lngContaAux & ") AND Centro = " & CentroCusto & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO)
                            Else
                                rstTemp("Or�ado").value = Soma("Valor", "[Or�amentos de Contas]", "Conta in (Select [Conta Cont�bil] from [Contas de Contas Auxiliares] where Conta = " & lngContaAux & ") AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO)
                            End If
                        Else
                            If CentrodeCusto(MFinanceiro) And (tabCtrlFinanc.SelectedItem.Key = "or�ado") And chkDiscCentroCusto.value = vbChecked Then
                                rstTemp("Or�ado").value = Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND Centro = " & CentroCusto & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO)
                            Else
                                'Protocolo 73669
                                'Checar se os Centros de Custo Inicial e Final foram informados
                                If Len(txtCtrlFinanc(9).Text) > 0 And Len(txtCtrlFinanc(10).Text) > 0 Then
                                    rstTemp("Or�ado").value = Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND Centro BETWEEN " & txtCtrlFinanc(9).Text & " AND " & txtCtrlFinanc(10).Text & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO)
                                Else
                                    'Checar se foi informado apenas o Centro de Custo Inicial
                                    If Len(txtCtrlFinanc(9).Text) > 0 Then
                                        rstTemp("Or�ado").value = Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND Centro >= " & txtCtrlFinanc(9).Text & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO)
                                    ElseIf Len(txtCtrlFinanc(10).Text) > 0 Then 'Sen�o foi informado apenas o Centro de Custo Final
                                        rstTemp("Or�ado").value = Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND Centro <= " & txtCtrlFinanc(10).Text & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO)
                                    Else
                                        rstTemp("Or�ado").value = Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO)
                                    End If
                                End If
                            End If
                        End If
                        'pt. 88454 - Ivo Sousa (17/09/2008)
                        If rstTemp("Or�ado").value <> 0 Then
                            curPercentual = ((curPercentual - rstTemp("Or�ado").value) / rstTemp("Or�ado").value) * 100
                        Else
                            curPercentual = 0
                        End If
                        rstTemp.Fields("Percentual").value = curPercentual
                        rstTemp.update
                        rstTemp.MoveFirst
                    End If
                    If Not UsandoModelo Then
                        curSaida = 0
                        curEntrada = 0
                        
                        curASair = 0
                        curAEntrar = 0
                    End If
                End If 'Fim do If tabCtrlFinanc.SelectedItem.Key = "or�ado"
            End If
            If CentrodeCusto(MFinanceiro) And (tabCtrlFinanc.SelectedItem.Key = "or�ado") And chkDiscCentroCusto.value = vbChecked Then
                If CentroCusto > 0 Then rstCentroCusto.MoveNext
            End If
            If CentrodeCusto(MFinanceiro) And (tabCtrlFinanc.SelectedItem.Key = "or�ado") And chkDiscCentroCusto.value = vbChecked Then
                CentroCusto = GetValue(rstCentroCusto, "C�digo", ZERO)
            End If
        Loop Until (Not CentrodeCusto(MFinanceiro)) Or Not (tabCtrlFinanc.SelectedItem.Key = "or�ado") Or Not chkDiscCentroCusto.value = vbChecked
        'Move a tabela origem para o pr�ximo registro
        rstSrc.MoveNext
    Loop Until rstSrc.EOF
    AppendTemp = True
    'Set rstTemp = Nothing
    Set fakedao = Nothing
End Function

'SUB.......: RelatorioOrcado
'Objetivo..: Monta e imprime o relat�rio de Controle Financeiro Sint�tico.
'Argumentos: [pdPrint]  : Destino da impress�o.
'            [rstSource]: Origem dos dados.
'            [strTitulo]: Subt�tulo do relat�rio.
Private Sub RelatorioOrcado(pdPrint As PrintDestinoEnum, rstSource As Object, strTitulo As String, strTitulo2 As String)
    Dim wrkSintetico  As KeybReport
    Dim rstGrupos     As Object
    Dim i             As Integer
    Dim Tamanho       As Single
    Dim strNomeTabela As String
    
    'Somente ser o recordset contiver algum registro
    If EstaVazio(rstSource) Then
        MsgBox LoadResString(146), vbInformation, MsgBoxCaption
        Exit Sub
    End If
  
    Set wrkSintetico = New KeybReport
    With wrkSintetico
        Set .DatabaseName = GlobalDataBase
        Set .Recordset = rstSource
        .Destino = pdPrint
        .Sentido = wrPSPaisagem
        .AutoRedraw = True
        .Tipo = wrObjectDraw
        .ScaleMode = vbMillimeters
        .WindowTitulo = "Controle Financeiro Or�ado"
        PageHeader wrkSintetico, "Controle Financeiro Or�ado"
        'Insere linha no Cabe�alho para Informar a Moeda
        If Len(txtCtrlFinanc(8).Text) > 0 Then
            .UltimaSecao.AddLinha "Moeda"
            .UltimaSecao(.UltimaSecao.LinhasCount).AddCampo , wrCSFixedText, "Valores Demonstrados em: " & txtCtrlFinanc(8).Text, wrTACentro
        End If
        'Acrescenta uma linha no cabe�alho para colocar o subt�tulo
        .Grupo("Cabe�alho").Header.AddLinha "sub"
        .Grupo("Cabe�alho").Header("sub").AddCampo , wrCSFixedText, strTitulo, wrTACentro
        .Grupo("Cabe�alho").Header.AddLinha "sub2"
        .Grupo("Cabe�alho").Header("sub2").AddCampo , wrCSFixedText, strTitulo2, wrTACentro
        .FontSize = 7
        .FontStyle = wrFSBold
        'Criando a estrutura do relat�rio
        .AddGrupo "1"
        .Grupo(1).AddSecao scHeader, 3, wrDBBottomBorder
        .Grupo(1).Quebra = "GrupoC�digo"      'Quebra do grupo por c�digo do grupo
        With .Grupo(1).Header.Linha(2)
            .AddCampo , wrCSFixedText, "Grupo:", , 15
            .AddCampo , , "GrupoC�digo", wrTADireito, 17
            .Campo(2).Formato = StrZero(0, 9)
            .AddCampo , , "GrupoNome"
        End With
        'pt. 84357 Abner Luidi Hempkemaier (30/11/2007)
        With .Grupo(1).Header.Linha(3)
            .AddCampo , wrCSFixedText, "Contas", , 19, 10
            If (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
                If CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
                    .AddCampo , wrCSFixedText, "Centro", wrTADireito, 10, 0
                End If
                .AddCampo , wrCSFixedText, "JAN", wrTADireito, 19, 35
                .AddCampo , wrCSFixedText, "FEV", wrTADireito, 19, 54
                .AddCampo , wrCSFixedText, "MAR", wrTADireito, 19, 73
                .AddCampo , wrCSFixedText, "ABR", wrTADireito, 19, 92
                .AddCampo , wrCSFixedText, "MAI", wrTADireito, 19, 111
                .AddCampo , wrCSFixedText, "JUN", wrTADireito, 19, 130
                .AddCampo , wrCSFixedText, "JUL", wrTADireito, 19, 149
                .AddCampo , wrCSFixedText, "AGO", wrTADireito, 19, 168
                .AddCampo , wrCSFixedText, "SET", wrTADireito, 19, 187
                .AddCampo , wrCSFixedText, "OUT", wrTADireito, 19, 206
                .AddCampo , wrCSFixedText, "NOV", wrTADireito, 19, 225
                .AddCampo , wrCSFixedText, "DEZ", wrTADireito, 19, 244
                .AddCampo , wrCSFixedText, "Total", wrTADireito, 19, 263
            End If
        End With
        'Criando a se��o de apresenta��o dos dados
        .FontStyle = wrFSNormal
        Tamanho = 35
        .Grupo(1).AddSecao scDetalhe, 2
        With .Grupo(1).Detalhe.Linha(1)
            .AddCampo , , "ContaC�digo", wrTADireito, 10
            .AddCampo , , "ContaNome", , 29
            For i = 1 To 12
                .AddCampo , , "Saldo" & i, wrTADireito, 19, Tamanho
                .Campo(i + 2).Formato = FMOEDA
                Tamanho = Tamanho + 19
            Next
            .AddCampo , , "TotalSaldo", wrTADireito, 40, Tamanho + 1
            .Campo(15).Formato = FMOEDA
        End With
        Tamanho = 35
        With .Grupo(1).Detalhe.Linha(2)
            For i = 1 To 12
                .AddCampo , , "Or�ado" & i, wrTADireito, 19, Tamanho
                .Campo(i).Formato = FMOEDA
                Tamanho = Tamanho + 19
            Next
            .AddCampo , , "TotalOrcado", wrTADireito, 40, Tamanho + 1
            .Campo(13).Formato = FMOEDA
        End With
        .Grupo(1).AddSecao scFooter, 1, wrDBBottomBorder
        With .Grupo(1).Footer.Linha(1)
            .DrawBorder = wrDBTopBorder
            .BorderStyle = wrDot
            .AddCampo , wrCSFixedText, "Total do Grupo:", , 35
            .Campo(1).FontStyle = wrFSBold
            Tamanho = 35
            For i = 2 To 13
                .AddCampo , wrCSDataLink, "SUM(Saldo" & (i - 1) & " - Or�ado" & (i - 1) & ")", wrTADireito, 19, Tamanho
'                If gTipoDB = Access Then
                    .Campo(i).TableLink = NomeAuxiliar
'                Else
'                    .Campo(i).TableLink = rstSource(0).Properties("BASETABLENAME")
'                End If
                .Campo(i).DataLink = "GrupoC�digo = {*GrupoC�digo}"
                .Campo(i).Formato = FMOEDA
                Tamanho = Tamanho + 19
            Next
            .AddCampo , wrCSDataLink, "SUM(TotalSaldo - TotalOrcado)", wrTADireito, 40, Tamanho + 1
'            If gTipoDB = Access Then
                .Campo(14).TableLink = NomeAuxiliar
'            Else
'                .Campo(14).TableLink = rstSource(0).Properties("BASETABLENAME")
'            End If
            .Campo(14).DataLink = "GrupoC�digo = {*GrupoC�digo}"
            .Campo(14).Formato = FMOEDA
        End With
'        If gTipoDB = Access Then
            strNomeTabela = NomeAuxiliar
'        Else
'            strNomeTabela = rstSource(0).Properties("BASETABLENAME")
'        End If
        .AddGrupo "2", wrDBTopBorder Or wrDBBottomBorder
        .Grupo(2).AddSecao scHeader, 4
        .Grupo(2).Header(2).DrawBorder = wrDBBottomBorder
        .Grupo(2).Header(2).BorderStyle = wrDot
        With .Grupo(2).Header.Linha(2)
            .DrawBorder = wrDBTopBorder
            .BorderStyle = wrDot
            .AddCampo , wrCSFixedText, "Realizado", , 50
            .Campo(1).FontStyle = wrFSBold
            Tamanho = 35
            For i = 2 To 13
                .AddCampo , wrCSDataLink, "SUM(Saldo" & (i - 1) & ")", wrTADireito, 19, Tamanho
                .Campo(i).TableLink = NomeAuxiliar
                .Campo(i).Formato = FMOEDA
                .Campo(i).FontStyle = wrFSBold
                Tamanho = Tamanho + 19
            Next
            .AddCampo , wrCSDataLink, "SUM(TotalSaldo)", wrTADireito, 40, Tamanho + 1
            .Campo(14).TableLink = NomeAuxiliar
            .Campo(14).FontStyle = wrFSBold
            .Campo(14).Formato = FMOEDA
        End With
        With .Grupo(2).Header.Linha(3)
            .DrawBorder = wrDBTopBorder
            .BorderStyle = wrDot
            .AddCampo , wrCSFixedText, "Or�ado", , 50
            .Campo(1).FontStyle = wrFSBold
            Tamanho = 35
            For i = 2 To 13
                .AddCampo , wrCSDataLink, "SUM(Or�ado" & (i - 1) & ")", wrTADireito, 19, Tamanho
                .Campo(i).TableLink = NomeAuxiliar
                .Campo(i).Formato = FMOEDA
                .Campo(i).FontStyle = wrFSBold
                Tamanho = Tamanho + 19
            Next
            .AddCampo , wrCSDataLink, "SUM(TotalOrcado)", wrTADireito, 40, Tamanho + 1
            .Campo(14).TableLink = NomeAuxiliar
            .Campo(14).FontStyle = wrFSBold
            .Campo(14).Formato = FMOEDA
        End With
        With .Grupo(2).Header.Linha(4)
            .DrawBorder = wrDBTopBorder
            .BorderStyle = wrDot
            .AddCampo , wrCSFixedText, "Acumulado", , 50
            .Campo(1).FontStyle = wrFSBold
            Tamanho = 35
            For i = 2 To 13
                .AddCampo , wrCSDataLink, "SUM(Or�ado" & (i - 1) & " - Saldo" & (i - 1) & ")", wrTADireito, 19, Tamanho
                .Campo(i).TableLink = NomeAuxiliar
                .Campo(i).Formato = FMOEDA
                .Campo(i).FontStyle = wrFSBold
                Tamanho = Tamanho + 19
            Next
            .AddCampo , wrCSDataLink, "SUM(TotalOrcado) - SUM(TotalSaldo)", wrTADireito, 40, Tamanho + 1
            .Campo(14).TableLink = NomeAuxiliar
            .Campo(14).FontStyle = wrFSBold
            .Campo(14).Formato = FMOEDA
        End With
    End With
    SetPtr vbDefault
    wrkSintetico.BeginPrint gTipoDB
    wrkSintetico.EndPrint
    Set wrkSintetico = Nothing
End Sub

'SUB.......: RelatorioSintetico
'Objetivo..: Monta e imprime o relat�rio de Controle Financeiro Sint�tico.
'Argumentos: [pdPrint]  : Destino da impress�o.
'            [rstSource]: Origem dos dados.
'            [strTitulo]: Subt�tulo do relat�rio.
Private Sub RelatorioSintetico(pdPrint As PrintDestinoEnum, rstSource As Object, strTitulo As String, strTitulo2 As String)
Dim wrkSintetico As KeybReport
Dim rstGrupos    As Object

  ' Somente ser o recordset contiver algum registro
  If EstaVazio(rstSource) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
    Exit Sub
  End If
  
  Set wrkSintetico = New KeybReport
  With wrkSintetico
    Set .DatabaseName = GlobalDataBase
    Set .Recordset = rstSource
    .Destino = pdPrint
    .AutoRedraw = True
    .Tipo = wrObjectDraw
    .ScaleMode = vbMillimeters
    .WindowTitulo = "Controle Financeiro Sint�tico"
    
    Const GRP_HEADER$ = "Cabe�alho"  'Nome para o grupo criado
    Dim sngWidth As Single

    sngWidth = .ClientWidth     'Largura da �rea imprim�vel da p�gina
    .AddGrupo GRP_HEADER, , wrVPNoTopo, True
    .Grupo(GRP_HEADER).AddSecao scHeader, 3, wrDBAllBorders
    .FontName = "Arial"
    .FontSize = 12
    .FontStyle = wrFSBold
    With .Grupo(GRP_HEADER).Header
      .Linha(1).AddCampo , wrCSFixedText, NomeDonaSistema, wrTACentro, sngWidth
       wrkSintetico.FontSize = 11

      .Linha(2).AddCampo , wrCSFixedText, UserName, wrTAEsquerdo, sngWidth, 0.2
      .Linha(2).AddCampo , wrCSData, , wrTADireito, sngWidth, 0.2

      wrkSintetico.FontSize = 10
      wrkSintetico.FontStyle = wrFSNormal

      .Linha(3).AddCampo , wrCSPagina, , wrTAEsquerdo, sngWidth, 0.2
      .Linha(3).AddCampo , wrCSFixedText, "Controle Financeiro Sint�tico", wrTACentro, sngWidth, 0.2
      .Linha(3).AddCampo , wrCSHora, , wrTADireito, sngWidth, 0.2
    End With
    
    'Insere linha no Cabe�alho para Informar a Moeda
    If Len(txtCtrlFinanc(8).Text) > 0 Then
      .UltimaSecao.AddLinha "Moeda"
      .UltimaSecao(.UltimaSecao.LinhasCount).AddCampo , wrCSFixedText, "Valores Demonstrados em: " & txtCtrlFinanc(8).Text, wrTACentro
    End If
    ' Acrescenta uma linha no cabe�alho para colocar o subt�tulo
    .Grupo("Cabe�alho").Header.AddLinha "sub"
    .Grupo("Cabe�alho").Header("sub").AddCampo , wrCSFixedText, strTitulo, wrTACentro
    .Grupo("Cabe�alho").Header.AddLinha "sub2"
    .Grupo("Cabe�alho").Header("sub2").AddCampo , wrCSFixedText, strTitulo2, wrTACentro
    .FontSize = 8
    .FontStyle = wrFSBold
    ' Criando a estrutura do relat�rio
    .AddGrupo "1"
    .Grupo(1).AddSecao scHeader, 3        'Cria se��o do GRUPO de Contas
    .Grupo(1).Quebra = "GrupoC�digo"      'Quebra do grupo por c�digo do grupo
    
    With .Grupo(1).Header.Linha(2)
      .AddCampo , wrCSFixedText, "Grupo:", , 10
      .AddCampo , , "GrupoC�digo", wrTADireito, 15
      .Campo(2).Formato = StrZero(0, 9)
      .AddCampo , , "GrupoNome"
    End With
    
    'Se for relat�rio Or�ado x Realizado
    If (tabCtrlFinanc.SelectedItem.Key = "or�ado") And CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
      .Grupo(1).AddSubGrupo "1"
      .Grupo(1).Subgrupo(1).AddSecao scHeader, 2    'Cria se��o do Subgrupo CONTAS
      .Grupo(1).Subgrupo(1).Quebra = "ContaC�digo"  'Quebra pelo c�digo da Conta
      
      With .Grupo(1).Subgrupo(1).Header.Linha(1)
        .AddCampo , wrCSFixedText, "Conta:", , 15, 5
        .AddCampo , , "ContaC�digo", wrTADireito, 10
        .Campo(2).Formato = StrZero(0, 9)
        .AddCampo , , "ContaNome"
      End With
      'Cabe�alho:  Contas (ou Centro)  Or�ado    Realizado   A Realizar    Varia��o
      With .Grupo(1).Subgrupo(1).Header.Linha(2)
        If CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
          .AddCampo , wrCSFixedText, "Centro", wrTAEsquerdo, 30, 10
        Else
          .AddCampo , wrCSFixedText, "Contas", wrTAEsquerdo, 30, 10
        End If
        .AddCampo , wrCSFixedText, "Or�ado", wrTADireito, 22, 95
        .AddCampo , wrCSFixedText, "Realizado", wrTADireito, 22
        .AddCampo , wrCSFixedText, "A Realizar", wrTADireito, 22
        .AddCampo , wrCSFixedText, "Varia��o", wrTADireito, 22
        'pt. 88454 - Ivo Sousa (17/09/2008)
        .AddCampo , wrCSFixedText, "%", wrTADireito, 22
      End With
    Else  'Se N�O for Ccusto = Financeiro e Discriminar CCusto = False
      With .Grupo(1).Header.Linha(3)
        'Se for relat�rio Or�ado x Realizado
        If (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
          .AddCampo , wrCSFixedText, "Contas", , 46, 17
          If CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
            .AddCampo , wrCSFixedText, "Centro", wrTADireito, 30, 54
          End If
          .AddCampo , wrCSFixedText, "Or�ado", wrTADireito, 22, 95
          .AddCampo , wrCSFixedText, "Realizado", wrTADireito, 22
          .AddCampo , wrCSFixedText, "A Realizar", wrTADireito, 22
          .AddCampo , wrCSFixedText, "Varia��o", wrTADireito, 22
          'pt. 88454 - Ivo Sousa (17/09/2008)
          .AddCampo , wrCSFixedText, "%", wrTADireito, 22
        Else 'Se for relat�rio Sint�tico
          .AddCampo , wrCSFixedText, "Contas", , 46, 17
          .AddCampo , wrCSFixedText, "Entradas", wrTADireito, 25, 65
          .AddCampo , wrCSFixedText, "A Entrar", wrTADireito, 25
          .AddCampo , wrCSFixedText, "Sa�das", wrTADireito, 25
          .AddCampo , wrCSFixedText, "A Sair", wrTADireito, 25
          .AddCampo , wrCSFixedText, "Saldo Realizado", wrTADireito, 25
        End If
      End With
    End If
    .FontStyle = wrFSNormal
    'Se for relat�rio Or�ado x Realizado, CCusto = Financeiro e Discriminar Centro de Custo = True
    If (tabCtrlFinanc.SelectedItem.Key = "or�ado") And CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
      'DETALHES    Or�ado x Realizado
      .Grupo(1).Subgrupo(1).AddSecao scDetalhe, 1
      With .Grupo(1).Subgrupo(1).Detalhe.Linha(1)
        .AddCampo , , "CentroC�digo", wrTAEsquerdo, 10, 10
        .AddCampo , , "CentroNome", , 50
        'Or�ado
        .AddCampo , , "Or�ado", wrTADireito, 22, 95
        .Campo(3).Formato = FMOEDA
        'Realizado
        .AddCampo , , "Saldo", wrTADireito, 22    'Saldo = Entrada - Sa�da
        .Campo(4).Formato = FMOEDA
        'A Realizar
        .AddCampo , , "ASaldo", wrTADireito, 22   'ASaldo = ACreditar - ADebitar
        .Campo(5).Formato = FMOEDA
        'Varia��o
        .AddCampo , wrCSDataLink, "Saldo + ASaldo - Or�ado", wrTADireito, 22
'        If gTipoDB = Access Then
          .Campo(6).TableLink = NomeTabeladoRST(rstSource)
'        Else
'          .Campo(6).TableLink = rstSource(0).Properties("BASETABLENAME")
'        End If
        .Campo(6).DataLink = "ContaC�digo = {*ContaC�digo} AND CentroC�digo = {CentroC�digo}"
        .Campo(6).Formato = FMOEDA
        
        'pt. 88454 - Ivo Sousa (17/09/2008)
        .AddCampo , , "Percentual", wrTADireito, 22
        .Campo(7).Formato = FMOEDA

      End With
    Else  'Se CCusto N�O for Financeiro e Discriminar Centro de Custo = False
      .Grupo(1).AddSecao scDetalhe, 1
      'D  E  T  A  L  H  E  S   para o Relat�rio Sint�tico
      With .Grupo(1).Detalhe.Linha(1)
        'C�digo e Nome da Conta
        .AddCampo , , "ContaC�digo", wrTADireito, 15, 1
        .AddCampo , , "ContaNome", , 50
        
        'D  E  T  A  L  H  E  S    Or�ado x Realizado quando:
        'CCusto n�o for Financeiro e Discriminar Centro de Custo = False
        If (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
          'Or�ado ---------------------------------------------------
          .AddCampo , , "Or�ado", wrTADireito, 22, 95
          .Campo(3).Formato = FMOEDA
          'Realizado -------------------------------------------------
          .AddCampo , , "Saldo", wrTADireito, 22   'Saldo = Entrada - Sa�da
          .Campo(4).Formato = FMOEDA
          'A Realizar -----------------------------------------------
          .AddCampo , , "ASaldo", wrTADireito, 22   'ASaldo = ACreditar - ADebitar
          .Campo(5).Formato = FMOEDA
          'Varia��o -------------------------------------------------
          .AddCampo , wrCSDataLink, "ABS(Saldo + ASaldo) - ABS(Or�ado)", wrTADireito, 22  'Parenteses mat�m o sinal do or�ado
'          If gTipoDB = Access Then
            .Campo(6).TableLink = NomeTabeladoRST(rstSource)
'          Else
'            .Campo(6).TableLink = rstSource(0).Properties("BASETABLENAME")
'          End If
          .Campo(6).DataLink = "ContaC�digo = {ContaC�digo}"
          .Campo(6).Formato = FMOEDA
          
          'Percentual ----------------------------------------------
          'pt. 88454 - Ivo Sousa (17/09/2008)
          .AddCampo , , "Percentual", wrTADireito, 22
          .Campo(7).Formato = FMOEDA
        Else
          'D  E  T  A  L  H  E  S    Relat�rio Sint�tico
          .AddCampo , , "Entrada", wrTADireito, 25, 65
          .Campo(3).Formato = FMOEDA
          .AddCampo , , "ACreditar", wrTADireito, 25
          .Campo(4).Formato = FMOEDA
          .AddCampo , , "Sa�da", wrTADireito, 25
          .Campo(5).Formato = FMOEDA
          .AddCampo , , "ADebitar", wrTADireito, 25
          .Campo(6).Formato = FMOEDA
          .AddCampo , , "Saldo", wrTADireito, 25
          .Campo(7).Formato = FMOEDA
        End If
      
      End With
    End If
    '
    ' Adiciona a se��o de Rodap� com os totais dos campos Sa�da, Entrada e Saldo
    '
    'Se for relat�rio Or�ado x Realizado, CCusto = Financeiro e Discriminar Centro de Custo = True
    If (tabCtrlFinanc.SelectedItem.Key = "or�ado") And CentrodeCusto(MFinanceiro) And chkDiscCentroCusto.value = vbChecked Then
      'R o d a p �      d a       S e � � o ------------------------------
      .Grupo(1).Subgrupo(1).AddSecao scFooter, 1, wrDBBottomBorder
      With .Grupo(1).Subgrupo(1).Footer.Linha(1)
        .DrawBorder = wrDBTopBorder
        .BorderStyle = wrDot
        .AddCampo , wrCSFixedText, "Total da Conta:", , 22, 70
        .Campo(1).FontStyle = wrFSBold
        'Or�ado -----------------------------------------------
        .AddCampo , wrCSSubTotal, "Or�ado", wrTADireito, 22, 95
        .Campo(2).Formato = FMOEDA
        'Realizado ---------------------------------------------
        .AddCampo , wrCSSubTotal, "Saldo", wrTADireito, 22  'Saldo = Entrada - Sa�da
        .Campo(3).Formato = FMOEDA
        'A Realizar -------------------------------------------
        .AddCampo , wrCSSubTotal, "ASaldo", wrTADireito, 22  'ASaldo = ACreditar - ADebitar
        .Campo(4).Formato = FMOEDA
        'Varia��o ----------------------------------------------
        .AddCampo , wrCSDataLink, "ABS(SUM(Saldo) + SUM(ASaldo)) - ABS(SUM(Or�ado))", wrTADireito, 22 'Parenteses mat�m o sinal do or�ado
'        If gTipoDB = Access Then
          .Campo(5).TableLink = NomeTabeladoRST(rstSource)
'        Else
'          .Campo(5).TableLink = rstSource(0).Properties("BASETABLENAME")
'        End If
        .Campo(5).DataLink = "ContaC�digo = {*ContaC�digo}"
        .Campo(5).Formato = FMOEDA
        'Percentual ----------------------------------------------
        'pt. 88454 - Ivo Sousa (17/09/2008)
        .AddCampo , wrCSSubTotal, "Percentual", wrTADireito, 22
        .Campo(6).Formato = FMOEDA
      End With
    Else
      'Se N�O for CCusto = Financeiro e Discriminar Centro de Custo = False
      'R o d a p �      d a       S e � � o ------------------------------
      If (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
         .Grupo(1).AddSecao scFooter, 1, wrDBBottomBorder
      Else
         .Grupo(1).AddSecao scFooter, 2, wrDBBottomBorder
      End If
      With .Grupo(1).Footer.Linha(1)
        .DrawBorder = wrDBTopBorder
        .BorderStyle = wrDot
        'Quando relat�rio for Or�ado x Realizado
        If (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
          .AddCampo , wrCSFixedText, "Total do Grupo:", , 62, 10
          .Campo(1).FontStyle = wrFSBold
          'Or�ado ----------------------------------------------
          .AddCampo , wrCSSubTotal, "Or�ado", wrTADireito, 22, 95
          .Campo(2).Formato = FMOEDA
          'Realizado ---------------------------------------------
          .AddCampo , wrCSSubTotal, "Saldo", wrTADireito, 22  'Saldo = Entrada + (- Sa�da)
          .Campo(3).Formato = FMOEDA
          'A Realizar ------------------------------------------
          .AddCampo , wrCSSubTotal, "ASaldo", wrTADireito, 22  'ASaldo = ACreditar + (- ADebitar)
          .Campo(4).Formato = FMOEDA
          'Varia��o --------------------------------------------
          .AddCampo , wrCSDataLink, "ABS(SUM(Saldo) + SUM(ASaldo)) - ABS(SUM(Or�ado))", wrTADireito, 22 'Parenteses mat�m o sinal do or�ado
'          If gTipoDB = Access Then
            .Campo(5).TableLink = NomeTabeladoRST(rstSource)
'          Else
'            .Campo(5).TableLink = rstSource(0).Properties("BASETABLENAME")
'          End If
          .Campo(5).DataLink = "GrupoC�digo = {*GrupoC�digo}"
          .Campo(5).Formato = FMOEDA
          
          'Percentual ----------------------------------------------
          'pt. 88454 - Ivo Sousa (17/09/2008)
          .AddCampo , wrCSSubTotal, "Percentual", wrTADireito, 22
          .Campo(6).Formato = FMOEDA

        Else  'Quando for relat�rio Sint�tico
          .AddCampo , wrCSFixedText, "", , 82, 10
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSSubTotal, "Entrada", wrTADireito, 25, 65
          .Campo(2).Formato = FMOEDA
          .AddCampo , wrCSSubTotal, "ACreditar", wrTADireito, 25
          .Campo(3).Formato = FMOEDA
          .AddCampo , wrCSSubTotal, "Sa�da", wrTADireito, 25
          .Campo(4).Formato = FMOEDA
          .AddCampo , wrCSSubTotal, "ADebitar", wrTADireito, 25
          .Campo(5).Formato = FMOEDA
          .AddCampo , wrCSSubTotal, "Saldo", wrTADireito, 25
          .Campo(6).Formato = FMOEDA
        End If
      End With
      If (tabCtrlFinanc.SelectedItem.Key <> "or�ado") Then
        'Protocolo 74461 - Acrescentado outro Total considerando tamb�m ADebitar e ACreditar
        With .Grupo(1).Footer.Linha(2)
            .DrawBorder = wrDBTopBorder
            .BorderStyle = wrDot
            .AddCampo , wrCSFixedText, "Total do Grupo:", , 82, 10
            .Campo(1).FontStyle = wrFSBold
            'Cr�ditos
            .AddCampo , wrCSFixedText, "Cr�ditos:", wrTADireito, 25, 65
            .Campo(2).FontStyle = wrFSBold
            '--------
            .AddCampo , wrCSDataLink, "SUM(Entrada) + SUM(ACreditar)", wrTADireito, 25
'            If gTipoDB = Access Then
              .Campo(3).TableLink = NomeTabeladoRST(rstSource)
'            Else
'              .Campo(3).TableLink = rstSource(0).Properties("BASETABLENAME")
'            End If
            .Campo(3).DataLink = "GrupoC�digo = {*GrupoC�digo}"
            .Campo(3).Formato = FMOEDA
            'D�bitos
            .AddCampo , wrCSFixedText, "D�bitos:", wrTADireito, 25
            .Campo(4).FontStyle = wrFSBold
            '--------
            .AddCampo , wrCSDataLink, "SUM(Sa�da) + SUM(ADebitar)", wrTADireito, 25
'            If gTipoDB = Access Then
              .Campo(5).TableLink = NomeTabeladoRST(rstSource)
'            Else
'              .Campo(5).TableLink = rstSource(0).Properties("BASETABLENAME")
'            End If
            .Campo(5).DataLink = "[GrupoC�digo] = {*GrupoC�digo}"
            .Campo(5).Formato = FMOEDA
        End With
      End If
    End If
    
    'Se relat�rio Or�ado x Realizado
    If (tabCtrlFinanc.SelectedItem.Key = "or�ado") Then
      Dim strNomeTabela         As String
      Dim curOrcado      As Currency
      Dim curRealizado   As Currency
      Dim curARealizar As Currency
      Dim a As Currency
      Dim b As Currency
      Dim c As Currency
      Dim curVariacao    As Currency       'Total de Entradas
      
'      If gTipoDB = Access Then
        strNomeTabela = NomeTabeladoRST(rstSource)
'      Else
'        strNomeTabela = rstSource(0).Properties("BASETABLENAME")
'      End If
      
      .AddGrupo "2", wrDBTopBorder Or wrDBBottomBorder
      .Grupo(2).AddSecao scHeader, 5
      .Grupo(2).Header(2).DrawBorder = wrDBBottomBorder
      .Grupo(2).Header(2).BorderStyle = wrDot
      
      With .Grupo(2).Header.Linha(2)
        .DrawBorder = wrDBTopBorder
        .BorderStyle = wrDot
        .AddCampo , wrCSFixedText, "TOTAL GERAL DE RECEITAS:", , 50, 30
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSDataLink, "SUM(Or�ado)", wrTADireito, 22, 95
        .Campo(2).TableLink = strNomeTabela
        .Campo(2).DataLink = "Or�ado > 0"
        .Campo(2).Formato = FMOEDA
        .Campo(2).FontStyle = wrFSBold
        'pt. 80949 - Moacir Pfau(24/04/2008)
        .AddCampo , wrCSDataLink, "SUM(Saldo)", wrTADireito, 22
        .Campo(3).TableLink = strNomeTabela
        .Campo(3).DataLink = "Saldo > 0"
        .Campo(3).Formato = FMOEDA
        .Campo(3).FontStyle = wrFSBold
        .AddCampo , wrCSDataLink, "SUM(ACreditar)", wrTADireito, 22
        .Campo(4).TableLink = strNomeTabela
        .Campo(4).DataLink = "ACreditar > 0"
        .Campo(4).Formato = FMOEDA
        .Campo(4).FontStyle = wrFSBold
        'pt. 80949 - Moacir Pfau(24/04/2008)
        .AddCampo , wrCSDataLink, "(SUM(Saldo) + SUM(ACreditar)) - (SUM(Or�ado))", wrTADireito, 22 'Parenteses mat�m o sinal do or�ado
        .Campo(5).TableLink = strNomeTabela
        .Campo(5).DataLink = "Saldo > 0 OR ACreditar > 0 OR Or�ado > 0"
        .Campo(5).FontStyle = wrFSBold
        .Campo(5).Formato = FMOEDA
      End With
        
      With .Grupo(2).Header.Linha(3)
        .DrawBorder = wrDBTopBorder
        .BorderStyle = wrDot
        .AddCampo , wrCSFixedText, "TOTAL GERAL DE DESPESAS:", , 50, 30
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSDataLink, "SUM(Or�ado)", wrTADireito, 22, 95
        .Campo(2).TableLink = strNomeTabela
        .Campo(2).DataLink = "Or�ado < 0"
        .Campo(2).Formato = FMOEDA
        .Campo(2).FontStyle = wrFSBold
        'pt. 80949 - Moacir Pfau(24/04/2008)
        .AddCampo , wrCSDataLink, "SUM(Saldo)", wrTADireito, 22
        .Campo(3).TableLink = strNomeTabela
        .Campo(3).DataLink = "Saldo < 0"
        .Campo(3).Formato = FMOEDA
        .Campo(3).FontStyle = wrFSBold
        .AddCampo , wrCSDataLink, "SUM(ADebitar)", wrTADireito, 22
        .Campo(4).TableLink = strNomeTabela
        .Campo(4).DataLink = "ADebitar < 0"
        .Campo(4).Formato = FMOEDA
        .Campo(4).FontStyle = wrFSBold
        'pt. 80949 - Moacir Pfau(24/04/2008)
        .AddCampo , wrCSDataLink, "(SUM(Saldo) + SUM(ADebitar)) - (SUM(Or�ado))", wrTADireito, 22 'Parenteses mat�m o sinal do or�ado
        .Campo(5).TableLink = strNomeTabela
        .Campo(5).DataLink = "Saldo < 0 OR ADebitar < 0 OR Or�ado < 0"
        .Campo(5).FontStyle = wrFSBold
        .Campo(5).Formato = FMOEDA
      End With
      
      curOrcado = Soma("Or�ado", strNomeTabela, "Or�ado > 0") - Abs(Soma("Or�ado", strNomeTabela, "Or�ado < 0"))
      curRealizado = Soma("Entrada", strNomeTabela, "Entrada > 0") - Abs(Soma("Sa�da", strNomeTabela, "Sa�da < 0"))
      curARealizar = Soma("ACreditar", strNomeTabela, "ACreditar > 0") - Abs(Soma("ADebitar", strNomeTabela, "ADebitar < 0"))
      
                  
      With .Grupo(2).Header.Linha(4)
        .DrawBorder = wrDBTopBorder
        .BorderStyle = wrDot
        .AddCampo , wrCSFixedText, "SALDO(RECEITAS - DESPESAS):", , 50, 30
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, Format$(curOrcado, FMOEDA), wrTADireito, 22, 95
        .Campo(2).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, Format$(curRealizado, FMOEDA), wrTADireito, 22
        .Campo(3).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, Format$(curARealizar, FMOEDA), wrTADireito, 22
        .Campo(4).FontStyle = wrFSBold

      End With
        
        
    ElseIf (tabCtrlFinanc.SelectedItem.Key = "sintetico") Then  'S I N T � T I C O
      '
      ' �ltimo Grupo do relat�rio: Totais dos Grupos de Contas
      ' Crio uma consulta que me retorna apenas o c�digo dos grupos existentes
      ' na tabela tempor�ria
      '
      strTitulo = "SELECT DISTINCT [GrupoC�digo], GrupoNome FROM " & NomeTabeladoRST(rstSource) & ";"
      
      If (AbreRecordset(rstGrupos, strTitulo, dbOpenSnapshot) = WL_OK) Then
        Dim curTotalCreditar  As Currency
        Dim curTotalACreditar As Currency
        Dim curTotalDebitar   As Currency
        Dim curTotalADebitar  As Currency
        Dim curDebito         As Currency       'Total de Sa�das
        Dim curASair       As Currency
        Dim curCredito        As Currency       'Total de Entradas
        Dim curAEntrar        As Currency
        Dim curSaldo          As Currency
        Dim curTotalSaldo     As Currency       'Total do total
        Dim intContador       As Integer
        
        .FontStyle = wrFSBold
        .AddGrupo "2", wrDBTopBorder Or wrDBBottomBorder, wrVPNoFinal
        
        .Grupo(2).AddSecao scHeader, 1
        With .Grupo(2).Header.Linha(1)
          .AddCampo , wrCSFixedText, "Saldo Inicial do Banco"
          .AddCampo , wrCSFixedText, Format(SaldoBanco(CDateDef(txtCtrlFinanc(0).Text)), FMOEDA), wrTADireito
        End With
              
        .Grupo(2).Header.AddLinha
        .Grupo(2).Header(2).DrawBorder = wrDBBottomBorder
        .Grupo(2).Header(2).BorderStyle = wrDot
        With .Grupo(2).Header.Linha(2)
          .AddCampo , wrCSFixedText, "Saldo Final do Banco"
          'Foi adicionado um dia por motivo da fun��o de saldo retornar o saldo inicial do dia
          ' assim n�o considerando os movimentos do dia atual
          .AddCampo , wrCSFixedText, Format(SaldoBanco(CDateDef(txtCtrlFinanc(1).Text) + 1), FMOEDA), wrTADireito
        End With
        
        .Grupo(2).Header.AddLinha
        .Grupo(2).Header(3).DrawBorder = wrDBBottomBorder
        .Grupo(2).Header(3).BorderStyle = wrDot
        With .Grupo(2).Header.Linha(3)
          .AddCampo "Teste", wrCSFixedText, "TOTAIS", wrTACentro
          .Campo(1).FontStyle = wrFSBold
          .Campo(1).FontSize = 9
        End With
        
        .Grupo(2).Header.AddLinha
        With .Grupo(2).Header.Linha(4)
          .AddCampo , wrCSFixedText, "Grupo de Contas", , 67, 1
          .AddCampo , wrCSFixedText, "Cr�dito", wrTADireito, 25, 65
          .AddCampo , wrCSFixedText, "A Creditar", wrTADireito, 25
          .AddCampo , wrCSFixedText, "D�bito", wrTADireito, 25
          .AddCampo , wrCSFixedText, "A Debitar", wrTADireito, 25
          .AddCampo , wrCSFixedText, "Saldo Realizado", wrTADireito, 25
        End With
        .FontStyle = wrFSNormal
        '
        ' Adicionando quantas linhas forem necess�rias para imprimir os valores
        ' R O D A P �
        '
        rstGrupos.MoveFirst
        intContador = 5
        Do Until rstGrupos.EOF
          .Grupo(2).Header.AddLinha
          With .Grupo(2).Header.Linha(intContador)
            .AddCampo , wrCSFixedText, CStr(GetValue(rstGrupos, 0)), wrTADireito, 17, 1
            .AddCampo , wrCSFixedText, GetValue(rstGrupos, 1), , 46
            
            If TypeOf rstSource Is dao.Recordset Then
              'C R � D I T O  -----------------------------------
              curTotalCreditar = Soma("Entrada", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
              'A   C R E D I T A R ------------------------------
              curTotalACreditar = Soma("ACreditar", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
              'D � B I T O  --------------------------------------
              curTotalDebitar = Soma("Sa�da", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value))
              'A   D E B I T A R  --------------------------------------
              curTotalADebitar = Soma("ADebitar", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value))
              'S A L D O  Entradas - Sa�das ----------------------------
              curTotalSaldo = Soma("Saldo", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
            Else
              'C R � D I T O  -----------------------------------
              curTotalCreditar = Soma("Entrada", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
              'A   C R E D I T A R ------------------------------
              curTotalACreditar = Soma("ACreditar", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
              'D � B I T O  --------------------------------------
              curTotalDebitar = Soma("Sa�da", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value))
              'A   D E B I T A R  --------------------------------------
              curTotalADebitar = Soma("ADebitar", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value))
              'S A L D O  Entradas - Sa�das -----------------------------
              curTotalSaldo = Soma("Saldo", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
            End If
            curCredito = curCredito + curTotalCreditar
            .AddCampo , wrCSFixedText, Format$(curTotalCreditar, FMOEDA), wrTADireito, 25, 65
            curAEntrar = curAEntrar + curTotalACreditar
            .AddCampo , wrCSFixedText, Format$(curTotalACreditar, FMOEDA), wrTADireito, 25
            curDebito = curDebito + curTotalDebitar
            .AddCampo , wrCSFixedText, Format$(curTotalDebitar, FMOEDA), wrTADireito, 25
            curASair = curASair + curTotalADebitar
            .AddCampo , wrCSFixedText, Format$(curTotalADebitar, FMOEDA), wrTADireito, 25
            curSaldo = curSaldo + curTotalSaldo
            .AddCampo , wrCSFixedText, Format$(curTotalSaldo, FMOEDA), wrTADireito, 25
          End With
          Inc intContador
          rstGrupos.MoveNext
        Loop
        
        '�ltima linha Totais
        If (tabCtrlFinanc.SelectedItem.Key = "sintetico") Then
           .Grupo(2).AddSecao scFooter, 2, wrDBTopBorder
        Else
           .Grupo(2).AddSecao scFooter, 1, wrDBTopBorder
        End If
        .Grupo(2).Footer.BorderStyle = wrDot
        With .Grupo(2).Footer.Linha(1)
          .AddCampo , wrCSFixedText, "Totais:"
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, Format$(curCredito, FMOEDA), wrTADireito, 25, 65
          .AddCampo , wrCSFixedText, Format$(curAEntrar, FMOEDA), wrTADireito, 25
          .AddCampo , wrCSFixedText, Format$(curDebito, FMOEDA), wrTADireito, 25
          .AddCampo , wrCSFixedText, Format$(curASair, FMOEDA), wrTADireito, 25
          .AddCampo , wrCSFixedText, Format$(curSaldo, FMOEDA), wrTADireito, 25
        End With
        If (tabCtrlFinanc.SelectedItem.Key = "sintetico") Then
          'Protocolo 74461 - Acrescentado outro Total considerando tamb�m ADebitar e ACreditar
          With .Grupo(2).Footer.Linha(2)
            .DrawBorder = wrDBTopBorder
            .BorderStyle = wrDot
            .AddCampo , wrCSFixedText, "Totais Gerais:"
            .Campo(1).FontStyle = wrFSBold
            .AddCampo , wrCSFixedText, "Cr�ditos:", wrTADireito, 25, 65
            .AddCampo , wrCSFixedText, Format$(curCredito + curAEntrar, FMOEDA), wrTADireito, 25
            .AddCampo , wrCSFixedText, "D�bitos:", wrTADireito, 25
            .AddCampo , wrCSFixedText, Format$(curDebito + curASair, FMOEDA), wrTADireito, 25
            .AddCampo , wrCSFixedText, Format$(curCredito + curAEntrar + curDebito + curASair, FMOEDA), wrTADireito, 25
          End With
        End If
      End If
      FechaRecordset rstGrupos
    End If
  End With
    
  SetPtr vbDefault
  wrkSintetico.BeginPrint gTipoDB
  wrkSintetico.EndPrint
  
  Set wrkSintetico = Nothing
End Sub

'FUNCTION..: AppendTempAnual
'Objetivo..: Grava a tabela tempor�ria para o relat�rio de Controle Financeiro
'            Anual.
'Argumentos: [rstTemp]     : Recordset da tabela tempor�ria
'            [rstContas]   : Recordset com os grupos e contas.
'            [Anterior]    : Saldo anterior
'            [lBancos]     : Matriz com os bancos inicial e final.
'            [dPeriodo]    : Matriz com as Datas inicial e final.
'Retorna...: False se o usu�rio cancelar, caso contr�rio True.
Private Function AppendTempAnual(rstTemp As Object, rstContas As Object, Anterior As Currency, lBancos() As Long, dPeriodo() As Date) As Boolean
Dim strCompare  As String
Dim strAnterior As String
Dim strMes      As String          'M�s atual do c�lculo
Dim strAno      As String          'Ano atual do c�lculo
Dim lngConta    As Double            'C�digo da conta
Dim strGrupo    As String          'Descri��o do Grupo
Dim dCalcMes    As Date            'M�s de c�lculo
Dim curDebito   As Currency        'Valor de Sa�da
Dim curCredito  As Currency        'Valor de Entrada
Dim curAnterior As Currency        'Valor Anterior
Dim lngContaA   As Double            'Conta anterior
Dim lngContaAux As Double
Dim genTemp As New CGenericRecordset   'Para permitir o FindFirst
Dim sSeparadorData As String
    
If gTipoDB = Access Then
    sSeparadorData = "#"
Else
    sSeparadorData = "'"
End If
    
  rstContas.MoveFirst
  Do Until rstContas.EOF
    curDebito = 0
    curCredito = 0
  
    '
    ' Resolvendo Conta e Grupo atual
    '
    lngConta = GetValue(rstContas, "C�digo")
    If UsandoModelo Then
      strGrupo = GetValue(rstContas, "DescGrupo")
    Else
      strGrupo = GetFieldValue("Descri��o", "Grupos", "[C�digo] = " & CStr(GetValue(rstContas, "Grupo")))
    End If
    
    dCalcMes = dPeriodo(IDX_INICIO)
    
    genTemp.Initialize rstTemp
    Do Until (DateDiff("m", dCalcMes, dPeriodo(IDX_FINAL)) < 0)
      curDebito = 0
      curCredito = 0

      
      If mbolCancelou Then Exit Function
      DoEvents
      
      strMes = CStr(Month(dCalcMes))
      strAno = CStr(Year(dCalcMes))
      
      SimpleMsgBar "Calculando M�s " & DataToMesAnoS(dCalcMes) & " da conta " & _
                   StrZero(lngConta, 9) & ESP & rstContas("Descri��o").value
                   
'      ********************** F�bio disse pra comentar *****************************
'      '
'      ' Selecionando os dados de Transfer�ncias Banc�rias com Banco de Origem
'      '
'      strCompare = AddTransfBancarias(lngConta, dPeriodo(), lBancos(), True)
'      If (Len(strCompare)) Then
'        Concat strCompare, " AND (Month(Data) = ", strMes, " AND Year(Data) = ", strAno, ")"
'        curDebito = Soma("Valor", "[Transf Banc�ria]", strCompare)
'      End If
'      If mbolCancelou Then Exit Function
'      DoEvents      'Possibilita ao usu�rio cancelar a gera��o

      '
      ' Seleciona os dados de Transfer�ncias Banc�rias com Banco de Destino
      '
      strCompare = AddTransfBancarias(lngConta, dPeriodo(), lBancos(), False)

      If (Len(strCompare)) Then
        
        ' Saldo Anterior
        If chkSaldoAnterior = vbChecked Then
          If lngContaA <> lngConta Then
            strAnterior = Left(strCompare, InStr(1, strCompare, "(Data") - 5) & Mid$(strCompare, InStr(1, strCompare, sSeparadorData & ")") + 3)
            Concat strAnterior, " AND Data <= " & InverteData(dPeriodo(IDX_INICIO), True)
            If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
              curAnterior = curAnterior + Soma("Valor", "[Transf Banc�ria]", strAnterior)
            Else
              curAnterior = curAnterior + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "[Transf Banc�ria]", strAnterior)
            End If
          End If
        End If
        '
        If Month(dPeriodo(IDX_INICIO)) = strMes And Year(dPeriodo(IDX_INICIO)) = strAno Then
          Concat strCompare, " AND Day(Data) >= ", Day(dPeriodo(IDX_INICIO))
        End If
        Concat strCompare, " AND (Month(Data) = ", strMes, " AND Year(Data) = ", strAno, ")"
        If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
          curCredito = Soma("Valor", "[Transf Banc�ria]", strCompare)
        Else
          curCredito = Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "[Transf Banc�ria]", strCompare)
        End If
      End If
      If mbolCancelou Then Exit Function
      DoEvents
      '
      ' Seleciona os dados de Aplica��es com o tipo Juros/Corre��o
      '
      strCompare = AddAplicacoes(lngConta, lBancos(), dPeriodo(), True)
      If (Len(strCompare)) Then
        ' Saldo Anterior
        If chkSaldoAnterior = vbChecked Then
          If lngContaA <> lngConta Then
            strAnterior = strCompare 'Left(strCompare, InStr(1, strCompare, "(Data") - 5 & Mid$(strCompare, InStr(1, strCompare, "#)") + 3))
            Concat strAnterior, " AND Data <= " & InverteData(dPeriodo(IDX_INICIO), True)
            If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
              curAnterior = curAnterior + Soma("Valor", "Aplica��es", strAnterior)
            Else
              curAnterior = curAnterior + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "Aplica��es", strAnterior)
            End If
          End If
        End If
        '
        If Month(dPeriodo(IDX_INICIO)) = strMes And Year(dPeriodo(IDX_INICIO)) = strAno Then
          Concat strCompare, " AND Day(Data) >= ", Day(dPeriodo(IDX_INICIO))
        End If
        Concat strCompare, " AND (Month(Data) = ", strMes, " AND Year(Data) = ", strAno, ")"
        
        If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
          curCredito = curCredito + Soma("Valor", "Aplica��es", strCompare)
        Else
          curCredito = curCredito + Soma("Valor /(SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "Aplica��es", strCompare)
        End If
      End If
      If mbolCancelou Then Exit Function
      DoEvents
      '
      ' Seleciona os dados de aplica��es com os tipo Taxa ou CPMF
      '
      strCompare = AddAplicacoes(lngConta, lBancos(), dPeriodo(), False)
      If (Len(strCompare)) Then
        ' Saldo Anterior
        If chkSaldoAnterior = vbChecked Then
          If lngContaA <> lngConta Then
            strAnterior = strCompare 'Left(strCompare, InStr(1, strCompare, "(Data") - 5) & Mid$(strCompare, InStr(1, strCompare, "#)") + 3)
            Concat strAnterior, " AND Data <= " & InverteData(dPeriodo(IDX_INICIO), True)
            If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
              curAnterior = curAnterior - Soma("Valor", "Aplica��es", strAnterior)
            Else
              curAnterior = curAnterior - Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "Aplica��es", strAnterior)
            End If
          End If
        End If
        '
        If Month(dPeriodo(IDX_INICIO)) = strMes And Year(dPeriodo(IDX_INICIO)) = strAno Then
          Concat strCompare, " AND Day(Data) >= ", Day(dPeriodo(IDX_INICIO))
        End If
        Concat strCompare, " AND (Month(Data) = ", strMes, " AND Year(Data) = ", strAno, ")"
        If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
          curDebito = curDebito + Soma("Valor", "Aplica��es", strCompare)
        Else
          curDebito = curDebito + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "Aplica��es", strCompare)
        End If
      End If
      If mbolCancelou Then Exit Function
      DoEvents
      
      strCompare = AddLancDupl(lngConta, lBancos(), dPeriodo(), True)
      
      '
      ' Seleciona os dados de Duplicatas a Pagar ou Pagas
      '
      If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Duplicatas")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Pagar")) Then
        If (Len(strCompare)) Then
          ' Saldo Anterior
          If chkSaldoAnterior = vbChecked Then
            If lngContaA <> lngConta Then
              strAnterior = Left(strCompare, InStr(1, strCompare, "(" & strData) - 5) & Mid$(strCompare, InStr(1, strCompare, sSeparadorData & ")") + 3)
              Concat strAnterior, " AND " & strData & " <= " & InverteData(dPeriodo(IDX_INICIO), True)
              
              'Protocolo 74572
              curAnterior = curAnterior - SomarMoedas("Duplicatas", strAnterior, txtCtrlFinanc(8).Text)
            End If
          End If
          '
          If Month(dPeriodo(IDX_INICIO)) = strMes And Year(dPeriodo(IDX_INICIO)) = strAno Then
            Concat strCompare, " AND Day(" & strData & ") >= ", Day(dPeriodo(IDX_INICIO))
          End If
          Concat strCompare, " AND (Month(" & strData & ") = ", strMes, " AND Year(" & strData & ") = ", strAno, ")"
          
          'Protocolo 74572
          curDebito = curDebito + SomarMoedas("Duplicatas", strCompare, txtCtrlFinanc(8).Text)
        End If
        If mbolCancelou Then Exit Function
        DoEvents
      End If
      
      strCompare = AddLancDupl(lngConta, lBancos(), dPeriodo(), True)
      
      '
      ' Seleciona os dados de Lan�amentos a Pagar ou Pagos
      '
      If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Lan�amentos")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Pagar")) Then
        
        If (Len(strCompare)) Then
          ' Saldo Anterior
          If chkSaldoAnterior = vbChecked Then
            If lngContaA <> lngConta Then
              strAnterior = Left(strCompare, InStr(1, strCompare, "(" & strData) - 5) & Mid$(strCompare, InStr(1, strCompare, sSeparadorData & ")") + 3)
              Concat strAnterior, " AND " & strData & " <= " & InverteData(dPeriodo(IDX_INICIO), True)
              'Protocolo 74572
              curAnterior = curAnterior - SomarMoedas("Lan�amentos", strAnterior, txtCtrlFinanc(8).Text)
            End If
          End If
          '
          If Month(dPeriodo(IDX_INICIO)) = strMes And Year(dPeriodo(IDX_INICIO)) = strAno Then
            Concat strCompare, " AND Day(" & strData & ") >= ", Day(dPeriodo(IDX_INICIO))
          End If
          Concat strCompare, " AND (Month(" & strData & ") = ", strMes, " AND Year(" & strData & ") = ", strAno, ")"
          'Protocolo 74572
          curAnterior = curAnterior - SomarMoedas("Lan�amentos", strCompare, txtCtrlFinanc(8).Text)
          
          'Projeto: #PT 125019 - Hist�ria: # - Desenvolvimento# -  Vinicius Alexandre Elyseu (22/11/2013)
          curDebito = curDebito + SomarMoedas("Lan�amentos", strCompare, txtCtrlFinanc(8).Text)
        End If
        If mbolCancelou Then Exit Function
        DoEvents
      End If
      
      strCompare = AddLancDupl(lngConta, lBancos(), dPeriodo(), False)
      
      '
      ' Seleciona os dados de Duplicatas a Receber ou Recebidas
      '
      If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Duplicatas")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Receber")) Then
        If (Len(strCompare)) Then
          ' Saldo Anterior
          If chkSaldoAnterior = vbChecked Then
            If lngContaA <> lngConta Then
              strAnterior = Left(strCompare, InStr(1, strCompare, "(" & strData) - 5) & Mid$(strCompare, InStr(1, strCompare, sSeparadorData & ")") + 3)
              Concat strAnterior, " AND " & strData & " <= " & InverteData(dPeriodo(IDX_INICIO), True)
             'Protocolo 74572
              curAnterior = curAnterior + SomarMoedas("Duplicatas", strAnterior, txtCtrlFinanc(8).Text)
            End If
          End If
          If Month(dPeriodo(IDX_INICIO)) = strMes And Year(dPeriodo(IDX_INICIO)) = strAno Then
            Concat strCompare, " AND Day(" & strData & ") >= ", Day(dPeriodo(IDX_INICIO))
          End If
          Concat strCompare, " AND (Month(" & strData & ") = ", strMes, " AND Year(" & strData & ") = ", strAno, ")"
          
          curCredito = curCredito + SomarMoedas("Duplicatas", strCompare, txtCtrlFinanc(8).Text)
      End If
        If mbolCancelou Then Exit Function
        DoEvents
      End If
      
      strCompare = AddLancDupl(lngConta, lBancos(), dPeriodo(), False)
      
      ' Seleciona os dados de Lan�amentos a Receber ou Recebidos
      If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Lan�amentos")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Receber")) Then
        If (Len(strCompare)) Then
          ' Saldo Anterior
          If chkSaldoAnterior = vbChecked Then
            If lngContaA <> lngConta Then
              strAnterior = Left(strCompare, InStr(1, strCompare, "(" & strData) - 5) & Mid$(strCompare, InStr(1, strCompare, sSeparadorData & ")") + 3)
              Concat strAnterior, " AND " & strData & " <= " & InverteData(dPeriodo(IDX_INICIO), True)
              
              'Protocolo 74572
              curAnterior = curAnterior + SomarMoedas("Lan�amentos", strAnterior, txtCtrlFinanc(8).Text)
              lngContaA = lngConta
            End If
          End If
          '
          If Month(dPeriodo(IDX_INICIO)) = strMes And Year(dPeriodo(IDX_INICIO)) = strAno Then
            Concat strCompare, " AND Day(" & strData & ") >= ", Day(dPeriodo(IDX_INICIO))
          End If
          Concat strCompare, " AND (Month(" & strData & ") = ", strMes, " AND Year(" & strData & ") = ", strAno, ")"
          'Protocolo 74572
          curCredito = curCredito + SomarMoedas("Lan�amentos", strCompare, txtCtrlFinanc(8).Text)
          'curCredito = curCredito + Soma("([Valor Original] + Acr�scimo - Abatimento) / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = " & strData & ")", _
                                           "Lan�amentos", strCompare)
        End If
        If mbolCancelou Then Exit Function
        DoEvents
      End If
      '
      ' Grava a tabela tempor�ria
      '
      If lngConta <> 0 And (curDebito <> 0 Or curCredito <> 0) Then
        If UsandoModelo Then
          lngContaAux = GetValue(rstContas, "ContaAuxiliar", ZERO)
        Else
          lngContaAux = lngConta
        End If
        
        genTemp.FindFirst "[GrupoC�digo]=" & rstContas("Grupo").value & " AND [ContaC�digo]=" & lngContaAux & " AND MesAno = " & InverteData(dCalcMes, True)
          
        If genTemp.NoMatch Then
          rstTemp.AddNew
        Else
          genTemp.Edit
          curDebito = curDebito + GetValue(rstTemp, "Sa�da", ZERO)
          curCredito = curCredito + GetValue(rstTemp, "Entrada", ZERO)
        End If
        
        rstTemp("GrupoC�digo").value = rstContas("Grupo").value
        rstTemp("GrupoNome").value = strGrupo
        rstTemp("ContaC�digo").value = lngContaAux
        rstTemp("ContaNome").value = rstContas("Descri��o").value
        rstTemp("Sa�da").value = curDebito
        rstTemp("Entrada").value = curCredito
        rstTemp("Saldo").value = (curCredito - curDebito)
        rstTemp("MesAno").value = dCalcMes
        rstTemp.update
        
        rstTemp.MoveFirst
      End If
      '
      ' Avan�a para o pr�ximo m�s no per�odo
      '
      dCalcMes = DateAdd("m", 1, dCalcMes)
      '
      ' Atualiza lngContaA
      '
      If lngContaA <> lngConta Then lngContaA = lngConta

    Loop
    rstContas.MoveNext      'Move para a pr�xima conta
  Loop
  
  Anterior = curAnterior
  
  AppendTempAnual = True
  
End Function

'SUB.......: RelatorioAnual
'Objetivo..: Gera o relat�rio de Controle Financeiro Anual.
'Argumentos: [pdImpressao]: Destino da impress�o
'            [rstOrigem]  : Recordset com a origem dos registros
'            [dtPeriodo]  : Matriz com as datas inicial e final.
Private Sub RelatorioAnual(pdImpressao As PrintDestinoEnum, rstOrigem As Object, dtPeriodo() As Date, curAnterior As Currency)
Dim wrkAnual        As KeybReport
Dim strSubTitulo    As String
  '
  ' Somente se o recordset tiver algum registro
  '
  If EstaVazio(rstOrigem) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
    Exit Sub
  End If
  
    strSubTitulo = CF_CONTAS & "em Aberto"  '  CF_CONTASQUITADAS

  ' Colocando a data
  Concat strSubTitulo, " de ", dtPeriodo(IDX_INICIO)
  Concat strSubTitulo, " at� ", dtPeriodo(IDX_FINAL)
  
  Set wrkAnual = New KeybReport
  With wrkAnual
    Set .DatabaseName = GlobalDataBase
    Set .Recordset = rstOrigem
    .Destino = pdImpressao
    .Tipo = wrObjectDraw
    .ScaleMode = vbMillimeters
    .WindowTitulo = "Controle Financeiro Anual"
    .AutoRedraw = True
    
    PageHeader wrkAnual, "Controle Financeiro Anual"
    
    'Insere linha no Cabe�alho para Informar a Moeda
    If Len(txtCtrlFinanc(8).Text) > 0 Then
      .UltimaSecao.AddLinha "Moeda"
      .UltimaSecao(.UltimaSecao.LinhasCount).AddCampo , wrCSFixedText, "Valores Demonstrados em: " & txtCtrlFinanc(8).Text, wrTACentro
    End If
    
        
    '
    ' Acrescenta uma linha no cabe�alho para colocar a data
    '
    .Grupo("Cabe�alho").Header.AddLinha "SubTitulo"
    With .Grupo("Cabe�alho").Header.Linha("SubTitulo")
      .AddCampo , wrCSFixedText, strSubTitulo, wrTACentro
    End With
    
    If chkSaldoAnterior.value = vbChecked Then
      .Grupo("Cabe�alho").Header.AddLinha "SaldoAnterior"
      With .Grupo("Cabe�alho").Header.Linha("SaldoAnterior")
        .AddCampo , wrCSFixedText, "Saldo Anterior:" & Format$(curAnterior, FCURRENCY), wrTACentro
      End With
    End If
    
    .FontSize = 8
    .FontStyle = wrFSBold
    '
    ' Grupo Principal, quebra por C�digo de Grupo
    '
    .AddGrupo "Principal", wrDBBottomBorder
    .Grupo("Principal").Quebra = "GrupoC�digo"
    .Grupo("Principal").AddSecao scHeader, 2
    With .Grupo("Principal").Header.Linha(2)
      .Height = (wrkAnual.TextHeight("W") * 2)
      .DrawBorder = wrDBAllBorders
      .AddCampo , wrCSFixedText, "Grupo:"
      .Campo(1).Top = ((.Height / 2) - (.Campo(1).Height / 2))
      .Campo(1).Width = 15
      .AddCampo , , "GrupoC�digo", wrTADireito
      .Campo(2).Top = ((.Height / 2) - (.Campo(1).Height / 2))
      .Campo(2).Width = 17
      .Campo(2).Formato = "000000000"
      .AddCampo , , "GrupoNome"
      .Campo(3).Top = ((.Height / 2) - (.Campo(1).Height / 2))
    End With
    '
    ' SubGrupo, quebra por C�digo da Conta
    '
    .FontStyle = wrFSBold Or wrFSItalic
    .Grupo("Principal").AddSubGrupo "SubGrupo", wrDBBottomBorder
    .Grupo("Principal").Subgrupo("SubGrupo").BorderStyle = wrDot
    .Grupo("Principal").Subgrupo("SubGrupo").AddSecao scHeader, 3
    .Grupo("Principal").Subgrupo("SubGrupo").Quebra = "ContaC�digo"
    With .Grupo("Principal").Subgrupo("SubGrupo").Header.Linha(2)
      .DrawBorder = wrDBBottomBorder
      .BorderStyle = wrDot
      .AddCampo , wrCSFixedText, "Conta:", , 15, 15
      .AddCampo , , "ContaC�digo", wrTADireito, 17
      .Campo(2).Formato = StrZero(0, 9)
      .AddCampo , , "ContaNome"
    End With
    
    .FontStyle = wrFSBold
    With .Grupo("Principal").Subgrupo("SubGrupo").Header.Linha(3)
      .AddCampo , wrCSFixedText, "Per�odo", , 12
      .AddCampo , wrCSFixedText, "Entrada", wrTADireito, 25
      .AddCampo , wrCSFixedText, "Sa�da", wrTADireito, 25
      .AddCampo , wrCSFixedText, "Saldo", wrTADireito, 25
      .AddCampo , wrCSFixedText, "Per�odo", wrTACentro, 12, (wrkAnual.ClientWidth / 2) - 5
      .AddCampo , wrCSFixedText, "Entrada", wrTADireito, 25
      .AddCampo , wrCSFixedText, "Sa�da", wrTADireito, 25
      .AddCampo , wrCSFixedText, "Saldo", wrTADireito, 25
    End With
    '
    ' Se��o de detalhes, impressa em duas colunas
    '
    .FontStyle = wrFSNormal
    .Grupo("Principal").Subgrupo("SubGrupo").AddSecao scDetalhe, 1
    .Grupo("Principal").Subgrupo("SubGrupo").Detalhe.MultiCol = True
    .Grupo("Principal").Subgrupo("SubGrupo").Detalhe.Width = 91
    With .Grupo("Principal").Subgrupo("SubGrupo").Detalhe.Linha(1)
      .AddCampo , , "MesAno", wrTAEsquerdo, 12
      .Campo(1).Formato = FMESANO
      .AddCampo , , "Entrada", wrTADireito, 25
      .Campo(2).Formato = FMOEDA
      .Campo(2).SuprimirZeros = True
      .AddCampo , , "Sa�da", wrTADireito, 25
      .Campo(3).Formato = FMOEDA
      .Campo(3).SuprimirZeros = True
      .AddCampo , , "Saldo", wrTADireito, 25
      .Campo(4).Formato = FMOEDA
      .Campo(4).SuprimirZeros = True
    End With
    '
    ' Rodap� do grupo principal
    '
    .Grupo("Principal").AddSecao scFooter, 1
    With .Grupo("Principal").Footer.Linha(1)
      .AddCampo , wrCSFixedText, "Total do Grupo:", wrTAEsquerdo, 25
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, "Entrada:", wrTADireito, 25
      .Campo(2).FontStyle = wrFSBold
      .AddCampo , wrCSSubTotal, "Entrada", wrTADireito, 25
      .Campo(3).Formato = FMOEDA
      .AddCampo , wrCSFixedText, "Sa�da:", wrTADireito, 25
      .Campo(4).FontStyle = wrFSBold
      .AddCampo , wrCSSubTotal, "Sa�da", wrTADireito, 25
      .Campo(5).Formato = FMOEDA
    End With
  
    ' �ltimo Grupo do relat�rio: Totais dos Grupos de Contas
    ' Crio uma consulta que me retorna apenas o c�digo dos grupos existentes
    ' na tabela tempor�ria
    Dim strTitulo As String
    Dim rstGrupos As Object
    strTitulo = "SELECT DISTINCT [GrupoC�digo], GrupoNome FROM " & NomeTabeladoRST(rstOrigem) & ";"
    If (AbreRecordset(rstGrupos, strTitulo, dbOpenSnapshot) = WL_OK) Then
      Dim curTotal    As Currency
      Dim curDebito   As Currency       'Total de Sa�das
      Dim curCredito  As Currency       'Total de Entradas
      Dim curSaldo    As Currency       'Total do total
      Dim intContador As Integer
      
      .FontStyle = wrFSBold
      .AddGrupo "2", wrDBTopBorder Or wrDBBottomBorder, wrVPNoFinal
      .Grupo(2).AddSecao scHeader, 1
      .Grupo(2).Header(1).DrawBorder = wrDBBottomBorder
      .Grupo(2).Header(1).BorderStyle = wrDot
      With .Grupo(2).Header.Linha(1)
        .AddCampo "Teste", wrCSFixedText, "TOTAIS", wrTACentro
        .Campo(1).FontStyle = wrFSBold
        .Campo(1).FontSize = 9
      End With
      
      .Grupo(2).Header.AddLinha
      With .Grupo(2).Header.Linha(2)
        .AddCampo , wrCSFixedText, "Grupo de Contas", , 30
        .AddCampo , wrCSFixedText, "Cr�dito", wrTADireito, 30, 101
        .AddCampo , wrCSFixedText, "D�bito", wrTADireito, 30
        .AddCampo , wrCSFixedText, "Saldo", wrTADireito, 30
      End With
      .FontStyle = wrFSNormal
      '
      ' Adicionando quantas linhas forem necess�rias para imprimir os valores
      '
      rstGrupos.MoveFirst
      intContador = 3
      Do Until rstGrupos.EOF
        .Grupo(2).Header.AddLinha
        With .Grupo(2).Header.Linha(intContador)
          .AddCampo , wrCSFixedText, CStr(GetValue(rstGrupos, 0)), wrTADireito, 17
          .AddCampo , wrCSFixedText, GetValue(rstGrupos, 1), , 81
          
          curTotal = Soma("Entrada", NomeTabeladoRST(rstOrigem), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
          curCredito = curCredito + curTotal
          .AddCampo , wrCSFixedText, Format$(curTotal, FMOEDA), wrTADireito, 30, 101
          
          curTotal = Soma("Sa�da", NomeTabeladoRST(rstOrigem), "[GrupoC�digo] = " & CStr(rstGrupos(0).value))
          curDebito = curDebito + curTotal
          .AddCampo , wrCSFixedText, Format$(curTotal, FMOEDA), wrTADireito, 30
                    
          curTotal = Soma("Saldo", NomeTabeladoRST(rstOrigem), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
          curSaldo = curSaldo + curTotal
          .AddCampo , wrCSFixedText, Format$(curTotal, FMOEDA), wrTADireito, 30
        End With
        Inc intContador
        rstGrupos.MoveNext
      Loop
      '
      ' �ltima linha Totais
      '
      .Grupo(2).AddSecao scFooter, 1, wrDBTopBorder
      .Grupo(2).Footer.BorderStyle = wrDot
      With .Grupo(2).Footer.Linha(1)
        .AddCampo , wrCSFixedText, "Totais:"
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, Format$(curCredito, FMOEDA), wrTADireito, 30, 101
        .AddCampo , wrCSFixedText, Format$(curDebito, FMOEDA), wrTADireito, 30
        .AddCampo , wrCSFixedText, Format$(curSaldo, FMOEDA), wrTADireito, 30
      End With
    End If
    FechaRecordset rstGrupos
  End With

  SetPtr vbDefault
  wrkAnual.BeginPrint gTipoDB
  wrkAnual.EndPrint
  Set wrkAnual = Nothing
End Sub

'FUNCTION..: AppendTempAnalitico
'Objetivo..: Grava os dados para o relat�rio de Controle Financeiro Anal�tico.
'Argumentos: [rstAux]   : Recordset auxiliar.
'            [rstContas]: Recordset com as Contas e Grupos.
'            [lngBancos]: Matriz com os Bancos inicial e final.
'            [datDatas] : Matriz com as Datas inicial e final.
'Retorna...: True se gravar a tabela corretamente, False se o usu�rio cancelar.
Private Function AppendTempAnalitico(rstAux As Object, rstContas As Object, lngBancos() As Long, datDatas() As Date) As Boolean
Dim strLanctos As String          'Instru��o para sele��o dos campos das tabelas
Dim strWhere   As String          'Instru��o de filtro
Dim rstLanctos As Object       'Recordset que receber� os dados
Dim strGrupo   As String          'Descri��o do Grupo
Dim lngConta   As Double           'C�digo da conta.
Dim lngContaAux  As Double

  rstContas.MoveFirst
  Do
    lngConta = GetValue(rstContas, "C�digo")
    If UsandoModelo Then
      strGrupo = GetValue(rstContas, "DescGrupo")
    Else
      strGrupo = GetFieldValue("Descri��o", "Grupos", "[C�digo] = " & CStr(GetValue(rstContas, "Grupo")))
    End If

    strLanctos = "SELECT [C�digo], Valor, Descri��o, Data FROM [Transf Banc�ria]"


    '
    ' Selecionando os dados de Transfer�ncia Banc�ria com banco de Origem
    '
    strWhere = AddTransfBancarias(lngConta, datDatas(), lngBancos(), True)
    If (Len(strWhere)) Then
      strWhere = strLanctos & " WHERE " & strWhere
      If (AbreRecordset(rstLanctos, strWhere, dbOpenForwardOnly) = WL_OK) Then
        GravaAuxAnalitico rstContas, rstLanctos, rstAux, False, "Transfer�ncia"
      End If
      FechaRecordset rstLanctos
    End If
    If mbolCancelou Then Exit Function
    
    DoEvents
    '
    ' Selecionando os dados de Transfer�ncia com o Banco de Destino
    '
    strWhere = AddTransfBancarias(lngConta, datDatas(), lngBancos(), False)
    If (Len(strWhere)) Then
      strWhere = strLanctos & " WHERE " & strWhere
      If (WL_OK = AbreRecordset(rstLanctos, strWhere, dbOpenForwardOnly)) Then
        GravaAuxAnalitico rstContas, rstLanctos, rstAux, True, "Transfer�ncia"
      End If
      FechaRecordset rstLanctos
    End If
    If mbolCancelou Then Exit Function
    DoEvents
    '
    ' Separando os dados da tabela de aplica��es tipo Juros e Corre��o
    '
    strLanctos = "SELECT [C�digo], Valor, [Descri��o], Data FROM [Aplica��es]"
    strWhere = AddAplicacoes(lngConta, lngBancos(), datDatas(), True)
    If (Len(strWhere)) Then
      strWhere = strLanctos & " WHERE " & strWhere
      If (WL_OK = AbreRecordset(rstLanctos, strWhere, dbOpenForwardOnly)) Then
        GravaAuxAnalitico rstContas, rstLanctos, rstAux, True, "Aplica��o"
      End If
      FechaRecordset rstLanctos
    End If
    If mbolCancelou Then Exit Function
    DoEvents
    '
    ' Separando dados de apli��o do tipo Taxas ou CPMF
    '
    strWhere = AddAplicacoes(lngConta, lngBancos(), datDatas(), False)
    If (Len(strWhere)) Then
      strWhere = strLanctos & " WHERE " & strWhere
      If (WL_OK = AbreRecordset(rstLanctos, strWhere, dbOpenForwardOnly)) Then
        GravaAuxAnalitico rstContas, rstLanctos, rstAux, False, "Aplica��o"
      End If
      FechaRecordset rstLanctos
    End If
    If mbolCancelou Then Exit Function
    DoEvents
    strLanctos = "SELECT Nota, ([Valor Original] + Acr�scimo - Abatimento) AS Soma, " & _
                 "Descri��o, Moeda, Emiss�o, Pagamento, Parcela, " & strData & " As Data, Empresa FROM Duplicatas"
    '
    ' Selecionando os dados para Duplicatas a Pagar ou Pagas
    '
    If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Duplicatas")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Pagar")) Then
      strWhere = AddLancDupl(lngConta, lngBancos(), datDatas(), True)
      If (Len(strWhere)) Then
        strWhere = strLanctos & " WHERE " & strWhere
        If (WL_OK = AbreRecordset(rstLanctos, strWhere, dbOpenForwardOnly)) Then
          GravaAuxAnalitico rstContas, rstLanctos, rstAux, False, "Duplicata"
        End If
        FechaRecordset rstLanctos
      End If
      If mbolCancelou Then Exit Function
      DoEvents
    End If
    '
    ' Selecionando os dados de Duplicatas a Receber ou Recebidas
    '
    If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Duplicatas")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Receber")) Then
      strWhere = AddLancDupl(lngConta, lngBancos(), datDatas(), False)
      If (Len(strWhere)) Then
        strWhere = strLanctos & " WHERE " & strWhere
        If (WL_OK = AbreRecordset(rstLanctos, strWhere, dbOpenForwardOnly)) Then
          GravaAuxAnalitico rstContas, rstLanctos, rstAux, True, "Duplicata"
        End If
        FechaRecordset rstLanctos
      End If
      If mbolCancelou Then Exit Function
      DoEvents
    End If
    '
    ' Selecionando os dados de Lan�amentos a Pagar ou Pagos
    '
    strLanctos = "SELECT [C�digo], ([Valor Original] + [Acr�scimo] - Abatimento) AS Soma, " & _
                 "[Descri��o], Moeda, [Emiss�o], Pagamento, " & strData & " As Data, Empresa FROM Lan�amentos"
    If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Lan�amentos")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Pagar")) Then
      strWhere = AddLancDupl(lngConta, lngBancos(), datDatas(), True)
      If (Len(strWhere)) Then
        strWhere = strLanctos & " WHERE " & strWhere
        If (WL_OK = AbreRecordset(rstLanctos, strWhere, dbOpenForwardOnly)) Then
          GravaAuxAnalitico rstContas, rstLanctos, rstAux, False, "Lan�amento"
        End If
        FechaRecordset rstLanctos
      End If
      If mbolCancelou Then Exit Function
      DoEvents
    End If
    '
    ' Selecionando os dados de Lan�amentos a Receber ou Recebidos
    '
    If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Lan�amentos")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Receber")) Then
      strWhere = AddLancDupl(lngConta, lngBancos(), datDatas(), False)
      If (Len(strWhere)) Then
        strWhere = strLanctos & " WHERE " & strWhere
        If (WL_OK = AbreRecordset(rstLanctos, strWhere, dbOpenForwardOnly)) Then
          GravaAuxAnalitico rstContas, rstLanctos, rstAux, True, "Lan�amento"
        End If
        FechaRecordset rstLanctos
      End If
      If mbolCancelou Then Exit Function
      DoEvents
    End If
    
    rstContas.MoveNext
    
  Loop Until rstContas.EOF
  
  NomeAuxiliar = NomeTabeladoRST(rstAux)
  NomeAuxiliar = "SELECT * FROM " & NomeAuxiliar & " ORDER BY [GrupoC�digo], [ContaC�digo], Data, [C�digo]"
  FechaRecordset rstAux
  Sleep (2000)
  If AbreRecordset(rstAux, NomeAuxiliar) = WL_OK Then
  'If AbreRecordset(rstAux, NomeAuxiliar, dbOpenDynaset, , , adUseClient) = WL_OK Then
  
    Dim SaldoFinal       As Double
    Dim Grupo            As Long
    Dim conta            As Long
    
    SaldoFinal = ZERO
    rstAux.MoveFirst
    
    Do
      If TypeOf rstAux Is dao.Recordset Then rstAux.Edit
      
      If Grupo <> GetValue(rstAux, "GrupoC�digo", ZERO) Or conta <> GetValue(rstAux, "ContaC�digo", ZERO) Then
        SaldoFinal = ZERO
      End If
      'Protocolo 74461 Trocado o sinal da saida (-)  por isso agora Entrada + Saida
      SaldoFinal = SaldoFinal + GetValue(rstAux, "Entrada", ZERO) + GetValue(rstAux, "Sa�da", ZERO)
      rstAux("Saldo").value = SaldoFinal
      rstAux.update

      conta = GetValue(rstAux, "ContaC�digo", ZERO)
      Grupo = GetValue(rstAux, "GrupoC�digo", ZERO)
      
      rstAux.MoveNext
    Loop Until rstAux.EOF
  End If
  
  AppendTempAnalitico = True
End Function

'FUNCTION..: GravaAuxAnalitico
'Objetivo..: Grava a tabela auxiliar para o relat�rio de Controle Financeiro
'            Anal�tico.
'Argumentos: [rstContas]: Recordset com as Contas e Grupos
'            [rstDados] : Recordset com os dados de Lan�amentos, Duplicatas,
'                         Tranfer�ncias Banc�rias ou Aplica��es.
'            [rstTemp]  : Recordset da tabela tempor�ria.
'            [bCredito] : True para cr�dito, False para d�bito.
'            [strOrigem]: String de origem dos dados.
'Retorna...: True se gravar a tabela, False se o usu�rio cancelar.
Private Function GravaAuxAnalitico(rstContas As Object, rstDados As Object, rstTemp As Object, bCredito As Boolean, strOrigem As String) As Boolean
  Dim strGrupo   As String      'Descri��o do grupo
  Dim lngConta   As Double
  Dim lngContaAux   As Double

  strGrupo = GetFieldValue("Descri��o", "Grupos", "[C�digo] = " & CStr(GetValue(rstContas, "Grupo")))
  
  
  lngConta = GetValue(rstContas, "C�digo")
  If UsandoModelo Then
    lngContaAux = GetValue(rstContas, "ContaAuxiliar")
  Else
    lngContaAux = lngConta
  End If
  
  If UsandoModelo Then
    strGrupo = GetValue(rstContas, "DescGrupo", NUL)
  Else
    strGrupo = GetFieldValue("Descri��o", "Grupos", "[C�digo] = " & GetValue(rstContas, "Grupo"))
  End If
  SimpleMsgBar "Calculando grupo " & CStr(GetValue(rstContas, "Grupo")) & _
               ESP & strGrupo & ", conta " & CStr(GetValue(rstContas, "C�digo")) & _
               ESP & GetValue(rstContas, "Descri��o")
  
  Do
    If mbolCancelou Then Exit Function
    DoEvents
    
    rstTemp.AddNew
    rstTemp("GrupoC�digo").value = rstContas("Grupo").value
    rstTemp("GrupoNome").value = strGrupo
    rstTemp("ContaC�digo").value = lngContaAux
    rstTemp("ContaNome").value = rstContas("Descri��o").value
    rstTemp("C�digo").value = rstDados(0).value
    
    
    If (CompStr(strOrigem, "Duplicata")) Then
      rstTemp("Parcela").value = rstDados("Parcela").value
    Else
      rstTemp("Parcela").value = 0
    End If
    
    rstTemp("Descri��o").value = rstDados("Descri��o").value
    
    If bCredito Then
      If strOrigem = "Duplicata" Or strOrigem = "Lan�amento" Then
         rstTemp("Entrada").value = ConvMoedaBase(rstDados(1).value, GetValue(rstDados, "Moeda"), GetValue(rstDados, "Emiss�o"), txtCtrlFinanc(8).Text, GetValue(rstDados, "Pagamento"))
         rstTemp("Sa�da").value = 0
      Else
         rstTemp("Entrada").value = rstDados(1).value / UltimaCotacao(txtCtrlFinanc(8).Text, GetValue(rstDados, "Data"))
         rstTemp("Sa�da").value = 0
      End If
    Else
      If strOrigem = "Duplicata" Or strOrigem = "Lan�amento" Then
         rstTemp("Sa�da").value = -1 * ConvMoedaBase(rstDados(1).value, GetValue(rstDados, "Moeda"), GetValue(rstDados, "Emiss�o"), txtCtrlFinanc(8).Text, GetValue(rstDados, "Pagamento"))
         rstTemp("Entrada").value = 0
      Else
         rstTemp("Sa�da").value = -1 * (rstDados(1).value / UltimaCotacao(txtCtrlFinanc(8).Text, GetValue(rstDados, "Data")))
         rstTemp("Entrada").value = 0
      End If

    End If
    
    rstTemp("Empresa").value = GetValue(rstDados, "Empresa", NUL)
    rstTemp("Origem").value = strOrigem
    rstTemp("Data").value = GetValue(rstDados, "Data")
    rstTemp.update
    
    rstDados.MoveNext
  Loop Until rstDados.EOF
  GravaAuxAnalitico = True
End Function

' SUB.......: RelatorioAnalitico
' Objetivo..: Imprime o relat�rio anal�tico.
' Argumento.: [rstSource]: Recordset com a origem dos dados.
'             [pdDestino]: Destino da impress�o.
'             [strTitulo]: Sub-T�tulo para o relat�rio.
' ---------------------------------------------------------------------------------
Private Sub RelatorioAnalitico(rstSource As Object, pdDestino As PrintDestinoEnum, strTitulo As String, strTitulo2 As String)
  
  Dim wrkAnalitico As KeybReport
  Dim strTiTuloData   As String
  
  ' Somente se houver algum registro no recordset
  If strData = "Vencimento" Then
    strTiTuloData = "Vencto"
  ElseIf strData = "Pagamento" Then
    strTiTuloData = "Pagto"
  ElseIf strData = "Emiss�o" Then
    strTiTuloData = "Emiss."
  ElseIf strData = "Libera��o" Then
    strTiTuloData = "Liber."
  End If

  
  If EstaVazio(rstSource) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
    Exit Sub
  End If
  
  Set wrkAnalitico = New KeybReport
  With wrkAnalitico
    Set .DatabaseName = GlobalDataBase
    Set .Recordset = rstSource
    .Destino = pdDestino
    .WindowTitulo = "Controle Financeiro Anal�tico"
    .Tipo = wrObjectDraw
    .ScaleMode = vbMillimeters
    .AutoRedraw = True
    
    PageHeader wrkAnalitico, "Controle Financeiro Anal�tico"
    
    'Insere linha no Cabe�alho para Informar a Moeda
    If Len(txtCtrlFinanc(8).Text) > 0 Then
      .UltimaSecao.AddLinha "Moeda"
      .UltimaSecao(.UltimaSecao.LinhasCount).AddCampo , wrCSFixedText, "Valores Demonstrados em: " & txtCtrlFinanc(8).Text, wrTACentro
    End If
    
    '
    ' Adiciona uma linha no cabe�alho para subt�tulo
    '
    .Grupo(WRK_HEADER).Header.AddLinha "sub"
    .Grupo(WRK_HEADER).Header("sub").AddCampo , wrCSFixedText, strTitulo, wrTACentro
    .Grupo(WRK_HEADER).Header.AddLinha "sub2"
    .Grupo(WRK_HEADER).Header("sub2").AddCampo , wrCSFixedText, strTitulo2, wrTACentro
    
    ' Criando o grupo principal, Quebra por grupo
    .FontSize = 8
    .FontStyle = wrFSBold
    .AddGrupo "1"
    .Grupo(1).Quebra = "GrupoC�digo"
    .Grupo(1).AddSecao scHeader, 2
    With .Grupo(1).Header.Linha(2)
      .Height = (wrkAnalitico.TextHeight("W") * 2)
      .DrawBorder = wrDBAllBorders
      .AddCampo , wrCSFixedText, "Grupo:", , 13
      .Campo(1).Top = ((.Height / 2) - (.Campo(1).Height / 2))
      .AddCampo , , "GrupoC�digo", wrTADireito, 17
      .Campo(2).Top = .Campo(1).Top
      .Campo(2).Formato = StrZero(0, 9)
      .AddCampo , , "GrupoNome"
      .Campo(3).Top = .Campo(1).Top
    End With
    '
    ' Criando o subGrupo, Quebra por Conta
    '
    .Grupo(1).AddSubGrupo "1"
    .Grupo(1).Subgrupo(1).Quebra = "ContaC�digo"
    .Grupo(1).Subgrupo(1).DrawBorder = wrDBBottomBorder
    .Grupo(1).Subgrupo(1).BorderStyle = wrDot
    .Grupo(1).Subgrupo(1).AddSecao scHeader, 3
    
    .FontStyle = wrFSBold Or wrFSItalic
    With .Grupo(1).Subgrupo(1).Header.Linha(2)
      .DrawBorder = wrDBBottomBorder
      .BorderStyle = wrDot
      .AddCampo , wrCSFixedText, "Conta:", , 13
      .AddCampo , , "ContaC�digo", wrTADireito, 17
      .Campo(2).Formato = StrZero(0, 9)
      .AddCampo , , "ContaNome"
    End With
    
    .FontStyle = wrFSBold
    With .Grupo(1).Subgrupo(1).Header.Linha(3)
    'Vinicius Elyseu(30/05/2016) - Demanda: #120791
      .AddCampo , wrCSFixedText, "C�digo", wrTAEsquerdo, 20  '20
      .AddCampo , wrCSFixedText, "Descri��o", , 50, 31       '69
      .AddCampo , wrCSFixedText, "Empresa", wrTAEsquerdo, 18
      .AddCampo , wrCSFixedText, "Origem", , 17
      .AddCampo , wrCSFixedText, strTiTuloData, wrTAEsquerdo, 10
      .AddCampo , wrCSFixedText, "Entradas", wrTADireito, 20, 130
      .AddCampo , wrCSFixedText, "Sa�das", wrTADireito, 20
      .AddCampo , wrCSFixedText, "Saldo", wrTADireito, 20
    End With
    '
    ' Se��o de detalhes do subgrupo
    '
    .FontStyle = wrFSNormal
    .Grupo(1).Subgrupo(1).AddSecao scDetalhe, 1
    With .Grupo(1).Subgrupo(1).Detalhe.Linha(1)
    'Vinicius Elyseu(12/10/2015) - Projeto: #0 - Hist�ria: #0 - Desenv: #94796
      .AddCampo , , "C�digo", wrTAEsquerdo, 25
      .Campo(1).Formato = StrZero(0, 15)  '6
      .AddCampo , wrCSFixedText, "-", , 1
      .AddCampo , , "Parcela", , 4
      '.Campo(3).Formato = StrZero(0, 2)
      .Campo(3).SuprimirZeros = True
      'Vinicius Elyseu(30/05/2016) - Demanda: #120791
      .AddCampo , , "Descri��o", , 50, 31
      .AddCampo , , "Empresa", , 18
      .AddCampo , , "Origem", , 17
      .AddCampo , , "Data", , 15
      .Campo(7).Formato = FDATA
      .AddCampo , , "Entrada", wrTADireito, 20, 130
      .Campo(8).Formato = FMOEDA
      .AddCampo , , "Sa�da", wrTADireito, 20
      .Campo(9).Formato = FMOEDA
      .AddCampo , , "Saldo", wrTADireito, 20
      .Campo(10).Formato = FMOEDA
    End With
    '
    ' Rodap� do grupo principal: Sub totais
    '
    .Grupo(1).Subgrupo(1).AddSecao scFooter, 1
    With .Grupo(1).Subgrupo(1).Footer.Linha(1)
      .AddCampo , wrCSFixedText, "Total da Conta:", , 30, 27
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSSubTotal, "Entrada", wrTADireito, 20, 130
      .Campo(2).Formato = FMOEDA
      .AddCampo , wrCSSubTotal, "Sa�da", wrTADireito, 20
      .Campo(3).Formato = FMOEDA
      'Protocolo 74461 Trocado o sinal das saidas (-) por isso Sum(entrada) + Sum(Saida)
      .AddCampo , wrCSDataLink, "SUM(Entrada) + SUM(Sa�da)", wrTADireito, 20
      .Campo(4).TableLink = NomeTabeladoRST(rstSource)
      .Campo(4).DataLink = "[ContaC�digo] = {*Quebra}"
      .Campo(4).Formato = FMOEDA
    End With
    
    .Grupo(1).AddSecao scFooter, 1
    With .Grupo(1).Footer.Linha(1)
      .DrawBorder = wrDBBottomBorder
      .AddCampo , wrCSFixedText, "Total do Grupo:", , 30, 27
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSSubTotal, "Entrada", wrTADireito, 20, 130
      .Campo(2).Formato = FMOEDA
      .AddCampo , wrCSSubTotal, "Sa�da", wrTADireito, 20
      .Campo(3).Formato = FMOEDA
      'Protocolo 74461 Trocado o sinal das saidas (-) por isso Sum(entrada) + Sum(Saida)
      .AddCampo , wrCSDataLink, "SUM(Entrada) + SUM(Sa�da)", wrTADireito, 20
      .Campo(4).TableLink = NomeTabeladoRST(rstSource)
      .Campo(4).DataLink = "[GrupoC�digo] = {*Quebra}"
      .Campo(4).Formato = FMOEDA
    End With
        
    ' �ltimo Grupo do relat�rio: Totais dos Grupos de Contas
    ' Crio uma consulta que me retorna apenas o c�digo dos grupos existentes
    ' na tabela tempor�ria
    Dim rstGrupos As Object
    strTitulo = "SELECT DISTINCT [GrupoC�digo], GrupoNome FROM " & NomeTabeladoRST(rstSource) & ";"
    If (AbreRecordset(rstGrupos, strTitulo, dbOpenSnapshot) = WL_OK) Then
      Dim curTotal    As Currency
      Dim curDebito   As Currency       'Total de Sa�das
      Dim curCredito  As Currency       'Total de Entradas
      Dim curSaldo    As Currency       'Total do total
      Dim intContador As Integer
                
                
      .FontStyle = wrFSBold
      .AddGrupo "2", wrDBTopBorder Or wrDBBottomBorder, wrVPNoFinal
      
      .Grupo(2).AddSecao scHeader, 1
      With .Grupo(2).Header.Linha(1)
        .AddCampo , wrCSFixedText, "Saldo Inicial do Banco"
        .AddCampo , wrCSFixedText, Format(SaldoBanco(CDateDef(txtCtrlFinanc(0).Text)), FMOEDA), wrTADireito
      End With
      
      .Grupo(2).Header.AddLinha
      .Grupo(2).Header(2).DrawBorder = wrDBBottomBorder
      .Grupo(2).Header(2).BorderStyle = wrDot
      With .Grupo(2).Header.Linha(2)
        .AddCampo , wrCSFixedText, "Saldo Final do Banco"
                  'Foi adicionado um dia por motivo da fun��o de saldo retornar o saldo inicial do dia
          ' assim n�o considerando os movimentos do dia atual
        .AddCampo , wrCSFixedText, Format(SaldoBanco(CDateDef(txtCtrlFinanc(1).Text) + 1), FMOEDA), wrTADireito
      End With
      
      .Grupo(2).Header.AddLinha
      .Grupo(2).Header(3).DrawBorder = wrDBBottomBorder
      .Grupo(2).Header(3).BorderStyle = wrDot
      With .Grupo(2).Header.Linha(3)
        .AddCampo "Teste", wrCSFixedText, "TOTAIS", wrTACentro
        .Campo(1).FontStyle = wrFSBold
        .Campo(1).FontSize = 9
      End With
      
      .Grupo(2).Header.AddLinha
      With .Grupo(2).Header.Linha(4)
        .AddCampo , wrCSFixedText, "Grupo de Contas", , 30
        .AddCampo , wrCSFixedText, "Cr�dito", wrTADireito, 20, 130
        .AddCampo , wrCSFixedText, "D�bito", wrTADireito, 20
        .AddCampo , wrCSFixedText, "Saldo", wrTADireito, 20
      End With
      .FontStyle = wrFSNormal
      
      ' Adicionando quantas linhas forem necess�rias para imprimir os valores
      rstGrupos.MoveFirst
      intContador = 5
      Do Until rstGrupos.EOF
        .Grupo(2).Header.AddLinha
        With .Grupo(2).Header.Linha(intContador)
          .AddCampo , wrCSFixedText, CStr(GetValue(rstGrupos, 0)), wrTADireito, 17
          .AddCampo , wrCSFixedText, GetValue(rstGrupos, 1), , 81
          
          curTotal = Soma("Entrada", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
          curCredito = curCredito + curTotal
          .AddCampo , wrCSFixedText, Format$(curTotal, FMOEDA), wrTADireito, 20, 130
          
          curTotal = Soma("Sa�da", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value))
          curDebito = curDebito + curTotal
          .AddCampo , wrCSFixedText, Format$(curTotal, FMOEDA), wrTADireito, 20
                    
          curTotal = Soma("Entrada + Sa�da", NomeTabeladoRST(rstSource), "[GrupoC�digo] = " & CStr(rstGrupos(0).value), 0)
          curSaldo = curSaldo + curTotal
          .AddCampo , wrCSFixedText, Format$(curTotal, FMOEDA), wrTADireito, 20
        End With
        Inc intContador
        rstGrupos.MoveNext
      Loop
      '
      ' �ltima linha Totais
      '
      .Grupo(2).AddSecao scFooter, 1, wrDBTopBorder
      .Grupo(2).Footer.BorderStyle = wrDot
      With .Grupo(2).Footer.Linha(1)
        .AddCampo , wrCSFixedText, "Totais:"
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, Format$(curCredito, FMOEDA), wrTADireito, 20, 130
        .AddCampo , wrCSFixedText, Format$(curDebito, FMOEDA), wrTADireito, 20
        .AddCampo , wrCSFixedText, Format$(curSaldo, FMOEDA), wrTADireito, 20
      End With
    End If
    FechaRecordset rstGrupos
  End With
    
  wrkAnalitico.BeginPrint gTipoDB
  wrkAnalitico.EndPrint
  Set wrkAnalitico = Nothing
End Sub

Private Function SaldoBanco(Data As Date) As Double
  Dim rstBancos    As Object
  Dim strBancos    As String
  Dim lBancos(1)   As Long        'Bancos Inicial e Final

  lBancos(IDX_INICIO) = Min(CLngDef(txtCtrlFinanc(4).Text), CLngDef(txtCtrlFinanc(5).Text))
  lBancos(IDX_FINAL) = Max(CLngDef(txtCtrlFinanc(4).Text), CLngDef(txtCtrlFinanc(5).Text))
  
  strBancos = "Select * from Bancos "
  If ((lBancos(IDX_INICIO) > 0) And (lBancos(IDX_FINAL) > 0)) Then
    If (lBancos(IDX_INICIO) = lBancos(IDX_FINAL)) Then
      Concat strBancos, " WHERE Banco = " & CStr(lBancos(IDX_INICIO))
    Else
      Concat strBancos, " WHERE (Banco BETWEEN ", CStr(lBancos(IDX_INICIO)), " AND ", CStr(lBancos(IDX_FINAL)), ")"
    End If
  ElseIf (lBancos(IDX_INICIO) > 0) Then
    Concat strBancos, " WHERE Banco >= " & CStr(lBancos(IDX_INICIO))
  ElseIf (lBancos(IDX_FINAL) > 0) Then
    Concat strBancos, " WHERE Banco <= ", CStr(lBancos(IDX_FINAL))
  End If
  Concat strBancos & " ORDER BY Banco"

  SaldoBanco = ZERO
  If AbreRecordset(rstBancos, strBancos, dbOpenSnapshot) = WL_OK Then
    Do
      SaldoBanco = SaldoBanco + SaldoInicial(GetValue(rstBancos, "Banco", ZERO), Data, False, strMoeda:=txtCtrlFinanc(8).Text, StrDescMoeda:=lblNomes(6).Caption, sConciliado:=cboConciliado.Text)
      rstBancos.MoveNext
    Loop Until rstBancos.EOF
  End If
  FechaRecordset rstBancos
  
End Function

'FUNCTION..: AppendTempOrcado
'Objetivo..: Adiciona os dados obtidos das tabelas de Lan�amentos e Duplicatas
'            na tabela tempor�ria criada para imprimir o relat�rio.
'Argumentos: [rstTemp]: Recordset da tabela auxiliar.
'            [rstSrc] : Recordset com os Grupos e Contas.
'            [lBco]   : Matriz com os bancos escolhidos pelo usu�rio.
'            [dDatas] : Matriz com as datas escolhidas pelo usu�rio.
'Retorna...: True se terminar, False se o usu�rio cancelar
Private Function AppendTempOrcado(rstTemp As Object, rstSrc As Object, lBco() As Long, dDatas() As Date) As Boolean
    Dim curEntrada     As Currency
    Dim curSaida       As Currency
    Dim curAEntrar   As Currency
    Dim curASair    As Currency
    Dim lngConta       As Double
    Dim lngGrupo       As Long
    Dim strGrupo       As String
    Dim strWhere       As String
    Dim strWhere1      As String
    Dim strNomeTabela  As String
    Dim X              As Integer
    Dim dDatasMes(1)   As Date
    Dim lngContaAux    As Double
    Dim genTemp As New CGenericRecordset
    
    genTemp.Initialize rstTemp
  
    For X = Month(dDatas(0)) To Month(dDatas(1))
        If X = Month(dDatas(0)) Then
            dDatasMes(0) = dDatas(0)
            dDatasMes(1) = LastDayS(dDatas(0))
        ElseIf X = Month(dDatas(1)) Then
            dDatasMes(0) = FirstDayS(dDatas(1))
            dDatasMes(1) = dDatas(1)
        Else
            dDatasMes(0) = CDateDef("01/" & str(X) & "/" & Year(dDatas(0)))
            dDatasMes(1) = LastDayS(dDatasMes(0))
        End If
        curEntrada = ZERO
        curSaida = ZERO
        curAEntrar = ZERO
        curASair = ZERO
        rstSrc.MoveFirst
        Do
            If (lngGrupo <> GetValue(rstSrc, "Grupo")) Then
                lngGrupo = GetValue(rstSrc, "Grupo")   'C�digo de Descri��o do Grupo
                If UsandoModelo Then
                    strGrupo = GetValue(rstSrc, "DescGrupo")
                Else
                    strGrupo = GetFieldValue("Descri��o", "Grupos", "[C�digo] = " & CStr(lngGrupo))
                End If
                SimpleMsgBar "Calculando Grupo: " & StrZero(lngGrupo, 9) & " - " & strGrupo & " - M�s : " & str(X)
            End If
            If mbolCancelou Then Exit Function
            DoEvents                          'Permite ao usu�rio cancelar a gera��o
            If lngContaAux <> GetValue(rstSrc, "ContaAuxiliar", ZERO) Then
                curSaida = 0
                curEntrada = 0
                curASair = 0
                curAEntrar = 0
            End If
            lngConta = GetValue(rstSrc, "C�digo")
            'Resolve a instru��o de Transfer�ncias com o Banco de Destino
            strWhere = AddTransfBancarias(lngConta, dDatasMes(), lBco(), False)
            If (Len(strWhere)) Then
                If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
                    curAEntrar = curAEntrar + Soma("Valor", "[Transf Banc�ria]", strWhere)
                    curEntrada = curEntrada + Soma("Valor", "[Transf Banc�ria]", strWhere)
                Else
                    curAEntrar = curAEntrar + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "[Transf Banc�ria]", strWhere)
                    curEntrada = curEntrada + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "[Transf Banc�ria]", strWhere)
                End If
            End If
            If mbolCancelou Then Exit Function
            DoEvents
            'Resolve a instru��o de Aplica��es para opera��es de cr�dito
            strWhere = AddAplicacoes(lngConta, lBco(), dDatasMes(), True)
            If (Len(strWhere)) Then
                If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
                    curAEntrar = curAEntrar + Soma("Valor", "Aplica��es", strWhere)
                    curEntrada = curEntrada + Soma("Valor", "Aplica��es", strWhere)
                Else
                    curAEntrar = curAEntrar + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "Aplica��es", strWhere)
                    curEntrada = curEntrada + Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "Aplica��es", strWhere)
                End If
            End If
            If mbolCancelou Then Exit Function
            DoEvents
            'Resolve a instru��o de Aplica��es para opera��es de D�bito
            strWhere = AddAplicacoes(lngConta, lBco(), dDatasMes(), False)
            If (Len(strWhere)) Then
                If TemMoeda(txtCtrlFinanc(8).Text, lblNomes(6).Caption) = False Then
                    curASair = curASair - Soma("Valor", "Aplica��es", strWhere)
                    curSaida = curSaida - Soma("Valor", "Aplica��es", strWhere)
                Else
                    curASair = curASair - Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "Aplica��es", strWhere)
                    curSaida = curSaida - Soma("Valor / (SELECT VALOR FROM COTA��ES  WHERE MOEDA = '" & txtCtrlFinanc(8).Text & "' AND DATA = Data)", "Aplica��es", strWhere)
                End If
            End If
            If mbolCancelou Then Exit Function
            DoEvents
            'Resolve a instru��o para Duplicatas Recebidas ou A Receber
            If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Duplicatas")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Receber")) Then
                strWhere = AddLancDupl(lngConta, lBco(), dDatasMes(), False, , 0) 'Realizado = 0 (AEntrar ou ASair)
                strWhere1 = AddLancDupl(lngConta, lBco(), dDatasMes(), False, , 1) 'Realizado = 1 (Entrada ou Sa�da)
                If (Len(strWhere)) And (Len(strWhere1)) Then
                    'Protocolo 74572
                    curAEntrar = curAEntrar + SomarMoedas("Duplicatas", strWhere, txtCtrlFinanc(8).Text)
                    curEntrada = curEntrada + SomarMoedas("Duplicatas", strWhere1, txtCtrlFinanc(8).Text)
                End If
                If mbolCancelou Then Exit Function
                DoEvents
            End If
            'Resolve a instru��o para Duplicatas Pagas ou A Pagar
            If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Duplicatas")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Pagar")) Then
                strWhere = AddLancDupl(lngConta, lBco(), dDatasMes(), True, , 0) 'Realizado = 0 (AEntrar ou ASair)
                strWhere1 = AddLancDupl(lngConta, lBco(), dDatasMes(), True, , 1) 'Realizado = 1 (Entrada ou Sa�da)
                If (Len(strWhere)) And (Len(strWhere1)) Then
                    'Protocolo 74572
                    curASair = curASair - SomarMoedas("Duplicatas", strWhere, txtCtrlFinanc(8).Text)
                    curSaida = curSaida - SomarMoedas("Duplicatas", strWhere1, txtCtrlFinanc(8).Text)
                End If
                If mbolCancelou Then Exit Function
                DoEvents
            End If
            'Resolve a instru��o para Lan�amentos Recebidos ou A Receber
            If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Lan�amentos")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Receber")) Then
                strWhere = AddLancDupl(lngConta, lBco(), dDatasMes(), False, , 0)  'Realizado = 0 (AEntrar ou ASair)
                strWhere1 = AddLancDupl(lngConta, lBco(), dDatasMes(), False, , 1) 'Realizado = 1 (Entrada ou Sa�da)
                If (Len(strWhere)) And (Len(strWhere1)) Then
                    'Protocolo 74572
                    curAEntrar = curAEntrar + SomarMoedas("Lan�amentos", strWhere, txtCtrlFinanc(8).Text)
                    curEntrada = curEntrada + SomarMoedas("Lan�amentos", strWhere1, txtCtrlFinanc(8).Text)
                End If
                If mbolCancelou Then Exit Function
                DoEvents
            End If
            'Resolve a instru��o para Lan�amentos Pagos ou A Pagar
            If ((cboOrigem.Text = "Ambos") Or (cboOrigem = "Lan�amentos")) And ((cboTipo.Text = "Todos") Or (cboTipo.Text = "� Pagar")) Then
                strWhere = AddLancDupl(lngConta, lBco(), dDatasMes(), True, , 0)  'Realizado = 0 (AEntrar ou ASair)
                strWhere1 = AddLancDupl(lngConta, lBco(), dDatasMes(), True, , 1) 'Realizado = 1 (Entrada ou Sa�da)
                If (Len(strWhere)) And (Len(strWhere1)) Then
                    'Protocolo 74572
                    curASair = curASair - SomarMoedas("Lan�amentos", strWhere, txtCtrlFinanc(8).Text)
                    curSaida = curSaida - SomarMoedas("Lan�amentos", strWhere1, txtCtrlFinanc(8).Text)
                End If
                If mbolCancelou Then Exit Function
                DoEvents
            End If
            'Grava os dados na tabela tempor�ria
            If lngConta <> 0 Then
                'PROTOCOLO 72265 - Sistema n�o tratava Soma dos Or�amentos por centro de custo.
                If (tabCtrlFinanc.SelectedItem.Key = "or�ado" And ((Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND (Centro BETWEEN " & txtCtrlFinanc(9).Text & " AND " & txtCtrlFinanc(10).Text & ")" & " AND (Per�odo BETWEEN " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(0).Text)), True) & " AND " & InverteData(FirstDayS(CDateDef(txtCtrlFinanc(1).Text)), True) & ")", ZERO) <> 0) Or ((curEntrada <> 0) Or (curSaida <> 0)) Or ((curAEntrar <> 0) Or (curASair <> 0)))) Then
                    If UsandoModelo Then
                        lngContaAux = GetValue(rstSrc, "ContaAuxiliar", ZERO)
                    Else
                        lngContaAux = lngConta
                    End If
                    genTemp.FindFirst "[GrupoC�digo]=" & lngGrupo & " AND [ContaC�digo]=" & lngContaAux
                    If genTemp.NoMatch Then
                        rstTemp.AddNew
                    Else
                        genTemp.Edit
                    End If
                    ' Zerando vari�veis de Saldo que n�o est�o preenchidos
                    If IsNull(rstTemp("Saldo1").value) Then
                        rstTemp("Saldo1").value = 0
                    End If
                    If IsNull(rstTemp("Saldo2").value) Then
                        rstTemp("Saldo2").value = 0
                    End If
                    If IsNull(rstTemp("Saldo3").value) Then
                        rstTemp("Saldo3").value = 0
                    End If
                    If IsNull(rstTemp("Saldo4").value) Then
                        rstTemp("Saldo4").value = 0
                    End If
                    If IsNull(rstTemp("Saldo5").value) Then
                        rstTemp("Saldo5").value = 0
                    End If
                    If IsNull(rstTemp("Saldo6").value) Then
                        rstTemp("Saldo6").value = 0
                    End If
                    If IsNull(rstTemp("Saldo7").value) Then
                        rstTemp("Saldo7").value = 0
                    End If
                    If IsNull(rstTemp("Saldo8").value) Then
                        rstTemp("Saldo8").value = 0
                    End If
                    If IsNull(rstTemp("Saldo9").value) Then
                        rstTemp("Saldo9").value = 0
                    End If
                    If IsNull(rstTemp("Saldo10").value) Then
                        rstTemp("Saldo10").value = 0
                    End If
                    If IsNull(rstTemp("Saldo11").value) Then
                        rstTemp("Saldo11").value = 0
                    End If
                    If IsNull(rstTemp("Saldo12").value) Then
                        rstTemp("Saldo12").value = 0
                    End If
                    'Zerando vari�veis de Or�ado que n�o est�o preenchidos
                    If IsNull(rstTemp("Or�ado1").value) Then
                        rstTemp("Or�ado1").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado2").value) Then
                        rstTemp("Or�ado2").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado3").value) Then
                        rstTemp("Or�ado3").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado4").value) Then
                        rstTemp("Or�ado4").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado5").value) Then
                        rstTemp("Or�ado5").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado6").value) Then
                        rstTemp("Or�ado6").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado7").value) Then
                        rstTemp("Or�ado7").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado8").value) Then
                        rstTemp("Or�ado8").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado9").value) Then
                        rstTemp("Or�ado9").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado10").value) Then
                        rstTemp("Or�ado10").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado11").value) Then
                        rstTemp("Or�ado11").value = 0
                    End If
                    If IsNull(rstTemp("Or�ado12").value) Then
                        rstTemp("Or�ado12").value = 0
                    End If
                    rstTemp("GrupoC�digo").value = lngGrupo
                    rstTemp("GrupoNome").value = strGrupo
                    rstTemp("ContaC�digo").value = lngContaAux
                    rstTemp("ContaNome").value = rstSrc("Descri��o").value
                    If UsandoModelo Then
                        rstTemp("Saldo" & X).value = (curEntrada - curSaida)
                        rstTemp("Or�ado" & X).value = Soma("Valor", "[Or�amentos de Contas]", "Conta in (Select [Conta Cont�bil] from [Contas de Contas Auxiliares] where Conta = " & lngContaAux & ") AND (Per�odo BETWEEN " & InverteData(dDatasMes(0), True) & " AND " & InverteData(dDatasMes(1), True) & ")", ZERO)
                    Else
                        rstTemp("Saldo" & X).value = (curEntrada - curSaida)
                        'PROTOCOLO 72265 - Sistema n�o tratava Soma dos Or�amentos por centro de custo.
                        rstTemp("Or�ado" & X).value = Soma("Valor", "[Or�amentos de Contas]", "Conta = " & CStr(lngConta) & " AND (Centro BETWEEN " & txtCtrlFinanc(9).Text & " AND " & txtCtrlFinanc(10).Text & ")" & " AND (Per�odo BETWEEN " & InverteData(dDatasMes(0), True) & " AND " & InverteData(dDatasMes(1), True) & ")", ZERO)
                    End If
                    rstTemp.update
                End If
            End If
            If Not UsandoModelo Then
                curSaida = 0
                curEntrada = 0
            End If
            'Move a tabela origem para o pr�ximo registro
            rstSrc.MoveNext
        Loop Until rstSrc.EOF
    Next
    'Soma os saldos e Or�ados
    'pt. 84357 Abner Luidi Hempkemaier (07/12/2007)
    If rstTemp.Recordcount <> 0 Then
        rstTemp.MoveFirst
        Do
            If TypeOf rstTemp Is dao.Recordset Then rstTemp.Edit
            rstTemp("TotalSaldo").value = GetValue(rstTemp, "Saldo1", ZERO) + GetValue(rstTemp, "Saldo2", ZERO) + GetValue(rstTemp, "Saldo3", ZERO) + GetValue(rstTemp, "Saldo4", ZERO) + GetValue(rstTemp, "Saldo5", ZERO) + GetValue(rstTemp, "Saldo6", ZERO) + GetValue(rstTemp, "Saldo7", ZERO) + GetValue(rstTemp, "Saldo8", ZERO) + GetValue(rstTemp, "Saldo9", ZERO) + GetValue(rstTemp, "Saldo10", ZERO) + GetValue(rstTemp, "Saldo11", ZERO) + GetValue(rstTemp, "Saldo12", ZERO)
            rstTemp("TotalOrcado").value = GetValue(rstTemp, "or�ado1", ZERO) + GetValue(rstTemp, "or�ado2", ZERO) + GetValue(rstTemp, "or�ado3", ZERO) + GetValue(rstTemp, "or�ado4", ZERO) + GetValue(rstTemp, "or�ado5", ZERO) + GetValue(rstTemp, "or�ado6", ZERO) + GetValue(rstTemp, "or�ado7", ZERO) + GetValue(rstTemp, "or�ado8", ZERO) + GetValue(rstTemp, "or�ado9", ZERO) + GetValue(rstTemp, "or�ado10", ZERO) + GetValue(rstTemp, "or�ado11", ZERO) + GetValue(rstTemp, "or�ado12", ZERO)
            rstTemp.update
            rstTemp.MoveNext
        Loop Until rstTemp.EOF
    End If
    NomeAuxiliar = NomeTabeladoRST(rstTemp)
    strNomeTabela = "SELECT * FROM " & NomeTabeladoRST(rstTemp) & " ORDER BY [GrupoC�digo],[ContaC�digo]"
    FechaRecordset rstTemp
    AbreRecordset rstTemp, strNomeTabela, dbOpenSnapshot
    AppendTempOrcado = True
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
