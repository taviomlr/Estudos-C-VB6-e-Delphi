VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frptFluxoCaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fluxo de Caixa"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "rptFluxoCaixa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFluxo 
      Caption         =   "Considerar títulos em atraso na composição do saldo anterior"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   330
      TabIndex        =   47
      Top             =   5790
      Value           =   1  'Checked
      Width           =   6435
   End
   Begin VB.Frame fraPedidos 
      Caption         =   "Considerar Pedidos Pendentes de:"
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
      Height          =   915
      Left            =   2820
      TabIndex        =   42
      Top             =   4800
      Width           =   4905
      Begin VB.CheckBox chkFluxo 
         Caption         =   "Serviços a Receber"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   150
         TabIndex        =   46
         Top             =   510
         Width           =   1725
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "Serviços a Pagar"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   9
         Left            =   2160
         TabIndex        =   45
         Top             =   510
         Width           =   1515
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "Vendas"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   150
         TabIndex        =   44
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "Compras"
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   7
         Left            =   2160
         TabIndex        =   43
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame FraOutros 
      Caption         =   "Outros"
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
      Height          =   915
      Left            =   240
      TabIndex        =   38
      Top             =   4800
      Width           =   2625
      Begin VB.ComboBox cboFluxo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "&Conciliado:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   360
         TabIndex        =   39
         Top             =   390
         Width           =   780
      End
   End
   Begin VB.Frame fraFluxo 
      Caption         =   "#"
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
      Height          =   4425
      Left            =   240
      TabIndex        =   32
      Top             =   360
      Width           =   7485
      Begin VB.CheckBox chkImprimeBancoSemMovimento 
         Caption         =   "Imprimir Banco Sem Movimentação"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4500
         TabIndex        =   48
         Top             =   810
         Width           =   2805
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   8
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   40
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "Quebrar por Data"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   4500
         TabIndex        =   9
         Top             =   510
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   7
         Left            =   1200
         TabIndex        =   26
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "Imprimir Ra&zão"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2730
         TabIndex        =   8
         Top             =   510
         Width           =   1455
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "I&mprimir Resumo"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   2730
         TabIndex        =   7
         Top             =   780
         Width           =   1455
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "Imprimir Descrição &completa  + Controle"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   2730
         TabIndex        =   6
         Top             =   1020
         Width           =   3495
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "&Quebrar por Banco"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   4500
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   6
         Left            =   1200
         TabIndex        =   23
         Top             =   3450
         Width           =   1335
      End
      Begin VB.CheckBox chkFluxo 
         Caption         =   "Atualizar &Saldos"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2730
         TabIndex        =   4
         Top             =   240
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   5
         Left            =   1200
         TabIndex        =   20
         Top             =   2910
         Width           =   1335
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   4
         Left            =   1200
         TabIndex        =   17
         Top             =   2550
         Width           =   1335
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   3
         Left            =   1200
         TabIndex        =   14
         Top             =   1950
         Width           =   1335
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   11
         Top             =   1590
         Width           =   1335
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtFluxo 
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFluxo 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   90
         TabIndex        =   37
         Top             =   3810
         Width           =   585
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo"
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
         Index           =   5
         Left            =   60
         TabIndex        =   35
         Top             =   3210
         Width           =   1380
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "Código do Banco"
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
         Index           =   2
         Left            =   60
         TabIndex        =   33
         Top             =   1290
         Width           =   1470
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "Código da Conta"
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
         Index           =   6
         Left            =   60
         TabIndex        =   34
         Top             =   2310
         Width           =   1425
      End
      Begin VB.Line lnhFluxo 
         BorderColor     =   &H80000010&
         Index           =   7
         X1              =   0
         X2              =   7425
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Line lnhFluxo 
         BorderColor     =   &H80000010&
         Index           =   5
         X1              =   0
         X2              =   7425
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Line lnhFluxo 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   30
         X2              =   7425
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "Nro pág. inicial:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   60
         TabIndex        =   41
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "&Moeda:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   600
         TabIndex        =   25
         Top             =   4020
         Width           =   540
      End
      Begin VB.Label lblNome 
         Caption         =   "lblNome(5)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2640
         TabIndex        =   31
         Top             =   3960
         UseMnemonic     =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblNome 
         Caption         =   "lblNome(4)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   24
         Top             =   3450
         UseMnemonic     =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "Códi&go:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   600
         TabIndex        =   22
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label lblNome 
         Caption         =   "lblNome(3)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   21
         Top             =   2910
         UseMnemonic     =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblNome 
         Caption         =   "lblNome(2)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   18
         Top             =   2550
         UseMnemonic     =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblNome 
         Caption         =   "lblNome(1)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   15
         Top             =   1950
         UseMnemonic     =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblNome 
         Caption         =   "lblNome(0)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         Top             =   1590
         UseMnemonic     =   0   'False
         Width           =   3615
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "Fina&l:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   780
         TabIndex        =   19
         Top             =   2940
         Width           =   375
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "I&nicial:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   690
         TabIndex        =   16
         Top             =   2610
         Width           =   450
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "&Final:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   750
         TabIndex        =   13
         Top             =   2010
         Width           =   375
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "&Inicial:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   690
         TabIndex        =   10
         Top             =   1650
         Width           =   450
      End
      Begin VB.Line lnhFluxo 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   7425
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "D&ata Final:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   2
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "&Data Inicial:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdFluxo 
      Cancel          =   -1  'True
      Caption         =   "#"
      Height          =   375
      Index           =   2
      Left            =   6630
      TabIndex        =   30
      Top             =   6270
      Width           =   1215
   End
   Begin VB.CommandButton cmdFluxo 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   29
      Top             =   6270
      Width           =   1215
   End
   Begin VB.CommandButton cmdFluxo 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   28
      Top             =   6270
      Width           =   1215
   End
   Begin ComctlLib.TabStrip tabFluxo 
      Height          =   6195
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   10927
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sintético por Data"
            Key             =   "sintetico"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Analítico"
            Key             =   "analitico"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sintético por Conta"
            Key             =   "sintetico por conta"
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
Attribute VB_Name = "frptFluxoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Qual a diferença entre o cálculo com a quebra e sem a quebra? Se você
'já deu uma olhada no código das funções que obtém os dados para o cál-
'culo não foi difícil perceber que há uma função para cada caso e que
'cada função chama uma mesma função que grava os dados na tabela auxiliar.
'Isto acontece porque a diferença está na forma como estes dados são
'gravados na tabela, ou seja, a seqüência da gravação. Nos relatórios
'cuja opção de quebrar por empresa está habilitada as função organizam
'os dados filtrando por banco e depois, dentro de cada banco, organiza
'por data. Quando não há quebra as funções, simplesmente, organizam por
'data, não há a preocupação em obter os saldos por banco, mas a função
'obtém um saldo inicial geral de todos os bancos solicitados. Note que
'só irão sair no cálculo os bancos cujo campo: Constar no Fluxo de Caixa
'esteja marcado como verdadeiro. Exceto quando o usuário escolhe um único
'banco em particular, esta opção deve estar assinalada para que o banco
'possa aparecer no relatório.
Option Explicit

Private Const IDS_SINTETICO = 172             'Índice do caption do frame no arquivo de recursos
Private Const IDS_ANALITICO = 173             'Ídem
Private Const IDS_EXTRATO = 192               'Ídem
Private Const IDS_EXTSINTETICO = 193          'Extrato Bancário Sintético
Private Const IDS_EXTANALITICO = 194          'Extrato Bancário Analítico
Private Const KEY_EXTRATO$ = "extrato"        'Chave do Tab um
Private Const KEY_SINTETICO$ = "sintetico"    'Chave do Tab dois
Private Const KEY_SINTETICO_CONTA$ = "sintetico por conta" 'Chave do Tab três
Private Const KEY_ANALITICO$ = "analitico"    'Chave do Tab três
Private Const KEY_SALDO$ = "UpdateSaldo"      'Chave do arquivo .ini
Private Const KEY_RAZAO$ = "PrintRazao"       'Chave do arquivo
Private Const KEY_DESC$ = "PrintDesc"         'Chave do arquivo
Private Const KEY_RESUMO$ = "Resumo"          'Chave do arquivo
Private Const KEY_QUEBRAR$ = "Quebra"         'Chave do arquivo
Private Const CAPLIC$ = "Aplicação"           'Define aplicação no fluxo analítico
Private Const CTRANS$ = "Transferência"       'Define transferência no fluxo analítico
Private Const TIPO_EXTRATO = 1                'Tipo do relatório
Private Const TIPO_FLUXO = 0
Private Const TIPO_MOVIMENTO = 2
Private Const DADOS_TRANSF = 1                'Define Transf Bancária para a tabela auxiliar
Private Const DADOS_APLIC = 2                 'Define Aplicação para a tabela auxiliar
Private Const DADOS_LANC = 3                  'Define Lançamentos ou Duplicatas para a tabela auxiliar
Private dtInicial        As Date              'Data Inicial Informada no Relatório
Private dtFinal          As Date              'Data Final Informada no Relatório
Private dblCotacao       As Double            'Valor da Cotação Na Data
Private mbolCancelou     As Boolean
Private mbitTipo         As Byte              'Tipo do Relatório
Private BQuebraData      As Boolean           'Define se há quebra por data
Private rstPrevisao      As Object
Private fdsPrevisao(8)   As FieldStruct
  

Private Sub chkFluxo_Click(Index As Integer)
    If Index = 4 Then       '4 == Quebra por Banco
        chkImprimeBancoSemMovimento.value = chkFluxo(Index).value
        chkFluxo(0).Enabled = EnableUpdate()
        'A rotina não pode efetuar a atualização dos Saldos Bancários quando
        'o relatório não é quebrado por Bancos. Também não é possível fazer a atualização
        'se o usuário filtra por Conta ou por Centro de Custo, já que o cálculo é efetuado
        'apenas nas movimentações solicitadas por ele. Desta forma, desabilito o campo
        'quando o usuário usa um desses filtros ou quando não quebra o relatório por
        'banco. Assim, a função que faz a atualização dos saldos tem como verificar se
        'pode ou não executar a atualização corretamente.
    End If
End Sub

Private Sub chkFluxo_GotFocus(Index As Integer)
    FluxoMsgBar chkFluxo(Index).TabIndex
End Sub

Private Sub cmdFluxo_Click(Index As Integer)
    mbolCancelou = False
    Screen.MousePointer = vbHourglass
    If Index < 2 Then
        If Not EData(txtFluxo(0).Text) Then
            MsgFunc "O campo 'Data Inicial' não contém uma data válida."
            Exit Sub
        End If
        If Not EData(txtFluxo(1).Text) Then
            MsgFunc "O campo 'Data Final' não contém uma data válida."
            Exit Sub
        End If
        dtInicial = txtFluxo(0).Text
        dtFinal = txtFluxo(1).Text
        If IsValid(txtFluxo(0).Text) And IsValid(txtFluxo(1).Text) Then
            If CDateDef(txtFluxo(1).Text) < CDateDef(txtFluxo(0).Text) Then
                MsgFunc "Data Final menor que Data Inicial"
                Exit Sub
            End If
        End If
        cmdFluxo(0).Enabled = False
        cmdFluxo(1).Enabled = False
        cmdFluxo(2).Caption = LoadResString(IDS_CANCELAR)
        If chkFluxo(6).value = vbChecked Or chkFluxo(7).value = vbChecked Or chkFluxo(8).value = vbChecked Or chkFluxo(9).value = vbChecked Then
            PedidosPendentes
        End If
        CriaFiltroFluxo IIf((Index = 0), wrToWindow, wrToPrinter)
        cmdFluxo(0).Enabled = True
        cmdFluxo(1).Enabled = True
        cmdFluxo(2).Caption = LoadResString(IDS_FECHAR)
    Else
        If cmdFluxo(0).Enabled Then
            Unload Me
        Else
            MsgBar LoadResString(171) & LoadResString(14)
            mbolCancelou = True
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

'PROPERTY..: Tipo
'Objetivo..: Define ou retorna uma string que define o tipo do relatório que
'            deve ser gerado. As opções possíveis são: Extrato Bancário ou
'            Fluxo de Caixa, sendo os números 0 (zero) e 1 (um) utilizados
'            para a definição respectivamente.
Public Property Get Tipo() As Byte
    Tipo = mbitTipo
End Property

Public Property Let Tipo(ByVal nTipo As Byte)
    mbitTipo = nTipo
    'Altera o Caption o Formulário conforme o tipo do relatório
    chkFluxo(10).value = vbChecked
    If nTipo = TIPO_EXTRATO Then
        Caption = LoadResString(IDS_EXTRATO)
        tabFluxo.Tabs(KEY_SINTETICO).Selected = True
        tabFluxo.Tabs.Remove KEY_SINTETICO_CONTA
        chkFluxo(10).value = vbUnchecked
        chkFluxo(10).Visible = False
    ElseIf nTipo = TIPO_MOVIMENTO Then
        Caption = "Movimento de Caixa"
        tabFluxo.Tabs.Remove KEY_SINTETICO
        tabFluxo.Tabs.Remove KEY_SINTETICO_CONTA
        fraFluxo.Caption = "Movimento de Caixa Analítico"
        chkFluxo(5).value = False
        txtFluxo(8).Text = Empty
        txtFluxo(8).Visible = False
        lblFluxo(13).Visible = False
        chkFluxo(10).value = vbUnchecked
        chkFluxo(10).Visible = False
    End If
    fraPedidos.Visible = (nTipo = TIPO_FLUXO)
End Property

Private Sub Form_Load()
    Dim strFoxIni As String
    
    'Configurando a abertura da janela
    CenterForm Me
    'Todos os campos que não são mais padrão são devido a solicitação do FM 31/07/2002
    lblNome(2).Caption = NUL
    lblNome(3).Caption = NUL
    lblNome(5).Caption = NUL
    'Exibindo valores padrão nos campos respectivos
    txtFluxo(0).Text = Format$(Date, FDATA)
    txtFluxo(1).Text = Format$(Date, FDATA)
    'Exibindo o menor e maior número de bancos encontrados
    txtFluxo(2).Text = MinValue("Banco", "Bancos", NUL)
    txtFluxo(3).Text = MaxValue("Banco", "Bancos", NUL)
    'Se o usuário não deseja controlar centro de custo não exibe o controle na tela
    If Not CentrodeCusto(MFinanceiro) Then
        lblFluxo(5).Visible = False
        lblFluxo(9).Visible = False
        lblNome(4).Visible = False
        txtFluxo(6).Visible = False
    Else
        lblNome(4).Caption = ""
    End If
    cmdFluxo(2).Caption = LoadResString(IDS_FECHAR)
    tabFluxo.Tabs(KEY_SINTETICO).Selected = True
    'Verifica se a Caixa de Verificação de atualização de Saldos deve ser marcada
    'ou não. Opção de impressão da razão social da empresa e impressão do campo
    'descrição completo.
    strFoxIni = IniFileName()
    chkFluxo(1).value = ((LerArquivoASCII(SEC_WKIF, KEY_RAZAO, strFoxIni) = "1") And vbChecked)
    chkFluxo(2).value = ((LerArquivoASCII(SEC_WKIF, KEY_DESC, strFoxIni) = "1") And vbChecked)
    chkFluxo(3).value = ((LerArquivoASCII(SEC_WKIF, KEY_RESUMO, strFoxIni) = "1") And vbChecked)
    lblNome(5).Caption = NUL
    cboFluxo.AddItem "Todos"
    cboFluxo.AddItem "Sim"
    cboFluxo.AddItem "Não"
    cboFluxo.Text = "Todos"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strFoxIni As String
  
    'Grava o valor da checkbox de atualização de saldos no arquivo .ini, Impressão
    'da razão social e impressão do campo descrição.
    strFoxIni = IniFileName()
    GravarArquivoASCII SEC_WKIF, KEY_RAZAO, chkFluxo(1).value, strFoxIni
    GravarArquivoASCII SEC_WKIF, KEY_DESC, chkFluxo(2).value, strFoxIni
    GravarArquivoASCII SEC_WKIF, KEY_RESUMO, chkFluxo(3).value, strFoxIni
    Set frptFluxoCaixa = Nothing
End Sub

Private Sub tabFluxo_Click()
    If mbitTipo = TIPO_FLUXO Then
        If tabFluxo.SelectedItem.Key = KEY_SINTETICO Then
            fraFluxo.Caption = LoadResString(IDS_SINTETICO)
        Else
            fraFluxo.Caption = LoadResString(IDS_ANALITICO)
        End If
    ElseIf mbitTipo = TIPO_EXTRATO Then
        If tabFluxo.SelectedItem.Key = KEY_SINTETICO Then
            fraFluxo.Caption = LoadResString(IDS_EXTSINTETICO)
        Else
            fraFluxo.Caption = LoadResString(IDS_EXTANALITICO)
        End If
    ElseIf mbitTipo = TIPO_MOVIMENTO Then
        fraFluxo.Caption = "Movimento de Caixa Analítico"
    End If
    If mbitTipo <> TIPO_MOVIMENTO Then
        chkFluxo(1).Visible = (tabFluxo.SelectedItem.Key = KEY_ANALITICO)
        chkFluxo(2).Visible = (tabFluxo.SelectedItem.Key = KEY_ANALITICO)
        chkFluxo(3).Visible = (tabFluxo.SelectedItem.Key = KEY_ANALITICO)
        chkFluxo(5).Visible = (tabFluxo.SelectedItem.Key = KEY_ANALITICO)
        'Adicionado no protoclo 74160
        txtFluxo(8).Visible = (tabFluxo.SelectedItem.Key = KEY_ANALITICO)
        If Not txtFluxo(8).Visible Then txtFluxo(8).Text = Empty
        lblFluxo(13).Visible = (tabFluxo.SelectedItem.Key = KEY_ANALITICO)
    End If
End Sub

Private Sub txtFluxo_Change(Index As Integer)
  Select Case Index
      Case 2, 3  'Código dos Bancos
            If IsValid(txtFluxo(Index).Text) Then
              GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtFluxo(Index).Text, _
                            lblNome(Index - 2)
            Else
              lblNome(Index - 2).Caption = NUL
            End If
      Case 4, 5  'Código das Contas
            If IsValid(txtFluxo(Index).Text) Then
              GetAssocValue "SELECT Descrição FROM Contas WHERE Código = " & txtFluxo(Index).Text, _
                            lblNome(Index - 2)
            Else
              lblNome(Index - 2).Caption = NUL
            End If
      Case 6     'Centro de Custo
            If (IsValid(txtFluxo(Index).Text)) Then
              GetAssocValue "SELECT Descrição FROM Centros WHERE Código = " & _
                            txtFluxo(Index).Text, lblNome(4)
            Else
              lblNome(4).Caption = NUL
            End If
      Case 7     'Moeda
        GetAssocValue "SELECT Descrição, Moeda FROM Moedas WHERE Moeda = '" & txtFluxo(7).Text & "'", _
                      lblNome(5), txtFluxo(7)
  End Select
  If Index > 1 Then
     chkFluxo(0).Enabled = EnableUpdate()
  End If
End Sub

Private Sub txtFluxo_GotFocus(Index As Integer)
    Selecione txtFluxo(Index)
    FluxoMsgBar txtFluxo(Index).TabIndex
End Sub

Private Sub txtFluxo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If Index > 1 And Index < 4 Then
            PCampo "Bancos", "Bancos", pbCampo, txtFluxo(Index), "Banco"
        ElseIf Index > 3 And Index < 6 Then
            PCampo "Contas", "Contas", pbCampo, txtFluxo(Index), "Código"
        ElseIf Index = 6 Then
            PCampo "Centro de Custo", "Centros", pbCampo, txtFluxo(Index), "Código"
        ElseIf Index = 7 Then
            PCampo "Moedas e Índices", "Moedas", PB_CAMPO, txtFluxo(7), "Moeda"
        End If
    End If
End Sub

Private Sub txtFluxo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 1 'Datas, Inicial e Final
            SetMascara KeyAscii, txtFluxo(Index).SelStart, MASK_DATE4
        Case 2 'Bancos, Inicial e Final
            SetMascara KeyAscii, txtFluxo(2).SelStart, fMask("Bancos", "Banco")
        Case 3
            SetMascara KeyAscii, txtFluxo(3).SelStart, fMask("Bancos", "Banco"), txtFluxo(2).hWnd
        Case 4 'Contas, Inicial e Final
            SetMascara KeyAscii, txtFluxo(4).SelStart, fMask("Contas", "Código")
        Case 5
            SetMascara KeyAscii, txtFluxo(5).SelStart, fMask("Contas", "Código"), txtFluxo(4).hWnd
        Case 6 'Centro de Custo
            SetMascara KeyAscii, txtFluxo(6).SelStart, fMask("Centros", "Código")
        Case 8
            SetMascara KeyAscii, txtFluxo(8).SelStart, "####"
    End Select
End Sub

'FUNCTION..: EnableUpdate
'Objetivo..: Verifica se o CheckBox de atualização de saldos pode ficar
'            habilitado ou não.
'Retorna...: True se o CheckBox puder ficar habilitado, False se não.
'Nota......: Só é possível habilitar esta opção se o usuário quebrar o relatório
'            por bancos, e se não houver filtro por Contas ou Centro de Custo.
Private Function EnableUpdate() As Boolean
    EnableUpdate = ((IsValid(txtFluxo(4).Text) = False And IsValid(txtFluxo(5).Text) = False And _
                    CentrodeCusto(MFinanceiro) = False) Or IsValid(txtFluxo(6).Text) = False)
End Function

'SUB.......: FluxoMsgBar
'Objetivo..: Exibe mensagens de ajuda ao usuário na barra de status do Sistema.
'Argumento.: [intTabIndex]: Propriedade TabIndex do controle que recebe o foco.
Private Sub FluxoMsgBar(intTabIndex As Integer)
    Select Case intTabIndex
        Case 2 'Campos de data
            MsgBar ResolveResString(161, resUM, "de Liberação")
        Case 4
            MsgBar ResolveResString(162, resUM, "de Liberação")
        Case 5 'Caixa de verificação de atualização do Saldo
            MsgBar "Marque para atualizar os Saldos Bancários"
        Case 6 'Caixa de verificação Quebrar Por Bancos
            MsgBar "Quebra o relatório por Banco"
        Case 7 'Caixa de verificação Imprimir Descrição completa
            MsgBar LoadResString(176)
        Case 8 'Caixa de verificação Imprimir Resumo
            MsgBar LoadResString(177)
        Case 9 'Caixa de verificação Imprimir Razão
            MsgBar LoadResString(175)
        Case 12, 15 'Campos de Códigos de Bancos
            MsgBar LoadResString(152) & ResolveResString(75, resUM, "Bancos")
        Case 19, 22 'Campos de Códigos de Contas
            MsgBar LoadResString(164) & ResolveResString(75, resUM, "Contas")
        Case 26 'Campo de Centro de Custo
            MsgBar LoadResString(156) & ResolveResString(75, resUM, "Centro de Custo")
    End Select
End Sub

'SUB.......: CriaFiltroFluxo
'Objetivo..: Cria o filtro para a geração do relatório de fluxo de caixa.
'Argumento.: [pdPraOnde]: Destino da impressão.
Private Sub CriaFiltroFluxo(pdPraOnde As PrintDestinoEnum)
    Dim rstBancos As Object        'Abre o cadastro de Bancos
    Dim rstTemp   As Object        'Tabela temporária
    Dim rstSaldos As Object        'Tabela temporária de Saldos
    Dim strBancos As String
    Dim lngBcoIni As Long          'Código do Banco inicial
    Dim lngBcoFim As Long          'Código do Banco final
    Dim DtData(1) As Date          'Resolve as datas inicial (elemento 0) e final (elemento 1)
    Dim strSe(5)  As String        'Utilizada para as instruções de abertura de cada tabela
    Dim strConta  As String        'Utilizada para adicionar as contas e centro de custo as instruções

    mbolCancelou = False
    dblCotacao = TemCotacao(txtFluxo(7).Text, lblNome(5).Caption, dtInicial, dtFinal)
    If tabFluxo.SelectedItem.Key = KEY_ANALITICO Then
        BQuebraData = (chkFluxo(5).value = vbChecked)
    Else
        BQuebraData = True
    End If
  
    'Verifica se a Moeda Informada é válida antes de executar a Conversão
    If Len(txtFluxo(7).Text) > 0 And lblNome(5).Caption = NUL Then
        MsgBox "Informe uma MOEDA válida para a Conversão de Valores", vbOKOnly Or vbExclamation, MsgBoxCaption
        LetFocus txtFluxo(7).hWnd
        Selecione txtFluxo(7)
        mbolCancelou = True
        Exit Sub
    End If
    'Verifica se a Moeda Informada tem Cotação
    If TemMoeda(txtFluxo(7).Text, lblNome(5).Caption) = True Then
        If dblCotacao = 0 Then
            MsgBox "Informe uma Cotação válida para a Moeda '" & txtFluxo(7).Text & "' na Data de " & txtFluxo(0).Text & " Até " & txtFluxo(1).Text, vbOKOnly Or vbExclamation, MsgBoxCaption
            LetFocus txtFluxo(7).hWnd
            Selecione txtFluxo(7)
            mbolCancelou = True
            Exit Sub
        End If
    End If
  
    'Verifica se existe cotação para da Moeda do Primeiro e do último Dia do Mês Anterior
    If TemMoeda(txtFluxo(7).Text, lblNome(5).Caption) = True Then
        If dblCotacao > 0 Then
            If (Cotacao(txtFluxo(7).Text, LastDay(DateAdd("M", -1, txtFluxo(0).Text))) = 0) Or (Cotacao(txtFluxo(7).Text, FirstDay(DateAdd("M", -1, txtFluxo(0).Text))) = 0) Then
                MsgBox "Informe uma Cotação válida para a Moeda '" & txtFluxo(7).Text & "' na Data de " & FirstDay(DateAdd("M", -1, txtFluxo(0).Text)) & " e/ou " & LastDay(DateAdd("M", -1, txtFluxo(0).Text)) & ".", vbOKOnly Or vbExclamation, MsgBoxCaption
                LetFocus txtFluxo(7).hWnd
                Selecione txtFluxo(7)
                mbolCancelou = True
                Exit Sub
            End If
        End If
    End If

    'Só verifica o campo Previsão quando for Fluxo de Caixa
    If mbitTipo = TIPO_FLUXO Then
        strBancos = "SELECT Banco, Nome FROM Bancos WHERE (Previsão = True OR Banco = 0) AND"
    Else
        strBancos = "SELECT Banco, Nome FROM Bancos WHERE"
    End If
    'Selecionando os bancos: final e inicial
    lngBcoIni = Min(CLngDef(txtFluxo(2).Text), CLngDef(txtFluxo(3).Text))
    lngBcoFim = Max(CLngDef(txtFluxo(2).Text), CLngDef(txtFluxo(3).Text))

    If lngBcoIni > ZERO And lngBcoFim > ZERO Then
        If lngBcoIni = lngBcoFim Then
            DeleteStr strBancos, "(Previsão = True OR Banco = 0) AND"  'Neste caso não devo fazer a comparação com o campo Previsão
            Concat strBancos, " Banco = ", lngBcoIni
        Else
            Concat strBancos, " (Banco BETWEEN ", CStr(lngBcoIni), " AND ", CStr(lngBcoFim), ")"
        End If
    ElseIf lngBcoIni > ZERO Then
        Concat strBancos, " Banco >= ", CStr(lngBcoIni)
    ElseIf lngBcoFim > ZERO Then
        Concat strBancos, " Banco <= ", CStr(lngBcoFim)
    Else
        AppendStr strBancos, " Banco > 0"
    End If
  

    'A data inicial é a mais difícil de definir. Se o usuário não informar esta
    'data o Sistema irá procurar, entre os cadastros de: Lançamentos, Duplicatas,
    'Aplicações e Tranferência Bancária a primeira data registrada. Dependendo do
    'tamanho das tabelas, esta busca pode demorar, mas, o usuário pode passar por
    'esta etapa se ele informar uma data inicial.
    DtData(0) = CDateDef(txtFluxo(0).Text, Empty)
    If (IsEmptyDate(DtData(0))) Then
        If (Not DataInicial(DtData(0))) Then    'Não há registros encontrados
            MsgBox LoadResString(146), vbInformation, MsgBoxCaption
            Exit Sub
        End If
    End If
  'Resolve as datas inicial e final
  If (IsValid(txtFluxo(1).Text)) Then
    If (EData(txtFluxo(1).Text)) Then
      DtData(1) = CDateDef(txtFluxo(1).Text, Date)
    Else
      MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, "data final")
      Exit Sub
    End If
  Else
    DtData(1) = Date
  End If
  
  If (Min(DtData(0), DtData(1)) = DtData(1)) Then
    Dim dtTemp As Date        '// Variável apenas para a troca

    dtTemp = DtData(0)
    DtData(1) = DtData(0)
    DtData(0) = dtTemp
  End If
  '//
  '// Resolvendo as instruções de seleção de dados para cada tabela
  '//
  strSe(0) = "Origem = <Banco> AND Data = " & Quote("<Data>", IIf(gTipoDB = Access, "##", "''")) '// Tabela de Transf Bancária como Banco de Origem
  strSe(1) = "Destino = <Banco> AND Data = " & Quote("<Data>", IIf(gTipoDB = Access, "##", "''"))  '// Tabela de Transf Bancária como Banco de Destino
  strSe(2) = "Banco = <Banco> AND Data = " & Quote("<Data>", IIf(gTipoDB = Access, "##", "''")) & _
             " AND Tipo = '" & _
             GetResOptions(1001, 1) & "'"             '// Tabela de Aplicações com o Tipo Juros/Correção
  strSe(3) = "Banco = <Banco> AND Data = " & Quote("<Data>", IIf(gTipoDB = Access, "##", "''")) & _
             " AND Tipo <> '" & _
             GetResOptions(1001, 1) & "'"             '// Tabela de Aplicações com o Tipo Taxas ou CPMF

  If (mbitTipo = TIPO_EXTRATO Or mbitTipo = TIPO_MOVIMENTO) Then                  '// Quando o tipo for extrato bancário
    If gTipoDB = Access Then
      strSe(4) = "Banco = <Banco> AND (Liberação = " & Quote("<Data>", "##") & _
                 " AND Pagamento IS NOT NULL) AND PagRec = " & _
                 "'R'" '// Tabela de Lançamentos ou Duplicatas Recebidas
      strSe(5) = "Banco = <Banco> AND (Liberação = " & Quote("<Data>", "##") & _
                 " AND Pagamento IS NOT NULL) AND PagRec = " & _
                 "'P'" '// Tabela de Lançamentos ou Duplicatas Pagas
    Else
      strSe(4) = "Banco = <Banco> AND (CONVERT(varchar(10),[Liberação],120) = " & Quote("<Data>", "''") & _
                 " AND (Pagamento IS NOT NULL)) AND PagRec = " & _
                 "'R'" '// Tabela de Lançamentos ou Duplicatas Recebidas
      strSe(5) = "Banco = <Banco> AND (CONVERT(varchar(10),[Liberação],120) = " & Quote("<Data>", "''") & _
                 " AND (Pagamento IS NOT NULL)) AND PagRec = " & _
                 "'P'" '// Tabela de Lançamentos ou Duplicatas Pagas
    End If
  Else                                                  '// Se Tipo for Fluxo de Caixa obtém todos os
                                                     '// dados, mesmo os não pagos
    #If FOXSQL = 1 Then
    strSe(4) = "Banco = <Banco> AND CONVERT(varchar(10),[Liberação],120) = " & Quote("<Data>", "''") & " AND PagRec = 'R'"
    strSe(5) = "Banco = <Banco> AND CONVERT(varchar(10),[Liberação],120) = " & Quote("<Data>", "''") & " AND PagRec = 'P'"
    #Else
    strSe(4) = "Banco = <Banco> AND Liberação = " & Quote("<Data>", "##") & " AND PagRec = 'R'"
    strSe(5) = "Banco = <Banco> AND Liberação = " & Quote("<Data>", "##") & " AND PagRec = 'P'"
    #End If
  End If
  '//
  '// Completa as instruções com código de Conta e Centro de Custo,
  '// se for necessário
  '//
  If ((IsValid(txtFluxo(4).Text)) Or (IsValid(txtFluxo(5).Text))) Then
    If (IsValid(txtFluxo(4).Text) And IsValid(txtFluxo(5).Text)) Then
      If (CompStr(txtFluxo(4).Text, txtFluxo(5).Text)) Then     'Se forem números iguais
        strConta = " AND Conta = " & txtFluxo(4).Text
      Else
        strConta = " AND (Conta BETWEEN " & txtFluxo(4).Text & _
                   " AND " & txtFluxo(5).Text & ")"
      End If
    ElseIf (Not IsValid(txtFluxo(4).Text) And IsValid(txtFluxo(5).Text)) Then
      strConta = " AND Conta <= " & txtFluxo(5).Text
    ElseIf (IsValid(txtFluxo(4).Text) And (Not IsValid(txtFluxo(5).Text))) Then
      strConta = " AND Conta >= " & txtFluxo(4).Text
    End If

    AppendStr strSe(0), strConta
    AppendStr strSe(1), strConta
    AppendStr strSe(2), strConta
    AppendStr strSe(3), strConta
    AppendStr strSe(4), strConta
    AppendStr strSe(5), strConta
    
  End If
  '// Centro de Custo
  '
  Dim strCCusto As String
  
  strCCusto = ""
  If ((txtFluxo(6).Visible) And (IsValid(txtFluxo(6).Text))) Then
    strConta = " AND Centro = " & txtFluxo(6).Text

    AppendStr strSe(0), strConta
    AppendStr strSe(1), strConta
    AppendStr strSe(2), strConta
    AppendStr strSe(3), strConta
    AppendStr strSe(4), strConta
    AppendStr strSe(5), strConta
    
    strCCusto = "Centro de Custo: " & txtFluxo(6).Text
    
  End If
  
  Dim sNomeTab As String
  
  
  If (AbreRecordset(rstBancos, strBancos, dbOpenSnapshot) = WL_OK) Then
  
    
    If (CrieTemp(rstTemp) And TempSaldos(rstSaldos)) Then       '// Se criar as tabelas temporárias
      Dim strTitulo    As String
      
      Select Case tabFluxo.SelectedItem.Key
      
      Case KEY_SINTETICO
        
          strTitulo = Me.Caption & " Sintético de " & DataToStr(DtData(0)) & " até " & DataToStr(DtData(1)) & IIf(Len(txtFluxo(7).Text) > 0, " (Moeda Base: " & txtFluxo(7).Text & ")", "") & " / " & strCCusto
          If (chkFluxo(4).value = vbChecked) Then                 '// Quebra por Bancos
            If (AddSinteticoA(rstBancos, rstTemp, rstSaldos, DtData(), strSe())) Then
              sNomeTab = GetTableSource(rstSaldos)
              fimpFluxoSintetico.Config rstTemp, strTitulo, chkFluxo(4).value, sNomeTab
            End If
          Else                                                    '// Sem Quebra
            sNomeTab = GetTableSource(rstSaldos)
            If (AddSinteticoB(rstBancos, rstTemp, rstSaldos, DtData(), strSe())) Then
              fimpFluxoSintetico.Config rstTemp, strTitulo, chkFluxo(4).value, sNomeTab
            End If
          End If
      
      Case "analitico"    '// Relatório Analítico
        sNomeTab = GetTableSource(rstSaldos)
        strTitulo = Me.Caption & " Analítico de " & DataToStr(DtData(0)) & " até " & DataToStr(DtData(1)) & IIf(Len(txtFluxo(7).Text) > 0, " (Moeda Base: " & txtFluxo(7).Text & ")", "") & " / " & strCCusto
        
        If (chkFluxo(4).value = vbChecked) Then                 '// Quebra por Bancos
          If (AddAnaliticoA(rstBancos, rstTemp, rstSaldos, DtData(), strSe())) Then
            
            If mbitTipo = TIPO_MOVIMENTO Then
              fimpFluxoMovimento.Config rstTemp, strTitulo, chkFluxo(4).value, sNomeTab
            Else
              fimpFluxoAnalitico.Config rstTemp, strTitulo, chkFluxo(5).value, chkFluxo(4).value, sNomeTab, chkFluxo(2).value, chkFluxo(1).value, chkFluxo(3).value, CIntDef(txtFluxo(8).Text, 0)
            End If
          End If
        Else                                                    '// Sem quebra
          If (AddAnaliticoB(rstBancos, rstTemp, rstSaldos, DtData(), strSe())) Then
            If mbitTipo = TIPO_MOVIMENTO Then
              fimpFluxoMovimento.Config rstTemp, strTitulo, chkFluxo(4).value, sNomeTab
            Else
              fimpFluxoAnalitico.Config rstTemp, strTitulo, chkFluxo(5).value, chkFluxo(4).value, sNomeTab, chkFluxo(2).value, chkFluxo(1).value, chkFluxo(3).value, CIntDef(txtFluxo(8).Text, 0)
            End If
          End If
        End If
      
      Case KEY_SINTETICO_CONTA
    
        sNomeTab = GetTableSource(rstSaldos)
        
        strTitulo = Me.Caption & " Sintético por Conta de " & DataToStr(DtData(0)) & " até " & DataToStr(DtData(1)) & IIf(Len(txtFluxo(7).Text) > 0, " (Moeda Base: " & txtFluxo(7).Text & ")", "") & " / " & strCCusto
        Dim rstSintConta As Object
        
        If (chkFluxo(4).value = vbChecked) Then                 '// Quebra por Bancos
          If (AddAnaliticoA(rstBancos, rstTemp, rstSaldos, DtData(), strSe())) Then
            Sleep 3000
            AbreRecordset rstSintConta, "SELECT Banco, Nome, Data, Conta, (SELECT T1.Descrição FROM Contas AS T1 WHERE T1.Código = Conta) AS DescConta, SUM(Saída) AS CSaída, SUM(Entrada) AS CEntrada " & _
                                        "FROM " & NomeTabeladoRST(rstTemp) & ESP & _
                                        "GROUP BY Banco, Nome, Data, Conta " & _
                                        "ORDER BY Banco, Data", dbOpenSnapshot
                                        

            fimpFluxoSinteticoConta.Config rstSintConta, strTitulo, chkFluxo(5).value, chkFluxo(4).value, sNomeTab, chkFluxo(2).value, chkFluxo(1).value, chkFluxo(3).value
            FechaRecordset rstSintConta
          End If
        Else                                                    '// Sem quebra
          If (AddAnaliticoB(rstBancos, rstTemp, rstSaldos, DtData(), strSe())) Then
            AbreRecordset rstSintConta, "SELECT Data, Conta, (SELECT T1.Descrição FROM Contas AS T1 WHERE T1.Código = Conta) AS DescConta, SUM(Saída) AS CSaída, SUM(Entrada) AS CEntrada " & _
                                        "FROM " & NomeTabeladoRST(rstTemp) & ESP & _
                                        "GROUP BY Data, Conta " & _
                                        "ORDER BY Data", dbOpenSnapshot
            fimpFluxoSinteticoConta.Config rstSintConta, strTitulo, chkFluxo(5).value, chkFluxo(4).value, sNomeTab, chkFluxo(2).value, chkFluxo(1).value, chkFluxo(3).value
            FechaRecordset rstSintConta
          End If
        End If
      End Select
      
      DeleteAux rstTemp, NUL
      DeleteAux rstSaldos, NUL
    
    Else
      MsgFunc LoadResString(174) ' Criatemp
    End If
  
  Else
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption ' Abrerecordset
  End If
  
  FechaRecordset rstBancos
  MsgBar LoadResString(116)
  
End Sub

' SUB.......: UpdateSaldoBanco
' Objetivo..: Atualiza o arquivo de Saldos Bancários conforme for necessário.
' Argumentos: [strBanco]: Código do Banco que deve ser atualizado.
'             [datData ]: Data para a atualização do Saldo.
'             [curSaldo]: Valor do saldo que deve ser gravado.
' -----------------------------------------------------------------------------------
Private Sub UpdateSaldoBanco(strBanco As String, datData As Date, curSaldo As Currency, Optional strMoeda As String, Optional dblCotacao)
    Dim strUpdate As String
  
    'Verifica se a caixa está marcada e visível.
    'pt. 88218 - Ivo Sousa (23/09/2008)
    If ((chkFluxo(0).value = vbChecked) And (chkFluxo(0).Enabled)) Then 'And (chkFluxo(0).Visible))
        'Somente se for o último dia do mês
        If (DateDiff("d", datData, LastDay(datData)) = ZERO) Then
            'Somente se não for uma data futura
            If (DateDiff("d", datData, Date) >= ZERO) Then
                If (Recordcount("SELECT Banco FROM [Saldos Bancários] WHERE Banco = " & strBanco & " AND Período = " & InverteData(FirstDay(datData), True) & ";")) Then
                    ' Se o registro já existir...
                    strUpdate = "UPDATE [Saldos Bancários] SET Valor = " & ValStr(curSaldo) & _
                                " WHERE Banco = " & strBanco & " AND Período = " & _
                                InverteData(FirstDay(datData), True) & ";"
                Else
                    strUpdate = "INSERT INTO [Saldos Bancários] (Banco, Período, Valor)" & _
                                " VALUES (" & strBanco & ", " & _
                                InverteData(FirstDay(datData), True) & ", " & ValStr(curSaldo) & ");"
                End If
                ExecuteSQL strUpdate
            End If
        End If
    End If
End Sub

' FUNCTION..: CrieTemp
' Objetivo..: Cria a tabela temporário onde os dados do relatório
'             serão gravados para serem impressos.
' Argumento.: [rstTemp]: Recordset que receberá a tabela temporária
' Retorna...: True se puder criar a tabela com sucesso, False se não.
' --------------------------------------------------------------------
Private Function CrieTemp(rstTemp As Object) As Boolean
Dim fsFluxo(17) As FieldStruct

  AppendVar fsFluxo(0), "Banco", dbLong           'Código do Banco
  AppendVar fsFluxo(1), "Nome", dbText, 40        'Nome do Banco
  AppendVar fsFluxo(2), "Mes", dbText, 9          'Mês (servirá como agrupador)
  AppendVar fsFluxo(3), "Data", dbDate            'Data do movimento
  AppendVar fsFluxo(4), "Empresa", dbText, 15     'Nome da empresa
  #If FOXSQL = 1 Then
  AppendVar fsFluxo(5), "Duplicata", dbFloat      'Código da Duplicata, Lançamento, Transferência ou Aplicação
  #Else
  AppendVar fsFluxo(5), "Duplicata", dbDouble      'Código da Duplicata, Lançamento, Transferência ou Aplicação
  #End If
  'Protocolo 73121 ---|
  AppendVar fsFluxo(6), "Parcela", dbText, 4      'Código da Duplicata, Lançamento, Transferência ou Aplicação
  '-------------------|
  AppendVar fsFluxo(7), "Tipo", dbText, 30        'Tipo da Duplicata ou Lançamento, Transferência ou Aplicação
  AppendVar fsFluxo(8), "Type", dbByte            'Utilizado para reconhecer quando o tipo é Aplicação, Transferência ou outro
  AppendVar fsFluxo(9), "Descrição", dbText, 80   'Descrição do movimento
  AppendVar fsFluxo(10), "Controle", dbText, 80   'Controle do movimento
  AppendVar fsFluxo(11), "Conta", dbLong          'Código da Conta
  AppendVar fsFluxo(12), "Cheque", dbLong         'Número do Cheque
  AppendVar fsFluxo(13), "Vencimento", dbDate     'Data do Vencimento
  AppendVar fsFluxo(14), "Pagamento", dbDate      'Data do Pagamento
  AppendVar fsFluxo(15), "Saída", dbCurrency      'Valor de saída (no caso de ser saída)
  AppendVar fsFluxo(16), "Entrada", dbCurrency    'Valor de entrada (no caso de ser entrada)
  AppendVar fsFluxo(17), "Saldo", dbCurrency      'Saldo após o movimento

  CrieTemp = CrieAux(rstTemp, fsFluxo())
  
End Function

' FUNCTION..: TempSaldos
' Objetivo..: Cria uma tabela auxiliar contendo os saldos Inicial e Final de
'             cada banco para impressão.
' Argumento.: [rstSaldos]: Recordset que receberá a tabela auxiliar.
' Retorna...: True se puder criar a tabela com sucesso False se não.
' -----------------------------------------------------------------------------
Private Function TempSaldos(rstSaldos As Object) As Boolean
Dim fsSaldos(3) As FieldStruct

  AppendVar fsSaldos(0), "Banco", dbLong          'Contém o código do Banco
  AppendVar fsSaldos(1), "Data", dbDate           'Contém a data do Saldo
  AppendVar fsSaldos(2), "Tipo", dbBoolean        'True para Saldo Final, False para Saldo Inicial
  AppendVar fsSaldos(3), "Valor", dbCurrency      'Valor do Saldo

  TempSaldos = CrieAux(rstSaldos, fsSaldos())
  
End Function

' FUNCTION..: AddSinteticoA
' Objetivo..: Grava os dados do relatório de Fluxo de Caixa e Extrato Bancário Sintético
'             na tabela auxiliar para impressão, quando deve quebrar por banco.
' Argumentos: [rstBcos]: Recordset com os bancos escolhidos pelo usuário.
'             [rstData]: Recordset que receberá os dados para o relatório.
'             [rstSld ]: Recordset com os saldos iniciais e finais de cada banco.
'             [dtData ]: Matriz com as datas inicial e final.
'             [strExp ]: Matriz com as strings de seleção para cada tabela.
' Retorno...: True se a função obtiver sucesso, False se não.
' ---------------------------------------------------------------------------------------
Private Function AddSinteticoA(rstBcos As Object, rstData As Object, rstSld As Object, DtData() As Date, strExp() As String) As Boolean
    Dim cTotal    As Currency
    Dim DtDia     As Date                 '// Data do dia da busca de dados
    Dim strBanco  As String               '// Banco atual da busca
    Dim MaisUm    As Boolean              '// Controle para banco zerado
    
    MaisUm = IsValid(txtFluxo(2).Text)
  
On Error GoTo AddSintetico_Erro

    'Move para o primeiro registro
    rstBcos.MoveFirst
    Do
        If (mbolCancelou) Then
            GoTo AddSintetico_Erro
        End If
        'Habilita ao usuário o cancelamento do cálculo
        DoEvents
        DtDia = DtData(0)
        strBanco = GetValue(rstBcos, "Banco", ZERO)
        'Protocolo 76585
        'Adicionado parametro TipoReg
        cTotal = SaldoInicial(CLng(strBanco), DtDia, strMoeda:=txtFluxo(7).Text, StrDescMoeda:=lblNome(5).Caption, sConciliado:=cboFluxo.Text, TipoRel:=mbitTipo, bConsiderarAtrasados:=chkFluxo(10).value)              '// Saldo inicial deste Banco
        
        If (TemMoeda(txtFluxo(7).Text, lblNome(5).Caption)) Then
            dblCotacao = UltimaCotacao(txtFluxo(7).Text, DtDia)
            If dblCotacao > 0 Then
                cTotal = cTotal / dblCotacao  'Saldo em Reais dividido pela cotacao da moeda base
            Else
                MsgFunc "Não há cotação para a data: " & DtDia
                cTotal = 0
            End If
        End If
        
        rstSld.AddNew                   '// Grava o Saldo inicial na tabela auxiliar dos saldos
        rstSld("Banco").value = CLng(strBanco)
        rstSld("Data").value = DtData(0)
        rstSld("Tipo").value = False    '// Tipo = False é usado para saldo inicial
        rstSld("Valor").value = cTotal
        rstSld.update
                
        'Adiciona um registro em branco por Banco
        'Para que os saldos sejam apresentados
        'Ainda que não exista movimentação no período especificado
        If chkImprimeBancoSemMovimento.value = vbChecked Then
            rstData.AddNew
            Dim i As Integer
            For i = 0 To rstData.Fields.Count - 1
                'Select Case TransDBTypeRetInt(rstData(i).Type)
                Select Case rstData(i).Type
                    Case dbText
                        ' by kleber 2305
                        ' Ajuste para funcionar com SQL
                        If gTipoDB = Access Then
                            rstData(i).value = CStr(DefaultValue(rstData(i).SourceTable, rstData.Fields(i)))
                        Else
                            rstData(i).value = CStr(DefaultValue(rstData.Source, rstData.Fields(i)))
                        End If
                    Case dbDate
                        If TypeOf rstData Is dao.Recordset Then
                            rstData(i).value = CDateDef(DefaultValue(rstData(i).SourceTable, rstData.Fields(i)))
                        Else
                            rstData(i).value = CDateDef(DefaultValue(rstData.Source, rstData.Fields(i)))
                        End If
                    Case dbInteger, dbLong, dbByte
                        If TypeOf rstData Is dao.Recordset Then
                            rstData(i).value = CLngDef(DefaultValue(rstData(i).SourceTable, rstData.Fields(i)))
                        Else
                            rstData(i).value = CLngDef(DefaultValue(rstData.Source, rstData.Fields(i)))
                        End If
                        'Empresa deve possuir algum conteúdo ou o gerador acusará erro(DataLink)
                        If rstData(i).name = "Empresa" Then
                            rstData(i).value = Space(1)
                        End If
                End Select
            Next i
            rstData("Banco").value = strBanco
            rstData("Nome").value = GetValue(rstBcos, "Nome", NUL)
            rstData.update
        End If
        
        
        Do Until (DateDiff("d", DtDia, DtData(1)) < ZERO)   '// Até que dtDia seja posterior a dtData(1)
        
            If (mbolCancelou) Then
                GoTo AddSintetico_Erro
            End If
            'Habilita ao usuário cancelar
            DoEvents
            
            SimpleMsgBar ResolveResString(1023, resUM, strBanco, resDOIS, GetValue(rstBcos, "Nome", NUL), resTRES, DataToStr(DtDia))
                                          
            If (Not GravaSintetico(rstData, strBanco, GetValue(rstBcos, "Nome", NUL), strExp(), DtDia, cTotal)) Then
                MsgFunc ResolveResString(247, resUM, Me.Caption)
                GoTo AddSintetico_Erro
            End If
            
            'Chama a função UpdateSaldoBanco que atualizará a tabela de Saldos Bancários
            'se for necessário.
            If TemMoeda(txtFluxo(7).Text, lblNome(5).Caption) = False Then
                UpdateSaldoBanco strBanco, DtDia, cTotal
            End If
            'Soma um dia a data atual
            DtDia = DateAdd("d", 1, DtDia)
        Loop
        If Not rstBcos.EOF Then
            rstBcos.MoveNext
        Else
            MaisUm = True
        End If
    Loop Until (rstBcos.EOF) And (MaisUm)
  
    If (EstaVazio(rstData)) Then
        MsgFunc LoadResString(IDS_NORECORD)
        AddSinteticoA = False
    Else
        AddSinteticoA = True
    End If
    Exit Function
  
AddSintetico_Erro:
    If (err.Number) Then
        If Not rstSld.EOF Or Not rstSld.BOF Then
            If (rstSld.EditMode <> dbEditNone) Then
                rstSld.CancelUpdate
            End If
        End If
    End If
    AddSinteticoA = False
    Resume
End Function

' FUNCTION..: AddSinteticoB
' Objetivo..: Grava os dados do relatório de Fluxo de Caixa e Extrato Bancário Sintético
'             na tabela auxiliar para impressão, quando não há quebras.
' Argumentos: [rstBcos]: Recordset com os bancos escolhidos pelo usuário.
'             [rstData]: Recordset que receberá os dados para o relatório.
'             [rstSld ]: Recordset com os saldos iniciais e finais de cada banco.
'             [dtData ]: Matriz com as datas inicial e final.
'             [strExp ]: Matriz com as strings de seleção para cada tabela.
' Retorna...: True se a função obtiver sucesso, False se não.
' ---------------------------------------------------------------------------------------
Private Function AddSinteticoB(rstBcos As Object, rstData As Object, rstSld As Object, DtData() As Date, strExp() As String) As Boolean
Dim cTotal    As Currency
Dim DtDia     As Date               '// Data do dia da busca de dados
Dim strBanco  As String             '// Banco atual da busca
Dim MaisUm    As Boolean            '// Controle para banco zerado

  MaisUm = IsValid(txtFluxo(2).Text)
  
  On Error GoTo AddSintetico_Erro
  '//
  '// Primeiro soma o saldo inicial de todos os bancos para o relatório
  '//
  rstBcos.MoveFirst
  Do
    'Protocolo 76585
    'Adicionado parametro TipoReg
    cTotal = cTotal + Round(SaldoInicial(GetValue(rstBcos, "Banco", ZERO), DtData(0), strMoeda:=txtFluxo(7).Text, StrDescMoeda:=lblNome(5).Caption, sConciliado:=cboFluxo.Text, TipoRel:=mbitTipo, bConsiderarAtrasados:=chkFluxo(10).value), 2)
    
    If (TemMoeda(txtFluxo(7).Text, lblNome(5).Caption)) Then
       dblCotacao = UltimaCotacao(txtFluxo(7).Text, DtData(0))
       If dblCotacao > 0 Then
          cTotal = Round((cTotal / dblCotacao), 2) 'Saldo em Reais dividido pela cotacao da moeda base
       Else
          MsgFunc "Não há cotação para a data: " & DtData(0)
          cTotal = 0
       End If
    End If
    
    
    If Not rstBcos.EOF Then
      rstBcos.MoveNext
    Else
      MaisUm = True
    End If
  
  Loop Until (rstBcos.EOF) And (MaisUm)
  
  rstSld.AddNew                              '// Salva o saldo inicial na tabela auxiliar de saldos
  rstSld("Banco").value = ZERO               '// Banco zero porque não há quebra
  rstSld("Data").value = DtData(0)           '// Data inicial do cálculo
  rstSld("Tipo").value = False               '// O Tipo Falso é usado para saldo inicial
  rstSld("Valor").value = Round(cTotal, 2)   '// Saldo calculado
  rstSld.update
  
  DtDia = DtData(0)
  Do Until (DateDiff("d", DtDia, DtData(1)) < ZERO)   '// Até que dtDia seja posterior a dtData(1)
    
    MaisUm = IsValid(txtFluxo(2).Text)
    
    rstBcos.MoveFirst               '// Move para o primeiro registro, novamente
    If (mbolCancelou) Then GoTo AddSintetico_Erro
    DoEvents                        '// Habilita ao usuário o cancelamento do cálculo
    
    Do

      If (mbolCancelou) Then GoTo AddSintetico_Erro
      DoEvents                      '// Habilita ao usuário cancelar
      
      strBanco = GetValue(rstBcos, "Banco", ZERO)
      SimpleMsgBar ResolveResString(1023, _
                                    resUM, strBanco, _
                                    resDOIS, GetValue(rstBcos, "Nome", NUL), _
                                    resTRES, DataToStr(DtDia))
                                    
      If (Not GravaSintetico(rstData, strBanco, _
                             GetValue(rstBcos, "Nome", NUL), _
                             strExp(), DtDia, cTotal)) Then
        MsgFunc ResolveResString(247, resUM, Me.Caption)
        GoTo AddSintetico_Erro
      End If
      
      If Not rstBcos.EOF Then
        rstBcos.MoveNext
      Else
        MaisUm = True
      End If
      
    Loop Until (rstBcos.EOF) And (MaisUm)
    
     'pt. 78772 - Dulcino Júnior
     'Atualização do saldo bancário
     If txtFluxo(2).Text = txtFluxo(3).Text Then
        If TemMoeda(txtFluxo(7).Text, lblNome(5).Caption) = False Then
           UpdateSaldoBanco strBanco, DtDia, cTotal
         End If
     End If
    
    DtDia = DateAdd("d", 1, DtDia)        '// Acrescenta um dia a dtDia
    
  Loop
  
  If (EstaVazio(rstData)) Then
    MsgFunc LoadResString(IDS_NORECORD)
    AddSinteticoB = False
  Else
    AddSinteticoB = True
  End If
  Exit Function
  
AddSintetico_Erro:

  If (err.Number) Then
    If (rstSld.EditMode <> dbEditNone) Then rstSld.CancelUpdate
    DAOErros erro(err.Number) & " AddSinteticoB"
  End If
  AddSinteticoB = False
  
End Function

' FUNCTION..: GravaSintetico
' Objetivo..: Obtém os valores das tabelas de movimentação bancária e
'             grava estes valores na tabela auxiliar.
' Argumentos: [rstAux ]: Recordset da tabela auxiliar.
'             [strBco ]: Código do Banco.
'             [strNome]: Nome do Banco.
'             [strExp ]: Matriz com as instruções já montadas.
'             [dtData ]: Dia para o cálculo.
'             [cSaldo ]: Saldo inicial.
' Retorna...: True se gravar os dados corretamente, False se não.
'             O argumento cSaldo retornará com o Saldo atualizado.
' -------------------------------------------------------------------------
Private Function GravaSintetico(rstAux As Object, strBco As String, strNome As String, strExp() As String, DtData As Date, cSaldo As Currency) As Boolean
Dim cCredito As Currency
Dim cDebito  As Currency
Dim strWhere As String
Dim strConciliado As String
  
  Select Case cboFluxo.Text
    Case "Todos"
       strConciliado = ""
    Case "Sim"
       strConciliado = " AND Conciliado = TRUE "
    Case "Não"
       strConciliado = " AND Conciliado = FALSE "
  End Select

  cCredito = ZERO: cDebito = ZERO
  '//
  '// Criando a consulta para a tabela de Transf Bancária
  '// como banco de origem. Esta instrução está no elemento ZERO
  '// da matriz strExp.
  '//
  If cboFluxo.Text <> "Não" Then
    strWhere = strExp(0)
    MidStr strWhere, "<Banco>", strBco
    MidStr strWhere, "<Data>", InverteData(DtData)

    cDebito = Round(Soma("Valor", "[Transf Bancária]", strWhere, ZERO) / UltimaCotacao(txtFluxo(7).Text, DtData), 2)
  End If
  '//
  '// Criando a Consulta para a tabela de Transf Bancária
  '// como banco de destino. Esta instrução está no elemento 1
  '// da matriz strExp.
  '//
  If cboFluxo.Text <> "Não" Then
     strWhere = strExp(1)
     MidStr strWhere, "<Banco>", strBco
     MidStr strWhere, "<Data>", InverteData(DtData)
   
     cCredito = Round(Soma("Valor", "[Transf Bancária]", strWhere, ZERO) / UltimaCotacao(txtFluxo(7).Text, DtData), 2)
  End If
  '//
  '// Criando a consulta para a tabela de Aplicações com o Tipo
  '// Juros/Correção. Esta instrução está no elemento 2 da
  '// matriz strExp.
  '//
  If cboFluxo.Text <> "Não" Then
      strWhere = strExp(2)
      MidStr strWhere, "<Banco>", strBco
      MidStr strWhere, "<Data>", InverteData(DtData)
    
      cCredito = cCredito + Round((Soma("Valor", "Aplicações", strWhere, ZERO) / UltimaCotacao(txtFluxo(7).Text, DtData)), 2)
  End If
  '//
  '// Criando a consulta para a tabela de Aplicações com o Tipo
  '// Taxas ou CPMF. Esta instrução está no elemento 3 da matriz strExp
  '//
  If cboFluxo.Text <> "Não" Then
    strWhere = strExp(3)
    MidStr strWhere, "<Banco>", strBco
    MidStr strWhere, "<Data>", InverteData(DtData)

      cDebito = cDebito + Round(Soma("Valor", "Aplicações", strWhere, ZERO) / UltimaCotacao(txtFluxo(7).Text, DtData), 2)
  End If
  '//
  '// Criando a consulta para a tabela de Duplicatas e Lançamentos Recebidos
  '// Esta instrução está no elemento 4 da matriz strExp.
  '//

  
  strWhere = strExp(4)
  MidStr strWhere, "<Banco>", strBco
  MidStr strWhere, "<Data>", InverteData(DtData)
    
  Concat strWhere, ESP, strConciliado
  
  'Protocolo 73606
  Concat strWhere, ESP, "AND Situação <> 'Cancelada'"
  
  cCredito = cCredito + Round(SomarMoedas("Duplicatas", strWhere, txtFluxo(7).Text), 2)
  cCredito = cCredito + Round(SomarMoedas("Lançamentos", strWhere, txtFluxo(7).Text), 2)
  
  'Considerar pedidos pendentes
  Dim strWhere2 As String
  
  strWhere2 = Replace(strWhere, "AND Situação <> 'Cancelada'", "", , , vbTextCompare)
  strWhere2 = Replace(strWhere2, "Liberação", "Vencimento", , , vbTextCompare)
  
  If chkFluxo(6).value = vbChecked Or _
       chkFluxo(7).value = vbChecked Or _
       chkFluxo(8).value = vbChecked Or _
       chkFluxo(9).value = vbChecked Then
    cCredito = cCredito + Round(Soma("Valor", NomeTabeladoRST(rstPrevisao), strWhere2, ZERO) / UltimaCotacao(txtFluxo(7).Text, DtData), 2)
  End If
  '//
  '// Criando a consulta para a tabela de Duplicatas e Lançamentos Pagos
  '// Esta instrução está no elemento 5 da matriz strExp.
  '//
  strWhere = strExp(5)
  MidStr strWhere, "<Banco>", strBco
  MidStr strWhere, "<Data>", InverteData(DtData)
  
  Concat strWhere, ESP, strConciliado
      
  'Protocolo 73606
  Concat strWhere, ESP, "AND Situação <> 'Cancelada'"
  
  cDebito = cDebito + Round(SomarMoedas("Duplicatas", strWhere, txtFluxo(7).Text), 2)
  cDebito = cDebito + Round(SomarMoedas("Lançamentos", strWhere, txtFluxo(7).Text), 2)
                           
  'Considerar pedidos pendentes
  If chkFluxo(6).value = vbChecked Or _
       chkFluxo(7).value = vbChecked Or _
       chkFluxo(8).value = vbChecked Or _
       chkFluxo(9).value = vbChecked Then
    'Pt. 102516 - Moacir Pfau(09/11/2010)
    cDebito = cDebito + Round(Soma("Valor", NomeTabeladoRST(rstPrevisao), Replace(Replace(strWhere2, "Liberação", "Vencimento", , , vbTextCompare), "PagRec = 'R'", "PagRec = 'P'"), ZERO) / UltimaCotacao(txtFluxo(7).Text, DtData), 2)
  End If

  'Protocolo 77590 --------------------------------------
  cSaldo = cSaldo + (cCredito - cDebito) '// Total do dia
  '-------------------------------------------------------
  '//
  '// Grava no arquivo auxiliar apenas os dias em que há movimento
  '//
  On Error GoTo GravaSintetico_Erro
  
  If ((cDebito > ZERO) Or (cCredito > ZERO)) Then
    rstAux.AddNew
    rstAux("Banco").value = CLng(strBco)
    rstAux("Nome").value = strNome
    rstAux("Data").value = DtData
    rstAux("Mes").value = MesExt(DtData, 9)
    'Verifica se existe Cotação e Se existe Moeda Caso não satisfaça a Cotação não emite o relatório
    rstAux("Saída").value = -Round(cDebito, 2)
    rstAux("Entrada").value = Round(cCredito, 2)
    'Protoclo 77590 ------------------------
    'A função Round alterava sinal de cSaldo ByRef
    'foi modificada para manter sinal (adotado ByVal)
    rstAux("Saldo").value = Round(cSaldo, 2)
    '---------------------------------------
    rstAux.update
  End If

  GravaSintetico = True
  
GravaSintetico_Erro:
  If (err.Number) Then
    If (rstAux.EditMode <> dbEditNone) Then rstAux.CancelUpdate
    DAOErros erro(err.Number) & " GravaSintetico"
    GravaSintetico = False
  End If
End Function

' SUB.......: RelatorioSintetico
' Objetivo..: Cria a extrutura do relatório de Fluxo de Caixa ou Extrato
'             Bancário Sintético, quebrando por bancos ou não.
' Argumentos: [pdDestino]: Destino da impressão.
'             [rstDados ]: Recordset com os dados de origem.
'             [dtDatas  ]: Matriz com as datas de geração.
'             [strSaldos]: Nome da Tabela que contém os saldos dos bancos.
' -------------------------------------------------------------------------
Private Sub RelatorioSintetico(pdDestino As Long, rstDados As Object, dtDatas() As Date, strSaldos As String)
Dim wrkSintetico As KeybReport    '// Variável do relatório
Dim strDataLink  As String
Dim bQuebra      As Boolean       '// Define se devo quebrar por bancos ou não

  SimpleMsgBar LoadResString(160)
  bQuebra = (chkFluxo(4).value = vbChecked)
  
  Set wrkSintetico = New KeybReport
  With wrkSintetico
    Set .DatabaseName = GlobalDataBase
    Set .Recordset = rstDados
    .AutoRedraw = True
    .Tipo = wrObjectDraw
    .ScaleMode = vbMillimeters
    .WindowTitulo = Me.Caption & " Sintético"
    .Destino = pdDestino

    PageHeader wrkSintetico, Me.Caption & " Sintético de " & _
                             DataToStr(dtDatas(0)) & " até " & _
                             DataToStr(dtDatas(1))
    
    If Len(txtFluxo(7).Text) > 0 Then
      .UltimaSecao.AddLinha "Moeda"
      .UltimaSecao(.UltimaSecao.LinhasCount).AddCampo , wrCSFixedText, "Valores Demonstrados em: " & txtFluxo(7).Text, wrTACentro
    End If
    
    .FontSize = 8
    .FontStyle = wrFSBold
    .AddGrupo "1", wrDBBottomBorder
    '//
    '// Se devo quebrar por Banco a propriedade "Quebra" do Grupo 1
    '// recebe o valor Banco que é o campo de código de Bancos, caso
    '// contrário não há quebra
    '//
    If (bQuebra) Then
      .Grupo(1).Quebra = "Banco"
    End If
    .Grupo(1).AddSecao scHeader, 3

    With .Grupo(1).Header.Linha(2)
      .DrawBorder = wrDBTopBorder
      If (bQuebra) Then
        .AddCampo , wrCSFixedText, "Banco:", , 15
        .AddCampo , , "Banco", wrTADireito, 17
        .Campo(2).Formato = StrZero(0, 9)
        .AddCampo , , "Nome", , 40
      End If
      .AddCampo , wrCSFixedText, "Saldo Anterior:", , 30, 114
      .AddCampo "saldo", wrCSDataLink, "Valor", wrTADireito
      .Campo("saldo").Formato = FMOEDA
      .Campo("saldo").TableLink = strSaldos

      If (bQuebra) Then
        .Campo("saldo").DataLink = "Banco = {*Banco} AND Data = " & _
                                   InverteData(dtDatas(0), True) & " AND Tipo = False"
      Else
        .Campo("saldo").DataLink = "Banco = 0"
      End If
    End With
    
    If (bQuebra) Then             '// Quando quebra por Banco
      With .Grupo(1).Header.Linha(3)
        .DrawBorder = wrDBTopBorder
        .BorderStyle = wrDot
        .AddCampo , wrCSFixedText, "Data", , 20, 16
        .AddCampo , wrCSFixedText, "Total de Entradas", wrTADireito, 30
        .AddCampo , wrCSFixedText, "Total de Saídas", wrTADireito, 30
        .AddCampo , wrCSFixedText, "Saldo Final", wrTADireito, 30
      End With
    Else                          '// Sem a quebra
      With .Grupo(1).Header.Linha(3)
        .DrawBorder = wrDBTopBorder
        .BorderStyle = wrDot
        .AddCampo , wrCSFixedText, "Banco", wrTADireito, 17
        .AddCampo , wrCSFixedText, "Nome", , 50
        .AddCampo , wrCSFixedText, "Data", , 20
        .AddCampo , wrCSFixedText, "Total de Entradas", wrTADireito, 30
        .AddCampo , wrCSFixedText, "Total de Saídas", wrTADireito, 30
        .AddCampo , wrCSFixedText, "Saldo Final", wrTADireito, 30
      End With
    End If
    .FontStyle = wrFSNormal
    '//
    '// Seção de Detalhes do Grupo.
    '//
    .Grupo(1).AddSecao scDetalhe, 1
    With .Grupo(1).Detalhe.Linha(1)
      If (Not bQuebra) Then
        .AddCampo , , "Banco", wrTADireito, 17
        .AddCampo , , "Nome", , 50
      End If
      .AddCampo "data", , "Data", , 20, IIf((bQuebra), 16, 0)
      .AddCampo "entr", , "Entrada", wrTADireito, 30
      .AddCampo "said", , "Saída", wrTADireito, 30
      .AddCampo "sald", , "Saldo", wrTADireito, 30
      .Campo("data").Formato = FDATA
      .Campo("entr").Formato = FMOEDA
      .Campo("said").Formato = FMOEDA
      .Campo("sald").Formato = FMOEDA
      .Campo("sald").SuprimirZeros = True
      .Campo("entr").SuprimirZeros = True
      .Campo("said").SuprimirZeros = True
      .Campo("sald").SuprimirZeros = True
    End With

'
'Seção de Total ((SOMA DE ENTRADAS - SOMA DAS SAÍDAS) + SOMA DO SALDO INICIAL))
'

    .FontSize = 8
    .FontStyle = wrFSBold
    .AddGrupo 2
    .Grupo(2).AddSecao scFooter, 2, wrDBNoBorders
    
    With wrkSintetico.Grupo(2).Footer(2)
      .BorderStyle = wrDot
      .DrawBorder = wrDBAllBorders
        
        .AddCampo , wrCSFixedText, "TOTAL GERAL:", wrTADireito, 33
        
        If chkFluxo(4).value = vbChecked Then
          .AddCampo "TotalEntrada", wrCSTotal, "Entrada", wrTADireito, 33
        Else
          .AddCampo "TotalEntrada", wrCSTotal, "Entrada", wrTADireito, 85
        End If
        .Campo("TotalEntrada").Formato = FMOEDA
        .Campo("TotalEntrada").SuprimirZeros = True
  
        If chkFluxo(4).value = vbChecked Then
          .AddCampo "TotalSaida", wrCSTotal, "Saída", wrTADireito, 31
        Else
          .AddCampo "TotalSaida", wrCSTotal, "Saída", wrTADireito, 30
        End If
        .Campo("TotalSaida").Formato = FMOEDA
        .Campo("TotalSaida").SuprimirZeros = True
      
      If TypeOf rstDados Is dao.Recordset Then
      
        If chkFluxo(4).value = vbChecked Then
        .AddCampo "TotalGeral", wrCSDataLink, "SUM(Saldo)", wrTADireito, 29
        Else
        .AddCampo "TotalGeral", wrCSDataLink, "SUM(Saldo)", wrTADireito, 30
        End If

        .Campo("TotalGeral").Formato = FMOEDA
        .Campo("TotalGeral").DataLink = "DATA = (SELECT MAX(T1.DATA) FROM " & rstDados.name & " AS T1 WHERE T1.Banco = " & rstDados.name & ".Banco)"
        .Campo("TotalGeral").TableLink = rstDados.name
        
      Else
        
        If chkFluxo(4).value = vbChecked Then
        .AddCampo "TotalGeral", wrCSDataLink, "SUM(Saldo)", wrTADireito, 29
        Else
        .AddCampo "TotalGeral", wrCSDataLink, "SUM(Saldo)", wrTADireito, 10
        End If
        
        .Campo("TotalGeral").Formato = FMOEDA
        .Campo("TotalGeral").DataLink = "DATA = (SELECT MAX(T1.DATA) FROM " & rstDados.Source & " AS T1 WHERE T1.Banco = " & rstDados.Source & ".Banco)"
        .Campo("TotalGeral").TableLink = rstDados.Source
        
      End If

    End With

  End With

  wrkSintetico.BeginPrint gTipoDB
  wrkSintetico.EndPrint
  Set wrkSintetico = Nothing
  
End Sub

' FUNCTION..: AddAnaliticoA
' Objetivo..: Adiciona os dados para o relatório analítico à tabela
'             auxiliar quando há quebra por Bancos.
' Argumentos: [rstBancos]: Recordset com os Bancos solicitados.
'             [rstTemp  ]: Recordset da tabela auxiliar.
'             [rstSaldos]: Recordset para gravação dos saldos.
'             [dtDatas  ]: Matriz com as datas (elemento 0: data inicial)
'                          (elemento 1: data final)
'             [strInstr ]: Matriz com as instruções de filtragem dos dados
' Retorna...: True se a função gerar o cálculo corretamente, False se algum
'             erro ocorrer durante a geração ou o usuário cancelar.
' --------------------------------------------------------------------------
Private Function AddAnaliticoA(rstBancos As Object, rstTemp As Object, rstSaldos As Object, dtDatas() As Date, strInstr() As String) As Boolean
    Dim cSaldo        As Currency     '// Acumula o saldo dia a dia
    Dim DtDia         As Date         '// Dia do cálculo
    Dim DtData        As Date         '// Data para Cálculo
    Dim MaisUm        As Boolean      '// Controle para banco zerado
    Dim dblCotacao    As Double       '// Valor da Cotacao
    
    'Se possui banco inicial não mostra banco zero
    MaisUm = IsValid(txtFluxo(2).Text)
    
On Error GoTo AddAnaliticoA_Erro
    
    rstBancos.MoveFirst
    Do
        If (mbolCancelou) Then
            GoTo AddAnaliticoA_Erro
        End If
        'Permitindo ao usuário cancelar o cálculo
        DoEvents
        'dtDia começa como a data inicial
        DtDia = dtDatas(0)
        
        'Protocolo 76585
        'Adicionado parametro TipoReg
        cSaldo = SaldoInicial(GetValue(rstBancos, "Banco", ZERO), dtDatas(0), strMoeda:=txtFluxo(7).Text, StrDescMoeda:=lblNome(5).Caption, sConciliado:=cboFluxo.Text, TipoRel:=mbitTipo, bConsiderarAtrasados:=chkFluxo(10).value)  '// Saldo inicial deste Banco
        
        If (TemMoeda(txtFluxo(7).Text, lblNome(5).Caption)) Then
            dblCotacao = UltimaCotacao(txtFluxo(7).Text, dtDatas(0))
            If dblCotacao > 0 Then
                cSaldo = Round(cSaldo / dblCotacao, 2) 'Saldo em Reais dividido pela cotacao da moeda base
            Else
                MsgFunc "Não há cotação para a data: " & dtDatas(0)
                cSaldo = 0
            End If
        End If
        
        'Grava o Saldo inicial deste Banco na tabela auxiliar
        rstSaldos.AddNew
        rstSaldos("Banco").value = GetValue(rstBancos, "Banco", ZERO)
        rstSaldos("Data").value = dtDatas(0)
        'False é usado para identificar saldo inicial
        rstSaldos("Tipo").value = False
        'Verifica se existe Cotação e Se existe Moeda Caso não satisfaça a Cotação não emite o relatório
        rstSaldos("Valor").value = cSaldo
        rstSaldos.update
        
        'Adiciona um registro em branco por Banco
        'Para que os saldos sejam apresentados
        'Ainda que não exista movimentação no período especificado
        If chkImprimeBancoSemMovimento.value = vbChecked Then
            rstTemp.AddNew
            Dim i As Integer
            For i = 0 To rstTemp.Fields.Count - 1
                'pt. 104498 - Ivo Sousa (26/01/2011)
                'Select Case TransDBTypeRetInt(rstTemp(i).Type)
                Select Case rstTemp(i).Type
                    Case dbText
                        'by kleber 2305
                        'ajustes para funcionar com SQL
                        If TypeOf rstTemp Is dao.Recordset Then
                            rstTemp(i).value = CStr(DefaultValue(rstTemp(i).SourceTable, rstTemp.Fields(i)))
                        Else
                            rstTemp(i).value = CStr(DefaultValue(rstTemp.Source, rstTemp.Fields(i)))
                        End If
                    Case dbDate
                        If TypeOf rstTemp Is dao.Recordset Then
                            rstTemp(i).value = CDateDef(DefaultValue(rstTemp(i).SourceTable, rstTemp.Fields(i)))
                        Else
                            rstTemp(i).value = CDateDef(DefaultValue(rstTemp.Source, rstTemp.Fields(i)))
                        End If
                    Case dbInteger, dbLong, dbByte
                        If TypeOf rstTemp Is dao.Recordset Then
                            rstTemp(i).value = CLngDef(DefaultValue(rstTemp(i).SourceTable, rstTemp.Fields(i)))
                        Else
                            rstTemp(i).value = CLngDef(DefaultValue(rstTemp.Source, rstTemp.Fields(i)))
                        End If
                        
                    Case dbCurrency
                      If TypeOf rstTemp Is dao.Recordset Then
                          rstTemp(i).value = CLngDef(DefaultValue(rstTemp(i).SourceTable, rstTemp.Fields(i)))
                      Else
                          rstTemp(i).value = CLngDef(DefaultValue(rstTemp.Source, rstTemp.Fields(i)))
                      End If
                End Select
                
                'Empresa deve possuir algum conteúdo ou o gerador acusará erro(DataLink)
                If rstTemp(i).name = "Empresa" Then
                    rstTemp(i).value = Space(1)
                End If
            Next i
            rstTemp("Banco").value = GetValue(rstBancos, "Banco", ZERO)
            rstTemp("Nome").value = GetValue(rstBancos, "Nome", NUL)
            rstTemp("Data").value = dtDatas(0)
            rstTemp.update
        End If
        
        'Até que dtdia seja posterior a data final
        Do Until (DateDiff("d", DtDia, dtDatas(1)) < ZERO)
            If (mbolCancelou) Then
                GoTo AddAnaliticoA_Erro
            End If
            'Possibilita ao usuário cancelar
            DoEvents
            SimpleMsgBar ResolveResString(1023, resUM, CStr(GetValue(rstBancos, "Banco", ZERO)), resDOIS, GetValue(rstBancos, "Nome", NUL), resTRES, DataToStr(DtDia))
            If (Not SelectDados(rstTemp, GetValue(rstBancos, "Banco", ZERO), GetValue(rstBancos, "Nome", NUL), DtDia, strInstr, cSaldo)) Then
                GoTo AddAnaliticoA_Erro
            End If
            
            'Grava o saldo final do dia para este banco
            rstSaldos.AddNew
            rstSaldos("Banco").value = GetValue(rstBancos, "Banco", ZERO)
            rstSaldos("Data").value = DtDia
            'True é usado para identificar o saldo final
            rstSaldos("Tipo").value = True
            rstSaldos("Valor").value = cSaldo
            rstSaldos.update
        
            'Chama a função UpdateSaldoBanco que atualizará a tabela de Saldos Bancários
            'se for necessário
            If TemMoeda(txtFluxo(7).Text, lblNome(5).Caption) = False Then
                UpdateSaldoBanco CStr(GetValue(rstBancos, "Banco", ZERO)), DtDia, cSaldo
            End If
            'Soma um dia a data atual
            DtDia = DateAdd("d", 1, DtDia)
        Loop
                
        If Not rstBancos.EOF Then
            'Move para o próximo Banco
            rstBancos.MoveNext
        Else
            MaisUm = True
        End If
    'Loop até chegar ao final da tabela de Bancos
    Loop Until (rstBancos.EOF) And (MaisUm)
    
    If (EstaVazio(rstTemp)) Then
        MsgFunc LoadResString(IDS_NORECORD)
        AddAnaliticoA = False
    Else
        AddAnaliticoA = True
    End If
    Exit Function
    
AddAnaliticoA_Erro:
    If (err.Number) Then
        If (rstSaldos.EditMode <> dbEditNone) Then
            rstSaldos.CancelUpdate
        End If
        DAOErros erro(err.Number) & " AddAnaliticoA"
    End If
    AddAnaliticoA = False
End Function

' FUNCTION..: AddAnaliticoB
' Objetivo..: Executa o cálculo do relatório analítico sem a quebra por Bancos.
' Argumentos: [rstBancos]: Recordset com os bancos.
'             [rstTemp  ]: Recordset que receberá os dados do relatório.
'             [rstSaldos]: Recordset onde serão gravados os saldos.
'             [dtDatas  ]: Matriz com as datas (elemento 0: data inicial)
'                          (elemento 1: data final)
'             [strInstr ]: Matriz com as intruções para abertura dos dados.
' Retorna...: True se a função executar todo o cálculo e o relatório puder
'             ser exibidi, False se algum erro impedir o término do cálculo ou
'             se o usuário cancelar.
' ----------------------------------------------------------------------------
Private Function AddAnaliticoB(rstBancos As Object, rstTemp As Object, rstSaldos As Object, dtDatas() As Date, strInstr() As String) As Boolean
Dim cSaldo  As Currency            '// Acumula o saldo dos dias
Dim DtDia   As Date                '// Dias de cálculo
Dim MaisUm  As Boolean

  MaisUm = IsValid(txtFluxo(2).Text)

  On Error GoTo AddAnaliticoB_Erro
  '//
  '// Primeiro soma o saldo inicial de todos os bancos, como não há quebra
  '// o saldo inicial é o saldo de todos os bancos que entram no cálculo
  '//
  rstBancos.MoveFirst
  Do
    'Protocolo 76585
    'Adicionado parametro TipoReg
  
    cSaldo = cSaldo + SaldoInicial(GetValue(rstBancos, "Banco", ZERO), dtDatas(0), strMoeda:=txtFluxo(7).Text, StrDescMoeda:=lblNome(5).Caption, sConciliado:=cboFluxo.Text, TipoRel:=mbitTipo, bConsiderarAtrasados:=chkFluxo(10).value)
    If (TemMoeda(txtFluxo(7).Text, lblNome(5).Caption)) Then
       dblCotacao = UltimaCotacao(txtFluxo(7).Text, dtDatas(0))
       If dblCotacao > 0 Then
          cSaldo = cSaldo / dblCotacao  'Saldo em Reais dividido pela cotacao da moeda base
       Else
          MsgFunc "Não há cotação para a data: " & dtDatas(0)
          cSaldo = 0
       End If
    End If
    
    If Not rstBancos.EOF Then
      rstBancos.MoveNext
    Else
      MaisUm = True
    End If
  Loop Until (rstBancos.EOF) And (MaisUm)
  
  rstSaldos.AddNew                      '// Gravando o Saldo inicial
  rstSaldos("Banco").value = ZERO       '// Zero porque só haverá um saldo inicial
  rstSaldos("Data").value = dtDatas(0)  '// Data inicial
  rstSaldos("Tipo").value = False       '// Só haverá saldo inicial
  rstSaldos("Valor").value = cSaldo
  rstSaldos.update
  
  DtDia = dtDatas(0)
  Do Until (DateDiff("d", DtDia, dtDatas(1)) < ZERO) '// Até dtDia seja posterior a data final
  
    MaisUm = IsValid(txtFluxo(2).Text)
    
    rstBancos.MoveFirst
    
    If (mbolCancelou) Then GoTo AddAnaliticoB_Erro
    DoEvents                        '// Possibilita ao usuário cancelar
    
    Do
      If (mbolCancelou) Then GoTo AddAnaliticoB_Erro
      DoEvents                      '// Possibilita, denovo, ao usuário cancelar

      SimpleMsgBar ResolveResString(1023, _
                                    resUM, CStr(GetValue(rstBancos, "Banco", NUL)), _
                                    resDOIS, GetValue(rstBancos, "Nome", NUL), _
                                    resTRES, DataToStr(DtDia))
      If (Not SelectDados(rstTemp, GetValue(rstBancos, "Banco", ZERO), _
                          GetValue(rstBancos, "Nome", NUL), DtDia, _
                          strInstr, cSaldo)) Then
        GoTo AddAnaliticoB_Erro
      End If
      
      If Not rstBancos.EOF Then
        rstBancos.MoveNext                '// Move para o próximo Banco
      Else
        MaisUm = True
      End If

    Loop Until (rstBancos.EOF) And (MaisUm)     '// Loop até o final da tabela
    '//
    '// Grava o saldo final deste dia na tabela auxiliar
    '//
    rstSaldos.AddNew
    rstSaldos("Banco").value = ZERO
    rstSaldos("Data").value = DtDia
    rstSaldos("Tipo").value = True
    rstSaldos("Valor").value = cSaldo
    rstSaldos.update

    DtDia = DateAdd("d", 1, DtDia)  '// Adiciona um dia a data atual
  Loop

  If (EstaVazio(rstTemp)) Then
    MsgFunc LoadResString(IDS_NORECORD)
    AddAnaliticoB = False
  Else
    AddAnaliticoB = True
  End If
  Exit Function
  
AddAnaliticoB_Erro:

  If (err.Number) Then
    If (rstSaldos.EditMode <> dbEditNone) Then rstSaldos.CancelUpdate
    DAOErros erro(err.Number) & " AddAnaliticoB"
  End If
  AddAnaliticoB = False
  
End Function

' FUNCTION..: SelectDados
' Objetivo..: Esta função seleciona os dados das tabelas de Duplicatas,
'             Lançamentos, Transf Bancárias e Aplicações para gravação
'             no arquivo temporário.
' Argumentos: [rstAux  ]: Recordset temporário para gravação dos dados.
'             [lngBanco]: Código do Banco atual da pesquisa.
'             [strBanco]: Nome do Banco.
'             [dtData  ]: Data para o movimento.
'             [strExp  ]: Matriz com parte das instruções de seleção.
'             [cSaldo  ]: Saldo inicial para o cálculo.
' Retorna...: Se a função obtiver sucesso, True, caso contrário False.
'             O argumento cSaldo retornará com o Saldo atualizado.
' ---------------------------------------------------------------------
Private Function SelectDados(rstAux As Object, lngBanco As Long, strBanco As String, DtData As Date, strExp() As String, cSaldo As Currency) As Boolean
Dim strWhere As String          '// Instrução de seleção de dados completa
Dim strConciliado As String
  
  
  Select Case cboFluxo.Text
    Case "Todos"
       strConciliado = ""
    Case "Sim"
       strConciliado = " AND Conciliado = TRUE "
    Case "Não"
       strConciliado = " AND Conciliado = FALSE "
  End Select

  
  '//
  '// Criando a consulta para a tabela de Transf Bancária como
  '// banco de origem. Esta instrução está no elemento zero da
  '// matriz strExp.
  '//
  
  If cboFluxo.Text <> "Não" Then
    strWhere = "SELECT *, '' AS MOEDA FROM [Transf Bancária] WHERE " & strExp(0) & " ORDER BY Código"
    MidStr strWhere, "<Banco>", CStr(lngBanco)
    MidStr strWhere, "<Data>", InverteData(DtData)
    If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, ZERO)) Then
      GoTo SeleMovimento_Erro
    End If
  End If
  '//
  '// Criando a consulta para a tabela de Transf Bancária como banco
  '// de Destino. Esta instrução está no elemento 1 da matriz strExp
  '//
  If cboFluxo.Text <> "Não" Then
    strWhere = "SELECT *, '' AS MOEDA FROM [Transf Bancária] WHERE " & strExp(1) & " ORDER BY Código"
    MidStr strWhere, "<Banco>", CStr(lngBanco)
    MidStr strWhere, "<Data>", InverteData(DtData)
    If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 1)) Then
      GoTo SeleMovimento_Erro
    End If
  End If
  '//
  '// Criando a instrução para a tabela de Aplicações com o tipo
  '// Juros/Correção, esta instrução está no elemento 2 da matriz
  '// strExp.
  '//
  If cboFluxo.Text <> "Não" Then
    strWhere = "SELECT *, '' AS MOEDA FROM Aplicações WHERE " & strExp(2) & " ORDER BY Código"
    MidStr strWhere, "<Banco>", CStr(lngBanco)
    MidStr strWhere, "<Data>", InverteData(DtData)
    If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 2)) Then
      GoTo SeleMovimento_Erro
    End If
  End If
  '//
  '// Cria a instrução para Aplicações com o Tipo Taxas ou CPMF
  '// Esta instrução está no elemento 3 da matriz strExp
  '//
  If cboFluxo.Text <> "Não" Then
    strWhere = "SELECT *, '' AS MOEDA FROM Aplicações WHERE " & strExp(3) & " ORDER BY Código"
    MidStr strWhere, "<Banco>", CStr(lngBanco)
    MidStr strWhere, "<Data>", InverteData(DtData)
    If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 3)) Then
      GoTo SeleMovimento_Erro
    End If
  End If
  '//
  '// Criando a instrução para os dados da tabela de Lançamentos a Receber.
  '// Esta instrução está no elemento 4 da Matriz strExp
  '//
  'Protocolo 73121: Criado campo parcela
  strWhere = "SELECT '' as Parcela, Código, Empresa, Tipo, Descrição, Controle, Pagamento, Vencimento, " & _
             "([Valor Original] + Acréscimo - Abatimento) AS Valor, " & _
             "Conta, Cheque, Moeda FROM Lançamentos WHERE " & strExp(4) & strConciliado & " AND Situação <> 'Cancelada' ORDER BY Código"
  MidStr strWhere, "<Banco>", CStr(lngBanco)
  MidStr strWhere, "<Data>", InverteData(DtData)
  If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 4)) Then
    GoTo SeleMovimento_Erro
  End If
  '//
  '// Criando a instrução para Duplicatas a Receber ou Recebidas. Troca o
  '// índice para a função SeleMovimento, Troca o campo de ordem e o nome da
  '// tabela.
  '//
  MidStr strWhere, "Código,", "Nota AS Código,"
  MidStr strWhere, "'' as Parcela,", "Parcela,"
  MidStr strWhere, "Lançamentos", "Duplicatas"
  MidStr strWhere, "BY Código", "BY Nota"
  If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 4)) Then
    GoTo SeleMovimento_Erro
  End If
  '//
  '// Criando a instrução para Lançamentos a Pagar ou Pagos. Esta instrução
  '// está no elemento 5 da matriz strExp.
  '//
  strWhere = "SELECT '' as Parcela, Código, Empresa, Tipo, Descrição, Controle, Pagamento, Vencimento, " & _
             "([Valor Original] + Acréscimo - Abatimento) AS Valor, " & _
             "Conta, Cheque, Moeda FROM Lançamentos WHERE " & strExp(5) & strConciliado & " AND Situação <> 'Cancelada' ORDER BY Código"
  MidStr strWhere, "<Banco>", CStr(lngBanco)
  MidStr strWhere, "<Data>", InverteData(DtData)
  If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 5)) Then
    GoTo SeleMovimento_Erro
  End If
  '//
  '// Criando a instrução para Duplicatas a Pagar ou Pagas. Troca o campo Código pelo
  '// Nota na instrução e altera o nome da tabela.
  '//
  MidStr strWhere, "Código,", "Nota AS Código,"
  MidStr strWhere, "'' as Parcela,", "Parcela,"
  MidStr strWhere, "Lançamentos", "Duplicatas"
  MidStr strWhere, "BY Código", "BY Nota"
  If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 5)) Then
    GoTo SeleMovimento_Erro
  End If
  
  If chkFluxo(6).value = vbChecked Or _
       chkFluxo(7).value = vbChecked Or _
       chkFluxo(8).value = vbChecked Or _
       chkFluxo(9).value = vbChecked Then
    '//
    '// Criando a consulta para a tabela de Pedidos sendo a Receber
    '//
    strWhere = "SELECT Número as Código, *, '' AS MOEDA FROM " & NomeTabeladoRST(rstPrevisao) & " WHERE " & Replace(strExp(4), "Liberação", "Vencimento", , , vbTextCompare) & " ORDER BY Número"
    MidStr strWhere, "<Banco>", CStr(lngBanco)
    MidStr strWhere, "<Data>", InverteData(DtData)
    If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 6)) Then
      GoTo SeleMovimento_Erro
    End If
    '//
    '// Criando a consulta para a tabela de Pedidos sendo a Pagar
    '//
    strWhere = "SELECT Número as Código, *, '' AS MOEDA FROM " & NomeTabeladoRST(rstPrevisao) & " WHERE " & Replace(strExp(5), "Liberação", "Vencimento", , , vbTextCompare) & " ORDER BY Número"
    MidStr strWhere, "<Banco>", CStr(lngBanco)
    MidStr strWhere, "<Data>", InverteData(DtData)
    If (Not GravaAnalitico(rstAux, strWhere, lngBanco, strBanco, DtData, cSaldo, 7)) Then
      GoTo SeleMovimento_Erro
    End If
  End If
  
  
  SelectDados = True
  Exit Function
  
SeleMovimento_Erro:
   
  If (err.Number) Then
    If (rstAux.EditMode <> dbEditNone) Then rstAux.CancelUpdate
    DAOErros erro(err.Number) & " SelMovimento"
  End If
  SelectDados = False
  
End Function

' FUNCTION..: GravaAnalitico
' Objetivo..: Grava os valores na tabela auxiliar para o relatório
'             analítico.
' Argumentos: [rstTemp]: Recordset da tabela temporária.
'             [strInst]: Instrução de abertura do Recordset dos dados.
'             [lBanco ]: Código do Banco atual.
'             [sBanco ]: Nome do Banco atual.
'             [datData]: Data do cálculo atual.
'             [cSaldo ]: Saldo Anterior.
'             [lngTipo]: Define qual é a tabela de origem. Os valores são
'                        os mesmos da matriz de consulta. Zero para Tranf Bancária
'                        como banco de Origem, 1 para Transf Bancária como banco de
'                        Destino, 2 para Aplicações com o Tipo Juros/Correção, 3
'                        para Aplicações com o Tipo Taxas ou CPMF, 4 para Duplicatas
'                        ou Lançamentos a Receber e 5 para Duplicatas ou Lançamentos
'                        a pagar.
' Retorna...: A função retorna True se gravar os dados com sucesso, False se
'             algum erro ocorrer ou o usuário cancelar. O argumento cSaldo retorna
'             com o Saldo atualizado.
' ----------------------------------------------------------------------------------
Private Function GravaAnalitico(rstTemp As Object, strInst As String, lBanco As Long, sBanco As String, datData As Date, cSaldo As Currency, lngTipo As Long) As Boolean
Dim curValor As Currency        '// Valor do movimento atual
Dim rstDados As Object          '// Recordset com os dados
Dim dblCotMoedaDoc As Double    '// Cotação da moeda do documento
Dim dblCotMoedaBase As Double   '// Cotação da moeda base

  On Error GoTo GravaAnalitico_Erro
  
  If (AbreRecordset(rstDados, strInst, dbOpenForwardOnly) = WL_OK) Then
    Do Until (rstDados.EOF)     '// Loop até o final do Recordset
      If (mbolCancelou) Then GoTo GravaAnalitico_Erro
      DoEvents                  '// Possibilita ao usuário cancelar o cálculo

      curValor = GetValue(rstDados, "Valor", ZERO)
      rstTemp.AddNew
      rstTemp("Banco").value = lBanco
      rstTemp("Nome").value = sBanco
      rstTemp("Data").value = datData
      
      If (lngTipo < 2) Then               '// Tabela de Transf Bancária
        rstTemp("Empresa").value = CTRANS
        rstTemp("Pagamento").value = GetValue(rstDados, "Data", Null)
        rstTemp("Type").value = DADOS_TRANSF
      ElseIf (lngTipo < 4) Then           '// Tabela de Aplicações
        rstTemp("Empresa").value = CAPLIC
        rstTemp("Pagamento").value = GetValue(rstDados, "Data", Null)
        rstTemp("Type").value = DADOS_APLIC
      ElseIf lngTipo = 6 Or lngTipo = 7 Then
        rstTemp("Empresa").value = GetValue(rstDados, "Empresa", NUL)
        rstTemp("Vencimento").value = GetValue(rstDados, "Vencimento")
        rstTemp("Pagamento").value = GetValue(rstDados, "Pagamento")
        rstTemp("Type").value = DADOS_LANC
      Else                                '// Tabela de Duplicatas ou Lançamentos
        rstTemp("Empresa").value = GetValue(rstDados, "Empresa", NUL)
        rstTemp("Vencimento").value = GetValue(rstDados, "Vencimento")
        rstTemp("Pagamento").value = GetValue(rstDados, "Pagamento")
        rstTemp("Type").value = DADOS_LANC
      End If
      rstTemp("Controle").value = GetValue(rstDados, "Controle", NUL)
      rstTemp("Duplicata").value = GetValue(rstDados, "Código", ZERO)
      'Protocolo 73121: Criado campo parcela
      rstTemp("Parcela").value = GetValue(rstDados, "Parcela", NUL)
      If rstTemp("Parcela").value <> NUL Then
         rstTemp("Parcela").value = "P" & rstTemp("Parcela").value
      End If
      rstTemp("Tipo").value = IIf((lngTipo > 1), _
                                        GetValue(rstDados, "Tipo", NUL), _
                                        NUL)        '// Transf não tem tipo
      If ((lngTipo = 6) Or (lngTipo = 7)) Then
        rstTemp("Descrição").value = "Ref. a Pedido"
      Else
        rstTemp("Descrição").value = GetValue(rstDados, "Descrição", NUL)
      End If
      rstTemp("Conta").value = GetValue(rstDados, "Conta", ZERO)
      
      If ((lngTipo = ZERO) Or (lngTipo = 5)) Then
        rstTemp("Cheque").value = GetValue(rstDados, "Cheque", ZERO)
      Else
        rstTemp("Cheque").value = ZERO
      End If
        'Se a moeda do lançamento/duplicatata for <> da moeda base
        'Calcule o valor com base na ultima cotacao encontrada até a data de
        'pagamento, senão manter o valor
        If GetValue(rstDados, "Moeda", ZERO) <> txtFluxo(7).Text Then
           If Not IsNull(rstTemp("Pagamento").value) Then
              dblCotMoedaDoc = UltimaCotacao(GetValue(rstDados, "Moeda"), rstTemp("Pagamento").value)
              dblCotMoedaBase = UltimaCotacao(txtFluxo(7).Text, rstTemp("Pagamento").value)
           Else
              dblCotMoedaDoc = UltimaCotacao(GetValue(rstDados, "Moeda"), rstTemp("Data").value)
              dblCotMoedaBase = UltimaCotacao(txtFluxo(7).Text, rstTemp("Data").value)
           End If
           If dblCotMoedaBase = 0 Then
              'Quando não houver cotação, para evitar erro de divisao por zero
              curValor = 0
           Else
              'Essa cálculo garante a conversão entre moedas: Exemplo Euro para Dolar ou Dolar para Peso, etc
              'pois converte primeiro o valor para reais e depois converte para a moeda base
              curValor = curValor * dblCotMoedaDoc / dblCotMoedaBase
           End If
        End If
      
      If ((lngTipo = ZERO) Or (lngTipo = 3) Or (lngTipo = 5) Or (lngTipo = 7)) Then  '// Saídas
        'Protocolo 77178 ----------------------------
        rstTemp("Saída").value = Round(curValor, 2)
        '--------------------------------------------
        rstTemp("Entrada").value = ZERO
        cSaldo = (cSaldo - curValor)
      Else                                                          '// Entradas
        rstTemp("Saída").value = ZERO
        'Protocolo 77178 ----------------------------
        rstTemp("Entrada").value = Round(curValor, 2)
        '--------------------------------------------
        cSaldo = (cSaldo + curValor)
      End If

      rstTemp.update
      rstDados.MoveNext
    Loop
  End If
  FechaRecordset rstDados
  GravaAnalitico = True
  Exit Function
  
GravaAnalitico_Erro:

  If (err.Number) Then
    If (rstTemp.EditMode <> dbEditNone) Then rstTemp.CancelUpdate
    DAOErros NUL
  End If
  GravaAnalitico = False
  
End Function

' SUB.......: RelatorioAnalito
' Objetivo..: Gera o relatório analítico
' Argumentos: [rstSource]: Recordset com os dados de origem
'             [pdDest   ]: Destino da impressão.
'             [datDatas ]: Matriz com as Datas inicial e final de filtro.
'             [strSaldos]: Nome da tabela auxiliar que contém os saldos dos Bancos.
'
' ----------------------------------------------------------------------------------
Private Sub RelatorioAnalitico(rstSource As Object, pdDest As PrintDestinoEnum, datDatas() As Date, strSaldos As String)
Dim wrkAnalitico As KeybReport
Dim strSubTitulo As String
Dim strSaldoCredor  As String       'String de DataLink para Campo de Saldo Credor
Dim strSaldoDevedor As String       'String de DataLink para Campo de Saldo Devedor
Dim bQuebra         As Boolean      'Define se há quebra por banco ou não
  
  bQuebra = (chkFluxo(4).value = vbChecked)
  SimpleMsgBar LoadResString(160)
  strSubTitulo = Me.Caption & " Analítico de " & DataToStr(datDatas(0))
  AppendStr strSubTitulo, " até " & DataToStr(datDatas(1))

  Set wrkAnalitico = New KeybReport
  With wrkAnalitico
    Set .DatabaseName = GlobalDataBase
    Set .Recordset = rstSource
    .ScaleMode = vbMillimeters
    .Tipo = wrObjectDraw
    .AutoRedraw = True
    .Destino = pdDest
    .WindowTitulo = "Fluxo de Caixa Analítico"
    
    PageHeader wrkAnalitico, strSubTitulo
    
    If Len(txtFluxo(7).Text) > 0 Then
      .UltimaSecao.AddLinha "Moeda"
      .UltimaSecao(.UltimaSecao.LinhasCount).AddCampo , wrCSFixedText, "Valores Demonstrados em: " & txtFluxo(7).Text, wrTACentro
    End If
    '
    '// Se não há quebra insere os bancos listado no cabeçalho
    '
    If Not (bQuebra) Then
      .UltimaSecao.AddLinha
      
      Dim strNomesBancos  As String
      strNomesBancos = "Bancos selecionados: "
      
      If IsValid(txtFluxo(2).Text) Or IsValid(txtFluxo(3).Text) Then
        
        Dim SQLBancos   As String
        
        SQLBancos = "SELECT Nome FROM Bancos WHERE Banco "
        
        If IsValid(txtFluxo(2).Text) And IsValid(txtFluxo(3).Text) Then
          Concat SQLBancos, " BETWEEN ", txtFluxo(2).Text, " AND ", txtFluxo(3).Text
          
        ElseIf IsValid(txtFluxo(2).Text) Then
          Concat SQLBancos, " >= ", txtFluxo(2).Text
          
        ElseIf IsValid(txtFluxo(3).Text) Then
          Concat SQLBancos, " <= ", txtFluxo(3).Text
          
        End If
        
        Dim rstNomesBancos   As Object
        
        If AbreRecordset(rstNomesBancos, SQLBancos, dbOpenSnapshot) = WL_OK Then
        
          Do
            
            Concat strNomesBancos, Trim$(GetValue(rstNomesBancos, "Nome", NUL)), ", "
            
            rstNomesBancos.MoveNext
            
          Loop Until (rstNomesBancos.EOF)
          
          strNomesBancos = Left$(strNomesBancos, Len(strNomesBancos) - 2)
          
        End If
        
        FechaRecordset rstNomesBancos
      Else
        Concat strNomesBancos, "Todos"
      End If
      
      
      .UltimaLinha.AddCampo , wrCSFixedText, strNomesBancos, wrTACentro
      
      .UltimaLinha.Campo(1).MultiLine = True
      .UltimaLinha.Campo(1).Left = 10
    End If
    
    .FontSize = 8
    .FontStyle = wrFSBold

    '
    ' Criando o grupo principal: quebra por Banco
    '
    .AddGrupo "1", wrDBBottomBorder
    
    If (bQuebra) Then                 '// Se há quebra o grupo principal quebra por Banco
      .Grupo(1).Quebra = "Banco"
    End If
    
    .Grupo(1).AddSecao scHeader, 2
    With .Grupo(1).Header.Linha(2)
      If (bQuebra) Then
        .Height = wrkAnalitico.TextHeight("W") * 2
        .DrawBorder = wrDBAllBorders
        .AddCampo , wrCSFixedText, "Banco:", , 15
        .Campo(1).Top = ((.Height / 2) - (.Campo(1).Height / 2))
        .AddCampo , , "Banco", wrTADireito, 17, 16
        .Campo(2).Formato = StrZero(0, 9)
        .Campo(2).Top = .Campo(1).Top
        .AddCampo , , "Nome", , 50, 34
        .Campo(3).Top = .Campo(1).Top
      End If
      .AddCampo "saldo", wrCSFixedText, "Saldo Anterior:", , 30, 138
      .Campo("saldo").Top = ((.Height / 2) - (.Campo(1).Height / 2))
      .AddCampo "valor", wrCSDataLink, "Valor", wrTADireito, , 146
      .Campo("valor").Top = ((.Height / 2) - (.Campo(1).Height / 2))
      .Campo("valor").Formato = FMOEDA
      .Campo("valor").TableLink = strSaldos
 
      If BQuebraData Then
        If (bQuebra) Then
          .Campo("valor").DataLink = "Banco = {*Quebra} AND Data = " & _
                                     InverteData(datDatas(0), True) & " AND " & _
                                     "Tipo = False"          '// Tipo False é igual ao saldo inicial
        Else
          .Campo("valor").DataLink = "Banco = 0 AND Tipo = False" '// Quando não há quebra o banco é igual a zero
        End If
      Else
        If (bQuebra) Then
          .Campo("valor").DataLink = "Banco = {*Quebra} AND Tipo = False"          '// Tipo False é igual ao saldo inicial
        Else
          .Campo("valor").DataLink = "Banco = 0 AND Tipo = False" '// Quando não há quebra o banco é igual a zero
        End If
      End If
    End With
    '
    ' SubGrupo: quebra por Data
    '
    .Grupo(1).AddSubGrupo "1"
    
    If (BQuebraData) Then
      .Grupo(1).Subgrupo(1).Quebra = "Data"
    End If
      
    .Grupo(1).Subgrupo(1).AddSecao scHeader, 4
    
    If BQuebraData Then
      With .Grupo(1).Subgrupo(1).Header.Linha(2)
        .AddCampo , wrCSFixedText, "Movimentação do dia", , 35
        .AddCampo , , "Data", wrTACentro, 17
        .Campo(2).Formato = FDATA
        .AddCampo , wrCSSimpleLine
        .Campo(3).BorderStyle = wrDot
      End With
    End If
    
    
    With .Grupo(1).Subgrupo(1).Header.Linha(4)
      .AddCampo , wrCSFixedText, "Empresa", , 22
      .AddCampo , wrCSFixedText, "Código", , 11
      .AddCampo , wrCSFixedText, "Tipo", , 11.5
      '
      ' Se o usuário não deseja imprimir a descrição completa
      '
      If (chkFluxo(2).value = vbUnchecked) Then
        .AddCampo , wrCSFixedText, "Descrição", , 30
        .AddCampo , wrCSFixedText, "Controle", , 14.5, 77.5
      End If
      .AddCampo , wrCSFixedText, "Conta", wrTADireito, 12
      .AddCampo , wrCSFixedText, "Cheque", wrTADireito, 14.5
      .AddCampo , wrCSFixedText, "Vencto.", wrTACentro, 13
      .AddCampo , wrCSFixedText, "Pagto.", wrTACentro, 13
      .AddCampo , wrCSFixedText, "Entrada", wrTADireito, 22, 146
      .AddCampo , wrCSFixedText, "Saída", wrTADireito, 22
    End With
    '
    ' Seção principal de detalhes
    '
    .FontStyle = wrFSNormal
    .Grupo(1).Subgrupo(1).AddSecao scDetalhe, 1
    With .Grupo(1).Subgrupo(1).Detalhe.Linha(1)
      .AddCampo , , "Empresa", , 22
      .AddCampo "codigo", , "Duplicata", , 11
      .Campo(2).Formato = StrZero(0, 6)
      .Campo(2).SuprimirZeros = True
      .AddCampo , , "Tipo", , 11.5
      
      If (chkFluxo(2).value = vbUnchecked) Then
        .AddCampo , , "Descrição", , 30
        .AddCampo , , "Controle", , 14.5, 77.5
      End If
      .AddCampo "conta", , "Conta", wrTADireito, 12
      .Campo("conta").SuprimirZeros = True
      .AddCampo "cheque", , "Cheque", wrTADireito, 14.5
      .Campo("cheque").Formato = StrZero(0, 6)
      .Campo("cheque").SuprimirZeros = True
      .AddCampo "vencto", , "Vencimento", wrTACentro, 13
      .Campo("vencto").Formato = FDDMMYY
      .AddCampo "pagto", , "Pagamento", wrTACentro, 13
      .Campo("pagto").Formato = FDDMMYY
      .AddCampo "entrada", , "Entrada", wrTADireito, 22, 146
      .Campo("entrada").Formato = FMOEDA
      .Campo("entrada").SuprimirZeros = True
      .AddCampo "saida", , "Saída", wrTADireito, 22
      .Campo("saida").Formato = FMOEDA
      .Campo("saida").SuprimirZeros = True
    End With
    ' Se o usuário deseja imprimir a razão social da empresa
    '
    If (chkFluxo(1).value = vbChecked) Then
      .Grupo(1).Subgrupo(1).Detalhe.DrawBorder = wrDBBottomBorder
      .Grupo(1).Subgrupo(1).Detalhe.BorderStyle = wrDot
      .Grupo(1).Subgrupo(1).Detalhe.AddLinha "razão"
      .Grupo(1).Subgrupo(1).Detalhe.Linha("razão").AddCampo , wrCSFixedText, "Razão Social:", , 20
      .UltimoCampo.Left = 23
      .Grupo(1).Subgrupo(1).Detalhe("razão").Campo(1).FontStyle = wrFSBold
      .Grupo(1).Subgrupo(1).Detalhe("razão").AddCampo , wrCSDataLink, "Razão"
      .Grupo(1).Subgrupo(1).Detalhe("razão").Campo(2).TableLink = "Empresas"
      .Grupo(1).Subgrupo(1).Detalhe("razão").Campo(2).DataLink = "Apel = {Empresa}"
    End If
    '
    ' Se o usuário deseja imprimir a descrição completa
    '
    If (chkFluxo(2).value = vbChecked) Then
      .Grupo(1).Subgrupo(1).Detalhe.DrawBorder = wrDBBottomBorder
      .Grupo(1).Subgrupo(1).Detalhe.BorderStyle = wrDot
      .Grupo(1).Subgrupo(1).Detalhe.AddLinha "desc"
      With .Grupo(1).Subgrupo(1).Detalhe.Linha("desc")
        .AddCampo , wrCSFixedText, "Descrição:", , 20
        .Campo(1).Left = 23
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , , "Descrição", , 60
        .AddCampo , wrCSFixedText, "Controle:", , 15
        .Campo(3).FontStyle = wrFSBold
        .AddCampo , , "Controle"
      End With
    End If
    '
    ' Seção de rodapé: SubTotais
    '
    .Grupo(1).Subgrupo(1).AddSecao scFooter, 4
    .FontStyle = wrFSBold
    With .Grupo(1).Subgrupo(1).Footer.Linha(2)
      .AddCampo , wrCSFixedText, "Totais", , 15, 135
      .AddCampo , wrCSSubTotal, "Entrada", wrTADireito, 22, 146
      .Campo(2).Formato = FMOEDA
      .AddCampo , wrCSSubTotal, "Saída", wrTADireito, 22
      .Campo(3).Formato = FMOEDA
    End With
    
    strSaldoDevedor = "IIf((Valor > 0), Null, Valor)"
    strSaldoCredor = "IIf((Valor < 0), Null, Valor)"
    
    If BQuebraData Then
      
    With .Grupo(1).Subgrupo(1).Footer.Linha(3)
      .AddCampo , wrCSFixedText, "Saldo do dia:", , 25, 105
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSDataLink, "Data", wrTACentro, 17
      .Campo(2).Formato = FDATA
      .Campo(2).TableLink = GetTableSource(rstSource)
      .Campo(2).DataLink = "Data = {*Quebra}"
      
      .AddCampo "SaldoCredor", wrCSDataLink, strSaldoCredor, wrTADireito, 22, 146
      .Campo("SaldoCredor").Formato = FMOEDA
      .Campo("SaldoCredor").TableLink = strSaldos
      .AddCampo "saldoDevedor", wrCSDataLink, strSaldoDevedor, wrTADireito, 22
      .Campo("saldoDevedor").Formato = FMOEDA
      .Campo("saldoDevedor").TableLink = strSaldos

      If BQuebraData Then
        If (bQuebra) Then           '// Quando há quebra por bancos
          .Campo("SaldoCredor").DataLink = "Banco = {**Banco} AND Data = {*Data} AND Tipo = True"
          .Campo("saldoDevedor").DataLink = "Banco = {**Banco} AND Data = {*Data} AND Tipo = True"
        Else                        '// Quando não há quebra por bancos
          .Campo("SaldoCredor").DataLink = "Data = {*Data} AND Tipo = True"
          .Campo("saldoDevedor").DataLink = "Data = {*Data} AND Tipo = True"
        End If
      Else
        If (bQuebra) Then           '// Quando há quebra por bancos
          .Campo("SaldoCredor").DataLink = "Banco = {**Banco} AND Tipo = True"
          .Campo("saldoDevedor").DataLink = "Banco = {**Banco} AND Tipo = True"
        Else                        '// Quando não há quebra por bancos
          .Campo("SaldoCredor").DataLink = "Tipo = True"
          .Campo("saldoDevedor").DataLink = "Tipo = True"
        End If
      End If
      
    End With
    End If
    
    With .Grupo(1).Subgrupo(1).Footer(4)
      .AddCampo , wrCSSimpleLine
      .Campo(1).BorderStyle = wrDash
    End With
    '
    ' SubGrupo: Resumo
    '
    If (chkFluxo(3).value = vbChecked) Then GrupoResumo wrkAnalitico, rstSource, strSaldos
    
  End With

  wrkAnalitico.BeginPrint gTipoDB
  wrkAnalitico.EndPrint
  
  Set wrkAnalitico = Nothing
  
End Sub

' SUB.......: GrupoResumo
' Objetivo..: Cria o grupo resumo na página de impressão do relatório de
'             Fluxo de Caixa Analítico.
' Argumentos: [wrkReport]: Referência ao objeto KeybReport.
'             [rstOrigem]: Recordset com os dados de origem.
'             [strSaldo ]: Nome da tabela que contém os saldos dos bancos.
' ---------------------------------------------------------------------------------
Private Sub GrupoResumo(wrkReport As KeybReport, rstOrigem As Object, strSaldo As String)
Dim strTableName  As String
Dim strDtLnkAplic As String     '// Instrução SQL para os dados de Aplicações
Dim strDtLnkTrans As String     '// Instrução SQL para os dados de Transf Bancária
Dim strDtLnkLanct As String     '// Instrução SQL para os dados de Lançamentos e Duplicatas

  If (chkFluxo(4).value = vbChecked) Then
    strDtLnkAplic = "Banco = {**Quebra} AND Type = " & CStr(DADOS_APLIC)
    strDtLnkTrans = "Banco = {**Quebra} AND Type = " & CStr(DADOS_TRANSF)
    strDtLnkLanct = "Banco = {**Quebra} AND Type = " & CStr(DADOS_LANC)
  Else
    strDtLnkAplic = "Type = " & CStr(DADOS_APLIC)
    strDtLnkTrans = "Type = " & CStr(DADOS_TRANSF)
    strDtLnkLanct = "Type = " & CStr(DADOS_LANC)
  End If
  
  strTableName = GetTableSource(rstOrigem)
  With wrkReport
  
    If (chkFluxo(4).value = vbChecked) Then       'Quebra por Banco
      .Grupo(1).AddSubGrupo "resumo"
      .Grupo(1).Subgrupo("resumo").AddSecao scHeader, 6
    Else
      .AddGrupo "resumo"
      .Grupo("resumo").AddSecao scHeader, 6
    End If
    
    .FontStyle = wrFSBold
    With .UltimaSecao.Linha(1)
      .AddCampo , wrCSFixedText, "Resumo", , 30
      .AddCampo , wrCSFixedText, "Entradas", wrTADireito, 22, 120
      .AddCampo , wrCSFixedText, "Saídas", wrTADireito, 22
      .AddCampo , wrCSFixedText, "Total", wrTADireito, 30
    End With
    
    .FontStyle = wrFSNormal
    
    With .UltimaSecao.Linha(2)
      .AddCampo , wrCSFixedText, "Aplicações:", wrTADireito, 30, 90
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSDataLink, "SUM(Entrada)", wrTADireito, 22
      .Campo(2).Formato = FMOEDA
      .Campo(2).TableLink = strTableName
      .Campo(2).DataLink = strDtLnkAplic
      .AddCampo , wrCSDataLink, "SUM(Saída)", wrTADireito, 22
      .Campo(3).Formato = FMOEDA
      .Campo(3).TableLink = strTableName
      .Campo(3).DataLink = strDtLnkAplic
      .AddCampo , wrCSDataLink, "SUM(Entrada) - SUM(Saída)", wrTADireito, 30
      .Campo(4).Formato = FMOEDA
      .Campo(4).TableLink = strTableName
      .Campo(4).DataLink = strDtLnkAplic
    End With
    
    With .UltimaSecao.Linha(3)
      .AddCampo , wrCSFixedText, "Transferências:", wrTADireito, 30, 90
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSDataLink, "SUM(Entrada)", wrTADireito, 22
      .Campo(2).Formato = FMOEDA
      .Campo(2).TableLink = strTableName
      .Campo(2).DataLink = strDtLnkTrans
      .AddCampo , wrCSDataLink, "SUM(Saída)", wrTADireito, 22
      .Campo(3).Formato = FMOEDA
      .Campo(3).TableLink = strTableName
      .Campo(3).DataLink = strDtLnkTrans
      .AddCampo , wrCSDataLink, "SUM(Entrada) - SUM(Saída)", wrTADireito, 30
      .Campo(4).Formato = FMOEDA
      .Campo(4).TableLink = strTableName
      .Campo(4).DataLink = strDtLnkTrans
    End With
    
    With .UltimaSecao.Linha(4)
      .AddCampo , wrCSFixedText, "Movimentação:", wrTADireito, 30, 90
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSDataLink, "SUM(Entrada)", wrTADireito, 22
      .Campo(2).Formato = FMOEDA
      .Campo(2).TableLink = strTableName
      .Campo(2).DataLink = strDtLnkLanct
      .AddCampo , wrCSDataLink, "SUM(Saída)", wrTADireito, 22
      .Campo(3).Formato = FMOEDA
      .Campo(3).TableLink = strTableName
      .Campo(3).DataLink = strDtLnkLanct
      .AddCampo , wrCSDataLink, "SUM(Entrada) - SUM(Saída)", wrTADireito, 30
      .Campo(4).Formato = FMOEDA
      .Campo(4).TableLink = strTableName
      .Campo(4).DataLink = strDtLnkLanct
    End With
    
    With .UltimaSecao.Linha(5)
      .AddCampo , wrCSFixedText, "Saldo:", wrTADireito, 30, 90
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSDataLink, "Valor", wrTADireito, 30, 170
      .Campo(2).Formato = FMOEDA
      .Campo(2).TableLink = strSaldo
      
      If BQuebraData Then
        If (chkFluxo(4).value = vbChecked) Then
          .Campo(2).DataLink = "Banco = {**Quebra} AND Tipo = True AND Data = (SELECT MAX(Data) FROM " & strSaldo & ")"
        Else
          .Campo(2).DataLink = "Tipo = True AND Data = (SELECT MAX(Data) FROM " & strSaldo & ")"
        End If
      Else
        If (chkFluxo(4).value = vbChecked) Then
          .Campo(2).DataLink = "Banco = {**Quebra} AND Tipo = True"
        Else
          .Campo(2).DataLink = "Tipo = True"
        End If
      End If
    End With
    
    With .UltimaSecao.Linha(6)
      .AddCampo , wrCSSimpleLine
      .Campo(1).BorderStyle = wrDash
    End With
    
  End With
  
End Sub

Private Function PedidosPendentes()

  Dim strSql           As String
  Dim strSql2          As String
  Dim sTabela          As String
  
  sTabela = GBL_PDV
  
  AppendVar fdsPrevisao(0), "PagRec", dbText, 1
  AppendVar fdsPrevisao(1), "Empresa", dbText, 15
  AppendVar fdsPrevisao(2), "Número", dbLong, 6
  AppendVar fdsPrevisao(3), "Tipo", dbText, 20
  AppendVar fdsPrevisao(4), "Banco", dbLong, 6
  AppendVar fdsPrevisao(5), "Conta", dbLong, 6
  AppendVar fdsPrevisao(6), "Centro", dbLong, 6
  AppendVar fdsPrevisao(7), "Vencimento", dbDate
  AppendVar fdsPrevisao(8), "Valor", dbDouble
  
  If CrieAux(rstPrevisao, fdsPrevisao) Then
    
    'Protocolo 73636  Criada nova consulta que substitui a anterior
    '(agora a condição de pagamento será verificada para cada parcela)
    strSql = strSql + "SELECT "
    strSql = strSql + "PV.Número, "
    strSql = strSql + "PV.[Tipo de Registro], "
    strSql = strSql + "PV.Fornecedor, "
    strSql = strSql + "PV.Empresa, "
    strSql = strSql + "PV.[Condição de Pagamento], "
    strSql = strSql + "IPV.[Data da Previsão], "
    strSql = strSql + "PV.Banco, "
    strSql = strSql + "PV.Conta, "
    strSql = strSql + "IPV.[Centro de Custo], "
    strSql = strSql + "(IPV.[Data da Previsão] + PAR.Parcela) as Vencimento, "
    strSql = strSql + "((IPV.[Valor Líquido] / IPV.Quantidade) * "
    strSql = strSql + "(IPV.Quantidade - IPV.[Quantidade Baixada]) * "
    strSql = strSql + "(PAR.Porcentagem /  100 ) - "
    
    'Se for access segue com IIF senao CASE
    If gTipoDB = Access Then
        strSql = strSql + "IIF((SELECT [Valor Original] "
        strSql = strSql + "FROM Duplicatas "
        strSql = strSql + "WHERE Parcela = -1 "
        strSql = strSql + "AND Nota = PV.Número) IS NOT NULL, "
        strSql = strSql + "(SELECT SUM([Valor Original] + Acréscimo - Abatimento) "
        strSql = strSql + "FROM Duplicatas "
        strSql = strSql + "WHERE Parcela = -1 AND Nota = PV.Número), 0)) AS ValorResult "
    Else
        strSql = strSql + "CASE WHEN (SELECT [Valor Original] "
        strSql = strSql + "FROM Duplicatas "
        strSql = strSql + "WHERE Parcela = -1 "
        strSql = strSql + "AND Nota = PV.Número) IS NOT NULL THEN "
        strSql = strSql + "(SELECT SUM([Valor Original] + Acréscimo - Abatimento) "
        strSql = strSql + "FROM Duplicatas "
        strSql = strSql + "WHERE Parcela = -1 AND Nota = PV.Número) ELSE 0 END ) AS ValorResult "
    End If
    
    strSql = strSql + "FROM "
    strSql = strSql + "[Pedidos de Venda] PV, "
    strSql = strSql + "[Itens de Pedidos de Venda] IPV, "
    strSql = strSql + "[Condições de Pagamento] CON, "
    strSql = strSql + "[Parcelas] PAR "
    
    strSql = strSql + "WHERE "
    strSql = strSql + "IPV.Situação = 'Pendente' "
    strSql = strSql + "AND IPV.[Data da Previsão] IS NOT NULL "
    strSql = strSql + "AND (IPV.Quantidade - IPV.[Quantidade Baixada]) > 0 "
    strSql = strSql + "AND PV.Número = IPV.Número "
    strSql = strSql + "AND PV.[Tipo de Registro] = IPV.[Tipo de Registro] "
    strSql = strSql + "AND PV.[Condição de Pagamento] = CON.Código "
    strSql = strSql + "AND CON.Código = PAR.Condição "
    strSql = strSql + "AND IPV.[Data da Previsão] + PAR.Parcela BETWEEN " & InverteData(dtInicial, True) & " AND " & InverteData(dtFinal, True)
    
    'Pedidos de Venda
    If chkFluxo(6).value = vbChecked Then
      RegistrarPendentes strSql, GBL_PDV
      
    End If
    
    'Pedidos de Compra
    If chkFluxo(7).value = vbChecked Then
       strSql2 = Replace(strSql, GBL_PDV, GBL_PDC, , , vbTextCompare)
       RegistrarPendentes strSql2, GBL_PDC
    End If
    
    'Pedidos de Serviços a Receber
    If chkFluxo(8).value = vbChecked Then
       strSql2 = Replace(strSql, GBL_PDV, GBL_PDSR, , , vbTextCompare)
       RegistrarPendentes strSql2, GBL_PDSR
    End If
    
    'Pedidos de Serviços a Pagar
    If chkFluxo(9).value = vbChecked Then
       strSql2 = Replace(strSql, GBL_PDV, GBL_PDSP, , , vbTextCompare)
       RegistrarPendentes strSql2, GBL_PDSP
    End If
  End If
  
End Function

Private Function RegistrarPendentes(strSql As String, Tabela As String)
  Dim rstPedidos    As Object
  Dim PrimeiroBanco As Long
  
  'Se não houver um banco infomado no Pedido
  'utilizar o primeiro banco do filtro e se não houver banco no filtro
  'utilizar o primeiro banco da tabela bancos
  Screen.MousePointer = vbHourglass
  
  If txtFluxo(2).Text <> "" Then
     PrimeiroBanco = CLng(txtFluxo(2).Text)
  Else
     PrimeiroBanco = CLng(GetFieldValue("Top 1 Banco", "Bancos", NUL, 0, 0))
  End If
  
  If AbreRecordset(rstPedidos, strSql, dbOpenSnapshot) = WL_OK Then
    Do
        rstPrevisao.AddNew
        If CompraVenda(Tabela) = "Venda" Then
          rstPrevisao("PagRec") = "R"
        Else
          rstPrevisao("PagRec") = "P"
        End If
        rstPrevisao("Empresa") = GetValue(rstPedidos, "Empresa", NUL)
        rstPrevisao("Número") = GetValue(rstPedidos, "Número", ZERO)
        rstPrevisao("Tipo") = GetValue(rstPedidos, "Tipo de Registro", NUL)
        
        If GetValue(rstPedidos, "Banco", ZERO) = ZERO Then
           rstPrevisao("Banco") = PrimeiroBanco
        Else
           rstPrevisao("Banco") = GetValue(rstPedidos, "Banco")
        End If
        
        rstPrevisao("Conta") = GetValue(rstPedidos, "Conta", ZERO)
        rstPrevisao("Centro") = GetValue(rstPedidos, "Centro de Custo", ZERO)
        rstPrevisao("Vencimento") = GetValue(rstPedidos, "Vencimento", NUL)
        rstPrevisao("Valor") = GetValue(rstPedidos, "ValorResult", ZERO)
        rstPrevisao.update
        
        rstPedidos.MoveNext
    Loop Until rstPedidos.EOF
  End If
  FechaRecordset rstPedidos
  
  Screen.MousePointer = vbDefault
End Function


Private Function RatearPendentes(strSql As String, Tabela As String)
  Dim rstPedidos    As Object
  Dim PrimeiroBanco As Long
  
  'Se não houver um banco infomado no Pedido
  'utilizar o primeiro banco do filtro e se não houver banco no filtro
  'utilizar o primeiro banco da tabela bancos
  If txtFluxo(2).Text <> "" Then
     PrimeiroBanco = CLng(txtFluxo(2).Text)
  Else
     PrimeiroBanco = CLng(GetFieldValue("Top 1 Banco", "Bancos", NUL, 0, 0))
  End If
  
  If AbreRecordset(rstPedidos, strSql, dbOpenSnapshot) = WL_OK Then
    Do
      If GetValue(rstPedidos, "Condição de Pagamento", ZERO) > 0 Then
        Dim rstCondPagamento    As Object
        Dim rstParcelas         As Object
        
        If AbreRecordset(rstCondPagamento, "Select * from [Condições de Pagamento] where [Código] = " & GetValue(rstPedidos, "Condição de Pagamento", ZERO), dbOpenSnapshot) = WL_OK Then
          If AbreRecordset(rstParcelas, "Select * from Parcelas Where Condição = " & GetValue(rstCondPagamento, "Código", ZERO), dbOpenSnapshot) = WL_OK Then
            Do
            
              Dim Dias              As Integer
              Dim Porcentagem       As Double
              Dim TipoDia           As String
              Dim ApenasDiasUteis   As Boolean
              Dim QtdPar            As Integer
              Dim Vencimento        As Date
              Dim UltData           As Date       ' Última data gerada
              Dim VrParcela         As Double

              Dias = GetValue(rstParcelas, "Parcela", 0)
              Porcentagem = GetValue(rstParcelas, "Porcentagem", 0)
              TipoDia = UCase(GetValue(rstCondPagamento, "Tipo de Dia", NUL))
              ApenasDiasUteis = GetValue(rstCondPagamento, "Considerar apenas dias úteis", False)
              QtdPar = GetValue(rstCondPagamento, "Número de Parcelas", ZERO)
              
              Select Case TipoDia
                Case UCase("Dias Corridos")
                  If ApenasDiasUteis Then Dias = NumeroDiasUteisNaoUteis(GetValue(rstPedidos, "Data da Previsão", NUL), Dias)
                  Vencimento = GetValue(rstPedidos, "Data da Previsão") + Dias
                Case UCase("Fixo")
                  ' Na primeira parcela, verifico se o dia da data inicial é maior do que o dia da
                  ' primeira parcela, se for o vencimento só pode começar no mês seguinte
                  If Not EData(UltData) Then
                    If (Day(GetValue(rstPedidos, "Data da Previsão")) > Dias) Or CBool(GetValue(rstCondPagamento, "Mês Atual") = False) Then
                      UltData = DateAdd("M", 1, GetValue(rstPedidos, "Data da Previsão"))
                    Else
                      UltData = GetValue(rstPedidos, "Data da Previsão")
                    End If
                  End If
                  If Dias = ZERO Then Dias = 1
                  If (Dias = 31) And (Not EData(Dias & "/" & Month(UltData) & "/" & Year(UltData))) Then
                    Vencimento = LastDay(CDate("01" & "/" & Month(UltData) & "/" & Year(UltData)))
                  Else
                    Vencimento = CDate(Dias & "/" & Month(UltData) & "/" & Year(UltData))
                  End If
                  If ApenasDiasUteis Then
                    Dias = NumeroDiasUteisNaoUteis(FirstDay(Vencimento), Dias)
                    Vencimento = DateAdd("D", Dias, Vencimento)
                  End If
                
                Case UCase("Semanal")
                  'Quando for a primeira passagem a data inicial deverá ser a data inicial, depois somente a ULTDATA
                  Vencimento = DatadaSemana(IIf(IsEmpty(UltData), GetValue(rstPedidos, "Data da Previsão"), UltData), Dias, True, 0)
    
                Case UCase("Fora Semana")
                  'Quando for a primeira passagem o dia inicial deverá ser o Domingo para o caso
                  'do usuário informar o primeiro vencimento na segunda feira.
                  If ApenasDiasUteis Then Dias = NumeroDiasUteisNaoUteis(GetValue(rstPedidos, "Data da Previsão"), Dias)
                  If Not EData(UltData) Then
                    UltData = ForaSemana(GetValue(rstPedidos, "Data da Previsão"), Dias)
                    Vencimento = UltData
                  Else
                    Vencimento = ForaSemana(UltData, Dias)
                  End If
                
              End Select
              UltData = Vencimento
              
              If (QtdPar Mod 3) = 0 And GetValue(rstCondPagamento, "Iguais", False) Then
                VrParcela = Kin_Round(GetValue(rstPedidos, "ValorLiquido", ZERO) / QtdPar, 2)
              Else
                VrParcela = Round(CSngDef((GetValue(rstPedidos, "ValorLiquido", ZERO) * (Porcentagem / 100))), 2)
              End If
              
              rstPrevisao.AddNew
              If CompraVenda(Tabela) = "Venda" Then
                rstPrevisao("PagRec") = "R"
              Else
                rstPrevisao("PagRec") = "P"
              End If
              rstPrevisao("Empresa") = GetValue(rstPedidos, "Empresa", NUL)
              rstPrevisao("Número") = GetValue(rstPedidos, "Número", ZERO)
              rstPrevisao("Tipo") = GetValue(rstPedidos, "Tipo de Registro", NUL)
              
              If GetValue(rstPedidos, "Banco", ZERO) = ZERO Then
                 rstPrevisao("Banco") = PrimeiroBanco
              Else
                 rstPrevisao("Banco") = GetValue(rstPedidos, "Banco")
              End If
              
              rstPrevisao("Conta") = GetValue(rstPedidos, "Conta", ZERO)
              rstPrevisao("Centro") = GetValue(rstPedidos, "Centro de Custo", ZERO)
              rstPrevisao("Vencimento") = Vencimento
              rstPrevisao("Valor") = VrParcela
              rstPrevisao.update
              
              rstParcelas.MoveNext
            Loop Until rstParcelas.EOF
          End If
          FechaRecordset rstParcelas
        End If
        FechaRecordset rstCondPagamento
        
      End If
      
      rstPedidos.MoveNext
    Loop Until rstPedidos.EOF
  End If
  FechaRecordset rstPedidos

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
