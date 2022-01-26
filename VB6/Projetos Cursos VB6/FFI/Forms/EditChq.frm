VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEditChq 
   Caption         =   "Editor para Impressão de Cheques"
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   810
   ClientWidth     =   9435
   Icon            =   "EditChq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   9435
   Tag             =   "Editor"
   Begin MSComDlg.CommonDialog dlgImpChq 
      Left            =   8640
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picEditor 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      ClipControls    =   0   'False
      Height          =   4680
      Left            =   0
      ScaleHeight     =   4620
      ScaleWidth      =   9375
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   735
      Width           =   9435
      Begin VB.PictureBox hRegua 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         DragIcon        =   "EditChq.frx":030A
         Height          =   375
         Left            =   375
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   625
         TabIndex        =   8
         Top             =   0
         Width           =   9375
         Begin VB.Shape hEscala 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1800
            Top             =   90
            Visible         =   0   'False
            Width           =   5295
         End
      End
      Begin VB.PictureBox vRegua 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         DragIcon        =   "EditChq.frx":0614
         Height          =   5295
         Left            =   0
         ScaleHeight     =   353
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   7
         Top             =   0
         Width           =   375
         Begin VB.Shape vEscala 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   3375
            Left            =   105
            Top             =   480
            Visible         =   0   'False
            Width           =   180
         End
      End
      Begin VB.PictureBox picDesktop 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4935
         Left            =   360
         MouseIcon       =   "EditChq.frx":091E
         MousePointer    =   99  'Custom
         ScaleHeight     =   4935
         ScaleWidth      =   9015
         TabIndex        =   9
         Top             =   360
         Width           =   9015
         Begin VB.Shape sAlcas 
            BorderColor     =   &H80000014&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   7
            Left            =   6120
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Shape sAlcas 
            BorderColor     =   &H80000014&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   6
            Left            =   5880
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Shape sAlcas 
            BorderColor     =   &H80000014&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   5
            Left            =   5640
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Shape sAlcas 
            BorderColor     =   &H80000014&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   4
            Left            =   5400
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Shape sAlcas 
            BorderColor     =   &H80000014&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   3
            Left            =   5160
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Shape sAlcas 
            BorderColor     =   &H80000014&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   2
            Left            =   4920
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Shape sAlcas 
            BorderColor     =   &H80000014&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   1
            Left            =   4800
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Shape sAlcas 
            BorderColor     =   &H80000014&
            FillColor       =   &H8000000D&
            FillStyle       =   0  'Solid
            Height          =   90
            Index           =   0
            Left            =   4560
            Top             =   840
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   1080
            MouseIcon       =   "EditChq.frx":0C28
            TabIndex        =   20
            Top             =   3480
            Width           =   825
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BANCO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   1
            Left            =   840
            MouseIcon       =   "EditChq.frx":0F32
            TabIndex        =   19
            Top             =   2640
            Width           =   2655
         End
         Begin VB.Shape sMove 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            BorderWidth     =   2
            Height          =   375
            Left            =   2400
            Top             =   3000
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Line hLine 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  'Dot
            Index           =   0
            X1              =   1200
            X2              =   7560
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line vLine 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  'Dot
            Index           =   0
            X1              =   480
            X2              =   480
            Y1              =   360
            Y2              =   4560
         End
         Begin VB.Label lblCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Num.Banc.-Num.Chq."
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   5520
            MouseIcon       =   "EditChq.frx":123C
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   3390
            Width           =   1740
         End
         Begin VB.Label lblCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ano"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   6930
            MouseIcon       =   "EditChq.frx":1546
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   2565
            Width           =   315
         End
         Begin VB.Label lblCheque 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mês"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   5340
            MouseIcon       =   "EditChq.frx":1850
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   2565
            Width           =   1380
         End
         Begin VB.Label lblCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "São Paulo"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   3300
            MouseIcon       =   "EditChq.frx":1B5A
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   2565
            Width           =   1935
         End
         Begin VB.Label lblCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nominal"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   840
            MouseIcon       =   "EditChq.frx":1E64
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   2205
            Width           =   6420
         End
         Begin VB.Label lblCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Extenso"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   840
            MouseIcon       =   "EditChq.frx":216E
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   1920
            Width           =   6420
         End
         Begin VB.Label lblCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Extenso"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1545
            MouseIcon       =   "EditChq.frx":2478
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   1590
            Width           =   5715
         End
         Begin VB.Label lblCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Informações do Cheque"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   840
            MouseIcon       =   "EditChq.frx":2782
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   1320
            Width           =   4575
         End
         Begin VB.Label lblCheque 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "#Valor#"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5445
            MouseIcon       =   "EditChq.frx":2A8C
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Shape sCheque 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   2775
            Left            =   720
            Top             =   1080
            Width           =   6735
         End
         Begin VB.Shape sSombra 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   2775
            Left            =   840
            Top             =   1200
            Width           =   6735
         End
      End
   End
   Begin VB.PictureBox hpicEdit 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   9435
      TabIndex        =   1
      Top             =   0
      Width           =   9435
      Begin VB.Frame fraEditor 
         Caption         =   "Informações"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   9375
         Begin VB.TextBox txtBco 
            Height          =   315
            Left            =   7320
            TabIndex        =   21
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtEditor 
            Height          =   315
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "Editor"
            Text            =   "Descrição"
            ToolTipText     =   "Descrição do Banco"
            Top             =   240
            Width           =   3975
         End
         Begin VB.ComboBox cboEditor 
            Height          =   315
            Index           =   0
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Tag             =   "Editor"
            ToolTipText     =   "Número do Banco"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblEditor 
            AutoSize        =   -1  'True
            Caption         =   "&Descrição:"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   6
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lblEditor 
            AutoSize        =   -1  'True
            Caption         =   "&Modelo:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   570
         End
      End
   End
   Begin VB.Menu mnuChqArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuChqArquivoNovo 
         Caption         =   "&Novo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuChqArquivoFechar 
         Caption         =   "&Fechar"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuChqArquivoSalvar 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuChqArquivoDeletar 
         Caption         =   "&Deletar"
      End
      Begin VB.Menu mnuChqArquivoBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChqArquivoSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuExibir 
      Caption         =   "E&xibir"
      Begin VB.Menu mnuExibirMaximizar 
         Caption         =   "&Maximizar Área de Trabalho"
      End
      Begin VB.Menu mnuExibirBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExibirReguas 
         Caption         =   "&Réguas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExibirPosicao 
         Caption         =   "&Posição"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExibirZoom 
         Caption         =   "&Opções..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuExibirBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExibirAtualizar 
         Caption         =   "&Atualizar"
      End
   End
   Begin VB.Menu mnuFormatar 
      Caption         =   "&Formatar"
      Begin VB.Menu mnuFormatarFonte 
         Caption         =   "&Fonte..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFormatarCompl 
         Caption         =   "&Complementos..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuFormatarPosicao 
         Caption         =   "&Posição..."
         Shortcut        =   {F8}
      End
   End
End
Attribute VB_Name = "frmEditChq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const fNORMAL = 0                   'Fonte normal
Private Const fITALICO = 1                  'Fonte itálica
Private Const fNEGRITO = 2                  'Fonte Negrito
Private Const fNEGRITALICO = 3              'Fonte Negrito/itálico

Private Const sblMM$ = " mm"                'Quando o usuário estiver trabalhando com milímetros
Private Const sblCM$ = " cm"                'Quando estiver trabalhando com centímetros
Private Const sblPO$ = " pol"               'Quando estiver trabalhando com polegadas

Private Const ptESQUERDO% = 0               'Texto com alinhamento esquerdo
Private Const ptDIREITO% = 1                'Texto com alinhamento direito
Private Const ptCENTRO% = 2                 'Texto com alinhamento centralizado

Private Const lmNENHUMA% = 0                 'Nenhuma letra maiúscula
Private Const lmPOR_FRASE% = 1               'Somente a primeira letra da frase
Private Const lmPOR_PALAVRA% = 2             'Somente a primeira letra de cada palavra
Private Const lmTODAS% = 3                   'Todas as letras maiúsculas

Private Const lBRANCO& = &HFFFFFF
Private Const lBORDA_CLARA& = &H80000014
Private Const lBORDA_ESCURA& = &H80000010
Private Const lCONTORNO& = &H80000002

Private Const DS_DESENHA% = -1          'Usada nas rotinas de desenho
Private Const DS_APAGA% = 0

Private mrecEditor As KINRECT           'Guarda as posições do cheque
Private mtriEditor(0 To 8) As KINTRI    'Guarda a posição da tabela
Private mrstCheque As Object         'Abre a tabela
Private msngZoom As Single              'Visualização com zoom
Private mdplMouse As MOUSEPOS           'Guarda a posição do mouse em relação ao controle
Private mintlblIndice As Integer        'Índice do controle selecionado
Private mintHLineIndice As Integer      'Conta as linhas Horizontais
Private mintVLineIndice As Integer      'Conta as linhas Verticais
Private mintLinCorrente As Integer      'Linha atual selecionada
Private msmcEscala As ScaleModeConstants  'Escala de medida utilizada pelo usuário
Private msblMM As String                'Completa a localização do mouse na barra de status
Private mintAlcaIndice As Integer       'Guarda a alça que foi selecionada
'
' Variáveis dos formulários que trabalham junto com esta janela
'
Private WithEvents mfComp As fComplementos  'Formulário de complementos
Attribute mfComp.VB_VarHelpID = -1
Private WithEvents mfOpt  As fOptCheque     'Formulário de opções e configuração
Attribute mfOpt.VB_VarHelpID = -1
Private WithEvents mfPos  As fPosiciona     'Formulário de posicionamento dos controles na tela
Attribute mfPos.VB_VarHelpID = -1

' FUNCTION..: LibProc
' Objetivo..: Recebe as mensagens da LIB conforme o necessário
' Argumentos: [sFuncao]: String com a rotina que deve ser executada
'             [lFuncao]: Flag adicional das funções.
' Retorna...: True se obtiver sucesso, False se não.
' ------------------------------------------------------------------
Public Function LibProc(sFuncao As String, lFuncao As Long) As Boolean

  Select Case sFuncao
  '
  Case WL_NOVO
    mnuChqArquivoNovo_Click
    LibProc = mnuChqArquivoSalvar.Enabled
  '
  Case WL_DELETAR
    mnuChqArquivoDeletar_Click
  '
  Case WL_SALVAR
    mnuChqArquivoSalvar_Click
    LibProc = (Not mnuChqArquivoSalvar.Enabled)
  '
  Case WL_SAIR
    Unload Me
  '
  Case WL_PESQUISAR
    If (mnuChqArquivoSalvar.Enabled) Then     'Se o usuário estiver editando
      If (Not LibProc(WL_SALVAR, ZERO)) Then
        Exit Function
      End If
    End If
    PCampo "Modelos de Cheques", _
           "SELECT Número, Descrição FROM ChqModelos;", _
           PB_REGISTRO, cboEditor(0), "Número"
  '
  Case WL_MENUCLICK, WL_MENUSELECT
    LibProc = False
  '
  Case Else
    MsgFunc LoadResString(246)
  '
  End Select
  
End Function

Private Sub cboEditor_Click(Index As Integer)
  DesenhaSelecao mintlblIndice, DS_APAGA
  mintlblIndice = 0
  CarregaModelo cboEditor(0).Text
  mnuChqArquivoSalvar.Enabled = False
End Sub

Private Sub cboEditor_DropDown(Index As Integer)
  If (mnuChqArquivoSalvar.Enabled) Then
    If (MsgFunc(LoadResString(89), vbQuestion Or vbYesNo) = vbYes) Then
      mnuChqArquivoSalvar_Click
    End If
  End If
  mnuChqArquivoSalvar.Enabled = False
End Sub

Private Sub Form_Load()
Dim intItem As Integer
  '
  ' Abrindo a tabela de modelos
  '
  If (AbreRecordset(mrstCheque, "ChqModelos") = WL_NORECORD) Then
    'Exibe um modelo padrão
    '
    cboEditor(0).AddItem LoadResString(200)
    AdicionaNovo cboEditor(0).List(0)
  ElseIf (UltimoRetorno = WL_OK) Then
    ' Carregando a caixa combo com os modelos
    '
    ComboAddItem cboEditor(0), "SELECT Número FROM ChqModelos", "Número"
  Else
    ' Não foi possível abrir a tabela
    MsgFunc ResolveResString(IDS_ERROPENTABLE, resUM, "ChqModelos")
    Exit Sub
  End If
  
  'Zoom e Escala inicial
  msngZoom = 1
  msmcEscala = vbMillimeters
  msblMM = sblMM
  '
  ' Configurando o lugar das reguas
  hline(0).Y1 = -100: hline(0).Y2 = -100
  hline(0).X1 = -100: hline(0).X2 = (picEditor.Width * 2)
  vLine(0).X1 = -100: vLine(0).X2 = -100
  vLine(0).Y1 = -100: vLine(0).Y2 = (picEditor.Height * 2)
  '
  ' Carregando o formulário de opções
  Set mfOpt = New fOptCheque
  Load mfOpt
  '
  ' Carregando o formulário de Complementos
  Set mfComp = New fComplementos
  Load mfComp
  '
  ' Carrega o formulário de posicionamento sem exibí-lo
  Set mfPos = New fPosiciona
  Load mfPos
  Set mfPos.EditForm = Me       'Passa uma referência deste formulário
  
  Me.Refresh                    'Garante que o formulário seja desenhado e posicionado
  '
  ' Recarrega a variável somente com o registro que interessa
  cboEditor(0).Text = cboEditor(0).List(0)
  '
  ' Reposiciona o cheque na tela
  '
  mfComp_Aplicar
  DoEvents
  '
  ' Desabilita o comando Salvar até que o usuário altera alguma coisa
  mnuChqArquivoSalvar.Enabled = False
  '
  ' Oculta a barra de ferramentas do formulário principal que não será
  ' utilizada
  '
End Sub

' Sub DesenhaSelecao
'
' Desenha alças de seleção em torno do controle
' Argumentos: [intId]: índice do Label clicado
'             [intFuncao]: -1 para desenhar, 0 para apagar
' ---------------------------------------------------------
Public Sub DesenhaSelecao(ByVal intId As Integer, ByVal intFuncao As Integer)
Const DS_7% = 105               'Sete pixels
Const DS_1% = 15                'Um pixel

Dim ptoCanto As MOUSEPOS
Dim intConta As Integer

  If intFuncao = DS_DESENHA Then
    'Superior esquerda
    sAlcas(0).Left = (lblCheque(intId).Left - DS_7)
    sAlcas(0).Top = (lblCheque(intId).Top - DS_7)
    'Superior central
    sAlcas(1).Left = ((lblCheque(intId).Left + (lblCheque(intId).Width / 2)) - (DS_7 / 2))
    sAlcas(1).Top = sAlcas(0).Top
    'Superior direita
    sAlcas(2).Left = (lblCheque(intId).Left + lblCheque(intId).Width + DS_1)
    sAlcas(2).Top = sAlcas(0).Top
    'Central direita
    sAlcas(3).Left = sAlcas(2).Left
    sAlcas(3).Top = ((lblCheque(intId).Top + (lblCheque(intId).Height / 2)) - (DS_7 / 2))
    'Central esquerda
    sAlcas(4).Left = sAlcas(0).Left
    sAlcas(4).Top = sAlcas(3).Top
    'Inferior esquerda
    sAlcas(5).Left = sAlcas(0).Left
    sAlcas(5).Top = (lblCheque(intId).Top + lblCheque(intId).Height + DS_1)
    'Inferior central
    sAlcas(6).Left = sAlcas(1).Left
    sAlcas(6).Top = sAlcas(5).Top
    'Inferior direita
    sAlcas(7).Left = sAlcas(3).Left
    sAlcas(7).Top = sAlcas(6).Top
    '
    ' É necessário reafirmar a seleção do label aqui para sincronizar
    ' o label selecionado neste form e no formulário de posicionamento
    mintlblIndice = intId
  End If
  
  For intConta = 0 To 7
    sAlcas(intConta).Visible = (intFuncao = DS_DESENHA)
  Next intConta
  
End Sub

' Sub CarregaModelo
'
' Carraga a variável de módulo somente com o modelo que o usuário escolhe
' Argumento: [strModelo]: Número do modelo
' ------------------------------------------------------------------------
Private Sub CarregaModelo(ByVal strModelo As String)
Dim strModelos As String

  strModelos = "SELECT * FROM ChqModelos WHERE Número = " & strModelo & ";"
  
  If (AbreRecordset(mrstCheque, strModelos) = WL_OK) Then
    ' Configurando as propriedades do cheque
    
    'Projeto: # - História: # - Desenvolvimento# -  Vinicius Alexandre Elyseu (24/04/2014)
    If Not IsNull(mrstCheque("Descrição")) Then
        txtEditor.Text = mrstCheque("Descrição")
    Else
        txtEditor.Text = "Sem Descrição"
    End If
        
    lblBanco(0).Caption = txtEditor.Text
    '
    ' Tamanho do cheque
    '
    mrecEditor.sWidth = (ScaleX(mrstCheque("Largura"), vbMillimeters, vbTwips))
    mrecEditor.sHeight = (ScaleY(mrstCheque("Altura"), vbMillimeters, vbTwips))
    '
    'Guardando o tamanho dos controles em variáveis
    ' Valor
    mtriEditor(1).sLateral = ScaleX(mrstCheque("VlrPosLeft"), vbMillimeters, vbTwips)
    mtriEditor(1).sBase = ScaleY(mrstCheque("VlrPosBase"), vbMillimeters, vbTwips)
    mtriEditor(1).sLargura = ScaleX(mrstCheque("VlrPosWidth"), vbMillimeters, vbTwips)
    '
    ' Informações do Banco
    mtriEditor(0).sBase = mtriEditor(1).sBase
    mtriEditor(0).sLargura = ScaleX(mrstCheque("InfPosWidth"), vbMillimeters, vbTwips)
    mtriEditor(0).sLateral = ScaleX(mrstCheque("InfPosLeft"), vbMillimeters, vbTwips)
    '
    ' Primeira linha do extenso
    mtriEditor(2).sBase = ScaleY(mrstCheque("ExtAPosBase"), vbMillimeters, vbTwips)
    mtriEditor(2).sLateral = ScaleX(mrstCheque("ExtAPosLeft"), vbMillimeters, vbTwips)
    mtriEditor(2).sLargura = ScaleX(mrstCheque("ExtAPosWidth"), vbMillimeters, vbTwips)
    '
    ' Segunda linha do extenso
    mtriEditor(3).sBase = ScaleY(mrstCheque("ExtBPosBase"), vbMillimeters, vbTwips)
    mtriEditor(3).sLateral = ScaleX(mrstCheque("ExtBPosLeft"), vbMillimeters, vbTwips)
    mtriEditor(3).sLargura = ScaleX(mrstCheque("ExtBPosWidth"), vbMillimeters, vbTwips)
    '
    ' Linha de Nominal do cheque
    mtriEditor(4).sBase = ScaleY(mrstCheque("NomPosBase"), vbMillimeters, vbTwips)
    mtriEditor(4).sLateral = ScaleX(mrstCheque("NomPosLeft"), vbMillimeters, vbTwips)
    mtriEditor(4).sLargura = ScaleX(mrstCheque("NomPosWidth"), vbMillimeters, vbTwips)
    '
    ' Linha de Localidade
    mtriEditor(5).sBase = ScaleY(mrstCheque("LocPosBase"), vbMillimeters, vbTwips)
    mtriEditor(5).sLateral = ScaleX(mrstCheque("LocPosLeft"), vbMillimeters, vbTwips)
    mtriEditor(5).sLargura = ScaleX(mrstCheque("LocPosWidth"), vbMillimeters, vbTwips)
    '
    ' Mês
    mtriEditor(6).sBase = mtriEditor(5).sBase
    mtriEditor(6).sLateral = ScaleX(mrstCheque("MesPosLeft"), vbMillimeters, vbTwips)
    mtriEditor(6).sLargura = ScaleX(mrstCheque("MesPosWidth"), vbMillimeters, vbTwips)
    '
    ' Ano
    mtriEditor(7).sBase = mtriEditor(6).sBase
    mtriEditor(7).sLateral = ScaleX(mrstCheque("AnoPosLeft"), vbMillimeters, vbTwips)
    mtriEditor(7).sLargura = ScaleX(mrstCheque("AnoPosWidth"), vbMillimeters, vbTwips)
    '
    ' Rodapé do cheque
    mtriEditor(8).sBase = ScaleY(mrstCheque("NumBanBase"), vbMillimeters, vbTwips)
    mtriEditor(8).sLateral = ScaleX(mrstCheque("NumBanLeft"), vbMillimeters, vbTwips)
    mtriEditor(8).sLargura = ScaleX(mrstCheque("NumBanWidth"), vbMillimeters, vbTwips)
    '
    ' Definindo o tipo da fonte exibida
    On Error Resume Next
    '
    ' Nome da fonte
    picEditor.FontName = GetValue(mrstCheque, "FonteNome")
    If (err.Number) Then
      'O usuário provavelmente exclui esta fonte de sua máquina
      VBErros ResolveResString(80, resUM, GetValue(mrstCheque, "FonteNome"), _
                               resDOIS, Me.FontName)
      picEditor.FontName = Me.FontName
    End If
    '
    ' Tamanho da fonte
    picEditor.FontSize = GetValue(mrstCheque, "FonteSize", 10)
    '
    ' Estilo da fonte
    picEditor.FontItalic = (mrstCheque("FonteTipo") = fITALICO) Or (mrstCheque("FonteTipo") = fNEGRITALICO)
    picEditor.FontBold = (mrstCheque("FonteTipo") = fNEGRITO) Or (mrstCheque("FonteTipo") = fNEGRITALICO)
    '
    ' Definindo as propriedades do formulário de complementos
    With mfComp
      .ComplValor = GetValue(mrstCheque, "CaracterSeguranca")
      .ComplExt = GetValue(mrstCheque, "CaracterComplemento", NUL)
      .CharCase = GetValue(mrstCheque, "LetrasMaiusculas")
      .MesCompleto = GetValue(mrstCheque, "MesCompleto")
      .AnoCompleto = GetValue(mrstCheque, "AnoCompleto")
      .FecharValor = GetValue(mrstCheque, "FecharValor")
      .FecharExt = GetValue(mrstCheque, "FecharExtenso")
    End With
    '
    ' Define o tamanho dos controles na tela
    mfComp_Aplicar
  Else
    MsgFunc LoadResString(81)
  End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim intResposta As Integer
  
  If (mnuChqArquivoSalvar.Enabled) Then   'Se a última alteração não foi salva
    intResposta = MsgFunc(ResolveResString(IDS_QUERYSAVE, resUM, Caption), vbQuestion Or vbYesNoCancel)
    If (intResposta = vbYes) Then
      mnuChqArquivoSalvar_Click
    ElseIf (intResposta = vbCancel) Then
      Cancel = True
    End If
  End If
    
End Sub

Private Sub Form_Resize()
  '
  'Reposiciona a picEditor
  '
  If ((WindowState <> vbMinimized) And (Height > 2000)) Then
    picEditor.Height = (ScaleHeight - picEditor.Top)
  End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '
  'Fecha o formulário de posicionamento, de complementos e de opções
  '
  Unload mfPos
  Set mfPos = Nothing
  
  Unload mfComp
  Set mfComp = Nothing
  
  Unload mfOpt
  Set mfOpt = Nothing
  
  MsgBar MsgBoxCaption
  
  Set frmEditChq = Nothing
  
End Sub

Private Sub hRegua_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move a linha junto com o mouse
  If (Button = vbLeftButton) Then
    hline(0).Y1 = (ScaleY(Y, vbPixels, vbTwips) - picDesktop.Top)
    hline(0).Y2 = hline(0).Y1
  End If
End Sub

Private Sub hRegua_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ((Button = vbLeftButton) And (hline(0).Y1 > 0)) Then
    Inc mintHLineIndice
    Load hline(mintHLineIndice)
    hline(mintHLineIndice).Y1 = hline(0).Y1
    hline(mintHLineIndice).Y2 = hline(0).Y2
    hline(mintHLineIndice).ZOrder vbBringToFront
    hline(mintHLineIndice).Visible = True
    hline(0).Y1 = -100
    hline(0).Y2 = -100
  End If
End Sub

Private Sub hRegua_Paint()
Dim sngLeft As Single, sngTop As Single
Dim intWidth As Integer, sngHeight As Single
Dim sngEtapa As Single, intEtapa As Integer

  'Desenha a borda da regua
  hRegua.Line (0, 0)-((hRegua.ScaleWidth - 1), 0), lBORDA_CLARA
  hRegua.Line -((hRegua.ScaleWidth - 1), (hRegua.ScaleHeight - 1)), lBORDA_ESCURA
  hRegua.Line -(0, (hRegua.ScaleHeight - 1)), lBORDA_ESCURA
  hRegua.Line -(0, 0), lBORDA_CLARA
  
  ' Definindo a escala
  sngLeft = hEscala.Left
  hRegua.Line (sngLeft, 6)-((sngLeft + hEscala.Width), 17), lBRANCO, BF
  hRegua.Line (sngLeft, 6)-((sngLeft + hEscala.Width), 6), lBORDA_ESCURA
  hRegua.Line -((sngLeft + hEscala.Width), 17), lBORDA_CLARA
  hRegua.Line -(sngLeft, 17), lBORDA_CLARA
  hRegua.Line -(sngLeft, 6), lBORDA_ESCURA
  
' sngHeight = (hRegua.ScaleHeight)
' If sngLeft > 0 Then Stop
' sngEtapa = ScaleX(5, vbMillimeters, vbPixels)
' For intWidth = 0 To hEscala.Width Step sngEtapa
'   intEtapa = (intWidth + sngLeft)
'   sngTop = ((hEscala.Top) Or (hEscala.Height))
'   sngHeight = ((hRegua.ScaleHeight) Or (hEscala.Height + 4))
'   hRegua.Line (intEtapa, sngTop)-(intEtapa, sngHeight), &H0
' Next intWidth
    
End Sub

Private Sub lblBanco_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sngMpX As Single, sngMpY As Single
Dim strMSG As String

  sngMpX = ScaleX(((X + lblBanco(Index).Left) - sCheque.Left), vbTwips, msmcEscala)
  sngMpX = (sngMpX / msngZoom)
  sngMpY = ScaleY(((Y + lblBanco(Index).Top) - sCheque.Top), vbTwips, msmcEscala)
  sngMpY = (sngMpY / msngZoom)
  strMSG = "(X:" & Format$(sngMpX, "Standard") & " ;Y:" & Format$(sngMpY, "Standard") & ")"
  Call MsgStbKin(strMSG)
    
End Sub

Private Sub lblCheque_DblClick(Index As Integer)
  'Exibe a janela de posicionamento com exibindo o label que foi clicado
  mintlblIndice = Index
  mnuFormatarPosicao_Click
End Sub

Private Sub lblCheque_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Inicia o Drag do controle
  '
  If ((Button = vbLeftButton) And (Index > 0)) Then
    mdplMouse.sngX = X
    mdplMouse.sngY = Y
    DesenhaSelecao mintlblIndice, DS_APAGA
    mintlblIndice = Index
    sMove.Move lblCheque(Index).Left, lblCheque(Index).Top, lblCheque(Index).Width, lblCheque(Index).Height
    sMove.ZOrder vbBringToFront
    sMove.Visible = True
  End If
  
End Sub

Private Sub lblCheque_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sngMX As Single, sngMY As Single
Dim strString As String
  '
  ' Exibe a posição do mouse em relação a área de trabalho
  '
  If (Button = 0) Then      'Nehum botão pressionado
    sngMX = ScaleX(((X + lblCheque(Index).Left) - sCheque.Left), vbTwips, msmcEscala)
    sngMX = (sngMX / msngZoom)
    sngMY = ScaleY(((Y + lblCheque(Index).Top) - sCheque.Top), vbTwips, msmcEscala)
    sngMY = (sngMY / msngZoom)
    strString = "(X:" & Format$(sngMX, "Standard") & " ;Y:" & Format$(sngMY, "Standard") & ")"
    Call MsgStbKin(strString)
  ElseIf Button = 1 And Index > 0 Then
    sMove.Move ((X + lblCheque(mintlblIndice).Left) - mdplMouse.sngX), ((Y + lblCheque(mintlblIndice).Top) - mdplMouse.sngY)
  End If
    
End Sub

Private Sub lblCheque_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Coloca o label no mesmo lugar que o shape
  '
  If ((Button = vbLeftButton) And (Index > 0)) Then
    sMove.Visible = False
    lblCheque(mintlblIndice).Move sMove.Left, sMove.Top
    mtriEditor(mintlblIndice).sBase = (((lblCheque(mintlblIndice).Top + lblCheque(mintlblIndice).Height) - sCheque.Top) / msngZoom)
    mtriEditor(mintlblIndice).sLargura = (lblCheque(mintlblIndice).Width / msngZoom)
    mtriEditor(mintlblIndice).sLateral = ((lblCheque(mintlblIndice).Left - sCheque.Left) / msngZoom)
    Call SincronizaForm(mintlblIndice)
    mnuChqArquivoSalvar.Enabled = True
  End If
    
End Sub

Private Sub mfComp_Aplicar()
Const ALT_AJUSTE! = 1.3

Dim sngCont As Single, strCharComp As String
Dim iBordaTipo As Integer

  ' Configurando as primeiras variáveis
  '
  ' Zoom
  '
  msngZoom = (mfOpt.Visual / 100)
  '
  ' Unidade de Escala
  '
  msmcEscala = mfOpt.Escala
  '
  ' Tipo da borda dos labels
  '
  If mfOpt.Bordas Then
    iBordaTipo = 1
  Else
    iBordaTipo = 0
  End If
  
  picDesktop.Visible = False
  '
  ' Definindo a disposição do cheque
  sCheque.Width = (mrecEditor.sWidth * msngZoom)
  sCheque.Height = (mrecEditor.sHeight * msngZoom)
  sCheque.Left = ((picDesktop.ScaleWidth / 2) - (sCheque.Width / 2))
  sCheque.Top = ((picDesktop.ScaleHeight / 2) - (sCheque.Height / 2))
  '
  ' Desenha uma sombra para o cheque na picDesktop
  sSombra.Move (sCheque.Left + 90), (sCheque.Top + 90), sCheque.Width, sCheque.Height
  '
  ' Define o caption de todos os labels
  DefineCaption
  '
  ' Posiciona todos os controles
  For sngCont = 0 To 8
    lblCheque(sngCont).Left = ((mtriEditor(sngCont).sLateral * msngZoom) + sCheque.Left)
    lblCheque(sngCont).Height = picDesktop.TextHeight(lblCheque(sngCont).Caption) * ALT_AJUSTE
    lblCheque(sngCont).Top = ((mtriEditor(sngCont).sBase * msngZoom) + sCheque.Top) - lblCheque(sngCont).Height
    lblCheque(sngCont).Width = (mtriEditor(sngCont).sLargura * msngZoom)
    lblCheque(sngCont).BorderStyle = iBordaTipo
  Next sngCont
  '
  ' Label banco de Descrição
  lblBanco(1).Move lblCheque(0).Left, (sCheque.Top + (sCheque.Height / 2))
  lblBanco(1).FontSize = (36 * msngZoom)
  lblBanco(0).Move lblCheque(0).Left, (lblBanco(1).Top + lblBanco(1).Height + 30)
  lblBanco(0).FontSize = (9 * msngZoom)
  '
  ' Completa o caption do campo de extenso
  strCharComp = mfComp.ComplExt
  'If strCharComp <> "Nenhum" Then
  If strCharComp <> "" Then
    sngCont = ((lblCheque(2).Width - picDesktop.TextWidth(lblCheque(2).Caption))) / picDesktop.TextWidth(strCharComp)
    sngCont = KDecimais(sngCont, 0)
    lblCheque(2).Caption = lblCheque(2).Caption & KString(strCharComp, sngCont)
    '
    ' Segunda linha do extenso
    sngCont = (lblCheque(3).Width / picDesktop.TextWidth(strCharComp))
    sngCont = KDecimais(sngCont, 0)
    lblCheque(3).Caption = KString(strCharComp, sngCont)
  End If
  '
  ' Definindo o tamanho dos controles de Escala
  hEscala.Move ScaleX(sCheque.Left, vbTwips, vbPixels), hEscala.Top, ScaleX(sCheque.Width, vbTwips, vbPixels)
  vEscala.Move vEscala.Left, ScaleY((sCheque.Top + picDesktop.Top), vbTwips, vbPixels), vEscala.Width, ScaleY(sCheque.Height, vbTwips, vbPixels)
  picDesktop.Visible = True
  hRegua.Refresh
  vRegua.Refresh
  
  mnuChqArquivoSalvar.Enabled = True
    
End Sub

Private Sub mfOpt_AplicarConfig()
  mfComp_Aplicar            'Sub que altera as configurações
End Sub

Private Sub mfPos_AplicarPos(Indice As Integer)

  If (Indice > 0) Then
    mtriEditor(Indice).sBase = ScaleY(mfPos.Base, msmcEscala, vbTwips)
    mtriEditor(Indice).sLargura = ScaleX(mfPos.largura, msmcEscala, vbTwips)
    mtriEditor(Indice).sLateral = ScaleX(mfPos.Lateral, msmcEscala, vbTwips)
    mfComp_Aplicar
    Call DesenhaSelecao(Indice, DS_DESENHA)
  Else
    mrecEditor.sWidth = ScaleX(mfPos.Lateral, msmcEscala, vbTwips)
    mrecEditor.sHeight = ScaleY(mfPos.Base, msmcEscala, vbTwips)
    mfComp_Aplicar
  End If
  mnuChqArquivoSalvar.Enabled = True
  
End Sub

Private Sub mnuChqArquivoDeletar_Click()
  If cboEditor(0).ListCount = 1 Then
    MsgFunc LoadResString(90), vbExclamation
    Exit Sub
  Else
    If (MsgFunc(LoadResString(27), vbQuestion Or vbYesNo) = vbYes) Then
      mrstCheque.Delete
      mrstCheque.MoveNext
      If (mrstCheque.EOF And (Not mrstCheque.BOF)) Then
        mrstCheque.MoveFirst
        cboEditor(0).RemoveItem cboEditor(0).ListIndex
        cboEditor(0).Text = cboEditor(0).List(0)
      Else
        cboEditor(0).Text = (cboEditor(0).ListIndex + 1)
        cboEditor(0).RemoveItem (cboEditor(0).ListIndex - 1)
      End If
    End If
  End If
  mnuChqArquivoSalvar.Enabled = False
End Sub

Private Sub mnuChqArquivoFechar_Click()
  Unload Me
End Sub

Private Sub mnuChqArquivoNovo_Click()
Dim strCodigo As String
Dim intCodigo As Integer
  '
  'Cria um novo formulário e exibe os bancos disponíveis para se criar o modelo
  '
  If mnuChqArquivoSalvar.Enabled Then
    If (MsgFunc(LoadResString(10), vbQuestion Or vbYesNo) = vbYes) Then
      mnuChqArquivoSalvar_Click
    End If
  End If
  '
  ' Abre a janela de pesquisa para o usuário retornar um Banco
  '
  strCodigo = "SELECT Banco, Nome, Conta, [Nome Conta], Câmara, Previsão " & _
              "FROM Bancos;"
  If (PCampo("Bancos", strCodigo, PB_CAMPO, txtBco, "Câmara")) Then
    ' O banco escolhido pelo usuário é colocado na caixa de texto txtBco, invisível ao
    ' usuário. Ela me serve apenas para retornar o valor escolhido pelo usuário.
    ' Verificando se já existe um modelo para este banco
    '
    intCodigo = IndexOf(txtBco.Text, cboEditor(0))
    If (intCodigo <> NENHUM) Then
      cboEditor(0).ListIndex = intCodigo
    Else
      cboEditor(0).AddItem txtBco.Text      'Adiciona o novo código a ComboBox
      AdicionaNovo txtBco.Text
      cboEditor(0).Text = txtBco.Text
      txtEditor.Text = GetFieldValue("Nome", "Bancos", "Câmara= " & txtBco.Text)   'Nome do Banco
      mnuChqArquivoSalvar.Enabled = True
    End If
    txtBco.Text = NUL                       'Reseta a caixa de texto, não é mais necessária
  End If
  
End Sub

Private Sub mnuChqArquivoSair_Click()
  Unload Me
End Sub

Private Sub mnuChqArquivoSalvar_Click()

  SetPtr vbHourglass
  
  If TypeOf mrstCheque Is dao.Recordset Then mrstCheque.Edit
  mrstCheque("Número") = CLngDef(cboEditor(0).Text)
  mrstCheque("Descrição") = txtEditor.Text
  '
  ' Tamanho do Cheque
  mrstCheque("Largura") = ScaleX(mrecEditor.sWidth, vbTwips, vbMillimeters)
  mrstCheque("Altura") = ScaleY(mrecEditor.sHeight, vbTwips, vbMillimeters)
  '
  ' Tamanho e posição do campo Valor
  mrstCheque("VlrPosLeft") = ScaleX(mtriEditor(1).sLateral, vbTwips, vbMillimeters)
  mrstCheque("VlrPosWidth") = ScaleX(mtriEditor(1).sLargura, vbTwips, vbMillimeters)
  mrstCheque("VlrPosBase") = ScaleY(mtriEditor(1).sBase, vbTwips, vbMillimeters)
  '
  ' Tamanho e posição do campo Informação
  mrstCheque("InfPosLeft") = ScaleX(mtriEditor(3).sLateral, vbTwips, vbMillimeters)

  '
  ' Tamanho e posição dos campos de Extenso
  mrstCheque("ExtAPosLeft") = ScaleX(mtriEditor(2).sLateral, vbTwips, vbMillimeters)
  mrstCheque("ExtAPosWidth") = ScaleX(mtriEditor(2).sLargura, vbTwips, vbMillimeters)
  mrstCheque("ExtAPosBase") = ScaleY(mtriEditor(2).sBase, vbTwips, vbMillimeters)
  mrstCheque("ExtBPosLeft") = ScaleX(mtriEditor(3).sLateral, vbTwips, vbMillimeters)
  mrstCheque("ExtBPosWidth") = ScaleX(mtriEditor(3).sLargura, vbTwips, vbMillimeters)
  mrstCheque("ExtBPosBase") = ScaleY(mtriEditor(3).sBase, vbTwips, vbMillimeters)
  '
  ' Tamanho e posição do campo Nominal
  mrstCheque("NomPosBase") = ScaleY(mtriEditor(4).sBase, vbTwips, vbMillimeters)
  mrstCheque("NomPosLeft") = ScaleX(mtriEditor(4).sLateral, vbTwips, vbMillimeters)
  mrstCheque("NomPosWidth") = ScaleX(mtriEditor(4).sLargura, vbTwips, vbMillimeters)
  '
  ' Tamanho e posição dos campos Local, Mês e Ano
  mrstCheque("LocPosBase") = ScaleY(mtriEditor(5).sBase, vbTwips, vbMillimeters)
  mrstCheque("LocPosLeft") = ScaleX(mtriEditor(5).sLateral, vbTwips, vbMillimeters)
  mrstCheque("LocPosWidth") = ScaleX(mtriEditor(5).sLargura, vbTwips, vbMillimeters)
  mrstCheque("MesPosLeft") = ScaleX(mtriEditor(6).sLateral, vbTwips, vbMillimeters)
  mrstCheque("MesPosWidth") = ScaleX(mtriEditor(6).sLargura, vbTwips, vbMillimeters)
  mrstCheque("AnoPosLeft") = ScaleX(mtriEditor(7).sLateral, vbTwips, vbMillimeters)
  mrstCheque("AnoPosWidth") = ScaleX(mtriEditor(7).sLargura, vbTwips, vbMillimeters)
  '
  ' Tamanho e posição do campo Adicional
  mrstCheque("NumBanBase") = ScaleY(mtriEditor(8).sBase, vbTwips, vbMillimeters)
  mrstCheque("NumBanLeft") = ScaleX(mtriEditor(8).sLateral, vbTwips, vbMillimeters)
  mrstCheque("NumBanWidth") = ScaleX(mtriEditor(8).sLargura, vbTwips, vbMillimeters)
  '
  ' Propriedades da Fonte
  mrstCheque("FonteNome") = picEditor.FontName
  mrstCheque("FonteSize") = picEditor.FontSize
  
  If ((Not picEditor.FontBold) And (Not picEditor.FontItalic)) Then
    mrstCheque("FonteTipo") = fNORMAL
  ElseIf (picEditor.FontBold And (Not picEditor.FontItalic)) Then
    mrstCheque("FonteTipo") = fNEGRITO
  ElseIf ((Not picEditor.FontBold) And picEditor.FontItalic) Then
    mrstCheque("FonteTipo") = fITALICO
  ElseIf (picEditor.FontBold And picEditor.FontItalic) Then
    mrstCheque("FonteTipo") = fNEGRITALICO
  End If
  '
  ' Outras configurações
  mrstCheque("LetrasMaiusculas") = mfComp.CharCase
  mrstCheque("CaracterSeguranca") = mfComp.ComplValor
  mrstCheque("CaracterComplemento") = mfComp.ComplExt
  mrstCheque("MesCompleto") = mfComp.MesCompleto
  mrstCheque("AnoCompleto") = mfComp.AnoCompleto
  mrstCheque("FecharValor") = mfComp.FecharValor
  mrstCheque("FecharExtenso") = mfComp.FecharExt
  
  mrstCheque.update
  mnuChqArquivoSalvar.Enabled = False
  SetPtr vbDefault
  
End Sub

Private Sub mnuExibirAtualizar_Click()
  Call mfComp_Aplicar  'Apenas repreenche o cheque
End Sub

Private Sub mnuExibirMaximizar_Click()
  '
  ' Oculta o Frame "Informações"
  mnuExibirMaximizar.Checked = (Not mnuExibirMaximizar.Checked)
  hpicEdit.Visible = (Not mnuExibirMaximizar.Checked)
  
  If hpicEdit.Visible Then
    picEditor.Move 0, hpicEdit.Height, ScaleWidth, (ScaleHeight - hpicEdit.Height)
  Else
    picEditor.Move 0, 0, ScaleWidth, ScaleHeight
  End If
  
End Sub

Private Sub mnuExibirPosicao_Click()
  mnuExibirPosicao.Checked = (Not mnuExibirPosicao.Checked)
End Sub

Private Sub mnuExibirReguas_Click()
  mnuExibirReguas.Checked = (Not mnuExibirReguas.Checked)
  hRegua.Visible = mnuExibirReguas.Checked
  vRegua.Visible = mnuExibirReguas.Checked
  If hRegua.Visible Then
    picDesktop.Move vRegua.Width, hRegua.Height, picEditor.ScaleWidth, picEditor.ScaleHeight
  Else
    picDesktop.Move 0, 0, picEditor.ScaleWidth, picEditor.ScaleHeight
  End If
End Sub

Private Sub mnuExibirZoom_Click()
  mfOpt.Show vbModal
End Sub

Private Sub mnuFormatarCompl_Click()
  mfComp.Exibe  'Exibe o formulário de complementos previamente carregado
End Sub

Private Sub mnuFormatarFonte_Click()

  With dlgImpChq
    .FontName = picEditor.FontName
    .FontSize = picEditor.FontSize
    .FontBold = picEditor.FontBold
    .FontItalic = picEditor.FontItalic
    .Flags = CF_BOTH Or CF_LIMITSIZE Or CF_FORCEFONTEXIST
    .Min = 8
    .Max = 36
  End With

  On Error Resume Next
  dlgImpChq.ShowFont
  If (err().Number) Then
    err().Clear
  Else
    picEditor.FontName = dlgImpChq.FontName
    picEditor.FontSize = dlgImpChq.FontSize
    picEditor.FontBold = dlgImpChq.FontBold
    picEditor.FontItalic = dlgImpChq.FontItalic
    mfComp_Aplicar
    mnuChqArquivoSalvar.Enabled = True
  End If
  
End Sub

Private Sub mnuFormatarPosicao_Click()
  '
  ' Exibindo o formulário
  SincronizaForm mintlblIndice
  mfPos.ExibeJanela
  
End Sub

Private Sub picDesktop_DblClick()
  'Exibe a janela de posicionamento com as medidas do cheque
  mintlblIndice = 0
  mnuFormatarPosicao_Click
End Sub

Private Sub picDesktop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim intL As Integer

  If ((Button = vbLeftButton) And (picDesktop.MousePointer = vbSizeNS)) Then
    For intL = 1 To mintHLineIndice
      If ((Y > (hline(intL).Y1 - 30)) And (Y < (hline(intL).Y2 + 30))) Then
        mintLinCorrente = hline(intL).Index
        Exit Sub
      End If
    Next intL
  ElseIf ((Button = vbLeftButton) And (picDesktop.MousePointer = vbSizeWE)) Then
    For intL = 1 To mintVLineIndice
      If ((X > (vLine(intL).X1 - 30)) And (X < (vLine(intL).X2 + 30))) Then
        mintLinCorrente = vLine(intL).Index
        Exit Sub
      End If
    Next intL
    
    If (((X > sAlcas(3).Left) And (X < (sAlcas(3).Width + sAlcas(3).Left))) And ((Y > sAlcas(3).Top) And (Y < (sAlcas(3).Height + sAlcas(3).Top)))) Then
      mintAlcaIndice = 3
      sMove.Move lblCheque(mintlblIndice).Left, lblCheque(mintlblIndice).Top, lblCheque(mintlblIndice).Width, lblCheque(mintlblIndice).Height
      sMove.ZOrder vbBringToFront
      Call DesenhaSelecao(mintlblIndice, DS_APAGA)
      sMove.Visible = True
      Exit Sub
    ElseIf (((X > sAlcas(4).Left) And (X < (sAlcas(4).Width + sAlcas(4).Left))) And ((Y > sAlcas(4).Top) And (Y < (sAlcas(4).Height + sAlcas(4).Top)))) Then
      mintAlcaIndice = 4
      sMove.Move lblCheque(mintlblIndice).Left, lblCheque(mintlblIndice).Top, lblCheque(mintlblIndice).Width, lblCheque(mintlblIndice).Height
      sMove.ZOrder vbBringToFront
      Call DesenhaSelecao(mintlblIndice, DS_APAGA)
      sMove.Visible = True
      Exit Sub
    End If
  Else
    'Tira a seleção dos Labels
    DesenhaSelecao mintlblIndice, DS_APAGA
    mintlblIndice = 0
  End If
  
End Sub

Private Sub picDesktop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sngMouseX As Single, sngMouseY As Single
Dim strMsgBar As String, intR As Integer
Dim sngControlaTam As Single

  'Exibe as coordenadas para a localização do mouse
  sngMouseX = ScaleX((X - sCheque.Left), vbTwips, msmcEscala)
  sngMouseX = (sngMouseX / msngZoom)
  sngMouseY = ScaleY((Y - sCheque.Top), vbTwips, msmcEscala)
  sngMouseY = (sngMouseY / msngZoom)
  strMsgBar = "(X:" & Format$(sngMouseX, "Standard") & " ;Y:" & Format$(sngMouseY, "Standard") & ")"
  MsgStbKin strMsgBar
      
  If (Button = 0) Then
    For intR = 1 To mintVLineIndice
      If ((X > (vLine(intR).X1 - 30)) And (X < (vLine(intR).X2 + 30))) Then
        picDesktop.MousePointer = vbSizeWE
        Exit Sub
      End If
    Next intR
    
    For intR = 1 To mintHLineIndice
      If ((Y > (hline(intR).Y1 - 30)) And (Y < (hline(intR).Y2 + 30))) Then
        picDesktop.MousePointer = vbSizeNS
        Exit Sub
      End If
    Next intR
    
    If (sAlcas(0).Visible) Then
      If ((X > sAlcas(3).Left And X < (sAlcas(3).Width + sAlcas(3).Left)) And (Y > sAlcas(3).Top And Y < (sAlcas(3).Height + sAlcas(3).Top))) Then
        picDesktop.MousePointer = vbSizeWE
        Exit Sub
      ElseIf ((X > sAlcas(4).Left And X < (sAlcas(4).Width + sAlcas(4).Left)) And (Y > sAlcas(4).Top And Y < (sAlcas(4).Height + sAlcas(4).Top))) Then
        picDesktop.MousePointer = vbSizeWE
        Exit Sub
      End If
    End If
    picDesktop.MousePointer = vbCustom
    
  ElseIf (Button = 1) Then
    If (picDesktop.MousePointer = vbSizeNS) Then
      hline(mintLinCorrente).Y1 = Y
      hline(mintLinCorrente).Y2 = Y
      MsgStbKin Format$(sngMouseY, "Standard") & msblMM
    ElseIf (picDesktop.MousePointer = vbSizeWE) Then
    
      If (sMove.Visible) Then
        If (mintAlcaIndice = 3) Then
          sngControlaTam = ((X - lblCheque(mintlblIndice).Left) - 45)
          If (sngControlaTam < 105) Then Exit Sub
          sMove.Width = sngControlaTam
          MsgStbKin Format$(sngControlaTam, "Standard") & msblMM
        ElseIf (mintAlcaIndice = 4) Then
          sngControlaTam = ((lblCheque(mintlblIndice).Left - X) + lblCheque(mintlblIndice).Width)
          If (sngControlaTam < 105) Then Exit Sub
          sMove.Move X, sMove.Top, sngControlaTam
          MsgStbKin Format$(sngMouseX, "Standard") & msblMM
        End If
      Else
        vLine(mintLinCorrente).X1 = X
        vLine(mintLinCorrente).X2 = X
        MsgStbKin Format$(sngMouseX, "Standard") & msblMM
      End If
      
    End If
  End If
  
End Sub

' Sub DefineCaption
'
' Define o tipo da fonte dos labels e o caption de cada um
' ----------------------------------------------------------
Private Sub DefineCaption()
Dim intLabel As Integer
Dim strRetorno As String
Dim strFechaFim As String
    
  'Redefinindo a fonte de exibição
  picDesktop.FontName = picEditor.FontName
  picDesktop.FontBold = picEditor.FontBold
  picDesktop.FontItalic = picEditor.FontItalic
  picDesktop.FontSize = (picEditor.FontSize * msngZoom)
  
  For intLabel = 0 To 8
    Set lblCheque(intLabel).Font = picDesktop.Font
  Next intLabel
  '
  ' Label de informações
  intLabel = mfComp.CharCase
  strRetorno = KeybUCase(LoadResString(84), intLabel)
  lblCheque(0).Caption = strRetorno
  strRetorno = NUL
  '
  ' Valor
  '
  If (mfComp.FecharValor) Then
    strRetorno = "("
    strFechaFim = ")"
  End If
  strRetorno = strRetorno & KeybUCase(LoadResString(85), intLabel) & strFechaFim
  
  'If (mfComp.ComplValor <> "Nenhum") Then
  If (mfComp.ComplValor <> "") Then
    strRetorno = mfComp.ComplValor & strRetorno & mfComp.ComplValor
  End If
  lblCheque(1).Caption = strRetorno
  strRetorno = NUL
  strFechaFim = NUL
  '
  ' Primeira linha do Extenso
  If (mfComp.FecharExt) Then
    strRetorno = "("
    strFechaFim = ")"
  End If
  strRetorno = strRetorno & KeybUCase(LoadResString(86), intLabel) & strFechaFim
  lblCheque(2).Caption = strRetorno
  strRetorno = NUL
  strFechaFim = NUL
  '
  ' Segunda linha do extenso sempre sai desta rotina nula
  lblCheque(3).Caption = NUL
  '
  ' Nominal
  lblCheque(4).Caption = LoadResString(87)
  '
  ' Local
  strRetorno = CidadePadrao()
  strRetorno = strRetorno & "   " & CStr(Day(Date))
  lblCheque(5).Caption = strRetorno
  '
  ' Mês
  If mfComp.MesCompleto Then
    strRetorno = MesExt(Date)
  Else
    strRetorno = MesExt(Date, 3)
  End If
  lblCheque(6).Caption = KeybUCase(strRetorno, PorPalavra)
  '
  ' Ano
  If mfComp.AnoCompleto Then
    strRetorno = Format$(Date, "yyyy")
  Else
    strRetorno = Format$(Date, "yy")
  End If
  lblCheque(7).Caption = strRetorno
  '
  ' Informações do cheque
  lblCheque(8).Caption = LoadResString(88)
  
End Sub

Private Sub picDesktop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If ((picDesktop.MousePointer = vbSizeWE) And sMove.Visible) Then
    lblCheque(mintlblIndice).Move sMove.Left, sMove.Top, sMove.Width, sMove.Height
    mtriEditor(mintlblIndice).sBase = (((lblCheque(mintlblIndice).Top + lblCheque(mintlblIndice).Height) - sCheque.Top) / msngZoom)
    mtriEditor(mintlblIndice).sLargura = (lblCheque(mintlblIndice).Width / msngZoom)
    mtriEditor(mintlblIndice).sLateral = ((lblCheque(mintlblIndice).Left - sCheque.Left) / msngZoom)
    sMove.Visible = False
    Call DesenhaSelecao(mintlblIndice, DS_DESENHA)
    mnuChqArquivoSalvar.Enabled = True
  End If
  
End Sub

Private Sub picEditor_Resize()
  ' Redefine o tamanho da área de trabalho do usuário e as réguas
  vRegua.Height = picEditor.ScaleHeight
  hRegua.Width = picEditor.ScaleWidth
  picDesktop.Width = picEditor.ScaleWidth - picDesktop.Left
  picDesktop.Height = picEditor.ScaleHeight - picDesktop.Top
End Sub

Private Sub txtEditor_Change()
  mnuChqArquivoSalvar.Enabled = True
End Sub

Private Sub vRegua_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Move a regua até onde precisar
  If (Button = vbLeftButton) Then
    vLine(0).X1 = (ScaleX(X, vbPixels, vbTwips) - picDesktop.Left)
    vLine(0).X2 = vLine(0).X1
  End If
End Sub

Private Sub vRegua_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ((Button = vbLeftButton) And (vLine(0).X1 > 0)) Then
    Inc mintVLineIndice
    Load vLine(mintVLineIndice)
    vLine(mintVLineIndice).X1 = vLine(0).X1
    vLine(mintVLineIndice).X2 = vLine(0).X2
    vLine(mintVLineIndice).ZOrder vbBringToFront
    vLine(mintVLineIndice).Visible = True
    vLine(0).X1 = -100
    vLine(0).X2 = -100
  End If
End Sub

Private Sub vRegua_Paint()
Dim sngTop As Single
    
  sngTop = ScaleY(hRegua.Height, vbTwips, vbPixels)
  ' Borda da regua
  vRegua.Line (0, sngTop)-((vRegua.ScaleWidth - 1), sngTop), lBORDA_CLARA
  vRegua.Line -((vRegua.ScaleWidth - 1), (vRegua.ScaleHeight - 1)), lBORDA_ESCURA
  vRegua.Line -(0, (vRegua.ScaleHeight - 1)), lBORDA_ESCURA
  vRegua.Line -(0, sngTop), lBORDA_CLARA
  '
  ' Símbolo do canto esquerdo
  vRegua.Line (0, 0)-((vRegua.ScaleWidth - 1), 0), lBORDA_CLARA
  vRegua.Line -((vRegua.ScaleWidth - 1), (sngTop - 1)), lBORDA_ESCURA
  vRegua.Line -(0, (sngTop - 1)), lBORDA_ESCURA
  vRegua.Line -(0, 0), lBORDA_CLARA
  vRegua.Line (4, 4)-((vRegua.ScaleWidth - 5), (sngTop - 5)), lBRANCO, BF
  vRegua.Line (4, 4)-((vRegua.ScaleWidth - 5), 4), lBORDA_ESCURA
  vRegua.Line -((vRegua.ScaleWidth - 5), (sngTop - 5)), lBORDA_CLARA
  vRegua.Line -(4, (hRegua.ScaleHeight - 5)), lBORDA_CLARA
  vRegua.Line -(4, 4), lBORDA_ESCURA
  '
  ' Escala
  vRegua.Line (7, vEscala.Top)-(17, (vEscala.Height + vEscala.Top)), lBRANCO, BF
  vRegua.Line (7, vEscala.Top)-(17, vEscala.Top), lBORDA_ESCURA
  vRegua.Line -(17, (vEscala.Top + vEscala.Height)), lBORDA_CLARA
  vRegua.Line -(7, (vEscala.Top + vEscala.Height)), lBORDA_CLARA
  vRegua.Line -(7, vEscala.Top), lBORDA_ESCURA
    
End Sub

' Sub MsgStbKin
'
' Exibe mensagens na barra de status do sistema
' Argumentos: [strMensagem]: Mensagem a ser exibida.
' -----------------------------------------------------
Private Sub MsgStbKin(ByVal strMensagem As String)
  If mnuExibirPosicao.Checked Then
    SimpleMsgBar strMensagem
  End If
End Sub

' Sub SincronizaForm
'
' Sincroniza o form de posicionamento para refletir o mesmo estado
' de seleção do principal.
' Argumento: [intIndiceCtl]: índice do controle selecionado.
' -------------------------------------------------------------------
Private Sub SincronizaForm(ByVal intIndiceCtl As Integer)
  'Configura as propriedades do formulário
  mfPos.Zoom = msngZoom
  mfPos.UserScale = msmcEscala
  mfPos.AtualizaPos intIndiceCtl
End Sub

' Sub AdicionaNovo
'
' Utilizada quando o usuário adiciona um novo modelo aos existentes
' exibe valores padrão para o mesmo.
' Argumento: [strChave]: A chave da tabela, se vier nula a chave exibida
'                        será do padrão
' --------------------------------------------------------------------
Private Sub AdicionaNovo(ByVal strChave As String)
Dim intCampos As Integer

  mrstCheque.AddNew
  
  If (Len(strChave)) Then
    mrstCheque(0).value = strChave
    mrstCheque(1).value = NUL
  Else
    mrstCheque(0).value = val(LoadResString(200))
    mrstCheque(1).value = LoadResString(201)
  End If
  
  For intCampos = 2 To 37
    mrstCheque(intCampos).value = LoadResString(intCampos + 200)
  Next intCampos
  
  mrstCheque.update
    
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
