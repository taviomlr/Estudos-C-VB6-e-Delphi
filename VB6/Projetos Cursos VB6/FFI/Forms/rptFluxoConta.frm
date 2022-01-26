VERSION 5.00
Begin VB.Form frptFluxoConta 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fluxo Semanal por Conta e Grupo"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "rptFluxoConta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
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
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   3120
      Width           =   5175
      Begin VB.ComboBox cboFluxoConta 
         Height          =   315
         Index           =   2
         ItemData        =   "rptFluxoConta.frx":0C42
         Left            =   1200
         List            =   "rptFluxoConta.frx":0C44
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "&Conciliados:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdFluxoConta 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   26
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdFluxoConta 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   25
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdFluxoConta 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   24
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame fraFluxoConta 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtFluxoConta 
         DataField       =   "C�digo"
         Height          =   315
         Index           =   5
         Left            =   1200
         TabIndex        =   16
         Tag             =   "Grupo"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox cboFluxoConta 
         Height          =   315
         Index           =   1
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2760
         Width           =   1695
      End
      Begin VB.ComboBox cboFluxoConta 
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtFluxoConta 
         DataField       =   "C�digo"
         Height          =   315
         Index           =   4
         Left            =   1200
         TabIndex        =   13
         Tag             =   "Grupo"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtFluxoConta 
         DataField       =   "C�digo"
         Height          =   315
         Index           =   3
         Left            =   1200
         TabIndex        =   10
         Tag             =   "Grupo"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtFluxoConta 
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtFluxoConta 
         DataField       =   "Banco"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Tag             =   "Bancos"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtFluxoConta 
         DataField       =   "Banco"
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Tag             =   "Bancos"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "&Moeda:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label lblFlxDesc 
         Caption         =   "lblFlxDesc(5)"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   17
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "P&agamento:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblFlxDesc 
         Caption         =   "lblFlxDesc(4)"
         Height          =   195
         Index           =   4
         Left            =   2520
         TabIndex        =   14
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "Grupo F&inal:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblFlxDesc 
         Caption         =   "lblFlxDesc(3)"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   11
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "&Grupo Inicial:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label lblFlxDesc 
         Caption         =   "lblFlxDesc(2)"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "&Data Inicial:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lblFlxDesc 
         Caption         =   "lblFlxDesc(1)"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblFlxDesc 
         Caption         =   "lblFlxDesc(0)"
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "Banco &Final:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblFluxoConta 
         AutoSize        =   -1  'True
         Caption         =   "Banco &Inicial:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   960
      End
   End
End
Attribute VB_Name = "frptFluxoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dblCotacao    As Double
Private mbolCancelou As Boolean       '// Verifica se o usu�rio cancelou a impress�o

' EVENT.....: cmdFluxoConta_Click
' Objetivo..: Executa as fun��es referentes aos bot�es da janela
' -----------------------------------------------------------------
Private Sub cmdFluxoConta_Click(Index As Integer)
  Select Case (Index)
    Case 0, 1                 '// Bot�es Visualizar ou Imprimir
      cmdFluxoConta(0).Enabled = False
      cmdFluxoConta(1).Enabled = False
      cmdFluxoConta(2).Caption = LoadResString(IDS_CANCELAR)

      Call FiltroFluxoConta(IIf((Index), wrToPrinter, wrToWindow))

      cmdFluxoConta(0).Enabled = True
      cmdFluxoConta(1).Enabled = True
      cmdFluxoConta(2).Caption = LoadResString(IDS_FECHAR)

    Case 2
      If (cmdFluxoConta(0).Enabled) Then
        Unload Me
      Else
        mbolCancelou = True
        SimpleMsgBar LoadResString(171) & LoadResString(14)
        DoEvents
      End If
  End Select
End Sub

' EVENT.....: Form_Load
' Objetivo..: Configura os controles na abertura do formul�rio
' ------------------------------------------------------------------
Private Sub Form_Load()

  '// Obtendo o MaxLenght dos campos conforme a estrutura da tabela

  ConfigCampos Me, "Bancos", "Bancos"     '// Banco inicial e final
  ConfigCampos Me, "Grupos", "Grupo"      '// Grupo inicial e final

  '// Carrega as op��es da ComboBox Tipo
  ComboAddItem cboFluxoConta(0), "SELECT * FROM Op��es WHERE Rotina = '" & _
                                 OPT_LANCAMENTOS & "';", "Texto"
  cboFluxoConta(0).AddItem "Todos"
  cboFluxoConta(0).ItemData(cboFluxoConta(0).NewIndex) = 1
  cboFluxoConta(0).ListIndex = cboFluxoConta(0).ListCount - 1

  '// Carrega as op��es da ComboBox Pagamento
  LoadResOptions 1032, cboFluxoConta(1), True, 2    '// 2 == Todos

  '// Carregando valores padr�o dos campos da janela
  txtFluxoConta(0).Text = CStr(MinValue("Banco", "Bancos", NUL))
  txtFluxoConta(1).Text = CStr(MaxValue("Banco", "Bancos", NUL))
  txtFluxoConta(2).Text = DataToStr(DateAdd(DD_SEMANA, -1, Date))
  txtFluxoConta(3).Text = CStr(MinValue("C�digo", "Grupos", NUL))
  txtFluxoConta(4).Text = CStr(MaxValue("C�digo", "Grupos", NUL))
  
  cboFluxoConta(2).AddItem "Todos"
  cboFluxoConta(2).AddItem "Sim"
  cboFluxoConta(2).AddItem "N�o"
  cboFluxoConta(2).Text = "Todos"
    
  CenterForm Me
    
  lblFlxDesc(5).Caption = NUL
End Sub

' EVENT.....: Form_Unload
' Objetivo..: Descarregar as vari�veis utilizada pela janela
' ---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
  Set frptFluxoConta = Nothing
End Sub

' EVENT.....: txtFluxoConta_Change
' Objetivo..: Exibe as descri��es referentes a cada campo nos Labels
'             correspondentes.
' ---------------------------------------------------------------------
Private Sub txtFluxoConta_Change(Index As Integer)
  Select Case (Index)
    Case 0, 1                       '// Banco Inicial e Final
      GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & _
                    CStr(txtFluxoConta(Index).Text), lblFlxDesc(Index)

    Case 2                          '// Data Inicial
      If (EData(txtFluxoConta(2).Text)) Then
        lblFlxDesc(2).Caption = DataToStr(DateAdd(DD_DIA, 5, txtFluxoConta(2).Text))
      Else
        lblFlxDesc(2).Caption = NUL
      End If

    Case 3, 4                       '// Grupo Inicial e Final
      GetAssocValue "SELECT Descri��o FROM Grupos WHERE C�digo = " & _
                    CStr(txtFluxoConta(Index).Text), lblFlxDesc(Index)
    
    Case 5                          '//Moeda
      GetAssocValue "SELECT Descri��o, Moeda FROM Moedas WHERE Moeda = '" & txtFluxoConta(5) & "'", _
                      lblFlxDesc(5), txtFluxoConta(5)
  End Select
End Sub

' EVENT.....: txtFluxoConta_GotFocus
' Objetivo..: Seleciona o conte�do do controle e exibe mensagens informativas
'             na barra de status do programa.
' ---------------------------------------------------------------------
Private Sub txtFluxoConta_GotFocus(Index As Integer)
  Selecione txtFluxoConta(Index)
  Select Case (Index)
    Case 0: MsgBar "C�digo do Banco Inicial"
    Case 1: MsgBar "C�digo do Banco Final"
    Case 2: MsgBar "Data Inicial do Per�odo"
    Case 3: MsgBar "C�digo do Grupo Inicial"
    Case 4: MsgBar "C�digo do Grupo Final"
  End Select
End Sub

' EVENT.....: txtFluxoConta_KeyDown
' Objetivo..: Abre a caixa de pesquisa para determinados campos.
' -------------------------------------------------------------------
Private Sub txtFluxoConta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If ((Shift = ZERO) And (KeyCode = vbKeyPageDown)) Then
    Select Case (Index)
      Case 0, 1                   '// Bancos
        PCampo "Bancos", "Bancos", PB_CAMPO, txtFluxoConta(Index), "Banco"
      Case 3, 4                   '// Grupos
        PCampo "Grupos", "Grupos", PB_CAMPO, txtFluxoConta(Index), "C�digo"
      Case 5                      '// Moeda
        PCampo "Moedas e �ndices", "Moedas", PB_CAMPO, txtFluxoConta(5), "Moeda"
    End Select
  End If
End Sub

' EVENT.....: txtFluxoConta_KeyPress
' Objetivo..: Mapear as teclas pressionadas para cada campo.
' --------------------------------------------------------------------
Private Sub txtFluxoConta_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case (Index)
    Case 0                    '// Banco Inicial
      SetMascara KeyAscii, txtFluxoConta(0).SelStart, fMask("Bancos", "Banco")
      
    Case 1                    '// Banco Final
      SetMascara KeyAscii, txtFluxoConta(1).SelStart, fMask("Bancos", "Banco"), txtFluxoConta(0).hWnd
      
    Case 2                    '// Data Inicial
      SetMascara KeyAscii, txtFluxoConta(2).SelStart, MASK_DATA
      
    Case 3                    '// Grupo Inicial
      SetMascara KeyAscii, txtFluxoConta(3).SelStart, fMask("Grupos", "C�digo")
      
    Case 4                    '// Grupo Final
      SetMascara KeyAscii, txtFluxoConta(4).SelStart, fMask("Grupos", "C�digo"), txtFluxoConta(3).hWnd
  End Select
End Sub

' SUB.......: FiltroFluxoConta
' Objetivo..: Verifica os dados digitados pelo usu�rio e cria os filtros
'             de grupo e conta.
' Argumento.: [pdeDestino]: Destino da impress�o.
' ------------------------------------------------------------------------
Private Sub FiltroFluxoConta(pdeDestino As PrintDestinoEnum)
Dim rstContas As Object           '// Recordset com os grupos e contas selecionadas
Dim strContas As String           '// Instru��o SELECT dos grupos
Dim dtPer(1)  As Date             '// 0 == Data Inicial, 1 == Data Final
Dim lTmpIni   As Long             '// C�digo do Banco ou Grupo Inicial
Dim lTmpFim   As Long             '// C�digo do Banco ou Grupo Final
Dim dtInicial As Date             '// Data Inicial
Dim dtFinal   As Date             '// Data Final
Dim rstTemp   As Object           '// Recordset auxiliar para a impress�o
Dim cMov(5)   As Currency         '// Acumula o movimento di�rio das contas no per�odo
Dim sWhere(5) As String           '// Matriz com as instru��es e compara��o:
                                  '// �ndice 0: Aplica��es recebidas
                                  '// �ndice 1: Aplica��es pagas
                                  '// �ndice 2: Tranf. Banc�ria com Bancos de Origem
                                  '// �ndice 3: Tranf. Banc�ria com Bancos de Destino
                                  '// �ndice 4: Duplicatas e Lan�amentos a receber e recebidos
                                  '// �ndice 5: Duplicatas e Lan�amentos a pagar e pagos
  mbolCancelou = False
  SetPtr vbHourglass

dtInicial = CDateDef(txtFluxoConta(2).Text)
dtFinal = CDateDef(lblFlxDesc(2).Caption)

  '
  'Verifica se a Moeda Informada � v�lida antes de executar a Convers�o
  '
  If Len(txtFluxoConta(5).Text) > 0 And lblFlxDesc(5).Caption = NUL Then
    MsgBox "Informe uma MOEDA v�lida para a Convers�o de Valores", vbOKOnly Or vbExclamation, MsgBoxCaption
    LetFocus txtFluxoConta(5).hWnd
    Selecione txtFluxoConta(5)
    mbolCancelou = True
    SetPtr vbDefault
    Exit Sub
  End If
  If TemMoeda(txtFluxoConta(5).Text, lblFlxDesc(5).Caption) Then
    If TemCotacao(txtFluxoConta(5).Text, lblFlxDesc(5).Caption, dtInicial, dtFinal) = Empty Then
      MsgBox "Informe a cota��o da Moeda " & txtFluxoConta(5).Text & " no per�odo de " & _
      dtPer(0) & " at� " & dtPer(1)
      LetFocus txtFluxoConta(5).hWnd
      Selecione txtFluxoConta(5)
      mbolCancelou = True
      SetPtr vbDefault
      Exit Sub
    End If
  End If
    
  
  '// Verifica se a data digitada est� correta
  If (Not EData(txtFluxoConta(2).Text)) Then
    MsgFunc ResolveResString(19, resUM, txtFluxoConta(2).Text)
    GoTo FiltroFluxoConta_Erro
  End If

  '// Verifica se os Bancos Inicial e Final est�o corretos
  If ((IsValid(txtFluxoConta(0).Text)) And (Len(lblFlxDesc(0).Caption) = ZERO)) Then
    MsgFunc ResolveResString(50, resUM, txtFluxoConta(0).Text, resDOIS, "Bancos")
    GoTo FiltroFluxoConta_Erro
  End If

  If ((IsValid(txtFluxoConta(1).Text)) And (Len(lblFlxDesc(1).Caption) = ZERO)) Then
    MsgFunc ResolveResString(50, resUM, txtFluxoConta(1).Text, resDOIS, "Bancos")
    GoTo FiltroFluxoConta_Erro
  End If

  '// Verifica se os Grupos Inicial e Final est�o corretos
  If ((IsValid(txtFluxoConta(3).Text)) And (Len(lblFlxDesc(3).Caption) = ZERO)) Then
    MsgFunc ResolveResString(50, resUM, txtFluxoConta(3).Text, resDOIS, "Grupos")
    GoTo FiltroFluxoConta_Erro
  End If

  If ((IsValid(txtFluxoConta(4).Text)) And (Len(lblFlxDesc(4).Caption) = ZERO)) Then
    MsgFunc ResolveResString(50, resUM, txtFluxoConta(4).Text, resDOIS, "Grupos")
    GoTo FiltroFluxoConta_Erro
  End If

  
  '// Iniciando a instru��o de sele��o de Grupos
  strContas = "SELECT * FROM Contas"

  lTmpIni = CLngDef(txtFluxoConta(3).Text)
  lTmpFim = CLngDef(txtFluxoConta(4).Text)

  If ((lTmpIni > 0) And (lTmpFim > 0)) Then
    If (lTmpIni = lTmpFim) Then
      AppendStr strContas, " WHERE Grupo = " & CStr(lTmpIni)
    Else
      Concat strContas, " WHERE (Grupo BETWEEN ", CStr(lTmpIni), " AND ", CStr(lTmpFim), ")"
    End If
  ElseIf ((lTmpIni > 0) And (lTmpFim = 0)) Then
    Concat strContas, " WHERE Grupo >= ", CStr(lTmpIni)
  ElseIf ((lTmpIni = 0) And (lTmpFim > 0)) Then
    Concat strContas, " WHERE Grupo <= ", CStr(lTmpFim)
  End If
  
  AppendStr strContas, " ORDER BY Grupo, C�digo;"
  
  '// Resolvendo as datas inicial e final
  dtPer(0) = CDateDef(txtFluxoConta(2).Text)
  dtPer(1) = DateAdd(DD_DIA, 5, dtPer(0))

  '// Resolvendo os Bancos inicial e final
  lTmpIni = CLngDef(txtFluxoConta(0).Text)
  lTmpFim = CLngDef(txtFluxoConta(1).Text)

  '// Definindo as instru��es de sele��o para cada tabela de dados
  '// nos elementos da matriz respectivos a cada uma
  '// Elemento Zero utilizado para a tabela de Aplica��es recebidas
  
  'Projeto: #4172 - Hist�ria: #4165 - Problema#4310 - Moacir Pfau(25/01/2013)
  #If FOXSQL = 1 Then
      sWhere(0) = "Convert(varchar(10),Data,120) = '|1' AND Conta = |2 AND Tipo = '" & GetResOptions(1001, 1) & _
                  "'" & ResolveBancos(lTmpIni, lTmpFim)
      
      '// Elemento Um utilizado para a tabela de Aplica��es pagas
      
      sWhere(1) = "Convert(varchar(10),Data,120) = '|1' AND Conta = |2 AND Tipo <> '" & GetResOptions(1001, 1) & _
                  "'" & ResolveBancos(lTmpIni, lTmpFim)
    
      '// Elemento Tres utilizado para a tabela de Transf. Banc�ria com
      '// Banco de Origem. Note que a fun��o ResolveBancos configura a
      '// express�o como: Banco = (?). Ent�o eu troco Banco por Origem
      '// que � o nome do campo nesta tabela.
      
      sWhere(2) = "Convert(varchar(10),Data,120) = '|1' AND Conta = |2 " & ResolveBancos(lTmpIni, lTmpFim)
      MidStr sWhere(2), "Banco", "Origem"
    
      '// Elemento Tres utilizado para a tabela de Transf. Banc�ria com
      '// Banco de Destino.
      
      sWhere(3) = "Convert(varchar(10),Data,120) = '|1' AND Conta = |2 " & ResolveBancos(lTmpIni, lTmpFim)
      MidStr sWhere(3), "Banco", "Destino"
      
      '// Elemento Quatro utilizado para tabelas de Duplicatas e Lan�amentos
      '// A Receber e Recebidos
      
      sWhere(4) = "PagRec = 'R' AND Convert(varchar(10),Libera��o,120)= '|1' AND Conta = |2" & _
                  ResolveBancos(lTmpIni, lTmpFim) & ESP & _
                  "AND Situa��o <> 'Cancelada'"  'Protocolo 73606 (Somente os titulos n�o cancelados)
    
      '// Elemento Cinco utilizado para tabelas de Duplicatas e Lan�amentos
      '// A Pagar e Pagos
      sWhere(5) = "PagRec = 'P' AND Convert(varchar(10),Libera��o,120) = '|1' AND Conta = |2" & _
                  ResolveBancos(lTmpIni, lTmpFim) & ESP & _
                  "AND Situa��o <> 'Cancelada'"  'Protocolo 73606 (Somente os titulos n�o cancelados)
  #Else
      sWhere(0) = "Data = #|1# AND Conta = |2 AND Tipo = '" & GetResOptions(1001, 1) & _
                  "'" & ResolveBancos(lTmpIni, lTmpFim)
      
      '// Elemento Um utilizado para a tabela de Aplica��es pagas
      
      sWhere(1) = "Data = #|1# AND Conta = |2 AND Tipo <> '" & GetResOptions(1001, 1) & _
                  "'" & ResolveBancos(lTmpIni, lTmpFim)
    
      '// Elemento Tres utilizado para a tabela de Transf. Banc�ria com
      '// Banco de Origem. Note que a fun��o ResolveBancos configura a
      '// express�o como: Banco = (?). Ent�o eu troco Banco por Origem
      '// que � o nome do campo nesta tabela.
      
      sWhere(2) = "Data = #|1# AND Conta = |2 " & ResolveBancos(lTmpIni, lTmpFim)
      MidStr sWhere(2), "Banco", "Origem"
    
      '// Elemento Tres utilizado para a tabela de Transf. Banc�ria com
      '// Banco de Destino.
      
      sWhere(3) = "Data = #|1# AND Conta = |2 " & ResolveBancos(lTmpIni, lTmpFim)
      MidStr sWhere(3), "Banco", "Destino"
      
      '// Elemento Quatro utilizado para tabelas de Duplicatas e Lan�amentos
      '// A Receber e Recebidos
      
      sWhere(4) = "PagRec = 'R' AND Libera��o = #|1# AND Conta = |2" & _
                  ResolveBancos(lTmpIni, lTmpFim) & ESP & _
                  "AND Situa��o <> 'Cancelada'"  'Protocolo 73606 (Somente os titulos n�o cancelados)
    
      '// Elemento Cinco utilizado para tabelas de Duplicatas e Lan�amentos
      '// A Pagar e Pagos
      sWhere(5) = "PagRec = 'P' AND Libera��o = #|1# AND Conta = |2" & _
                  ResolveBancos(lTmpIni, lTmpFim) & ESP & _
                  "AND Situa��o <> 'Cancelada'"  'Protocolo 73606 (Somente os titulos n�o cancelados)
  #End If

  '// Verifica se o usu�rio est� filtrando por tipo de Lan�amento. Na ComboBox
  '// a �nica op��o que tem um valor de ItemData diferente de zero � a
  '// op��o Todos
  
  If (GetItemData(cboFluxoConta(0)) = ZERO) Then
    AppendStr sWhere(4), " AND Tipo = '" & cboFluxoConta(0).Text & "'"
    AppendStr sWhere(5), " AND Tipo = '" & cboFluxoConta(0).Text & "'"
  End If
  
  If (cboFluxoConta(1).ListIndex = 0) Then              '// 0 == Quitados
    AppendStr sWhere(4), " AND Pagamento IS NOT NULL"
    AppendStr sWhere(5), " AND Pagamento IS NOT NULL"
  ElseIf (cboFluxoConta(1).ListIndex = 1) Then          '// 1 == Em Aberto
    AppendStr sWhere(4), " AND Pagamento IS NULL"
    AppendStr sWhere(5), " AND Pagamento IS NULL"
  End If

  If cboFluxoConta(2).Text <> "Todos" Then
    If cboFluxoConta(2).Text = "Sim" Then
      AppendStr sWhere(4), " AND Conciliado=True"
      AppendStr sWhere(5), " AND Conciliado=True"
    Else
      AppendStr sWhere(4), " AND Conciliado=False"
      AppendStr sWhere(5), " AND Conciliado=False"
    End If
  End If
  
  If (AbreRecordset(rstContas, strContas, dbOpenSnapshot) = WL_OK) Then
    If (CriaAuxFluxo(rstTemp)) Then
      If (UpdateAux(rstContas, rstTemp, dtPer, sWhere, cMov)) Then
        Call PrintFluxoConta(rstTemp, pdeDestino, dtPer, cMov)
      End If
    End If
    DeleteAux rstTemp, NUL
  ElseIf (UltimoRetorno() = WL_NORECORD) Then
    MsgFunc "N�o foi encontrado nenhum Grupo com os valores indicados"
  End If
  
FiltroFluxoConta_Erro:
  If (Err().Number) Then
    DAOErros NUL
  End If
  FechaRecordset rstContas
  MsgBar Me.Caption
  SetPtr vbDefault
End Sub

' FUNCTION..: ResolveBancos
' Objetivo..: Resolve as instru��es de sele��o das tabelas de Aplica��es
'             Transfer�ncias, Duplicatas e Lan�amentos.
' Argumentos: [lBcoIni]: C�digo do Banco Inicial.
'             [lBcoFim]: C�digo do Banco Final.
' Retorna...: A string de sele��o resolvida.
' ---------------------------------------------------------------------------
Private Function ResolveBancos(lBcoIni As Long, lBcoFim As Long) As String
Const PREV$ = " AND ((SELECT Previs�o FROM Bancos WHERE Bancos.Banco = |3) = True)"
Dim strPrev As String   '

Dim sResult As String

  '
  ' Prevendo Banco = Zero
  '
  If Not IsValid(txtFluxoConta(0).Text) Then
    strPrev = " AND (((SELECT Previs�o FROM Bancos WHERE Bancos.Banco = |3) = True) OR (|3 = 0))"
  Else
    strPrev = PREV
  End If
 
  If ((lBcoIni > 0) And (lBcoFim > 0)) Then
    If (lBcoIni = lBcoFim) Then
      sResult = " AND Banco = " & CStr(lBcoIni)
    Else
      sResult = " AND (Banco BETWEEN " & CStr(lBcoIni) & " AND " & CStr(lBcoFim) & ")" & strPrev
    End If
  ElseIf ((lBcoIni > 0) And (lBcoFim = 0)) Then
    sResult = " AND Banco >= " & CStr(lBcoIni) & strPrev
  ElseIf ((lBcoIni = 0) And (lBcoFim > 0)) Then
    sResult = " AND Banco <= " & CStr(lBcoFim) & strPrev
  End If
  ResolveBancos = sResult
  
End Function

' FUNCTION..: CriaAuxFluxo
' Objetivo..: Cria a tabela auxiliar que ser� utilizada para impress�o dos
'             dos dados do relat�rio.
' Argumento.: [rstAux]: Recordset que receber� uma refer�ncia a tabela criada
' Retorna...: True se obtiver sucesso, False se n�o.
' ----------------------------------------------------------------------------
Private Function CriaAuxFluxo(rstAux As Object) As Boolean
Dim fsFluxo(8) As FieldStruct

  AppendVar fsFluxo(0), "Grupo", dbLong
  AppendVar fsFluxo(1), "Conta", dbLong
  AppendVar fsFluxo(2), "Desc", dbText, 40
  AppendVar fsFluxo(3), "Dia1", dbCurrency
  AppendVar fsFluxo(4), "Dia2", dbCurrency
  AppendVar fsFluxo(5), "Dia3", dbCurrency
  AppendVar fsFluxo(6), "Dia4", dbCurrency
  AppendVar fsFluxo(7), "Dia5", dbCurrency
  AppendVar fsFluxo(8), "Dia6", dbCurrency

  If (CrieAux(rstAux, fsFluxo)) Then
    CriaAuxFluxo = True
  Else
    MsgFunc LoadResString(174)
  End If
  
End Function

' FUNCTION..: UpdateAux
' Objetivo..: Grava os dados no arquivo auxiliar. A fun��o totaliza
'             os valores movimentados para as Contas nos dias correspondentes
'             ordenando por data e conta.
' Argumentos: [rstContas]: Recordset que cont�m os grupos e as contas selecionadas.
'             [rstAux   ]: Recordset da tabela auxiliar.
'             [dtDatas  ]: Matriz com as datas Inicial e final.
'             [strWhere ]: Matriz com as instru��es de compara��o para as
'                          tabelas de Lan�amentos, Duplicatas, Aplica��es e
'                          Tranfer�ncias Banc�rias.
'             [cMov     ]: Matriz para c�lculo do movimento di�rio.
' Retorna...: True se obtiver sucesso, False se n�o.
' -----------------------------------------------------------------------------
Private Function UpdateAux(rstContas As Object, rstAux As Object, dtDatas() As Date, strWhere() As String, cMov() As Currency) As Boolean
Dim DtDia       As Date            '// Dia sendo calculado
Dim sTmp        As String          '// Guarda, temporariamente, as clausulas de compara��o das tabelas
Dim lConta      As Long            '// Conta atual
Dim cTotal      As Currency        '// Total de cada conta a cada dia
Dim bDia        As Byte            '// N�mero dos campos (Dia1, Dia2, Dia3, ...)
Dim sConta      As String          '// Descri��o da Conta
Dim bEmAberto   As Boolean         '// Situa��o dos pagamentos

  On Error GoTo UpdateAux_Erro

  bEmAberto = (cboFluxoConta(1).ListIndex = 1)
  Do
    lConta = GetValue(rstContas, "C�digo", ZERO)
    sConta = GetValue(rstContas, "Descri��o", ZERO)
    DtDia = dtDatas(0)             '// Inicia na data Inicial ( l�gico ).
    cTotal = ZERO
    bDia = ZERO
    
    SimpleMsgBar "Calculando Conta: " & sConta
    
    If (mbolCancelou) Then GoTo UpdateAux_Erro
    DoEvents
    
    rstAux.AddNew
    rstAux("Grupo").Value = GetValue(rstContas, "Grupo", ZERO)
    rstAux("Conta").Value = lConta
    rstAux("Desc").Value = sConta
    
    While (DateDiff(DD_DIA, DtDia, dtDatas(1)) >= ZERO)
      If (mbolCancelou) Then GoTo UpdateAux_Erro
      DoEvents
      
      Inc bDia
      
      '// S� posso acrescentar os valores de Aplica��es Financeiras e Transf.
      '// Banc�rias se o usu�rio N�O escolheu Pagamento Em Aberto
      If (Not bEmAberto) Then
        
        If txtFluxoConta(2).Text <> "N�o" Then
            '// Atualizando a instru��o Select para constar a conta atual na data atual
            '// na tabela de Aplica��es com tipo = 'Juros/Corre��o' (entradas)
            
            sTmp = strWhere(0)
            MidStr sTmp, "|1", InverteData(DtDia)
            MidStr sTmp, "|2", CStr(lConta)
            MidAll sTmp, "|3", "Aplica��es.Banco"
            cTotal = Soma("Valor", "Aplica��es", sTmp, ZERO) / Cotacao(txtFluxoConta(5).Text, DtDia)
            
    
            '// Atualiza a instru��o para constar a conta atual na data atual
            '// em Aplica��es com tipo <> 'Juros/Corre��o' (sa�das)
    
            sTmp = strWhere(1)
            MidStr sTmp, "|1", InverteData(DtDia)
            MidStr sTmp, "|2", CStr(lConta)
            MidAll sTmp, "|3", "Aplica��es.Banco"
            cTotal = cTotal - Soma("Valor", "Aplica��es", sTmp, ZERO) / Cotacao(txtFluxoConta(5).Text, DtDia)
    
            '// Atualiza a instru��o para a tabela de Transf Banc�ria
            '// com o Banco de Destino (entrada)
    
            sTmp = strWhere(2)
            MidStr sTmp, "|1", InverteData(DtDia)
            MidStr sTmp, "|2", CStr(lConta)
            MidAll sTmp, "|3", "[Transf Banc�ria].Destino"
            cTotal = cTotal + Soma("Valor", "[Transf Banc�ria]", sTmp, ZERO) / Cotacao(txtFluxoConta(5).Text, DtDia)
            
            '// Atualiza a instru��o para constar a conta atual na data atual
            '// em Transf Banc�ria com Banco de Origem (sa�das)
            
            sTmp = strWhere(3)
            MidStr sTmp, "|1", InverteData(DtDia)
            MidStr sTmp, "|2", CStr(lConta)
            MidAll sTmp, "|3", "[Transf Banc�ria].Origem"
            cTotal = cTotal - Soma("Valor", "[Transf Banc�ria]", sTmp, ZERO) / Cotacao(txtFluxoConta(5).Text, DtDia)
          End If
      End If
      
      '// Altera a instru��o para a tabela de Duplicatas a Receber
      sTmp = strWhere(4)
      MidStr sTmp, "|1", InverteData(DtDia)
      MidStr sTmp, "|2", CStr(lConta)
      
      MidAll sTmp, "|3", "Duplicatas.Banco"                 ' Previs�o
      'cTotal = cTotal + Soma("([Valor Original] + Acr�scimo - Abatimento)", _
                             "Duplicatas", sTmp, ZERO)
      cTotal = cTotal + SomarMoedas("Duplicatas", sTmp, txtFluxoConta(5).Text)
      
      MidAll sTmp, "Duplicatas.Banco", "Lan�amentos.Banco"  ' Previs�o
      'cTotal = cTotal + Soma("([Valor Original] + Acr�scimo - Abatimento)", _
                             "Lan�amentos", sTmp, ZERO)
      cTotal = cTotal + SomarMoedas("Lan�amentos", sTmp, txtFluxoConta(5).Text)
                             

      '// Altera a instru��o para a tabela de Duplicatas a pagar
      sTmp = strWhere(5)
      MidStr sTmp, "|1", InverteData(DtDia)
      MidStr sTmp, "|2", CStr(lConta)

      MidAll sTmp, "|3", "Duplicatas.Banco"                 ' Previs�o
      'cTotal = cTotal - Soma("([Valor Original] + Acr�scimo - Abatimento)", _
                             "Duplicatas", sTmp, ZERO)
      cTotal = cTotal - SomarMoedas("Duplicatas", sTmp, txtFluxoConta(5).Text)
      
      MidAll sTmp, "Duplicatas.Banco", "Lan�amentos.Banco"  ' Previs�o
      'cTotal = cTotal - Soma("([Valor Original] + Acr�scimo - Abatimento)", _
                             "Lan�amentos", sTmp, ZERO)
      cTotal = cTotal - SomarMoedas("Lan�amentos", sTmp, txtFluxoConta(5).Text)

      
      rstAux("Dia" & CStr(bDia)).Value = cTotal
      
      cMov(bDia - 1) = cMov(bDia - 1) + cTotal    '// Guarda a movimenta��o di�ria
      
      DtDia = DateAdd(DD_DIA, 1, DtDia)           '// Avan�a para o pr�ximo dia
    Wend
    rstAux.update                                 '// Grava o registro atual
    rstContas.MoveNext                            '// Avan�a para a pr�xima conta
  Loop Until (rstContas.EOF)

  If (Not EstaVazio(rstAux)) Then
    UpdateAux = True                              '// Retorna SUCESSO!
  Else
    MsgFunc LoadResString(257)
  End If
  
UpdateAux_Erro:
  If (Err().Number) Then
    DAOErros NUL
    If (rstAux.EditMode <> dbEditNone) Then rstAux.CancelUpdate
    UpdateAux = False
  End If
End Function

' SUB.......: PrintFluxoConta
' Objetivo..: Configura o Gerador de Relat�rio para a impress�o.
' Argumentos: [rstDados]: Recordset que cont�m os dados que devem ser impressos.
'             [pdDest  ]: Destino da impress�o.
'             [dPeriodo]: Matriz com as datas Inicial e Final.
'             [curMov  ]: Matriz com o ac�mulo da movimenta��o de cada dia.
' -------------------------------------------------------------------------------
Private Sub PrintFluxoConta(rstDados As Object, pdDest As Long, dPeriodo() As Date, curMov() As Currency)
Dim wrFluxoContas As KeybReport
Dim curSaldos     As Currency         '// Calcula os saldos inicia e final de cada dia

  If (CreateReport(wrFluxoContas, pdDest, "Fluxo por Conta e Grupo")) Then
    With wrFluxoContas
      Set .Recordset = rstDados
      
      PageHeader wrFluxoContas, "Fluxo por Conta e Grupo"
      .UltimaSecao.AddLinha
      .UltimaLinha.AddCampo , wrCSFixedText, "Per�odo de " & DataToStr(dPeriodo(0)) & _
                              " a " & DataToStr(dPeriodo(1)), wrTACentro
      'Insere linha no Cabe�alho para Informar a Moeda
      If Len(txtFluxoConta(5).Text) > 0 Then
        .UltimaSecao.AddLinha "Moeda"
        .UltimaSecao(.UltimaSecao.LinhasCount).AddCampo , wrCSFixedText, "Valores Demonstrados em: " & txtFluxoConta(5).Text, wrTACentro
      End If
    
      
      .UltimoGrupo.AddSecao scFooter, 1       '// Adiciona uma linha em branco
      
      .FontSize = 9
      .FontStyle = wrFSBold
      
      '// Montando o grupo principal. Este grupo quebra por c�digo de Grupo
      
      .AddGrupo "1"
      .Grupo(1).Quebra = "Grupo"
      .Grupo(1).AddSecao scHeader, 2
      With .Grupo(1).Header.Linha(1)
        .Height = .Height * 2                   '// Altura de duas linhas
        .DrawBorder = wrDBAllBorders
        .AddCampo , wrCSDataLink, "Descri��o"
        .Campo(1).Top = ((.Height \ 2) - (.Campo(1).Height \ 2))
        .Campo(1).TableLink = "Grupos"
        .Campo(1).DataLink = "C�digo = {Grupo}"
      End With
      .FontSize = 8
      
      With .Grupo(1).Header.Linha(2)
        .AddCampo , wrCSFixedText, "Conta", , 50
        .AddCampo , wrCSFixedText, DataToStr(dPeriodo(0)), wrTADireito, 20
        .AddCampo , wrCSFixedText, DataToStr(DateAdd(DD_DIA, 1, dPeriodo(0))), wrTADireito, 20
        .AddCampo , wrCSFixedText, DataToStr(DateAdd(DD_DIA, 2, dPeriodo(0))), wrTADireito, 20
        .AddCampo , wrCSFixedText, DataToStr(DateAdd(DD_DIA, 3, dPeriodo(0))), wrTADireito, 20
        .AddCampo , wrCSFixedText, DataToStr(DateAdd(DD_DIA, 4, dPeriodo(0))), wrTADireito, 20
        .AddCampo , wrCSFixedText, DataToStr(DateAdd(DD_DIA, 5, dPeriodo(0))), wrTADireito, 20
      End With
      .FontStyle = wrFSNormal
      
      .Grupo(1).AddSecao scDetalhe, 1           '// Se��o de detalhe
      With .Grupo(1).Detalhe.Linha(1)
        .AddCampo , , "Desc", , 50                    '// Descri��o da Conta
        .AddCampo , , "Dia1", wrTADireito, 20         '// Valor da primeira data
        .AddCampo , , "Dia2", wrTADireito, 20         '// Valor da segunda data
        .AddCampo , , "Dia3", wrTADireito, 20         '// Valor da terceira data
        .AddCampo , , "Dia4", wrTADireito, 20
        .AddCampo , , "Dia5", wrTADireito, 20
        .AddCampo , , "Dia6", wrTADireito, 20
        .Campo(2).Formato = FMOEDA
        .Campo(2).SuprimirZeros = True
        .Campo(3).Formato = FMOEDA
        .Campo(3).SuprimirZeros = True
        .Campo(4).Formato = FMOEDA
        .Campo(4).SuprimirZeros = True
        .Campo(5).Formato = FMOEDA
        .Campo(5).SuprimirZeros = True
        .Campo(6).Formato = FMOEDA
        .Campo(6).SuprimirZeros = True
        .Campo(7).Formato = FMOEDA
        .Campo(7).SuprimirZeros = True
      End With

      '// Rodap� do Grupo: Calcula o total do Grupo por dia

      .Grupo(1).AddSecao scFooter, 2
      With .Grupo(1).Footer.Linha(1)
        .DrawBorder = wrDBBottomBorder
        .BorderStyle = wrDot
        .AddCampo , wrCSFixedText, "Total do Grupo:", , 50
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSSubTotal, "Dia1", wrTADireito, 20
        .Campo(2).SuprimirZeros = True
        .Campo(2).Formato = FMOEDA
        .AddCampo , wrCSSubTotal, "Dia2", wrTADireito, 20
        .Campo(3).SuprimirZeros = True
        .Campo(3).Formato = FMOEDA
        .AddCampo , wrCSSubTotal, "Dia3", wrTADireito, 20
        .Campo(4).SuprimirZeros = True
        .Campo(4).Formato = FMOEDA
        .AddCampo , wrCSSubTotal, "Dia4", wrTADireito, 20
        .Campo(5).SuprimirZeros = True
        .Campo(5).Formato = FMOEDA
        .AddCampo , wrCSSubTotal, "Dia5", wrTADireito, 20
        .Campo(6).SuprimirZeros = True
        .Campo(6).Formato = FMOEDA
        .AddCampo , wrCSSubTotal, "Dia6", wrTADireito, 20
        .Campo(7).SuprimirZeros = True
        .Campo(7).Formato = FMOEDA
      End With

      '// Cria o Grupo de resumo
      .AddGrupo "2"

      If (Not GrupoResumo(.Grupo(2), dPeriodo(0), curMov)) Then
        '// Se a fun��o retorna False � porque o usu�rio cancelou
        GoTo PrintFluxoConta_Erro
      End If
      
    End With
    wrFluxoContas.BeginPrint gTipoDB
    wrFluxoContas.EndPrint
  End If

PrintFluxoConta_Erro:
  Set wrFluxoContas = Nothing
End Sub

' FUNCTION..: GrupoResumo
' Objetivo..: Configura o Grupo de resumo do relat�rio.
' Argumentos: [grpResumo   ]: Objeto Grupo do Gerador de relat�rios.
'             [dtInicial   ]: Data Inicial.
'             [curMovimento]: Matriz com o valor da movimenta��o di�ria.
' Retorna...: True se obtiver sucesso, False se o usu�rio cancelar.
' -----------------------------------------------------------------------
Private Function GrupoResumo(grpResumo As Grupo, dtInicial As Date, curMovimento() As Currency) As Boolean
Dim cSaldoInicial As Currency         '// Saldo inicial do c�lculo
Dim cSaldoDia(5)  As Currency         '// Saldos di�rios
Dim lSaldo        As Long             '// Utilizado no Loop
Dim DtDia(5)      As Date             '// Data Base para a Cota��o da moeda



  If (SaldoInicialGeral(dtInicial, cSaldoInicial, False, strMoeda:=txtFluxoConta(5).Text, StrDescMoeda:=lblFlxDesc(5).Caption, sConciliado:=cboFluxoConta(2).Text) = WL_OK) Then

    cSaldoDia(0) = cSaldoInicial + curMovimento(0)
    For lSaldo = 1 To 5
      cSaldoDia(lSaldo) = (cSaldoDia(lSaldo - 1) + curMovimento(lSaldo))
      
    Next
    
    grpResumo.AddSecao scHeader, 5
    grpResumo.Header.Linha(1).AddCampo , wrCSSimpleLine
    grpResumo.Header.Linha(5).AddCampo , wrCSSimpleLine
  
    With grpResumo.Header.Linha(2)
      .AddCampo , wrCSFixedText, "Saldo Anterior:", , 50
      .Campo(1).FontStyle = wrFSBold
      
      .AddCampo , wrCSFixedText, Format$(cSaldoInicial / Cotacao(txtFluxoConta(5), dtInicial), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(0) / Cotacao(txtFluxoConta(5), DateAdd("D", 1, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(1) / Cotacao(txtFluxoConta(5), DateAdd("D", 2, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(2) / Cotacao(txtFluxoConta(5), DateAdd("D", 3, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(3) / Cotacao(txtFluxoConta(5), DateAdd("D", 4, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(4) / Cotacao(txtFluxoConta(5), DateAdd("D", 5, dtInicial)), FMOEDA), wrTADireito, 20
      .Campo(2).SuprimirZeros = True
      .Campo(3).SuprimirZeros = True
      .Campo(4).SuprimirZeros = True
      .Campo(5).SuprimirZeros = True
      .Campo(6).SuprimirZeros = True
      .Campo(7).SuprimirZeros = True
    End With

    With grpResumo.Header.Linha(3)
      .AddCampo , wrCSFixedText, "Movimenta��o do Dia:", , 50
      .Campo(1).FontStyle = wrFSBold

      .AddCampo , wrCSFixedText, Format$(curMovimento(0) / Cotacao(txtFluxoConta(5), dtInicial), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(curMovimento(1) / Cotacao(txtFluxoConta(5), DateAdd("D", 1, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(curMovimento(2) / Cotacao(txtFluxoConta(5), DateAdd("D", 2, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(curMovimento(3) / Cotacao(txtFluxoConta(5), DateAdd("D", 3, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(curMovimento(4) / Cotacao(txtFluxoConta(5), DateAdd("D", 4, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(curMovimento(5) / Cotacao(txtFluxoConta(5), DateAdd("D", 5, dtInicial)), FMOEDA), wrTADireito, 20
      .Campo(2).SuprimirZeros = True
      .Campo(3).SuprimirZeros = True
      .Campo(4).SuprimirZeros = True
      .Campo(5).SuprimirZeros = True
      .Campo(6).SuprimirZeros = True
      .Campo(7).SuprimirZeros = True
    End With

    With grpResumo.Header.Linha(4)
      .AddCampo , wrCSFixedText, "Saldo Final do Dia:", , 50
      .Campo(1).FontStyle = wrFSBold

      .AddCampo , wrCSFixedText, Format$(cSaldoDia(0) / Cotacao(txtFluxoConta(5), dtInicial), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(1) / Cotacao(txtFluxoConta(5), DateAdd("D", 1, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(2) / Cotacao(txtFluxoConta(5), DateAdd("D", 2, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(3) / Cotacao(txtFluxoConta(5), DateAdd("D", 3, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(4) / Cotacao(txtFluxoConta(5), DateAdd("D", 4, dtInicial)), FMOEDA), wrTADireito, 20
      .AddCampo , wrCSFixedText, Format$(cSaldoDia(5) / Cotacao(txtFluxoConta(5), DateAdd("D", 5, dtInicial)), FMOEDA), wrTADireito, 20
      .Campo(2).SuprimirZeros = True
      .Campo(3).SuprimirZeros = True
      .Campo(4).SuprimirZeros = True
      .Campo(5).SuprimirZeros = True
      .Campo(6).SuprimirZeros = True
      .Campo(7).SuprimirZeros = True
    End With
    GrupoResumo = True
  End If
  
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
