VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCalculos 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expurgo de Dados"
   ClientHeight    =   5010
   ClientLeft      =   345
   ClientTop       =   1125
   ClientWidth     =   7815
   Icon            =   "Calculos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCalculos 
      Caption         =   "Exclus�o de Duplicatas e Lan�amentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2175
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox cboCalculos 
         Height          =   315
         Index           =   2
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cboCalculos 
         Height          =   315
         Index           =   1
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCalculos 
         DataField       =   "Apel"
         Height          =   315
         Index           =   4
         Left            =   1080
         TabIndex        =   18
         Tag             =   "Empresas"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtCalculos 
         DataField       =   "C�digo"
         Height          =   315
         Index           =   3
         Left            =   1080
         TabIndex        =   15
         Tag             =   "Contas"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cboCalculos 
         Height          =   315
         Index           =   0
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtCalculos 
         DataField       =   "C�digo"
         Height          =   315
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Tag             =   "Lan�amentos"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCalculos 
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCalculos 
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblCalculos 
         AutoSize        =   -1  'True
         Caption         =   "Sit&ua��o:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDesc 
         Caption         =   "lblDesc(2)"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   19
         Top             =   1680
         Width           =   3240
      End
      Begin VB.Label lblCalculos 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblCalculos 
         AutoSize        =   -1  'True
         Caption         =   "&Empresa:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label lblDesc 
         Caption         =   "lblDesc(1)"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   16
         Top             =   1320
         Width           =   3240
      End
      Begin VB.Label lblDesc 
         Caption         =   "lblDesc(0)"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   13
         Top             =   960
         Width           =   3600
      End
      Begin VB.Label lblCalculos 
         AutoSize        =   -1  'True
         Caption         =   "&Conta:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label lblCalculos 
         AutoSize        =   -1  'True
         Caption         =   "C�&digo:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   540
      End
      Begin VB.Label lblCalculos 
         AutoSize        =   -1  'True
         Caption         =   "Data &Final:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblCalculos 
         AutoSize        =   -1  'True
         Caption         =   "Data &Inicial:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lblCalculos 
         AutoSize        =   -1  'True
         Caption         =   "&Sele��o Por:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Frame fraCalculos 
      Caption         =   "Re&gistros Selecionados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   7575
      Begin ComctlLib.ListView lvwCalculos 
         Height          =   1935
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdCalculos 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   23
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculos 
      Caption         =   "E&xibir"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   22
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "lblDesc(3)"
      Height          =   195
      Index           =   3
      Left            =   4320
      TabIndex        =   24
      Top             =   4560
      Width           =   705
   End
   Begin ComctlLib.ImageList imgCalculos 
      Left            =   120
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Calculos.frx":08CA
            Key             =   "check"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Calculos.frx":0BE4
            Key             =   "uncheck"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCalculos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ACT_EXIBIR = 0      '// A a��o do bot�o padr�o � exibir os Lan�amentos selecionados
Private Const ACT_EXCLUIR = 1     '// A a��o do bot�o padr�o � excluir os Lan�amentos selelcionados
Private Const IDI_OK = 1          '// �cone do registro quando deve ser exclu�do
Private Const IDI_NO = 2          '// �cone do registro quando n�o deve ser exclu�do
Private Const CMD_EXIBIR$ = "E&xibir"   '// Caption do bot�o
Private Const STR_TOTAL$ = "Total de Registros: "
Private Const SOURCE_LANC = UM    '// Origem da sele��o � a tabela de Lan�amentos
Private Const SOURCE_DUPL = 2     '// Origem da sele��o � a tabela de Duplicatas

Private mlSource As Long          '// Origem da sele��o atual
Private mlAction As Long          '// A��o atual da janela
Private mrstLanc As Object     '// Recordset com os dados selecionados pelo usu�rio

' EVENT.....: cboCalculos_GotFocus
' Objetivo..: Exibe mensagens de ajuda na barra de status do programa
' ---------------------------------------------------------------------
Private Sub cboCalculos_GotFocus(Index As Integer)
  CalcMsg cboCalculos(Index).TabIndex
End Sub

' EVENT.....: cmdCalculos_Click
' Objetivo..: Executa as fun��o dos bot�es.
' ---------------------------------------------------------------------
Private Sub cmdCalculos_Click(Index As Integer)
  Select Case Index
    Case 0                '// Bot�o Exibir/Excluir
      If (mlAction = ACT_EXIBIR) Then
        If (ShowInListView(True)) Then
          cmdCalculos(0).Caption = LoadResString(IDS_EXCLUIR)
          cmdCalculos(1).Caption = LoadResString(IDS_CANCELAR)
          lblDesc(3).Caption = STR_TOTAL & CStr(lvwCalculos.ListItems.Count)
          mlAction = ACT_EXCLUIR
        End If
      Else
        If (ExportarDados()) Then
          FechaRecordset mrstLanc
          cmdCalculos(0).Caption = CMD_EXIBIR
          cmdCalculos(1).Caption = LoadResString(IDS_FECHAR)
          lblDesc(3).Caption = STR_TOTAL & CStr(lvwCalculos.ListItems.Count)
          mlAction = ACT_EXIBIR
        End If
      End If

    Case 1                '// Bot�o Fechar
      If (mlAction = ACT_EXIBIR) Then
        Unload Me
      Else
        FechaRecordset mrstLanc
        lvwCalculos.ListItems.Clear
        cmdCalculos(0).Caption = CMD_EXIBIR
        cmdCalculos(1).Caption = LoadResString(IDS_FECHAR)
        lblDesc(3).Caption = NUL
        mlAction = ACT_EXIBIR
      End If
  End Select
End Sub

' EVENT.....: Form_Load
' Objetivo..: Configura a posi��o do formul�rio na tela e outras
'             configura��es.
' ----------------------------------------------------------------
Private Sub Form_Load()

  CenterForm Me
  
  txtCalculos(0).Text = Format$(FirstDay(DateAdd(DD_MES, -1, Date)), FDATA)
  txtCalculos(1).Text = Format$(LastDay(DateAdd(DD_MES, -1, Date)), FDATA)

  '// Carregando as strings da caixa de combina��o. A quarta op��o (de �ndice 3)
  '// � trazida como padr�o (3 - Mensal). A segunda op��o (de �ndice 1) �
  '// trazida como padr�o (1 - Vencimento).

  LoadResOptions 1026, cboCalculos(0), True, 3    '// 3 == Pagamento
  LoadResOptions 1030, cboCalculos(1), True, 1    '// 1 == Duplicatas
  LoadResOptions 1031, cboCalculos(2), True, 2    '// 2 == Ambos

  '// Configura o MaxLength do campo Empresa

  txtCalculos(4).MaxLength = GetPropValueEx("Empresas", "Apel", "Size", 15)

  '// Configurando o controle ListView.

  lvwCalculos.ColumnHeaders.add 1, , "C�digo", 795
  lvwCalculos.ColumnHeaders.add 2, , "Empresa", 1440
  lvwCalculos.ColumnHeaders.add 3, , "Tipo", 795
  lvwCalculos.ColumnHeaders.add 4, , "Descri��o", 1440
  lvwCalculos.ColumnHeaders.add 5, , "Emiss�o", 960, lvwColumnCenter
  lvwCalculos.ColumnHeaders.add 6, , "Vencimento", 960, lvwColumnCenter
  lvwCalculos.ColumnHeaders.add 7, , "Pagamento", 960, lvwColumnCenter
  lvwCalculos.ColumnHeaders.add 8, , "Libera��o", 960, lvwColumnCenter
  lvwCalculos.ColumnHeaders.add 9, , "Valor Original", 960, lvwColumnRight
  lvwCalculos.ColumnHeaders.add 10, , "Acr�scimo", 960, lvwColumnRight
  lvwCalculos.ColumnHeaders.add 11, , "Abatimento", 960, lvwColumnRight
  lvwCalculos.ColumnHeaders.add 12, , "Banco", 795, lvwColumnRight
  lvwCalculos.ColumnHeaders.add 13, , "Conta", 795, lvwColumnRight
  lvwCalculos.ColumnHeaders.add 14, , "Controle", 960
  lvwCalculos.ColumnHeaders.add 15, , "Situa��o", 960

  '// Conectando o controle ListView com o controle ImageList do Formul�rio

  lvwCalculos.SmallIcons = imgCalculos
  lblDesc(0).Caption = NUL
  lblDesc(1).Caption = NUL
  lblDesc(2).Caption = NUL
  lblDesc(3).Caption = NUL
  
  mlAction = ACT_EXIBIR
  
End Sub

' EVENT.....: Form_Unload
' Objetivo..: Descarrega a vari�vel impl�cita do VB da mem�ria.
' -------------------------------------------------------------------------
Private Sub Form_Unload(cancel As Integer)
  FechaRecordset mrstLanc
  Set frmCalculos = Nothing
End Sub

' EVENT.....: lvwCalculos_ColumnClick
' Objetivo..: Classifica o controle ListView conforme a coluna escolhida
'             pelo usu�rio.
' -------------------------------------------------------------------------
Private Sub lvwCalculos_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
Dim strField As String              '// Nome do campo que deve ser classificado

  If (Not EstaVazio(mrstLanc)) Then
  
    '// Ao contr�rio da maioria das vezes em que fiz esta rotina de classifica��o
    '// do controle ListView, nesta n�o posso, realmente, classificar o controle.
    '// Como as linhas do Recordset devem corresponder �s linhas do controle eu
    '// classifico o Recordset e ent�o preencho novamente o controle.

    If (ColumnHeader.Index = 1) Then

      '// A primeiro coluna sempre � C�digo, por�m, se a tabela exibida no
      '// momento for Duplicatas o nome do campo que deve ser classificado �
      '// Nota.

      If (mlSource = SOURCE_LANC) Then
        strField = "C�digo"
      Else
        strField = "Nota"
      End If
    Else
      strField = "[" & ColumnHeader.Text & "]"
    End If
  
    If (SortRecordset(mrstLanc, strField) = WL_OK) Then
      lvwCalculos.ListItems.Clear
      CompleteListView mrstLanc
    End If
  End If
  
End Sub

' EVENT.....: lvwCalculos_DblClick
' Objetivo..: Alterna a propriedade SmallIcon do �ten que identifica se ele deve
'             ou n�o ser exclu�do.
' -------------------------------------------------------------------------
Private Sub lvwCalculos_DblClick()
  If (Not IsNothing(lvwCalculos.SelectedItem)) Then
    If (lvwCalculos.SelectedItem.SmallIcon = IDI_OK) Then
      lvwCalculos.SelectedItem.SmallIcon = IDI_NO
    Else
      lvwCalculos.SelectedItem.SmallIcon = IDI_OK
    End If
  End If
End Sub

' EVENT.....: lvwCalculos_KeyDown
' Objetivo..: Alterna a propriedade SmallIcon do �ten que estiver atualmente
'             selecionado.
' -------------------------------------------------------------------------
Private Sub lvwCalculos_KeyDown(KeyCode As Integer, Shift As Integer)
  If ((Shift = ZERO) And (KeyCode = vbKeySpace)) Then
    If (Not IsNothing(lvwCalculos.SelectedItem)) Then
      lvwCalculos_DblClick
    End If
  End If
End Sub

' EVENT.....: txtCalculos_Change
' Objetivo..: Exibe a descri��o do Lan�amento ou Nome da empresa quando
'             o usu�rio altera o conte�dos dos campos 2 e/ou 3.
' -------------------------------------------------------------------------
Private Sub txtCalculos_Change(Index As Integer)
  Select Case Index
  '
  ' C�digo do Lan�amento
  Case 2
    GetAssocValue "SELECT Descri��o FROM Lan�amentos WHERE C�digo = " & _
                  txtCalculos(2).Text & " AND PagRec = 'P';", lblDesc(0)
  '
  ' C�digo da Conta
  Case 3
    GetAssocValue "SELECT Descri��o FROM Contas WHERE C�digo = " & _
                  txtCalculos(3).Text & ";", lblDesc(1)
  '
  ' Fantasia da Empresa
  Case 4
    GetAssocValue "SELECT Raz�o, Apel FROM Empresas WHERE Apel = '" & _
                  txtCalculos(4).Text & "';", lblDesc(2), txtCalculos(4)
  '
  End Select
End Sub

' EVENT.....: txtCalculos_GotFocus
' Objetivo..: Seleciona todo o texto do controle e exibe mensagens
'             explicativas na barra de status do Sistema.
' -------------------------------------------------------------------
Private Sub txtCalculos_GotFocus(Index As Integer)
  Selecione txtCalculos(Index)
  CalcMsg txtCalculos(Index).TabIndex
End Sub

' EVENT.....: txtCalculos_KeyDown
' Objetivo..: Abre a janela de pesquisa para que o usu�rio possa selecionar
'             um Lan�amentos existente.
' -----------------------------------------------------------------------
Private Sub txtCalculos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If ((Shift = ZERO) And (KeyCode = vbKeyPageDown)) Then
    Select Case Index
    '
    ' C�digo do Lan�amento
    Case 2
      Dim strInst As String       '// Instru��o Select
      
      strInst = "SELECT C�digo, Tipo, Empresa, Descri��o, Banco, Conta, " & _
                "[Valor Original], Acr�scimo, Abatimento, Emiss�o, Vencimento, " & _
                "Pagamento, Libera��o FROM Lan�amentos WHERE PagRec = 'P'"

      If (EData(txtCalculos(0).Text)) Then    '// Se o campo de data inicial contiver uma data v�lida
        AppendStr strInst, " AND Vencimento >= #" & InverteData(txtCalculos(0).Text) & "#"
      End If
      If (EData(txtCalculos(1).Text)) Then    '// Se o campo de data final contiver uma data v�lida
        AppendStr strInst, " AND Vencimento <= #" & InverteData(txtCalculos(1).Text) & "#"
      End If
      strInst = strInst & ";"

      PCampo "Lan�amentos", strInst, PB_CAMPO, txtCalculos(2), "C�digo"
    '
    ' C�digo da Conta
    Case 3
      PCampo "Contas", "Contas", PB_CAMPO, txtCalculos(3), "C�digo"
    '
    ' Fantasia da Empresa
    Case 4
      PCampo "Empresas", "Empresas", PB_CAMPO, txtCalculos(4), "Apel"
    '
    End Select
  End If
End Sub

' SUB.......: txtCalculos_KeyPress
' Objetivo..: Mapear as teclas que o usu�rio utiliza sobre um determinado
'             campo da janela.
' ---------------------------------------------------------------------------
Private Sub txtCalculos_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
  '
  ' Data Inicial e Final
  Case 0, 1
    SetMascara KeyAscii, txtCalculos(Index).SelStart, MASK_DATE
  '
  ' C�digo do Lan�amento
  Case 2
    SetMascara KeyAscii, txtCalculos(2).SelStart, fMask("Lan�amentos", "C�digo")
  '
  ' C�digo da Conta
  Case 3
    SetMascara KeyAscii, txtCalculos(3).SelStart, fMask("Contas", "C�digo")
  '
  End Select
End Sub

' SUB.......: CalcMsg
' Objetivo..: Exibe as mensagens na barra de status conforme o TabIndex
'             do controle.
' Argumento.: [iTabIndex]: TabIndex do controle que recebeu o foco.
' -----------------------------------------------------------------------
Private Sub CalcMsg(iTabIndex As Integer)
  Select Case iTabIndex
    Case 2              '// Sele��o Por
      MsgBar "Tipos da data filtrada"
    '
    Case 4              '// Tipo
      MsgBar "Tipos de Lan�amentos"
    '
    Case 6              '// Situa��o
      MsgBar "Seleciona por contas Pagas e/ou Recebidas"
    '
    Case 8              '// Data Inicial
      MsgBar ResolveResString(161, resUM, cboCalculos(0).Text)
    '
    Case 10             '// Data Final
      MsgBar ResolveResString(162, resUM, cboCalculos(0).Text)
    '
    Case 12             '// Lan�amento
      MsgBar "C�digo do Lan�amento e/ou Duplicata"
    '
    Case 15             '// Conta
      MsgBar "C�digo da Conta"
    '
    Case 18             '// Empresa
      MsgBar "Nome Fantasia da Empresa"
  End Select
End Sub

' FUNCTION..: VerFiltro
' Objetivo..: Verifica o filtro utilizado pelo usu�rio, cria a instru��o
'             de sele��o de dados se estiver tudo correto ou exibe mensagens
'             de alerta quando o usu�rio perfaz um filtro que n�o � v�lido.
' Argumento.: [strString]: Vari�vel string que receber� a instru��o select
'                          montada.
' Retorna...: True se obtiver sucesso, caso contr�rio False.
' -----------------------------------------------------------------------
Private Function VerFiltro(strString As String) As Boolean
Dim dtDatas(1) As Date        '// Cont�m as datas inicial e final
Dim dtMesAno   As Date        '// Vari�vel utilizada para testar o movimento conferido
Dim bReturn    As Boolean     '// Retorno da Fun��o

  If (EData(txtCalculos(0).Text)) Then  '// Se a data inicial for uma data v�lida
    dtDatas(0) = CDate(txtCalculos(0).Text)
  Else
    dtDatas(0) = Empty
  End If

  If (EData(txtCalculos(1).Text)) Then  '// Se a data final for uma data v�lida
    dtDatas(1) = CDate(txtCalculos(1).Text)
  Else                                  '// A data final n�o pode ficar em branco
    MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, "Data Inicial")
    bReturn = False
    GoTo VerFiltro_Erro
  End If
  '// Verifica se a data inicial n�o � posterior a data final
  '//
  If (DateDiff("d", dtDatas(0), dtDatas(1)) < ZERO) Then
    MsgFunc ResolveResString(139, resUM, "Final", resDOIS, "Inicial")
    bReturn = False
    GoTo VerFiltro_Erro
  End If
  '// Verifica se o per�odo especificado n�o cont�m o movimento conferido
  '//
  dtMesAno = dtDatas(0)
  While (DateDiff("d", dtMesAno, FirstDay(dtDatas(1))) >= ZERO)
    If (MovConferido(DataToStr(dtMesAno), "KIF")) Then    '// Movimento do M�s j� conferido
      bReturn = False                                     '// N�o � poss�vel fazer a gera��o
      GoTo VerFiltro_Erro
    End If
    dtMesAno = DateAdd("m", 1, dtMesAno)
  Wend
  '// Inicia a montagem da instru��o
  '//
  strString = "SELECT * FROM Lan�amentos WHERE "
  If ((Not IsEmptyDate(dtDatas(0))) And (Not IsEmptyDate(dtDatas(1)))) Then
    If (DateDiff("d", dtDatas(0), dtDatas(1)) = ZERO) Then
      AppendStr strString, "Vencimento = #" & InverteData(dtDatas(0)) & "#"
    Else
    
    End If
  ElseIf ((IsEmptyDate(dtDatas(0))) And (Not IsEmptyDate(dtDatas(1)))) Then
  End If
  
VerFiltro_Erro:
  If (Err.Number) Then
    VBErros NUL
  End If
  VerFiltro = bReturn
End Function

' FUNCTION..: SeleDados
' Objetivo..: Verifica os dados digitados pelo usu�rio e cria a instru��o
'             SELECT para a sele��o de dados.
' Argumento.: [strInst]: Ponteiro string que receber� a instru��o de sele��o
' Retorna...: True se a fun��o obtiver sucesso e criar a instru��o corretamente
'             False se algum erro ocorrer e n�o for poss�vel criar a instru��o
'             de sele��o.
' ------------------------------------------------------------------------------
Private Function SeleDados(strInst As String) As Boolean
Dim datInit As Date               '// Data Inicial
Dim datFim  As Date               '// Data Final
Dim strData As String             '// Instru��o de Sele��o pelas datas

  SeleDados = False

  strInst = "SELECT * FROM " & cboCalculos(1).Text & " WHERE Pagamento IS NOT NULL"

  If (cboCalculos(2).ListIndex < 2) Then            '// 2 == Ambos
    If (cboCalculos(2).ListIndex = 0) Then          '// 0 == Pagos
      AppendStr strInst, " AND PagRec = 'P'"
    Else                                            '// 1 == Recebidos
      AppendStr strInst, " AND PagRec = 'R'"
    End If
  End If

  '// Verificando os dados digitados pelo usu�rio...

  If (IsValid(txtCalculos(2).Text)) Then      '// Campo c�digo do Lan�amento
    AppendStr strInst, " AND C�digo = " & txtCalculos(2).Text
  End If

  If (IsValid(txtCalculos(3).Text)) Then      '// Campo c�digo da Conta
    Concat strInst, " AND Conta = ", txtCalculos(3).Text
  End If

  If (Len(txtCalculos(4).Text)) Then          '// Campo Empresa
    Concat strInst, " AND Empresa = '", txtCalculos(4).Text, "'"
  End If

  If (IsValid(txtCalculos(0).Text)) Then
    If (EData(txtCalculos(0).Text)) Then
      datInit = CDateDef(txtCalculos(0).Text, Empty)
    Else
      MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, lblCalculos(1).Caption)
      Exit Function
    End If
  End If

  If (IsValid(txtCalculos(1).Text)) Then
    If (EData(txtCalculos(1).Text)) Then
      datFim = CDateDef(txtCalculos(1).Text, Empty)
    Else
      MsgFunc ResolveResString(IDS_DATAINVALIDA, resDOIS, lblCalculos(2).Caption)
      Exit Function
    End If
  End If

  If ((Not IsEmptyDate(datInit)) And (Not IsEmptyDate(datFim))) Then
    If (DateDiff(DD_DIA, datInit, datFim) = ZERO) Then
      Concat strInst, " AND ", cboCalculos(0).Text, " = #", InverteData(datInit), "#"
    Else
      Concat strInst, " AND (", cboCalculos(0).Text, " BETWEEN #", InverteData(datInit), _
                      "# AND #", InverteData(datFim), "#)"
    End If
  ElseIf ((Not IsEmptyDate(datInit)) And (IsEmptyDate(datFim))) Then
    Concat strInst, " AND ", cboCalculos(0).Text, " >= #", InverteData(datInit), "#"
  ElseIf ((IsEmptyDate(datInit)) And (Not IsEmptyDate(datFim))) Then
    Concat strInst, " AND ", cboCalculos(0).Text, " <= #", InverteData(datFim), "#"
  End If

  '// Instru��o de Sele��o criada com sucesso

  SeleDados = True
  
End Function

' FUNCTION..: ShowInListView
' Objetivo..: Exibe os registros selecionados no controle ListView
' Retorna...: True se encontrar algum registro, False se n�o.
' --------------------------------------------------------------------------
Private Function ShowInListView(Msg As Boolean) As Boolean
Dim strSelect As String           '// Instru��o de sele��o de Dados

  If (SeleDados(strSelect)) Then
    Dim lIndex   As Long          '// �ndice dos �tens no ListView

    '// Preenchendo o controle ListView com os dados encontrados no cadastro

    SetPtr vbArrowHourglass
    SimpleMsgBar LoadResString(13) & LoadResString(14)
    lvwCalculos.ListItems.Clear
    If (AbreRecordset(mrstLanc, strSelect) = WL_OK) Then

      '// Configura a origem da sele��o
      
      mlSource = IIf((cboCalculos(1).ListIndex = ZERO), SOURCE_LANC, SOURCE_DUPL)
      CompleteListView mrstLanc
      ShowInListView = True
    ElseIf (UltimoRetorno() = WL_NORECORD) Then
      If Msg Then MsgFunc LoadResString(IDS_RECORDNOTFOUND)
    End If
  End If
  SetPtr vbDefault
  MsgBar Me.Caption
  
End Function

' SUB.......: CompleteListView
' Objetivo..: Completa o controle ListView com os dados do Recordset.
' Argumento.: [rstDados]: Recordset com os dados que ser�o exibidos.
' -------------------------------------------------------------------------
Private Sub CompleteListView(rstDados As Object)
Dim lIndex As Long            '// �ndice das linhas do ListView
Dim sCod   As String          '// Usada para obter o valor do campo de chave das
                              '// tabelas

  sCod = IIf((mlSource = SOURCE_LANC), "C�digo", "Nota")
  rstDados.MoveFirst
  Do
    Inc lIndex
    lvwCalculos.ListItems.add lIndex, , StrZero(GetValue(rstDados, sCod, 0), 6), , IDI_OK
    lvwCalculos.ListItems(lIndex).SubItems(1) = GetValue(rstDados, "Empresa", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(2) = GetValue(rstDados, "Tipo", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(3) = GetValue(rstDados, "Descri��o", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(4) = GetValue(rstDados, "Emiss�o", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(5) = GetValue(rstDados, "Vencimento", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(6) = GetValue(rstDados, "Pagamento", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(7) = GetValue(rstDados, "Libera��o", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(8) = Format$(GetValue(rstDados, "Valor Original", 0), FCURRENCY)
    lvwCalculos.ListItems(lIndex).SubItems(9) = Format$(GetValue(rstDados, "Acr�scimo", 0), FCURRENCY)
    lvwCalculos.ListItems(lIndex).SubItems(10) = Format$(GetValue(rstDados, "Abatimento", 0), FCURRENCY)
    lvwCalculos.ListItems(lIndex).SubItems(11) = GetValue(rstDados, "Banco", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(12) = GetValue(rstDados, "Conta", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(13) = GetValue(rstDados, "Controle", NUL)
    lvwCalculos.ListItems(lIndex).SubItems(14) = GetValue(rstDados, "Situa��o", NUL)
    rstDados.MoveNext
  Loop Until (rstDados.EOF)
  
End Sub

' FUNCTION..: ExportarDados
' Objetivo..: Realiza o expurgo dos dados, de acordo com as sele��es do usu�rio.
' Retorna...: True se a fun��o obtiver sucesso, False se n�o.
' ----------------------------------------------------------------------------------
Private Function ExportarDados() As Boolean
Dim dbDatabase As Object        '// Banco de Dados externo, onde os dados ser�o gravados
Dim osdSave    As OPENSAVEDIALOG  '// Para a caixa de di�logo Salvar Como...

  osdSave.lnghWnd = Me.hWnd
  osdSave.strFiltro = "Banco de Dados Access (*.mdb)|*.mdb|"
  osdSave.strFile = AddSepDir(CurDir$()) & "Exp" & Format$(Date, "dd-mm-yyyy") & ".mdb"
  osdSave.strInitialDir = CurDir$()
  osdSave.strTitulo = "Salvar Backup Como..."
  osdSave.lngFlags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_OVERWRITEPROMPT Or OFN_NOCHANGEDIR

  '// Abre a janela Salvar Como... para que o usu�rio indique um nome de arquivo
  '// para o Banco de Dados de Backup. A Fun��o gera um nome com a data atual assim,
  '// se o usu�rio preferir, pode gerar um arquivo diferente de backup e separ�-los por
  '// data. Crio um arquivo de Banco de Dados com o nome e caminho indicados e tento
  '// abr�-lo em modo exclusivo. O usu�rio pode ser idiota o suficiente para indicar o
  '// mesmo nome do arquivo utilizado atualmente no Sistema. Abrindo o arquivo no modo
  '// exclusivo far� com que a abertura falhe - pois o arquivo j� est� aberto - e a
  '// tabela n�o seja gerada.
  
  If (SaveAsDialogBox(osdSave)) Then

    SetPtr vbHourglass
    If (Not ArquivoExiste(osdSave.strFile)) Then  '// Se o arquivo n�o existir, cria
      CrieDatabase Left$(osdSave.strFile, osdSave.intFileOffset), osdSave.strFileTitle
    End If
    
    If (AbreDatabase(dbDatabase, osdSave.strFile, True) = WL_OK) Then
      Dim lTmp As Long
      Dim sTmp As String
      
      On Error Resume Next

      If (mlSource = SOURCE_LANC) Then
        sTmp = "Lan�amentos"
      Else
        sTmp = "Duplicatas"
      End If
      
      lTmp = dbDatabase.TableDefs(sTmp).Fields.Count
      If (Err().Number) Then
        Err().Clear
        If (Not CriaTabelaBackup(dbDatabase, sTmp)) Then
          MsgFunc ResolveResString(IDS_CRIETABELAERRO, resUM, sTmp)
          GoTo ExportaDados_Erro
        End If
      End If

      '// Passa os registros selecionados pelo usu�rio para a tabela no arquivo
      '// de Backup

      If (Not GravaDados(dbDatabase, sTmp)) Then
        MsgFunc ResolveResString(255, resUM, sTmp)
        GoTo ExportaDados_Erro
      Else
        MsgFunc ResolveResString(256, resUM, osdSave.strFile)

        '// Assim que o usu�rio fecha a caixa de mensagens procedo o refresh
        '// da lista de registros do controle ListView. Ele deve mostrar, agora
        '// somente os dados que n�o foram exclu�dos pelo usu�rio.
        
        ShowInListView False
        ExportarDados = True
      End If
    End If
    
  End If

ExportaDados_Erro:
  If (Err().Number) Then
    DAOErros NUL
    ExportarDados = False
  End If
  CloseDatabase dbDatabase
  SetPtr vbDefault
End Function

' FUNCTION..: CriaTabelaBackup
' Objetivo..: Cria a tabela no Banco de Dados de Backup com a mesma estrutura da
'             tabela atual.
' Argumentos: [dbBackup]: Vari�vel do arquivo de backup.
'             [strNome ]: Nome da tabela que deve ser criada.
' Retorna...: True se obtiver sucesso, False se n�o.
' -------------------------------------------------------------------------------
Private Function CriaTabelaBackup(dbBackup As Object, strNome As String) As Boolean
Dim tdfBackup As TableDef             '// Estrutura da tabela
Dim fldBackup As Object                '// Campo da tabela
Dim lFields   As Long                 '// Utilizada no Loop

  On Error GoTo CriaTabelaBackup_Erro
  Set tdfBackup = dbBackup.CreateTableDef(strNome)
  
  For lFields = ZERO To mrstLanc.Fields.Count - 1
    Set fldBackup = tdfBackup.CreateField(mrstLanc(lFields).Name, mrstLanc(lFields).Type)

    If (mrstLanc(lFields).Type > dbDate) Then     '// (dbBinary, dbText, dbLongBinary, dbMemo)
      fldBackup.Size = mrstLanc(lFields).Size
      fldBackup.AllowZeroLength = True
    End If
    
    fldBackup.Required = False
    tdfBackup.Fields.Append fldBackup
  Next

  dbBackup.TableDefs.Append tdfBackup
  CriaTabelaBackup = True
  
CriaTabelaBackup_Erro:
  If (Err().Number) Then
    DAOErros NUL
    CriaTabelaBackup = False
  End If
  Set fldBackup = Nothing
  Set tdfBackup = Nothing
End Function

' FUNCTION..: GravaDados
' Objetivo..: Grava os dados escolhidos pelo usu�rio na tabela de Backup no
'             arquivo de Banco de Dados criado.
' Argumentos: [dbBackup]: Banco de Dados de backup.
'             [sName   ]: Nome da tabela que deve ser aberta.
' Retorna...: True se obtiver sucesso, False se n�o.
' ---------------------------------------------------------------------------
Private Function GravaDados(dbBackup As Object, sName As String) As Boolean
Dim rstDest As Object          '// Recordset de destino
Dim lCount  As Long               '// Contador do Loop
Dim fldBkp  As Object              '// Utilizada no Loop dos Backups

  On Error GoTo GravaDados_Erro
  Set rstDest = dbBackup.OpenRecordset(sName, dbOpenDynaset)
  mrstLanc.MoveFirst

  '// O Loop � executado para todos os registros encontrados no Recordset.
  '// Como o registro est� na mesma posi��o das linhas do ListView, virifico
  '// se a linha correspondente est� marcada para exclus�o (SmallIcon = IDI_OK).
  '// Se estiver o registro � gravado na tabela do banco de dados de backup e
  '// o registro � exclu�do.

  Do
    Inc lCount

    If (lvwCalculos.ListItems(lCount).SmallIcon = IDI_OK) Then
      rstDest.AddNew
      For Each fldBkp In mrstLanc.Fields
        rstDest(fldBkp.Name).Value = fldBkp.Value
      Next
      rstDest.Update
      Set fldBkp = Nothing        '// Libera a �ltima refer�ncia ao campo
      mrstLanc.Delete             '// Exclui o registro atual
    End If
    
    mrstLanc.MoveNext
  Loop Until (mrstLanc.EOF)
  GravaDados = True
  
GravaDados_Erro:
  If (Err().Number) Then
    DAOErros NUL
    GravaDados = False
  End If
  FechaRecordset rstDest
  Set fldBkp = Nothing
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
