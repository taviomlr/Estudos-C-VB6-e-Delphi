VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frptCheque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatórios de Cheques"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   Icon            =   "RptChq.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTab 
      Caption         =   "Cópia de Cheque"
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
      Height          =   2250
      Left            =   90
      TabIndex        =   26
      Top             =   315
      Width           =   7755
      Begin VB.CheckBox chkCheque 
         Caption         =   "Imprimir Centro de Custo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   5235
         TabIndex        =   8
         Top             =   645
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkCheque 
         Caption         =   "Imprimir Conta"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   5235
         TabIndex        =   7
         Top             =   375
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   0
         Top             =   315
         Width           =   975
      End
      Begin VB.CheckBox chkCheque 
         Caption         =   "Um cheque por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   5235
         TabIndex        =   9
         Top             =   1605
         Width           =   2055
      End
      Begin VB.CheckBox chkCheque 
         Caption         =   "Imprimir já impressos?"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   5235
         TabIndex        =   10
         Top             =   1860
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "RptChq.frx":0C42
         Left            =   1200
         List            =   "RptChq.frx":0C44
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1380
         Width           =   1575
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Index           =   4
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Index           =   3
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1020
         Width           =   975
      End
      Begin VB.ComboBox cboCheque 
         Height          =   315
         ItemData        =   "RptChq.frx":0C46
         Left            =   1200
         List            =   "RptChq.frx":0C48
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Index           =   2
         Left            =   3480
         MaxLength       =   9
         TabIndex        =   2
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   1
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "Relatório:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   23
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label lblDescCheque 
         Caption         =   "lblDescCheque(0)"
         Height          =   195
         Index           =   0
         Left            =   2280
         TabIndex        =   19
         Top             =   375
         UseMnemonic     =   0   'False
         Width           =   3240
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "Histórico:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   25
         Top             =   1860
         Width           =   660
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "Chq. Final:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2400
         TabIndex        =   21
         Top             =   720
         Width           =   750
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "Chq. Inicial:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblCheque 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   375
         Width           =   510
      End
   End
   Begin VB.Frame fraBotoes 
      Height          =   2775
      Left            =   7950
      TabIndex        =   34
      Top             =   -90
      Width           =   1395
      Begin VB.CommandButton cmdCheque 
         Cancel          =   -1  'True
         Caption         =   "#"
         Height          =   375
         Index           =   2
         Left            =   75
         TabIndex        =   17
         Top             =   930
         Width           =   1215
      End
      Begin VB.CommandButton cmdCheque 
         Caption         =   "Im&primir"
         Height          =   375
         Index           =   1
         Left            =   75
         TabIndex        =   16
         Top             =   540
         Width           =   1215
      End
      Begin VB.CommandButton cmdCheque 
         Caption         =   "&Visualizar..."
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   15
         Top             =   150
         Width           =   1215
      End
   End
   Begin VB.Frame fraImpressoraCheque 
      Caption         =   "Impressora de Cheque"
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
      Height          =   2265
      Left            =   90
      TabIndex        =   28
      Top             =   315
      Width           =   7770
      Begin VB.ComboBox cboImpressoraCheque 
         Height          =   315
         Index           =   1
         ItemData        =   "RptChq.frx":0C4A
         Left            =   1080
         List            =   "RptChq.frx":0C69
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cboImpressoraCheque 
         Height          =   315
         Index           =   0
         ItemData        =   "RptChq.frx":0C88
         Left            =   1080
         List            =   "RptChq.frx":0C8F
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtImpressoraCheque 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtImpressoraCheque 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDescBanco 
         Caption         =   "lblDescBanco"
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblTxtImpressaoCheque 
         Caption         =   "Porta  COM:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblTxtImpressaoCheque 
         Caption         =   "Modelo:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblTxtImpressaoCheque 
         Caption         =   "Cheque:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblTxtImpressaoCheque 
         Caption         =   "Banco:"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   615
      End
   End
   Begin ComctlLib.TabStrip tabCheque 
      Height          =   2670
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4710
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cópia de Cheque"
            Key             =   "copia"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Relatório de Cheques"
            Key             =   "relatorio"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Impressão de Cheques"
            Key             =   "impressao"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Impressora de Cheque"
            Key             =   "impressoracheque"
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
   Begin VB.CheckBox chkCheque 
      Caption         =   "Quebrar por Banco"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   5445
      TabIndex        =   35
      Top             =   1530
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frptCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SEC_RPTCHEQUES$ = "RptCheques"        '// Seção do relatório de cheques no .ini
Private Const KEY_BANCO$ = "Banco"                  '// Chave para o número do Banco
Private Const KEY_CHQINI$ = "ChqIni"                '// Chave para o cheque inicial
Private Const KEY_CHQFIM$ = "ChqFim"                '// Chave para o cheque final
Private Const KEY_DTINI$ = "DtIni"                  '// Chave para a data inicial
Private Const KEY_DTFIM$ = "DtFim"                  '// Chave para a data final
Private Const KEY_ORDEM$ = "Ordem"                  '// Chave para a ordem dos dados
Private Const KEY_DESC$ = "Desc"                    '// Chave para a descrição da cópia
Private Const KEY_UMPORFOLHA$ = "UmPorPagina"       '// Chave do CheckBox de Um Cheque por Página

Private Const F_NORMAL = 0            'Fonte Normal
Private Const F_ITALICO = 1           'Fonte Itálico
Private Const F_NEGRITO = 2           'Fonte Negrito
Private Const F_NEGRITOITALICO = 3    'Fonte Negrito Itálico

Private m_intOrdem As Integer         ' Ordem padrão para Relatório de Cheques
Private m_intDesc  As Integer         ' Define o padrão para Cópia de Cheques (Descrição ou Lançamentos)

' EVENT.....: cboCheque_Click
' Objetivo..: Grava as alterações do usuário na combo para utilizar na
'             gravação dos valores padrão no encerramento da janela.
' ------------------------------------------------------------------------------------
Private Sub cboCheque_Click()
  If (tabCheque.SelectedItem.Key = "copia") Then
    m_intDesc = GetItemData(cboCheque)
    If cboCheque.Text <> "Lançamentos" Then
      chkCheque(3).Visible = False
      chkCheque(4).Visible = False
    Else
      If CentrodeCusto(MFinanceiro) Then chkCheque(3).Visible = True
      chkCheque(4).Visible = True
    End If
  Else
    m_intOrdem = GetItemData(cboCheque)
  End If
End Sub

' EVENT.....: cboCheque_GotFocus
' Objetivo..: Exibe mensagens descritivas na barra de status
'             do Sistema
' ------------------------------------------------------------------------------------
Private Sub cboCheque_GotFocus()
  ChequeMsgStatus cboCheque.TabIndex
End Sub



  

' EVENT.....: cmdCheque_Click
' Objetivo..: Executa as funções do botões da janela.
' ------------------------------------------------------------------------------------
Private Sub cmdCheque_Click(Index As Integer)

  If (Index < 2) Then
    cmdCheque(0).Enabled = False
    cmdCheque(1).Enabled = False
    cmdCheque(2).Caption = LoadResString(IDS_CANCELAR)

    ImprimeCheques IIf((Index), wrToPrinter, wrToWindow)

    cmdCheque(0).Enabled = True
    cmdCheque(1).Enabled = True
    cmdCheque(2).Caption = LoadResString(IDS_FECHAR)
  Else
    If (cmdCheque(0).Enabled) Then
      Unload Me
    Else
      SimpleMsgBar LoadResString(171)
    End If
  End If

End Sub

' EVENT.....: Form_Load
' Objetivo..: Configura a janela para sua abertura
' ------------------------------------------------------------------------------------
Private Sub Form_Load()
  
  Dim sTmp                    As String
  Dim UtilizaImpressoraCheque As Boolean
  
  lblDescCheque(0).Caption = NUL

  ' Configurando valores padrão para os campos da janela

  m_intOrdem = ReadSettings(SEC_RPTCHEQUES, KEY_ORDEM, "1")
  m_intDesc = ReadSettings(SEC_RPTCHEQUES, KEY_DESC, "1")

  ' Campo Código do Banco

  txtCheque(0).Text = ReadSettings(SEC_RPTCHEQUES, KEY_BANCO, NUL)

  ' Campo Cheque Inicial e Final

  txtCheque(1).Text = ReadSettings(SEC_RPTCHEQUES, KEY_CHQINI, NUL)
  txtCheque(2).Text = ReadSettings(SEC_RPTCHEQUES, KEY_CHQFIM, NUL)

  '// Campo Data Inicial e Data Final

  txtCheque(3).Text = ReadSettings(SEC_RPTCHEQUES, KEY_DTINI, NUL)
  txtCheque(4).Text = ReadSettings(SEC_RPTCHEQUES, KEY_DTFIM, NUL)

  ' Um ou mais cheques por página

  sTmp = ReadSettings(SEC_RPTCHEQUES, KEY_UMPORFOLHA, "0")
  chkCheque(0).value = CIntDef(sTmp, vbUnchecked)

  ' Caption para o botão fechar do formulário
  cmdCheque(2).Caption = LoadResString(IDS_FECHAR)

  ' Definindo o Tab visível inicial

  tabCheque.Tabs("copia").Selected = True
  LoadResOptions 1084, cboTipo, True, 1
  CenterForm Me                   'Centraliza o formulário na tela
  
  UtilizaImpressoraCheque = Configuracao("Utiliza Impressora de Cheque", False)
  
  If UtilizaImpressoraCheque = False Then
    tabCheque.Tabs.Remove 4
  Else
    cboImpressoraCheque(0).Text = "Chronos"
    cboImpressoraCheque(1).Text = "1"
    lblDescBanco.Caption = " "
  End If

End Sub

' EVENT.....: Form_QueryUnload
' Objetivo..: Verifica se o formulário pode ser fechado.
' ------------------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (UnloadMode = vbFormControlMenu) Then
    If (Not cmdCheque(0).Enabled) Then
      Call SendKeysEx(Chr$(vbKeyEscape))
      Cancel = True
    End If
  End If
  MsgBar MsgBoxCaption
End Sub

' EVENT.....: Form_Unload
' Objetivo..: Grava as configurações atuais da janela e finaliza a
'             variável global
' ------------------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)

  WriteSettings SEC_RPTCHEQUES, KEY_BANCO, txtCheque(0).Text
  WriteSettings SEC_RPTCHEQUES, KEY_CHQINI, txtCheque(1).Text
  WriteSettings SEC_RPTCHEQUES, KEY_CHQFIM, txtCheque(2).Text
  WriteSettings SEC_RPTCHEQUES, KEY_DTINI, txtCheque(3).Text
  WriteSettings SEC_RPTCHEQUES, KEY_DTFIM, txtCheque(4).Text
  WriteSettings SEC_RPTCHEQUES, KEY_UMPORFOLHA, chkCheque(0).value
  WriteSettings SEC_RPTCHEQUES, KEY_ORDEM, CStr(m_intOrdem)
  WriteSettings SEC_RPTCHEQUES, KEY_DESC, CStr(m_intDesc)

  Set frptCheque = Nothing

End Sub

Private Sub Label1_Click()

End Sub

' EVENT.....: tabCheque_Click
' Objetivo..: Exibe os controles correspondentes a cada tipo de
'             relatório.
' ------------------------------------------------------------------------------------
Private Sub tabCheque_Click()

  If (tabCheque.SelectedItem.Key = "copia") Then
    fraImpressoraCheque.Visible = False
    fraTab.Visible = True
    LoadResOptions 1019, cboCheque, True          '// Lançamento[2] ou Cheque[1]
    cboCheque.ListIndex = IndexOfItemData(cboCheque, CLng(m_intDesc))    '// Valor padrão
    lblCheque(3).Caption = LoadResString(198)     '// "Histórico"
    fraTab.Caption = LoadResString(197)           '// "Cópia de Cheques"
    cboTipo.Visible = False
    lblCheque(6).Visible = False
    If Not CentrodeCusto(MFinanceiro) Then chkCheque(3).Visible = False
    If cboCheque.Text <> "Lançamentos" Then
      chkCheque(3).Visible = False
      chkCheque(4).Visible = False
    Else
      If CentrodeCusto(MFinanceiro) Then chkCheque(3).Visible = True
      chkCheque(4).Visible = True
    End If
  ElseIf (tabCheque.SelectedItem.Key = "relatorio") Then
    fraImpressoraCheque.Visible = False
    fraTab.Visible = True
    LoadResOptions 1020, cboCheque, True          '// Cheque[1] ou Data[2]
    cboCheque.ListIndex = IndexOfItemData(cboCheque, CLng(m_intOrdem))   '// Valor padrão
    lblCheque(3).Caption = LoadResString(199)     '// "Ordem"
    fraTab.Caption = LoadResString(196)           '// "Relatório de Cheques"
    cboTipo.Visible = True
    lblCheque(6).Visible = True
  ElseIf (tabCheque.SelectedItem.Key = "impressoracheque") Then
    fraImpressoraCheque.Visible = True           '// Imprimir na Impressora de Cheque
    fraTab.Visible = False
  Else
    fraImpressoraCheque.Visible = False
    fraTab.Visible = True
    cboTipo.Visible = False
    fraTab.Caption = LoadResString(243)           '// "Impressão de Cheques"
    lblCheque(6).Visible = False
  End If

  '// Para impressão de Cheques a caixa de combinação não deve aparecer

  lblCheque(3).Visible = (tabCheque.SelectedItem.Key <> "impressao")
  cboCheque.Visible = (tabCheque.SelectedItem.Key <> "impressao")
  chkCheque(2).Visible = (tabCheque.SelectedItem.Key <> "impressao")

  '// O CheckBox de quantidade de folhas por página só na Cópia de Cheque

  chkCheque(0).Visible = (tabCheque.SelectedItem.Key = "copia")
  'chkCheque(1).Visible = (tabCheque.SelectedItem.Key = "relatorio")
    
End Sub

' EVENT.....: txtCheque_Change
' Objetivo..: Exibe o nome do banco no Label de descrição.
' ------------------------------------------------------------------------------------
Private Sub txtCheque_Change(Index As Integer)
  If (Index = 0) Then
    AssocValue "Nome", "Bancos", "Banco = %s", Array(txtCheque(0).Text), lblDescCheque(0)
  End If
End Sub

' EVENT.....: txtCheque_GotFocus
' Objetivo..: Exibe mensagens descritivas na barra de Status do programa
' ------------------------------------------------------------------------------------
Private Sub txtCheque_GotFocus(Index As Integer)
  Selecione txtCheque(Index)
  ChequeMsgStatus txtCheque(Index).TabIndex
End Sub

' EVENT.....: txtCheque_KeyDown
' Objetivo..: Abre a janela de pesquisa para bancos e cheques.
' ------------------------------------------------------------------------------------
Private Sub txtCheque_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim strSelDados As String

  If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
    Select Case Index
    '
    ' Campo código do Banco
    Case 0
      If tabCheque.Tabs(3).Selected = True Then
        PCampo "Bancos", "SELECT * FROM Bancos WHERE (Câmara BETWEEN (SELECT TOP 1 Número FROM ChqModelos ORDER BY Número) AND (SELECT TOP 1 Número FROM ChqModelos ORDER BY Número DESC))", pbCampo, txtCheque(0), "Banco"
      Else
        PCampo "Bancos", "Bancos", pbCampo, txtCheque(0), "Banco"
      End If
    '
    ' Campo Cheque Inicial, Cheque Final
    Case 1, 2
      If (IsValid(txtCheque(0).Text)) Then
        strSelDados = "SELECT * FROM Cheque WHERE Banco = " & txtCheque(0).Text
      Else
        strSelDados = "Cheque"
      End If
      PCampo "Cheques", strSelDados, pbCampo, txtCheque(Index), "Cheque"
    '
    End Select
  End If

End Sub

' EVENT.....: txtCheque_KeyPress
' Objetivo..: Faz a validação dos caracteres digitados pelo usuário.
' ------------------------------------------------------------------------------------
Private Sub txtCheque_KeyPress(Index As Integer, KeyAscii As Integer)

  Select Case Index
  '
  ' Campo Código do Banco
  Case 0
    SetMascara KeyAscii, txtCheque(0).SelStart, fMask("Bancos", "Banco")
  '
  ' Campos Cheque Inicial e Cheque Final
  Case 1
    SetMascara KeyAscii, txtCheque(1).SelStart, fMask("Cheque", "Cheque")
  '
  Case 2
    SetMascara KeyAscii, txtCheque(2).SelStart, fMask("Cheque", "Cheque"), txtCheque(1).hWnd
  '
  ' Campo de Data Inicial e Final
  Case 3, 4
    SetMascara KeyAscii, txtCheque(Index).SelStart, MASK_DATA
  '
  End Select

End Sub

' SUB.......: ChequeMsgStatus
' Objetivo..: Exibe mensagens de auxilio ao usuário na barra de status do
'             programa.
' Argumento.: [intTabIndex]: Propriedade TabIndex do controle que recebe o foco.
' ---------------------------------------------------------------------------------
Private Sub ChequeMsgStatus(intTabIndex As Integer)

  Select Case intTabIndex
  '
  ' Campo Código do Banco
  Case 2
    MsgBar LoadResString(152) & ResolveResString(75, resUM, "Bancos")
  '
  ' Cheque Inicial e Final
  Case 5, 7
    MsgBar LoadResString(190) & ResolveResString(75, resUM, "Cheques")
  '
  ' Datas Inicial e Final
  Case 9: MsgBar ResolveResString(161, resUM, "de emissão")

  Case 11: MsgBar ResolveResString(162, resUM, "de emissão")
  '
  ' Histórico, Ordem    Dependendo do Tab que estiver a frente
  Case 13
    If (tabCheque.SelectedItem.Key = "copia") Then
      MsgBar "Como imprimir o histórico do cheque"
    Else
      MsgBar "Ordem do relatório"
    End If
  '
  ' Imprimir conta
  Case 10
    MsgBar "Imprime Código e Descrição da Conta"
  '
  ' Um cheque por página
  Case 14
    MsgBar "Faz uma quebra de página a cada cheque impresso"
  '
  End Select

End Sub

' SUB.......: ImprimeCheques
' Objetivo..: Cria o filtro para os cheques selecionados pelo usuário e,
'             dependendo do tipo do relatório, gera o arquivo auxiliar
'             que servirá de base de dados para a impressão dos cheques.
' Argumento.: [pdeDestino]: Destino da impressão.
' -------------------------------------------------------------------------
Private Sub ImprimeCheques(pdeDestino As PrintDestinoEnum)

  SetPtr vbHourglass

  Select Case (tabCheque.SelectedItem.Key)
    Case "relatorio"
      If cboTipo.Text = "Sintético" Then
        Call FiltraCheques(pdeDestino)
      Else
        Call FiltraAnalitico(pdeDestino)
      End If

    Case "copia"
      If (GetItemData(cboCheque) = 1) Or (GetItemData(cboCheque) = 3) Then          '// 1 = Descrição do Cheque
        Call FiltraCheques(pdeDestino)
      Else
        Call FiltroCopiaCheques(pdeDestino)
      End If

    Case "impressao"
      If (CLngDef(txtCheque(0).Text) = ZERO) Then
        MsgFunc "É necessário escolher um único banco para a impressão dos cheques"
      Else
        Call FiltraCheques(pdeDestino)
      End If
    Case "impressoracheque"
      'Call FiltraImpressoraCheques
  End Select

  MsgBar MsgBoxCaption
  SetPtr vbDefault

End Sub

' FUNCTION..: TempRelatorio
' Objetivo..: Cria a tabela auxiliar para o Relatório de Cheques.
' Argumento.: [rsTemp]: Recordset que receberá a tabela auxiliar.
' Retorna...: True se obtiver sucesso, False se não.
' ------------------------------------------------------------------------------------
Private Function TempRelatorio(rstEmp As Object) As Boolean
Dim fsRpt(5) As FieldStruct

  AppendVar fsRpt(0), "Banco", dbLong
  AppendVar fsRpt(1), "Nome", dbText, 40
  AppendVar fsRpt(2), "Cheque", dbLong
  AppendVar fsRpt(3), "Data", dbDate
  AppendVar fsRpt(4), "Valor", dbCurrency
  AppendVar fsRpt(5), "Nominal", dbText, 60

  If (CrieAux(rstEmp, fsRpt())) Then
    TempRelatorio = True
  Else
    MsgFunc LoadResString(174), vbExclamation
  End If

End Function

' FUNCTION..: TempCopia
' Objetivo..: Cria a tabela auxiliar responsável pelos dados de
'             impressão da cópia de cheque.
' Argumento.: [rsCopia]: Recordset que receberá a tabela.
' Retorna...: True se obtiver sucesso, False se não.
' ------------------------------------------------------------------------------------
Private Function TempCopia(rsCopia As Object) As Boolean
Dim fsCp(10) As FieldStruct

  AppendVar fsCp(0), "Banco", dbLong              '// Código do Banco
  AppendVar fsCp(1), "Cheque", dbLong             '// Número do Cheque
  AppendVar fsCp(2), "Valor", dbCurrency          '// Valor do Cheque
  AppendVar fsCp(3), "Data", dbDate               '// Data do Cheque
  AppendVar fsCp(4), "Nome", dbText, 40           '// Nome do Banco
  AppendVar fsCp(5), "Nominal", dbText, 60        '// Nominativo do cheque
  AppendVar fsCp(6), "Extenso", dbText, 255       '// Extenso do valor do cheque
  AppendVar fsCp(7), "DtExt", dbText, 70          '// Extenso da data (Cidade, Dia ' de ' Mês ' de ' Ano)
  AppendVar fsCp(8), "Desc", dbMemo               '// Descrição do Cheque
  AppendVar fsCp(9), "Valor Total", dbCurrency          '// Valor do Cheque
  AppendVar fsCp(10), "Extenso Total", dbText, 255       '// Extenso do valor do cheque
  If (CrieAux(rsCopia, fsCp())) Then
    TempCopia = True
  Else
    MsgFunc LoadResString(174), vbExclamation
  End If

End Function

' FUNCTION..: TempCopiaLan
' Objetivo..: Cria a tabela auxiliar responsável pela impressão do
'             relatório de Cópia de Cheque. Esta tabela contém os
'             lançamentos efetuados com os cheques quando o relatório
'             deve apresentar estes dados.
' Argumento.: [rsLanc]: Recordset que receberá a tabela auxiliar.
' Retorna...: True se obtiver sucesso, False se não.
' ------------------------------------------------------------------------------------
Private Function TempCopiaLan(rsLanc As Object) As Boolean
Dim fsCp(12) As FieldStruct

  AppendVar fsCp(0), "Banco", dbLong           '// Código do Banco
  AppendVar fsCp(1), "Nome", dbText, 40        '// Nome do Banco
  AppendVar fsCp(2), "Cheque", dbLong          '// Número do cheque
  AppendVar fsCp(3), "Data", dbDate            '// Data do lançamento
  AppendVar fsCp(4), "Valor", dbCurrency       '// Valor do lançamento
  AppendVar fsCp(5), "Lancto", dbText, 40      '// Código do lançamento e Tipo
  AppendVar fsCp(6), "Emp", dbText, 15         '// Nome Fantasia da Empresa
  AppendVar fsCp(7), "Desc", dbText, 80        '// Descrição do lançamento
  AppendVar fsCp(8), "Conta", dbLong           '// Código da conta contábil
  AppendVar fsCp(9), "CtDesc", dbText, 40      '// Descrição da conta contábil
  AppendVar fsCp(10), "Custo", dbLong          '// Código do Centro de Custo
  AppendVar fsCp(11), "CsDesc", dbText, 40     '// Descrição do Centro de Custo
  'Protocolo Nr 102985 - Carlos Felippe Vernizze - 10/01/2011
  AppendVar fsCp(12), "Controle", dbText, 50   '// Controle

  If (CrieAux(rsLanc, fsCp())) Then
    TempCopiaLan = True
  Else
    MsgFunc LoadResString(174), vbExclamation
  End If

End Function

' FUNCTION..: TempImpressao
' Objetivo..: Cria a tabela temporária para a impressão de cheque.
' Argumento.: [rstTemp]: Recordset que receberá a tabela.
' Retorna...: True se obtiver sucesso, False se não.
' ------------------------------------------------------------------------------------
Private Function TempImpressao(rstTemp As Object) As Boolean
Dim fsImp(7) As FieldStruct

  AppendVar fsImp(0), "BcoChq", dbText, 20           '// Número do Banco e do Cheque
  AppendVar fsImp(1), "Valor", dbText, 25            '// Valor do cheque
  AppendVar fsImp(2), "Local", dbText, 50            '// Normalmente a cidade padrão e o dia
  AppendVar fsImp(3), "Mês", dbText, 20              '// Mês por extenso
  AppendVar fsImp(4), "Ano", dbText, 4               '// Ano da emissão
  AppendVar fsImp(5), "Nominal", dbText, 60          '// Nominativo do cheque
  AppendVar fsImp(6), "Extenso1", dbText, 100        '// Primeiro linha do extenso do cheque
  AppendVar fsImp(7), "Extenso2", dbText, 150        '// Segunda linha do extenso do cheque

  If (CrieAux(rstTemp, fsImp())) Then
    TempImpressao = True
  Else
    MsgFunc LoadResString(174), vbExclamation
  End If

End Function

' FUNCTION..: Envio
' Objetivo..: Verificar a necessidade de adicionar a instrução SQL
'             o filtro para não exibir registros já enviados
' Argumentos: [CampoData]:  String com o
' ----------------------------------------------------------------
Private Function Envio(CampoData As String) As String
  If chkCheque(2).value = vbUnchecked Then
    Envio = " AND ((" & CampoData & " is null) OR (" & CampoData & " = ''))"
  Else
    Envio = NUL
  End If
End Function

' SUB.......: FiltraCheques
' Objetivo..: Filtra os dados para o Relatório de Cheques
' Argumento.: [pdeDestino]: Destino da impressão.
' ------------------------------------------------------------------------------------
Private Sub FiltraCheques(pdeDestino As PrintDestinoEnum)
Dim strDupls   As String        '// Instrução de seleção de dados para Duplicatas
Dim strLanctos As String        '// Instrução de seleção de dados para Lançamentos
Dim strTransf  As String        '// Instrução de seleção de dados para Transf. Bancária
Dim rstDados   As Object        '// Recordset com os dados dos lançamentos
Dim rstAux     As Object        '// Recordset da tabela auxiliar
Dim qdfTemp    As QueryDef      '// Consulta da seleção dos dados
Dim nCodIni    As Long          '// Código Inicial
Dim nCodFim    As Long          '// Código Final
Dim dInicial   As Date          '// Data Inicial
Dim dFinal     As Date          '// Data Final


  SimpleMsgBar "Selecionando dados, aguarde..."

  strDupls = "SELECT D.Banco, D.Cheque, D.Pagamento As Data, " & _
               "SUM(D.[Valor Original] + D.Acréscimo - D.Abatimento) As Valor, " & _
               "D.Enviada As Impresso, D.Descrição  " & _
               "FROM Duplicatas As D WHERE D.PagRec = 'P'" & Envio("D.Enviada")
               
  strLanctos = "SELECT L.Banco, L.Cheque, L.Pagamento As Data, " & _
               "SUM(L.[Valor Original] + L.Acréscimo - L.Abatimento) As Valor, " & _
               "L.Enviado As Impresso, L.Descrição " & _
               "FROM Lançamentos As L WHERE L.PagRec = 'P'" & Envio("L.Enviado")

  strTransf = "SELECT T.Origem, T.Cheque, T.Data, SUM(T.Valor), " & _
              "T.Enviada As Impresso, T.Descrição " & _
              "FROM [Transf Bancária] As T"

  '// Verificando se o usuário indicou um Banco

  nCodIni = CLngDef(txtCheque(0).Text)
  If (nCodIni) Then
    Concat strDupls, " AND D.Banco = ", CStr(nCodIni)
    Concat strLanctos, " AND L.Banco = ", CStr(nCodIni)
    Concat strTransf, " WHERE T.Origem = ", CStr(nCodIni) & Envio("T.Enviada")
  Else
    Concat strTransf, " WHERE T.Origem > 0 " & Envio("T.Enviada")
  End If

  '// Verificando se o usuário filtrou por cheque

  nCodIni = CLngDef(txtCheque(1).Text)
  nCodFim = CLngDef(txtCheque(2).Text)

  If (CBool(nCodIni) And CBool(nCodFim)) Then
    If (nCodIni = nCodFim) Then
      Concat strDupls, " AND D.Cheque = ", CStr(nCodIni)
      Concat strLanctos, " AND L.Cheque = ", CStr(nCodIni)
      Concat strTransf, " AND T.Cheque = ", CStr(nCodIni)
    Else
      Concat strDupls, wsprintf(" AND (D.Cheque BETWEEN %l AND %l)", nCodIni, nCodFim)
      Concat strLanctos, wsprintf(" AND (L.Cheque BETWEEN %l AND %l)", nCodIni, nCodFim)
      Concat strTransf, wsprintf(" AND (T.Cheque BETWEEN %l AND %l)", nCodIni, nCodFim)
    End If
  ElseIf (CBool(nCodIni) And Not CBool(nCodFim)) Then
    Concat strDupls, " AND D.Cheque >= ", CStr(nCodIni)
    Concat strLanctos, " AND L.Cheque >= ", CStr(nCodIni)
    Concat strTransf, " AND T.Cheque >= ", CStr(nCodIni)
  ElseIf (Not CBool(nCodIni) And CBool(nCodFim)) Then
    Concat strDupls, " AND D.Cheque <= ", CStr(nCodFim)
    Concat strLanctos, " AND L.Cheque <= ", CStr(nCodFim)
    Concat strTransf, " AND T.Cheque <= ", CStr(nCodFim)
  Else
    Concat strDupls, " AND D.Cheque > 0"          '// Evita que sejam recuparados
    Concat strLanctos, " AND L.Cheque > 0"        '// registros que não possuam
    Concat strTransf, " AND T.Cheque > 0"         '// cheque
  End If

  '// Verificando se o usuário filtrou por datas

  dInicial = CDateDef(txtCheque(3).Text)
  dFinal = CDateDef(txtCheque(4).Text)
  
  If IsValid(txtCheque(3).Text) And IsValid(txtCheque(4).Text) Then
    If EData(dInicial) And EData(dFinal) Then
      If dFinal < dInicial Then
        MsgFunc "Data Final menor que Data Inicial"
        Exit Sub
      End If
    End If
  End If
  
  If gTipoDB = Access Then

    If (Not IsEmptyDate(dInicial) And Not IsEmptyDate(dFinal)) Then
      If (DateDiff(DD_DIA, dInicial, dFinal) = ZERO) Then
        Concat strDupls, wsprintf(" AND D.Pagamento = #%q#", dInicial)
        Concat strLanctos, wsprintf(" AND L.Pagamento = #%q#", dInicial)
        Concat strTransf, wsprintf(" AND T.Data = #%q#", dInicial)
      Else
        Concat strDupls, wsprintf(" AND (D.Pagamento BETWEEN #%q# AND #%q#)", dInicial, dFinal)
        Concat strLanctos, wsprintf(" AND (L.Pagamento BETWEEN #%q# AND #%q#)", dInicial, dFinal)
        Concat strTransf, wsprintf(" AND (T.Data BETWEEN #%q# AND #%q#)", dInicial, dFinal)
      End If
    ElseIf (Not IsEmptyDate(dInicial) And IsEmptyDate(dFinal)) Then
      Concat strDupls, wsprintf(" AND D.Pagamento >= #%q#", dInicial)
      Concat strLanctos, wsprintf(" AND L.Pagamento >= #%q#", dInicial)
      Concat strTransf, wsprintf(" AND T.Data >= #%q#", dInicial)
    ElseIf (IsEmptyDate(dInicial) And Not IsEmptyDate(dFinal)) Then
      Concat strDupls, wsprintf(" AND D.Pagamento <= #%q#", dFinal)
      Concat strLanctos, wsprintf(" AND L.Pagamento <= #%q#", dFinal)
      Concat strTransf, wsprintf(" AND T.Data <= #%q#", dFinal)
    End If
  
  Else
  
    If (Not IsEmptyDate(dInicial) And Not IsEmptyDate(dFinal)) Then
      If (DateDiff(DD_DIA, dInicial, dFinal) = ZERO) Then
        Concat strDupls, wsprintf(" AND D.Pagamento = '%q'", dInicial)
        Concat strLanctos, wsprintf(" AND L.Pagamento = '%q'", dInicial)
        Concat strTransf, wsprintf(" AND T.Data = '%q'", dInicial)
      Else
        Concat strDupls, wsprintf(" AND (D.Pagamento BETWEEN '%q' AND '%q')", dInicial, dFinal)
        Concat strLanctos, wsprintf(" AND (L.Pagamento BETWEEN '%q' AND '%q')", dInicial, dFinal)
        Concat strTransf, wsprintf(" AND (T.Data BETWEEN '%q' AND '%q')", dInicial, dFinal)
      End If
    ElseIf (Not IsEmptyDate(dInicial) And IsEmptyDate(dFinal)) Then
      Concat strDupls, wsprintf(" AND D.Pagamento >= '%q'", dInicial)
      Concat strLanctos, wsprintf(" AND L.Pagamento >= '%q'", dInicial)
      Concat strTransf, wsprintf(" AND T.Data >= '%q'", dInicial)
    ElseIf (IsEmptyDate(dInicial) And Not IsEmptyDate(dFinal)) Then
      Concat strDupls, wsprintf(" AND D.Pagamento <= '%q'", dFinal)
      Concat strLanctos, wsprintf(" AND L.Pagamento <= '%q'", dFinal)
      Concat strTransf, wsprintf(" AND T.Data <= '%q'", dFinal)
    End If
  
  End If
  '// Agrupando os dados em Banco e Cheque

  Concat strDupls, " GROUP BY D.Banco, D.Cheque, D.Pagamento, D.Enviada, D.Descrição "
  Concat strLanctos, " GROUP BY L.Banco, L.Cheque, L.Pagamento, L.Enviado, L.Descrição "
  Concat strTransf, " GROUP BY T.Origem, T.Cheque, T.Data, T.Enviada, T.Descrição "

  '// Finaliza a instrução unindo as "SELECT's" e acrescentado a
  '// ordem dos dados.
  Dim sOrderBy As String
  If (tabCheque.SelectedItem.Key = "relatorio") Then
    'Protocolo Nr 102985 - Carlos Felippe Vernizze - 14/12/2010
    #If FOXSQL = 0 Then
    sOrderBy = IIf((chkCheque(1).value = vbChecked), "D.Banco, " & cboCheque.Text, cboCheque.Text)
    #Else
    sOrderBy = IIf((chkCheque(1).value = vbChecked), "Banco, " & cboCheque.Text, cboCheque.Text)
    #End If
    strDupls = wsprintf("%s UNION ALL %s UNION ALL %s", strDupls, strLanctos, strTransf)
  Else
    #If FOXSQL = 0 Then
    sOrderBy = "D.Banco, D.Cheque"
    #Else
    sOrderBy = "Banco, Cheque"
    #End If
    strDupls = wsprintf("%s UNION ALL %s UNION ALL %s", strDupls, strLanctos, strTransf)
  End If

  ExecuteSQL "DROP TABLE Temp", False
  ExecuteSQL "DROP VIEW Temp", False
  #If FOXSQL = 0 Then
    strDupls = strDupls & " ORDER BY " & sOrderBy
    If ConsultaExiste("Temp") Then DeleteQuery Nothing, "Temp"
    '// Cria uma Consulta temporária para a seleção dos dados
    If (CreateQuery(qdfTemp, "Temp", strDupls) = WL_OK) Then
  #Else
    Dim s As String
    If ExecuteSQL("CREATE TABLE Temp (Banco INT, Cheque INT, Data Date, Valor MONEY, Impresso CHAR(1), [Descrição] VarChar(80))") = 0 Then
        MsgBox "Erro criando tabela temporária"
        Exit Sub
    End If
    s = "INSERT INTO Temp SELECT * FROM (" & strDupls & ") AS TMP ORDER BY " & sOrderBy
    If ExecuteSQL(s) Then
  #End If

    '// Cria uma segunda instrução "SELECT" para somar os valores dos
    '// cheques agrupados na consulta

    If (tabCheque.SelectedItem.Key = "relatorio") Then
      strDupls = "SELECT Banco, Cheque, Data, SUM(Valor) As Total FROM Temp " & _
                 "GROUP BY Banco, Cheque, Data ORDER BY " & cboCheque.Text & ";"
    ElseIf (tabCheque.SelectedItem.Key = "impressao") Then
        If gTipoDB = MsSql Then
            strDupls = "SELECT Banco, Cheque, Data, SUM(Valor) As Total, MIN(Impresso) as Imp FROM Temp GROUP BY Banco, Cheque, Data ORDER BY Banco, Cheque;"
        Else
            strDupls = "SELECT Banco, Cheque, Data, SUM(Valor) As Total, First(Impresso) as Imp FROM Temp GROUP BY Banco, Cheque, Data ORDER BY Banco, Cheque;"
        End If
    Else
      strDupls = "SELECT Banco, Cheque, Data, SUM(Valor) As Total, Impresso, Descrição FROM Temp " & _
                 "GROUP BY Banco, Cheque, Data, Impresso, Descrição ORDER BY Banco, Cheque;"
    End If
    'Pt. 95368 - Moacir Pfau(12/11/2009)
    If (AbreRecordset(rstDados, strDupls, dbOpenSnapshot) = WL_OK) Then
      If (tabCheque.SelectedItem.Key = "relatorio") Then

        If (TempRelatorio(rstAux)) Then                             '// Cria a tabela temporária
          If (AddRegRelatorio(rstDados, rstAux)) Then               '// Adiciona os dados
            Call RelatorioCheques(rstAux, pdeDestino)
          End If
        End If
        Call DeleteAux(rstAux, NUL)

      ElseIf (tabCheque.SelectedItem.Key = "copia") Then

        '// Esta função só irá imprimir o relatório de Cópia de Cheques
        '// quando o usuário NÃO solicitar impressão dos lançamentos
        '// correspondentes

        If (TempCopia(rstAux)) Then                                 '// Cria a tabela temporária
          If (AddRegCopia(rstDados, rstAux)) Then
            Call RelatorioCopia(rstAux, rstAux, pdeDestino)
          End If
        End If
        Call DeleteAux(rstAux, NUL)

      ElseIf (tabCheque.SelectedItem.Key = "impressao") Then

        '// Esta parte é utilizada para impressão de cheque em formulário
        '// contínuo ou avulso.
        If (TempImpressao(rstAux)) Then
          If (AddRegImpressao(rstAux, rstDados)) Then
          

            '**************************************
            'Referente ao protocolo: 71664
            'Devido aos problemas de impressão
            'o relatório foi refeito com o ReportX
            fimpCheque.Config rstAux, CLngDef(txtCheque(0).Text) ', (pdeDestino <> wrToPrinter)
            '*********************************

            'Call ChequeImpressao(rstAux, pdeDestino)
          End If
        End If
        Call DeleteAux(rstAux, NUL)

      End If
    ElseIf (UltimoRetorno() = WL_NORECORD) Then
      MsgFunc LoadResString(146)
    End If
    Call FechaRecordset(rstDados)

  End If
  #If FOXSQL = 1 Then
  ExecuteSQL "DROP TABLE Temp"
  #Else
  Call DeleteQuery(qdfTemp)
  #End If

  MsgBar MsgBoxCaption

End Sub

' FUNCTION..: AddRegRelatorio
' Objetivo..: Adiciona os dados à tabela temporária para o Relatório de Cheques.
' Argumentos: [rsData]: Recordset com os dados de Duplicatas, Lançamentos e Transferências.
'             [rsTemp]: Recordset com a tabela temporária.
' Retorna...: True se obtiver sucesso, False se não.
' ------------------------------------------------------------------------------------
Private Function AddRegRelatorio(rsData As Object, rstEmp As Object) As Boolean
Dim lBanco  As Long                     '// Código do Banco
Dim lCheque As Long                     '// Número do Cheque
Dim sBanco  As String                   '// Nome do Banco no cadastro de Bancos

  Call InKey(vbKeyEscape)               '// Limpa o buffer do teclado
  On Error GoTo AddRegRelatorio_Erro

  Call InitTrans
  Do
    lBanco = GetValue(rsData, "Banco", ZERO)
    lCheque = GetValue(rsData, "Cheque", ZERO)
    sBanco = GetFieldValue("Nome", "Bancos", "Banco = " & CStr(lBanco), , NUL)

    If (InKey(vbKeyEscape)) Then            '// Habilita ao usuário cancelar a operação
      GoTo AddRegRelatorio_Erro
    End If

    SimpleMsgBar wsprintf("Pesquisando Cheque %l do Banco %l %s", lCheque, lBanco, sBanco)

    rstEmp.AddNew
    rstEmp("Banco").value = lBanco
    rstEmp("Nome").value = sBanco
    rstEmp("Cheque").value = lCheque
    rstEmp("Data").value = GetValue(rsData, "Data", Null)
    rstEmp("Valor").value = GetValue(rsData, "Total", ZERO)
    rstEmp("Nominal").value = GetFieldValue("Nominal", "Cheque", _
                                            "Banco = " & CStr(lBanco) & _
                                            " AND Cheque = " & CStr(lCheque), , NUL)
    rstEmp.update
    rsData.MoveNext               '// Move para o próximo registro de lançamentosg
  Loop Until (rsData.EOF)
  Call UpdateTrans(FORCE_WRITE)
  AddRegRelatorio = True
  Exit Function

AddRegRelatorio_Erro:
  If (err().Number) Then
    #If (DebugInfo) Then
      MsgErro wsprintf("Erro: %l\n%s\nAddRegRelatorio", err.Number, err.Description)
    #Else
      DAOErros NUL
    #End If
  End If
  Call CancelTrans
  AddRegRelatorio = False
End Function

' FUNCTION..: AddRegCopia
' Objetivo..: Grava a tabela auxiliar responsável pela parte correspondente
'             ao cheque no relatório de Cópia de Cheques.
' Argumentos: [rsLanctos]: Recordset com os lançamentos.
'             [rsAux    ]: Recordset com a tabela auxiliar.
' Retorna...: True se obtiver sucesso, False se não.
' ------------------------------------------------------------------------------------
Private Function AddRegCopia(rsLanctos As Object, rsAux As Object) As Boolean
Dim lBco As Long                  '// Código do Banco
Dim lChq As Long                  '// Código do Cheque
Dim sBco As String                '// Nome do Banco
Dim rsCh As Object             '// Recordset para os dados do cadastro de Cheques
Dim sChq As String                '// Instrução de seleção de dados para Cheques
Dim ValorTotal    As Double
Dim fakedao As New CGenericRecordset

If Not rsAux Is Nothing Then
    fakedao.Initialize rsAux
End If

  Call InKey(vbKeyEscape)         '// Limpa o buffer do teclado

  On Error GoTo AddRegCopia_Erro
  Call InitTrans
  
  rsLanctos.MoveFirst
  Do
    If (InKey(vbKeyEscape)) Then      '// Se o usuário pressionar ESC
      GoTo AddRegCopia_Erro           '// Cancela a geração
    End If

    lBco = GetValue(rsLanctos, "Banco", ZERO)
    lChq = GetValue(rsLanctos, "Cheque", ZERO)
    sBco = GetFieldValue("Nome", "Bancos", "Banco = " & CStr(lBco), , NUL)

    SimpleMsgBar wsprintf("Pesquisando Cheque %l do Banco %l %s", lChq, lBco, sBco)
    
    sChq = wsprintf("SELECT * FROM Cheque WHERE Banco = %l AND Cheque = %l;", lBco, lChq)
    If (AbreRecordset(rsCh, sChq, dbOpenSnapshot) = WL_OK) Then
      ValorTotal = Soma("Valor", "Temp", "Banco = " & lBco & " and Cheque = " & lChq, ZERO)
      If Not EstaVazio(rsAux) Then
        rsAux.MoveFirst
        fakedao.FindFirst "Banco = " & lBco & " AND Cheque = " & lChq
        If fakedao.NoMatch Then
          rsAux.AddNew
        Else
          fakedao.Edit
        End If
      Else
        rsAux.AddNew
      End If
      rsAux("Banco").value = lBco
      rsAux("Cheque").value = lChq
      rsAux("Valor").value = ValorTotal
      rsAux("Data").value = GetValue(rsLanctos, "Data", Null)
      rsAux("Nome").value = sBco
      rsAux("Nominal").value = GetValue(rsCh, "Nominal", NUL)
      If IsValid(GetValue(rsAux, "Desc", NUL)) Then
        rsAux("Desc").value = GetValue(rsAux, "Desc", NUL) & ", " & GetValue(rsLanctos, "Descrição", NUL)
      Else
        rsAux("Desc").value = GetValue(rsLanctos, "Descrição", NUL)
      End If

      '// Resolvendo o extenso do valor do cheque

      rsAux("Extenso").value = KeybUCase(KeybExtenso(GetValue(rsLanctos, "Total", ZERO)), PorPalavra)

      'Valores Totais
      rsAux("Valor Total").value = ValorTotal
      rsAux("Extenso Total").value = KeybUCase(KeybExtenso(ValorTotal), PorPalavra)
      '// Resolvendo o extenso da data do cheque

      rsAux("DtExt").value = DataLongaExt(GetValue(rsLanctos, "Data", Empty))
      rsAux.update
    End If
    Call FechaRecordset(rsCh)
    rsLanctos.MoveNext                '// Move para o próximo cheque
  Loop Until (rsLanctos.EOF)
  Call UpdateTrans(FORCE_WRITE)
  AddRegCopia = True
  Set fakedao = Nothing
  Exit Function

AddRegCopia_Erro:
  If (err().Number) Then
    #If (DebugInfo) Then
      MsgErro wsprintf("Erro: %l\n%s\nAddRegCopia", err.Number, err.Description)
    #Else
      DAOErros NUL
    #End If
  End If
  Call CancelTrans
  Call FechaRecordset(rsCh)
  Resume
End Function

' FUNCTION..: FiltroCopiaCheques
' Objetivo..: Cria a instrução que filtra os dados dos lançamentos
'             selecionados pelo usuário para o relatório de Cópia de
'             Cheques quando o usuário seleciona que os lançamentos devem
'             aparecer no relatório.
' Argumento.: [pdeDest]: Destino da impressão.
' ------------------------------------------------------------------------------------
Private Sub FiltroCopiaCheques(pdeDest As PrintDestinoEnum)
Dim sDupl As String                       '// String de seleção de Duplicatas
Dim sLanc As String                       '// String de seleção de Lançamentos
Dim sTran As String                       '// String de seleção de Transf. Bancária
Dim niCod As Long                         '// Código Inicial
Dim nfCod As Long                         '// Código Final
Dim diDat As Date                         '// Data Inicial
Dim dfDat As Date                         '// Data Final
Dim rsLan As Object                    '// Recordset com os lançamentos selecionados


  ' Ah se eu pudesse matar o kra que fez isso com FORMAT...
  ' Devido o SQL não reconhecer o FORMAT dentro da Select terei sempre 2 versões da mesma select
  ' fique atento com alterações na mesma, sempre deve ser feito nas duas versões.

  SimpleMsgBar "Filtrando dados, aguarde..."
  
  
  
  If gTipoDB = Access Then

    sDupl = wsprintf("SELECT FORMAT(D.Nota, \'000000\') & '-' & D.Tipo AS Cod, D.Empresa, " & _
                   "D.Descrição, D.Conta, D.Centro, D.Controle, (D.[Valor Original] + D.Acréscimo " & _
                   "- D.Abatimento) AS Valor, D.Banco, D.Cheque, D.Pagamento AS Data " & _
                   "FROM Duplicatas AS D")

    sLanc = wsprintf("SELECT FORMAT(L.Código, \'000000\') & '-' & L.Tipo AS Cod, L.Empresa, " & _
                   "L.Descrição, L.Conta, L.Centro, L.Controle, (L.[Valor Original] + L.Acréscimo " & _
                   "- L.Abatimento) AS Valor, L.Banco, L.Cheque, L.Pagamento AS Data " & _
                   "FROM Lançamentos AS L")

    sTran = wsprintf("SELECT FORMAT(T.Código, \'000000\') & '-Tranferência' AS Cod, T.Destino, " & _
                   "T.Descrição, T.Conta, T.Centro, ' ' AS Controle, T.Valor, T.Origem, T.Cheque, T.Data " & _
                   "FROM [Transf Bancária] AS T")
  Else
  
    sDupl = wsprintf("SELECT replicate('0',15-len(cast(cast(D.Nota as bigint) as varchar(15))))+ cast(cast(D.Nota as bigint) as varchar(15)) + '-' + convert(varchar,D.Tipo) AS Cod, D.Empresa, " & _
                   "D.Descrição, D.Conta, D.Centro, D.Controle, (D.[Valor Original] + D.Acréscimo " & _
                   "- D.Abatimento) AS Valor, D.Banco, D.Cheque, D.Pagamento AS Data " & _
                   "FROM Duplicatas AS D")
                   
    sLanc = wsprintf("SELECT replicate('0',15-len(cast(cast(L.Código as bigint) as varchar(15))))+ cast(cast(L.Código as bigint) as varchar(15)) + '-' + convert(varchar,L.Tipo) AS Cod, L.Empresa, " & _
                   "L.Descrição, L.Conta, L.Centro, L.Controle, (L.[Valor Original] + L.Acréscimo " & _
                   "- L.Abatimento) AS Valor, L.Banco, L.Cheque, L.Pagamento AS Data " & _
                   "FROM Lançamentos AS L")
                   
    sTran = wsprintf("SELECT replicate('0',15-len(cast(cast(T.Código as bigint) as varchar(15))))+ cast(cast(T.Código as bigint) as varchar(15)) + '-Tranferência' AS Cod, convert(varchar,T.Destino), " & _
                   "T.Descrição, T.Conta, T.Centro, ' ' as Controle, T.Valor, T.Origem, T.Cheque, T.Data " & _
                   "FROM [Transf Bancária] AS T")
    
  End If

  '// Verificando se o usuário filtrou por banco

  niCod = CLngDef(txtCheque(0).Text)
  If (niCod) Then
    Concat sDupl, " WHERE D.Banco = ", CStr(niCod)
    Concat sLanc, " WHERE L.Banco = ", CStr(niCod)
    Concat sTran, " WHERE T.Origem = ", CStr(niCod)
  Else
    Concat sDupl, " WHERE D.Banco > 0"
    Concat sLanc, " WHERE L.Banco > 0"
    Concat sTran, " WHERE T.Origem > 0"
  End If

  '// Verificando se o usuário filtrou por cheque

  niCod = CLngDef(txtCheque(1).Text)
  nfCod = CLngDef(txtCheque(2).Text)

  If (CBool(niCod) And CBool(nfCod)) Then
    If (niCod = nfCod) Then
      Concat sDupl, " AND D.Cheque = ", CStr(niCod)
      Concat sLanc, " AND L.Cheque = ", CStr(niCod)
      Concat sTran, " AND T.Cheque = ", CStr(niCod)
    Else
      Concat sDupl, wsprintf(" AND (D.Cheque BETWEEN %l AND %l)", niCod, nfCod)
      Concat sLanc, wsprintf(" AND (L.Cheque BETWEEN %l AND %l)", niCod, nfCod)
      Concat sTran, wsprintf(" AND (T.Cheque BETWEEN %l AND %l)", niCod, nfCod)
    End If
  ElseIf (CBool(niCod) And Not CBool(nfCod)) Then
    Concat sDupl, " AND D.Cheque >= ", CStr(niCod)
    Concat sLanc, " AND L.Cheque >= ", CStr(niCod)
    Concat sTran, " AND T.Cheque >= ", CStr(niCod)
  ElseIf (Not CBool(niCod) And CBool(nfCod)) Then
    Concat sDupl, " AND D.Cheque <= ", CStr(nfCod)
    Concat sLanc, " AND L.Cheque <= ", CStr(nfCod)
    Concat sTran, " AND T.Cheque <= ", CStr(nfCod)
  End If

  '// Verificando se o usuário filtrou por data

  diDat = CDateDef(txtCheque(3).Text, Empty)
  dfDat = CDateDef(txtCheque(4).Text, Empty)

  If IsValid(txtCheque(3).Text) And IsValid(txtCheque(4).Text) Then
    If EData(diDat) And EData(dfDat) Then
      If dfDat < diDat Then
        MsgFunc "Data Final menor que Data Inicial"
        Exit Sub
      End If
    End If
  End If
  
  If (Not IsEmptyDate(diDat) And Not IsEmptyDate(dfDat)) Then
    If (DateDiff(DD_DIA, diDat, dfDat) = ZERO) Then
        If gTipoDB = Access Then
            Concat sDupl, wsprintf(" AND D.Pagamento = #%q#", diDat)
            Concat sLanc, wsprintf(" AND L.Pagamento = #%q#", diDat)
            Concat sTran, wsprintf(" AND T.Data = #%q#", diDat)
        Else
            Concat sDupl, wsprintf(" AND D.Pagamento = '%q'", diDat)
            Concat sLanc, wsprintf(" AND L.Pagamento = '%q'", diDat)
            Concat sTran, wsprintf(" AND T.Data = '%q'", diDat)
        End If
    Else
        If gTipoDB = Access Then
            Concat sDupl, wsprintf(" AND (D.Pagamento BETWEEN #%q# AND #%q#)", diDat, dfDat)
            Concat sLanc, wsprintf(" AND (L.Pagamento BETWEEN #%q# AND #%q#)", diDat, dfDat)
            Concat sTran, wsprintf(" AND (T.Data BETWEEN #%q# AND #%q#)", diDat, dfDat)
        Else
            Concat sDupl, wsprintf(" AND (D.Pagamento BETWEEN '%q' AND '%q')", diDat, dfDat)
            Concat sLanc, wsprintf(" AND (L.Pagamento BETWEEN '%q' AND '%q')", diDat, dfDat)
            Concat sTran, wsprintf(" AND (T.Data BETWEEN '%q' AND '%q')", diDat, dfDat)
        End If
    End If
  ElseIf (Not IsEmptyDate(diDat) And IsEmptyDate(dfDat)) Then
    If gTipoDB = Access Then
        Concat sDupl, wsprintf(" AND D.Pagamento >= #%q#", diDat)
        Concat sLanc, wsprintf(" AND L.Pagamento >= #%q#", diDat)
        Concat sTran, wsprintf(" AND T.Data >= #%q#", diDat)
    Else
        Concat sDupl, wsprintf(" AND D.Pagamento >= '%q'", diDat)
        Concat sLanc, wsprintf(" AND L.Pagamento >= '%q'", diDat)
        Concat sTran, wsprintf(" AND T.Data >= '%q'", diDat)
    End If
  ElseIf (IsEmptyDate(diDat) And Not IsEmptyDate(dfDat)) Then
    If gTipoDB = Access Then
        Concat sDupl, wsprintf(" AND D.Pagamento <= #%q#", dfDat)
        Concat sLanc, wsprintf(" AND L.Pagemento <= #%q#", dfDat)
        Concat sTran, wsprintf(" AND T.Data <= #%q#", dfDat)
    Else
        Concat sDupl, wsprintf(" AND D.Pagamento <= '%q'", dfDat)
        Concat sLanc, wsprintf(" AND L.Pagemento <= '%q'", dfDat)
        Concat sTran, wsprintf(" AND T.Data <= '%q'", dfDat)
    End If
  End If

  '// Concatenando as instruções para retornarem um único recordset

  sDupl = wsprintf("%s UNION ALL %s UNION ALL %s ORDER BY Banco, Cheque;", sDupl, sLanc, sTran)

  '// Chama a função que abre as consultas e verifica se existem cheques

  If (AbreRecordset(rsLan, sDupl, dbOpenSnapshot, , , adUseClient) = WL_OK) Then
    Call AddRegCopiaLanctos(rsLan, pdeDest)
  ElseIf (UltimoRetorno() = WL_NORECORD) Then
    MsgFunc LoadResString(IDS_RECORDNOTFOUND)
  End If
  FechaRecordset rsLan

End Sub

' SUB.......: AddRegCopiaLanctos
' Objetivo..: Abre as consultas criadas para selecionar os lançamentos do
'             do usuário, cria e grava as tabelas auxiliares e exibe o relatório.
' Argumentos: [rstLanc]: Recordset com os dados para impressão
'             [nDest  ]: Destino da impressão.
' ------------------------------------------------------------------------------------
Private Sub AddRegCopiaLanctos(rstLanc As Object, nDest As Long)
Dim rsChAx As Object             '// Tabela auxiliar com os dados dos cheques
Dim rsAux  As Object             '// Tabela auxiliar com os dados dos lançamentos
Dim cValor As Currency              '// Valor total do cheque
Dim dtChq  As Date                  '// Data da emissão do cheque
Dim nBco   As Long                  '// Código do Banco
Dim nChq   As Long                  '// Número do Cheque
Dim sBco   As String                '// Nome do Banco

  On Error GoTo AddRegCopiaLanctos_Erro

  Call InKey(vbKeyEscape)           '// Limpa o buffer do teclado

  If (TempCopia(rsChAx) And TempCopiaLan(rsAux)) Then
    
    Call InitTrans
    While (Not rstLanc.EOF)
      nBco = GetValue(rstLanc, "Banco", ZERO)
      nChq = GetValue(rstLanc, "Cheque", ZERO)
      sBco = GetFieldValue("Nome", "Bancos", "Banco = " & CStr(nBco), , NUL)
      cValor = ZERO
      dtChq = GetValue(rstLanc, "Data", Empty)

      SimpleMsgBar wsprintf("Pesquisando Cheque %l do Banco %l %s", nChq, nBco, sBco)

      Do While ((nBco = GetValue(rstLanc, "Banco", 0)) And (nChq = GetValue(rstLanc, "Cheque", 0)))
        DoEvents
        If (InKey(vbKeyEscape)) Then GoTo AddRegCopiaLanctos_Erro

        Call GravaAuxLanc(rsAux, rstLanc)

        cValor = cValor + GetValue(rstLanc, "Valor", ZERO)
        rstLanc.MoveNext
        If (rstLanc.EOF) Then Exit Do
      Loop
      
      Dim rstContasL  As Object    ' SELECT DISTINC das contas em Lançamentos
      Dim rstContasD  As Object    ' SELECT DISTINC das contas em Duplicatas
      Dim fdsCts(0)   As FieldStruct  ' Campo Conta da tabela auxiliar
      Dim rstCts      As Object    ' Tabela auxliar: conterá contas diferentes de cada cheque(Lançamentos e Duplicatas)
      
      AppendVar fdsCts(0), "Conta", dbLong
      CrieAux rstCts, fdsCts
      
      '
      ' Para mostrar o total das contas, primeiro
      If (AbreRecordset(rstContasL, "SELECT DISTINCT Conta FROM Lançamentos " & _
              "WHERE Banco = " & CStr(nBco) & " AND Cheque = " & CStr(nChq), dbOpenSnapshot, , , adUseClient) = WL_OK) Or _
        (AbreRecordset(rstContasD, "SELECT DISTINCT Conta FROM Duplicatas " & _
            "WHERE Banco = " & CStr(nBco) & " AND Cheque = " & CStr(nChq), dbOpenSnapshot, , , adUseClient) = WL_OK) Then
        '
        ' Gravando contas distintas de Lançamentos
        While Not rstContasL.EOF
          
          rstCts.AddNew
          rstCts("Conta").value = GetValue(rstContasL, "Conta")
          rstCts.update
          
          rstContasL.MoveNext
          
        Wend
        
        '
        ' Gravando contas distintas de Duplicatas
        While Not rstContasD.EOF
          If TypeOf rstCts Is dao.Recordset Then
            If Recordcount("SELECT Conta FROM " & rstCts.name & " WHERE Conta = " & CStr(GetValue(rstContasD, "Conta"))) = 0 Then
              rstCts.AddNew
              rstCts("Conta").value = GetValue(rstContasD, "Conta")
              rstCts.update
            End If
          Else
            If Recordcount(rstCts.Source & " WHERE Conta = " & CStr(GetValue(rstContasD, "Conta"))) = 0 Then
              rstCts.AddNew
              rstCts("Conta").value = GetValue(rstContasD, "Conta")
              rstCts.update
            End If
          End If
          
          rstContasD.MoveNext
        Wend
        
        rstCts.MoveFirst
        
        '
        ' Pula uma linha
        '
        rsAux.AddNew
        rsAux("Banco").value = nBco
        rsAux("Cheque").value = nChq
        rsAux("Data").value = Null
        rsAux("Valor").value = ZERO
        rsAux("Lancto").value = NUL
        rsAux("Emp").value = NUL
        rsAux("Desc").value = NUL
        rsAux("Conta").value = ZERO
        rsAux("CtDesc").value = NUL
        rsAux("Custo").value = ZERO
        rsAux("CsDesc").value = NUL
        rsAux("Controle").value = NUL

        rsAux.update
        
        '
        ' Agora grava os registros como se fossem Lançamentos/Duplicatas
        ' Para aparecer o total da conta ao final
        '
        Do
        
          rsAux.AddNew
          rsAux("Banco").value = nBco
          rsAux("Cheque").value = nChq
          rsAux("Data").value = Null
          rsAux("Valor").value = (Soma("[Valor Original]", "Lançamentos", "Conta = " & GetValue(rstCts, "Conta") & " AND Banco = " & CStr(nBco) & " AND Cheque = " & CStr(nChq))) + (Soma("[Valor Original]", "Duplicatas", "Conta = " & GetValue(rstCts, "Conta") & " AND Banco = " & CStr(nBco) & " AND Cheque = " & CStr(nChq)))
          rsAux("Lancto").value = NUL
          rsAux("Emp").value = NUL
          rsAux("Desc").value = "TOTAL DA CONTA"
          rsAux("Conta").value = GetValue(rstCts, "Conta", ZERO)
          rsAux("CtDesc").value = GetFieldValue("Descrição", "Contas", "Código = " & _
                                                 GetValue(rstCts, "Conta"))
          rsAux("Custo").value = ZERO
          rsAux("CsDesc").value = NUL
          rsAux("Controle").value = NUL
          rsAux.update
          
          rstCts.MoveNext
          
        Loop Until (rstCts.EOF)
      End If
      FechaRecordset rstContasL
      FechaRecordset rstContasD
      DeleteAux rstCts, NUL
      
      rsChAx.AddNew
      rsChAx("Banco").value = nBco
      rsChAx("Cheque").value = nChq
      rsChAx("Valor").value = cValor
      rsChAx("Data").value = dtChq
      rsChAx("Nome").value = sBco
      rsChAx("Nominal").value = GetFieldValue("Nominal", "Cheque", _
                                              wsprintf("Banco = %l AND Cheque = %l", _
                                                       nBco, nChq), , NUL)
      rsChAx("Extenso").value = KeybUCase(KeybExtenso(cValor), PorPalavra)
      rsChAx("DtExt").value = DataLongaExt(dtChq)
      rsChAx("Desc").value = GetFieldValue("Histórico", "Cheque", _
                                              wsprintf("Banco = %l AND Cheque = %l", _
                                                       nBco, nChq), , NUL)
      rsChAx.update
    Wend
    Call UpdateTrans(FORCE_WRITE)
    Call RelatorioCopia(rsAux, rsChAx, nDest)
  End If
  Call DeleteAux(rsChAx, NUL)
  Call DeleteAux(rsAux, NUL)
  Exit Sub

AddRegCopiaLanctos_Erro:
  If err().Number <> 0 Then
    #If (DebugInfo) Then
      MsgErro wsprintf("Erro: %l\n%s\nAddRegCopiaLanctos", err.Number, err.Description)
    #Else
      DAOErros NUL
    #End If
  End If
  Call CancelTrans
  Call DeleteAux(rsChAx, NUL)
  Call DeleteAux(rsAux, NUL)
End Sub

' SUB.......: GravaAuxLanc
' Objetivo..: Grava a tabela auxiliar com dados de lançamentos para o
'             relatório de Cópia de Cheque quando o usuário necessita
'             que sejam exibidos os lançamentos no relatório.
' Argumentos: [rstAux]: Recordset da tabela auxiliar.
'             [rstSrc]: Recordset com os dados dos lançamentos.
' ------------------------------------------------------------------------------------
Private Sub GravaAuxLanc(rstAux As Object, rstSrc As Object)
  rstAux.AddNew
  rstAux("Banco").value = GetValue(rstSrc, "Banco", ZERO)
  rstAux("Nome").value = GetFieldValue("Nome", "Bancos", "Banco = " & GetValue(rstSrc, "Banco", 0), , ZERO)
  rstAux("Cheque").value = GetValue(rstSrc, "Cheque", ZERO)
  rstAux("Data").value = GetValue(rstSrc, "Data", Null)
  rstAux("Valor").value = GetValue(rstSrc, "Valor", ZERO)
  rstAux("Lancto").value = GetValue(rstSrc, "Cod", NUL)
  rstAux("Emp").value = GetValue(rstSrc, "Empresa", NUL)
  rstAux("Desc").value = GetValue(rstSrc, "Descrição", NUL)
  rstAux("Conta").value = GetValue(rstSrc, "Conta", ZERO)
  rstAux("CtDesc").value = GetFieldValue("Descrição", "Contas", "Código = " & _
                                         GetValue(rstSrc, "Conta", 0), , NUL)
  rstAux("Custo").value = GetValue(rstSrc, "Centro", ZERO)
  rstAux("CsDesc").value = GetFieldValue("Descrição", "Centros", "Código = " & _
                                         GetValue(rstSrc, "Centro", 0), , NUL)
  rstAux("Controle").value = GetValue(rstSrc, "Controle", ZERO)
  
  rstAux.update
End Sub

' FUNCTION..: AddRegImpressao
' Objetivo..: Grava os dados para impressão do cheque na tabela auxiliar.
' Argumento.: [rstAux]: Recordset da tabela auxiliar.
'             [rstLan]: Recordset com os lançamentos.
' Retorna...: True se obtiver sucesso, False se não.
' ------------------------------------------------------------------------------------
Private Function AddRegImpressao(rstAux As Object, rstLan As Object) As Boolean
Dim rstModel As Object           '// Recordset com o modelo de impressão
Dim nBanco   As Long                '// Número do Banco
Dim sBanco   As String              '// Nome do Banco
Dim nCamara  As Long                '// Código de Compensação
Dim strExt1  As String              '// Primeira linha de extenso
Dim strExt2  As String              '// Segunda  linha de extenso

  On Error GoTo AddRegImpressao_Erro

  Call InKey(vbKeyEscape)           '// Limpa o buffer do teclado

  nBanco = CLngDef(txtCheque(0).Text)
  nCamara = GetFieldValue("Câmara", "Bancos", "Banco = " & CStr(nBanco), , ZERO)
  sBanco = GetFieldValue("Nome", "Bancos", "Banco = " & CStr(nBanco), , NUL)

  If (AbreRecordset(rstModel, wsprintf("SELECT * FROM ChqModelos WHERE Número = %l", nCamara), _
                    dbOpenSnapshot) = WL_OK) Then
    Call InitTrans
    Do Until (rstLan.EOF)
      DoEvents
      If (InKey(vbKeyEscape)) Then GoTo AddRegImpressao_Erro

      SimpleMsgBar wsprintf("Configurando Cheque %l do Banco %l %s", _
                            GetValue(rstLan, "Cheque", 0), nBanco, sBanco)
      
      Dim Imprime   As Boolean    ' Indica se o registro será ou não impresso"
      Imprime = True

      If GetValue(rstLan, "Imp", NUL) <> NUL Then
        Imprime = (MsgBox("O cheque " & CStr(GetValue(rstLan, "Cheque", 0)) & " já foi impresso." & vbCrLf & vbCrLf & "Deseja imprimí-lo novamente?", _
                        vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes)
      End If

      If Imprime Then
      
        rstAux.AddNew
        rstAux("BcoChq").value = wsprintf("%03l - %06l", _
                                           GetFieldValue("Câmara", _
                                                         "Bancos", _
                                                         "Banco = " & CStr(nBanco), , 0), _
                                           GetValue(rstLan, "Cheque", 0))
        rstAux("Local").value = wsprintf("%s, %d", CidadePadrao(), GetValue(rstLan, "Data", Empty))
        rstAux("Nominal").value = GetFieldValue("Nominal", "Cheque", _
                                                wsprintf("Banco = %l AND Cheque = %l", _
                                                         nBanco, _
                                                         GetValue(rstLan, "Cheque", 0)), , NUL)
        
        If (GetValue(rstModel, "MesCompleto", False)) Then
          rstAux("Mês").value = wsprintf("%M", GetValue(rstLan, "Data", Empty))
        Else
          rstAux("Mês").value = wsprintf("%.3M", GetValue(rstLan, "Data", Empty))
        End If
        
        If (GetValue(rstModel, "AnoCompleto", False)) Then
          rstAux("Ano").value = Format$(GetValue(rstLan, "Data", Empty), "yyyy")
        Else
          rstAux("Ano").value = Format$(GetValue(rstLan, "Data", Empty), "yy")
        End If
  
        If (GetValue(rstModel, "FecharValor", False)) Then
          rstAux("Valor").value = wsprintf("%s(%C)%s", GetValue(rstModel, "CaracterSeguranca", NUL), _
                                                       GetValue(rstLan, "Total", ZERO), _
                                                       GetValue(rstModel, "CaracterSeguranca", NUL))
        Else
          rstAux("Valor").value = wsprintf("%s%C%s", GetValue(rstModel, "CaracterSeguranca", NUL), _
                                                       GetValue(rstLan, "Total", ZERO), _
                                                       GetValue(rstModel, "CaracterSeguranca", NUL))
        End If
  
        '// Resolve o extenso do cheque
  
        strExt1 = KeybUCase(KeybExtenso(GetValue(rstLan, "Total", 0)), _
                            GetValue(rstModel, "LetrasMaiusculas", PorFrase))
        strExt2 = NUL
  
        Call SeparaExtenso(strExt1, strExt2, rstModel)
  
        If (GetValue(rstModel, "FecharExtenso", False)) Then
          strExt1 = "(" & strExt1
          If (Len(strExt2)) Then
            strExt2 = strExt2 & ")"
          Else
            strExt1 = strExt1 & ")"
          End If
        End If
  
        '// Completa a primeira string com os caracteres de complemento. A primeira
        '// linha de extenso pode conter até 100 caracteres
  
        strExt1 = strExt1 & KString(GetValue(rstModel, "CaracterComplemento", " "), 100)
        rstAux("Extenso1").value = Left$(strExt1, 100)
  
        '// Completa a segunda string com os caracteres de complemento. A segunda
        '// linha de extenso pode conter até 150 caracteres
  
        strExt2 = strExt2 & KString(GetValue(rstModel, "CaracterComplemento", " "), 150)
        rstAux("Extenso2").value = Left$(strExt2, 150)
         
        '
        ' Atualizando o campo que diz que o registro já foi impresso
        '
        Call ExecuteSQL("UPDATE [Lançamentos]      SET Enviado = 'C' WHERE Banco   = " & CStr(nBanco) & " AND Cheque = " & CStr(GetValue(rstLan, "Cheque", 0)) & " AND Pagamento = " & Quote(InverteData(GetValue(rstLan, "Data", Empty)), "##"))
        Call ExecuteSQL("UPDATE [Duplicatas]       SET Enviada = 'C' WHERE Banco   = " & CStr(nBanco) & " AND Cheque = " & CStr(GetValue(rstLan, "Cheque", 0)) & " AND Pagamento = " & Quote(InverteData(GetValue(rstLan, "Data", Empty)), "##"))
        Call ExecuteSQL("UPDATE [Transf Bancária]  SET Enviada = 'C' WHERE Origem  = " & CStr(nBanco) & " AND Cheque = " & CStr(GetValue(rstLan, "Cheque", 0)) & " AND Data      = " & Quote(InverteData(GetValue(rstLan, "Data", Empty)), "##"))
        
        rstAux.update
      End If
      
      rstLan.MoveNext
      
    Loop
    
    Call UpdateTrans(FORCE_WRITE)
    AddRegImpressao = Not EstaVazio(rstAux)                '// Retorna se obteve...SUCESSO!

  ElseIf (UltimoRetorno() = WL_NORECORD) Then
    MsgFunc wsprintf("Não foi encontrado um modelo de cheque para o banco %l", _
                     GetValue(rstLan, "Banco", ZERO)), vbExclamation
  End If
  Call FechaRecordset(rstModel)
  Exit Function

AddRegImpressao_Erro:
  If (err().Number) Then
    #If (DebugInfo) Then
      MsgErro wsprintf("Erro: %l\n%s\nAddRegImpressao", err.Number, err.Description)
    #Else
      DAOErros NUL
    #End If
  End If
  FechaRecordset rstModel
  Call CancelTrans
End Function

' FUNCTION..: SeparaExtenso
' Objetivo..: Separa a String de Extenso em duas para preencher em todo
'             o campo reservado do cheque.
' Argumentos: [strVlrExt1]: Texto total do extenso.
'             [strVlrExt2]: Ponteiro para uma segunda string.
'             [rstModelo ]: Recordset do modelo para impressão.
' Retorna...: True se obtiver sucesso, False se não.
'             O argumento strVlrExt1 retorna a primeiro linha do extenso e o
'             argumento strVlrExt2 retorna a segunda linha do extenso.
' ----------------------------------------------------------------------------------
Private Function SeparaExtenso(strVlrExt1 As String, strVlrExt2 As String, rstModelo As Object) As Boolean
Dim fntFont   As Font         '// Salva a fonte atual do formulário
Dim sngExt1   As Single       '// Largura da primeira linha do extenso em millímetros
Dim sTmp      As String       '// String temporária
Dim iSpace    As Integer      '// Localização dos espaços no Loop
Dim iPos      As Integer      '// Posição em que a primeira linha será separada
Dim nScale    As Long         '// Escala deste formulário

  Set fntFont = New StdFont

  ' Salva a escala atual do formulário e a altera para Milímetros, que é a escala utilizada
  ' na impressão dos cheques. Também salva a fonte atual e altera para a fonte utilizada na
  ' impressão.

  nScale = Me.ScaleMode
  Me.ScaleMode = vbMillimeters

  fntFont.name = Me.FontName
  fntFont.Size = Me.FontSize
  fntFont.Bold = Me.FontBold
  fntFont.Italic = Me.FontItalic

  On Error Resume Next

  Me.FontName = GetValue(rstModelo, "FonteNome", "Arial")

  If (err().Number) Then

    '// Se a fonte especificada não for encontrada escolho
    '// uma fonte padrão

    Me.FontName = "Ms Sans Serif"
    err().Clear
  End If

  Me.FontSize = GetValue(rstModelo, "FonteSize", 10)
  Me.FontBold = GetValue(rstModelo, "FonteTipo", 0) And 2
  Me.FontItalic = GetValue(rstModelo, "FonteTipo", 0) And 1

  sngExt1 = GetValue(rstModelo, "ExtAPosWidth", 200)
  If (sngExt1 < Me.TextWidth(strVlrExt1)) Then

    '// O texto não cabe todo na primeira linha. Concateno, palavra por palavra da
    '// string, até um tamanho máximo que se acomode na primeira linha, o restante é
    '// colocado na segunda linha.

    Do
      iSpace = InStr(iSpace + 1, strVlrExt1, ESP)
      If (iSpace) Then
        sTmp = Left$(strVlrExt1, iSpace)
        If (Me.TextWidth(sTmp) > sngExt1 - 2) Then Exit Do '// Final da primeira linha encontrado
        iPos = iSpace
      End If
    Loop While (iSpace)

    '// Sai do loop quando a primeira String ultrapassa o tamanho do espaço
    '// reservado para ela no form. Ela é dividida então em duas e a parte que
    '// ficar de fora é colocada na segunda String.

    strVlrExt2 = Right$(strVlrExt1, (Len(strVlrExt1) - iPos))
    strVlrExt1 = Left$(strVlrExt1, iPos)
  End If

  '// Restaura as propriedades do formulário que foram alteradas
  '// no início da função

  Me.FontName = fntFont.name
  Me.FontSize = fntFont.Size
  Me.FontBold = fntFont.Bold
  Me.FontItalic = fntFont.Italic
  Me.ScaleMode = nScale

  SeparaExtenso = True

End Function

' SUB.......: RelatorioCheques
' Objetivo..: Imprime o relatório de cheques
' Argumentos: [rstCheque]: Recordset com os dados a serem impressos.
'             [pdeDest  ]: Destino do relatório.
' ----------------------------------------------------------------------------
Private Sub RelatorioCheques(rstCheques As Object, pdeDest As PrintDestinoEnum)
Dim wrkCheques As KeybReport
Dim strSorted  As String

  On Error GoTo RelatorioCheques_Erro

  If (CreateReport(wrkCheques, pdeDest, "Relatório de Cheques")) Then

    SimpleMsgBar "Gerando Relatório, aguarde..."

    Set wrkCheques.Recordset = rstCheques

    Call PageHeader(wrkCheques, "Relatório de Cheques")

    With wrkCheques
      If (IsValid(txtCheque(0).Text)) And (chkCheque(1).value = vbUnchecked) Then     '// Se o usuário escolheu um Banco em particular
        .UltimaSecao.AddLinha
        .UltimaLinha.AddCampo , wrCSFixedText, wsprintf("Banco: %s %s", _
                                                        txtCheque(0).Text, _
                                                        lblDescCheque(0).Caption), wrTACentro
      End If

      .FontStyle = wrFSBold
      .FontSize = 8

      '// Adiciona o grupo com as colunas. Se o usuário não escolheu um Banco é
      '// adicionado dois campos com o código e nome do banco

      .AddGrupo 1
      .Grupo(1).AddSecao scHeader, 2, wrDBBottomBorder
      If chkCheque(1).value = vbChecked Then
        .Grupo(1).Quebra = "Banco"
        
        With .Grupo(1).Header(2)
          .BorderStyle = wrDot
          .DrawBorder = wrDBAllBorders
          .AddCampo , wrCSFixedText, "Banco:", wrTAEsquerdo, 25
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSDataField, "Banco", wrTAEsquerdo, 10
          .AddCampo , wrCSDataLink, "Nome", wrTAEsquerdo
          .Campo(3).TableLink = "Bancos"
          .Campo(3).DataLink = "Banco = {Banco}"
          .Campo(3).FontStyle = wrFSBold Or wrFSItalic
        End With
        
        .Grupo(1).Header.AddLinha
        .Grupo(1).Header.AddLinha
        
      End If

      With .Grupo(1).Header(.Grupo(1).Header.LinhasCount)
        If (Not IsValid(txtCheque(0).Text)) And (chkCheque(1).value = vbUnchecked) Then
          .AddCampo , wrCSFixedText, "Banco", wrTADireito, 15
          .AddCampo , wrCSFixedText, "Nome", wrTAEsquerdo, 40
        End If
        .AddCampo , wrCSFixedText, "Cheque", wrTADireito, 15
        .AddCampo , wrCSFixedText, "Data", wrTACentro, 20
        .AddCampo , wrCSFixedText, "Valor", wrTADireito, 30
        .AddCampo , wrCSFixedText, "Nominal"
      End With
      
      '// Grupo que imprime os dados dos campos

      .Grupo(1).AddSubGrupo "1"

      .FontStyle = wrFSNormal
      .Grupo(1).Subgrupo(1).AddSecao scDetalhe, 1
      With .Grupo(1).Subgrupo(1).Detalhe.Linha(1)
        If (Not IsValid(txtCheque(0).Text)) And (chkCheque(1).value = vbUnchecked) Then
          .AddCampo "Banco", wrCSDataField, "Banco", wrTADireito, 15
          .AddCampo "Nome", wrCSDataField, "Nome", wrTAEsquerdo, 40
          .Campo("Banco").Formato = String$(6, 48)      '// 48 == "0"
        End If
        .AddCampo "Cheque", wrCSDataField, "Cheque", wrTADireito, 15
        .AddCampo "Data", wrCSDataField, "Data", wrTACentro, 20
        .AddCampo "Valor", wrCSDataField, "Valor", wrTADireito, 30
        .AddCampo "Nominal", wrCSDataField, "Nominal"
        .Campo("Cheque").Formato = String$(6, 48)
        .Campo("Data").Formato = FDATA
        .Campo("Valor").Formato = FMOEDA
      End With
      
      If (chkCheque(1).value = vbChecked) Then
        .Grupo(1).Subgrupo(1).AddSecao scFooter, 2
        With .Grupo(1).Subgrupo(1).Footer(2)
          .BorderStyle = wrDot
          .DrawBorder = wrDBAllBorders
          .Left = 130
          .AddCampo , wrCSFixedText, "Total do Banco:", wrTADireito, 25
          .AddCampo , wrCSSubTotal, "Valor", wrTADireito, 25
        End With
      End If
      
      .Grupo(1).AddSecao scFooter, 2
      With .Grupo(1).Footer(2)
        .DrawBorder = wrDBAllBorders
        .BorderStyle = wrSolid
        .Left = 130
        .AddCampo , wrCSFixedText, "TOTAL GERAL:", wrTADireito, 25
        .AddCampo , wrCSTotal, "Valor", wrTADireito, 25
        .Campo(2).Formato = FMOEDA
      End With
      
    End With
    Set wrkCheques.DatabaseName = GlobalDataBase
    wrkCheques.BeginPrint gTipoDB             '// Exibe a janela ou manda para impressora
    wrkCheques.EndPrint             '// Encerra a rotina de impressão
  End If

  Set wrkCheques = Nothing

RelatorioCheques_Erro:
  If (err().Number) Then
    #If (DebugInfo) Then
      MsgErro wsprintf("Erro: %l\n%s\nRelatorioCheques", err.Number, err.Description)
    #Else
      VBErros NUL
    #End If
  End If
End Sub

' SUB.......: RelatorioCopia
' Objetivo..: Imprime o relatório de cópia de cheque
' Argumentos: [rstCheque ]: Recordset com os dados a serem impressos
'             [rstValores]: Recordset com os valores dos cheques
'             [pdePrint  ]: Destino da impressão.
' --------------------------------------------------------------------------------
Private Sub RelatorioCopia(rstCheque As Object, rstValores As Object, pdePrint As PrintDestinoEnum)
Dim wrkCopia As KeybReport

On Error GoTo erro

  If (CreateReport(wrkCopia, pdePrint, "Cópia de Cheque")) Then

    SimpleMsgBar wsprintf("Gerando relatório, aguarde...")

    With wrkCopia
      Set .Recordset = rstCheque
      .FontSize = 9
      .MargemEsquerda = 20
      .MargemSuperior = 20

      '// Cria o grupo principal

      .AddGrupo "cheque"
      .Grupo("cheque").Quebra = "Cheque"          '// Quebra pelo campo Cheque
      .Grupo("cheque").AddSecao scHeader, 3, wrDBBottomBorder Or wrDBTopBorder
      .FontStyle = wrFSBold
      .FontSize = 9

      With .Grupo("cheque").Header(2)
        .AddCampo , wrCSFixedText, NomeDonaSistema(), , (wrkCopia.ClientWidth / 2)
        .AddCampo , wrCSFixedText, "Cópia de Cheque", , 30, 125
        .AddCampo , wrCSDataField, "Cheque", wrTADireito
        .Campo(3).Formato = String$(6, "0")
      End With

      '// Criando o SubGrupo com os dados do cheque

      .FontStyle = wrFSNormal
      .FontSize = 12
      .Grupo("cheque").AddSubGrupo "detalhe"
      With .Grupo("cheque").Subgrupo("detalhe")
        .Quebra = "Cheque"                 '// Quebra por número de cheque, também
        .PageBreak = IIf((chkCheque(0).value And vbChecked), wrQuebrarDepois, wrSemQuebra)
        .AddSecao scHeader, 12

        With .Header(1)
          .AddCampo , wrCSDataLink, IIf(cboCheque.Text = "Descrição" Or cboCheque.Text = "Histórico", "[Valor Total]", "sum(Valor)"), wrTADireito, 50, 140
          .Campo(1).TableLink = GetTableSource(rstValores, True)
          .Campo(1).DataLink = "Banco = {Banco} AND Cheque = {Cheque}"
          .Campo(1).Formato = FCURRENCY
        End With

        With .Header(3)
          .AddCampo , wrCSDataLink, IIf(cboCheque.Text = "Descrição" Or cboCheque.Text = "Histórico", "[Extenso Total]", "Extenso"), , 140, 30
          .Campo(1).TableLink = GetTableSource(rstValores, True)
          .Campo(1).DataLink = "Banco = {Banco} AND Cheque = {Cheque}"
          .Campo(1).MultiLine = True
        End With

        With .Header(4)
          .AddCampo , wrCSDataLink, "Nominal", , 140, 30
          .Campo(1).TableLink = GetTableSource(rstValores, True)
          .Campo(1).DataLink = "Banco = {Banco} AND Cheque = {Cheque}"
        End With

        With .Header(6)
          .AddCampo , wrCSDataLink, "DtExt", wrTADireito, 100, 70
          .Campo(1).TableLink = GetTableSource(rstValores, True)
          .Campo(1).DataLink = "Banco = {Banco} AND Cheque = {Cheque}"
        End With
        .Header(7).AddCampo , wrCSSimpleLine
        .Header(7).Campo(1).BorderStyle = wrDot

        wrkCopia.FontSize = 9

        With .Header(8)
          .AddCampo , wrCSFixedText, "Número:", , 30
          .AddCampo , wrCSDataField, "Cheque"
          .Campo(1).FontStyle = wrFSBold
          .Campo(2).Formato = String$(6, "0")
        End With

        With .Header(9)
          .AddCampo , wrCSFixedText, "Valor:", , 30
          .AddCampo , wrCSDataLink, "sum(Valor)"
          
          .Campo(1).FontStyle = wrFSBold
          .Campo(2).TableLink = GetTableSource(rstValores, True)
          .Campo(2).DataLink = "Banco = {Banco} AND Cheque = {Cheque}"
          .Campo(2).Formato = FCURRENCY
        End With

        With .Header(10)
          .AddCampo , wrCSFixedText, "Nominal a:", , 30
          .AddCampo , wrCSDataLink, "Nominal"
          .Campo(1).FontStyle = wrFSBold
          .Campo(2).TableLink = GetTableSource(rstValores, True)
          .Campo(2).DataLink = "Banco = {Banco} AND Cheque = {Cheque}"
        End With

        With .Header(11)
          .AddCampo , wrCSFixedText, "Data:", , 30
          .AddCampo , wrCSDataField, "Data"
          .Campo(1).FontStyle = wrFSBold
          .Campo(2).Formato = FDATA
        End With

        With .Header(12)
          .AddCampo , wrCSFixedText, "Banco:", , 30
          .AddCampo , wrCSDataField, "Banco", wrTADireito, 17
          .AddCampo , wrCSDataField, "Nome"
          .Campo(1).FontStyle = wrFSBold
        End With

        '// Se o usuário solicitou que os lançamentos aparecessem no
        '// relatório de cópia de cheques crio um subgrupo que imprimirá
        '// os dados. Caso contrário é criada uma seção detalhes que
        '// conterá, apenas o campo descrição do cheque.

        If (GetItemData(cboCheque) = 2) Then        '// 2 = "Lançamentos"
          wrkCopia.FontSize = 8
          wrkCopia.FontStyle = wrFSBold
          .AddSubGrupo "lanctos", wrDBTopBorder Or wrDBBottomBorder
          .Subgrupo("lanctos").Quebra = "Cheque"    '// Também quebra pelo número do cheque
          .Subgrupo("lanctos").AddSecao scHeader, 1, wrDBBottomBorder
          .Subgrupo("lanctos").Header.BorderStyle = wrDot

          With .Subgrupo("lanctos").Header(1)
            .AddCampo , wrCSFixedText, "Lançamentos", , 26
            .Campo(1).DrawBorder = wrDBRightBorder
            .Campo(1).BorderStyle = wrSolid
            .AddCampo , wrCSFixedText, "Empresa", , 20
            .Campo(2).DrawBorder = wrDBRightBorder
            .Campo(2).BorderStyle = wrSolid
            .AddCampo , wrCSFixedText, "Descrição", , 40
            .Campo(3).DrawBorder = wrDBRightBorder
            .Campo(3).BorderStyle = wrSolid

            If (CentrodeCusto(MFinanceiro)) Then       '// Se o centro de custo de aparecer
              If chkCheque(3).value = vbChecked Then
                .AddCampo "Centro", wrCSFixedText, "Centro de Custo", wrTACentro, 30
                .Campo("Centro").DrawBorder = wrDBRightBorder
                .Campo("Centro").BorderStyle = wrSolid
              End If
            Else
              .AddCampo "Controle", wrCSFixedText, "Controle", wrTACentro, 30
              .Campo("Controle").DrawBorder = wrDBRightBorder
              .Campo("Controle").BorderStyle = wrSolid
            End If

            If chkCheque(4).value = vbChecked Then
              .AddCampo "Conta", wrCSFixedText, "Conta", wrTACentro, 40
              .Campo("Conta").DrawBorder = wrDBRightBorder
              .Campo("Conta").BorderStyle = wrSolid
            End If
            .AddCampo , wrCSFixedText, "Valor", wrTADireito
          End With

          wrkCopia.FontStyle = wrFSNormal
          .Subgrupo("lanctos").AddSecao scDetalhe, 1
          With .Subgrupo("lanctos").Detalhe(1)
            .AddCampo "Lancto", , "Lancto", , 26
            .Campo("Lancto").SuprimirZeros = True
            .AddCampo , , "Emp", , 20
            .AddCampo , , "Desc", , 40

            If (CentrodeCusto(MFinanceiro)) Then
              If chkCheque(3).value = vbChecked Then
                .AddCampo "Centro", , "Custo", wrTADireito, 12
                .Campo("Centro").SuprimirZeros = True
                .AddCampo , , "CsDesc", , 18
              End If
            Else
              .AddCampo "Controle", , "Controle", wrTAEsquerdo, 30, 91
            End If
            
            If chkCheque(4).value = vbChecked Then
              .AddCampo "Conta", , "Conta", wrTADireito, 8
              .Campo("Conta").SuprimirZeros = True
              .AddCampo , , "CtDesc", , 32
            End If
            .AddCampo "vl", , "Valor", wrTADireito
            .Campo("vl").Formato = FMOEDA
            .Campo("vl").SuprimirZeros = True
          End With
        Else
          .AddSecao scDetalhe, 1

          With .Detalhe(1)
            If cboCheque.Text = "Histórico" Then
              .AddCampo , wrCSFixedText, "Histórico:", , 40
              .AddCampo , wrCSDataLink, "Histórico"
              .Campo(2).TableLink = "Cheque"
              .Campo(2).DataLink = "Banco = {Banco} and Cheque = {Cheque}"
              .Campo(1).FontStyle = wrFSBold
              .Campo(2).MultiLine = True
            Else
              .AddCampo , wrCSFixedText, "Descrição:", , 40
              .AddCampo , wrCSDataField, "Desc"
              .Campo(1).FontStyle = wrFSBold
              .Campo(2).MultiLine = True
            End If
          End With
        End If

        wrkCopia.FontSize = 9
        .AddSecao scFooter, 5, wrDBBottomBorder

        With .Footer(2)
          .AddCampo , wrCSFixedText, "Visto:", , 20
          .AddCampo , wrCSFixedText, NUL, , 60
          .AddCampo , wrCSFixedText, NUL, , 80
          .Campo(1).FontStyle = wrFSBold
          .Campo(2).DrawBorder = wrDBBottomBorder
          .Campo(2).BorderStyle = wrDot
          .Campo(3).DrawBorder = wrDBBottomBorder
          .Campo(3).BorderStyle = wrDot
        End With

        With .Footer(3)
          .AddCampo , wrCSSimpleLine
          .Campo(1).BorderStyle = wrDot
        End With

      End With
      
    End With
    
    wrkCopia.BeginPrint gTipoDB
    wrkCopia.EndPrint
    
  End If

  Set wrkCopia = Nothing
  
erro:
  If err.Number <> 0 Then
    MsgBox err.Description
  End If

End Sub

' SUB.......: ChequeImpressao
' Objetivo..: Configura o gerador de relatórios para a impressão do cheque.
' Argumentos: [rstAux]: Recordset auxiliar usado na impressão.
'             [nDest ]: Destino da impressão.
' ------------------------------------------------------------------------------------
Private Sub ChequeImpressao(rstAux As Object, nDest As Long)
Dim wrkImp  As KeybReport         '// Objeto KeybReport
Dim nBco    As Long               '// Código do Banco
Dim rsMdl   As Object          '// Recordset do modelo do cheque
Dim nCamara As Long               '// Código de Compensação

  nBco = CLngDef(txtCheque(0).Text)   '// Código do Banco selecionado
  nCamara = GetFieldValue("Câmara", "Bancos", "Banco = " & CLngDef(txtCheque(0).Text))

  If (AbreRecordset(rsMdl, wsprintf("SELECT * FROM ChqModelos WHERE Número = %l", nCamara), _
                    dbOpenSnapshot) = WL_OK) Then
                    
    If (CreateReport(wrkImp, nDest, "Impressão de Cheques")) Then

      Set wrkImp.Recordset = rstAux

      '// Não há margens neste tipo de impressão

'      wrkImp.MargemEsquerda = ZERO
'      wrkImp.MargemDireita = ZERO
'      wrkImp.MargemSuperior = ZERO
'      wrkImp.MargemInferior = ZERO

      '// Configurando a altura, largura e fonte da página

      wrkImp.PageHeight = GetValue(rsMdl, "Altura", 90)
      wrkImp.PageWidth = GetValue(rsMdl, "Largura", 150)
      wrkImp.FontName = GetValue(rsMdl, "FonteNome", "Arial")
      wrkImp.FontSize = GetValue(rsMdl, "FonteSize", 9)

      Select Case (GetValue(rsMdl, "FonteTipo", 0))
        Case 0: wrkImp.FontStyle = wrFSNormal
        Case 1: wrkImp.FontStyle = wrFSItalic
        Case 2: wrkImp.FontStyle = wrFSBold
        Case 3: wrkImp.FontStyle = wrFSBold Or wrFSItalic
      End Select

      '// Cria o único grupo do relatório e configura para quebra página ao
      '// final da impressão

      wrkImp.AddGrupo "cheque", , , , wrQuebrarDepois

      '// Apenas uma linha é necessária, já que os campos terão suas posições
      '// definidas manualmente

      wrkImp.Grupo("cheque").AddSecao scDetalhe, 1

      With wrkImp.Grupo("cheque").Detalhe(1)
        .Height = wrkImp.PageHeight - (wrkImp.MargemInferior + wrkImp.MargemSuperior)         '// Altura da linha = altura da página
        .AddCampo , , "Valor", wrTACentro, GetValue(rsMdl, "VlrPosWidth"), GetValue(rsMdl, "VlrPosLeft")
        .AddCampo , , "Extenso1", , GetValue(rsMdl, "ExtAPosWidth"), GetValue(rsMdl, "ExtAPosLeft")
        .AddCampo , , "Extenso2", , GetValue(rsMdl, "ExtBPosWidth"), GetValue(rsMdl, "ExtBPosLeft")
        .AddCampo , , "Nominal", , GetValue(rsMdl, "NomPosWidth"), GetValue(rsMdl, "NomPosLeft")
        .AddCampo , , "Local", , GetValue(rsMdl, "LocPosWidth"), GetValue(rsMdl, "LocPosLeft")
        .AddCampo , , "Mês", wrTACentro, GetValue(rsMdl, "MesPosWidth"), GetValue(rsMdl, "MesPosLeft")
        .AddCampo , , "Ano", , GetValue(rsMdl, "AnoPosWidth"), GetValue(rsMdl, "AnoPosLeft")
        .AddCampo , , "BcoChq", , GetValue(rsMdl, "NumBanWidth"), GetValue(rsMdl, "NumBanLeft")
        .Campo(1).Top = GetValue(rsMdl, "VlrPosBase") - .Campo(1).Height
        .Campo(2).Top = GetValue(rsMdl, "ExtAPosBase") - .Campo(2).Height
        .Campo(3).Top = GetValue(rsMdl, "ExtBPosBase") - .Campo(3).Height
        .Campo(4).Top = GetValue(rsMdl, "NomPosBase") - .Campo(4).Height
        .Campo(5).Top = GetValue(rsMdl, "LocPosBase") - .Campo(5).Height
        .Campo(6).Top = GetValue(rsMdl, "LocPosBase") - .Campo(6).Height
        .Campo(7).Top = GetValue(rsMdl, "LocPosBase") - .Campo(7).Height
        .Campo(8).Top = GetValue(rsMdl, "NumBanBase") - .Campo(8).Height
      End With
      wrkImp.BeginPrint gTipoDB
      wrkImp.EndPrint
    End If
    Set wrkImp = Nothing
  End If
  FechaRecordset rsMdl

End Sub

Public Sub FiltraAnalitico(pdeDestino As PrintDestinoEnum)
Dim strDupls   As String        '// Instrução de seleção de dados para Duplicatas
Dim strLanctos As String        '// Instrução de seleção de dados para Lançamentos
Dim strTransf  As String        '// Instrução de seleção de dados para Transf. Bancária
Dim rstDados   As Object        '// Recordset com os dados dos lançamentos
Dim rstAux     As Object        '// Recordset da tabela auxiliar
Dim qdfTemp    As QueryDef      '// Consulta da seleção dos dados
Dim nCodIni    As Long          '// Código Inicial
Dim nCodFim    As Long          '// Código Final
Dim dInicial   As Date          '// Data Inicial
Dim dFinal     As Date          '// Data Final


  SimpleMsgBar "Selecionando dados, aguarde..."

  strDupls = "SELECT CONVERT(VARCHAR, D.Nota) + '-' + CONVERT(VARCHAR, D.Parcela) AS Número, " & _
             "       D.Empresa, D.Emissão, D.Vencimento, D.Liberação," & _
             "       D.Banco, D.Cheque, D.Pagamento As Data, " & _
             "       SUM(D.[Valor Original] + D.Acréscimo - D.Abatimento) As Valor, " & _
             "       D.Enviada As Impresso, 'Duplicata' As Tipo, " & _
             "       C.Nominal " & _
             "  FROM Duplicatas AS D " & _
             " INNER JOIN Cheque AS C ON C.Banco = D.Banco AND C.Cheque = D.Cheque " & _
             " WHERE D.PagRec = 'P' " & Envio("D.Enviada")
               
  strLanctos = "SELECT CONVERT(VARCHAR, L.Código) + '-' + CONVERT(VARCHAR, L.Parcela) AS Número, " & _
               "       L.Empresa, L.Emissão, L.Vencimento, L.Liberação, " & _
               "       L.Banco, L.Cheque, L.Pagamento As Data, " & _
               "       SUM(L.[Valor Original] + L.Acréscimo - L.Abatimento) As Valor, " & _
               "       L.Enviado AS Impresso, 'Lançamento', " & _
               "       C.Nominal " & _
               "  FROM Lançamentos AS L " & _
               " INNER JOIN Cheque C ON C.Banco = L.Banco AND C.Cheque = L.Cheque" & _
               " WHERE L.PagRec = 'P' " & Envio("L.Enviado")

  strTransf = "SELECT CONVERT(VARCHAR, T.Origem) + '- 0' AS Número, " & _
              "       ' ', T.Data As Emissão, T.Data As Vencimento, " & _
              "       T.Data As Liberação ,  0 , T.Cheque,T.Data As Data, SUM(T.Valor) As Valor, " & _
              "       T.Enviada As Impresso, 'Transf. Bancária', " & _
              "       C.Nominal " & _
              "  FROM [Transf Bancária] As T " & _
              " INNER JOIN Cheque AS C ON C.Cheque = T.Cheque AND C.Banco = T.Origem"

  '// Verificando se o usuário indicou um Banco

  nCodIni = CLngDef(txtCheque(0).Text)
  If (nCodIni) Then
    Concat strDupls, " AND D.Banco = ", CStr(nCodIni)
    Concat strLanctos, " AND L.Banco = ", CStr(nCodIni)
    Concat strTransf, " WHERE T.Origem = ", CStr(nCodIni) & Envio("T.Enviada")
  Else
    Concat strTransf, " WHERE T.Origem > 0 " & Envio("T.Enviada")
  End If

  '// Verificando se o usuário filtrou por cheque

  nCodIni = CLngDef(txtCheque(1).Text)
  nCodFim = CLngDef(txtCheque(2).Text)

  If (CBool(nCodIni) And CBool(nCodFim)) Then
    If (nCodIni = nCodFim) Then
      Concat strDupls, " AND D.Cheque = ", CStr(nCodIni)
      Concat strLanctos, " AND L.Cheque = ", CStr(nCodIni)
      Concat strTransf, " AND T.Cheque = ", CStr(nCodIni)
    Else
      Concat strDupls, wsprintf(" AND (D.Cheque BETWEEN %l AND %l)", nCodIni, nCodFim)
      Concat strLanctos, wsprintf(" AND (L.Cheque BETWEEN %l AND %l)", nCodIni, nCodFim)
      Concat strTransf, wsprintf(" AND (T.Cheque BETWEEN %l AND %l)", nCodIni, nCodFim)
    End If
  ElseIf (CBool(nCodIni) And Not CBool(nCodFim)) Then
    Concat strDupls, " AND D.Cheque >= ", CStr(nCodIni)
    Concat strLanctos, " AND L.Cheque >= ", CStr(nCodIni)
    Concat strTransf, " AND T.Cheque >= ", CStr(nCodIni)
  ElseIf (Not CBool(nCodIni) And CBool(nCodFim)) Then
    Concat strDupls, " AND D.Cheque <= ", CStr(nCodFim)
    Concat strLanctos, " AND L.Cheque <= ", CStr(nCodFim)
    Concat strTransf, " AND T.Cheque <= ", CStr(nCodFim)
  Else
    Concat strDupls, " AND D.Cheque > 0"          '// Evita que sejam recuparados
    Concat strLanctos, " AND L.Cheque > 0"        '// registros que não possuam
    Concat strTransf, " AND T.Cheque > 0"         '// cheque
  End If

  '// Verificando se o usuário filtrou por datas

  dInicial = CDateDef(txtCheque(3).Text)
  dFinal = CDateDef(txtCheque(4).Text)
  
  If IsValid(txtCheque(3).Text) And IsValid(txtCheque(4).Text) Then
    If EData(dInicial) And EData(dFinal) Then
      If dFinal < dInicial Then
        MsgFunc "Data Final menor que Data Inicial"
        Exit Sub
      End If
    End If
  End If
  
  If gTipoDB = Access Then

    If (Not IsEmptyDate(dInicial) And Not IsEmptyDate(dFinal)) Then
      If (DateDiff(DD_DIA, dInicial, dFinal) = ZERO) Then
        Concat strDupls, wsprintf(" AND D.Pagamento = #%q#", dInicial)
        Concat strLanctos, wsprintf(" AND L.Pagamento = #%q#", dInicial)
        Concat strTransf, wsprintf(" AND T.Data = #%q#", dInicial)
      Else
        Concat strDupls, wsprintf(" AND (D.Pagamento BETWEEN #%q# AND #%q#)", dInicial, dFinal)
        Concat strLanctos, wsprintf(" AND (L.Pagamento BETWEEN #%q# AND #%q#)", dInicial, dFinal)
        Concat strTransf, wsprintf(" AND (T.Data BETWEEN #%q# AND #%q#)", dInicial, dFinal)
      End If
    ElseIf (Not IsEmptyDate(dInicial) And IsEmptyDate(dFinal)) Then
      Concat strDupls, wsprintf(" AND D.Pagamento >= #%q#", dInicial)
      Concat strLanctos, wsprintf(" AND L.Pagamento >= #%q#", dInicial)
      Concat strTransf, wsprintf(" AND T.Data >= #%q#", dInicial)
    ElseIf (IsEmptyDate(dInicial) And Not IsEmptyDate(dFinal)) Then
      Concat strDupls, wsprintf(" AND D.Pagamento <= #%q#", dFinal)
      Concat strLanctos, wsprintf(" AND L.Pagamento <= #%q#", dFinal)
      Concat strTransf, wsprintf(" AND T.Data <= #%q#", dFinal)
    End If
  
  Else
  
    If (Not IsEmptyDate(dInicial) And Not IsEmptyDate(dFinal)) Then
      If (DateDiff(DD_DIA, dInicial, dFinal) = ZERO) Then
        Concat strDupls, wsprintf(" AND D.Pagamento = '%q'", dInicial)
        Concat strLanctos, wsprintf(" AND L.Pagamento = '%q'", dInicial)
        Concat strTransf, wsprintf(" AND T.Data = '%q'", dInicial)
      Else
        Concat strDupls, wsprintf(" AND (D.Pagamento BETWEEN '%q' AND '%q')", dInicial, dFinal)
        Concat strLanctos, wsprintf(" AND (L.Pagamento BETWEEN '%q' AND '%q')", dInicial, dFinal)
        Concat strTransf, wsprintf(" AND (T.Data BETWEEN '%q' AND '%q')", dInicial, dFinal)
      End If
    ElseIf (Not IsEmptyDate(dInicial) And IsEmptyDate(dFinal)) Then
      Concat strDupls, wsprintf(" AND D.Pagamento >= '%q'", dInicial)
      Concat strLanctos, wsprintf(" AND L.Pagamento >= '%q'", dInicial)
      Concat strTransf, wsprintf(" AND T.Data >= '%q'", dInicial)
    ElseIf (IsEmptyDate(dInicial) And Not IsEmptyDate(dFinal)) Then
      Concat strDupls, wsprintf(" AND D.Pagamento <= '%q'", dFinal)
      Concat strLanctos, wsprintf(" AND L.Pagamento <= '%q'", dFinal)
      Concat strTransf, wsprintf(" AND T.Data <= '%q'", dFinal)
    End If
  
  End If
  
  '// Agrupando os dados em Banco e Cheque
  Concat strDupls, " GROUP BY D.Nota, D.Parcela, D.Empresa, D.Emissão, D.Vencimento, D.Liberação,D.Banco, D.Cheque, D.Pagamento, D.Enviada, C.Nominal"
  Concat strLanctos, " GROUP BY L.Código, L.Parcela, L.Empresa, L.Emissão, L.Vencimento, L.Liberação,L.Banco, L.Cheque, L.Pagamento, L.Enviado, C.Nominal"
  Concat strTransf, " GROUP BY T.Código, T.Data,T.Origem, T.Cheque, T.Enviada, C.Nominal"

  '// Finaliza a instrução unindo as "SELECT's" e acrescentado a
  '// ordem dos dados.
  If (tabCheque.SelectedItem.Key = "relatorio") Then
    strDupls = wsprintf("%s UNION ALL %s UNION ALL %s ORDER BY %s;", _
                        strDupls, strLanctos, strTransf, _
                        IIf((chkCheque(1).value = vbChecked), _
                        "Banco, " & cboCheque.Text, cboCheque.Text))
  Else
    strDupls = wsprintf("%s UNION ALL %s UNION ALL %s ORDER BY Banco, Cheque;", _
                        strDupls, strLanctos, strTransf)
  End If

  If ConsultaExiste("Temp") Then DeleteQuery Nothing, "Temp"

  '// Cria uma Consulta temporária para a seleção dos dados
  If (CreateQuery(qdfTemp, "Temp", strDupls) = WL_OK) Then

    '// Cria uma segunda instrução "SELECT" para somar os valores dos
    '// cheques agrupados na consulta

    If (tabCheque.SelectedItem.Key = "relatorio") Then
      strDupls = "SELECT Banco, Cheque, Data, SUM(Valor) As VlrCheque, Número, Empresa, Emissão, Vencimento, Liberação, Tipo, Nominal FROM Temp " & _
                 "GROUP BY  Banco, Cheque, Data, Número, Empresa, Emissão, Vencimento, " & _
                 "Liberação, Tipo, Nominal ORDER BY " & cboCheque.Text & ";"
    End If
    
    'pt.96014 - Fernando Luís Paludo - (20/11/2009)
    'Alterado a função de conexão de AbreRecordset para AbreRecordsetDAO devido a uma limitação do KeybReport
    If (AbreRecordsetDAO(rstDados, strDupls, dbOpenSnapshot) = WL_OK) Then
      If (tabCheque.SelectedItem.Key = "relatorio") Then
             '// Adiciona os dados
            RelatorioAnalitico rstDados, pdeDestino
      End If
    ElseIf (UltimoRetorno() = WL_NORECORD) Then
      MsgFunc LoadResString(146)
    End If
    
  End If
  
  Call FechaRecordset(rstDados)
  Call DeleteQuery(qdfTemp)
  
 
End Sub

' SUB.......: RelatorioAnalitico
' Objetivo..: Imprime o Relatório Analítico
' Argumentos: [rstCheque]: Recordset com os dados a serem impressos.
'             [pdeDest  ]: Destino do relatório.
' ----------------------------------------------------------------------------
Public Sub RelatorioAnalitico(rstAnalitico As Object, pdeDest As PrintDestinoEnum)
Dim wrkAnalitico As KeybReport
Dim strSorted  As String

  On Error GoTo RelatorioAnalitico_Erro

  If (CreateReport(wrkAnalitico, pdeDest, "Relatório Analítico")) Then

    SimpleMsgBar "Gerando Relatório, aguarde..."

    Set wrkAnalitico.Recordset = rstAnalitico

    Call PageHeader(wrkAnalitico, "Relatório Analítico")

    With wrkAnalitico
      If (IsValid(txtCheque(0).Text)) And (chkCheque(1).value = vbUnchecked) Then     '// Se o usuário escolheu um Banco em particular
        .UltimaSecao.AddLinha
        .UltimaLinha.AddCampo , wrCSFixedText, wsprintf("Banco: %s %s", _
                                                        txtCheque(0).Text, _
                                                        lblDescCheque(0).Caption), wrTACentro
      End If

      .FontStyle = wrFSBold
      .FontSize = 8

      ' Adiciona o grupo com as colunas. Se o usuário não escolheu um Banco são
      ' adicionados dois campos com o código e nome do banco

      .AddGrupo "1"
      .Grupo(1).AddSecao scHeader, 5, wrDBBottomBorder
        
      .Grupo(1).Quebra = "Cheque"
        

      '// Grupo que imprime os dados dos campos
      wrkAnalitico.FontStyle = wrFSNormal
      With .Grupo(1).Header
        wrkAnalitico.Grupo(1).Header(2).DrawBorder = wrDBTopBorder Or wrDBLeftBorder Or wrDBRightBorder
        wrkAnalitico.Grupo(1).Header(3).DrawBorder = wrDBBottomBorder Or wrDBLeftBorder Or wrDBRightBorder
        .Linha(2).AddCampo , wrCSFixedText, "Banco:", wrTAEsquerdo, 25
        .Linha(2).Campo(1).FontStyle = wrFSBold
        .Linha(2).AddCampo , wrCSDataField, "Banco", wrTAEsquerdo, 10
        .Linha(2).AddCampo , wrCSDataLink, "Nome", wrTAEsquerdo
        .Linha(2).Campo(3).TableLink = "Bancos"
        .Linha(2).Campo(3).DataLink = "Banco = {Banco}"
        .Linha(2).Campo(3).FontStyle = wrFSBold Or wrFSItalic
        wrkAnalitico.FontStyle = wrFSBold
        .Linha(3).AddCampo , wrCSFixedText, "Cheque:", wrTAEsquerdo, 15
        wrkAnalitico.FontStyle = wrFSNormal
        .Linha(3).AddCampo "Cheque", wrCSDataField, "Cheque", wrTAEsquerdo, 15
        .Linha(3).Campo("Cheque").Formato = String$(6, 48)
        
        wrkAnalitico.FontStyle = wrFSBold
        .Linha(3).AddCampo , wrCSFixedText, "Data:", wrTACentro, 8
        wrkAnalitico.FontStyle = wrFSNormal
        .Linha(3).AddCampo "Data", wrCSDataField, "Data", wrTACentro, 20
        .Linha(3).Campo("Data").Formato = FDATA

        wrkAnalitico.FontStyle = wrFSBold
        .Linha(3).AddCampo , wrCSFixedText, "Nominal:", wrTAEsquerdo, 15
        wrkAnalitico.FontStyle = wrFSNormal
        .Linha(3).AddCampo "Nominal", wrCSDataField, "Nominal", wrTAEsquerdo, 100
      
        wrkAnalitico.FontStyle = wrFSBold
        .Linha(5).AddCampo , wrCSFixedText, "Número", wrTAEsquerdo, 15
        .Linha(5).AddCampo , wrCSFixedText, "Empresa", wrTAEsquerdo, 75
        .Linha(5).AddCampo , wrCSFixedText, "Emissão", wrTACentro, 20
        .Linha(5).AddCampo , wrCSFixedText, "Vencto", wrTACentro, 20
        .Linha(5).AddCampo , wrCSFixedText, "Liberação", wrTACentro, 20
        .Linha(5).AddCampo , wrCSFixedText, "Tipo", wrTAEsquerdo, 20
        .Linha(5).AddCampo , wrCSFixedText, "Valor", wrTADireito, 40
      End With
 
      'Informa o Cabeçalho dos Valores de Lançamento
      .Grupo(1).AddSecao scDetalhe, 1
      With .Grupo(1).Detalhe.Linha(1)
        wrkAnalitico.FontStyle = wrFSNormal
        .AddCampo "Número", wrCSDataField, "Número", wrTAEsquerdo, 15
        .AddCampo "Empresa", wrCSDataField, "Empresa", wrTAEsquerdo, 75
        .AddCampo "Emissão", wrCSDataField, "Emissão", wrTACentro, 20
        .AddCampo "Vencimento", wrCSDataField, "Vencimento", wrTACentro, 20
        .AddCampo "Liberação", wrCSDataField, "Liberação", wrTACentro, 20
        .AddCampo "Tipo", wrCSDataField, "Tipo", wrTAEsquerdo, 20
        .AddCampo "Valor", wrCSDataField, "VlrCheque", wrTADireito, 40
        .Campo(7).Formato = FMOEDA
      End With
      
      .Grupo(1).AddSecao scFooter, 1

      With .Grupo(1).Footer.Linha(1)
        .DrawBorder = wrDBTopBorder
        .BorderStyle = wrSolid
        .AddCampo , wrCSFixedText, "Subtotal:", wrTAEsquerdo, 173
        .Campo(1).FontStyle = wrFSBold
        .AddCampo "SubTotal", wrCSSubTotal, "VlrCheque", wrTADireito
        .Campo(2).Formato = FMOEDA
      End With
      
      'Exibe o Total de Todos os Cheques
      
      
      With .AddGrupo(2, wrDBBottomBorder Or wrDBTopBorder, wrVPNoFinal, False, wrSemQuebra).AddSecao(scFooter, 1).Linha(1)
        .AddCampo , wrCSFixedText, "Total Geral:", , 30, 125
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSTotal, "VlrCheque", wrTADireito, 50
        .Campo(2).Formato = FMOEDA
      End With
    End With
    
    
    wrkAnalitico.BeginPrint gTipoDB   '// Exibe a janela ou manda para impressora
    wrkAnalitico.EndPrint             '// Encerra a rotina de impressão

  
  Set wrkAnalitico = Nothing

  End If
  
RelatorioAnalitico_Erro:
  If (err().Number) Then
    #If (DebugInfo) Then
      MsgErro wsprintf("Erro: %l\n%s\nRelatorioAnalitico", err.Number, err.Description)
    #Else
      VBErros NUL
    #End If
  End If
End Sub

Private Sub txtImpressoraCheque_Change(Index As Integer)
     If Index = 0 Then
      GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtImpressoraCheque(Index).Text, _
                    lblDescBanco
     End If
End Sub

Private Sub txtImpressoraCheque_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   
   Dim strSelDados As String
   
   ' Campo Banco
   If Index = 0 Then
    PCampo "Bancos", "Bancos", pbCampo, txtImpressoraCheque(Index), "Banco"
   End If
   
    'Campo Cheque
   If Index = 1 Then
    If (IsValid(txtImpressoraCheque(0).Text)) Then
      strSelDados = "SELECT * FROM Cheque WHERE Banco = " & txtImpressoraCheque(0).Text
    Else
      strSelDados = "Cheque"
    End If
    PCampo "Cheques", strSelDados, pbCampo, txtImpressoraCheque(Index), "Cheque"
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
