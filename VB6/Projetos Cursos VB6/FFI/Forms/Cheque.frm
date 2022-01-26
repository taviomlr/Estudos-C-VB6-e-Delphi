VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cheques"
   ClientHeight    =   5145
   ClientLeft      =   285
   ClientTop       =   1770
   ClientWidth     =   10665
   Icon            =   "Cheque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10665
   Tag             =   "Cheques"
   Begin VB.Frame Frame 
      Height          =   5025
      Index           =   1
      Left            =   9240
      TabIndex        =   26
      Top             =   60
      Width           =   1365
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   2100
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   11
         Top             =   2490
         Width           =   1185
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   1710
         Width           =   1185
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   930
         Width           =   1185
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   540
         Width           =   1185
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   150
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5025
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   9165
      Begin VB.Frame fraCheques 
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
         Height          =   2775
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   2070
         Width           =   8895
         Begin ComctlLib.ListView lvwCheques 
            Height          =   2055
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblChqInfo 
            AutoSize        =   -1  'True
            Caption         =   "Data:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   390
         End
         Begin VB.Label lblChqInfo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "#"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   24
            Tag             =   "Cheques"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblChqInfo 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Index           =   2
            Left            =   2760
            TabIndex        =   23
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblChqInfo 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "#"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   22
            Tag             =   "Cheques"
            Top             =   240
            Width           =   1455
         End
         Begin ComctlLib.ImageList imgCheques 
            Left            =   5400
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            MaskColor       =   12632256
            _Version        =   327682
         End
      End
      Begin VB.Frame fraCheques 
         Caption         =   "&Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   1
         Left            =   5160
         TabIndex        =   19
         Top             =   270
         Width           =   3855
         Begin VB.TextBox txtCheques 
            DataField       =   "Histórico"
            Height          =   1455
            Index           =   4
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Tag             =   "Cheques"
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame fraCheques 
         Caption         =   "Principais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   270
         Width           =   4935
         Begin VB.TextBox txtCheques 
            DataField       =   "Banco"
            Height          =   315
            Index           =   0
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   0
            Tag             =   "Cheques"
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtCheques 
            DataField       =   "Cheque"
            Height          =   315
            Index           =   1
            Left            =   1080
            MaxLength       =   9
            TabIndex        =   1
            Tag             =   "Cheques"
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox cboCheques 
            DataField       =   "Situação"
            Height          =   315
            Index           =   2
            ItemData        =   "Cheque.frx":030A
            Left            =   1080
            List            =   "Cheque.frx":030C
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Tag             =   "Cheques"
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox txtCheques 
            DataField       =   "Nominal"
            Height          =   315
            Index           =   3
            Left            =   1080
            MaxLength       =   60
            TabIndex        =   3
            Tag             =   "Cheques"
            Top             =   1320
            Width           =   3615
         End
         Begin VB.Label lblCheques 
            AutoSize        =   -1  'True
            Caption         =   "&Banco:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   270
            Width           =   510
         End
         Begin VB.Label lblCheques 
            AutoSize        =   -1  'True
            Caption         =   "Ch&eque:"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   600
         End
         Begin VB.Label lblCheques 
            AutoSize        =   -1  'True
            Caption         =   "&Situação:"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   16
            Top             =   960
            Width           =   675
         End
         Begin VB.Label lblCheques 
            AutoSize        =   -1  'True
            Caption         =   "&Nominal:"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label lblChqInfo 
            Caption         =   "#"
            Height          =   195
            Index           =   4
            Left            =   2760
            TabIndex        =   14
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   2085
         End
      End
   End
   Begin VB.Menu mnuRegistro 
      Caption         =   "&Registro"
      Begin VB.Menu mnuRegistroNovo 
         Caption         =   "&Novo"
      End
      Begin VB.Menu mnuRegistroSalvar 
         Caption         =   "&Salvar"
      End
      Begin VB.Menu mnuRegistroExcluir 
         Caption         =   "&Excluir"
      End
      Begin VB.Menu mnuRegistroSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegistroPrimeiro 
         Caption         =   "Primeiro"
      End
      Begin VB.Menu mnuRegistroAnterior 
         Caption         =   "&Anterior"
      End
      Begin VB.Menu mnuRegistroProximo 
         Caption         =   "Próximo"
      End
      Begin VB.Menu mnuRegistroUltimo 
         Caption         =   "Último"
      End
      Begin VB.Menu mnuRegistroSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegistroFechar 
         Caption         =   "&Fechar"
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Begin VB.Menu mnuConsultasLocalizar 
         Caption         =   "&Localizar"
      End
      Begin VB.Menu mnuConsultasPesquisar 
         Caption         =   "&Pesquisar"
      End
      Begin VB.Menu mnuConsultasSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultasFiltrar 
         Caption         =   "&Filtrar"
      End
   End
   Begin VB.Menu MnuCadastros 
      Caption         =   "Ca&dastros"
      Begin VB.Menu MnuCadastrosBancos 
         Caption         =   "&Bancos"
      End
   End
   Begin VB.Menu mnuGeracaoCheques 
      Caption         =   "Gerar &Numeração de Cheques"
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu mnuAjudaConteudo 
         Caption         =   "&Conteúdo"
      End
      Begin VB.Menu mnuAjudaWinHelp 
         Caption         =   "Como &usar a Ajuda..."
      End
      Begin VB.Menu mnuAjudaSuporte 
         Caption         =   "Suporte Técnico..."
      End
      Begin VB.Menu mnuAjudaSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjudaSobre 
         Caption         =   "&Sobre..."
      End
   End
End
Attribute VB_Name = "frmCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LIST_ADD = 1    'Para o controle ListView
Private Const LIST_DEL = -1

Private Const IDB_TRANSF = 509          'Imagem para o ListView para Cheques em Transferências
Private Const IDB_DUPLS = 510           'Ídem para Duplicatas
Private Const IDB_LANCTOS = 511         'Ídem para Lançamentos

Private Const IDM_CHQBANCOS& = 32000          '// Cadastro de Bancos

Private mrstCheques As Object
Private mlngCheques As Long

' FUNCTION..: LibProc
' Objetivo..: Função de chamada de retorno para a Lib
' Argumentos: [sFuncao]: Constante com a função a ser executada;
'             [lFuncao]: Informação adicional.
' Retorna...: True se puder executar as funções corretamente, False se não.
' -----------------------------------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean

  Select Case sFuncao
  '
  ' Botão Novo
  Case WL_NOVO
    If LimpaControles(mrstCheques, Me, Tag, mlngCheques) = WL_OK Then
      ChequeInfo LIST_DEL
      LibProc = True
    End If
  '
  ' Botão Excluir
  Case WL_DELETAR
    MsgFunc LoadResString(242)
  '
  ' Botão Localizar
  Case WL_LOCALIZAR
    If (WL_OK = localizar(mrstCheques, Me, "Cheque", Tag, mlngCheques)) Then
      ChequeInfo LIST_DEL
    End If
  '
  ' Botão Pesquisar
  Case WL_PESQUISAR
    If (WL_OK = PRegistro(mrstCheques, Me, "Cheques", "Cheque", "Cheque", _
                          Tag, mlngCheques, pbRegistro)) Then
      ChequeInfo LIST_DEL
    End If
  '
  ' Botão Primerio Registro
  Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
    If (MoveRecordset(mrstCheques, Me, Tag, mlngCheques, lFuncao) <> MC_NOMOVE) Then
      ChequeInfo LIST_DEL
    End If
  '
  ' Botão Sair
  Case WL_SAIR
    Unload Me
    Exit Function
  '
  ' Botão Navegar
  Case WL_NAVEGAR
    If (Browse(mrstCheques, Me, Tag, mlngCheques, "Cheque") = WL_OK) Then
      ChequeInfo LIST_DEL
    End If
  '
  ' Botão Salvar
  Case WL_SALVAR
    If ChqVerifique() Then
      LibProc = (SalvaRegistro(mrstCheques, Me, Tag, mlngCheques) = WL_OK)
    End If
    Exit Function
  '
  ' Botão Cancelar
  Case WL_CANCELAR
    If (LimpaControles(mrstCheques, Me, Tag, mlngCheques) = WL_OK) Then
    'If (CancelaEdicao(mrstCheques, Me, Tag, mlngCheques) = WL_LIMPA) Then
      ChequeInfo LIST_DEL
    End If
  '
  ' Opção Filtrar
  Case WL_FILTRAR
    If (Filtrar(mrstCheques, Me, Tag, "Cheque", mlngCheques) = WL_OK) Then
      ChequeInfo LIST_DEL
    End If
  '
  ' Opção Exibir
  Case WL_EXIBIR
    Dim strChq As String
    strChq = "SELECT * FROM Cheque WHERE Banco = {Banco} AND Cheque = {Cheque};"
    If (RetornaRegs(mrstCheques, Me, Tag, strChq, mlngCheques) = WL_OK) Then
      ChequeInfo LIST_ADD
    ElseIf (UltimoRetorno = WL_LIMPA) Or (UltimoRetorno = WL_ADDNEW) Then
      ChequeInfo LIST_DEL
    End If
    Exit Function
  '
  ' Opção Cadastro de Bancos
  Case "Bancos"
      If (KeybAcesso(LoadResString(2003))) Then
        frmBancos.Show
        CallChange frmBancos.hWnd, txtCheques(0).hWnd
      LibProc = True
    End If

'  Case WL_MENUSELECT
'    If (lFuncao = IDM_CHQBANCOS) Then
'      MsgBar LoadResString(IDM_KIN_BANCOS)
'      LibProc = True: Exit Function
'    End If
'  '
  End Select

End Function

Private Sub cboCheques_Click(Index As Integer)
  AlteraValor mlngCheques
End Sub

Private Sub cboCheques_GotFocus(Index As Integer)
  MsgBar DescCampo(mrstCheques, 2)
End Sub

'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(24/05/2012)
Private Sub cmdAjuda_Click()
    Call LibProc(WL_AJUDA)
End Sub

'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(24/05/2012)
Private Sub cmdCancelar_Click()
    Call LibProc(WL_CANCELAR)
End Sub

'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(24/05/2012)
Private Sub cmdExcluir_Click()
    Call LibProc(WL_DELETAR)
End Sub

'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(24/05/2012)
Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
End Sub

'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(24/05/2012)
Private Sub cmdNovo_Click()
    Call LibProc(WL_NOVO)
End Sub

'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(24/05/2012)
Private Sub cmdPesquisar_Click()
    Call LibProc(WL_PESQUISAR)
End Sub

'Projeto: #1203 - História: # - Desenvolvimento# - João Henrique(24/05/2012)
Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
End Sub

'' EVENT.....: Form_Activate
'' Objetivo..: Cria e exibe os menus do cadastro
'' ------------------------------------------------------------------------------------
'Private Sub Form_Activate()
'Dim mit() As MENUITEMTEMPLATE
'
'  If (LoadMenus(Me)) Then
'    AddMit mit(), MF_STRING, IDM_CHQBANCOS, "&Bancos..."
'    AddMenu Me, "&Cadastros", mit()
'  End If
'
'End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oHelpHtml As New clsHelp
    If KeyCode = vbKeyF1 Then
        oHelpHtml.Origem = 0
        oHelpHtml.hWnd = Me.hWnd
        oHelpHtml.HelpContext = Me.HelpContextID
        Call oHelpHtml.ShowHelp
        Set oHelpHtml = Nothing
    End If
  GetKeyDown Me, KeyCode, Shift
End Sub

Private Sub Form_Load()

  LoadResOptions 1002, cboCheques(2)  'Carrega a lista de opções do campo Situação
  ConfigCampos Me, "Cheque", Tag

  '// Preferi configurar o controle ListView no código para ficar mais fácil
  '// fazer alterações

  lvwCheques.ColumnHeaders.add 1, , "Número", 975, lvwColumnLeft
  lvwCheques.ColumnHeaders.add 2, , "Tipo", 1440, lvwColumnLeft
  lvwCheques.ColumnHeaders.add 3, , "Empresa", 1440, lvwColumnLeft
  lvwCheques.ColumnHeaders.add 4, , "Data", 960, lvwColumnCenter
  lvwCheques.ColumnHeaders.add 5, , "Valor", 1440, lvwColumnRight

  ' A opção Verificar Saldos do menu Utilitários só é visível quando o sistema
  ' está sendo executado na Keyb

  AbreRecordset mrstCheques, "Cheque"
  lblChqInfo(1).Caption = NUL
  lblChqInfo(3).Caption = NUL
  lblChqInfo(4).Caption = NUL

  '// Configurando o controle ImageList

  imgCheques.ImageHeight = 16
  imgCheques.ImageWidth = 16
  imgCheques.MaskColor = vbWhite
  imgCheques.UseMaskColor = True
  imgCheques.ListImages.add 1, "duplicata", LoadResBitmap(IDB_DUPLS)
  imgCheques.ListImages.add 2, "lancamento", LoadResBitmap(IDB_LANCTOS)
  imgCheques.ListImages.add 3, "transferencia", LoadResBitmap(IDB_TRANSF)

  lvwCheques.SmallIcons = imgCheques
  cboCheques(2).ListIndex = 0
  DoEvents
  DefAddNew mlngCheques
  DefineAcesso mlngCheques, Acesso
  'DeleteFlag AC_CADASTRAR, mlngCheques      'O usuário não tem acesso a adicionar cheques
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = UnloadForm(mrstCheques, Me, Tag, mlngCheques)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmCheque = Nothing
End Sub

' FUNCTION..: ChqVerifique
' Objetivo..: Verfica se o cadastro pode ser salvo.
' Verifica se os campos do formulário estão preenchidos corretamente pelo usuário.
' Retorna...: True se puder salvar, False se não.
' ---------------------------------------------------------------------------------
Private Function ChqVerifique() As Boolean

  ' Verifica se o banco cadastrado pelo usuário existe
  If Len(lblChqInfo(4).Caption) = 0 And CLngDef(txtCheques(0).Text) > 0 Then
    If MsgBox(ResolveResString(35, resUM, txtCheques(0).Text, resDOIS, "Bancos"), _
              vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
      LibProc "Bancos", 0
    End If
    Exit Function
  End If
  '
  ' Verifica se o usuário está cancelando um cheque
  '
  If (cboCheques(2).ListIndex = 1) Then             'Cancelado
    Dim lngLanctos As Long

    If (IsValid(txtCheques(0).Text) And IsValid(txtCheques(1).Text)) Then
      SetPtrWait Me
      lngLanctos = Recordcount("FROM [Transf Bancária] WHERE Origem = " & _
                               txtCheques(0).Text & " AND Cheque = " & _
                               txtCheques(1).Text)
      lngLanctos = lngLanctos + Recordcount("FROM Lançamentos WHERE Banco = " & _
                                            txtCheques(0).Text & " AND Cheque = " & _
                                            txtCheques(1).Text)
      lngLanctos = lngLanctos + Recordcount("FROM Duplicatas WHERE Banco = " & _
                                            txtCheques(0).Text & " AND Cheque = " & _
                                            txtCheques(1).Text)
      SetPtrDef Me
      If (lngLanctos) Then
        If MsgFunc("Existem Lançamentos cadastrados com este número de Cheque." & vbCrLf & _
                "Deseja cancelar os pagamentos referentes a este cheque?", vbQuestion + vbYesNo) = vbYes Then
          ExecuteSQL "UPDATE [Transf Bancária] SET Cheque = 0 WHERE Origem = " & _
                               txtCheques(0).Text & " AND Cheque = " & _
                               txtCheques(1).Text
          ExecuteSQL "UPDATE Lançamentos Set Cheque = 0, Pagamento = '' WHERE Banco = " & _
                                            txtCheques(0).Text & " AND Cheque = " & _
                                            txtCheques(1).Text
          ExecuteSQL "UPDATE Duplicatas Set Cheque = 0, Pagamento = '' WHERE Banco = " & _
                                            txtCheques(0).Text & " AND Cheque = " & _
                                            txtCheques(1).Text
        Else
          Exit Function
        End If
        'If (MsgFunc(LoadResString(145), vbQuestion Or vbYesNo) = vbYes) Then
        '  ChequeInfo LIST_ADD
        'End If
        'Exit Function
      End If
    End If
  End If
  ChqVerifique = True

End Function

Private Sub mnuCadastrosBancos_Click()
    If (KeybAcesso(LoadResString(2003))) Then
        frmBancos.Show
        CallChange frmBancos.hWnd, txtCheques(0).hWnd
    End If
End Sub

Private Sub mnuGeracaoCheques_Click()
  'fcalcNumeracaoCheques.Show vbModal
End Sub

Private Sub tabCheques_Click()

End Sub

Private Sub txtCheques_Change(Index As Integer)
  If (Index = 0) Then
    AssocValue "Nome", "Bancos", "Banco = %s", Array(txtCheques(0).Text), lblChqInfo(4)
  ElseIf Index > 1 Then
    AlteraValor mlngCheques
  End If

End Sub

Private Sub txtCheques_GotFocus(Index As Integer)
  Selecione txtCheques(Index)
  If Index = 0 Then
    MsgBar DescCampo(mrstCheques, 0) & ResolveResString(75, resUM, "Bancos")
  Else
    MsgBar DescCampo(mrstCheques, txtCheques(Index).DataField)
  End If
End Sub

Private Sub txtCheques_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index < 2 Then
    If ControlaChave(KeyCode, Shift, txtCheques(Index), mlngCheques) Then
      If (Shift = 0) And (KeyCode = vbKeyPageDown) Then
        Dim lBco As Long        '// Código do Banco atual

        Select Case (Index)
          Case 0                '// Banco
            PCampo "Bancos", "Bancos", pbCampo, txtCheques(0), 0
          Case 1                '// Cheques
            lBco = CLngDef(txtCheques(0).Text)    '// Obtém o código do Banco atual
            If (lBco) Then
              PCampo "Cheque", "SELECT * FROM Cheque WHERE Banco = " & CStr(lBco), _
                     PB_CAMPO, txtCheques(1), "Cheque"
            Else
              PCampo "Cheque", "Cheque", PB_CAMPO, txtCheques(1), "Cheque"
            End If
        End Select

      End If
    End If
  End If
End Sub

Private Sub txtCheques_KeyPress(Index As Integer, KeyAscii As Integer)

  If (Index = 0) Then           'Campo Banco
    SetMascara KeyAscii, txtCheques(Index).SelStart, fMask("Bancos", "Banco")
  ElseIf (Index = 1) Then       'Campo Cheque
    SetMascara KeyAscii, txtCheques(Index).SelStart, fMask("Cheque", "Cheque")
  End If

End Sub

Private Sub txtCheques_LostFocus(Index As Integer)
  If Index < 2 And CLngDef(txtCheques(0).Text) > 0 And CLngDef(txtCheques(1).Text) > 0 Then
    LibProc WL_EXIBIR, 0
  End If
End Sub

' SUB.......: ChequeInfo
' Objetivo..: Traz informação dos cheques para o usuário, ou limpa os campos.
' -------------------------------------------------------------------------------
Private Sub ChequeInfo(intAc As Integer)
Dim nBanco  As Long           '// Código do Banco
Dim nCheque As Long           '// Código do Cheque
Dim cValor  As Currency       '// Valor total do Cheque
Dim strInfo As String         '// Instruções de seleção
Dim rstInfo As Object      '// Variável Recordset com a data do cheque

  If (lvwCheques.ListItems.Count) Then
    lvwCheques.ListItems.Clear
  End If

  nBanco = CLngDef(txtCheques(0).Text)
  nCheque = CLngDef(txtCheques(1).Text)

  If ((nBanco = ZERO) And (nCheque = ZERO)) Then Exit Sub

  SetPtr vbHourglass

  'Criando a primeira instrução para obter os dados do cadastro de Duplicatas

  If gTipoDB = Access Then
    strInfo = wsprintf("SELECT FORMAT(Nota, \'000000\') & ' - ' & FORMAT(Parcela, \'00\'), " & _
                       "Tipo, Empresa, Pagamento, " & _
                       "FORMAT(([Valor Original] + Acréscimo - Abatimento), \'%s\') " & _
                       "AS Valor FROM Duplicatas WHERE PagRec = 'P' AND Banco = %l " & _
                       "AND Cheque = %l;", FMOEDA, nBanco, nCheque)
  Else
    strInfo = wsprintf("SELECT CONVERT(VARCHAR(MAX),Nota) +  ' - ' + CONVERT(VARCHAR(MAX),Parcela), " & _
                       "Tipo, Empresa, Pagamento, " & _
                       "([Valor Original] + Acréscimo - Abatimento) " & _
                       "AS Valor FROM Duplicatas WHERE PagRec = 'P' AND Banco = %l " & _
                       "AND Cheque = %l;", nBanco, nCheque)
  End If
  
  Call ListViewAddItem(lvwCheques, strInfo, "duplicata")

  ' Segunda instrução: Abre o cadastro de Lançamentos

  If gTipoDB = Access Then
    wvsprintf strInfo, "SELECT FORMAT(Código, \'000000\'), Tipo, Empresa, Pagamento, " & _
                      "FORMAT(([Valor Original] + Acréscimo - Abatimento), \'%s\') " & _
                      "AS Valor FROM Lançamentos WHERE PagRec = 'P' AND Banco = %l " & _
                      "AND Cheque = %l;", FMOEDA, nBanco, nCheque
  Else
    wvsprintf strInfo, "SELECT Código, Tipo, Empresa, Pagamento, " & _
                      "([Valor Original] + Acréscimo - Abatimento) " & _
                      "AS Valor FROM Lançamentos WHERE PagRec = 'P' AND Banco = %l " & _
                      "AND Cheque = %l;", nBanco, nCheque
  End If

  Call ListViewAddItem(lvwCheques, strInfo, "lancamento")

  ' Terceira instrução: Abre o cadastro de Transferências Bancárias
  If gTipoDB = Access Then
    wvsprintf strInfo, "SELECT FORMAT(T.Código, \'000000\'), 'Transferência', B.Nome, " & _
                       "T.Data, FORMAT(T.Valor, \'%s\') FROM [Transf Bancária] As T, " & _
                       "Bancos AS B WHERE B.Banco = T.Origem AND T.Origem = %l AND " & _
                       "T.Cheque = %l;", FMOEDA, nBanco, nCheque
  Else
    wvsprintf strInfo, "SELECT T.Código, 'Transferência', B.Nome, " & _
                       "T.Data, T.Valor FROM [Transf Bancária] As T, " & _
                       "Bancos AS B WHERE B.Banco = T.Origem AND T.Origem = %l AND " & _
                       "T.Cheque = %l;", nBanco, nCheque
  End If

  Call ListViewAddItem(lvwCheques, strInfo, "transferencia")

  '// Somando o valor do Cheque no cadastro de Duplicatas

  wvsprintf strInfo, "PagRec = 'P' AND Banco = %l AND Cheque = %l", nBanco, nCheque

  cValor = Soma("([Valor Original] + Acréscimo - Abatimento)", "Duplicatas", _
                strInfo, ZERO)

  '// Somando o valor do Cheque no cadastro de Lançamentos

  cValor = cValor + Soma("([Valor Original] + Acréscimo - Abatimento)", _
                         "Lançamentos", strInfo, ZERO)

  '// Somando o valor do Cheque no cadastro de Transferências Bancárias

  wvsprintf strInfo, "Origem = %l AND Cheque = %l", nBanco, nCheque

  cValor = cValor + Soma("Valor", "[Transf Bancária]", strInfo, ZERO)

  lblChqInfo(3).Caption = Format$(cValor, FCURRENCY)

  '// Trazendo a data do Cheque exibir na janela

  strInfo = "SELECT Pagamento As Data FROM Duplicatas WHERE Banco = %l AND Cheque = %l UNION " & _
            "SELECT Pagamento As Data FROM Lançamentos WHERE Banco = %l AND Cheque = %l UNION " & _
            "SELECT Data FROM [Transf Bancária] WHERE Origem = %l AND Cheque = %l;"

  wvsprintf strInfo, strInfo, nBanco, nCheque, nBanco, nCheque, nBanco, nCheque
  If (WL_OK = AbreRecordset(rstInfo, strInfo, dbOpenSnapshot)) Then
    rstInfo.Move ZERO                 '// Um Refresh nos registros
    If (rstInfo.Recordcount > UM) Then
      MsgFunc wsprintf("Foi detectado que há mais de uma data para este cheque\n" & _
                       "Por favor contate o suporte técnico e relate o problema")
    End If
    lblChqInfo(1).Caption = GetValue(rstInfo, ZERO, NUL)
  End If
  FechaRecordset rstInfo

  SetPtr vbDefault

End Sub

Private Sub mnuRegistroNovo_Click()
  LibProc WL_NOVO
End Sub

Private Sub mnuRegistroSalvar_Click()
  LibProc WL_SALVAR
End Sub

Private Sub mnuRegistroExcluir_Click()
  LibProc WL_DELETAR
End Sub

Private Sub mnuRegistroPrimeiro_Click()
  LibProc WL_PRIMEIRO, MC_MOVEFIRST
End Sub

Private Sub mnuRegistroAnterior_Click()
  LibProc WL_ANTERIOR, MC_MOVEPREV
End Sub

Private Sub mnuRegistroProximo_Click()
  LibProc WL_PROXIMO, MC_MOVENEXT
End Sub

Private Sub mnuRegistroUltimo_Click()
  LibProc WL_ULTIMO, MC_MOVELAST
End Sub

Private Sub mnuRegistroFechar_Click()
  LibProc WL_SAIR
End Sub

Private Sub mnuConsultasLocalizar_Click()
  LibProc WL_LOCALIZAR
End Sub

Private Sub mnuConsultasPesquisar_Click()
  LibProc WL_PESQUISAR
End Sub

Private Sub mnuConsultasFiltrar_Click()
  LibProc WL_FILTRAR
End Sub
