VERSION 5.00
Begin VB.Form frmTransfBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferências Bancárias"
   ClientHeight    =   5115
   ClientLeft      =   2430
   ClientTop       =   3360
   ClientWidth     =   8580
   Icon            =   "TransBco.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   8580
   Tag             =   "Transf"
   Begin VB.Frame Frame2 
      Height          =   5050
      Left            =   7140
      TabIndex        =   34
      Top             =   30
      Width           =   1410
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   40
         Top             =   1390
         Width           =   1215
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   39
         Top             =   1795
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   38
         Top             =   2200
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   37
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   36
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   35
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame fraTransfBanco 
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
      ForeColor       =   &H80000006&
      Height          =   3015
      Index           =   1
      Left            =   45
      TabIndex        =   19
      Top             =   2070
      Width           =   7080
      Begin VB.TextBox txtTransf 
         DataField       =   "cd_operacao_contabil"
         Height          =   315
         Index           =   10
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "Transf"
         Top             =   2550
         Width           =   975
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Controle"
         Height          =   315
         Index           =   9
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   10
         Tag             =   "Transf"
         Top             =   1890
         Width           =   1455
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Conta"
         Height          =   315
         Index           =   6
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   7
         Tag             =   "Transf"
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Cheque"
         Height          =   315
         Index           =   8
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   9
         Tag             =   "Transf"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Valor"
         Height          =   315
         Index           =   7
         Left            =   1710
         MaxLength       =   18
         TabIndex        =   8
         Tag             =   "Transf"
         Top             =   1230
         Width           =   2175
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Centro"
         Height          =   315
         Index           =   5
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   11
         Tag             =   "Transf"
         Top             =   2220
         Width           =   975
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Descrição"
         Height          =   315
         Index           =   4
         Left            =   1710
         MaxLength       =   60
         TabIndex        =   6
         Tag             =   "Transf"
         Top             =   570
         Width           =   4815
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Data"
         Height          =   315
         Index           =   3
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Transf"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDescTransf 
         Caption         =   "lblDescTransf(4)"
         Height          =   255
         Index           =   4
         Left            =   2715
         TabIndex        =   30
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Operação Contábil"
         Height          =   195
         Left            =   285
         TabIndex        =   29
         Top             =   2640
         Width           =   1320
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Controle"
         Height          =   195
         Index           =   9
         Left            =   285
         TabIndex        =   26
         Top             =   1965
         Width           =   1320
      End
      Begin VB.Label lblDescTransf 
         Caption         =   "lblDescTransf(3)"
         Height          =   255
         Index           =   3
         Left            =   3270
         TabIndex        =   23
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&onta Financ."
         Height          =   195
         Index           =   6
         Left            =   285
         TabIndex        =   22
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label lblDescTransf 
         Caption         =   "lblDescTransf(2)"
         Height          =   255
         Index           =   2
         Left            =   2700
         TabIndex        =   28
         Top             =   2295
         Width           =   3255
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&heque"
         Height          =   195
         Index           =   8
         Left            =   285
         TabIndex        =   25
         Top             =   1620
         Width           =   1320
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Valor"
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
         Index           =   7
         Left            =   1155
         TabIndex        =   24
         Top             =   1290
         Width           =   450
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Centro de C&usto"
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
         Left            =   225
         TabIndex        =   27
         Top             =   2295
         Width           =   1380
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "De&scrição"
         Height          =   195
         Index           =   4
         Left            =   285
         TabIndex        =   21
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D&ata"
         Height          =   195
         Index           =   3
         Left            =   285
         TabIndex        =   20
         Top             =   300
         Width           =   1320
      End
   End
   Begin VB.Frame fraTransfBanco 
      Caption         =   "Geral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2025
      Index           =   0
      Left            =   45
      TabIndex        =   13
      Top             =   30
      Width           =   7080
      Begin VB.TextBox txtTransf 
         DataField       =   "integracao_bi"
         Height          =   315
         Index           =   12
         Left            =   6540
         TabIndex        =   41
         Tag             =   "Transf"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "empresa_favorecida"
         Height          =   315
         Index           =   11
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   4
         Tag             =   "Transf"
         Top             =   1560
         Width           =   1665
      End
      Begin VB.ComboBox cboTipoRegistro 
         DataField       =   "Tipo_registro"
         Height          =   315
         Left            =   1710
         TabIndex        =   1
         Tag             =   "Transf"
         Text            =   "Fatura"
         Top             =   570
         Width           =   1605
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Destino"
         Height          =   315
         Index           =   2
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   3
         Tag             =   "Transf"
         Top             =   1230
         Width           =   1575
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Origem"
         Height          =   315
         Index           =   1
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "Transf"
         Top             =   900
         Width           =   1575
      End
      Begin VB.TextBox txtTransf 
         DataField       =   "Código"
         Height          =   315
         Index           =   0
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "Transf"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblProdutos 
         AutoSize        =   -1  'True
         Caption         =   "Integração BI"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   88
         Left            =   5250
         TabIndex        =   42
         Top             =   300
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lblDescTransf 
         Caption         =   "lblDescTransf(5)"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   33
         Top             =   1650
         Width           =   3255
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Empresa Favorecida"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   10
         Left            =   165
         TabIndex        =   32
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label lblTpreg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   10
         Left            =   195
         TabIndex        =   31
         Top             =   630
         Width           =   1425
      End
      Begin VB.Label lblDescTransf 
         Caption         =   "lblDescTransf(1)"
         Height          =   255
         Index           =   1
         Left            =   3390
         TabIndex        =   18
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label lblDescTransf 
         Caption         =   "lblDescTransf(0)"
         Height          =   255
         Index           =   0
         Left            =   3390
         TabIndex        =   16
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco Des&tino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   1290
         Width           =   1260
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco &Origem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   15
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label lblTransf 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Códi&go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   1020
         TabIndex        =   14
         Top             =   300
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmTransfBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IDM_TRNBANCOS& = 32000      '// Abre o cadastro de Bancos
Private Const IDM_TRNCUSTOS& = 32001      '// Abre o cadastro de Centros de Custo
Private Const IDM_TRNCONTAS& = 32002      '// Abre o cadastro de Contas Contábeis
Private Const IDM_TRNREPORT& = 32003      '// Abre a janela do Relatório de Transferências

Private mrstTransfB As Object
Private mlngTransfB As Long
Private mrstCheques As Object        'Abre a tabela de cheques
Private mlngCheques As Long             'Controle de Cheques
Private mlngOperacao As Long 'pt. 82335 - Dulcino Júnior

' FUNCTION..: LibProc
' Objetivo..: Função de retorno de chamada da Lib.
' Argumentos: [sFuncao]: Função que deve ser executada.
'             [lFuncao]: Parâmetro adicional, varia conforme a função.
' Retorna...: True se executar a função com sucesso, False, se não.
' ----------------------------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
Dim strTransfB As String
Dim nBanco     As Long        '// Código do Banco atual
Dim nCheque    As Long        '// Número do Cheque atual
Dim blnSalvo   As Boolean     'Identifica se o usuário salvou a alteração para
                              'atualizar a tabela de cheques
Dim cValor     As Currency    'Valor atual da Transferência
Dim objMatrizDAO As New cMatrizContabilizacaoDAO
Dim objMatriz As cMatrizContabilizacao
Dim strTipoRegistro As String
Dim strEmpresa As String
Dim intNumTransf As Long
Dim intOrigem As Long
Dim intDestino As Long

  Select Case sFuncao
  '
  ' Botão Novo
  Case WL_NOVO
    LibProc = (LimpaControles(mrstTransfB, Me, Tag, mlngTransfB) = WL_OK)
    Set objMatriz = objMatrizDAO.Carregar("Fatura")
    If Not objMatriz Is Nothing Then
        mlngOperacao = objMatriz.Transferencia
    End If
    txtTransf(10).Text = mlngOperacao
    Set objMatrizDAO = Nothing
    Set objMatriz = Nothing
    
  '
  Case WL_SETFOCUS: Call FirstFocus(txtTransf(0))
  '
  ' Botão deletar
  Case WL_DELETAR
    Dim sDta As String

    sDta = GetValue(mrstTransfB, "Data", NUL)
    If Not ValidaDatasDiasUteis(0, 0, CDate(sDta), True) Then
        Exit Function
    End If
    
    strTipoRegistro = cboTipoRegistro.Text
    strEmpresa = txtTransf(11).Text
    intNumTransf = txtTransf(0).Text
    intOrigem = txtTransf(1).Text
    intDestino = txtTransf(2).Text
    
    nBanco = GetValue(mrstTransfB, "Origem", ZERO)
    nCheque = GetValue(mrstTransfB, "Cheque", ZERO)
    If (DeletaRegistro(mrstTransfB, Me, Tag, mlngTransfB) = WL_OK) Then
        Call GravarHistoricoTransf(strTipoRegistro, intNumTransf, intOrigem, intDestino, strEmpresa)
        If ((nBanco > 0) And (nCheque > 0)) Then
          If (ExisteCheque(nBanco, nCheque) = ZERO) Then
              DeleteAll "Cheque", wsprintf("Banco = %l AND Cheque = %l", nBanco, nCheque)
          End If
        End If
    End If
  '
  ' Botão Localizar
  Case WL_LOCALIZAR
    localizar mrstTransfB, Me, "Transf Bancária", Tag, mlngTransfB
  '
  ' Botão Pesquisar
  Case WL_PESQUISAR
    PRegistro mrstTransfB, Me, "Transferências Bancárias", "Transf Bancária", _
              "Transf Bancária", Tag, mlngTransfB, PB_REGISTRO
    If EAddNew(mlngTransfB) Then
        txtTransf(10).Text = mlngOperacao
    End If
  '
  ' Botão Primeiro Registro
  Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
    MoveRecordset mrstTransfB, Me, Tag, mlngTransfB, lFuncao
  '
  ' Botão Sair
  Case WL_SAIR
    Unload Me
    Exit Function
  '
  ' Botão Navegar
  Case WL_NAVEGAR
    Browse mrstTransfB, Me, Tag, mlngTransfB, "Transf Bancária"
  '
  ' Botão Salvar
  Case WL_SALVAR
    If (TransfBVerifique()) Then
      nBanco = GetValue(mrstTransfB, "Origem", ZERO)
      nCheque = GetValue(mrstTransfB, "Cheque", ZERO)
      If (SalvaRegistro(mrstTransfB, Me, Tag, mlngTransfB) = WL_OK) Then
        'Vinicius Elyseu (07/03/2016) - Projeto: #100340 / História: #104582
        #If FOXSQL = 1 Then
        If DateDiff("m", txtTransf(3).Text, Now()) > 0 Then
            Call ConfigSys.GravaUltimoLancDup(Lancamento, Format(txtTransf(3).Text, "dd/mm/yyyy"))
            If MsgBox("Este lançamento de transferência bancária tem data anterior a data atual e será necessário fazer o Reprocessamento dos Saldos Bancários. Deseja fazer agora?", vbYesNo, "Alerta para Reprocessamento de Saldo") = vbYes Then
                frmReprocessaSaldo.Show
                frmReprocessaSaldo.etxBanco.valorInteiro = txtTransf(2).Text
                frmReprocessaSaldo.etxBancoFinal.valorInteiro = txtTransf(2).Text
            End If
        End If
        #End If
        If ((nBanco > 0) And (nCheque > 0)) Then

          '// Verifica se ainda existen lançamentos com o cheque
          '// anterior, se não existe, exclui o cheque

          If (ExisteCheque(nBanco, nCheque) = ZERO) Then
            DeleteAll "Cheque", wsprintf("Banco = %l AND Cheque = %l", nBanco, nCheque)
          End If
        End If

        nBanco = GetValue(mrstTransfB, "Origem", ZERO)
        nCheque = GetValue(mrstTransfB, "Cheque", ZERO)

        If ((nBanco > 0) And (nCheque > 0)) Then

          '// Verifica se já existe um registro para o cheque atual
          '// se não existir a função acrescenta.

          strTransfB = wsprintf("FROM Cheque WHERE Banco = %l AND Cheque = %l", nBanco, nCheque)
          If (Recordcount(strTransfB) = ZERO) Then    '// Não há registros
            strTransfB = "INSERT INTO Cheque (Banco, Cheque, Nominal) " & _
                         wsprintf("VALUES (%l, %l, '%s');", nBanco, nCheque, GetFieldValue("[Nome Conta]", "Bancos", "Banco = " & txtTransf(2).Text, , NUL))
            Call ExecuteSQL(strTransfB)
          End If
        End If
        LibProc = True
      End If
    End If
  '
  ' Botão Cancelar
  Case WL_CANCELAR
    CancelaEdicao mrstTransfB, Me, Tag, mlngTransfB
  '
  ' Opção Exibir
  Case WL_EXIBIR
    strTransfB = "SELECT * FROM [Transf Bancária] WHERE Código = {Código};"
    RetornaRegs mrstTransfB, Me, Tag, strTransfB, mlngTransfB
    If EAddNew(mlngTransfB) Then
        If txtTransf(10).Text = "0" Then
            txtTransf(10).Text = mlngOperacao
        End If
    End If
  '
  ' Opção Filtrar
  Case WL_FILTRAR
    Filtrar mrstTransfB, Me, Tag, "Transf Bancária", mlngTransfB
  '
  ' Registro Duplicado
  Case WL_DUPLICADO
    ResolveDuplicacao Me, txtTransf(0), "Transf Bancária"
  
  '
  Case "Bancos"       '// Cadastro de Bancos quando chamado via código, não via menu
    If (KeybAcesso(LoadResString(2003))) Then
      frmBancos.Show
      CallChange frmBancos.hWnd, txtTransf(lFuncao).hWnd
    End If
    Exit Function

  '
  Case WL_MENUCLICK
    Select Case (lFuncao)
      Case IDM_TRNBANCOS
        If (KeybAcesso(LoadResString(2003))) Then
          frmBancos.Show
        End If

      Case IDM_TRNCUSTOS
        If (KeybAcesso(LoadResString(2029))) Then
          frmCusto.Show
          CallChange frmCusto.hWnd, txtTransf(5).hWnd
        End If

      Case IDM_TRNCONTAS
        If (KeybAcesso(LoadResString(2007))) Then
          frmContas.Show
          CallChange frmContas.hWnd, txtTransf(6).hWnd
        End If

      Case IDM_TRNREPORT
        If (KeybAcesso(LoadResString(2032))) Then
          frptTransfBanco.Show vbModal
        End If
        
      End Select
      
  End Select

End Function

Private Sub cboTipoRegistro_Click()
    Dim objMatrizDAO As New cMatrizContabilizacaoDAO
    Dim objMatriz As cMatrizContabilizacao
    
    Set objMatriz = objMatrizDAO.Carregar(cboTipoRegistro.Text)
    If Not objMatriz Is Nothing Then
        mlngOperacao = objMatriz.Transferencia
    End If
    txtTransf(10).Text = mlngOperacao
    Set objMatrizDAO = Nothing
    Set objMatriz = Nothing
    AlteraValor mlngTransfB
End Sub

Private Sub cboTipoRegistro_KeyPress(KeyAscii As Integer)
    If cboTipoRegistro.ListIndex = -1 Then
        KeyAscii = 0
        cboTipoRegistro.Text = "Fatura"
    End If
End Sub

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

'pt. 88289 - Ivo Sousa (07/10/2008)
Private Sub cmdCancelar_Click()
    LibProc (WL_CANCELAR)
End Sub

'pt. 88289 - Ivo Sousa (07/10/2008)
Private Sub cmdExcluir_Click()
    If LibProc(WL_DELETAR) Then
        MsgBox "Registro excluído com sucesso.", vbInformation, NomeModulo
    End If
End Sub

'pt. 88289 - Ivo Sousa (07/10/2008)
Private Sub cmdGravar_Click()
    If LibProc(WL_SALVAR) Then
        MsgBox "Registro gravado com sucesso.", vbInformation, NomeModulo
    End If
End Sub

'pt. 88289 - Ivo Sousa (07/10/2008)
Private Sub cmdNovo_Click()
    LibProc (WL_NOVO)
End Sub

'pt. 88289 - Ivo Sousa (07/10/2008)
Private Sub cmdSair_Click()
    LibProc (WL_SAIR)
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
  GetKeyDown Me, KeyCode, Shift
End Sub

Private Sub Form_Load()
    ConfigCampos Me, "Transf Bancária", Tag
    AbreRecordset mrstTransfB, "Transf Bancária"  'Abre a tabela de Transferências
    
    Label1.Enabled = Configuracao("Utiliza Integração Contábil", False)
    txtTransf(10).Enabled = Configuracao("Utiliza Integração Contábil", False)
    lblDescTransf(4).Enabled = Configuracao("Utiliza Integração Contábil", False)
    
    LibProc WL_NOVO
    mlngTransfB = WL_USERADDNEW
    DefineAcesso mlngTransfB, Acesso
    
    lblDescTransf(0).Caption = vbNullString
    lblDescTransf(1).Caption = vbNullString
    lblDescTransf(2).Caption = vbNullString
    lblDescTransf(3).Caption = vbNullString
    lblDescTransf(5).Caption = vbNullString
    
    ' Verificando se o usuário possui ou não centro de custo
    
    If (Not CentrodeCusto(MFinanceiro)) Then
        txtTransf(5).Enabled = False
        lblTransf(5).Enabled = False
        lblDescTransf(2).Enabled = False
    End If
    'pt. 82335 - Leandro Mesquita
    Call preencheCombo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = UnloadForm(mrstTransfB, Me, Tag, mlngTransfB)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmTransfBanco = Nothing
End Sub

Private Sub txtTransf_Change(Index As Integer)
    Dim strProcBanco As String

    Select Case Index
        'Campo Banco Origem, Destino
        Case 1, 2
            AssocValue "Nome", "Bancos", "Banco = %s", Array(txtTransf(Index).Text), lblDescTransf(Index - 1)
        'Campo Centro de Custo
        Case 5
            AssocValue "Descrição", "Centros", "Código = %s", Array(txtTransf(5).Text), lblDescTransf(2)
        'Campo Conta
        Case 6
            AssocValue "Descrição", "Contas", "Código = %s", Array(Iif(txtTransf(6).Text = "", "0", txtTransf(6).Text)), lblDescTransf(3)
        'Campo Operação Contabil
        Case 10
            If Len(txtTransf(Index).Text) > 0 Then
                lblDescTransf(4).Caption = GetFieldValue("descricao", "OperacaoContabil", "cd_operacao = " & txtTransf(Index).Text)
            Else
                lblDescTransf(4).Caption = vbNullString
            End If
        'pt. 88289 - Ivo Sousa (07/10/2008)
        'Campo Empresa
        Case 11
            If Len(txtTransf(Index).Text) > 0 Then
                lblDescTransf(5).Caption = GetFieldValue("Razão", "Empresas", "Apel = '" & txtTransf(Index).Text & "'")
            Else
                lblDescTransf(5).Caption = vbNullString
            End If
    End Select
    If Index > 0 And Index <> 10 Then
        AlteraValor mlngTransfB
    End If
End Sub

Private Sub txtTransf_GotFocus(Index As Integer)

  Selecione txtTransf(Index)
  Select Case Index
  '
  ' Campos Banco Origem e Destino
  Case 1, 2
    MsgBar DescCampo(mrstTransfB, txtTransf(Index).DataField) & ResolveResString(75, resUM, "Bancos")
  '
  ' Campo Centro de Custo
  Case 5
    MsgBar DescCampo(mrstTransfB, txtTransf(Index).DataField) & ResolveResString(75, resUM, "Custos")
  '
  ' Campo Conta
  Case 6
    MsgBar DescCampo(mrstTransfB, txtTransf(Index).DataField) & ResolveResString(75, resUM, "Contas")
  '
  ' Campo Cheque
  Case 8
    ' Exibe o próximo número de cheque para o banco atual
    '
    'pt. 82431 - Dulcino Júnior (Alterado conforme solicitação do Carlos Dias 23/06/2007)
'    If IsValid(txtTransf(1).Text) And Not IsValid(txtTransf(8).Text) And EstaEditando(mlngTransfB) Then
'      txtTransf(8).Text = ProximoNumero("Cheque", "Cheque", "Banco = " & txtTransf(1).Text)
'    End If
    Selecione txtTransf(8)
    MsgBar DescCampo(mrstTransfB, 8) & ResolveResString(75, resUM, "Cheques")
  '
  ' Qualquer outro campo
  Case Else
    MsgBar DescCampo(mrstTransfB, txtTransf(Index).DataField)
  '
  End Select

End Sub

Private Sub txtTransf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Index = 0) Then
        ControlaChave KeyCode, Shift, txtTransf(0), mlngTransfB
    Else
        If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
            Select Case Index
                Case 1, 2 ' Campo Banco de Origem ou Destino
                    PCampo "Bancos", "Bancos", PB_CAMPO, txtTransf(Index), 0
                Case 5 ' Campo Centro de Custo
                    PCampo "Centros de Custo", "Centros", PB_CAMPO, txtTransf(5), 0
                Case 6 ' Campo Conta
                    PCampo "Contas", "select * from Contas where Ctaati='S'", PB_CAMPO, txtTransf(6), "Código"
                Case 8 ' Campo Cheque
                    ' Exibe somente os cheques existentes com este banco
                    If IsValid(txtTransf(1).Text) Then
                        PCampo "Cheque", "SELECT * FROM Cheque WHERE Banco = " & txtTransf(1).Text & ";", pbCampo Or pbNoFiltro, txtTransf(8), "Cheque"
                    Else
                        PCampo "Cheque", "Cheque", pbCampo Or pbNoFiltro, txtTransf(8), "Cheque"
                    End If
                Case 10 'Campo Operação Contábil
                    PCampo "Operações Contabeis", "OperacaoContabil", pbCampo, txtTransf(Index), "cd_operacao"
                'pt. 88289 - Ivo Sousa (07/10/2008)
                Case 11 'Campo Empresa
                    Call PCampo("Transferência Bancária", "SELECT Apel, Razão, Tipo FROM Empresas", pbCampo, txtTransf(11), "Apel")
            End Select
        End If
    End If
   
End Sub

Private Sub txtTransf_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0 ' Campo Código
        SetMascara KeyAscii, txtTransf(Index).SelStart, fMask("Transf Bancária", "Código")
    Case 1, 2 ' Campo Banco de Origem e Destino
        SetMascara KeyAscii, txtTransf(Index).SelStart, fMask("Bancos", "Banco")
    Case 4:
        If KeyAscii = 60 Or KeyAscii = 62 Then 'Bloquear caracteres "<" e ">"
            KeyAscii = 0
        End If
    
    Case 5 ' Campo Conta
        SetMascara KeyAscii, txtTransf(Index).SelStart, fMask("Contas", "Código")
    Case 6 ' Campo Custo
        SetMascara KeyAscii, txtTransf(Index).SelStart, fMask("Centros", "Código")
    Case 3 ' Campo Data
        SetMascara KeyAscii, txtTransf(3).SelStart, MASK_DATE4
    Case 8 ' Campo Número do Cheque
        SetMascara KeyAscii, txtTransf(8).SelStart, fMask("Cheque", "Cheque")
    Case 7 ' Campo Valor
        DMoeda KeyAscii
    Case 10 ' Campo Operação Contábil
        SetMascara KeyAscii, txtTransf(Index).SelStart, fMask("Centros", "Código")
End Select
End Sub

Private Sub txtTransf_LostFocus(Index As Integer)
  If Index = 0 Then
    LibProc WL_EXIBIR, 0
  End If
End Sub

' FUNCTION..: TransfBVerifique
' Objetivo..: Verifica se o usuário digitou códigos de bancos corretos e se
'             a data é válida
' Retorna...: True se estiver tudo correto, False se não.
' ------------------------------------------------------------------------------
Private Function TransfBVerifique() As Boolean
Dim nReturn As Long         '// Retorno da função ConfRelation

  If cboTipoRegistro.Text = "" Then
    MsgBox "O tipo global é obrigatório.", vbInformation, "Validação de campos"
    cboTipoRegistro.SetFocus
    Exit Function
    TransfBVerifique = False
  End If
  
  '// Obrigatoriedade do codigo.
  If val(txtTransf(0).Text) = 0 Then
    MsgBox "O código é obrigatório.", vbInformation, "Validação de campos"
    Exit Function
  End If
  
  
  '// Obrigatoriedade do primeiro banco.
  If val(txtTransf(1).Text) = 0 Then
    MsgBox "O banco origem é obrigatório.", vbInformation, "Validação de campos"
    Exit Function
  End If
  '// Verificando o código do primeiro banco
  nReturn = ConfRelation(txtTransf(1).Text, lblDescTransf(0).Caption, "Bancos")
  If (nReturn) Then
    If (nReturn = vbYes) Then Call LibProc("Bancos", 1)
    Exit Function
  End If

  '// Obrigatoriedade do segundo banco.
  If val(txtTransf(2).Text) = 0 Then
    MsgBox "O banco destino é obrigatório.", vbInformation, "Validação de campos"
    Exit Function
  End If
  '// Verificando o segundo banco
  nReturn = ConfRelation(txtTransf(2).Text, lblDescTransf(1).Caption, "Bancos")
  If (nReturn) Then
    If (nReturn = vbYes) Then Call LibProc("Bancos", 2)
    Exit Function
  End If

  '// Verificando se os bancos são iguais

  If (IsValid(txtTransf(1).Text) And IsValid(txtTransf(2).Text)) Then
    If (txtTransf(1).Text = txtTransf(2).Text) Then
      MsgBox LoadResString(143), vbInformation, MsgBoxCaption
      Exit Function
    End If
  End If

'  '// Verificando se a data é válida
'
'  If (EEdicao(mlngTransfB)) Then
'    Dim sTmp As String
'    sTmp = GetValue(mrstTransfB, "Data", NUL)
'    If Not ValidaDatasDiasUteis(0, 0, CDate(sTmp)) Then
'        Exit Function
'    End If
'  End If
  
  
    ' Verificando se a data informada é uma data válida para o Movimento Conferico

    'If (IsValid(txtTransf(3).Text)) Then
        If (Not EData(txtTransf(3).Text)) Then
            MsgBox ResolveResString(26, resUM, txtTransf(3).Text), vbInformation, MsgBoxCaption
            Exit Function
        Else
            'pt. 86132 - Ivo Sousa (01/04/2008)
            ' Verifica se o movimento deste período já foi conferido.
            If Not ValidaDatasDiasUteis(0, 0, txtTransf(3).Text) Then
                txtTransf(3).SetFocus
                Exit Function
            End If
            If (CLngDef(txtTransf(5).Text) > 0) And Len(txtTransf(3).Text) Then
                ' Verifica se a data de liberação está dentro da data limite do centro de custo
                If DataLimiteCentroCusto(CLngDef(txtTransf(5).Text), txtTransf(3).Text) Then
                    Exit Function
                End If
            End If
        End If
    'Else
    '    MsgBox ResolveResString(26, resUM, txtTransf(3).Text), vbInformation, MsgBoxCaption
    '    Exit Function
    'End If

  ' Verificando se o centro de custo está cadastrado. Apenas se o campo estiver
  ' visível

  If (txtTransf(5).Enabled) Then
    If (IsValid(txtTransf(5).Text)) Then
      nReturn = ConfRelation(txtTransf(5).Text, lblDescTransf(2).Caption, "Centros de Custo")
      If (nReturn) Then
        If (nReturn = vbYes) Then Call LibProc(WL_MENUCLICK, IDM_TRNCUSTOS)
        Exit Function
      End If
    Else
      ' O campo não pode ser deixado em Branco
      MsgFunc ResolveResString(IDS_COMPLETECAMPO, resUM, "Centro de Custo")
      Exit Function
    End If
  End If

  ' Verificando a Conta

  nReturn = ConfRelation(Iif(txtTransf(6).Text = "", "0", txtTransf(6).Text), lblDescTransf(3).Caption, "Contas")
  If (nReturn) Then
    If (nReturn = vbYes) Then Call LibProc(WL_MENUCLICK, IDM_TRNCONTAS)
    Exit Function
  End If
  
  'Verificar se conta é ativa ou nao
  If GetFieldValue("Ctaati", "Contas", " [Código]=" & Iif(txtTransf(6).Text = "", "0", txtTransf(6).Text)) = "N" Then
    MsgBox "Conta " & txtTransf(6).Text & " não está ativa", vbCritical, MsgBoxCaption
    txtTransf(6).SetFocus
    Exit Function
  End If

  ' Verfica se não há datas diferentes para um mesmo cheque

  If (IsValid(txtTransf(8).Text) And IsValid(txtTransf(1).Text) And _
     IsValid(txtTransf(3).Text)) Then
    If (Not ConfDataCheque(txtTransf(1).Text, txtTransf(8).Text, txtTransf(3).Text, mlngTransfB)) Then
      Exit Function
    End If
  End If
  
  If txtTransf(10).Enabled Then
    If Len(lblDescTransf(4).Caption) = 0 Then
        MsgBox "O campo Operação Contábil deve ser preenchido!", vbInformation, "Validação de Campos"
        Exit Function
    End If
  End If

  '// Obrigatoriedade do valor.
  If Not IsNumeric(txtTransf(7).Text) Then
    txtTransf(7).Text = 0
  End If
  If txtTransf(7).Text = 0 Then
    MsgBox "O valor é obrigatório.", vbInformation, "Validação de campos"
    Exit Function
  End If
  Transform txtTransf(7), mlngTransfB, FMOEDA
  
  TransfBVerifique = True

End Function

Private Sub preencheCombo()
    Dim rdResult As IDBReader
    Dim selCmd As IDBSelectCommand
    
    Aplicacao.Connect
    Set selCmd = Aplicacao.CreateSelectCommand
    selCmd.Table.TableName = "[Tipos Globais]"
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, selCmd)
    While Not rdResult.EOF
        cboTipoRegistro.AddItem rdResult.GetString("tipo")
        rdResult.MoveNext
    Wend
    rdResult.CloseReader
    Set rdResult = Nothing
    Set selCmd = Nothing
    Aplicacao.Disconnect
End Sub


Public Function GravarHistoricoTransf(strTpRegistro As String, intNumTransf As Long, intOrigem As Long, intDestino As Long, Optional strEmpresa As String) As Boolean
    Dim cmd         As IDBInsertCommand
    Dim booExcluido As Boolean
    
On Error GoTo erro_inserindo
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateInsertCommand
    With cmd
        .Table = "FFITransfBancHistorico"
                
        Call .AddValue("[id_seq]", "@pIdSequencial")
        Call .Parameters.add(.CreateParameter("@pIdSequencial", ProximoNumero("id_seq", "FFITransfBancHistorico", NUL)))
        
        Call .AddValue("[enterprise_id]", "@pEnterpriseId")
        Call .Parameters.add(.CreateParameter("@pEnterpriseId", EnterpriseID, dbFieldTypeInt))
        
        Call .AddValue("[cd_estabelecimento]", "@pCdEstabelecimento")
        Call .Parameters.add(.CreateParameter("@pCdEstabelecimento", CdEstabelecimento, dbFieldTypeInt))
        
        Call .AddValue("[tp_registro]", "@pTipoRegistro")
        Call .Parameters.add(.CreateParameter("@pTipoRegistro", strTpRegistro, dbFieldTypeString))
        
        Call .AddValue("[empresa]", "@pEmpresa")
        Call .Parameters.add(.CreateParameter("@pEmpresa", IIf(strEmpresa = "", NUL, strEmpresa), dbFieldTypeString))
                       
        Call .AddValue("[nr_transf]", "@pNumTransf")
        Call .Parameters.add(.CreateParameter("@pNumTransf", intNumTransf, dbFieldTypeLong))
        
        Call .AddValue("[origem]", "@pOrigem")
        Call .Parameters.add(.CreateParameter("@pOrigem", intOrigem, dbFieldTypeLong))
                
        Call .AddValue("[destino]", "@pDestino")
        Call .Parameters.add(.CreateParameter("@pDestino", intDestino, dbFieldTypeLong))
        
        Call .AddValue("[usuario]", "@pUsuario")
        Call .Parameters.add(.CreateParameter("@pUsuario", UserName, dbFieldTypeString))
        
        Call .AddValue("[dataHora]", "@pDataHora")
        Call .Parameters.add(.CreateParameter("@pDataHora", Now, dbFieldTypeDateTime))
        
        Call .AddValue("[integracao_bi]", "@pIntegracaoBI")
        Call .Parameters.add(.CreateParameter("@pIntegracaoBI", 0, dbFieldTypeInt))
       
    End With
    
    booExcluido = Aplicacao.ExecuteUpdate(Aplicacao.GetInternalAuthorization, cmd) > 0
    
    GravarHistoricoTransf = booExcluido
    Aplicacao.Disconnect
    Exit Function
erro_inserindo:
    Call Throw(err)
    Aplicacao.Disconnect
End Function


