VERSION 5.00
Begin VB.Form frmAplicacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Aplicações Financeiras"
   ClientHeight    =   3525
   ClientLeft      =   1620
   ClientTop       =   2220
   ClientWidth     =   7560
   Icon            =   "AplicFin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   7560
   Tag             =   "Aplicacao"
   Begin VB.Frame Frame 
      Height          =   3405
      Index           =   1
      Left            =   6150
      TabIndex        =   23
      Top             =   60
      Width           =   1365
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   30
         Top             =   2100
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   29
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   28
         Top             =   2490
         Width           =   1185
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   27
         Top             =   1710
         Width           =   1185
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   26
         Top             =   930
         Width           =   1185
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   25
         Top             =   540
         Width           =   1185
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   24
         Top             =   150
         Width           =   1185
      End
   End
   Begin VB.Frame fraAplicacao 
      Caption         =   "Gerais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   40
      TabIndex        =   9
      Top             =   45
      Width           =   6105
      Begin VB.TextBox txtAplicacao 
         DataField       =   "cd_operacao_contabil"
         Height          =   315
         Index           =   3
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   6
         Tag             =   "Aplicacao"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtAplicacao 
         DataField       =   "Código"
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Aplicacao"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtAplicacao 
         DataField       =   "Banco"
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "Aplicacao"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtAplicacao 
         DataField       =   "Data"
         Height          =   315
         Index           =   2
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Aplicacao"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox cboAplicacao 
         DataField       =   "Tipo"
         Height          =   315
         Index           =   3
         ItemData        =   "AplicFin.frx":030A
         Left            =   3480
         List            =   "AplicFin.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Aplicacao"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtAplicacao 
         DataField       =   "Descrição"
         Height          =   315
         Index           =   6
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Aplicacao"
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox txtAplicacao 
         DataField       =   "Valor"
         Height          =   315
         Index           =   7
         Left            =   1200
         MaxLength       =   18
         TabIndex        =   8
         Tag             =   "Aplicacao"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtAplicacao 
         DataField       =   "Centro"
         Height          =   315
         Index           =   4
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   5
         Tag             =   "Aplicacao"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtAplicacao 
         DataField       =   "Conta"
         Height          =   315
         Index           =   5
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   4
         Tag             =   "Aplicacao"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblAplication 
         Caption         =   "lblAplication(1)"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   22
         Top             =   2205
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Op. Contabil"
         Height          =   195
         Left            =   165
         TabIndex        =   21
         Top             =   2205
         Width           =   990
      End
      Begin VB.Label lblAplicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Códi&go"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   20
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lblAplicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Banco"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lblAplicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "D&ata"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   18
         Top             =   1080
         Width           =   990
      End
      Begin VB.Label lblAplicacao 
         AutoSize        =   -1  'True
         Caption         =   "Ti&po:"
         Height          =   195
         Index           =   3
         Left            =   3000
         TabIndex        =   17
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lblAplicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "De&scrição"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   16
         Top             =   2550
         Width           =   990
      End
      Begin VB.Label lblAplicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Valor"
         Height          =   195
         Index           =   7
         Left            =   165
         TabIndex        =   15
         Top             =   2910
         Width           =   990
      End
      Begin VB.Label lblAplication 
         Caption         =   "lblAplication(0)"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblAplicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C&onta Financ."
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   13
         Top             =   1470
         Width           =   990
      End
      Begin VB.Label lblAplication 
         Caption         =   "lblAplication(1)"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   12
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label lblAplication 
         Caption         =   "lblAplication(2)"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label lblAplicacao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "C. C&usto"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   10
         Top             =   1830
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmAplicacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If FOXSQL Then
    Private mrstAplicFin As ADODB.Recordset
#Else
    Private mrstAplicFin As Object
#End If
Private mlngAplicFin As Long

' SUB.......: LibProc
' Objetivo..: Função de chamada de retorno para a Lib
' Argumentos: [sFuncao]: Constantes com a função a ser executada
'             [lFuncao]: Constantes com parâmetros adicionais
' Retorna...: True se obtiver sucesso, False se não.
' ----------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
    Dim objMatrizDAO As New cMatrizContabilizacaoDAO
    Dim objMatriz As cMatrizContabilizacao
    
    Select Case sFuncao
    
        Case WL_NOVO
            LibProc = (LimpaControles(mrstAplicFin, Me, Tag, mlngAplicFin) = WL_OK)
            FirstFocus txtAplicacao(0)
            If txtAplicacao(3).Enabled Then
                Set objMatriz = objMatrizDAO.carregarPrimeiro
                If Not objMatriz Is Nothing Then
                    txtAplicacao(3).Text = objMatriz.Aplicacao
                Else
                    txtAplicacao(3).Text = "0"
                End If
            Else
                txtAplicacao(3).Text = "0"
            End If
            Set objMatrizDAO = Nothing
            Set objMatriz = Nothing
      
        Case WL_DELETAR
            Dim strData As String
            
            'Devo verificar se o movimento referente a data da aplicação já está
            'conferido. Caso esteja, o usuário não pode excluir o registro
            strData = GetValue(mrstAplicFin, "Data", NUL)
            If Not ValidaDatasDiasUteis(0, 0, CDate(strData), True) Then
                Exit Function
            End If
            DeletaRegistro mrstAplicFin, Me, Tag, mlngAplicFin
            
        Case WL_EDITAR
            AlteraValor mlngAplicFin
       
        Case WL_LOCALIZAR
            LibProc = (localizar(mrstAplicFin, Me, "Aplicações", Tag, mlngAplicFin) = WL_OK)
        
        Case WL_PESQUISAR
            LibProc = (PRegistro(mrstAplicFin, Me, "Aplicações Financeiras", _
                "Aplicações", "Aplicações", Tag, mlngAplicFin, _
                pbRegistro) = WL_OK)
   
        Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
            MoveRecordset mrstAplicFin, Me, Tag, mlngAplicFin, lFuncao
     
        Case WL_SAIR
            Unload Me
            Exit Function
     
        Case WL_NAVEGAR
            LibProc = (Browse(mrstAplicFin, Me, Tag, mlngAplicFin, "Aplicações") = WL_OK)
        
        Case WL_SALVAR
            If AplicVerifique() Then
                LibProc = (SalvaRegistro(mrstAplicFin, Me, Tag, mlngAplicFin) = WL_OK)
                If LibProc Then
                    'Vinicius Elyseu (07/03/2016) - Projeto: #100340 / História: #104582
                    #If FOXSQL = 1 Then
                    If DateDiff("m", txtAplicacao(2).Text, Now()) > 0 Then
                        Call ConfigSys.GravaUltimoLancDup(Lancamento, Format(txtAplicacao(2).Text, "dd/mm/yyyy"))
                        If MsgBox("Este lançamento de aplicação financeira tem data anterior a data atual e será necessário fazer o Reprocessamento dos Saldos Bancários. Deseja fazer agora?", vbYesNo, "Alerta para Reprocessamento de Saldo") = vbYes Then
                            frmReprocessaSaldo.Show
                            frmReprocessaSaldo.etxBanco.valorInteiro = txtAplicacao(1).Text
                            frmReprocessaSaldo.etxBancoFinal.valorInteiro = txtAplicacao(1).Text
                        End If
                    End If
                    #End If
                    
                    MsgBox "Registro gravado com sucesso!", vbInformation, NomeModulo
                End If
            End If
    
        Case WL_CANCELAR
            CancelaEdicao mrstAplicFin, Me, Tag, mlngAplicFin
        
        Case WL_EXIBIR
            Dim strAplicFin As String
            
            strAplicFin = "SELECT * FROM Aplicações WHERE Código = {Código};"
            LibProc = (RetornaRegs(mrstAplicFin, Me, Tag, strAplicFin, mlngAplicFin) = WL_OK)

        Case WL_FILTRAR
            Filtrar mrstAplicFin, Me, Tag, "Aplicações", mlngAplicFin
        
        Case "Bancos"
            If (KeybAcesso(LoadResString(2003))) Then
                frmBancos.Show
                frmBancos.ZOrder vbBringToFront
                CallChange frmBancos.hWnd, txtAplicacao(1).hWnd
            End If

        Case "Custos"
            If (KeybAcesso(LoadResString(2029))) Then
                frmCusto.Show
                frmCusto.ZOrder vbBringToFront
                CallChange frmCusto.hWnd, txtAplicacao(4).hWnd
            End If
       
        Case "Relatorio"
            If (KeybAcesso(LoadResString(2030))) Then
                frptAplicacoes.Show vbModal
            End If
      
        Case "Contas"
            If (KeybAcesso(LoadResString(2007))) Then
                frmContas.Show
                frmContas.ZOrder vbBringToFront
                CallChange frmContas.hWnd, txtAplicacao(5).hWnd
            End If
   
        Case "Configuração"
            If (KeybAcesso(LoadResString(2199))) Then
                FrmConfCad.Configura "Aplicações Financeiras"
                FrmConfCad.Show vbModal
            End If
    End Select
  
End Function

Private Sub cboAplicacao_Click(Index As Integer)
  AlteraValor mlngAplicFin
End Sub

Private Sub cboAplicacao_GotFocus(Index As Integer)
  MsgBar DescCampo(mrstAplicFin, cboAplicacao(Index).DataField)
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
    LoadMenuTitulos Me        'Carrega os captions dos menus padrão da Lib
    LoadResOptions 1001, cboAplicacao(3)  'Carrega a lista de opções do campo Tipo
    ConfigCampos Me, "Aplicações", Tag
    
    ' Se o usuário não controla Centro de Custo, oculto o controle e redimensiono
    ' o formulário de forma que o usuário não perceba a diferença. Oculta, também,
    ' a opção do menu para abrir o cadastro de Centros de Custo.
    If (Not CentrodeCusto(MFinanceiro)) Then
        lblAplicacao(4).Enabled = False
        txtAplicacao(4).Enabled = False
        lblAplication(1).Enabled = False
    End If
    
    AbreRecordset mrstAplicFin, "Aplicações"
    'PT. 81189 - Dulcino Júnior
    'Integração contábil
    Label1.Enabled = Configuracao("Utiliza Integração Contábil", False)
    txtAplicacao(3).Enabled = Configuracao("Utiliza Integração Contábil", False)
    lblAplication(1).Enabled = Configuracao("Utiliza Integração Contábil", False)
    DoEvents
    DefAddNew mlngAplicFin
    
    DefineAcesso mlngAplicFin, Acesso
    Me.LibProc WL_NOVO
    
    lblAplication(0).Caption = NUL
    lblAplication(1).Caption = NUL
    lblAplication(2).Caption = NUL
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = UnloadForm(mrstAplicFin, Me, Tag, mlngAplicFin)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAplicacao = Nothing
End Sub

' FUNCTION..: AplicVerifique
'
' Faz verificações de rotina nos dados alterados do cadastro.
' Retorna...: True se tudo estiver certo. False se não.
' ------------------------------------------------------------------------------
Private Function AplicVerifique() As Boolean

    ' Verificando se o Banco está cadastrado
    If Len(lblAplication(0).Caption) = 0 Then
        If IsValid(txtAplicacao(1).Text) Then
            If MsgBox(ResolveResString(35, resUM, "O Banco " & txtAplicacao(1).Text, resDOIS, "Bancos"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                LibProc "Bancos", 0
            End If
            Exit Function
        End If
    End If
  
    ' Verificando o Código da Conta
    If (IsValid(txtAplicacao(5).Text)) Then
        If (GetFieldValue("Código", "Contas", "Código = " & txtAplicacao(5).Text) = 0) Then
            If (MsgBox(ResolveResString(35, resUM, "a conta " & txtAplicacao(5).Text, resDOIS, "Contas"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes) Then
                LibProc "Contas"
            End If
            Exit Function
        End If
    End If
  
    'Verificando se a conta está ativa
    If GetFieldValue("Ctaati", "Contas", " [Código]=" & txtAplicacao(5).Text) = "N" Then
        MsgBox "Conta " & txtAplicacao(5).Text & " não está ativa", vbCritical, MsgBoxCaption
        txtAplicacao(5).SetFocus
        Exit Function
    End If
  
    ' Verificando se o Centro de Custo é válido
    If (txtAplicacao(4).Enabled) Then
        If (IsValid(txtAplicacao(4).Text)) Then
            If (Recordcount("FROM Centros WHERE Código = " & txtAplicacao(4).Text) = 0) Then
                If (MsgBox(ResolveResString(35, resUM, "o centro " & txtAplicacao(4).Text, resDOIS, "Centro de Custo"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes) Then
                    LibProc "Custos"
                End If
                Exit Function
            End If
        Else
            ' O usuário deve preencher este campo
            MsgFunc ResolveResString(IDS_COMPLETECAMPO, resUM, "Centro de Custo")
            Exit Function
        End If
    End If

  
    If EEdicao(mlngAplicFin) Then
        If Not ValidaDatasDiasUteis(0, 0, CDate(GetValue(mrstAplicFin, "Data")), True) Then
            Exit Function
        End If
    End If
  
  '// Verificando se a data é válida
  
    If Len(txtAplicacao(2).Text) > 0 Then
        If Not EData(txtAplicacao(2).Text) Then
            MsgBox ResolveResString(26, resUM, txtAplicacao(2).Text), vbInformation, MsgBoxCaption
            Exit Function
        End If
        
        'pt. 86132 - Ivo Sousa (26/03/2008)
        'Validação da data(Dias Úteis)
        If Not ValidaDatasDiasUteis(0, 0, txtAplicacao(2).Text) Then
            txtAplicacao(2).SetFocus
            Exit Function
        End If
        
        If (CLngDef(txtAplicacao(4).Text) > 0) And Len(txtAplicacao(2).Text) Then
            ' Verifica se a data de liberação está dentro da data limite do centro de custo
            If DataLimiteCentroCusto(CLngDef(txtAplicacao(4).Text), txtAplicacao(2).Text) Then
                Exit Function
            End If
        End If
    End If
  
    'PT. 81189 - Dulcino Júnior
    'Integração contábil
    If txtAplicacao(3).Enabled Then
        If Len(lblAplication(3).Caption) = 0 Then
            MsgBox "O campo Operação Contábil deve ser preenchido.", vbInformation, "Validação de Campos"
            txtAplicacao(3).SetFocus
            Exit Function
        End If
    End If
    AplicVerifique = True
End Function

Private Sub txtAplicacao_Change(Index As Integer)
    Dim strBco As String

    Select Case Index
    
        ' Campo Banco
        Case 1
            If Len(txtAplicacao(1).Text) > 0 Then
                strBco = "SELECT Nome FROM Bancos WHERE Banco = " & txtAplicacao(1).Text & ";"
                GetAssocValue strBco, lblAplication(0)
            Else
                lblAplication(0).Caption = vbNullString
            End If
        
        ' Campo Operação Contabil
        Case 3
            If Len(txtAplicacao(Index).Text) > 0 Then
                lblAplication(3).Caption = GetFieldValue("descricao", "OperacaoContabil", "cd_operacao = " & txtAplicacao(Index).Text)
            Else
                lblAplication(3).Caption = vbNullString
            End If
     
        ' Campo Centro de Custo
        Case 4
            If (IsValid(txtAplicacao(4).Text)) Then
                strBco = "SELECT Descrição FROM Centros WHERE Código = " & txtAplicacao(4).Text & ";"
                GetAssocValue strBco, lblAplication(1)
            Else
                lblAplication(1).Caption = NUL
            End If
     
        ' Campo Conta Contábil
        Case 5
            If (IsValid(txtAplicacao(5).Text)) Then
                strBco = "SELECT Descrição FROM Contas WHERE Código = " & txtAplicacao(5).Text & ";"
                GetAssocValue strBco, lblAplication(2)
            Else
                lblAplication(2).Caption = NUL
            End If
    End Select
    If Index > 0 Then
        AlteraValor mlngAplicFin
    End If
End Sub

Private Sub txtAplicacao_GotFocus(Index As Integer)
    Selecione txtAplicacao(Index)
    Select Case Index
     ' Banco
     Case 1
        MsgBar DescCampo(mrstAplicFin, 1) & ResolveResString(75, resUM, "Bancos")
     ' Centro de Custo
     Case 4
        MsgBar DescCampo(mrstAplicFin, 4) & ResolveResString(75, resUM, "Centro de Custo")
     ' Conta
     Case 5
        MsgBar DescCampo(mrstAplicFin, 5) & ResolveResString(75, resUM, "Contas")
     ' Outros campos
     Case Else
        MsgBar DescCampo(mrstAplicFin, txtAplicacao(Index).DataField)
    End Select
End Sub

Private Sub txtAplicacao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Index = 0) Then
        ControlaChave KeyCode, Shift, txtAplicacao(0), mlngAplicFin
    Else
        If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
            Select Case Index
                ' Banco
                Case 1
                    PCampo "Bancos", "Bancos", PB_CAMPO, txtAplicacao(1), "Banco"
                'Operação contabil
                Case 3
                    PCampo "Operações Contabeis", "OperacaoContabil", pbCampo, txtAplicacao(3), "cd_operacao"
                ' Centro de Custo
                Case 4
                    PCampo "Centro de Custo", "Centros", pbCampo, txtAplicacao(4), "Código"
                ' Conta
                Case 5
                    PCampo "Contas", "select * from contas where Ctaati='S'", pbCampo, txtAplicacao(5), "Código"
            End Select
        End If
    End If
End Sub

Private Sub txtAplicacao_KeyPress(Index As Integer, KeyAscii As Integer)
  
    Select Case Index
        ' Campo Código
        Case 0
            SetMascara KeyAscii, txtAplicacao(Index).SelStart, InputMask(mrstAplicFin, "Código")
        ' Campo Banco
        Case 1
            SetMascara KeyAscii, txtAplicacao(Index).SelStart, fMask("Bancos", "Banco")
        ' Campo Conta
        Case 4
            SetMascara KeyAscii, txtAplicacao(Index).SelStart, fMask("Contas", "Código")
        ' Campo Centro de Custo
        Case 5
            SetMascara KeyAscii, txtAplicacao(Index).SelStart, fMask("Centros", "Código")
        ' Campo Data
        Case 2
            SetMascara KeyAscii, txtAplicacao(2).SelStart, MASK_DATE4
        ' Campo Valor
        Case 7
            DMoeda KeyAscii
    End Select
  
End Sub

Private Sub txtAplicacao_LostFocus(Index As Integer)
    If Index = 0 Then
        LibProc WL_EXIBIR, 0
    End If
End Sub
