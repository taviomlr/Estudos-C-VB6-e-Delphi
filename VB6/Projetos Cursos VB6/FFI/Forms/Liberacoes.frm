VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmLiberacoes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Liberações"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "Liberacoes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   7860
   Tag             =   "CAD"
   Begin VB.Frame Frame 
      Height          =   3105
      Index           =   1
      Left            =   6450
      TabIndex        =   10
      Top             =   60
      Width           =   1365
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   2100
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   16
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   15
         Top             =   2490
         Width           =   1185
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   14
         Top             =   1710
         Width           =   1185
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   13
         Top             =   930
         Width           =   1185
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   12
         Top             =   540
         Width           =   1185
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   11
         Top             =   150
         Width           =   1185
      End
   End
   Begin VB.Frame fraPrincipal 
      Height          =   3105
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   6375
      Begin VB.TextBox txtCAD 
         DataField       =   "Empresa"
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Tag             =   "CAD"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCAD 
         DataField       =   "Validade"
         Height          =   315
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Tag             =   "CAD"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtCAD 
         DataField       =   "Responsável"
         Height          =   315
         Index           =   2
         Left            =   1200
         TabIndex        =   5
         Tag             =   "CAD"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtCAD 
         DataField       =   "Motivo"
         Height          =   1125
         Index           =   3
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Tag             =   "CAD"
         Top             =   1320
         Width           =   4995
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblDesc 
         Height          =   210
         Left            =   2880
         OleObjectBlob   =   "Liberacoes.frx":0582
         TabIndex        =   9
         Top             =   240
         Width           =   3345
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   225
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "Liberacoes.frx":05EE
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   225
         Index           =   1
         Left            =   120
         OleObjectBlob   =   "Liberacoes.frx":065E
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   345
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "Liberacoes.frx":06D0
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   225
         Index           =   3
         Left            =   120
         OleObjectBlob   =   "Liberacoes.frx":0748
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Menu mnuRegistro 
      Caption         =   "&Registro"
      Begin VB.Menu mnuRegistroAdicionar 
         Caption         =   "Adicionar"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuRegistroEditar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuRegistroDeletar 
         Caption         =   "Deletar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuRegistroSalvar 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuRegistroBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegistroPrimeiro 
         Caption         =   "Primeiro"
      End
      Begin VB.Menu mnuRegistroAnterior 
         Caption         =   "Anterior"
      End
      Begin VB.Menu mnuRegistroProximo 
         Caption         =   "Próximo "
      End
      Begin VB.Menu mnuRegistroUltimo 
         Caption         =   "Último "
      End
      Begin VB.Menu mnuRegistroBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegistroFechar 
         Caption         =   "Fechar"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "&Consultas"
      Begin VB.Menu mnuConsultasLocalizar 
         Caption         =   "Localizar"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuConsultasPesquisa 
         Caption         =   "Pesquisar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuConsultasBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultasFiltrar 
         Caption         =   "Filtrar"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frmLiberacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCAD As Object
Dim lngCAD As Long
Const Tabela$ = "Liberações"

' FUNCTION..: LibProc
' Objetivo..: Função de retorno de chamada da Lib.
' Argumentos: [sFuncao]: Função que deve ser executada.
'             [lFuncao]: Parâmetro adicional, varia conforme a função.
' Retorna...: True se executar a função com sucesso, False, se não.
' ----------------------------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
Dim lngResult As Long
Dim strFunc As String
Dim rstTab  As Object
  
  Select Case sFuncao
  '
  ' Botão Novo
  Case WL_NOVO
    LibProc = (LimpaControles(rstCAD, Me, Tag, lngCAD) = WL_OK)
    LetFocus txtCAD(0).hWnd
  '
  ' Botão Deletar
  Case WL_DELETAR
    DeletaRegistro rstCAD, Me, Tag, lngCAD
  '
  ' Botão Editar
  Case WL_EDITAR
    AlteraValor lngCAD
  '
  ' Botão Localizar
  Case WL_LOCALIZAR
    localizar rstCAD, Me, Tabela$, Tag, lngCAD
  '
  ' Botão Pesquisar
  Case WL_PESQUISAR
    strFunc = "SELECT * FROM " & Tabela$ & ";"
    PRegistro rstCAD, Me, Caption, Tabela$, strFunc, Tag, _
              lngCAD, pbRegistro
  '
  ' Botão Primeiro Registro
  Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
    MoveRecordset rstCAD, Me, Tag, lngCAD, lFuncao
  '
  ' Botão Navegar
  Case WL_NAVEGAR
    Browse rstCAD, Me, Tag, lngCAD, Tabela$
  '
  ' Botão Sair
  Case WL_SAIR
    Unload Me
    Exit Function
  '
  ' Botão Salvar
  Case WL_SALVAR
   If VerificaCampos Then
    LibProc = (SalvaRegistro(rstCAD, Me, Tag, lngCAD) = WL_OK)
   End If
  '
  ' Botão Cancelar
  Case WL_CANCELAR
    CancelaEdicao rstCAD, Me, Tag, lngCAD
  '
  ' Opção Exibir
  Case WL_EXIBIR
    'Pt. 95368 - Moacir Pfau(21/10/2009)
    If gTipoDB = Access Then
      'strFunc = "SELECT * FROM " & Quote(Tabela$, "[]") & " WHERE Empresa = '{Empresa}' and Validade = #{Validade}#;"
    'Else
      
      
        If IsDate(txtCAD(1).Text) And Trim(txtCAD(0).Text) <> "" Then
            'Pt. 95368 - Moacir Pfau(30/10/2009)
            strFunc = "SELECT * FROM " & Quote(Tabela$, "[]") & " WHERE Empresa = '" & txtCAD(0).Text & "' and Validade = #" & Format(txtCAD(1).Text, "DD/MM/YYYY") & "#;"
            If (AbreRecordset(rstTab, strFunc, dbOpenDynaset) = WL_OK) Then
                RetornaRegs rstCAD, Me, Tag, strFunc, lngCAD
            End If
            FechaRecordset (rstTab)
        End If
        
    End If
    
  '
  ' Opção Filtrar
  Case WL_FILTRAR
    Filtrar rstCAD, Me, Tag, Tabela$, lngCAD
  
  Case "Empresas"
    If KeybAcesso(LoadResString(2037)) Then
      frmEmpresas.Show
      frmEmpresas.ZOrder
      CallChange frmEmpresas.hWnd, txtCAD(0).hWnd
    End If

  End Select
  
End Function

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = UnloadForm(rstCAD, Me, Tag, lngCAD)
End Sub

Private Sub mnuRegistroAdicionar_Click()
  LibProc WL_NOVO, 0
End Sub

Private Sub mnuRegistroEditar_Click()
  LibProc WL_EDITAR, 0
End Sub

Private Sub mnuRegistroDeletar_Click()
  LibProc WL_DELETAR, 0
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

Private Sub mnuRegistroSalvar_Click()
    LibProc WL_SALVAR, 0
End Sub

Private Sub mnuRegistroUltimo_Click()
  LibProc WL_ULTIMO, MC_MOVELAST
End Sub

Private Sub mnuRegistroFechar_Click()
  LibProc WL_SAIR, 0
End Sub

Private Sub mnuConsultasLocalizar_Click()
  LibProc WL_LOCALIZAR, 0
End Sub

Private Sub mnuConsultasPesquisa_Click()
  LibProc WL_PESQUISAR, 0
End Sub

Private Sub mnuConsultasFiltrar_Click()
  LibProc WL_FILTRAR, 0
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
Dim i As Integer
    '
    ' Define o acesso
    DefineAcesso lngCAD, Acesso
    '
    ' Abre a tabela
    If AbreRecordset(rstCAD, Tabela$) = WL_ERRO Then
        MsgBox LoadResString(IDS_ERR)
        Unload Me
        Exit Sub
    End If
    
    '
    ' Limpa as Descrições
    lblDesc.Caption = ""
        
    '
    ' Configura os campos
    ConfigCampos Me, Tabela$, Tag
    
    '
    ' Define status inicial
    Call Me.LibProc(WL_NOVO)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmLiberacoes = Nothing
End Sub

Private Sub txtCAD_Change(Index As Integer)
    If Index > 1 Then
        AlteraValor lngCAD
    ElseIf Index = 0 Then
        GetAssocValue "Select Razão From empresas where apel = " & Quote(txtCAD(Index).Text, "''"), lblDesc
    End If
End Sub

Private Sub txtCAD_Click(Index As Integer)
    LetFocus txtCAD(Index).hWnd
End Sub

Private Sub txtCAD_GotFocus(Index As Integer)
    If (EAdicao(lngCAD) Or EAddNew(lngCAD)) And Index = 2 Then
        txtCAD(Index).Text = UserName
    End If
    Selecione txtCAD(Index)
    MsgBar DescCampo(rstCAD, txtCAD(Index).DataField)
End Sub

Private Sub txtCAD_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index < 2 Then
        If Index = 0 Then
            If (KeyCode = vbKeyPageDown And Shift = ZERO) Then
                PCampo "Empresas", "Empresas", pbCampo, txtCAD(Index), "Apel"
            End If
        End If
        ControlaChave KeyCode, Shift, txtCAD(Index), lngCAD
    End If
End Sub

Private Sub txtCAD_KeyPress(Index As Integer, KeyAscii As Integer)
    SetMascara KeyAscii, txtCAD(Index).SelStart, fMask(Tabela, txtCAD(Index).DataField)
End Sub

Private Sub txtCAD_LostFocus(Index As Integer)
    If Index < 2 Then
        Call Me.LibProc(WL_EXIBIR)
    End If
End Sub


Private Function VerificaCampos() As Boolean

  VerificaCampos = True
  If ((Len(txtCAD(0).Text) > 0) And (Len(lblDesc.Caption) = 0)) Then
    VerificaCampos = False
    If (MsgBox(ResolveResString(IDS_DADONAOENCONTRADO, "|1", txtCAD(0).Text, "|2", "Empresas"), vbQuestion Or vbYesNo, _
        MsgBoxCaption) = vbYes) Then
        LibProc "Empresas", 0
        Exit Function
    End If
  End If

End Function
