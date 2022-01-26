VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmGruposAuxiliares 
   Caption         =   "Cadastro de Grupos Auxiliares"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   6360
   Tag             =   "Grupos"
   Begin VB.Frame fraGrupos 
      Caption         =   "Grupos"
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
      Height          =   1215
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   5895
      Begin VB.TextBox txtGrupos 
         DataField       =   "Código"
         Height          =   315
         Index           =   0
         Left            =   960
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Grupos"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtGrupos 
         DataField       =   "Descrição"
         Height          =   315
         Index           =   1
         Left            =   960
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "Grupos"
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label lblGrupos 
         AutoSize        =   -1  'True
         Caption         =   "Có&digo:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblGrupos 
         AutoSize        =   -1  'True
         Caption         =   "&Descrição:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   765
      End
   End
   Begin ComctlLib.TabStrip tabGrupos 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2990
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Grupos de Contas Auxiliares"
            Key             =   ""
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
Attribute VB_Name = "frmGruposAuxiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrstGrupos As Object
Private mlngGrupos As Long

' FUNCTION..: LibProc
' Objetivo..: Função de retorno de chamada da Lib.
' Argumentos: [sFuncao]: Função que deve ser executada.
'             [lFuncao]: Parâmetro adicional, varia conforme a função.
' Retorna...: True se executar a função com sucesso, False, se não.
' ----------------------------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
Dim STRgRUPOS    As String

  Select Case sFuncao
  '
  ' Botão Novo
  Case WL_NOVO
    LibProc = (LimpaControles(mrstGrupos, Me, Tag, mlngGrupos) = WL_OK)
    FirstFocus txtGrupos(0)
  '
  ' Botão Deletar
  Case WL_DELETAR
    DeletaRegistro mrstGrupos, Me, Tag, mlngGrupos
  '
  ' Botão Editar
  Case WL_EDITAR
    AlteraValor mlngGrupos
  '
  ' Botão Localizar
  Case WL_LOCALIZAR
    Localizar mrstGrupos, Me, "Grupos Auxiliares", Tag, mlngGrupos
  '
  ' Botão Pesquisar
  Case WL_PESQUISAR
    PRegistro mrstGrupos, Me, "Grupo Auxiliares", "Grupos Auxiliares", "Grupos Auxiliares", _
              Tag, mlngGrupos, PB_REGISTRO
  '
  ' Botão Primeiro Registro
  Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
    MoveRecordset mrstGrupos, Me, Tag, mlngGrupos, lFuncao
  '
  ' Botão Sair
  Case WL_SAIR
    Unload Me
    Exit Function
  '
  ' Botão Navegar
  Case WL_NAVEGAR
    Browse mrstGrupos, Me, Tag, mlngGrupos, "Grupos Auxiliares"
  '
  ' Botão Salvar
  Case WL_SALVAR
    LibProc = (SalvaRegistro(mrstGrupos, Me, Tag, mlngGrupos) = WL_OK)
  '
  ' Botão Cancelar
  Case WL_CANCELAR
    CancelaEdicao mrstGrupos, Me, Tag, mlngGrupos
  '
  ' Opção Exibir
  Case WL_EXIBIR
    STRgRUPOS = "SELECT * FROM [Grupos Auxiliares] WHERE Código = {Código};"
    RetornaRegs mrstGrupos, Me, Tag, STRgRUPOS, mlngGrupos
  '
  ' Opção filtrar
  Case WL_FILTRAR
    Filtrar mrstGrupos, Me, Tag, "Grupos Auxiliares", mlngGrupos
  '
  ' Registro Duplicado
  Case WL_DUPLICADO
    ResolveDuplicacao Me, txtGrupos(0)
  
  End Select
  
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
  GetKeyDown Me, KeyCode, Shift
End Sub

Private Sub Form_Load()

  If KeybAcesso(LoadResString(2338)) Then
  
  End If

  PosForm Me
  ConfigCampos Me, Tag, Tag
  AbreRecordset mrstGrupos, "Grupos Auxiliares"
  txtGrupos(0).Text = ProximoNumero("Código", "Grupos Auxiliares", vbNullString)
  DoEvents
  mlngGrupos = WL_USERADDNEW
  DefineAcesso mlngGrupos, Acesso
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = UnloadForm(mrstGrupos, Me, Tag, mlngGrupos)
End Sub

Private Sub Form_Resize()
  AlinControle Me, tabGrupos, 6255, 2175
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SavePosForm Me
  Set frmGrupos = Nothing
End Sub

Private Sub txtGrupos_Change(Index As Integer)
  If Index > 0 Then
    AlteraValor mlngGrupos
  End If
End Sub

Private Sub txtGrupos_GotFocus(Index As Integer)
  Selecione txtGrupos(Index)
  MsgBar DescCampo(mrstGrupos, txtGrupos(Index).DataField)
End Sub

Private Sub txtGrupos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 0 Then
    ControlaChave KeyCode, Shift, txtGrupos(0), mlngGrupos
  End If
End Sub

Private Sub txtGrupos_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Then
    SetMascara KeyAscii, txtGrupos(0).SelStart, InputMask(mrstGrupos, "Código")
  End If
End Sub

Private Sub txtGrupos_LostFocus(Index As Integer)
  If Index = 0 Then
    LibProc WL_EXIBIR
  End If
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


