VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmModelos 
   Caption         =   "Cadastro de Modelos"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   6120
   Tag             =   "Modelos"
   Begin VB.Frame fratab 
      Height          =   4335
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   5655
      Begin VB.Frame fraContas 
         Caption         =   "Grupos do Modelo"
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
         Height          =   3135
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   1200
         Width           =   5655
         Begin FOX.WDBGrid wdbGruposModelos 
            Height          =   2775
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   7646
            CorFonteFixa    =   -2147483640
            BackColor       =   -2147483636
            RowHeightMin    =   225
            ColWidthMin     =   1440
            RowHeight       =   225
            BeginProperty FonteFixa {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ScrollBars      =   3
            NumColunas      =   1
            ColWidth(1)     =   1605
            ColCaption(1)   =   "Grupo"
            ColFieldName(1) =   "Grupo"
         End
      End
      Begin VB.Frame fraModelos 
         Caption         =   "Cadastro de Modelos"
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
         Height          =   1335
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   5655
         Begin VB.TextBox txtModelos 
            DataField       =   "Descrição"
            Height          =   315
            Index           =   2
            Left            =   960
            MaxLength       =   30
            TabIndex        =   3
            Tag             =   "Modelos"
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox txtModelos 
            DataField       =   "Código"
            Height          =   315
            Index           =   0
            Left            =   960
            MaxLength       =   9
            TabIndex        =   1
            Tag             =   "Modelos"
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblContas 
            AutoSize        =   -1  'True
            Caption         =   "D&escrição:"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   765
         End
         Begin VB.Label lblContas 
            AutoSize        =   -1  'True
            Caption         =   "Códi&go:"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   540
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Modelos"
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
Attribute VB_Name = "frmModelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrstModelos As Object
Private mlngModelos As Long

' FUNCTION..: LibProc
' Objetivo..: Função de retorno de chamada da Lib.
' Argumentos: [sFuncao]: Função que deve ser executada.
'             [lFuncao]: Parâmetro adicional, varia conforme a função.
' Retorna...: True se executar a função com sucesso, False se não.
' ---------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
  
  Select Case sFuncao
  '
  ' Botão Novo
  Case WL_NOVO
    LibProc = (LimpaControles(mrstModelos, Me, Tag, mlngModelos) = WL_OK)
    FirstFocus txtModelos(0)
  '
  ' Botão Deletar
  Case WL_DELETAR
    DeletaRegistro mrstModelos, Me, Tag, mlngModelos
  '
  ' Botão Editar
  Case WL_EDITAR
    AlteraValor mlngModelos
  '
  ' Botão Localizar
  Case WL_LOCALIZAR
    Localizar mrstModelos, Me, "Modelos", Tag, mlngModelos
  '
  ' Botão Pesquisar
  Case WL_PESQUISAR
    PRegistro mrstModelos, Me, "Modelos", "Modelos", "Modelos", Tag, mlngModelos, pbRegistro
  '
  ' Botão Primeiro Registro
  Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
    MoveRecordset mrstModelos, Me, Tag, mlngModelos, lFuncao
  '
  ' Botão Sair
  Case WL_SAIR
    Unload Me
    Exit Function
  '
  ' Botão Navegar
  Case WL_NAVEGAR
    Browse mrstModelos, Me, Tag, mlngModelos, "Modelos"
  '
  ' Botão Salvar
  Case WL_SALVAR
    LibProc = (SalvaRegistro(mrstModelos, Me, Tag, mlngModelos) = WL_OK)
  '
  ' Botão Cancelar
  Case WL_CANCELAR
    CancelaEdicao mrstModelos, Me, Tag, mlngModelos
  '
  ' Opção Exibir
  Case WL_EXIBIR
    Dim strModelos As String
    
    strModelos = "SELECT * FROM Modelos WHERE Código = {Código};"
    RetornaRegs mrstModelos, Me, Tag, strModelos, mlngModelos
    
  '
  ' Opção Filtrar
  Case WL_FILTRAR
    Filtrar mrstModelos, Me, Tag, "Modelos", mlngModelos
  '
  ' Registro Duplicado
  Case WL_DUPLICADO
    ResolveDuplicacao Me, txtModelos(0)
  
  Case WL_GRIDMSG
    LibProc = wdbGruposModelos.Sincronize(mrstModelos, lFuncao, Tag, LibProc)
  
  End Select
  
End Function
' EVENT.....: Form_Activate
' Objetivo..: Ativa os menus do formulário.
' ------------------------------------------------------------
Private Sub Form_Activate()

  Dim mit(0) As MENUITEMTEMPLATE
  mit(0).mtString = "&Modelos..."
  
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
  
  PosForm Me
  ConfigCampos Me, Tag, Tag
  AbreRecordset mrstModelos, "Modelos"
  txtModelos(0).Text = ProximoNumero("Código", "Modelos", vbNullString)
  
  DoEvents
  
  SetGridInfo mlngModelos
  DefineAcesso mlngModelos, Acesso

  With wdbGruposModelos
    If KeybAcesso(LoadResString(2337)) Then
      .TipoAcesso = Acesso()
      .RecordRelation = "SELECT * FROM [Grupos de Modelos] WHERE Modelo = {Código}"
      .RecordSource = "SELECT * FROM [Grupos de Modelos] WHERE Modelo = {Código}"
    Else
      .Enabled = False
    End If
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = UnloadForm(mrstModelos, Me, Tag, mlngModelos)
End Sub

Private Sub Form_Resize()
  AlinControle Me, TabStrip1, 6255, 2415
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SavePosForm Me
  Set frmModelos = Nothing
End Sub

Private Sub txtModelos_Change(Index As Integer)
  If Index > 0 Then
    AlteraValor mlngModelos
  End If
End Sub

Private Sub txtModelos_GotFocus(Index As Integer)

  Selecione txtModelos(Index)
  
  MsgBar DescCampo(mrstModelos, txtModelos(Index).DataField)

  
End Sub

Private Sub txtModelos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 0 Then
    ControlaChave KeyCode, Shift, txtModelos(0), mlngModelos
  End If
End Sub

Private Sub txtModelos_KeyPress(Index As Integer, KeyAscii As Integer)
  If (Index = 0) Then
    SetMascara KeyAscii, txtModelos(Index).SelStart, InputMask(mrstModelos, "Código")
  Else
    SetMascara KeyAscii, txtModelos(Index).SelStart, InputMask(mrstModelos, txtModelos(Index).DataField)
  End If
End Sub

Private Sub txtModelos_LostFocus(Index As Integer)
  If Index = 0 Then
    LibProc WL_EXIBIR, 0
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
 
Private Sub wdbGruposModelos_ChangeRecord(Tipo As Long)
  wdbGruposModelos.AlteraValor Tipo, mlngModelos
End Sub

Private Sub wdbGruposModelos_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageDown And Shift = 0 Then
    If wdbGruposModelos.Col = 1 Then
      PCampo "Grupos Auxiliares", "Grupos Auxiliares", pbCampo, wdbGruposModelos, "Código"
    End If
  End If
End Sub

Private Sub wdbGruposModelos_NeedStatus(UserStatus As Long)
  UserStatus = mlngModelos
End Sub

Private Sub wdbGruposModelos_UpdateRecord(Tipo As Long, Cancel As Boolean)
  If RecordCount("Select * from [Grupos Auxiliares] where Código = " & wdbGruposModelos.ColValue(1)) = 0 Then
    MsgFunc "Grupo Auxiliar informado não existe."
    Cancel = True
    Exit Sub
  End If
End Sub
