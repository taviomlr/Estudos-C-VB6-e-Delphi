VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContasAuxiliares 
   Caption         =   "Contas Auxiliares"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   6135
   Tag             =   "Contas"
   Begin VB.Frame fratab 
      Height          =   4335
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   5655
      Begin VB.Frame fraContas 
         Caption         =   "Contas Cont�beis do Modelo"
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
         Height          =   2775
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   1560
         Width           =   5655
         Begin FOX.WDBGrid wdbContasAuxiliares 
            Height          =   2415
            Left            =   120
            TabIndex        =   6
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
            ColCaption(1)   =   "Conta Cont�bil"
            ColFieldName(1) =   "Conta Cont�bil"
         End
      End
      Begin VB.Frame fraModelos 
         Caption         =   "Cadastro de Contas Auxiliares"
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
         Height          =   1815
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   5655
         Begin VB.TextBox txtContas 
            DataField       =   "Grupo"
            Height          =   315
            Index           =   1
            Left            =   960
            MaxLength       =   30
            TabIndex        =   3
            Tag             =   "Contas"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtContas 
            DataField       =   "C�digo"
            Height          =   315
            Index           =   0
            Left            =   960
            MaxLength       =   9
            TabIndex        =   1
            Tag             =   "Contas"
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtContas 
            DataField       =   "Descri��o"
            Height          =   315
            Index           =   2
            Left            =   960
            MaxLength       =   30
            TabIndex        =   5
            Tag             =   "Contas"
            Top             =   1080
            Width           =   4575
         End
         Begin VB.Label lblGrupo 
            Caption         =   "Grupo"
            Height          =   255
            Left            =   2160
            TabIndex        =   11
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label lblContas 
            AutoSize        =   -1  'True
            Caption         =   "Grupo:"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblContas 
            AutoSize        =   -1  'True
            Caption         =   "C�di&go:"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   540
         End
         Begin VB.Label lblContas 
            AutoSize        =   -1  'True
            Caption         =   "D&escri��o:"
            ForeColor       =   &H80000002&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   1080
            Width           =   765
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contas Auxiliares"
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
Attribute VB_Name = "frmContasAuxiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrstContas As Object
Private mlngContas As Long

' FUNCTION..: LibProc
' Objetivo..: Fun��o de retorno de chamada da Lib.
' Argumentos: [sFuncao]: Fun��o que deve ser executada.
'             [lFuncao]: Par�metro adicional, varia conforme a fun��o.
' Retorna...: True se executar a fun��o com sucesso, False se n�o.
' ---------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
  
  Select Case sFuncao
  '
  ' Bot�o Novo
  Case WL_NOVO
    LibProc = (LimpaControles(mrstContas, Me, Tag, mlngContas) = WL_OK)
    FirstFocus txtContas(0)
  '
  ' Bot�o Deletar
  Case WL_DELETAR
    DeletaRegistro mrstContas, Me, Tag, mlngContas
  '
  ' Bot�o Editar
  Case WL_EDITAR
    AlteraValor mlngContas
  '
  ' Bot�o Localizar
  Case WL_LOCALIZAR
    Localizar mrstContas, Me, "Contas Auxiliares", Tag, mlngContas
  '
  ' Bot�o Pesquisar
  Case WL_PESQUISAR
    PRegistro mrstContas, Me, "Contas Auxiliares", "Contas Auxiliares", "Contas Auxiliares", Tag, mlngContas, pbRegistro
  '
  ' Bot�o Primeiro Registro
  Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
    MoveRecordset mrstContas, Me, Tag, mlngContas, lFuncao
  '
  ' Bot�o Sair
  Case WL_SAIR
    Unload Me
    Exit Function
  '
  ' Bot�o Navegar
  Case WL_NAVEGAR
    Browse mrstContas, Me, Tag, mlngContas, "Contas Auxiliares"
  '
  ' Bot�o Salvar
  Case WL_SALVAR
    If Len(txtContas(1).Text) And (Not IsValid(lblGrupo.Caption)) Then
      MsgFunc "Grupo n�o est� cadastrado."
      FirstFocus txtContas(1)
      Exit Function
    End If
    LibProc = (SalvaRegistro(mrstContas, Me, Tag, mlngContas) = WL_OK)
  '
  ' Bot�o Cancelar
  Case WL_CANCELAR
    CancelaEdicao mrstContas, Me, Tag, mlngContas
  '
  ' Op��o Exibir
  Case WL_EXIBIR
    Dim strContas As String
    
    strContas = "SELECT * FROM [Contas Auxiliares] WHERE C�digo = {C�digo};"
    RetornaRegs mrstContas, Me, Tag, strContas, mlngContas
    
  '
  ' Op��o Filtrar
  Case WL_FILTRAR
    Filtrar mrstContas, Me, Tag, "Contas Auxiliares", mlngContas
  '
  ' Registro Duplicado
  Case WL_DUPLICADO
    ResolveDuplicacao Me, txtContas(0)
  
  Case WL_GRIDMSG
    LibProc = wdbContasAuxiliares.Sincronize(mrstContas, lFuncao, Tag, LibProc)
  
  End Select
  
End Function
' EVENT.....: Form_Activate
' Objetivo..: Ativa os menus do formul�rio.
' ------------------------------------------------------------
Private Sub Form_Activate()

  Dim mit(0) As MENUITEMTEMPLATE
  mit(0).mtString = "&Contas Auxiliares..."
  
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
  
  If KeybAcesso(LoadResString(2339)) Then
  
  End If
  
  PosForm Me
  ConfigCampos Me, Tag, Tag
  AbreRecordset mrstContas, "Contas Auxiliares"
  txtContas(0).Text = ProximoNumero("C�digo", "Contas Auxiliares", vbNullString)
  
  lblGrupo.Caption = NUL
  
  DoEvents
  
  DefineAcesso mlngContas, Acesso

  With wdbContasAuxiliares
    If KeybAcesso(LoadResString(2340)) Then
      .TipoAcesso = Acesso()
      .RecordRelation = "SELECT * FROM [Contas de Contas Auxiliares] WHERE Conta = {C�digo}"
      .RecordSource = "SELECT * FROM [Contas de Contas Auxiliares] WHERE Conta = {C�digo}"
    Else
      .Enabled = False
    End If
  End With
  
  SetGridInfo mlngContas
  
  LibProc WL_NOVO
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = UnloadForm(mrstContas, Me, Tag, mlngContas)
End Sub

Private Sub Form_Resize()
  AlinControle Me, TabStrip1, 6255, 2415
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SavePosForm Me
  Set frmContasAuxiliares = Nothing
End Sub

Private Sub txtContas_Change(Index As Integer)
  If Index > 0 Then
    AlteraValor mlngContas
  End If
  
  If Index = 1 Then
    GetAssocValue "Select Descri��o from [Grupos Auxiliares] where C�digo = " & txtContas(Index).Text, lblGrupo
  End If
End Sub

Private Sub txtContas_GotFocus(Index As Integer)

  Selecione txtContas(Index)
  
  MsgBar DescCampo(mrstContas, txtContas(Index).DataField)

End Sub

Private Sub txtContas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 0 Then
    ControlaChave KeyCode, Shift, txtContas(0), mlngContas
  End If
  
  If KeyCode = vbKeyPageDown And Shift = ZERO Then
    If Index = 1 Then
      PCampo "Grupos Auxiliares", "Grupos Auxiliares", pbCampo, txtContas(Index), "C�digo"
    End If
  End If
End Sub

Private Sub txtContas_KeyPress(Index As Integer, KeyAscii As Integer)
  If (Index = 0) Then
    SetMascara KeyAscii, txtContas(Index).SelStart, InputMask(mrstContas, "C�digo")
  Else
    SetMascara KeyAscii, txtContas(Index).SelStart, InputMask(mrstContas, txtContas(Index).DataField)
  End If
End Sub

Private Sub txtContas_LostFocus(Index As Integer)
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
 
Private Sub wdbContasAuxiliares_ChangeRecord(Tipo As Long)
  wdbContasAuxiliares.AlteraValor Tipo, mlngContas
End Sub

Private Sub wdbContasAuxiliares_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageDown And Shift = 0 Then
    If wdbContasAuxiliares.Col = 1 Then
      PCampo "Contas Cont�beis", "Select C�digo, Grupo, Descri��o from Contas", pbCampo, wdbContasAuxiliares, "C�digo"
      KeyCode = 0
      Exit Sub
    End If
  End If
End Sub

Private Sub wdbContasAuxiliares_NeedStatus(UserStatus As Long)
  UserStatus = mlngContas
End Sub

Private Sub wdbContasAuxiliares_UpdateRecord(Tipo As Long, Cancel As Boolean)
  If RecordCount("Select * from Contas where C�digo = " & wdbContasAuxiliares.ColValue(1)) = 0 Then
    MsgFunc "Conta Cont�bil informada n�o existe."
    Cancel = True
    Exit Sub
  End If
End Sub
