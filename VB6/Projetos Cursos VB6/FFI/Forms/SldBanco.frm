VERSION 5.00
Begin VB.Form frmSldBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saldos Banc�rios"
   ClientHeight    =   2250
   ClientLeft      =   2430
   ClientTop       =   3360
   ClientWidth     =   5640
   Icon            =   "SldBanco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2250
   ScaleWidth      =   5640
   Tag             =   "SldBanco"
   Begin VB.Frame fraSldBanco 
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
      Height          =   2150
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5535
      Begin VB.TextBox txtSldBanco 
         DataField       =   "Valor"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   6
         Tag             =   "SldBanco"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtSldBanco 
         DataField       =   "Per�odo"
         Height          =   315
         Index           =   1
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "SldBanco"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtSldBanco 
         DataField       =   "Banco"
         Height          =   315
         Index           =   0
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "SldBanco"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image imgInformativa 
         Height          =   480
         Left            =   70
         Picture         =   "SldBanco.frx":030A
         Top             =   1610
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "              Para incluir Saldos Banc�rios, utilize os cadastros de Lan�amentos                  a Pagar ou Lan�amentos a Receber"
         Height          =   495
         Left            =   20
         TabIndex        =   8
         Top             =   1600
         Width           =   5470
      End
      Begin VB.Label lblSldBcoDesc 
         Caption         =   "lblSldBcoDesc"
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblSldBanco 
         AutoSize        =   -1  'True
         Caption         =   "&Valor:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblSldBanco 
         AutoSize        =   -1  'True
         Caption         =   "&Per�odo:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblSldBanco 
         AutoSize        =   -1  'True
         Caption         =   "&Banco:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmSldBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrstSldBanco As Object
Private mlngSldBanco As Long

' FUNCTION..: LibProc
' Objetivo..: Fun��o de retorno de chamada da Lib.
' Argumentos: [sFuncao]: Fun��o que deve ser executada.
'             [lFuncao]: Par�metro adicional, varia conforme a fun��o.
' Retorna...: True se executar a fun��o com sucesso, False, se n�o.
' ----------------------------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
Dim strSldBanco As String
  
  Select Case sFuncao
  '
  ' Bot�o Novo
  Case WL_NOVO
    LibProc = (LimpaControles(mrstSldBanco, Me, Tag, mlngSldBanco) = WL_OK)
    FirstFocus txtSldBanco(0)
  '
  ' Bot�o Excluir
  Case WL_DELETAR
    Dim sTmp As String

'    sTmp = GetValue(mrstSldBanco, "Per�odo", NUL)
'    If (Len(sTmp)) Then
'      If (MovConferido(sTmp, "KIF")) Then Exit Function
'    End If
    'Projeto: #218 - Hist�ria: #164 - Desenvolvimento#418 - Moacir Pfau(18/09/2012)
    MsgBox "N�o poder� ser realizado exclus�o na rotina de Saldos Banc�rios." & vbCrLf & vbCrLf & "Para atualizar os saldos banc�rios favor utilizar os cadastros de " & vbCrLf & "Lan�amentos a Pagar ou Lan�amentos a Receber."
    
  ' DeletaRegistro mrstSldBanco, Me, Tag, mlngSldBanco
  '
  ' Bot�o Editar
  Case WL_EDITAR
    AlteraValor mlngSldBanco
  '
  ' Bot�o Localizar
  Case WL_LOCALIZAR
    localizar mrstSldBanco, Me, "Saldos Banc�rios", Tag, mlngSldBanco
  '
  ' Bot�o Pesquisar
  Case WL_PESQUISAR
    'Pt. 95368 - Moacir Pfau(11/11/2009)
    txtSldBanco(1).MaxLength = 10
    PRegistro mrstSldBanco, Me, "Saldos Banc�rios", "Saldos Banc�rios", _
              "Saldos Banc�rios", Tag, mlngSldBanco, PB_REGISTRO
    'Pt. 95368 - Moacir Pfau(11/11/2009)
    txtSldBanco(1).MaxLength = 7
    If Len(txtSldBanco(1).Text) = 10 Then
        txtSldBanco(1).Text = Mid(txtSldBanco(1).Text, 4, 10)
    End If
  '
  ' Bot�o Primeiro Registro
  Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
    MoveRecordset mrstSldBanco, Me, Tag, mlngSldBanco, lFuncao
  '
  ' Bot�o Sair
  Case WL_SAIR
    Unload Me
    Exit Function
  '
  ' Bot�o Navegar
  Case WL_NAVEGAR
    Browse mrstSldBanco, Me, Tag, mlngSldBanco, "Saldos Banc�rios"
  '
  ' Bot�o Salvar
  Case WL_SALVAR
    'If SldBancoVerifique() Then LibProc = (SalvaRegistro(mrstSldBanco, Me, Tag, mlngSldBanco) = WL_OK)
    MsgBox "N�o poder� ser realizado altera��o na rotina de Saldos Banc�rios." & vbCrLf & vbCrLf & "Para atualizar os saldos banc�rios favor utilizar os cadastros de " & vbCrLf & "Lan�amentos a Pagar ou Lan�amentos a Receber."
  '
  ' Bot�o Cancelar
  Case WL_CANCELAR
    CancelaEdicao mrstSldBanco, Me, Tag, mlngSldBanco
  '
  ' Op��o Exibir
  Case WL_EXIBIR
    'Pt. 95368 - Moacir Pfau(26/10/2009)
    
    If IsDate(txtSldBanco(1).Text) Then
        strSldBanco = "SELECT * FROM [Saldos Banc�rios] WHERE Banco = {Banco} AND Per�odo = " & IIf(gTipoDB = Access, "#{Per�odo}#;", "'{Per�odo}';")
    Else
        strSldBanco = "SELECT * FROM [Saldos Banc�rios] WHERE Banco = {Banco}"
    End If
    RetornaRegs mrstSldBanco, Me, Tag, strSldBanco, mlngSldBanco
    txtSldBanco(1).Text = Mid(GetValue(mrstSldBanco, "Per�odo"), 4, 10)
  ''
  ' Op��o Filtrar
  Case WL_FILTRAR
    Filtrar mrstSldBanco, Me, Tag, "Saldos Banc�rios", mlngSldBanco
  '
  ' Op��o Bancos
  Case "Bancos"
    If (KeybAcesso(LoadResString(2003))) Then
      frmBancos.Show
      frmBancos.ZOrder vbBringToFront
      CallChange frmBancos.hWnd, txtSldBanco(0).hWnd
    End If
  '
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
    LoadMenuTitulos Me
    ConfigCampos Me, "Saldos Banc�rios", Tag
    AbreRecordset mrstSldBanco, "Saldos Banc�rios"
    DefAddNew mlngSldBanco
    DefineAcesso mlngSldBanco, acSomenteLeitura
    lblSldBcoDesc.Caption = vbNullString
    'Pt. 95368 - Moacir Pfau(11/11/2009)
    txtSldBanco(1).MaxLength = 7
    If Len(txtSldBanco(1).Text) = 10 Then
        txtSldBanco(1).Text = Mid(txtSldBanco(1).Text, 4, 10)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = UnloadForm(mrstSldBanco, Me, Tag, mlngSldBanco)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmSldBancos = Nothing
End Sub


' FUNCTION: SldBancoVerifique
'
' Executa verifica��es de Banco e Data evitando que o usu�rio indique dados
' inconsistentes.
' Retorna: True se estiver tudo correto False se n�o.
' ------------------------------------------------------------------------------
Private Function SldBancoVerifique() As Boolean
  '
  ' Verifica se o Banco existe no cadastro de Bancos
  '
  If Len(lblSldBcoDesc.Caption) = 0 Then
    If MsgBox(ResolveResString(35, resUM, "O Banco " & txtSldBanco(0).Text, _
                               resDOIS, "Bancos"), vbQuestion Or vbYesNo, _
                               MsgBoxCaption) = vbYes Then
      LibProc "Bancos", 0
    End If
    Exit Function
  End If

  '// Verificando se o per�odo � v�lido.

  If (EEdicao(mlngSldBanco)) Then
    Dim sTmp As String

    sTmp = GetValue(mrstSldBanco, "Per�odo", NUL)
    If (Len(sTmp)) Then
      If (MovConferido(sTmp, "KIF")) Then Exit Function
    End If
  End If
  
  If Len(txtSldBanco(1).Text) > 0 Then
    If Not EMesAno(txtSldBanco(1).Text) Then
      MsgBox ResolveResString(26, resUM, txtSldBanco(1).Text), vbInformation, _
             MsgBoxCaption
      Exit Function
    End If

    '// Verifica se o movimento do M�s j� est� conferido

    If calendario.PermiteLancamento(txtSldBanco(1).Text) = "X" Then
      Exit Function
    End If
  End If
  
  SldBancoVerifique = True
  
End Function

Private Sub txtSldBanco_Change(Index As Integer)
Dim strBanco As String

  If Index = 0 Then
    If Len(txtSldBanco(0).Text) > 0 Then
      strBanco = "SELECT Nome FROM Bancos WHERE Banco = " & txtSldBanco(0).Text & ";"
      GetAssocValue strBanco, lblSldBcoDesc
    Else
      lblSldBcoDesc.Caption = vbNullString
    End If
  ElseIf Index = 2 Then
    AlteraValor mlngSldBanco
  End If
  
End Sub

Private Sub txtSldBanco_GotFocus(Index As Integer)
  Selecione txtSldBanco(Index)
  If Index = 0 Then
    MsgBar DescCampo(mrstSldBanco, 0) & ResolveResString(75, resUM, "Bancos")
  Else
    MsgBar DescCampo(mrstSldBanco, txtSldBanco(Index).DataField)
  End If
End Sub

Private Sub txtSldBanco_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index < 2 Then ControlaChave KeyCode, Shift, txtSldBanco(0), mlngSldBanco
  
  If (Index = 0) And (Shift = ZERO) And (KeyCode = vbKeyPageDown) Then
    PCampo "Bancos", "Bancos", PB_CAMPO, txtSldBanco(0), 0
    KeyCode = 0
  End If
  
End Sub

Private Sub txtSldBanco_KeyPress(Index As Integer, KeyAscii As Integer)

  Select Case Index
  '
  ' Campo Banco
  Case 0
    SetMascara KeyAscii, txtSldBanco(0).SelStart, fMask("Bancos", "Banco")
  '
  ' Campo Per�odo
  Case 1
    SetMascara KeyAscii, txtSldBanco(1).SelStart, MASK_MESANO4
  '
  ' Campo Valor
  Case 2
    DMoeda KeyAscii
  '
  End Select
  
End Sub

Private Sub txtSldBanco_LostFocus(Index As Integer)
  
  
  If Index = 0 Or Index = 1 Then LibProc WL_EXIBIR, 0
End Sub
