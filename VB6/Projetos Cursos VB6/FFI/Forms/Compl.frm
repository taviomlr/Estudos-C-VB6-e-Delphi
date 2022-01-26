VERSION 5.00
Begin VB.Form fComplementos 
   KeyPreview      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caracteres"
   ClientHeight    =   3975
   ClientLeft      =   4740
   ClientTop       =   1605
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Frame fraCompl 
      Caption         =   "Complementos"
      ClipControls    =   0   'False
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
      Height          =   3555
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CheckBox chkCompl 
         Caption         =   "Valo&r entre parênteses"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox cboCompl 
         Height          =   315
         Index           =   4
         ItemData        =   "Compl.frx":0000
         Left            =   120
         List            =   "Compl.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2520
         Width           =   2055
      End
      Begin VB.ComboBox cboCompl 
         Height          =   315
         Index           =   3
         ItemData        =   "Compl.frx":0004
         Left            =   120
         List            =   "Compl.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox cboCompl 
         Height          =   315
         Index           =   2
         ItemData        =   "Compl.frx":0008
         Left            =   120
         List            =   "Compl.frx":0018
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cboCompl 
         Height          =   315
         Index           =   1
         ItemData        =   "Compl.frx":0054
         Left            =   1320
         List            =   "Compl.frx":007C
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox cboCompl 
         Height          =   315
         Index           =   0
         ItemData        =   "Compl.frx":00B4
         Left            =   1320
         List            =   "Compl.frx":00D3
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkCompl 
         Caption         =   "E&xtenso entre parênteses"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   2145
      End
      Begin VB.Label lblCompl 
         AutoSize        =   -1  'True
         Caption         =   "Formato do &Ano:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label lblCompl 
         AutoSize        =   -1  'True
         Caption         =   "Formato dos &Meses:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1425
      End
      Begin VB.Label lblCompl 
         AutoSize        =   -1  'True
         Caption         =   "&Letras Maiúsculas:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lblCompl 
         AutoSize        =   -1  'True
         Caption         =   "Compl. &Extenso:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lblCompl 
         AutoSize        =   -1  'True
         Caption         =   "Compl. do &Valor:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "fComplementos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const CC_NENHUM$ = "Nenhum"

Event Aplicar()               'Evento gerado quando o usuário deseja aplicar as configurações

Private Sub cboCompl_GotFocus(Index As Integer)
  Call ImprimeMsgStatus(Index)
End Sub

Private Sub cboCompl_KeyPress(Index As Integer, KeyAscii As Integer)
  If (KeyAscii <> vbKeyBack) Then       'BackSpace
    If (Len(cboCompl(Index).Text) > 1) Then
      KeyAscii = 0
      Beep
    End If
  End If
End Sub

Private Sub chkCompl_GotFocus(Index As Integer)
  ImprimeMsgStatus (Index + 5)
End Sub

Private Sub cmdOk_Click()
  RaiseEvent Aplicar
End Sub

Private Sub Form_Load()

  CenterForm Me
  ' Preenche as caixas de combinação que não puderam ser completadas
  ' em tempo de projeto
  With cboCompl(3)
    .AddItem "Completo: (" & MesExt(Date) & ")"
    .AddItem "Abreviado: (" & MesExt(Date, 3) & ")"
  End With
  
  With cboCompl(4)
    .AddItem "Completo: (" & Format$(Date, "yyyy") & ")"
    .AddItem "Abreviado: (" & Format$(Date, "yy") & ")"
  End With
  
End Sub

Public Property Get ComplValor() As String
  If (cboCompl(0).Text <> CC_NENHUM) Then
    ComplValor = cboCompl(0).Text
  Else
    ComplValor = NUL
  End If
End Property

Public Property Let ComplValor(ByVal strValor As String)
Dim intCombo As Integer

  If (Len(strValor)) Then
    cboCompl(0).Text = strValor
    For intCombo = 0 To cboCompl(0).ListCount - 1
      If (cboCompl(0).List(intCombo) = strValor) Then
        cboCompl(0).Text = cboCompl(0).List(intCombo)
        Exit Property
      End If
    Next intCombo
    '
    ' Se a nova opção não existir adiciona a esta seção do programa
    cboCompl(0).AddItem strValor
  Else
    cboCompl(0).Text = CC_NENHUM
  End If
End Property

Public Property Get ComplExt() As String
  If (cboCompl(1).Text = CC_NENHUM) Then
    ComplExt = NUL
  Else
    ComplExt = cboCompl(1).Text
  End If
End Property

Public Property Let ComplExt(ByVal strExtenso As String)
Dim intExt As Integer

  If (Len(strExtenso)) Then
    cboCompl(1).Text = strExtenso
    For intExt = 0 To cboCompl(1).ListCount - 1
      If (cboCompl(1).List(intExt) = strExtenso) Then
        cboCompl(1).Text = cboCompl(1).List(intExt)
        Exit Property
      End If
    Next intExt
    '
    ' Se não houver na lista adiciona nesta seção do programa
    cboCompl(1).AddItem strExtenso
  Else
    cboCompl(1).Text = CC_NENHUM
  End If
  
End Property

Public Property Get CharCase() As KUCase
  CharCase = cboCompl(2).ListIndex
End Property

Public Property Let CharCase(ByVal lngChar As KUCase)
  'Prevenção contra erros de digitação
  If ((lngChar < 0) And (lngChar > 3)) Then
    lngChar = NenhumaLetra
  End If
  '
  ' Definindo na combo
  cboCompl(2).Text = cboCompl(2).List(lngChar)
    
End Property

Public Property Get MesCompleto() As Boolean
  MesCompleto = (cboCompl(3).ListIndex = 0)
End Property

Public Property Let MesCompleto(ByVal blnSim As Boolean)
  If blnSim Then
    cboCompl(3).Text = cboCompl(3).List(0)
  Else
    cboCompl(3).Text = cboCompl(3).List(1)
  End If
End Property

Public Property Get AnoCompleto() As Boolean
  AnoCompleto = (cboCompl(4).ListIndex = 0)
End Property

Public Property Let AnoCompleto(ByVal bSim As Boolean)
  If bSim Then
    cboCompl(4).Text = cboCompl(4).List(0)
  Else
    cboCompl(4).Text = cboCompl(4).List(1)
  End If
End Property

Public Property Get FecharValor() As Boolean
  FecharValor = (chkCompl(0).Value = 1)
End Property

Public Property Let FecharValor(ByVal blnFechar As Boolean)
  chkCompl(0).Value = (blnFechar And vbChecked)
End Property

Public Property Get FecharExt() As Boolean
  FecharExt = (chkCompl(1).Value = vbChecked)
End Property

Public Property Let FecharExt(ByVal bFechar As Boolean)
  If bFechar Then
    chkCompl(1).Value = 1
  Else
    chkCompl(1).Value = 0
  End If
End Property

' PROPERTY..: EditForm
' Objetivo..: Obtém uma referência do formularío de edição de Cheques
'             para trabalhos dentro da janela.
' ---------------------------------------------------------------------
Public Property Set EditForm(frmEdit As Form)

End Property

' SUB.......: Exibe
' Objetivo..: Exibe o form e o mantém sempre em primeiro plano
' -------------------------------------------------------------
Public Sub Exibe()

  SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' Verifica como o formulário está sendo descarregado
  ' vbFormControlMenu = Botão fechar da janela
  If (UnloadMode = vbFormControlMenu) Then
    Cancel = True
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_HIDEWINDOW
  End If
End Sub

' Sub ImprimeMsgStatus
'
' Imprime mensagens de ajuda na barra de status do programa
' Argumento: [intCtlInd]: Índice do controle que recebeu o foco
' --------------------------------------------------------------
Private Sub ImprimeMsgStatus(ByVal intCtlInd As Integer)
    
  Select Case intCtlInd
  'Complemento do valor
  Case 0
    SimpleMsgBar "Caracteres que ficarão ao redor do valor do cheque"
  'Complemento do Extenso
  Case 1
    SimpleMsgBar "Caracteres que completarão as linhas de Extenso"
  'Letras Maiúsculas
  Case 2
    SimpleMsgBar "Define como serão utilizadas as letras maiúsculas para o extenso do cheque"
  'Formato dos Meses
  Case 3
    SimpleMsgBar "Os meses podem ser abreviados, com três letras, ou completos"
  'Formato dos Anos
  Case 4
    SimpleMsgBar "Os anos podem ser abreviados, somente os dois últimos dígitos, ou completos"
  'Valor entre parênteses
  Case 5
    SimpleMsgBar "Acrescentar parênteses ao redor do valor do cheque"
  'Extenso entre parênteses
  Case 6
    SimpleMsgBar "Acrescentar parênteses no início e final do Extenso do cheque"
  End Select
  
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
