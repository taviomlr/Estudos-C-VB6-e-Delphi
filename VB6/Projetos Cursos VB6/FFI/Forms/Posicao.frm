VERSION 5.00
Begin VB.Form fPosiciona 
   KeyPreview      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Posição"
   ClientHeight    =   2640
   ClientLeft      =   4065
   ClientTop       =   1665
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraPosiciona 
      Caption         =   "Campos"
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
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.ComboBox cboPosiciona 
         Height          =   315
         ItemData        =   "Posicao.frx":0000
         Left            =   840
         List            =   "Posicao.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtPosiciona 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPosiciona 
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtPosiciona 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin VB.VScrollBar vsclPosiciona 
         Height          =   315
         Index           =   0
         Left            =   2280
         Max             =   0
         Min             =   32767
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   200
      End
      Begin VB.VScrollBar vsclPosiciona 
         Height          =   315
         Index           =   1
         Left            =   2280
         Max             =   0
         Min             =   32767
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   200
      End
      Begin VB.VScrollBar vsclPosiciona 
         Height          =   315
         Index           =   2
         Left            =   2280
         Max             =   0
         Min             =   32767
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1560
         Width           =   200
      End
      Begin VB.Label lblPosiciona 
         AutoSize        =   -1  'True
         Caption         =   "&Campo:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblPosiciona 
         AutoSize        =   -1  'True
         Caption         =   "&Lateral Esquerda:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label lblPosiciona 
         AutoSize        =   -1  'True
         Caption         =   "&Base:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   405
      End
      Begin VB.Label lblPosiciona 
         AutoSize        =   -1  'True
         Caption         =   "La&rgura:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
      Height          =   285
      Left            =   1275
      TabIndex        =   13
      Top             =   2340
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   2340
      Width           =   1245
   End
End
Attribute VB_Name = "fPosiciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlEscala As ScaleModeConstants
Private msZoom As Single
Private mintAntigaSelecao As Integer

Private mfEdit As frmEditChq        'Referência ao formulário de edição de Cheque
Event AplicarPos(Indice As Integer) 'Evento gerado quando o usuário deseja aplicar as configurações

Public Property Get UserScale() As ScaleModeConstants
  UserScale = mlEscala
End Property

Public Property Let UserScale(ByVal NovaEscala As ScaleModeConstants)
  mlEscala = NovaEscala
End Property

Public Property Get Zoom() As Single
  Zoom = msZoom
End Property

Public Property Let Zoom(ByVal NovoValor As Single)
  msZoom = NovoValor
End Property

Public Property Get Lateral() As Single
  Lateral = CSng(txtPosiciona(0).Text)
End Property

Public Property Get Base() As Single
  Base = CSng(txtPosiciona(1).Text)
End Property

Public Property Get Largura() As Single
  Largura = CSng(txtPosiciona(2).Text)
End Property

' PROPERTY..: EditForm
' Objetivo..: Obtém uma referência ao formulário de Edição de Cheques
' ---------------------------------------------------------------------------
Public Property Set EditForm(frmEdit As frmEditChq)
  Set mfEdit = frmEdit
End Property

Public Property Get EditForm() As frmEditChq
  Set EditForm = mfEdit
End Property

' Sub AtualizaPos
'
' Atualiza a posição exibida no form
' Argumento: [intObjInd]: índice do objeto que deve ser atualizado
' ----------------------------------------------------------------
Public Sub AtualizaPos(ByVal intObjInd As Integer)
  cboPosiciona.Text = cboPosiciona.List(intObjInd)
End Sub

Private Sub cboPosiciona_Click()
Dim triObjetos As KINTRI
Dim intIndObj As Integer
Dim sngTemp As Single
    
  intIndObj = cboPosiciona.ListIndex
  
  If (intIndObj) Then
    '
    'Propriedade left
    sngTemp = (mfEdit.lblCheque(intIndObj).Left - mfEdit.sCheque.Left)
    triObjetos.sLateral = ScaleX((sngTemp / msZoom), vbTwips, mlEscala)
    '
    'Propriedade Width
    sngTemp = mfEdit.lblCheque(intIndObj).Width
    triObjetos.sLargura = ScaleX((sngTemp / msZoom), vbTwips, mlEscala)
    '
    'Valor da Base
    sngTemp = ((mfEdit.lblCheque(intIndObj).Top + mfEdit.lblCheque(intIndObj).Height) - mfEdit.sCheque.Top)
    triObjetos.sBase = ScaleY((sngTemp / msZoom), vbTwips, mlEscala)
    
    lblPosiciona(1).Caption = "&Lateral Esquerda:"
    lblPosiciona(2).Caption = "&Base:"
    
    txtPosiciona(2).Enabled = True
    vsclPosiciona(2).Enabled = True
    
    vsclPosiciona(0).Value = (triObjetos.sLateral * 100)
    vsclPosiciona(1).Value = (triObjetos.sBase * 100)
    vsclPosiciona(2).Value = (triObjetos.sLargura * 100)
    
    mfEdit.DesenhaSelecao intIndObj, -1
  Else
    ' Quando o índice é maior que 8 é a posição do cheque
    sngTemp = mfEdit.sCheque.Width
    triObjetos.sLargura = ScaleX((sngTemp / msZoom), vbTwips, mlEscala)
    
    sngTemp = mfEdit.sCheque.Height
    triObjetos.sBase = ScaleY((sngTemp / msZoom), vbTwips, mlEscala)
    
    lblPosiciona(1).Caption = "&Largura:"
    lblPosiciona(2).Caption = "&Altura:"
    
    txtPosiciona(2).Enabled = False
    vsclPosiciona(2).Enabled = False
    
    vsclPosiciona(0).Value = (triObjetos.sLargura * 100)
    vsclPosiciona(1).Value = (triObjetos.sBase * 100)
    mfEdit.DesenhaSelecao mintAntigaSelecao, 0
  End If
  
End Sub

Private Sub cboPosiciona_DropDown()
  mintAntigaSelecao = cboPosiciona.ListIndex
End Sub

Private Sub cboPosiciona_GotFocus()
  ExibeMsg 0
End Sub

Private Sub cmdAplicar_Click()
  'Aplica as alterações do usuário
  RaiseEvent AplicarPos(cboPosiciona.ListIndex)
End Sub

Private Sub cmdCancelar_Click()
  'Cancela as alterações do usuário
  AtualizaPos cboPosiciona.ListIndex
End Sub

Private Sub Form_Load()
  ' Centraliza o formulário
  CenterForm Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Verifica o tipo de fechamento do form
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_HIDEWINDOW
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mfEdit = Nothing          'Defaz a referência do formulário EditChq
  Set fPosiciona = Nothing
End Sub

Private Sub txtPosiciona_GotFocus(Index As Integer)
Dim intMeuIndice As Integer
    
  Selecione txtPosiciona(Index)
  If lblPosiciona(Index + 1).Caption = "&Altura" Then
    intMeuIndice = 4
  Else
    intMeuIndice = Index + 1
  End If
  ExibeMsg intMeuIndice
    
End Sub

Private Sub txtPosiciona_KeyPress(Index As Integer, KeyAscii As Integer)
  DValor KeyAscii
End Sub

Private Sub txtPosiciona_LostFocus(Index As Integer)
  'Verificando o que o usuário está digitando
  If (Not EValido(txtPosiciona(Index).Text)) Then
    MsgFunc "Este campo não pode conter um valor negativo"
    vsclPosiciona(Index).Value = 5000
  Else
    vsclPosiciona(Index).Value = Int(CSng(txtPosiciona(Index).Text) * 100)
  End If
End Sub

Private Sub vsclPosiciona_Change(Index As Integer)
Dim sngValor As Single
  '
  'Altera o valor da caixa de texto correspondente
  '
  sngValor = (vsclPosiciona(Index).Value / 100)
  txtPosiciona(Index).Text = Format$(sngValor, "Standard")
  
End Sub

' Sub ExibeJanela
'
' Exibe a janela ao usuário tornando-a sempre visível
' ---------------------------------------------------
Public Sub ExibeJanela()
  SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
End Sub

' Sub ExibeMsg
'
' Exibe mensagens de ajuda na barra de Status do Sistema
' Argumento: [intControleId]: índice do controle que recebe o foco
' ----------------------------------------------------------------
Private Sub ExibeMsg(ByVal intControleId As Integer)
  Select Case intControleId
  'Campo
  Case 0
    SimpleMsgBar "Campos editáveis do cheque"
  'Lateral esquerda
  Case 1
    SimpleMsgBar "Posição da borda esquerda do campo"
  'Base
  Case 2
    SimpleMsgBar "Posição da borda inferior do campo"
  'Largura
  Case 3
    SimpleMsgBar "Largura total do campo ou cheque"
  'Altura do cheque
  Case 4
    SimpleMsgBar "Altura total do cheque"
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
