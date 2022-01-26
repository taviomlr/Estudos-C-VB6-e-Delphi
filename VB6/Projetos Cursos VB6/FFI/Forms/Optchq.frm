VERSION 5.00
Begin VB.Form fOptCheque 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opções do Editor"
   ClientHeight    =   4425
   ClientLeft      =   3285
   ClientTop       =   1185
   ClientWidth     =   4335
   Icon            =   "Optchq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCheque 
      Caption         =   "Elementos"
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
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   4095
      Begin VB.CheckBox chkOpt 
         Caption         =   "&Borda nos campos"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraEscala 
      Caption         =   "Escalas"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   2895
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "P&olegadas"
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&Centímetros"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&Milímetros"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame fraOpt 
      Caption         =   "Visualização"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.Frame fraPercent 
         Caption         =   "Porcentagem"
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
         Height          =   855
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtVisualizacao 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            MaxLength       =   5
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&Personalizar"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&50%"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&60%"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&75%"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&85%"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&100%"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&150%"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optVisualizacao 
         Caption         =   "&200%"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Aplicar"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Ok"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   19
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "fOptCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Chaves do arquivo de Inicialização
'
Private Const KEY_VISUALIZACAO$ = "Visualização"
Private Const KEY_BORDAS$ = "Bordas"
Private Const KEY_ESCALA$ = "Escala"

Event AplicarConfig()                  'Evento quando o usuário deseja aplicar as configurações

Private Sub chkOpt_GotFocus(Index As Integer)
  BarMsg 11
End Sub

Private Sub cmdVisualizar_Click(Index As Integer)

  Select Case Index
  '
  'Cancelar: Oculta o Form sem gravar nada
  Case 0
    Me.Hide
    Configurar
  '
  'Ok      : Aplica as alterações e oculta a janela
  Case 1
    RaiseEvent AplicarConfig
    Me.Hide
  '
  'Aplicar : Aplica as alterações sem ocultar
  Case 2
    RaiseEvent AplicarConfig
  End Select
  
End Sub

Private Sub Form_Load()
  CenterForm Me
  Configurar
End Sub

Public Property Get Visual() As Single
  Visual = CSngDef(txtVisualizacao.Text)
End Property

Public Property Get Escala() As ScaleModeConstants
  If (optVisualizacao(8).Value) Then Escala = vbInches
  If (optVisualizacao(9).Value) Then Escala = vbMillimeters
  If (optVisualizacao(10).Value) Then Escala = vbCentimeters
End Property

Public Property Get Bordas() As Boolean
  Bordas = (chkOpt(0).Value = 1)
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Não permite que o usuário feche a janela
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
  Else
    GravaAlteracao
  End If
End Sub

Private Sub optVisualizacao_Click(Index As Integer)
Dim strPorcentagem As String
    
  If Index > -1 And Index < 7 Then
    strPorcentagem = optVisualizacao(Index).Caption
    strPorcentagem = Right$(strPorcentagem, (Len(strPorcentagem) - 1))
    strPorcentagem = Left$(strPorcentagem, (Len(strPorcentagem) - 1))
    txtVisualizacao.Text = strPorcentagem
    fraPercent.Enabled = False
  ElseIf Index = 7 Then
    fraPercent.Enabled = True
  End If
    
End Sub

Private Sub optVisualizacao_GotFocus(Index As Integer)
  BarMsg Index
End Sub

Private Sub txtVisualizacao_GotFocus()
  BarMsg 12
End Sub

Private Sub txtVisualizacao_LostFocus()
  'Verifica a digitação do usuário
  If Not EValido(txtVisualizacao.Text) Then
    MsgFunc ResolveResString(55, resUM, fraPercent.Caption)
    Exit Sub
  Else
    If ((Val(txtVisualizacao.Text) > 200) Or (Val(txtVisualizacao.Text) < 50)) Then
      MsgFunc LoadResString(79)
      Exit Sub
    End If
  End If
End Sub

Private Sub Configurar()
Dim strValores As String
  '
  ' Lendo o arquivo .ini para buscar as configurações anteriores
  ' do editor
  strValores = LerArquivoASCII(SEC_WKIF, KEY_VISUALIZACAO, LocIni("fox.ini"))
  'Verificando se o usuário não editou o arquivo e colocou um valor não válido
  If EValido(strValores) Then
    Select Case Val(strValores)
    Case 200
      optVisualizacao(0).Value = True
    Case 150
      optVisualizacao(1).Value = True
    Case 100
      optVisualizacao(2).Value = True
    Case 85
      optVisualizacao(3).Value = True
    Case 75
      optVisualizacao(4).Value = True
    Case 60
      optVisualizacao(5).Value = True
    Case 50
      optVisualizacao(6).Value = True
    Case Else
      optVisualizacao(7).Value = True
      txtVisualizacao.Text = "100"
    End Select
  Else
    optVisualizacao(2).Value = True
  End If
  '
  ' Elementos
  strValores = LerArquivoASCII(SEC_WKIF, KEY_BORDAS, LocIni("fox.ini"))
  If IsNumeric(strValores) Then
    chkOpt(0).Value = Val(strValores)
  Else
    chkOpt(0).Value = 1
  End If
  '
  ' Escala
  strValores = LerArquivoASCII(SEC_WKIF, KEY_ESCALA, LocIni("fox.ini"))
  If EValido(strValores) Then
    Dim intTemp As Integer
    
    intTemp = CIntDef(strValores)
    If (intTemp < 8) Or (intTemp > 10) Then
      optVisualizacao(8).Value = True
    Else
      optVisualizacao(intTemp).Value = True
    End If
  Else
    optVisualizacao(8).Value = True
  End If
     
End Sub

' Sub GravaAlteracao
'
' Grava as alterações no arquivo fox.ini
' -------------------------------------------
Private Sub GravaAlteracao()
Dim iAlteracao As Integer
  '
  ' Gravando a visualização
  GravarArquivoASCII SEC_WKIF, KEY_VISUALIZACAO, txtVisualizacao.Text, LocIni("fox.ini")
  '
  ' Gravando os elementos
  GravarArquivoASCII SEC_WKIF, KEY_BORDAS, CStr(chkOpt(0).Value), LocIni("fox.ini")
  '
  ' Gravando a escala
  Select Case Escala
  'Polegadas
  Case 5
    iAlteracao = 8
  'Milímetros
  Case 6
    iAlteracao = 9
  'Centímetros
  Case 7
    iAlteracao = 10
  End Select
  
  GravarArquivoASCII SEC_WKIF, KEY_ESCALA, CStr(iAlteracao), LocIni("fox.ini")
    
End Sub

'Sub BarMsg
'
'Imprime mensagens de ajuda na barra de status do Sistema
'Argumentos: [intCtlId]: índice do controle que recebe o foco
'------------------------------------------------------------
Private Sub BarMsg(ByVal intCtlId As Integer)
  
  Select Case intCtlId
  'Caixas de zoom
  Case 0 To 6
    SimpleMsgBar "Opções de Zoom do Editor"
  Case 7
    SimpleMsgBar "Abilita a personalização do zoom"
  Case 8 To 10
    SimpleMsgBar "Escalas de medida do editor"
  Case 11
    SimpleMsgBar "Exibe ou oculta as bordas dos campos"
  Case 12
    SimpleMsgBar "Valor do zoom"
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
