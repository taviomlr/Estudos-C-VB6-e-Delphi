VERSION 5.00
Begin VB.Form frptTransfBanco 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferências Bancárias"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "rptTBco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFinanc 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   29
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinanc 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   28
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinanc 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   27
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame fraFinanc 
      Caption         =   "Transferências Bancárias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   8
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   30
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   7
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   6
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   5
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   24
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   4
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cboFinanc 
         Height          =   315
         ItemData        =   "rptTBco.frx":0C42
         Left            =   1200
         List            =   "rptTBco.frx":0C52
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   3
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   2
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtFinanc 
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Controle:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "lblDescCampo(4)"
         Height          =   195
         Index           =   4
         Left            =   2640
         TabIndex        =   5
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   2520
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Ordem:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   6
         Top             =   2760
         Width           =   510
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "lblDescCampo(3)"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   2
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   2520
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Conta Fina&l:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "lblDescCampo(1)"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   25
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   2520
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "lblDescCampo(0)"
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   22
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   2520
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "&Destino:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   585
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Ori&gem:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   540
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Con&ta Inicial:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   0
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Data Fin&al:"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   18
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Data I&nicial:"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   16
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Código &Final:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   915
      End
      Begin VB.Label lblFinanc 
         AutoSize        =   -1  'True
         Caption         =   "Código &Inicial:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame fraFinanc 
      Caption         =   "Centro de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   3360
      Width           =   5295
      Begin VB.TextBox txtCusto 
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "lblDescCampo(2)"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   3360
      End
      Begin VB.Label lblCusto 
         AutoSize        =   -1  'True
         Caption         =   "&Código:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   540
      End
   End
End
Attribute VB_Name = "frptTransfBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboFinanc_GotFocus()
  TransfBcoMsg cboFinanc.TabIndex
End Sub

Private Sub cmdFinanc_Click(Index As Integer)
  If (Index < 2) Then
    Dim strTransf As String
    
    cmdFinanc(0).Enabled = False
    cmdFinanc(1).Enabled = False
    
    SetPtr vbArrowHourglass
    SimpleMsgBar LoadResString(160)
    If MontaInstrucao(strTransf) Then
      PrintTransf IIf((Index = 0), wrToWindow, wrToPrinter), strTransf
    End If
    SetPtr vbDefault
    MsgBar Caption
    
    cmdFinanc(0).Enabled = True
    cmdFinanc(1).Enabled = True
  Else
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  '
  ' Configurando o formulário:
  '
  CenterForm Me
  cboFinanc.ListIndex = 0
  lblDescCampo(0).Caption = NUL
  lblDescCampo(1).Caption = NUL
  lblDescCampo(2).Caption = NUL
  lblDescCampo(3).Caption = NUL
  lblDescCampo(4).Caption = NUL
  '
  ' Se o usuário não controla Centro de Custo
  '
  If Not CentrodeCusto(MFinanceiro) Then
    fraFinanc(1).Visible = False
    cmdFinanc(0).Top = (cmdFinanc(0).Top - fraFinanc(1).Height)
    cmdFinanc(1).Top = cmdFinanc(0).Top
    cmdFinanc(2).Top = cmdFinanc(0).Top
    Me.Height = (Me.Height - fraFinanc(1).Height)
  End If
  '
  ' Configurando os valores padrão de cada campo
  ' Código da Transferência
  '
  txtFinanc(0).Text = MinValue("Código", "Transf Bancária", NUL)
  txtFinanc(1).Text = MaxValue("Código", "Transf Bancária", NUL)
  '
  ' Data Inicial e Final
  '
  txtFinanc(2).Text = Format$(Date, FDATA)
  txtFinanc(3).Text = Format$(Date, FDATA)
  '
  ' Conta Inicial e Final
  '
  txtFinanc(6).Text = MinValue("Código", "Contas", NUL)
  txtFinanc(7).Text = MaxValue("Código", "Contas", NUL)
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
  MsgBar App.ProductName
  Set frptTransfBanco = Nothing
End Sub


Private Sub txtCusto_Change(Index As Integer)
  Select Case Index
    Case 0
      GetAssocValue "SELECT Descrição FROM Centros WHERE Código = " & txtCusto(0).Text, _
                  lblDescCampo(2)
  End Select
End Sub
Private Sub txtCusto_GotFocus(Index As Integer)
  Selecione txtCusto(0)
  TransfBcoMsg txtCusto(0).TabIndex
End Sub
Private Sub txtCusto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then

        PCampo "Centro de Custo", "Centros", pbCampo, txtCusto(0), 0

  End If
End Sub

Private Sub txtCusto_KeyPress(Index As Integer, KeyAscii As Integer)
  If txtCusto(Index).Index = 0 Then SetMascara KeyAscii, txtCusto(0).SelStart, fMask("Centros", "Código")
End Sub

Private Sub txtFinanc_Change(Index As Integer)
  If ((Index = 4) Or (Index = 5)) Then
    GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtFinanc(Index).Text, _
                  lblDescCampo(Index - 4)
  ElseIf ((Index = 6) Or (Index = 7)) Then
    GetAssocValue "SELECT Descrição FROM Contas WHERE Código = " & txtFinanc(Index).Text, _
                  lblDescCampo(Index - 3)
  End If
End Sub

Private Sub txtFinanc_GotFocus(Index As Integer)
  Selecione txtFinanc(Index)
  TransfBcoMsg txtFinanc(Index).TabIndex
End Sub

Private Sub txtFinanc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
    '
    ' Códigos das Transferências
    '
    If (Index < 2) Then
      PCampo "Transferências Bancárias", "Transf Bancária", pbCampo, txtFinanc(Index), 0
    '
    ' Banco de Origem e Destino
    '
    ElseIf ((Index = 4) Or (Index = 5)) Then
      PCampo "Bancos", "Bancos", pbCampo, txtFinanc(Index), 0
    '
    ' Conta Inicial e Final
    '
    ElseIf ((Index = 6) Or (Index = 7)) Then
      PCampo "Contas", "Contas", pbCampo, txtFinanc(Index), "Código"
    End If
  End If
End Sub

Private Sub txtFinanc_KeyPress(Index As Integer, KeyAscii As Integer)

  Select Case Index
  '
  ' Código das contas
  Case 0
    SetMascara KeyAscii, txtFinanc(0).SelStart, fMask("Transf Bancária", "Código")
  '
  Case 1
    SetMascara KeyAscii, txtFinanc(1).SelStart, fMask("Transf Bancária", "Código"), txtFinanc(0).hWnd
  '
  ' Datas Inicial e Final
  Case 2, 3
    SetMascara KeyAscii, txtFinanc(Index).SelStart, MASK_DATE4
  '
  ' Banco de Origem e Destino
  Case 4, 5
    SetMascara KeyAscii, txtFinanc(Index).SelStart, fMask("Bancos", "Banco")
  '
  ' Conta Inicial e Final
  Case 6
    SetMascara KeyAscii, txtFinanc(6).SelStart, fMask("Contas", "Código")
  '
  Case 7
    SetMascara KeyAscii, txtFinanc(7).SelStart, fMask("Contas", "Código"), txtFinanc(6).hWnd
  '
  End Select
  
End Sub

' SUB.......: TransfBcoMsg
' Objetivo..: Exibe mensagens na barra de Status do Sistema
' Argumento.: [intTabIndex]: Prop. TabIndex do controle.
' --------------------------------------------------------------
Private Sub TransfBcoMsg(intTabIndex As Integer)
  Select Case intTabIndex
  '
  ' Código da Transferência
  Case 2, 4
    MsgBar LoadResString(153) & ResolveResString(75, resUM, "Transferências Bancárias")
  '
  ' Data da Transferência
  Case 6, 8
    MsgBar LoadResString(154)
  '
  ' Código para o Banco de Origem
  Case 10
    MsgBar LoadResString(152) & " de Origem" & ResolveResString(75, resUM, "Bancos")
  '
  ' Código para o Banco de Destino
  Case 13
    MsgBar LoadResString(152) & " de Destino" & ResolveResString(75, resUM, "Banco")
  '
  ' Código da Conta inicial e Final
  Case 16, 19
    MsgBar LoadResString(164) & ResolveResString(75, resUM, "Contas")
  '
  ' Ordem do relatório
  Case 22
    MsgBar LoadResString(155)
  '
  ' Código do Centro de Custo
  Case 25
    MsgBar LoadResString(156)
  '
  End Select
End Sub

' FUNCTION..: MontaInstrucao
' Objetivo..: Monta a instrução Select que será passada ao
'             gerador de relatórios.
' Argumento.: [strResult]: Variável que retornará com a Instrução.
' Retorna...: True se os filtros do usuário estiverem corretos e se
'             a instrução retornar algum registros, caso contrário False.
' -----------------------------------------------------------------------
Private Function MontaInstrucao(strResult As String) As Boolean
Dim lngInicial As Long          'Para os códigos de filtro
Dim lngFinal   As Long
Dim datInicial As Date          'Para as datas de filtro
Dim datFinal   As Date
  
  strResult = "SELECT * FROM [Transf Bancária] WHERE"
  ' Códigos Inicial e Final
  '
  lngInicial = CLngDef(txtFinanc(0).Text, 1)
  AppendStr strResult, " Código >= " & CStr(lngInicial)
  
  lngFinal = CLngDef(txtFinanc(1).Text)
  If lngFinal Then
    Concat strResult, " AND Código <= ", CStr(lngFinal)
  End If
  '
  ' Data inicial e Data final
  '
  
  If IsValid(txtFinanc(2).Text) And IsValid(txtFinanc(3).Text) Then
    If EData(txtFinanc(2).Text) And EData(txtFinanc(3).Text) Then
     If CDateDef(txtFinanc(3).Text) < CDateDef(txtFinanc(2).Text) Then
        MsgFunc "Data Final menor que Data inicial"
        Exit Function
     End If
    End If
  End If
  
  If IsValid(txtFinanc(2).Text) Then
    If EData(txtFinanc(2).Text) Then
      Concat strResult, " AND Data >= ", InverteData(txtFinanc(2).Text, True)
    Else
      Exit Function
    End If
  End If
  
  If IsValid(txtFinanc(3).Text) Then
    If EData(txtFinanc(3).Text) Then
      Concat strResult, " AND Data <= ", InverteData(txtFinanc(3).Text, True)
    Else
      Exit Function
    End If
  End If
  '
  ' Filtrando por Banco Origem
  '
  If IsValid(txtFinanc(4).Text) Then
    Concat strResult, " AND Origem = ", txtFinanc(4).Text
  End If
  '
  ' Filtrando por Banco Destino
  '
  If IsValid(txtFinanc(5).Text) Then
    Concat strResult, " AND Destino = ", txtFinanc(5).Text
  End If
  '
  ' Código de Conta Inicial
  '
  If (IsValid(txtFinanc(6).Text) And IsValid(txtFinanc(7).Text)) Then
    Concat strResult, " AND (Conta BETWEEN ", Min(CLngDef(txtFinanc(6).Text, 1), CLngDef(txtFinanc(7).Text, 1))
    Concat strResult, " AND ", Max(CLngDef(txtFinanc(6).Text, 1), CLngDef(txtFinanc(7).Text, 1)), ")"
  ElseIf ((Not IsValid(txtFinanc(6).Text)) And IsValid(txtFinanc(7).Text)) Then
    Concat strResult, " AND Conta <= ", CLngDef(txtFinanc(7).Text, 1)
  ElseIf (IsValid(txtFinanc(6).Text) And (Not IsValid(txtFinanc(7).Text))) Then
    Concat strResult, " AND Conta >= ", CLngDef(txtFinanc(6).Text, 1)
  End If
  '
  ' Filtrando por Centro de Custo
  '
  If (IsValid(txtCusto(0).Text) And fraFinanc(1).Visible) Then
    Concat strResult, " AND Centro = ", txtCusto(0).Text
  End If
  '
  ' Filtrando por Controle
  '
  If IsValid(txtFinanc(8).Text) Then
    Concat strResult, " AND Controle = ", Quote(txtFinanc(8).Text)
  End If
  '
  ' Ordenando a instrução (Nota: Aqui eu faço (cboFinanc.ListIndex + 1)
  ' na função Choose porque seu primeiro parâmetro tem que estar entre
  ' 1 e a última opção)
  '
  Concat strResult, " ORDER BY ", Choose((cboFinanc.ListIndex + 1), _
                                         "Código", _
                                         "Data", _
                                         "Origem", _
                                         "Destino")
  ' Retorna True
  MontaInstrucao = True
  
End Function

' SUB.......: PrintTransf
' Objetivo..: Imprime o relatório de Transferência Bancária.
' Argumentos: [pdeDestino]: Destino da Impressão.
'             [strTransf] : Instrução Select dos dados.
' ----------------------------------------------------------------
Private Sub PrintTransf(pdeDestino As PrintDestinoEnum, strTransf As String)
Dim rstTransf As Object

  If (AbreRecordsetDAO(rstTransf, strTransf, dbOpenSnapshot) = WL_OK) Then
    Dim wkrTransf As KeybReport
    Dim secTemp   As Secao
    Dim strGrupo  As String     'Guarda o nome do grupo de cabeçalho
    
    strGrupo = "Transferências Bancárias"
    Set wkrTransf = New KeybReport
    With wkrTransf
      Set .DatabaseName = GlobalDataBase
      Set .Recordset = rstTransf
      .WindowTitulo = strGrupo
      .ScaleMode = vbMillimeters
      .AutoRedraw = True
      .Destino = pdeDestino
      
      PageHeader wkrTransf, strGrupo
      '
      ' Se o usuário controla centro de custo
      '
      If (IsValid(txtCusto(0).Text) And fraFinanc(1).Visible) Then
        '
        ' Acrescenta uma linha do grupo de cabeçalho
        '
        If ((CentrodeCusto(MFinanceiro)) And (IsValid(txtCusto(0).Text))) Then
          .Grupo(strGrupo).Header.AddLinha "1"
          With .Grupo(strGrupo).Header.Linha("1")
            .AddCampo , wrCSFixedText
            .Campo(1).FontSize = 9
            .Campo(1).FontStyle = wrFSBold
            .Campo(1).Text = "Centro do Custo" & " " & _
                             StrZero(txtCusto(0).Text, 9) & " - " & _
                             lblDescCampo(2).Caption
            .Campo(1).Alinhamento = wrTACentro
          End With
        End If
      End If
      
      .FontSize = 8
      .AddGrupo "1", , , True
      Set secTemp = .Grupo(1).AddSecao(scHeader, 3, wrDBBottomBorder)

      .FontStyle = wrFSBold
      With secTemp.Linha(2)
        .AddCampo , wrCSFixedText, "Código", wrTADireito, 12
        .AddCampo , wrCSFixedText, "Data", wrTACentro, 15
        .AddCampo , wrCSFixedText, "Descrição", , 35
        If ((CentrodeCusto(MFinanceiro)) And (Not IsValid(txtCusto(0).Text))) Then
          .AddCampo , wrCSFixedText, "C.Custo", wrTADireito, 17
        End If
        .AddCampo , wrCSFixedText, "Banco Origem", , 35
        .AddCampo , wrCSFixedText, "Cheque", wrTACentro, 14
        .AddCampo , wrCSFixedText, "Banco Destino", , 35
        .AddCampo , wrCSFixedText, "Valor", wrTADireito
      End With
      .FontStyle = wrFSNormal
      .AddGrupo "2"
      Set secTemp = .Grupo(2).AddSecao(scDetalhe, 1)
      With secTemp.Linha(1)
        .AddCampo , , "Código", wrTADireito, 12
        .Campo(1).Formato = String$(6, "0")
        .AddCampo , , "Data", wrTACentro, 15
        .Campo(2).Formato = FDDMMYY
        .AddCampo , , "Descrição", , 35
        If ((CentrodeCusto(MFinanceiro)) And (Not IsValid(txtCusto(0).Text))) Then
          .AddCampo "Centro", , "Centro", wrTADireito, 17
          .Campo("Centro").Formato = String$(9, "0")
          .Campo("Centro").SuprimirZeros = True
        End If
        .AddCampo "Nome", wrCSDataLink, "Nome", , 35
        .Campo("Nome").DataLink = "Banco = {Origem}"
        .Campo("Nome").TableLink = "Bancos"
        .AddCampo "Cheque", , "Cheque", wrTACentro, 14
        .Campo("Cheque").Formato = String$(6, "0")
        .Campo("Cheque").SuprimirZeros = True
        .AddCampo "Nome2", wrCSDataLink, "Nome", , 35
        .Campo("Nome2").DataLink = "Banco = {Destino}"
        .Campo("Nome2").TableLink = "Bancos"
        .AddCampo "Valor", , "Valor", wrTADireito
        .Campo("Valor").Formato = FMOEDA
      End With
    With .Grupo(2)
      
      .AddSecao scFooter, 1, wrDBTopBorder
      With .Footer(1)
        .AddCampo , wrCSFixedText, "TOTAL:", wrTAEsquerdo, 25
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSTotal, "Valor", wrTADireito
        .Campo(2).Formato = FMOEDA
        .Campo(2).FontStyle = wrFSBold
      End With
      
    End With
      
      '
      ' Encerrando o relatório
      '
      .AddGrupo "3", wrDBTopBorder, wrVPNoFinal, True
      .Grupo(3).Height = .TextHeight("W")
    End With
    
    SetPtr vbDefault

    wkrTransf.BeginPrint gTipoDB
    wkrTransf.EndPrint

    Set secTemp = Nothing
    Set wkrTransf = Nothing
  Else
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
  End If
  
  FechaRecordset rstTransf
  
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
