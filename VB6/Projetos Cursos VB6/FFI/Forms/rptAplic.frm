VERSION 5.00
Begin VB.Form frptAplicacoes 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicações Financeiras"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "rptAplic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAplic 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   29
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmdAplic 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   28
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmdAplic 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   27
      Top             =   3780
      Width           =   1215
   End
   Begin VB.Frame fraAplic 
      Caption         =   "Aplicações Financeiras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   5295
      Begin VB.TextBox txtAplic 
         Height          =   315
         Index           =   7
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   18
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtAplic 
         Height          =   315
         Index           =   6
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cboAplic 
         Height          =   315
         ItemData        =   "rptAplic.frx":0C42
         Left            =   1320
         List            =   "rptAplic.frx":0C4C
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtAplic 
         Height          =   315
         Index           =   5
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtAplic 
         Height          =   315
         Index           =   4
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtAplic 
         Height          =   315
         Index           =   3
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtAplic 
         Height          =   315
         Index           =   2
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAplic 
         Height          =   315
         Index           =   1
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtAplic 
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDecr 
         Caption         =   "lblDecr(3)"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   19
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "Conta Fina&l:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label lblDecr 
         Caption         =   "lblDecr(2)"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   16
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "Co&nta Inicial:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label lblDecr 
         Caption         =   "lblDecr(1)"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   13
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblDecr 
         Caption         =   "lblDecr(0)"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   10
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "Orde&m:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "Bc&o. Final:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "&Bco. Inicial:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "Data Fin&al:"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "Data I&nicial:"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "Código &Final:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   915
      End
      Begin VB.Label lblAplic 
         AutoSize        =   -1  'True
         Caption         =   "Código &Inicial:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame fraCusto 
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
      Left            =   120
      TabIndex        =   26
      Top             =   2940
      Width           =   5295
      Begin VB.TextBox txtCusto 
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCusto 
         AutoSize        =   -1  'True
         Caption         =   "&Código:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblDescCampo 
         Caption         =   "lblDescCampo(0)"
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   24
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   3360
      End
   End
End
Attribute VB_Name = "frptAplicacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAplic_GotFocus()
  AplicStatusMsg cboAplic.TabIndex
End Sub

Private Sub cmdAplic_Click(Index As Integer)
  If (Index < 2) Then           'Visualizar ou Imprimir
    cmdAplic(0).Enabled = False
    cmdAplic(1).Enabled = False
    FiltroAplic IIf(Index, wrToPrinter, wrToWindow)
    cmdAplic(0).Enabled = True
    cmdAplic(1).Enabled = True
  Else
    Unload Me
  End If
End Sub

Private Sub Form_Load()

  CenterForm Me
  '
  ' Configurando valores padrão de alguns campos
  '
  txtAplic(0).Text = MinValue("Código", "Aplicações", NUL)
  txtAplic(1).Text = MaxValue("Código", "Aplicações", NUL)
  txtAplic(2).Text = Format$(Date, FDATA)
  txtAplic(3).Text = Format$(Date, FDATA)
  txtAplic(4).Text = MinValue("Banco", "Bancos", NUL)
  txtAplic(5).Text = MaxValue("Banco", "Bancos", NUL)
  txtAplic(6).Text = MinValue("Código", "Contas", NUL)
  txtAplic(7).Text = MaxValue("Código", "Contas", NUL)
  lblDescCampo(0).Caption = NUL
  cboAplic.ListIndex = 0
  '
  ' Verificando se o usuário possui Centro de Custo
  '
  If (Not CentrodeCusto(MFinanceiro)) Then
    fraCusto.Visible = False
    cmdAplic(0).Top = (cmdAplic(0).Top - fraCusto.Height)
    cmdAplic(1).Top = cmdAplic(0).Top
    cmdAplic(2).Top = cmdAplic(0).Top
    Me.Height = (Me.Height - fraCusto.Height)
  End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MsgBar App.ProductName
End Sub

Private Sub txtAplic_Change(Index As Integer)
  If ((Index = 4) Or (Index = 5)) Then
    GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtAplic(Index).Text, _
                  lblDecr(Index - 4)
  ElseIf ((Index = 6) Or (Index = 7)) Then
    GetAssocValue "SELECT Descrição FROM Contas WHERE Código = " & txtAplic(Index).Text, _
                  lblDecr(Index - 4)
  End If
End Sub

Private Sub txtAplic_GotFocus(Index As Integer)
  Selecione txtAplic(Index)
  AplicStatusMsg txtAplic(Index).TabIndex
End Sub

Private Sub txtAplic_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
    If (Index < 2) Then
      PCampo "Aplicações", "Aplicações", pbCampo, txtAplic(Index), "Código"
    ElseIf ((Index = 4) Or (Index = 5)) Then
      PCampo "Bancos", "Bancos", pbCampo, txtAplic(Index), "Banco"
    ElseIf ((Index = 6) Or (Index = 7)) Then
      PCampo "Contas", "Contas", pbCampo, txtAplic(Index), "Código"
    End If
  End If
End Sub

Private Sub txtAplic_KeyPress(Index As Integer, KeyAscii As Integer)

  Select Case Index
  '
  ' Código da Aplicação
  Case 0
    SetMascara KeyAscii, txtAplic(ZERO).SelStart, fMask("Aplicações", "Código")
  '
  Case 1
    SetMascara KeyAscii, txtAplic(1).SelStart, fMask("Aplicações", "Código"), txtAplic(0).hWnd
  '
  ' Datas
  Case 2, 3
    SetMascara KeyAscii, txtAplic(Index).SelStart, MASK_DATE4
  '
  ' Código do Banco
  Case 4
    SetMascara KeyAscii, txtAplic(4).SelStart, fMask("Bancos", "Banco")
  '
  Case 5
    SetMascara KeyAscii, txtAplic(5).SelStart, fMask("Bancos", "Banco"), txtAplic(4).hWnd
  '
  ' Código da Conta
  Case 6
    SetMascara KeyAscii, txtAplic(6).SelStart, fMask("Contas", "Código")
  '
  Case 7
    SetMascara KeyAscii, txtAplic(7).SelStart, fMask("Contas", "Código"), txtAplic(6).hWnd
  '
  End Select
  
End Sub

' SUB.......: AplicStatusMsg
' Objetivo..: Exibe mensagens na barra de Status do Sistema
' Argumento.: [iTabIndex]: Propriedade TabIndex do controle
' ---------------------------------------------------------------------------------
Private Sub AplicStatusMsg(iTabIndex As Integer)
  Select Case iTabIndex
  '
  ' Código da Aplicação
  Case 2, 4
    MsgBar LoadResString(157) & ResolveResString(75, resUM, "Aplicações")
  '
  ' Data da Aplicação
  Case 6, 8
    MsgBar LoadResString(158)
  '
  ' Código do Banco
  Case 10, 13
    MsgBar LoadResString(152) & ResolveResString(75, resUM, "Bancos")
  '
  ' Código da Conta
  Case 16, 19
    MsgBar LoadResString(164) & ResolveResString(75, resUM, "Contas")
  '
  ' Ordem
  Case 25
    MsgBar LoadResString(155)
  '
  End Select
End Sub

' SUB.......: FiltroAplic
' Objetivo..: Cria um filtro que cria um arquivo temporário para a impressão do
'             relatório.
' Argumento.: [pdeDestino]: Destino da impressão.
' Retorna...: True se puder cria o arquivo, False se não.
' ---------------------------------------------------------------------------------
Private Sub FiltroAplic(pdeDestino As PrintDestinoEnum)
Dim strSelect As String
Dim rstAplic  As Object
  
  strSelect = "SELECT * FROM Aplicações WHERE"
  '
  ' Código inicial e final da aplicação
  '
  AppendStr strSelect, " Código >= " & CStr(CLngDef(txtAplic(0).Text, 1))
  If IsValid(txtAplic(1).Text) Then
    Concat strSelect, " AND Código <= ", txtAplic(1).Text
  End If
  '
  ' Datas inicial e final
  '
  If IsValid(txtAplic(2).Text) Then
    If EData(txtAplic(2).Text) Then
      Concat strSelect, " AND Data >= ", InverteData(txtAplic(2).Text, True)
    End If
  End If
  
  If IsValid(txtAplic(3).Text) Then
    If EData(txtAplic(3).Text) Then
      Concat strSelect, " AND Data <= ", InverteData(txtAplic(3).Text, True)
    End If
  End If
  
  If IsValid(txtAplic(2).Text) And IsValid(txtAplic(3).Text) Then
    If CDateDef(txtAplic(3).Text) < CDateDef(txtAplic(2).Text) Then
      MsgFunc "Data Final menor que Data Inicial"
      Exit Sub
    End If
  End If
  
  ' Bancos Inicial e Final
  '
  If IsValid(txtAplic(4).Text) Then
    Concat strSelect, " AND Banco >= ", txtAplic(4).Text
  End If
  
  If IsValid(txtAplic(5).Text) Then
    Concat strSelect, " AND Banco <= ", txtAplic(5).Text
  End If
  '
  ' Contas Inicial e Final
  '
  If (IsValid(txtAplic(6).Text)) Then
    Concat strSelect, " AND Conta >= ", txtAplic(6).Text
  End If
  
  If (IsValid(txtAplic(7).Text)) Then
    Concat strSelect, " AND Conta <= ", txtAplic(7).Text
  End If
  '
  ' Centro de Custo, apenas se o usuário possuir
  '
  If (fraCusto.Visible And IsValid(txtCusto(0).Text)) Then
    Concat strSelect, " AND Centro = ", txtCusto(0).Text
  End If
  
  '
  ' Ordem dos dados
  '
  Concat strSelect, " ORDER BY Banco, ", IIf(cboAplic.ListIndex, "Data;", "Código;")
  '
  ' Abrindo o recordset
  '
  If (AbreRecordset(rstAplic, strSelect, dbOpenSnapshot) = WL_OK) Then
    If CriaAplicAux(rstAplic) Then
      ImprimeAplic rstAplic, pdeDestino
    End If
    
    DeleteAux rstAplic, NUL
    
  Else
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
  End If
  
End Sub

' FUNCTION..: CriaAplicAux
' Objetivo..: Cria um arquivo auxiliar para a impressão do relatório.
' Argumento.: [rstSource]: Recordset com os dados de origem, retorna com o recordset
'                          auxiliar com os dados do relatório.
' Retorna...: True se a instrução retorna algum dado e se a tabela auxiliar puder
'             ser criada com sucesso, False se não.
' ----------------------------------------------------------------------------------
Private Function CriaAplicAux(rstSource As Object) As Boolean
Dim fdAux(8) As FieldStruct
Dim rstAux   As Object
Dim strMsg   As String
  
  strMsg = LoadResString(159)
  SetPtr vbArrowHourglass
  SimpleMsgBar strMsg & LoadResString(14)
  
  AppendVar fdAux(0), "Banco", dbLong
  AppendVar fdAux(1), "Código", dbLong
  AppendVar fdAux(2), "Data", dbDate
  AppendVar fdAux(3), "Taxas", dbCurrency
  AppendVar fdAux(4), "Juros", dbCurrency
  AppendVar fdAux(5), "CPMF", dbCurrency
  AppendVar fdAux(6), "Descrição", dbText, 40
  AppendVar fdAux(7), "Centro", dbLong
  AppendVar fdAux(8), "Conta", dbLong
  
  If CrieAux(rstAux, fdAux()) Then
    '
    ' Adicionando os dados a tabela temporária
    '
    rstSource.MoveFirst
    Do
      rstAux.AddNew
      rstAux("Código").Value = rstSource("Código").Value
      rstAux("Banco").Value = rstSource("Banco").Value
      
      SimpleMsgBar strMsg & "...Banco " & CStr(rstSource("Banco").Value)
      
      rstAux("Centro").Value = rstSource("Centro").Value
      rstAux("Conta").Value = rstSource("Conta").Value
      rstAux("Data").Value = rstSource("Data").Value
      rstAux("Descrição").Value = rstSource("Descrição").Value
      Select Case rstSource("Tipo").Value
      '
      Case "Taxas Bancárias"
        rstAux("Taxas").Value = rstSource("Valor").Value
        rstAux("Juros").Value = 0
        rstAux("CPMF").Value = 0
      '
      Case "Juros/Correção"
        rstAux("Taxas").Value = 0
        rstAux("Juros").Value = rstSource("Valor").Value
        rstAux("CPMF").Value = 0
      '
      Case "CPMF"
        rstAux("Taxas").Value = 0
        rstAux("Juros").Value = 0
        rstAux("CPMF").Value = rstSource("Valor").Value
      '
      End Select
      rstAux.update
      rstSource.MoveNext
    Loop Until rstSource.EOF
    FechaRecordset rstSource
    Set rstSource = rstAux
    
    SetPtr vbDefault
    SimpleMsgBar LoadResString(160)
    
    CriaAplicAux = True
  End If
  
End Function

' SUB.......: ImprimeAplic
' Objetivo..: Imprime os dados do relatório de aplicações.
' Argumentos: [rstDados]: Recordset que contém os dados a serem impressos.
'             [pdeDest] : Destino da impressão.
' ---------------------------------------------------------------------------------
Private Sub ImprimeAplic(rstDados As Object, pdeDest As PrintDestinoEnum)
Dim wrkAplic As KeybReport
Dim secAplic As Secao

  Set wrkAplic = New KeybReport
  With wrkAplic
    Set .DatabaseName = GlobalDataBase
    Set .Recordset = rstDados
    .AutoRedraw = True
    .ScaleMode = vbMillimeters
    .WindowTitulo = "Aplicações Financeiras"
    .Tipo = wrObjectDraw
    .Destino = pdeDest
    
    PageHeader wrkAplic, "Aplicações Financeiras"
    
    
    .FontSize = 8
    .AddGrupo "1"
    .Grupo(1).Quebra = "Banco"
    Set secAplic = .Grupo(1).AddSecao(scHeader, 3, wrDBBottomBorder)
    With secAplic.Linha(2)                      'Seção de cabeçalho do grupo
      .AddCampo , wrCSFixedText, "Banco", , 13
      .Campo(1).FontStyle = wrFSBold Or wrFSItalic
      .AddCampo , , "Banco", wrTADireito, 17
      .Campo(2).Formato = StrZero(0, 9)
      .Campo(2).FontStyle = wrFSItalic Or wrFSBold
      .AddCampo , wrCSDataLink, "Nome"
      .Campo(3).FontStyle = wrFSBold Or wrFSItalic
      .Campo(3).TableLink = "Bancos"
      .Campo(3).DataLink = "Banco = {Banco}"
    End With
    
    .FontStyle = wrFSBold
    
    With .Grupo(1).Header.Linha(3)
      .AddCampo , wrCSFixedText, "Código", wrTADireito, 12
      .AddCampo , wrCSFixedText, "Data", , 17, 13
      .AddCampo , wrCSFixedText, "Taxas", wrTADireito, 25
      .AddCampo , wrCSFixedText, "Juros", wrTADireito, 25
      .AddCampo , wrCSFixedText, "CPMF", wrTADireito, 25
      .AddCampo , wrCSFixedText, "Conta", wrTADireito, 15
      If (CentrodeCusto(MFinanceiro)) Then
        .AddCampo , wrCSFixedText, "C.Custo", wrTADireito, 15
      End If
      .AddCampo , wrCSFixedText, "Descrição"
    End With
    
    .FontStyle = wrFSNormal
    .Grupo(1).AddSecao scDetalhe, 1
    With .Grupo(1).Detalhe.Linha(1)           'Seção que imprime os dados
      .AddCampo , , "Código", wrTADireito, 12
      .AddCampo , , "Data", , 17, 13
      .Campo(2).Formato = FDATA
      .AddCampo , , "Taxas", wrTADireito, 25
      .Campo(3).Formato = FMOEDA
      .AddCampo , , "Juros", wrTADireito, 25
      .Campo(4).Formato = FMOEDA
      .AddCampo , , "CPMF", wrTADireito, 25
      .Campo(5).Formato = FMOEDA
      .AddCampo , , "Conta", wrTADireito, 15
      If (CentrodeCusto(MFinanceiro)) Then
        .AddCampo , , "Centro", wrTADireito, 15
        .Campo(7).SuprimirZeros = True
      End If
      .AddCampo , , "Descrição"
    End With
    
    .Grupo(1).AddSecao scFooter, 1, wrDBBottomBorder
    With .Grupo(1).Footer.Linha(1)                  'Rodapé do grupo
      .DrawBorder = wrDBTopBorder
      .BorderStyle = wrDot
      .AddCampo , wrCSFixedText, "Totais:", wrTADireito, 30
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSSubTotal, "Taxas", wrTADireito, 25
      .Campo(2).Left = wrkAplic.Grupo(1).Detalhe(1).Campo(3).Left
      .Campo(2).Formato = FMOEDA
      .AddCampo , wrCSSubTotal, "Juros", wrTADireito, 25
      .Campo(3).Formato = FMOEDA
      .AddCampo , wrCSSubTotal, "CPMF", wrTADireito, 25
      .Campo(4).Formato = FMOEDA
    End With
  End With
  wrkAplic.BeginPrint gTipoDB
  wrkAplic.EndPrint
  
  Set secAplic = Nothing
  Set wrkAplic = Nothing
  
  MsgBar Caption
  
End Sub

Private Sub txtCusto_Change(Index As Integer)
Select Case Index
  Case 0
    GetAssocValue "SELECT Descrição FROM Centros WHERE Código = " & _
                  txtCusto(0).Text & ";", lblDescCampo(0)
End Select
End Sub

Private Sub txtCusto_GotFocus(Index As Integer)
Select Case Index
  Case 0
    Selecione txtCusto(0)
    MsgBar LoadResString(156) & ResolveResString(75, resUM, "Centro de Custo")
End Select
End Sub


Private Sub txtCusto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
  
    Select Case Index
      Case 0
        PCampo "Centros de Custo", "Centros", pbCampo, txtCusto, "Código"
    End Select
    
  End If
    
  End Sub

Private Sub txtCusto_KeyPress(Index As Integer, KeyAscii As Integer)
  If txtCusto(Index).Index = 0 Then SetMascara KeyAscii, txtCusto(0).SelStart, fMask("Centros", "Código")  'Centro de Custo
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
