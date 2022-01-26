VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frptRegistrodeDuplicatas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório de Registro de Duplicatas"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "rptRegistrodeDuplicatas.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6795
   Begin VB.Frame fraRegistroDuplicatas 
      Caption         =   "Relatório de Registro de Duplicatas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   6495
      Begin VB.ComboBox cboDuplicatas 
         Height          =   315
         ItemData        =   "rptRegistrodeDuplicatas.frx":0442
         Left            =   1560
         List            =   "rptRegistrodeDuplicatas.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtFil 
         DataField       =   "Empresa"
         Height          =   315
         Index           =   3
         Left            =   1560
         TabIndex        =   13
         Top             =   2520
         Width           =   1695
      End
      Begin VB.ComboBox cboTipoDuplicata 
         Height          =   315
         ItemData        =   "rptRegistrodeDuplicatas.frx":0446
         Left            =   1560
         List            =   "rptRegistrodeDuplicatas.frx":0448
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtFil 
         Height          =   315
         Index           =   1
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtFil 
         DataField       =   "Empresa"
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox chkTotaliza 
         Caption         =   "Total por Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   1695
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtFil 
         Height          =   315
         Index           =   2
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblDesc 
         Height          =   210
         Left            =   3360
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":044A
         TabIndex        =   18
         Top             =   1080
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   210
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":04B6
         TabIndex        =   0
         Top             =   360
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   210
         Index           =   1
         Left            =   120
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":052E
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   215
         Index           =   2
         Left            =   120
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":05A2
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   210
         Index           =   3
         Left            =   120
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":0610
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   210
         Index           =   4
         Left            =   120
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":0692
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDescMoeda 
         Height          =   210
         Left            =   3360
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":0714
         TabIndex        =   19
         Top             =   2520
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   210
         Index           =   5
         Left            =   120
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":078A
         TabIndex        =   12
         Top             =   2520
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   210
         Index           =   6
         Left            =   120
         OleObjectBlob   =   "rptRegistrodeDuplicatas.frx":07F6
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "&Visualizar"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   15
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   16
      Top             =   3420
      Width           =   1215
   End
   Begin VB.CommandButton cmdImp 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   2
      Left            =   5400
      TabIndex        =   17
      Top             =   3420
      Width           =   1215
   End
End
Attribute VB_Name = "frptRegistrodeDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImp_Click(Index As Integer)
    
    If Index < 2 Then
        cmdImp(0).Enabled = False
        cmdImp(1).Enabled = False
        cmdImp(2).Caption = LoadResString(IDS_CANCELAR)
    
        FiltraRegistroDuplicata CBool(Index)
    
        cmdImp(0).Enabled = True
        cmdImp(1).Enabled = True
        cmdImp(2).Caption = LoadResString(IDS_FECHAR)
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    CenterForm Me
    
    lblDesc.Caption = NUL
    LblDescMoeda.Caption = NUL
    
    cboTipo.AddItem "Simples"
    cboTipo.AddItem "Completo"
    cboTipo.Text = "Simples"
    
    cboDuplicatas.AddItem "A Receber"
    cboDuplicatas.AddItem "Recebidas"
    cboDuplicatas.AddItem "A Pagar"
    cboDuplicatas.AddItem "Pagas"
    
    'Tipo de Registro
    '
    ' Carrega as opções de Tipo de Registro
    '
    ComboAddItem cboTipoDuplicata, "SELECT Texto FROM Opções WHERE Rotina = '" & OPT_DUPLICATAS & "';", "Texto"
    cboTipoDuplicata.AddItem "Todos"
    cboTipoDuplicata.Text = "Todos"
    cboDuplicatas.Text = "A Receber"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frptRegistrodeDuplicatas = Nothing
End Sub

Private Sub txtFil_Change(Index As Integer)

    If Index = 0 Then
        GetAssocValue "Select Razão, Apel from empresas where apel = " & Quote(txtFil(Index).Text, "''"), lblDesc, txtFil(Index)
    End If
    
    If Index = 3 Then
      If Len(txtFil(3).Text) Then
        GetAssocValue "SELECT Descrição, Moeda FROM Moedas WHERE Moeda = '" & txtFil(3).Text & "'", _
                      LblDescMoeda, txtFil(3)
      End If
    End If
    
End Sub

Private Sub txtFil_GotFocus(Index As Integer)
    Selecione txtFil(Index)
End Sub

Private Sub txtFil_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If (KeyCode = vbKeyPageDown And Shift = ZERO) Then
      If Index = 0 Then
        PCampo "Empresas", "Empresas", pbCampo, txtFil(Index), "Apel"
      End If
      
      If Index = 3 Then
        PCampo "Moedas e Índices", "Moedas", PB_CAMPO, txtFil(3), "Moeda"
      End If
    End If

      
End Sub

Private Sub txtFil_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        SetMascara KeyAscii, txtFil(Index).SelStart, fMask("Empresas", txtFil(Index).DataField)
    ElseIf Index < 3 Then
        SetMascara KeyAscii, txtFil(Index).SelStart, fMask("Duplicatas", "Vencimento")
    Else
        SetMascara KeyAscii, txtFil(Index).SelStart, fMask("Moedas", "Moeda")
    End If
End Sub

Private Function FiltraRegistroDuplicata(ByVal Destino As Long)
Dim sSql               As String
Dim sWhere             As String
Dim rstTemp            As Object
Dim lRet               As Long
Dim dtIni, dtFim       As Date
Dim dblCotacao         As Double
Dim dtInicial          As Date
Dim dtFinal            As Date

Dim SqlTabNFServico     As String
Dim SqlTabNF            As String

dtInicial = CDateDef(txtFil(1).Text)
dtFinal = CDateDef(txtFil(2).Text)


  dblCotacao = TemCotacao(txtFil(3).Text, LblDescMoeda.Caption, dtInicial, dtFinal)

  'Verifica se a Moeda Informada tem Cotação
  If TemMoeda(txtFil(3).Text, LblDescMoeda.Caption) = False And (Len(txtFil(3).Text) > 0) Then
    MsgBox "Informe uma MOEDA válida para a Conversão de Valores", vbOKOnly Or vbExclamation, MsgBoxCaption
    LetFocus txtFil(3).hWnd
    Selecione txtFil(3)
    Exit Function
  End If
  
  If dblCotacao = 0 And TemMoeda(txtFil(3).Text, LblDescMoeda.Caption) Then
    MsgBox "Não existe Cotação da Moeda " & txtFil(3).Text & " no Dia " & dtInicial & ", Você deve cadastrar uma Cotação para esta data para o Relatório ser " & _
    "Emitido.", vbOKOnly Or vbExclamation, MsgBoxCaption
    LetFocus txtFil(3).hWnd
    Selecione txtFil(3)
    Exit Function
  End If
  
  Dim Tabela As String
  Dim TabelaServico As String
  Dim ehServico As Boolean

  If cboDuplicatas.Text = "A Receber" Or cboDuplicatas.Text = "Recebidas" Then
    TabelaServico = "Notas Fiscais de Serviços a Receber"
    Tabela = "Notas Fiscais de Saída"
  ElseIf cboDuplicatas.Text = "A Pagar" Or cboDuplicatas.Text = "Pagas" Then
    TabelaServico = "Notas Fiscais de Serviços a Pagar"
    Tabela = "Notas Fiscais de Entrada"
  End If
  
  'caso tenha sido informada uma moeda Monta a SELECT com os valores divididos pelo valor da Cotação na data da Emissão da Duplicata
  If (TemMoeda(txtFil(3).Text, LblDescMoeda.Caption)) And (dblCotacao > 0) Then
    SqlTabNF = MontaSqlFiltroDuplicata(Tabela, True)
    SqlTabNFServico = MontaSqlFiltroDuplicata(TabelaServico, True)
  Else
    SqlTabNF = MontaSqlFiltroDuplicata(Tabela, False)
    SqlTabNFServico = MontaSqlFiltroDuplicata(TabelaServico, False)
  End If
    
    '// Começo o filtro
    
    ' Filtra data de emissão da Duplicata
    dtIni = CDateDef(txtFil(1).Text)
    dtFim = CDateDef(txtFil(2).Text)
    
    If (EData(dtIni) And EData(dtFim)) Then
        If dtIni = dtFim Then
            Concat sWhere, " AND Duplicatas.Emissão = " & InverteData(txtFil(1).Text, True)
        Else
            Concat sWhere, " AND Duplicatas.Emissão BETWEEN " & InverteData(txtFil(1).Text, True) & _
                            " AND " & InverteData(txtFil(2).Text, True)
        End If
    ElseIf (EData(dtIni) And Not EData(dtFim)) Then
        Concat sWhere, " AND Duplicatas.Emissão >= " & InverteData(txtFil(1).Text, True)
    ElseIf (Not EData(dtIni) And EData(dtFim)) Then
        Concat sWhere, " AND Duplicatas.Emissão <= " & InverteData(txtFil(2).Text, True)
    End If
    
    ' Filtro a Empresa
    If Len(txtFil(0).Text) > 0 Then
        Concat sWhere, " AND Empresas.Apel = " & Quote(txtFil(0).Text, "''")
    End If
    
    ' Tipo
    If cboTipoDuplicata.Text <> "Todos" Then
        Concat sWhere, " AND Duplicatas.Tipo = " & Quote(cboTipoDuplicata.Text, "''")
        'Concat sWhere, " AND [" & Tabela & "].[Tipo de Registro] = " & Quote(cboTipoDuplicata.Text, "''")
    End If
    
    If cboDuplicatas.Text = "A Receber" Then
        Concat sWhere, " AND Duplicatas.PagRec = 'R'"
        Concat sWhere, " AND Duplicatas.Pagamento IS NULL"
    ElseIf cboDuplicatas.Text = "A Pagar" Then
        Concat sWhere, " AND Duplicatas.PagRec = 'P'"
        Concat sWhere, " AND Duplicatas.Pagamento IS NULL"
    ElseIf cboDuplicatas.Text = "Recebidas" Then
        Concat sWhere, " AND Duplicatas.PagRec = 'R'"
        Concat sWhere, " AND Duplicatas.Pagamento IS NOT NULL"
    ElseIf cboDuplicatas.Text = "Pagas" Then
        Concat sWhere, " AND Duplicatas.PagRec = 'P'"
        Concat sWhere, " AND Duplicatas.Pagamento IS NOT NULL"
    End If
    
    ' Monta a Consulta
    Concat SqlTabNF, sWhere
    Concat SqlTabNFServico, sWhere
    
    'faco o union para juntar as duplicatas
    'referentes as nf de saidas e as nf de servicos
    sSql = SqlTabNF & " UNION " & SqlTabNFServico
    
       
    
    ' Ordena a Consulta
    Concat sSql, " ORDER BY Apel, Tipo, [Número], Parcela"
    
    lRet = AbreRecordset(rstTemp, sSql, dbOpenSnapshot)
    
    If lRet = WL_NORECORD Then
        MsgFunc LoadResString(IDS_NORECORD)
    ElseIf lRet = WL_ERRO Then
        MsgFunc LoadResString(IDS_ERR)
    Else
        '
        ' Executa o relatório
        Call fimpRegistroDuplicatas.Config(rstTemp, cboTipo.Text, Destino, _
                                    txtFil(1).Text & " de " & txtFil(2).Text & " - Duplicatas " & cboDuplicatas.Text, _
                                    IIf(chkTotaliza.value = 1, True, False))
    End If
  
    FechaRecordset rstTemp
End Function


Private Function MontaSqlFiltroDuplicata(Tabela As String, converteMoeda As Boolean) As String
    
    Dim s As String
    
    s = "SELECT Empresas.Apel, "
    s = s & "Empresas.[CNPJ/CPF], "
    s = s & "Empresas.Cidade, "
    
    s = s & "[" & Tabela & "].Número, "
    
    '[Valor Total]
    If converteMoeda Then
        s = s & "[" & Tabela & "].[Valor Total] / "
        s = s & "(SELECT VALOR FROM [COTAÇÕES] "
        s = s & "WHERE MOEDA = '" & txtFil(3).Text & "' "
        s = s & "AND DATA = Duplicatas.Emissão) As [Valor Total], "
    Else
        s = s & "[" & Tabela & "].[Valor Total], "
    End If
    
    s = s & "Duplicatas.Parcela, "
    s = s & "Duplicatas.Tipo, "
    s = s & "Duplicatas.Emissão, "
    s = s & "Duplicatas.Vencimento, "
    s = s & "Duplicatas.Pagamento, "
    
    '[Valor Original]
    If converteMoeda Then
        s = s & "Duplicatas.[Valor Original] / "
        s = s & "(SELECT VALOR FROM COTAÇÕES "
        s = s & "WHERE MOEDA = '" & txtFil(3).Text & "' "
        s = s & "AND DATA = Duplicatas.Emissão) As [Valor Original] "
    Else
        s = s & "Duplicatas.[Valor Original] "
    End If
    
    'FROM
    s = s & "FROM (Duplicatas "
    
    'JOIN Tabela NF
    s = s & "INNER JOIN [" & Tabela & "] "
    s = s & "ON Duplicatas.Nota = [" & Tabela & "].Número "
    s = s & "AND Duplicatas.Tipo = [" & Tabela & "].[Tipo de Registro])"
    
    'JOIN Tabela Empresa
    s = s & "INNER JOIN Empresas "
    s = s & "ON Duplicatas.Empresa = Empresas.Apel "
    
    s = s & "WHERE [" & Tabela & "].Número > 0 "
    
    MontaSqlFiltroDuplicata = s

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
End Sub
