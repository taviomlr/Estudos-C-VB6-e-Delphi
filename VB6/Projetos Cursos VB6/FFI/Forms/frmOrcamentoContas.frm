VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmOrcamentoContas 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Orçamento de Contas"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9045
   Begin VB.Frame fraUnico 
      Height          =   5905
      Left            =   40
      TabIndex        =   11
      Top             =   -40
      Width           =   7520
      Begin VB.Frame fraOrcamentos 
         Caption         =   "Orçamentos"
         ForeColor       =   &H80000008&
         Height          =   5295
         Index           =   1
         Left            =   60
         TabIndex        =   12
         Top             =   555
         Width           =   7400
         Begin VB.Frame fraOrdemDados 
            Caption         =   "Ordem dos Dados"
            Height          =   915
            Left            =   5730
            TabIndex        =   13
            Top             =   285
            Width           =   1575
            Begin VB.OptionButton optCrescente 
               Caption         =   "Crescente"
               Height          =   315
               Left            =   100
               TabIndex        =   3
               Top             =   240
               Width           =   1275
            End
            Begin VB.OptionButton optDecrescente 
               Caption         =   "Decrescente"
               Height          =   315
               Left            =   120
               TabIndex        =   4
               Top             =   520
               Width           =   1395
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdOrcamentos 
            Height          =   3825
            Left            =   60
            TabIndex        =   9
            Top             =   1400
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   6747
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin Fox.EBSText etxValor 
            Height          =   330
            Left            =   960
            TabIndex        =   2
            Top             =   960
            Width           =   1890
            _ExtentX        =   4048
            _ExtentY        =   582
            Tipo            =   1
            CasasDecimais   =   2
            PermiteNegativo =   -1  'True
            TipoTexto       =   0
            MaxLength       =   16
            Caption         =   "Valor"
            TipoCriterio    =   6
            Alinhamento     =   1
            Mascara         =   "##,##0.00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxCentroCusto 
            Height          =   330
            Left            =   195
            TabIndex        =   1
            Top             =   600
            Width           =   5335
            _ExtentX        =   436139
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
            Caption         =   "Centro de Custo"
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Centros"
            TamanhoDescricao=   3100
            Alinhamento     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Fox.EBSData edtPeriodo 
            Height          =   330
            Left            =   760
            TabIndex        =   0
            Top             =   240
            Width           =   1500
            _ExtentX        =   65511
            _ExtentY        =   582
            HabilitaCalendario=   0   'False
            Formato         =   1
            Caption         =   "Período"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Fox.EBSText etxConta 
         Height          =   330
         Left            =   980
         TabIndex        =   14
         Top             =   240
         Width           =   5550
         _ExtentX        =   440531
         _ExtentY        =   582
         TipoTexto       =   0
         Caption         =   "Conta"
         Enabled         =   0   'False
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Código"
         TipoCriterio    =   4
         CampoDescricao  =   "Descrição"
         TabelaConsulta  =   "Contas"
         TamanhoDescricao=   4000
         Alinhamento     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5905
      Left            =   7590
      TabIndex        =   10
      Top             =   -40
      Width           =   1410
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   1400
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmOrcamentoContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlngCodConta           As Long
Private mblnAlterando          As Boolean
Private mblnGravando           As Boolean
Private mbolbuscaPelaGrid      As Boolean
Private mbolNavegador          As Boolean
Private mblnLoad               As Boolean
Private mblnGravou             As Boolean
Private Const strTituloGrid$ = "campo=Período;label=Período;tamanho=1000|" & _
                                "campo=Centro;label=Centro;tamanho=1000|" & _
                                "campo=Valor;label=Valor;tamanho=1000"

Public Property Let CodigoConta(ByVal NewVal As Long)
    mlngCodConta = NewVal
End Property

Private Sub cmdExcluir_Click()
    Call LibProc(WL_DELETAR)
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
    If mblnGravou Then
        Call LibProc(WL_NOVO)
        edtPeriodo.SetFocus
    End If
End Sub

Private Sub cmdNovo_Click()
    LibProc (WL_NOVO)
    edtPeriodo.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub edtPeriodo_LostFocus()
    If edtPeriodo.MesAno <> "" Then
        If ExisteRegistro(etxConta.valorInteiro, edtPeriodo.MesAno, etxCentroCusto.valorInteiro) Then
            Call PreencheCampos(etxConta.valorInteiro, edtPeriodo.MesAno, etxCentroCusto.valorInteiro)
            cmdGravar.Enabled = True
        Else
            cmdGravar.Enabled = True
        End If
    End If
End Sub

Private Sub etxCentroCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCentroCusto.ValorDescricao = "" Then
            etxCentroCusto.valorInteiro = 0
        End If
        PCampo "Centro de Custo", "SELECT * FROM Centros", pbCampo, etxCentroCusto, "Código"
    End If
End Sub

Private Sub etxCentroCusto_LostFocus()
    If edtPeriodo.MesAno <> "" Then
        If ExisteRegistro(etxConta.valorInteiro, edtPeriodo.MesAno, etxCentroCusto.valorInteiro) Then
            Call PreencheCampos(etxConta.valorInteiro, edtPeriodo.MesAno, etxCentroCusto.valorInteiro)
        End If
    End If
End Sub

Private Sub Form_Load()
    Call etxCentroCusto.AddConexao(Aplicacao)
    Call etxConta.AddConexao(Aplicacao)
    etxConta.valorInteiro = mlngCodConta
    etxCentroCusto.Enabled = ConfigSys.ControlarCentrodeCusto
    optCrescente.Value = True
    mblnLoad = True
    LibProc (WL_NOVO)
    mblnLoad = False
    Call carregaGrid
End Sub

Public Function LibProc(strFuncao As String, Optional lngFuncao As Long) As Boolean
    Dim strSqlGravar     As String
    Dim strSqlDeletar    As String
    Dim strTipo          As String
    Dim lngCodigo        As Long
    Dim strSinal         As String
    Dim strStatus        As String
    
    Select Case strFuncao
        Case WL_NOVO
            mblnAlterando = False
            Call LimpaTela
            cmdExcluir.Enabled = False
            cmdGravar.Enabled = False
            edtPeriodo.Enabled = True
            etxCentroCusto.Enabled = ConfigSys.ControlarCentrodeCusto
        Case WL_CANCELAR
            If MsgBox("Tem certeza que deseja cancelar?", vbQuestion + vbYesNo, NomeModulo) = vbYes Then
                Call LimpaTela
            End If
        Case WL_SALVAR
            If Not mblnGravando Then
                If ValidaCampos Then
                    mblnGravando = True
                    SetPtr vbHourglass
                    If Not mblnAlterando Then
                        strSqlGravar = "INSERT INTO [Orçamentos de Contas] (Conta,Centro,Período,Valor) VALUES (" & _
                                        etxConta.valorInteiro & "," & etxCentroCusto.valorInteiro & ",#" & _
                                        Format(edtPeriodo.MesAno, "mm/yyyy") & "#," & Replace(etxValor.valorMoeda, ",", ".") & ");"
                    Else
                        strSqlGravar = "UPDATE [Orçamentos de Contas] SET "
                        strSqlGravar = strSqlGravar & " Valor = " & Replace(etxValor.valorMoeda, ",", ".")
                        strSqlGravar = strSqlGravar & " WHERE Conta = " & CLng(etxConta.valorInteiro)
                        strSqlGravar = strSqlGravar & " AND Centro = " & etxCentroCusto.valorInteiro
                        strSqlGravar = strSqlGravar & " AND Período = #" & Format(edtPeriodo.MesAno, "mm/yyyy") & "#;"
                    End If
                    Call ExecuteSQL(strSqlGravar)
                    SetPtr vbDefault
                    mblnGravando = False
                    Call carregaGrid
                    mblnAlterando = True
                    cmdExcluir.Enabled = True
                    mblnGravou = True
                Else
                    mblnGravou = False
                End If
            Else
                MsgBox "Aguarde o termino da gravação antes de gravar novamente.", vbInformation, NomeModulo
            End If
        Case WL_DELETAR
            If mblnAlterando Then
                If MsgBox("Tem certeza que deseja excluir este registro?", vbQuestion + vbYesNo, NomeModulo) = vbYes Then
                    strSqlDeletar = "DELETE FROM [Orçamentos de Contas] WHERE Conta = " & CLng(etxConta.valorInteiro) & " AND Período = #" & Format(edtPeriodo.MesAno, "mm/yyyy") & "# AND Centro = " & etxCentroCusto.valorInteiro & ";"
                    If ExecuteSQL(strSqlDeletar) Then
                        Call LibProc(WL_NOVO)
                        mblnAlterando = False
                        etxConta.SetFocus
                        Call carregaGrid
                    End If
                End If
            Else
                MsgBox "Escolha um Lançamento antes de tentar excluír.", vbInformation, NomeModulo
            End If
            cmdExcluir.Enabled = False
    End Select
End Function

Private Sub LimpaTela()
    edtPeriodo.clear
    etxCentroCusto.clear
    etxValor.clear
    mblnAlterando = False
End Sub

Private Function ValidaCampos() As Boolean
    If edtPeriodo.MesAno = "" Then
        MsgBox "O campo Período deve ser informado.", vbInformation, NomeModulo
        edtPeriodo.SetFocus
        Exit Function
    Else
        ValidaCampos = True
    End If
End Function

Private Sub carregaGrid()
    Dim rsOrcamentoContas    As Object
    Dim sqlOrcamentoContas   As String
    Dim i           As Integer
    Dim strLinha    As String
    
    Call PreparaGrid
    If optDecrescente.Value Then
        sqlOrcamentoContas = "SELECT Centro,Período,Valor FROM [Orçamentos de Contas] WHERE Conta = " & etxConta.valorInteiro & " ORDER BY Período DESC"
    Else
        sqlOrcamentoContas = "SELECT Centro,Período,Valor FROM [Orçamentos de Contas] WHERE Conta = " & etxConta.valorInteiro & " ORDER BY Período ASC"
    End If
    If AbreRecordset(rsOrcamentoContas, sqlOrcamentoContas) = WL_OK Then
        With rsOrcamentoContas
            .MoveFirst
            i = 1
            While Not .EOF
                strLinha = "" & Chr(vbKeyTab) & Format(.Fields("Período").Value, "mm/yyyy") & _
                                Chr(vbKeyTab) & GetFieldValue("Descrição", "Centros", "Código = " & .Fields("Centro").Value) & _
                                Chr(vbKeyTab) & Format(.Fields("Valor").Value, "R$ 0.00#,##")
                grdOrcamentos.AddItem (strLinha)
                .MoveNext
                i = i + 1
            Wend
            If grdOrcamentos.Rows > 2 Then
                If grdOrcamentos.TextMatrix(1, 1) = "" Then
                    grdOrcamentos.RemoveItem (1)
                End If
            End If
        End With
    End If
    Set rsOrcamentoContas = Nothing
End Sub

Private Sub PreparaGrid()
    Dim intIndex As Integer

    With grdOrcamentos
        .Cols = 4
        .FixedCols = 1
        .Rows = 2
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 120
        
        .TextMatrix(0, 1) = "Período"
        .ColWidth(1) = 800
        
        .TextMatrix(0, 2) = "Centro de Custos"
        .ColWidth(2) = 4500
        .ColAlignment(2) = flexAlignLeftCenter
        
        .TextMatrix(0, 3) = "Valor"
        .ColWidth(3) = 1500
        .ColAlignment(3) = flexAlignRightCenter
        
        
        For intIndex = 0 To .Cols - 1
            .TextMatrix(1, intIndex) = ""
        Next
    End With
End Sub

Private Sub grdOrcamentos_DblClick()
    Dim lngCodigo       As Long
    Dim lngCentroCusto  As Long
    Dim dtaPeriodo      As Date
    'Se tem algo no grid entao preenche os campos acima dele
    If grdOrcamentos.TextMatrix(grdOrcamentos.Row, 1) <> Empty Then
        lngCodigo = CLng(etxConta.valorInteiro)
        dtaPeriodo = grdOrcamentos.TextMatrix(grdOrcamentos.Row, 1)
        lngCentroCusto = GetFieldValue("Código", "Centros", "Descrição = '" & grdOrcamentos.TextMatrix(grdOrcamentos.Row, 2) & "'")
        mbolbuscaPelaGrid = True
        Call PreencheCampos(lngCodigo, dtaPeriodo, lngCentroCusto)
        mbolbuscaPelaGrid = False
        etxValor.SetFocus
    End If
End Sub

Private Sub optCrescente_Click()
    Call carregaGrid
End Sub

Private Sub PreencheCampos(lngConta As Long, dtaPeriodo As Date, lngCentroCusto As Long)
    Dim strSqlPreenche As String
    Dim rsOrcamentos As Object
    
    mblnAlterando = True
    strSqlPreenche = "SELECT Centro,Período,Valor FROM [Orçamentos de Contas] WHERE Conta = " & lngConta & " AND Período = #" & Format(dtaPeriodo, "mm/yyyy") & "# AND Centro = " & lngCentroCusto
    If AbreRecordset(rsOrcamentos, strSqlPreenche) = WL_OK Then
        With rsOrcamentos
            edtPeriodo.MesAno = .Fields("Período").Value
            etxCentroCusto.valorInteiro = .Fields("Centro").Value
            etxValor.valorMoeda = .Fields("Valor").Value
            mblnAlterando = True
            cmdGravar.Enabled = True
            cmdExcluir.Enabled = True
            edtPeriodo.Enabled = False
            etxCentroCusto.Enabled = False
        End With
    End If
End Sub

Private Sub optDecrescente_Click()
    Call carregaGrid
End Sub

Private Function ExisteRegistro(lngConta As Long, datPeriodo As Date, lngCentroCusto As Long) As Boolean
    Dim strOpFinanceira As String
    Dim rsOpFinanceira  As Object
    
    If lngConta > 0 Then
        strOpFinanceira = "SELECT * FROM [Orçamentos de Contas] WHERE Conta = " & lngConta & " AND Período = #" & Format(datPeriodo, "mm/dd/yyyy") & "# AND Centro = " & lngCentroCusto
        If AbreRecordset(rsOpFinanceira, strOpFinanceira) = WL_OK Then
            ExisteRegistro = True
        Else
            ExisteRegistro = False
        End If
    Else
        ExisteRegistro = False
    End If
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
