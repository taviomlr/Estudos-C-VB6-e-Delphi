VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmPesquisarLancamentoContaCorrente 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisar Lançamentos de Conta Corrente"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   11025
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   12030
   Begin VB.Frame fraUnico 
      Height          =   4830
      Left            =   40
      TabIndex        =   9
      Top             =   -40
      Width           =   10515
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPesquisaLancamentos 
         Height          =   3150
         Left            =   60
         TabIndex        =   11
         Top             =   1605
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   5556
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin FOX.EBSText etxEstabelecimento 
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   180
         Width           =   9525
         _ExtentX        =   444923
         _ExtentY        =   582
         Tipo            =   4
         Caption         =   "Estabelecimento"
         Enabled         =   0   'False
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Apel"
         CampoDescricao  =   "Razão"
         TabelaConsulta  =   "Empresas"
         TamanhoDescricao=   6500
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
      Begin FOX.EBSData edtLancamentoFin 
         Height          =   330
         Left            =   3060
         TabIndex        =   4
         Top             =   1220
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
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
      Begin FOX.EBSData edtLancamentoIni 
         Height          =   330
         Left            =   1500
         TabIndex        =   3
         Top             =   1220
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
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
      Begin FOX.EBSText etxCliente 
         Height          =   330
         Left            =   930
         TabIndex        =   0
         Top             =   525
         Width           =   8835
         _ExtentX        =   444923
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   15
         Caption         =   "Cliente"
         ValorSelecionado=   -1  'True
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Apel"
         CampoDescricao  =   "Razão"
         TabelaConsulta  =   "Empresas"
         TamanhoDescricao=   6500
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
      Begin FOX.EBSText etxCodigoIni 
         Height          =   330
         Left            =   1500
         TabIndex        =   1
         Top             =   870
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
         ValorSelecionado=   -1  'True
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "id_fficontacorrente"
         TipoCriterio    =   4
         CampoDescricao  =   "id_fficontacorrente"
         TabelaConsulta  =   "FFIContaCorrente"
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
         ExibeDescricao  =   0   'False
      End
      Begin FOX.EBSText etxCodigoFin 
         Height          =   330
         Left            =   3060
         TabIndex        =   2
         Top             =   870
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
         ValorSelecionado=   -1  'True
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "id_fficontacorrente"
         TipoCriterio    =   4
         CampoDescricao  =   "id_fficontacorrente"
         TabelaConsulta  =   "FFIContaCorrente"
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
         ExibeDescricao  =   0   'False
      End
      Begin FOX.EBSText etxDocumento 
         Height          =   330
         Left            =   4920
         TabIndex        =   5
         Top             =   1200
         Width           =   2505
         _ExtentX        =   19764
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   20
         Caption         =   "Documento"
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
      Begin VB.Label lblCodigo 
         Alignment       =   1  'Right Justify
         Caption         =   "Código"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   190
         TabIndex        =   16
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label lblCentroA 
         Caption         =   "à"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   915
         Width           =   135
      End
      Begin VB.Label lblLancamento 
         Alignment       =   1  'Right Justify
         Caption         =   "Dt.Lançamento"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   315
         TabIndex        =   14
         Top             =   1265
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "à"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   1265
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4830
      Left            =   10590
      TabIndex        =   10
      Top             =   -40
      Width           =   1410
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "&Alterar"
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   585
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPesquisarLancamentoContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mblnBuscaPesquisar As Boolean

Private Sub cmdAlterar_Click()
    Call grdPesquisaLancamentos_DblClick
End Sub

Private Sub cmdPesquisar_Click()
    If ValidaCampos Then
        mblnBuscaPesquisar = True
        Call carregaGrid
        mblnBuscaPesquisar = False
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub etxCodigoIni_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyPageDown Then
        If etxCodigoIni.valorInteiro > 0 Then
            etxCodigoIni.valorInteiro = 0
        End If
        PCampo "Lançamentos Conta Corrente", "SELECT * FROM FFIContaCorrente", pbCampo, etxCodigoIni, "id_fficontacorrente"
    End If
End Sub

Private Sub etxCodigoFin_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyPageDown Then
        If etxCodigoFin.valorInteiro > 0 Then
            etxCodigoFin.valorInteiro = 0
        End If
        PCampo "Lançamentos Conta Corrente", "SELECT * FROM FFIContaCorrente", pbCampo, etxCodigoFin, "id_fficontacorrente"
    End If
End Sub


Private Sub Form_Load()
    Call etxEstabelecimento.AddConexao(Aplicacao)
    Call etxCliente.AddConexao(Aplicacao)
    Call etxCodigoIni.AddConexao(Aplicacao)
    Call etxCodigoFin.AddConexao(Aplicacao)
    etxEstabelecimento.valorTexto = DonaSistema
    Call preparaGrid
End Sub

Private Sub etxCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCliente.ValorDescricao = "" Then
            etxCliente.valorTexto = ""
        End If
        PCampo "Empresas", "SELECT * FROM Empresas", pbCampo, etxCliente, "Apel"
    End If
End Sub

Private Sub etxCliente_LostFocus()
    If etxCliente.valorTexto <> "" Then
        etxCliente.valorTexto = GetFieldValue("Apel", "Empresas", "Razão = '" & etxCliente.ValorDescricao & "'")
    End If
End Sub

Private Sub preparaGrid()
    Dim intIndex As Integer

    With grdPesquisaLancamentos
        .Cols = 9
        .FixedCols = 1
        .Rows = 2
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 120
        
        .TextMatrix(0, 1) = "Código"
        .ColWidth(1) = 900
        
        .TextMatrix(0, 2) = "Data"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Documento"
        .ColWidth(3) = 1000
        .ColAlignment(3) = flexAlignLeftCenter
        
        .TextMatrix(0, 4) = "Cliente"
        .ColWidth(4) = 2100
        .ColAlignment(4) = flexAlignLeftCenter
        
        .TextMatrix(0, 5) = "Operação"
        .ColWidth(5) = 800
        .ColAlignment(5) = flexAlignRightCenter
        
        .TextMatrix(0, 6) = "Sinal"
        .ColWidth(6) = 450
        .ColAlignment(6) = flexAlignLeftCenter
        
        .TextMatrix(0, 7) = "Valor"
        .ColWidth(7) = 1400
        .ColAlignment(7) = flexAlignRightCenter
                
        .TextMatrix(0, 8) = "Observação"
        .ColWidth(8) = 2300
        .ColAlignment(8) = flexAlignLeftCenter
        
        For intIndex = 0 To .Cols - 1
            .TextMatrix(1, intIndex) = ""
        Next
    End With
End Sub

Public Sub carregaGrid(Optional strCliente As String)
    Dim rsPesquisa       As Object
    Dim sqlPesquisa      As String
    Dim i                As Integer
    Dim strLinha         As String
    Dim strFiltro        As String
    
    Call preparaGrid
    If strCliente <> "" Then
        strFiltro = BuscaFiltro(strCliente)
    Else
        strFiltro = BuscaFiltro
    End If
    sqlPesquisa = "SELECT id_fficontacorrente, dt_lancamento, nr_documento, apel, cd_op_financeira, vl_lancamento, obs FROM FFIContaCorrente"
    If strFiltro <> "" Then
        sqlPesquisa = sqlPesquisa & " WHERE " & strFiltro
    End If
    If AbreRecordset(rsPesquisa, sqlPesquisa) = WL_OK Then
        With rsPesquisa
            .MoveFirst
            i = 1
            intContRegistros = 0
            While Not .EOF
                strLinha = "" & Chr(vbKeyTab) & .Fields("id_fficontacorrente").Value & _
                                Chr(vbKeyTab) & .Fields("dt_lancamento").Value & _
                                Chr(vbKeyTab) & .Fields("nr_documento").Value & _
                                Chr(vbKeyTab) & .Fields("apel").Value & _
                                Chr(vbKeyTab) & .Fields("cd_op_financeira").Value & _
                                Chr(vbKeyTab) & GetFieldValue("sinal", "FFIOperacaoFinanceira", "cd_op_financeira = " & .Fields("cd_op_financeira").Value) & _
                                Chr(vbKeyTab) & FormatCurrency(.Fields("vl_lancamento").Value, 2) & _
                                Chr(vbKeyTab) & .Fields("obs").Value
                grdPesquisaLancamentos.AddItem (strLinha)
                .MoveNext
                i = i + 1
            Wend
            If grdPesquisaLancamentos.Rows > 2 Then
                If grdPesquisaLancamentos.TextMatrix(1, 1) = "" Then
                    grdPesquisaLancamentos.RemoveItem (1)
                End If
            End If
        End With
    Else
        If mblnBuscaPesquisar Then
            MsgBox "Não há registros para o filtro selecionado.", vbOKOnly, NomeModulo
            etxCliente.SetFocus
        End If
    End If
    Set rsPesquisa = Nothing
End Sub

Private Function BuscaFiltro(Optional strCliente As String) As String
    Dim strFiltro As String
    If strCliente <> "" Then
        strFiltro = " apel = '" & strCliente & "'"
    ElseIf etxCliente.valorTexto <> "" Then
        strFiltro = " apel = '" & etxCliente.valorTexto & "'"
    End If
    If Not IsEmptyDate(edtLancamentoIni.data) Then
        If strFiltro = "" Then
            strFiltro = " dt_lancamento BETWEEN #" & Format(edtLancamentoIni.data, "mm/dd/yyyy") & "# AND #" & Format(edtLancamentoFin.data, "mm/dd/yyyy") & "#"
        Else
            strFiltro = strFiltro & " AND dt_lancamento BETWEEN #" & Format(edtLancamentoIni.data, "mm/dd/yyyy") & "# AND #" & Format(edtLancamentoFin.data, "mm/dd/yyyy") & "#"
        End If
    End If
    If etxDocumento.valorTexto <> "" Then
        If strFiltro = "" Then
            strFiltro = " nr_documento = '" & etxDocumento.valorTexto & "'"
        Else
            strFiltro = strFiltro & " AND nr_documento = '" & etxDocumento.valorTexto & "'"
        End If
    End If
    If etxCodigoIni.valorInteiro > 0 Then
        If strFiltro = "" Then
            strFiltro = " id_fficontacorrente BETWEEN " & etxCodigoIni.valorInteiro & " AND " & etxCodigoFin.valorInteiro
        Else
            strFiltro = strFiltro & " AND id_fficontacorrente BETWEEN " & etxCodigoIni.valorInteiro & " AND " & etxCodigoFin.valorInteiro
        End If
    End If
    BuscaFiltro = strFiltro
End Function

Private Function ValidaCampos() As Boolean
    If (IsEmptyDate(edtLancamentoIni.data) And Not IsEmptyDate(edtLancamentoFin.data)) Or (Not IsEmptyDate(edtLancamentoIni.data) And IsEmptyDate(edtLancamentoFin.data)) Then
        MsgBox "Não é possível exeutar a pesquisa com somente uma das datas preenchida.", vbOKOnly, NomeModulo
        edtLancamentoIni.SetFocus
        Exit Function
    ElseIf edtLancamentoIni.data > edtLancamentoFin.data Then
        MsgBox "A data inicial não pode ser maior que a data final.", vbOKOnly, NomeModulo
        edtLancamentoIni.SetFocus
        Exit Function
    ElseIf (etxCodigoIni.valorInteiro > 0 And etxCodigoFin.valorInteiro = 0) Or (etxCodigoIni.valorInteiro = 0 And etxCodigoFin.valorInteiro > 0) Then
        MsgBox "Não é possível executar a pesquisa com somente um dos códigos preenchido.", vbOKOnly, NomeModulo
        etxCodigoIni.SetFocus
        Exit Function
    Else
        ValidaCampos = True
    End If
End Function

Private Sub grdPesquisaLancamentos_DblClick()
    If IsValid(grdPesquisaLancamentos.TextMatrix(grdPesquisaLancamentos.Row, 1)) Then
        frmLancamentoContaCorrente.Codigo = grdPesquisaLancamentos.TextMatrix(grdPesquisaLancamentos.Row, 1)
        Unload Me
    Else
        MsgBox "Selecione um registro para ser alterado.", vbOKOnly, NomeModulo
    End If
End Sub

Public Property Let Cliente(ByVal NewVal As String)
    etxCliente.valorTexto = NewVal
End Property

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
