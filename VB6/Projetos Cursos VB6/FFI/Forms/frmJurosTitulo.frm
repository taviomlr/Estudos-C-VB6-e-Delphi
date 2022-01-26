VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmJurosTitulo 
   Caption         =   "Juros do(s) Título(s)"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSelecionar 
      Caption         =   "Selecionar"
      Height          =   615
      Left            =   30
      TabIndex        =   22
      Top             =   1680
      Width           =   9165
      Begin VB.CommandButton cmdNenhum 
         Caption         =   "&Nenhum"
         Height          =   345
         Left            =   7890
         TabIndex        =   5
         Top             =   180
         Width           =   1125
      End
      Begin VB.CommandButton cmdTodos 
         Caption         =   "&Todos"
         Height          =   345
         Left            =   6720
         TabIndex        =   4
         Top             =   180
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   9210
      TabIndex        =   18
      Top             =   -60
      Width           =   1395
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   1050
         Width           =   1215
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "C&onfirmar"
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Canc&elar"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   1215
      End
      Begin ComctlLib.ImageList imgCheck 
         Left            =   420
         Top             =   3450
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmJurosTitulo.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmJurosTitulo.frx":0352
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraGeral 
      Height          =   1755
      Left            =   30
      TabIndex        =   9
      Top             =   -60
      Width           =   9165
      Begin VB.Frame fraSubMenu 
         Height          =   555
         Left            =   6660
         TabIndex        =   21
         Top             =   1110
         Width           =   2415
         Begin VB.CommandButton cmdCancelaTitulo 
            Caption         =   "&Cancelar"
            Height          =   345
            Left            =   1230
            TabIndex        =   3
            Top             =   150
            Width           =   1125
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Alterar"
            Height          =   345
            Left            =   60
            TabIndex        =   2
            Top             =   150
            Width           =   1125
         End
      End
      Begin Fox.EBSText etxBanco 
         Height          =   330
         Left            =   2460
         TabIndex        =   10
         Top             =   570
         Width           =   795
         _ExtentX        =   1323
         _ExtentY        =   582
         TipoTexto       =   0
         Enabled         =   0   'False
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Banco"
         TipoCriterio    =   3
         CampoDescricao  =   "Nome"
         TabelaConsulta  =   "Bancos"
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
      Begin Fox.EBSText etxDocumento 
         Height          =   330
         Left            =   1140
         TabIndex        =   11
         Top             =   210
         Width           =   2115
         _ExtentX        =   3016
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         Enabled         =   0   'False
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
      Begin Fox.EBSText etxParcela 
         Height          =   330
         Left            =   1140
         TabIndex        =   12
         Top             =   570
         Width           =   525
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         Enabled         =   0   'False
         TipoCriterio    =   0
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
      Begin Fox.EBSText etxTaxaJuros 
         Height          =   330
         Left            =   1140
         TabIndex        =   0
         Top             =   1290
         Width           =   705
         _ExtentX        =   265
         _ExtentY        =   582
         Tipo            =   2
         CasasDecimais   =   2
         TipoTexto       =   0
         MaxLength       =   5
         TipoCriterio    =   5
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
      End
      Begin Fox.EBSText etxJuros 
         Height          =   330
         Left            =   2880
         TabIndex        =   1
         Top             =   1290
         Width           =   1065
         _ExtentX        =   265
         _ExtentY        =   582
         Tipo            =   2
         CasasDecimais   =   2
         TipoTexto       =   0
         MaxLength       =   9
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
      End
      Begin Fox.EBSData edtPagamento 
         Height          =   330
         Left            =   1140
         TabIndex        =   20
         Top             =   930
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         HabilitaCalendario=   0   'False
         Enabled         =   0   'False
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
      Begin VB.Label lblPagamento 
         Alignment       =   1  'Right Justify
         Caption         =   "Pagamento"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblJuros 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor Juros"
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label lblTxJuros 
         Alignment       =   1  'Right Justify
         Caption         =   "Taxa %"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Banco"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lblParcela 
         Alignment       =   1  'Right Justify
         Caption         =   "Parcela"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblDocumento 
         Alignment       =   1  'Right Justify
         Caption         =   "Documento"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   885
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTitulo 
      Height          =   2415
      Left            =   30
      TabIndex        =   6
      Top             =   2310
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmJurosTitulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mcurValorTitulo() As Currency
Private mdatVencimento()  As Date
Private mlngBanco         As Long
'Vinicius Elyseu(24/05/2016) - Projeto: #100340 - Demanda: #120791
Private mlngDocumento()   As Double
Private mintParcela()     As Integer
Private mstrTipo()        As String
Private mstrTabela()      As String
Private mstrCampoNumero() As String
Private mstrPagRec()      As String
Private mcurTaxaJuros()   As Currency
Private mdatPagamento     As Date
Private mstrOrigem()      As String
Private mintIndex         As Integer
Private mcurValorJuros    As Currency
Private Const grdChecked = 2
Private Const grdUnchecked = 1
Private Const colCheck = 1

Public Property Let ValorTitulo(intIndex As Integer, ByVal curValorTitulo As Currency)
    ReDim Preserve mcurValorTitulo(intIndex)
    mcurValorTitulo(intIndex) = curValorTitulo
End Property

Public Property Let Vencimento(intIndex As Integer, ByVal datNovoValor As Date)
    ReDim Preserve mdatVencimento(intIndex)
    mdatVencimento(intIndex) = datNovoValor
End Property

Public Property Let Pagamento(ByVal datNovoValor As Date)
    mdatPagamento = datNovoValor
End Property

Public Property Let documento(intIndex As Integer, ByVal NovoValor As Double)
    ReDim Preserve mlngDocumento(intIndex)
    mlngDocumento(intIndex) = NovoValor
End Property

Public Property Let Parcela(intIndex As Integer, ByVal NovoValor As Double)
    ReDim Preserve mintParcela(intIndex)
    mintParcela(intIndex) = NovoValor
End Property

Public Property Let Tipo(intIndex As Integer, ByVal NovoValor As String)
    ReDim Preserve mstrTipo(intIndex)
    mstrTipo(intIndex) = NovoValor
End Property

Public Property Let Banco(ByVal NovoValor As Long)
    mlngBanco = NovoValor
End Property

Public Property Let PagRec(intIndex As Integer, NovoValor As String)
    ReDim Preserve mstrPagRec(intIndex)
    mstrPagRec(intIndex) = NovoValor
End Property

Public Property Let TaxaJuros(intIndex As Integer, ByVal NovoValor As Currency)
    ReDim Preserve mcurTaxaJuros(intIndex)
    mcurTaxaJuros(intIndex) = NovoValor
End Property

Public Property Let Origem(intIndex As Integer, ByVal NovoValor As String)
    ReDim Preserve mstrOrigem(intIndex)
    mstrOrigem(intIndex) = NovoValor
    If NovoValor = "Dupl" Then
        ReDim Preserve mstrTabela(intIndex)
        mstrTabela(intIndex) = "Duplicatas"
        ReDim Preserve mstrCampoNumero(intIndex)
        mstrCampoNumero(intIndex) = "Nota"
    Else
        ReDim Preserve mstrTabela(intIndex)
        mstrTabela(intIndex) = "Lançamentos"
        ReDim Preserve mstrCampoNumero(intIndex)
        mstrCampoNumero(intIndex) = "Código"
    End If
End Property

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Sub carregaRegistro()
    edtPagamento.Data = mdatPagamento
    etxBanco.valorInteiro = mlngBanco
    cmdAlterar.Enabled = False
    Call PreparaGrid
    Call CarregaRegistrosGrid
End Sub

Private Sub cmdAlterar_Click()
    mcurTaxaJuros(mintIndex) = etxTaxaJuros.valorDecimal
    With grdTitulo
        .Row = mintIndex + 1
        .TextMatrix(.Row, 7) = etxTaxaJuros.valorDecimal
        .TextMatrix(.Row, 9) = FormatCurrency(etxJuros.valorDecimal, 2)
        .TextMatrix(.Row, 10) = FormatCurrency(etxJuros.valorDecimal + mcurValorTitulo(mintIndex), 2)
    End With
    Call LimpaCampos
    cmdConfirmar.SetFocus
    mintIndex = 0
End Sub

Private Sub cmdCancelaTitulo_Click()
    Call LimpaCampos
End Sub

Private Sub cmdConfirmar_Click()
    Dim strSql  As String
    Dim intCont As Integer
    Dim i       As Integer
    'Projeto: 100340 - Desenv.: 146186 - Ueder Budni (14/10/2016)
    Dim objLogLancDup   As clsLogLancamentosDuplicatas
    Dim strEmpresa      As String
    Dim strTipo         As String
    
    intCont = 0
    If ValidaCampos Then
        For i = 1 To grdTitulo.Rows - 1
            grdTitulo.Row = i
            grdTitulo.col = 1
            If grdTitulo.CellPicture = imgCheck.ListImages(grdChecked).Picture And (grdTitulo.TextMatrix(i, 9) > 0) Then
                strSql = "UPDATE " & mstrTabela(intCont) & " SET Acréscimo = " & Replace(CCur(grdTitulo.TextMatrix(i, 9)), ",", ".") & " WHERE " & mstrCampoNumero(intCont) & " = " & mlngDocumento(intCont) & " AND Parcela = " & mintParcela(intCont) & " AND PagRec = '" & mstrPagRec(intCont) & "'"
                If Not ExecuteSQL(strSql) > 0 Then
                    MsgBox "Não foi possível atualizar a tabela de " & mstrTabela(intCont) & ".", vbInformation, NomeModulo
                    Exit Sub
                'Projeto: 100340 - Desenv.: 146186 - Ueder Budni (14/10/2016)
                Else
                    Set objLogLancDup = New clsLogLancamentosDuplicatas
                    With objLogLancDup
                        strEmpresa = GetFieldValue("Empresa", mstrTabela(intCont), mstrCampoNumero(intCont) & " = " & mlngDocumento(intCont) & " AND Parcela = " & mintParcela(intCont) & " AND PagRec = '" & mstrPagRec(intCont) & "'")
                        strTipo = GetFieldValue("Tipo", mstrTabela(intCont), mstrCampoNumero(intCont) & " = " & mlngDocumento(intCont) & " AND Parcela = " & mintParcela(intCont) & " AND PagRec = '" & mstrPagRec(intCont) & "'")
                        
                        Call .SetKey(mstrPagRec(intCont), CDbl(mlngDocumento(intCont)), strEmpresa, strTipo, CLng(mintParcela(intCont)), IIf(mstrTabela(intCont) = "Lançamentos", Lancamento, Duplicata))
                        Call .InsertMsg("Acréscimo de " & grdTitulo.TextMatrix(i, 9) & " (" & grdTitulo.TextMatrix(i, 7) & "%) referente à juros ao baixar título através da rotina de Baixas.")
                    End With
                End If
            End If
            intCont = intCont + 1
        Next
        Set objLogLancDup = Nothing
        Unload Me
    End If
End Sub

Private Sub cmdNenhum_Click()
    Dim i As Long

    If grdTitulo.TextMatrix(1, 2) <> "" Then
        grdTitulo.col = 1
        For i = 1 To grdTitulo.Rows - 1
            grdTitulo.Row = i
            Set grdTitulo.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
        Next
    End If
End Sub

Private Sub cmdTodos_Click()
    Dim i As Long
    
    If grdTitulo.TextMatrix(1, 2) <> "" Then
        grdTitulo.col = 1
        For i = 1 To grdTitulo.Rows - 1
            grdTitulo.Row = i
            Set grdTitulo.CellPicture = imgCheck.ListImages(grdChecked).Picture
        Next
    End If
End Sub

Private Sub etxJuros_LostFocus()
    Dim intDiasAtraso As Integer
    Dim curTaxaJuros As Currency
    
    intDiasAtraso = mdatPagamento - mdatVencimento(mintIndex)
    If etxJuros.valorDecimal > 0 And etxDocumento.valorTexto <> "" Then
        'Se o valor for igual recalcula a taxa sobre o valor antigo por quetões de arredondamento
        If FormatNumber(mcurValorJuros, 2) = etxJuros.valorDecimal Then
            curTaxaJuros = ((mcurValorJuros / intDiasAtraso) / (mcurValorTitulo(mintIndex) / 30)) * 100
        Else
            curTaxaJuros = ((etxJuros.valorDecimal / intDiasAtraso) / (mcurValorTitulo(mintIndex) / 30)) * 100
        End If
        etxTaxaJuros.valorDecimal = curTaxaJuros
    Else
        etxTaxaJuros.valorDecimal = 0
        etxJuros.valorDecimal = 0
    End If
End Sub

Private Sub etxTaxaJuros_LostFocus()
    Dim lngDiasAtraso As Long
    
    If etxTaxaJuros.valorDecimal > 0 And etxDocumento.valorTexto <> "" Then
        If IsValid(edtPagamento.Data) Then
            mdatPagamento = edtPagamento.Data
            lngDiasAtraso = CInt(mdatPagamento - mdatVencimento(mintIndex))
            mcurValorJuros = ((mcurValorTitulo(mintIndex) / 30) * (etxTaxaJuros.valorDecimal / 100)) * lngDiasAtraso
            etxJuros.valorDecimal = FormatNumber(mcurValorJuros, 2)
        End If
    Else
        etxTaxaJuros.valorDecimal = 0
        etxJuros.valorDecimal = 0
    End If
End Sub

Private Sub Form_Load()
    Call etxBanco.AddConexao(Aplicacao)
End Sub

Private Function ValidaCampos() As Boolean
    Dim intIndex As Integer
    
    If grdTitulo.TextMatrix(1, 2) <> "" Then
        With grdTitulo
            For i = 1 To .Rows - 1
                .Row = i
                If Trim(.TextMatrix(i, 9)) = "" Or .TextMatrix(i, 9) = 0 Then
                    If MsgBox("O título " & .TextMatrix(i, 3) & " está marcado porém não foi informado o valor dos Juros. Confirma assim mesmo?", vbQuestion + vbYesNo, NomeModulo) = vbNo Then
                        ValidaCampos = False
                        Exit Function
                    End If
                End If
            Next
        End With
    End If
    ValidaCampos = True
End Function

Private Function CalculaJuros(intIndex As Integer) As Currency
    Dim lngDiasAtraso As Long
    
    If mcurTaxaJuros(intIndex) > 0 Then
        If mdatVencimento(intIndex) < mdatPagamento Then
            lngDiasAtraso = CInt(mdatPagamento - mdatVencimento(intIndex))
        Else
            lngDiasAtraso = 1
        End If
        mcurValorJuros = ((mcurValorTitulo(intIndex) / 30) * (mcurTaxaJuros(intIndex) / 100)) * lngDiasAtraso
        CalculaJuros = FormatCurrency(mcurValorJuros, 2)
    Else
        CalculaJuros = 0
    End If
End Function

Private Sub PreparaGrid()
    Dim intIndex As Integer

    With grdTitulo
        .Cols = 11
        .FixedCols = 1
        .Rows = 2
        
        'Configura a coluna fixa
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 150
        
        'Configura a coluna de seleção
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        
        'Origem
        .TextMatrix(0, 2) = "Origem"
        .ColWidth(2) = 630
        .ColAlignment(2) = flexAlignLeftCenter
                
        'Configura a coluna de Número
        .TextMatrix(0, 3) = "Número"
        .ColWidth(3) = 1500
        .ColAlignment(3) = flexAlignLeftCenter
        
        'Configura a coluna Parcela
        .TextMatrix(0, 4) = "Parc."
        .ColWidth(4) = 450
        .ColAlignment(4) = flexAlignRightCenter
                
        'Configura a coluna de Vencimento
        .TextMatrix(0, 5) = "Vencimento"
        .ColWidth(5) = 1000
        .ColAlignment(5) = flexAlignCenterCenter
        
        'Configura a coluna de Valor
        .TextMatrix(0, 6) = "Valor"
        .ColWidth(6) = 1050
        .ColAlignment(6) = flexAlignRightCenter
        
        'Configura a coluna de Taxa de Juros
        .TextMatrix(0, 7) = "Taxa %"
        .ColWidth(7) = 700
        .ColAlignment(7) = flexAlignRightCenter
        
        'Configura a coluna de Dias em Atraso
        .TextMatrix(0, 8) = "Dias Atraso"
        .ColWidth(8) = 900
        .ColAlignment(8) = flexAlignRightCenter
        
        'Configura a coluna de Valor dos Juros
        .TextMatrix(0, 9) = "Juros"
        .ColWidth(9) = 900
        .ColAlignment(9) = flexAlignRightCenter
        
        'Configura a coluna de Valor com os Juros
        .TextMatrix(0, 10) = "Total"
        .ColWidth(10) = 1200
        .ColAlignment(10) = flexAlignRightCenter
            
        For intIndex = 0 To .Cols - 1
            .TextMatrix(1, intIndex) = ""
        Next
        .col = colCheck
        .Row = 1
        Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
    End With
End Sub

Private Sub CarregaRegistrosGrid()
    Dim intCont As Integer
    Dim i As Integer
    
    i = 1
    For intCont = 0 To UBound(mlngDocumento)
        grdTitulo.AddItem ("")
        grdTitulo.col = 1
        grdTitulo.Row = grdTitulo.Rows - 1
        Set grdTitulo.CellPicture = imgCheck.ListImages(grdChecked).Picture
        
        'Origem
        grdTitulo.TextMatrix(i, 2) = mstrOrigem(intCont)
        grdTitulo.ColAlignment(2) = flexAlignLeftCenter
        'Documento
        grdTitulo.TextMatrix(i, 3) = mlngDocumento(intCont)
        grdTitulo.ColAlignment(3) = flexAlignLeftCenter
        'Parcela
        grdTitulo.TextMatrix(i, 4) = mintParcela(intCont)
        grdTitulo.ColAlignment(4) = flexAlignRightCenter
        'Vencimento
        grdTitulo.TextMatrix(i, 5) = mdatVencimento(intCont)
        grdTitulo.ColAlignment(5) = flexAlignCenterCenter
        'Valor
        grdTitulo.TextMatrix(i, 6) = FormatCurrency(mcurValorTitulo(intCont), 2)
        grdTitulo.ColAlignment(6) = flexAlignRightCenter
        'Taxa de Juros
        grdTitulo.TextMatrix(i, 7) = mcurTaxaJuros(intCont)
        grdTitulo.ColAlignment(7) = flexAlignRightCenter
        'Dias em Atraso
        grdTitulo.TextMatrix(i, 8) = CInt(mdatPagamento - mdatVencimento(intCont))
        grdTitulo.ColAlignment(8) = flexAlignRightCenter
        'Valor dos Juros
        grdTitulo.TextMatrix(i, 9) = FormatCurrency(CalculaJuros(intCont), 2)
        grdTitulo.ColAlignment(9) = flexAlignRightCenter
        'Valor Total
        grdTitulo.TextMatrix(i, 10) = FormatCurrency(grdTitulo.TextMatrix(i, 9) + mcurValorTitulo(intCont), 2)
        grdTitulo.ColAlignment(10) = flexAlignRightCenter
        
        i = i + 1
    Next
    If grdTitulo.Rows > 2 Then
        grdTitulo.RemoveItem (grdTitulo.Rows - 1)
    End If
End Sub

Private Sub grdTitulo_DblClick()
    If grdTitulo.TextMatrix(grdTitulo.Row, 2) <> "" Then
        If Not grdTitulo.col = 1 Then
            With grdTitulo
                cmdAlterar.Enabled = True
                mintIndex = .Row - 1
                etxDocumento.valorTexto = .TextMatrix(.Row, 3)
                etxParcela.valorInteiro = .TextMatrix(.Row, 4)
                etxBanco.valorInteiro = mlngBanco
                edtPagamento.Data = mdatPagamento
                etxTaxaJuros.valorDecimal = .TextMatrix(.Row, 7)
                etxJuros.valorDecimal = .TextMatrix(.Row, 9)
            End With
        End If
    End If
End Sub

Private Sub grdTitulo_Click()
    grdTitulo.col = grdTitulo.ColSel
    grdTitulo.Row = grdTitulo.RowSel
    If grdTitulo.TextMatrix(grdTitulo.Row, 2) <> "" Then
        If grdTitulo.col = 1 Then
            If grdTitulo.CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                Set grdTitulo.CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            Else
                Set grdTitulo.CellPicture = imgCheck.ListImages(grdChecked).Picture
                Call CalculaJuros(grdTitulo.Row - 1)
            End If
        End If
    End If
End Sub

Private Sub LimpaCampos()
    etxDocumento.Clear
    etxParcela.Clear
    etxBanco.valorInteiro = mlngBanco
    edtPagamento.Data = mdatPagamento
    etxTaxaJuros.Clear
    etxJuros.Clear
    cmdAlterar.Enabled = False
    mcurValorJuros = 0
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
