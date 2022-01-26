VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmAlteracaoTipoLancamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração Tipo de Serviço"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   11775
   Begin VB.Frame fraBotoes 
      Height          =   5955
      Left            =   10290
      TabIndex        =   13
      Top             =   30
      Width           =   1455
      Begin VB.Frame fraBaixas 
         Caption         =   "Espéc&ie"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1035
         Index           =   1
         Left            =   60
         TabIndex        =   16
         Top             =   5970
         Visible         =   0   'False
         Width           =   1320
         Begin VB.OptionButton optBaixas 
            Caption         =   "À Receber"
            Enabled         =   0   'False
            ForeColor       =   &H80000006&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   18
            Top             =   570
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optBaixas 
            Caption         =   "À Pagar"
            Enabled         =   0   'False
            ForeColor       =   &H80000006&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   17
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Ca&ncelar"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.ImageList imgGrid 
         Left            =   420
         Top             =   4230
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAlteraTipoLancamento.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAlteraTipoLancamento.frx":0352
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraBaixas 
      Caption         =   "Selecionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   1515
      Index           =   5
      Left            =   7230
      TabIndex        =   11
      Top             =   30
      Width           =   3000
      Begin VB.CommandButton cmdNenhum 
         Caption         =   "Nenhum"
         Height          =   345
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1200
      End
      Begin VB.CommandButton cmdTodos 
         Caption         =   "Todos"
         Height          =   345
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1200
      End
   End
   Begin VB.Frame fraAlteracao 
      Caption         =   "Alteração"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   45
      TabIndex        =   10
      Top             =   30
      Width           =   7170
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "&Alterar"
         Height          =   345
         Left            =   5730
         TabIndex        =   3
         Top             =   600
         Width           =   1200
      End
      Begin Fox.EBSText etxTipoServico 
         Height          =   330
         Left            =   1830
         TabIndex        =   0
         Top             =   270
         Width           =   3690
         _extentx        =   11774
         _extenty        =   582
         font            =   "frmAlteraTipoLancamento.frx":06A4
         tipo            =   4
         tipotexto       =   0
         maxlength       =   2
         possuidescricao =   -1  'True
         campocriterio   =   "cd_tipo_servico"
         campodescricao  =   "desc_tipo_servico"
         tabelaconsulta  =   "FFICamaraTipoServico"
         tamanhodescricao=   3000
      End
      Begin Fox.EBSText etxFormaLancamento 
         Height          =   330
         Left            =   1830
         TabIndex        =   1
         Top             =   630
         Width           =   3690
         _extentx        =   11774
         _extenty        =   582
         font            =   "frmAlteraTipoLancamento.frx":06D0
         tipo            =   4
         tipotexto       =   0
         maxlength       =   2
         possuidescricao =   -1  'True
         campocriterio   =   "cd_forma_lancamento"
         campodescricao  =   "desc_forma_lancamento"
         tabelaconsulta  =   "FFICamaraFormaLancamento"
         tamanhodescricao=   3000
      End
      Begin Fox.EBSText etxCodLancamento 
         Height          =   330
         Left            =   1830
         TabIndex        =   2
         Top             =   990
         Width           =   3690
         _extentx        =   11774
         _extenty        =   582
         font            =   "frmAlteraTipoLancamento.frx":06FC
         tipo            =   4
         tipotexto       =   0
         maxlength       =   5
         possuidescricao =   -1  'True
         campocriterio   =   "cd_cod_lancamento"
         campodescricao  =   "desc_cod_lancamento"
         tabelaconsulta  =   "FFICamaraCodigoLancamento"
         tamanhodescricao=   3000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Código de Lançamento"
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   1050
         Width           =   1665
      End
      Begin VB.Label lblFormaLancamento 
         Alignment       =   1  'Right Justify
         Caption         =   "Forma de Lançamento"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   690
         Width           =   1665
      End
      Begin VB.Label lblTipoServico 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Serviço"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   330
         Width           =   1665
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResultado 
      Height          =   4380
      Left            =   30
      TabIndex        =   12
      Top             =   1590
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7726
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmAlteracaoTipoLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const colUnCheck = 1
Private Const colCheck = 2
Private mlngCamara As Long
Private mlngNumeroArquivo As Long
Private mblnConfirmou As Boolean
Private mstrDescricao As String
Private mintQtdLotes As Integer

Public Property Get Camara() As Long
    Camara = mlngCamara
End Property

Public Property Let Camara(ByVal NewVal As Long)
    mlngCamara = NewVal
End Property

Public Property Get NumeroArquivo() As Long
    NumeroArquivo = mlngNumeroArquivo
End Property

Public Property Let NumeroArquivo(ByVal NewVal As Long)
    mlngNumeroArquivo = NewVal
End Property

Public Property Get Descricao() As String
    Descricao = mstrDescricao
End Property

Public Property Let Descricao(ByVal NewVal As String)
    mstrDescricao = NewVal
End Property

Private Sub cmdCancelar_Click()
    Call ConfigureGrid
    Call CarregaRegistrosGrid
End Sub

Private Sub cmdConfirmar_Click()
    Dim strErro As String
    Dim intQtd As Integer
    
    If ValidaRegistros Then
        If AtualizaRegistro(strErro, intQtd) Then
            Load frmGeracaoArquivoRemessaPagamento
            frmGeracaoArquivoRemessaPagamento.CamaraOrigem = mlngCamara
            frmGeracaoArquivoRemessaPagamento.QtdLotes = intQtd
            frmGeracaoArquivoRemessaPagamento.MostraRegistro (mlngNumeroArquivo)
            mblnConfirmou = True
            Unload Me
            Call mostrarForm(frmGeracaoArquivoRemessaPagamento, 2866, False)
        Else
            MsgBox "Não foi possível atualizar os registros." & strErro & ".", vbInformation, NomeModulo
            mblnConfirmou = False
        End If
    Else
        mblnConfirmou = False
    End If
End Sub

Private Sub cmdSair_Click()
    Dim intRow As Integer
    
    Unload Me
    Set frmAlteracaoTipoLancamento = Nothing
    DoEvents
    If Not mblnConfirmou Then
        Call ExecuteSQL("DELETE FROM FFIItemPagamento WHERE cd_arquivoPagamento = " & mlngNumeroArquivo)
        Call ExecuteSQL("DELETE FROM FFIArquivoPagamento WHERE cd_arquivoPagamento = " & mlngNumeroArquivo)
        frmPagamentoDigitalFornecedores.Show
    End If
End Sub

Private Sub etxCodLancamento_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intLinhaSelecionada As Integer
    
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxCodLancamento.ValorDescricao = "" Then
            etxCodLancamento.valorTexto = ""
        End If
        Call PCampo("Código de Lançamento", "SELECT * FROM FFICamaraCodigoLancamento" & IIf(mlngCamara > 0, " WHERE cd_camara = " & mlngCamara & IIf(Trim(etxFormaLancamento.valorTexto) <> Empty, " AND cd_forma_lancamento = " & etxFormaLancamento.valorTexto, ""), ""), pbCampo, etxCodLancamento, "cd_cod_lancamento")
    End If
End Sub

Private Sub Form_Load()
    Call ConfigureGrid
    Call etxTipoServico.AddConexao(Aplicacao)
    Call etxFormaLancamento.AddConexao(Aplicacao)
    Call etxCodLancamento.AddConexao(Aplicacao)
End Sub

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdAlterar_Click()
    Dim intCont As Integer
    
    If etxTipoServico.valorTexto <> "" And etxFormaLancamento.valorTexto <> "" Then
        With grdResultado
            .col = 1
            For intCont = 1 To .Rows - 1
                .Row = intCont
                If .CellPicture = imgGrid.ListImages(colCheck).Picture Then
                    .TextMatrix(.Row, 7) = etxTipoServico.valorTexto
                    .TextMatrix(.Row, 8) = etxFormaLancamento.valorTexto
                    .TextMatrix(.Row, 11) = etxCodLancamento.valorTexto
                End If
            Next
        End With
    ElseIf Trim(etxTipoServico.valorTexto) = "" Then
        MsgBox "O campo Tipo de Serviço do lote não foi preenchido.", vbInformation, NomeModulo
    ElseIf Trim(etxFormaLancamento.valorTexto) = "" Then
        MsgBox "O campo Forma de Lançamento do lote não foi preenchido.", vbInformation, NomeModulo
    Else
        MsgBox "O campo Câmara Centralizadora do lote não foi preenchido.", vbInformation, NomeModulo
    End If
End Sub

Private Sub etxFormaLancamento_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intLinhaSelecionada As Integer
    
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxFormaLancamento.ValorDescricao = "" Then
            etxFormaLancamento.valorTexto = ""
        End If
        Call PCampo("Forma de Lançamento", "SELECT * FROM FFICamaraFormaLancamento" & IIf(mlngCamara > 0, " WHERE cd_camara = " & mlngCamara, ""), pbCampo, etxFormaLancamento, "cd_forma_lancamento")
    End If
End Sub

Private Sub etxTipoServico_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intLinhaSelecionada As Integer
    
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If etxTipoServico.ValorDescricao = "" Then
            etxTipoServico.valorTexto = ""
        End If
        Call PCampo("Tipo de Serviço", "SELECT * FROM FFICamaraTipoServico" & IIf(mlngCamara > 0, " WHERE cd_camara = " & mlngCamara, ""), pbCampo, etxTipoServico, "cd_tipo_servico")
    End If
End Sub

Private Sub cmdNenhum_Click()
    Dim intCont As Integer
    
    With grdResultado
        If .TextMatrix(1, 2) <> "" Then
            .col = 1
            For intCont = 1 To grdResultado.Rows - 1
                .Row = intCont
                Set .CellPicture = imgGrid.ListImages(colUnCheck).Picture
            Next
        End If
    End With
End Sub

Private Sub cmdTodos_Click()
    Dim intCont As Integer
    
    With grdResultado
        If .TextMatrix(1, 2) <> "" Then
            .col = 1
            For intCont = 1 To grdResultado.Rows - 1
                .Row = intCont
                Set .CellPicture = imgGrid.ListImages(colCheck).Picture
            Next
        End If
    End With
End Sub

'Data.......: 10/10/2008
'Autor......: Ivo Sousa
'Descrição..: Procedimento utilizado para listar os dados localizados pelo filtro.
Private Sub ConfigureGrid()
    Dim intColuna As Integer
    
    With grdResultado
        .Rows = 2
        .Cols = 12
        
        'Coluna Fixa
        .ColWidth(0) = 150
        .TextMatrix(0, 0) = ""
        
        'Coluna de seleção
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        .ColAlignment(1) = flexAlignCenterCenter
        
        'Coluna Câmara
        .ColWidth(2) = 600
        .TextMatrix(0, 2) = "Doc."
        .ColAlignment(2) = flexAlignLeftCenter
        
        'Coluna Agência
        .ColWidth(3) = 900
        .TextMatrix(0, 3) = "Número"
        .ColAlignment(3) = flexAlignRightCenter
        
        'Coluna Digito da Agência
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "Tipo"
        .ColAlignment(4) = flexAlignLeftCenter
        
        'Coluna Digito da Agência
        .ColWidth(5) = 650
        .TextMatrix(0, 5) = "Parcela"
        .ColAlignment(5) = flexAlignRightCenter
        
        'Coluna Conta Corrente
        .ColWidth(6) = 2950
        .TextMatrix(0, 6) = "Empresa"
        .ColAlignment(6) = flexAlignLeftCenter
                
        'Coluna Código do Serviço
        .ColWidth(7) = 750
        .TextMatrix(0, 7) = "Serviço"
        .ColAlignment(7) = flexAlignRightCenter
        
        'Coluna Forma de Lançamento
        .ColWidth(8) = 1000
        .TextMatrix(0, 8) = "Lançamento"
        .ColAlignment(8) = flexAlignRightCenter
                
        'Sequência na tabela
        .ColWidth(9) = 0
        .TextMatrix(0, 9) = ""
        
        'Codigo do Lote
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = ""
        
        'Coluna Código de Lançamento
        .ColWidth(11) = 1400
        .TextMatrix(0, 11) = "Cod. Lançamento"
        .ColAlignment(11) = flexAlignRightCenter
        
        For intColuna = 0 To .Cols - 1
            .TextMatrix(1, intColuna) = ""
            If intColuna = 1 Then
                .col = intColuna
                Set .CellPicture = imgGrid.ListImages(colUnCheck).Picture
            End If
        Next
    End With
End Sub

Public Sub CarregaRegistrosGrid()
    Dim intCont   As Integer
    Dim rstResult As Object
    
    If mlngNumeroArquivo > 0 Then
        If AbreRecordset(rstResult, "SELECT * FROM FFIItemPagamento WHERE cd_arquivoPagamento = " & mlngNumeroArquivo) = WL_OK Then
            With rstResult
                intCont = 1
                Set grdResultado.CellPicture = imgGrid.ListImages(colCheck).Picture
                While Not .EOF
                    grdResultado.AddItem ("")
                    grdResultado.Row = grdResultado.Rows - 1
                    grdResultado.col = 1
                    Set grdResultado.CellPicture = imgGrid.ListImages(colCheck).Picture
                    grdResultado.TextMatrix(intCont, 2) = .Fields("tp_documento").value
                    grdResultado.TextMatrix(intCont, 3) = .Fields("nr_documento").value
                    grdResultado.TextMatrix(intCont, 4) = .Fields("tipo_registro").value
                    grdResultado.TextMatrix(intCont, 5) = .Fields("nr_parcela").value
                    grdResultado.TextMatrix(intCont, 6) = .Fields("cd_empresa").value
                    grdResultado.TextMatrix(intCont, 9) = .Fields("cd_itemPagamento").value
                    grdResultado.TextMatrix(intCont, 10) = "0"
                    intCont = intCont + 1
                    .MoveNext
                Wend
                If grdResultado.Rows > 2 Then
                    grdResultado.RemoveItem (grdResultado.Rows - 1)
                End If
            End With
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call cmdSair_Click
End Sub

Private Sub grdResultado_Click()
    With grdResultado
        If .TextMatrix(.Row, 2) <> "" Then
            .CellPictureAlignment = flexAlignCenterCenter
            If .col = 1 Then
                If .CellPicture = imgGrid.ListImages(colUnCheck).Picture Then
                    Set .CellPicture = imgGrid.ListImages(colCheck).Picture
                Else
                    Set .CellPicture = imgGrid.ListImages(colUnCheck).Picture
                End If
            End If
        End If
    End With
End Sub

Private Function ValidaRegistros() As Boolean
    Dim intCont As Integer
    With grdResultado
        For intCont = 1 To .Rows - 1
            If Trim(.TextMatrix(intCont, 7)) = "" Then
                MsgBox "O documento " & .TextMatrix(intCont, 3) & " parcela " & .TextMatrix(intCont, 5) & " está sem as informações referentes ao Código do Serviço.", vbInformation, NomeModulo
                Exit Function
            End If
        Next
    End With
    ValidaRegistros = True
End Function

Private Function AtualizaRegistro(ByRef strErro As String, ByRef intQtd As Integer) As Boolean
    Dim strSql As String
    Dim intCont As Integer
    
On Error GoTo ErroAtualizacao
    With grdResultado
        For intCont = 1 To .Rows - 1
            strSql = "UPDATE FFIItemPagamento SET cd_tipo_servico = " & .TextMatrix(intCont, 7) & ", cd_forma_lancamento = " & .TextMatrix(intCont, 8) & ", cd_cod_lancamento = '" & .TextMatrix(intCont, 11) & "' WHERE " & _
                     "cd_arquivoPagamento = " & mlngNumeroArquivo & " AND cd_itemPagamento = " & .TextMatrix(intCont, 9)
            If Not ExecuteSQL(strSql) > 0 Then
                AtualizaRegistro = False
            End If
        Next
    End With
    Call MontaLotes(intQtd)
    AtualizaRegistro = True
    Exit Function
    
ErroAtualizacao:
    AtualizaRegistro = False
    strErro = err.Description
End Function

Private Sub MontaLotes(ByRef intQtd As Integer)
    Dim rstResult      As Object
    Dim intLote        As Integer
    Dim lngBanco       As Long
    Dim strTipoServico As String
    Dim strFormaLanc   As String
    
    If AbreRecordset(rstResult, "SELECT * FROM FFIItemPagamento WHERE cd_arquivoPagamento = " & CStr(mlngNumeroArquivo) & " ORDER BY cd_banco,cd_tipo_servico,cd_forma_lancamento") = WL_OK Then
        With rstResult
            .MoveFirst
            intLote = 1
            lngBanco = .Fields("cd_banco").value
            strTipoServico = .Fields("cd_tipo_servico").value
            strFormaLanc = .Fields("cd_forma_lancamento").value
            Call ExecuteSQL("INSERT INTO FFIArquivoPagamento (cd_arquivoPagamento,cd_lotePagamento,descricao,nr_camara)VALUES (" & mlngNumeroArquivo & "," & intLote & ",'" & mstrDescricao & "'," & mlngCamara & ")")
            While Not .EOF
                '.Edit
                If Not (lngBanco = .Fields("cd_banco").value And strTipoServico = .Fields("cd_tipo_servico").value And strFormaLanc = .Fields("cd_forma_lancamento").value) Then
                    intLote = intLote + 1
                    Call ExecuteSQL("INSERT INTO FFIArquivoPagamento (cd_arquivoPagamento,cd_lotePagamento,descricao,nr_camara)VALUES (" & mlngNumeroArquivo & "," & intLote & ",'" & mstrDescricao & "'," & mlngCamara & ")")
                End If
                .Fields("cd_lotePagamento").value = intLote
                .update
                .MoveNext
            Wend
            intQtd = intLote
        End With
    End If
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
