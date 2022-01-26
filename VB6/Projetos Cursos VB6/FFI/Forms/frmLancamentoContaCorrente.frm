VERSION 5.00
Begin VB.Form frmLancamentoContaCorrente 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de Conta Corrente"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   10050
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   8610
      TabIndex        =   15
      Top             =   -40
      Width           =   1415
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   1400
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   40
      TabIndex        =   12
      Top             =   -40
      Width           =   8535
      Begin Fox.EBSText etxEmpUser 
         Height          =   330
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7695
         _ExtentX        =   442278
         _ExtentY        =   582
         Tipo            =   4
         Caption         =   "Estabelecimento"
         Enabled         =   0   'False
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Apel"
         CampoDescricao  =   "Razão"
         TabelaConsulta  =   "Empresas"
         TamanhoDescricao=   5000
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
      Begin Fox.EBSText etxCliente 
         Height          =   330
         Left            =   810
         TabIndex        =   0
         Top             =   960
         Width           =   7110
         _ExtentX        =   442463
         _ExtentY        =   582
         Tipo            =   4
         MaxLength       =   15
         Caption         =   "Cliente"
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Apel"
         CampoDescricao  =   "Razão"
         TabelaConsulta  =   "Empresas"
         TamanhoDescricao=   5100
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
      Begin Fox.EBSData edtDataLote 
         Height          =   330
         Left            =   1380
         TabIndex        =   1
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
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
      Begin Fox.EBSText etxDocumento 
         Height          =   330
         Left            =   3000
         TabIndex        =   2
         Top             =   1320
         Width           =   2475
         _ExtentX        =   134805
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
         ExibeDescricao  =   0   'False
      End
      Begin Fox.EBSText etxSaldo 
         Height          =   330
         Left            =   6000
         TabIndex        =   3
         Top             =   1320
         Width           =   2085
         _ExtentX        =   71146
         _ExtentY        =   582
         Tipo            =   1
         CasasDecimais   =   2
         TipoTexto       =   0
         Caption         =   "Saldo"
         TipoCriterio    =   6
         Alinhamento     =   1
         Mascara         =   "##,##0.00"
         Locked          =   -1  'True
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
      Begin Fox.EBSText etxOperacao 
         Height          =   330
         Left            =   585
         TabIndex        =   4
         Top             =   1680
         Width           =   5955
         _ExtentX        =   440955
         _ExtentY        =   582
         MaxLength       =   5
         Caption         =   "Operação"
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "cd_op_financeira"
         TipoCriterio    =   4
         CampoDescricao  =   "descricao"
         TabelaConsulta  =   "FFIOperacaoFinanceira"
         TamanhoDescricao=   4250
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
      Begin Fox.EBSText etxValor 
         Height          =   330
         Left            =   930
         TabIndex        =   5
         Top             =   2040
         Width           =   1830
         _ExtentX        =   62521
         _ExtentY        =   582
         Tipo            =   1
         CasasDecimais   =   2
         TipoTexto       =   0
         MaxLength       =   15
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
      Begin Fox.EBSText etxDescricao 
         Height          =   330
         Left            =   570
         TabIndex        =   6
         Top             =   2400
         Width           =   7800
         _ExtentX        =   125756
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   250
         Caption         =   "Descrição"
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
      Begin Fox.EBSText etxCodigo 
         Height          =   330
         Left            =   800
         TabIndex        =   18
         Top             =   600
         Width           =   1365
         _ExtentX        =   29078
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   20
         Caption         =   "Código"
         Enabled         =   0   'False
         TipoCriterio    =   4
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
      Begin VB.Label lblSiglaSinal 
         Caption         =   "Sinal"
         Height          =   255
         Left            =   6840
         TabIndex        =   17
         Top             =   1725
         Width           =   375
      End
      Begin VB.Label lblSinal 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7320
         TabIndex        =   16
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblDtLote 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Lote"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   195
         TabIndex        =   13
         Top             =   1380
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLancamentoContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mlngIdFFIContaCorrente As Long
Private mlngEnterpriseId       As Long
Private mlngCdEstabelecimento  As Long
Private mrsContaCorrente       As Object
Private mblnAlterando          As Boolean
Private mblnGravando           As Boolean
Private mbolbuscaPelaGrid      As Boolean
Private mbolNavegador          As Boolean
Private mblnLoad               As Boolean
Private mblnGravou             As Boolean
Private mlngCodigo             As Long
Private lWnd                   As Long

Private Sub cmdExcluir_Click()
    LibProc (WL_DELETAR)
End Sub

Private Sub cmdNovo_Click()
    LibProc (WL_NOVO)
End Sub

Private Sub cmdPesquisar_Click()
    Load frmPesquisarLancamentoContaCorrente
    If etxCliente.valorTexto <> "" Then
        frmPesquisarLancamentoContaCorrente.Cliente = etxCliente.valorTexto
        Call frmPesquisarLancamentoContaCorrente.carregaGrid(etxCliente.valorTexto)
    End If
    Call mostrarForm(frmPesquisarLancamentoContaCorrente, 2800, True)
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
    If mblnGravou Then
        Call CarregaRecordset
        LibProc (WL_NOVO)
    End If
End Sub

Public Function LibProc(strFuncao As String, Optional lngFuncao As Long) As Boolean
    Dim strSqlGravar          As String
    Dim strSqlDeletar         As String
    Dim strTipo               As String
    Dim lngCodigo             As Long
    Dim lngEnterpriseId       As Long
    Dim lngIdFFIContaCorrente As Long
    Dim strSinal              As String
    Dim strStatus             As String
    Dim curCredito            As Currency
    Dim curDebito             As Currency
    Dim rsAtualizaSaldo       As Object
    Dim strSqlAtualizaSaldo   As String
    Dim rsContaCorrente       As Object
    Dim strContaCorrente      As String
    
    With mrsContaCorrente
        Select Case strFuncao
            Case WL_NOVO
                mblnAlterando = False
                Call LimpaTela
                cmdExcluir.Enabled = False
                cmdGravar.Enabled = False
                etxOperacao.Locked = False
                etxValor.Locked = False
                etxCodigo.valorInteiro = mlngCodigo + 1
                etxCliente.SetFocus
            Case WL_PRIMEIRO
                .MoveFirst
                lngCodigo = .Fields("cd_estabelecimento").value
                lngEnterpriseId = .Fields("enterprise_id").value
                lngIdFFIContaCorrente = .Fields("id_fficontacorrente").value
                mbolNavegador = True
                Call PreencheCampos(lngCodigo, lngEnterpriseId, lngIdFFIContaCorrente)
                mbolNavegador = False
            Case WL_ANTERIOR
                .MovePrevious
                If Not .BOF Then
                    lngCodigo = .Fields("cd_estabelecimento").value
                    lngEnterpriseId = .Fields("enterprise_id").value
                    lngIdFFIContaCorrente = .Fields("id_fficontacorrente").value
                    mbolNavegador = True
                    Call PreencheCampos(lngCodigo, lngEnterpriseId, lngIdFFIContaCorrente)
                    mbolNavegador = False
                Else
                    MsgBox "Já está no primeiro registro.", vbInformation + vbOKOnly, NomeModulo
                End If
            
            Case WL_PROXIMO
                .MoveNext
                If Not .EOF Then
                    lngCodigo = .Fields("cd_estabelecimento").value
                    lngEnterpriseId = .Fields("enterprise_id").value
                    lngIdFFIContaCorrente = .Fields("id_fficontacorrente").value
                    mbolNavegador = True
                    Call PreencheCampos(lngCodigo, lngEnterpriseId, lngIdFFIContaCorrente)
                    mbolNavegador = False
                Else
                    MsgBox "Já está no último registro.", vbInformation + vbOKOnly, NomeModulo
                End If
            Case WL_ULTIMO
                .MoveLast
                lngCodigo = .Fields("cd_estabelecimento").value
                lngEnterpriseId = .Fields("enterprise_id").value
                lngIdFFIContaCorrente = .Fields("id_fficontacorrente").value
                mbolNavegador = True
                Call PreencheCampos(lngCodigo, lngEnterpriseId, lngIdFFIContaCorrente)
                mbolNavegador = False
            Case WL_SAIR
                Unload Me
                Exit Function
            Case WL_CANCELAR
                If MsgBox("Tem certeza que deseja cancelar?", vbQuestion + vbYesNo, NomeModulo) = vbYes Then
                    Call LimpaTela
                End If
            Case WL_SALVAR
                If Not mblnGravando Then
                    If ValidaCampos Then
                        mblnGravando = True
                        SetPtr vbHourglass
                        If UCase(GetFieldValue("sinal", "FFIOperacaoFinanceira", "cd_op_financeira = " & etxOperacao.valorInteiro)) = "C" Then
                            curCredito = etxValor.valorMoeda
                            curDebito = 0
                        Else
                            curCredito = 0
                            curDebito = etxValor.valorMoeda
                        End If
                        If Not mblnAlterando Then
                            strSqlGravar = "INSERT INTO FFIContaCorrente ("
                            strSqlGravar = strSqlGravar & " enterprise_id, "
                            strSqlGravar = strSqlGravar & " cd_estabelecimento, "
                            #If FOXSQL Then
                                strSqlGravar = strSqlGravar & " id_fficontacorrente, "
                            #End If
                            strSqlGravar = strSqlGravar & " apel, "
                            strSqlGravar = strSqlGravar & " dt_lancamento, "
                            strSqlGravar = strSqlGravar & " nr_documento, "
                            strSqlGravar = strSqlGravar & " vl_lancamento, "
                            strSqlGravar = strSqlGravar & " obs, "
                            strSqlGravar = strSqlGravar & " cd_op_financeira"
                            strSqlGravar = strSqlGravar & " )VALUES ("
                            strSqlGravar = strSqlGravar & EnterpriseId & ", "
                            strSqlGravar = strSqlGravar & CdEstabelecimento & ", "
                            #If FOXSQL Then
                                strSqlGravar = strSqlGravar & "'" & etxCodigo.valorInteiro & "',"
                            #End If
                            strSqlGravar = strSqlGravar & "'" & etxCliente.valorTexto & "',"
                            #If FOXSQL Then
                                strSqlGravar = strSqlGravar & "'" & Format(edtDataLote.Data, "MM/DD/YYYY") & "', "
                            #Else
                                strSqlGravar = strSqlGravar & "#" & Format(edtDataLote.Data, "mm/dd/yyyy") & "#, "
                            #End If
                            If etxDocumento.valorTexto <> "" Then
                                strSqlGravar = strSqlGravar & "'" & etxDocumento.valorTexto & "', "
                            Else
                                strSqlGravar = strSqlGravar & "'', "
                            End If
                            strSqlGravar = strSqlGravar & Replace(Format(etxValor.valorMoeda, "0.00#,##"), ",", ".") & ", "
                            If etxDescricao.valorTexto <> "" Then
                                strSqlGravar = strSqlGravar & "'" & Trim(etxDescricao.valorTexto) & "',"
                            Else
                                strSqlGravar = strSqlGravar & "'', "
                            End If
                            strSqlGravar = strSqlGravar & etxOperacao.valorInteiro & ");"
                        Else
                            strSqlGravar = "UPDATE FFIContaCorrente SET "
                            strSqlGravar = strSqlGravar & " apel = '" & etxCliente.valorTexto & "'"
                            #If FOXSQL Then
                                strSqlGravar = strSqlGravar & " ,dt_lancamento = '" & Format(edtDataLote.Data, "MM/DD/YYYY") & "'"
                            #Else
                                strSqlGravar = strSqlGravar & " ,dt_lancamento = #" & Format(edtDataLote.Data, "mm/dd/yyyy") & "#"
                            #End If
                            If etxDocumento.valorTexto <> "" Then
                                strSqlGravar = strSqlGravar & " ,nr_documento = '" & etxDocumento.valorTexto & "'"
                            End If
                            If etxDescricao.valorTexto <> "" Then
                                strSqlGravar = strSqlGravar & " ,obs = '" & Trim(etxDescricao.valorTexto) & "'"
                            End If
                            strSqlGravar = strSqlGravar & " WHERE id_fficontacorrente = " & etxCodigo.valorInteiro
                            strSqlGravar = strSqlGravar & " AND enterprise_id = " & EnterpriseId
                            strSqlGravar = strSqlGravar & " AND cd_estabelecimento = " & CdEstabelecimento & ";"
                        End If
                        If ExecuteSQL(strSqlGravar) And Not mblnAlterando Then
                            strSqlAtualizaSaldo = "SELECT vl_credito, vl_debito FROM FFISaldoCCO WHERE enterprise_id = " & EnterpriseId & " AND cd_estabelecimento = " & CdEstabelecimento & " AND apel = '" & etxCliente.valorTexto & "'"
                            If AbreRecordset(rsAtualizaSaldo, strSqlAtualizaSaldo) = WL_OK Then
                                strSqlGravar = "UPDATE FFISaldoCCO SET "
                                strSqlGravar = strSqlGravar & "vl_credito = " & Replace(Format(curCredito + CCur(rsAtualizaSaldo.Fields("vl_credito").value), "0.00#,##"), ",", ".")
                                strSqlGravar = strSqlGravar & ", vl_debito = " & Replace(Format(curDebito + CCur(rsAtualizaSaldo.Fields("vl_debito").value), "0.00#,##"), ",", ".")
                                strSqlGravar = strSqlGravar & " WHERE enterprise_id = " & EnterpriseId & " AND cd_estabelecimento = " & CdEstabelecimento & " AND apel = '" & etxCliente.valorTexto & "'"
                                Call ExecuteSQL(strSqlGravar)
                            Else
                                strSqlGravar = "INSERT INTO FFISaldoCCO VALUES ("
                                strSqlGravar = strSqlGravar & EnterpriseId & ", "
                                strSqlGravar = strSqlGravar & CdEstabelecimento & ", "
                                strSqlGravar = strSqlGravar & "'" & etxCliente.valorTexto & "', "
                                #If FOXSQL Then
                                    strSqlGravar = strSqlGravar & "'" & Format(edtDataLote.Data, "MM/DD/YYYY") & "', "
                                #Else
                                    strSqlGravar = strSqlGravar & "#" & Format(edtDataLote.Data, "mm/dd/yyyy") & "#, "
                                #End If
                                strSqlGravar = strSqlGravar & Replace(Format(curCredito, "0.00#,##"), ",", ".") & ", "
                                strSqlGravar = strSqlGravar & Replace(Format(curDebito, "0.00#,##"), ",", ".") & ");"
                                Call ExecuteSQL(strSqlGravar)
                            End If
                        End If
                        SetPtr vbDefault
                        mblnGravando = False
                        mblnAlterando = True
                        cmdExcluir.Enabled = True
                        etxCliente.SetFocus
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
                        strSqlDeletar = "DELETE * FROM FFIContaCorrente WHERE id_fficontacorrente = " & mlngIdFFIContaCorrente & _
                                        " AND enterprise_id = " & EnterpriseId & _
                                        " AND cd_estabelecimento = " & CdEstabelecimento
                        If ExecuteSQL(strSqlDeletar) Then
                            strSqlAtualizaSaldo = "SELECT vl_credito, vl_debito FROM FFISaldoCCO WHERE enterprise_id = " & EnterpriseId & " AND cd_estabelecimento = " & CdEstabelecimento & " AND apel = '" & etxCliente.valorTexto & "'"
                            If AbreRecordset(rsAtualizaSaldo, strSqlAtualizaSaldo) = WL_OK Then
                                If UCase(GetFieldValue("sinal", "FFIOperacaoFinanceira", "cd_op_financeira = " & etxOperacao.valorInteiro)) = "C" Then
                                    curCredito = CCur(rsAtualizaSaldo.Fields("vl_credito").value) - etxValor.valorMoeda
                                    curDebito = CCur(rsAtualizaSaldo.Fields("vl_debito").value)
                                Else
                                    curCredito = CCur(rsAtualizaSaldo.Fields("vl_credito").value)
                                    curDebito = CCur(rsAtualizaSaldo.Fields("vl_debito").value) - etxValor.valorMoeda
                                End If
                                strSqlDeletar = "UPDATE FFISaldoCCO SET "
                                strSqlDeletar = strSqlDeletar & " vl_debito = " & Replace(curDebito, ",", ".")
                                strSqlDeletar = strSqlDeletar & ", vl_credito = " & Replace(curCredito, ",", ".")
                                strSqlDeletar = strSqlDeletar & " WHERE enterprise_id = " & EnterpriseId & " AND cd_estabelecimento = " & CdEstabelecimento & " AND apel = '" & etxCliente.valorTexto & "'"
                                Call ExecuteSQL(strSqlDeletar)
                            End If
                            Set rsAtualizaSaldo = Nothing
                        End If
                        Call LibProc(WL_NOVO)
                        mblnAlterando = False
                        etxCliente.SetFocus
                    End If
                Else
                    MsgBox "Escolha um Lançamento antes de tentar excluír.", vbInformation, NomeModulo
                End If
                cmdExcluir.Enabled = False
        End Select
    End With
End Function

Private Sub LimpaTela()
    etxCliente.clear
    edtDataLote.Data = Date
    etxDocumento.clear
    etxSaldo.clear
    etxOperacao.clear
    lblSinal.Caption = Empty
    etxValor.clear
    etxDescricao.clear
    etxCodigo.clear
    mblnAlterando = False
End Sub

Private Sub PreencheCampos(lngCodigo As Long, lngEnterpriseId As Long, lngIdFFIContaCorrente As Long)
    Dim strSqlPreenche      As String
    Dim rsLancContaCorrente As Object
    
    'mblnAlterando = True
    strSqlPreenche = "SELECT id_fficontacorrente, " & _
                     "apel, " & _
                     "dt_lancamento, " & _
                     "nr_documento, " & _
                     "vl_lancamento, " & _
                     "obs, " & _
                     "cd_op_financeira " & _
                     "FROM FFIContaCorrente " & _
                     "WHERE cd_estabelecimento = " & lngCodigo & _
                     " AND enterprise_id = " & lngEnterpriseId & _
                     " AND id_fficontacorrente = " & lngIdFFIContaCorrente
                     
    If AbreRecordset(rsLancContaCorrente, strSqlPreenche) = WL_OK Then
        With rsLancContaCorrente
            etxCliente.valorTexto = .Fields("apel").value
            edtDataLote.Data = .Fields("dt_lancamento").value
            etxDocumento.valorTexto = .Fields("nr_documento").value
            etxOperacao.valorInteiro = .Fields("cd_op_financeira").value
            etxValor.valorMoeda = .Fields("vl_lancamento").value
            etxDescricao.valorTexto = .Fields("obs").value
            mlngIdFFIContaCorrente = .Fields("id_fficontacorrente").value
            etxSaldo.valorMoeda = SaldoConta(.Fields("apel").value)
            mblnAlterando = True
            cmdGravar.Enabled = True
            cmdExcluir.Enabled = True
            etxOperacao.Locked = True
            etxValor.Locked = True
        End With
        Set rsLancContaCorrente = Nothing
    End If
End Sub

Private Sub etxCliente_Change()
    If etxCliente.valorTexto <> "" Then
        cmdGravar.Enabled = True
        etxSaldo.valorMoeda = SaldoConta(etxCliente.valorTexto)
    Else
        etxSaldo.valorMoeda = 0
        cmdGravar.Enabled = False
    End If
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

Private Sub etxCodigo_Change()
    If etxCodigo.valorInteiro > 0 Then
        Call PreencheCampos(CdEstabelecimento, EnterpriseId, etxCodigo.valorInteiro)
    End If
End Sub

Private Sub etxOperacao_Change()
    If etxOperacao.ValorDescricao <> "" Then
        lblSinal.Caption = GetFieldValue("sinal", "FFIOperacaoFinanceira", "cd_op_financeira = " & etxOperacao.valorInteiro)
    Else
        lblSinal.Caption = ""
    End If
End Sub

Private Sub etxOperacao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxOperacao.ValorDescricao = "" Then
            etxOperacao.valorInteiro = 0
        End If
        PCampo "Operação Financeira", "SELECT * FROM FFIOperacaoFinanceira", pbCampo, etxOperacao, "cd_op_financeira"
    End If
End Sub

Private Sub CarregaRecordset()
    Dim strSqlContaCorrente As String
    
    'strSqlContaCorrente = "SELECT * FROM FFIContaCorrente"
    strSqlContaCorrente = "SELECT * FROM FFIContaCorrente ORDER BY id_fficontacorrente"
    If AbreRecordset(mrsContaCorrente, strSqlContaCorrente) = WL_OK Then
        mrsContaCorrente.MoveLast
        mlngCodigo = mrsContaCorrente!id_fficontacorrente
        mrsContaCorrente.MoveFirst
    End If
End Sub

Private Sub Form_Load()
    Call etxEmpUser.AddConexao(Aplicacao)
    Call etxCliente.AddConexao(Aplicacao)
    Call etxOperacao.AddConexao(Aplicacao)
    etxEmpUser.valorTexto = DonaSistema
    edtDataLote.Data = Date
    Set mrsContaCorrente = Nothing
    mblnAlterando = False
    cmdGravar.Enabled = False
    cmdExcluir.Enabled = False
    Call CarregaRecordset
    If mlngCodigo > 0 Then
        etxCodigo.valorInteiro = mlngCodigo
    End If
End Sub

Private Function ValidaCampos() As Boolean
    If Trim(etxCliente.valorTexto) = "" Then
        MsgBox "O campo Cliente deve ser informado.", vbInformation, NomeModulo
        etxCliente.SetFocus
        Exit Function
    ElseIf IsEmptyDate(edtDataLote.Data) Then
        MsgBox "O campo Data deve ser informado.", vbInformation, NomeModulo
        edtDataLote.SetFocus
        Exit Function
    ElseIf etxOperacao.valorInteiro = 0 Then
        MsgBox "O campo Operação deve ser informado.", vbInformation, NomeModulo
        etxOperacao.SetFocus
        Exit Function
    ElseIf etxValor.valorMoeda = 0 Then
        MsgBox "O campo Valor deve ser informado.", vbInformation, NomeModulo
        etxValor.SetFocus
        Exit Function
    Else
        ValidaCampos = True
    End If
End Function

Public Property Let Codigo(ByVal NewVal As Long)
    etxCodigo.valorInteiro = NewVal
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
