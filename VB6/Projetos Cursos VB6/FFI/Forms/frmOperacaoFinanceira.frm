VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmOperacaoFinanceira 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operação Financeira"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   8925
   Begin VB.Frame Frame1 
      Height          =   3150
      Left            =   45
      TabIndex        =   18
      Top             =   -40
      Width           =   7410
      Begin VB.Frame fraSituacao 
         Caption         =   "Situação"
         Height          =   550
         Left            =   5500
         TabIndex        =   32
         Top             =   195
         Width           =   1815
         Begin VB.OptionButton optAtivo 
            Caption         =   "Ativo"
            Height          =   200
            Left            =   100
            TabIndex        =   8
            Top             =   240
            Width           =   675
         End
         Begin VB.OptionButton optInativo 
            Caption         =   "Inativo"
            Height          =   200
            Left            =   840
            TabIndex        =   9
            Top             =   255
            Width           =   790
         End
      End
      Begin VB.Frame fraOperacao 
         Caption         =   "Operação"
         Height          =   550
         Left            =   5500
         TabIndex        =   31
         Top             =   760
         Width           =   1815
         Begin VB.OptionButton optDebito 
            Caption         =   "Débito"
            Height          =   200
            Left            =   960
            TabIndex        =   11
            Top             =   220
            Width           =   790
         End
         Begin VB.OptionButton optCredito 
            Caption         =   "Crédito"
            Height          =   200
            Left            =   100
            TabIndex        =   10
            Top             =   220
            Width           =   790
         End
      End
      Begin Fox.EBSText etxCodigo 
         Height          =   330
         Left            =   1320
         TabIndex        =   0
         Top             =   200
         Width           =   1095
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   5
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
      Begin Fox.EBSText etxSigla 
         Height          =   330
         Left            =   1320
         TabIndex        =   2
         Top             =   910
         Width           =   1095
         _ExtentX        =   265
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   5
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
      Begin Fox.EBSText etxContaFinanceira 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   1635
         Width           =   1095
         _ExtentX        =   741
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Código"
         TipoCriterio    =   4
         CampoDescricao  =   "Descrição"
         TabelaConsulta  =   "Contas"
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
      Begin Fox.EBSText etxCentroCusto 
         Height          =   330
         Left            =   1320
         TabIndex        =   5
         Top             =   1995
         Width           =   1095
         _ExtentX        =   1455
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Código"
         TipoCriterio    =   4
         CampoDescricao  =   "Descrição"
         TabelaConsulta  =   "Centros"
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
      Begin Fox.EBSText etxOpContabil 
         Height          =   330
         Left            =   1320
         TabIndex        =   6
         Top             =   2355
         Width           =   1095
         _ExtentX        =   661
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   5
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "cd_operacao"
         TipoCriterio    =   4
         CampoDescricao  =   "cd_operacao"
         TabelaConsulta  =   "OperacaoContabil"
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
      Begin Fox.EBSText etxOpEstorno 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   2715
         Width           =   1095
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   5
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
         ExibeDescricao  =   0   'False
      End
      Begin Fox.EBSText etxDescricao 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   555
         Width           =   4095
         _ExtentX        =   2461
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   40
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
      Begin Fox.EBSText etxBanco 
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Top             =   1275
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Banco"
         TipoCriterio    =   4
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
         ExibeDescricao  =   0   'False
      End
      Begin VB.Label lblBanco 
         Height          =   195
         Left            =   2520
         TabIndex        =   34
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   810
         TabIndex        =   33
         Top             =   1335
         Width           =   465
      End
      Begin VB.Label lblDescOpEstorno 
         Height          =   195
         Left            =   2520
         TabIndex        =   30
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label lblDescOpContabil 
         Height          =   195
         Left            =   2520
         TabIndex        =   29
         Top             =   2415
         Width           =   4695
      End
      Begin VB.Label lblDescCentroCusto 
         Height          =   195
         Left            =   2520
         TabIndex        =   28
         Top             =   2055
         Width           =   4695
      End
      Begin VB.Label lblDescContaFinanceira 
         Height          =   195
         Left            =   2520
         TabIndex        =   27
         Top             =   1695
         Width           =   4695
      End
      Begin VB.Label lblOpEstorno 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Op. Estorno"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   435
         TabIndex        =   26
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Conta Financeira"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   75
         TabIndex        =   25
         Top             =   1695
         Width           =   1200
      End
      Begin VB.Label lblOpContabil 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Op. Contábil"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         TabIndex        =   24
         Top             =   2415
         Width           =   870
      End
      Begin VB.Label lblCentroCusto 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   2055
         Width           =   1140
      End
      Begin VB.Label lblCodIPI 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   21
         Top             =   260
         Width           =   495
      End
      Begin VB.Label lblDesIPI 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   555
         TabIndex        =   20
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblAliquota 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sigla"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   930
         TabIndex        =   19
         Top             =   970
         Width           =   345
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6075
      Left            =   7485
      TabIndex        =   17
      Top             =   -40
      Width           =   1415
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   15
         Top             =   1400
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   16
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   14
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   13
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdOpFinanceira 
      Height          =   2895
      Left            =   45
      TabIndex        =   22
      Top             =   3120
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmOperacaoFinanceira"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mblnAlterando          As Boolean
Private mblnGravando           As Boolean
Private mbolbuscaPelaGrid      As Boolean
Private mbolNavegador          As Boolean
Private mblnLoad               As Boolean
Private mblnGravou             As Boolean
Private Const strTituloGrid$ = "campo=cd_op_financeira;label=Código;tamanho=670|" & _
                                "campo=descricao;label=Descrição;tamanho=6000|" & _
                                "campo=status;label=Sit;tamanho=400"
                                
Public Function LibProc(strFuncao As String, Optional lngFuncao As Long) As Boolean
    Dim strSqlGravar     As String
    Dim strSqlDeletar    As String
    Dim strTipo          As String
    Dim lngCodigo        As Long
    Dim strSinal         As String
    Dim strStatus        As String
    
    With rsTipoLancamento
        Select Case strFuncao
            Case WL_NOVO
                mblnAlterando = False
                Call LimpaTela
                cmdExcluir.Enabled = False
                cmdGravar.Enabled = False
                If ultimoCodigo(lngCodigo) Then
                    etxCodigo.valorInteiro = lngCodigo
                Else
                    etxCodigo.valorInteiro = 1
                End If
                If Not mblnLoad Then
                    etxCodigo.SetFocus
                End If
                etxCodigo.Enabled = True
            Case WL_PRIMEIRO
                .MoveFirst
                lngCodigo = .Fields("cd_codigo").Value
                mbolNavegador = True
                Call PreencheCampos(lngCodigo)
                mbolNavegador = False
            Case WL_ANTERIOR
                .MovePrevious
                If Not .BOF Then
                    lngCodigo = .Fields("cd_codigo").Value
                    mbolNavegador = True
                    Call PreencheCampos(lngCodigo)
                    mbolNavegador = False
                Else
                    MsgBox "Já está no primeiro registro.", vbInformation + vbOKOnly, NomeModulo
                End If
            
            Case WL_PROXIMO
                .MoveNext
                If Not .EOF Then
                    lngCodigo = .Fields("cd_codigo").Value
                    mbolNavegador = True
                    Call PreencheCampos(lngCodigo)
                    mbolNavegador = False
                Else
                    MsgBox "Já está no último registro.", vbInformation + vbOKOnly, NomeModulo
                End If
            Case WL_ULTIMO
                .MoveLast
                lngCodigo = .Fields("cd_codigo").Value
                mbolNavegador = True
                Call PreencheCampos(lngCodigo)
                mbolNavegador = False
            Case WL_SAIR
                Unload Me
                Exit Function
            Case WL_CANCELAR
                If MsgBox("Tem certeza que deseja cancelar?", vbQuestion + vbYesNo, NomeModulo) = vbYes Then
                    Call LimpaTela
                End If
            Case WL_SALVAR
                If mblnAlterando Then
                    strTipo = "alterar"
                Else
                    strTipo = "incluir"
                End If
                If Not mblnGravando Then
                    If ValidaCampos Then
                        mblnGravando = True
                        SetPtr vbHourglass
                        If optCredito.Value Then
                            strSinal = "C"
                        Else
                            strSinal = "D"
                        End If
                        If optAtivo.Value Then
                            strStatus = "A"
                        Else
                            strStatus = "C"
                        End If
                        If Not mblnAlterando Then
                            strSqlGravar = "INSERT INTO FFIOperacaoFinanceira " & _
                                            "(cd_op_financeira, descricao, sigla," & _
                                            "sinal, cd_operacao_contabil, cd_op_financeira_est," & _
                                            "status, cd_contafinanceira, cd_centrocusto, cd_banco)" & _
                                            " VALUES (" & _
                                            etxCodigo.valorInteiro & ",'" & etxDescricao.valorTexto & "','" & _
                                            Trim(etxSigla.valorTexto) & "','" & strSinal & "'," & _
                                            etxOpContabil.valorInteiro & "," & etxOpEstorno.valorInteiro & ",'" & _
                                            strStatus & "'," & etxContaFinanceira.valorInteiro & "," & etxCentroCusto.valorInteiro & "," & etxBanco.valorInteiro & ");"
                        Else
                            strSqlGravar = "UPDATE FFIOperacaoFinanceira SET "
                            strSqlGravar = strSqlGravar + " descricao = '" & Trim(etxDescricao.valorTexto) & "'"
                            strSqlGravar = strSqlGravar + " ,sigla = '" & Trim(etxSigla.valorTexto) & "'"
                            strSqlGravar = strSqlGravar + " ,sinal = '" & Trim(strSinal) & "'"
                            strSqlGravar = strSqlGravar + " ,cd_operacao_contabil = " & etxOpContabil.valorInteiro
                            strSqlGravar = strSqlGravar + " ,cd_op_financeira_est = " & etxOpEstorno.valorInteiro
                            strSqlGravar = strSqlGravar + " ,status = '" & Trim(strStatus) & "'"
                            strSqlGravar = strSqlGravar + " ,cd_contafinanceira = " & etxContaFinanceira.valorInteiro
                            strSqlGravar = strSqlGravar + " ,cd_centrocusto = " & etxCentroCusto.valorInteiro
                            strSqlGravar = strSqlGravar + " ,cd_banco = " & etxBanco.valorInteiro
                            strSqlGravar = strSqlGravar + " WHERE cd_op_financeira = " & CLng(etxCodigo.valorInteiro) & ";"
                        End If
                        ExecuteSQL strSqlGravar
                        SetPtr vbDefault
                        mblnGravando = False
                        Call carregaGrid
                        mblnAlterando = True
                        cmdExcluir.Enabled = True
                        etxCodigo.SetFocus
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
                        strSqlDeletar = "DELETE FROM FFIOperacaoFinanceira WHERE cd_op_financeira = " & CLng(etxCodigo.valorInteiro) & ";"
                        ExecuteSQL strSqlDeletar
                        Call LibProc(WL_NOVO)
                        mblnAlterando = False
                        Call carregaGrid
                        etxCodigo.SetFocus
                    End If
                Else
                    MsgBox "Escolha um Lançamento antes de tentar excluír.", vbInformation, NomeModulo
                End If
                cmdExcluir.Enabled = False
        End Select
    End With
End Function

Private Sub cmdCancelar_Click()
    Call LibProc(WL_NOVO)
End Sub

Private Sub cmdExcluir_Click()
    If ExisteRegistro(etxCodigo.valorInteiro) Then
        Call LibProc(WL_DELETAR)
        etxCodigo.SetFocus
    End If
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
    If mblnGravou Then
        Call LibProc(WL_NOVO)
        etxCodigo.SetFocus
    End If
End Sub

Private Function ExisteRegistro(lngCodigo As Long) As Boolean
    Dim strOpFinanceira As String
    Dim rsOpFinanceira  As Object
    
    If lngCodigo > 0 Then
        strOpFinanceira = "SELECT cd_op_financeira FROM FFIOperacaoFinanceira WHERE cd_op_financeira = " & etxCodigo.valorInteiro
        If AbreRecordset(rsOpFinanceira, strOpFinanceira) = WL_OK Then
            ExisteRegistro = True
        Else
            ExisteRegistro = False
        End If
    Else
        ExisteRegistro = False
        Call LimpaTela
    End If
End Function

Private Function ultimoCodigo(ByRef lngCodigo As Long) As Boolean
    Dim strSqlCodigo As String
    Dim rsCodigo     As Object
    
    strSqlCodigo = "SELECT MAX (cd_op_financeira) AS codigo FROM FFIOperacaoFinanceira"
    If AbreRecordset(rsCodigo, strSqlCodigo) Then
        If Not IsNull(rsCodigo.Fields("codigo").Value) Then
            If Len(rsCodigo.Fields("codigo").Value) > 6 Then
                etxCodigo.MaxLength = Len(rsCodigo.Fields("codigo").Value)
                lngCodigo = rsCodigo.Fields("codigo").Value + 1
                ultimoCodigo = True
            Else
                lngCodigo = rsCodigo.Fields("codigo").Value + 1
                ultimoCodigo = True
            End If
        Else
            ultimoCodigo = False
        End If
    Else
        ultimoCodigo = False
    End If
End Function

Private Sub carregaGrid()
    Dim rsOpFinanceira    As Object
    Dim sqlOpFinanceira   As String
    
    sqlOpFinanceira = "SELECT * FROM FFIOperacaoFinanceira"
    Call AbreRecordset(rsOpFinanceira, sqlOpFinanceira)
    If Not rsOpFinanceira.EOF Then
        Call CarregaHFlexGrid(grdOpFinanceira, rsOpFinanceira, strTituloGrid)
    Else
        Call CarregaHFlexGrid(grdOpFinanceira, Nothing, strTituloGrid)
    End If
    Set rsOpFinanceira = Nothing
End Sub

Private Sub LimpaTela()
    Dim lngCodigo As Long
    
    If ultimoCodigo(lngCodigo) Then
        etxCodigo.valorInteiro = lngCodigo
    Else
        etxCodigo.valorInteiro = 1
    End If
    etxDescricao.clear
    etxSigla.clear
    etxCentroCusto.clear
    lblDescCentroCusto.Caption = Empty
    etxContaFinanceira.clear
    lblDescContaFinanceira.Caption = Empty
    etxOpContabil.clear
    lblDescOpContabil.Caption = Empty
    etxOpEstorno.clear
    lblDescOpEstorno.Caption = Empty
    etxBanco.clear
    optAtivo.Value = True
    optCredito.Value = True
    mblnAlterando = False
End Sub

Private Sub PreencheCampos(lngCodigo As Long)
    Dim strSqlPreenche As String
    Dim rsOpFinanceira As Object
    
    strSqlPreenche = "SELECT * FROM FFIOperacaoFinanceira WHERE cd_op_financeira = " & lngCodigo
    If AbreRecordset(rsOpFinanceira, strSqlPreenche) = WL_OK Then
        With rsOpFinanceira
            If mbolbuscaPelaGrid Then
                etxCodigo.valorInteiro = .Fields("cd_op_financeira").Value
            End If
            If UCase((Trim(.Fields("sinal").Value))) = "C" Then
                optCredito.Value = True
            Else
                optDebito.Value = True
            End If
            If UCase((Trim(.Fields("status").Value))) = "A" Then
                optAtivo.Value = True
            Else
                optInativo.Value = True
            End If
            etxDescricao.valorTexto = .Fields("descricao").Value
            etxSigla.valorTexto = .Fields("sigla").Value
            If Not IsNull(.Fields("cd_contafinanceira").Value) Then
                etxContaFinanceira.valorInteiro = .Fields("cd_contafinanceira").Value
            End If
            If Not IsNull(.Fields("cd_centrocusto").Value) Then
                etxCentroCusto.valorInteiro = .Fields("cd_centrocusto").Value
            End If
            If Not IsNull(.Fields("cd_operacao_contabil").Value) Then
                etxOpContabil.valorInteiro = .Fields("cd_operacao_contabil").Value
            End If
            If Not IsNull(.Fields("cd_op_financeira_est").Value) Then
                etxOpEstorno.valorInteiro = .Fields("cd_op_financeira_est").Value
            End If
            If Not IsNull(.Fields("cd_banco").Value) Then
                etxBanco.valorInteiro = .Fields("cd_banco").Value
            End If
            etxCodigo.Enabled = False
            mblnAlterando = True
            cmdGravar.Enabled = True
            cmdExcluir.Enabled = True
        End With
    End If
End Sub

Private Sub cmdNovo_Click()
    LibProc (WL_NOVO)
End Sub

Private Sub cmdSair_Click()
    LibProc (WL_SAIR)
End Sub

Private Sub etxCentroCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If lblDescCentroCusto.Caption = "" Then
            etxCentroCusto.valorInteiro = 0
        End If
        PCampo "Centro de Custo", "SELECT * FROM Centros", pbCampo, etxCentroCusto, "Código"
    End If
End Sub

Private Sub etxCentroCusto_Change()
    If etxCentroCusto.valorInteiro > 0 Then
        lblDescCentroCusto.Caption = GetFieldValue("Descrição", "Centros", "Código = " & etxCentroCusto.valorInteiro)
    Else
        lblDescCentroCusto.Caption = ""
    End If
End Sub

Private Sub etxCodigo_LostFocus()
    Dim lngCodigo As Long
    
    If etxCodigo.valorInteiro > 0 Then
        lngCodigo = etxCodigo.valorInteiro
        If ExisteRegistro(lngCodigo) Then
           Call PreencheCampos(lngCodigo)
        Else
            cmdGravar.Enabled = True
            cmdExcluir.Enabled = False
        End If
    Else
        cmdGravar.Enabled = True
        cmdExcluir.Enabled = False
    End If
End Sub

Private Function ValidaCampos() As Boolean
    If etxCodigo.valorInteiro = 0 Then
        MsgBox "O campo Código deve ser informado.", vbInformation, NomeModulo
        etxCodigo.valorInteiro = 0
        etxCodigo.SetFocus
        Exit Function
    ElseIf Trim(etxDescricao.valorTexto) = "" Then
        MsgBox "O campo Descrição deve ser informado.", vbInformation, NomeModulo
        etxDescricao.SetFocus
        Exit Function
    ElseIf Trim(etxSigla.valorTexto) = "" Then
        MsgBox "O campo Sigla deve ser informado.", vbInformation, NomeModulo
        etxSigla.SetFocus
        Exit Function
    ElseIf Trim(etxContaFinanceira.valorInteiro) = 0 Then
        MsgBox "O campo Conta Financeira deve ser informado.", vbInformation, NomeModulo
        etxContaFinanceira.SetFocus
        Exit Function
    ElseIf Trim(etxCentroCusto.valorInteiro) = 0 And ConfigSys.ControlarCentrodeCusto Then
        MsgBox "O campo Centro de Custo deve ser informado.", vbInformation, NomeModulo
        etxCentroCusto.SetFocus
        Exit Function
    Else
        ValidaCampos = True
    End If
End Function

Private Sub etxContaFinanceira_Change()
    If etxContaFinanceira.valorInteiro > 0 Then
        lblDescContaFinanceira.Caption = GetFieldValue("Descrição", "Contas", "Código = " & etxContaFinanceira.valorInteiro)
    Else
        lblDescContaFinanceira.Caption = ""
    End If
End Sub

Private Sub etxContaFinanceira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If lblDescContaFinanceira.Caption = "" Then
            etxContaFinanceira.valorInteiro = 0
        End If
        PCampo "Conta Financeira", "SELECT * FROM Contas", pbCampo, etxContaFinanceira, "Código"
    End If
End Sub

Private Sub etxOpContabil_Change()
    If etxOpContabil.valorInteiro > 0 Then
        lblDescOpContabil.Caption = GetFieldValue("descricao", "OperacaoContabil", "cd_operacao = " & etxOpContabil.valorInteiro)
    Else
        lblDescOpContabil.Caption = ""
    End If
End Sub

Private Sub etxOpContabil_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If lblDescOpContabil.Caption = "" Then
            etxOpContabil.valorInteiro = 0
        End If
        PCampo "Operação Contábil", "SELECT * FROM OperacaoContabil", pbCampo, etxOpContabil, "cd_operacao"
    End If
End Sub

Private Sub etxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If lblBanco.Caption = "" Then
            etxBanco.valorInteiro = 0
        End If
        PCampo "Bancos", "SELECT * FROM Bancos", pbCampo, etxBanco, "Banco"
    End If
End Sub

Private Sub etxBanco_Change()
    If etxBanco.valorInteiro > 0 Then
        lblBanco.Caption = GetFieldValue("Nome", "Bancos", "Banco = " & etxBanco.valorInteiro)
    Else
        lblBanco.Caption = ""
    End If
End Sub

Private Sub etxOpEstorno_Change()
    Dim strOperacao As String
    
    If optCredito.Value Then
        strOperacao = "C"
    Else
        strOperacao = "D"
    End If
    If etxOpEstorno.valorInteiro > 0 Then
        lblDescOpEstorno.Caption = GetFieldValue("descricao", "FFIOperacaoFinanceira", "cd_op_financeira = " & etxOpEstorno.valorInteiro & " AND sinal <> '" & strOperacao & "'")
    Else
        lblDescOpEstorno.Caption = ""
    End If
End Sub

Private Sub etxOpEstorno_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strOperacao As String
    
    If KeyCode = vbKeyPageDown Then
        If lblDescOpEstorno.Caption = "" Then
            etxOpEstorno.valorInteiro = 0
        End If
        If optCredito.Value Then
            strOperacao = "C"
        Else
            strOperacao = "D"
        End If
        PCampo "Operação Estorno", "SELECT * FROM FFIOperacaoFinanceira WHERE sinal <> '" & strOperacao & "'", pbCampo, etxOpEstorno, "cd_op_financeira"
    End If
End Sub

Private Sub etxOpEstorno_LostFocus()
    If etxOpEstorno.valorInteiro > 0 Then
        If lblDescOpEstorno.Caption = "" Then
            If Screen.ActiveForm Is frmOperacaoFinanceira Then
                MsgBox "A operação de estorno informada é inválida.", vbInformation, NomeModulo
                etxOpEstorno.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Call etxCentroCusto.AddConexao(Aplicacao)
    Call etxContaFinanceira.AddConexao(Aplicacao)
    Call etxOpContabil.AddConexao(Aplicacao)
    Call etxOpEstorno.AddConexao(Aplicacao)
    Call etxBanco.AddConexao(Aplicacao)
    etxCentroCusto.Enabled = ConfigSys.ControlarCentrodeCusto
    lblCentroCusto.Enabled = ConfigSys.ControlarCentrodeCusto
    etxOpContabil.Enabled = ConfigSys.UtilizaIntegracaoContabil
    lblOpContabil.Enabled = ConfigSys.UtilizaIntegracaoContabil
    mblnLoad = True
    LibProc (WL_NOVO)
    mblnLoad = False
    Call carregaGrid
End Sub

Private Sub grdOpFinanceira_DblClick()
    Dim lngCodigo    As Long
    
    If grdOpFinanceira.TextMatrix(grdOpFinanceira.Row, 0) <> "" Then
        lngCodigo = grdOpFinanceira.TextMatrix(grdOpFinanceira.Row, 0)
        mbolbuscaPelaGrid = True
        Call PreencheCampos(lngCodigo)
        mbolbuscaPelaGrid = False
    End If
    etxCodigo.SetFocus
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
