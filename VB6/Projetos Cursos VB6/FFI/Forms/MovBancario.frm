VERSION 5.00
Begin VB.Form frmMovBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimento Banc�rio"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7635
   Icon            =   "MovBancario.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   7635
   Tag             =   "MovBancario"
   Begin VB.Frame Frame9 
      Height          =   4620
      Left            =   6230
      TabIndex        =   36
      Top             =   -60
      Width           =   1380
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   43
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   42
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   41
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   40
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   90
         TabIndex        =   39
         Top             =   1740
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   38
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   37
         Top             =   2130
         Width           =   1215
      End
   End
   Begin VB.Frame fraMovBancario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4620
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   6165
      Begin VB.TextBox txtMovBancario 
         DataField       =   "cd_operacao_baixa"
         Height          =   330
         Index           =   15
         Left            =   1680
         TabIndex        =   23
         Tag             =   "MovBancario"
         Top             =   3780
         Width           =   1095
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "C�digo"
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         Tag             =   "MovBancario"
         Top             =   180
         Width           =   1455
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Empresa"
         Height          =   315
         Index           =   2
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "MovBancario"
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Pagamento"
         Height          =   315
         Index           =   3
         Left            =   1680
         TabIndex        =   9
         Tag             =   "MovBancario"
         Top             =   1260
         Width           =   1455
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Controle"
         Height          =   315
         Index           =   4
         Left            =   1680
         TabIndex        =   11
         Tag             =   "MovBancario"
         Top             =   1620
         Width           =   2175
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Descri��o"
         Height          =   315
         Index           =   5
         Left            =   1680
         TabIndex        =   13
         Tag             =   "MovBancario"
         Top             =   1980
         Width           =   4095
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Valor Original"
         Height          =   315
         Index           =   6
         Left            =   1680
         TabIndex        =   15
         Tag             =   "MovBancario"
         Top             =   2340
         Width           =   1455
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Banco"
         Height          =   315
         Index           =   7
         Left            =   1680
         TabIndex        =   17
         Tag             =   "MovBancario"
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Conta"
         Height          =   315
         Index           =   8
         Left            =   1680
         TabIndex        =   19
         Tag             =   "MovBancario"
         Top             =   3060
         Width           =   1095
      End
      Begin VB.ComboBox cboMovBancario 
         DataField       =   "Tipo"
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "MovBancario"
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Centro"
         Height          =   315
         Index           =   13
         Left            =   1680
         TabIndex        =   21
         Tag             =   "MovBancario"
         Top             =   3420
         Width           =   1095
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Cheque"
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         Tag             =   "MovBancario"
         Top             =   4140
         Width           =   1095
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Usu�rio"
         Height          =   315
         Index           =   14
         Left            =   4500
         TabIndex        =   35
         Tag             =   "MovBancario"
         Top             =   4140
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkConciliado 
         Alignment       =   1  'Right Justify
         Caption         =   "Conciliado"
         DataField       =   "Conciliado"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4485
         TabIndex        =   3
         Tag             =   "MovBancario"
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "PagRec"
         Height          =   315
         Index           =   9
         Left            =   4965
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "MovBancario"
         Top             =   540
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Emiss�o"
         Height          =   315
         Index           =   12
         Left            =   4965
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "MovBancario"
         Top             =   1620
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Libera��o"
         Height          =   315
         Index           =   10
         Left            =   4965
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "MovBancario"
         Top             =   900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtMovBancario 
         DataField       =   "Vencimento"
         Height          =   315
         Index           =   11
         Left            =   4965
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "MovBancario"
         Top             =   1260
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblDescMovBancario 
         AutoSize        =   -1  'True
         Caption         =   "lblDescMovBancario(0)"
         Height          =   195
         Index           =   0
         Left            =   2895
         TabIndex        =   34
         Top             =   3825
         Width           =   1650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Op. Cont�bil"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   705
         TabIndex        =   22
         Top             =   3825
         Width           =   870
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Lan�amento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   525
         TabIndex        =   1
         Top             =   210
         Width           =   1050
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Tipo Conta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   4
         Top             =   570
         Width           =   945
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   6
         Top             =   930
         Width           =   735
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   1155
         TabIndex        =   8
         Top             =   1290
         Width           =   420
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "N� Documento"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   525
         TabIndex        =   10
         Top             =   1650
         Width           =   1050
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Descri��o"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   855
         TabIndex        =   12
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   1125
         TabIndex        =   14
         Top             =   2370
         Width           =   450
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Banco"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   1110
         TabIndex        =   16
         Top             =   2730
         Width           =   465
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Conta Financeira"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   3090
         Width           =   1455
      End
      Begin VB.Label lblDescMovBancario 
         Caption         =   "lblDescMovBancario(2)"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   26
         Top             =   900
         Width           =   2280
      End
      Begin VB.Label lblDescMovBancario 
         Caption         =   "lblDescMovBancario(7)"
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   31
         Top             =   2700
         Width           =   2280
      End
      Begin VB.Label lblDescMovBancario 
         Caption         =   "lblDescMovBancario(8)"
         Height          =   195
         Index           =   8
         Left            =   2880
         TabIndex        =   32
         Top             =   3060
         Width           =   2280
      End
      Begin VB.Label lblDescMovBancario 
         Caption         =   "lblDescMovBancario(13)"
         Height          =   195
         Index           =   13
         Left            =   2880
         TabIndex        =   33
         Top             =   3420
         Width           =   3240
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Centro de Custo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   435
         TabIndex        =   20
         Top             =   3450
         Width           =   1140
      End
      Begin VB.Label lblMovBancario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Che&que"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   1020
         TabIndex        =   24
         Top             =   4170
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmMovBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gstrPagRec         As String         ' Indica se s�o Lan�amentos � pagar ou � receber

Private mstrSelect        As String         ' Select pra tabela certa
Private mlngMovBancario   As Long           ' Controla as a��es do usu�rio
Private mrstMovBancario   As Object         ' Recordset da tabela de Lan�amentos

Private SeqLancamentos  As Boolean          'Configura��o para sugerir seq��ncia de Lan�amentos
Private mlngOperacao    As Long

' FUNCTION..: LibProc
' Objetivo..: Executa comandos e fun��o da Lib
' Argumentos: [sFuncao]: Fun��o que deve ser executada
'             [lFuncao]: Par�metro adicional, dependente da fun��o.
' Retorna...: Se obtiver sucesso retorna True, caso contr�rio False
' ------------------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
    Dim strMovBancario      As String
    Dim Banco               As Long
    Dim Cheque              As Long
    Dim strSql              As String
    Dim strWhere            As String
    Dim EraEdicao           As Boolean
    Dim strProximoNumero    As String
    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
    Dim objOldStateObj      As New VoLancamentoDuplicata
    Dim objNewStateObj      As New VoLancamentoDuplicata
    Dim objBizLancDup       As New BizLancamentoDuplicata
    Dim objLogLancDup       As New clsLogLancamentosDuplicatas
    Dim strEmpresa          As String
    Dim dblCodigo           As Double
    Dim strTipo             As String
    
    Select Case sFuncao
        'Bot�o Novo
        Case WL_NOVO
            If (LimpaControles(mrstMovBancario, Me, Tag, mlngMovBancario) = WL_OK) Then
                Call sugereOperacao
                LibProc = True
                
                If SeqLancamentos Then
                    strWhere = "PagRec = " & Quote(gstrPagRec, "''")
                End If
                
                txtMovBancario(9).Text = gstrPagRec
                txtMovBancario(7).Text = LastValue("Banco", "Lan�amentos", strWhere, , 0)
                txtMovBancario(8).Text = LastValue("Conta", "Lan�amentos", strWhere, , 0)
                txtMovBancario(3).Text = LastValue("Pagamento", "Lan�amentos", strWhere, , Date)
                
                strProximoNumero = ProximoNumero("C�digo", "Lan�amentos", "PagRec = " & Quote(gstrPagRec, IIf(gTipoDB = Access, """", "''")))
                #If FOXSQL = 1 Then
                If Len(CDec(Trim(strProximoNumero))) <= 15 Then
                #Else
                If Len(CDec(Trim(strProximoNumero))) <= 9 Then
                #End If
                    txtMovBancario(0).Text = strProximoNumero
                Else
                    txtMovBancario(0).Text = ProximoGapDeNumero(gstrPagRec)
                End If
                FirstFocus txtMovBancario(0)
                DefAddNew mlngMovBancario
            End If
    
        'Bot�o Deletar
        Case WL_DELETAR
            Banco = GetValue(mrstMovBancario, "Banco", ZERO)
            Cheque = GetValue(mrstMovBancario, "Cheque", ZERO)
            
            dblCodigo = CDblDef(txtMovBancario(0).Text, 0)
            strEmpresa = txtMovBancario(2).Text
            strTipo = cboMovBancario.Text
            If DeletaRegistro(mrstMovBancario, Me, Tag, mlngMovBancario) = WL_OK Then
                'Projeto: 100340 - Desenv.: 142890 - Ueder Budni (23/09/2016)
                With objLogLancDup
                    Call .SetKey(gstrPagRec, dblCodigo, strEmpresa, strTipo, 1, Lancamento)
                    Call .InsertMsg("T�tulo exclu�do atrav�s da rotina de " & Me.Caption & ".")
                End With
                If ExisteCheque(Banco, Cheque) = 0 Then
                    Call ExecuteSQL("Delete from Cheque where Banco = " & Banco & " and Cheque = " & Cheque)
                End If
            End If
    
        'Bot�o Editar
        Case WL_EDITAR
            Call AlteraValor(mlngMovBancario)
    
        'Bot�o Localizar
        Case WL_LOCALIZAR
            Call localizar(mrstMovBancario, Me, mstrSelect, Tag, mlngMovBancario)
    
        'Bot�o Pesquisar
        Case WL_PESQUISAR
            DefAddNew mlngMovBancario
            Call PRegistro(mrstMovBancario, Me, "Lan�amentos", mstrSelect, mstrSelect, Tag, mlngMovBancario, pbRegistro)
            'Projeto: #218 - Hist�ria: #268 - Desenvolvimento#621 - Moacir Pfau(21/09/2012)
            If val(txtMovBancario(0).Text) > 0 Then
                If txtMovBancario(15).Text = 0 Then
                    txtMovBancario(15).Text = GetFieldValue("cd_operacao_baixa", "Lan�amentos", " PagRec = '" & gstrPagRec & "' AND C�digo = " & txtMovBancario(0).Text)
                End If
            End If
        'Bot�es Primeiro, anterior, Pr�ximo e �ltimo Registro
        Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
            If MoveRecordset(mrstMovBancario, Me, Tag, mlngMovBancario, lFuncao) = MC_ADDNEW Then
                txtMovBancario(9).Text = gstrPagRec
            End If
          
        'Bot�o Navegar
        Case WL_NAVEGAR
            Call Browse(mrstMovBancario, Me, Tag, mlngMovBancario, mstrSelect)
    
        'Bot�o Sair
        Case WL_SAIR
            Unload Me
            Exit Function
    
        'Bot�o Salvar
        Case WL_SALVAR
            If verificaCampos Then
                EraEdicao = False
                If EEdicao(mlngMovBancario) Then
                    EraEdicao = True
                    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
                    Set objOldStateObj = objBizLancDup.Carregar(IIf(gstrPagRec = "P", Pagamento, Recebimento), txtMovBancario(0).Text, cboMovBancario.Text, 1, txtMovBancario(2).Text, Lancamento)
                Else
                    'pt. 104930 - Ivo Sousa (02/05/2011)
                    'Vinicius Elyseu(22/12/2015) - Protocolo #374105
                    'strProximoNumero = ProximoNumero("C�digo", "Lan�amentos", "PagRec = " & Quote(gstrPagRec, IIf(gTipoDB = Access, """", "''")))
                    'If strProximoNumero <> txtMovBancario(0).Text Then
                    If txtMovBancario(0).Text <> "" Then
                        #If FOXSQL = 1 Then
                        If Len(CDec(Trim(txtMovBancario(0).Text))) > 15 Then
                        #Else
                        If Len(CDec(Trim(txtMovBancario(0).Text))) > 9 Then
                        #End If
                            'txtMovBancario(0).Text = strProximoNumero
                        'Else
                            txtMovBancario(0).Text = ProximoGapDeNumero(gstrPagRec)
                        End If
                    End If
                End If
                txtMovBancario(14).Text = UserName()
     
                If SalvaRegistro(mrstMovBancario, Me, Tag, mlngMovBancario) = WL_OK Then
                    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
                    If EraEdicao Then
                        Set objNewStateObj = objBizLancDup.Carregar(IIf(gstrPagRec = "P", Pagamento, Recebimento), txtMovBancario(0).Text, cboMovBancario.Text, 1, txtMovBancario(2).Text, Lancamento)
                    End If
                    LibProc = True
                    If gstrPagRec = "P" Then
                        Banco = CLngDef(txtMovBancario(7).Text)
                        Cheque = CLngDef(txtMovBancario(1).Text)
                        
                        'Pt. 96078 - Moacir Pfau(27/11/2009)
                        If Cheque > 0 Then
                            strSql = "Select * from Cheque where Banco = " & Banco & " and Cheque = " & Cheque
                            If EraEdicao Then
                                If ExisteCheque(Banco, Cheque) = 0 Then
                                    ExecuteSQL "Delete from Cheque where Banco = " & Banco & " and Cheque = " & Cheque
                                End If
                            End If
                            If (EraEdicao = False) Or (GetValue(mrstMovBancario, "Cheque", ZERO) = ZERO And IsValid(txtMovBancario(1).Text)) Then
                                If Recordcount(strSql) > 0 Then
                                    ExecuteSQL "Delete from Cheque where Banco = " & Banco & " and Cheque = " & Cheque
                                End If
                                ExecuteSQL "Insert into Cheque (Banco, Cheque, Nominal) " & " Values (" & Banco & ", " & Cheque & ", '" & GetFieldValue("Raz�o", "Empresas", "Apel = " & Quote(GetValue(mrstMovBancario, "Empresa", NUL), "'"), , NUL) & "')"
                            ElseIf EraEdicao Then
                                If Recordcount(strSql) = 0 Then
                                    ExecuteSQL "Insert into Cheque (Banco, Cheque, Nominal) " & " Values (" & Banco & ", " & Cheque & ", '" & GetFieldValue("Raz�o", "Empresas", "Apel = " & Quote(GetValue(mrstMovBancario, "Empresa", NUL), "'"), , NUL) & "')"
                                End If
                            End If
                        End If
                    End If
                    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
                    Call objLogLancDup.SetKey(gstrPagRec, txtMovBancario(0).Text, txtMovBancario(2).Text, cboMovBancario.Text, 1, Lancamento)
                    If EraEdicao Then
                        Call objLogLancDup.InsertDiffObject(objOldStateObj, objNewStateObj, Me.Caption)
                    Else
                        Call objLogLancDup.InsertMsg("Titulo criado pela rotina " & Me.Caption & ".")
                    End If
                    
                    Call MsgBox("Registro gravado com sucesso!", vbInformation, NomeModulo)
                End If
            End If
    
        'Bot�o Cancelar
        Case WL_CANCELAR
            Call CancelaEdicao(mrstMovBancario, Me, Tag, mlngMovBancario)
    
        'Op��o Exibir
        Case WL_EXIBIR
            strMovBancario = "SELECT PagRec, C�digo, Empresa, Emiss�o, Vencimento, Libera��o, Pagamento," & _
                              "[Valor Original], Tipo, Descri��o, Controle, Banco, Conta, Centro, Cheque, Usu�rio, Conciliado, cd_operacao_baixa FROM Lan�amentos " & _
                              "WHERE Pagamento IS NOT NULL AND PagRec = '{PagRec}' AND C�digo = {C�digo};"
            Call RetornaRegs(mrstMovBancario, Me, Tag, strMovBancario, mlngMovBancario)
            If EAddNew(mlngMovBancario) Then
                Call sugereOperacao
            End If
    
        'Op��o Filtrar
        Case WL_FILTRAR
          Call Filtrar(mrstMovBancario, Me, Tag, mstrSelect, mlngMovBancario)

        'Cadastro de Empresas
        Case "Empresas"
            If KeybAcesso(LoadResString(2102)) Then
                frmEmpresas.Show
                frmEmpresas.ZOrder
                CallChange frmBancos.hWnd, txtMovBancario(2).hWnd
            End If
    
        'Cadastro de Bancos
        Case "Bancos"
            If KeybAcesso(LoadResString(2102)) Then
                frmBancos.Show
                frmBancos.ZOrder
                CallChange frmBancos.hWnd, txtMovBancario(7).hWnd
            End If
    
        'Cadastro de Contas
        Case "Contas"
            If KeybAcesso(LoadResString(2103)) Then
                frmContas.Show
                frmContas.ZOrder
                CallChange frmContas.hWnd, txtMovBancario(8).hWnd
            End If

        'Cadastro de Centros de Custos
        Case "Centros de Custos"
            If KeybAcesso(LoadResString(2029)) Then
                frmCusto.Show
                frmCusto.ZOrder
                CallChange frmCusto.hWnd, txtMovBancario(13).hWnd
            End If
    End Select
    'Projeto: 100340 - Desenv.: 145973 - Ueder Budni (13/10/2016)
    Set objOldStateObj = Nothing
    Set objBizLancDup = Nothing
    Set objLogLancDup = Nothing
End Function

' FUNCTION..: VerificaCampos
' Objetivo..: Verificar a validade dos dados inseridos pelo usu�rio nos campos
' Retorna...: True se os dados estiverem corretos, False se n�o.
' ----------------------------------------------------------------------------
Private Function verificaCampos() As Boolean
  
    'Pt. 95368 - Moacir Pfau(26/10/2009)
    If Not IsDate(txtMovBancario(3).Text) Then
        MsgBox "O campo data deve ser preenchido."
        txtMovBancario(3).SetFocus
        Exit Function
    End If
  
    'pt. 86132 - Ivo Sousa(01/04/2008)
    'Valida o movimento conferido para a data informada
    If Not ValidaDatasDiasUteis(0, 0, CDate(txtMovBancario(3).Text)) Then
        txtMovBancario(3).SetFocus
        Exit Function
    End If

    'Data
    If Not EData(txtMovBancario(3).Text) Then
        MsgFunc "O campo 'Data' n�o cont�m uma data v�lida."
        txtMovBancario(3).SetFocus
        Exit Function
    End If
    
    'Valor Original
    'Pt. 96078 - Moacir Pfau(27/11/2009)
    If IsNumeric(txtMovBancario(6).Text) Then
        If txtMovBancario(6).Text <= 0 Then
            MsgBox "O campo valor deve ser preenchido."
            txtMovBancario(6).SetFocus
            Exit Function
        End If
    Else
        MsgBox "O campo valor deve ser preenchido."
        txtMovBancario(6).SetFocus
        Exit Function
    End If
    
    'Banco
    If (IsValid(txtMovBancario(7).Text) And Len(lblDescMovBancario(7).Caption) = 0) Then
        If MsgBox(ResolveResString(35, "|1", txtMovBancario(7).Text, "|2", "Bancos"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
            LibProc "Bancos"
            Exit Function
        End If
    End If
    
    'Verificar se conta � ativa ou n�o
    If GetFieldValue("Ctaati", "Contas", " [C�digo]=" & txtMovBancario(8).Text) = "N" Then
        MsgBox "Conta " & txtMovBancario(8).Text & " n�o est� ativa", vbCritical, MsgBoxCaption
        txtMovBancario(8).SetFocus
        Exit Function
    End If
  
    'Centros de Custos
    If ConfigSys.ControlarCentrodeCusto Then
        If Not IsValid(txtMovBancario(13).Text) Then
            MsgFunc "O campo 'Centro de Custo' deve ser preenchido"
            txtMovBancario(13).SetFocus
            Exit Function
        End If
        If (IsValid(txtMovBancario(13).Text) And Len(lblDescMovBancario(13).Caption) = 0) Then
            If MsgBox(ResolveResString(35, "|1", txtMovBancario(13).Text, "|2", "Centros de Custos"), vbQuestion Or vbYesNo, MsgBoxCaption) = vbYes Then
                LibProc "Centros de Custos"
                Exit Function
            End If
        End If
    End If
    
    'Empresa
    If lblDescMovBancario(2).Caption = "" Then
        MsgBox "Empresa nao cadastrada no sistema", vbInformation
        txtMovBancario(2).SetFocus
        Exit Function
    End If

    ' Verificando se o tipo do Lan�amento digitado � um novo tipo
    If Len(cboMovBancario.Text) > 0 Then
        If IndexOf(cboMovBancario.Text, cboMovBancario) = NENHUM Then
            ' Salvando o novo tipo na tabela de �tens de Lista de Op��o
            Dim strOptions  As String
            strOptions = "INSERT INTO Op��es (Rotina, Texto, Descri��o) " & "VALUES ('" & OPT_LANCAMENTOS & "', '" & cboMovBancario.Text & "', '');"
            ExecuteSQL strOptions
            cboMovBancario.AddItem cboMovBancario.Text
        End If
    End If
    
    If (CLngDef(txtMovBancario(13).Text) > 0) And Len(txtMovBancario(3).Text) Then
        ' Verifica se a data de libera��o est� dentro da data limite do centro de custo
        If DataLimiteCentroCusto(CLngDef(txtMovBancario(13).Text), txtMovBancario(3).Text) Then
            txtMovBancario(3).SetFocus
            Exit Function
        End If
    End If
    
    'Opera��o Contabil
    If txtMovBancario(15).Enabled Then
        If txtMovBancario(15).Text <> "" Or txtMovBancario(15).Text <> "0" Then
            If lblDescMovBancario(0).Caption = "" Then
                MsgBox "O campo c�digo da Opera��o cont�bil � obrigat�rio.", vbInformation, "Valida��o de Campos"
                txtMovBancario(15).SetFocus
                Exit Function
            End If
        End If
    End If
    
    'pt. 86791 - Ivo Sousa(02/06/2008)
    'Conta Financeira
    If txtMovBancario(8).Text = "" Or txtMovBancario(8).Text = "0" Then
        MsgBox "O campo c�digo da conta financeira � obrigat�rio.", vbInformation, NomeModulo
        txtMovBancario(8).SetFocus
        Exit Function
    ElseIf Not IsValid(lblDescMovBancario(8).Caption) Then
        MsgBox "O c�digo da conta financeira informado n�o existe.", vbInformation, NomeModulo
        txtMovBancario(8).SetFocus
        Exit Function
    End If
        
    'pt. 86728 - Moacir Pfau(09/06/2008)
     If Not (fEmpresaBloqueada(txtMovBancario(2).Text, CDate(Format(Now, "DD/MM/YYYY")))) Then
        verificaCampos = False
     End If
     
    verificaCampos = True
End Function

Private Sub cboMovBancario_Click()
    If EAddNew(mlngMovBancario) Then
        Call sugereOperacao
    End If
End Sub

Private Sub chkConciliado_Click()
  AlteraValor mlngMovBancario
End Sub

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Call LibProc(WL_CANCELAR)
End Sub

Private Sub cmdExcluir_Click()
    Call LibProc(WL_DELETAR)
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
End Sub

Private Sub cmdNovo_Click()
    Call LibProc(WL_NOVO)
End Sub

Private Sub cmdPesquisar_Click()
    Call LibProc(WL_PESQUISAR)
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
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
  GetKeyDown Me, KeyCode, Shift
End Sub

Private Sub Form_Load()
    Dim strOptCombo     As String
    
    lblDescMovBancario(2).Caption = NUL
    lblDescMovBancario(7).Caption = NUL
    lblDescMovBancario(8).Caption = NUL
    lblDescMovBancario(13).Caption = NUL
  
    Call ConfigCampos(Me, "Lan�amentos", Tag)
    
    txtMovBancario(9).Text = gstrPagRec
    
    If gstrPagRec <> "P" Then
        lblMovBancario(9).Enabled = False
        txtMovBancario(1).Enabled = False
    End If
        
    strOptCombo = "SELECT Texto FROM Op��es WHERE Rotina = '" & OPT_LANCAMENTOS & "';"
    Call ComboAddItem(cboMovBancario, strOptCombo, "Texto")
    'Pt. 95368 - Moacir Pfau(26/10/2009)
    Call AbreRecordset(mrstMovBancario, "SELECT * FROM Lan�amentos WHERE PagRec = '" & gstrPagRec & "' AND Pagamento IS NOT NULL", dbOpenDynaset)
    
    mstrSelect = "SELECT PagRec, C�digo, Empresa, Emiss�o, Vencimento, Libera��o, Pagamento," & _
                 "[Valor Original], Tipo, Descri��o, Controle, Banco, Conta, Centro, Cheque,Usu�rio, Conciliado, cd_operacao_baixa FROM Lan�amentos " & _
                 "WHERE Pagamento IS NOT NULL AND PagRec = '" & gstrPagRec & "'"
                
    ' Centro de Custos
    lblMovBancario(13).Enabled = ConfigSys.ControlarCentrodeCusto
    lblDescMovBancario(13).Enabled = ConfigSys.ControlarCentrodeCusto
    txtMovBancario(13).Enabled = ConfigSys.ControlarCentrodeCusto
    If txtMovBancario(13).Enabled Then
        lblMovBancario(13).FontBold = True
    End If
    
    Call DefineAcesso(mlngMovBancario, Acesso())
    Call DefAddNew(mlngMovBancario)
    SeqLancamentos = Configuracao("Seq�enciar Lan�amentos de Entrada e de Sa�da", False)
    Label1.Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    txtMovBancario(15).Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    lblDescMovBancario(0).Enabled = Configuracao("Utiliza Integra��o Cont�bil", False)
    txtMovBancario(0).MaxLength = 15
    Call LibProc(WL_NOVO)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = UnloadForm(mrstMovBancario, Me, Tag, mlngMovBancario)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmMovBancario = Nothing
End Sub

Private Sub txtMovBancario_Change(Index As Integer)
    If Index > 0 Then AlteraValor mlngMovBancario
    Select Case Index
        Case 2 'Campo Empresa
            If IsValid(txtMovBancario(2).Text) Then
                GetAssocValue "SELECT Raz�o, Apel FROM Empresas WHERE Apel = '" & txtMovBancario(2).Text & "'", lblDescMovBancario(2)
            Else
                lblDescMovBancario(2).Caption = ""
            End If
        Case 3 'Campo Data
            ' Atualiza todas as datas
            txtMovBancario(10).Text = txtMovBancario(3).Text
            txtMovBancario(11).Text = txtMovBancario(3).Text
            txtMovBancario(12).Text = txtMovBancario(3).Text
        Case 7 'Campo Banco
            GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtMovBancario(7).Text, lblDescMovBancario(7)
        Case 8 'Campo Conta
            GetAssocValue "SELECT Descri��o FROM Contas WHERE C�digo = " & txtMovBancario(8).Text, lblDescMovBancario(8)
        Case 13 'Campo Centro de Custo
            GetAssocValue "SELECT Descri��o FROM Centros WHERE C�digo = " & txtMovBancario(13).Text, lblDescMovBancario(13)
        Case 15 'Campo Opera��o Cont�bil
            GetAssocValue "SELECT descricao FROM OperacaoContabil WHERE cd_operacao=" & txtMovBancario(Index).Text, lblDescMovBancario(0)
    End Select
End Sub

Private Sub txtMovBancario_GotFocus(Index As Integer)
    Selecione txtMovBancario(Index)
    Select Case Index
        Case 7 'Banco
            MsgBar DescCampo(mrstMovBancario, txtMovBancario(Index).DataField) & ResolveResString(75, "|1", "Bancos")
        Case 8 'Conta
            MsgBar DescCampo(mrstMovBancario, txtMovBancario(Index).DataField) & ResolveResString(75, "|1", "Contas")
        Case 13 'Centro de Custo
            MsgBar DescCampo(mrstMovBancario, txtMovBancario(Index).DataField) & ResolveResString(75, "|1", "Centros de Custos")
        Case Else 'Outros
            MsgBar DescCampo(mrstMovBancario, txtMovBancario(Index).DataField)
    End Select
End Sub


Private Sub txtMovBancario_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    ' Controla o �ndice
    If Index = 0 Then ControlaChave KeyCode, Shift, txtMovBancario(0), mlngMovBancario
    ' Pesquisa de Campo
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Select Case Index
            Case 2 'Empresa
                PCampo "Empresas", "Empresas", pbCampo, txtMovBancario(2), "Apel"
            Case 7 'Banco
                PCampo "Bancos", "Bancos", pbCampo, txtMovBancario(7), "Banco"
            Case 8 'Conta
                PCampo "Contas", "select * from Contas where Ctaati='S'", pbCampo, txtMovBancario(8), "C�digo"
            Case 13 'Centro de Custos
                PCampo "Centros de Custos", "Centros", pbCampo, txtMovBancario(13), "C�digo"
            Case 15 'Opera��es Cont�beis
                PCampo "Opera��o Cont�bil", "OperacaoContabil", pbCampo, txtMovBancario(Index), "cd_operacao"
        End Select
    End If
End Sub

Private Sub txtMovBancario_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0 'C�digo
            SetMascara KeyAscii, txtMovBancario(0).SelStart, InputMask(mrstMovBancario, "C�digo")
        Case 3 'Data
            SetMascara KeyAscii, txtMovBancario(3).SelStart, MASK_DATA
        Case 6 'Valor
            DValor KeyAscii
        Case 7 'Banco
            SetMascara KeyAscii, txtMovBancario(7).SelStart, fMask("Bancos", "Banco")
        Case 8 'Conta
            SetMascara KeyAscii, txtMovBancario(8).SelStart, fMask("Contas", "Conta")
        Case 13 'Centros de Custo
            SetMascara KeyAscii, txtMovBancario(13).SelStart, fMask("Centros", "C�digo")
        Case 15 'Operacao Cont�bil
            SetMascara KeyAscii, txtMovBancario(15).SelStart, fMask("OperacaoContabil", "cd_operacao")
    End Select
End Sub

Private Sub txtMovBancario_LostFocus(Index As Integer)
    Dim seq          As New CSequenciadorApl
    Dim sSeq         As String
    Dim strProcura   As String
    Dim rstBanco     As Object
    
    If Index = 0 Then
        'ao sair do campo lan�amento vou tratar o autonumercao
        If EAddNew(mlngMovBancario) And CLngDef(txtMovBancario(Index).Text, -1) = 0 Then
            seq.Construtor
            sSeq = seq.PegaSequencial(SEQ_APL_MOVBANCARI0)
            Set seq = Nothing
            txtMovBancario(Index).Text = sSeq
        End If
        LibProc WL_EXIBIR
    End If
    If Index = 2 Then
        If lblDescMovBancario(2).Caption <> "" Then
            GetAssocValue "SELECT Raz�o, Apel FROM Empresas WHERE Apel = '" & txtMovBancario(2).Text & "'", lblDescMovBancario(2), txtMovBancario(2)
        Else
            txtMovBancario(2).Text = ""
        End If
        'pt. 79561 - Moacir Pfau(04/04/2008)
        If EAdicao(mlngMovBancario) Or (Not EAdicao(mlngMovBancario) And strToLng(txtMovBancario(7).Text) = 0 And strToLng(txtMovBancario(8).Text) = 0) Then
            strProcura = "SELECT Banco, Conta FROM Empresas WHERE Apel = '" & txtMovBancario(2).Text & "';"
            AbreRecordset rstBanco, strProcura
            txtMovBancario(7).Text = strToLng(GetValue(rstBanco, "Banco"))
            txtMovBancario(8).Text = strToLng(GetValue(rstBanco, "Conta"))
            FechaRecordset (rstBanco)
        End If
    End If
End Sub

'Data.......: 25/06/2007
'Autor......: Dulcino J�nior
'Descri��o..: Procedimento utilizado para sugerir a opera��o cont�bil de acordo
'               com o tipo global selecionado.
Private Sub sugereOperacao()
    Dim objMatrizDAO As New cMatrizContabilizacaoDAO
    Dim objMatriz    As cMatrizContabilizacao
    
    Set objMatriz = objMatrizDAO.Carregar(cboMovBancario.Text)
    If Not objMatriz Is Nothing Then
        If gstrPagRec = "P" Then
            mlngOperacao = objMatriz.bancoSaida
        Else
            mlngOperacao = objMatriz.bancoEntrada
        End If
    Else
        mlngOperacao = 0
    End If
    txtMovBancario(15).Text = mlngOperacao
    Set objMatrizDAO = Nothing
    Set objMatriz = Nothing
End Sub

Private Function ProximoGapDeNumero(strPagRec As String) As Long
    Dim strSql As String
    Dim rstResult As Object
    Dim seq          As New CSequenciadorApl
    Dim sSeq         As String
        
    #If FOXSQL = 1 Then
        strSql = ""
        strSql = strSql & "SELECT TOP 1 cont as NextGap "
        strSql = strSql & "FROM ( "
        strSql = strSql & "      SELECT C�digo, cont "
        strSql = strSql & "      FROM (SELECT C�digo , pagrec, ROW_NUMBER() OVER (ORDER BY [C�digo]) as cont "
        strSql = strSql & "            FROM (SELECT DISTINCT C�digo, pagRec "
        strSql = strSql & "                  FROM Lan�amentos "
        strSql = strSql & "                  WHERE PagRec = '" & strPagRec & "') as X) as Y  WHERE C�digo <> cont) as Z  "
        
        If AbreRecordset(rstResult, strSql) = WL_OK Then
            ProximoGapDeNumero = rstResult.Fields("NextGap").value
        Else
            ProximoGapDeNumero = 0
        End If
    #Else
        seq.Construtor
        ProximoGapDeNumero = seq.PegaSequencial(SEQ_APL_MOVBANCARI0)
        Set seq = Nothing
    #End If
End Function
