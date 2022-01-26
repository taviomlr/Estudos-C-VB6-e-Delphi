VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD_old.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmLancamentoDuplicataRateio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rateio"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9135
   Begin VB.Frame Frame 
      Height          =   5955
      Index           =   1
      Left            =   7710
      TabIndex        =   21
      Top             =   0
      Width           =   1365
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   11
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   930
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   12
         Top             =   1710
         Width           =   1185
      End
      Begin VB.CommandButton cmdRatear 
         Caption         =   "&Ratear"
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   2490
         Width           =   1185
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   540
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdicionar 
         Caption         =   "&Adicionar"
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   150
         Width           =   1185
      End
      Begin ComctlLib.ImageList imgRateio 
         Left            =   330
         Top             =   1920
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
               Picture         =   "frmLancamentoDuplicataRateio.frx":0000
               Key             =   "Checked"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmLancamentoDuplicataRateio.frx":005E
               Key             =   "Unchecked"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5955
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7695
      Begin VB.Frame Frame 
         Height          =   3675
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   2220
         Width           =   7545
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgGrid 
            Height          =   3375
            Left            =   60
            TabIndex        =   24
            Top             =   180
            Width           =   7380
            _ExtentX        =   13018
            _ExtentY        =   5953
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   0
            ScrollBars      =   2
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame frmRateio 
         Caption         =   "Rateio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   90
         TabIndex        =   14
         Top             =   120
         Width           =   7545
         Begin Fox.EBSText etxCentroCustoRateio 
            Height          =   330
            Left            =   1365
            TabIndex        =   0
            Top             =   330
            Width           =   3300
            _ExtentX        =   437171
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Centros"
            TamanhoDescricao=   2100
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
         Begin Fox.EBSText etxContaFinancRateio 
            Height          =   330
            Left            =   1365
            TabIndex        =   1
            Top             =   750
            Width           =   3300
            _ExtentX        =   437171
            _ExtentY        =   582
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Código"
            TipoCriterio    =   4
            CampoDescricao  =   "Descrição"
            TabelaConsulta  =   "Contas"
            TamanhoDescricao=   2100
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
         Begin Fox.EBSText etxPorcentagemRateio 
            Height          =   330
            Left            =   1350
            TabIndex        =   2
            Top             =   1200
            Width           =   1215
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   2
            CasasDecimais   =   2
            MaxLength       =   15
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
         Begin Fox.EBSText etxValorRateio 
            Height          =   330
            Left            =   5580
            TabIndex        =   3
            Top             =   330
            Width           =   1845
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   2
            CasasDecimais   =   2
            MaxLength       =   16
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
         Begin Fox.EBSText etxAbatimentoRateio 
            Height          =   330
            Left            =   5580
            TabIndex        =   5
            Top             =   1200
            Width           =   1845
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   2
            CasasDecimais   =   2
            MaxLength       =   16
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
         Begin Fox.EBSText etxAcrescimoRateio 
            Height          =   330
            Left            =   5580
            TabIndex        =   4
            Top             =   780
            Width           =   1845
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   2
            CasasDecimais   =   2
            MaxLength       =   16
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
         Begin Fox.EBSText etxSaldoRestanteRateio 
            Height          =   330
            Left            =   5580
            TabIndex        =   7
            Top             =   1620
            Width           =   1845
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   2
            CasasDecimais   =   2
            MaxLength       =   16
            Enabled         =   0   'False
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
         Begin VB.Label lblSaldoRestanteRateio 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Restante:"
            Height          =   195
            Left            =   4380
            TabIndex        =   23
            Top             =   1680
            Width           =   1140
         End
         Begin VB.Label lblValorRateio 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Left            =   5100
            TabIndex        =   20
            Top             =   420
            Width           =   405
         End
         Begin VB.Label lblAbatimentoRateio 
            AutoSize        =   -1  'True
            Caption         =   "Abatimento:"
            Height          =   195
            Left            =   4680
            TabIndex        =   19
            Top             =   1260
            Width           =   840
         End
         Begin VB.Label lblAcrescimoRateio 
            AutoSize        =   -1  'True
            Caption         =   "Acréscimo:"
            Height          =   195
            Left            =   4740
            TabIndex        =   18
            Top             =   840
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Porcentagem:"
            Height          =   195
            Left            =   315
            TabIndex        =   17
            Top             =   1260
            Width           =   990
         End
         Begin VB.Label lblContaFinancRateio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Conta Financ.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   45
            TabIndex        =   16
            Top             =   810
            Width           =   1260
         End
         Begin VB.Label lblCentroCustoRateio 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C. Custo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   510
            TabIndex        =   15
            Top             =   390
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frmLancamentoDuplicataRateio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjVo              As VoLancamentoDuplicata
Private mCol                As colRateio
Private mVoRateio           As VoRateio
Private mblnAlterando       As Boolean
Private mblnRateio          As Boolean

Public Enum ENUMRateio
    e_Centro = 1
    e_Conta = 2
    e_Porcentual = 3
    e_Valor = 4
    e_Acrescimo = 5
    e_Abatimento = 6
End Enum
Private Const strGrid = "campo=SeqGrid;label=;tamanho=250|" & _
                           "campo=Centro;label=Centro;tamanho=800;tipo=tpColGridInteger|" & _
                           "campo=Conta;label=Conta;tamanho=800;tipo=tpColGridInteger|" & _
                           "campo=Percentual;label=Porcentual;tamanho=1000;formato=###,##0.00;tipo=tpColGridInteger|" & _
                           "campo=Valor;label=Valor;tamanho=1500;formato=###,##0.00;tipo=tpColGridInteger|" & _
                           "campo=Acrescimo;label=Acrescimo;tamanho=1500;formato=###,##0.00;tipo=tpColGridInteger|" & _
                           "campo=Abatimento;label=Abatimento;tamanho=1500;formato=###,##0.00;tipo=tpColGridInteger"

Private Sub LimparCampos()
    etxCentroCustoRateio.Clear
    etxContaFinancRateio.Clear
    etxPorcentagemRateio.Clear
    etxValorRateio.Clear
    etxAbatimentoRateio.Clear
    etxAcrescimoRateio.Clear
End Sub

Private Sub cmdAdicionar_Click()
    Call LibProc(WL_ADICIONAR)
End Sub

Private Sub cmdAjuda_Click()
    Call LibProc(WL_AJUDA)
End Sub

Private Sub cmdCancelar_Click()
    Call LibProc(WL_CANCELAR)
End Sub

Private Sub cmdExcluir_Click()
    Call LibProc(WL_DELETAR)
End Sub

Private Sub cmdRatear_Click()
    Call LibProc(WL_PROCESSO)
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
End Sub

Private Sub etxCentroCustoRateio_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10575 - João Henrique(30/03/2012)
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Código, Descrição, [Data Limite],[cd_conta_contabil], [cd_centro_crd] " _
        & "FROM Centros"
        PCampo "C.Custo", strSql, PB_CAMPO, etxCentroCustoRateio, "Código"
    End If
End Sub

Private Sub etxContaFinancRateio_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql  As String
    
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10575 - João Henrique(30/03/2012)
    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Contas.Código as Conta, Contas.Descrição as [Descrição da Conta], " _
        & "Grupos.Código as Grupo, Grupos.Descrição as [Descrição do Grupo] " _
        & "FROM Grupos INNER JOIN Contas ON Grupos.Código = Contas.Grupo where Contas.Ctaati='S' " _
        & "ORDER BY Grupos.Código,Contas.Código"
        PCampo "Conta Financ.", strSql, PB_CAMPO, etxContaFinancRateio, "Conta"
    End If
End Sub

Private Sub fgGrid_DblClick()
    'Projeto: #1203 - História: #10564 - Problema#11881 - João Henrique(12/04/2012)
    If itemSelecionado(e_Conta) <> "" Then
        Set mVoRateio = mCol.GetItem(CLng(itemSelecionado(e_Conta)), CLng(itemSelecionado(e_Centro)))
        If Not mVoRateio Is Nothing Then
            Call carregaRegistro
        End If
    End If
End Sub

Private Function itemSelecionado(col As ENUMRateio) As String
    With fgGrid
       If .Row > 0 Then
          If .TextMatrix(.Row, col) <> "" Then
            itemSelecionado = .TextMatrix(.Row, col)
          Else
            itemSelecionado = ""
          End If
       End If
    End With
End Function

Private Sub Form_Load()
    
    Aplicacao.Connect
    'Projeto: #1203 - História: #10564 - Desenvolvimento#10575 - João Henrique(30/03/2012)
    Call etxCentroCustoRateio.AddConexao(Aplicacao)
    Call etxContaFinancRateio.AddConexao(Aplicacao)
    
    Aplicacao.Disconnect
    
    Set mCol = mobjVo.Col_Rateio
    
    Call CarregaHFlexGrid(fgGrid, , strGrid, , , mCol)
    'Projeto: #1203 - História: #10582 - Desenvolvimento#10595 - João Henrique(10/04/2012)
    Call calculoVlrSaldoRestanteRateio
End Sub

Public Function LibProc(strFuncao As String) As Boolean
    Dim biz                 As BizLancamentoDuplicata
    Dim blnGravar           As Boolean
    Dim strCodigo           As String
    Dim lngParcela          As Long
    Dim strTipo             As String
    Dim strEmpresa          As String
    Dim enumPagRec          As enuPagRec
    
On Error GoTo err
    
    Set biz = New BizLancamentoDuplicata
    
    
    Select Case strFuncao
        Case WL_SAIR
            Unload Me
            Exit Function

        Case WL_ADICIONAR
            Call fcarregaClasse
            If fValidaCampos() Then
                If mblnAlterando Then
                    Call mCol.update(mVoRateio)
                Else
                    Call mCol.add(mVoRateio)
                End If
                Call calculoVlrSaldoRestanteRateio
                mblnAlterando = False
                Call CarregaGrid
                etxCentroCustoRateio.SetFocus
                LimparCampos
            End If
            
        Case WL_DELETAR
            Call fcarregaClasse
            If mCol.Remove(mVoRateio) Then
                Call CarregaHFlexGrid(fgGrid, , strGrid, , , mCol)
                Call calculoVlrSaldoRestanteRateio
                LimparCampos
                mblnAlterando = False
            End If
        Case WL_CANCELAR
            Call NovoRegistro
            mblnAlterando = False
            
        Case WL_AJUDA
            Dim oHelpHtml As New clsHelp
            
            oHelpHtml.Origem = 0
            oHelpHtml.hWnd = Me.hWnd
            oHelpHtml.HelpContext = Me.HelpContextID
            Call oHelpHtml.ShowHelp
            Set oHelpHtml = Nothing
                
        Case WL_PROCESSO
            strCodigo = mobjVo.Codigo_Nota: lngParcela = mobjVo.Parcela: strTipo = mobjVo.Tipo: strEmpresa = mobjVo.Empresa
            If mobjVo.PagRec = "R" Then
                enumPagRec = Recebimento
            Else
                enumPagRec = Pagamento
            End If
            If biz.ProcessoRateioLancamentoDuplicata(mCol, mobjVo) Then
                mblnRateio = True
                Call AbreTelaMessengerBox(Ok, "Rateio realizado com sucesso.", NomeModulo, True)
                Call frmLancamentoDuplicata.CarregarLancamentoDuplicataOutrasRotinas(strCodigo, strTipo, lngParcela, strEmpresa, enumPagRec, mobjVo.LancDup)
                Unload Me
            End If
    End Select
    Exit Function
err:
End Function

Private Sub NovoRegistro()
    LimparCampos
    mblnAlterando = False
End Sub

Private Function fValidaCampos() As Boolean
    Dim objBiz              As New BizLancamentoDuplicata
    Dim col                 As New Collection
    Dim colTemp             As New colRateio

    If mCol.Count > 0 Then
        mCol.MoveFirst
        While Not mCol.EOF
            Call colTemp.add(mCol.CurrentObject)
            mCol.MoveNext
        Wend
    End If
    
    If mblnAlterando Then
        Call colTemp.Remove(mVoRateio)
    End If
    
    Call objBiz.validarCampoObrigatorioRateio(etxContaFinancRateio.valorInteiro, etxCentroCustoRateio.valorInteiro, _
                                              etxPorcentagemRateio.valorDecimal, etxValorRateio.valorDecimal, col)
    
    Call objBiz.validarCampoDiversosRateio(etxPorcentagemRateio.valorDecimal, etxValorRateio.valorDecimal, _
                                           etxAcrescimoRateio.valorDecimal, etxAbatimentoRateio.valorDecimal, _
                                           colTemp.Percentual, colTemp.Valor, etxContaFinancRateio.valorInteiro, _
                                           etxCentroCustoRateio.valorInteiro, colTemp, mobjVo.ValorTotal, mVoRateio.ValorTotal, col)
  
    
    'Exibe eventuais mensagens para o usuário.
    fValidaCampos = AbreTelaMensagem(col, True)

    Set objBiz = Nothing
    Set colTemp = Nothing
End Function

'Projeto: #1203 - História: #10564 - Desenvolvimento#10575 - João Henrique(05/04/2012)
Private Sub calculoVlrSaldoRestanteRateio()
    etxSaldoRestanteRateio.valorDecimal = Round(mobjVo.ValorTotal - calculoVlrRateio, 2)
End Sub

'Projeto: #1203 - História: #10564 - Desenvolvimento#10575 - João Henrique(05/04/2012)
Private Function calculoVlrRateio() As Double
    calculoVlrRateio = mCol.Valor + mCol.Acrescimo - mCol.Abatimento
End Function

Public Function ObjetoVo(ByRef Valor As VoLancamentoDuplicata)
    Set mobjVo = Valor
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mobjVo.Col_Rateio = mCol
    Call frmLancamentoDuplicata.setVo(mobjVo, mblnRateio)
    Set mCol = Nothing
End Sub

Public Sub CarregaGrid()
    Call CarregaHFlexGrid(fgGrid, , strGrid, , , Nothing)
    If Not mCol Is Nothing Then
        If mCol.EOF Then
            mCol.MoveFirst
        End If
        Call CarregaHFlexGrid(fgGrid, , strGrid, , , mCol)
    End If
End Sub

Private Sub fcarregaClasse()
    Set mVoRateio = New VoRateio
    With mVoRateio
        .Centro = etxCentroCustoRateio.valorInteiro
        .conta = etxContaFinancRateio.valorInteiro
        .Abatimento = etxAbatimentoRateio.valorDecimal
        .Acrescimo = etxAcrescimoRateio.valorDecimal
        .Percentual = etxPorcentagemRateio.valorDecimal
        .Valor = etxValorRateio.valorDecimal
    End With
End Sub

Private Sub carregaRegistro()
    mblnAlterando = True
    With mVoRateio
        etxCentroCustoRateio.valorInteiro = .Centro
        etxContaFinancRateio.valorInteiro = .conta
        etxPorcentagemRateio.valorDecimal = .Percentual
        etxValorRateio.valorDecimal = .Valor
        etxAbatimentoRateio.valorDecimal = .Abatimento
        etxAcrescimoRateio.valorDecimal = .Acrescimo
    End With
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
