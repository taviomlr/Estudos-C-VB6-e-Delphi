VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmImpDigExtratoBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar/Digitar Extrato Bancário"
   ClientHeight    =   8190
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   10470
   Begin VB.Frame fraGeral 
      Height          =   8175
      Left            =   60
      TabIndex        =   21
      Top             =   -30
      Width           =   8895
      Begin VB.CommandButton cmdImpExtrato 
         Caption         =   "&Importar Extrato"
         Height          =   375
         Left            =   6960
         TabIndex        =   2
         Top             =   210
         Width           =   1845
      End
      Begin VB.Frame fraLanc 
         Caption         =   "Lançamentos"
         Enabled         =   0   'False
         Height          =   6810
         Left            =   90
         TabIndex        =   24
         Top             =   1290
         Width           =   8745
         Begin VB.CommandButton cmdCadHist 
            Caption         =   "..."
            Height          =   285
            Left            =   2070
            TabIndex        =   35
            Top             =   630
            Width           =   255
         End
         Begin VB.Frame Frame 
            Height          =   1875
            Left            =   7200
            TabIndex        =   32
            Top             =   120
            Width           =   1455
            Begin VB.CommandButton cmdExcluirLanc 
               Caption         =   "E&xcluir"
               Height          =   375
               Left            =   120
               TabIndex        =   14
               Top             =   930
               Width           =   1215
            End
            Begin VB.CommandButton cmdConfirmar 
               Caption         =   "&Confirmar"
               Height          =   375
               Left            =   120
               TabIndex        =   13
               Top             =   540
               Width           =   1215
            End
            Begin VB.CommandButton cmdNovoLanc 
               Caption         =   "N&ovo"
               Height          =   375
               Left            =   120
               TabIndex        =   12
               Top             =   150
               Width           =   1215
            End
         End
         Begin VB.Frame fraTipoOperacao 
            Caption         =   "Tipo de Operação"
            Height          =   735
            Left            =   3420
            TabIndex        =   31
            Top             =   1260
            Width           =   3705
            Begin VB.OptionButton optCredito 
               Caption         =   "Crédito"
               Height          =   225
               Left            =   2280
               TabIndex        =   11
               Top             =   330
               Width           =   855
            End
            Begin VB.OptionButton optDebito 
               Caption         =   "Débito"
               Height          =   225
               Left            =   690
               TabIndex        =   10
               Top             =   330
               Value           =   -1  'True
               Width           =   1275
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResultado 
            Height          =   4695
            Left            =   75
            TabIndex        =   25
            Top             =   2040
            Width           =   8580
            _ExtentX        =   15134
            _ExtentY        =   8281
            _Version        =   393216
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin Fox.EBSText etxHistorico 
            Height          =   330
            Left            =   1245
            TabIndex        =   6
            Top             =   570
            Width           =   795
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
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
         Begin Fox.EBSText etxDescricao 
            Height          =   330
            Left            =   1245
            TabIndex        =   7
            Top             =   930
            Width           =   5865
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
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
         Begin Fox.EBSText etxDocumento 
            Height          =   330
            Left            =   1245
            TabIndex        =   8
            Top             =   1290
            Width           =   2040
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
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
         Begin Fox.EBSText etxValor 
            Height          =   330
            Left            =   1245
            TabIndex        =   9
            Top             =   1650
            Width           =   2010
            _ExtentX        =   265
            _ExtentY        =   582
            Tipo            =   2
            CasasDecimais   =   2
            TipoTexto       =   0
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
            ExibeDescricao  =   0   'False
         End
         Begin Fox.EBSText etxDia 
            Height          =   330
            Left            =   1245
            TabIndex        =   5
            Top             =   210
            Width           =   780
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   2
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
         Begin VB.Label lblDescricaoHistorico 
            Height          =   225
            Left            =   2430
            TabIndex        =   33
            Tag             =   "Desc"
            Top             =   630
            UseMnemonic     =   0   'False
            Width           =   4485
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor"
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
            Left            =   210
            TabIndex        =   30
            Top             =   1740
            Width           =   975
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
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
            Left            =   210
            TabIndex        =   29
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Desc. Extrato"
            Height          =   195
            Left            =   225
            TabIndex        =   28
            Top             =   1020
            Width           =   960
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Histórico"
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
            Left            =   210
            TabIndex        =   27
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Dia"
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
            Left            =   210
            TabIndex        =   26
            Top             =   300
            Width           =   975
         End
      End
      Begin Fox.EBSData edtEmissaoInicial 
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Top             =   570
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         Formato         =   1
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
      Begin Fox.EBSText etxBancoInicial 
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         TipoTexto       =   0
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
      End
      Begin Fox.EBSText etxExtrato 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   1275
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
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
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Origem da chamada da tela"
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
         Left            =   6090
         TabIndex        =   37
         Top             =   810
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label lblOrigemConciliacao 
         Caption         =   "0"
         Height          =   315
         Left            =   8610
         TabIndex        =   36
         Top             =   810
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Extrato"
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
         Left            =   615
         TabIndex        =   34
         Top             =   1010
         Width           =   650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Mês/Ano"
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
         Index           =   1
         Left            =   285
         TabIndex        =   23
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Banco"
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
         Left            =   270
         TabIndex        =   22
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Frame fraBotoes 
      Height          =   8175
      Left            =   9000
      TabIndex        =   0
      Top             =   -30
      Width           =   1455
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1380
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1770
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
               Picture         =   "frmImpDigExtratoBancario.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmImpDigExtratoBancario.frx":0352
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmImpDigExtratoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CHAR_DEBITO = "D"
Private Const CHAR_CREDITO = "C"

Private mbizExtratoBanc     As BizImpDigExtratoBancario
Private mcolLancamentos     As New ColImpDigExtratoBancario
Private mlngSeq             As Long
Private mblnAlterando        As Boolean
Private mblnGravar        As Boolean

Private Sub cmdCadHist_Click()
    frmCadHistBancario.etxBanco.valorInteiro = etxBancoInicial.valorInteiro
    frmCadHistBancario.CarregaLancamentosBanco
    Call mostrarForm(frmCadHistBancario, frmCadHistBancario.HelpContextID, False)
End Sub

Private Sub cmdConfirmar_Click()
    Dim voItem      As New VoImpDigExtratoBancario
    Dim strMSG      As String
    
    strMSG = ValidaObrigatorios
    If Len(strMSG) = 0 Then
        Call CarregaVO(voItem)
        'If mblnAlterando Then
        If mcolLancamentos.Find(voItem) > 0 Then
            Call mcolLancamentos.update(voItem)
        Else
            Call mcolLancamentos.add(voItem)
        End If
        Call CarregaGrid
        Call LimpaCampos
                
        Set voItem = New VoImpDigExtratoBancario
        cmdNovoLanc_Click
        etxDia.SetFocus
    Else
        MsgBox strMSG, vbInformation
    End If
    Set voItem = Nothing
    mblnGravar = False
End Sub

Private Sub cmdExcluir_Click()
    Dim objDaoDuplicLanc As DaoLancamentoDuplicata
    
    Set objDaoDuplicLanc = New DaoLancamentoDuplicata
    If etxBancoInicial.valorInteiro > 0 And IsDate(edtEmissaoInicial.MesAno) And etxExtrato.valorInteiro > 0 Then
        If Not objDaoDuplicLanc.ExisteExtratoConciliado("Duplicatas", etxExtrato.valorInteiro, etxBancoInicial.valorInteiro) Then
            If Not objDaoDuplicLanc.ExisteExtratoConciliado("Lançamentos", etxExtrato.valorInteiro, etxBancoInicial.valorInteiro) Then
                If MsgBox("Deseja excluir o histórico deste mês para este extrato e banco?", vbQuestion + vbYesNo) = vbYes Then
                    Call mcolLancamentos.Clear
                    If mbizExtratoBanc.SalvaColecao(mcolLancamentos, etxBancoInicial.valorInteiro, edtEmissaoInicial.MesAno, etxExtrato.valorInteiro) Then
                        MsgBox "Registro excluído com sucesso.", vbInformation
                    Else
                        MsgBox "Erro ao excluir registro.", vbCritical
                    End If
                Else
                    Exit Sub
                End If
            Else
                MsgBox "Não é possível excluir este extrato. " & vbNewLine & "Existe(m) lançamento(s) conciliado(s) a este extrato bancário.", vbInformation, "Exclusão de extrato bancário"
            End If
        Else
            MsgBox "Não é possível excluir este extrato. " & vbNewLine & "Existe(m) duplicata(s)/lançamento(s) conciliada(s) a este extrato bancário.", vbInformation, "Exclusão de extrato bancário"
        End If
    End If
    cmdNovo_Click
End Sub

Private Sub cmdExcluirLanc_Click()
    Dim objVO As VoImpDigExtratoBancario
    With grdResultado
        If .Rows = 2 And mlngSeq > 0 Then
            cmdExcluir_Click
            Exit Sub
        End If
        If mlngSeq > 0 Then
            Set objVO = mcolLancamentos.GetItem(etxBancoInicial.valorInteiro, mlngSeq)
            If Not objVO Is Nothing Then
                Call mcolLancamentos.Remove(objVO)
                CarregaGrid
            End If
        Else
            MsgBox "Favor selecionar um lançamento para excluir.", vbInformation, "Atenção"
        End If
    End With
    LimpaCampos
    mblnGravar = False
End Sub

Private Sub cmdGravar_Click()
    If etxBancoInicial.valorInteiro <> 0 And IsDate(edtEmissaoInicial.MesAno) Then
        If mbizExtratoBanc.SalvaColecao(mcolLancamentos, etxBancoInicial.valorInteiro, edtEmissaoInicial.MesAno, etxExtrato.valorInteiro) Then
            MsgBox "Extrato gravado com sucesso. Número: " & etxExtrato.valorInteiro, vbInformation, "Importar/Digitar Extrato Bancário"
        Else
            MsgBox "Erro ao gravar registros.", vbCritical
        End If
    Else
        MsgBox "Para gravar um extrato, insira um valor no campo 'Banco' e no campo 'Mês/Ano'", vbInformation
    End If
    mblnGravar = True
End Sub

Private Sub cmdImpExtrato_Click()
    If etxBancoInicial.valorInteiro > 0 Then
        Call mostrarForm(frmImpArqExtratoBancario, frmImpArqExtratoBancario.HelpContextID)
    Else
        MsgBox "Favor preencher o código do banco corretamente!", vbInformation, "Importar/Digitar Extrato Bancário"
    End If
End Sub

Private Sub cmdNovo_Click()
    LimpaTodosCampos
    Call CarregaHeaderGrid(mcolLancamentos.Count)
    etxBancoInicial.SetFocus
End Sub

Private Sub cmdNovoLanc_Click()
    LimpaCamposLancamento
End Sub

Private Sub cmdPesquisar_Click()
    Dim strSql As String
    Dim strMesAux As String
    Dim strAnoAux As String
    Dim lngExtratoAux As Long
    
    strSql = "SELECT FEB.CD_BANCO as [Codigo do Banco] , B.Nome, cd_extrato as Extrato, MONTH(FEB.data_extrato) as Mes , year(FEB.data_extrato) as Ano " & _
             "FROM FFIExtratoBancario as FEB INNER JOIN Bancos B on B.Banco = FEB.cd_banco " & _
             "GROUP BY FEB.cd_banco, B.Nome, MONTH(FEB.data_extrato) ,year(FEB.data_extrato), cd_extrato " & _
             "ORDER BY 1,3,2 "
    
    If PMultiCampo("Consulta - " & Me.Caption, strSql, pbCampo, "Codigo do Banco;Mes;Ano;Extrato", etxBancoInicial, strMesAux, strAnoAux, etxExtrato) Then
        edtEmissaoInicial.MesAno = Format(strMesAux, "00") & "/" & strAnoAux
        
        Call etxExtrato_LostFocus
    End If
End Sub

Private Sub edtEmissaoInicial_LostFocus()
    If etxBancoInicial.valorInteiro > 0 And IsDate(edtEmissaoInicial.MesAno) Then
        'CarregaGridLancamentos
        fraLanc.Enabled = True
    End If
End Sub


Private Sub etxBancoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        PCampo "Bancos", "Bancos", pbCampo, etxBancoInicial, "Banco"
    End If
End Sub

Private Sub etxBancoInicial_LostFocus()
    
    If etxBancoInicial.valorInteiro > 0 Then
        If Not ModGeral.ReadOnly Then cmdImpExtrato.Enabled = True
        'etxCodigo.valorInteiro = frmImpArqExtratoBancario.ProximoCodigoExtrato(etxBancoInicial.valorInteiro)
        If IsDate(edtEmissaoInicial.MesAno) Then
            'CarregaGridLancamentos
            fraLanc.Enabled = True
        End If
    End If
End Sub



Private Sub etxExtrato_LostFocus()
    If etxBancoInicial.valorInteiro > 0 And IsDate(edtEmissaoInicial.MesAno) And etxExtrato.valorInteiro > 0 Then
        CarregaGridLancamentos
        fraLanc.Enabled = True
    End If
End Sub

Private Sub etxHistorico_Change()
    CarregaHistorico
End Sub
Private Sub etxHistorico_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    
    If KeyCode = vbKeyPageDown Then
        If etxBancoInicial.valorInteiro <> 0 Then
            strSql = "SELECT cd_historico, descricao_extrato, complemento_descricao, tipo_operacao FROM FFIExtratoBancarioHistorico WHERE cd_banco = " & etxBancoInicial.valorInteiro & " and (tipo_operacao = '" & IIf(optDebito.value, "D", "C") & "' or tipo_operacao = 'A')"
            Call PMultiCampo("Históricos", strSql, pbCampo, "cd_historico;descricao_extrato", etxHistorico, lblDescricaoHistorico)
            Call etxHistorico_LostFocus
        End If
    End If
End Sub

Private Sub etxHistorico_LostFocus()
    CarregaHistorico
End Sub

Private Sub Form_Load()
    IniciaEBSTexts
    Set mbizExtratoBanc = New BizImpDigExtratoBancario
    Call CarregaGrid
    mblnGravar = True
    
    If ModGeral.ReadOnly Then cmdImpExtrato.Enabled = False
End Sub

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
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

Private Sub CarregaHeaderGrid(intCont As Integer)
    Dim intIndex As Long

    With grdResultado
        .Cols = 10
        .FixedCols = 1
        
        If intCont > 0 Then
            .Rows = 1
        Else
            .Rows = 2
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 120
                
        .TextMatrix(0, 1) = "Dia"
        .ColWidth(1) = 400
        .ColAlignment(1) = flexAlignCenterCenter
        
        .TextMatrix(0, 2) = "Histórico"
        .ColWidth(2) = 2200
        .ColAlignment(2) = flexAlignLeftCenter
        
        .TextMatrix(0, 3) = "Desc. Extrato"
        .ColWidth(3) = 2100
        .ColAlignment(3) = flexAlignLeftCenter
        
        .TextMatrix(0, 4) = "Documento"
        .ColWidth(4) = 1250
        .ColAlignment(4) = flexAlignLeftCenter
        
        .TextMatrix(0, 5) = "Valor"
        .ColWidth(5) = 900
        .ColAlignment(5) = flexAlignRightCenter
        
        .TextMatrix(0, 6) = "Débito/Crédito"
        .ColWidth(6) = 1150
        .ColAlignment(6) = flexAlignLeftCenter
        
        .TextMatrix(0, 7) = "Sequencial"
        .ColWidth(7) = 0    'Esconder a coluna
        
        .TextMatrix(0, 8) = "Cod. Historico"
        .ColWidth(8) = 0   'Esconder a coluna
        
        .TextMatrix(0, 9) = "Conciliado"
        .ColWidth(9) = 0   'Esconder a coluna
        
'        For intIndex = 0 To .Cols - 1
'            .TextMatrix(1, intIndex) = ""
'        Next
    End With
End Sub

Private Sub cmdSair_Click()
    If Not mblnGravar Then
        If MsgBox("Os lançamentos do extrato não foram gravados. Tem certeza que deseja sair?", vbYesNo, "Extrato Bancário") = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Public Sub CarregaGridLancamentos()
    Dim strOptTipo As String
    
    Dim blnMostraConciliado As Boolean
    
    blnMostraConciliado = True
    Set mcolLancamentos = mbizExtratoBanc.carregarColecao(etxBancoInicial.valorInteiro, edtEmissaoInicial.MesAno, etxExtrato.valorInteiro, blnMostraConciliado)
    If Not mcolLancamentos Is Nothing Then
        Call CarregaGrid
    End If
End Sub

Private Sub CarregaGrid()
    Dim objVO        As VoImpDigExtratoBancario
    Dim daoExtrato   As DaoImpDigExtratoBancario
    Dim strItem      As String
    Dim i            As Integer
    Dim strHistorico As String

On Error GoTo Erro
    grdResultado.Clear
    CarregaHeaderGrid mcolLancamentos.Count
    If Not mcolLancamentos Is Nothing Then
        If mcolLancamentos.Count > 0 Then
            mcolLancamentos.MoveFirst
            etxExtrato.valorInteiro = mcolLancamentos.CurrentObject.CdExtrato
            While Not mcolLancamentos.EOF
                Set objVO = mcolLancamentos.CurrentObject
                Set daoExtrato = New DaoImpDigExtratoBancario
                strHistorico = daoExtrato.BuscaDescricaoHistorico(objVO.CdBanco, objVO.CdHistorico)
                With objVO
                    strItem = vbTab & Day(.DataExtrato) & vbTab & strHistorico & vbTab & .Descricao & vbTab & .Documento & vbTab & Format(.Valor, "##,##0.00") & vbTab & .TipoOperacao & vbTab & .SeqLancExtrato & vbTab & .Conciliado
                    grdResultado.AddItem strItem
                End With
                Set objVO = Nothing
                mcolLancamentos.MoveNext
            Wend
        End If
    End If
    grdResultado.FixedRows = 1
    grdResultado.Sort = flexSortNumericAscending
    Exit Sub
Erro:
    MsgBox "Erro ao carregar tabela: " & err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mbizExtratoBanc = Nothing
    Set mcolLancamentos = Nothing
End Sub

Private Function CarregaVO(ByRef vo As VoImpDigExtratoBancario) As Boolean
        
    With vo
        .EnterpriseId = ModGeral.EnterpriseId
        .CdEstabelecimento = ModGeral.CdEstabelecimento
        .CdExtrato = IIf(etxExtrato.valorInteiro = 0, frmImpArqExtratoBancario.ProximoCodigoExtrato(etxBancoInicial.valorInteiro), etxExtrato.valorInteiro)
        .CdBanco = etxBancoInicial.valorInteiro
        .SeqLancExtrato = IIf(mlngSeq = 0, mcolLancamentos.Count + 1, mlngSeq)
        .CdHistorico = etxHistorico.valorInteiro
        .DescricaoHistorico = lblDescricaoHistorico.Caption
        .DataExtrato = Format(etxDia.valorInteiro & "/" & Month(edtEmissaoInicial.MesAno) & "/" & Year(edtEmissaoInicial.MesAno), "dd/MM/YYYY")
        .Descricao = Trim(etxDescricao.valorTexto)
        .Documento = etxDocumento.valorTexto
        .Valor = etxValor.valorDecimal
        .TipoOperacao = IIf(optDebito.value = True, "D", "C")
        .ValorInterno = 0 'Implementar
        .Conciliado = False 'Implementar
        .DataConciliacao = "00:00:00" 'Implementar
    End With

End Function

Private Function ValidaObrigatorios() As String
    Dim strMSG  As String
    
    strMSG = vbNullString
    If edtEmissaoInicial.MesAno = "00:00:00" Then
        strMSG = strMSG & "O campo 'Mês/Ano' é obrigatório." & vbCrLf
    ElseIf Not IsDate(edtEmissaoInicial.MesAno) Then
        strMSG = strMSG & "O campo 'Mês/Ano' não é uma data válida." & vbCrLf
    End If
    
    If etxBancoInicial.valorInteiro = 0 Then
        strMSG = strMSG & "O campo 'Banco' é obrigatório." & vbCrLf
    End If
    
    If etxDia.valorInteiro = 0 Then
        strMSG = strMSG & "O campo 'Dia' é obrigatório." & vbCrLf
    ElseIf Not DiaValido Then
        strMSG = strMSG & "O campo 'Dia' não é válido para o mês e o ano informado no campo 'Mês/Ano'." & vbCrLf
    End If
    
    If Trim(lblDescricaoHistorico.Caption) = "" Then
        strMSG = strMSG & "O campo 'Histórico' é obrigatório." & vbCrLf
    End If
    
    If Len(Trim(etxDocumento.valorTexto)) = 0 Then
        strMSG = strMSG & "O campo 'Documento' é obrigatório." & vbCrLf
    End If
    
    If etxValor.valorDecimal = 0 Then
        strMSG = strMSG & "O campo 'Valor' é obrigatório." & vbCrLf
    End If
    ValidaObrigatorios = strMSG
End Function

Private Function DiaValido() As Boolean
    DiaValido = IIf(IsDate(Format(etxDia.valorInteiro & "/" & Month(edtEmissaoInicial.MesAno) & "/" & Year(edtEmissaoInicial.MesAno), "dd/MM/YYYY")) = True, True, False)
End Function

Private Sub IniciaEBSTexts()
    Dim ctrl As Control

    Aplicacao.Connect
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "EBSText" Then
            Call ctrl.AddConexao(Aplicacao)
        End If
    Next
    Aplicacao.Disconnect
End Sub

Private Sub LimpaCampos()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "EBSText" Then
            If ctrl.Container.name = fraLanc.name Then
                ctrl.Clear
            End If
        End If
    Next
    lblDescricaoHistorico.Caption = ""
    mlngSeq = 0
    mblnAlterando = False
End Sub

Private Sub grdResultado_DblClick()
    mblnAlterando = True
    CarregaDados
End Sub

Private Sub CarregaDados()
    Dim objVO      As VoImpDigExtratoBancario
    If mcolLancamentos.Count > 0 Then
        With grdResultado
            If .TextMatrix(.Row, 8) Then
                cmdConfirmar.Enabled = False
                cmdExcluirLanc.Enabled = False
            Else
                cmdConfirmar.Enabled = True
                cmdExcluirLanc.Enabled = True
            End If
            Set objVO = mcolLancamentos.GetItem(etxBancoInicial.valorInteiro, .TextMatrix(.Row, 7))
        End With
        If Not objVO Is Nothing Then
            With objVO
                etxDia.valorInteiro = Day(.DataExtrato)
                etxHistorico.valorInteiro = .CdHistorico
                etxDescricao.valorTexto = .Descricao
                etxDocumento.valorTexto = .Documento
                etxValor.valorDecimal = .Valor
                mlngSeq = .SeqLancExtrato
                If .TipoOperacao = CHAR_CREDITO Then
                    optCredito.value = True
                ElseIf .TipoOperacao = CHAR_DEBITO Then
                    optDebito.value = True
                End If
            End With
        End If
    End If
End Sub

Private Sub LimpaTodosCampos()
    LimpaCampos
    
    mcolLancamentos.Clear
    grdResultado.Clear
    etxBancoInicial.Clear
    edtEmissaoInicial.Clear
    etxExtrato.Clear
    mlngSeq = 0
End Sub
Private Sub LimpaCamposLancamento()
    etxDia.Clear
    etxHistorico.Clear
    etxDescricao.Clear
    etxDocumento.Clear
    etxValor.Clear
    mblnAlterando = False
    cmdConfirmar.Enabled = True
    cmdExcluirLanc.Enabled = True
    mlngSeq = 0
End Sub

Private Sub optDebito_Click()
    Dim objVO   As VoCadHistBancario
    Dim objBiz  As New BizCadHistBancario
    
    If grdResultado.Rows = 0 Then
        etxHistorico.Clear
    ElseIf etxHistorico.valorInteiro > 0 Then
        Set objVO = objBiz.CarregarHistorico(etxBancoInicial.valorInteiro, etxHistorico.valorInteiro, IIf(optDebito.value, "D", "C"))
        If Not objVO Is Nothing Then
            lblDescricaoHistorico.Caption = objVO.DescricaoExtrato & IIf(Len(objVO.ComplementoDescricao) > 0, " " & objVO.ComplementoDescricao, vbNullString)
        Else
            lblDescricaoHistorico.Caption = ""
        End If
    End If
End Sub

Private Sub optCredito_Click()
    Dim objVO   As VoCadHistBancario
    Dim objBiz  As New BizCadHistBancario
    
    If grdResultado.Rows = 0 Then
        etxHistorico.Clear
    ElseIf etxHistorico.valorInteiro > 0 Then
        Set objVO = objBiz.CarregarHistorico(etxBancoInicial.valorInteiro, etxHistorico.valorInteiro, IIf(optDebito.value, "D", "C"))
        If Not objVO Is Nothing Then
            lblDescricaoHistorico.Caption = objVO.DescricaoExtrato & IIf(Len(objVO.ComplementoDescricao) > 0, " " & objVO.ComplementoDescricao, vbNullString)
        Else
            lblDescricaoHistorico.Caption = ""
        End If
    End If
End Sub
Private Sub CarregaHistorico()
    Dim objVO   As VoCadHistBancario
    Dim objBiz  As New BizCadHistBancario
    
    Set objVO = objBiz.CarregarHistorico(etxBancoInicial.valorInteiro, etxHistorico.valorInteiro, IIf(optDebito.value, "D", "C"), True)
    If Not objVO Is Nothing Then
        lblDescricaoHistorico.Caption = objVO.DescricaoExtrato & IIf(Len(objVO.ComplementoDescricao) > 0, " " & objVO.ComplementoDescricao, vbNullString)
    Else
        lblDescricaoHistorico.Caption = ""
    End If
End Sub
