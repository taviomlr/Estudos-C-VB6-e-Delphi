VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmCadHistBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Históricos Bancários"
   ClientHeight    =   7380
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   9930
   Begin VB.Frame fraBotoes 
      Height          =   7365
      Left            =   8430
      TabIndex        =   17
      Top             =   -30
      Width           =   1485
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1380
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1770
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   600
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
               Picture         =   "frmCadHistBancario.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadHistBancario.frx":0352
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame frmLanc 
      Height          =   7365
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   8385
      Begin VB.Frame Frame 
         Height          =   6825
         Index           =   1
         Left            =   60
         TabIndex        =   18
         Top             =   480
         Width           =   8265
         Begin VB.Frame fraTipoOperacao 
            Caption         =   "Tipo de Operação"
            Height          =   675
            Left            =   60
            TabIndex        =   20
            Top             =   1260
            Width           =   4155
            Begin VB.OptionButton optCredito 
               Caption         =   "Crédito"
               Height          =   225
               Left            =   1560
               TabIndex        =   6
               Top             =   270
               Width           =   855
            End
            Begin VB.OptionButton optDebito 
               Caption         =   "Débito"
               Height          =   225
               Left            =   420
               TabIndex        =   5
               Top             =   270
               Width           =   975
            End
            Begin VB.OptionButton optAmbas 
               Caption         =   "Ambas"
               Height          =   225
               Left            =   2730
               TabIndex        =   7
               Top             =   270
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame Frame 
            Height          =   675
            Index           =   0
            Left            =   4230
            TabIndex        =   19
            Top             =   1260
            Width           =   3975
            Begin VB.CommandButton cmdNovoLanc 
               Caption         =   "N&ovo"
               Height          =   375
               Left            =   120
               TabIndex        =   8
               Top             =   180
               Width           =   1215
            End
            Begin VB.CommandButton cmdConfirmar 
               Caption         =   "&Confirmar"
               Height          =   375
               Left            =   1380
               TabIndex        =   9
               Top             =   180
               Width           =   1215
            End
            Begin VB.CommandButton cmdExcluirDescr 
               Caption         =   "E&xcluir"
               Height          =   375
               Left            =   2640
               TabIndex        =   10
               Top             =   180
               Width           =   1215
            End
         End
         Begin Fox.EBSText etxCodHist 
            Height          =   330
            Left            =   2325
            TabIndex        =   2
            Top             =   180
            Width           =   750
            _ExtentX        =   265
            _ExtentY        =   582
            TipoTexto       =   0
            MaxLength       =   9
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
         Begin Fox.EBSText etxDescricao 
            Height          =   330
            Left            =   2325
            TabIndex        =   3
            Top             =   540
            Width           =   5790
            _ExtentX        =   9419
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   60
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
         Begin Fox.EBSText etxCompDescricao 
            Height          =   330
            Left            =   2325
            TabIndex        =   4
            Top             =   900
            Width           =   5790
            _ExtentX        =   9419
            _ExtentY        =   582
            Tipo            =   4
            TipoTexto       =   0
            MaxLength       =   60
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResultado 
            Height          =   4740
            Left            =   60
            TabIndex        =   24
            Top             =   2010
            Width           =   8130
            _ExtentX        =   14340
            _ExtentY        =   8361
            _Version        =   393216
            FixedRows       =   0
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Complemento da descrição"
            Height          =   195
            Left            =   330
            TabIndex        =   23
            Top             =   990
            Width           =   1920
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Descrição no extrato"
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
            Left            =   480
            TabIndex        =   22
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código"
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
            Left            =   1665
            TabIndex        =   21
            Top             =   270
            Width           =   600
         End
      End
      Begin Fox.EBSText etxBanco 
         Height          =   330
         Left            =   2385
         TabIndex        =   1
         Top             =   150
         Width           =   780
         _ExtentX        =   1376
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
      End
      Begin VB.Label Label3 
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
         Left            =   1770
         TabIndex        =   16
         Top             =   240
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmCadHistBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CHAR_DEBITO = "D"
Private Const CHAR_CREDITO = "C"
Private Const CHAR_AMBOS = "A"

Private mbizCadHist     As BizCadHistBancario
Private mcolCadHist     As ColCadHistBancario
Private mblnAlteracao   As Boolean
Private mblnGravar      As Boolean

Private Sub cmdConfirmar_Click()
    Dim objVO   As New VoCadHistBancario
    Dim strMSG  As String
    Dim mbrResp As VbMsgBoxResult
    
    mbrResp = vbYes
    strMSG = ValidaObrigatorios
    If Len(strMSG) = 0 Then
        If mcolCadHist.DescricaoRepetida(etxDescricao.valorTexto) Then
            mbrResp = MsgBox("A descrição do extrato informada já existe na tabela abaixo. Deseja continuar?", vbQuestion + vbYesNo)
        End If
        If mbrResp = vbYes Then
            Call CarregaVO(objVO)
            If mcolCadHist.Find(objVO) > 0 Then
               Call mcolCadHist.update(objVO)
            Else
                Call mcolCadHist.add(objVO)
            End If
            CarregaGrid
            LimpaCampos
            etxCodHist.valorInteiro = mbizCadHist.NovoCodigo(mcolCadHist, etxBanco.valorInteiro)
            etxDescricao.SetFocus
        End If
    Else
        MsgBox strMSG, vbInformation
    End If
    mblnGravar = False

End Sub

Private Sub cmdExcluir_Click()
    Dim objDAO As DaoExtratoBancario
    Dim blnPodeExcluir As Boolean
    Dim i As Integer
    
    Set objDAO = New DaoExtratoBancario
    
    blnPodeExcluir = True
    If mcolCadHist.Count > 0 And etxBanco.valorInteiro <> 0 Then
        For i = 1 To grdResultado.Rows - 1
            If objDAO.ExisteExtratoVinculado(grdResultado.TextMatrix(i, 2), etxBanco.valorInteiro) Then
                blnPodeExcluir = False
                Exit For
            End If
        Next
        If blnPodeExcluir Then
            If MsgBox("Deseja excluir o histórico deste banco?", vbQuestion + vbYesNo) = vbYes Then
                Call mcolCadHist.Clear
                If mbizCadHist.SalvaColecao(mcolCadHist, etxBanco.valorInteiro) Then
                    MsgBox "Registro excluído com sucesso.", vbInformation
                Else
                    MsgBox "Erro ao excluir registro.", vbCritical
                End If
                LimpaTodosCampos
            End If
        Else
            MsgBox "Não é possível excluir histórico com vinculação a um extrato.", vbInformation, "Atenção"
        End If
    Else
        MsgBox "Não há registros para serem excluídos.", vbInformation
    End If
End Sub

Private Sub cmdExcluirDescr_Click()
    Dim objVO As VoCadHistBancario
    Dim objDAO As DaoExtratoBancario
    
    If grdResultado.Rows = 2 And etxCodHist.valorInteiro > 0 And etxDescricao.valorTexto <> "" Then
        cmdExcluir_Click
        Exit Sub
    End If
    
    Set objDAO = New DaoExtratoBancario
    
    If etxCodHist.valorInteiro > 0 And etxDescricao.valorTexto <> "" Then
        If Not objDAO.ExisteExtratoVinculado(etxCodHist.valorInteiro, etxBanco.valorInteiro) Then
            With grdResultado
               If Len(.TextMatrix(.Row, 1)) <> 0 Then
                    Set objVO = mcolCadHist.GetItem(.TextMatrix(.Row, 1), etxCodHist.valorInteiro)
                    Call mcolCadHist.Remove(objVO)
                    CarregaGrid
               End If
            End With
        Else
            MsgBox "Não é possível excluir histórico vinculado a extrato.", vbInformation, "Atenção"
        End If
    Else
        MsgBox "Favor selecionar um histórico para excluir.", vbInformation, "Atenção"
    End If
    LimpaCampos
    mblnGravar = False
    
End Sub

Private Sub cmdGravar_Click()
    If etxBanco.valorInteiro <> 0 And mcolCadHist.Count <> 0 Then
        If etxCodHist.valorInteiro > 0 And etxDescricao.valorTexto <> "" Then
            If MsgBox("Existe um histórico a confirmar." & Chr(13) & "Deseja desconsiderá-lo e salvar?", vbQuestion + vbYesNo) = vbYes Then
                If mbizCadHist.SalvaColecao(mcolCadHist, etxBanco.valorInteiro) Then
                    MsgBox "Registros gravados com sucesso.", vbInformation
                Else
                    MsgBox "Erro ao gravar registros.", vbCritical
                End If
            End If
        Else
            If mbizCadHist.SalvaColecao(mcolCadHist, etxBanco.valorInteiro) Then
                MsgBox "Registros gravados com sucesso.", vbInformation
            Else
                MsgBox "Erro ao gravar registros.", vbCritical
            End If
        End If
    Else
        MsgBox "Preencha os campos obrigatórios e a tabela de históricos para gravar."
    End If
    mblnGravar = True
End Sub

Private Sub cmdNovo_Click()
    LimpaTodosCampos
    etxCodHist.valorInteiro = mbizCadHist.NovoCodigo(mcolCadHist, etxBanco.valorInteiro)
End Sub

Private Sub cmdNovoLanc_Click()
    LimpaCamposLancamento
    CarregaGrid
End Sub

Private Sub etxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        PCampo "Bancos", "Bancos", PB_CAMPO, etxBanco, "Banco"
        If etxBanco.valorInteiro <> 0 Then
            etxBanco_LostFocus
        End If
    End If
End Sub

Private Sub etxBanco_LostFocus()
    CarregaLancamentosBanco
End Sub

Private Sub Form_Load()
    IniciaEBSTexts
    Set mcolCadHist = New ColCadHistBancario
    Set mbizCadHist = New BizCadHistBancario
    Call CarregaGrid
    mblnGravar = True
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
        .Cols = 5
        .FixedCols = 1
        
        If intCont > 0 Then
            .Rows = 1
        Else
            .Rows = 2
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 120
                
        .TextMatrix(0, 1) = "Banco"
        .ColWidth(1) = 800
        .ColAlignment(1) = flexAlignCenterCenter
        
        .TextMatrix(0, 2) = "Código"
        .ColWidth(2) = 1200
        .ColAlignment(2) = flexAlignRightCenter
        
        .TextMatrix(0, 3) = "Descrição"
        .ColWidth(3) = 4350
        .ColAlignment(3) = flexAlignLeftCenter
        
        .TextMatrix(0, 4) = "Débito/Crédito"
        .ColWidth(4) = 1150
        .ColAlignment(4) = flexAlignLeftCenter
    End With
End Sub

Private Sub cmdSair_Click()
    If Not mblnGravar Then
        If MsgBox("Os históricos não foram gravados. Tem certeza que deseja sair?", vbYesNo, "Cadastro de Históricos") = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub CarregaGrid()
    Dim objVO   As VoCadHistBancario
    Dim strItem As String
    Dim i       As Integer
    Dim intUltimoCdHistorico As Integer

On Error GoTo Erro
    grdResultado.Clear
    CarregaHeaderGrid mcolCadHist.Count
    If Not mcolCadHist Is Nothing Then
        If mcolCadHist.Count > 0 Then
            mcolCadHist.MoveFirst
            While Not mcolCadHist.EOF
                Set objVO = mcolCadHist.CurrentObject
                With objVO
                    strItem = vbTab & .CdBanco & vbTab & .CdHistorico & vbTab & .DescricaoExtrato & vbTab & .TipoOperacao
                    grdResultado.AddItem strItem
                    intUltimoCdHistorico = .CdHistorico
                End With
                Set objVO = Nothing
                mcolCadHist.MoveNext
            Wend
        End If
    End If
    etxCodHist.valorInteiro = intUltimoCdHistorico + 1
    grdResultado.FixedRows = 1
    Exit Sub
Erro:
    MsgBox "Erro ao carregar tabela: " & err.Description
End Sub

Private Function ValidaObrigatorios() As String
    Dim strMSG  As String
    
    strMSG = vbNullString
    If etxBanco.valorInteiro = 0 Then
        strMSG = strMSG & "O campo 'Banco' é obrigatório." & vbCrLf
    End If
    
    If etxCodHist.valorInteiro = 0 Then
        strMSG = strMSG & "O campo 'Código' é obrigatório." & vbCrLf
    End If
    
    If Len(Trim(etxDescricao.valorTexto)) = 0 Then
        strMSG = strMSG & "O campo 'Descrição no extrato' é obrigatório." & vbCrLf
    End If
    ValidaObrigatorios = strMSG
End Function

Private Sub CarregaVO(ByRef objVO As VoCadHistBancario)

    With objVO
        .EnterpriseId = ModGeral.EnterpriseId
        .CdEstabelecimento = ModGeral.CdEstabelecimento
        .CdBanco = etxBanco.valorInteiro
        .CdHistorico = etxCodHist.valorInteiro
        .DescricaoExtrato = Trim(etxDescricao.valorTexto)
        .ComplementoDescricao = Trim(etxCompDescricao.valorTexto)
        .TipoOperacao = GetOperacaoChar()
    End With
End Sub

Private Function GetOperacaoChar() As String

    If optDebito.value = True Then
        GetOperacaoChar = CHAR_DEBITO
    ElseIf optCredito.value = True Then
        GetOperacaoChar = CHAR_CREDITO
    Else
        GetOperacaoChar = CHAR_AMBOS
    End If
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

Private Sub Form_Unload(Cancel As Integer)
    Set mcolCadHist = Nothing
    Set mbizCadHist = Nothing
End Sub

Private Sub grdResultado_DblClick()
    CarregaDados
End Sub

Private Sub CarregaDados()
    Dim objVO      As VoCadHistBancario

    If mcolCadHist.Count > 0 Then
        With grdResultado
            Set objVO = mcolCadHist.GetItem(.TextMatrix(.Row, 1), .TextMatrix(.Row, 2))
        End With
        
        If Not objVO Is Nothing Then
            With objVO
                etxBanco.valorInteiro = .CdBanco
                etxCodHist.valorInteiro = .CdHistorico
                etxDescricao.valorTexto = .DescricaoExtrato
                etxCompDescricao.valorTexto = .ComplementoDescricao
                If .TipoOperacao = CHAR_CREDITO Then
                    optCredito.value = True
                ElseIf .TipoOperacao = CHAR_DEBITO Then
                    optDebito.value = True
                Else
                    optAmbas.value = True
                End If
                mblnAlteracao = True
            End With
        End If
    End If
End Sub

Private Sub LimpaCampos()
    etxDescricao.Clear
    etxCompDescricao.Clear
    mblnAlteracao = False
End Sub

Private Sub LimpaTodosCampos()
    LimpaCampos
    etxBanco.Clear
    etxCodHist.Clear
    mcolCadHist.Clear
    etxCodHist.valorInteiro = 0
    CarregaGrid
    mblnGravar = True
End Sub
Private Sub LimpaCamposLancamento()
    LimpaCampos
End Sub
Public Sub CarregaLancamentosBanco()
    If etxBanco.valorInteiro > 0 Then
        Set mcolCadHist = mbizCadHist.carregarColecao(etxBanco.valorInteiro)
        CarregaGrid
        etxCodHist.valorInteiro = mbizCadHist.NovoCodigo(mcolCadHist, etxBanco.valorInteiro)
    End If
End Sub
