VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmRetorno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carregar Retorno Bancário"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   Icon            =   "frmRetorno.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   8595
   Begin VB.Frame fraBotoes 
      Height          =   1515
      Left            =   7200
      TabIndex        =   7
      Top             =   -30
      Width           =   1350
      Begin VB.CommandButton cmdGerar 
         Caption         =   "&Carregar"
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   960
         Width           =   1185
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   570
         Width           =   1185
      End
      Begin ComctlLib.ImageList imgCheck 
         Left            =   660
         Top             =   900
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
               Picture         =   "frmRetorno.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmRetorno.frx":06DC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame 
      Height          =   1515
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   7155
      Begin Fox.EBSText etxBanco 
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   270
         Width           =   5385
         _ExtentX        =   132265
         _ExtentY        =   582
         MaxLength       =   9
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Banco"
         TipoCriterio    =   4
         CampoDescricao  =   "Nome"
         TabelaConsulta  =   "Bancos"
         TamanhoDescricao=   3500
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
      Begin Fox.EBSText etxCarteira 
         Height          =   330
         Left            =   1680
         TabIndex        =   4
         Top             =   630
         Width           =   5385
         _ExtentX        =   132265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   6
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "id_carteira"
         TipoCriterio    =   4
         CampoDescricao  =   "desc_carteira"
         TabelaConsulta  =   "FFICarteira"
         TamanhoDescricao=   3500
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
      Begin Fox.EBSArquivo etxCaminhoRetorno 
         Height          =   330
         Left            =   1665
         TabIndex        =   6
         Top             =   990
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   582
         TipoTratamento  =   2
         Filter          =   ""
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Caminho R&etorno"
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
         Left            =   150
         TabIndex        =   5
         Top             =   1065
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Carteira"
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
         Left            =   945
         TabIndex        =   3
         Top             =   705
         Width           =   675
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "&Banco/Conta"
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
         TabIndex        =   1
         Top             =   345
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngEnterpriseId                                    As Long
Private mlngCdEstabelecimento                               As Long
Private mobjRetorno                                         As clsCarregarRetorno
Private mobjTitulo                                          As New clsTituloCobrebem
Private Const grdChecked = 2
Private Const grdUnchecked = 1
Private Enum ENUMColBoleto
    e_nota = 1
    e_Parcela = 2
    e_Tipo = 3
    e_Empresa = 4
    e_Vencimento = 5
    e_Valor = 6
End Enum
Private Const strConsulta = "campo=SeqGrid;label=;tamanho=250|" & _
                           "campo=nota;label=Nota;tamanho=1000;tipo=tpColGridInteger|" & _
                           "campo=Parcela;label=Parcela;tamanho=700;tipo=tpColGridInteger|" & _
                           "campo=Tipo;label=Tipo;tamanho=1000|" & _
                           "campo=Empresa;label=Empresa;tamanho=1700|" & _
                           "campo=Vencimento;label=Vencimento;tamanho=1200|" & _
                           "campo=[valor Original];label=Valor;tamanho=1200;formato=###,##0.00;tipo=tpColGridInteger"

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdGerar_Click()
    Call LibProc(WL_NOVO)
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
End Sub

Private Sub etxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxBanco.ValorDescricao = "" Then
            etxBanco.valorInteiro = 0
        End If
        PCampo "Banco", "SELECT [Banco], [Nome], [Agência], [Conta] FROM [Bancos]", pbCampo, etxBanco, "Banco"
    End If
End Sub

Private Sub etxBanco_Change()
    If etxBanco.valorInteiro > 0 Then
        etxCarteira.Enabled = True
    Else
        etxCarteira.Enabled = False
        etxCarteira.Clear
    End If
End Sub

Private Sub etxCarteira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCarteira.ValorDescricao = "" Then
            etxCarteira.valorInteiro = 0
        End If
        'PCampo "Banco", "SELECT [id_carteira], [desc_carteira] FROM [FFICarteira]", pbCampo, etxCarteira, "id_carteira"
        PCampo "Banco", "SELECT Bancos.Banco, Bancos.Nome, FFICarteira.id_carteira, FFICarteira.desc_carteira " _
                      & "FROM (Bancos INNER JOIN FFIBanco_carteira ON Bancos.Banco = FFIBanco_carteira.Banco)  " _
                      & "INNER JOIN FFICarteira ON FFIBanco_carteira.id_carteira = FFICarteira.id_carteira WHERE Bancos.Banco = " & etxBanco.valorInteiro, pbCampo, etxCarteira, "id_carteira"
    End If
End Sub

Private Sub etxCarteira_LostFocus()
    Dim objCarteira                     As New clsCarteira
    Dim objCarteiraDao                  As New clsCarteiraDAO
    
    If etxCarteira.valorInteiro > 0 Then
        If Not mobjRetorno.ExisteCarteira(etxBanco.valorInteiro, etxCarteira.valorInteiro) Then
            MsgBox "Carteira não pertence ao banco selecionado.", vbInformation, NomeModulo
            etxCarteira.Clear
        End If
    End If
    
    If Trim(etxCaminhoRetorno.Valor) = "" Then
        Call objCarteiraDao.init(Aplicacao)
        Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, etxCarteira.valorInteiro)
        If Not objCarteira Is Nothing Then
            If Trim(objCarteira.Caminho_arquivo_retorno_padrao) = "" Then
                etxCaminhoRetorno.Valor = App.Path & "\Retorno.ret"
            Else
                etxCaminhoRetorno.Valor = objCarteira.Caminho_arquivo_retorno_padrao & "\Retorno.ret"
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Aplicacao.Connect
    Set mobjRetorno = New clsCarregarRetorno
    Call mobjRetorno.init(Aplicacao)
    Call mobjTitulo.init(Aplicacao)
    Call CenterForm(Me)
    Call fLoadEnterprise_estabelecimento
    Call etxCarteira.AddConexao(Aplicacao)
    Call etxBanco.AddConexao(Aplicacao)
    etxCarteira.Enabled = False
End Sub

'CARREGA O ENTERPRISE_ID E ESTABELECIMENTO.
Private Sub fLoadEnterprise_estabelecimento()
    mlngEnterpriseId = GetFieldValue("enterprise_id", "Usuários", "usuário = '" & UserName & "'")
    mlngCdEstabelecimento = GetFieldValue("cd_estabelecimento", "Usuários", "usuário = '" & UserName & "'")
End Sub

Private Function CarregaClasse()
    With mobjRetorno
        .Enterprise_id = mlngEnterpriseId
        .Cd_estabelecimento = mlngCdEstabelecimento
        .Banco = etxBanco.valorInteiro
        .Id_carteira = etxCarteira.valorInteiro
        .CaminhoRetorno = etxCaminhoRetorno.Valor
    End With
End Function

Private Function fLimpaCampo()
    etxBanco.Clear
    etxCarteira.Clear
    etxCaminhoRetorno.Clear
End Function

Public Function LibProc(strFuncao As String, Optional lngFuncao As Long) As Boolean
    Dim i                                       As Long
    Dim lngcount                                As Long
    Dim intIndex                                As Long
    Dim objBoleto                               As clsTituloCobrebem
    
    Select Case strFuncao
        Case WL_NOVO
            'Total de registros processados.
            'lngcount = 0
            If fValidaCampos Then
                CarregaClasse
                If mobjTitulo.CarregarRetorno(mobjRetorno.Id_carteira, mobjRetorno.Banco, mobjRetorno.CaminhoRetorno) Then
                    fLimpaCampo
                End If
            End If
            If mobjTitulo.MensagemValidacao <> "" Then
                frmMensagemBoleto.mensagem = mobjTitulo.MensagemValidacao
                frmMensagemBoleto.Caption = "Validação Título Bancário"
                'Validação emissão de relatório.
                'Pt 000000 - Moacir Pfau - 20/08/2009
                If mobjTitulo.MensagemBaixado = "" Then
                    frmMensagemBoleto.cmdImprimir.Visible = True
                    frmMensagemBoleto.NossoNumero = mobjTitulo.RelacaoNossoNumero
                End If
                
                Load frmMensagemBoleto
                frmMensagemBoleto.Show vbModal
            End If
            If mobjTitulo.MensagemBaixado <> "" Then
                frmMensagemBoleto.mensagem = mobjTitulo.MensagemBaixado
                frmMensagemBoleto.Caption = "Relação de Título(s) Baixado(s)"
                frmMensagemBoleto.cmdImprimir.Visible = True
                frmMensagemBoleto.NossoNumero = mobjTitulo.RelacaoNossoNumero
                Load frmMensagemBoleto
                frmMensagemBoleto.Show vbModal
            End If
        Case WL_SAIR
            Unload Me
    End Select
End Function

Private Function fValidaCampos() As Boolean
Dim strMensagem             As String

On Error GoTo err
    If etxBanco.valorInteiro = 0 Then
        strMensagem = strMensagem & "Banco é de preenchimento obrigatório." & vbCrLf
    End If
   
    If etxCarteira.valorInteiro = 0 Then
        strMensagem = strMensagem & "Carteira é de preenchimento obrigatório." & vbCrLf
    End If
    
    If Trim(etxCaminhoRetorno.Valor) = "" Then
        strMensagem = strMensagem & "Caminho do arquivo de retorno é de preenchimento obrigatório." & vbCrLf
    End If
    
    If strMensagem = "" Then
        fValidaCampos = True
    Else
        MsgBox strMensagem, vbInformation, NomeModulo
    End If
    
    Exit Function
err:
    fValidaCampos = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    Aplicacao.Disconnect
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
