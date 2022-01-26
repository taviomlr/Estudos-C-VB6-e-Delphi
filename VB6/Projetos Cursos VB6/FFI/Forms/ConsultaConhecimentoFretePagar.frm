VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmConsultaConhecimentoFretePagar 
   KeyPreview      =   -1  'True
   Caption         =   "Consulta de conhecimento de frete (A Pagar)"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4875
      Left            =   8880
      TabIndex        =   1
      Top             =   -45
      Width           =   1380
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar"
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton cmdExecutar 
         Caption         =   "&Executar"
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   165
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4875
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   8820
      Begin VB.TextBox txtConhecimentoFinal 
         Height          =   315
         Left            =   2970
         MaxLength       =   9
         TabIndex        =   5
         Top             =   195
         Width           =   1230
      End
      Begin VB.TextBox txtConhecimentoInicial 
         Height          =   315
         Left            =   1395
         MaxLength       =   9
         TabIndex        =   3
         Top             =   195
         Width           =   1230
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConhecimentos 
         Height          =   4230
         Left            =   45
         TabIndex        =   8
         Top             =   585
         Width           =   8730
         _ExtentX        =   15399
         _ExtentY        =   7461
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Fox.EBSData txtDataEmissaoInicial 
         Height          =   330
         Left            =   5520
         TabIndex        =   11
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Fox.EBSData txtDataEmissaoFinal 
         Height          =   330
         Left            =   7080
         TabIndex        =   12
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "à"
         Height          =   195
         Left            =   6915
         TabIndex        =   7
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   4500
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "à"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   2745
         TabIndex        =   4
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conhecimento"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   300
         TabIndex        =   2
         Top             =   240
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmConsultaConhecimentoFretePagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const strTituloGrid$ = "campo=nr_conhecimento;label=Conhecimento;tamanho=1200|" & _
                                "campo=tp_registro;label=Tipo Reg.;tamanho=1200|" & _
                                "campo=cd_transportadora;label=Transp.;tamanho=800|" & _
                                "campo=dt_emissao;label=Data;tamanho=1100|" & _
                                "campo=cd_remetente;label=Remetente;tamanho=2100|" & _
                                "campo=cd_destinatario;label=Destinatário;tamanho=2100|" & _
                                "campo=vl_conhecimento;label=Valor;tamanho=1500"
Private Enum ENUMCOLGRID
    col_numero = 0
    col_tipoRegistro = 1
    col_Transportadora = 2
    col_emissao = 3
    col_remetente = 4
    col_destinatario = 5
    col_valor = 6
End Enum

Private Sub cmdExecutar_Click()
    If ValidaCampos Then
        consultaConhecimentos
    End If
End Sub

Private Sub cmdVoltar_Click()
    Dim objFrete As cFretePagar
    
    Set objFrete = conhecimentoSelecionado
    If Not objFrete Is Nothing Then
        Call frmConhecimentoFretePagar.setConhecimento(objFrete)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = 10635
    Me.Height = 5265
    CenterForm Me
    Call LimpaCampos
End Sub

Private Sub grdConhecimentos_DblClick()
    Call cmdVoltar_Click
End Sub

Private Sub txtConhecimentoFinal_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub txtConhecimentoInicial_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii)
End Sub

Private Sub LimpaCampos()
    txtConhecimentoInicial.Text = ""
    txtConhecimentoFinal.Text = ""
    txtDataEmissaoInicial.clear
    txtDataEmissaoFinal.clear
    Call preencheConhecimentos
End Sub

Private Sub preencheConhecimentos(Optional rdResult As IDBReader = Nothing)
    If Not rdResult Is Nothing Then
        Call CarregaHFlexGrid(grdConhecimentos, rdResult.GetRecordset, strTituloGrid)
    Else
        Call CarregaHFlexGrid(grdConhecimentos, rdResult, strTituloGrid)
    End If
    grdConhecimentos.SelectionMode = flexSelectionByRow
End Sub

Private Sub consultaConhecimentos()
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "FreteEntrada"
    If txtConhecimentoInicial.Text <> "" Then
        Call cmd.Filter.Append("nr_conhecimento >= @pNrConhecimentoInicial")
        Call cmd.Parameters.add(cmd.CreateParameter("@pNrConhecimentoInicial", txtConhecimentoInicial.Text, dbFieldTypeLong))
    End If
    If txtConhecimentoFinal.Text <> "" Then
        Call cmd.Filter.Append("nr_conhecimento <= @pNrConhecimentoFinal")
        Call cmd.Parameters.add(cmd.CreateParameter("@pNrConhecimentoFinal", txtConhecimentoFinal.Text, dbFieldTypeLong))
    End If
    If txtDataEmissaoInicial.IsValidDate Then
        Call cmd.Filter.Append("dt_emissao >= @pDtEmissaoInicial")
        Call cmd.Parameters.add(cmd.CreateParameter("@pDtEmissaoInicial", txtDataEmissaoInicial.Data, dbFieldTypeDate))
    End If
    If txtDataEmissaoFinal.IsValidDate Then
        Call cmd.Filter.Append("dt_emissao <= @pDtEmissaoFinal")
        Call cmd.Parameters.add(cmd.CreateParameter("@pDtEmissaoFinal", txtDataEmissaoFinal.Data, dbFieldTypeDate))
    End If
    cmd.OrderByClause = "dt_emissao, nr_conhecimento"
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        Call preencheConhecimentos(rdResult)
    Else
        Call preencheConhecimentos
    End If
    rdResult.CloseReader
    Set rdResult = Nothing
    Aplicacao.Disconnect
End Sub

Private Function ValidaCampos() As Boolean
    ValidaCampos = True
    If Trim(txtConhecimentoInicial.Text) <> "" Then
        If Not IsNumeric(txtConhecimentoInicial.Text) Then
            MsgBox "O campo conhecimento inicial deve ser um número.", vbInformation, Me.Caption
            ValidaCampos = False
            Exit Function
        Else
            If Trim(txtConhecimentoInicial.Text) = "0" Then
                MsgBox "O campo conhecimento inicial deve ser um número maior do que ZERO.", vbInformation, Me.Caption
                ValidaCampos = False
                Exit Function
            End If
        End If
    End If
    
    If Trim(txtConhecimentoFinal.Text) <> "" Then
        If Not IsNumeric(txtConhecimentoFinal.Text) Then
            MsgBox "O campo conhecimento final deve ser um número.", vbInformation, Me.Caption
            ValidaCampos = False
            Exit Function
        Else
            If Trim(txtConhecimentoFinal.Text) = "0" Then
                MsgBox "O campo conhecimento final deve ser um número maior do que ZERO.", vbInformation, Me.Caption
                ValidaCampos = False
                Exit Function
            End If
        End If
    End If
        
    If Not txtDataEmissaoInicial.IsValidDate Then
        MsgBox "O campo data da emissão inicial deve ser uma data válida.", vbInformation, Me.Caption
        ValidaCampos = False
        Exit Function
    End If
    
    If Not txtDataEmissaoFinal.IsValidDate Then
        MsgBox "O campo data da emissão final deve ser uma data válida.", vbInformation, Me.Caption
        ValidaCampos = False
        Exit Function
    End If
End Function

Private Function conhecimentoSelecionado() As cFretePagar
Dim dao As New cFretePagarDAO
    With grdConhecimentos
        If .Row > 0 Then
            Set conhecimentoSelecionado = dao.Carregar(strToLng(.TextMatrix(.Row, col_numero)), .TextMatrix(.Row, col_tipoRegistro), strToLng(.TextMatrix(.Row, col_Transportadora)))
        End If
    End With
End Function

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
