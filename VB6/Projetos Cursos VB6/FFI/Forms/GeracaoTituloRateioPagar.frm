VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmGeracaoTituloRateioPagar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rateio de Geração de Títulos Pagar"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3390
      Left            =   4140
      TabIndex        =   9
      Top             =   -45
      Width           =   1410
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3390
      Left            =   0
      TabIndex        =   2
      Top             =   -45
      Width           =   4110
      Begin VB.CommandButton cmdCancelarRateio 
         Caption         =   "Ca&ncelar"
         Height          =   375
         Left            =   2655
         TabIndex        =   8
         Top             =   855
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   "&Remover"
         Height          =   375
         Left            =   1395
         TabIndex        =   7
         Top             =   855
         Width           =   1215
      End
      Begin VB.CommandButton cmdInserir 
         Caption         =   "&Inserir"
         Height          =   375
         Left            =   135
         TabIndex        =   4
         Top             =   855
         Width           =   1215
      End
      Begin Fox.EBSText etxCentroCusto 
         Height          =   330
         Left            =   90
         TabIndex        =   0
         Top             =   450
         Width           =   1200
         _ExtentX        =   2117
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
      Begin Fox.EBSText etxConta 
         Height          =   330
         Left            =   1485
         TabIndex        =   1
         Top             =   450
         Width           =   1230
         _ExtentX        =   2170
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
      Begin Fox.EBSText etxPercentual 
         Height          =   330
         Left            =   2880
         TabIndex        =   3
         Top             =   450
         Width           =   1050
         _ExtentX        =   265
         _ExtentY        =   582
         Tipo            =   1
         CasasDecimais   =   2
         TipoTexto       =   0
         MaxLength       =   6
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdRatFin 
         Height          =   1995
         Left            =   90
         TabIndex        =   13
         Top             =   1305
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   3519
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Percentual"
         Height          =   195
         Left            =   3015
         TabIndex        =   12
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Conta"
         Height          =   195
         Left            =   1755
         TabIndex        =   11
         Top             =   180
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C.Custo"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   180
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmGeracaoTituloRateioPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objGerTitPagar          As New cGeracaoTituloPagar
Private mobjGerTitPagar         As New cGeracaoTituloPagar
Private objCRateioTitPagar      As New cGeracaoTituloPagar
Public mbolObj                  As Boolean

Private booAlterando            As Boolean
'pt. 85684 - Moacir Pfau(02/07/2008) - Geração de Titulos.
Private Const NomeTabela$ = "FFITituloPagarRateio"
Private Const strTituloGrid$ = "campo=R_Cd_titulo;label=Cd_Titulo;tamanho=100|" & _
                    "campo=R_Cd_centro_custo;label=C.Centro;tamanho=1000|" & _
                    "campo=R_Cd_conta;label=Conta;tamanho=1000|" & _
                    "campo=R_Percentual;label=Percentual;tamanho=1000;formato=###,##0.00"
Private mblnGerouDuplicatas As Boolean

'pt. 85684 - Ivo Sousa(15/07/2008)
Public Property Get GerouDuplicatas() As Boolean
    GerouDuplicatas = mblnGerouDuplicatas
End Property

Public Property Let GerouDuplicatas(ByVal blnGerouDuplicatas As Boolean)
    mblnGerouDuplicatas = blnGerouDuplicatas
End Property

Private Sub cmdCancelar_Click()
    Set objCRateioTitPagar = Nothing
    frmGeracaoTitulosPagar.CarregaColRateio objCRateioTitPagar
    Unload Me
End Sub

Private Sub cmdCancelarRateio_Click()
    etxCentroCusto.valorInteiro = 0
    etxConta.valorInteiro = 0
    etxPercentual.valorMoeda = 0
    etxCentroCusto.SetFocus
End Sub

Private Sub cmdInserir_Click()
    Dim strSql As String
    Dim strConta_Grid As String
    Dim strAgrupamentoConta As Boolean
    Dim strAgrupamentoContaGrid As Boolean
    Dim rstResult As Object
    
    If ValidaCampos Then
        
        'Início - Marcel Henrique (Data: 14/01/2015 Projeto: #59699 Desenvolvimento: #62974)
        If grdRatFin.TextMatrix(1, 2) <> "" Then
                
            'Obtenção do campo agrupamento C.C. da conta a inserir
            strSql = "SELECT agrupa_centro_custo FROM contas " & _
            " WHERE código=" & etxConta.valorInteiro
            
            If AbreRecordset(rstResult, strSql) = WL_OK Then
                strAgrupamentoConta = rstResult.Fields("agrupa_centro_custo").value
            End If
            Call FechaRecordset(rstResult)
        
            'Obtenção do campo agrupamento C.C. da conta a inserida
            strConta_Grid = grdRatFin.TextMatrix(1, 2)
           
            strSql = "SELECT agrupa_centro_custo FROM contas " & _
            " WHERE código=" & strConta_Grid
            
            If AbreRecordset(rstResult, strSql) = WL_OK Then
                strAgrupamentoContaGrid = rstResult.Fields("agrupa_centro_custo").value
            End If
            Call FechaRecordset(rstResult)
            
            'Comparação dos conteúdos de agrupamento, obtidos
            If strAgrupamentoConta <> strAgrupamentoContaGrid Then
                MsgBox "As contas utilizadas no rateio devem ter a mesma configuração" & vbCrLf & _
                "de agrupamento por centro de custo." & vbCrLf & vbCrLf & _
                "A conta que está tentando inserir possui configuração diferente" & vbCrLf & _
                "das já inseridas. Verifique a conta."
                etxConta.SetFocus
                Exit Sub
            End If
    
        End If
        'Fim - Marcel Henrique
        
        fInserir
        etxCentroCusto.SetFocus
    End If
End Sub

Private Sub cmdRemover_Click()
    Call objCRateioTitPagar.Rateio.Remove(objGerTitPagar)
    Call limpaCamposTitulos
    Call CarregaGrid
End Sub

Private Sub cmdConfirmar_Click()
    If objCRateioTitPagar.Rateio.Count > 0 Then
        'Demanda 103241 - Davi Brito - 01/07/2016
        If CDbl(CStr(objCRateioTitPagar.Rateio.totalValor)) <> 100 Then
            MsgBox "Soma do valor de rateio esta diferente de 100%"
            Exit Sub
        End If
    Else
        mbolObj = True
    End If
    frmGeracaoTitulosPagar.CarregaColRateio objCRateioTitPagar
    Unload Me
End Sub

Private Sub etxConta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxConta.ValorDescricao = "" Then
            etxConta.valorInteiro = 0
        End If
        PCampo "Contas", "SELECT Código, Grupo, Descrição FROM Contas", pbCampo, etxConta, "Código"
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    Call etxCentroCusto.AddConexao(Aplicacao)
    Call etxConta.AddConexao(Aplicacao)
    Call limpaCamposTitulos
    If objCRateioTitPagar.Rateio.Count = 0 Then
        Set objCRateioTitPagar = New cGeracaoTituloPagar
        Call CarregaColecao
    End If
    Call CarregaGrid
    'pt. 85684 - Ivo Sousa(15/07/2008)
    If mblnGerouDuplicatas Then
        etxCentroCusto.Enabled = False
        etxConta.Enabled = False
        etxPercentual.Enabled = False
    End If
End Sub

Private Sub CarregaColecao()
    Dim strSql       As String
    Dim rstTab       As Object
    Dim i            As Integer
    Dim GerTitPagar As New cGeracaoTituloPagar
    
    If mbolObj = False Then
        strSql = ""
        strSql = strSql & "SELECT * "
        strSql = strSql & "FROM " & NomeTabela & " "
        strSql = strSql & "WHERE cd_titulo=" & objGerTitPagar.Cd_Titulo
        If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
            rstTab.MoveFirst
            While Not rstTab.EOF
                Set GerTitPagar = New cGeracaoTituloPagar
                With GerTitPagar
                    .R_Cd_titulo = objGerTitPagar.Cd_Titulo
                    .R_Cd_centro_custo = GetValue(rstTab, "Cd_Centro_Custo")
                    .R_Cd_conta = GetValue(rstTab, "Cd_Conta_Financeira")
                    .R_Percentual = GetValue(rstTab, "pr_percentual")
                End With
                Call objCRateioTitPagar.Rateio.add(GerTitPagar)
                Set GerTitPagar = Nothing
                rstTab.MoveNext
            Wend
        End If
        FechaRecordset (rstTab)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objRateioTitPagar = Nothing
    Set objCRateioTitPagar = Nothing
End Sub

Public Sub CarregaObj(mobjGerTitPagar As Object)
    Set objGerTitPagar = mobjGerTitPagar
End Sub

Public Sub CarregaCol(mobjGerTitPagar As Object)
    Set objCRateioTitPagar = mobjGerTitPagar
End Sub

Private Sub etxCentroCusto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCentroCusto.ValorDescricao = "" Then
            etxCentroCusto.valorInteiro = 0
        End If
        PCampo "Centro", "SELECT Código, Descrição FROM Centros", pbCampo, etxCentroCusto, "Código"
    End If
End Sub

Private Sub preencheClasse()
    Call preencheTitPagarClasse
End Sub

Private Sub preencheTitPagarClasse()
    With objGerTitPagar
        .R_Cd_titulo = objGerTitPagar.Cd_Titulo
        .R_Cd_centro_custo = etxCentroCusto.valorInteiro
        .R_Cd_conta = etxConta.valorInteiro
        .R_Percentual = etxPercentual.valorMoeda
    End With
End Sub

Private Sub limpaCamposTitulos()
    etxCentroCusto.valorInteiro = 0
    etxConta.valorInteiro = 0
    etxPercentual.valorMoeda = 0
    grdRatFin.Clear
End Sub

Private Sub CarregaGrid()
    grdRatFin.Clear
    If objCRateioTitPagar.Rateio.Count = 0 Then
        Call CarregaHFlexGrid(grdRatFin, Nothing, strTituloGrid)
    Else
        objCRateioTitPagar.Rateio.MoveFirst
        Call CarregaHFlexGrid(grdRatFin, , strTituloGrid, , , objCRateioTitPagar.Rateio)
    End If
    grdRatFin.FixedCols = 1
End Sub

Private Sub fInserir()
    Dim strSql       As String
    Dim rstTab       As Object
    Dim i            As Integer
    Dim GerTitPagar As New cGeracaoTituloPagar
    
    If booAlterando Then
        Call objCRateioTitPagar.Rateio.Remove(objGerTitPagar)
    End If
    
    preencheClasse
    If objCRateioTitPagar.Rateio.Find(objGerTitPagar) Then
        MsgBox "Registro já lançado, não será possível realizar o lançamento.", vbInformation
        Exit Sub
    End If
    Set GerTitPagar = New cGeracaoTituloPagar
    With GerTitPagar
        .R_Cd_titulo = objGerTitPagar.Cd_Titulo
        .R_Cd_centro_custo = objGerTitPagar.R_Cd_centro_custo
        .R_Cd_conta = objGerTitPagar.R_Cd_conta
        .R_Percentual = objGerTitPagar.R_Percentual
    End With
    Call objCRateioTitPagar.Rateio.add(GerTitPagar)
    Set GerTitPagar = Nothing
    
    Call limpaCamposTitulos
    Call CarregaGrid
    booAlterando = False
End Sub

Private Sub Excluir()
    Dim strSql       As String
    Dim rstTab       As Object
    Dim i            As Integer
    Dim GerTitPagar As New cGeracaoTituloPagar
    
    'Verifica se existe parcelas geradas.
    strSql = ""
    strSql = strSql & "SELECT * "
    strSql = strSql & "FROM " & NomeTabela & " "
    strSql = strSql & "WHERE cd_titulo=" & objGerTitPagar.Cd_Titulo
    If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
        rstTab.MoveFirst
        While Not rstTab.EOF
            Set GerTitPagar = New cGeracaoTituloPagar
            With GerTitPagar
                .R_Cd_titulo = objGerTitPagar.Cd_Titulo
                .R_Cd_centro_custo = GetValue(rstTab, "Cd_Centro_Custo")
                .R_Cd_conta = GetValue(rstTab, "Cd_Conta_Financeira")
                .R_Percentual = GetValue(rstTab, "pr_percentual")
            End With
            Call objCRateioTitPagar.Rateio.add(GerTitPagar)
            Set GerTitPagar = Nothing
            rstTab.MoveNext
        Wend
    End If
    Call CarregaGrid
    FechaRecordset (rstTab)
End Sub

Private Sub mostraCamposClasse()
    Call carregaCamposRateioPagar
    Call mostraCamposRateioPagar
End Sub

Private Sub carregaCamposRateioPagar()
    With objGerTitPagar
        .R_Cd_titulo = grdRatFin.TextMatrix(grdRatFin.Row, 0)
        .R_Cd_centro_custo = grdRatFin.TextMatrix(grdRatFin.Row, 1)
        .R_Cd_conta = grdRatFin.TextMatrix(grdRatFin.Row, 2)
        .R_Percentual = grdRatFin.TextMatrix(grdRatFin.Row, 3)
    End With
    booAlterando = True
End Sub

Private Sub mostraCamposRateioPagar()
    With objGerTitPagar
        etxCentroCusto.valorInteiro = .R_Cd_centro_custo
        etxConta.valorInteiro = .R_Cd_conta
        etxPercentual.valorMoeda = .R_Percentual
    End With
End Sub

Private Sub grdRatFin_DblClick()
    mostraCamposClasse
End Sub

Private Function ValidaCampos() As Boolean
    'pt. 85684 - Moacir Pfau(02/07/2008) - VALIDAÇÃO AO INSERIR.
    Dim strMensagem As String
    strMensagem = ""
    
    If etxCentroCusto.valorInteiro = 0 And etxCentroCusto.Enabled = True Then
        strMensagem = strMensagem & "Preenchimento do campo centro de custo é obrigatório." & vbCrLf
    End If
    If etxConta.valorInteiro = 0 Then
        strMensagem = strMensagem & "Preenchimento do campo conta é obrigatório." & vbCrLf
    End If
    If etxPercentual.valorMoeda = 0 Then
        strMensagem = strMensagem & "Preenchimento do campo valor é obrigatório." & vbCrLf
    End If
    If etxPercentual.valorMoeda > 100 Then
        strMensagem = strMensagem & "Preenchimento do campo valor tem que ser igual ou menor que 100." & vbCrLf
    End If

    If strMensagem = "" Then
        ValidaCampos = True
    Else
        MsgBox strMensagem, vbInformation
    End If
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
