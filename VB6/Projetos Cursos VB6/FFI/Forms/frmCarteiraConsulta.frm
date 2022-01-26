VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmCarteiraConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Carteira"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
   HelpContextID   =   2784
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraConsulta 
      Height          =   780
      Left            =   40
      TabIndex        =   5
      Top             =   -45
      Width           =   6870
      Begin Fox.EBSText etxNumeroInicial 
         Height          =   330
         Left            =   900
         TabIndex        =   0
         Top             =   255
         Width           =   1230
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
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
      Begin Fox.EBSText etxNumeroFinal 
         Height          =   330
         Left            =   2385
         TabIndex        =   1
         Top             =   255
         Width           =   1230
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
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
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   2190
         TabIndex        =   7
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lblNumero 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3780
      Left            =   6915
      TabIndex        =   4
      Top             =   -45
      Width           =   1410
      Begin VB.CommandButton cmdExecutar 
         Caption         =   "&Executar"
         Height          =   375
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdVoltar 
         Caption         =   "&Voltar"
         Height          =   375
         Left            =   90
         TabIndex        =   3
         Top             =   585
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPesquisa 
      Height          =   2970
      Left            =   45
      TabIndex        =   8
      Top             =   735
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   5239
      _Version        =   393216
      FixedRows       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmCarteiraConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const strTituloGrid$ = "campo=id_carteira;label=Número;tamanho=1500;tipo=tpColGridInteger|" & _
                                "campo=desc_carteira;label=Nome;tamanho=5000"
                                
Private Sub cmdExecutar_Click()
    If ValidaCampos Then
        SetPtr vbHourglass
        ConsultaRegistro
        SetPtr vbDefault
    End If
End Sub

Private Sub cmdVoltar_Click()
    Dim intNumero       As Integer
    
    With grdPesquisa
        intNumero = strToLng(.TextMatrix(.Row, 0))
    End With
    If intNumero > 0 Then
            Call frmCarteira.fCarregaPesquisa(intNumero)
        Call Unload(Me)
    Else
        Call Unload(Me)
    End If
End Sub

Private Sub Form_Load()
    Aplicacao.Connect
    Call CarregaHFlexGrid(grdPesquisa, Nothing, strTituloGrid)
    AjusteGrid
    Call ConsultaRegistro
End Sub

Private Function ValidaCampos() As Boolean
    If etxNumeroInicial.valorInteiro > 0 And etxNumeroFinal.valorInteiro > 0 Then
        If etxNumeroInicial.valorInteiro > etxNumeroFinal.valorInteiro Then
            Call MsgBox("O campo número inicial dever ser menor do que o campo número final.", vbInformation, NomeModulo)
            etxNumeroInicial.SetFocus
            ValidaCampos = False
        Else
            ValidaCampos = True
        End If
    Else
        ValidaCampos = True
    End If
End Function

Private Sub ConsultaRegistro(Optional strOrdenar As String)
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Dim rsResult As Object
    
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "[FFiCarteira]"
    
    If etxNumeroInicial.valorInteiro <> 0 Then
        Call cmd.Filter.Append("id_carteira >= @pModCodigoInicial")
        Call cmd.Parameters.add(cmd.CreateParameter("@pModCodigoInicial", etxNumeroInicial.valorInteiro, dbFieldTypeLong))
    End If
    If etxNumeroFinal.valorInteiro <> 0 Then
        Call cmd.Filter.Append("id_carteira <= @pModCodigoFinal")
        Call cmd.Parameters.add(cmd.CreateParameter("@pModCodigoFinal", etxNumeroFinal.valorInteiro, dbFieldTypeLong))
    End If
    
    If strOrdenar = "" Then
        cmd.OrderByClause = "id_carteira" ', DESC desc_carteira "
    Else
        cmd.OrderByClause = strOrdenar
    End If
    
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    Set rsResult = rdResult.GetRecordset
    Call CarregaHFlexGrid(grdPesquisa, rsResult, strTituloGrid)
    Call AjusteGrid
    Set rsResult = Nothing
    rdResult.CloseReader
    Set rdResult = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Aplicacao.Disconnect
End Sub

Private Sub grdPesquisa_dblClick()
    Call cmdVoltar_Click
End Sub

Private Sub grdPesquisa_Click()
    Dim strCampo  As String
    
    If grdPesquisa.Row = 0 Then
        Select Case grdPesquisa.col
            Case 0
                strCampo = "ModCodigo"
            Case 1
                strCampo = "ModNome"
        End Select
        If IsNumeric(grdPesquisa.TextMatrix(Row, 0)) Then
            Call ConsultaRegistro(strCampo)
        End If
    End If
    
End Sub

Public Sub AjusteGrid()
    Dim i As Integer
    grdPesquisa.FixedRows = 0
    grdPesquisa.Row = 0
    For i = 0 To 1
        grdPesquisa.col = i
        grdPesquisa.CellBackColor = &H8000000F
    Next
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
