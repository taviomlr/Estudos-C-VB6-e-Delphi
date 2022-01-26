VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmModeloConsulta 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Modelo"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   HelpContextID   =   2784
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraConsulta 
      Height          =   780
      Left            =   40
      TabIndex        =   5
      Top             =   -45
      Width           =   5940
      Begin Fox.EBSText etxNumeroInicial 
         Height          =   330
         Left            =   900
         TabIndex        =   0
         Top             =   255
         Width           =   1230
         _extentx        =   265
         _extenty        =   582
         tipotexto       =   0
         maxlength       =   9
         tipocriterio    =   4
         alinhamento     =   1
         font            =   "frmModeloConsulta.frx":0000
      End
      Begin Fox.EBSText etxNumeroFinal 
         Height          =   330
         Left            =   2385
         TabIndex        =   1
         Top             =   255
         Width           =   1230
         _extentx        =   265
         _extenty        =   582
         tipotexto       =   0
         maxlength       =   9
         tipocriterio    =   4
         alinhamento     =   1
         font            =   "frmModeloConsulta.frx":002C
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
      Left            =   6015
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
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   5239
      _Version        =   393216
      FixedRows       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmModeloConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const strTituloGrid$ = "campo=ModCodigo;label=Número;tamanho=1000;tipo=tpColGridInteger|" & _
                                "campo=ModNome;label=Nome;tamanho=2500"
                                
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
            Call frmModelo.fCarregaPesquisa(intNumero)
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
    ValidaCampos = False
    If etxNumeroInicial.valorInteiro > 0 And etxNumeroFinal.valorInteiro > 0 Then
        If etxNumeroInicial.valorInteiro > etxNumeroFinal.valorInteiro Then
            Call MsgBox("O campo número inicial dever ser menor do que o campo número final.", vbInformation, NomeModulo)
            etxNumeroInicial.SetFocus
            ValidaCampos = False
        Else
            ValidaCampos = True
        End If
    End If
End Function

Private Sub ConsultaRegistro(Optional strOrdenar As String)
    Dim cmd As IDBSelectCommand
    Dim rdResult As IDBReader
    Dim rsResult As Object
    
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.Table.TableName = "[Modelo]"
    
    If etxNumeroInicial.valorInteiro <> 0 Then
        Call cmd.Filter.Append("ModCodigo >= @pModCodigoInicial")
        Call cmd.Parameters.Add(cmd.CreateParameter("@pModCodigoInicial", etxNumeroInicial.valorInteiro, dbFieldTypeLong))
    End If
    If etxNumeroFinal.valorInteiro <> 0 Then
        Call cmd.Filter.Append("ModCodigo <= @pModCodigoFinal")
        Call cmd.Parameters.Add(cmd.CreateParameter("@pModCodigoFinal", etxNumeroFinal.valorInteiro, dbFieldTypeLong))
    End If
    
    If strOrdenar = "" Then
        cmd.OrderByClause = "ModCodigo DESC, ModNome "
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
        
        Call ConsultaRegistro(strCampo)
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
