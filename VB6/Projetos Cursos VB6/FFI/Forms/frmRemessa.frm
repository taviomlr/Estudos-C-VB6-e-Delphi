VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmRemessa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão Remessa Bancária"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   Icon            =   "frmRemessa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   10740
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgConsulta 
      Height          =   3435
      Left            =   30
      TabIndex        =   39
      Top             =   4830
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   6059
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraBotoes 
      Height          =   8325
      Left            =   9360
      TabIndex        =   34
      Top             =   -30
      Width           =   1350
      Begin VB.CommandButton cmdConsulta 
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   90
         TabIndex        =   35
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "&Gerar"
         Height          =   375
         Left            =   90
         TabIndex        =   36
         Top             =   570
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   38
         Top             =   1350
         Width           =   1185
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   37
         Top             =   960
         Width           =   1185
      End
      Begin MSComctlLib.ImageList imgCheck 
         Left            =   240
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRemessa.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRemessa.frx":06DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRemessa.frx":0A2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmRemessa.frx":0CD9
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame 
      Height          =   4005
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   9285
      Begin Fox.EBSCombo cboTipo 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   1710
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Dados           =   ""
         DadosAssist     =   ""
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
      Begin Fox.EBSData etxDataInicial 
         Height          =   330
         Left            =   1680
         TabIndex        =   12
         Top             =   2070
         Width           =   1905
         _ExtentX        =   3360
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
      Begin VB.CheckBox chkImpressao 
         Caption         =   "Emitir remessa já gerada"
         Height          =   435
         Left            =   1620
         TabIndex        =   25
         Top             =   3540
         Width           =   2865
      End
      Begin Fox.EBSText etxBanco 
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   270
         Width           =   5385
         _ExtentX        =   439632
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
         _ExtentX        =   439632
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
      Begin Fox.EBSText etxNumeroInicial 
         Height          =   330
         Left            =   1680
         TabIndex        =   16
         Top             =   2430
         Width           =   1905
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
      End
      Begin Fox.EBSText etxNumeroFinal 
         Height          =   330
         Left            =   1680
         TabIndex        =   20
         Top             =   2790
         Width           =   1905
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
      End
      Begin Fox.EBSText etxParcelaInicial 
         Height          =   330
         Left            =   4380
         TabIndex        =   18
         Top             =   2430
         Width           =   1905
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
      End
      Begin Fox.EBSText etxParcelaFinal 
         Height          =   330
         Left            =   4380
         TabIndex        =   22
         Top             =   2790
         Width           =   1905
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
      End
      Begin Fox.EBSData etxDataFinal 
         Height          =   330
         Left            =   4380
         TabIndex        =   14
         Top             =   2070
         Width           =   1905
         _ExtentX        =   3360
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
      Begin Fox.EBSCombo cboOrigem 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   990
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Dados           =   ""
         DadosAssist     =   ""
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
      Begin Fox.EBSText etxEmpresa 
         Height          =   330
         Left            =   1680
         TabIndex        =   8
         Top             =   1350
         Width           =   5385
         _ExtentX        =   439632
         _ExtentY        =   582
         Tipo            =   4
         MaxLength       =   15
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Apel"
         CampoDescricao  =   "Razão"
         TabelaConsulta  =   "Empresas"
         TamanhoDescricao=   3500
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
      Begin Fox.EBSArquivo etxCaminhoRemessa 
         Height          =   330
         Left            =   1665
         TabIndex        =   24
         Top             =   3150
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   582
         Filter          =   ""
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Caminho R&emessa"
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
         Left            =   60
         TabIndex        =   23
         Top             =   3225
         Width           =   1560
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "&Empresa"
         Height          =   195
         Left            =   1005
         TabIndex        =   7
         Top             =   1425
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Origem"
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
         Left            =   1020
         TabIndex        =   5
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "P&arcela"
         Height          =   195
         Left            =   3750
         TabIndex        =   21
         Top             =   2858
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Parcela"
         Height          =   195
         Left            =   3750
         TabIndex        =   17
         Top             =   2505
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tít&ulo Final"
         Height          =   195
         Left            =   825
         TabIndex        =   19
         Top             =   2865
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "T&ítulo Inicial"
         Height          =   195
         Left            =   750
         TabIndex        =   15
         Top             =   2475
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   3930
         TabIndex        =   13
         Top             =   2160
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Intervalo de Data"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2145
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo de Filtro"
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
         TabIndex        =   9
         Top             =   1770
         Width           =   1140
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
   Begin VB.Frame Pedidos 
      Caption         =   "Títulos"
      Height          =   855
      Left            =   30
      TabIndex        =   29
      Top             =   3960
      Width           =   3900
      Begin Fox.EBSText etxTotalPedidos 
         Height          =   330
         Left            =   435
         TabIndex        =   32
         Top             =   435
         Width           =   1245
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
         Enabled         =   0   'False
         TipoCriterio    =   4
         Alinhamento     =   1
         Locked          =   -1  'True
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
      Begin Fox.EBSText etxTotalSelecionados 
         Height          =   330
         Left            =   2115
         TabIndex        =   33
         Top             =   435
         Width           =   1245
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   9
         Enabled         =   0   'False
         TipoCriterio    =   4
         Alinhamento     =   1
         Locked          =   -1  'True
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
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Marcados"
         Height          =   195
         Left            =   2115
         TabIndex        =   31
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Todos"
         Height          =   195
         Left            =   435
         TabIndex        =   30
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Frame Frame 
      Height          =   855
      Index           =   1
      Left            =   3945
      TabIndex        =   26
      Top             =   3960
      Width           =   5370
      Begin VB.CommandButton cmdNaoEnviados 
         Caption         =   "Nã&o Enviados"
         Height          =   375
         Left            =   1470
         TabIndex        =   41
         Top             =   330
         Width           =   1185
      End
      Begin VB.CommandButton cmdEnviados 
         Caption         =   "&Enviados"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   330
         Width           =   1185
      End
      Begin VB.CommandButton cmdTodos 
         Caption         =   "T&odos"
         Height          =   375
         Left            =   2700
         TabIndex        =   27
         Top             =   330
         Width           =   1185
      End
      Begin VB.CommandButton cmdNenhum 
         Caption         =   "&Nenhum"
         Height          =   375
         Left            =   3930
         TabIndex        =   28
         Top             =   330
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngEnterpriseId                                    As Long
Private mlngCdEstabelecimento                               As Long
Private mlngOrdemAscDesc                                    As Long
Private mbolOrdDtAscendente                                 As Boolean
Private mobjRemessa                                         As clsEmissaoRemessa
Private mobjTitulo                                          As New clsTituloCobrebem

Private Const grdChecked = 2
Private Const grdUnchecked = 1
Private Const grdAsc = 3
Private Const grdDesc = 4

Private Const strConsulta = "campo=SeqGrid;label=;tamanho=250|" & _
                            "campo=nota;label=Título;tamanho=1500;tipo=tpColGridInteger|" & _
                            "campo=Parcela;label=Parc.;tamanho=700;tipo=tpColGridInteger|" & _
                            "campo=Tipo;label=Tipo;tamanho=900|" & _
                            "campo=Empresa;label=Empresa;tamanho=1600|" & _
                            "campo=Vencimento;label=Vencimento;tamanho=1200|" & _
                            "campo=[valor Original];label=Valor;tamanho=1000;formato=###,##0.00;tipo=tpColGridInteger|" & _
                            "campo=Remessa;label=Status Remessa;tamanho=1600"

Private Enum ENUMColBoleto
    e_nota = 1
    e_Parcela = 2
    e_Tipo = 3
    e_Empresa = 4
    e_Vencimento = 5
    e_Valor = 6
    e_RemessaStatus = 7
End Enum

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdConsulta_Click()
    Call LibProc(WL_CONSULTA)
    Call CountMarcados
End Sub

Private Sub cmdEnviados_Click()
    Dim intIndex           As Integer
    
    With fgConsulta
        If .TextMatrix(1, e_Tipo) <> "" Then
            For intIndex = 1 To .Rows - 1
                If .TextMatrix(intIndex, e_RemessaStatus) = "Enviado" Then
                    .Row = intIndex
                    .col = 0
                    Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                Else
                    .Row = intIndex
                    .col = 0
                    Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                End If
            Next
        End If
    End With
    Call CountMarcados
End Sub

Private Sub cmdGerar_Click()
    Call LibProc(WL_NOVO)
End Sub

Private Sub cmdNaoEnviados_Click()
    Dim intIndex           As Integer
    
    With fgConsulta
        If .TextMatrix(1, e_Tipo) <> "" Then
            For intIndex = 1 To .Rows - 1
                If (.TextMatrix(intIndex, e_RemessaStatus) = "Não Enviado") Or (.TextMatrix(intIndex, e_RemessaStatus) = "Boleto Gerado") Then
                    .Row = intIndex
                    .col = 0
                    Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                Else
                    .Row = intIndex
                    .col = 0
                    Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                End If
            Next
        End If
    End With
    Call CountMarcados
End Sub

Private Sub cmdNenhum_Click()
    Dim intIndex           As Integer
    
    With fgConsulta
        For intIndex = 1 To .Rows - 1
            .Row = intIndex
            .col = 0
            Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
        Next
    End With
    Call CountMarcados
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
End Sub

Private Sub cmdTodos_Click()
    Dim intIndex            As Integer
    
    With fgConsulta
        If .TextMatrix(1, e_Tipo) <> "" Then
            For intIndex = 1 To .Rows - 1
                .Row = intIndex
                .col = 0
                Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
            Next
        End If
    End With
    Call CountMarcados
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

Private Sub etxCaminhoRemessa_LostFocus()
    If Trim(etxCaminhoRemessa.Valor) <> "" Then
        If Mid(etxCaminhoRemessa.Valor, Len(etxCaminhoRemessa.Valor) - 3, 1) <> "." Then
            etxCaminhoRemessa.Valor = etxCaminhoRemessa.Valor & "\remessa.txt"
        End If
    End If
End Sub

Private Sub etxCarteira_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxCarteira.ValorDescricao = "" Then
            etxCarteira.valorInteiro = 0
        End If
        PCampo "Banco", "SELECT Bancos.Banco, Bancos.Nome, FFICarteira.id_carteira, FFICarteira.desc_carteira " _
                      & "FROM (Bancos INNER JOIN FFIBanco_carteira ON Bancos.Banco = FFIBanco_carteira.Banco)  " _
                      & "INNER JOIN FFICarteira ON FFIBanco_carteira.id_carteira = FFICarteira.id_carteira WHERE Bancos.Banco = " & etxBanco.valorInteiro, pbCampo, etxCarteira, "id_carteira"
    End If
End Sub

Private Sub etxCarteira_LostFocus()
    Dim objCarteira                     As New clsCarteira
    Dim objCarteiraDao                  As New clsCarteiraDAO
    
    If etxCarteira.valorInteiro > 0 Then
        If Not mobjRemessa.ExisteCarteira(etxBanco.valorInteiro, etxCarteira.valorInteiro) Then
            MsgBox "Carteira não pertence ao banco selecionado.", vbInformation, NomeModulo
            etxCarteira.Clear
        End If
    End If
    
    If Trim(etxCaminhoRemessa.Valor) = "" And etxCarteira.valorInteiro > 0 Then
        Call objCarteiraDao.init(Aplicacao)
        Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, etxCarteira.valorInteiro)
        If Not objCarteira Is Nothing Then
            If Trim(objCarteira.Caminho_arquivo_remessa_padrao) = "" Then
                etxCaminhoRemessa.Valor = App.Path & "\Remessa.txt"
            Else
                etxCaminhoRemessa.Valor = objCarteira.Caminho_arquivo_remessa_padrao & "\Remessa.txt"
            End If
        End If
    End If
End Sub

Private Sub etxDataInicial_LostFocus()
    If Not etxDataFinal.IsValidDate Then
        etxDataFinal.Data = etxDataInicial.Data
    End If
End Sub

Private Sub etxEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxEmpresa.ValorDescricao = "" Then
            etxEmpresa.valorTexto = ""
        End If
        PCampo "Banco", "SELECT [APEL], [RAZÃO], [Tipo], [CNPJ/CPF] FROM [Empresas]", pbCampo, etxEmpresa, "Apel"
    End If
End Sub

Private Sub etxNumeroFinal_Change()
    If etxNumeroFinal.valorInteiro > 0 Then
        etxParcelaFinal.Enabled = True
    Else
        etxParcelaFinal.Enabled = False
        etxParcelaFinal.Clear
    End If
End Sub

Private Sub etxNumeroInicial_Change()
    If etxNumeroInicial.valorInteiro > 0 Then
        etxParcelaInicial.Enabled = True
    Else
        etxParcelaInicial.Enabled = False
        etxParcelaInicial.Clear
    End If
End Sub

Private Sub fgConsulta_Click()
    With fgConsulta
        If (.Row > 0) And (.TextMatrix(.Row, e_nota) <> "") Then
            If LinhaSelecionada(.Row) Then
                Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
            Else
                Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
            End If
            
            Call CountMarcados
        Else
            If .col > 0 Then
                Dim i As Integer
                Dim colSelecionada As Integer
                            
                colSelecionada = .col
                For i = 1 To .Cols - 1
                    If colSelecionada <> i Then
                        .col = i
                        Set .CellPicture = Nothing
                    End If
                Next
            End If
            
            mbolOrdDtAscendente = Not mbolOrdDtAscendente
        End If
    End With
End Sub

' Luiz Satto - 22/09/2016 - Protocolo 402923
Private Sub OrdenarGrid(fgConsulta As MSHFlexGrid, ByVal lngColuna As Long)
    With fgConsulta
        .Redraw = False
        .col = lngColuna
        .ColSel = lngColuna
                
        .Row = 0
        .RowSel = 0
        
        If lngColuna = 5 Then ' Coluna 5 = Data de vencimento
            mlngOrdemAscDesc = flexSortCustom
        Else
            If mlngOrdemAscDesc <> flexSortGenericAscending Then
                Set .CellPicture = imgCheck.ListImages(grdAsc).Picture
                mlngOrdemAscDesc = flexSortGenericAscending
            Else
                Set .CellPicture = imgCheck.ListImages(grdDesc).Picture
                mlngOrdemAscDesc = flexSortGenericDescending
            End If
        End If
        
        .CellPictureAlignment = flexAlignRightCenter
        .Sort = mlngOrdemAscDesc
        .Redraw = True
    End With
End Sub

Private Sub CountMarcados()
    Dim lngcount                    As Long
    Dim intIndex                    As Long
    Dim i                           As Long
                  
    lngcount = 0
    With fgConsulta
        If .TextMatrix(1, e_Tipo) <> "" Then
            For intIndex = 1 To .Rows - 1
                .Row = intIndex
                If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                    lngcount = lngcount + 1
                End If
                i = i + 1
            Next
        End If
    End With
    etxTotalSelecionados.valorInteiro = lngcount
End Sub

Private Function LinhaSelecionada(lngLinha As Long) As Boolean
    If lngLinha <= fgConsulta.Rows - 1 Then
        fgConsulta.Row = lngLinha
        fgConsulta.col = 0
        LinhaSelecionada = (fgConsulta.CellPicture = imgCheck.ListImages(2).Picture)
    Else
        LinhaSelecionada = False
    End If
End Function

Private Sub fgConsulta_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    ' Luiz Satto - 22/09/2016 - Protocolo 402923
    With fgConsulta ' Coluna 5 = Data de vencimento
        If mbolOrdDtAscendente Then
            Set .CellPicture = imgCheck.ListImages(grdAsc).Picture
            Cmp = IIf(CDate(.TextMatrix(Row1, 5)) > CDate(.TextMatrix(Row2, 5)), 1, -1)
        Else
            Set .CellPicture = imgCheck.ListImages(grdDesc).Picture
            Cmp = IIf(CDate(.TextMatrix(Row1, 5)) < CDate(.TextMatrix(Row2, 5)), 1, -1)
        End If
    End With
End Sub

Private Sub fgConsulta_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Luiz Satto - 22/09/2016 - Protocolo 402923
    If (fgConsulta.MouseRow = 0) And (fgConsulta.MouseCol > 0) Then
        OrdenarGrid fgConsulta, fgConsulta.MouseCol
    ElseIf fgConsulta.MouseRow = 0 And fgConsulta.MouseCol = 0 Then
        fgConsulta.Row = 0
    End If
End Sub

Private Sub Form_Load()
    Aplicacao.Connect
    Set mobjRemessa = New clsEmissaoRemessa
    Call mobjRemessa.init(Aplicacao)
    Call mobjTitulo.init(Aplicacao)
    Call CenterForm(Me)
    Call fLoadEnterprise_estabelecimento
    Call CarregaCombo
    Call etxCarteira.AddConexao(Aplicacao)
    Call etxBanco.AddConexao(Aplicacao)
    Call etxEmpresa.AddConexao(Aplicacao)
    Call CarregaHFlexGrid(fgConsulta, Nothing, strConsulta)
    
    fgConsulta.RowHeight(0) = fgConsulta.RowHeight(0) + 32
    etxParcelaInicial.Enabled = False
    etxParcelaFinal.Enabled = False
    etxCarteira.Enabled = False
    mlngOrdemAscDesc = 1 ' flexSortGenericAscending
    
    If ModGeral.ReadOnly Then cmdGerar.Enabled = False
End Sub

Private Function CarregaCombo()
    cboOrigem.AddItem "Lançamentos"
    cboOrigem.AddItem "Duplicatas"
    cboOrigem.SelectItem "Duplicatas"
    
    cboTipo.AddItem "Emissão"
    cboTipo.AddItem "Vencimento"
    cboTipo.AddItem "Liberação"
    cboTipo.SelectItem "Vencimento"
End Function

'CARREGA O ENTERPRISE_ID E ESTABELECIMENTO.
Private Sub fLoadEnterprise_estabelecimento()
    mlngEnterpriseId = GetFieldValue("enterprise_id", "Usuários", "usuário = '" & UserName & "'")
    mlngCdEstabelecimento = GetFieldValue("cd_estabelecimento", "Usuários", "usuário = '" & UserName & "'")
End Sub

Private Function CarregaClasse()
    With mobjRemessa
        .Enterprise_id = mlngEnterpriseId
        .Cd_estabelecimento = mlngCdEstabelecimento
        .Banco = etxBanco.valorInteiro
        .Id_carteira = etxCarteira.valorInteiro
        .Origem = cboOrigem.SelectedItem
        .Tipofiltro = cboTipo.SelectedItem
        .Data_inicial = etxDataInicial.Data
        .Data_final = etxDataFinal.Data
        .Numero_inicial = etxNumeroInicial.valorInteiro
        .Parcela_inicial = etxParcelaInicial.valorInteiro
        .Numero_final = etxNumeroFinal.valorInteiro
        .Parcela_final = etxParcelaFinal.valorInteiro
        .Empresa = etxEmpresa.valorTexto
        .Reimpressao = chkImpressao.value
        .CaminhoRemessa = etxCaminhoRemessa.Valor
    End With
End Function

Private Sub CarregaGrid(col As clsColTituloCobrebem)
    Dim CurrentObject                   As clsTituloCobrebem
    Dim i                               As Integer
    
    Call CarregaHFlexGrid(fgConsulta, Nothing, strConsulta)
    
    With fgConsulta
        .AddItem ("")
        .col = 0
        .Row = 0
        If col.Count > 0 Then
            .Rows = col.Count + 1
            col.MoveFirst
            i = 1
            While Not col.EOF
                .Row = i
                Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                .TextMatrix(i, e_nota) = col.CurrentObject.NumeroDocumento
                .TextMatrix(i, e_Parcela) = col.CurrentObject.Parcela
                .TextMatrix(i, e_Tipo) = col.CurrentObject.Tipo
                .TextMatrix(i, e_Empresa) = col.CurrentObject.Empresa
                .TextMatrix(i, e_Vencimento) = col.CurrentObject.DataVencimento
                .TextMatrix(i, e_Valor) = Format(col.CurrentObject.ValorDocumento, "#,##0.00")
                'Demanda 131996 - Davi Brito - 22/07/2016
                .TextMatrix(i, e_RemessaStatus) = col.CurrentObject.RemessaStatus
                i = i + 1
                col.MoveNext
            Wend
        End If
    End With
End Sub

Public Function LibProc(strFuncao As String, Optional lngFuncao As Long) As Boolean
    Dim i                                       As Long
    Dim lngcount                                As Long
    Dim intIndex                                As Long
    Dim objBoleto                               As clsTituloCobrebem
    
    If ModGeral.ReadOnly And strFuncao = WL_NOVO Then
        MsgBox "Sistema em modo Somente Leitura!", vbInformation, NomeModulo
        LibProc = False
        Exit Function
    End If
    
    Select Case strFuncao
        Case WL_CONSULTA
            If fValidaCampos Then
                Set mobjRemessa = New clsEmissaoRemessa
                CarregaClasse
                Call mobjRemessa.init(Aplicacao)
                mobjRemessa.colTitulo = mobjRemessa.carregarConsulta
                If mobjRemessa.MensagemValidacao <> "" Then
                    frmMensagemBoleto.mensagem = mobjRemessa.MensagemValidacao
                    Load frmMensagemBoleto
                    frmMensagemBoleto.Show vbModal
                End If
                If mobjRemessa.colTitulo.Count > 0 Then
                    If Not ModGeral.ReadOnly Then cmdGerar.Enabled = True
                    Call CarregaGrid(mobjRemessa.colTitulo)
                Else
                    Call CarregaHFlexGrid(fgConsulta, Nothing, strConsulta)
                    MsgBox "Nenhum registro encontrado.", vbInformation, NomeModulo
                End If
                etxTotalPedidos.valorInteiro = mobjRemessa.colTitulo.Count
            End If
        Case WL_NOVO
            'Total de registros processados.
            lngcount = 0
            CarregaClasse
            With fgConsulta
                If .TextMatrix(1, e_Tipo) <> "" Then
                    For intIndex = 1 To .Rows - 1
                        .Row = intIndex
                        If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                            lngcount = lngcount + 1
                        Else
                            Set objBoleto = New clsTituloCobrebem
                            objBoleto.NumeroDocumento = .TextMatrix(intIndex, e_nota)
                            objBoleto.Parcela = .TextMatrix(intIndex, e_Parcela)
                            Call mobjRemessa.colTitulo.Remove(objBoleto)
                            Set objBoleto = Nothing
                        End If
                        i = i + 1
                    Next
                End If
                If lngcount = 0 Then
                    MsgBox "Favor selecionar o(s) algum título(s).", vbInformation, NomeModulo
                    Exit Function
                End If
            End With
            If mobjRemessa.colTitulo.Count > 0 Then
                If mobjTitulo.verificaProximoNN(mobjRemessa.Id_carteira) Then
                    mobjTitulo.Origem = mobjRemessa.Origem
                    If mobjTitulo.GeracaoRemessa(mobjRemessa.colTitulo, mobjRemessa.Id_carteira, mobjRemessa.Banco, mobjRemessa.CaminhoRemessa) Then
                        Call CarregaHFlexGrid(fgConsulta, Nothing, strConsulta)
                        etxTotalPedidos.valorInteiro = 0
                        MsgBox "ARQUIVO DE REMESSA GERADO COM SUCESSO." & vbCrLf _
                               & "Arquivo encontra-se no caminho: " & mobjRemessa.CaminhoRemessa, vbInformation, NomeModulo
                    End If
                Else
                    MsgBox mobjTitulo.MensagemValidacao, vbInformation, NomeModulo
                End If
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
    
    If Not (etxDataInicial.IsValidDate And etxDataFinal.IsValidDate) Then
        strMensagem = strMensagem & "Intervalo de datas é de preenchimento obrigatório." & vbCrLf
    End If
    
    If Trim(etxCaminhoRemessa.Valor) = "" Then
        strMensagem = strMensagem & "Caminho do arquivo de remessa é de preenchimento obrigatório." & vbCrLf
    Else
        If Not PathExiste(fSeparaDiretorioArquivo(etxCaminhoRemessa.Valor)) Then
            strMensagem = strMensagem & "Caminho do arquivo de remessa é inválido." & vbCrLf
        End If
    End If
    
    If (etxDataInicial.IsValidDate And etxDataFinal.IsValidDate) Then
        If DateDiff("D", etxDataInicial.Data, etxDataFinal.Data) < 0 Then
            strMensagem = strMensagem & "Data inicial devera ser menor que a data final." & vbCrLf
        End If
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

Private Function fSeparaDiretorioArquivo(ByVal strDiretorioArquivo As String) As String
    Dim i                   As Integer
    Dim intTemp             As Integer
    Dim strDiretorio        As String
    
    i = 1: strDiretorio = ""
    For i = 0 To Len(strDiretorioArquivo)
        intTemp = Len(strDiretorioArquivo) - i
        If Mid(strDiretorioArquivo, intTemp, 1) = "\" Then
            strDiretorio = Mid(strDiretorioArquivo, 1, intTemp)
            i = Len(strDiretorioArquivo)
        End If
    Next
    fSeparaDiretorioArquivo = strDiretorio
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
