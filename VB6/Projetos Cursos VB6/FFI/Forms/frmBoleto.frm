VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD_old.OCX"
Begin VB.Form frmBoleto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão Boleto"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   Icon            =   "frmBoleto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10455
   Begin VB.Frame FrameEmail 
      Caption         =   "Status E-mail"
      Height          =   855
      Left            =   5700
      TabIndex        =   43
      Top             =   4200
      Width           =   3165
      Begin VB.CommandButton cmdEnviadosEmail 
         Caption         =   "&Enviados"
         Height          =   375
         Left            =   1830
         TabIndex        =   45
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton cmdPendentesEmail 
         Caption         =   "&Pendentes"
         Height          =   375
         Left            =   150
         TabIndex        =   44
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Frame 
      Height          =   855
      Index           =   1
      Left            =   2865
      TabIndex        =   24
      Top             =   4200
      Width           =   2800
      Begin VB.CommandButton cmdTodos 
         Caption         =   "T&odos"
         Height          =   375
         Left            =   150
         TabIndex        =   25
         Top             =   300
         Width           =   1185
      End
      Begin VB.CommandButton cmdNenhum 
         Caption         =   "&Nenhum"
         Height          =   375
         Left            =   1470
         TabIndex        =   26
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame Pedidos 
      Caption         =   "Títulos"
      Height          =   855
      Left            =   30
      TabIndex        =   27
      Top             =   4200
      Width           =   2800
      Begin Fox.EBSText etxTotalPedidos 
         Height          =   330
         Left            =   135
         TabIndex        =   30
         Top             =   435
         Width           =   1155
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
         Left            =   1455
         TabIndex        =   31
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
         Left            =   1455
         TabIndex        =   29
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Todos"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   240
         Width           =   450
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgConsulta 
      Height          =   3435
      Left            =   30
      TabIndex        =   38
      Top             =   5070
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6059
      _Version        =   393216
      HighLight       =   2
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraBotoes 
      Height          =   8535
      Left            =   8910
      TabIndex        =   32
      Top             =   -30
      Width           =   1560
      Begin VB.CommandButton cmdEmail 
         Caption         =   "&Enviar por E-mail"
         Height          =   375
         Left            =   90
         TabIndex        =   35
         Top             =   990
         Width           =   1365
      End
      Begin VB.CommandButton cmdConsulta 
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   90
         TabIndex        =   33
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "&Gerar"
         Height          =   375
         Left            =   90
         TabIndex        =   34
         Top             =   570
         Width           =   1365
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   37
         Top             =   1800
         Width           =   1365
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   36
         Top             =   1410
         Width           =   1365
      End
      Begin MSComctlLib.ImageList imgCheck 
         Left            =   360
         Top             =   3840
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
               Picture         =   "frmBoleto.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBoleto.frx":06DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBoleto.frx":0A2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBoleto.frx":0CD9
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Filtros"
      Height          =   4215
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   8835
      Begin VB.Frame FrameFiltroEmail 
         Caption         =   "E-mail"
         Height          =   675
         Left            =   60
         TabIndex        =   39
         Top             =   3480
         Width           =   8715
         Begin VB.OptionButton optEmail 
            Caption         =   "Todos"
            Height          =   435
            Index           =   0
            Left            =   750
            TabIndex        =   42
            Top             =   180
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optEmail 
            Caption         =   "Enviados"
            Height          =   435
            Index           =   1
            Left            =   3660
            TabIndex        =   41
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton optEmail 
            Caption         =   "Pendentes"
            Height          =   435
            Index           =   2
            Left            =   6900
            TabIndex        =   40
            Top             =   180
            Width           =   1065
         End
      End
      Begin Fox.EBSCombo cboTipo 
         Height          =   315
         Left            =   1620
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
         Left            =   1620
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
         Caption         =   "Re-Impressão"
         Height          =   495
         Left            =   1620
         TabIndex        =   23
         Top             =   3120
         Width           =   1485
      End
      Begin Fox.EBSText etxBanco 
         Height          =   330
         Left            =   1620
         TabIndex        =   2
         Top             =   270
         Width           =   4890
         _ExtentX        =   438759
         _ExtentY        =   582
         MaxLength       =   9
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Banco"
         TipoCriterio    =   4
         CampoDescricao  =   "Nome"
         TabelaConsulta  =   "Bancos"
         TamanhoDescricao=   3000
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
         Left            =   1620
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
         Left            =   1620
         TabIndex        =   16
         Top             =   2430
         Width           =   1905
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   6
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
         Left            =   1620
         TabIndex        =   20
         Top             =   2790
         Width           =   1905
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   6
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
         Left            =   5250
         TabIndex        =   18
         Top             =   2430
         Width           =   1905
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   3
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
         Left            =   5250
         TabIndex        =   22
         Top             =   2790
         Width           =   1905
         _ExtentX        =   265
         _ExtentY        =   582
         TipoTexto       =   0
         MaxLength       =   3
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
         Left            =   5250
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
         Left            =   1620
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
         Left            =   1620
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "&Empresa"
         Height          =   195
         Left            =   945
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
         Left            =   960
         TabIndex        =   5
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "P&arcela"
         Height          =   195
         Left            =   4140
         TabIndex        =   21
         Top             =   2865
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Parcela"
         Height          =   195
         Left            =   4140
         TabIndex        =   17
         Top             =   2505
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tít&ulo Final"
         Height          =   195
         Left            =   765
         TabIndex        =   19
         Top             =   2865
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "T&ítulo Inicial"
         Height          =   195
         Left            =   690
         TabIndex        =   15
         Top             =   2475
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   4320
         TabIndex        =   13
         Top             =   2145
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
         Left            =   60
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
         Left            =   420
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
         Left            =   885
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
         Left            =   420
         TabIndex        =   1
         Top             =   345
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngEnterpriseId                                    As Long
Private mlngCdEstabelecimento                               As Long
Private mlngOrdemAscDesc                                    As Long
Private mbolOrdDtAscendente                                 As Boolean
Private mobjBoleto                                          As clsImpressaoBoleto
Private mobjTitulo                                          As New clsTituloCobrebem

Private Const grdChecked = 2
Private Const grdUnchecked = 1
Private Const grdAsc = 3
Private Const grdDesc = 4

Private Const PDF_PRINTERNAME = "PDF Writer - bioPDF"
Private Const PRINTER_PROGID = "bioPDF.PDFPrinterSettings"

Private Const strConsulta = "campo=SeqGrid;label=;tamanho=250|" & _
                            "campo=nota;label=Título;tamanho=1500;tipo=tpColGridInteger|" & _
                            "campo=Parcela;label=Parc.;tamanho=700;tipo=tpColGridInteger|" & _
                            "campo=Tipo;label=Tipo;tamanho=900|" & _
                            "campo=Empresa;label=Empresa;tamanho=1600|" & _
                            "campo=Vencimento;label=Vencimento;tamanho=1200|" & _
                            "campo=[valor Original];label=Valor;tamanho=1000;formato=###,##0.00;tipo=tpColGridInteger|" & _
                            "campo=emailboleto_enviado;label=Status E-mail;tamanho=1300"

Private Enum ENUMColBoleto
    e_nota = 1
    e_Parcela = 2
    e_Tipo = 3
    e_Empresa = 4
    e_Vencimento = 5
    e_Valor = 6
    e_Email = 7
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

Private Sub cmdEmail_Click()
    Call LibProc(WL_PROCESSO)
End Sub

Private Sub cmdGerar_Click()
    Call LibProc(WL_NOVO)
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

Private Sub cmdPendentesEmail_Click()
    Dim intIndex            As Integer
    
    With fgConsulta
        If .TextMatrix(1, e_Tipo) <> "" Then
            For intIndex = 1 To .Rows - 1
                .Row = intIndex
                .col = 0
                If fgConsulta.TextMatrix(intIndex, e_Email) = "Pendente" Then
                    Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                Else
                    Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                End If
            Next
        End If
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

Private Sub cmdEnviadosEmail_Click()
    Dim intIndex            As Integer
    
    With fgConsulta
        If .TextMatrix(1, e_Tipo) <> "" Then
            For intIndex = 1 To .Rows - 1
                .Row = intIndex
                .col = 0
                If fgConsulta.TextMatrix(intIndex, e_Email) = "Enviado" Then
                    Set .CellPicture = imgCheck.ListImages(grdChecked).Picture
                Else
                    Set .CellPicture = imgCheck.ListImages(grdUnchecked).Picture
                End If
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
    Dim objCarteira             As New clsCarteira
    Dim objCarteiraDao          As New clsCarteiraDAO
    
    If etxCarteira.valorInteiro > 0 Then
        If Not mobjBoleto.ExisteCarteira(etxBanco.valorInteiro, etxCarteira.valorInteiro) Then
            MsgBox "Carteira não pertence ao banco selecionado.", vbInformation, NomeModulo
            etxCarteira.Clear
        End If
    End If
    
    If etxCarteira.valorInteiro > 0 Then
        Set objCarteira = New clsCarteira
        Set objCarteiraDao = New clsCarteiraDAO
        Call objCarteiraDao.init(Aplicacao)
        Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, etxCarteira.valorInteiro)
        If Not ModGeral.ReadOnly Then cmdGerar.Enabled = Not objCarteira.banco_Emite_boleto
        #If FOXSQL = 1 Then
            cmdEmail.Enabled = Not objCarteira.banco_Emite_boleto
        #End If
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
    Set mobjBoleto = New clsImpressaoBoleto
    
    Call mobjBoleto.init(Aplicacao)
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
    
    #If FOXSQL = 0 Then
        cmdEmail.ToolTipText = "Funcionalidade disponível somente na versão FOX SQL"
        cmdEmail.Enabled = False
        FrameFiltroEmail.Enabled = False
        FrameEmail.Enabled = False
        optEmail(0).Enabled = False
        optEmail(1).Enabled = False
        optEmail(2).Enabled = False
        cmdPendentesEmail.Enabled = False
        cmdEnviadosEmail.Enabled = False
    #End If
    
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
    With mobjBoleto
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
        If optEmail(0).value Then
            .StatusEmail = e_Todos
        End If
        If optEmail(1).value Then
            .StatusEmail = e_Enviados
        End If
        If optEmail(2).value Then
            .StatusEmail = e_Pendentes
        End If
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
                .TextMatrix(i, ENUMColBoleto.e_nota) = col.CurrentObject.NumeroDocumento
                .TextMatrix(i, ENUMColBoleto.e_Parcela) = col.CurrentObject.Parcela
                .TextMatrix(i, ENUMColBoleto.e_Tipo) = col.CurrentObject.Tipo
                .TextMatrix(i, ENUMColBoleto.e_Empresa) = col.CurrentObject.Empresa
                .TextMatrix(i, ENUMColBoleto.e_Vencimento) = col.CurrentObject.DataVencimento
                .TextMatrix(i, ENUMColBoleto.e_Valor) = Format(col.CurrentObject.ValorDocumento, "#,##0.00")
                .TextMatrix(i, ENUMColBoleto.e_Email) = IIf(col.CurrentObject.Emailboleto_enviado = 0, "Pendente", "Enviado")

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
    Dim objCarteira                             As clsCarteira
    Dim objCarteiraDao                          As clsCarteiraDAO
    
    If ModGeral.ReadOnly And strFuncao = WL_NOVO Then
        MsgBox "Sistema em modo Somente Leitura!", vbInformation, NomeModulo
        LibProc = False
        Exit Function
    End If
    
    Select Case strFuncao
        Case WL_CONSULTA
            If fValidaCampos Then
                Set mobjBoleto = New clsImpressaoBoleto
                CarregaClasse
                Call mobjBoleto.init(Aplicacao)
                mobjBoleto.colTitulo = mobjBoleto.carregarConsulta
                If mobjBoleto.MensagemValidacao <> "" Then
                    frmMensagemBoleto.mensagem = mobjBoleto.MensagemValidacao
                    Load frmMensagemBoleto
                    frmMensagemBoleto.Show vbModal
                End If
                If mobjBoleto.colTitulo.Count > 0 Then
                    If etxCarteira.valorInteiro > 0 Then
                        Set objCarteira = New clsCarteira
                        Set objCarteiraDao = New clsCarteiraDAO
                        Call objCarteiraDao.init(Aplicacao)
                        Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, etxCarteira.valorInteiro)
                        If Not ModGeral.ReadOnly Then cmdGerar.Enabled = Not objCarteira.banco_Emite_boleto
                        #If FOXSQL = 1 Then
                            cmdEmail.Enabled = Not objCarteira.banco_Emite_boleto
                        #End If
                    End If
                    Call CarregaGrid(mobjBoleto.colTitulo)
                Else
                    Call CarregaHFlexGrid(fgConsulta, Nothing, strConsulta)
                    MsgBox "Nenhum registro encontrado.", vbInformation, NomeModulo
                End If
                etxTotalPedidos.valorInteiro = mobjBoleto.colTitulo.Count
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
                            Call mobjBoleto.colTitulo.Remove(objBoleto)
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
            If mobjBoleto.colTitulo.Count > 0 Then
                'Pt. 102597 - Moacir Pfau(09/11/2010)
                If mobjTitulo.verificaProximoNN(mobjBoleto.Id_carteira) Then
                    mobjTitulo.Origem = mobjBoleto.Origem
                    Call mobjTitulo.EmissaoBoleto(mobjBoleto.colTitulo, mobjBoleto.Id_carteira, mobjBoleto.Banco)
                        Call CarregaHFlexGrid(fgConsulta, Nothing, strConsulta)
                        etxTotalPedidos.valorInteiro = 0
                Else
                    MsgBox mobjTitulo.MensagemValidacao, vbInformation, NomeModulo
                End If
            End If
            
        Case WL_PROCESSO
            Dim colRelacaoAnexo     As clsColTituloCobrebem
            Dim colCobrebem         As clsColTituloCobrebem
            Dim objCobreBem         As clsTituloCobrebem
            Dim PastaBoleto         As String
            Dim oEmpresa            As New CEmpresas
            
            'Valida se a impressora BioPDF esta instalada
            If Not PrinterIndex(PDF_PRINTERNAME) >= 0 Then
                MsgBox "Atenção!!!" & vbCrLf & "A Impressora 'PDF Writer - bioPDF' não está instalada." & vbCrLf & "Para utilizar essa rotina é necessário realizar a instalação", vbExclamation, NomeModulo
                Exit Function
            End If
    
            lngcount = 0
            CarregaClasse
            With fgConsulta
                If .TextMatrix(1, e_Tipo) <> "" Then
                    'Gerar os arquivos em PDF.
                    Set colRelacaoAnexo = New clsColTituloCobrebem
                    For intIndex = 1 To .Rows - 1
                        .Row = intIndex
                        If .CellPicture = imgCheck.ListImages(grdChecked).Picture Then
                            lngcount = lngcount + 1
                            
                            Set objBoleto = New clsTituloCobrebem
                            objBoleto.NumeroDocumento = .TextMatrix(intIndex, e_nota)
                            objBoleto.Parcela = .TextMatrix(intIndex, e_Parcela)
                            objBoleto.Empresa = .TextMatrix(intIndex, e_Empresa)
                            objBoleto.Tipo = .TextMatrix(intIndex, e_Tipo)
                            
                            PastaBoleto = CaminhoPasta(pastaProgramas) & "Cobrebem\Boleto\"
                            
                            If Not PathExiste(PastaBoleto) Then
                                MkDir PastaBoleto
                            End If
                            
                            objBoleto.CaminhoPDF = PastaBoleto & objBoleto.Empresa & "_" & objBoleto.NumeroDocumento & "_" & objBoleto.Parcela & ".pdf"
                            
                            If ArquivoExiste(objBoleto.CaminhoPDF) Then
                                Kill objBoleto.CaminhoPDF
                            End If
                            
                            
                            Call colRelacaoAnexo.updateAnexo(objBoleto, objBoleto.CaminhoPDF, objBoleto.NumeroDocumento)

                            Set objCobreBem = New clsTituloCobrebem
                            Set colCobrebem = New clsColTituloCobrebem
                            Set objCobreBem = mobjBoleto.colTitulo.GetItem(objBoleto)
                            
                            Call colCobrebem.add(objCobreBem)
                            mobjTitulo.Origem = mobjBoleto.Origem
                            Call mobjTitulo.EmissaoBoletoEmail(colCobrebem, mobjBoleto.Id_carteira, mobjBoleto.Banco, objBoleto.CaminhoPDF)
                            Call mobjBoleto.colTitulo.Remove(objBoleto)
                            Set objBoleto = Nothing
                        Else
                        End If
                        i = i + 1
                    Next
                    
                    Call CarregaHFlexGrid(fgConsulta, Nothing, strConsulta)
                    etxTotalPedidos.valorInteiro = 0
                End If
                
                Dim sAnexos             As String
                Dim sEmpresa            As String
                Dim item                As New clsTituloCobrebem
                Call Sleep(3000)

                If Not colRelacaoAnexo Is Nothing Then
                    If colRelacaoAnexo.Count > 0 Then
                        colRelacaoAnexo.MoveFirst
                        While Not colRelacaoAnexo.EOF
                            Set item = colRelacaoAnexo.CurrentObject
                        
                            Call oEmpresa.CarregarRegistro(item.Empresa)

                            Call EncaminharEmail(RetornaEmailPara(item.Empresa), oEmpresa.Razao, item.Anexos, cboOrigem.SelectedItem)
                            colRelacaoAnexo.MoveNext
                        Wend
                    End If
                End If
                If lngcount = 0 Then
                    MsgBox "Favor selecionar o(s) algum título(s).", vbInformation, NomeModulo
                    Exit Function
                Else
                    MsgBox "Boleto(s) enviado(s) com sucesso.", vbInformation, NomeModulo
                    Exit Function
                End If
            End With
       
        Case WL_SAIR
            Unload Me
    End Select
End Function

Public Function PrinterIndex(ByVal printername As String) As Integer
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).DeviceName) Like LCase(printername) Then
            PrinterIndex = i
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function

Private Function RetornaEmailPara(ByVal Empresa As String) As String
    Dim oEmpresaContato         As New EmpresasContatos
    Dim colEmpresaContato       As New cColEmpresasContatos
    Dim EmailContato            As String
    Dim oEmpresa                 As New CEmpresas
    
    EmailContato = ""
    Set colEmpresaContato = oEmpresaContato.CarregarColecao(Empresa)
    
    If colEmpresaContato.Count > 0 Then
        colEmpresaContato.MoveFirst
        While Not colEmpresaContato.EOF
            Set oEmpresaContato = colEmpresaContato.CurrentObject
            
            If oEmpresaContato.Enviar_boleto Then
                If EmailContato = "" Then
                    EmailContato = oEmpresaContato.e_mail
                Else
                    EmailContato = EmailContato & ";" & oEmpresaContato.e_mail
                End If
            End If
            colEmpresaContato.MoveNext
        Wend
    End If
    
    'If EmailContato = "" Then
    '    If oEmpresa.CarregarRegistro(Empresa) Then
    '        EmailContato = oEmpresa.Email_nfe
    '    End If
    'End If
    
    RetornaEmailPara = EmailContato
End Function

Private Function fValidaCampos() As Boolean
    Dim strMensagem             As String
    'Pt. 96589 - Moacir Pfau(05/02/2010)
    Dim objCarteira             As New clsCarteira
    Dim objCarteiraDao          As New clsCarteiraDAO

On Error GoTo err
    If etxBanco.valorInteiro = 0 Then
        strMensagem = strMensagem & "Banco é de preenchimento obrigatório." & vbCrLf
    End If
   
    If etxCarteira.valorInteiro = 0 Then
        strMensagem = strMensagem & "Carteira é de preenchimento obrigatório." & vbCrLf
    Else
        'Pt. 96589 - Moacir Pfau(05/02/2010)
        Call objCarteiraDao.init(Aplicacao)
        Set objCarteira = objCarteiraDao.Carregar(mlngEnterpriseId, mlngCdEstabelecimento, etxCarteira.valorInteiro)
        If objCarteira.Banco_gera_nosso_numero Then
            strMensagem = strMensagem & "Não é possível realizar a emissão do(s) boleto(s), pois o banco gera o nosso número." & vbCrLf
        End If
    End If
    
    If Not (etxDataInicial.IsValidDate And etxDataFinal.IsValidDate) Then
        strMensagem = strMensagem & "Intervalo de datas é de preenchimento obrigatório." & vbCrLf
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Aplicacao.Disconnect
End Sub

Private Function EncaminharEmail(ByVal strEmail As String, ByVal strRazao As String, ByVal anexo As clscolAnexosEmail, ByVal Origem As String) As Boolean
    Dim objReq                  As New EnvioEmailReq
    Dim objEnvio                As New EnvioEmail
    Dim objRetorno              As New EnvioEmailResp
    Dim objEnderecoEmailPara    As New EnderecoEmail
    Dim objEnderecoEmailCopia   As New EnderecoEmail
    Dim objAnexo                As anexo
    Dim objListAnexo            As New ListAnexo
    Dim objEmpresa              As New CEmpresas
    Dim strMsg                  As String
    Dim intCont                 As Integer
    Dim strNota                 As String
    Dim strPara()               As String
    Dim i                       As Integer
    
On Error GoTo err
  
    Set objReq.ConfigContaEmail = New ConfigContaEmail
    With objReq.ConfigContaEmail
        .NomeConta = EmpresaUsuaria.Email
        .Servidor = ConfigSys.S_SMTP
        .Protocolo = "Smtp"
        .PORTA = ConfigSys.S_PORT
        .UtilizarConexaoSegura = ConfigSys.UtilizaConexaoSegura
        .usuario = ConfigSys.S_USER
        .Senha = ConfigSys.S_PASS
    End With
        
    
    strPara = Split(strEmail, ";")
        
    Set objReq.MensagemEmail = New MensagemEmail
    With objReq.MensagemEmail
        .Assunto = "Boleto de cobrança"
        
        'Endereços
        Set .De = New EnderecoEmail
        .De.Endereco = EmpresaUsuaria.Email
        For i = 0 To UBound(strPara)
            If i = 0 Then
                Set .Para = New ListEnderecoEmail
                objEnderecoEmailPara.Endereco = strPara(i)
                Call .Para.add(objEnderecoEmailPara)
            Else
                Set .Copia = New ListEnderecoEmail
                objEnderecoEmailCopia.Endereco = strPara(i)
                Call .Copia.add(objEnderecoEmailCopia)
            End If
        Next
        
        
        .TipoCorpo = TipoCorpo_Html
        Set .Anexos = New ListAnexo
        
        'LOOP INICIO
        If anexo.Count > 0 Then
            anexo.MoveFirst
            While Not anexo.EOF
                Set objAnexo = New anexo
                objAnexo.Caminho = anexo.CurrentObject.AnexoEmail
                
                If strNota = "" Then
                    strNota = anexo.CurrentObject.DocumentoEmail
                Else
                    If InStr(1, strNota, anexo.CurrentObject.DocumentoEmail, vbTextCompare) = 0 Then
                        strNota = strNota & ", " & anexo.CurrentObject.DocumentoEmail
                    End If
                End If

                
                objAnexo.TipoAnexo = EnumTipoAnexo_Arquivo
                Call .Anexos.add(objAnexo)
                Set objAnexo = Nothing
                anexo.MoveNext
            Wend
        End If
        'LOOP FINAL
        .Corpo = CompoemMensagem(strRazao, strNota, "", IIf(Origem = "Duplicatas", True, False))
    End With
    
    Set objRetorno = objEnvio.Processar(objReq)
    If Not objRetorno.Erros Is Nothing Then
        If objRetorno.Erros.Count > 0 Then
            For intCont = 0 To objRetorno.Erros.Count - 1
                strMsg = strMsg + objRetorno.Erros.ElementAt(intCont).mensagem + vbNewLine
            Next
            MsgBox "Não foi possível enviar o e-mail: " & vbNewLine & strMsg
            EncaminharEmail = False
        Else
            EncaminharEmail = True
        End If
    Else
        EncaminharEmail = True
    End If
    
    Set objReq = Nothing
    Set objEnvio = Nothing
    Set objRetorno = Nothing
    Set objEnderecoEmailPara = Nothing
    Set objEmpresa = Nothing
err:
    If err.Number <> 0 Then
        MsgBox "Ocorreu um erro ao enviar o email: " & err.Description
    End If
    Set objReq = Nothing
    Set objEnvio = Nothing
    Set objRetorno = Nothing
    Set objEnderecoEmailPara = Nothing
    Set objEmpresa = Nothing
End Function

Public Function CompoemMensagem(strRazao As String, NrNota As String, strTp_Registro As String, Optional blnNota As Boolean = False) As String
    Dim strMensagemHTML     As String
    Dim objTipoGlobal       As CTiposGlobais
    Dim objEmpresaUsuaria   As New CEmpresas
    
    Set objEmpresaUsuaria = EmpresaUsuaria.GetCadastroEmpresa
    
    strMensagemHTML = "<!DOCTYPE html PUBLIC " & Chr(34) & "-//W3C//DTD XHTML 1.0 Transitional//EN" & Chr(34) & " " & Chr(34) & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" & Chr(34) & "><html xmlns=" & Chr(34) & "http://www.w3.org/1999/xhtml" & Chr(34) & "><head><meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=utf-8" & Chr(34) & " /><title>"
    strMensagemHTML = strMensagemHTML & "</title></head><body>"
    strMensagemHTML = strMensagemHTML & "<table width=" & Chr(34) & "700" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " align=" & Chr(34) & "center" & Chr(34) & " cellpadding=" & Chr(34) & "0" & Chr(34) & " cellspacing=" & Chr(34) & "0" & Chr(34) & ""
    strMensagemHTML = strMensagemHTML & "style=" & Chr(34) & "font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif; font-size:12px;" & Chr(34) & ">"
    strMensagemHTML = strMensagemHTML & "<tr> <td colspan=" & Chr(34) & "3" & Chr(34) & "> <img src=" & Chr(34) & "http://www.ebs.com.br/boletim/email_nfe_fox/topo_padrao.jpg" & Chr(34) & ""
    strMensagemHTML = strMensagemHTML & "alt=" & Chr(34) & "" & Chr(34) & " width=" & Chr(34) & "700" & Chr(34) & " height=" & Chr(34) & "110" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " /></td></tr><tr>"
    strMensagemHTML = strMensagemHTML & "<td width=" & Chr(34) & "61" & Chr(34) & " rowspan=" & Chr(34) & "3" & Chr(34) & "><img src=" & Chr(34) & "http://www.ebs.com.br/boletim/email_nfe_fox/lateral_esquerda.jpg" & Chr(34) & "/>"
    strMensagemHTML = strMensagemHTML & "</td><td width=" & Chr(34) & "578" & Chr(34) & ">&#160;</td><td width=" & Chr(34) & "61" & Chr(34) & " rowspan=" & Chr(34) & "3" & Chr(34) & "><img src=" & Chr(34) & "http://www.ebs.com.br/boletim/email_nfe_fox/lateral_direita.jpg" & Chr(34) & "/>"
    strMensagemHTML = strMensagemHTML & "</td></tr><tr>"
    strMensagemHTML = strMensagemHTML & "<td>"

    strMensagemHTML = strMensagemHTML & " <br /><br />Destinatário: <strong>" & UCase(strRazao) & " </strong>"
    strMensagemHTML = strMensagemHTML & " <br /><br />Prezado(a) cliente <br />"
    If blnNota Then
        strMensagemHTML = strMensagemHTML & "Segue(m) o(s) boleto(s) bancário(s) emitido(s) pela nossa empresa e de acordo com a(s) Nota(s) Fiscal(is) número " & NrNota & ".<br />"
    Else
        strMensagemHTML = strMensagemHTML & "Segue(m) o(s) boleto(s) bancário(s) emitidos pela nossa empresa e de acordo com o(s) documento(s) " & NrNota & ".<br />"
    End If
   
    strMensagemHTML = strMensagemHTML & " <br />Atenciosamente<br /><br />"
    strMensagemHTML = strMensagemHTML & "<strong> " & objEmpresaUsuaria.Razao & "<br />" & objEmpresaUsuaria.Endereco & "<br />" & objEmpresaUsuaria.Cidade & "/" & objEmpresaUsuaria.GetEstado.Sigla & "<br />" & objEmpresaUsuaria.Fone1 & " </strong></td>"
    strMensagemHTML = strMensagemHTML & "</tr>"
    strMensagemHTML = strMensagemHTML & "<tr><td>&#160;</td></tr><tr><td colspan=" & Chr(34) & "3" & Chr(34) & "><img src=" & Chr(34) & "http://www.ebs.com.br/boletim/email_nfe_fox/rodape.jpg" & Chr(34) & " alt=" & Chr(34) & "" & Chr(34) & " width=" & Chr(34) & "700" & Chr(34) & " height=" & Chr(34) & "73" & Chr(34) & " border=" & Chr(34) & "0" & Chr(34) & " /></td></tr></table>"
    strMensagemHTML = strMensagemHTML & "</body></html>"
    CompoemMensagem = strMensagemHTML
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
