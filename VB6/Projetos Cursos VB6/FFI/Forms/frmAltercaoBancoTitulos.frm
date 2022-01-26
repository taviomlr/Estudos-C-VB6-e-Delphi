VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmAltercaoBancoTitulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração de Banco em Títulos"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   11055
   Begin VB.Frame fraBotoes 
      Height          =   6825
      Left            =   9600
      TabIndex        =   39
      Top             =   -60
      Width           =   1455
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Espéc&ie"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   1035
         Index           =   1
         Left            =   60
         TabIndex        =   43
         Top             =   5730
         Visible         =   0   'False
         Width           =   1320
         Begin VB.OptionButton optBaixas 
            Caption         =   "À Receber"
            Enabled         =   0   'False
            ForeColor       =   &H80000006&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   45
            Top             =   570
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optBaixas 
            Caption         =   "À Pagar"
            Enabled         =   0   'False
            ForeColor       =   &H80000006&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   44
            Top             =   270
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Ca&ncelar"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton cmdVisualizar 
         Caption         =   "&Visualizar"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1860
         Width           =   1215
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1440
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
               Picture         =   "frmAltercaoBancoTitulos.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAltercaoBancoTitulos.frx":0352
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraGeral 
      Height          =   6825
      Left            =   30
      TabIndex        =   27
      Top             =   -60
      Width           =   9570
      Begin VB.Frame fraAlteracao 
         Caption         =   "Alteração"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   46
         Top             =   2220
         Width           =   5070
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Alterar"
            Height          =   335
            Left            =   3120
            TabIndex        =   20
            Top             =   270
            Width           =   1485
         End
         Begin Fox.EBSText etxNovoBanco 
            Height          =   330
            Left            =   1620
            TabIndex        =   19
            Top             =   270
            Width           =   885
            _ExtentX        =   1561
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
            ExibeDescricao  =   0   'False
         End
         Begin VB.Label lblNovoBanco 
            AutoSize        =   -1  'True
            Caption         =   "Novo Banco"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   630
            TabIndex        =   47
            Top             =   330
            Width           =   900
         End
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Ordem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   795
         Index           =   3
         Left            =   5190
         TabIndex        =   38
         Top             =   1380
         Width           =   4320
         Begin VB.OptionButton optTipo 
            Caption         =   "Tipo"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   1560
            TabIndex        =   16
            Top             =   450
            Width           =   915
         End
         Begin VB.OptionButton optEmissao 
            Caption         =   "Emissão"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   210
            Width           =   915
         End
         Begin VB.OptionButton optNotaCod 
            Caption         =   "Nota/Código"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   210
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optEmpresa 
            Caption         =   "Empresa"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   450
            Width           =   1335
         End
         Begin VB.OptionButton optVenc 
            Caption         =   "Vencimento"
            ForeColor       =   &H80000006&
            Height          =   255
            Left            =   3000
            TabIndex        =   17
            Top             =   210
            Width           =   1155
         End
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Tipos de Registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   795
         Index           =   4
         Left            =   60
         TabIndex        =   37
         Top             =   1380
         Width           =   5070
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   3660
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton optDup 
            Caption         =   "Duplicatas"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   480
            TabIndex        =   10
            Top             =   360
            Width           =   1395
         End
         Begin VB.OptionButton optLanc 
            Caption         =   "Lançamentos"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   2040
            TabIndex        =   11
            Top             =   360
            Width           =   1395
         End
      End
      Begin VB.Frame fraBaixas 
         Caption         =   "Selecionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   735
         Index           =   5
         Left            =   5190
         TabIndex        =   36
         Top             =   2220
         Width           =   4320
         Begin VB.CommandButton cmdTodos 
            Caption         =   "Todos"
            Height          =   345
            Left            =   390
            TabIndex        =   21
            Top             =   270
            Width           =   1725
         End
         Begin VB.CommandButton cmdNenhum 
            Caption         =   "Nenhum"
            Height          =   345
            Left            =   2220
            TabIndex        =   22
            Top             =   270
            Width           =   1725
         End
      End
      Begin VB.ComboBox cboBaixas 
         Height          =   315
         Index           =   0
         ItemData        =   "frmAltercaoBancoTitulos.frx":06A4
         Left            =   6150
         List            =   "frmAltercaoBancoTitulos.frx":06A6
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   1830
      End
      Begin Fox.EBSText etxBancoInicial 
         Height          =   330
         Left            =   1230
         TabIndex        =   2
         Top             =   570
         Width           =   1215
         _ExtentX        =   2090
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
         ExibeDescricao  =   0   'False
      End
      Begin Fox.EBSData edtDataVencimentoInicial 
         Height          =   330
         Left            =   6150
         TabIndex        =   8
         Top             =   900
         Width           =   1230
         _ExtentX        =   2170
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
      Begin Fox.EBSData edtDataVencimentoFinal 
         Height          =   330
         Left            =   7650
         TabIndex        =   9
         Top             =   900
         Width           =   1230
         _ExtentX        =   2170
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
      Begin Fox.EBSText etxBancoFinal 
         Height          =   330
         Left            =   2760
         TabIndex        =   3
         Top             =   570
         Width           =   1230
         _ExtentX        =   2090
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
         ExibeDescricao  =   0   'False
      End
      Begin Fox.EBSText etxCodInicial 
         Height          =   330
         Left            =   1230
         TabIndex        =   0
         Top             =   210
         Width           =   1230
         _ExtentX        =   265
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   15
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
      Begin Fox.EBSText etxEmpresas 
         Height          =   330
         Left            =   1230
         TabIndex        =   4
         Top             =   930
         Width           =   3525
         _ExtentX        =   222753
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   15
         PossuiDescricao =   -1  'True
         CampoCriterio   =   "Apel"
         CampoDescricao  =   "Razão"
         TabelaConsulta  =   "Empresas"
         TamanhoDescricao=   2000
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
      Begin Fox.EBSText etxCodFinal 
         Height          =   330
         Left            =   2760
         TabIndex        =   1
         Top             =   210
         Width           =   1230
         _ExtentX        =   265
         _ExtentY        =   582
         Tipo            =   4
         TipoTexto       =   0
         MaxLength       =   15
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
      Begin Fox.EBSData edtDataEmissaoInicial 
         Height          =   330
         Left            =   6150
         TabIndex        =   6
         Top             =   540
         Width           =   1230
         _ExtentX        =   2170
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
      Begin Fox.EBSData edtDataEmissaoFinal 
         Height          =   330
         Left            =   7650
         TabIndex        =   7
         Top             =   540
         Width           =   1230
         _ExtentX        =   2170
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdResultado 
         Height          =   3780
         Left            =   45
         TabIndex        =   42
         ToolTipText     =   "Clique para alterar o banco"
         Top             =   2985
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   6668
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblBaixas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emissão"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   5
         Left            =   5220
         TabIndex        =   41
         Top             =   600
         Width           =   840
      End
      Begin VB.Label lblBaixas 
         AutoSize        =   -1  'True
         Caption         =   "à"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   16
         Left            =   7455
         TabIndex        =   40
         Top             =   600
         Width           =   90
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "à"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   2550
         TabIndex        =   35
         Top             =   270
         Width           =   90
      End
      Begin VB.Label lblBaixas 
         AutoSize        =   -1  'True
         Caption         =   "à"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   7
         Left            =   7455
         TabIndex        =   34
         Top             =   975
         Width           =   90
      End
      Begin VB.Label lblBaixas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vencimento"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   6
         Left            =   5220
         TabIndex        =   33
         Top             =   975
         Width           =   840
      End
      Begin VB.Label lblBaixas 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   32
         Top             =   975
         Width           =   615
      End
      Begin VB.Label lblBaixas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   2
         Left            =   5220
         TabIndex        =   31
         Top             =   255
         Width           =   840
      End
      Begin VB.Label lblBaixas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nota/Código"
         ForeColor       =   &H80000006&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   30
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   645
         TabIndex        =   29
         Top             =   615
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "à"
         ForeColor       =   &H80000006&
         Height          =   195
         Left            =   2535
         TabIndex        =   28
         Top             =   615
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmAltercaoBancoTitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const colUnCheck = 1
Private Const colCheck = 2

'Data.......: 10/10/2008
'Autor......: Ivo Sousa
'Descrição..: Procedimento utilizado para listar os dados localizados pelo filtro.
Private Sub ConfigureGrid()
    Dim intColuna As Integer
    
    With grdResultado
        .Rows = 2
        .Cols = 11
        
        'Coluna Fixa
        .ColWidth(0) = 150
        .TextMatrix(0, 0) = ""
        
        'Coluna de seleção
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 250
        .ColAlignment(1) = flexAlignCenterCenter
        
        'Coluna Câmara
        .ColWidth(2) = 600
        .TextMatrix(0, 2) = "Origem"
        .ColAlignment(2) = flexAlignLeftCenter
        
        'Coluna Agência
        'Vinicius Elyseu(24/05/2016) - Demanda: # 120791
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = "Número"
        .ColAlignment(3) = flexAlignLeftCenter
        
        'Coluna Digito da Agência
        'Vinicius Elyseu(24/05/2016) - Demanda: # 120791
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Tipo"
        .ColAlignment(4) = flexAlignLeftCenter
        
        'Coluna Digito da Agência
        .ColWidth(5) = 450
        .TextMatrix(0, 5) = "Parc."
        .ColAlignment(5) = flexAlignRightCenter
        
        'Coluna Conta Corrente
        .ColWidth(6) = 2400
        .TextMatrix(0, 6) = "Empresa"
        .ColAlignment(6) = flexAlignLeftCenter
        
        'Coluna Digito da Conta Corrente
        .ColWidth(7) = 600
        .TextMatrix(0, 7) = "Banco"
        .ColAlignment(7) = flexAlignRightCenter
        
        'Coluna Digito da Conta Corrente
        .ColWidth(8) = 1050
        .TextMatrix(0, 8) = "Emissão"
        .ColAlignment(8) = flexAlignCenterCenter
        
        'Coluna Digito da Conta Corrente
        .ColWidth(9) = 1050
        .TextMatrix(0, 9) = "Vencimento"
        .ColAlignment(9) = flexAlignCenterCenter
        
        'Coluna Verificadora de Alteração
        .ColWidth(10) = 0
        .TextMatrix(0, 10) = ""
        
        For intColuna = 0 To .Cols - 1
            .TextMatrix(1, intColuna) = ""
            If intColuna = 1 Then
                .col = intColuna
                Set .CellPicture = imgGrid.ListImages(colUnCheck).Picture
            End If
        Next
    End With
End Sub

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdAlterar_Click()
    Dim intCont As Integer
    
    If etxNovoBanco.valorInteiro > 0 Then
        With grdResultado
            .col = 1
            For intCont = 1 To .Rows - 1
                .Row = intCont
                If .CellPicture = imgGrid.ListImages(colCheck).Picture Then
                    .TextMatrix(.Row, 7) = etxNovoBanco.valorInteiro
                    .TextMatrix(.Row, 10) = "True"
                End If
            Next
            etxNovoBanco.valorInteiro = 0
        End With
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call cmdVisualizar_Click
End Sub

Private Sub cmdConfirmar_Click()
    Dim intRegistrosAlt As Integer

    Call GravaAlteracoes(intRegistrosAlt)
    If intRegistrosAlt > 0 Then
        MsgBox "Alterado(s) " & intRegistrosAlt & " registro(s) com sucesso.", vbInformation, NomeModulo
    Else
        MsgBox "Nenhum registro foi alterado.", vbInformation, NomeModulo
    End If
End Sub

Private Sub cmdNenhum_Click()
    Dim intCont As Integer
    
    With grdResultado
        If .TextMatrix(1, 2) <> "" Then
            .col = 1
            For intCont = 1 To grdResultado.Rows - 1
                .Row = intCont
                Set .CellPicture = imgGrid.ListImages(colUnCheck).Picture
            Next
        End If
    End With
End Sub

Private Sub cmdTodos_Click()
    Dim intCont As Integer
    
    With grdResultado
        If .TextMatrix(1, 2) <> "" Then
            .col = 1
            For intCont = 1 To grdResultado.Rows - 1
                .Row = intCont
                Set .CellPicture = imgGrid.ListImages(colCheck).Picture
            Next
        End If
    End With
End Sub

Private Sub cmdVisualizar_Click()
    If etxCodFinal.valorTexto < etxCodInicial.valorTexto Then
        MsgBox "Código inicial não pode ser maior que o código final.", vbInformation, NomeModulo
        etxCodFinal.SetFocus
    Else
        Dim strSql    As String
        Dim rstResult As Object
        
        Call ResolveExp(strSql)
        Call ConfigureGrid
        If AbreRecordset(rstResult, strSql) = WL_OK Then
            Call CarregaRegistrosGrid(rstResult)
        End If
    End If
End Sub

Private Sub etxBancoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If Not etxBancoInicial.ValorDescricao <> "" Then
            etxBancoInicial.valorInteiro = 0
        End If
        Call PCampo("Bancos", "SELECT Banco, Nome FROM Bancos", pbCampo, etxBancoInicial, "Banco")
    End If
End Sub

Private Sub etxBancoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If Not etxBancoFinal.ValorDescricao <> "" Then
            etxBancoFinal.valorInteiro = 0
        End If
        Call PCampo("Bancos", "SELECT Banco, Nome FROM Bancos", pbCampo, etxBancoFinal, "Banco")
    End If
End Sub

Private Sub etxCodFinal_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_inteiro)
End Sub

Private Sub etxCodInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call ConsultaNotas(etxCodInicial)
    End If
End Sub

Private Sub etxCodFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        Call ConsultaNotas(etxCodFinal)
    End If
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub etxCodInicial_KeyPress(KeyAscii As Integer)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_inteiro)
End Sub

Private Sub etxEmpresas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If Not etxEmpresas.ValorDescricao <> "" Then
            etxEmpresas.valorInteiro = 0
        End If
        Call PCampo("Empresas", "SELECT Apel, Razão, Tipo FROM Empresas", pbCampo, etxEmpresas, "Apel")
    End If
End Sub

Private Sub etxNovoBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        If Not etxNovoBanco.ValorDescricao <> "" Then
            etxNovoBanco.valorInteiro = 0
        End If
        Call PCampo("Bancos", "SELECT Banco, Nome FROM Bancos", pbCampo, etxNovoBanco, "Banco")
    End If
End Sub

Private Sub Form_Load()
    Dim strOpcoes As String

    If mstrPagRecAlteraBanco = "P" Then
        optBaixas(0).value = True
    Else
        optBaixas(1).value = True
    End If
    Call ConfigureGrid
    strOpcoes = "SELECT * FROM [Tipos Globais]"
    ComboAddItem cboBaixas(0), strOpcoes, "Tipo"
    cboBaixas(0).AddItem "Todos"
    cboBaixas(0).Text = "Todos"
    Call etxBancoInicial.AddConexao(Aplicacao)
    Call etxBancoFinal.AddConexao(Aplicacao)
    Call etxEmpresas.AddConexao(Aplicacao)
    Call etxNovoBanco.AddConexao(Aplicacao)
End Sub

'SUB.......: ResolveExp
'Objetivo..: Resolve a expressão final de consulta.
'Argumento.: [strVarExpLanc]: Variável que receberá a expressão.
Private Sub ResolveExp(ByRef strVarExp As String)
    Dim dblValorIni    As Double
    Dim dblValorFin    As Double
    Dim strVarExpLanc  As String
    Dim strVarExpDupl  As String
    
    'Verificando qual o tipo da baixa, iniciando a expressão
    If optDup.value Then
        Call ResolveExpDupl(strVarExpDupl)
    ElseIf optLanc.value Then
        Call ResolveExpLancto(strVarExpLanc)
    Else
        Call ResolveExpDuplLanc(strVarExpLanc, strVarExpDupl)
    End If
    
    'Tipo da duplicata
    If Not CompStr(cboBaixas(0).Text, "Todos") And Len(cboBaixas(0).Text) > 0 Then
        Call Concat(strVarExpLanc, " AND Tipo = '", cboBaixas(0).Text, "'")
        Call Concat(strVarExpDupl, " AND Tipo = '", cboBaixas(0).Text, "'")
    End If
    
    'Empresa
    If etxEmpresas.valorTexto <> "" Then
        Call Concat(strVarExpLanc, " AND Empresa = '", etxEmpresas.valorTexto, "'")
        Call Concat(strVarExpDupl, " AND Empresa = '", etxEmpresas.valorTexto, "'")
    End If
    
    'Filtrando Vencimento
    If Not IsEmptyDate(edtDataVencimentoInicial.Data) And Not IsEmptyDate(edtDataVencimentoFinal.Data) Then
        Call Concat(strVarExpLanc, " AND Vencimento BETWEEN ", InverteData(edtDataVencimentoInicial.Data, True), " AND ", InverteData(edtDataVencimentoFinal.Data, True))
        Call Concat(strVarExpDupl, " AND Vencimento BETWEEN ", InverteData(edtDataVencimentoInicial.Data, True), " AND ", InverteData(edtDataVencimentoFinal.Data, True))
    ElseIf Not IsEmptyDate(edtDataVencimentoInicial.Data) Then
        Call Concat(strVarExpLanc, " AND Vencimento >= ", InverteData(edtDataVencimentoInicial.Data, True))
        Call Concat(strVarExpDupl, " AND Vencimento >= ", InverteData(edtDataVencimentoInicial.Data, True))
    ElseIf Not IsEmptyDate(edtDataVencimentoFinal.Data) Then
        Call Concat(strVarExpLanc, " AND Vencimento <= ", InverteData(edtDataVencimentoFinal.Data, True))
        Call Concat(strVarExpDupl, " AND Vencimento <= ", InverteData(edtDataVencimentoFinal.Data, True))
    End If
  
    'Filtrando Emissão
    If Not IsEmptyDate(edtDataEmissaoInicial.Data) And Not IsEmptyDate(edtDataEmissaoFinal.Data) Then
        Call Concat(strVarExpLanc, " AND Emissão BETWEEN ", InverteData(edtDataEmissaoInicial.Data, True), " AND ", InverteData(edtDataEmissaoFinal.Data, True))
        Call Concat(strVarExpDupl, " AND Emissão BETWEEN ", InverteData(edtDataEmissaoInicial.Data, True), " AND ", InverteData(edtDataEmissaoFinal.Data, True))
    ElseIf Not IsEmptyDate(edtDataEmissaoInicial.Data) Then
        Call Concat(strVarExpLanc, " AND Emissão >= ", InverteData(edtDataEmissaoInicial.Data, True))
        Call Concat(strVarExpDupl, " AND Emissão >= ", InverteData(edtDataEmissaoInicial.Data, True))
    ElseIf Not IsEmptyDate(edtDataEmissaoFinal.Data) Then
        Call Concat(strVarExpLanc, " AND Emissão <= ", InverteData(edtDataEmissaoFinal.Data, True))
        Call Concat(strVarExpDupl, " AND Emissão <= ", InverteData(edtDataEmissaoFinal.Data, True))
    End If

    'Filtrando entre pagos e recebidos
    If optBaixas(0).value Then
        Call AppendStr(strVarExpLanc, " AND PagRec = 'P'")
        Call AppendStr(strVarExpDupl, " AND PagRec = 'P'")
    Else
        Call AppendStr(strVarExpLanc, " AND PagRec = 'R'")
        Call AppendStr(strVarExpDupl, " AND PagRec = 'R'")
    End If
    
    'pt. 86113 - Dulcino Júnior(25/03/2008)
    If etxBancoInicial.valorInteiro > 0 Then
        Call AppendStr(strVarExpLanc, " AND Banco >=" & etxBancoInicial.valorInteiro)
        Call AppendStr(strVarExpDupl, " AND Banco >=" & etxBancoInicial.valorInteiro)
    End If
    If etxBancoFinal.valorInteiro > 0 Then
        Call AppendStr(strVarExpLanc, " AND Banco <=" & etxBancoFinal.valorInteiro)
        Call AppendStr(strVarExpDupl, " AND Banco <=" & etxBancoFinal.valorInteiro)
    End If
    
    'Especificando apenas os registros não pagos
    Call AppendStr(strVarExpLanc, " AND (Pagamento IS NULL)")
    Call AppendStr(strVarExpDupl, " AND (Pagamento IS NULL)")
  
    If optDup.value Then
        Call Concat(strVarExpDupl, " ORDER BY ", getOrderBy, ";")
        strVarExp = strVarExpDupl
    ElseIf optLanc.value Then
        'pt. 80029
        'Receber expressao Order By conforme o OptioButton Selecionado
        Call Concat(strVarExpLanc, " ORDER BY ", getOrderBy, ";")
        strVarExp = strVarExpLanc
    Else
        strVarExp = "(" & strVarExpDupl & ") UNION (" & strVarExpLanc & ") ORDER BY " & getOrderBy
    End If
End Sub

'Date.......: 10/10/2008
'Author.....: Ivo Sousa
'Descrição..: Resove a expressão de consulta quando o usuário deseja ver as duplicatas
'Parametros.: [String] Retorno da expressão
Private Sub ResolveExpDupl(strRetorno As String)
    strRetorno = "SELECT 'Dupl' AS Origem, Nota as cod_id, Parcela, Tipo, Empresa, Emissão, Vencimento, Banco FROM Duplicatas WHERE "
    If etxCodInicial.valorTexto <> "" And etxCodFinal.valorTexto <> "" Then
        Call Concat(strRetorno, "Nota BETWEEN ", etxCodInicial.valorTexto, " AND ", etxCodFinal.valorTexto)
    ElseIf etxCodInicial.valorTexto <> "" Then
        Call AppendStr(strRetorno, "Nota > " & etxCodInicial.valorTexto)
    ElseIf etxCodFinal.valorTexto <> "" Then
        Call AppendStr(strRetorno, "Nota < " & etxCodFinal.valorTexto)
    Else
        Call AppendStr(strRetorno, "Nota > 0")
    End If
    Call Concat(strRetorno, " AND Situação <> 'Cancelada'")
End Sub

'Date.......: 10/10/2008
'Author.....: Ivo Sousa
'Descrição..: Resove a expressão de consulta quando o usuário deseja ver os lançamentos
'Parametros.: [String] Retorno da expressão
Private Sub ResolveExpLancto(strResult As String)
    strResult = "SELECT 'Lanc' AS Origem, Código as cod_id, Parcela, Tipo, Empresa, Emissão, Vencimento, Banco FROM Lançamentos WHERE "
    If etxCodInicial.valorTexto <> "" And etxCodFinal.valorTexto <> "" Then
        Call Concat(strResult, "Código BETWEEN ", etxCodInicial.valorTexto, " AND ", etxCodFinal.valorTexto)
    ElseIf etxCodInicial.valorTexto <> "" Then
        Call AppendStr(strResult, "Código > " & etxCodInicial.valorTexto)
    ElseIf etxCodFinal.valorTexto <> "" Then
        Call AppendStr(strResult, "Código < " & etxCodFinal.valorTexto)
    Else
        Call AppendStr(strResult, "Código > 0")
    End If
    Call Concat(strResult, " AND Situação <> 'Cancelada'")
End Sub

'Date.......: 10/10/2008
'Author.....: Ivo Sousa
'Descrição..: Resove a expressão de consulta quando o usuário deseja ver os lançamentos e as duplicatas
'Parametros.: [String] Retorno da expressão
Private Sub ResolveExpDuplLanc(strResultLanc As String, strResultDupl As String, Optional blnConsulta As Boolean)
    strResultLanc = "SELECT 'Lanc' AS Origem, Código as cod_id, Parcela, Lançamentos.Tipo, Empresa, Emissão, Vencimento, Banco FROM Lançamentos WHERE "
    strResultDupl = "SELECT 'Dupl' AS Origem, Nota as cod_id, Parcela, Duplicatas.Tipo, Empresa, Emissão, Vencimento, Banco FROM Duplicatas WHERE "

    If etxCodInicial.valorTexto <> "" And etxCodFinal.valorTexto <> "" Then
        Call Concat(strResultLanc, "Lançamentos.Código BETWEEN ", etxCodInicial.valorTexto, " AND ", etxCodFinal.valorTexto)
        Call Concat(strResultDupl, "Duplicatas.Nota BETWEEN ", etxCodInicial.valorTexto, " AND ", etxCodFinal.valorTexto)
    ElseIf etxCodInicial.valorTexto <> "" Then
        Call AppendStr(strResultLanc, "Lançamentos.Código > " & etxCodInicial.valorTexto)
        Call AppendStr(strResultDupl, "Duplicatas.Nota > " & etxCodInicial.valorTexto)
    ElseIf etxCodFinal.valorTexto <> "" Then
        Call AppendStr(strResultLanc, "Lançamentos.Código < " & etxCodFinal.valorTexto)
        Call AppendStr(strResultDupl, "Duplicatas.Nota < " & etxCodFinal.valorTexto)
    Else
        Call AppendStr(strResultLanc, "Lançamentos.Código > 0")
        Call AppendStr(strResultDupl, "Duplicatas.Nota > 0")
    End If
    Call Concat(strResultLanc, " AND Situação <> 'Cancelada'")
    Call Concat(strResultDupl, " AND Situação <> 'Cancelada'")
End Sub

Private Function getOrderBy() As String
    If optNotaCod.value Then
        If optDup.value Then
            getOrderBy = "Nota"
        ElseIf optLanc.value Then
            getOrderBy = "Código"
        Else
            getOrderBy = "cod_id"
        End If
    End If
    If optEmpresa.value Then
        getOrderBy = "Empresa"
    End If
    If optEmissao.value Then
        getOrderBy = "Emissão"
    End If
    If optVenc.value Then
        getOrderBy = "Vencimento"
    End If
    If optTipo.value Then
        getOrderBy = "Tipo"
    End If
End Function
Private Sub CarregaRegistrosGrid(rstResult As Object)
    Dim intCont As Integer
    
    With rstResult
        intCont = 1
        .MoveFirst
        While Not .EOF
            grdResultado.AddItem ("")
            grdResultado.col = 1
            grdResultado.Row = grdResultado.Rows - 1
            Set grdResultado.CellPicture = imgGrid.ListImages(colUnCheck).Picture
            grdResultado.TextMatrix(intCont, 2) = .Fields("Origem").value
            grdResultado.TextMatrix(intCont, 3) = .Fields("cod_id").value
            grdResultado.TextMatrix(intCont, 4) = .Fields("Tipo").value
            grdResultado.TextMatrix(intCont, 5) = .Fields("Parcela").value
            grdResultado.TextMatrix(intCont, 6) = .Fields("Empresa").value
            grdResultado.TextMatrix(intCont, 7) = .Fields("Banco").value
            grdResultado.TextMatrix(intCont, 8) = .Fields("Emissão").value
            grdResultado.TextMatrix(intCont, 9) = .Fields("Vencimento").value
            grdResultado.TextMatrix(intCont, 10) = "False"
            .MoveNext
            intCont = intCont + 1
        Wend
        If grdResultado.Rows > 2 Then
            grdResultado.RemoveItem (grdResultado.Rows - 1)
        End If
    End With
End Sub

Private Sub ConsultaNotas(txtDestino As EBSText)
    Dim strCodigo     As String
    Dim strVarExpDupl As String
    Dim strVarExpLanc As String
    Dim lngValor      As Long
    
    strCodigo = "SELECT Nota, Código, Tipo, Parcela, Empresa, Descrição, " & _
                "[Valor Original], Acréscimo, Abatimento, Controle, " & _
                "Situação FROM <Tabela> WHERE PagRec = '" & _
                IIf(optBaixas(0).value, "P", "R") & "' AND (Pagamento IS NULL);"

    If optDup.value Then
        DeleteStr strCodigo, ", Código"
        InsereStr strCodigo, "Duplicatas", DeleteStr(strCodigo, "<Tabela>")
        Call PCampo("Duplicatas", strCodigo, pbCampo, txtDestino, "Nota")
    ElseIf optLanc.value Then
        DeleteStr strCodigo, "Nota, "
        DeleteStr strCodigo, "Parcela, "
        InsereStr strCodigo, "Lançamentos", DeleteStr(strCodigo, "<Tabela>")
        Call PCampo("Lançamentos", strCodigo, pbCampo, txtDestino, Código)
    Else
        If etxCodFinal.valorTexto < etxCodInicial.valorTexto Then
            MsgBox "Código inicial não pode ser maior que o código final.", vbInformation, NomeModulo
            etxCodFinal.SetFocus
        Else
            Call ResolveExpDuplLanc(strVarExpDupl, strVarExpLanc, True)
            strCodigo = Replace("(" & strVarExpDupl & ") UNION (" & strVarExpLanc & ") ORDER BY " & getOrderBy, "cod_id", "Numero")
            Call PCampo("Lançamentos/Duplicatas", strCodigo, pbCampo, txtDestino, "Numero")
        End If
    End If
End Sub

Private Sub grdResultado_Click()
    If grdResultado.TextMatrix(grdResultado.Row, 2) <> "" Then
        If grdResultado.col = 1 Then
            If grdResultado.CellPicture = imgGrid.ListImages(colCheck).Picture Then
                Set grdResultado.CellPicture = imgGrid.ListImages(colUnCheck).Picture
            Else
                Set grdResultado.CellPicture = imgGrid.ListImages(colCheck).Picture
            End If
        End If
    End If
End Sub

Private Sub GravaAlteracoes(ByRef intRegistrosAlt As Integer)
    Dim intCont     As Integer
    'Projeto: 100340 - Desenv.: 142888 - Ueder Budni (22/09/2016)
    Dim lngOldBanco As Long
    
    With grdResultado
        If Trim(.TextMatrix(1, 2)) <> "" Then
            For intCont = 1 To .Rows - 1
                If CBool(.TextMatrix(intCont, 10)) Then
                    'Projeto: 100340 - Desenv.: 142888 - Ueder Budni (22/09/2016)
                    lngOldBanco = GetFieldValue("Banco", IIf(.TextMatrix(intCont, 2) = "Dupl", "Duplicatas", "Lançamentos"), IIf(.TextMatrix(intCont, 2) = "Dupl", " Nota = ", " Código = ") & .TextMatrix(intCont, 3) & " AND Empresa = '" & .TextMatrix(intCont, 6) & "' AND Tipo = '" & .TextMatrix(intCont, 4) & "' AND Parcela = " & .TextMatrix(intCont, 5) & " AND PagRec = '" & mstrPagRecAlteraBanco & "'")
                    
                    If .TextMatrix(intCont, 2) = "Dupl" Then
                        Call ExecuteSQL("UPDATE Duplicatas SET Banco = " & .TextMatrix(intCont, 7) & " WHERE Nota = " & .TextMatrix(intCont, 3) & " AND Empresa = '" & .TextMatrix(intCont, 6) & "' AND Tipo = '" & .TextMatrix(intCont, 4) & "' AND Parcela = " & .TextMatrix(intCont, 5) & " AND PagRec = '" & mstrPagRecAlteraBanco & "'")
                    Else
                        Call ExecuteSQL("UPDATE Lançamentos SET Banco = " & .TextMatrix(intCont, 7) & " WHERE Código = " & .TextMatrix(intCont, 3) & " AND Empresa = '" & .TextMatrix(intCont, 6) & "' AND Tipo = '" & .TextMatrix(intCont, 4) & "' AND Parcela = " & .TextMatrix(intCont, 5) & " AND PagRec = '" & mstrPagRecAlteraBanco & "'")
                    End If
                    'Projeto: 100340 - Desenv.: 142888 - Ueder Budni (22/09/2016)
                    Call RegistraLogLancDup(strToDbl(.TextMatrix(intCont, 3)), .TextMatrix(intCont, 6), .TextMatrix(intCont, 4), .TextMatrix(intCont, 5), mstrPagRecAlteraBanco, IIf(.TextMatrix(intCont, 2) = "Dupl", enuLancDup.Duplicata, enuLancDup.Lancamento), lngOldBanco, .TextMatrix(intCont, 7))
                    
                    intRegistrosAlt = intRegistrosAlt + 1
                    .TextMatrix(intCont, 10) = "False"
                End If
            Next
        End If
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

'Projeto: 100340 - Desenv.: 142888 - Ueder Budni (22/09/2016)
Private Sub RegistraLogLancDup(dblNumero As Double, strEmpresa As String, strTipo As String, lngParcela As Long, strPagRec As String, enuTabela As enuLancDup, lngOldBanco As Long, lngNewBanco As Long)
    Dim objLogLancDup   As New clsLogLancamentosDuplicatas

On Error GoTo erro
    Call objLogLancDup.SetKey(strPagRec, dblNumero, strEmpresa, strTipo, lngParcela, enuTabela)
    Call objLogLancDup.InsertCustomMsg("Alterado o banco de {0} para {1} - Via '{2}'.", lngOldBanco, lngNewBanco, Me.Caption)
erro:
    Set objLogLancDup = Nothing
End Sub
