VERSION 5.00
Begin VB.Form frptDuplLancAtrasoNovo 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicatas e Lançamentos em Atraso - Analitico"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7350
   Begin VB.Frame Frame 
      Height          =   4125
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7275
      Begin VB.Frame Frame 
         Height          =   765
         Index           =   3
         Left            =   60
         TabIndex        =   25
         Top             =   3300
         Width           =   7155
         Begin VB.CommandButton cmdFechar 
            Caption         =   "Fechar"
            Height          =   405
            Left            =   5190
            TabIndex        =   27
            Top             =   210
            Width           =   1785
         End
         Begin VB.CommandButton cmdVisualizar 
            Caption         =   "Visualizar"
            Height          =   405
            Left            =   3330
            TabIndex        =   26
            Top             =   210
            Width           =   1785
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Filtros"
         Height          =   3135
         Index           =   2
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   7155
         Begin Fox.EBSText etxNumeroI 
            Height          =   330
            Left            =   1290
            TabIndex        =   9
            Top             =   1320
            Width           =   1335
            _ExtentX        =   767
            _ExtentY        =   582
            MaxLength       =   15
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
         Begin Fox.EBSText etxNotaI 
            Height          =   330
            Left            =   1290
            TabIndex        =   13
            Top             =   1710
            Width           =   1335
            _ExtentX        =   767
            _ExtentY        =   582
            MaxLength       =   15
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
         Begin Fox.EBSData etxEmissaoI 
            Height          =   330
            Left            =   1290
            TabIndex        =   17
            Top             =   2100
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
         Begin Fox.EBSData etxVencimentoI 
            Height          =   330
            Left            =   1290
            TabIndex        =   21
            Top             =   2490
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
         Begin Fox.EBSText etxEmpresa 
            Height          =   330
            Left            =   1290
            TabIndex        =   3
            Top             =   210
            Width           =   5535
            _ExtentX        =   440531
            _ExtentY        =   582
            Tipo            =   4
            MaxLength       =   15
            PossuiDescricao =   -1  'True
            CampoCriterio   =   "Apel"
            CampoDescricao  =   "Razão"
            TabelaConsulta  =   "Empresas"
            TamanhoDescricao=   4000
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
         Begin Fox.EBSText etxNumeroF 
            Height          =   330
            Left            =   3240
            TabIndex        =   11
            Top             =   1350
            Width           =   1335
            _ExtentX        =   767
            _ExtentY        =   582
            MaxLength       =   15
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
         Begin Fox.EBSData etxEmissaoF 
            Height          =   330
            Left            =   3240
            TabIndex        =   19
            Top             =   2100
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
         Begin Fox.EBSData etxVencimentoF 
            Height          =   330
            Left            =   3240
            TabIndex        =   23
            Top             =   2490
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
         Begin Fox.EBSText etxNotaF 
            Height          =   330
            Left            =   3240
            TabIndex        =   15
            Top             =   1710
            Width           =   1335
            _ExtentX        =   767
            _ExtentY        =   582
            MaxLength       =   15
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
         Begin Fox.EBSReport ertRelatorio 
            Height          =   795
            Left            =   90
            TabIndex        =   24
            Top             =   2880
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1402
            NomeRelatorio   =   "FOXFVF10174.ERC"
         End
         Begin Fox.EBSCombo cboOrigem 
            Height          =   315
            Left            =   1290
            TabIndex        =   5
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            OrigemDados     =   2
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
         Begin Fox.EBSCombo cboTipo 
            Height          =   315
            Left            =   1290
            TabIndex        =   7
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            OrigemDados     =   2
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
         Begin VB.Label lblOrigem 
            AutoSize        =   -1  'True
            Caption         =   "Ori&gem:"
            Height          =   195
            Index           =   0
            Left            =   660
            TabIndex        =   4
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblTipo 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Index           =   1
            Left            =   810
            TabIndex        =   6
            Top             =   1020
            Width           =   360
         End
         Begin VB.Label lblNota 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nota:"
            Height          =   195
            Left            =   825
            TabIndex        =   12
            Top             =   1785
            Width           =   390
         End
         Begin VB.Label lblNumero 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Left            =   615
            TabIndex        =   8
            Top             =   1395
            Width           =   600
         End
         Begin VB.Label lblEmissao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Emissão:"
            Height          =   195
            Left            =   585
            TabIndex        =   16
            Top             =   2175
            Width           =   630
         End
         Begin VB.Label lblVencimento 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            Height          =   195
            Left            =   330
            TabIndex        =   20
            Top             =   2565
            Width           =   885
         End
         Begin VB.Label lblEmpresa 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Empresa:"
            Height          =   195
            Left            =   540
            TabIndex        =   2
            Top             =   285
            Width           =   660
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   2880
            TabIndex        =   10
            Top             =   1395
            Width           =   90
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   2880
            TabIndex        =   14
            Top             =   1785
            Width           =   90
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   2910
            TabIndex        =   18
            Top             =   2175
            Width           =   90
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "á"
            Height          =   195
            Left            =   2910
            TabIndex        =   22
            Top             =   2565
            Width           =   90
         End
      End
   End
End
Attribute VB_Name = "frptDuplLancAtrasoNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Projeto: #218 - História: # - Problema# - João Henrique(05/10/2012)

Private Sub cmdFechar_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdVisualizar_Click()
    Dim strRelatorio As String

    ertRelatorio.ClearParametro
    
    If etxNumeroI.valorInteiro <> 0 Then
        ertRelatorio.AddParametro "vNumIni", etxNumeroI.valorInteiro
    End If

    If etxNumeroF.valorInteiro <> 0 Then
        ertRelatorio.AddParametro "vNumFim", etxNumeroF.valorInteiro
    End If

    If etxNotaI.valorInteiro <> 0 Then
        ertRelatorio.AddParametro "vNotaIni", etxNotaI.valorInteiro
    End If

    If etxNotaF.valorInteiro <> 0 Then
        ertRelatorio.AddParametro "vNotaFim", etxNotaF.valorInteiro
    End If

    If etxEmpresa.valorTexto <> "" Then
        ertRelatorio.AddParametro "vEmp", etxEmpresa.valorTexto
    End If

    If etxEmissaoI.IsValidDate Then
        ertRelatorio.AddParametro "vEmiIni", etxEmissaoI.Data
    End If

    If etxEmissaoF.IsValidDate Then
        ertRelatorio.AddParametro "vEmiFim", etxEmissaoF.Data
    End If

    If etxVencimentoI.IsValidDate Then
        ertRelatorio.AddParametro "vVencIni", etxVencimentoI.Data
    End If

    If etxVencimentoF.IsValidDate Then
        ertRelatorio.AddParametro "vVencFim", etxVencimentoF.Data
    End If
    
    If cboTipo.SelectedItem <> "Todos" Then
        ertRelatorio.AddParametro "vTipo", Left(cboTipo.SelectedItem, 1)
    End If

    If cboOrigem.SelectedItem <> "Todos" Then
        ertRelatorio.AddParametro "vOrigem", Left(cboOrigem.SelectedItem, 1)
    End If
    
    ertRelatorio.EnterpriseId = EnterpriseId
    ertRelatorio.UserGroup = GetFieldValue("grupo", "Usuários", "usuário = '" & UserName & "'", , "")

    If ReadSettings("PARAMETROS", "app_remoto", "") <> "" Then
        strRelatorio = ReadSettings("PARAMETROS", "app_remoto", "") & "Programas\ERC\RELS\FOXFIN00222.ERC"
    Else
        strRelatorio = ReadSettings("PARAMETROS", "app_local", "") & "Programas\ERC\RELS\FOXFIN00222.ERC"
    End If
    ertRelatorio.NomeRelatorio = "FOXFIN00222.ERC"
    ertRelatorio.CaminhoConfiguracao = ArquivoConfiguracao

    ertRelatorio.NumeroCopias = 1
    ertRelatorio.CaminhoImpressora = "PDF Writer - bioPDF"
    ertRelatorio.EscModel = emNone
    ertRelatorio.OEMConvert = False
    'ertRelatorio.NumeroCopias = lngNrCopia
    ertRelatorio.Visualizador = ReadSettings("PARAMETROS", "app_local", "") & "Programas\fre.exe"
    ertRelatorio.ArquivoExecucao = ReadSettings("PARAMETROS", "app_local", "") & "Programas\LancamentoDuplicataAtrasoAnalitico.xml"
    ertRelatorio.LoginUsuario = UserName
    ertRelatorio.SenhaUsuario = GetFieldValue("senha", "Usuários", "usuário = '" & UserName & "'", , "")
    ertRelatorio.Visualizar
End Sub

Private Sub etxEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSql As String

    If KeyCode = vbKeyPageDown Then
        strSql = "SELECT Apel, Razão, Pessoa, Tipo, [CNPJ/CPF], [IEst/RG], CCM, " _
        & "Ramo, Endereço, Bairro, CEP, Cidade, Estado, " _
        & "Região, País, Fone1, Ramal1, Contato, Dpto " _
        & "FROM Empresas"
        ' Verifica a configuração para separar as empresas por tipo
        PCampo "Empresas", strSql, PB_CAMPO, etxEmpresa, "Apel"
    End If
End Sub

Private Sub etxEmpresa_LostFocus()
    If Trim(etxEmpresa.valorTexto) <> "" Then
       ' Call DemonstrarInformacaoAdicional
    End If
End Sub

Private Function preencheCombo()
    Call preencheComboTipo
    Call preencheComboOrigem
End Function

Private Sub Form_Load()
    Aplicacao.Connect
    preencheCombo
    Call etxEmpresa.AddConexao(Aplicacao)
    Call etxNumeroI.AddConexao(Aplicacao)
    Call etxNumeroF.AddConexao(Aplicacao)
    Call etxNotaI.AddConexao(Aplicacao)
    Call etxNotaF.AddConexao(Aplicacao)
    Aplicacao.Disconnect
End Sub

Private Sub preencheComboOrigem()
    Dim strDefault          As String
    Dim i                   As Integer
    Dim ArrOrigem()         As String

    strDefault = "Todos"
    ArrOrigem = Split("Todos;Duplicatas;Lançamentos", ";")
    For i = 0 To UBound(ArrOrigem)
        cboOrigem.AddItem ArrOrigem(i)
    Next

    cboOrigem.SelectItem strDefault

End Sub

Private Sub preencheComboTipo()
    Dim strDefault          As String
    Dim i                   As Integer
    Dim ArrTipo()           As String

    strDefault = "Todos"
    ArrTipo = Split("Todos;Pagar;Receber", ";")
    For i = 0 To UBound(ArrTipo)
        cboTipo.AddItem ArrTipo(i)
    Next

    cboTipo.SelectItem strDefault

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
