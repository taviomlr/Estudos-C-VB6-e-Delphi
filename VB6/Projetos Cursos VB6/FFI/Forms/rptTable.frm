VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmTabelas 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatórios Gerais"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "rptTable.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTabelas 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   31
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdTabelas 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   30
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdTabelas 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   29
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame fraTab 
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
      Height          =   2535
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin VB.ComboBox cboTabelas 
         Height          =   315
         Index           =   1
         ItemData        =   "rptTable.frx":0C42
         Left            =   1320
         List            =   "rptTable.frx":0C4C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cboTabelas 
         Height          =   315
         Index           =   0
         ItemData        =   "rptTable.frx":0C5E
         Left            =   1320
         List            =   "rptTable.frx":0C6E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtTabelas 
         Height          =   315
         Index           =   1
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtTabelas 
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblTblDesc 
         Caption         =   "lblTblDesc(1)"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   34
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   2130
      End
      Begin VB.Label lblTblDesc 
         Caption         =   "lblTblDesc(0)"
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   33
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   2130
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "&Ordem:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "Código &Final:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "Código &Inicial:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame fraTab 
      Caption         =   "Centro de Custo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2535
      Index           =   2
      Left            =   240
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cboTabelas 
         Height          =   315
         Index           =   4
         ItemData        =   "rptTable.frx":0CA1
         Left            =   1680
         List            =   "rptTable.frx":0CAB
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtTabelas 
         Height          =   315
         Index           =   7
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtTabelas 
         Height          =   315
         Index           =   6
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblTblDesc 
         Caption         =   "lblTblDesc(7)"
         Height          =   195
         Index           =   7
         Left            =   3000
         TabIndex        =   40
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblTblDesc 
         Caption         =   "lblTblDesc(6)"
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   39
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1890
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "&Ordem do Relatório:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "Código &Final:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "Código &Inicial:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   990
      End
   End
   Begin VB.Frame fraTab 
      Caption         =   "Grupos de Contas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2535
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cboTabelas 
         Height          =   315
         Index           =   3
         ItemData        =   "rptTable.frx":0CC2
         Left            =   1560
         List            =   "rptTable.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2040
         Width           =   1935
      End
      Begin VB.ComboBox cboTabelas 
         Height          =   315
         Index           =   2
         ItemData        =   "rptTable.frx":0CE3
         Left            =   1560
         List            =   "rptTable.frx":0CED
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtTabelas 
         Height          =   315
         Index           =   5
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   17
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtTabelas 
         Height          =   315
         Index           =   4
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtTabelas 
         Height          =   315
         Index           =   3
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtTabelas 
         Height          =   315
         Index           =   2
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblTblDesc 
         Caption         =   "lblTblDesc(5)"
         Height          =   195
         Index           =   5
         Left            =   2760
         TabIndex        =   38
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   2130
      End
      Begin VB.Label lblTblDesc 
         Caption         =   "lblTblDesc(4)"
         Height          =   195
         Index           =   4
         Left            =   2760
         TabIndex        =   37
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   2130
      End
      Begin VB.Label lblTblDesc 
         Caption         =   "lblTblDesc(3)"
         Height          =   195
         Index           =   3
         Left            =   2760
         TabIndex        =   36
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   2130
      End
      Begin VB.Label lblTblDesc 
         Caption         =   "lblTblDesc(2)"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   35
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   2130
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "Ordem das Conta&s:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "&Ordem dos Grupos:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "Conta &Final:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "&Conta Inicial:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   915
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "Gr&upo Final:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblrptBancos 
         AutoSize        =   -1  'True
         Caption         =   "G&rupo Inicial:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   930
      End
   End
   Begin ComctlLib.TabStrip tabTabelas 
      Height          =   3015
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5318
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Bancos"
            Key             =   "bancos"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Grupos de Contas"
            Key             =   "contas"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Centro de Custo"
            Key             =   "centros"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTabelas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboTabelas_GotFocus(Index As Integer)
  StatusMsg cboTabelas(Index).TabIndex
End Sub

Private Sub cmdTabelas_Click(Index As Integer)
  Select Case Index
  '
  ' Botão Visualizar
  Case 0
    PTabelas wrToWindow
  '
  Case 1
    PTabelas wrToPrinter
  '
  Case 2
    Unload Me
  End Select
  
End Sub

Private Sub Form_Load()
  'Centraliza o formulário na tela
  CenterForm Me
  
  cboTabelas(0).ListIndex = 0     'Tab de Bancos
  cboTabelas(1).ListIndex = 0
  cboTabelas(2).ListIndex = 0     'Tab de Grupos e Contas
  cboTabelas(3).ListIndex = 0
  cboTabelas(4).ListIndex = 0     'Tab de Centro de Custo
  '
  ' Se o usuário não desenha controlar o Centro de Custo eu
  ' removo a guia que exibe este relatório
  '
  If (Not CentrodeCusto(MFinanceiro)) Then
    tabTabelas.Tabs.Remove 3    'Centros
  End If
  '
  ' Carregando os números dos primeiros bancos, Contas, Grupos e Centro de Custo
  '
  txtTabelas(0).Text = MinValue("Banco", "Bancos", NUL)
  txtTabelas(1).Text = MaxValue("Banco", "Bancos", NUL)
  
  txtTabelas(2).Text = MinValue("Código", "Grupos", NUL)
  txtTabelas(3).Text = MaxValue("Código", "Grupos", NUL)
  txtTabelas(4).Text = MinValue("Código", "Contas", NUL)
  txtTabelas(5).Text = MaxValue("Código", "Contas", NUL)
  
  txtTabelas(6).Text = MinValue("Código", "Centros", NUL)
  txtTabelas(7).Text = MaxValue("Código", "Centros", NUL)
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MsgBar App.ProductName
End Sub

Private Sub tabTabelas_Click()
  fraTab(0).Visible = (tabTabelas.SelectedItem.Key = "bancos")
  fraTab(1).Visible = (tabTabelas.SelectedItem.Key = "contas")
  fraTab(2).Visible = (tabTabelas.SelectedItem.Key = "centros")
End Sub

Private Sub txtTabelas_Change(Index As Integer)

  Select Case Index
  '
  Case 0, 1          'Bancos
    If IsValid(txtTabelas(Index).Text) Then
      GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtTabelas(Index).Text, _
                    lblTblDesc(Index)
    Else
      lblTblDesc(Index).Caption = NUL
    End If
  '
  Case 2, 3           'Grupos
    If IsValid(txtTabelas(Index).Text) Then
      GetAssocValue "SELECT Descrição FROM Grupos WHERE Código = " & txtTabelas(Index).Text, _
                    lblTblDesc(Index)
    Else
      lblTblDesc(Index).Caption = NUL
    End If
  '
  Case 4, 5           'Contas
    If IsValid(txtTabelas(Index).Text) Then
      GetAssocValue "SELECT Descrição FROM Contas WHERE Código = " & txtTabelas(Index).Text, _
                    lblTblDesc(Index)
    Else
      lblTblDesc(Index).Caption = NUL
    End If
  '
  Case 6, 7           'Centro de Custo
    If IsValid(txtTabelas(Index).Text) Then
      GetAssocValue "SELECT Descrição FROM Centros WHERE Código = " & txtTabelas(Index).Text, _
                    lblTblDesc(Index)
    Else
      lblTblDesc(Index).Caption = NUL
    End If
  '
  End Select
  
End Sub

Private Sub txtTabelas_GotFocus(Index As Integer)
  Selecione txtTabelas(Index)
  StatusMsg txtTabelas(Index).TabIndex
End Sub

' SUB.......: StatusMsg
' Objetivo..: Exibe mensagens na barra de status do Sistema.
' Argumento.: [intTabIndex]: TabIndex do controle atual.
' --------------------------------------------------------------------------
Private Sub StatusMsg(intTabIndex As Integer)
  Select Case intTabIndex
  ' Código inicial e código final de Bancos
  Case 2, 4
    MsgBar "Código do Banco" & ResolveResString(75, resUM, "Bancos")
  ' Tipo de relatório de Banco
  Case 6
    MsgBar "Típos de relatório disponíveis"
  ' Ordem do relatório de Bancos e Centros de Custo
  Case 8, 28
    MsgBar "Lista os campos em que o relatório será ordenado"
  ' Código do grupo
  Case 11, 13
    MsgBar "Código do Grupo de Contas" & ResolveResString(75, resUM, "Grupos de Contas")
  ' Código da conta
  Case 15, 17
    MsgBar "Código das Contas" & ResolveResString(75, resUM, "Contas")
  ' Ordem dos campos dos grupos
  Case 19
    MsgBar "Lista os campos em que os Grupos podem ser ordenados"
  ' Ordem dos campos das contas
  Case 21
    MsgBar "Lista os campos em que as Contas podem ser ordenadas"
  ' Código do Centro de Custo
  Case 24, 25
    MsgBar "Código do Centro de Custo" & ResolveResString(75, resUM, "Centro de Custo")
  '
  End Select
End Sub

' SUB.......: PTabelas
' Objetivo..: Verifica qual o relatório que deve ser impresso e direciona
'             para a função correta.
' Argumento.: [lDestino]: Destino da impressão.
' --------------------------------------------------------------------------
Private Sub PTabelas(lDestino As Long)
  Select Case tabTabelas.SelectedItem.Key
  '
  Case "bancos"
    ImprimeBanco lDestino
  '
  Case "contas"
    ImprimeGrupos lDestino
  '
  Case "centros"
    ImprimeCentros lDestino
  '
  End Select
End Sub

' SUB.......: ImprimeBanco
' Objetivo..: Imprime o relatório de bancos conforme o necessário.
' Argumento.: [pDestino]: Destino da impressão.
' ----------------------------------------------------------------------------------
Private Sub ImprimeBanco(pDestino As Long)
Dim lngBcoInicial As Long
Dim lngBcoFinal   As Long
Dim strORDERBY    As String
  '
  ' Verificando as informações digitadas pelo usuário
  '
  lngBcoInicial = CLngDef(txtTabelas(0).Text, 0)
  lngBcoFinal = CLngDef(txtTabelas(1).Text, 999999999)
  strORDERBY = IIf(cboTabelas(1).ListIndex, "Nome", "Banco")
  
  Select Case cboTabelas(0).ListIndex
  '
  ' Relatório Simples
  Case 0
    BancoSimples pDestino, lngBcoInicial, lngBcoFinal, strORDERBY
  '
  ' Relatório de Endereços
  Case 1
    BancosEnderecos pDestino, CStr(lngBcoInicial), CStr(lngBcoFinal), strORDERBY
  '
  ' Relatório de Contatos
  Case 2
    BancosContatos pDestino, CStr(lngBcoInicial), CStr(lngBcoFinal), strORDERBY
  '
  ' Relatório de Ficha Cadastral
  Case 3
    BancosCadastro pDestino, CStr(lngBcoInicial), CStr(lngBcoFinal), strORDERBY
  '
  End Select
  
End Sub

' SUB.......: BancoSimples
' Objetivo..: Imprime o relatório bancário Simples.
' Argumentos: [lDevice]    : Destino da impressão.
'             [lBcoInicial]: Número do Banco inicial.
'             [lBcoFinal]  : Número do Banco final.
'             [strOrdem]   : Campo que ordenará a tabela.
' -----------------------------------------------------------------------------------
Private Sub BancoSimples(lDevice As Long, lBcoInicial As Long, lBcoFinal As Long, strOrdem As String)
Dim strSelectBanco As String
Dim rstBcoSimples  As Object

  strSelectBanco = "SELECT Banco, Nome, Fone, Contato, Fax FROM Bancos " & _
                   "WHERE Banco BETWEEN " & CStr(lBcoInicial) & " AND " & _
                   CStr(lBcoFinal) & " ORDER BY " & strOrdem
  'Pt. 96013 - Moacir Pfau(20/11/2009)
  If (AbreRecordset(rstBcoSimples, strSelectBanco, dbOpenSnapshot) = WL_OK) Then
    Const SUB_TIT$ = "SubTitulo"
    Const GRP_DADOS$ = "Dados"
    Const GRP_FOT$ = "Rodapé"
    
    Dim rptTabelas     As KeybReport
    Dim cmpTemp        As Campo
    Dim secTemp        As Secao
    Dim sngPos As Single          'Configura a posição Left de cada Campo
    
    Set rptTabelas = New KeybReport
    Set rptTabelas.Recordset = rstBcoSimples
    '
    ' Configurando algumas propriedades do gerador
    '
    rptTabelas.Destino = lDevice
    rptTabelas.Tipo = wrObjectDraw
    rptTabelas.AutoRedraw = True
    rptTabelas.WindowTitulo = "Bancos - Relatório Simples"
    rptTabelas.ScaleMode = vbMillimeters
    '
    ' Cria o cabeçalho do relatório
    '
    PageHeader rptTabelas, "Bancos - Relatório Simples"
    '
    ' Criando o primeiro grupo como sub-cabeçalho
    '
    rptTabelas.FontSize = 8
    rptTabelas.FontStyle = wrFSBold
    
    rptTabelas.AddGrupo SUB_TIT, , , True
    Set secTemp = rptTabelas.Grupo(SUB_TIT).AddSecao(scHeader, 1)
    secTemp(1).DrawBorder = wrDBBottomBorder
    '
    ' Coluna Banco
    '
    Set cmpTemp = secTemp.Linha(1).AddCampo(, wrCSFixedText, "Banco", wrTAEsquerdo)
    cmpTemp.Width = 70
    '
    ' Coluna Telefone
    '
    Set cmpTemp = secTemp.Linha(1).AddCampo(, wrCSFixedText, "Fone", wrTAEsquerdo)
    cmpTemp.Width = 27
    '
    ' Coluna Contato
    '
    Set cmpTemp = secTemp.Linha(1).AddCampo(, wrCSFixedText, "Contato", wrTAEsquerdo)
    cmpTemp.Width = 44
    '
    ' Coluna Fax
    '
    Set cmpTemp = secTemp.Linha(1).AddCampo(, wrCSFixedText, "Fax", wrTAEsquerdo)
    '
    ' Adicionando o grupo de impressão dos registros
    '
    rptTabelas.FontStyle = wrFSNormal
    
    rptTabelas.AddGrupo GRP_DADOS
    Set secTemp = rptTabelas.Grupo(GRP_DADOS).AddSecao(scDetalhe, 1)
    '
    ' Coluna Número do Banco
    '
    Set cmpTemp = secTemp.Linha(1).AddCampo(, wrCSDataField, "Banco", wrTADireito)
    cmpTemp.Width = 17
    cmpTemp.Formato = "000000000"
    '
    ' Coluna Nome do Banco
    '
    secTemp.Linha(1).AddCampo , , "Nome", wrTAEsquerdo
    secTemp.Linha(1).Campo(2).Width = 50
    Set cmpTemp = secTemp.Linha(1).Campo(2)
    '
    ' Coluna Fone
    '
    secTemp.Linha(1).AddCampo , , "Fone", wrTAEsquerdo
    secTemp.Linha(1).Campo(3).Width = rptTabelas.Grupo(SUB_TIT).Header.Linha(1).Campo(2).Width
    '
    ' Coluna Contato
    '
    secTemp.Linha(1).AddCampo , , "Contato", wrTAEsquerdo
    secTemp.Linha(1).Campo(4).Width = rptTabelas.Grupo(SUB_TIT).Header.Linha(1).Campo(3).Width
    '
    ' Coluna Fax
    '
    secTemp.Linha(1).AddCampo , , "Fax", wrTAEsquerdo
    secTemp.Linha(1).Campo(5).Width = rptTabelas.Grupo(SUB_TIT).Header.Linha(1).Campo(4).Width
    '
    ' Finalizando o relatório
    '
    rptTabelas.AddGrupo GRP_FOT, wrDBTopBorder, wrVPNoFinal, True
    rptTabelas(GRP_FOT).Height = rptTabelas.TextHeight("W")
    '
    ' Exibindo os dados
    '
    Set rptTabelas.DatabaseName = GlobalDataBase
    rptTabelas.BeginPrint gTipoDB
    rptTabelas.EndPrint
    
    FechaRecordset rstBcoSimples
    Set cmpTemp = Nothing
    Set secTemp = Nothing
    Set rptTabelas = Nothing
    
  ElseIf (UltimoRetorno = WL_NORECORD) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
  End If
  
End Sub

' SUB.......: BancosEnderecos
' Objetivo..: Imprime o relatório de endereços dos Bancos
' Argumentos: [pdDest]  : Destino da impressão.
'             [sInicial]: Código do Banco Inicial.
'             [sFinal]  : Código do Banco Final.
'             [sOrdem]  : Campo que ordenará a tabela.
' ----------------------------------------------------------------------------------
Private Sub BancosEnderecos(pdDest As PrintDestinoEnum, sInicial As String, sFinal As String, sOrdem As String)
Dim strBancos As String
Dim rstBancos As Object

  strBancos = "SELECT Banco, Nome, Agência, Bairro, Endereço FROM Bancos WHERE " & _
              "Banco BETWEEN " & sInicial & " AND " & sFinal & " ORDER BY " & _
              sOrdem & ";"
  If (AbreRecordsetDAO(rstBancos, strBancos, dbOpenSnapshot) = WL_OK) Then
    Dim kgrBcoEnd As KeybReport
    Dim secTemp   As Secao
    
    Set kgrBcoEnd = New KeybReport
    
    With kgrBcoEnd
      .AutoRedraw = True
      .ScaleMode = vbMillimeters
      Set .Recordset = rstBancos
      .Tipo = wrObjectDraw
      .Destino = pdDest
      .WindowTitulo = "Bancos - Relatório de Endereços"
      '
      ' Criando o cabeçalho do relatório
      '
      PageHeader kgrBcoEnd, "Bancos - Relatório de Endereços"
      '
      ' Grupo de Sub-Titulo
      '
      .FontSize = 8
      .FontStyle = wrFSBold
      .AddGrupo "1", , , True
      Set secTemp = .Grupo("1").AddSecao(scHeader, 1, wrDBBottomBorder)
      ' Coluna Banco
      secTemp(1).AddCampo , wrCSFixedText, "Banco", wrTAEsquerdo
      secTemp(1).Campo(1).Width = 65
      ' Coluna Agência
      secTemp(1).AddCampo , wrCSFixedText, "Agência", wrTAEsquerdo
      secTemp(1).Campo(2).Width = 27
      ' Coluna Bairro
      secTemp(1).AddCampo , wrCSFixedText, "Bairro", wrTAEsquerdo
      secTemp(1).Campo(3).Width = 27
      ' Coluna Endereço
      secTemp(1).AddCampo , wrCSFixedText, "Endereço", wrTAEsquerdo
      '
      ' Criando o grupo de exibição dos dados
      '
      .FontStyle = wrFSNormal
      .AddGrupo "2"
      Set secTemp = .Grupo("2").AddSecao(scDetalhe, 1)
      ' Coluna Banco
      secTemp(1).AddCampo , wrCSDataField, "Banco", wrTADireito
      secTemp(1).Campo(1).Width = 17
      secTemp(1).Campo(1).Formato = "000000000"
      ' Coluna Nome
      secTemp(1).AddCampo , wrCSDataField, "Nome", wrTAEsquerdo
      secTemp(1).Campo(2).Width = 48
      ' Coluna Agência
      secTemp(1).AddCampo , wrCSDataField, "Agência", wrTAEsquerdo
      secTemp(1).Campo(3).Width = 27
      ' Coluna Bairro
      secTemp(1).AddCampo , wrCSDataField, "Bairro", wrTAEsquerdo
      secTemp(1).Campo(4).Width = 27
      ' Coluna Endereço
      secTemp(1).AddCampo , wrCSDataField, "Endereço", wrTAEsquerdo
      '
      ' Finalizando com uma linha
      '
      .AddGrupo "3", wrDBTopBorder, wrVPNoFinal, True
      .Grupo("3").Height = .TextHeight("W")
    End With

    kgrBcoEnd.BeginPrint gTipoDB
    kgrBcoEnd.EndPrint
    
    Set secTemp = Nothing
    Set kgrBcoEnd = Nothing
    
  ElseIf (UltimoRetorno = WL_NORECORD) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
  End If
  
  FechaRecordset rstBancos
    
End Sub

' SUB.......: BancosContatos
' Objetivo..: Imprime o relatório de bancos por contato.
' Argumentos: [pdeDestino]: Destino da impressão.
'             [strInicio] : Código do Banco Inicial.
'             [strFinal]  : Código do Banco final.
'             [strOrdem]  : Nome do campo de ordem.
' ----------------------------------------------------------------------------------
Private Sub BancosContatos(pdeDestino As PrintDestinoEnum, strINICIO As String, strFinal As String, strOrdem As String)
Dim strBco As String
Dim rstBco As Object

  strBco = "SELECT Banco, Nome, Contato, Departamento, Fone, Fax FROM Bancos " & _
           "WHERE Banco BETWEEN " & strINICIO & " AND " & strFinal & " ORDER BY " & _
           strOrdem & ";"
  'Pt. 96013 - Moacir Pfau(20/11/2009)
  If (AbreRecordsetDAO(rstBco, strBco, dbOpenSnapshot) = WL_OK) Then
    Dim kgrBcoContato As KeybReport
    Dim secContatos   As Secao
    
    Set kgrBcoContato = New KeybReport
    With kgrBcoContato
      Set .Recordset = rstBco
      .WindowTitulo = "Bancos - Relatório de Contatos"
      .ScaleMode = vbMillimeters
      .AutoRedraw = True
      .Tipo = wrObjectDraw
      .Destino = pdeDestino
      '
      ' Cabeçalho do relatório
      '
      PageHeader kgrBcoContato, "Bancos - Relatório de Contatos"
      '
      ' Título das colunas
      '
      .FontSize = 8
      .FontStyle = wrFSBold
      .AddGrupo "1", , , True
      Set secContatos = .Grupo(1).AddSecao(scHeader, 1, wrDBBottomBorder)
      ' Coluna Banco
      secContatos(1).AddCampo , wrCSFixedText, "Banco", wrTAEsquerdo
      secContatos(1).Campo(1).Width = 65
      ' Coluna Contato
      secContatos(1).AddCampo , wrCSFixedText, "Contato", wrTAEsquerdo
      secContatos(1).Campo(2).Width = 35
      ' Coluna Departamento
      secContatos(1).AddCampo , wrCSFixedText, "Departamento", wrTAEsquerdo
      secContatos(1).Campo(3).Width = 35
      ' Coluna Fone
      secContatos(1).AddCampo , wrCSFixedText, "Fone", wrTAEsquerdo
      secContatos(1).Campo(4).Width = 25
      ' Coluna Fax
      secContatos(1).AddCampo , wrCSFixedText, "Fax", wrTAEsquerdo
      '
      ' Criando o Grupo de exibição dos dados
      '
      .FontStyle = wrFSNormal
      .AddGrupo "2"
      Set secContatos = .Grupo(2).AddSecao(scDetalhe, 1)
      ' Coluna Banco
      secContatos(1).AddCampo , wrCSDataField, "Banco", wrTAEsquerdo
      secContatos(1).Campo(1).Width = 17
      secContatos(1).Campo(1).Formato = "000000000"
      ' Coluna Nome
      secContatos(1).AddCampo , wrCSDataField, "Nome", wrTAEsquerdo
      secContatos(1).Campo(2).Width = 47
      ' Coluna Contato
      secContatos(1).AddCampo , wrCSDataField, "Contato", wrTAEsquerdo
      secContatos(1).Campo(3).Width = 35
      ' Coluna Departamento
      secContatos(1).AddCampo , wrCSDataField, "Departamento", wrTAEsquerdo
      secContatos(1).Campo(4).Width = 35
      ' Coluna Fone
      secContatos(1).AddCampo , wrCSDataField, "Fone", wrTAEsquerdo
      secContatos(1).Campo(5).Width = 25
      ' Coluna Fax
      secContatos(1).AddCampo , wrCSDataField, "Fax", wrTAEsquerdo
      '
      ' Finalizando o relatório
      '
      .AddGrupo "3", wrDBTopBorder, wrVPNoFinal, True
      .Grupo(3).Height = .TextHeight("W")
    End With
    kgrBcoContato.BeginPrint gTipoDB
    kgrBcoContato.EndPrint
    Set secContatos = Nothing
    Set kgrBcoContato = Nothing
    
  ElseIf (UltimoRetorno = WL_NORECORD) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
  End If
  
  FechaRecordset rstBco
  
End Sub

' SUB.......: BancosCadastro
' Objetivo..: Imprime o relatório de Bancos - Ficha Cadastral
' Argumentos: [pdeDest]: Destino da impressão.
'             [strInit]: Código do Banco Inicial.
'             [strFim] : Código do Banco Final.
'             [strOrd] : Campo em que a select deve ser ordenada.
' ----------------------------------------------------------------------------------
Private Sub BancosCadastro(pdeDest As PrintDestinoEnum, strInit As String, strFim As String, strOrd As String)
Dim strCadastro As String
Dim rstCadastro As Object

  strCadastro = "SELECT Banco, Nome, Endereço, Bairro, Agência, Conta, Contato, " & _
                "Fone, Fax FROM Bancos WHERE Banco BETWEEN " & strInit & " AND " & _
                strFim & " ORDER BY " & strOrd & ";"
    'Pt. 96013 - Moacir Pfau(20/11/2009)
  If (AbreRecordsetDAO(rstCadastro, strCadastro, dbOpenSnapshot) = WL_OK) Then
    Dim kgrCadastro As KeybReport
    Dim secCadastro As Secao
    
    Set kgrCadastro = New KeybReport
    With kgrCadastro
      Set .Recordset = rstCadastro
      .WindowTitulo = "Bancos - Ficha Cadastral"
      .ScaleMode = vbMillimeters
      .AutoRedraw = True
      .Tipo = wrObjectDraw
      .Destino = pdeDest
      '
      ' Cabeçalho do relatório
      '
      PageHeader kgrCadastro, "Bancos - Ficha Cadastral"
      '
      ' Grupo de exibição dos dados do relatório
      '
      .FontSize = 8
      .AddGrupo "1"
      Set secCadastro = .Grupo(1).AddSecao(scDetalhe, 4, wrDBBottomBorder)
      secCadastro.BorderStyle = wrDot
      ' Linha 1: Banco e Nome
      With secCadastro(1)
        .AddCampo , wrCSFixedText, "Banco:"
        .Campo(1).Width = 20
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSDataField, "Banco", wrTADireito
        .Campo(2).Width = 17
        .AddCampo , wrCSDataField, "Nome"
      End With
      ' Linha 2: Endereço e Bairro
      With secCadastro(2)
        .AddCampo , wrCSFixedText, "Endereço:"
        .Campo(1).Width = 20
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSDataField, "Endereço"
        .Campo(2).Width = 80
        .AddCampo , wrCSDataField, "Bairro"
      End With
      ' Linha 3: Agência e Conta
      With secCadastro(3)
        .AddCampo , wrCSFixedText, "Agência:"
        .Campo(1).Width = 20
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSDataField, "Agência"
        .Campo(2).Width = 80
        .AddCampo , wrCSFixedText, "Conta:"
        .Campo(3).FontStyle = wrFSBold
        .Campo(3).Width = 20
        .AddCampo , , "Conta"
      End With
      ' Linha 4: Contato, Fone e Fax
      With secCadastro(4)
        .AddCampo , wrCSFixedText, "Contato"
        .Campo(1).Width = 20
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , , "Contato"
        .Campo(2).Width = 80
        .AddCampo , wrCSFixedText, "Fone e Fax:"
        .Campo(3).FontStyle = wrFSBold
        .Campo(3).Width = 20
        .AddCampo , , "Fone"
        .Campo(4).Width = 28
        .AddCampo , , "Fax"
      End With
      ' Finalizando o relatório
      .AddGrupo "2", wrDBTopBorder, wrVPNoFinal, True
      .Grupo(2).Height = .TextHeight("W")
    End With

    kgrCadastro.BeginPrint gTipoDB
    kgrCadastro.EndPrint
    
    Set secCadastro = Nothing
    Set kgrCadastro = Nothing
  
  ElseIf (UltimoRetorno = WL_NORECORD) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
  End If
  
  FechaRecordset rstCadastro
      
End Sub

' SUB.......: ImprimeGrupos
' Objetivo..: Imprime o relatório de contas e grupos de contas.
' Argumento.: [pdeDestino]: Destino da impressão
' -------------------------------------------------------------------------
Private Sub ImprimeGrupos(pdeDestino As PrintDestinoEnum)
Dim strContas As String
Dim rstContas As Object
  '
  ' Resolvendo os filtros do usuário
  '
  strContas = "SELECT Grupos.Código as Grupos_Código, Grupos.Descrição as Grupos_Descrição, Contas.Código as Contas_Código, " & _
              "Contas.Descrição as Contas_Descrição FROM Grupos, Contas WHERE " & _
              "Contas.Grupo = Grupos.Código"
  ' Seleção de grupos
  '
  If IsValid(txtTabelas(2).Text) Then     'Grupo Inicial
    AppendStr strContas, " AND Grupos.Código >= " & txtTabelas(2).Text
  End If
  
  If IsValid(txtTabelas(3).Text) Then     'Grupo Final
    AppendStr strContas, " AND Grupos.Código <= " & txtTabelas(3).Text
  End If
  '
  ' Seleção de Contas
  '
  If IsValid(txtTabelas(4).Text) Then     'Conta Inicial
    AppendStr strContas, " AND Contas.Código >= " & txtTabelas(4).Text
  End If
  
  If IsValid(txtTabelas(5).Text) Then     'Conta Final
    AppendStr strContas, " AND Contas.Código <= " & txtTabelas(5).Text
  End If
  '
  ' Ordem dos dados
  '
  Concat strContas, " ORDER BY Grupos.", IIf(cboTabelas(2).ListIndex, "Descrição", "Código")
  Concat strContas, ", Contas.", IIf(cboTabelas(3).ListIndex, "Descrição;", "Código;")
  '
  ' Abre o recordset e verifica se dados foram encontrados
  '
  'Pt. 96013 - Moacir Pfau(20/11/2009)
  If (AbreRecordsetDAO(rstContas, strContas, dbOpenSnapshot) = WL_OK) Then
    Dim kgrContas As KeybReport
    Dim secContas As Secao
    '
    ' Montando o relatório
    '
    Set kgrContas = New KeybReport
    With kgrContas
      .AutoRedraw = True
      .Destino = pdeDestino
      .ScaleMode = vbMillimeters
      .Tipo = wrObjectDraw
      .WindowTitulo = "Grupos e Contas"
      Set .Recordset = rstContas
      
      PageHeader kgrContas, "Grupos e Contas"       'Cabeçalho do relatório
      
      .FontSize = 8
      .AddGrupo "1", wrDBBottomBorder
      .Grupo(1).Quebra = "Grupos_Código"
      .Grupo(1).BorderStyle = wrDashDot
      '
      ' Seção com o número dos grupos
      '
      Set secContas = .Grupo(1).AddSecao(scHeader, 3)
      '
      ' Nota: A primeira linha da seção é apenas uma linha em branco
      '
      With secContas.Linha(2)
        .DrawBorder = wrDBBottomBorder
        .BorderStyle = wrSolid
        .AddCampo , wrCSFixedText, "Grupo:"
        .Campo(1).FontStyle = wrFSBold
        .Campo(1).Width = 15
        .AddCampo , , "Grupos_Código"
        .Campo(2).Left = 16
        .Campo(2).Width = 20
        .Campo(2).Formato = "000000000"
        .AddCampo , , "Grupos_Descrição"
        .Campo(3).Left = 38
      End With
      With secContas.Linha(3)
        .AddCampo , wrCSFixedText, "Contas:"
        .Campo(1).FontStyle = wrFSBold
        .Campo(1).Left = 16
      End With
      '
      ' Seção com os dados das contas
      '
      Set secContas = .Grupo(1).AddSecao(scDetalhe, 1)
      With secContas.Linha(1)
        .AddCampo , , "Contas_Código"
        .Campo(1).Left = 38
        .Campo(1).Width = 20
        .Campo(1).Formato = "000000000"
        .AddCampo , , "Contas_Descrição"
        .Campo(2).Left = 60
      End With
      '
      ' Finalizando o relatório
      '
      .AddGrupo "2", wrDBTopBorder, wrVPNoFinal, True
      .Grupo(2).Height = .TextHeight("W")
      
    End With

    kgrContas.BeginPrint gTipoDB
    kgrContas.EndPrint
    
    Set secContas = Nothing
    Set kgrContas = Nothing
  
  ElseIf (UltimoRetorno = WL_NORECORD) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
  End If
  
  FechaRecordset rstContas
  
End Sub

' SUB.......: ImprimeCentros
' Objetivo..: Imprime o relatório de centro de custo
' Argumento.: [pdeDest]: Destino da impressão.
' ---------------------------------------------------------------------------------
Private Sub ImprimeCentros(pdeDest As PrintDestinoEnum)
Dim strCentros As String
Dim rstCentros As Object

  strCentros = "SELECT * FROM Centros WHERE Código BETWEEN "
  '
  ' Centro de custo inicial
  '
  AppendStr strCentros, IIf(IsValid(txtTabelas(6).Text), txtTabelas(6).Text, "0")
  '
  ' Centro de custo final
  '
  Concat strCentros, " AND ", IIf(IsValid(txtTabelas(7).Text), txtTabelas(7).Text, "9999")
  '
  ' Ordem dos dados
  '
  Concat strCentros, " ORDER BY ", IIf(cboTabelas(4).ListIndex, "Descrição", "Código")
  'Pt. 96013 - Moacir Pfau(20/11/2009)
  If (AbreRecordsetDAO(rstCentros, strCentros, dbOpenSnapshot) = WL_OK) Then
    Dim kgrCentros As KeybReport
    Dim secCentros As Secao
    
    Set kgrCentros = New KeybReport
    With kgrCentros
      .AutoRedraw = True
      .WindowTitulo = "Centro de Custo"
      .Tipo = wrObjectDraw
      .ScaleMode = vbMillimeters
      .Destino = pdeDest
      Set .Recordset = rstCentros
      
      PageHeader kgrCentros, "Centro de Custo"    'Cabeçalho do relatório
      '
      ' Títulos das colunas
      '
      .FontSize = 8
      .FontStyle = wrFSBold
      .AddGrupo "Titulos", wrDBBottomBorder, , True
      .Grupo("Titulos").AddSecao scHeader, 2
      .Grupo("Titulos").Header(2).AddCampo , wrCSFixedText, "Código"
      .Grupo("Titulos").Header(2).Campo(1).Width = 12
      .Grupo("Titulos").Header(2).AddCampo , wrCSFixedText, "Descrição"
      .Grupo("Titulos").Header(2).Campo(2).Left = 14
      '
      ' Grupo de dados
      '
      .AddGrupo "1"
      .FontStyle = wrFSNormal
      Set secCentros = .Grupo(1).AddSecao(scDetalhe, 1)
      With secCentros.Linha(1)
        .AddCampo , , "Código", wrTADireito
        .Campo(1).Width = 12
        .Campo(1).Formato = "0000"
        .AddCampo , , "Descrição"
        .Campo(2).Left = 14
      End With
      '
      ' Finalizando a impressão
      '
      .AddGrupo "Final", wrDBTopBorder, wrVPNoFinal, True
      .Grupo("Final").Height = .TextHeight("W")
      
    End With

    kgrCentros.BeginPrint gTipoDB
    kgrCentros.EndPrint
    
    Set secCentros = Nothing
    Set kgrCentros = Nothing
  
  ElseIf (UltimoRetorno = WL_NORECORD) Then
    MsgBox LoadResString(146), vbInformation, MsgBoxCaption
  End If
  
  FechaRecordset rstCentros
  
End Sub

Private Sub txtTabelas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If (Shift = 0) And (KeyCode = vbKeyPageDown) Then
    Select Case UCase(tabTabelas.SelectedItem.Key)
    '
    Case "BANCOS"
      PCampo "Bancos", "Bancos", pbCampo, txtTabelas(Index), "Banco"
    '
    Case "CONTAS"
      If ((Index = 2) Or (Index = 3)) Then
        PCampo "Grupos de Contas", "Grupos", pbCampo, txtTabelas(Index), "Código"
      Else
        PCampo "Contas", "Contas", pbCampo, txtTabelas(Index), "Código"
      End If
    '
    Case "CENTROS"
      PCampo "Centro de Custo", "Centros", pbCampo, txtTabelas(Index), "Código"
    '
    End Select
  End If
End Sub

Private Sub txtTabelas_KeyPress(Index As Integer, KeyAscii As Integer)
  If (tabTabelas.SelectedItem.Key = "bancos") Then
    SetMascara KeyAscii, txtTabelas(Index).SelStart, fMask("Bancos", "Banco"), IIf((Index = 1), txtTabelas(0).hWnd, ZERO)
  ElseIf (tabTabelas.SelectedItem.Key = "contas") Then
    If (Index < 2) Then       '// Index < 2 == Contas
      SetMascara KeyAscii, _
                 txtTabelas(Index).SelStart, _
                 fMask("Contas", "Código"), _
                 IIf((Index = 1), txtTabelas(0).hWnd, ZERO)
    Else                      '// Index > 2 == Grupos de Contas
      SetMascara KeyAscii, _
                 txtTabelas(Index).SelStart, _
                 fMask("Grupos", "Código"), _
                 IIf((Index = 3), txtTabelas(2).hWnd, ZERO)
    End If
  ElseIf (tabTabelas.SelectedItem.Key = "centros") Then
    SetMascara KeyAscii, txtTabelas(Index).SelStart, fMask("Centros", "Código"), IIf((Index = 1), txtTabelas(0).hWnd, ZERO)
  End If
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
