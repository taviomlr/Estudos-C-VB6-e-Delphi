VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frptRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibos"
   ClientHeight    =   5160
   ClientLeft      =   495
   ClientTop       =   1935
   ClientWidth     =   8175
   Icon            =   "rtRecibo.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTab 
      Caption         =   "Recibo de Duplicatas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   7695
      Begin VB.ComboBox cboRecibos 
         Height          =   315
         Index           =   2
         ItemData        =   "rtRecibo.frx":0C42
         Left            =   840
         List            =   "rtRecibo.frx":0C4C
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chkRecibos 
         Caption         =   "&Gerar lançamento desta cobrança"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   42
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtRecibos 
         Height          =   975
         Index           =   7
         Left            =   3720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   6
         Left            =   840
         MaxLength       =   20
         TabIndex        =   39
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   5
         Left            =   840
         MaxLength       =   40
         TabIndex        =   37
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboRecibos 
         Height          =   315
         Index           =   3
         ItemData        =   "rtRecibo.frx":0C62
         Left            =   840
         List            =   "rtRecibo.frx":0C64
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   4
         Left            =   840
         MaxLength       =   15
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   3
         Left            =   840
         MaxLength       =   2
         TabIndex        =   35
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   2
         Left            =   840
         MaxLength       =   9
         TabIndex        =   32
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Pag/Rec:"
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblDescRecibos 
         Caption         =   "lblDescRecibos(2)"
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   33
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   5370
      End
      Begin VB.Label lblDescRecibos 
         Caption         =   "lblDescRecibos(1)"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   30
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   4650
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Obs.:"
         Height          =   195
         Index           =   11
         Left            =   3240
         TabIndex        =   40
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "D&epto.:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   38
         Top             =   2520
         Width           =   525
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Contato:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "Empres&a:"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "Parce&la:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Nota:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   390
      End
   End
   Begin VB.TextBox txtParcela 
      Height          =   315
      Left            =   1680
      TabIndex        =   63
      Top             =   4680
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtLanca 
      Height          =   285
      Left            =   360
      TabIndex        =   62
      Top             =   4710
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox chkLancamentos 
      Caption         =   "Ver..."
      Height          =   315
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Caption         =   "Cartorais"
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
      Height          =   855
      Index           =   2
      Left            =   360
      TabIndex        =   45
      Top             =   3480
      Width           =   7455
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   13
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   53
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   12
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   51
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   11
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   49
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   10
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   47
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Juros:"
         Height          =   195
         Index           =   17
         Left            =   4200
         TabIndex        =   52
         Top             =   480
         Width           =   420
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "De&spesas Cartorais:"
         Height          =   195
         Index           =   16
         Left            =   4200
         TabIndex        =   50
         Top             =   120
         Width           =   1410
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "Ta&xa de Envio ao Cartório:"
         Height          =   195
         Index           =   15
         Left            =   0
         TabIndex        =   48
         Top             =   480
         Width           =   1890
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "Condição do Tít&ulo:"
         Height          =   195
         Index           =   14
         Left            =   0
         TabIndex        =   46
         Top             =   120
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdRecibos 
      Cancel          =   -1  'True
      Caption         =   "#"
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   60
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdRecibos 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   59
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdRecibos 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   58
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame fraLancamentos 
      Caption         =   "Lançamentos selecionados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   2760
      TabIndex        =   54
      Top             =   540
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdAdiciona 
         Caption         =   "Adicionar..."
         Height          =   315
         Left            =   2520
         TabIndex        =   55
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remover"
         Height          =   315
         Left            =   3840
         TabIndex        =   56
         Top             =   3600
         Width           =   1215
      End
      Begin ComctlLib.ListView lvwLancamentos 
         Height          =   3255
         Left            =   120
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         SmallIcons      =   "imgNotaDebitp"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Código"
            Object.Tag             =   ""
            Text            =   "Código"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   "Tipo"
            Object.Tag             =   ""
            Text            =   "Tipo"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   "Descrição"
            Object.Tag             =   ""
            Text            =   "Descrição"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            SubItemIndex    =   3
            Key             =   "Vencimento"
            Object.Tag             =   ""
            Text            =   "Vencimento"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   "Empresa"
            Object.Tag             =   ""
            Text            =   "Empresa"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   "Valor"
            Object.Tag             =   ""
            Text            =   "Valor"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   "Controle"
            Object.Tag             =   ""
            Text            =   "Controle"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   7
            Key             =   "Pagamento"
            Object.Tag             =   ""
            Text            =   "Pagamento"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   8
            Key             =   "Parc"
            Object.Tag             =   ""
            Text            =   "Parcela"
            Object.Width           =   706
         EndProperty
      End
      Begin ComctlLib.ImageList imgNotaDebitp 
         Left            =   240
         Top             =   3480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "rtRecibo.frx":0C66
               Key             =   "lanca"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraTab 
      Caption         =   "Recibo de Contas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7695
      Begin VB.Frame fraTab 
         BorderStyle     =   0  'None
         Height          =   1575
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   3735
         Begin VB.TextBox txtRecibos 
            Height          =   315
            Index           =   17
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   44
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox txtRecibos 
            Height          =   315
            Index           =   16
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   22
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtRecibos 
            Height          =   315
            Index           =   15
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   20
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtRecibos 
            Height          =   315
            Index           =   14
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   18
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label lblRecibos 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   195
            Index           =   20
            Left            =   0
            TabIndex        =   43
            Top             =   1080
            Width           =   465
         End
         Begin VB.Label lblRecibos 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Index           =   19
            Left            =   0
            TabIndex        =   21
            Top             =   720
            Width           =   405
         End
         Begin VB.Label lblRecibos 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            Height          =   195
            Index           =   18
            Left            =   0
            TabIndex        =   19
            Top             =   360
            Width           =   885
         End
         Begin VB.Label lblRecibos 
            AutoSize        =   -1  'True
            Caption         =   "Proveniencia:"
            Height          =   195
            Index           =   4
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.CheckBox chkRecibos 
         Caption         =   "&Gerar lançamento destas custas"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   13
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   9
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   10
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   8
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtRecibos 
         Height          =   1455
         Index           =   1
         Left            =   3960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1080
         Width           =   3615
      End
      Begin VB.ComboBox cboRecibos 
         Height          =   315
         Index           =   1
         ItemData        =   "rtRecibo.frx":0F80
         Left            =   1200
         List            =   "rtRecibo.frx":0F8A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cboRecibos 
         Height          =   315
         Index           =   0
         ItemData        =   "rtRecibo.frx":0FA0
         Left            =   4320
         List            =   "rtRecibo.frx":0FAA
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtRecibos 
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblDescRecibos 
         Caption         =   "lblDescRecibos(0)"
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   8
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   5010
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "Con&tato:"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "Departa&mento:"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Obs.:"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   14
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Lançamento:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   930
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "Tipo da &Conta:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lblRecibos 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         Height          =   195
         Index           =   0
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   360
      End
   End
   Begin ComctlLib.TabStrip tabRecibos 
      Height          =   4575
      Left            =   120
      TabIndex        =   61
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8070
      MultiRow        =   -1  'True
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Recibo de Contas"
            Key             =   "contas"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Recibo de Duplicatas"
            Key             =   "duplicatas"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Despesas Cartorais de Contas"
            Key             =   "contas cartorio"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Despesas Cartorais de Duplicatas"
            Key             =   "duplicatas cartorio"
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
Attribute VB_Name = "frptRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
' Constantes de valores gravados no arquivo .ini
Private Const KEY_GERARDUPL$ = "GerarLanctoDupl"         'chkRecibos(0)
Private Const KEY_GERARCONTA$ = "GerarLanctoContas"      'chkRecibos(1)

Private mstrPagRec    As String

Private Sub cboRecibos_Click(Index As Integer)

  If (Index = 1) Then
    ' Exibe o frame para as contas pagas.
    fraTab(3).Visible = (cboRecibos(1).ListIndex = 0)   ' Contas Pagas
  ElseIf (Index = 0) Then
    If (cboRecibos(0).ListIndex = 1) Then               ' Nota de Débito
      cboRecibos(1).ListIndex = 1                       ' Contas pagas
      lblDescRecibos(0).Visible = False
    Else
      lblDescRecibos(0).Visible = True
    End If
  ElseIf Index = 2 Then
    mstrPagRec = IIf((cboRecibos(2).ListIndex = 0), "R", "P")
  End If
  
  cboRecibos(1).Enabled = (cboRecibos(0).ListIndex = 0)
  lvwLancamentos.ListItems.Clear
  
End Sub

Private Sub chkLancamentos_Click()
  fraLancamentos.Visible = (chkLancamentos.value = vbChecked)
  fraTab(0).Enabled = Not (chkLancamentos.value = vbChecked)
  tabRecibos.Enabled = fraTab(0).Enabled
  chkLancamentos.Caption = IIf((chkLancamentos.value = vbChecked), "Ocultar", "Ver...")
End Sub

Private Sub cmdAdiciona_Click()
Dim strSql    As String
Dim lngLanca  As Double
Dim strPagRec As String
Dim intParcela As Integer
  '
  ' A pagar ou a receber conforme o informado
  '
  strPagRec = IIf((cboRecibos(1).Text = "Pagas"), "P", "R")


  '
  ' Instrução de Consulta
  '
  strSql = "SELECT Código,Parcela, Descrição, Vencimento, Pagamento, Empresa , Banco, " & _
            "[Valor Original],Acréscimo,Abatimento,Tipo," & _
            "Controle FROM Lançamentos WHERE PagRec = " & Quote(strPagRec, "''") & " AND Pagamento IS NOT NULL "


   
  If lvwLancamentos.ListItems.Count > 0 Then
    '
    ' Todos os lançamentos devem ser da mesma empresa
    ' Se o usuário quiser trocar de empresa, basta limpar a lista
    ' que todos os lançamentos serão listados.
    ' Senão, apenas os da empresa do primeiro lançamento selecionado
    '
    Concat strSql, " AND Empresa = '", lvwLancamentos.ListItems(1).SubItems(4), "'"
    
    Dim i As Integer
    For i = 1 To lvwLancamentos.ListItems.Count
      Concat strSql, " AND Código <> ", CStr(lvwLancamentos.ListItems(i).Text)
    Next i
  End If
  
  'If PCampo("Contas " & cboRecibos(1).Text, strSql, pbCampo, lngLanca, "Código") Then
   If PMultiCampo("Contas " & cboRecibos(1).Text, strSql, pbCampo, "Código;Parcela", txtLanca, txtParcela) Then
    'Vinicius Elyseu(25/05/2016) - Projeto: #100340 Demanda: 120791
    lngLanca = CDbl(txtLanca.Text)
    intParcela = CInt(txtParcela.Text)
    lvwLancamentos.ListItems.add , , StrZero(CStr(lngLanca), 6), , "lanca"
    
    With lvwLancamentos.ListItems(lvwLancamentos.ListItems.Count)
      .SubItems(1) = GetFieldValue("Tipo", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))
      .SubItems(2) = GetFieldValue("Descrição", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))
      .SubItems(3) = GetFieldValue("Vencimento", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))
      .SubItems(4) = GetFieldValue("Empresa", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))
      'pt. 87769 - Ivo Sousa (26/09/2008)
      .SubItems(5) = Format$((GetFieldValue("[Valor Original]", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))) + (GetFieldValue("Acréscimo", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela)) - GetFieldValue("Abatimento", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))), FMOEDA)
      .SubItems(6) = GetFieldValue("Controle", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))
      .SubItems(7) = GetFieldValue("Pagamento", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))
      .SubItems(8) = GetFieldValue("Parcela", "Lançamentos", "PagRec = " & Quote(strPagRec, "''") & " AND Código = " & CStr(lngLanca) & " AND parcela=" & CStr(intParcela))
      .Selected = True
    End With
  End If
  
End Sub

Private Sub cmdRecibos_Click(Index As Integer)
  If (Index < 2) Then
    cmdRecibos(0).Enabled = False
    cmdRecibos(1).Enabled = False
    cmdRecibos(2).Caption = LoadResString(IDS_CANCELAR)
    
    FiltroRecibo IIf(Index, wrToPrinter, wrToWindow)
    
    cmdRecibos(0).Enabled = True
    cmdRecibos(1).Enabled = True
    cmdRecibos(2).Caption = LoadResString(IDS_FECHAR)
  Else
    If cmdRecibos(0).Enabled Then
      Unload Me
    Else
      SimpleMsgBar LoadResString(171) & LoadResString(14)
    End If
  End If
End Sub

Private Sub cmdRemove_Click()
  If lvwLancamentos.ListItems.Count > 0 Then lvwLancamentos.ListItems.Remove (lvwLancamentos.SelectedItem.Index)
End Sub

Private Sub Form_Load()
Dim strOpcoes As String
  '
  ' Configurando a abertura do formulário
  '
  cmdRecibos(2).Caption = LoadResString(IDS_FECHAR)
  lblDescRecibos(0).Caption = NUL
  lblDescRecibos(1).Caption = NUL
  lblDescRecibos(2).Caption = NUL
  '
  ' Carregando os tipos de duplicatas da tabela de opções
  '
  strOpcoes = "SELECT Texto FROM Opções WHERE Rotina = '" & OPT_DUPLICATAS & "';"
  ComboAddItem cboRecibos(3), strOpcoes, "Texto"
  If cboRecibos(3).ListCount Then
    cboRecibos(3).ListIndex = 0
  End If
  '
  ' Trazendo a opção gravada do CheckBox
  '
  strOpcoes = IniFileName()
  chkRecibos(0).value = ((LerArquivoASCII(SEC_WKIF, KEY_GERARDUPL, strOpcoes) = "1") And vbChecked)
  chkRecibos(1).value = ((LerArquivoASCII(SEC_WKIF, KEY_GERARDUPL, strOpcoes) = "1") And vbChecked)
  '
  ' Opções padrão das outras caixas combinadas.
  '
  cboRecibos(0).ListIndex = 0           'Recibo de Contas
  cboRecibos(1).ListIndex = 1           'Recebidas
  cboRecibos(2).ListIndex = 0           'Recebidas
  mstrPagRec = "R"
  
  '
  ' Define os primeiros controles visíveis na janela
  '
  tabRecibos.Tabs(1).Selected = True
 
  CenterForm Me
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim strFoxIni As String

  ' Gravando o valor do CheckBox no arquivo .ini
  
  strFoxIni = IniFileName()
  GravarArquivoASCII SEC_WKIF, KEY_GERARDUPL, CStr(chkRecibos(0).value), strFoxIni
  GravarArquivoASCII SEC_WKIF, KEY_GERARCONTA, CStr(chkRecibos(1).value), strFoxIni
  
  Set frptRecibos = Nothing
  
End Sub

Private Sub lvwLancamentos_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  lvwLancamentos.Sorted = True
  lvwLancamentos.SortKey = (ColumnHeader.Index - 1)
  lvwLancamentos.Sorted = False
End Sub

Private Sub lvwLancamentos_KeyDown(KeyCode As Integer, Shift As Integer)
  Call cmdRemove_Click
End Sub

Private Sub tabRecibos_Click()
  ' Alternando os controles conforme o tab
  '
  Select Case tabRecibos.SelectedItem.Key
  '
  Case "contas", "contas cartorio"
    fraTab(0).Caption = tabRecibos.SelectedItem.Caption
  '
  Case "duplicatas", "duplicatas cartorio"
    fraTab(1).Caption = tabRecibos.SelectedItem.Caption
  '
  End Select
  '
  ' Cobrança de despesas cartorais de contas habilitado somente para contas
  ' recebidas. Desabilita a caixa combinada de contas.
  '
  If (tabRecibos.SelectedItem.Key = "contas cartorio") Then
    cboRecibos(1).ListIndex = 1   'Contas Recebidas
    cboRecibos(1).Enabled = False
    chkRecibos(1).Move txtRecibos(8).Left, (txtRecibos(8).Top + 360) 'Posiciona o checkbox
  Else
    cboRecibos(1).Enabled = True
  End If
  '
  ' CheckBox de geração de Lançamento visível apenas quando o usuário selecionar
  ' Despesas Cartorais de Duplicatas ou Desepesas Cartorais de Contas
  '
  chkRecibos(0).Visible = (tabRecibos.SelectedItem.Key = "duplicatas cartorio")
  chkRecibos(1).Visible = (tabRecibos.SelectedItem.Key = "contas cartorio")
  '
  ' ComboBox de Tipos de recibo de Contas visível somente quando Recibo
  '
  lblRecibos(0).Visible = (InStr(1, tabRecibos.SelectedItem.Key, "cartorio") = 0)
  cboRecibos(0).Visible = (InStr(1, tabRecibos.SelectedItem.Key, "cartorio") = 0)
  
  fraTab(0).Visible = (InStr(1, tabRecibos.SelectedItem.Key, "contas") > 0)
  fraTab(1).Visible = (InStr(1, tabRecibos.SelectedItem.Key, "duplicatas") > 0)
  fraTab(2).Visible = (InStr(1, tabRecibos.SelectedItem.Key, "cartorio") > 0)
  If fraTab(2).Visible Then fraTab(2).ZOrder vbBringToFront
  
  chkLancamentos.Visible = ((tabRecibos.Tabs("contas").Selected))

End Sub

Private Sub txtRecibos_Change(Index As Integer)

  Select Case Index
  '
  ' Lançamentos
  Case 0
    If IsValid(txtRecibos(0).Text) Then
      GetAssocValue "SELECT Descrição FROM Lançamentos WHERE Código = " & _
                    txtRecibos(0).Text & " AND PagRec = " & _
                    IIf((cboRecibos(1).ListIndex = 0), "'P'", "'R'"), _
                    lblDescRecibos(0)
    Else
      lblDescRecibos(0).Caption = NUL
    End If
  '
  ' Nota
  Case 2
    If IsValid(txtRecibos(2).Text) Then
      Dim strSelDupl As String              'Select de Duplicatas
      
      If Len(txtRecibos(4).Text) Then       'Nome da Empresa
        strSelDupl = "SELECT Descrição, Tipo, Parcela FROM Duplicatas WHERE " & _
                     "Nota = " & txtRecibos(2).Text & " AND Empresa = '" & _
                     txtRecibos(4).Text & "' AND PagRec = '" & mstrPagRec & "' AND Tipo = '" & _
                     cboRecibos(3).Text & "';"
                     
        GetAssocValue strSelDupl, lblDescRecibos(2), cboRecibos(3), _
                      txtRecibos(3)
      Else
        strSelDupl = "SELECT Descrição, Tipo, Parcela, Empresa FROM Duplicatas " & _
                     "WHERE Nota = " & txtRecibos(2).Text & " AND PagRec = " & _
                     "'" & mstrPagRec & "' AND Tipo = '" & cboRecibos(3).Text & "';"
                     
        GetAssocValue strSelDupl, lblDescRecibos(2), cboRecibos(3), _
                                  txtRecibos(3), txtRecibos(4)
      End If
    End If
  '
  ' Empresa em Duplicatas
  Case 4
    If IsValid(txtRecibos(4).Text) Then
      GetAssocValue "SELECT Razão, Apel FROM Empresas WHERE Apel = '" & _
                    txtRecibos(4).Text & "';", lblDescRecibos(1), txtRecibos(4)
    Else
      lblDescRecibos(1).Caption = NUL
    End If
  '
  End Select
  
End Sub

Private Sub txtRecibos_GotFocus(Index As Integer)
  Selecione txtRecibos(Index)
End Sub

Private Sub txtRecibos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  
  If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
    Dim strSelDados As String     'Para seleção de dados
    
    Select Case Index
    '
    ' Lançamentos
    Case 0
      strSelDados = "SELECT Código, Empresa, Tipo, Descrição, Emissão, " & _
                    "Vencimento, Pagamento, Liberação, [Valor Original], " & _
                    "Acréscimo, Abatimento, Banco, Conta, Centro, Cheque, " & _
                    "Moeda, [Valor da Moeda], Controle FROM " & _
                    "Lançamentos WHERE PagRec = " & _
                    IIf((cboRecibos(1).ListIndex = 0), "'P'", "'R'")
      If PCampo("Lançamentos", strSelDados, pbCampo, txtRecibos(0), "Código") Then
        strSelDados = GetFieldValue("Empresa", "Lançamentos", "Código = " & _
                                    txtRecibos(0).Text & " AND PagRec = " & _
                                    IIf(cboRecibos(1).ListIndex, "'R'", "'P'"))
        strSelDados = "SELECT Contato, Dpto FROM Empresas WHERE Apel = '" & _
                      strSelDados & "';"
        ImportData strSelDados, txtRecibos(9), txtRecibos(8)
      End If
    '
    ' Duplicata
    Case 2
      If Len(txtRecibos(4).Text) Then
        strSelDados = "SELECT Nota, Empresa, Tipo, Parcela, Descrição, Emissão, " & _
                      "Vencimento, Pagamento, Liberação, [Valor Original], " & _
                      "Acréscimo, Abatimento, Banco, Conta, Centro, Cheque, " & _
                      "Moeda, [Valor da Moeda], Controle, Situação FROM " & _
                      "Duplicatas WHERE Empresa = '" & txtRecibos(4).Text & _
                      "' AND PagRec = '" & mstrPagRec & "' AND Tipo = '" & cboRecibos(3).Text & "';"
      Else
        strSelDados = "SELECT Nota, Empresa, Tipo, Parcela, Descrição, Emissão, " & _
                      "Vencimento, Pagamento, Liberação, [Valor Original], " & _
                      "Acréscimo, Abatimento, Banco, Conta, Centro, Cheque, " & _
                      "Moeda, [Valor da Moeda], Controle, Situação FROM " & _
                      "Duplicatas WHERE PagRec = '" & mstrPagRec & "' AND Tipo = '" & _
                      cboRecibos(3).Text & "';"
      End If
      If PMultiCampo("Duplicatas", strSelDados, pbCampo, "Nota;Parcela", txtRecibos(2), txtRecibos(3)) Then
      'If PCampo("Duplicatas", strSelDados, pbCampo, txtRecibos(2), 0) Then
                strSelDados = GetFieldValue("Empresa", "Duplicatas", "Nota = " & _
                                    txtRecibos(2).Text & " AND PagRec = '" & mstrPagRec & "'" & _
                                    " AND Tipo = '" & cboRecibos(3).Text & "'")
        txtRecibos(4).Text = strSelDados
        strSelDados = "SELECT Contato, Dpto FROM Empresas WHERE Apel = '" & _
                      strSelDados & "';"
        ImportData strSelDados, txtRecibos(5), txtRecibos(6)
      End If
      Call txtRecibos_Change(4)
    '
    ' Empresa em Duplicatas
    Case 4
      strSelDados = "SELECT Apel, Razão, Tipo, Cadastro, Pessoa, [CNPJ/CPF], " & _
                    "[IEst/RG], Endereço, Bairro, Cidade, Estado, Fone1, " & _
                    "Ramal1, Contato, Dpto FROM Empresas"
                    
      If (LerArquivoASCII("KinSys", "Separar Empresa por tipo", gstrTempSys) = "S") Then
        AppendStr strSelDados, " WHERE Tipo <> 'Fornecedor';"
      Else
        AppendStr strSelDados, ";"
      End If
      PCampo "Empresas", strSelDados, pbCampo, txtRecibos(4), "Apel"
    '
    End Select
  End If
  
End Sub

Private Sub txtRecibos_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
  '
  ' Lançamento, Nota
  Case 0, 2
    SetMascara KeyAscii, txtRecibos(Index).SelStart, fMask("Duplicatas", "Nota")
  '
  ' Parcela
  Case 3
    SetMascara KeyAscii, txtRecibos(3).SelStart, "##"
  '
  ' Taxa de Envio ao Cartório, Despesas Cartorais, Juros e Valor
  Case 11 To 13, 16
    DMoeda KeyAscii
  '
  ' Vencimento
  Case 15
    SetMascara KeyAscii, txtRecibos(15).SelStart, MASK_DATE4
  '
  End Select
End Sub

Private Sub txtRecibos_LostFocus(Index As Integer)
  ' Importa os dados da empresa para a janela apenas se os campos estiverem
  ' vazios
  Select Case Index
  '
  ' Lançamentos
  Case 0
    If IsValid(txtRecibos(0).Text) Then
      If ((Len(txtRecibos(9).Text) = 0) And (Len(txtRecibos(8).Text) = 0)) Then
        Dim strApel As String             'Chave da empresa
        
        strApel = GetFieldValue("Empresa", "Lançamentos", "Código = " & _
                                txtRecibos(0).Text & " AND PagRec = " & _
                                IIf((cboRecibos(1).ListIndex = 0), "'P'", "'R'"))
        If Len(strApel) Then
          ImportData "SELECT Contato, Dpto FROM Empresas WHERE Apel = '" & _
                     strApel & "';", txtRecibos(9), txtRecibos(8)
        End If
      End If
    End If
  '
  ' Empresa
  Case 4
    If Len(txtRecibos(4).Text) Then
      If ((Len(txtRecibos(5).Text) = 0) And (Len(txtRecibos(6).Text) = 0)) Then
        Dim strEmp As String
        
        strEmp = "SELECT Contato, Dpto FROM Empresas WHERE Apel = '" & _
                 txtRecibos(4).Text & "';"
        ImportData strEmp, txtRecibos(5), txtRecibos(6)
      End If
    End If
  '
  ' Valor
  Case 11 To 13, 16
    Transform txtRecibos(Index), WL_USEREDITNONE, FCURRENCY
  '
  Case 2
    Call txtRecibos_Change(4)

  End Select
  
End Sub

' SUB.......: MsgStatusRecibos
' Objetivo..: Exibe pequenas mensagens de ajuda ao usuário na barra de Status
'             do programa.
' Argumento.: [intTabIndex]: Propriedade TabIndex dos controles.
' --------------------------------------------------------------------------------
Private Sub MsgStatusRecibos(intTabIndex As Integer)

  Select Case intTabIndex
  '
  ' Tipo de Conta
  Case 2
    MsgBar "Conta paga ou recebida"
  '
  ' Tipo
  Case 4
    MsgBar "Tipos de Lançamentos encontrados"
  '
  ' Lançamento
  Case 6
    MsgBar "Código do Lançamento" & ResolveResString(75, resUM, "Lançamentos")
  '
  ' Contato
  Case 9
    MsgBar "Nome da pessoa para contato"
  '
  ' Departamento
  Case 11
    MsgBar "Nome do departamento para contato"
  '
  ' Proveniência
  Case 14
    MsgBar "Identificação do Recibo"
  '
  ' Vencimento
  Case 16
    MsgBar "Data de vencimento do débito"
  '
  ' Valor
  Case 18
    MsgBar "Valor do Lançamento"
  '
  ' Nome
  Case 20
    MsgBar "Nome do assinante do Recibo"
  '
  ' Obs
  Case 22, 40
    MsgBar "Observação para ser impressa"
  '
  ' Tipo (Duplicata)
  Case 25
    MsgBar "Tipos de Duplicatas encontradas"
  '
  ' Empresa
  Case 27
    MsgBar "Nome fantasia da empresa" & ResolveResString(75, resUM, "Empresas")
  '
  ' Nota
  Case 30
    MsgBar "Número da Duplicata" & ResolveResString(75, resUM, "Duplicatas")
  '
  ' Parcela
  Case 33
    MsgBar "Parcela da Duplicata"
  '
  ' Contato (Duplicata)
  Case 35
    MsgBar "Nome da pessoa para contato"
  '
  ' Depto
  Case 37
    MsgBar "Nome do departamento para contato"
  '
  ' Gerar Lançamento...
  Case 38
    MsgBar "Gera um Lançamento ref. à cobrança de despesas cartorais"
  '
  ' Condição do Título
  Case 43
    MsgBar "Forma de pagamento do débito"
  '
  ' Taxa de Envio
  Case 45
    MsgBar "Valor da taxa de envio de duplicata/conta ao cartório"
  '
  ' Taxa de despesa
  Case 47
    MsgBar "Valor das taxas de despesas cartorais"
  '
  ' Juros
  Case 49
    MsgBar "Valor dos juros do débito"
  '
  Case Else
    MsgBar NUL
    
  End Select
  
End Sub

' SUB.......: FiltroRecibo
' Objetivo..: Cria o Filtro e abre o Recordset utilizado para impressão do
'             recibo.
' Argumento.: [pdeDestino]: Constante com o destino da impressão.
' --------------------------------------------------------------------------------
Private Sub FiltroRecibo(pdeDestino As PrintDestinoEnum)
Dim strTabela As String           'Nome da tabela origem
Dim strSelRec As String           'String com o filtro dos dados
Dim rstRecibo As Object        'Recordset para a abertura do registro
Dim wrkGeral  As KeybReport       'Objeto KeybReport
Dim Ok        As Boolean          'Indica se pode seguir

  SetPtr vbArrowHourglass
  
  Set wrkGeral = New KeybReport
  wrkGeral.Tipo = wrObjectDraw
  wrkGeral.AutoRedraw = True
  wrkGeral.ScaleMode = vbMillimeters
  wrkGeral.Destino = pdeDestino
  wrkGeral.WindowTitulo = tabRecibos.SelectedItem.Caption
  wrkGeral.FontName = "Arial"
  wrkGeral.FontSize = 10
  
  Select Case tabRecibos.SelectedItem.Key
  '
  ' Recibo de Contas
  Case "contas"
    If lvwLancamentos.ListItems.Count > 0 Then
      EstruturaContas wrkGeral, cboRecibos(0).Text
      Ok = True
    Else
      MsgFunc LoadResString(183)
      Exit Sub
    End If
  '
  ' Recibo de Duplicatas
  Case "duplicatas"
    If (FiltraDuplicatas(rstRecibo)) Then
      EstruturaRecibo wrkGeral, rstRecibo
      Ok = True
    End If
  '
  ' Despesas Cartorais de Contas
  Case "contas cartorio"
    If (FiltraLancamentos(rstRecibo)) Then
      If (chkRecibos(1).value) Then
        GeraLancamento rstRecibo
      End If
      EstruturaCartorio wrkGeral, rstRecibo
      Ok = True
    End If
  '
  ' Despesas Cartorais de Duplicatas
  Case "duplicatas cartorio"
    If (FiltraDuplicatas(rstRecibo)) Then
      If (chkRecibos(0).value) Then
        GeraLancamento rstRecibo
      End If
      EstruturaCartorio wrkGeral, rstRecibo
      Ok = True
    End If
  '
  End Select
  
  If Ok Then
    
    RodapeReport wrkGeral, rstRecibo
    
    SimpleMsgBar LoadResString(160)
  
    wrkGeral.BeginPrint gTipoDB
    wrkGeral.EndPrint
    
  End If
  
FiltroRecibo_Erro:

  FechaRecordset rstRecibo
  Set wrkGeral = Nothing
  MsgBar LoadResString(IDS_PRONTO)
  SetPtr vbDefault
  
End Sub

' FUNCTION..: FiltraLancamentos
' Objetivo..: Cria a instrução SELECT quando a origem é a tabela de Lançamentos
' Argumento.: [rstDestino]: Recordset que receberá o registro.
' Retorna...: True se abrir a tabela e encontrar algum registro, False se não.
' --------------------------------------------------------------------------------
Private Function FiltraLancamentos(rstDestino As Object) As Boolean
Dim strLancto As String

  SimpleMsgBar LoadResString(13) & LoadResString(14)
  '
  ' Iniciando a instrução
  '
  strLancto = "SELECT Código, Descrição, Vencimento, Pagamento, Empresa , Banco, " & _
              "[Valor Original], Tipo, ([Valor Original] + Acréscimo - Abatimento) " & _
              "As Soma FROM Lançamentos WHERE "
  '
  ' Verificando se o usuário indicou o código do Lançamento
  '
  If IsValid(txtRecibos(0).Text) Then
    Concat strLancto, "Código = ", txtRecibos(0).Text
  Else
    MsgBox LoadResString(183), vbInformation, MsgBoxCaption
    Exit Function
  End If
  '
  ' Conta paga ou recebida
  '
  Concat strLancto, " AND PagRec = ", IIf(cboRecibos(1).ListIndex, "'R';", "'P';")
  '
  ' Se puder abrir o Recordset...
  '
  If (AbreRecordset(rstDestino, strLancto, dbOpenSnapshot) = WL_ERRO) Then
    Exit Function
  ElseIf (UltimoRetorno = WL_NORECORD) Then
    MsgBox LoadResString(IDS_FILTRONORETURN), vbInformation, MsgBoxCaption
    Exit Function
  End If
  
  FiltraLancamentos = True
  
End Function

' FUNCTION..: FiltraDuplicatas
' Objetivo..: Cria a instrução Select para encontrar os dados da Nota indicada pelo
'             usuário.
' Argumento.: [rstDuplicata]: Recordset que conterá os dados.
' Retorna...: Se encontrar o registro especificado retorna True, caso contrário
'             a função retorna False.
' --------------------------------------------------------------------------------
Private Function FiltraDuplicatas(rstDuplicata As Object) As Boolean
Dim strDuplicata As String

  SimpleMsgBar LoadResString(13) & LoadResString(14)
  '
  If Len(txtRecibos(4).Text) > 0 And Len(lblDescRecibos(1).Caption) = 0 Then
    MsgBox "Empresa não cadastrada", vbCritical, "Recibos"
    FiltraDuplicatas = False
    Exit Function
  End If
  
  ' Iniciando a instrução
  '
  strDuplicata = "SELECT Nota, Parcela, Tipo, Descrição, Vencimento, Pagamento, " & _
                 "Empresa, Banco, Conta, Situação, [Valor Original], " & _
                 "([Valor Original] + Acréscimo - Abatimento) AS Soma " & _
                 "FROM Duplicatas WHERE PagRec = '" & mstrPagRec & "' AND "
  ' Verificando se o usuário indicou o número da nota
  '
  If (IsValid(txtRecibos(2).Text)) Then
    Concat strDuplicata, "Nota = ", txtRecibos(2).Text
  Else
    MsgBox LoadResString(183), vbInformation, MsgBoxCaption
    Exit Function
  End If
  
  'Verificando se o usuário digitou o nome da Empresa
  If (IsValid(txtRecibos(4).Text)) Then
    Concat strDuplicata, " AND Empresa = ", Quote(txtRecibos(4).Text, "'")
  End If
  
  'Protocolo 72684: Verificando se a parcela foi informada
  If (IsValid(txtRecibos(3).Text)) Then
    Concat strDuplicata, " AND Parcela = ", Trim$(txtRecibos(3).Text)
  End If
  
  '
  ' Abrindo o Recordset
  '
  If (AbreRecordset(rstDuplicata, strDuplicata, dbOpenSnapshot) = WL_ERRO) Then
    Exit Function
  ElseIf (UltimoRetorno = WL_NORECORD) Then
    MsgBox LoadResString(IDS_FILTRONORETURN), vbInformation, MsgBoxCaption
    FechaRecordset rstDuplicata
    Exit Function
  End If
  
  FiltraDuplicatas = True
  
End Function

' SUB.......: TituloReport
' Objetivo..: Cria os grupos que serão usados como título do relatório.
' Argumentos: [strTitulo]: Título do relatório.
'             [wrkDest]  : KeybReport da impressão.
'             [rstEmp]   : Recordset com os dados da empresa.
' --------------------------------------------------------------------------------
Private Sub TituloReport(strTitulo As String, wrkDest As KeybReport, rstEmp As Object)
Dim strTexto As String          'Concate os textos de alguns campos

  With wrkDest
    .FontSize = 12
    .AddPageHeader
    .PageHeader.AddSecao scHeader, 8
    With .PageHeader.Header.Linha(1)
      .AddCampo , wrCSSimpleLine
      .Campo(1).BorderStyle = wrDash
    End With

    With .PageHeader.Header.Linha(2)
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Razão", NUL)
      .Campo(1).FontStyle = wrFSBold
    End With

    With .PageHeader.Header.Linha(3)
      strTexto = GetValue(rstEmp, "Endereço", NUL)
      Concat strTexto, " - ", GetValue(rstEmp, "CEP", NUL)
      .AddCampo , wrCSFixedText, strTexto
    End With

    With .PageHeader.Header.Linha(4)
      strTexto = GetValue(rstEmp, "Bairro", NUL)
      Concat strTexto, " - ", GetValue(rstEmp, "Cidade", NUL)
      Concat strTexto, " - ", GetValue(rstEmp, "Estado", NUL)
      .AddCampo , wrCSFixedText, strTexto
    End With

    With .PageHeader.Header.Linha(5)
      strTexto = "Fone: " & GetValue(rstEmp, "Fone1", NUL)
      Concat strTexto, " - ", GetValue(rstEmp, "Fone2", NUL)
      Concat strTexto, "  -  ", "Fax: ", GetValue(rstEmp, "Fax", NUL)
      .AddCampo , wrCSFixedText, strTexto
    End With

    With .PageHeader.Header.Linha(6)
      strTexto = "Inscrição no CNPJ nº.: "
      AppendStr strTexto, GetValue(rstEmp, "CNPJ/CPF", NUL)
      .AddCampo , wrCSFixedText, strTexto
    End With

    With .PageHeader.Header.Linha(7)
      strTexto = "Inscrição Estadual nº.: "
      AppendStr strTexto, GetValue(rstEmp, "IEst/RG", NUL)
      .AddCampo , wrCSFixedText, strTexto
    End With

    With .PageHeader.Header.Linha(8)
      .AddCampo , wrCSSimpleLine
      .Campo(1).BorderStyle = wrDash
    End With
    '
    ' Seção de rodapé do cabeçalho com o título do recibo
    '
    .PageHeader.AddSecao scFooter, 4
    
    With .PageHeader.Footer.Linha(2)
      .DrawBorder = wrDBTopBorder
      .BorderStyle = wrDash
      .DrawWidth = 2
    End With

    With .PageHeader.Footer.Linha(3)
      .AddCampo , wrCSFixedText, strTitulo, wrTACentro
      .Campo(1).FontStyle = wrFSBold
      .Campo(1).FontSize = 16
    End With

    With .PageHeader.Footer.Linha(4)
      .DrawBorder = wrDBBottomBorder
      .BorderStyle = wrDash
      .DrawWidth = 2
    End With

  End With
  
End Sub

' SUB.......: RodapeReport
' Objetivo..: Cria os grupos que serão usados como título do relatório.
' Argumentos: [wrkDest]  : KeybReport da impressão.
'             [rstEmp]   : Recordset com os dados da empresa.
' --------------------------------------------------------------------------------
Private Sub RodapeReport(wrkDest As KeybReport, rstEmp As Object)
  '
  ' Adicionando o grupo de rodapé da página
  '
  With wrkDest
  
    .Grupo(1).AddSecao scFooter, 5
    With .Grupo(1).Footer
      .Linha(1).AddCampo , wrCSFixedText, Format$(Date, "Long Date"), wrTADireito
      .Linha(4).AddCampo , wrCSSimpleLine, , , , (wrkDest.ClientWidth / 2)
      .Linha(5).AddCampo , wrCSFixedText, GetValue(rstEmp, "Razão", NUL), wrTACentro, , (wrkDest.ClientWidth / 2)
    
    
      '
      ' Se o recibo for para conta paga devo trocar o nome do assinante do recibo
      '
      If ((tabRecibos.SelectedItem.Key = "contas") And (cboRecibos(1).ListIndex = 0)) Then
        .Linha(5).Campo(1).Text = txtRecibos(17).Text
      End If
      
    End With
  End With
  
End Sub

' SUB.......: EstruturaRecibo
' Objetivo..: Estrutura o relatório para Recibos diversos e Nota de Débito.
' Argumento.: [grkRecibo]: KeybRelatorio para Recibo ou Nota de Débito.
'             [rstSrc]   : Recordset que contém os dados da conta.
' -------------------------------------------------------------------------------
Private Sub EstruturaRecibo(grkRecibo As KeybReport, rstSrc As Object)
Dim strHeader As String         'Para seleção dos dados do cabeçalho
Dim rstHeader As Object      'Abre os dados do cabeçalho.
Dim rstEmp    As Object      'Abre os dados da empresa para o recibo.
Dim strEmp    As String         'Usada na instrução select da empresa para o recibo
Dim strTitulo As String         'Título do relatório
Dim curValor  As Currency       'Valor do recibo
Dim strProv   As String         'Proveniente
Dim strDesc   As String         'Descrição do Lançamento/Duplicata
Dim strVenc   As String         'Data de vencimento
Dim strPgto   As String         'Data de Pagamento
  '
  ' Construi duas instruções Select. Uma para o cabeçalho da página e outra
  ' para os dados que serão apresentados no corpo do recibo.
  '
  strHeader = "SELECT Empresa As Razão, Endereço, Bairro, CEP, Cidade, " & _
              "Estado, Fone1, Fone2, Fax, CNPJ AS [CNPJ/CPF], " & _
              "[Inscrição Estadual] AS [IEst/RG] FROM KinSys;"
              
  strEmp = "SELECT Razão, Endereço, Bairro, Cidade, Estado, CEP, Cidade, " & _
           "Estado, Fone1, Fone2, Fax, [CNPJ/CPF], [IEst/RG] FROM Empresas WHERE Apel = '" & _
           GetValue(rstSrc, "Empresa", NUL) & "';"
  '
  ' Verifica qual o tipo do recibo a ser enviado.
  '
  If ((tabRecibos.SelectedItem.Key = "contas") And (cboRecibos(1).ListIndex = 0)) Or _
    (((tabRecibos.SelectedItem.Key = "duplicatas")) And (mstrPagRec = "P")) Then
    ' Se for recibo de contas pagas
    AbreRecordset rstHeader, strEmp, dbOpenSnapshot
    AbreRecordset rstEmp, strHeader, dbOpenSnapshot
    strTitulo = "Recibo de " & KeybUCase(tabRecibos.SelectedItem.Key, PorFrase) & " Pagas"
  Else
    'Em qualquer outro caso de recibo
    AbreRecordset rstHeader, strHeader, dbOpenSnapshot
    AbreRecordset rstEmp, strEmp, dbOpenSnapshot
    If ((tabRecibos.SelectedItem.Key = "contas") And (cboRecibos(0).ListIndex = 1)) Then
      strTitulo = "Nota de Débito"
    Else
      strTitulo = "Recibo"
    End If
  End If
  '
  ' Cria o grupo de cabeçalho dos recibos
  '
  TituloReport strTitulo, grkRecibo, rstHeader
  '
  ' Configurando o corpo do relatório
  '
  With grkRecibo
    .FontStyle = wrFSNormal
    .FontSize = 12
    .AddGrupo "1"
    .Grupo(1).AddSecao scHeader, 17
    With .Grupo(1).Header.Linha(3)
      .AddCampo , wrCSFixedText, "Empresa:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Razão", NUL)
      .Campo(2).FontStyle = wrFSBold
    End With
    
    With .Grupo(1).Header.Linha(4)
      .AddCampo , wrCSFixedText, "Endereço:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Endereço", NUL), , 100
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "CEP", NUL)
    End With
    
    With .Grupo(1).Header.Linha(5)
      .AddCampo , wrCSFixedText, "Bairro:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Bairro", NUL), , 50
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Cidade", NUL), , 50
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Estado", NUL)
    End With
    ' Verificando se os campos de contato e departamento estão preenchidos
    ' dependendo do tab que está visível
    '
    If (tabRecibos.SelectedItem.Key = "contas") Then
      If Len(txtRecibos(9).Text) Then
        With .Grupo(1).Header.Linha(6)
          .AddCampo , wrCSFixedText, "Contato:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, txtRecibos(9).Text
        End With
      End If
      If Len(txtRecibos(8).Text) Then
        With .Grupo(1).Header.Linha(7)
          .AddCampo , wrCSFixedText, "Departamento:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, txtRecibos(8).Text
        End With
      End If
    ElseIf (tabRecibos.SelectedItem.Key = "duplicatas") Then
      If Len(txtRecibos(5).Text) Then
        With .Grupo(1).Header.Linha(6)
          .AddCampo , wrCSFixedText, "Contato:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, txtRecibos(5).Text
        End With
      End If
      If Len(txtRecibos(6).Text) Then
        With .Grupo(1).Header.Linha(7)
          .AddCampo , wrCSFixedText, "Departamento:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, txtRecibos(6).Text
        End With
      End If
    End If
    '
    ' Obtém o valor dependendo do tipo de recibo
    '
    If ((tabRecibos.SelectedItem.Key = "contas") And (cboRecibos(1).ListIndex = 0)) Then
      curValor = CMoeda(txtRecibos(16).Text)
    Else
      curValor = GetValue(rstSrc, "Soma", 0)
    End If
    
    With .Grupo(1).Header.Linha(9)
      .AddCampo , wrCSFixedText, "A importância de:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, Format$(curValor, "Currency")
    End With
    
    With .Grupo(1).Header.Linha(10)
      .AddCampo , wrCSFixedText, "Extenso:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, KeybUCase(KeybExtenso(curValor), PorPalavra)
      .Campo(2).MultiLine = True
    End With
    
    If ((tabRecibos.SelectedItem.Key = "contas") And (cboRecibos(1).ListIndex = 0)) Then
      'Contas pagas
      strProv = txtRecibos(14).Text
      strDesc = NUL
      strVenc = txtRecibos(15).Text
      strTitulo = txtRecibos(1).Text      'strTitulo utilizado para Observação
    ElseIf (tabRecibos.SelectedItem.Key = "contas") Then
      'Contas recebidas
      strProv = "Lançamento " & StrZero(GetValue(rstSrc, "Código", 0), 6)
      strDesc = GetValue(rstSrc, "Descrição", NUL)
      strVenc = GetValue(rstSrc, "Vencimento", NUL)
      strTitulo = txtRecibos(1).Text      'strTitulo utilizado para Observação
    ElseIf (tabRecibos.SelectedItem.Key = "duplicatas") Then
      'Duplicatas recebidas
      strProv = "Nota Fiscal " & StrZero(GetValue(rstSrc, "Nota", 0), 6) & IIf(txtRecibos(3).Text <> "", " / Parcela nº " & txtRecibos(3).Text, "")
      strDesc = GetValue(rstSrc, "Descrição", NUL)
      strVenc = GetValue(rstSrc, "Vencimento", NUL)
      strPgto = GetValue(rstSrc, "Pagamento", NUL)
      strTitulo = txtRecibos(7).Text      'strTitulo utilizado para Observação
    End If
    
    With .Grupo(1).Header.Linha(13)
      .AddCampo , wrCSFixedText, "Proveniente do(a):", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, strProv
    End With
    
    With .Grupo(1).Header.Linha(14)
      .AddCampo , wrCSFixedText, "Vencimento:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, strVenc
    End With
    
    With .Grupo(1).Header.Linha(15)
      .AddCampo , wrCSFixedText, "Pagamento:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, strPgto
    End With
    
    If (Len(strDesc)) Then
      With .Grupo(1).Header.Linha(16)
        .AddCampo , wrCSFixedText, "Descrição:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, strDesc
      End With
    End If
    
    If (Len(strTitulo)) Then
      With .Grupo(1).Header.Linha(17)
        .AddCampo , wrCSFixedText, "Observação:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, strTitulo
        .Campo(2).MultiLine = True
      End With
    End If
  
  End With
  ' Fechando os dois recordsets que não são mais necessários aqui
  '
  FechaRecordset rstEmp
  FechaRecordset rstHeader
    
End Sub

' SUB.......: EstruturaCartorio
' Objetivo..: Cria a estrutura do recibo de pagamentos via cartório para duplicatas
'             e contas.
' Argumentos: [grkCart]: Variável do Gerador de relatórios.
'             [rstSrc] : Recordset com os dados de origem.
' --------------------------------------------------------------------------------
Private Sub EstruturaCartorio(grkCart As KeybReport, rstSrc As Object)
Dim rstEmp     As Object               'Recordset com os dados da empresa
Dim rstDonaSis As Object               'Recordset com os dados da empresa usuária
Dim strSelect  As String                  'Instrução Select para as tabelas
Dim blnDupl    As Boolean                 'Configura se é duplicata ou lançamento
Dim curCart    As Currency                'Total de custas do cartório

  blnDupl = (InStr(1, tabRecibos.SelectedItem.Key, "duplicatas") > 0)
  '
  ' Seleciona os dados da empresa usuária do sistema para o cabeçalho do
  ' relatório
  '
  strSelect = "SELECT Empresa As Razão, Endereço, Bairro, CEP, Cidade, " & _
              "Estado, Fone1, Fone2, Fax, CNPJ AS [CNPJ/CPF], " & _
              "[Inscrição Estadual] AS [IEst/RG] FROM KinSys;"
              
  AbreRecordset rstDonaSis, strSelect, dbOpenSnapshot
  '
  ' Seleciona os dados da empresa que está no registro da conta
  '
  strSelect = "SELECT Razão, Endereço, Bairro, Cidade, Estado, CEP, Fone1, " & _
              "Fone2, Fax, [CNPJ/CPF], [IEst/RG], Contato, Dpto FROM " & _
              "Empresas WHERE Apel = '" & GetValue(rstSrc, "Empresa", NUL) & "';"
              
  AbreRecordset rstEmp, strSelect, dbOpenSnapshot
  '
  ' Título do Recibo
  '
  strSelect = "Cobrança de Despesas de Cartório de "
  AppendStr strSelect, IIf(blnDupl, "Duplicatas", "Contas")
  
  TituloReport strSelect, grkCart, rstDonaSis
  '
  ' Configurando o corpo do relatório
  '
  With grkCart
    .FontStyle = wrFSNormal
    .FontSize = 12
    .AddGrupo "1"
    .Grupo(1).AddSecao scHeader, 18
    With .Grupo(1).Header
      With .Linha(2)
        .AddCampo , wrCSFixedText, "Empresa:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, GetValue(rstEmp, "Razão", NUL)
        .Campo(2).FontStyle = wrFSBold
      End With
      
      With .Linha(3)
        .AddCampo , wrCSFixedText, "Aos cuidados de:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, GetValue(rstEmp, "Contato", NUL)
      End With
      
      With .Linha(4)
        .AddCampo , wrCSFixedText, "Fax:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, GetValue(rstEmp, "Fax", NUL)
      End With
      
      With .Linha(5)
        .AddCampo , wrCSFixedText, "Contato/Depto:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, txtRecibos(IIf(blnDupl, 5, 9)).Text, , 50
        .AddCampo , wrCSFixedText, txtRecibos(IIf(blnDupl, 6, 8)).Text
      End With
      '
      ' Fecha o recordset da empresa para liberar memória. Esta variável será
      ' utilizada posteriormente para obter os dados do Banco
      '
      FechaRecordset rstEmp
      
      With .Linha(6)
        .AddCampo , wrCSFixedText, "Documento:", , 40
        .Campo(1).FontStyle = wrFSBold
        
        strSelect = IIf(blnDupl, "Nota Fiscal ", "Lançamento ")
        AppendStr strSelect, StrZero(GetValue(rstSrc, IIf(blnDupl, "Nota", "Código"), 0), 6)
        If blnDupl Then
          Concat strSelect, " - Parcela ", StrZero(GetValue(rstSrc, "Parcela", 1), 2)
        End If
        
        .AddCampo , wrCSFixedText, strSelect
      End With
      
      With .Linha(7)
        .AddCampo , wrCSFixedText, "Vencimento:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, GetValue(rstSrc, "Vencimento", NUL)
      End With
      
      With .Linha(8)
        .AddCampo , wrCSFixedText, "Valor:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, Format$(GetValue(rstSrc, "Valor Original", 0), "Currency")
      End With
      
      With .Linha(9)
        .AddCampo , wrCSFixedText, "Pagamento:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, GetValue(rstSrc, "Pagamento", NUL)
      End With
      
      With .Linha(10)
        .AddCampo , wrCSFixedText, "Valor Pago:", , 40
        .Campo(1).FontStyle = wrFSBold
        If IsNull(rstSrc("Pagamento").value) Then
          .AddCampo , wrCSFixedText, txtRecibos(10).Text
        Else
          .AddCampo , wrCSFixedText, Format$(GetValue(rstSrc, "Soma", 0), "Currency"), , 30
        End If
      End With
      
      .Linha(11).AddCampo , wrCSFixedText, "(  ) Cartório    (  ) Depósito Bancário", , , 90
      
      .Linha(13).AddCampo , wrCSFixedText, LoadResString(184)
      
      With .Linha(14)
        .AddCampo , wrCSFixedText, "Taxa de envio a Cartório:", , 55
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, _
                    Format$(CMoeda(txtRecibos(11).Text), FMOEDA), wrTADireito, 30
      End With
      
      With .Linha(15)
        .AddCampo , wrCSFixedText, "Despesas Cartorais:", , 55
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, Format$(CMoeda(txtRecibos(12).Text), FMOEDA), _
                    wrTADireito, 30
      End With
      
      With .Linha(16)
        .AddCampo , wrCSFixedText, "Juros:", , 55
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, Format$(CMoeda(txtRecibos(13).Text), FMOEDA), _
                    wrTADireito, 30
      End With
      
      With .Linha(17)
        .AddCampo , wrCSFixedText, "Total a ser depositado:", , 55
        .Campo(1).FontStyle = wrFSBold
        
        curCart = CMoeda(txtRecibos(11).Text) + CMoeda(txtRecibos(12).Text) + CMoeda(txtRecibos(13).Text)
        
        .AddCampo , wrCSFixedText, Format$(curCart, FMOEDA), wrTADireito, 30
        .Campo(2).FontStyle = wrFSBold
      End With
      '
      ' Obtendo os dados do Banco da empresa usuária do Sistema
      '
      strSelect = "SELECT Nome, Agência, Conta FROM Bancos WHERE Banco = " & _
                  GetValue(rstSrc, "Banco", 0) & ";"
      If (AbreRecordset(rstEmp, strSelect, dbOpenSnapshot) = WL_OK) Then
        .AddLinha
        With .Linha(19)
          .AddCampo , wrCSFixedText, LoadResString(185)
          .Campo(1).FontStyle = wrFSBold
        End With
        
        .AddLinha
        With .Linha(20)
          .AddCampo , wrCSFixedText, "Banco:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, GetValue(rstEmp, "Nome", NUL)
        End With
        
        .AddLinha
        With .Linha(21)
          .AddCampo , wrCSFixedText, "Agência:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, GetValue(rstEmp, "Agência", NUL)
        End With
        
        .AddLinha
        With .Linha(22)
          .AddCampo , wrCSFixedText, "Conta:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, GetValue(rstEmp, "Conta", NUL)
        End With
      End If
      
      .AddLinha "Obs" 'Linha 19 se não houver Banco, 23 se houver
      With .Linha("Obs")
        .AddCampo , wrCSFixedText, "Observação:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, txtRecibos(IIf(blnDupl, 7, 1)).Text
        .Campo(2).MultiLine = True
      End With
    End With
  End With
  '
  ' Fecha as variáveis Recordset não mais necessárias
  '
  FechaRecordset rstEmp
  FechaRecordset rstDonaSis
  
End Sub

' SUB.......: GeraLancamento
' Objetivo..: Quanto o usuário está tirando um relatório de Cobrança de Despesas
'             cartorais de Duplicatas o sistema pode gerar um lançamento com o
'             valor recebido. Esta função gera este lançamento.
' Argumento.: [rstDuplicata]: Recordset com os dados da duplicata.
' ---------------------------------------------------------------------------------
Private Sub GeraLancamento(rstDuplicata As Object)
Dim rstNovoLancto As Object      'Para o novo Lançamento gerado
Dim datVencto     As Date           'Data de Vencimento
Dim strTemp       As String         'Para retorno das InputBox
  '
  ' Obtendo a data de vencimento do usuário.
  '
  Do While (Len(strTemp) = 0)
    strTemp = InputBox("Informe a data de Vencimento:", "Vencimento", NUL)
    If (Len(strTemp) = 0) Then
      If MsgBox(LoadResString(186), vbQuestion Or vbYesNo Or vbDefaultButton2, _
                MsgBoxCaption) = vbYes Then Exit Sub
    Else
      If (Not EData(strTemp)) Then
        MsgBox ResolveResString(IDS_INVALIDDATEVALUE, resUM, strTemp), _
               vbInformation, MsgBoxCaption
        strTemp = NUL
      End If
    End If
  Loop
  
  datVencto = CDate(strTemp)
  '
  ' Obtendo a data de Pagamento. Esta data não é obrigatória. Se o usuário cancelar
  ' posso sair do loop.
  '
  Do
    strTemp = InputBox("Informe a data de Pagamento:", "Pagamento", NUL)
    If (Len(strTemp) = 0) Then Exit Do
    If (Not EData(strTemp)) Then
      MsgBox ResolveResString(IDS_INVALIDDATEVALUE, resUM, strTemp), _
             vbInformation, MsgBoxCaption
      strTemp = NUL
    End If
  Loop While (Len(strTemp) = 0)
  
  SimpleMsgBar LoadResString(187)
  
  If (AbreRecordset(rstNovoLancto, "Lançamentos", dbOpenDynaset) = WL_ERRO) Then
    MsgBox ResolveResString(IDS_ERROPENTABLE, resUM, "Lançamentos"), _
           vbInformation, MsgBoxCaption
    Exit Sub
  Else
    Dim lngCodLancto As Long
    
    lngCodLancto = CLng(ProximoNumero("Código", "Lançamentos", "PagRec = 'R'"))
    
    rstNovoLancto.AddNew
    rstNovoLancto("PagRec").value = "R"
    rstNovoLancto("Código").value = lngCodLancto
    rstNovoLancto("Empresa").value = rstDuplicata("Empresa").value
    rstNovoLancto("Tipo").value = rstDuplicata("Tipo").value
    '
    ' Verificando se estou gerando o Lançamento de Duplicata ou de Lançamento
    '
    If (Trim$(ExtractStr(rstDuplicata.name, "FROM", "WHERE")) = "Duplicatas") Then
      rstNovoLancto("Descrição").value = ResolveResString(188, resUM, _
                                         StrZero(rstDuplicata("Nota").value, 6), _
                                         resDOIS, GetValue(rstDuplicata, "Tipo"), _
                                         "|3", StrZero(rstDuplicata("Parcela").value, 2))
    Else
      rstNovoLancto("Descrição").value = ResolveResString(195, resUM, _
                                         StrZero(rstDuplicata("Código").value, 6), _
                                         resDOIS, GetValue(rstDuplicata, "Tipo"))
    End If
                                         
    rstNovoLancto("Emissão").value = Date
    rstNovoLancto("Vencimento").value = datVencto
    If Len(strTemp) Then
      rstNovoLancto("Pagamento").value = CDate(strTemp)
    End If
    rstNovoLancto("Liberação").value = datVencto
    rstNovoLancto("Valor Original").value = CMoeda(txtRecibos(11).Text) + _
                                            CMoeda(txtRecibos(12).Text) + _
                                            CMoeda(txtRecibos(13).Text)
    rstNovoLancto("Acréscimo").value = 0
    rstNovoLancto("Abatimento").value = 0
    rstNovoLancto("Banco").value = GetValue(rstDuplicata, "Banco", 1)
    rstNovoLancto("Conta").value = GetValue(rstDuplicata, "Conta", 1)
    rstNovoLancto("Centro").value = 0
    rstNovoLancto("Cheque").value = 0
    rstNovoLancto("Moeda").value = NUL
    rstNovoLancto("Valor da Moeda").value = 0
    rstNovoLancto("Controle").value = NUL
    rstNovoLancto("Marcação").value = False
    rstNovoLancto("Obs").value = NUL
    
    rstNovoLancto.update
    
    MsgBox ResolveResString(189, resUM, StrZero(lngCodLancto, 6)), vbInformation, MsgBoxCaption
    
  End If
  
  FechaRecordset rstNovoLancto
  
End Sub

' SUB.......: EstruturaContas
' Objetivo..: Estrutura o relatório para Nota de Débito e Recibo.
' Argumentos: [grkRecibo]: KeybRelatorio para Nota de Débito.
'             [strEmit  ]: String com o nome do Emitente
' -------------------------------------------------------------------------------
Private Sub EstruturaContas(grkNotaDebito As KeybReport, TipoRelatorio As String)
Dim strHeader As String         ' Para seleção dos dados do cabeçalho
Dim rstHeader As Object      ' Abre os dados do cabeçalho.
Dim rstEmp    As Object      ' Abre os dados da empresa para o Nota de Débito.
Dim strEmp    As String         ' Usada na instrução select da empresa para o Nota de Débito
Dim strTitulo As String         ' Título do relatório
Dim curValor  As Currency       ' Valor Total da Nota de Débito
Dim strProv   As String         ' Proveniente
Dim strDesc   As String         ' Descrição do Lançamento/Duplicata
Dim strVenc   As String         ' Data de vencimento
Dim strPgto   As String         ' Data de vencimento
Dim strParc   As String         ' Parcela
Dim strVal    As String         ' Valor do Lançamento
Dim strCtrl   As String         ' Controle do Lançamento
Dim i         As Integer        ' Contador da ListView
Dim strEmit   As String         ' Nome do Emitente
  '
  ' Construi duas instruções Select. Uma para o cabeçalho da página e outra
  ' para os dados que serão apresentados no corpo da Nota de Débito.
  '
  strHeader = "SELECT Empresa As Razão, Endereço, Bairro, CEP, Cidade, " & _
              "Estado, Fone1, Fone2, Fax, CNPJ AS [CNPJ/CPF], " & _
              "[Inscrição Estadual] AS [IEst/RG] FROM KinSys;"
              
  strEmp = "SELECT Razão, Endereço, Bairro, Cidade, Estado, CEP, Fone1, " & _
           "Fone2, Fax, [CNPJ/CPF], [IEst/RG] FROM Empresas WHERE Apel = '" & _
            lvwLancamentos.ListItems(1).SubItems(4) & "'"
  '
  ' Verifica qual o tipo da Nota de Débito a ser enviado.
  '
  AbreRecordset rstHeader, strHeader, dbOpenSnapshot
  AbreRecordset rstEmp, strEmp, dbOpenSnapshot
  
  If TipoRelatorio = "Nota de Débito" Then
    strEmit = InputBox$("Insira o nome do Emitente da Nota de Débito", "Nota de Débito")
  End If
  
  strTitulo = TipoRelatorio
  '
  ' Cria o grupo de cabeçalho dos recibos
  '
  TituloReport strTitulo, grkNotaDebito, rstHeader
  '
  ' Configurando o corpo do relatório
  '
  With grkNotaDebito
    .FontStyle = wrFSNormal
    .FontSize = 12
    .AddGrupo "1"
    .Grupo(1).AddSecao scHeader, 12
    With .Grupo(1).Header.Linha(3)
      .AddCampo , wrCSFixedText, "Empresa:", , 40
      .Campo(1).FontStyle = wrFSBold
      If IsValid(GetValue(rstEmp, "Razão", NUL)) Then
        .AddCampo , wrCSFixedText, GetValue(rstEmp, "Razão", NUL)
      Else
        .AddCampo , wrCSFixedText, lvwLancamentos.ListItems(1).SubItems(4)
      End If
      .Campo(2).FontStyle = wrFSBold
    End With
    
    With .Grupo(1).Header.Linha(4)
      .AddCampo , wrCSFixedText, "Endereço:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Endereço", NUL), , 100
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "CEP", NUL)
    End With
    
    With .Grupo(1).Header.Linha(5)
      .AddCampo , wrCSFixedText, "Bairro:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Bairro", NUL), , 50
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Cidade", NUL), , 50
      .AddCampo , wrCSFixedText, GetValue(rstEmp, "Estado", NUL)
    End With
    '
    ' Verificando se os campos de contato e departamento estão preenchidos
    ' dependendo do tab que está visível
    '
    If (tabRecibos.SelectedItem.Key = "contas") Then
      If Len(txtRecibos(9).Text) Then
        With .Grupo(1).Header.Linha(6)
          .AddCampo , wrCSFixedText, "Contato:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, txtRecibos(9).Text
        End With
      End If
      If Len(txtRecibos(8).Text) Then
        With .Grupo(1).Header.Linha(7)
          .AddCampo , wrCSFixedText, "Departamento:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, txtRecibos(8).Text
        End With
      End If
    ElseIf (tabRecibos.SelectedItem.Key = "duplicatas") Then
      If Len(txtRecibos(5).Text) Then
        With .Grupo(1).Header.Linha(6)
          .AddCampo , wrCSFixedText, "Contato:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, txtRecibos(5).Text
        End With
      End If
      If Len(txtRecibos(6).Text) Then
        With .Grupo(1).Header.Linha(7)
          .AddCampo , wrCSFixedText, "Departamento:", , 40
          .Campo(1).FontStyle = wrFSBold
          .AddCampo , wrCSFixedText, txtRecibos(6).Text
        End With
      End If
    End If
    '
    ' Obtém o valor
    '
    For i = 1 To lvwLancamentos.ListItems.Count
      curValor = curValor + CCurDef(lvwLancamentos.ListItems(i).SubItems(5))
    Next i
    
    With .Grupo(1).Header.Linha(9)
      .AddCampo , wrCSFixedText, "A importância de:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, Format$(curValor, "Currency")
    End With
    
    With .Grupo(1).Header.Linha(10)
      .AddCampo , wrCSFixedText, "Extenso:", , 40
      .Campo(1).FontStyle = wrFSBold
      .AddCampo , wrCSFixedText, KeybUCase(KeybExtenso(curValor), PorPalavra)
      .Campo(2).MultiLine = True
    End With
    
    .FontSize = 9
    .FontStyle = wrFSBold
    With .Grupo(1).Header.Linha(12)
    'Vinicius Elyseu(25/05/2016) - Demanda: #120791
      .AddCampo , wrCSFixedText, "Lançamento", , 25
      .AddCampo , wrCSFixedText, "Parcela", , 15
      .AddCampo , wrCSFixedText, "Descrição", , 77
      .AddCampo , wrCSFixedText, "Vencimento", wrTACentro, 20
      .AddCampo , wrCSFixedText, "Pagamento", wrTACentro, 20
      .AddCampo , wrCSFixedText, "Controle", , 20
      .AddCampo , wrCSFixedText, "Valor", wrTADireito, 20
      .BorderStyle = wrSolid
      .DrawBorder = wrDBBottomBorder
    End With

    For i = 1 To lvwLancamentos.ListItems.Count
      With lvwLancamentos
        strProv = StrZero(lvwLancamentos.ListItems(i), 6)
        strDesc = lvwLancamentos.ListItems(i).SubItems(2)
        strVenc = lvwLancamentos.ListItems(i).SubItems(3)
        strVal = lvwLancamentos.ListItems(i).SubItems(5)
        strCtrl = lvwLancamentos.ListItems(i).SubItems(6)
        strPgto = lvwLancamentos.ListItems(i).SubItems(7)
        strParc = StrZero(lvwLancamentos.ListItems(i).SubItems(8), 3)
      End With
      
      .FontStyle = wrFSNormal
      .Grupo(1).Header.AddLinha
      With .Grupo(1).Header.Linha(.Grupo(1).Header.LinhasCount)
      'Vinicius Elyseu(25/05/2016) - Demanda: #120791
        .AddCampo , wrCSFixedText, strProv, , 25
        .AddCampo , wrCSFixedText, strParc, , 15
        .AddCampo , wrCSFixedText, strDesc, , 77
        .AddCampo , wrCSFixedText, strVenc, wrTACentro, 20
        .AddCampo , wrCSFixedText, strPgto, wrTACentro, 20
        .AddCampo , wrCSFixedText, strCtrl, , 20
        .AddCampo , wrCSFixedText, Format$(strVal, FMOEDA), wrTADireito, 20
      End With
    Next i
    
    .FontSize = 12
    
    If Len(strEmit) > 0 Then
      .Grupo(1).Header.AddLinha
      .Grupo(1).Header.AddLinha
      With .Grupo(1).Header.Linha(.Grupo(1).Header.LinhasCount)
        .AddCampo , wrCSFixedText, "Emitente:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, strEmit
      End With
    End If
            
    strTitulo = txtRecibos(1).Text      'strTitulo utilizado para Observação
    If (Len(strTitulo)) Then
      .Grupo(1).Header.AddLinha
      .Grupo(1).Header.AddLinha
      With .Grupo(1).Header.Linha(.Grupo(1).Header.LinhasCount)
        .AddCampo , wrCSFixedText, "Observação:", , 40
        .Campo(1).FontStyle = wrFSBold
        .AddCampo , wrCSFixedText, strTitulo
        .Campo(2).MultiLine = True
      End With
    End If
    .Grupo(1).Header.AddLinha
  End With
  '
  ' Fechando os dois recordsets que não são mais necessários aqui
  '
  FechaRecordset rstEmp
  FechaRecordset rstHeader
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
