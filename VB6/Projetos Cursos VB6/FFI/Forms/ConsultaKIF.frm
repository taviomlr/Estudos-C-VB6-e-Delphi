VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form fdConsultasKIF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "ConsultaKIF.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraConsultas 
      Caption         =   "Saldos Bancários"
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
      Index           =   3
      Left            =   240
      TabIndex        =   44
      Top             =   360
      Width           =   5775
      Begin VB.CheckBox chkImprimir 
         Caption         =   "Imprimir Relatório?"
         Height          =   255
         Left            =   3960
         TabIndex        =   51
         Top             =   2040
         Width           =   1695
      End
      Begin ComctlLib.ListView lvwConsultas 
         Height          =   1695
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         _Version        =   327682
         SmallIcons      =   "imgConsultas"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Banco"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Nome"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Conta"
            Object.Width           =   1693
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Nome da Conta"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   19
         Left            =   600
         MaxLength       =   10
         TabIndex        =   47
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
         Height          =   195
         Index           =   26
         Left            =   120
         TabIndex        =   46
         Top             =   2040
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdConsultas 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   49
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsultas 
      Caption         =   "E&xibir..."
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   48
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame fraConsultas 
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
      TabIndex        =   22
      Top             =   360
      Width           =   5775
      Begin VB.ComboBox cboConsultas 
         Height          =   315
         Index           =   8
         ItemData        =   "ConsultaKIF.frx":0C42
         Left            =   3600
         List            =   "ConsultaKIF.frx":0C5E
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2040
         Width           =   2055
      End
      Begin VB.ComboBox cboConsultas 
         Height          =   315
         Index           =   3
         ItemData        =   "ConsultaKIF.frx":0CB4
         Left            =   1080
         List            =   "ConsultaKIF.frx":0CCD
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   13
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cboConsultas 
         Height          =   315
         Index           =   2
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   12
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   30
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   11
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   35
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   10
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   33
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   9
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   39
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   8
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   37
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   7
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   41
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Ordem:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   27
         Left            =   3000
         TabIndex        =   42
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Consultar:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   15
         Left            =   2520
         TabIndex        =   27
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblDesc 
         Caption         =   "#"
         Height          =   195
         Index           =   12
         Left            =   2640
         TabIndex        =   31
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   2985
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Emissão de:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "até"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   12
         Left            =   2400
         TabIndex        =   34
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Vencto. de:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "até"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   10
         Left            =   2400
         TabIndex        =   38
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Controle:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   40
         Top             =   2040
         Width           =   630
      End
   End
   Begin VB.Frame fraConsultas 
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
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin VB.ComboBox cboConsultas 
         Height          =   315
         Index           =   7
         ItemData        =   "ConsultaKIF.frx":0D26
         Left            =   3600
         List            =   "ConsultaKIF.frx":0D42
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   6
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   5
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   4
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   3
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   2
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   1
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cboConsultas 
         Height          =   315
         Index           =   1
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtConsultas 
         Height          =   315
         Index           =   0
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cboConsultas 
         Height          =   315
         Index           =   0
         ItemData        =   "ConsultaKIF.frx":0D96
         Left            =   1080
         List            =   "ConsultaKIF.frx":0DAF
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Ordem:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   25
         Left            =   3000
         TabIndex        =   20
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Controle:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "até"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   7
         Left            =   2400
         TabIndex        =   16
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Vencto. de:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "até"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Emissão de:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblDesc 
         Caption         =   "#"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   2985
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   5
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Nota:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   390
      End
      Begin VB.Label lblConsultas 
         AutoSize        =   -1  'True
         Caption         =   "Consultar:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   705
      End
   End
   Begin ComctlLib.TabStrip tabConsultas 
      Height          =   3135
      Left            =   120
      TabIndex        =   50
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5530
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Duplicatas"
            Key             =   "Duplicatas"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Lançamentos"
            Key             =   "Lançamentos"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Saldos"
            Key             =   "Saldos"
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
   Begin ComctlLib.ImageList imgConsultas 
      Left            =   240
      Top             =   3000
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
            Picture         =   "ConsultaKIF.frx":0E08
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fdConsultasKIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboConsultas_Click(Index As Integer)
  
  Select Case Index
  '
  ' Campo Consulta de Duplicatas
  Case 0
    If CompStr(cboConsultas(0).Text, "Todos") Then
      fraConsultas(0).Caption = "Todas as Duplicatas"
    Else
      fraConsultas(0).Caption = "Duplicatas " & cboConsultas(0).Text
    End If
  '
  ' Campo Consulta de Lançamentos
  Case 3
    If CompStr(cboConsultas(3).Text, "Todos") Then
      fraConsultas(1).Caption = "Todos os Lançamentos"
    Else
      fraConsultas(1).Caption = "Lançamentos " & cboConsultas(1).Text
    End If
  '
  End Select
    
End Sub

Private Sub cboConsultas_GotFocus(Index As Integer)
  '
  ' Mensagens da barra de status
  '
  Select Case Index
  '
  ' Consultas duplicatas
  Case 0
    MsgBar "Define como consultar as duplicatas"
  '
  ' Tipo da duplicata
  Case 1
    MsgBar "Tipo da duplicata"
  '
  ' Consultas lançamentos
  Case 3
    MsgBar "Define como consultar os lançamentos"
  '
  ' Tipo do Lançamento
  Case 2
    MsgBar "Tipo do lançamento"
  '
  End Select
  
End Sub

Private Sub cmdConsultas_Click(Index As Integer)
  Select Case Index
  '
  ' Botão Exibir
  Case 0
    cmdConsultas(0).Enabled = False
    cmdConsultas(1).Enabled = False
    'Select Case tabConsultas.SelectedItem.Key
    '
    'Case "Duplicatas"
      'ConsultaDuplicatas
    '
    'Case "Lançamentos"
      'ConsultaLanctos
    '
    'Case "Saldos"
      ConsultaSaldos
    '
    'End Select
    cmdConsultas(0).Enabled = True
    cmdConsultas(1).Enabled = True
  '
  ' Botão Fechar
  Case 1
    Unload Me
    
  
  
  End Select
End Sub

Private Sub Form_Load()
Dim strProc As String

  CenterForm Me
  '
  ' Carregando os tipos de duplicatas e lançamentos nas caixas combinadas0
  ' Duplicatas (Valor padrão: "Todos")
  '
  cboConsultas(0).ListIndex = (cboConsultas(0).ListCount - 1)
  '
  ' Tipo de duplicatas (Valor Padrão: "Todos")
  '
  strProc = "SELECT * FROM Opções WHERE Rotina = 'Dupl. a Pagar';"
  ComboAddItem cboConsultas(1), strProc, "Texto"
  cboConsultas(1).AddItem "Todos"
  cboConsultas(1).ListIndex = (cboConsultas(1).ListCount - 1)
  '
  ' Tipo de Lançamentos (Valor padrão: "Todos")
  '
  strProc = "SELECT * FROM Opções WHERE Rotina = 'Lanctos a Pagar';"
  ComboAddItem cboConsultas(2), strProc, "Texto"
  cboConsultas(2).AddItem "Todos"
  cboConsultas(2).ListIndex = (cboConsultas(2).ListCount - 1)
  '
  ' Lançamentos (Valor padrão: "Todos")
  '
  cboConsultas(3).ListIndex = (cboConsultas(3).ListCount - 1)
  '
  ' Ordem dos dados (Duplicatas e Lançamentos)
  '
  cboConsultas(7).ListIndex = (cboConsultas(7).ListCount - 1)
  cboConsultas(8).ListIndex = (cboConsultas(8).ListCount - 1)
  '
  ' Bancos disponíveis no Sistema
  'OBS .: como a select muda d+, achei melhor criar as duas versões separadas
  If gTipoDB = Access Then
    strProc = "SELECT Format(Banco, ""000000"") AS Bco, Nome, Conta, [Nome Conta] FROM Bancos;"
  Else
    strProc = "SELECT replicate('0',6-len(BANCO))+convert(char(6),BANCO) AS Bco, Nome, Conta, [Nome Conta] FROM Bancos;"
  End If
  
  If ListViewAddItem(lvwConsultas, strProc, 1) Then
    lvwConsultas.ListItems(1).Selected = True
  End If
  '
  ' Limpando os Labels de descrição
  '
  lblDesc(1).Caption = NUL
  lblDesc(12).Caption = NUL
  '
  ' Sugerindo a última data em todos os campos de data
  '
  txtConsultas(2).Text = LerArquivoASCII("Consultas", "txtConsultas(2)", "Fox.ini")
  txtConsultas(3).Text = LerArquivoASCII("Consultas", "txtConsultas(3)", "Fox.ini")
  txtConsultas(4).Text = LerArquivoASCII("Consultas", "txtConsultas(4)", "Fox.ini")
  txtConsultas(5).Text = LerArquivoASCII("Consultas", "txtConsultas(5)", "Fox.ini")
  txtConsultas(8).Text = LerArquivoASCII("Consultas", "txtConsultas(8)", "Fox.ini")
  txtConsultas(9).Text = LerArquivoASCII("Consultas", "txtConsultas(9)", "Fox.ini")
  txtConsultas(10).Text = LerArquivoASCII("Consultas", "txtConsultas(10)", "Fox.ini")
  txtConsultas(11).Text = LerArquivoASCII("Consultas", "txtConsultas(11)", "Fox.ini")
  txtConsultas(19).Text = LerArquivoASCII("Consultas", "txtConsultas(19)", "Fox.ini")
  '
  ' Exibindo o primeiro controle tab
  '
  'tabConsultas.Tabs(1).Selected = True
   tabConsultas.Tabs.Remove (1)
   tabConsultas.Tabs.Remove (2)
   tabConsultas.Tabs(1).Caption = "Saldos"

  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not cmdConsultas(0).Enabled Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

  GravarArquivoASCII "Consultas", "txtConsultas(2)", txtConsultas(2).Text, "Fox.ini"
  GravarArquivoASCII "Consultas", "txtConsultas(3)", txtConsultas(3).Text, "Fox.ini"
  GravarArquivoASCII "Consultas", "txtConsultas(4)", txtConsultas(4).Text, "Fox.ini"
  GravarArquivoASCII "Consultas", "txtConsultas(5)", txtConsultas(5).Text, "Fox.ini"
  GravarArquivoASCII "Consultas", "txtConsultas(8)", txtConsultas(8).Text, "Fox.ini"
  GravarArquivoASCII "Consultas", "txtConsultas(9)", txtConsultas(9).Text, "Fox.ini"
  GravarArquivoASCII "Consultas", "txtConsultas(10)", txtConsultas(10).Text, "Fox.ini"
  GravarArquivoASCII "Consultas", "txtConsultas(11)", txtConsultas(11).Text, "Fox.ini"
  GravarArquivoASCII "Consultas", "txtConsultas(19)", txtConsultas(19).Text, "Fox.ini"

  Set fdConsultasKIF = Nothing
  MsgBar MsgBoxCaption
End Sub


Private Sub tabConsultas_Click()
  'fraConsultas(0).Visible = (tabConsultas.SelectedItem.Key = "Duplicatas")
  'fraConsultas(1).Visible = (tabConsultas.SelectedItem.Key = "Lançamentos")
  'fraConsultas(3).Visible = (tabConsultas.SelectedItem.Key = "Saldos")
  
End Sub

Private Sub txtConsultas_Change(Index As Integer)
  
  Select Case Index
  '
  ' Campos de empresa
  Case 1, 12
    If Len(txtConsultas(Index).Text) Then
      GetAssocValue "SELECT Razão, Apel FROM Empresas WHERE Apel = '" & _
                    txtConsultas(Index).Text & "';", lblDesc(Index), txtConsultas(Index)
    Else
      lblDesc(Index).Caption = NUL
    End If
    
  End Select
  
End Sub

Private Sub txtConsultas_GotFocus(Index As Integer)
  Selecione txtConsultas(Index)
  '
  ' Mensagens da barra de status do programa
  '
  Select Case Index
  '
  ' Campo Nota
  Case 0
    MsgBar "Número da nota" & ResolveResString(75, resUM, "Duplicatas")
  '
  ' Campo Empresa em duplicatas, Empresa em Lançamentos e Empresa em Empresas
  Case 1, 12
    MsgBar "Nome fantasia da empresa"
  '
  ' Campos data inicial da emissão em duplicatas e lançamentos
  Case 2, 10
    MsgBar "Data inicial de Emissão"
  '
  ' Campos data final de emissão em duplicatas e lançamentos
  Case 3, 11
    MsgBar "Data final da Emissão"
  '
  ' Campos data inicial de vencimento em duplicatas e lançamentos
  Case 4, 8
    MsgBar "Data inicial de vencimento"
  '
  ' Campos data final de vencimento em duplicatas e lançamentos
  Case 5, 9
    MsgBar "Data final de vencimento"
  '
  ' Campos Controle em duplicatas e lançamentos
  Case 6, 7
    MsgBar "Código de controle"
  '
  ' Campo Código em Lançamentos
  Case 13
    MsgBar "Código do Lançamento" & ResolveResString(75, resUM, "Lançamentos")
  '
  ' Campo data para o saldo
  Case 19
    MsgBar "Data para o saldo final"
  '
  End Select
  
End Sub

Private Sub txtConsultas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim strSelect As String

  If (Shift = 0) And (KeyCode = vbKeyPageDown) Then
    Select Case Index
    '
    ' Duplicatas.Nota, Lançamentos.Código
    Case 0, 13
      Select Case IIf((Index = 0), cboConsultas(0).ListIndex, cboConsultas(3).ListIndex)
      '
      ' A pagar
      Case 0
        strSelect = "WHERE PagRec = 'P' AND (Pagamento IS NULL);"
      '
      ' Pagas
      Case 1
        strSelect = "WHERE PagRec = 'P' AND Not (Pagamento IS NULL);"
      '
      ' A Receber
      Case 2
        strSelect = "WHERE PagRec = 'R' AND (Pagamento IS NULL);"
      '
      ' Recebidas
      Case 3
        strSelect = "WHERE PagRec = 'R' AND Not (Pagamento IS NULL);"
      '
      ' A Pagar e Pagas
      Case 4
        strSelect = "WHERE PagRec = 'P';"
      '
      ' A Receber e Recebidas
      Case 5
        strSelect = "WHERE PagRec = 'R';"
      '
      ' Todos
      Case 6
        strSelect = NUL
      '
      End Select
      
      If (Index = 0) Then
        PCampo "Duplicatas", "SELECT * FROM Duplicatas " & strSelect, _
               pbCampo, txtConsultas(0), "Nota"
      Else
        PCampo "Lançamentos", "SELECT * FROM Lançamentos " & strSelect, _
               pbCampo, txtConsultas(13), "Código"
      End If
    '
    ' Empresas
    Case 1, 12
      PCampo "Empresas Ativas", "SELECT * FROM Empresas;", _
               pbCampo, txtConsultas(Index), "Apel"
    End Select
  End If
  
End Sub

Private Sub txtConsultas_KeyPress(Index As Integer, KeyAscii As Integer)
  
  Select Case Index
  '
  ' Campos de data
  Case 2 To 5, 8 To 11, 19
    SetMascara KeyAscii, txtConsultas(Index).SelStart, MASK_DATE4
  '
  ' Campos Duplicatas.Nota e Lançamentos.Código
  Case 0, 13
    SetMascara KeyAscii, txtConsultas(Index).SelStart, fMask("Duplicatas", "Nota")
  '
  End Select
  
End Sub

' SUB.......: ConsultaDuplicatas
' Objetivo..: Cria a expressão de consulta para exibir os dados da tabela de duplicatas
' ----------------------------------------------------------------------------------------
Private Sub ConsultaDuplicatas()
Dim strFinal    As String         'Termina a expressão
Dim datInicial  As Date           'Para as expressões com data
Dim datFinal    As Date

  SimpleMsgBar LoadResString(13) & LoadResString(14)
  '
  ' Criando a expressão de Consulta
  '
  SetPtr vbHourglass
  
  If gTipoDB = Access Then
  
    strFinal = "SELECT Nota, Parcela, Tipo, Empresa, Descrição, Emissão, Vencimento, " & _
             "Pagamento, Liberação, " & _
             "([Valor Original] + Acréscimo - Abatimento) " & _
             "AS [Valor Total], DateDiff(""d"", Vencimento, IIf(Pagamento IS NULL, Now(), Pagamento)) " & _
             "AS [Dias de Atraso], " & _
             "[Valor Original], Acréscimo, Abatimento, PerMul, VlrMul, VlrMrd, VlrDsp FROM Duplicatas"
  Else
  
    strFinal = "SELECT Nota, Parcela, Tipo, Empresa, Descrição, Emissão, Vencimento, " & _
             "Pagamento, Liberação, " & _
             "([Valor Original] + Acréscimo - Abatimento) " & _
             "AS [Valor Total], DateDiff(d, Vencimento, CASE WHEN(Pagamento IS Null) THEN getdate() ELSE Pagamento END) " & _
             "AS [Dias de Atraso], " & _
             "[Valor Original], Acréscimo, Abatimento, PerMul, VlrMul, VlrMrd, VlrDsp  FROM Duplicatas"
  
  End If
  '
  ' Campo Nota
  If IsValid(txtConsultas(0).Text) Then
    Concat strFinal, " WHERE Nota = ", txtConsultas(0).Text
  Else
    Concat strFinal, " WHERE Nota > 0"
  End If
  
  Select Case cboConsultas(0).ListIndex
  '
  ' A Pagar
  Case 0
    Concat strFinal, " AND PagRec = 'P' AND Pagamento IS NULL"
  '
  ' Pagas
  Case 1
    Concat strFinal, " AND PagRec = 'P' AND Pagamento IS NOT NULL"
  '
  ' A Receber
  Case 2
    Concat strFinal, " AND PagRec = 'R' AND Pagamento IS NULL"
  '
  ' Recebidas
  Case 3
        Concat strFinal, " AND PagRec = 'R' AND Pagamento IS NOT NULL"
  '
  ' A Pagar e Pagas
  Case 4
    Concat strFinal, " AND PagRec = 'P'"
  '
  ' A Receber e Recebidas
  Case 5
    Concat strFinal, " AND PagRec = 'R'"
  '
  End Select
  
  If Len(txtConsultas(1).Text) Then       'Campo Empresa
    Concat strFinal, " AND Empresa = '", txtConsultas(1).Text, "'"
  End If
  '
  ' Confere as datas de emissão
  '
  If CompDatas(txtConsultas(2), txtConsultas(3), datInicial, datFinal) Then
    Concat strFinal, " AND Emissão BETWEEN ", InverteData(datInicial, True), " AND ", _
           InverteData(datFinal, True)
  End If
  '
  ' Confere as datas de Vencimento
  '
  If CompDatas(txtConsultas(4), txtConsultas(5), datInicial, datFinal) Then
    Concat strFinal, " AND Vencimento BETWEEN ", InverteData(datInicial, True), _
           " AND ", InverteData(datFinal, True)
  End If
  '
  ' Campo Controle
  '
  If Len(txtConsultas(6).Text) Then
    Concat strFinal, " AND Controle = '", txtConsultas(6).Text, "'"
  End If
  '
  ' Campo Tipo
  '
  If Not CompStr(cboConsultas(1).Text, "Todos") Then
    Concat strFinal, " AND Tipo = '", cboConsultas(1).Text, "'"
  End If
 
  '
  ' Ordem final da consulta
  '
  If Not CompStr(cboConsultas(7).Text, "Sem Ordem") Then
    Concat strFinal, " ORDER BY [", cboConsultas(7).Text, "];"
  Else
    AppendStr strFinal, ";"
  End If
  
  PCampo "Consulta Duplicatas", strFinal, 0, Nothing, 0, "Valor Total;Currency"
  SetPtr vbDefault
  MsgBar NUL
  
End Sub

' SUB.......: ConsultaLanctos
' Objetivo..: Exibe a janela de consulta para os lançamentos solicitados pelo usuário
' ------------------------------------------------------------------------------------
Private Sub ConsultaLanctos()
Dim strLanctos As String
Dim datInit    As Date
Dim datFim     As Date
  
  SimpleMsgBar LoadResString(13) & LoadResString(14)
  SetPtr vbHourglass
  
  ' Devido as diferenças, há duas versões desta select
  If gTipoDB = Access Then
  
        strLanctos = "SELECT Código, Tipo, Empresa, Descrição, Emissão, Vencimento, Pagamento, " & _
                 "Liberação, " & _
                 "([Valor Original] + Acréscimo - Abatimento) AS Total, " & _
                 "DateDiff(""d"", Vencimento, IIf(Pagamento IS NULL, Now(), Pagamento)) " & _
                 "AS [Dias de Atraso], [Valor Original], " & _
                 "Acréscimo, Abatimento, PerMul, VlrMul, VlrMrd, VlrDsp FROM Lançamentos WHERE"
  Else
        strLanctos = "SELECT Código, Tipo, Empresa, Descrição, Emissão, Vencimento, Pagamento, " & _
                 "Liberação, " & _
                 "([Valor Original] + Acréscimo - Abatimento) AS Total, " & _
                 "DateDiff(d, Vencimento, CASE WHEN (Pagamento Is Not Null) THEN getdate() ELSE Pagamento END) " & _
                 "AS [Dias de Atraso], [Valor Original], " & _
                 "Acréscimo, Abatimento, PerMul, VlrMul, VlrMrd, VlrDsp FROM Lançamentos WHERE"
  End If
  '
  ' Código do Lançamento
  If IsValid(txtConsultas(13).Text) Then
    Concat strLanctos, " Código = ", txtConsultas(13).Text
  Else
    Concat strLanctos, " Código > 0"
  End If
  '
  ' Verificando os dados que devem ser retornados
  '
  Select Case cboConsultas(3).ListIndex
  '
  ' A Pagar
  Case 0
    Concat strLanctos, " AND PagRec = 'P' AND Pagamento IS NULL"
  '
  ' Pagos
  Case 1
    Concat strLanctos, " AND PagRec = 'P' AND Pagamento IS NOT NULL"
  '
  ' A Receber
  Case 2
    Concat strLanctos, " AND PagRec = 'R' AND Pagamento IS NULL"
  '
  ' Recebidos
  Case 3
    Concat strLanctos, " AND PagRec = 'R' AND Pagamento IS NOT NULL"
  '
  ' A Pagar e Pagos
  Case 4
    Concat strLanctos, " AND PagRec = 'P'"
  '
  ' A Receber e Recebidos
  Case 5
    Concat strLanctos, " AND PagRec = 'R'"
  '
  End Select
  '
  ' Tipo do Lançamento
  '
  If Not CompStr(cboConsultas(2).Text, "Todos") Then
    Concat strLanctos, " AND Tipo = '", cboConsultas(2).Text, "'"
  End If
  '
  ' Empresa
  '
  If Len(txtConsultas(12).Text) Then
    Concat strLanctos, " AND Empresa = '", txtConsultas(12).Text, "'"
  End If
  '
  ' Datas de Emissão
  '
  If CompDatas(txtConsultas(10), txtConsultas(11), datInit, datFim) Then
    Concat strLanctos, " AND Emissão BETWEEN ", InverteData(datInit, True), " AND ", _
           InverteData(datFim, True)
  End If
  '
  ' Datas de Vecimento
  '
  If CompDatas(txtConsultas(8), txtConsultas(9), datInit, datFim) Then
    Concat strLanctos, " AND Vencimento BETWEEN ", InverteData(datInit, True), _
                       " AND ", InverteData(datFim, True)
  End If
  '
  ' Campo Controle
  '
  If Len(txtConsultas(7).Text) Then
    Concat strLanctos, " AND Controle = '", txtConsultas(7).Text, "'"
  End If
  
  '
  ' Ordenando os dados da pesquisa
  '
  If Not CompStr(cboConsultas(8).Text, "Sem Ordem") Then
    Concat strLanctos, " ORDER BY [", cboConsultas(8).Text, "];"
  Else
    AppendStr strLanctos, ";"
  End If
  
  PCampo "Consultando Lançamentos", strLanctos, 0, Nothing, 0, "Total;Currency"
  SetPtr vbDefault
  MsgBar NUL
  
End Sub

' SUB.......: ConsultaSaldos
' Objetivo..: Cria a consulta de saldos bancários sobre os bancos selecionados
'             pelo usuário. Exibe a janela de pesquisa.
' -------------------------------------------------------------------------------
Private Sub ConsultaSaldos()
Dim strTableTemp As String          'Para a tabela temporária
Dim strSaldo     As String          'Para a string de consulta
Dim rstAux       As Object

  SimpleMsgBar LoadResString(13) & LoadResString(14)
  SetPtr vbHourglass
  '
  ' Cria a tabela auxiliar para exibir a janela de pesquisa
  '
  strTableTemp = CriaTabela()
  Call Sleep(1000)
  If Len(strTableTemp) = 0 Then   'Se estiver vazia nenhum dado foi encontrado
    MsgBox LoadResString(2024), vbInformation, MsgBoxCaption
  Else
    MsgBar Caption
    strSaldo = "SELECT Banco, Nome, Inicial AS [Saldo Inicial], " & _
               "Entradas, Saidas, Aplicações, Resgates, Taxas, Juros, " & _
               "Final AS [Saldo Final] FROM " & strTableTemp & ";"
    
    If chkImprimir.value = vbChecked Then
       AbreRecordset rstAux, "select * from " & strTableTemp, dbOpenSnapshot
       fimpSaldosBancarios.Config rstAux
       FechaRecordset rstAux
    Else
      PCampo "Consulta de Saldos", strSaldo, 0, Nothing, 0, "Saldo Final;Currency"
    End If
  
  End If
  '
  ' Exclui a tabela auxiliar criada
  '
  DeleteAux Nothing, strTableTemp
  MsgBar NUL
  SetPtr vbDefault
  
End Sub


' FUNCTION..: CriaTabela
' Objetivo..: Cria a tabela auxiliar e completa-a com os dados necessários.
' Retorna...: Uma String com o nome da tabela criada.
' -------------------------------------------------------------------------------
Private Function CriaTabela() As String
Dim rstSaldos As Object        'Recordset para a tabela auxiliar
Dim rstTemp   As Object        'Recordset para obter os resultados
Dim intBancos As Integer          'Utilizado no Loop dos bancos
Dim strSaldos As String           'Instruções de seleção
Dim fsTemp(9) As FieldStruct      'Estrutura para montar os campos da tabela aux.
Dim strCodBco As String           'Código do Banco atual.
Dim dtInit    As Date             'Data inicial. Último dia do mês anterior.
Dim dtFim     As Date             'Data final. Data informada pelo usuário.
Dim curSaldo  As Currency         'Calcula o Saldo Final.
Dim curValor  As Currency         'Obtém os valore intermediários
  '
  ' Configurando as datas para seleção
  ' Data final é a data informada pelo usuário
  If Len(txtConsultas(19).Text) Then
    If EData(txtConsultas(19).Text) Then
      dtFim = CDate(txtConsultas(19).Text)
    Else
      MsgBox ResolveResString(26, resUM, txtConsultas(19).Text), vbInformation, MsgBoxCaption
      Exit Function
    End If
  Else
    dtFim = Date
  End If
  '
  ' Data inicial é o primeiro dia do mês da data final
  '
  dtInit = FirstDay(dtFim)
  '
  ' Cria uma tabela temporária para armazenar os dados para a consulta
  '
  AppendVar fsTemp(0), "Banco", dbLong
  AppendVar fsTemp(1), "Nome", dbText, 40
  AppendVar fsTemp(2), "Inicial", dbCurrency
  AppendVar fsTemp(3), "Entradas", dbCurrency
  AppendVar fsTemp(4), "Saidas", dbCurrency
  AppendVar fsTemp(5), "Aplicações", dbCurrency
  AppendVar fsTemp(6), "Resgates", dbCurrency
  AppendVar fsTemp(7), "Taxas", dbCurrency
  AppendVar fsTemp(8), "Juros", dbCurrency
  AppendVar fsTemp(9), "Final", dbCurrency
  
  If CrieAux(rstSaldos, fsTemp) Then
    '
    ' Cria uma instrução Select para cada Banco selecionado e inclui os dados na
    ' tabela auxiliar
    '
    For intBancos = 1 To lvwConsultas.ListItems.Count
      
      If lvwConsultas.ListItems(intBancos).Selected Then
        
        strCodBco = lvwConsultas.ListItems(intBancos).Text
        
        rstSaldos.AddNew
        '
        ' Código e nome do Banco
        '
        strSaldos = "SELECT Banco, Nome FROM Bancos WHERE Banco = " & strCodBco & ";"
        
        If AbreRecordset(rstTemp, strSaldos, dbOpenSnapshot) = WL_OK Then
          rstSaldos("Banco").value = GetValue(rstTemp, "Banco")
          rstSaldos("Nome").value = GetValue(rstTemp, "Nome")
          SimpleMsgBar wsprintf("Calculando Banco %s %s", strCodBco, GetValue(rstTemp, "Nome", NUL))
        End If
        '
        ' Saldo Inicial (Saldo do mês anterior)
        '
        curSaldo = SaldoInicial(CLng(strCodBco), dtInit, sConciliado:="Sim")
        rstSaldos("Inicial").value = curSaldo
        '
        ' Entradas em Lançamentos
        '
        #If FOXSQL Then
            strSaldos = "PagRec = 'R' AND (Liberação BETWEEN '%q' AND '%q') AND Banco = %s" & _
                        " AND Pagamento IS NOT NULL"
        #Else
            strSaldos = "PagRec = 'R' AND (Liberação BETWEEN #%q# AND #%q#) AND Banco = %s" & _
                        " AND Pagamento IS NOT NULL"
        #End If
        wvsprintf strSaldos, strSaldos, dtInit, dtFim, strCodBco
        
        curValor = Soma("([Valor Original] + Acréscimo - Abatimento)", _
                        "Lançamentos", strSaldos)
        ' Entradas em Duplicatas
        '
        curValor = curValor + Soma("([Valor Original] + Acréscimo - Abatimento)", _
                                   "Duplicatas", strSaldos)
        curSaldo = (curSaldo + curValor)
        
        rstSaldos("Entradas").value = curValor
        '
        ' Saidas em Lançamentos
        '
        MidStr strSaldos, "'R'", "'P'"
        curValor = Soma("([Valor Original] + Acréscimo - Abatimento)", _
                        "Lançamentos", strSaldos)
        ' Saídas em Duplicatas
        '
        curValor = curValor + Soma("([Valor Original] + Acréscimo - Abatimento)", _
                                   "Duplicatas", strSaldos)
        curSaldo = (curSaldo - curValor)
        rstSaldos("Saidas").value = curValor
        '
        ' Aplicações/Transferências
        '
        #If FOXSQL Then
            strSaldos = "Destino = %s AND (Data BETWEEN '%q' AND '%q')"
        #Else
            strSaldos = "Destino = %s AND (Data BETWEEN #%q# AND #%q#)"
        #End If
        wvsprintf strSaldos, strSaldos, strCodBco, dtInit, dtFim
        
        curValor = Soma("Valor", "[Transf Bancária]", strSaldos, 0)
        curSaldo = (curSaldo + curValor)
        rstSaldos("Aplicações").value = curValor
        '
        ' Resgates/Transferências
        '
        InsereStr strSaldos, "Origem", DeleteStr(strSaldos, "Destino")
        curValor = Soma("Valor", "[Transf Bancária]", strSaldos, 0)
        curSaldo = (curSaldo - curValor)
        rstSaldos("Resgates").value = curValor
        '
        ' Taxas
        '
        #If FOXSQL Then
            strSaldos = "Banco = %s AND (Data BETWEEN '%q' AND '%q') AND Tipo <> '%s'"
        #Else
            strSaldos = "Banco = %s AND (Data BETWEEN #%q# AND #%q#) AND Tipo <> '%s'"
        #End If
        wvsprintf strSaldos, strSaldos, strCodBco, dtInit, dtFim, GetResOptions(1001, 1)
        
        curValor = Soma("Valor", "Aplicações", strSaldos)
        curSaldo = (curSaldo - curValor)
        rstSaldos("Taxas").value = curValor
        '
        ' Juros
        '
        MidStr strSaldos, "Tipo <> '", "Tipo = '"
        curValor = Soma("Valor", "Aplicações", strSaldos)
        curSaldo = (curSaldo + curValor)
        rstSaldos("Juros").value = curValor
        '
        ' Saldo Final
        '
        rstSaldos("Final").value = curSaldo

        rstSaldos.update
      End If
    Next
    '
    ' Fecha a tabela temporária e retorna o seu nome
    '
    CriaTabela = GetTableSource(rstSaldos, True)
    FechaRecordset rstSaldos
    
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
