VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmConsultaAtrasos 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Títulos em Atraso"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11715
   Icon            =   "frmConsultaAtrasos.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11715
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtBanco 
      DataField       =   "CodDef"
      Height          =   315
      Left            =   1020
      TabIndex        =   20
      ToolTipText     =   "Código do Banco"
      Top             =   60
      Width           =   1815
   End
   Begin VB.Frame fraEmpresa 
      Caption         =   "Empresa Devedora"
      Height          =   525
      Left            =   30
      TabIndex        =   10
      Top             =   6030
      Width           =   11640
      Begin VB.Label lblCalcTotFinal 
         Alignment       =   1  'Right Justify
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   9480
         TabIndex        =   19
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label lblTotal 
         Caption         =   "Total:"
         Height          =   285
         Left            =   8880
         TabIndex        =   18
         Top             =   180
         Width           =   825
      End
      Begin VB.Label lblDescEmp 
         Caption         =   "Apresentar o nome da empresa deverdora do título selecionado na grid"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   180
         Width           =   8340
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1545
      Left            =   10140
      TabIndex        =   9
      Top             =   0
      Width           =   1515
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   150
         TabIndex        =   7
         Top             =   1050
         Width           =   1215
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   150
         TabIndex        =   6
         Top             =   630
         Width           =   1215
      End
   End
   Begin VB.TextBox txtEmpresa 
      DataField       =   "CodDef"
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      ToolTipText     =   "Código da empresa"
      Top             =   1140
      Width           =   1815
   End
   Begin VB.TextBox txtMora 
      Height          =   315
      Left            =   4530
      TabIndex        =   4
      ToolTipText     =   "Utilizado quando o valor da mora diária não estiver informado no título."
      Top             =   810
      Width           =   1200
   End
   Begin VB.TextBox txtDataBasePagto 
      Height          =   315
      Left            =   4530
      TabIndex        =   3
      ToolTipText     =   "Data base para cálculo do atraso."
      Top             =   420
      Width           =   1200
   End
   Begin VB.ComboBox cboOrigem 
      Height          =   315
      ItemData        =   "frmConsultaAtrasos.frx":0C42
      Left            =   1020
      List            =   "frmConsultaAtrasos.frx":0C4F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   1815
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "frmConsultaAtrasos.frx":0C73
      Left            =   1020
      List            =   "frmConsultaAtrasos.frx":0C7D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   780
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4365
      Left            =   30
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1590
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   7699
      _Version        =   393216
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      Height          =   225
      Left            =   5790
      TabIndex        =   23
      Top             =   870
      Width           =   165
   End
   Begin VB.Label lblBanco 
      Caption         =   "#"
      Height          =   255
      Left            =   2970
      TabIndex        =   22
      Top             =   90
      Width           =   6555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Banco:"
      Height          =   255
      Left            =   60
      TabIndex        =   21
      Top             =   90
      Width           =   900
   End
   Begin VB.Label lblEmpresa 
      Caption         =   "#"
      Height          =   255
      Left            =   2970
      TabIndex        =   17
      Top             =   1200
      Width           =   6555
   End
   Begin VB.Label LabLote 
      Alignment       =   1  'Right Justify
      Caption         =   "Empresa:"
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   1170
      Width           =   900
   End
   Begin VB.Label LabDatEntLot 
      Alignment       =   1  'Right Justify
      Caption         =   "Mora Diária Padrão:"
      Height          =   255
      Left            =   2970
      TabIndex        =   15
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Dt.Base para Pagto.:"
      Height          =   255
      Left            =   2970
      TabIndex        =   14
      Top             =   450
      Width           =   1485
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   810
      Width           =   900
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Origem:"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   450
      Width           =   900
   End
End
Attribute VB_Name = "frmConsultaAtrasos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_APEL = 3
Private mColunasGrid As String
Private mCodApelDesc As String
Private mModal As Boolean

Private Property Get ColunasGrid() As String

    If mColunasGrid = Empty Then
        mColunasGrid = "campo=SeqGrid;label=;tamanho=600" & _
                                "|campo=Banco;Label=Banco;tamanho=600;" & _
                                "|campo=Origem;label=Origem;tamanho=600" & _
                                "|campo=Empresa;Label=Empresa;tamanho=1100" & _
                                "|campo=Titulo;label=Tit.nº;tamanho=800" & _
                                "|campo=Parcela;Label=Parcela;tamanho=700" & _
                                "|campo=Valor Original;Label=Valor Original;tamanho=1100;tipo=tpColGridDouble; formato=###,###,##0.00" & _
                                "|campo=Vencimento;Label=Vencimento;tamanho=1000" & _
                                "|campo=Dias;Label=Dias Atraso;tamanho=900;tipo=tpColGridInteger;formato=###,###,##0" & _
                                "|campo=VlrMRD;Label= Vlr.Mr.Diária;tamanho=1000;tipo=tpColGridDouble;formato=###,###,##0.00" & _
                                "|campo=VlrMoraCalc;Label=Mora Calc.;tamanho=900;tipo=tpColGridDouble;formato=###,###,##0.00" & _
                                "|campo=VlrMul;Label=Multa;tamanho=900;tipo=tpColGridDouble;formato=###,###,##0.00" & _
                                "|campo=VlrFinal;Label=Valor Final;tamanho=1100;tipo=tpColGridDouble;formato=###,###,##0.00"
                      
                                
    End If

    ColunasGrid = mColunasGrid
    
End Property

Private Sub LimpaCampos()
    txtEmpresa.Text = Empty
    lblEmpresa.Caption = Empty
    txtMora.Text = Empty
    cboOrigem.Text = "Duplicatas"
    cboTipo.Text = "A Receber"
    GridVazia
    lblBanco.Caption = Empty
End Sub

Private Sub cboTipo_LostFocus()
  Select Case cboTipo.Text
     Case "A Receber"
       fraEmpresa.Caption = "Empresa Devedora"
     Case "A Pagar"
       fraEmpresa.Caption = "Empresa Credora"
  End Select
End Sub

Private Sub cmdCancelar_Click()
    LimpaCampos
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
txtDataBasePagto.Text = CStr(Date)
LimpaCampos

PosForm Me, bRedimensiona:=False

End Sub

Private Sub GridVazia()

lblDescEmp.Caption = ""
lblCalcTotFinal.Caption = "0,00"
Call CarregaHFlexGrid(MSHFlexGrid1, Nothing, ColunasGrid)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SavePosForm Me
    Set frmConsultaAtrasos = Nothing
End Sub

Private Sub MostraDescEmpresa()
    'se trocou o produto
    If StrComp(mCodApelDesc, MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, COL_APEL)) <> 0 Then
        ModControlesAux.GetAssocValueEmpresa MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, COL_APEL), lblDescEmp
        mCodApelDesc = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, COL_APEL)
    End If
End Sub

Private Sub MSHFlexGrid1_Click()
    MostraDescEmpresa
End Sub

Private Sub MSHFlexGrid1_RowColChange()
    MostraDescEmpresa
End Sub

Private Sub cmdConsultar_Click()
    If IsValid(txtBanco.Text) And lblBanco.Caption = "" Then
        MsgBox "Banco informado nao esta cadastrado!", vbInformation
        Exit Sub
    End If
    Consultar
End Sub

Private Sub txtBanco_Change()
     GetAssocValue "Select nome from bancos where banco = " & txtBanco.Text, lblBanco
End Sub

Private Sub txtBanco_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 And KeyCode = vbKeyPageDown Then
     PCampo "Bancos", "SELECT * FROM Bancos", pbCampo, txtBanco, "Banco"
  End If
End Sub

Private Sub txtBanco_KeyPress(KeyAscii As Integer)
    DValor KeyAscii
End Sub

Private Sub txtBanco_LostFocus()
  Dim dMora As Single
  
  'pt 77609
  dMora = GetFieldValue("Mora", "Bancos", IIf(txtBanco.Text = NUL, "banco = 0", "banco = " & txtBanco.Text), , 0)
  If dMora > 0 Then
    txtMora.Text = Format(CStr(dMora / 30), F4CASAS)
  End If
    
End Sub

Private Sub txtEmpresa_Change()
   ModControlesAux.GetAssocValueEmpresa txtEmpresa.Text, lblEmpresa
End Sub

Private Sub txtEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 0 And KeyCode = vbKeyPageDown Then
     PCampo "Empresas", "SELECT * FROM Empresas", pbCampo, txtEmpresa, "Apel"
  End If
End Sub

Private Sub txtDataBasePagto_KeyPress(KeyAscii As Integer)
   SetMascara KeyAscii, txtDataBasePagto.SelStart, MASK_DATA
End Sub

Private Sub txtDataBasePagto_LostFocus()
   If (txtDataBasePagto.Text <> "") Then
      If Not (IsDate(txtDataBasePagto.Text)) Then
         MsgBox ("Data fora do padrão DD/MM/AAAA")
         txtDataBasePagto.SetFocus
      End If
   End If
End Sub

Private Sub txtMora_KeyPress(KeyAscii As Integer)
    DValor KeyAscii
   ' DMoeda KeyAscii
End Sub

Private Sub Consultar()
    Dim oTitulos As CTitulos
    Dim rs As Object
    Dim tpOrigem As OrigemTitulo
    Dim tpTitulo As TipoTituloAtrasado

On Error GoTo Error_Handler

    'limpo a grid
    GridVazia
    SetPtr vbHourglass
    cmdCancelar.Enabled = False
    cmdConsultar.Enabled = False
    DoEvents
    Select Case cboOrigem
        Case "Lançamentos"
            tpOrigem = ori_Lancamentos
        Case "Duplicatas"
            tpOrigem = ori_Duplicatas
        Case "Ambos"
            tpOrigem = ori_Ambos
    End Select
    
    Select Case cboTipo
        Case "A Pagar"
            tpTitulo = tip_A_Pagar_Vencido
        Case "A Receber"
            tpTitulo = tip_A_Receber_Vencido
        Case "Todos"
            tpTitulo = tip_Todos_Vencidos
    End Select
      
    Set oTitulos = New CTitulos
    Set rs = oTitulos.ConsultaTitulosAtrasados(tpOrigem, tpTitulo, IIf(txtEmpresa.Text <> NUL, txtEmpresa.Text, ""), IIf(txtDataBasePagto.Text <> NUL, CDateDef(txtDataBasePagto.Text), Date), IIf(txtMora.Text <> NUL, CSngDef(txtMora.Text), 0), IIf(txtBanco.Text <> NUL, CLngDef(txtBanco.Text), 0))
    Call CarregaHFlexGrid(MSHFlexGrid1, rs, mColunasGrid)
    lblCalcTotFinal.Caption = Format(oTitulos.TotFinal, "###,###,##0.00")
    Set rs = Nothing
    Set oTitulos = Nothing
    cmdCancelar.Enabled = True
    cmdConsultar.Enabled = True
    cmdConsultar.SetFocus
    SetPtr vbDefault
    Exit Sub
    
Error_Handler:
    cmdCancelar.Enabled = True
    cmdConsultar.Enabled = True
    cmdConsultar.SetFocus
    SetPtr vbDefault
    FinallyConnection Aplicacao
    TrataErroEx
End Sub

Private Sub Form_Deactivate()
If mModal Then Me.SetFocus
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
