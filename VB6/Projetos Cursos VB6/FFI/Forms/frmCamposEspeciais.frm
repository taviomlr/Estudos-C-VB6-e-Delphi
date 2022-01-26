VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmCamposEspeciais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Campos Especiais"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9795
   LinkTopic       =   "frmCamposEspeciais"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   9795
   Begin VB.Frame Frame 
      Height          =   5055
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   8325
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   1170
         MaxLength       =   250
         TabIndex        =   0
         Top             =   360
         Width           =   6885
      End
      Begin VB.TextBox txtValor 
         Height          =   285
         Left            =   1170
         MaxLength       =   200
         TabIndex        =   1
         Top             =   720
         Width           =   6885
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgCamposEspeciais 
         Height          =   3615
         Left            =   60
         TabIndex        =   8
         Top             =   1380
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6376
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblDescricao 
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   11
         Top             =   735
         Width           =   525
      End
   End
   Begin VB.Frame fraBotoes 
      Height          =   5055
      Left            =   8400
      TabIndex        =   9
      Top             =   0
      Width           =   1350
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   3
         Top             =   570
         Width           =   1185
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   4
         Top             =   960
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   7
         Top             =   2130
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   1350
         Width           =   1185
      End
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   6
         Top             =   1740
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmCamposEspeciais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private registroSelecionado As voCamposEspeciais
Private Const strCpEspeciais = "campo=#vazio;label=;tamanho=150|" & _
                               "campo=Descricao;label=Descrição;tamanho=4500|" & _
                               "campo=Valor;label=Nome;tamanho=3000"
                               
Private isAlterando As Boolean
                               
Private Sub CarregarGrid()
    Dim CampoEspecial       As voCamposEspeciais
    Dim CamposEspeciais     As colCamposEspeciais
    Dim dao                 As CamposEspeciaisDAO
    
    Set dao = New CamposEspeciaisDAO
    Call dao.Initialize
    
    Set registroAntes = New voCamposEspeciais
    
    Set CamposEspeciais = dao.Carregar
    
    Call dgCamposEspeciais.Clear
    If CamposEspeciais Is Nothing Then
        Call CarregaHFlexGrid(dgCamposEspeciais, Nothing, strCpEspeciais)
    Else
        If CamposEspeciais.Count = 0 Then
            Call CarregaHFlexGrid(dgCamposEspeciais, Nothing, strCpEspeciais)
        Else
            CamposEspeciais.MoveFirst
            Call CarregaHFlexGrid(dgCamposEspeciais, , strCpEspeciais, , , CamposEspeciais)
        End If
    End If
    
    Call dao.Terminate
    
    dgCamposEspeciais.Row = 1
    Set CamposEspeciais = Nothing
    Set dao = Nothing
End Sub


Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdCancelar_Click()
    If MsgBox("Você deseja realmente cancelar este registro?", vbYesNo, "Atenção") = vbYes Then
        Call CarregarGrid
        Call cmdNovo_Click
    End If
End Sub

Private Sub cmdExcluir_Click()
    Dim dao As CamposEspeciaisDAO
    Dim Registro As voCamposEspeciais
    Dim objCpEspecial As clsCarteiraCpEspecialDAO
    
    If IsNothing(objCpEspecial) Then
       Set objCpEspecial = New clsCarteiraCpEspecialDAO
    End If
    
    Aplicacao.Connect
    Call objCpEspecial.init(Aplicacao)
    
    If Not objCpEspecial.ExisteVinculacaoCampo(EnterpriseID, CdEstabelecimento, txtValor.Text) Then
        If MsgBox("Você deseja realmente excluir este registro?", vbYesNo, "Atenção") = vbYes Then
        
            Set Registro = New voCamposEspeciais
            Registro.Descricao = txtDescricao.Text
            Registro.Valor = txtValor.Text
        
            Set dao = New CamposEspeciaisDAO
            dao.Initialize
        
            If dao.Excluir(Registro) Then
                MsgBox "Registro excluído com sucesso! ", vbInformation
            Else
                MsgBox "Falha ao excluir o registro! ", vbCritical
            End If
            
            Call CarregarGrid
            Set Registro = Nothing
        
            dao.Terminate
            Set dao = Nothing
        
            Call cmdNovo_Click
        End If
    Else
        MsgBox "Não foi possível excluir o campo '" & txtDescricao.Text & "' pois o mesmo possui vinculação em uma ou mais carteiras.", vbInformation, NomeModulo
    End If
    Aplicacao.Disconnect
End Sub

Private Sub cmdGravar_Click()
    Dim dao As CamposEspeciaisDAO
    Dim Registro As voCamposEspeciais
    
    If txtDescricao.Text = "" Then
        MsgBox "Favor preencher o campo descrição corretamente!", vbInformation, NomeModulo
        Exit Sub
    End If
    If txtValor.Text = "" Then
        MsgBox "Favor preencher o campo nome corretamente!", vbInformation, NomeModulo
        Exit Sub
    End If
    
    Set dao = New CamposEspeciaisDAO
    dao.Initialize
    
    If Not dao.ExisteRegistroDuplicado(txtDescricao.Text, txtValor.Text) Then
        If MsgBox("Você deseja gravar este registro?", vbYesNo, NomeModulo) = vbYes Then
            Set Registro = New voCamposEspeciais
            Registro.Descricao = txtDescricao.Text
            Registro.Valor = txtValor.Text
        
            If isAlterando And dao.ExisteRegistro(registroSelecionado) Then
                Call dao.Excluir(registroSelecionado)
            End If
            If dao.Gravar(Registro) Then
                MsgBox "Registro gravado com sucesso! ", vbInformation
                Call CarregarGrid
            Else
                MsgBox "Falha ao gravar o registro! ", vbCritical
            End If
        
            Call cmdNovo_Click
            Set Registro = Nothing
        End If
    Else
        MsgBox "Não foi possível incluir o campo especial pois já existe um registro equivalente!", vbInformation, NomeModulo
    End If
    Call dao.Terminate
    Set dao = Nothing
End Sub

Private Sub cmdNovo_Click()
    txtDescricao.Text = ""
    txtValor.Text = ""
    isAlterando = False
    dgCamposEspeciais.Row = 1
    
    Call txtDescricao.SetFocus
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub dgCamposEspeciais_Click()
     With dgCamposEspeciais
       If .Row > 0 Then
          If .TextMatrix(.Row, 1) <> "" Then
            txtDescricao.Text = .TextMatrix(.Row, 1)
            txtValor.Text = .TextMatrix(.Row, 2)
                            
            registroSelecionado.Descricao = .TextMatrix(.Row, 1)
            registroSelecionado.Valor = .TextMatrix(.Row, 2)
            isAlterando = True
          End If
       End If
    End With
End Sub

Private Sub Form_Load()
    isAlterando = False
    Set registroSelecionado = New voCamposEspeciais
    Call CarregarGrid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set registroSelecionado = Nothing
End Sub
