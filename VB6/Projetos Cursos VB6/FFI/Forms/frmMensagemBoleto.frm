VERSION 5.00
Begin VB.Form frmMensagemBoleto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validação Título Bancário"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14010
   Icon            =   "frmMensagemBoleto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBotoes 
      Height          =   4485
      Left            =   12600
      TabIndex        =   1
      Top             =   0
      Width           =   1365
      Begin Fox.EBSReport ertRelatorio 
         Height          =   795
         Left            =   240
         TabIndex        =   5
         Top             =   2790
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1402
         NomeRelatorio   =   "FOXFCO00041.ERC"
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   375
         Left            =   90
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   2
         Top             =   195
         Width           =   1185
      End
   End
   Begin VB.Frame fraOutras 
      Height          =   4485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12585
      Begin VB.TextBox emeObs 
         Height          =   4245
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   180
         Width           =   12465
      End
   End
End
Attribute VB_Name = "frmMensagemBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrMensagem                                  As String
Private mstrNossoNumero                                  As String

Private Sub cmdImprimir_Click()
    Call fInformacaoImpressao
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    MostraDados
End Sub

Private Sub MostraDados()
    emeObs.Text = mstrMensagem
End Sub

Public Property Let mensagem(Valor As String)
   mstrMensagem = Valor
End Property

Public Property Let NossoNumero(Valor As String)
   mstrNossoNumero = Valor
End Property

Private Function fInformacaoImpressao()
    Dim strNusNum() As String
    Dim intCont     As Integer
    
    'ertRelatorio.AddParametro "VNOSNUM", mstrNossoNumero
    If Trim(mstrNossoNumero <> "") Then
        Call ExecuteSQL("DELETE FROM FFIRetornoAux")
        strNusNum = Split(mstrNossoNumero, ",")
        For intCont = 0 To UBound(strNusNum)
            Call ExecuteSQL("INSERT INTO FFIRetornoAux VALUES(" & strNusNum(intCont) & ")")
        Next
    End If
    DoEvents
    Call Sleep(1500)
    
    'Projeto - 1139 - Fernando Paludo(19/09/2011)
    ertRelatorio.EnterpriseId = EnterpriseId
    ertRelatorio.UserGroup = GetFieldValue("grupo", "Usuários", "usuário = '" & UserName & "'", , "")
    
    ertRelatorio.NomeRelatorio = "FOXFIN00180"
    
    
    'Ivo Sousa (08/07/2013) - Alteração para buscar o Config do Temp do Usuário do Windows
    If ArquivoExiste(ArquivoConfiguracao) Then
        ertRelatorio.CaminhoConfiguracao = ArquivoConfiguracao
    Else
        ertRelatorio.CaminhoConfiguracao = App.Path & "\..\Configurações\CONFIG.INI"
    End If
    
    'Ivo Sousa(23/12/2008) - Alterado o Caminho do Config
    'ertRelatorio.CaminhoConfiguracao = "C:\Fox\Configurações\CONFIG.INI" 'ArquivoConfiguracao      '"C:\Fox\Configurações\CONFIG.INI"
    
    ertRelatorio.EscModel = emNone
    ertRelatorio.OEMConvert = False
    'ertRelatorio.CaminhoImpressora = mobjDAOImpressao.Impressora
    'ertRelatorio.NumeroCopias = mobjDAOImpressao.Nr_copia
    ertRelatorio.Visualizador = CaminhoPasta(pastaProgramas) & "fre.exe"
    ertRelatorio.ArquivoExecucao = CaminhoPasta(pastaProgramas) & "Bordero.xml"
    ertRelatorio.LoginUsuario = UserName
    ertRelatorio.SenhaUsuario = GetFieldValue("senha", "Usuários", "usuário = '" & UserName & "'", , "")
    ertRelatorio.Visualizar
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
