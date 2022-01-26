VERSION 5.00
Begin VB.Form frmMensagemAtivacaoRotina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validação de Ativação da Rotina"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   7665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPrincipal 
      Height          =   1455
      Left            =   30
      TabIndex        =   0
      Top             =   -60
      Width           =   7635
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   6180
         TabIndex        =   3
         Top             =   870
         Width           =   1215
      End
      Begin VB.CheckBox chkAlerta 
         Caption         =   "Não receber mais este alerta"
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   420
         TabIndex        =   1
         Tag             =   "Agenda"
         Top             =   960
         Width           =   2805
      End
      Begin VB.Label lblMsg 
         AutoSize        =   -1  'True
         Caption         =   "0"
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   420
         TabIndex        =   2
         Top             =   390
         Width           =   6930
      End
   End
End
Attribute VB_Name = "frmMensagemAtivacaoRotina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mIntIdForm            As Integer
'Public mstrNomeForm         As String
Public frm                   As Form
Public mlngHelpContextId     As Long
Private mblnTemAcessoDireto  As Boolean
Private mblnAcaoBotaoOk      As Boolean


Public Property Get TemAcessoDireto() As Boolean
    TemAcessoDireto = mblnTemAcessoDireto
End Property

Public Property Let TemAcessoDireto(ByVal bTemAcessoDireto As Boolean)
    mblnTemAcessoDireto = bTemAcessoDireto
End Property

Public Sub cmdOk_Click()
    If Not mblnAcaoBotaoOk Then
        If Not mblnTemAcessoDireto Then
            If chkAlerta.value Then
                Call AtualizaAlertaAcessoRotina(mIntIdForm)
            End If
        End If
        Call ChamaTelaDireto
    Else
        Unload Me
    End If
End Sub

Public Sub MostrarFormulario()
    Call ValidacaoInicial
End Sub

Private Sub ChamaTelaDireto()
    Unload frmMensagemAtivacaoRotina
    Call escreveIdFormArquivo(gstrArquivoIdForms, gstrModuloFinanceiro, CStr(mlngHelpContextId), frm.name, NomeModulo)
    Call mostrarForm(frm, mlngHelpContextId)
End Sub

Private Sub ValidacaoInicial()
    Dim intDiasFaltantes  As Integer
    Dim blnPrimeiroAcesso As Boolean
    Dim blnNaoAlerta      As Boolean
            
    mblnAcaoBotaoOk = False
    TemAcessoDireto = False

    If Not frm Is Nothing And mIntIdForm > 0 And mlngHelpContextId > 0 Then
        If ValidaPrimeiroAcessoRotina(mIntIdForm) Then
            InserePrimeiroAcessoRotina (mIntIdForm)
            lblMsg.Caption = "A tela de Conciliação Bancária Automática estará disponível de forma demonstrativa por 60 dias."
            blnPrimeiroAcesso = True
        End If
        intDiasFaltantes = ValidaDiasAcessoRotina(mIntIdForm)
        If intDiasFaltantes > 0 Then
            If Not blnPrimeiroAcesso Then
                'Se estiver faltando mais de 10 dias sempre mostra o alerta
                If intDiasFaltantes >= 10 Then
                    If ValidaAlertaAcessoRotina(mIntIdForm) Then
                        lblMsg.Caption = "Você ainda tem " & intDiasFaltantes & " dia(s) para utilizar esta rotina de forma demonstrativa. "
                        chkAlerta.Enabled = True
                    Else
                        mblnTemAcessoDireto = True
                        TemAcessoDireto = True
                    End If
                Else
                     lblMsg.Caption = "Você ainda tem " & intDiasFaltantes & " dia(s) para utilizar esta rotina de forma demonstrativa. "
                     chkAlerta.Enabled = False
                End If
            End If
        Else
           lblMsg.Caption = "O seu acesso a esta rotina de forma demonstrativa expirou. " & vbNewLine & "Para utilizar novamente entre em contato com o departamento Comercial."
           chkAlerta.Enabled = False
           mblnAcaoBotaoOk = True
           Exit Sub
        End If
    Else
        lblMsg.Caption = "Ocorreu um problema para acessar esta rotina. Entre em contato com o Administrador do Sistema."
        chkAlerta.Enabled = False
        mblnAcaoBotaoOk = True
        Exit Sub
    End If
End Sub
