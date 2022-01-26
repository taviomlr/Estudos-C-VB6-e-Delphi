VERSION 5.00
Begin VB.Form frmReprocessaSaldo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reprocessamento de Saldos Bancários"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Banco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   30
      TabIndex        =   8
      Top             =   -30
      Width           =   4575
      Begin Fox.EBSText etxBanco 
         Height          =   330
         Left            =   960
         TabIndex        =   0
         Top             =   330
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
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
      Begin Fox.EBSText etxBancoFinal 
         Height          =   330
         Left            =   2520
         TabIndex        =   1
         Top             =   330
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
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
      Begin Fox.EBSData etdInicial 
         Height          =   330
         Left            =   960
         TabIndex        =   2
         Top             =   690
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Formato         =   1
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
      Begin Fox.EBSData etdFinal 
         Height          =   330
         Left            =   2520
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   690
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         Formato         =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin VB.Label lblInformativa 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Se nenhum banco for informado, o sistema irá realizar o reprocessamento de todos os bancos cadastrados."
         Height          =   465
         Left            =   510
         TabIndex        =   13
         Top             =   1290
         Width           =   4020
      End
      Begin VB.Image imgInformativa 
         Height          =   480
         Left            =   30
         Picture         =   "frmReprocessaSaldo.frx":0000
         Top             =   1260
         Width           =   480
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "Banco"
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   12
         Top             =   390
         Width           =   465
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Index           =   1
         Left            =   2250
         TabIndex        =   11
         Top             =   390
         Width           =   90
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "Período"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   10
         Top             =   780
         Width           =   570
      End
      Begin VB.Label lblFluxo 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Index           =   4
         Left            =   2250
         TabIndex        =   9
         Top             =   750
         Width           =   90
      End
   End
   Begin VB.Frame fraBotoes 
      Height          =   1785
      Left            =   4635
      TabIndex        =   7
      Top             =   -30
      Width           =   1425
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "&Ajuda"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   630
         Width           =   1215
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "&Gerar"
         Height          =   390
         Left            =   120
         TabIndex        =   4
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmReprocessaSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdGerar_Click()
  On Error GoTo err_Handler
  Dim blnTodosBancos As Boolean
  
    If etxBanco.valorInteiro = 0 And etxBancoFinal.valorInteiro = 0 Then
        blnTodosBancos = True
    End If
    If CDate(FirstDay(etdInicial.Data)) And CDate(LastDay(etdFinal.Data)) Then
        If (DateDiff("m", FirstDay(etdInicial.Data), FirstDay(etdFinal.Data)) <= 12) Then
            If (etxBanco.valorInteiro > 0 And etxBancoFinal.valorInteiro > 0) Or blnTodosBancos Or etxBanco.valorInteiro >= etxBancoFinal.valorInteiro Then
                MousePointer = 11
                cmdGerar.Enabled = False
                Call ReprocessamentoSaldoBancario(FirstDay(etdInicial.Data), LastDay(etdFinal.Data), etxBanco.valorInteiro, 0, , , , , etxBancoFinal.valorInteiro, blnTodosBancos)
                MsgBox "Processo finalizado com sucesso!", vbInformation, NomeModulo
                cmdGerar.Enabled = True
                MousePointer = 0
            Else
                MsgBox "Favor informar um intervalo de banco válido.", vbInformation, NomeModulo
            End If
        Else
            MsgBox "Favor digitar uma data inicial com no máximo 12 meses retroativos a data de hoje.", vbInformation, NomeModulo
        End If
    Else
    MsgBox "Favor informar uma data válida.", vbInformation, NomeModulo
    End If
    Exit Sub
err_Handler:
    MousePointer = 0
    MsgBox "Ocorreu um erro ao reprocessar Saldos Bancários", vbInformation, NomeModulo
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub etxBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxBanco.ValorDescricao = "" Then
            etxBanco.valorInteiro = 0
        End If
        PCampo "Banco", "SELECT [Banco], [Nome], [Agência], [Conta] FROM [Bancos]", pbCampo, etxBanco, "Banco"
    End If
End Sub

Private Sub etxBancoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown Then
        If etxBancoFinal.ValorDescricao = "" Then
            etxBancoFinal.valorInteiro = 0
        End If
        PCampo "Banco", "SELECT [Banco], [Nome], [Agência], [Conta] FROM [Bancos]", pbCampo, etxBancoFinal, "Banco"
    End If
End Sub

Private Sub Form_Load()
    Call etxBanco.AddConexao(Aplicacao)
    Call etxBancoFinal.AddConexao(Aplicacao)
    Call CenterForm(Me)
    etdFinal.MesAno = Format(Now(), "mm/yyyy")
    'Pega o ultimo para sugerir data inicial de reprocessamento
    ConfigSys.CarregarRegistro
    If ConfigSys.UltimaDuplicataRetro = "" Then ConfigSys.UltimaDuplicataRetro = Format(Now(), "mm/dd/yyyy")
    If ConfigSys.UltimaLancamentoRetro = "" Then ConfigSys.UltimaLancamentoRetro = Format(Now(), "mm/dd/yyyy")
    
    If CDate(ConfigSys.UltimaDuplicataRetro) <= CDate(ConfigSys.UltimaLancamentoRetro) Then
        If ConfigSys.UltimaDuplicataRetro <> "" Then
            etdInicial.MesAno = Format(ConfigSys.UltimaDuplicataRetro, "mm/yyyy")
        End If
    Else
        etdInicial.MesAno = Format(ConfigSys.UltimaLancamentoRetro, "mm/yyyy")
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

Private Sub Label4_Click()

End Sub
