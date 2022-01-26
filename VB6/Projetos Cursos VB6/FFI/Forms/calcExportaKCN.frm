VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSmapi32.ocx"
Begin VB.Form fcalcExportaKCN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportação e Envio de Movimentação"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "calcExportaKCN.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sai&r"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar Dados..."
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame fraDatas 
      Caption         =   "Informações"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Identificação da Empresa"
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtDatas 
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtDatas 
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   4080
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line lin1 
         BorderColor     =   &H80000010&
         BorderWidth     =   2
         X1              =   120
         X2              =   4080
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblKCN 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label lblKCN 
         AutoSize        =   -1  'True
         Caption         =   "Data &Inicial:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblKCN 
         AutoSize        =   -1  'True
         Caption         =   "Data &Final:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   765
      End
   End
   Begin MSMAPI.MAPIMessages Mmsg 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession Msession 
      Left            =   720
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   0   'False
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "fcalcExportaKCN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnviar_Click()

  '
  ' Validando
  '
  If Not EData(txtDatas(0).Text) Then
    MsgFunc "O campo 'Data Inicial' não contém uma data válida"
    Exit Sub
  End If
  
  If Not EData(txtDatas(1).Text) Then
    MsgFunc "O campo 'Data Final' não contém uma data válida"
    Exit Sub
  End If

  If Len(txtCodigo.Text) = 0 Then
    MsgFunc "O campo 'Código' se refere a identificação de sua empresa. Esse deve ser preenchido"
    Exit Sub
  End If

  Msession.SignOn
  Mmsg.SessionID = Msession.SessionID
  Mmsg.Compose


  'Endereço o qual o e-mail será enviado
  Mmsg.RecipAddress = "online@balan-set.com.br"

  Mmsg.AddressResolveUI = True

  'Exibe o Assunto da Mensagem
  Mmsg.MsgSubject = txtCodigo.Text & " - Período de " & txtDatas(0).Text & " até " & txtDatas(1).Text

  'Exibe o Valor de Texo de E-mails padrões como texto do E-mail
  Mmsg.MsgNoteText = txtCodigo.Text

  ExportaKCN CDateDef(txtDatas(0).Text), CDateDef(txtDatas(1).Text), NUL

  'anexa no final da mensagem
  Mmsg.AttachmentPosition = Len(Mmsg.MsgNoteText)

  'da um nome ao anexo
  Mmsg.AttachmentName = "L" & CStr((Format(txtDatas(0).Text, "ddmmyy")) & ".TXZ")

  'define o caminho e nome do arquivo a anexar
  If ArquivoExiste(AddSepDir(DiretorioTemp) & CStr("L" & CStr(Format(txtDatas(1).Text, "ddmmyy")) & ".TXZ")) Then
    Mmsg.AttachmentPathName = AddSepDir(DiretorioTemp) & CStr("L" & CStr(Format(txtDatas(1).Text, "ddmmyy")) & ".TXZ")
  End If

  'Não exibe a Mensagem antes de Enviar
  Mmsg.send False

  'Desconecta da Conta de E-mail
  Msession.SignOff

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
  CenterForm Me
  
  ' Carregando configurações
  If EData(LerArquivoASCII("KCN", "LastData", AddSepDir(App.Path) & "Fox.ini")) Then
    txtDatas(0).Text = DateAdd("d", 1, CDateDef(LerArquivoASCII("KCN", "LastData", AddSepDir(App.Path) & "Fox.ini")))
  End If
  txtCodigo.Text = LerArquivoASCII("KCN", "Codigo", AddSepDir(App.Path) & "Fox.ini")

End Sub

Private Sub Form_Unload(Cancel As Integer)
  GravarArquivoASCII "KCN", "LastData", txtDatas(1).Text, AddSepDir(App.Path) & "Fox.ini"
  GravarArquivoASCII "KCN", "Codigo", txtCodigo.Text, AddSepDir(App.Path) & "Fox.ini"

  Set fcalcExportaKCN = Nothing
End Sub

Private Sub txtDatas_GotFocus(Index As Integer)
  Selecione txtDatas(Index)
End Sub

Private Sub txtDatas_KeyPress(Index As Integer, KeyAscii As Integer)
  SetMascara KeyAscii, txtDatas(Index).SelStart, MASK_DATA
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
