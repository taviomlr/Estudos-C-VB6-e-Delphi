VERSION 5.00
Begin VB.Form frmManutencaoOCR 
   KeyPreview      =   -1  'True
   Caption         =   "Manutenção de OCR"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2760
      Left            =   7425
      TabIndex        =   8
      Top             =   -45
      Width           =   1500
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   375
         Left            =   135
         TabIndex        =   7
         Top             =   630
         Width           =   1215
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         Height          =   375
         Left            =   135
         TabIndex        =   6
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2760
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   7395
      Begin VB.TextBox txtMotorista 
         Height          =   330
         Left            =   1500
         MaxLength       =   40
         TabIndex        =   5
         Top             =   2145
         Width           =   5775
      End
      Begin VB.TextBox txtPlacaVeiculo 
         Height          =   330
         Left            =   1500
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1760
         Width           =   1230
      End
      Begin VB.TextBox txtPesoBruto 
         BackColor       =   &H8000000E&
         Height          =   330
         Left            =   1500
         MaxLength       =   15
         TabIndex        =   2
         Top             =   970
         Width           =   1230
      End
      Begin VB.TextBox txtPesoLiquido 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         ForeColor       =   &H80000001&
         Height          =   330
         Left            =   1500
         TabIndex        =   3
         Top             =   1365
         Width           =   1230
      End
      Begin VB.TextBox txtTaraCaminhao 
         Height          =   330
         Left            =   1500
         MaxLength       =   15
         TabIndex        =   1
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Motorista:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   15
         Top             =   2250
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Placa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   885
         TabIndex        =   14
         Top             =   1845
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Peso Bruto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   13
         Top             =   1045
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peso Líquido:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tara Caminhão:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   650
         Width           =   1350
      End
      Begin VB.Label lblNumeroOCR 
         AutoSize        =   -1  'True
         Caption         =   "lblNumeroOCR"
         Height          =   195
         Left            =   1470
         TabIndex        =   10
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Left            =   795
         TabIndex        =   9
         Top             =   270
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmManutencaoOCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objCarregamento As COrdCarregamento

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Sub setNumeroCarregamento(lngNumero As Long)
    If lngNumero > 0 Then
        Set objCarregamento = New COrdCarregamento
        Call objCarregamento.carregar(lngNumero)
        Call PreencheCampos
    End If
End Sub

Private Sub PreencheCampos()
    With objCarregamento
        lblNumeroOCR.Caption = .numeroOCR
        txtTaraCaminhao.Text = .pesoCaminhao
        txtPesoLiquido.Text = .PesoLiquido
        txtPesoBruto.Text = .PesoBruto
        txtPlacaVeiculo.Text = .PlacaVeiculo
        txtMotorista.Text = .nomeMotorista
    End With
End Sub

Private Sub cmdConfirmar_Click()
    If ValidaCampos Then
        With objCarregamento
            .pesoCaminhao = txtTaraCaminhao.Text
            .PesoLiquido = txtPesoLiquido.Text
            .PesoBruto = txtPesoBruto.Text
            .PlacaVeiculo = txtPlacaVeiculo.Text
            .nomeMotorista = txtMotorista.Text
        End With
        Call objCarregamento.atualiza
        Call cmdCancelar_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Width = 9060
    Me.Height = 3120
    CenterForm Me
End Sub

Private Sub txtPesoBruto_Change()
    If IsNumeric(txtTaraCaminhao.Text) And IsNumeric(txtPesoBruto.Text) Then
        txtPesoLiquido.Text = CDbl(txtPesoBruto.Text) - CDbl(txtTaraCaminhao.Text)
    End If
End Sub

Private Sub txtPesoBruto_KeyPress(KeyAscii As Integer)
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_inteiro, txtPesoBruto)
    '---------------------------------------------------------------------------
End Sub

Private Sub txtPesoBruto_Validate(Cancel As Boolean)
    
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    If val(txtPesoBruto.Text) < val(txtTaraCaminhao.Text) Then
        MsgBox "O campo Peso Bruto não pode ser menor que o campo Tara Caminhão.", vbInformation, Me.Caption
        Cancel = True
    End If
    '-----------------------------------------------------------------------------
End Sub

Private Sub txtPlacaVeiculo_KeyPress(KeyAscii As Integer)
    Call mascaraPlaca(KeyAscii, txtPlacaVeiculo)
End Sub

Private Sub txtTaraCaminhao_Change()
    If (txtTaraCaminhao.Text) = "" And Trim(txtPesoBruto.Text) = "" Then
        txtTaraCaminhao.Text = 0
        txtPesoBruto.Text = 0
    End If
    If IsNumeric(txtTaraCaminhao.Text) And IsNumeric(txtPesoBruto.Text) Then
        txtPesoLiquido.Text = CDbl(txtPesoBruto.Text) - CDbl(txtTaraCaminhao.Text)
    End If
End Sub

Public Function ValidaCampos() As Boolean
    ValidaCampos = False
    
    If (txtTaraCaminhao.Text) = "" And Trim(txtPesoBruto.Text) = "" Then
        txtTaraCaminhao.Text = 0
        txtPesoBruto.Text = 0
    End If
    
    If val(txtTaraCaminhao.Text) <= 0 Then
        MsgBox "O campo tara do caminhão deve conter um número maior do que ZERO.", vbInformation, Me.Caption
        txtTaraCaminhao.SetFocus
    ElseIf Not IsNumeric(txtPesoLiquido.Text) Then
        MsgBox "o campo peso líquido deve conter um número.", vbInformation, Me.Caption
        txtPesoLiquido.SetFocus
    ElseIf CLng(txtPesoLiquido.Text) <= 0 Then
        MsgBox "O campo Peso líquido deve conter um número maior do que ZERO.", vbInformation, Me.Caption
        'txtPesoLiquido.SetFocus
    ElseIf Trim(txtMotorista.Text) = "" Then
        MsgBox "O campo nome do motorista deve ser preenchido.", vbInformation, Me.Caption
        txtMotorista.SetFocus
    ElseIf Trim(txtPlacaVeiculo.Text) = "" Then
        MsgBox "O campo placa do veiculo deve ser preenchido.", vbInformation, Me.Caption
        txtPlacaVeiculo.SetFocus
    ElseIf Len(txtPlacaVeiculo.Text) < 8 Then
        MsgBox "O campo placa do veiculo deve conter uma placa válida.", vbInformation, Me.Caption
        txtPlacaVeiculo.SetFocus
    Else
        ValidaCampos = True
    End If
End Function

Private Sub txtTaraCaminhao_KeyPress(KeyAscii As Integer)
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    Call validaNumeros(KeyAscii, enumValidaNumero.tipo_inteiro, txtTaraCaminhao)
    '----------------------------------------------------------------------------
End Sub

Private Sub txtTaraCaminhao_Validate(Cancel As Boolean)
    
    'Pt.97090  - Fernando Paludo - (22/02/2010)
    If val(txtTaraCaminhao.Text) > val(txtPesoBruto.Text) Then
        MsgBox "O campo Tara Caminhão não pode ser maior que o campo Peso Bruto.", vbInformation, Me.Caption
        Cancel = True
    End If
    '-----------------------------------------------------------------------------
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
