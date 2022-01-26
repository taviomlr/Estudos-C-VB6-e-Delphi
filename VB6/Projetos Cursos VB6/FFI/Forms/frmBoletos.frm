VERSION 5.00
Begin VB.Form frmBoletos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processamento de Boletos Bancários"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmBoletos.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBoleto 
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
      Height          =   3915
      Left            =   30
      TabIndex        =   15
      Top             =   -30
      Width           =   9255
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   0
         Left            =   1290
         TabIndex        =   4
         Top             =   1620
         Width           =   1005
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   1
         Left            =   2340
         TabIndex        =   5
         Top             =   1620
         Width           =   1005
      End
      Begin VB.TextBox txtInstrucoes 
         Height          =   1065
         Left            =   3420
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2610
         Width           =   5685
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   3
         Left            =   2340
         TabIndex        =   7
         Top             =   1980
         Width           =   1005
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   2
         Left            =   1290
         TabIndex        =   6
         Top             =   1980
         Width           =   1005
      End
      Begin VB.TextBox txtLocalPagamento 
         Height          =   705
         Left            =   3420
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1590
         Width           =   5685
      End
      Begin VB.ComboBox cboOrigem 
         Height          =   315
         ItemData        =   "frmBoletos.frx":0442
         Left            =   1290
         List            =   "frmBoletos.frx":044F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtBanco 
         Height          =   315
         Left            =   1290
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmBoletos.frx":0473
         Left            =   1290
         List            =   "frmBoletos.frx":0475
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   4
         Left            =   1290
         TabIndex        =   8
         Top             =   2640
         Width           =   1005
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   5
         Left            =   2340
         TabIndex        =   9
         Top             =   2640
         Width           =   1005
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   6
         Left            =   1290
         TabIndex        =   10
         Top             =   3000
         Width           =   1005
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   7
         Left            =   2340
         TabIndex        =   11
         Top             =   3000
         Width           =   1005
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   8
         Left            =   1290
         TabIndex        =   12
         Top             =   3360
         Width           =   1005
      End
      Begin VB.TextBox txtBoleto 
         Height          =   315
         Index           =   9
         Left            =   2340
         TabIndex        =   13
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label lblBoletos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial:"
         Height          =   195
         Index           =   5
         Left            =   1365
         TabIndex        =   32
         Top             =   2370
         Width           =   855
      End
      Begin VB.Label lblBoletos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Data Final:"
         Height          =   195
         Index           =   3
         Left            =   2415
         TabIndex        =   31
         Top             =   2370
         Width           =   795
      End
      Begin VB.Label lblFra 
         Alignment       =   1  'Right Justify
         Caption         =   "Lanc./Duplic.:"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   30
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label lblBoletos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Banco:"
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   29
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblBoletos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Inicial:"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   28
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label lblBoletos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Final:"
         Height          =   195
         Index           =   2
         Left            =   2610
         TabIndex        =   27
         Top             =   1380
         Width           =   405
      End
      Begin VB.Label lblFra 
         Alignment       =   1  'Right Justify
         Caption         =   "Parcelas:"
         Height          =   195
         Index           =   3
         Left            =   375
         TabIndex        =   26
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label lblFra 
         AutoSize        =   -1  'True
         Caption         =   "Local de Pagamento"
         Height          =   195
         Index           =   4
         Left            =   3450
         TabIndex        =   25
         Top             =   1350
         Width           =   1470
      End
      Begin VB.Label lblFra 
         AutoSize        =   -1  'True
         Caption         =   "Instruções"
         Height          =   195
         Index           =   2
         Left            =   3450
         TabIndex        =   24
         Top             =   2370
         Width           =   735
      End
      Begin VB.Label lblBoletos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Origem:"
         Height          =   195
         Index           =   13
         Left            =   375
         TabIndex        =   23
         Top             =   660
         Width           =   885
      End
      Begin VB.Label lblDescBanco 
         Caption         =   "lblDescBanco"
         Height          =   195
         Left            =   2730
         TabIndex        =   22
         Top             =   300
         Width           =   6180
      End
      Begin VB.Label lblBoletos 
         Alignment       =   1  'Right Justify
         Caption         =   "&Tipo:"
         Height          =   195
         Index           =   4
         Left            =   375
         TabIndex        =   21
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label lblFra 
         Alignment       =   1  'Right Justify
         Caption         =   "Liberação:"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   20
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblFra 
         Alignment       =   1  'Right Justify
         Caption         =   "Vencimento:"
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   19
         Top             =   3060
         Width           =   885
      End
      Begin VB.Label lblFra 
         Alignment       =   1  'Right Justify
         Caption         =   "Emissão:"
         Height          =   195
         Index           =   6
         Left            =   375
         TabIndex        =   18
         Top             =   3420
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdBoleto 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   14
      Top             =   4020
      Width           =   1335
   End
   Begin VB.CommandButton cmdBoleto 
      Caption         =   "&Processar"
      Height          =   375
      Index           =   0
      Left            =   6480
      TabIndex        =   0
      Top             =   4020
      Width           =   1335
   End
End
Attribute VB_Name = "frmBoletos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strDatas As String
Private intCamara As Integer

Private Sub ProcessaBoletos()
    Dim sSql       As String
    Dim rsBoletos  As Object
    Dim sTabela    As String
    Dim sWhere     As String
    Dim lngProSeq  As Double
    Dim objCNAB240 As clsCNAB
    Dim strUpdate  As String

On Error GoTo erroBoletos
    If txtBanco.Text = "" Then
       MsgFunc "Banco não informado!", vbInformation
       Exit Sub
    End If
    
 
    Dim d As New CDuplicata
    d.AtualizarValorMoraDiaria
    Set d = Nothing
    
    Dim l As New CLancamento
    l.AtualizarValorMoraDiaria
    Set l = Nothing
    
    SetPtr vbHourglass
     
    sSql = ""
    If cboOrigem.Text <> "Lançamentos" Then
        sSql = sSql & "SELECT "
        sSql = sSql & "  D.NOTA AS NUM, "
        sSql = sSql & "  D.EMPRESA, "
        sSql = sSql & "  D.BANCO, "
        sSql = sSql & " 'D' AS ORIGEM, "
        sSql = sSql & "  D.TIPO, "
        sSql = sSql & "  D.PARCELA, "
        sSql = sSql & "  D.VENCIMENTO, "
        sSql = sSql & "  D.SEQNOSSONUMERO, "
        sSql = sSql & "  B.CONVÊNIO, "
        sSql = sSql & "  B.REGISTRO, "
        sSql = sSql & "  B.[Código da Empresa] as CodEmpresa, "
        sSql = sSql & "  B.CÂMARA, "
        sSql = sSql & "  B.[NÚMERO AGÊNCIA], "
        sSql = sSql & "  B.[DÍGITO DA AGÊNCIA], "
        sSql = sSql & "  B.[NÚMERO CONTA], "
        sSql = sSql & "  B.[DÍGITO DA CONTA], "
        sSql = sSql & "  B.CODCED, "
        sSql = sSql & "  D.AGECCE, "
        sSql = sSql & "  D.NOSNUM, "
        sSql = sSql & "  D.LINDIG, "
        sSql = sSql & "  D.CODBAR, "
        sSql = sSql & "  B.CARTEIRA, "
        sSql = sSql & "  (D.[Valor Original] + D.[Acréscimo] - D.Abatimento) AS VALOR "
        sSql = sSql & "FROM "
        sSql = sSql & "  DUPLICATAS As D "
        sSql = sSql & "INNER JOIN "
        sSql = sSql & "  BANCOS As B "
        sSql = sSql & "ON "
        sSql = sSql & "  B.BANCO = D.BANCO "
        sSql = sSql & "WHERE D.PAGREC = 'R' "
        sSql = sSql & "  AND (D.LINDIG = '' OR D.LINDIG IS NULL) "
        sSql = sSql & "  AND D.BANCO = " & txtBanco.Text & " "
        If cboTipo.Text <> "Todos" Then sSql = sSql & "  AND D.TIPO = '" & cboTipo.Text & "' "
        'Número do Documento
        If txtBoleto(0).Text <> "" Then
            sSql = sSql & "  AND D.NOTA >= " & txtBoleto(0).Text & " "
        End If
        If txtBoleto(1).Text <> "" Then
            sSql = sSql & "  AND D.NOTA <= " & txtBoleto(1).Text & " "
        End If
        'Parcela
        If txtBoleto(2).Text <> "" Then
            sSql = sSql & "  AND D.PARCELA >= " & txtBoleto(2).Text & " "
        End If
        If txtBoleto(3).Text <> "" Then
            sSql = sSql & "  AND D.PARCELA <= " & txtBoleto(3).Text & " "
        End If
        'Data de Liberação
        If txtBoleto(4).Text <> "" Then
            sSql = sSql & "  AND D.LIBERAÇÃO >= " & InverteData(txtBoleto(4).Text, True) & " "
        End If
        If txtBoleto(5).Text <> "" Then
            sSql = sSql & "  AND D.LIBERAÇÃO <= " & InverteData(txtBoleto(5).Text, True) & " "
        End If
        'Data de Vencimento
        If txtBoleto(6).Text <> "" Then
            sSql = sSql & "  AND D.VENCIMENTO >= " & InverteData(txtBoleto(6).Text, True) & " "
        End If
        If txtBoleto(7).Text <> "" Then
            sSql = sSql & "  AND D.VENCIMENTO <= " & InverteData(txtBoleto(7).Text, True) & " "
        End If
        'Data de Emissão
        If txtBoleto(8).Text <> "" Then
            sSql = sSql & "  AND D.EMISSÃO >= " & InverteData(txtBoleto(8).Text, True) & " "
        End If
        If txtBoleto(9).Text <> "" Then
            sSql = sSql & "  AND D.EMISSÃO <= " & InverteData(txtBoleto(9).Text, True) & " "
        End If
        'pt. 84059 - Abner Luidi Hempkemaier(29/10/2007)
        If cboOrigem.Text <> "Todos" Then
            If ConfigSys.OrdemGeracaoBoleto = "V" Then
                sSql = sSql & "ORDER BY VENCIMENTO "
            Else
                sSql = sSql & "ORDER BY D.NOTA "
            End If
        End If
    End If
    
    If cboOrigem.Text = "Todos" Then
        sSql = sSql & "UNION ALL "
    End If
    
    If cboOrigem.Text <> "Duplicatas" Then
        sSql = sSql & "SELECT "
        sSql = sSql & "  L.CÓDIGO AS NUM, "
        sSql = sSql & "  L.EMPRESA, "
        sSql = sSql & "  L.BANCO, "
        sSql = sSql & " 'L' AS ORIGEM, "
        sSql = sSql & "  L.TIPO, "
        sSql = sSql & "  L.PARCELA, "
        sSql = sSql & "  L.VENCIMENTO, "
        sSql = sSql & "  L.SEQNOSSONUMERO, "
        sSql = sSql & "  B.CONVÊNIO, "
        sSql = sSql & "  B.REGISTRO, "
        sSql = sSql & "  B.[Código da Empresa] as CodEmpresa, "
        sSql = sSql & "  B.CÂMARA, "
        sSql = sSql & "  B.[NÚMERO AGÊNCIA], "
        sSql = sSql & "  B.[DÍGITO DA AGÊNCIA], "
        sSql = sSql & "  B.[NÚMERO CONTA], "
        sSql = sSql & "  B.[DÍGITO DA CONTA], "
        sSql = sSql & "  B.CODCED, "
        sSql = sSql & "  L.AGECCE, "
        sSql = sSql & "  L.NOSNUM, "
        sSql = sSql & "  L.LINDIG, "
        sSql = sSql & "  L.CODBAR, "
        sSql = sSql & "  B.CARTEIRA, "
        sSql = sSql & "  (L.[Valor Original] + L.[Acréscimo] - L.Abatimento) AS VALOR "
        sSql = sSql & "FROM "
        sSql = sSql & "  LANÇAMENTOS As L "
        sSql = sSql & "INNER JOIN "
        sSql = sSql & "  BANCOS As B "
        sSql = sSql & "ON "
        sSql = sSql & "  B.BANCO = L.BANCO "
        sSql = sSql & "WHERE L.PAGREC = 'R' "
        sSql = sSql & " AND (L.LINDIG = '' OR L.LINDIG IS NULL) "
        sSql = sSql & "  AND L.BANCO = " & txtBanco.Text & " "
        If cboTipo.Text <> "Todos" Then sSql = sSql & "  AND L.TIPO = '" & cboTipo.Text & "' "
        'Número do Documento
        If txtBoleto(0).Text <> "" Then
            sSql = sSql & "  AND L.CÓDIGO >= " & txtBoleto(0).Text & " "
        End If
        If txtBoleto(1).Text <> "" Then
            sSql = sSql & "  AND L.CÓDIGO <= " & txtBoleto(1).Text & " "
        End If
        'Data de Liberação
        If txtBoleto(4).Text <> "" Then
            sSql = sSql & "  AND L.LIBERAÇÃO >= " & InverteData(txtBoleto(4).Text, True) & " "
        End If
        If txtBoleto(5).Text <> "" Then
            sSql = sSql & "  AND L.LIBERAÇÃO <= " & InverteData(txtBoleto(5).Text, True) & " "
        End If
        'Data de Vencimento
        If txtBoleto(6).Text <> "" Then
            sSql = sSql & "  AND L.VENCIMENTO >= " & InverteData(txtBoleto(6).Text, True) & " "
        End If
        If txtBoleto(7).Text <> "" Then
            sSql = sSql & "  AND L.VENCIMENTO <= " & InverteData(txtBoleto(7).Text, True) & " "
        End If
        'Data de Emissão
        If txtBoleto(8).Text <> "" Then
            sSql = sSql & "  AND L.EMISSÃO >= " & InverteData(txtBoleto(8).Text, True) & " "
        End If
        If txtBoleto(9).Text <> "" Then
            sSql = sSql & "  AND L.EMISSÃO <= " & InverteData(txtBoleto(9).Text, True) & " "
        End If
        If cboOrigem.Text = "Todos" Then
            If ConfigSys.OrdemGeracaoBoleto = "V" Then
                sSql = sSql & "ORDER BY VENCIMENTO "
            Else
                sSql = sSql & "ORDER BY NUM "
            End If
        Else
            If ConfigSys.OrdemGeracaoBoleto = "V" Then
                sSql = sSql & "ORDER BY VENCIMENTO "
            Else
                sSql = sSql & "ORDER BY L.CÓDIGO "
            End If
        End If
    End If
    
    If AbreRecordset(rsBoletos, sSql, dbOpenDynaset) = WL_OK Then
        lngProSeq = GetFieldValue("PROSEQ", "BANCOS", "BANCO = " & txtBanco.Text, , 1)
        If lngProSeq = 0 Then
            lngProSeq = 1
        End If
        Do
            Set objCNAB240 = New clsCNAB
            With objCNAB240
                .IdentificadorBanco = GetValue(rsBoletos, "CÂMARA", "0")
                .Agencia = GetValue(rsBoletos, "NÚMERO AGÊNCIA", "0")
                .DigitoAgencia = GetValue(rsBoletos, "DÍGITO DA AGÊNCIA", "0")
                .NumeroConta = GetValue(rsBoletos, "NÚMERO CONTA", "0")
                .DigitoConta = GetValue(rsBoletos, "DÍGITO DA CONTA", "0")
                .CodigoCedente = GetValue(rsBoletos, "CODCED", "0")
                .Carteira = GetValue(rsBoletos, "CARTEIRA", "0")
                .ValorDocumento = GetValue(rsBoletos, "VALOR", "0")
                .DataVencimento = GetValue(rsBoletos, "VENCIMENTO", "0")
                .NumeroDocumento = Format(GetValue(rsBoletos, "NUM", "0"), "000000") & Format(GetValue(rsBoletos, "Parcela", "0"), "00")
                .Convenio = GetValue(rsBoletos, "CONVÊNIO", "0")
                .TipoCobranca = GetValue(rsBoletos, "REGISTRO", "CR")
                .CodigoEmpresa = GetValue(rsBoletos, "CodEmpresa", "")
                .PadraoCNAB = GetFieldValue("tipo_layout", "BANCOS", "BANCO = " & txtBanco.Text, , 1)
            
                sWhere = "WHERE 1=1"
                If GetValue(rsBoletos, "ORIGEM") = "D" Then
                    sTabela = "DUPLICATAS"
                    sWhere = sWhere & " AND NOTA = " & GetValue(rsBoletos, "NUM")
                    sWhere = sWhere & " AND EMPRESA = '" & GetValue(rsBoletos, "EMPRESA") & "'"
                    sWhere = sWhere & " AND TIPO = '" & GetValue(rsBoletos, "TIPO") & "'"
                    sWhere = sWhere & " AND PARCELA = " & GetValue(rsBoletos, "PARCELA")
                Else
                    sTabela = "LANÇAMENTOS"
                    sWhere = sWhere & " AND CÓDIGO = " & GetValue(rsBoletos, "NUM")
                    sWhere = sWhere & " AND PARCELA = " & GetValue(rsBoletos, "PARCELA")
                End If
                If .CriarBoleto(lngProSeq) Then
                    If Len(.LinhaDigitavel) = 54 Then
                        strUpdate = "UPDATE " & sTabela & " SET LINDIG = '" & .LinhaDigitavel & "', CODBAR = '"
                        strUpdate = strUpdate & .CodigoBarras & "', NOSNUM='" & .NossoNumero & "', LOCPAG = '"
                        strUpdate = strUpdate & txtLocalPagamento.Text & "', INSTRU='" & txtInstrucoes.Text & "', "
                        strUpdate = strUpdate & " AGECCE = '" & .AgenciaCodigoCedente & "' " & sWhere
                        ExecuteSQL strUpdate
                        'Registra o número da próxima sequencia de boleto
                        lngProSeq = lngProSeq + 1
                    End If
                End If
            End With
            rsBoletos.MoveNext
        Loop Until rsBoletos.EOF
    Else
        Call MsgBox("Nenhum boleto foi processado!", vbInformation, NomeModulo)
        FechaRecordset rsBoletos
        SetPtr vbDefault
        Exit Sub
    End If
    FechaRecordset rsBoletos
    ExecuteSQL "UPDATE BANCOS SET PROSEQ = " & lngProSeq & " WHERE BANCO = " & txtBanco.Text
    Call MsgBox("Processamento de boletos realizado com sucesso!", vbInformation, NomeModulo)
    SetPtr vbDefault
    Exit Sub

erroBoletos:
    MsgBox str(err.Number) & " - " & err.Description, vbCritical, NomeModulo
End Sub

'Descrição..: Sub utilizada para remover o ultimo caracter de uma String
'               a sub foi criada para resolver o problema da EBSBoleto que
'               aloca caracteres loucos no fim da string, e é necessário
'               remove-los.
'Parametros.: [String] Texto retornado pela chamada da função em Delphi.
Private Sub RemoveUltimo(ByRef sTexto As String)
    sTexto = Trim$(sTexto)
    sTexto = Left(sTexto, Len(sTexto) - 1)
End Sub

Private Sub cboOrigem_Click()
    Dim Habilitado  As Boolean
  
    Habilitado = (cboOrigem.Text = "Duplicatas" Or cboOrigem.Text = "Todos")
    lblFra(3).Enabled = Habilitado
    txtBoleto(2).Enabled = Habilitado
    txtBoleto(3).Enabled = Habilitado
End Sub

Private Sub cmdBoleto_Click(Index As Integer)
    If Index < 1 Then
        Call ProcessaBoletos
    Else
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    lblDescBanco.Caption = NUL
    
    'Valores padrão
    cboOrigem.Text = "Duplicatas"
    ComboAddItem cboTipo, "Select Tipo from [Tipos Globais];", "Tipo"
    cboTipo.AddItem "Todos"
    cboTipo.Text = "Todos"
    'Os últimos usados
    txtBoleto(0).Text = LerArquivoASCII("Boletos", "NotaI", "Fox.ini")
    txtBoleto(1).Text = LerArquivoASCII("Boletos", "NotaF", "Fox.ini")
    txtInstrucoes.Text = LerArquivoASCII("Boletos", "Instrucoes", "Fox.ini")
    If Len(LerArquivoASCII("Boletos", "Tipo", "Fox.ini")) > 0 Then
        cboOrigem.Text = LerArquivoASCII("Boletos", "Tipo", "Fox.ini")
    End If
    MsgBar "Verificando cadastro de duplicatas e lançamentos"
    MsgBar ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Gravando para a próxima abertura
    GravarArquivoASCII "Boletos", "Tipo", cboOrigem.Text, "Fox.ini"
    GravarArquivoASCII "Boletos", "NotaI", txtBoleto(0).Text, "Fox.ini"
    GravarArquivoASCII "Boletos", "NotaF", txtBoleto(1).Text, "Fox.ini"
    GravarArquivoASCII "Boletos", "Instrucoes", txtInstrucoes.Text, "Fox.ini"
    Set frmBoletos = Nothing
End Sub

Private Sub txtBanco_Change()
    GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtBanco.Text, lblDescBanco
End Sub

Private Sub txtBanco_GotFocus()
    Selecione txtBanco
End Sub

Private Sub txtBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        PCampo "Bancos", "Bancos", pbCampo, txtBanco, "Banco"
    End If
End Sub

Private Sub txtBanco_LostFocus()
    If txtBanco.Text <> NUL Then
        txtLocalPagamento.Text = GetFieldValue("[Local de Pagamento]", "Bancos", "Banco =" & txtBanco.Text, , NUL)
        txtInstrucoes.Text = GetFieldValue("[Instruções]", "Bancos", "Banco =" & txtBanco.Text, , NUL)
    Else
        txtLocalPagamento.Text = NUL
        txtInstrucoes.Text = NUL
    End If
End Sub

Private Sub txtBoleto_GotFocus(Index As Integer)
    Selecione txtBoleto(Index)
End Sub

Private Sub txtBoleto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strSql As String
    Dim lngDup As Long
    
    'Instrução Padrão de pesquisa
    'Pt. 95368 - Moacir Pfau(03/11/2009)
    strSql = "SELECT " & IIf((cboOrigem.Text = "Duplicatas"), "Nota, Parcela ", "Código") & ", " & _
              "Tipo, Descrição, Vencimento, Pagamento, Empresa , Banco, " & _
              IIf((cboOrigem.Text = "Duplicatas"), "Enviada", "Enviado") & ", " & _
              "[Valor Original], Controle, Abatimento, Acréscimo FROM " & cboOrigem.Text & " WHERE [Valor Original] > 0 " & _
              "AND isnull(" & cboOrigem.Text & ".Pagamento) AND " & cboOrigem.Text & ".PagRec = 'R'"
    If (Shift = 0) And (KeyCode = vbKeyPageDown) Then
        If txtBoleto(Index).Text <> "0" And txtBoleto(Index).Text <> "" Then
            lngDup = txtBoleto(Index).Text
            txtBoleto(Index) = 0
        End If
        Select Case Index
            Case 0, 1
                Call PCampo(cboOrigem.Text, strSql, pbCampo, txtBoleto(Index), IIf((cboOrigem.Text = "Duplicatas"), "Nota", "Código"))
            Case 2, 3
                Call PCampo(cboOrigem.Text, strSql, pbCampo, txtBoleto(Index), "Parcela")
        End Select
        If txtBoleto(Index).Text = "0" Or txtBoleto(Index).Text = "" Then
            txtBoleto(Index).Text = lngDup
        End If
    End If
End Sub

Private Sub txtBoleto_KeyPress(Index As Integer, KeyAscii As Integer)
    'Máscaras dos campos
    Select Case Index
        Case 0, 1
            SetMascara KeyAscii, txtBoleto(Index).SelStart, fMask(cboOrigem.Text, "Código")
        Case 2, 3
            SetMascara KeyAscii, txtBoleto(Index).SelStart, fMask("Duplicatas", "Parcela")
        Case 4 To 9
            SetMascara KeyAscii, txtBoleto(Index).SelStart, MASK_DATA
    End Select
End Sub

'pt. 86383 - Ivo Sousa (07/04/2008)
Private Sub txtBoleto_LostFocus(Index As Integer)

On Error Resume Next
    If Index = 0 Or Index = 1 Then
        If txtBoleto(Index).Text <> "" And txtBoleto(Index).Text <> "0" Then
            If cboOrigem.Text = "Duplicatas" Then
                If GetFieldValue("Nota", "Duplicatas", "Nota = " & txtBoleto(Index).Text & " AND ISNULL(Pagamento) AND PagRec = 'R'", , 0) = 0 Then
                    MsgBox "O código da Duplicata informado não é valido.", vbInformation + vbOKOnly, NomeModulo
                    txtBoleto(Index).Text = 0
                    txtBoleto(Index).SetFocus
                End If
            ElseIf cboOrigem.Text = "Lançamentos" Then
                If GetFieldValue("Código", "Lançamentos", "Código = " & txtBoleto(Index).Text & " AND ISNULL(Pagamento) AND PagRec = 'R'", , 0) = 0 Then
                    MsgBox "O código do Lançamento informado não é valido.", vbInformation + vbOKOnly, NomeModulo
                    txtBoleto(Index).Text = 0
                    txtBoleto(Index).SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub txtInstrucoes_GotFocus()
    Selecione txtInstrucoes
End Sub

Private Sub txtLocalPagamento_GotFocus()
    Selecione txtLocalPagamento
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
