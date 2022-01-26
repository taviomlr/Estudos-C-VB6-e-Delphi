VERSION 5.00
Begin VB.Form fdlgDuplicatas 
   KeyPreview      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " em aberto"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "DlgDupls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   9
      Left            =   3360
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   8
      Left            =   3360
      TabIndex        =   16
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDupl 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   25
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdDupl 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   24
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   7
      Left            =   840
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   6
      Left            =   840
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   5
      Left            =   3360
      TabIndex        =   23
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   4
      Left            =   3360
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   3
      Left            =   840
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   2
      Left            =   840
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtDupl 
      Height          =   315
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000014&
      Index           =   11
      X1              =   2535
      X2              =   2535
      Y1              =   120
      Y2              =   2880
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000010&
      Index           =   10
      X1              =   2520
      X2              =   2520
      Y1              =   120
      Y2              =   2880
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Final:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   14
      Left            =   2760
      TabIndex        =   17
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Inicial:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   13
      Left            =   2760
      TabIndex        =   15
      Top             =   240
      Width           =   450
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Número do &Banco"
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
      Height          =   195
      Index           =   12
      Left            =   2760
      TabIndex        =   14
      Top             =   0
      Width           =   1530
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000014&
      Index           =   9
      X1              =   2640
      X2              =   4920
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000010&
      Index           =   8
      X1              =   2640
      X2              =   4920
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Final:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Inicial:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "V&alores"
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
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   645
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000014&
      Index           =   7
      X1              =   120
      X2              =   2400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000010&
      Index           =   6
      X1              =   120
      X2              =   2400
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Data de &Emissão"
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
      Height          =   195
      Index           =   6
      Left            =   2760
      TabIndex        =   19
      Top             =   960
      Width           =   1440
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000014&
      Index           =   5
      X1              =   2640
      X2              =   4920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   2640
      X2              =   4920
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Final:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   2760
      TabIndex        =   22
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Inicial:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   2760
      TabIndex        =   20
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Final:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Inicial:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Data de &Vencimento"
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
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1740
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   120
      X2              =   2400
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   120
      X2              =   2400
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Final:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Inicial:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   450
   End
   Begin VB.Label lblDupl 
      AutoSize        =   -1  'True
      Caption         =   "Número da &Nota"
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
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   30
      Width           =   1395
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   2400
      Y1              =   135
      Y2              =   135
   End
   Begin VB.Line lnDupl 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   120
      X2              =   2400
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "fdlgDuplicatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gstrTabela   As String

Private Const SEC_FLTDUPL$ = "Filtro_de_Duplicatas"
Private Const KEY_NOTAINI$ = "NotaIni"
Private Const KEY_NOTAFIM$ = "NoraFim"
Private Const KEY_VENCINI$ = "VencIni"
Private Const KEY_VENCFIM$ = "VencFim"
Private Const KEY_EMISINI$ = "EmisIni"
Private Const KEY_EMISFIM$ = "EmisFim"
Private Const KEY_VLRINI$ = "ValorIni"
Private Const KEY_VLRFIM$ = "ValorFim"

Private m_strTipo   As String         '// Define se são duplicatas/lançamentos a receber ou a pagar
Private m_bolCancel As Boolean        '// Configura se o usuário cancelou ou não
Private m_strExp    As String         '// Instrução SQL formada

' PROPERTY..: Tipo
' Objetivo..: Configura a janela quando devem ser selecionados
'             registro a pagar ou a receber.
' ---------------------------------------------------------------------
Public Property Let tipo(strTipo As String)
  m_strTipo = strTipo
  Me.Caption = Me.Caption & IIf((strTipo = "P"), " a pagar", " a receber")
End Property

' PROPERTY..: Cancel
' Objetivo..: Informa se o usuário cancelou ou não o filtro.
' ---------------------------------------------------------------------
Public Property Get Cancel() As Boolean
  Cancel = m_bolCancel
End Property

' PROPERTY..: Expressao
' Objetivo..: Retorna a expressão de consulta formada pelas informações
'             fornecidas pelo usuário.
' ---------------------------------------------------------------------
Public Property Get Expressao() As String
  Expressao = m_strExp
End Property

' EVENT.....: cmdDupl_Click
' Objetivo..: Executa as funções dos botões da janela.
' ---------------------------------------------------------------------
Private Sub cmdDupl_Click(Index As Integer)
  Select Case (Index)
    Case ZERO
      SetPtr vbHourglass
      If (FiltroExp()) Then
        m_bolCancel = False
        Me.Hide
      End If
      SetPtr vbDefault
    Case 1
      m_bolCancel = True
      Me.Hide
  End Select
End Sub

' EVENT.....: Form_Load
' Objetivo..: Configura o formulário para sua abertura
' -----------------------------------------------------------------------
Private Sub Form_Load()

  CenterForm Me
  
  Me.Caption = gstrTabela & Me.Caption

  txtDupl(0).MaxLength = Len(fMask(gstrTabela, "Nota"))   '// Nota
  txtDupl(1).MaxLength = txtDupl(0).MaxLength

  txtDupl(2).MaxLength = 10               '// Data de Vencimento
  txtDupl(3).MaxLength = 10

  txtDupl(4).MaxLength = 10               '// Data de Emissão
  txtDupl(5).MaxLength = 10

  '// Carregando os valores gravados no arquivo .ini

  txtDupl(0).Text = ReadSettings(SEC_FLTDUPL, KEY_NOTAINI, NUL)
  txtDupl(1).Text = ReadSettings(SEC_FLTDUPL, KEY_NOTAFIM, NUL)

  txtDupl(2).Text = ReadSettings(SEC_FLTDUPL, KEY_VENCINI, NUL)
  txtDupl(3).Text = ReadSettings(SEC_FLTDUPL, KEY_VENCFIM, NUL)

  txtDupl(4).Text = ReadSettings(SEC_FLTDUPL, KEY_EMISINI, NUL)
  txtDupl(5).Text = ReadSettings(SEC_FLTDUPL, KEY_EMISFIM, NUL)

  txtDupl(6).Text = ReadSettings(SEC_FLTDUPL, KEY_VLRINI, NUL)
  txtDupl(7).Text = ReadSettings(SEC_FLTDUPL, KEY_VLRFIM, NUL)

End Sub

' EVENT.....: Form_QueryUnload
' Objetivo..: Verifica se o formulário pode ser fechado.
' ---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (UnloadMode <> vbFormCode) Then
    m_bolCancel = True
    Cancel = True
    Me.Hide
  End If
End Sub

' EVENT.....: Form_Unload
' Objetivo..: Grava os valores padrão no arquivo .ini.
' ---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)

  WriteSettings SEC_FLTDUPL, KEY_NOTAINI, txtDupl(0).Text
  WriteSettings SEC_FLTDUPL, KEY_NOTAFIM, txtDupl(1).Text
  WriteSettings SEC_FLTDUPL, KEY_VENCINI, txtDupl(2).Text
  WriteSettings SEC_FLTDUPL, KEY_VENCFIM, txtDupl(3).Text
  WriteSettings SEC_FLTDUPL, KEY_EMISINI, txtDupl(4).Text
  WriteSettings SEC_FLTDUPL, KEY_EMISFIM, txtDupl(5).Text
  WriteSettings SEC_FLTDUPL, KEY_VLRINI, txtDupl(6).Text
  WriteSettings SEC_FLTDUPL, KEY_VLRFIM, txtDupl(7).Text

  Set fdlgDuplicatas = Nothing

End Sub

' EVENT.....: txtDupl_GotFocus
' Objetivo..: Exibe mensagens descritivas na barra de status do programa
' ---------------------------------------------------------------------
Private Sub txtDupl_GotFocus(Index As Integer)
  Selecione txtDupl(Index)

  Select Case (Index)
    Case 0, 1: MsgBar "Número da nota" & ResolveResString(IDS_PGDN, resUM, gstrTabela)
    Case 2, 3: MsgBar "Data de Vencimento" & ResolveResString(IDS_PGDN, resUM, gstrTabela)
    Case 4, 5: MsgBar "Data de Emissão" & ResolveResString(IDS_PGDN, resUM, gstrTabela)
    Case 6, 7: MsgBar "Valores"
    Case 8, 9: MsgBar "Código do Banco" & ResolveResString(IDS_PGDN, resUM, gstrTabela)
  End Select

End Sub

' EVENT.....: txtDupl_KeyDown
' Objetivo..: Abre a janela de pesquisa para os campos da janela.
' ---------------------------------------------------------------------
Private Sub txtDupl_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sExpSearch As String            '// Expressão de Busca

  If ((Shift = ZERO) And (KeyCode = vbKeyPageDown)) Then

    If (Index < 6) Then
      sExpSearch = "SELECT Nota, Parcela, Tipo, Empresa, Descrição, " & _
                   "Vencimento, Emissão, ([Valor Original] + Acréscimo - Abatimento) " & _
                   "AS Valor FROM " & gstrTabela & " WHERE PagRec = '" & m_strTipo & "';"
    Else
      sExpSearch = "Bancos"
    End If

    Select Case (Index)
      Case 0 To 5
        PCampo gstrTabela, sExpSearch, PB_CAMPO, txtDupl(Index), "Nota"

      Case 8, 9
        PCampo "Bancos", sExpSearch, PB_CAMPO, txtDupl(Index), "Banco"
    End Select

  End If
End Sub

' EVENT.....: txtDupl_KeyPress
' Objetivo..: Define a máscara para cada campo conforme a estrutura
'             do Banco de Dados.
' ---------------------------------------------------------------------
Private Sub txtDupl_KeyPress(Index As Integer, KeyAscii As Integer)
Dim iSelStart As Integer

  iSelStart = txtDupl(Index).SelStart

  Select Case (Index)
    Case 0: SetMascara KeyAscii, iSelStart, fMask(gstrTabela, IIf((gstrTabela = "Duplicatas"), "Nota", "Código"))
    Case 1: SetMascara KeyAscii, iSelStart, fMask(gstrTabela, IIf((gstrTabela = "Duplicatas"), "Nota", "Código")), txtDupl(0).hWnd
    Case 2 To 5: SetMascara KeyAscii, iSelStart, MASK_DATA
    Case 6, 7: DMoeda KeyAscii
    Case 8: SetMascara KeyAscii, iSelStart, fMask("Bancos", "Banco")
    Case 9: SetMascara KeyAscii, iSelStart, fMask("Bancos", "Banco"), txtDupl(8).hWnd
  End Select

End Sub

' FUNCTION..: FiltroExp
' Objetivo..: Forma a instrução de seleção de dados com as informações
'             fornecidas pelo usuário.
' Retorna...: True se obtiver sucesso, False se não.
' ---------------------------------------------------------------------
Private Function FiltroExp() As Boolean
Const V_TOTAL$ = "([Valor Original] + Acréscimo - Abatimento)"

Dim nCodIni As Long               '// Código Inicial
Dim nCodFim As Long               '// Código Final
Dim dDatIni As Date               '// Data Inicial
Dim dDatFim As Date               '// Data Final
Dim cValIni As Currency           '// Valor Inicial
Dim cValFim As Currency           '// Valor Final

  If gstrTabela = "Duplicatas" Then
    If gTipoDB = Access Then
        m_strExp = "SELECT Nota, Parcela, Empresa, Emissão, Vencimento, " & _
                   "FORMAT(" & V_TOTAL & ", ""#,##0.00""), 'D' as Origem FROM Duplicatas " & _
                   "WHERE PagRec = '" & m_strTipo & "' AND Pagamento IS NULL AND (Borderô = 0 OR Borderô IS NULL)"
    Else
        m_strExp = "SELECT Nota, Parcela, Empresa, Emissão, Vencimento, " & _
                   "CONVERT(varchar,CAST((" & V_TOTAL & ") AS MONEY),1), 'D' as Origem FROM Duplicatas " & _
                   "WHERE PagRec = '" & m_strTipo & "' AND Pagamento IS NULL AND (Borderô = 0 OR Borderô IS NULL)"
    End If
  Else
    If gTipoDB = Access Then
        m_strExp = "SELECT Código, '' as Parcela, Empresa, Emissão, Vencimento, " & _
                   "FORMAT(" & V_TOTAL & ", ""#,##0.00""), 'L' as Origem FROM Lançamentos " & _
                   "WHERE PagRec = '" & m_strTipo & "' AND Pagamento IS NULL AND (Borderô = 0 OR Borderô IS NULL)"
    Else
        m_strExp = "SELECT Código, '' as Parcela, Empresa, Emissão, Vencimento, " & _
                   "CONVERT(varchar,CAST((" & V_TOTAL & ") AS MONEY),1), 'L' as Origem FROM Lançamentos " & _
                   "WHERE PagRec = '" & m_strTipo & "' AND Pagamento IS NULL AND (Borderô = 0 OR Borderô IS NULL)"
    End If
  End If

  '// Verificando se o usuário filtrou por Número de Nota

  nCodIni = CLngDef(txtDupl(0).Text)
  nCodFim = CLngDef(txtDupl(1).Text)

  Dim strCampo As String
  strCampo = IIf(gstrTabela = "Duplicatas", "Nota", "Código")

  If ((nCodIni > 0) And (nCodFim > 0)) Then
    If (nCodIni = nCodFim) Then
      Concat m_strExp, " AND " & strCampo & " = ", CStr(nCodIni)
    Else
      Concat m_strExp, wsprintf(" AND (" & strCampo & " BETWEEN %l AND %l)", nCodIni, nCodFim)
    End If
  ElseIf ((nCodIni > 0) And (nCodFim = 0)) Then
    Concat m_strExp, " AND " & strCampo & " >= ", CStr(nCodIni)
  ElseIf ((nCodIni = 0) And (nCodFim > 0)) Then
    Concat m_strExp, " AND " & strCampo & " <= ", CStr(nCodFim)
  End If

  '// Verificando se o usuário filtrou por código de Banco

  nCodIni = CLngDef(txtDupl(8).Text)
  nCodFim = CLngDef(txtDupl(9).Text)

  If ((nCodIni > 0) And (nCodFim > 0)) Then
    If (nCodIni = nCodFim) Then
      Concat m_strExp, " AND Banco = ", CStr(nCodIni)
    Else
      Concat m_strExp, wsprintf(" AND (Banco BETWEEN %l AND %l)", nCodIni, nCodFim)
    End If
  ElseIf ((nCodIni > 0) And (nCodFim = 0)) Then
    Concat m_strExp, " AND Banco >= ", CStr(nCodIni)
  ElseIf ((nCodIni = 0) And (nCodFim > 0)) Then
    Concat m_strExp, " AND Banco <= ", CStr(nCodFim)
  End If

  '// Verificando se o usuário filtrou por data de vencimento

  If (IsValid(txtDupl(2).Text)) Then
    If (Not EData(txtDupl(2).Text)) Then
      MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, "Vencimento Inicial")
      Exit Function
    End If
  End If

  If (IsValid(txtDupl(3).Text)) Then
    If (Not EData(txtDupl(3).Text)) Then
      MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, "Vencimento Final")
      Exit Function
    End If
  End If

  dDatIni = CDateDef(txtDupl(2).Text)
  dDatFim = CDateDef(txtDupl(3).Text)

  If ((Not IsEmptyDate(dDatIni)) And (Not IsEmptyDate(dDatFim))) Then
    If (DateDiff(DD_DIA, dDatIni, dDatFim) = ZERO) Then
      Concat m_strExp, wsprintf(" AND Vencimento = #%q#", dDatIni)
    Else
      Concat m_strExp, wsprintf(" AND (Vencimento BETWEEN #%q# AND #%q#)", dDatIni, dDatFim)
    End If
  ElseIf ((Not IsEmptyDate(dDatIni)) And (IsEmptyDate(dDatFim))) Then
    Concat m_strExp, wsprintf(" AND Vencimento >= #%q#", dDatIni)
  ElseIf ((IsEmptyDate(dDatIni)) And (Not IsEmptyDate(dDatFim))) Then
    Concat m_strExp, wsprintf(" AND Vencimento <= #%q#", dDatFim)
  End If

  '// Verificando se o usuário filtrou por data de emissão

  If (IsValid(txtDupl(4).Text)) Then
    If (Not EData(txtDupl(4).Text)) Then
      MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, "Emissão Inicial")
      Exit Function
    End If
  End If

  If (IsValid(txtDupl(5).Text)) Then
    If (Not EData(txtDupl(5).Text)) Then
      MsgFunc ResolveResString(IDS_DATAINVALIDA, resUM, "Emissão Final")
      Exit Function
    End If
  End If

  dDatIni = CDateDef(txtDupl(4).Text)
  dDatFim = CDateDef(txtDupl(5).Text)

  If ((Not IsEmptyDate(dDatIni)) And (Not IsEmptyDate(dDatFim))) Then
    If (DateDiff(DD_DIA, dDatIni, dDatFim) = ZERO) Then
      Concat m_strExp, wsprintf(" AND Emissão = #%q#", dDatIni)
    Else
      Concat m_strExp, wsprintf(" AND (Emissão BETWEEN #%q# AND #%q#)", dDatIni, dDatFim)
    End If
  ElseIf ((Not IsEmptyDate(dDatIni)) And (IsEmptyDate(dDatFim))) Then
    Concat m_strExp, wsprintf(" AND Emissão >= #%q#", dDatIni)
  ElseIf ((IsEmptyDate(dDatIni)) And (Not IsEmptyDate(dDatFim))) Then
    Concat m_strExp, wsprintf(" AND Emissão <= #%q#", dDatFim)
  End If

  '// Verificando se o usuário filtrou por valor

  cValIni = CCurDef(txtDupl(6).Text)
  cValFim = CCurDef(txtDupl(7).Text)

  If ((cValIni > 0) And (cValFim > 0)) Then
    If (cValIni = cValFim) Then
      Concat m_strExp, " AND ", V_TOTAL, " = ", ValStr(cValIni)
    Else
      Concat m_strExp, wsprintf(" AND (%s BETWEEN %s AND %s)", V_TOTAL, ValStr(cValIni), ValStr(cValFim))
    End If
  ElseIf ((cValIni > 0) And (cValFim = 0)) Then
    Concat m_strExp, " AND ", V_TOTAL, " >= ", ValStr(cValIni)
  ElseIf ((cValIni = 0) And (cValFim > 0)) Then
    Concat m_strExp, " AND ", V_TOTAL, " <= ", ValStr(cValFim)
  End If
  FiltroExp = True

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
