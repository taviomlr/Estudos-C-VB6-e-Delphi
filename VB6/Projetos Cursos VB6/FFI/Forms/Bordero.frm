VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmBordero 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitação de Borderôs"
   ClientHeight    =   5235
   ClientLeft      =   2430
   ClientTop       =   3375
   ClientWidth     =   9075
   Icon            =   "Bordero.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Tag             =   "Borderôs"
   Begin VB.ComboBox cboOrigem 
      Height          =   315
      ItemData        =   "Bordero.frx":058A
      Left            =   1320
      List            =   "Bordero.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdBordero 
      Caption         =   "&Gravar Borderô"
      Height          =   375
      Index           =   4
      Left            =   2880
      TabIndex        =   16
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdBordero 
      Caption         =   "&Selecionar..."
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdBordero 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   2
      Left            =   7920
      TabIndex        =   19
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdBordero 
      Caption         =   "&Imprimir"
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   18
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdBordero 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   17
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame fraBordero 
      Caption         =   "Principal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8655
      Begin VB.TextBox txtBordero 
         DataField       =   "Código"
         Height          =   315
         Index           =   2
         Left            =   6720
         TabIndex        =   8
         Tag             =   "Borderôs"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cboBordero 
         Height          =   315
         Index           =   1
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboBordero 
         Height          =   315
         Index           =   0
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin ComctlLib.ListView lvwBordero 
         Height          =   1935
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtBordero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   225
         Index           =   1
         Left            =   6120
         MultiLine       =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtBordero 
         DataField       =   "Código"
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Tag             =   "Borderôs"
         Top             =   240
         Width           =   1095
      End
      Begin ComctlLib.ListView lvwBordero 
         Height          =   1215
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblBordero 
         AutoSize        =   -1  'True
         Caption         =   "&Banco:"
         Height          =   195
         Index           =   5
         Left            =   6120
         TabIndex        =   7
         Top             =   240
         Width           =   510
      End
      Begin VB.Label lblBordero 
         AutoSize        =   -1  'True
         Caption         =   "Si&tuação:"
         Height          =   195
         Index           =   4
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblBordero 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   360
      End
      Begin VB.Label lblBordero 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Index           =   2
         Left            =   5520
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblBordero 
         AutoSize        =   -1  'True
         Caption         =   "D&uplicatas"
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
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   915
      End
      Begin VB.Line lnBordero 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   120
         X2              =   7800
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line lnBordero 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   120
         X2              =   7800
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Label lblBordero 
         AutoSize        =   -1  'True
         Caption         =   "Có&digo:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   540
      End
   End
   Begin ComctlLib.TabStrip tabBordero 
      Height          =   4695
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8281
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Borderô"
            Key             =   ""
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
   Begin ComctlLib.ImageList imgBordero 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmBordero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IDB_DUPL = 510    '// Ícone para duplicatas
Private Const IDB_LANC = 511    '// Ícone para lançamentos
Private Const IDC_DUPL = 512    '// Cursor usado nas operações de Drag & Drop

Private nDragFlag As Long       '// Flag para as operações de Drag & Drop

' EVENT.....: cboBordero_Click
' Objetivo..: Mapea as alterações dos controles ComboBox
' ------------------------------------------------------
Private Sub cboBordero_Click(Index As Integer)
  If (Index = ZERO) Then
    If (CLngDef(txtBordero(0).Text)) Then
      Call txtBordero_LostFocus(ZERO)
    End If
  End If
End Sub

' EVENT.....: cboBordero_GotFocus
' Objetivo..: Exibe mensagens de ajuda na barra de status.
' ---------------------------------------------------------------------
Private Sub cboBordero_GotFocus(Index As Integer)
  Select Case (Index)
    Case 0: MsgBar "Tipo das duplicatas/lançamentos: A Pagar ou A Receber"
    Case 1: MsgBar "Situação das duplicatas/lançamentos no borderô"
  End Select
End Sub

Private Sub cboOrigem_LostFocus()
  If cboOrigem.Enabled Then
    Call cboBordero_Click(0)
  End If
End Sub

' EVENT.....: cmdBordero_Click
' Objetivo..: Executa as funções dos botões da janela.
' ---------------------------------------------------------------------
Private Sub cmdBordero_Click(Index As Integer)
  Select Case (Index)
    Case 0, 1           '// Visualizar e Imprimir
     
     If txtBordero(2).Text = "" Then
        MsgBox "Digite um banco válido", vbInformation
        Exit Sub
     End If
    
      cmdBordero(0).Enabled = False
      cmdBordero(1).Enabled = False
      cmdBordero(2).Caption = LoadResString(IDS_CANCELAR)
      cmdBordero(3).Enabled = False
      cmdBordero(4).Enabled = False

      ImprimirBordero IIf(Index, WL_TOPRINTER, WL_TOWINDOW)

      cmdBordero(0).Enabled = True
      cmdBordero(1).Enabled = True
      cmdBordero(2).Caption = LoadResString(IDS_FECHAR)
      cmdBordero(3).Enabled = True
      cmdBordero(4).Enabled = True

    Case 2              '// Cancelar/Fechar
      If (cmdBordero(0).Enabled) Then
        Unload Me
      Else
        Call SendKeysEx(Chr$(27))       '// Simula o pressionamento da tecla ESC
        DoEvents
      End If

    Case 3              '// Selecionar Duplicatas/Lançamentos
      If lvwBordero(1).ListItems.Count > 0 Then
        If MsgBox("Deseja limpar os itens já selecionados?", vbYesNo Or vbQuestion, MsgBoxCaption) = vbYes Then
          lvwBordero(1).ListItems.Clear
          lvwBordero(0).ListItems.Clear
        End If
      End If

      DoEvents          '// Necessário para processar o evento LostFocus do primeiro TextBox
      If (ShowDuplLancFiltro()) Then
        Call LetFocus(lvwBordero(1).hWnd)
      End If

    Case 4              '// Gravar Borderô
      Call GravarBordero

  End Select
End Sub

' EVENT.....: Form_Load
' Objetivo..: Configura o formulário para sua abertura.
' -----------------------------------------------------
Private Sub Form_Load()

  SetPtr vbHourglass

  CenterForm Me

  '// Carregando os valores da caixa de combinação de tipo
  '// Utiliza as mesmas opções do cadastro de Contas Fixas

  Call LoadResOptions(1028, cboBordero(0), True, 0)
  Call LoadResOptions(1000, cboBordero(1), True, 0)

  cboOrigem.Text = cboOrigem.List(0)

  imgBordero.ImageHeight = 16
  imgBordero.ImageWidth = 16
  imgBordero.UseMaskColor = True
  imgBordero.MaskColor = vbWhite
  imgBordero.ListImages.add 1, "duplicatas", LoadResBitmap(IDB_DUPL)
  imgBordero.ListImages.add 2, "lançamentos", LoadResBitmap(IDB_LANC)

  lvwBordero(0).ColumnHeaders.add 1, NUL, "Código", 600, lvwColumnLeft
  lvwBordero(0).ColumnHeaders.add 2, NUL, "Parcela", 480, lvwColumnRight
  lvwBordero(0).ColumnHeaders.add 3, NUL, "Empresa", 1440, lvwColumnLeft
  lvwBordero(0).ColumnHeaders.add 4, NUL, "Emissão", 840, lvwColumnCenter
  lvwBordero(0).ColumnHeaders.add 5, NUL, "Vencimento", 840, lvwColumnCenter
  lvwBordero(0).ColumnHeaders.add 6, NUL, "Valor", 1440, lvwColumnRight
  lvwBordero(0).ColumnHeaders.add 7, NUL, "Origem", 300, lvwColumnLeft

  lvwBordero(1).ColumnHeaders.add 1, NUL, "Código", 600, lvwColumnLeft
  lvwBordero(1).ColumnHeaders.add 2, NUL, "Parcela", 480, lvwColumnRight
  lvwBordero(1).ColumnHeaders.add 3, NUL, "Empresa", 1440, lvwColumnLeft
  lvwBordero(1).ColumnHeaders.add 4, NUL, "Emissão", 840, lvwColumnCenter
  lvwBordero(1).ColumnHeaders.add 5, NUL, "Vencimento", 840, lvwColumnCenter
  lvwBordero(1).ColumnHeaders.add 6, NUL, "Valor", 1440, lvwColumnRight
  lvwBordero(1).ColumnHeaders.add 7, NUL, "Origem", 300, lvwColumnLeft

  lvwBordero(0).SmallIcons = imgBordero
  lvwBordero(1).SmallIcons = imgBordero

  Set Me.MouseIcon = LoadResCursor(IDC_DUPL) '// Carrega o cursor personalizado

  '// Procurando nos registros de duplicatas/lançamentos qual o próximo número para
  '// o borderô atual

  txtBordero(0).Text = ProximoNumero("Borderô", cboOrigem.Text, NUL)

  SetPtr vbDefault
  
'  Dim i As Long
'  For i = lvwBordero(0).ColumnHeaders.Count To 1 Step -1
'    lvwBordero(0).ColumnHeaders.Remove i
'  Next i
'
'  For i = lvwBordero(1).ColumnHeaders.Count To 1 Step -1
'    lvwBordero(1).ColumnHeaders.Remove i
'  Next i
'
'  lvwBordero(0).ColumnHeaders.Add 1, NUL, "Código", 600, lvwColumnLeft
'  lvwBordero(0).ColumnHeaders.Add 2, NUL, "Parcela", 480, lvwColumnRight
'  lvwBordero(0).ColumnHeaders.Add 3, NUL, "Empresa", 1440, lvwColumnLeft
'  lvwBordero(0).ColumnHeaders.Add 4, NUL, "Emissão", 840, lvwColumnCenter
'  lvwBordero(0).ColumnHeaders.Add 5, NUL, "Vencimento", 840, lvwColumnCenter
'  lvwBordero(0).ColumnHeaders.Add 6, NUL, "Valor", 1440, lvwColumnRight
'
'  lvwBordero(1).ColumnHeaders.Add 1, NUL, "Código", 600, lvwColumnLeft
'  lvwBordero(1).ColumnHeaders.Add 2, NUL, "Parcela", 480, lvwColumnRight
'  lvwBordero(1).ColumnHeaders.Add 3, NUL, "Empresa", 1440, lvwColumnLeft
'  lvwBordero(1).ColumnHeaders.Add 4, NUL, "Emissão", 840, lvwColumnCenter
'  lvwBordero(1).ColumnHeaders.Add 5, NUL, "Vencimento", 840, lvwColumnCenter
'  lvwBordero(1).ColumnHeaders.Add 6, NUL, "Valor", 1440, lvwColumnRight
  
End Sub

' EVENT.....: Form_QueryUnload
' Objetivo..: Verifica se o formulário pode ser fechado.
' ---------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If (Not cmdBordero(0).Enabled) Then   '// Impressão em curso
    SendKeysEx Chr$(27)                 '// Cancela a impressão atual
    Cancel = True
    DoEvents
  End If
End Sub

' EVENT.....: Form_Unload
' Objetivo..: Encerra a referência ao formulário.
' ---------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
  Set frmBordero = Nothing
End Sub

' EVENT.....: lvwBordero_DblClick
' Objetivo..: Move uma duplicata para o primeiro ListView.
' ---------------------------------------------------------------------
Private Sub lvwBordero_DblClick(Index As Integer)
  AddDuplicatasLancamentos
End Sub

' EVENT.....: lvwBordero_GotFocus
' Objetivo..: Exibe mensagens de ajuda na barra de status.
' -----------------------------------------------------------------------
Private Sub lvwBordero_GotFocus(Index As Integer)
  If (Index = ZERO) Then
    MsgBar cboOrigem.Text & " deste borderô. <PgDn> filtra " & cboOrigem.Text & " em aberto"
  End If
End Sub

' EVENT.....: lvwBordero_KeyDown
' Objetivo..: Abre a janela de filtro de duplicatas/lançamentos em aberto.
' -----------------------------------------------------------------------
Private Sub lvwBordero_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim nCount As Long

  If ((Index = ZERO) And (Shift = ZERO)) Then
    Select Case (KeyCode)
      Case vbKeyPageDown
        Call cmdBordero_Click(3)        '// Executa o evento do botão diretamente

      Case vbKeyDelete
        If (Not IsNothing(lvwBordero(0).SelectedItem)) Then
          lvwBordero(0).ListItems.Remove lvwBordero(0).SelectedItem.Index
          SomaBordero
        End If
    End Select
  ElseIf ((Index = UM) And (Shift = ZERO)) Then
    Select Case (KeyCode)
      Case vbKeyReturn: AddDuplicatasLancamentos
    End Select
  End If

End Sub

' EVENT.....: lvwBordero_MouseMove
' Objetivo..: Inicia a operação de "dragagem" de duplicatas/lançamentos para o
'             primeiro ListView.
' ---------------------------------------------------------------------
Private Sub lvwBordero_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ptCur As POINT                '// Posição atual do Mouse
Dim rc    As RECT                 '// Retângulo das janelas

  If (nDragFlag) Then

    Call GetCursorPos(ptCur)                    '// Obtém a posição do cursor na tela
    Call GetWindowRect(lvwBordero(1).hWnd, rc)  '// Obtém a posição do ListView

    If (PtInRect(rc, ptCur)) Then               '// Se o mouse estiver nesta área
      Me.MousePointer = vbCustom
    Else
      Call GetWindowRect(lvwBordero(0).hWnd, rc) '// Posição do primeiro ListView

      If (PtInRect(rc, ptCur)) Then             '// Se o mouse estiver nesta área
        Me.MousePointer = vbCustom
      Else
        Me.MousePointer = vbNoDrop              '// Não é possível soltar neste local
      End If
    End If

  ElseIf ((Index = 1) And (Button = vbLeftButton) And (Shift = 0) And (nDragFlag = False)) Then
    If (Not IsNothing(lvwBordero(1).SelectedItem)) Then
      nDragFlag = True
      Call SetCapture(lvwBordero(1).hWnd)          '// Captura os eventos de mouse
      Me.MousePointer = vbCustom
    End If
  End If
End Sub

' EVENT.....: lvwBordero_MouseUp
' Objetivo..: Finaliza a operação de Drag & Drop se possível.
' ---------------------------------------------------------------------
Private Sub lvwBordero_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim rcLvw As RECT           '// Localização do controle
Dim ptCur As POINT          '// Localização do mouse

  If ((Index = 1) And (nDragFlag)) Then

    Call GetCursorPos(ptCur)                      '// Obtém a posição do Cursor
    Call GetWindowRect(lvwBordero(0).hWnd, rcLvw) '// Obtém a posição do Controle

    If (PtInRect(rcLvw, ptCur)) Then
      AddDuplicatasLancamentos         '// Adiciona as duplicatas/lançamentos selecionadas
    End If
    Call ReleaseCapture                 '// Libera a captura do mouse
    nDragFlag = False                   '// Reseta o Flag da operação
    Me.MousePointer = vbDefault         '// Retorna o cursor padrão

  End If

End Sub

' EVENT.....: txtBordero_GotFocus
' Objetivo..: Exibe mensagens de descrição na barra de status do programa
' -----------------------------------------------------------------------
Private Sub txtBordero_GotFocus(Index As Integer)
  Selecione txtBordero(Index)
  MsgBar "Número do Borderô"
End Sub

Private Sub txtBordero_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Shift = 0 And KeyCode = vbKeyPageDown Then
    
    If Index = 2 Then PCampo "Bancos", "Bancos", pbCampo, txtBordero(2), "Banco"
    
  End If
End Sub

' EVENT.....: txtBordero_KeyPress
' Objetivo..: Mapear as teclas digitadas pelo usuário sobre os
'             campos da janela.
' ---------------------------------------------------------------------
Private Sub txtBordero_KeyPress(Index As Integer, KeyAscii As Integer)
Dim iSelStart As Integer

  iSelStart = txtBordero(Index).SelStart

  Select Case (Index)
    Case 0: SetMascara KeyAscii, iSelStart, "######"
    Case 1: KeyAscii = ZERO           '// Não é permitido entrar informações
    Case 2: SetMascara KeyAscii, iSelStart, fMask("Bancos", "Banco")
  End Select

End Sub

' EVENT.....: txtBordero_LostFocus
' Objetivo..: Seleciona as duplicatas ou lançamentos cujo borderô
'             é o selecionado.
' -----------------------------------------------------------------------
Private Sub txtBordero_LostFocus(Index As Integer)
Dim sExp As String      '// String de seleção de dados
Dim nCod As Long        '// Código do borderô
Dim sPR  As String      '// Pagamento ou Recebimento
Dim rst  As Object   '// Recordset com os dados do borderô

  If (Index = ZERO) Then
    nCod = CLngDef(txtBordero(0).Text)
    If (nCod) Then
      SetPtr vbHourglass

      sPR = IIf((GetItemData(cboBordero(0)) = 1), "P", "R")

      If cboOrigem.Text = "Duplicatas" Then
        Dim d As New CDuplicata
        sExp = d.SelecionaBordero(nCod, sPR)
        Set d = Nothing
      Else
        Dim l As New CLancamento
        sExp = l.SelecionaBordero(nCod, sPR)
        Set l = Nothing
      End If
      
      
      lvwBordero(1).ListItems.Clear
      
      Call ListViewAddItem(lvwBordero(1), sExp, "duplicatas")

      sExp = wsprintf("SELECT DISTINCT [Situação], Banco FROM " & cboOrigem.Text & _
                      " WHERE [Borderô] = %l AND PagRec = '%s';", nCod, sPR)
      If (AbreRecordset(rst, sExp, dbOpenSnapshot) = WL_OK) Then
        cboBordero(1).Text = GetValue(rst, "Situação", "Normal")
        txtBordero(2).Text = GetValue(rst, "Banco", ZERO)
      End If
      FechaRecordset rst
      SomaBordero

      SetPtr vbDefault

    End If
  End If

End Sub

' FUNCTION..: ShowDuplFiltro
' Objetivo..: Exibe a janela de filtro de duplicatas/lançamentos para o
'             usuário selecionar um conjunto de registros.
'             Carrega os dados e exíbe-os no segundo controle
'             ListView.
' Retorna...: True se obtiver sucesso, False se não.
' ---------------------------------------------------------------------
Private Function ShowDuplLancFiltro() As Boolean
Dim fdlg As fdlgDuplicatas      '// Instancia do diálogo de filtro
Dim sTmp As String              '// Instrução de seleção de dados

  Set fdlg = New fdlgDuplicatas
  Load fdlg

  fdlg.Tipo = IIf((GetItemData(cboBordero(0)) = 1), "P", "R")
  fdlg.gstrTabela = cboOrigem.Text
  fdlg.Show vbModal

  DoEvents

  If (fdlg.Cancel) Then
    sTmp = NUL
  Else
    sTmp = fdlg.Expressao
  End If

  Unload fdlg
  Set fdlg = Nothing

  If (Len(sTmp)) Then
    lvwBordero(1).ListItems.Clear
    Call ListViewAddItem(lvwBordero(1), sTmp, "duplicatas")

    If (lvwBordero(1).ListItems.Count = ZERO) Then
      MsgFunc LoadResString(IDS_RECORDNOTFOUND)
    End If

    ShowDuplLancFiltro = True
  End If

End Function

' SUB.......: AddDuplicatasLancamentos
' Objetivo..: Adiciona as duplicatas/Lançamentos selecionadas no segundo
'             ListView ao primeiro ListView que contém
'             as duplicatas que fazem parte do boleto
' ---------------------------------------------------------------------
Private Sub AddDuplicatasLancamentos()
Dim iItem As Long         '// Índice do ítem selecionado
Dim nItem As Long         '// Índice do ítem adicionado na lista de duplicatas
Dim nCont As Long         '// Contador dos subítens

  nItem = lvwBordero(0).ListItems.Count

  For iItem = lvwBordero(1).ListItems.Count To 1 Step -1
    If (lvwBordero(1).ListItems(iItem).Selected) Then
      nItem = nItem + 1
      lvwBordero(0).ListItems.add nItem, , lvwBordero(1).ListItems(iItem).Text, , "duplicatas"

      For nCont = 1 To lvwBordero(0).ColumnHeaders.Count - 1
          lvwBordero(0).ListItems(nItem).SubItems(nCont) = lvwBordero(1).ListItems(iItem).SubItems(nCont)
      Next

      lvwBordero(1).ListItems.Remove iItem
    End If
  Next
  SomaBordero

End Sub

' SUB.......: SomaBordero
' Objetivo..: Soma as duplicatas/lançamentos constantes no borderô.
' ---------------------------------------------------------------------
Private Sub SomaBordero()
Dim cValor As Currency
Dim n      As Long

  For n = 1 To lvwBordero(0).ListItems.Count
    cValor = cValor + CMoedaFormatoAmericano(lvwBordero(0).ListItems(n).SubItems(lvwBordero(0).ColumnHeaders.Count - 2))
  Next
  txtBordero(1).Text = Format$(cValor, FCURRENCY)

End Sub

' FUNCTION..: GravarBordero
' Objetivo..: Grava o número do borderô nos registros de duplicatas/lançamentos
'             selecionados pelo usuário.
' Retorna...: True se obtiver sucesso, False se não.
' ---------------------------------------------------------------------
Private Function GravarBordero() As Boolean
Dim nItems   As Long          '// Conta os ítens selecionados para um borderô
Dim sUpdt    As String        '// Instrução de atualização
Dim nBordero As Long          '// Código do Borderô atual
Dim nNota    As Long          '// Número da nota atual
Dim nParcela As Long          '// Número da parcela atual
Dim sEmpresa As String        '// Nome da empresa
Dim sPagRec  As String        '// Pagar ou Receber
Dim nBanco   As Long          '// Código do Banco
Dim sTipo    As String        '// Situação do Borderô
Dim sData    As String        '// Data de pagamento

  SetPtr vbHourglass

  nBordero = CLngDef(txtBordero(0).Text)

  If (nBordero = ZERO) Then
    MsgFunc "Número de Borderô inválido!", vbExclamation
    SetPtr vbDefault
    Exit Function
  End If

  nBanco = CLngDef(txtBordero(2).Text)
  If (nBanco = ZERO) Then
    MsgFunc "Você deve informar um Banco para gravar o Borderô"
    SetPtr vbDefault
    Exit Function
  Else
    '// Verificando se o código informado é um código válido

    If (ConfRelation(txtBordero(2).Text, "Banco = %s", "Bancos", CRN_ON_TABLE Or CRN_NO_QUERY)) Then
      SetPtr vbDefault
      Exit Function
    End If
  End If

  sTipo = cboBordero(1).Text            '// Situação da Duplicata
  sPagRec = IIf((GetItemData(cboBordero(0)) = 1), "P", "R")
  
  
  '// Obtendo a Data de Pagamento do usuário
  Do
    sData = InputBox(vbCrLf & "Insira a 'Data de Pagamento' para atualizar em todos os registros de " & cboOrigem.Text & " deste Borderô.", "Atualizar Data de Pagamento das Duplicatas")
    If Len(sData) And Not EData(sData) Then
      MsgFunc "A 'Data de Pagamento' inserida não é válida!"
    End If
  Loop Until EData(sData) Or Len(sData) = 0
  

  '// Inicialmente exclui todas as referências a este número de borderô
  '// na tabela

  sUpdt = wsprintf("UPDATE " & "Duplicatas" & " SET Borderô = 0 WHERE Borderô = %l " & _
                  "AND PagRec = '%s';", nBordero, sPagRec)
  ExecuteSQL sUpdt
  sUpdt = wsprintf("UPDATE " & "Lançamentos" & " SET Borderô = 0 WHERE Borderô = %l " & "AND PagRec = '%s';", nBordero, sPagRec)
  ExecuteSQL sUpdt
  
  '// Determina a string da Data de Pagamento para atualizar as Duplicatas do Borderô
  If EData(sData) Then
    sData = "Pagamento = " & InverteData(sData, True) & ", "
  Else
    sData = ""
  End If

  If (lvwBordero(0).ListItems.Count) Then
    For nItems = 1 To lvwBordero(0).ListItems.Count
      If lvwBordero(0).ListItems(nItems).SubItems(6) = "D" Then    'Duplicatas
        sUpdt = "UPDATE Duplicatas SET Borderô = %l, " & sData & "Banco = %l, Situação = '%s' " & _
                "WHERE Nota = %l AND Parcela = %l AND Empresa = '%s' AND PagRec = '%s';"
                
        nNota = CLngDef(lvwBordero(0).ListItems(nItems).Text)
        nParcela = CLngDef(lvwBordero(0).ListItems(nItems).SubItems(1))
        sEmpresa = lvwBordero(0).ListItems(nItems).SubItems(2)
  
        wvsprintf sUpdt, sUpdt, nBordero, nBanco, sTipo, nNota, nParcela, sEmpresa, sPagRec
        ExecuteSQL sUpdt
      ElseIf lvwBordero(0).ListItems(nItems).SubItems(6) = "L" Then   'Lançamentos
        sUpdt = "UPDATE Lançamentos SET Borderô = %l, " & sData & "Banco = %l, Situação = '%s' " & _
                "WHERE Código = %l AND PagRec = '%s';"
                
        nNota = CLngDef(lvwBordero(0).ListItems(nItems).Text)
  
        wvsprintf sUpdt, sUpdt, nBordero, nBanco, sTipo, nNota, sPagRec
        ExecuteSQL sUpdt
        
      End If

    Next
  End If
  SetPtr vbDefault
  GravarBordero = True

End Function

' FUNCTION..: TempDupliLanc
' Objetivo..: Cria uma tabela auxiliar para impressão dos ítens do
'             borderô.
' Argumento.: [rsDupls]: Recordset que receberá a tabela auxiliar.
' Retorna...: True se obtiver sucesso, False se não.
' ---------------------------------------------------------------------
Private Function TempDupliLanc(rsDupls As Object) As Boolean
    Dim fsDupls(3) As FieldStruct
    
    ' 08/04/2019 - FBMI:63 - Yuji F. - Alteração do tamanho da coluna Código,
    'para aceitar lançamentos com + de 9 caracteres
    AppendVar fsDupls(0), "Código", dbText, 25     '// Nota + Parcela
    AppendVar fsDupls(1), "Razão", dbText, 15     '// Fantasia da empresa
    AppendVar fsDupls(2), "Vencto", dbDate        '// Data de Vencimento da Duplicata
    AppendVar fsDupls(3), "Valor", dbCurrency     '// Valor da duplicata
    
    ' 08/04/2019 - FBMI:63 - Yuji F. - Mudança de método para criar uma tabela temporária
    If (CrieTempTable(rsDupls, fsDupls(), "#Bordero" & UserName())) Then
        TempDupliLanc = True
    Else
        MsgFunc LoadResString(174), vbExclamation
    End If

End Function

' SUB.......: ImprimirBordero
' Objetivo..: Imprime o Borderô na tela ou impressora.
' Argumento.: [nDst]: Destino da impressão.
' ---------------------------------------------------------------------
Private Sub ImprimirBordero(nDst As PrintDestinoEnum)
Dim rstAux As Object           '// Recordset da tabela auxiliar
Dim nItens As Long                '// Ítens da ListView

'  If (Not GravarBordero()) Then
'    MsgBar MsgBoxCaption
'    SetPtr vbDefault
'    Exit Sub        '// Grava o borderô incondicionalmente antes de gerar a impressão
'  End If

  On Error GoTo ImprimirBordero_Erro
  SetPtr vbHourglass

  Call InKey(vbKeyEscape)         '// Limpa o buffer do teclado

  SimpleMsgBar "Gerando relatório, aguarde..."

  If (lvwBordero(0).ListItems.Count = ZERO) Then
    MsgFunc "Não há nenhum registro selecionado"
    MsgBar MsgBoxCaption
    SetPtr vbDefault
    Exit Sub
  End If

  If (TempDupliLanc(rstAux)) Then
    Call InitTrans
    For nItens = 1 To lvwBordero(0).ListItems.Count
      DoEvents
      If (InKey(vbKeyEscape)) Then GoTo ImprimirBordero_Erro:

      rstAux.AddNew
      If lvwBordero(0).ListItems(nItens).SubItems(6) = "D" Then
        rstAux("Código").value = lvwBordero(0).ListItems(nItens).Text & _
                                 "-" & lvwBordero(0).ListItems(nItens).SubItems(1)
        rstAux("Razão").value = lvwBordero(0).ListItems(nItens).SubItems(2)
        rstAux("Vencto").value = CDateDef(lvwBordero(0).ListItems(nItens).SubItems(4), Null)
        rstAux("Valor").value = CMoedaFormatoAmericano(lvwBordero(0).ListItems(nItens).SubItems(5))
        
      Else
        rstAux("Código").value = lvwBordero(0).ListItems(nItens).Text
        rstAux("Razão").value = lvwBordero(0).ListItems(nItens).SubItems(2)
        rstAux("Vencto").value = CDateDef(lvwBordero(0).ListItems(nItens).SubItems(4), Null)
        rstAux("Valor").value = CMoedaFormatoAmericano(lvwBordero(0).ListItems(nItens).SubItems(5))
      End If
      rstAux.update
    Next
    Call UpdateTrans(FORCE_WRITE)

    Call FormataRelatorio(rstAux, nDst)
  End If

ImprimirBordero_Erro:
  If (err().Number) Then
    #If (DebugInfo) Then
      MsgErro wsprintf("Erro: %l\n%s\nImprimirBoleto", err.Number, err.Description)
    #Else
      DAOErros NUL
    #End If
    Call CancelTrans
  End If
  DeleteAux rstAux, NUL
  SetPtr vbDefault
  MsgBar MsgBoxCaption
End Sub

' SUB.......: FormataRelatorio
' Objetivo..: Configura o objeto KeybReport e imprime o relatório.
' Argumentos: [rstDados]: Recordset com os dados que devem ser impressos.
'             [nDest   ]: Destino da impressão.
' ---------------------------------------------------------------------
Private Sub FormataRelatorio(rstDados As Object, nDest As Long)
Dim wrBordero As KeybReport
Dim rstBanco  As Object          '// Dados do Banco selecionado


  If (CreateReport(wrBordero, nDest, "Borderô")) Then
    Set wrBordero.Recordset = rstDados

    wrBordero.AddGrupo "1"
    wrBordero.FontSize = 10

    With wrBordero.Grupo(1)
      .AddSecao scHeader, 8

      If (AbreRecordset(rstBanco, "SELECT * FROM Bancos WHERE Banco = " & _
                                   txtBordero(2).Text, dbOpenSnapshot) <> WL_OK) Then
        wrBordero.EndPrint
        Set wrBordero = Nothing
        Exit Sub
      End If

      With .Header.Linha(1)       '// Nome do Banco
        .AddCampo , wrCSFixedText, "Ao Banco:", , 20
        .AddCampo , wrCSFixedText, GetValue(rstBanco, "Nome", NUL)
        .Campo(1).FontStyle = wrFSBold
      End With

      With .Header.Linha(2)       '// Agência e Conta
        .AddCampo , wrCSFixedText, "Agência:", , 20
        .AddCampo , wrCSFixedText, GetValue(rstBanco, "Agência", NUL), , 80
        .AddCampo , wrCSFixedText, "Conta:", , 15
        .AddCampo , wrCSFixedText, GetValue(rstBanco, "Conta", NUL)
        .Campo(1).FontStyle = wrFSBold
        .Campo(3).FontStyle = wrFSBold
      End With

      With .Header.Linha(3)       '// Bairro do Banco
        .AddCampo , wrCSFixedText, "Bairro:", , 20
        .AddCampo , wrCSFixedText, GetValue(rstBanco, "Bairro", NUL)
        .Campo(1).FontStyle = wrFSBold
      End With

      With .Header.Linha(4)       '// Número do Borderô
        .AddCampo , wrCSFixedText, "Borderô:", , 20
        .AddCampo , wrCSFixedText, txtBordero(0).Text
        .Campo(1).FontStyle = wrFSBold
      End With

      With .Header.Linha(5)       '// Situação das Duplicatas/Lançamentos neste Borderô
        .AddCampo , wrCSFixedText, "Situação:", , 20
        .AddCampo , wrCSFixedText, cboBordero(1).Text
        .Campo(1).FontStyle = wrFSBold
      End With

      With .Header.Linha(6)       '// Linha de Texto
        If (GetItemData(cboBordero(0)) = UM) Then        '// 1 == A Pagar
          .AddCampo , wrCSFixedText, LoadResString(260)
        Else
          .AddCampo , wrCSFixedText, LoadResString(259)
        End If
        .Campo(1).MultiLine = True
      End With

      .Header.Linha(7).DrawBorder = wrDBBottomBorder
      wrBordero.FontStyle = wrFSBold Or wrFSItalic

      With .Header.Linha(8)       '// Cabeçalho das colunas da lista de duplicatas
        .AddCampo , wrCSFixedText, "Dupl/Lanc", , 22

        If (GetItemData(cboBordero(0)) = UM) Then     '// 1 == A Pagar
          .AddCampo , wrCSFixedText, "Sacador", , 68
        Else
          .AddCampo , wrCSFixedText, "Sacado", , 68
        End If

        .AddCampo , wrCSFixedText, "Vencimento", wrTACentro, 40
        .AddCampo , wrCSFixedText, "Valor", wrTADireito
        .DrawBorder = wrDBBottomBorder
        .BorderStyle = wrDot
      End With

      wrBordero.FontStyle = wrFSNormal

      .AddSecao scDetalhe, 1

      With .Detalhe.Linha(1)
        .AddCampo , , "Código", wrTADireito, 22
        .AddCampo , , "Razão", , 68
        .AddCampo , , "Vencto", wrTACentro, 40
        .AddCampo , , "Valor", wrTADireito
        .Campo(4).Formato = FMOEDA
      End With

      .AddSecao scFooter, 10

      With .Footer.Linha(1)           '// Valor Total do Borderô
        .DrawBorder = wrDBTopBorder
        .AddCampo , wrCSFixedText, "Total:", , 40, 100
        .AddCampo , wrCSFixedText, txtBordero(1).Text, wrTADireito
        .Campo(1).FontStyle = wrFSBold
      End With

      With .Footer.Linha(2)           '// Extenso do valor
        ' 08/04/2019 - FBMI:63 - Yuji F. - Mudança de método que estava alterando o valor a ser convertido para extenso
        .AddCampo , wrCSFixedText, KeybUCase(KeybExtenso(CCurDef(txtBordero(1).Text)), PorPalavra)
        .Campo(1).MultiLine = True
      End With

      With .Footer.Linha(3)           '// Quantidade de duplicatas/lançamentos
        .AddCampo , wrCSFixedText, "Quantidade de Títulos:", , 40
        .AddCampo , wrCSFixedText, CStr(lvwBordero(0).ListItems.Count), , 15
        .Campo(1).FontStyle = wrFSBold
      End With

      With .Footer.Linha(5)           '// Data da impressão
        .AddCampo , wrCSFixedText, "Data:", , 15
        .AddCampo , wrCSData
        .Campo(1).FontStyle = wrFSBold
      End With

      With .Footer.Linha(6)
        .AddCampo , wrCSFixedText, "Atenciosamente", , , wrBordero.ClientWidth / 2
      End With

      With .Footer.Linha(10)
        .AddCampo , wrCSFixedText, NomeDonaSistema(), wrTACentro, , wrBordero.ClientWidth / 2
        .DrawBorder = wrDBBottomBorder
      End With
    End With

    wrBordero.BeginPrint gTipoDB
    wrBordero.EndPrint
  End If
  FechaRecordset rstBanco
  Set wrBordero = Nothing

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
