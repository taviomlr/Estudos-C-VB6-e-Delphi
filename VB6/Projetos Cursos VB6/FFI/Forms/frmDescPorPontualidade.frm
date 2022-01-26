VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComctl.ocx"
Begin VB.Form frmDescPorPontualidade 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconto por Pontualidade"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   10485
   Begin VB.Frame Frame3 
      Height          =   3495
      Left            =   9000
      TabIndex        =   29
      Top             =   30
      Width           =   1485
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1050
         Width           =   1215
      End
      Begin VB.CommandButton cmdExecutar 
         Caption         =   "Executar"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   1215
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   1215
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   480
         Top             =   2730
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDescPorPontualidade.frx":0000
               Key             =   "marcado"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDescPorPontualidade.frx":015A
               Key             =   "Desmarcado"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Atualização do Desconto por Pontualidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   30
      TabIndex        =   28
      Top             =   2160
      Width           =   8955
      Begin VB.TextBox txtPer 
         Height          =   315
         Left            =   2250
         TabIndex        =   12
         Tag             =   "Baixas"
         Top             =   930
         Width           =   1815
      End
      Begin VB.TextBox txtVlrFix 
         Height          =   315
         Left            =   2250
         TabIndex        =   10
         Tag             =   "Baixas"
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optPer 
         Caption         =   "Desconto por Percentual"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   2205
      End
      Begin VB.OptionButton optVlrFix 
         Caption         =   "Desconto por Valor Fixo"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   630
         Width           =   2115
      End
      Begin VB.OptionButton optCadCli 
         Caption         =   "Percentual de desconto configurado no cadastro do Cliente"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   5145
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Duplicatas/Lançamentos a Receber"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   30
      TabIndex        =   17
      Top             =   30
      Width           =   8955
      Begin VB.ComboBox cboTipDoc 
         Height          =   315
         ItemData        =   "frmDescPorPontualidade.frx":05AC
         Left            =   1410
         List            =   "frmDescPorPontualidade.frx":05B6
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1815
      End
      Begin VB.TextBox txtBanco 
         Height          =   315
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Baixas"
         Top             =   1710
         Width           =   1815
      End
      Begin VB.TextBox txtVenFim 
         Height          =   315
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Baixas"
         Top             =   1380
         Width           =   1005
      End
      Begin VB.TextBox txtVenIni 
         Height          =   315
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Baixas"
         Top             =   1380
         Width           =   1005
      End
      Begin VB.ComboBox cboTipReg 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   1815
      End
      Begin VB.TextBox txtEmp 
         Height          =   315
         Left            =   1410
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "Baixas"
         Top             =   990
         Width           =   1815
      End
      Begin VB.TextBox txtNroIni 
         Height          =   315
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Baixas"
         Top             =   1350
         Width           =   1005
      End
      Begin VB.TextBox txtNroFim 
         Height          =   315
         Left            =   2610
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Baixas"
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label lblBanDes 
         Caption         =   "lblDescricao"
         Height          =   255
         Left            =   3300
         TabIndex        =   27
         Top             =   1740
         Width           =   5445
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   26
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   25
         Top             =   1770
         Width           =   1215
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "a"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   1
         Left            =   5850
         TabIndex        =   24
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label lblCalculo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         Height          =   195
         Index           =   6
         Left            =   3855
         TabIndex        =   23
         Top             =   1410
         Width           =   885
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Registro:"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   22
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   21
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   20
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label lblEmpDes 
         Caption         =   "lblDescricao"
         Height          =   255
         Left            =   3270
         TabIndex        =   19
         Top             =   1020
         Width           =   5505
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "a"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   8
         Left            =   2460
         TabIndex        =   18
         Top             =   1410
         Width           =   90
      End
   End
   Begin MSComctlLib.ListView lvwDupLan 
      Height          =   2145
      Left            =   30
      TabIndex        =   13
      Top             =   3510
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   3784
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmDescPorPontualidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MARCADO = 1
Private Const DESMARCADO = 2

Private mTipPagRec As String

Private Function ValidaParametrosExecuta() As Boolean
Dim i As Long
    
    ValidaParametrosExecuta = False
    
    If optVlrFix.Value Then
        If txtVlrFix.Text = Empty Then
            MsgBox "O valor deve ser informado."
            txtVlrFix.SetFocus
            Exit Function
        End If
        
        If CDblDef(txtVlrFix.Text, 0) < 0 Then
            MsgBox "O VALOR deve ser maior que 0."
            txtVlrFix.SetFocus
            Exit Function
        End If
    End If
    
    If optPer.Value Then
        If txtPer.Text = Empty Then
            MsgBox "O percentual deve ser informado."
            txtPer.SetFocus
            Exit Function
        End If
        
        If CDblDef(txtPer.Text, 0) <= 0 Or CDblDef(txtPer.Text, 0) > 100 Then
            MsgBox "O percentual deve ser maior que 0 e menor que 100."
            optPer.SetFocus
            Exit Function
        End If
    End If
    
    'faco o loop para ver se tem algum registro marcado
    For i = 1 To lvwDupLan.ListItems.Count
        If lvwDupLan.ListItems(i).SmallIcon = MARCADO Then
            Exit For
        End If
    Next
    If i > lvwDupLan.ListItems.Count Then
        MsgBox "Nenhum registro selecionado para atualização."
        Exit Function
    End If
    
    ValidaParametrosExecuta = True
End Function

Private Function ValidaParametroscConsulta() As Boolean
    ValidaParametroscConsulta = False
    
    If txtEmp.Text <> Empty And lblEmpDes.Caption = Empty Then
        MsgBox "A empresa informada não existe."
        txtEmp.SetFocus
        Exit Function
    End If
    
    If txtNroIni.Text <> Empty And txtNroFim.Text <> Empty Then
        If CLngDef(txtNroIni.Text, 0) > CLngDef(txtNroFim.Text, 0) Then
            MsgBox "O número inicial deve ser menor ou igual ao número final."
            txtNroIni.SetFocus
            Exit Function
        End If
    End If
    
    If txtVenIni.Text <> Empty And CDateDef(txtVenIni.Text, 0) = 0 Then
        MsgBox "Data inicial de vencimento inválida."
        txtVenIni.SetFocus
        Exit Function
    End If
    
    If txtVenFim.Text <> Empty And CDateDef(txtVenFim.Text, 0) = 0 Then
        MsgBox "Data final de vencimento inválida."
        txtVenFim.SetFocus
        Exit Function
    End If
    
    If txtVenIni.Text <> Empty And txtVenFim.Text <> Empty Then
        If CDateDef(txtVenIni.Text, 0) > CDateDef(txtVenFim.Text, 0) Then
            MsgBox "A data de vencimento inicial deve ser menor ou igual a data de vencimento final."
            txtVenIni.SetFocus
            Exit Function
        End If
    End If
    
    
    If txtBanco.Text <> Empty And lblBanDes.Caption = Empty Then
        MsgBox "O banco informado não existe."
        txtBanco.SetFocus
        Exit Function
    End If
    
    ValidaParametroscConsulta = True
End Function

Private Sub cmdConsultar_Click()
    If ValidaParametroscConsulta Then
        CarregaLista False
    End If
End Sub

Private Sub CarregaLista(bSomenteCabecalho As Boolean)
    Dim rs As ADODB.Recordset
    
    If cboTipDoc.Text = "Duplicatas" Then
        If Not bSomenteCabecalho Then
            Set rs = modFinanceiro.CarregaDuplicatas(mTipPagRec, _
                        pTipReg:=cboTipReg.Text, _
                        pEmp:=txtEmp.Text, _
                        pNroIni:=CLngDef(txtNroIni.Text, 0), _
                        pNroFim:=CLngDef(txtNroFim.Text, 0), _
                        pVctIni:=CDateDef(txtVenIni.Text, 0), _
                        pVctFim:=CDateDef(txtVenFim.Text, 0), _
                        pBan:=CLngDef(txtBanco.Text, 0), _
                        pSomenteNaoPagas:=True)
        End If
        
        CarregaListaDuplicatas rs
        
    Else
        
        If Not bSomenteCabecalho Then
            Set rs = modFinanceiro.CarregaLancamentos(mTipPagRec, _
                        pTipReg:=cboTipReg.Text, _
                        pEmp:=txtEmp.Text, _
                        pNroIni:=CLngDef(txtNroIni.Text, 0), _
                        pNroFim:=CLngDef(txtNroFim.Text, 0), _
                        pVctIni:=CDateDef(txtVenIni.Text, 0), _
                        pVctFim:=CDateDef(txtVenFim.Text, 0), _
                        pBan:=CLngDef(txtBanco.Text, 0), _
                        pSomenteNaoPagas:=True)
        End If
        
        CarregaListaLancamentos rs
    End If
    
    Set rs = Nothing
End Sub

Private Sub CarregaListaLancamentos(pRs As ADODB.Recordset)

    Screen.MousePointer = vbHourglass
    
    lvwDupLan.ColumnHeaders.clear
    lvwDupLan.ListItems.clear
          
    'utilizo o tag com o nome do campo para facilitar o preeenchimento
    lvwDupLan.ColumnHeaders.add(Key:="Codigo", Text:="Código", Width:=1100).Tag = "Código"
    lvwDupLan.ColumnHeaders.add(Key:="Descricao", Text:="Descrição", Width:=2000).Tag = "Descrição"
    lvwDupLan.ColumnHeaders.add(Key:="Tipo", Text:="Tipo").Tag = "Tipo"
    lvwDupLan.ColumnHeaders.add(Key:="Empresa", Text:="Empresa").Tag = "Empresa"
    lvwDupLan.ColumnHeaders.add(Key:="ValorOriginal", Text:="Valor", Width:=1350, Alignment:=AlignmentConstants.vbRightJustify).Tag = "Valor Original"
    lvwDupLan.ColumnHeaders.add(Key:="Vencimento", Text:="Vencimento", Width:=1100).Tag = "Vencimento"
    lvwDupLan.ColumnHeaders.add(Key:="Emissao", Text:="Emissão", Width:=1100).Tag = "Emissão"
    
    DoEvents
    
    PreencheListView pRs
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub CarregaListaDuplicatas(pRs As ADODB.Recordset)
   
    Screen.MousePointer = vbHourglass
    
    lvwDupLan.ColumnHeaders.clear
    lvwDupLan.ListItems.clear
          
    'utilizo o tag com o nome do campo para facilitar o preeenchimento
    lvwDupLan.ColumnHeaders.add(Key:="Nota", Text:="Nota", Width:=1100).Tag = "Nota"
    lvwDupLan.ColumnHeaders.add(Key:="Parcela", Text:="Parcela", Width:=800).Tag = "Parcela"
    lvwDupLan.ColumnHeaders.add(Key:="Descricao", Text:="Descrição", Width:=2000).Tag = "Descrição"
    lvwDupLan.ColumnHeaders.add(Key:="Tipo", Text:="Tipo", Width:=1100).Tag = "Tipo"
    lvwDupLan.ColumnHeaders.add(Key:="Empresa", Text:="Empresa").Tag = "Empresa"
    lvwDupLan.ColumnHeaders.add(Key:="ValorOriginal", Text:="Valor", Width:=1350, Alignment:=AlignmentConstants.vbRightJustify).Tag = "Valor Original"
    lvwDupLan.ColumnHeaders.add(Key:="Vencimento", Text:="Vencimento", Width:=1100).Tag = "Vencimento"
    lvwDupLan.ColumnHeaders.add(Key:="Emissao", Text:="Emissão", Width:=1100).Tag = "Emissão"
    
    DoEvents

    PreencheListView pRs
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub PreencheListView(pRs As ADODB.Recordset)

Dim i As Long
Dim Linha As Long

    lvwDupLan.ListItems.clear
    Linha = 0
    
If Not pRs Is Nothing And lvwDupLan.ColumnHeaders.Count > 0 Then
    Do While Not pRs.EOF
        
        Linha = Linha + 1
        lvwDupLan.ListItems.add Linha, , pRs(lvwDupLan.ColumnHeaders(1).Tag)
        
        For i = 2 To lvwDupLan.ColumnHeaders.Count
            If lvwDupLan.ColumnHeaders(i).Tag <> Empty Then
                If lvwDupLan.ColumnHeaders(i).Key = "ValorOriginal" Then
                    lvwDupLan.ListItems(Linha).SubItems(i - 1) = Format(pRs(lvwDupLan.ColumnHeaders(i).Tag), "#,##0.00")
                Else
                    If Not IsNull(pRs(lvwDupLan.ColumnHeaders(i).Tag)) Then
                        lvwDupLan.ListItems(Linha).SubItems(i - 1) = pRs(lvwDupLan.ColumnHeaders(i).Tag)
                    End If
                End If
            End If
            
            'MARCO COMO SELECIONADO
            lvwDupLan.ListItems(Linha).SmallIcon = MARCADO
        Next
        
        pRs.MoveNext
        
    Loop
End If

End Sub

Private Sub cmdExecutar_Click()
Dim i As Long
Dim vlr As Double
Dim tip As FIN_TIP_DESC_PONTUALIDADE

On Error GoTo Error_Handler

    If Not ValidaParametrosExecuta Then
        Exit Sub
    End If
                
                
    Screen.MousePointer = vbHourglass
    
    If optCadCli.Value Then
        vlr = 0
        tip = tdpCadCliente
    ElseIf optVlrFix.Value Then
        vlr = CDblDef(txtVlrFix.Text)
        tip = tdpVlrFixo
    Else
        vlr = CDblDef(txtPer.Text)
        tip = tdpPercentual
    End If
        
    'faco o loop atualizado os registros marcados
    For i = 1 To lvwDupLan.ListItems.Count
        If lvwDupLan.ListItems(i).SmallIcon = MARCADO Then
            If cboTipDoc.Text = "Duplicatas" Then
                Call AtualizaDescPorPontualidadeDuplicata(mTipPagRec, _
                        lvwDupLan.ListItems(i).SubItems(lvwDupLan.ColumnHeaders("Tipo").Index - 1), _
                        lvwDupLan.ListItems(i).SubItems(lvwDupLan.ColumnHeaders("Empresa").Index - 1), _
                        CLngDef(lvwDupLan.ListItems(i).Text, 0), _
                        CIntDef(lvwDupLan.ListItems(i).SubItems(lvwDupLan.ColumnHeaders("Parcela").Index - 1), 1), _
                        tip, vlr)
            ElseIf cboTipDoc.Text = "Lançamentos" Then
                Call AtualizaDescPorPontualidadeLancamentos(mTipPagRec, _
                        CLngDef(lvwDupLan.ListItems(i).Text, 0), _
                        tip, vlr)
            End If
        End If
    Next
    Screen.MousePointer = vbDefault
    
    MsgBox "Registros atualizados com sucesso."
    
    Exit Sub

Error_Handler:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
    mTipPagRec = "R"
    
    lblEmpDes.Caption = Empty
    lblBanDes.Caption = Empty
    
    cboTipDoc.Text = "Duplicatas"
    ComboAddItem cboTipReg, "SELECT Tipo FROM [Tipos Globais]", "Tipo"
    cboTipReg.AddItem Empty, 0
    

    'para marcar a opção do cadastro do cliente como padrão
    'ja que o vb não dispara o evento no load do form
    optCadCli.Value = True
    Call optCadCli_Click
    
    CarregaLista True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SavePosForm Me
    Set frmDescPorPontualidade = Nothing
End Sub

Private Sub lvwDupLan_DblClick()
Dim i As Long
    
    If Not lvwDupLan.SelectedItem Is Nothing Then
        i = lvwDupLan.SelectedItem.Index
        If (lvwDupLan.ListItems(i).SmallIcon = MARCADO) Then
          lvwDupLan.ListItems(i).SmallIcon = DESMARCADO
        Else
          lvwDupLan.ListItems(i).SmallIcon = MARCADO
        End If
    End If
End Sub

Private Sub optCadCli_Click()
    If optCadCli.Value Then
        txtPer.Text = Empty
        txtPer.Enabled = False
        txtVlrFix.Text = Empty
        txtVlrFix.Enabled = False
    End If
End Sub

Private Sub optPer_Click()
    If optPer.Value Then
        txtPer.Enabled = True
        txtVlrFix.Text = Empty
        txtVlrFix.Enabled = False
    End If
End Sub

Private Sub optVlrFix_Click()
    If optVlrFix.Value Then
        txtPer.Text = Empty
        txtPer.Enabled = False
        txtVlrFix.Enabled = True
    End If
End Sub

Private Sub txtBanco_Change()
    GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & txtBanco.Text, lblBanDes
End Sub

Private Sub txtBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        PCampo "Bancos", "Bancos", pbCampo, txtBanco, 0
    End If
End Sub

Private Sub txtEmp_Change()
    GetAssocValue "Select Razão, Apel from Empresas where Apel = " & Quote(txtEmp.Text, "'"), lblEmpDes
End Sub

Private Sub txtEmp_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 0 And KeyCode = vbKeyPageDown Then
        PCampo "Empresas", "Empresas", pbCampo, txtEmp, "Apel"
    End If

End Sub

Private Sub txtEmp_LostFocus()
    GetAssocValue "Select Razão, Apel from Empresas where Apel = " & Quote(txtEmp.Text, "'"), lblEmpDes, txtEmp
End Sub

Private Sub txtNroFim_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If cboTipDoc.Text = "Duplicatas" Then
            PCampo "Duplicatas", "Select * from Duplicatas where PagRec = " & Quote(mTipPagRec, "'"), pbCampo, txtNroFim, "Nota"
        Else
            PCampo "Lançamentos", "Select * from Lançamentos where PagRec = " & Quote(mTipPagRec, "'"), pbCampo, txtNroFim, "Código"
        End If
    End If
    
End Sub

Private Sub txtNroFim_KeyPress(KeyAscii As Integer)
    SetMascara KeyAscii, txtNroFim.SelStart, "######"
End Sub

Private Sub txtNroIni_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 0 And KeyCode = vbKeyPageDown Then
        If cboTipDoc.Text = "Duplicatas" Then
            PCampo "Duplicatas", "Select * from Duplicatas where PagRec = " & Quote(mTipPagRec, "'"), pbCampo, txtNroIni, "Nota"
        Else
            PCampo "Lançamentos", "Select * from Lançamentos where PagRec = " & Quote(mTipPagRec, "'"), pbCampo, txtNroIni, "Código"
        End If
    End If
  
End Sub

Private Sub txtNroIni_KeyPress(KeyAscii As Integer)
  SetMascara KeyAscii, txtNroIni.SelStart, "######"
End Sub

Private Sub txtPer_KeyPress(KeyAscii As Integer)
    DValor KeyAscii
End Sub

Private Sub txtPer_LostFocus()
    txtPer.Text = CDblDef(txtPer.Text, 0)
End Sub

Private Sub txtVenFim_KeyPress(KeyAscii As Integer)
    SetMascara KeyAscii, txtVenFim.SelStart, MASK_DATA
End Sub

Private Sub txtVenIni_KeyPress(KeyAscii As Integer)
    SetMascara KeyAscii, txtVenIni.SelStart, MASK_DATA
End Sub

Private Sub txtVlrFix_KeyPress(KeyAscii As Integer)
    DValor KeyAscii
End Sub


Private Sub txtVlrFix_LostFocus()
    txtVlrFix.Text = CDblDef(txtVlrFix.Text, 0)
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
