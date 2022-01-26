VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form fcalcGeracaoNotasFinanceiro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de Notas Fiscais a partir de Duplicatas e Lançamentos"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "calcGeracaoNotasFinanceiro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Height          =   735
      Index           =   1
      Left            =   3480
      TabIndex        =   13
      Top             =   3600
      Width           =   3375
      Begin VB.CommandButton cmdFechar 
         Caption         =   "&Fechar"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1545
      End
      Begin VB.CommandButton cmdGerar 
         Caption         =   "&Gerar"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Marcar"
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   3375
      Begin VB.CommandButton cmdDesmarcarTodos 
         Caption         =   "&Nenhum"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   1545
      End
      Begin VB.CommandButton cmdMarcarTodos 
         Caption         =   "&Todos"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione quais Duplicatas/Lançamentos deverão ser gerados Nota"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6735
      Begin VB.Frame fraServico 
         Caption         =   "Serviço"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         TabIndex        =   10
         Top             =   2280
         Width           =   6735
         Begin VB.ComboBox cboTipoRegistro 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtServico 
            Height          =   315
            Left            =   1440
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblcalcFinanc 
            AutoSize        =   -1  'True
            Caption         =   "&Tipo de Registro:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblServico 
            Caption         =   "Serviço"
            Height          =   255
            Left            =   3240
            TabIndex        =   11
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label lblcalcFinanc 
            AutoSize        =   -1  'True
            Caption         =   "&Serviço:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Width           =   585
         End
      End
      Begin ComctlLib.ListView lvwLancDup 
         Height          =   1935
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327682
         Icons           =   "imgComissoes"
         SmallIcons      =   "imgComissoes"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Origem"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Núm."
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Empresa"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Tipo"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Parc."
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Valor"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Emissão"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Vencimento"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   8
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Pagamento"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   9
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Liberação"
            Object.Width           =   1411
         EndProperty
      End
      Begin ComctlLib.ImageList imgComissoes 
         Left            =   6240
         Top             =   1800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   2
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "calcGeracaoNotasFinanceiro.frx":000C
               Key             =   "Checked"
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "calcGeracaoNotasFinanceiro.frx":0326
               Key             =   "Unchecked"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "fcalcGeracaoNotasFinanceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngItem   As Long

Private Sub cmdFechar_Click()
  LimparLista
  
  Unload Me
  Exit Sub
End Sub

Private Sub cmdGerar_Click()
  On Error GoTo Erro
  
  cmdGerar.Enabled = False
  cmdFechar.Enabled = False
  
  If GerarNotas Then
    MsgFunc "Nota(s) gerada(s) com sucesso."
    
    LimparLista
  
    Unload Me
    Exit Sub
  End If
  
  cmdGerar.Enabled = True
  cmdFechar.Enabled = True

  Exit Sub
Erro:
  MsgFunc "Ocorreu um erro ao Iniciar a geração." & vbCrLf & "Erro: " & err.Number & " - " & err.Description & " - " & err.Source
  err.clear
  Exit Sub
End Sub



Private Sub Form_Load()
  On Error GoTo Erro
  
  CenterForm Me
  
  ComboAddItem cboTipoRegistro, "SELECT Tipo FROM [Tipos Globais];", "Tipo"
  cboTipoRegistro.Text = "Fatura"
  lblServico.Caption = NUL
  
  Exit Sub
Erro:
  MsgFunc "Ocorreu um erro no Form_load." & vbCrLf & "Erro: " & err.Number & " - " & err.Description & " - " & err.Source
  err.clear
  Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
  LimparLista
  Set fcalcGeracaoNotasFinanceiro = Nothing
End Sub

Private Sub lvwLancDup_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  On Error GoTo Erro
  
  lvwLancDup.Sorted = True
  lvwLancDup.SortKey = (ColumnHeader.Index - 1)
  lvwLancDup.Sorted = False
  
  Exit Sub
Erro:
  MsgFunc "Ocorreu um erro na função lvwLancDup_ColumnClick." & vbCrLf & "Erro: " & err.Number & " - " & err.Description & " - " & err.Source
  err.clear
  Exit Sub
End Sub
Private Sub lvwLancDup_DblClick()
  On Error GoTo Erro
  'Marcar e desmarcar o item
  Call MarcaDesmarca(lngItem)
  DoEvents

  Exit Sub
Erro:
  MsgFunc "Ocorreu um erro na função lvwLancDup_DblClick." & vbCrLf & "Erro: " & err.Number & " - " & err.Description & " - " & err.Source
  err.clear
  Exit Sub
End Sub
Private Sub lvwLancDup_ItemClick(ByVal item As ComctlLib.ListItem)
  On Error GoTo Erro
  
  ' aribuindo o item atual a variável que controla os itens
  lngItem = item.Index
  
  Exit Sub
Erro:
  MsgFunc "Ocorreu um erro na função lvwLancDup_ItemClick." & vbCrLf & "Erro: " & err.Number & " - " & err.Description & " - " & err.Source
  err.clear
  Exit Sub
End Sub
Private Sub MarcaDesmarca(item As Long)
  On Error GoTo Erro
  
  If item > 0 Then
    If lvwLancDup.ListItems(item).SmallIcon = 1 Then
      lvwLancDup.ListItems(item).SmallIcon = 2
    Else
      lvwLancDup.ListItems(item).SmallIcon = 1
    End If
  End If
  
  Exit Sub
Erro:
  MsgFunc "Ocorreu um erro na função MarcaDesmarca." & vbCrLf & "Erro: " & err.Number & " - " & err.Description & " - " & err.Source
  err.clear
  Exit Sub
End Sub
Private Function LimparLista()
  On Error GoTo Erro
  
  Dim Linha    As Long
  
  For Linha = 1 To lvwLancDup.ListItems.Count
    lvwLancDup.ListItems.Remove 1
  Next Linha

  Exit Function
Erro:
  MsgFunc "Ocorreu um erro na função LimparLista." & vbCrLf & "Erro: " & err.Number & " - " & err.Description & " - " & err.Source
  err.clear
  Exit Function
End Function

Private Sub txtServico_Change()
    GetAssocValue "Select Descrição from Serviços where Código = " & txtServico.Text, lblServico
End Sub

Private Sub txtServico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyPageDown And Shift = 0 Then
        PCampo "Serviços", "Serviços", pbCampo, txtServico, "Código"
    End If
End Sub
Private Function GerarNotas() As Boolean
  
    Dim Linha             As Long
    Dim bolMarcado        As Boolean
    Dim Origem            As String
    Dim Numero            As Long
    Dim TipoReg           As String
    Dim Empresa           As String
    Dim Parcela           As Long
    Dim rstNotas          As Object
    Dim rstItensNotas     As Object
    Dim ProxNumero        As Long
    Dim strWhere          As String
    Dim CodigoServico     As Long
    Dim intOpContabil     As Integer
    Dim strRetorno        As String
    Dim rstCond           As Object
    Dim strMensagem       As String
    Dim strSql            As String
    Dim rstTab            As Object
    Dim intCondPagto      As Integer

On Error GoTo Erro
  
    GerarNotas = False
    strMensagem = ""
    ReDim Preserve DuplCC(0)
    DuplCC(0).Valor = 0
  
    If Not IsValid(txtServico.Text) Then
        MsgFunc "Código de Serviço não informado."
        Exit Function
    ElseIf IsValid(txtServico.Text) And Not IsValid(lblServico.Caption) Then
        MsgFunc "Código de Serviço não cadastroado."
        Exit Function
    End If
    For Linha = 1 To lvwLancDup.ListItems.Count
        If lvwLancDup.ListItems(Linha).SmallIcon = 1 Then
            bolMarcado = True
            Exit For
        End If
    Next Linha
    If bolMarcado = False Then
        MsgFunc "Nenhuma Duplicata/Lançamento foi selecionado."
        Exit Function
    End If
            
    If AbreRecordset(rstNotas, GBL_NFSR, dbOpenDynaset) <> WL_ERRO Then
        For Linha = 1 To lvwLancDup.ListItems.Count
      
            If lvwLancDup.ListItems(Linha).SmallIcon = 1 Then
    
                Origem = lvwLancDup.ListItems(Linha)
                Numero = CLngDef(lvwLancDup.ListItems(Linha).SubItems(1))
                Empresa = lvwLancDup.ListItems(Linha).SubItems(2)
                TipoReg = lvwLancDup.ListItems(Linha).SubItems(3)
                Parcela = CLngDef(lvwLancDup.ListItems(Linha).SubItems(4))
                intOpContabil = val(fcalcGeracaoViaFinanceiro.txtOpContabil.Text)
                
                ProxNumero = ProximoNumero("Número", GBL_NFSR, "[Tipo de Registro] = " & Quote(cboTipoRegistro.Text, "'"))
                CodigoServico = GetFieldValue("[Código de Serviço]", "Serviços", "Código = " & txtServico.Text, , NUL)
            
                If Origem = "Duplicatas" Then
                      strWhere = "PagRec = 'R' and Nota= " & Numero & " and Tipo = " & Quote(TipoReg, "'") & " and Parcela = " & Parcela
                Else
                      strWhere = "PagRec = 'R' and Código= " & Numero & " and Tipo = " & Quote(TipoReg, "'")
                End If

                rstNotas.AddNew
                rstNotas("Tipo de Registro") = cboTipoRegistro.Text
                rstNotas("Número") = ProxNumero
                rstNotas("Fornecedor") = Left(DonaSistema, 15)
                rstNotas("Emissão") = Date
                rstNotas("Tipo da Empresa") = "Ativa"
                rstNotas("Empresa") = Left(Empresa, 15)
                rstNotas("Moeda") = IIf(IsValid(GetFieldValue("Moeda", Origem, strWhere)), GetFieldValue("Moeda", Origem, strWhere), "REAL")
                rstNotas("Cadastrado por") = UserName
                rstNotas("Registro Impresso") = False
                rstNotas("Situação") = "Normal"
                rstNotas("Data") = Date
                rstNotas("Hora") = Time
                rstNotas("Contato") = GetFieldValue("Contato", "[Empresas Contatos]", "Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Departamento") = GetFieldValue("Dpto", "[Empresas Contatos]", "Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Fone") = GetFieldValue("Fone1", "[Empresas Contatos]", "Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Fax") = GetFieldValue("Fax", "[Empresas Contatos]", "Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Observações") = "Nota gerada através de " & Mid(Origem, 1, Len(Origem) - 1) & ": " & Numero & IIf(Parcela > 0, "/" & Parcela, "") & " - Empresa: " & Empresa & " - Tipo: " & TipoReg
                rstNotas("Endereço de Cobrança") = GetFieldValue("Endereço", "[Empresas Endereços]", "Tipo = 'Cobrança' and Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Bairro de Cobrança") = GetFieldValue("Bairro", "[Empresas Endereços]", "Tipo = 'Cobrança' and Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("CEP de Cobrança") = GetFieldValue("CEP", "[Empresas Endereços]", "Tipo = 'Cobrança' and Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Cidade de Cobrança") = GetFieldValue("Cidade", "[Empresas Endereços]", "Tipo = 'Cobrança' and Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Estado de Cobrança") = GetFieldValue("Estado", "[Empresas Endereços]", "Tipo = 'Cobrança' and Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Fone de Cobrança") = GetFieldValue("Fone", "[Empresas Endereços]", "Tipo = 'Cobrança' and Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Ramal de Cobrança") = GetFieldValue("Ramal", "[Empresas Endereços]", "Tipo = 'Cobrança' and Apel = " & Quote(Left(Empresa, 15), "'"), , NUL)
                rstNotas("Total de Mão de Obra") = GetFieldValue("[Valor Original]", Origem, strWhere)
                rstNotas("Total de INSS") = GetFieldValue("[Valor Original]", Origem, strWhere) * (GetFieldValue("[Alíquota de INSS]", "[Códigos de Serviços]", "Código = " & CodigoServico, , ZERO) / 100)
                rstNotas("Total de ISS") = GetFieldValue("[Valor Original]", Origem, strWhere) * (GetFieldValue("[Alíquota de ISS]", "[Códigos de Serviços]", "Código = " & CodigoServico, , ZERO) / 100)
                rstNotas("Total de IRF") = GetFieldValue("[Valor Original]", Origem, strWhere) * (GetFieldValue("[Alíquota de IRRF]", "[Códigos de Serviços]", "Código = " & CodigoServico, , ZERO) / 100)
                rstNotas("Valor Total") = GetFieldValue("[Valor Original]", Origem, strWhere)
                rstNotas("Banco") = GetFieldValue("Banco", Origem, strWhere)
                rstNotas("Conta") = GetFieldValue("Conta", Origem, strWhere)
                strSql = "SELECT Empresas.Apel, Empresas.[Tabela de Preços], Empresas.CondPag, [Tabelas de Preços].CONDPAGTO "
                strSql = strSql & "FROM [Empresas] LEFT JOIN [Tabelas de Preços] ON Empresas.[Tabela de Preços]=[Tabelas de Preços].Código  "
                strSql = strSql & "WHERE Empresas.Apel='" & Left(Empresa, 15) & "'"
                If (AbreRecordset(rstTab, strSql, dbOpenSnapshot) = WL_OK) Then
                    rstNotas("TABPRECO") = GetValue(rstTab, "Tabela de Preços", ZERO)
                    If GetValue(rstTab, "CONDPAGTO", ZERO) > 0 Then
                        rstNotas("Condição de Pagamento") = GetValue(rstTab, "CONDPAGTO", ZERO)
                    ElseIf GetValue(rstTab, "CondPag", ZERO) > 0 Then
                        rstNotas("Condição de Pagamento") = GetValue(rstTab, "CondPag", ZERO)
                    End If
                End If
                FechaRecordset (rstTab)
                rstNotas("cd_operacao_contabil") = intOpContabil
                rstNotas.update
                
                AbreRecordset rstItensNotas, GBL_ITENS & GBL_NFSR, dbOpenDynaset
                rstItensNotas.AddNew
                rstItensNotas("Tipo de Registro") = cboTipoRegistro.Text
                rstItensNotas("Fornecedor") = Left(DonaSistema, 15)
                rstItensNotas("Número") = ProxNumero
                rstItensNotas("Item") = "1"
                rstItensNotas("Serviço") = txtServico.Text
                rstItensNotas("Descrição do Serviço") = GetFieldValue("Descrição", "Serviços", "Código = " & txtServico.Text, , NUL)
                rstItensNotas("Unidade") = GetFieldValue("Unidade", "Serviços", "Código = " & txtServico.Text, , NUL)
                rstItensNotas("Destinação") = "Outros"
                rstItensNotas("Quantidade") = "1"
                rstItensNotas("Tipo da Venda") = "Normal"
                rstItensNotas("Valor Original") = GetFieldValue("[Valor Original]", Origem, strWhere)
                rstItensNotas("Valor Líquido") = GetFieldValue("[Valor Original]", Origem, strWhere)
                rstItensNotas("IRF") = GetFieldValue("[Alíquota de IRRF]", "[Códigos de Serviços]", "Código = " & CodigoServico, , ZERO)
                rstItensNotas("INSS") = GetFieldValue("[Alíquota de INSS]", "[Códigos de Serviços]", "Código = " & CodigoServico, , ZERO)
                rstItensNotas("ISS") = GetFieldValue("[Alíquota de ISS]", "[Códigos de Serviços]", "Código = " & CodigoServico, , ZERO)
                rstItensNotas("Situação") = "Normal"
                rstItensNotas("Data da Previsão") = Date
                rstItensNotas("Tipo de desconto") = "Incondicional"
                rstItensNotas("Data da Cotação") = Date
                rstItensNotas("Centro de Custo") = GetFieldValue("Centro", Origem, strWhere)
                rstItensNotas.update

                FechaRecordset rstItensNotas
                
                'pt. 75606 - Moacir Pfau(23/06/2008) - Gerar duplicatas.
                intCondPagto = 0
                strSql = ""
                strSql = "SELECT Empresas.Apel, Empresas.[Tabela de Preços], Empresas.CondPag, [Tabelas de Preços].CONDPAGTO "
                strSql = strSql & "FROM [Empresas] LEFT JOIN [Tabelas de Preços] ON Empresas.[Tabela de Preços]=[Tabelas de Preços].Código  "
                strSql = strSql & "WHERE Empresas.Apel='" & Left(Empresa, 15) & "'"
                AbreRecordset rstTab, strSql
                If GetValue(rstTab, "CONDPAGTO", ZERO) > 0 Then
                    intCondPagto = GetValue(rstTab, "CONDPAGTO", ZERO)
                ElseIf GetValue(rstTab, "CondPag", ZERO) > 0 Then
                    intCondPagto = GetValue(rstTab, "CondPag", ZERO)
                Else
                    'strMensagem = strMensagem & "Não será possível gerar a duplicata da empresa: " & Empresa & ", número: " & Numero & ", favor gerar manualmente. " & vbCrLf
                End If
                FechaRecordset (rstTab)
                If intCondPagto > 0 Then
                    RateiodeDuplicatas True, False, "R", Left(Empresa, 15), ProxNumero, cboTipoRegistro.Text, intCondPagto, GetFieldValue("[Valor Original]", Origem, strWhere), Format(Now, "DD/MM/YYYY"), True, , , , , , , , GetFieldValue("Banco", Origem, strWhere), GetFieldValue("Conta", Origem, strWhere), , , , , GetFieldValue("Centro", Origem, strWhere)
                End If

                If Origem = "Lançamentos" Then
                    ExecuteSQL "DELETE FROM Lançamentos WHERE PagRec='R' AND Código=" & Numero & " AND Parcela=" & Parcela
                ElseIf Origem = "Duplicatas" Then
                    ExecuteSQL "DELETE FROM Duplicatas WHERE PagRec='R' AND Nota=" & Numero & " AND Parcela=" & Parcela & " AND Empresa='" & Empresa & "' AND Tipo='" & TipoReg & "'"
                End If
            End If
        Next Linha
    End If
    FechaRecordset rstNotas
   
    GerarNotas = True
  
    If GerarNotas And strMensagem <> "" Then
        MsgBox strMensagem
    End If
    
    Exit Function
Erro:
    MsgFunc "Ocorreu um erro na função GerarNota(s)." & vbCrLf & "Erro: " & err.Number & " - " & err.Description & " - " & err.Source
    err.clear
    Exit Function
End Function

'ComctlLib.ListItem
'  If item > 0 Then
'    If lvwLancDup.ListItems(item).SmallIcon = 1 Then
'      lvwLancDup.ListItems(item).SmallIcon = 2
'    Else
'      lvwLancDup.ListItems(item).SmallIcon = 1
'    End If
'  End If

Private Sub cmdDesmarcarTodos_Click()
    Dim intIndex           As Integer
On Error GoTo err
    For intIndex = 1 To lvwLancDup.ListItems.Count
        lvwLancDup.ListItems(intIndex).SmallIcon = 2
    Next
err:
End Sub

Private Sub cmdMarcarTodos_Click()
    Dim intIndex           As Integer
On Error GoTo err
    For intIndex = 1 To lvwLancDup.ListItems.Count
        lvwLancDup.ListItems(intIndex).SmallIcon = 1
    Next
err:
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
