VERSION 5.00
Begin VB.Form frptRazao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Raz�o Auxiliar"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "Razao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRazao 
      Cancel          =   -1  'True
      Caption         =   "#"
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdRazao 
      Caption         =   "Im&primir"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdRazao 
      Caption         =   "&Visualizar..."
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame fraRazao 
      Caption         =   "Raz�o"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtRazao 
         Height          =   315
         Index           =   2
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtRazao 
         Height          =   315
         Index           =   1
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtRazao 
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cboRazao 
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblDescEmp 
         Caption         =   "lblDescEmp"
         Height          =   195
         Left            =   3000
         TabIndex        =   8
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   2880
      End
      Begin VB.Label lblRazao 
         AutoSize        =   -1  'True
         Caption         =   "&Empresa:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label lblRazao 
         AutoSize        =   -1  'True
         Caption         =   "Data &Final:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblRazao 
         AutoSize        =   -1  'True
         Caption         =   "Data &Inicial:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblRazao 
         AutoSize        =   -1  'True
         Caption         =   "&Tipo:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   360
      End
   End
End
Attribute VB_Name = "frptRazao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Verifica se o usu�rio cancelou
Private mbolCancel As Boolean

Private Sub cboRazao_GotFocus(Index As Integer)
    RazaoMsgStatus cboRazao(Index).TabIndex
End Sub

Private Sub cmdRazao_Click(Index As Integer)
    If (Index < 2) Then         'Bot�es Visualizar ou Imprimir
        cmdRazao(0).Enabled = False
        cmdRazao(1).Enabled = False
        cmdRazao(2).Caption = LoadResString(IDS_CANCELAR)
        FiltraEmpresa IIf((Index > 0), wrToPrinter, wrToWindow)
        cmdRazao(0).Enabled = True
        cmdRazao(1).Enabled = True
        cmdRazao(2).Caption = LoadResString(IDS_FECHAR)
    Else
        If (cmdRazao(0).Enabled) Then
            Unload Me
        Else
            mbolCancel = True
            MsgBar LoadResString(171) & LoadResString(14)
        End If
    End If
End Sub

Private Sub Form_Load()
    'Carrega a lista de op��es das caixas combo Tipo
    LoadResOptions 1021, cboRazao(0), , 0
    'Carrega o caption do bot�o Fechar/Cancelar
    cmdRazao(2).Caption = LoadResString(IDS_FECHAR)
    'Valores padr�o para os campos de data -
    txtRazao(0).Text = FirstDayS(Date)
    '�ltimo dia do m�s corrente
    txtRazao(1).Text = LastDayS(Date)
    'Limpando o Label de descri��o de empresa
    lblDescEmp.Caption = NUL
    'Centraliza e exibe o formul�rio
    CenterForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frptRazao = Nothing
End Sub

Private Sub txtRazao_Change(Index As Integer)
    'Campo Empresa
    If (Index = 2) Then
        GetAssocValue "SELECT Raz�o, Apel FROM Empresas WHERE Apel = '" & txtRazao(2).Text & "';", lblDescEmp, txtRazao(2)
    End If
End Sub

Private Sub txtRazao_GotFocus(Index As Integer)
    Selecione txtRazao(Index)
    RazaoMsgStatus txtRazao(Index).TabIndex
End Sub

Private Sub txtRazao_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strTipo   As String       'Tipo da Empresa
    Dim strSelEmp As String       'Instru��o de Sele��o das empresas para pesquisa
    
    If ((Shift = 0) And (KeyCode = vbKeyPageDown)) Then
        'Campo Empresa
        If (Index = 2) Then
            'Cliente ou Fonecedor
            strTipo = GetResOptions(1003, IIf((cboRazao(0).ListIndex > 0), 1, 2))
            strSelEmp = "SELECT Apel, Raz�o, Tipo, Pessoa, [CNPJ/CPF], [IEst/RG], " & _
                        "Endere�o, Bairro, CEP, Cidade, Estado, Regi�o, Fone1, Contato, " & _
                        "Dpto FROM Empresas WHERE Tipo <> '" & strTipo & "';"
                        
            PCampo "Empresas", strSelEmp, pbCampo, txtRazao(2), "Apel"
        End If
    End If
End Sub

Private Sub txtRazao_KeyPress(Index As Integer, KeyAscii As Integer)
    'Datas Inicial e Final
    If (Index < 2) Then
        SetMascara KeyAscii, txtRazao(Index).SelStart, MASK_DATE4
    End If
    If (Index = 3) Then
        SetMascara KeyAscii, txtRazao(Index).SelStart, fMask("Moedas", "Moeda")
    End If
End Sub

' SUB.......: RazaoMsgStatus
' Objetivo..: Exibe mensagens informativas na barra de status do programa para
'             cada controle que recebe o foco.
' Argumento.: [iTabIdx]: Propriedade TabIndex do controle.
Private Sub RazaoMsgStatus(iTabIdx As Integer)
    Select Case iTabIdx
        'ComboBox Tipo
        Case 2
            MsgBar "Tipo das empresas"
        'Data Inicial
        Case 4
            MsgBar "Data Inicial do per�do de apura��o"
        'Data Final
        Case 6
            MsgBar "Data Final do per�odo de apura��o"
        'Empresa
        Case 8
            MsgBar "Nome Fantasia da Empresa" & ResolveResString(75, resUM, "Empresas")
        'ComboBox de Ordem
        Case 11
            MsgBar "Ordem do Relat�rio"
    End Select
End Sub

' SUB.......: FiltraEmpresa
' Objetivo..: Filtra o cadastro de empresas para obter a empresa selecionada
'             pelo usu�rio. Usa a fun��o AddValores para completar a tabela
'             auxiliar para a gera��o do relat�rio.
' Argumento.: [dImpressao]: Destino da impress�o.
Private Sub FiltraEmpresa(dImpressao As Long)
    Dim rstEmp          As Object
    Dim strEmp          As String
    Dim dInit           As Variant
    Dim dFinal          As Variant
    Dim rstDados        As Object
    Dim rstSaldo        As Object
    Dim rstSaldoGeral   As Object

    'Reseta Cancel
    mbolCancel = False
    SetPtr vbArrowHourglass
    dInit = CDateDef(txtRazao(0).Text, Empty)
    dFinal = CDateDef(txtRazao(1).Text, LastDay(Date))
    strEmp = "SELECT Apel, Raz�o FROM Empresas WHERE Tipo <> '"
    If (cboRazao(0).ListIndex) Then
        AppendStr strEmp, GetResOptions(1003, 1) & "'"
    Else
        AppendStr strEmp, GetResOptions(1003, 2) & "'"
    End If
    If (Len(txtRazao(2).Text)) Then
        AppendStr strEmp, " AND Apel = '" & txtRazao(2).Text & "'"
    End If
    strEmp = strEmp & " ORDER BY Apel;"
    'Pt. 95368 - Moacir Pfau(04/11/2009)
    If (WL_OK = AbreRecordsetDAO(rstEmp, strEmp, dbOpenSnapshot)) Then
        If (CrieAuxRazao(rstDados)) Then
            If (CriaAuxSaldo(rstSaldo)) Then
                'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                If (CriaAuxSaldoGeral(rstSaldoGeral)) Then
                    If (AddRegistros(rstEmp, rstDados, rstSaldo, dInit, dFinal, rstSaldoGeral)) Then
                        RazaoAuxiliar rstDados, rstSaldo, dImpressao, rstSaldoGeral
                    End If
                End If
            End If
            DeleteAux rstSaldo, NUL
        End If
        DeleteAux rstDados, NUL
        FechaRecordset rstDados
        FechaRecordset rstSaldo
    Else
        MsgFunc LoadResString(IDS_RECORDNOTFOUND)
    End If
    FechaRecordset rstEmp
    MsgBar Caption
    SetPtr vbDefault
End Sub

' FUNCTION..: CrieAuxRazao
' Objetivo..: Cria a tabela auxiliar para impress�o do relat�rio
' Argumento.: [rstAux]: Recordset que retorna com a tabela aberta.
' Retorna...: True se criar a tabela com sucesso, False se n�o.
Private Function CrieAuxRazao(rstAux As Object) As Boolean
    Dim fsRazao(10) As FieldStruct
  
    AppendVar fsRazao(0), "Apel", dbText, 20
    AppendVar fsRazao(1), "Empresa", dbText, 95
    AppendVar fsRazao(2), "Data", dbDate
    AppendVar fsRazao(3), "Descri��o", dbText, 35
    AppendVar fsRazao(4), "N�mero", dbDouble
    AppendVar fsRazao(5), "Tipo", dbText, 35
    AppendVar fsRazao(6), "D�bito", dbCurrency
    AppendVar fsRazao(7), "Cr�dito", dbCurrency
    AppendVar fsRazao(8), "Saldo", dbCurrency
    AppendVar fsRazao(9), "D�bito/Cr�dito", dbText, 1
    AppendVar fsRazao(10), "Parcela", dbLong
    
    CrieAuxRazao = CrieAux(rstAux, fsRazao())
End Function

' FUNCTION..: CriaAuxSaldo
' Objetivo..: Cria uma tabela auxiliar para gravar os saldos das empresas.
' Argumento.: [rstSaldos]: Recordset que retorna com a tabela aberta.
' Retorna...: True se criar a tabela, False se n�o.
Private Function CriaAuxSaldo(rstSaldos As Object) As Boolean
    Dim fsSaldo(2) As FieldStruct

    AppendVar fsSaldo(0), "Apel", dbText, 15
    AppendVar fsSaldo(1), "Final", dbBoolean
    AppendVar fsSaldo(2), "Saldo", dbCurrency
    
    CriaAuxSaldo = CrieAux(rstSaldos, fsSaldo())
End Function

Private Function CriaAuxSaldoGeral(rstSaldosGeral As Object) As Boolean
    Dim fsSaldoGeral(2) As FieldStruct

    AppendVar fsSaldoGeral(0), "Debito", dbCurrency
    AppendVar fsSaldoGeral(1), "Credito", dbCurrency
    AppendVar fsSaldoGeral(2), "Saldo", dbCurrency
    
    CriaAuxSaldoGeral = CrieAux(rstSaldosGeral, fsSaldoGeral())
End Function

' FUNCTION..: AddRegistros
' Objetivo..: Completa a tabela auxiliar com os dados necess�rios
'             para impress�o do relat�rio.
' Argumentos: [rstEmpresa]: Recordset com as empresas escolhidas.
'             [rstTemp]   : Recordset da tabela tempor�ria.
'             [rstAuxSld] : Recordset da tabela de saldos de empresas.
'             [dtInicial] : Data inicial
'             [dtFinal]   : Data Final
' Retorna...: True se puder gravar a tabela com sucesso. False se um
'             erro ocorrer ou se o usu�rio cancelar.
Private Function AddRegistros(rstEmpresa As Object, rstTemp As Object, rstAuxSld As Object, dtInicial, dtFinal, rstAuxSldGeral As Object) As Boolean
    Dim rstLanctos As Object           'Para os lan�amentos encontrados
    Dim strLanctos As String              'Instru��o de sele��o dos lan�amentos
    Dim cSaldo     As Currency            'Valor do Saldo inicial
    Dim strApel    As String              'Nome Fantasia da empresa
    Dim blnCliente As Boolean             'Relat�rio para empresas Cliente ou Fonecedores
    Dim sMensLanc  As String              'String de mensagem para lan�amentos
    Dim sMensDupl  As String              'String de mensagem para duplicatas
    Dim strTabela  As String
    Dim dblTotDed  As Double
    Dim dblTotCre  As Double
    
On Error GoTo AddRegistro_Erro
    
    SetPtr vbHourglass
    blnCliente = (cboRazao(0).ListIndex = 1)
    If (blnCliente) Then        'Se Cliente s�o lan�amentos a receber
        sMensLanc = "Lan�amentos Recebidos:"
        sMensDupl = "Duplicatas Recebidas:"
    Else                        'Se Fornecedos s�o lan�amentos a pagar
        sMensLanc = "Lan�amentos Pagos:"
        sMensDupl = "Duplicatas Pagas:"
    End If
    dblTotDed = 0
    dblTotCre = 0
    Do
        strApel = GetValue(rstEmpresa, "Apel")
        If (IsEmpty(dtInicial)) Then
            'Se a data inicial n�o foi informada o saldo inicial ser� sempre zero
            cSaldo = 0
        Else
            SimpleMsgBar "Pesquisando saldo inicial da empresa " & strApel
            strLanctos = "Empresa = '" & strApel & "' AND [Emiss�o] < " & InverteData(dtInicial, True) & _
                         " AND (Pagamento IS NULL OR Pagamento >= " & InverteData(dtInicial, True) & ")" & _
                         " AND PagRec = "
            AppendStr strLanctos, IIf(blnCliente, "'R'", "'P'")
            cSaldo = Soma("([Valor Original] + [Acr�scimo] - Abatimento)", "Duplicatas", strLanctos)
            cSaldo = (cSaldo + Soma("([Valor Original] + [Acr�scimo] - Abatimento)", "[Lan�amentos]", strLanctos))
            If (blnCliente) Then
                cSaldo = -cSaldo
            End If
        End If
        If mbolCancel Then
            GoTo AddRegistro_Erro
        End If
        'Habilita ao usu�rio cancelar
        DoEvents
        
        'Grava o saldo desta empresa na tabela auxiliar de saldos
        rstAuxSld.AddNew
        rstAuxSld("Apel").value = strApel
        rstAuxSld("Saldo").value = cSaldo
        rstAuxSld("Final").value = False
        rstAuxSld.update
        
        'Selecionando os dados de Duplicatas emitidos no per�odo espec�ficado
        strLanctos = "SELECT Nota, Parcela, Tipo, [Emiss�o], Parcela, ([Valor Original] " & _
                     "+ [Acr�scimo] - Abatimento) AS Total FROM Duplicatas WHERE " & _
                     "Empresa = '" & strApel & "'"
        'Se h� data inicial
        If (Not IsEmpty(dtInicial)) Then
            AppendStr strLanctos, " AND [Emiss�o] BETWEEN " & InverteData(dtInicial, True)
            AppendStr strLanctos, " AND " & InverteData(dtFinal, True)
        Else
            AppendStr strLanctos, " AND [Emiss�o] <= " & InverteData(dtFinal, True)
        End If
        AppendStr strLanctos, " AND PagRec = " & IIf(blnCliente, "'R'", "'P'")
        
        'pt. 75830 - Dulcino J�nior
        'Ordenando por data de emiss�o
        AppendStr strLanctos, " ORDER BY [Emiss�o]"
        SimpleMsgBar "Obtendo movimenta��o em Duplicatas"
        'Pt. 95368 - Moacir Pfau(04/11/2009)
        If (WL_OK = AbreRecordsetDAO(rstLanctos, strLanctos, dbOpenSnapshot)) Then
            Do
                If mbolCancel Then
                    GoTo AddRegistro_Erro
                End If
                DoEvents
                rstTemp.AddNew
                rstTemp("Apel").value = strApel
                rstTemp("Empresa").value = GetValue(rstEmpresa, "Raz�o")
                rstTemp("Data").value = GetValue(rstLanctos, "[Emiss�o]")
                rstTemp("Descri��o").value = "Duplicatas do Per�odo:"
                rstTemp("N�mero").value = GetValue(rstLanctos, "Nota")
                rstTemp("Parcela").value = GetValue(rstLanctos, "Parcela")
                rstTemp("Tipo").value = GetValue(rstLanctos, "Tipo")
                If (blnCliente) Then
                    rstTemp("D�bito").value = GetValue(rstLanctos, "Total")
                    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                    rstTemp("D�bito/Cr�dito").value = "D"
                    dblTotDed = dblTotDed + GetValue(rstLanctos, "Total")
                    cSaldo = cSaldo - GetValue(rstLanctos, "Total")       'Calcula o Saldo atual
                Else
                    rstTemp("Cr�dito").value = GetValue(rstLanctos, "Total")
                    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                    rstTemp("D�bito/Cr�dito").value = "C"
                    dblTotCre = dblTotCre + GetValue(rstLanctos, "Total")
                    cSaldo = cSaldo + GetValue(rstLanctos, "Total")
                End If
                rstTemp("Saldo").value = cSaldo
                rstTemp.update
                rstLanctos.MoveNext
            Loop Until rstLanctos.EOF
        End If
        FechaRecordset rstLanctos
        
        ' Selecionando os dados de Lan�amentos emitidos no per�odo espec�ficado
        strLanctos = "SELECT C�digo, Tipo, [Emiss�o], Parcela, ([Valor Original] " & _
                     "+ [Acr�scimo] - Abatimento) AS Total FROM [Lan�amentos] WHERE " & _
                     "Empresa = '" & strApel & "'"
        ' Se h� data inicial
        If (Not IsEmpty(dtInicial)) Then
            AppendStr strLanctos, " AND [Emiss�o] BETWEEN " & InverteData(dtInicial, True)
            AppendStr strLanctos, " AND " & InverteData(dtFinal, True)
        Else
            AppendStr strLanctos, " AND [Emiss�o] <= " & InverteData(dtFinal, True)
        End If
        AppendStr strLanctos, " AND PagRec = " & IIf(blnCliente, "'R'", "'P'")
        
        'pt. 75830 - Dulcino J�nior
        'Ordenando por data de emiss�o
        AppendStr strLanctos, " ORDER BY [Emiss�o]"
        'Pt. 95368 - Moacir Pfau(04/11/2009)
        If (WL_OK = AbreRecordsetDAO(rstLanctos, strLanctos, dbOpenSnapshot)) Then
            Do
                If mbolCancel Then
                    GoTo AddRegistro_Erro
                End If
                DoEvents
                rstTemp.AddNew
                rstTemp("Apel").value = strApel
                rstTemp("Empresa").value = GetValue(rstEmpresa, "Raz�o")
                rstTemp("Data").value = GetValue(rstLanctos, "[Emiss�o]")
                rstTemp("Descri��o").value = "Lan�amentos do Per�odo:"
                rstTemp("N�mero").value = GetValue(rstLanctos, "C�digo")
                rstTemp("Parcela").value = GetValue(rstLanctos, "Parcela")
                rstTemp("Tipo").value = GetValue(rstLanctos, "Tipo")
                If (blnCliente) Then
                    rstTemp("D�bito").value = GetValue(rstLanctos, "Total")
                    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                    rstTemp("D�bito/Cr�dito").value = "D"
                    dblTotDed = dblTotDed + GetValue(rstLanctos, "Total")
                    cSaldo = cSaldo - GetValue(rstLanctos, "Total")       'Calcula o Saldo atual
                Else
                    rstTemp("Cr�dito").value = GetValue(rstLanctos, "Total")
                    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                    rstTemp("D�bito/Cr�dito").value = "C"
                    dblTotCre = dblTotCre + GetValue(rstLanctos, "Total")
                    cSaldo = cSaldo + GetValue(rstLanctos, "Total")
                End If
                rstTemp("Saldo").value = cSaldo
                rstTemp.update
                rstLanctos.MoveNext
            Loop Until rstLanctos.EOF
        End If
        FechaRecordset rstLanctos
        
        ' Selecionando os dados de Duplicatas cujo pagamento se encontra no per�odo especificado
        strLanctos = "SELECT Nota, Tipo, Pagamento, Parcela, ([Valor Original] " & _
                     "+ [Acr�scimo] - Abatimento) AS Total FROM Duplicatas WHERE " & _
                     "Empresa = '" & strApel & "'"
        ' Se h� data inicial
        If (Not IsEmpty(dtInicial)) Then
            AppendStr strLanctos, " AND Pagamento BETWEEN " & InverteData(dtInicial, True)
            AppendStr strLanctos, " AND " & InverteData(dtFinal, True)
        Else
            AppendStr strLanctos, " AND Pagamento <= " & InverteData(dtFinal, True)
        End If
        AppendStr strLanctos, " AND PagRec = " & IIf(blnCliente, "'R'", "'P'")
        
        'pt. 75830 - Dulcino J�nior
        'Ordenando por data de emiss�o
        AppendStr strLanctos, " ORDER BY [Emiss�o]"
        SimpleMsgBar "Obtendo movimenta��o em Lan�amentos"
        'Pt. 95368 - Moacir Pfau(04/11/2009)
        If (WL_OK = AbreRecordsetDAO(rstLanctos, strLanctos, dbOpenSnapshot)) Then
            Do
                If mbolCancel Then
                    GoTo AddRegistro_Erro
                End If
                DoEvents
                rstTemp.AddNew
                rstTemp("Apel").value = strApel
                rstTemp("Empresa").value = GetValue(rstEmpresa, "Raz�o")
                rstTemp("Data").value = GetValue(rstLanctos, "Pagamento")
                rstTemp("Descri��o").value = sMensDupl
                rstTemp("N�mero").value = GetValue(rstLanctos, "Nota")
                rstTemp("Parcela").value = GetValue(rstLanctos, "Parcela")
                rstTemp("Tipo").value = GetValue(rstLanctos, "Tipo")
                If (blnCliente) Then
                    rstTemp("Cr�dito").value = GetValue(rstLanctos, "Total")
                    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                    rstTemp("D�bito/Cr�dito").value = "C"
                    dblTotCre = dblTotCre + GetValue(rstLanctos, "Total")
                    cSaldo = cSaldo + GetValue(rstLanctos, "Total")
                Else
                    rstTemp("D�bito").value = GetValue(rstLanctos, "Total")
                    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                    rstTemp("D�bito/Cr�dito").value = "D"
                    dblTotDed = dblTotDed + GetValue(rstLanctos, "Total")
                    cSaldo = cSaldo - GetValue(rstLanctos, "Total")
                End If
                rstTemp("Saldo").value = cSaldo
                rstTemp.update
                rstLanctos.MoveNext
            Loop Until rstLanctos.EOF
        End If
        FechaRecordset rstLanctos
        
        ' Selecionando os dados de Lan�amentos cujo pagamento se encontra no per�odo especificado
        strLanctos = "SELECT C�digo, Tipo, Pagamento, Parcela, ([Valor Original] " & _
                     "+ [Acr�scimo] - Abatimento) AS Total FROM [Lan�amentos] WHERE " & _
                     "Empresa = '" & strApel & "'"
        ' Se h� data inicial
        If (Not IsEmpty(dtInicial)) Then
            AppendStr strLanctos, " AND Pagamento BETWEEN " & InverteData(dtInicial, True)
            AppendStr strLanctos, " AND " & InverteData(dtFinal, True)
        Else
            AppendStr strLanctos, " AND Pagamento <= " & InverteData(dtFinal, True)
        End If
        AppendStr strLanctos, " AND PagRec = " & IIf(blnCliente, "'R'", "'P'")
        
        'pt. 75830 - Dulcino J�nior
        'Ordenando por data de emiss�o
        AppendStr strLanctos, " ORDER BY [Emiss�o]"
        'Pt. 95368 - Moacir Pfau(04/11/2009)
        If (WL_OK = AbreRecordsetDAO(rstLanctos, strLanctos, dbOpenSnapshot)) Then
            Do
                If mbolCancel Then
                    GoTo AddRegistro_Erro
                End If
                DoEvents
                rstTemp.AddNew
                rstTemp("Apel").value = strApel
                rstTemp("Empresa").value = GetValue(rstEmpresa, "Raz�o")
                rstTemp("Data").value = GetValue(rstLanctos, "Pagamento")
                rstTemp("Descri��o").value = sMensLanc
                rstTemp("N�mero").value = GetValue(rstLanctos, "C�digo")
                rstTemp("Parcela").value = GetValue(rstLanctos, "Parcela")
                rstTemp("Tipo").value = GetValue(rstLanctos, "Tipo")
                If (blnCliente) Then
                    rstTemp("Cr�dito").value = GetValue(rstLanctos, "Total")
                    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                    rstTemp("D�bito/Cr�dito").value = "C"
                    dblTotCre = dblTotCre + GetValue(rstLanctos, "Total")
                    cSaldo = cSaldo + GetValue(rstLanctos, "Total")       'Calcula o Saldo atual
                Else
                    rstTemp("D�bito").value = GetValue(rstLanctos, "Total")
                    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
                    rstTemp("D�bito/Cr�dito").value = "D"
                    dblTotDed = dblTotDed + GetValue(rstLanctos, "Total")
                    cSaldo = cSaldo - GetValue(rstLanctos, "Total")
                End If
                rstTemp("Saldo").value = cSaldo
                rstTemp.update
                rstLanctos.MoveNext
            Loop Until rstLanctos.EOF
        End If
        FechaRecordset rstLanctos
        
        ' Gravando o Saldo Final desta empresa
        rstAuxSld.AddNew
        rstAuxSld("Apel").value = strApel
        rstAuxSld("Final").value = True
        rstAuxSld("Saldo").value = cSaldo
        rstAuxSld.update
        rstEmpresa.MoveNext
    Loop Until rstEmpresa.EOF
    
    'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(08/05/2013)
    If dblTotCre > 0 Or dblTotDed > 0 Then
        rstAuxSldGeral.AddNew
        rstAuxSldGeral("Debito").value = dblTotDed
        rstAuxSldGeral("Credito").value = dblTotCre
        rstAuxSldGeral("Saldo").value = dblTotCre - dblTotDed
        rstAuxSldGeral.update
    End If
    
    'pt. 84768 - Ivo Sousa (28/10/2008)
    #If FOXSQL = 1 Then
    strTabela = ExtractTableName(rstTemp.Source)
    #Else
    strTabela = rstTemp.name
    #End If
    rstTemp.Close
    'Pt. 95368 - Moacir Pfau(04/11/2009)
    Call AbreRecordsetDAO(rstTemp, "SELECT * FROM " & strTabela & " ORDER BY Apel,Data")
    Call AtualizaSaldos(rstTemp, rstEmpresa, CDate(dtInicial), blnCliente)
    'Se encontrou algum registro
    If (Not EstaVazio(rstTemp)) Then
        'Tabela completa e pronta
        AddRegistros = True
    Else
        MsgFunc LoadResString(IDS_RECORDNOTFOUND)
    End If
    Exit Function
  
AddRegistro_Erro:
    If (err.Number) Then
      DAOErros NUL
    End If
    FechaRecordset rstLanctos
    SetPtr vbDefault
    AddRegistros = False
End Function

'Data.......: 05/11/2008
'Autor......: Ivo Sousa
'Descri��o..: Atualiza os saldos das empresas que ser�o mostrados no relat�rio.
'Parametros.: [Object] Recordset com os registros a serem alterados.
'...........: [Object] Recordset com as empresas.
'...........: [Date]   Data Incial da consulta para buscar o saldo anterior.
'...........: [Boolean]Se a consulta � para clientes ou fornecedores para atualizar o saldo.
Private Sub AtualizaSaldos(ByRef rstPrincipal As Object, rstEmpresas As Object, dtInicial As Date, blnCliente As Boolean)
    Dim curSaldo As Currency
    Dim strLanctos As String
    Dim fakedao As New CGenericRecordset
    'Pt. 95368 - Moacir Pfau(03/11/2009)
    fakedao.Initialize rstPrincipal
    If fakedao.Recordcount > 0 Then
        fakedao.MoveFirst
        rstEmpresas.MoveFirst
        While Not rstEmpresas.EOF
            strLanctos = "Empresa = '" & GetValue(rstEmpresas, "Apel") & "' AND [Emiss�o] < " & InverteData(dtInicial, True) & " AND (Pagamento IS NULL OR Pagamento >= " & InverteData(dtInicial, True) & ") AND PagRec = " & IIf(blnCliente, "'R'", "'P'")
            curSaldo = Soma("([Valor Original] + [Acr�scimo] - Abatimento)", "Duplicatas", strLanctos)
            curSaldo = (curSaldo + Soma("([Valor Original] + [Acr�scimo] - Abatimento)", "Lan�amentos", strLanctos))
            While Not (GetValue(rstEmpresas, "Apel") <> GetValue(fakedao, "Apel"))
                'Pt. 95368 - Moacir Pfau(21/10/2009)
                fakedao.Edit
                If Not blnCliente Then
                    curSaldo = curSaldo + (GetValue(fakedao, "Cr�dito") - GetValue(fakedao, "D�bito"))
                    fakedao("D�bito/Cr�dito").value = IIf(curSaldo > 0, "C", "D")
                Else
                    curSaldo = curSaldo - (GetValue(fakedao, "Cr�dito") - GetValue(fakedao, "D�bito"))
                    fakedao("D�bito/Cr�dito").value = IIf(curSaldo > 0, "D", "C")
                End If
                fakedao("Saldo").value = Format(curSaldo, "#,##0.00")
                fakedao.update
                fakedao.MoveNext
            Wend
            rstEmpresas.MoveNext
        Wend
    End If
    Set fakedao = Nothing
End Sub

' SUB.......: RazaoAuxiliar
' Objetivo..: Configura o gerador de relat�rios para imprimir o resultado do filtro criado pelo usu�rio.
' Argumentos: [rstSource]: Recordset de origem dos dados.
'             [rstSaldos]: Recordset com os saldos das empresas.
'             [lDestino] : Destino da impress�o.
Private Sub RazaoAuxiliar(rstSource As Object, rstSaldos As Object, lDestino As Long, rstSaldoGeral As Object)
    Dim wrkRazao  As KeybReport
    Dim strTitulo As String           'Subt�tulo do relat�rio
  
    'Resolvendo o subt�tulo
    'Gerando Relat�rio..."
    SimpleMsgBar LoadResString(160)
    strTitulo = "Per�odo:"
    If (EData(txtRazao(0).Text)) Then
        AppendStr strTitulo, " de " & txtRazao(0).Text
    End If
    If (EData(txtRazao(1).Text)) Then
        AppendStr strTitulo, " at� " & txtRazao(1).Text
    Else
        AppendStr strTitulo, " at� " & DataToStr(Date)
    End If
    Set wrkRazao = New KeybReport
        
    With wrkRazao
        Set .DatabaseName = GlobalDataBase
        'Pt. 95368 - Moacir Pfau(21/10/2009)
        'If gTipoDB = Access Then
            If (Not (rstSource Is Nothing)) And (TypeOf rstSource Is ADODB.Recordset) Then
                Set .Recordset = rstSource
            Else
                Set .Recordset = rstSource.OpenRecordset()
            End If
        'Else
        '    Set .Recordset = rstSource
        'End If
        .Destino = lDestino
        .ScaleMode = vbMillimeters
        .AutoRedraw = True
        .Tipo = wrObjectDraw
        'pt. 84768 - Ivo Sousa (29/10/2008)
        .WindowHeight = Screen.Height
        .WindowWidth = Screen.Width
        
        .WindowTitulo = "Raz�o Auxiliar por Empresa"
        PageHeader wrkRazao, "Raz�o Auxiliar de " & cboRazao(0).Text
        
        'Adiciona uma linha ao cabe�alho para o subt�tulo do relat�rio
        .UltimaSecao.AddLinha
        .UltimaLinha.AddCampo , wrCSFixedText, strTitulo, wrTACentro
        .FontStyle = wrFSBold
        .FontSize = 8
        
        'Criando o Grupo principal, quebra por Empresa
        .AddGrupo "1"
        .Grupo(1).Quebra = "Apel"
        .Grupo(1).AddSecao scHeader, 3
        With .Grupo(1).Header.Linha(2)
            .Height = .Height * 2
            .DrawBorder = wrDBAllBorders
            .AddCampo , wrCSFixedText, "Empresa:", , 15
            .Campo(1).Top = ((.Height / 2) - (.Campo(1).Height / 2))
            .AddCampo , , "Apel", , 25
            .Campo(2).Top = .Campo(1).Top
            .AddCampo , , "Empresa", , 80
            .Campo(3).Top = .Campo(1).Top
            .AddCampo , wrCSFixedText, "Saldo Anterior:", , 25
            .Campo(4).Top = .Campo(1).Top
            .AddCampo , wrCSDataLink, "Saldo", wrTADireito, 27, 159
            .Campo(5).Top = .Campo(1).Top
            .Campo(5).Formato = "#,##0.00"" C"";#,##0.00"" D"";0.00"
            If TypeOf rstSaldos Is ADODB.Recordset Then
                .Campo(5).TableLink = ExtractTableName(rstSaldos.Source)
            Else
                .Campo(5).TableLink = rstSaldos.name
            End If
            .Campo(5).DataLink = "Apel = {*Quebra} AND Final = False"
        End With
        
        'T�tulos das colunas
        With .Grupo(1).Header.Linha(3)
            .DrawBorder = wrDBBottomBorder
            .AddCampo , wrCSFixedText, "Data", , 17
            .AddCampo , wrCSFixedText, "Hist�rico", , 45
            .AddCampo , wrCSFixedText, "Documento", , 33, 62
            .AddCampo , wrCSFixedText, "D�bito", wrTADireito, 26, 95
            .AddCampo , wrCSFixedText, "Cr�dito", wrTADireito, 26, 121
            .AddCampo , wrCSFixedText, "Saldo", wrTADireito, 26, 147
            'pt. 84768 - Ivo Sousa (28/10/2008)
            .AddCampo , wrCSFixedText, "D�bito/Cr�dito", wrTADireito, 25, 172
        End With
        .FontStyle = wrFSNormal
        
        'Se��o de impress�o dos dados
        .Grupo(1).AddSecao scDetalhe, 1
        With .Grupo(1).Detalhe.Linha(1)
            .AddCampo , , "Data", , 17
            .Campo(1).Formato = FDATA
            .AddCampo , , "Descri��o", , 45
            'Vinicius Elyseu(30/05/2016) - Projeto: #0 - Hist�ria: #0 - Desenv: #0
            .AddCampo , , "N�mero", wrTADireito, 27, 62
            'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(07/05/2013)
            '.Campo(3).Formato = "000000000"" - """
            .Campo(3).Formato = StrZero(0, 15) & " - "
            .AddCampo , , "Parcela", , 7
            .Campo(4).Formato = "000"
            .AddCampo , , "Tipo", , 20
            .AddCampo , , "D�bito", wrTADireito, 26, 95
            .Campo(6).Formato = FMOEDA
            .AddCampo , , "Cr�dito", wrTADireito, 26, 121
            .Campo(7).Formato = FMOEDA
            .AddCampo , , "Saldo", wrTADireito, 26, 147
            .Campo(8).Formato = FMOEDA
            'pt. 84768 - Ivo Sousa (28/10/2008)
            .AddCampo , , "D�bito/Cr�dito", wrTADireito, 15, 172
        End With
        
        'Se��o de Rodap�: Sub Totais por empresa
        .Grupo(1).AddSecao scFooter, 1, wrDBTopBorder Or wrDBBottomBorder
        With .Grupo(1).Footer.Linha(1)
            .Height = .Height * 2
            .AddCampo , wrCSFixedText, "Total da Empresa:", , 35
            .Campo(1).Top = ((.Height / 2) - (.Campo(1).Height / 2))
            '17/02/2003 - Fabricio
            'Este campo teria de buscar o Nome da Empresa mas gera um erro e trava o Sistema em algumas situa��es
            '.AddCampo , wrCSDataLink, "{*Quebra}", , 27
            '.Campo(2).Top = .Campo(1).Top
            .AddCampo , wrCSTotal, "D�bito", wrTADireito, 26, 95
            .Campo(2).Formato = FMOEDA
            '.Campo(2).Left = wrkRazao(1).Detalhe(1).Campo(5).Left
            .Campo(2).Top = .Campo(1).Top
            .AddCampo , wrCSTotal, "Cr�dito", wrTADireito, 52
            .Campo(3).Formato = FMOEDA
            .Campo(3).Left = wrkRazao(1).Detalhe(1).Campo(6).Left
            .Campo(3).Top = .Campo(1).Top
            .AddCampo , wrCSDataLink, "Saldo", wrTADireito, 54.4
            If TypeOf rstSaldos Is ADODB.Recordset Then
                .Campo(4).TableLink = ExtractTableName(rstSaldos.Source)
            Else
                .Campo(4).TableLink = rstSaldos.name
            End If
            .Campo(4).DataLink = "Apel = {*Quebra} AND Final = True"
            'Projeto: #7373 - Hist�ria: #6134 - Desenvolvimento: #7416 - Ivo Sousa(07/05/2013)
            .Campo(4).Formato = FMOEDA
            .Campo(4).Left = 118.5 'wrkRazao(1).Detalhe(1).Campo(8).Left
            .Campo(4).Top = .Campo(1).Top
        End With
        
        Call GrupoResumo(wrkRazao, rstSaldoGeral)

    End With
    'Pt. 95368 - Moacir Pfau(21/10/2009)
    wrkRazao.BeginPrint gTipoDB
    wrkRazao.EndPrint
    Set wrkRazao = Nothing
    MsgBar Me.Caption
End Sub

Private Sub GrupoResumo(wrkReport As KeybReport, rstSaldoGeral As Object)
    Dim strTable As String
        
    If TypeOf rstSaldoGeral Is ADODB.Recordset Then
        strTable = ExtractTableName(rstSaldoGeral.Source)
    Else
        strTable = rstSaldoGeral.name
    End If
    
    With wrkReport
        .AddGrupo "resumo"
        .Grupo("resumo").AddSecao scHeader, 1
        With .UltimaSecao.Linha(1)
            .AddCampo , wrCSFixedText, "Total Geral", , 30
            
            .AddCampo , wrCSDataLink, "Debito", wrTADireito, 26, 95
            .Campo(2).Formato = FMOEDA
            .Campo(2).TableLink = strTable
            
            .AddCampo , wrCSDataLink, "Credito", wrTADireito, 26, 121
            .Campo(3).Formato = FMOEDA
            .Campo(3).TableLink = strTable
            
            .AddCampo , wrCSDataLink, "Saldo", wrTADireito, 26, 147
            .Campo(4).Formato = FMOEDA
            .Campo(4).TableLink = strTable
        End With
        .FontStyle = wrFSNormal
    End With
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
