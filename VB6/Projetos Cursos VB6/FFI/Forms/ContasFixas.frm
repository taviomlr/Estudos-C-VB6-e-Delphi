VERSION 5.00
Begin VB.Form frmContasFixas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Contas Fixas"
   ClientHeight    =   5235
   ClientLeft      =   2430
   ClientTop       =   3360
   ClientWidth     =   9615
   Icon            =   "ContasFixas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   9615
   Tag             =   "CFixas"
   Begin VB.TextBox txtCFixas 
      DataField       =   "vencimento_regra_excecao"
      Height          =   285
      Index           =   11
      Left            =   4200
      TabIndex        =   49
      Tag             =   "CFixas"
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   5180
      Left            =   8160
      TabIndex        =   47
      Top             =   30
      Width           =   1410
      Begin VB.CommandButton cmdGerarLancamento 
         Caption         =   "&Gerar Lan�."
         Height          =   375
         Left            =   90
         TabIndex        =   18
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   90
         TabIndex        =   20
         Top             =   1785
         Width           =   1215
      End
      Begin VB.CommandButton cmdNovo 
         Caption         =   "&Novo"
         Height          =   375
         Left            =   90
         TabIndex        =   16
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdGravar 
         Caption         =   "&Gravar"
         Height          =   375
         Left            =   90
         TabIndex        =   17
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton cmdExcluirLanc 
         Caption         =   "&Excluir Lan�."
         Height          =   375
         Left            =   90
         TabIndex        =   19
         Top             =   1380
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   2590
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   90
         TabIndex        =   21
         Top             =   2190
         Width           =   1215
      End
   End
   Begin VB.Frame fraTab 
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   60
      TabIndex        =   23
      Top             =   40
      Width           =   8055
      Begin VB.Frame fraCFixas 
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
         Height          =   1420
         Index           =   0
         Left            =   0
         TabIndex        =   24
         Top             =   -20
         Width           =   8055
         Begin VB.ComboBox cboCFixas 
            DataField       =   "Preced�ncia"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            ItemData        =   "ContasFixas.frx":08CA
            Left            =   6000
            List            =   "ContasFixas.frx":08CC
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Tag             =   "CFixas"
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "Descri��o"
            Height          =   315
            Index           =   9
            Left            =   1200
            TabIndex        =   4
            Tag             =   "CFixas"
            Top             =   960
            Width           =   3615
         End
         Begin VB.ComboBox cboCFixas 
            DataField       =   "Tipo"
            Height          =   315
            Index           =   0
            Left            =   3000
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Tag             =   "CFixas"
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "Empresa"
            Height          =   315
            Index           =   1
            Left            =   1200
            TabIndex        =   3
            Tag             =   "CFixas"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "C�digo"
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   0
            Tag             =   "CFixas"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblCFixas 
            AutoSize        =   -1  'True
            Caption         =   "Proced�nci&a"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   4960
            TabIndex        =   27
            Top             =   270
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lblCFixas 
            AutoSize        =   -1  'True
            Caption         =   "Descr&i��o"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   360
            TabIndex        =   30
            Top             =   960
            Width           =   720
         End
         Begin VB.Label lblCFixas 
            AutoSize        =   -1  'True
            Caption         =   "&Tipo"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   2565
            TabIndex        =   26
            Top             =   270
            Width           =   315
         End
         Begin VB.Label lblContasFixas 
            Caption         =   "lblContasFixas(0)"
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   29
            Top             =   630
            Width           =   4665
         End
         Begin VB.Label lblCFixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Empresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   345
            TabIndex        =   28
            Top             =   630
            Width           =   735
         End
         Begin VB.Label lblCFixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C�di&go"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   25
            Top             =   270
            Width           =   600
         End
      End
      Begin VB.Frame fraCFixas 
         Caption         =   "Datas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Index           =   1
         Left            =   0
         TabIndex        =   31
         Top             =   1410
         Width           =   8055
         Begin VB.Frame fraData 
            Caption         =   "Vencimento - Regra Exce��o"
            Height          =   525
            Left            =   4680
            TabIndex        =   48
            Top             =   650
            Width           =   2535
            Begin VB.OptionButton optProximo 
               Caption         =   "Prorrogar"
               Height          =   195
               Left            =   1320
               TabIndex        =   10
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optAnterior 
               Caption         =   "Antecipar"
               Height          =   195
               Left            =   240
               TabIndex        =   9
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "In�cio"
            Height          =   315
            Index           =   10
            Left            =   1560
            TabIndex        =   5
            Tag             =   "CFixas"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "T�rmino"
            Height          =   315
            Index           =   3
            Left            =   5280
            TabIndex        =   6
            Tag             =   "CFixas"
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cboCFixas 
            DataField       =   "Periodicidade"
            Height          =   315
            Index           =   1
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "CFixas"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "Vencimento"
            Height          =   315
            Index           =   2
            Left            =   1560
            TabIndex        =   8
            Tag             =   "CFixas"
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblContasFixas 
            Caption         =   "lblContasFixas(5)"
            Height          =   195
            Index           =   5
            Left            =   3000
            TabIndex        =   33
            Top             =   270
            Width           =   1545
         End
         Begin VB.Label lblCFixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "In�&cio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   945
            TabIndex        =   32
            Top             =   270
            Width           =   510
         End
         Begin VB.Label lblContasFixas 
            Caption         =   "lblContasFixas(4)"
            Height          =   195
            Index           =   4
            Left            =   6720
            TabIndex        =   35
            Top             =   270
            Width           =   1185
         End
         Begin VB.Label lblCFixas 
            AutoSize        =   -1  'True
            Caption         =   "T�&rmino"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   4545
            TabIndex        =   34
            Top             =   270
            Width           =   690
         End
         Begin VB.Label lblCFixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Periodicidade"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   495
            TabIndex        =   36
            Top             =   630
            Width           =   960
         End
         Begin VB.Label lblCFixas 
            AutoSize        =   -1  'True
            Caption         =   "Dia do &vencimento"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   990
            Width           =   1335
         End
      End
      Begin VB.Frame fraCFixas 
         Caption         =   "Informa��es Financeiras"
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
         Index           =   2
         Left            =   0
         TabIndex        =   38
         Top             =   2820
         Width           =   8055
         Begin VB.TextBox txtCFixas 
            DataField       =   "Valor"
            Height          =   315
            Index           =   8
            Left            =   5070
            TabIndex        =   15
            Tag             =   "CFixas"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "Controle"
            Height          =   315
            Index           =   7
            Left            =   1230
            TabIndex        =   14
            Tag             =   "CFixas"
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "Centro"
            Height          =   315
            Index           =   6
            Left            =   1230
            TabIndex        =   13
            Tag             =   "CFixas"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "Conta"
            Height          =   315
            Index           =   5
            Left            =   1230
            TabIndex        =   12
            Tag             =   "CFixas"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtCFixas 
            DataField       =   "Banco"
            Height          =   315
            Index           =   4
            Left            =   1230
            TabIndex        =   11
            Tag             =   "CFixas"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblCFixas 
            AutoSize        =   -1  'True
            Caption         =   "Va&lor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   4590
            TabIndex        =   46
            Top             =   1350
            Width           =   450
         End
         Begin VB.Label lblCFixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C&ontrole"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   540
            TabIndex        =   45
            Top             =   1350
            Width           =   585
         End
         Begin VB.Label lblContasFixas 
            Caption         =   "lblContasFixas(3)"
            Height          =   195
            Index           =   3
            Left            =   2670
            TabIndex        =   44
            Top             =   990
            Width           =   5265
         End
         Begin VB.Label lblCFixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "C. C&usto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   390
            TabIndex        =   43
            Top             =   990
            Width           =   735
         End
         Begin VB.Label lblContasFixas 
            Caption         =   "lblContasFixas(2)"
            Height          =   195
            Index           =   2
            Left            =   2670
            TabIndex        =   42
            Top             =   630
            Width           =   5265
         End
         Begin VB.Label lblCFixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Co&nta Financ."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   45
            TabIndex        =   41
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label lblContasFixas 
            Caption         =   "lblContasFixas(1)"
            Height          =   195
            Index           =   1
            Left            =   2670
            TabIndex        =   40
            Top             =   270
            Width           =   5265
         End
         Begin VB.Label lblCFixas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "&Banco"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   570
            TabIndex        =   39
            Top             =   270
            Width           =   555
         End
      End
   End
   Begin VB.Image imgInformativa 
      Height          =   480
      Left            =   120
      Picture         =   "ContasFixas.frx":08CE
      Top             =   4680
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"ContasFixas.frx":1510
      Height          =   495
      Left            =   70
      TabIndex        =   50
      Top             =   4680
      Width           =   8020
   End
End
Attribute VB_Name = "frmContasFixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrstCFixas      As Object     '// Abre a tabela de Contas Fixas
Private mlngCFixas      As Long          '// Controle de altera��es do usu�rio
Private mblnNaoAltera   As Boolean
Private mIntTipoConta   As Integer
Private mintProcedencia As Integer

' FUNCTION..: LibProc
' Objetivo..: Fun��o de retorno de chamada da Lib.
' Argumentos: [sFuncao]: Fun��o que deve ser executada.
'             [lFuncao]: Par�metro adicional, varia conforme a fun��o.
' Retorna...: True se executar a fun��o com sucesso, False, se n�o.
' ----------------------------------------------------------------------------------------------
Public Function LibProc(sFuncao As String, Optional lFuncao As Long) As Boolean
    Dim strExibir        As String
    Dim lCFixa           As Long
    Dim intContRegistros As Integer
    Dim blnVencProximo   As Boolean
    
    Select Case sFuncao
  
        'Bot�o Novo
        Case WL_NOVO
            frmContasFixas.cmdSair.Enabled = False
            LibProc = (LimpaControles(mrstCFixas, Me, Tag, mlngCFixas) = WL_OK)
            FirstFocus txtCFixas(0)
            cmdExcluirLanc.Enabled = False
            cmdExcluir.Enabled = False
            cmdGerarLancamento.Enabled = False
            frmContasFixas.cmdSair.Enabled = True
        'Bot�o Excluir
        Case WL_DELETAR
            'pt. 81971 - Ivo Sousa(30/04/2008)
            If Not ExisteLancamentos Then
                lCFixa = GetValue(mrstCFixas, "C�digo", 0)
                If (DeletaRegistro(mrstCFixas, Me, Tag, mlngCFixas) = WL_OK) Then
                    'Exclui todos os registros da tabela de Gera��es Fixas que tem
                    'o mesmo C�digo desta conta
                    Call DeleteAll("Gera��es Fixas", "Conta = " & CStr(lCFixa))
                End If
            Else
                MsgBox "A Conta Fixa possui lan�amentos e n�o pode ser excluida.", vbInformation + vbOKOnly, NomeModulo
            End If

        'Bot�o Localizar
        Case WL_LOCALIZAR
            localizar mrstCFixas, Me, "SELECT * FROM [Contas Fixas] WHERE [Preced�ncia] = " & mintProcedencia, Tag, mlngCFixas

        'Bot�o Pesquisar
        Case WL_PESQUISAR
            PRegistro mrstCFixas, Me, "Contas Fixas", "SELECT * FROM [Contas Fixas] WHERE [Preced�ncia] = " & mintProcedencia, NUL, Tag, mlngCFixas, PB_REGISTRO

        'Bot�o Primerio Registro
        Case WL_PRIMEIRO, WL_ANTERIOR, WL_PROXIMO, WL_ULTIMO
            MoveRecordset mrstCFixas, Me, Tag, mlngCFixas, lFuncao
    
        'Bot�o Sair
        Case WL_SAIR
            Unload Me
            Exit Function

        'Bot�o Navegar
        Case WL_NAVEGAR
            Browse mrstCFixas, Me, Tag, mlngCFixas, "Contas Fixas"

        'Bot�o Salvar
        Case WL_SALVAR
            If (VerContasFixas()) Then
                If txtCFixas(11).Text = "" Then
                    'pt. 86144 - Ivo Sousa(02/05/2008)
                    If optAnterior.value Then
                        txtCFixas(11).Text = "A"
                    Else
                        txtCFixas(11).Text = "P"
                    End If
                End If
                LibProc = (SalvaRegistro(mrstCFixas, Me, Tag, mlngCFixas) = WL_OK)
                cmdExcluir.Enabled = True
                cmdGerarLancamento.Enabled = True
            End If

        'Bot�o Cancelar
        Case WL_CANCELAR
            CancelaEdicao mrstCFixas, Me, Tag, mlngCFixas

        'Op��o Exibir
        Case WL_EXIBIR
            If GetFieldValue("C�digo", "[Contas Fixas]", "C�digo = " & txtCFixas(0).Text & " AND [Preced�ncia] = " & mintProcedencia, , 0) > 0 Then
                strExibir = "SELECT * FROM [Contas Fixas] WHERE C�digo = {C�digo} AND [Preced�ncia] = " & mintProcedencia & ";"
                If RetornaRegs(mrstCFixas, Me, Tag, strExibir, mlngCFixas) = WL_OK Then
                    cmdExcluir.Enabled = True
                    cmdGerarLancamento.Enabled = True
                    If ExisteLancamentos Then
                        cmdExcluirLanc.Enabled = True
                    End If
                End If
            End If
            
        'Op��o Filtrar
        Case WL_FILTRAR
            Filtrar mrstCFixas, Me, Tag, "Contas Fixas", mlngCFixas
            
        'Registro Duplicado
        Case WL_DUPLICADO
            ResolveDuplicacao Me, txtCFixas(0)
        
        'Cadastro de Empresas
        Case "empresas"
            If (KeybAcesso(LoadResString(2037))) Then
                frmEmpresas.Show
                CallChange frmEmpresas.hWnd, txtCFixas(1).hWnd
                Exit Function
            End If

        'Cadastro de Bancos
        Case "bancos"
            If (KeybAcesso(LoadResString(2003))) Then
                frmBancos.Show
                CallChange frmBancos.hWnd, txtCFixas(4).hWnd
                Exit Function
            End If

        'Cadastro de Contas
        Case "contas"
            If (KeybAcesso(LoadResString(2007))) Then
                frmContas.Show
                CallChange frmContas.hWnd, txtCFixas(5).hWnd
                Exit Function
            End If

        'Cadastro de Centro de Custo
        Case "custos"
            If (KeybAcesso(LoadResString(2029))) Then
                frmCusto.Show
                CallChange frmCusto.hWnd, txtCFixas(6).hWnd
                Exit Function
            End If

        'Gerar Lan�amento
        Case "gerar", "gerartodas"
            SetPtrWait Me
            strExibir = "SELECT * FROM [Contas Fixas] WHERE C�digo = " & txtCFixas(0).Text
          
            If (EstaEditando(mlngCFixas)) Then            'Verifica se o usu�rio est� editando
                If (MsgFunc(LoadResString(250), vbQuestion Or vbYesNo) = vbYes) Then
                    If (LibProc(WL_SALVAR, ZERO)) Then
                        LibProc "gerar", ZERO
                    End If
                End If
            ElseIf (Not IsVisibleRecord(mlngCFixas)) And sFuncao = "gerar" Then '// Verifica se h� um registro atual
                MsgFunc LoadResString(251)
            Else 'Gera os lan�amentos para esta conta
                If sFuncao = "gerar" Then
                    If (Not IsEmptyDate(GetValue(mrstCFixas, "T�rmino", Empty))) Then
                        'pt. 86144 - Ivo Sousa(02/05/2008)
                        If optAnterior.value Then
                            blnVencProximo = False
                        Else
                            blnVencProximo = True
                        End If
                        Call GerarContasFixas(strExibir, GCF_TODOS, , intContRegistros, blnVencProximo)
                        'pt. 81971 - Ivo Sousa(02/05/2008)
                        If intContRegistros > 0 Then
                            MsgBox "Foram gerados " & intContRegistros & " lan�amentos.", vbInformation + vbOKOnly, NomeModulo
                            If Not ExisteLancamentos Then
                                cmdExcluirLanc.Enabled = False
                            Else
                                cmdExcluirLanc.Enabled = True
                            End If
                        Else
                            MsgBox "N�o foi possivel gerar os lan�amentos.", vbInformation + vbOKOnly, NomeModulo
                        End If
                    Else 'If (IsEmptyDate(GetValue(mrstCFixas, "T�rmino", Empty))) Then
                        Call GerarContasFixas(strExibir, GCF_UNICO)
                    End If
                Else
                    'caso o usu�rio informe para gerar todos os lan�amentos de todas as contas o sistema pede uma data para gera��o
                    'de contas que n�o tem data de t�rmino mas precisam ser geradas para meses posteriores.
                    Dim DataTermino     As String
                    
                    DataTermino = InputBox("Data final para gera��o...:", "Informe a data Final para gera��o", LastDay(Date))
                    If IsValid(DataTermino) Then
                        If EData(DataTermino) Then
                            GerarContasFixas NUL, GCF_TODOS, CDateDef(DataTermino)
                        Else
                            MsgFunc "Data informada n�o � v�lida."
                            SetPtrDef Me
                            LibProc = False:    Exit Function
                        End If
                    Else
                        MsgFunc "Data n�o informada."
                        SetPtrDef Me
                        LibProc = False:      Exit Function
                    End If
                End If
            End If
            SetPtrDef Me
        
        'Visualizar a gera��o
        Case "ver"
            'Abre a janela de pesquisa com todos os Lan�amentos gerados a partir da
            'conta atual
            If (IsVisibleRecord(mlngCFixas)) Then
                strExibir = "SELECT * FROM [Gera��es Fixas] WHERE Conta = " & txtCFixas(0).Text & ";"
                PCampo "Gera��es", strExibir, ZERO, Nothing, NUL
            Else
                MsgFunc LoadResString(251)
            End If

        ' Configura��o do Cadastro
        Case "Config"
            If KeybAcesso(LoadResString(2104)) Then
                FrmConfCad.Configura "Contas Fixas"
                FrmConfCad.Show vbModal
            End If
          
        Case "atualizarlanc"
            If IsValid(GetValue(mrstCFixas, "C�digo", ZERO)) And IsValid(GetValue(mrstCFixas, "Preced�ncia", ZERO)) Then
                fdlgTaxaSobreConta.mlngConta = GetValue(mrstCFixas, "C�digo", ZERO)
                fdlgTaxaSobreConta.Show vbModal
            End If
    End Select
End Function

Private Sub cboCFixas_Click(Index As Integer)
    If Index = 2 Then
        If Not mblnNaoAltera Then
            AlteraValor mlngCFixas
        End If
    Else
        AlteraValor mlngCFixas
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LibProc(WL_CANCELAR)
End Sub

Private Sub cmdExcluir_Click()
    Call LibProc(WL_DELETAR)
End Sub

Private Sub cmdExcluirLanc_Click()
    Call DeletaLancamentos
End Sub

Private Sub cmdGerarLancamento_Click()
    Call LibProc("gerar")
End Sub

Private Sub cmdGravar_Click()
    Call LibProc(WL_SALVAR)
End Sub

Private Sub cmdNovo_Click()
    Call Configure(mintProcedencia - 1)
End Sub

Private Sub cmdSair_Click()
    Call LibProc(WL_SAIR)
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
    GetKeyDown Me, KeyCode, Shift
End Sub

Private Sub Form_Load()
    LoadMenuTitulos Me
    ConfigCampos Me, "Contas Fixas", Tag      'Configura os controles pela estrutura da tabela
    
    ComboAddItem cboCFixas(0), "SELECT * FROM Op��es WHERE Rotina = '" & _
                               OPT_LANCAMENTOS & "';", "Texto"
                               
    'Projeto: #218 - Hist�ria: #268 - Desenvolvimento#592 - Moacir Pfau(23/09/2012)
    LoadResOptions 1027, cboCFixas(1), True, 0  'Carrega a lista de op��es para Periodicidade
    LoadResOptions 1028, cboCFixas(2), True, 0  'Carrega a lisat de op��es para Preced�ncia
    mlngCFixas = 0
    'Limpando os Labels de descri��o
    lblContasFixas(0).Caption = NUL
    lblContasFixas(1).Caption = NUL
    lblContasFixas(2).Caption = NUL
    lblContasFixas(3).Caption = NUL
    lblContasFixas(4).Caption = NUL
    lblContasFixas(5).Caption = NUL
    
    'Configurando o campo Centro de Custos
    If (Not CentrodeCusto(MFinanceiro)) Then
        lblCFixas(8).Enabled = False               'Label do campo Centro
        txtCFixas(6).Enabled = False               'Campo Centro
        lblContasFixas(3).Enabled = False
    End If
    'Acrescenta os bot�es de gera��o e ver gera��es na barra de ferramentas
    DoEvents
    Call AddToolbarButton("gerar", "Gerar Lan�amento da Conta Fixa Atual", NUL, 505, ATB_IMAGERES Or ATB_IMAGEICON Or ATB_SEPBEFORE)
    Call AddToolbarButton("ver", "Visualizar L�t. gerados pela Conta Fixa Atual", NUL, 506, ATB_IMAGERES Or ATB_IMAGEICON Or ATB_SEPBEFORE)
    Call AddToolbarButton("gerartodas", "Gerar Lan�amentos de todas as Contas Fixas", NUL, IDI_GLOBREC, ATB_IMAGERES Or ATB_IMAGEICON Or ATB_SEPAFTER)
    Call AddToolbarButton("atualizarlanc", "Atualizar lan�amentos da conta fixa com Taxa", NUL, IDI_OPIDATA, ATB_IMAGERES Or ATB_IMAGEICON Or ATB_SEPAFTER)
    
    DefineAcesso mlngCFixas, Acesso()
    'LibProc WL_NOVO, ZERO
    cmdExcluir.Enabled = False
    cmdExcluirLanc.Enabled = False
    cmdGerarLancamento.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = UnloadForm(mrstCFixas, Me, Tag, mlngCFixas)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteToolbarButton "gerar", ATB_SEPBEFORE       '// Exclui o bot�o criado para este form
    DeleteToolbarButton "ver"
    DeleteToolbarButton "gerartodas", ATB_SEPAFTER
    DeleteToolbarButton "atualizarlanc", ATB_SEPAFTER
    
    Set frmContasFixas = Nothing
End Sub

Private Sub optAnterior_Click()
    txtCFixas(11).Text = "A"
End Sub

Private Sub optProximo_Click()
    txtCFixas(11).Text = "P"
End Sub

Private Sub txtBuscaDiasUteis_Change()
End Sub

Private Sub txtCFixas_Change(Index As Integer)
    Select Case (txtCFixas(Index).DataField)
        Case "Empresa"
            GetAssocValue "SELECT Raz�o, Apel FROM Empresas WHERE Apel LIKE '" & _
            txtCFixas(Index).Text & "';", lblContasFixas(0), txtCFixas(Index)
        Case "In�cio"
            lblContasFixas(5).Caption = Semana(txtCFixas(Index).Text)
        Case "T�rmino"
            lblContasFixas(4).Caption = Semana(txtCFixas(Index).Text)
        Case "Banco"
            GetAssocValue "SELECT Nome FROM Bancos WHERE Banco = " & _
            txtCFixas(Index).Text & ";", lblContasFixas(1)
        Case "Conta"
            GetAssocValue "SELECT Descri��o FROM Contas WHERE C�digo = " & _
            txtCFixas(Index).Text & ";", lblContasFixas(2)
        Case "Centro"
            GetAssocValue "SELECT Descri��o FROM Centros WHERE C�digo = " & _
            txtCFixas(Index).Text & ";", lblContasFixas(3)
        Case "vencimento_regra_excecao"
            If txtCFixas(Index).Text = "A" Then
                optAnterior.value = True
            ElseIf txtCFixas(Index).Text = "P" Then
                optProximo.value = True
            End If
    End Select
    If (Index > ZERO) Then
        AlteraValor mlngCFixas
    End If
End Sub

Private Sub txtCFixas_GotFocus(Index As Integer)
    Selecione txtCFixas(Index)
    MsgBar DescCampo(mrstCFixas, txtCFixas(Index).DataField)
End Sub

Private Sub txtCFixas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If (Index = 0) Then
        ControlaChave KeyCode, Shift, txtCFixas(Index), mlngCFixas
    ElseIf ((KeyCode = vbKeyPageDown) And (Shift = ZERO)) Then
        Select Case (txtCFixas(Index).DataField)
            Case "Empresa"
                PCampo "Empresas", "Empresas", PB_CAMPO, txtCFixas(Index), "Apel"
            Case "Banco"
                PCampo "Bancos", "Bancos", PB_CAMPO, txtCFixas(Index), "Banco"
            Case "Conta"
                'pt. 83864 - Dulcino J�nior (11/10/2007)
                PCampo "Contas", "SELECT Contas.C�digo as Conta, Contas.Descri��o as [Descri��o da Conta], Grupos.C�digo as Grupo, Grupos.Descri��o as [Descri��o do Grupo] " & _
                               " FROM Grupos INNER JOIN Contas ON Grupos.C�digo = Contas.Grupo where Contas.Ctaati='S' " & _
                               " ORDER BY Grupos.C�digo,Contas.C�digo", PB_CAMPO, txtCFixas(Index), "Conta"
            Case "Centro"
                PCampo "Centro de Custo", "Centros", PB_CAMPO, txtCFixas(Index), "C�digo"
        End Select
    End If
End Sub

Private Sub txtCFixas_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case (txtCFixas(Index).DataField)
        Case "Empresa"
            SetMascara KeyAscii, txtCFixas(Index).SelStart, MaskEmpresa
        Case "C�digo"
            SetMascara KeyAscii, txtCFixas(Index).SelStart, InputMask(mrstCFixas, "C�digo")
        Case "Vencimento"
            SetMascara KeyAscii, txtCFixas(Index).SelStart, InputMask(mrstCFixas, "Vencimento")
        Case "T�rmino", "In�cio"
            SetMascara KeyAscii, txtCFixas(Index).SelStart, MASK_DATA
        Case "Banco"
            SetMascara KeyAscii, txtCFixas(Index).SelStart, fMask("Bancos", "Banco")
        Case "Conta"
            SetMascara KeyAscii, txtCFixas(Index).SelStart, fMask("Contas", "C�digo")
        Case "Centro"
            SetMascara KeyAscii, txtCFixas(Index).SelStart, fMask("Custos", "C�digo")
        Case "Valor"
            DMoeda KeyAscii
    End Select
End Sub

Private Sub txtCFixas_LostFocus(Index As Integer)
    'Projeto: #8404 - Hist�ria: #9679 - Desenvolvimento#9869 - Jo�o Henrique(02/07/2013)
    Dim strProcura   As String
    Dim rstBanco     As Object

    If (CompStr(txtCFixas(Index).DataField, "C�digo")) Then
        If txtCFixas(Index).Text <> Empty Then
            LibProc WL_EXIBIR
        End If
    ElseIf Index = 1 Then
        txtCFixas(Index).Text = FormataEmpresa(txtCFixas(Index).Text)
        'Projeto: #8404 - Hist�ria: #9679 - Desenvolvimento#9869 - Jo�o Henrique(02/07/2013)
        If EAdicao(mlngCFixas) Or (Not EAdicao(mlngCFixas) And strToLng(txtCFixas(4).Text) = 0 And strToLng(txtCFixas(5).Text) = 0) Then
            strProcura = "SELECT Banco, Conta FROM Empresas WHERE Apel = '" & txtCFixas(1).Text & "';"
            AbreRecordset rstBanco, strProcura
            txtCFixas(4).Text = strToLng(GetValue(rstBanco, "Banco"))
            txtCFixas(5).Text = strToLng(GetValue(rstBanco, "Conta"))
            FechaRecordset (rstBanco)
        End If
    End If
End Sub

' FUNCTION..: VerContasFixas
' Objetivo..: Verificar os dados cadastrados pelo usu�rio.
' Retorna...: Retorna True se estiver tudo correto, False se n�o.
' ------------------------------------------------------------------------------
Private Function VerContasFixas() As Boolean

    'pt. 72415 - Ivo Sousa(30/04/2008)
    'Valida a Empresa
    If txtCFixas(1).Text <> "" And txtCFixas(1).Text <> "0" And lblContasFixas(0).Caption = "" Then
        MsgBox "A Empresa informada n�o � valida.", vbInformation + vbOKOnly, NomeModulo
        txtCFixas(1).SetFocus
        Exit Function
    ElseIf txtCFixas(1).Text = "" Or txtCFixas(1).Text = "0" Then
        MsgBox "O campo Empresa deve ser informado.", vbInformation + vbOKOnly, NomeModulo
        txtCFixas(1).SetFocus
        Exit Function
    End If

    'Verificando as datas digitadas pelo usu�rio
    '�ndice 10 == Data de In�cio
    '�ndice  3 == Data de T�rmino
    If ((IsValid(txtCFixas(10).Text)) And (IsValid(txtCFixas(3).Text))) Then
        If ((EData(txtCFixas(10).Text)) And (EData(txtCFixas(3).Text))) Then
            'Verifica se a data de t�rmino � inferior a data de in�cio
            If (DateDiff(DD_DIA, txtCFixas(10).Text, txtCFixas(3).Text) < ZERO) Then
                MsgFunc LoadResString(252)
                Exit Function
            End If
        End If
    End If

    'Verificando se o Banco mensionado existe no cadastro de Bancos
    If ((IsValid(txtCFixas(4).Text)) And (Len(lblContasFixas(1).Caption) = ZERO)) Then
        If (vbYes = MsgFunc(ResolveResString(IDS_DADONAOENCONTRADO, resUM, txtCFixas(4).Text, _
            resDOIS, "Bancos"), vbQuestion Or vbYesNo)) Then
            Call LibProc("bancos")
        End If
        Exit Function
    'pt. 72415 - Ivo Sousa(30/04/2008)
    ElseIf txtCFixas(4).Text = "" Or txtCFixas(4).Text = "0" Then
        MsgBox "O campo Banco deve ser informado.", vbInformation + vbOKOnly, NomeModulo
        txtCFixas(4).SetFocus
        Exit Function
    End If

    'Verificando o c�digo da Conta Cont�bil
    If ((IsValid(txtCFixas(5).Text)) And (Len(lblContasFixas(2).Caption) = ZERO)) Then
        If (vbYes = MsgFunc(ResolveResString(IDS_DADONAOENCONTRADO, resUM, txtCFixas(5).Text, _
            resDOIS, "Contas"), vbQuestion Or vbYesNo)) Then
            Call LibProc("contas")
        End If
        Exit Function
    'pt. 72415 - Ivo Sousa(30/04/2008)
    ElseIf txtCFixas(5).Text = "" Or txtCFixas(5).Text = "0" Then
        MsgBox "O campo Conta Financeira deve ser informado.", vbInformation + vbOKOnly, NomeModulo
        txtCFixas(5).SetFocus
        Exit Function
    End If

    'Verificar se conta � ativa ou nao
    If txtCFixas(5).Text <> "" Then
        If GetFieldValue("Ctaati", "Contas", " [C�digo]=" & txtCFixas(5).Text) = "N" Then
            MsgBox "Conta " & txtCFixas(5).Text & " n�o est� ativa", vbCritical, MsgBoxCaption
            txtCFixas(5).SetFocus
            Exit Function
        End If
    End If

    'Verifica o c�digo do Centro de Custo
    If ((IsValid(txtCFixas(6).Text)) And (Len(lblContasFixas(3).Caption) = ZERO)) Then
        If (vbYes = MsgFunc(ResolveResString(IDS_DADONAOENCONTRADO, resUM, txtCFixas(6).Text, _
            resDOIS, "Centro de Custo"), vbQuestion Or vbYesNo)) Then
            Call LibProc("custos")
        End If
        Exit Function
    'pt. 72415 - Ivo Sousa(30/04/2008)
    ElseIf txtCFixas(6).Enabled And (txtCFixas(6).Text = "" Or txtCFixas(6).Text = "0") Then
        MsgBox "O campo Centro de Custo deve ser informado.", vbInformation + vbOKOnly, NomeModulo
        txtCFixas(6).SetFocus
        Exit Function
    End If
    'Valor da Conta
    If Not IsValid(txtCFixas(8).Text) Then
        MsgBox "Valor da Conta deve ser preenchido.", vbOKOnly + vbInformation, MsgBoxCaption
        txtCFixas(8).SetFocus
        Exit Function
    End If
    'Data de termino
    If txtCFixas(3).Text = "" Then
        MsgBox "A Data de Termino deve ser preenchida.", vbOKOnly + vbInformation, MsgBoxCaption
        txtCFixas(3).SetFocus
        Exit Function
    End If
    
    'pt. 86728 - Moacir Pfau(09/06/2008)
     If Not (fEmpresaBloqueada(txtCFixas(1).Text, CDate(Format(Now, "DD/MM/YYYY")))) Then
        Exit Function
     End If
     
    VerContasFixas = True
End Function

Public Sub Configure(intIndex As Integer)
    mblnNaoAltera = True
    mIntTipoConta = intIndex
    'pt. 86770 - Ivo Sousa(06/05/2008)
    mintProcedencia = intIndex + 1
    AbreRecordset mrstCFixas, "SELECT * FROM [Contas Fixas] WHERE [Preced�ncia] = " & mintProcedencia
    Call LibProc(WL_NOVO)
    cboCFixas(2).ListIndex = intIndex  ' ListIndex = 0  - A Pagar  ListIndex = 1 - A Receber
    mblnNaoAltera = False
End Sub

'Date.......: 02/05/2008
'Author.....: Ivo Sousa
'Descri��o..: Deleta os lan�amento da conta fixa que n�o tiverem data de pagamento
Private Sub DeletaLancamentos()
    Dim strSql          As String
    Dim rsLancamentos   As Object
    Dim intContReg      As Integer
    Dim strPagRec       As String
    'Projeto: 100340 - Desenv.: 143674 - Ueder Budni (26/09/2016)
    Dim objLogLancDup   As clsLogLancamentosDuplicatas
    Dim strEmpresa      As String
    Dim strTipo         As String
    
    
    Set objLogLancDup = New clsLogLancamentosDuplicatas
    
    If txtCFixas(0).Text <> 0 And txtCFixas(0).Text <> "0" Then
        strSql = "SELECT C�digo FROM [Gera��es Fixas] WHERE Conta = " & txtCFixas(0).Text
        If AbreRecordset(rsLancamentos, strSql) = WL_OK Then
            With rsLancamentos
                If cboCFixas(2).Text = "A Receber" Then
                    strPagRec = "R"
                Else
                    strPagRec = "P"
                End If
                .MoveFirst
                While Not .EOF
                    strEmpresa = GetFieldValue("Empresa", "Lan�amentos", "PagRec = '" & strPagRec & "' AND C�digo = " & rsLancamentos("C�digo").value & " AND Parcela = 1 AND Pagamento IS NULL")
                    strTipo = GetFieldValue("Tipo", "Lan�amentos", "PagRec = '" & strPagRec & "' AND C�digo = " & rsLancamentos("C�digo").value & " AND Parcela = 1 AND Pagamento IS NULL")
                    If ExecuteSQL("DELETE * FROM Lan�amentos WHERE PagRec = '" & strPagRec & "' AND C�digo = " & rsLancamentos("C�digo").value & " AND Parcela = 1 AND Pagamento IS NULL") Then
                        'Projeto: 100340 - Desenv.: 143674 - Ueder Budni (26/09/2016)
                        With objLogLancDup
                            Call .SetKey(strPagRec, rsLancamentos("C�digo").value, strEmpresa, strTipo, 1, enuLancDup.Lancamento)
                            Call .InsertMsg("T�tulo exclu�do atrav�s da rotina de Gera��o de Contas Fixas.")
                        End With
                        
                        Call ExecuteSQL("DELETE * FROM [Gera��es Fixas] WHERE Conta = " & txtCFixas(0).Text & " AND C�digo = " & rsLancamentos("C�digo").value)
                        intContReg = intContReg + 1
                    End If
                    .MoveNext
                Wend
                If intContReg > 0 Then
                    Call ExecuteSQL("UPDATE [Contas Fixas] SET Gera��o = NULL WHERE C�digo = " & txtCFixas(0).Text)
                    MsgBox "Foram exclu�dos " & intContReg & " lan�amentos.", vbInformation + vbOKOnly, NomeModulo
                    If ExisteLancamentos Then
                        cmdExcluirLanc.Enabled = True
                    Else
                        cmdExcluirLanc.Enabled = False
                    End If
                Else
                    MsgBox "N�o foi poss�vel excluir nenhum lan�amento.", vbInformation + vbOKOnly, NomeModulo
                End If
            End With
        Else
            MsgBox "N�o h� lan�amentos para excluir.", vbInformation + vbOKOnly, NomeModulo
        End If
    End If
    'Projeto: 100340 - Desenv.: 143674 - Ueder Budni (26/09/2016)
    Set objLogLancDup = Nothing
End Sub

'Date.......: 02/05/2008
'Author.....: Ivo Sousa
'Descri��o..: Verifica se existem lan�amento vinculados a conta.
Private Function ExisteLancamentos() As Boolean
    Dim lngLancamento As Double
    
    lngLancamento = GetFieldValue("C�digo", "[Gera��es Fixas]", "Conta = " & txtCFixas(0).Text, , 0)
    If lngLancamento = 0 Then
        ExisteLancamentos = False
    Else
        ExisteLancamentos = True
    End If
End Function
