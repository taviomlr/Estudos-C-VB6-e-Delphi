VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComctl.ocx"
Begin VB.Form frptDuplLancAtraso 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duplicatas e Lançamentos em Atraso - Sintético"
   ClientHeight    =   2010
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   6285
   Icon            =   "frptDuplLancAtraso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6285
   Begin MSComctlLib.ProgressBar Progresso 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1275
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdDuplLancAtraso 
      Caption         =   "&Visualizar"
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   1590
      Width           =   1335
   End
   Begin VB.CommandButton cmdDuplLancAtraso 
      Caption         =   "&Imprimir"
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   1590
      Width           =   1335
   End
   Begin VB.CommandButton cmdDuplLancAtraso 
      Cancel          =   -1  'True
      Caption         =   "Fecha&r"
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Duplicatas e Lançamentos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox cboDuplLancAtraso 
         Height          =   315
         ItemData        =   "frptDuplLancAtraso.frx":000C
         Left            =   960
         List            =   "frptDuplLancAtraso.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtDuplLancAtraso 
         Height          =   315
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblDuplLancAtraso 
         Caption         =   "DescricaoEmpresa"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblDuplLancAtraso 
         Caption         =   "&Tipo:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblDuplLancAtraso 
         Caption         =   "&Empresa:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frptDuplLancAtraso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDuplLancAtraso_Click(Index As Integer)
  If Index < 2 Then
    cmdDuplLancAtraso(0).Enabled = False
    cmdDuplLancAtraso(1).Enabled = False
    cmdDuplLancAtraso(2).Caption = LoadResString(IDS_CANCELAR)

    FiltraDuplLancAtraso

    cmdDuplLancAtraso(0).Enabled = True
    cmdDuplLancAtraso(1).Enabled = True
    cmdDuplLancAtraso(2).Caption = LoadResString(IDS_FECHAR)
  Else
    Unload Me
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
  
  PosForm Me
  cboDuplLancAtraso.Text = "Todos"
  lblDuplLancAtraso(2).Caption = NUL
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SavePosForm Me
  Set frptDuplLancAtraso = Nothing
End Sub

Private Sub FiltraDuplLancAtraso()
  
  Dim rstAux             As Object
  Dim fdsAtraso(11)      As FieldStruct
  Dim strDuplicatas      As String
  Dim strLancamentos     As String
  Dim strEmpresa         As String
  Dim strSql             As String
  Dim rstSql             As Object
  Dim strAux             As String
  Dim X                  As Integer
  Dim strDias            As String
  Dim strDiasFinal       As String
  Dim curSomaDupl        As Currency
  Dim curSomaLanc        As Currency
  Dim nCont              As Long
  Dim strWhere           As String
  Dim strApel            As String
  Dim fakedao            As New CGenericRecordset
  
  If Len(txtDuplLancAtraso.Text) > 0 And Len(lblDuplLancAtraso(2).Caption) = 0 Then
    MsgBox "Empresa não cadastrada!", vbCritical, "Relatório de Duplicatas/Lançamentos em Atraso"
    Exit Sub
  End If
    
  strDuplicatas = "SELECT DISTINCT(Empresa) FROM Duplicatas WHERE Vencimento <" & InverteData(Date, True) & "  AND Pagamento IS NULL AND PAGREC = 'R' AND Situação <> 'Cancelada'"
  strLancamentos = "SELECT DISTINCT(Empresa) FROM Lançamentos WHERE Vencimento <" & InverteData(Date, True) & "  AND Pagamento IS NULL AND PAGREC = 'R' AND Situação <> 'Cancelada'"
  strEmpresa = " AND Empresa=" & Quote(txtDuplLancAtraso.Text, "'")
  
  If IsValid(txtDuplLancAtraso.Text) Then
    If cboDuplLancAtraso.Text = "Duplicatas" Or cboDuplLancAtraso.Text = "Todos" Then
     Concat strDuplicatas, strEmpresa
    End If
    If cboDuplLancAtraso.Text = "Lançamentos" Or cboDuplLancAtraso.Text = "Todos" Then
     Concat strLancamentos, strEmpresa
    End If
  
  End If
  
  strSql = ""
  If cboDuplLancAtraso.Text = "Todos" Then
    Concat strSql, strDuplicatas, " UNION ALL ", strLancamentos
  ElseIf cboDuplLancAtraso.Text = "Duplicatas" Then
    Concat strSql, strDuplicatas
  Else
    Concat strSql, strLancamentos
  End If
  Concat strSql, " ORDER BY Empresa"
  
  On Error GoTo trataErro
  
  If AbreRecordset(rstSql, strSql, dbOpenSnapshot) = WL_OK Then
    
    AppendVar fdsAtraso(0), "Empresa", dbText, 15
    AppendVar fdsAtraso(1), "Ate30", dbCurrency
    AppendVar fdsAtraso(2), "Ate60", dbCurrency
    AppendVar fdsAtraso(3), "Ate90", dbCurrency
    AppendVar fdsAtraso(4), "Ate120", dbCurrency
    AppendVar fdsAtraso(5), "Apos120", dbCurrency
    AppendVar fdsAtraso(6), "Contato", dbText, 50
    AppendVar fdsAtraso(7), "Fone1", dbText, 15
    AppendVar fdsAtraso(8), "Ramal1", dbText, 10
    AppendVar fdsAtraso(9), "foneCobranca", dbText, 15
    AppendVar fdsAtraso(10), "ramalCobranca", dbText, 10
    AppendVar fdsAtraso(11), "razaoEmpresa", dbText, 90

    CrieAux rstAux, fdsAtraso()
    
    nCont = 0
    
    Progresso.Max = Recordcount(strSql)
    
    Do
    
      
     nCont = nCont + 1
     curSomaDupl = 0
     curSomaLanc = 0
     
     For X = 1 To 5
      
        Select Case X
          Case 1
            strDias = "<=30"
          Case 2
            strDias = "<=60"
            strDiasFinal = ">30"
          Case 3
            strDias = "<=90"
            strDiasFinal = ">60"
          Case 4
            strDias = "<=120"
            strDiasFinal = ">90"
          Case 5
            strDias = ">120"
        End Select
        
        If cboDuplLancAtraso.Text = "Duplicatas" Or cboDuplLancAtraso.Text = "Todos" Then
            strWhere = ""
            strWhere = " PagRec = 'R' AND Empresa=" & Quote(GetValue(rstSql, "Empresa", NUL), "'") & _
                 " AND Vencimento <" & InverteData(Date, True) & _
                 " AND Pagamento IS NULL AND Situação <> 'Cancelada'" & _
                 " AND DateDiff(""d"",Vencimento," & InverteData(Date, True) & ")" & strDias
            If X <> 1 And X <> 5 Then
              Concat strWhere, " AND DateDiff(""d"",Vencimento," & InverteData(Date, True) & ")" & strDiasFinal
            End If
            curSomaDupl = Soma("[Valor Original]", "DUPLICATAS", strWhere)
            curSomaDupl = curSomaDupl + Soma("[Acréscimo]", "DUPLICATAS", strWhere)
            curSomaDupl = curSomaDupl - Soma("Abatimento", "DUPLICATAS", strWhere)
        End If
        
        If cboDuplLancAtraso.Text = "Lançamentos" Or cboDuplLancAtraso.Text = "Todos" Then
            strWhere = ""
            strWhere = " PagRec = 'R' AND Empresa=" & Quote(GetValue(rstSql, "Empresa", NUL), "'") & _
                 " AND Vencimento <" & InverteData(Date, True) & _
                 " AND Pagamento IS NULL AND Situação <> 'Cancelada'" & _
                 " AND DateDiff(""d"",Vencimento," & InverteData(Date, True) & ")" & strDias
            If X <> 1 And X <> 5 Then
              Concat strWhere, " AND DateDiff(""d"",Vencimento," & InverteData(Date, True) & ")" & strDiasFinal
            End If
            curSomaLanc = Soma("[Valor Original]", "Lançamentos", strWhere)
            curSomaLanc = curSomaLanc + Soma("[Acréscimo]", "Lançamentos", strWhere)
            curSomaLanc = curSomaLanc - Soma("Abatimento", "Lançamentos", strWhere)
        End If
        
        fakedao.Initialize rstAux
        fakedao.FindFirst " Empresa=" & Quote(GetValue(rstSql, "Empresa", NUL), "'")
        If fakedao.NoMatch Then
          fakedao.AddNew
        Else
          fakedao.Edit
        End If
        
        strApel = GetValue(rstSql, "Empresa", NUL)
        rstAux("Empresa").Value = strApel
        
        Select Case X
          Case 1
            rstAux("Ate30").Value = curSomaDupl + curSomaLanc
          Case 2
            rstAux("Ate60").Value = curSomaDupl + curSomaLanc
          Case 3
            rstAux("Ate90").Value = curSomaDupl + curSomaLanc
          Case 4
            rstAux("Ate120").Value = curSomaDupl + curSomaLanc
          Case 5
            rstAux("Apos120").Value = curSomaDupl + curSomaLanc
        End Select
        
        Call dadosEmpresa(strApel, rstAux)
        fakedao.update
        
      Next
      
      Progresso = nCont
      rstSql.MoveNext
      
    Loop Until rstSql.EOF
    
    fimpDuplLancAtraso.Config rstAux
    DeleteAux rstAux, NUL
    Progresso.Value = 0
  Else
    MsgFunc LoadResString(IDS_RECORDNOTFOUND)
  End If
  FechaRecordset rstSql
  
  
  Set fakedao = Nothing
  Exit Sub
  
trataErro:
  MsgFunc err.Number & err.Description
  
End Sub

Private Sub txtDuplLancAtraso_Change()
  GetAssocValue "Select Razão,Apel From Empresas Where Apel =" & Quote(txtDuplLancAtraso.Text, "'"), lblDuplLancAtraso(2), txtDuplLancAtraso
End Sub

Private Sub txtDuplLancAtraso_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyPageDown And Shift = 0 Then
      PCampo "Empresas", "Empresas", pbCampo, txtDuplLancAtraso, "Apel"
      KeyCode = 0
  End If

End Sub

'Autor: Dulcino Júnior
'Data: 16/11/2006
'Função que preenche os dados da empresa
Private Sub dadosEmpresa(strApel As String, ByRef rst As Object)
    Dim cmd      As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "Razão, Contato, Fone1, Ramal1"
    cmd.Table.TableName = "Empresas"
    Call cmd.Filter.Append("Apel = @pApel")
    Call cmd.Parameters.add(cmd.CreateParameter("@pApel", strApel, dbFieldTypeString, 15))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        rst("razaoEmpresa").Value = rdResult.GetString("Razão")
        rst("Contato").Value = rdResult.GetString("Contato")
        rst("Fone1").Value = rdResult.GetString("Fone1")
        rst("Ramal1").Value = rdResult.GetString("Ramal1")
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Call dadosCobranca(strApel, rst)
    Aplicacao.Disconnect
End Sub

'Autor: Dulcino Júnior
'Data: 16/11/2006
'Função que busca os dados de cobrança
Private Sub dadosCobranca(strApel As String, ByRef rst As Object)
    Dim cmd      As IDBSelectCommand
    Dim rdResult As IDBReader
    
    Aplicacao.Connect
    Set cmd = Aplicacao.CreateSelectCommand
    cmd.SelectClause = "TOP 1 Fone, Ramal"
    cmd.Table.TableName = "[Empresas Endereços]"
    Call cmd.Filter.Append("Apel = @pApel")
    Call cmd.Parameters.add(cmd.CreateParameter("@pApel", strApel, dbFieldTypeString, 15))
    Call cmd.Filter.Append("Tipo = @pTipo")
    Call cmd.Parameters.add(cmd.CreateParameter("@pTipo", "Cobrança", dbFieldTypeString, 13))
    Set rdResult = Aplicacao.ExecuteReader(Aplicacao.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        rst("foneCobranca").Value = rdResult.GetString("Fone")
        rst("ramalCobranca").Value = rdResult.GetString("Ramal")
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Aplicacao.Disconnect
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
