VERSION 5.00
Begin VB.Form fcalcDuplLanc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Geração de Controle de Duplicatas e Lançamentos"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "fcalcDuplLanc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancela&r"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Lançamentos e Duplicatas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5775
      Begin VB.Frame Frame2 
         Caption         =   "Vencimento"
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
         TabIndex        =   11
         Top             =   2640
         Width           =   5775
         Begin VB.TextBox txtCalculo 
            Height          =   315
            Index           =   2
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Baixas"
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtCalculo 
            Height          =   315
            Index           =   1
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Baixas"
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblCalculo 
            AutoSize        =   -1  'True
            Caption         =   "&Final:"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblCalculo 
            AutoSize        =   -1  'True
            Caption         =   "&Inicial:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Duplicata/Lançamentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   0
         TabIndex        =   14
         Top             =   840
         Width           =   5775
         Begin VB.ComboBox cboCalculo 
            Height          =   315
            Index           =   0
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtCalculo 
            Height          =   315
            Index           =   3
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   17
            Tag             =   "Baixas"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtCalculo 
            Height          =   315
            Index           =   0
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "Baixas"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtCalculo 
            Height          =   315
            Index           =   5
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Baixas"
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblCalculo 
            AutoSize        =   -1  'True
            Caption         =   "&Tipo de Registro:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblCalculo 
            AutoSize        =   -1  'True
            Caption         =   "&Empresa:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   660
         End
         Begin VB.Label lblDescricao 
            Caption         =   "lblDescricao"
            Height          =   255
            Left            =   3360
            TabIndex        =   19
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblCalculo 
            AutoSize        =   -1  'True
            Caption         =   "&Final:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblCalculo 
            AutoSize        =   -1  'True
            Caption         =   "&Inicial:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   450
         End
      End
      Begin VB.TextBox txtCalculo 
         Height          =   315
         Index           =   4
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   5
         Tag             =   "Baixas"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.ComboBox cboCalculo 
         Height          =   315
         Index           =   1
         ItemData        =   "fcalcDuplLanc.frx":000C
         Left            =   1440
         List            =   "fcalcDuplLanc.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "&Último Controle:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   4080
         Width           =   1110
      End
      Begin VB.Label lblCalculo 
         AutoSize        =   -1  'True
         Caption         =   "&Tabela:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   540
      End
   End
End
Attribute VB_Name = "fcalcDuplLanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSql      As String
Dim rstDuplLanc As Object
Private Sub cmdCancel_Click()
  Unload Me
  Exit Sub
End Sub

Private Sub cmdOk_Click()
    
' If cboCalculo(1).Text = "Duplicatas" Or cboCalculo(1).Text = "Lançamentos" Then
'  If txtCalculo(0).Text = "" Or txtCalculo(5).Text = "" Then
'    MsgBox "Número Inicial e Final devem ser preenchidos!", vbCritical, "Número"
'    txtCalculo(0).SetFocus
'    Exit Sub
'  End If
' End If
 
' If Not IsValid(txtCalculo(1).Text) Then
'  MsgBox "Data de Vencimento Inicial em branco!", vbCritical, "Vencimento Inicial em Branco"
'  txtCalculo(1).SetFocus
'  Exit Sub
' Else
  If IsValid(txtCalculo(1).Text) Then
    If Not EData(txtCalculo(1).Text) Then
      MsgBox "Data Inválida", vbCritical, "Data"
      Exit Sub
    End If
  End If
' End If
 
 'If Not IsValid(txtCalculo(2).Text) Then
 ' MsgBox "Data de Vencimento Final em branco!", vbCritical, "Vencimento Final em Branco"
 ' txtCalculo(2).SetFocus
 ' Exit Sub
 'Else
   If IsValid(txtCalculo(2).Text) Then
      If EData(txtCalculo(2).Text) Then
        If InverteData(txtCalculo(1).Text) > InverteData(txtCalculo(2).Text) Then
          MsgBox "Data de Vencimento Inicial é maior que  Vencimento Final!", vbCritical, "Vencimento"
          Exit Sub
        End If
      Else
        MsgBox "Data Inválida!", vbCritical, "Vencimento"
        Exit Sub
      End If
   End If
 'End If
 
 
 If Not IsValid(txtCalculo(4).Text) Then
  MsgBox "Número de Controle em branco!", vbCritical, "Controle"
  txtCalculo(4).SetFocus
  Exit Sub
 End If
 
 SelecionaRegistros
 AtualizaControle
 
End Sub

Private Sub Form_Load()
  
  Dim TiposCombo As String
  
  CenterForm Me
  cboCalculo(1).Text = "Duplicatas"


  TiposCombo = "SELECT Tipo FROM [Tipos Globais];"
  ComboAddItem cboCalculo(0), TiposCombo, "Tipo"

  cboCalculo(0).AddItem "Todos"
  
  cboCalculo(0).Text = "Todos"
  lblDescricao.Caption = ""
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set fcalcDuplLanc = Nothing
End Sub

Private Sub SelecionaRegistros()

  Dim strTabela As String
  Dim strPagamento As String
  Dim strWhere As String
  Dim strPagRec As String
  
  strTabela = Trim(cboCalculo(1).Text)
  strSql = NUL
  strPagamento = " WHERE Pagamento IS NULL AND (Controle IS NULL OR Controle = '') "
  strPagRec = " AND PagRec = 'R'"
  
  If strTabela = "Duplicatas" Then
    strSql = "SELECT 'Duplicatas' as Tabela,PagRec,Nota as Número,Parcela,Empresa,Tipo,Vencimento,Pagamento,Controle FROM [Duplicatas]" & strPagamento & strPagRec
  ElseIf strTabela = "Lançamentos" Then
    strSql = "SELECT 'Lançamentos' as Tabela,PagRec,Código as Número,'0' as Parcela,Empresa,Tipo,Vencimento,Pagamento,Controle FROM [Lançamentos]" & strPagamento & strPagRec
  End If
  
  If IsValid(txtCalculo(0).Text) Or IsValid(txtCalculo(5).Text) Then   ' Numero da Nota
    If IsValid(txtCalculo(0).Text) And IsValid(txtCalculo(5).Text) Then   ' Numero da Nota
      strWhere = " AND " & IIf(strTabela = "Duplicatas", "Nota", "Código") & " >= " & CLngDef(txtCalculo(0).Text) & " AND " & IIf(strTabela = "Duplicatas", "Nota", "Código") & "<= " & CLngDef(txtCalculo(5).Text)
    ElseIf IsValid(txtCalculo(0).Text) Then
      strWhere = " AND " & IIf(strTabela = "Duplicatas", "Nota", "Código") & " >= " & CLngDef(txtCalculo(0).Text)
    Else
      strWhere = " AND " & IIf(strTabela = "Duplicatas", "Nota", "Código") & "<= " & CLngDef(txtCalculo(5).Text)
    End If
  End If
  
  If cboCalculo(0).Text <> "Todos" Then ' Tipo de Registro
    Concat strWhere, " AND [Tipo] = " & Quote(cboCalculo(0).Text, "'")
  End If
  
  If IsValid(txtCalculo(3).Text) Then ' Empresa
    Concat strWhere, " AND [Empresa] = " & Quote(txtCalculo(3).Text, "'")
  End If
  
  If IsValid(txtCalculo(1).Text) Then  ' Vencimento 1
    Concat strWhere, " AND [Vencimento] >= " & InverteData(txtCalculo(1).Text, True)
  End If

  If IsValid(txtCalculo(2).Text) Then ' Vencimento 2
    Concat strWhere, " AND [Vencimento] <= " & InverteData(txtCalculo(2).Text, True)
  End If

  Concat strSql, strWhere
  If cboCalculo(1).Text = "Duplicatas" Then
    Concat strSql, " ORDER BY Nota, Vencimento"
  Else
    Concat strSql, " ORDER BY Código, Vencimento"
  End If
  
End Sub


Private Sub AtualizaControle()

  'Vinicius Elyseu(24/05/2016) - Projeto: #100340 Demanda: #120791
  Dim nNumeroControle As Double
  Dim strUpdate       As String
  
  'Vinicius Elyseu(24/05/2016) - Projeto: #100340 Demanda: #120791
  nNumeroControle = CDblDef(txtCalculo(4).Text) + 1
  
  If AbreRecordset(rstDuplLanc, strSql, dbOpenDynaset) = WL_OK Then
  
    Do
      
      If GetValue(rstDuplLanc, "Tabela", NUL) = "Duplicatas" Then
      
        strUpdate = "UPDATE Duplicatas SET Controle=" & str(nNumeroControle) & " WHERE PagRec = " & _
                         Quote(GetValue(rstDuplLanc, "PagRec", NUL), "'") & " AND Nota = " & _
                         GetValue(rstDuplLanc, "NÚmero", ZERO) & " AND Empresa = " & _
                         Quote(GetValue(rstDuplLanc, "Empresa", NUL), "'") & " AND Tipo = " & _
                         Quote(GetValue(rstDuplLanc, "Tipo", ZERO), "'") & " AND parcela = " & _
                         GetValue(rstDuplLanc, "Parcela", NUL)

      ElseIf GetValue(rstDuplLanc, "Tabela", NUL) = "Lançamentos" Then
      
        strUpdate = "UPDATE Lançamentos SET Controle=" & str(nNumeroControle) & " WHERE PagRec = " & _
                         Quote(GetValue(rstDuplLanc, "PagRec", NUL), "'") & " AND Código = " & _
                         GetValue(rstDuplLanc, "Número", ZERO)
      
      End If
      ExecuteSQL (strUpdate)
      
      nNumeroControle = nNumeroControle + 1
      
      rstDuplLanc.MoveNext
      
    Loop Until rstDuplLanc.EOF
  Else
    MsgBox "Não há Registros a serem atualizados!!", vbCritical, "Cálculo"
    Exit Sub
  End If
  
  MsgBox "Registros Atualizados com sucesso!!", vbInformation, "Cálculo"
  
End Sub

Private Sub txtCalculo_Change(Index As Integer)
   If Index = 3 Then
     GetAssocValue "Select Razão, Apel from Empresas where Apel = " & Quote(txtCalculo(Index).Text, "'"), lblDescricao
   End If
End Sub

Private Sub txtCalculo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  
  If Shift = 0 And KeyCode = vbKeyPageDown Then
    If Index = 0 Or Index = 5 Then
      If cboCalculo(1).Text = "Duplicatas" Then
        PCampo "Duplicatas", "Select * from Duplicatas where PagRec = 'R'", pbCampo, txtCalculo(Index), "Nota"
      Else
        PCampo "Lançamentos", "Select * from Lançamentos where PagRec = 'R'", pbCampo, txtCalculo(Index), "Código"
      End If
    End If
  
    If Index = 3 Then
      PCampo "Empresas", "Empresas", pbCampo, txtCalculo(3), "Apel"
    End If
     
  End If
End Sub

Private Sub txtCalculo_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 1 Or Index = 2 Then
    SetMascara KeyAscii, txtCalculo(Index).SelStart, MASK_DATE4
  End If

End Sub

Private Sub txtCalculo_LostFocus(Index As Integer)
    If Index = 3 Then
        GetAssocValue "Select Razão, Apel from Empresas where Apel = " & Quote(txtCalculo(Index).Text, "'"), lblDescricao, txtCalculo(Index)
    End If
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
