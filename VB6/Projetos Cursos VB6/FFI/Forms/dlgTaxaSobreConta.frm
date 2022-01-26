VERSION 5.00
Begin VB.Form fdlgTaxaSobreConta 
   KeyPreview      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Adicionar taxa nos Lançamentos da Conta Fixa"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancela&r"
      Height          =   330
      Left            =   3480
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame fraPrincipal 
      Caption         =   "Conta"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtContas 
         Height          =   315
         Index           =   3
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtContas 
         Height          =   315
         Index           =   4
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame fraPrinc 
         Caption         =   "Taxa"
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
         Height          =   855
         Left            =   0
         TabIndex        =   13
         Top             =   2160
         Width           =   4335
         Begin VB.TextBox txtContas 
            Height          =   315
            Index           =   2
            Left            =   840
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblForm 
            Caption         =   "Taxa:"
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
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
         ForeColor       =   &H80000002&
         Height          =   1215
         Left            =   0
         TabIndex        =   14
         Top             =   1080
         Width           =   4335
         Begin VB.TextBox txtContas 
            Height          =   315
            Index           =   0
            Left            =   840
            TabIndex        =   5
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtContas 
            Height          =   315
            Index           =   1
            Left            =   840
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblForm 
            Caption         =   "Inicial:"
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblForm 
            Caption         =   "Final:"
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Label lblDescConta 
         Caption         =   "Conta Final"
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblDescConta 
         Caption         =   "Conta Inicial"
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblForm 
         Caption         =   "Inicial:"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblForm 
         Caption         =   "Final:"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
   End
End
Attribute VB_Name = "fdlgTaxaSobreConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngConta     As Long

Private Sub cmdCancel_Click()
  Unload Me
  Exit Sub
End Sub

Private Sub cmdOk_Click()
  Dim strContas         As String
  Dim strData           As String
  Dim dblTaxa           As Double
  Dim PagRec            As String
  
  strContas = NUL
  If IsValid(txtContas(3).Text) Or IsValid(txtContas(4).Text) Then
    If IsValid(txtContas(3).Text) And IsValid(lblDescConta(3).Caption) Then
      strContas = " (Código >= " & CLngDef(txtContas(3).Text) & ") "
    End If
    If IsValid(txtContas(4).Text) And IsValid(lblDescConta(4).Caption) Then
      Concat strContas, IIf(IsValid(strContas), " AND ", "") & " (Código <= " & CLngDef(txtContas(4).Text) & ") "
    End If
  Else
    If MsgFunc("Não foi informado nenhuma Conta (Inicial e Final) para Atualização." & vbCrLf & _
         "Deseja alterar os Lançamentos referentes a todas as Contas?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
    End If
  End If
    
  'Verificando datas
  strData = NUL
  If IsValid(txtContas(0).Text) Or IsValid(txtContas(1).Text) Then
    If IsValid(txtContas(0).Text) Then
      If Not EData(txtContas(0).Text) Then
        MsgFunc "Vencimento Inicial informado é inválido."
        Exit Sub
      End If
      strData = "(Vencimento >= " & InverteData(txtContas(0).Text, True) & ")"
    End If
    
    If IsValid(txtContas(1).Text) Then
      If Not EData(txtContas(1).Text) Then
        MsgFunc "Vencimento Final informado é inválido."
        Exit Sub
      End If
      Concat strData, IIf(IsValid(strData), " and ", ""), "(Vencimento <= " & InverteData(txtContas(1).Text, True) & ")"
    End If
    
    If IsValid(txtContas(1).Text) And IsValid(txtContas(0).Text) Then
      If CDateDef(txtContas(1).Text) < CDateDef(txtContas(0).Text) Then
        MsgFunc "Vencimento Final informado é maior que Vencimento Inicial."
        Exit Sub
      End If
    End If
  Else
    If MsgFunc("Não foi informado nenhum Vencimento (Inicial e Final)." & vbCrLf & _
         "Deseja adicionar a Taxa em todos os Lançamentos referentes a Conta Fixa?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
    End If
  End If
  
  'Verificando Taxa
  dblTaxa = 0
  If Not IsValid(txtContas(2).Text) Then
    MsgFunc "Taxa não informada."
    Exit Sub
  End If
  dblTaxa = CDblDef(txtContas(2).Text)
  
  Dim strSQLGeracoes    As String
  Dim strSQLContas      As String
  
  Dim rstContas         As Object
  Dim rstGeracoes       As Object
  Dim rstLancamentos    As Object
  
  Dim lngResult         As Long
  
  
  strSQLContas = "Select Código, Precedência from [Contas Fixas] " & IIf(IsValid(strContas), " WHERE " & strContas, "")
  
  lngResult = AbreRecordset(rstContas, strSQLContas, dbOpenSnapshot)
  If lngResult = WL_OK Then
    Do
    
      mlngConta = GetValue(rstContas, "Código", ZERO)
      PagRec = IIf(GetValue(rstContas, "Precedência", ZERO) = 1, "P", "R")
      
      strSQLGeracoes = "Select * from [Gerações Fixas] where Conta = " & mlngConta & IIf(IsValid(strData), " and " & strData, "")
      
      lngResult = AbreRecordset(rstGeracoes, strSQLGeracoes, dbOpenSnapshot)
      
      If lngResult = WL_OK Then
        Do
          AbreRecordset rstLancamentos, "Select [Valor Original] from Lançamentos where PagRec = '" & PagRec & "' and Código = " & GetValue(rstGeracoes, "Código", ZERO), dbOpenDynaset
        
          If Not rstLancamentos.EOF Then
            If TypeOf rstLancamentos Is dao.Recordset Then rstLancamentos.Edit
            
            rstLancamentos("Valor Original").Value = GetValue(rstLancamentos, "Valor Original", ZERO) + (GetValue(rstLancamentos, "Valor Original", ZERO) * (dblTaxa / 100))
            rstLancamentos.Update
          End If
          FechaRecordset rstLancamentos
          
          rstGeracoes.MoveNext
        Loop Until rstGeracoes.EOF
        
      End If
      FechaRecordset rstGeracoes
      
      rstContas.MoveNext
      
    Loop Until rstContas.EOF
  
    MsgFunc "Lançamentos atualizados com Sucesso."
  ElseIf lngResult = WL_NORECORD Then
    MsgFunc "Não foram encontradas Contas Fixas com os parâmetros informados."
  End If
  FechaRecordset rstContas
  Unload Me
  Exit Sub
  
End Sub

Private Sub Form_Load()
  CenterForm Me
  
  lblDescConta(3).Caption = NUL
  lblDescConta(4).Caption = NUL
  
  txtContas(3).Text = mlngConta
  txtContas(4).Text = mlngConta
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set fdlgTaxaSobreConta = Nothing
End Sub

Private Sub txtContas_Change(Index As Integer)
  If Index = 3 Or Index = 4 Then
    GetAssocValue "Select Descrição from [Contas Fixas] where Código = " & txtContas(Index).Text, lblDescConta(Index)
  End If
End Sub

Private Sub txtContas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 3 Or Index = 4 Then
    If KeyCode = vbKeyPageDown And Shift = 0 Then
      PCampo "Contas Fixas", "Contas Fixas", pbCampo, txtContas(Index), "Código"
    End If
  End If
End Sub

Private Sub txtContas_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Or Index = 1 Then
    SetMascara KeyAscii, txtContas(Index).SelStart, MASK_DATA
  ElseIf Index = 2 Then
    DValor KeyAscii
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
