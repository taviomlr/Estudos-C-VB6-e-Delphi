VERSION 5.00
Object = "{D2618305-B2BB-11D2-925E-444553540000}#1.3#0"; "ReportX.ocx"
Begin VB.Form fimpCheque 
   KeyPreview      =   -1  'True
   Caption         =   "Relatório de Movimentação de Produtos por Centro de Custo"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   ScaleHeight     =   104.246
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   262.732
   StartUpPosition =   3  'Windows Default
   Begin ReportX.ReportMain Imprimir 
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   5102
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Pagina          =   9
      Divisao         =   5
      Regua           =   -1  'True
      MargemEsquerda  =   5
      MargemDireita   =   5
      MargemSuperior  =   5
      MargemInferior  =   5
      Titulo          =   ""
   End
   Begin ReportX.ReportSection ReportSection2 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   4455
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   450
      Tipo            =   7
      Mostrar         =   0   'False
   End
   Begin ReportX.ReportSection ReportSection1 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   450
      Tipo            =   2
      Mostrar         =   0   'False
   End
   Begin ReportX.ReportSection Detalhe 
      Align           =   1  'Align Top
      Height          =   4200
      Left            =   0
      Top             =   255
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   7408
      Begin ReportX.ReportField rfdInfBanco 
         Height          =   240
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   423
         Campo           =   "InfBanco"
         Caption         =   "InfBanco"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rfdValor 
         Height          =   240
         Left            =   7080
         TabIndex        =   2
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         Campo           =   "Valor"
         Caption         =   "Valor"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rfdExtensoA 
         Height          =   240
         Left            =   600
         TabIndex        =   3
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   423
         Campo           =   "Extenso1"
         Caption         =   "ExtensoA"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rfdExtensoB 
         Height          =   240
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   423
         Campo           =   "Extenso2"
         Caption         =   "ExtensoB"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rfdNominal 
         Height          =   240
         Left            =   0
         TabIndex        =   5
         Top             =   1080
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   423
         Campo           =   "Nominal"
         Caption         =   "Nominal"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rfdLocal 
         Height          =   240
         Left            =   2520
         TabIndex        =   6
         Top             =   2640
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   423
         Campo           =   "Local"
         Caption         =   "Local"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rfdMes 
         Height          =   240
         Left            =   6600
         TabIndex        =   7
         Top             =   2640
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   423
         Campo           =   "Mês"
         Caption         =   "Mes"
         Alignment       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rfdAno 
         Height          =   240
         Left            =   8400
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Campo           =   "Ano"
         Caption         =   "Ano"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
      Begin ReportX.ReportField rfdBancoCheque 
         Height          =   240
         Left            =   6480
         TabIndex        =   9
         Top             =   3720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   423
         Campo           =   "BcoChq"
         Caption         =   "BancoCheque"
         Alignment       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
      End
   End
End
Attribute VB_Name = "fimpCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mComplementoValor As String
Private mValorEntreParenteses As Boolean
Private mComplementoExtenso As String
Private mExtensoEntreParentes As Boolean
Private mMesCompleto As Boolean
Private mAnoCompleto As Boolean
Private mCaixaLetras As Integer

Public Sub Config(rsDados As Object, pBanco As Long, Optional Visualizar As Boolean = True)
 
    ConfiguraCheque pBanco
    
    'configuro se visualiza o rel
    Imprimir.Visualizar = Visualizar
    
    'imprimir sempre true
    'se visualizar, habilita o botão
    'se não visualizar, irá diretó para impressora
    Imprimir.Imprimir = True
    
    'se não visualizar, abre tela para selecionar a impressora
    Imprimir.SelecionarImpressora = Not Visualizar
  
    Set Imprimir.Recordset = rsDados
    
    Imprimir.Ativar
  
    Unload Me
  
End Sub

Private Sub ConfiguraFonteControles(pFontName As String, pFontSize As Single, pFontStyle As Integer)

Dim c As Control
Dim rfd As ReportField
Dim bBold As Boolean
Dim bItalic As Boolean

bBold = False
bItalic = False

Select Case pFontStyle
  
  Case 1
    bItalic = True
    
  Case 2
    bBold = True
  
  Case 3
    bBold = True
    bItalic = True
    
End Select


For Each c In Detalhe.Controls
    If TypeOf c Is ReportField Then
        Set rfd = c
        rfd.Font.Name = pFontName
        rfd.Font.Size = pFontSize
        rfd.Font.Bold = bBold
        rfd.Font.Italic = bItalic
    End If
Next

End Sub

Private Sub ConfiguraCheque(pBco As Long)

'recordset do modelo do cheque
Dim rsMdl   As Object
'código de Compensação
Dim nCamara As Long

    nCamara = GetFieldValue("Câmara", "Bancos", "Banco = " & str(pBco))

    If (AbreRecordset(rsMdl, wsprintf("SELECT * FROM ChqModelos WHERE Número = %l", nCamara), dbOpenSnapshot) = WL_OK) Then
                    
        'configura a altura da seção detalhe
        'para a altura do cheque
        
        Detalhe.Height = CDbl(GetValue(rsMdl, "Altura", 90)) - 1
        
        'Imprimir.UsarPapelAtual = False
        'Imprimir.AlturaPapel = CDbl(GetValue(rsMdl, "Altura", 90))
        'Imprimir.LarguraPapel = CDbl(GetValue(rsMdl, "Largura", 150))
        
        'configura a margem
        Imprimir.MargemDireita = 0
        Imprimir.MargemEsquerda = 0
        Imprimir.MargemSuperior = 0
        Imprimir.MargemInferior = 0
        
        'configuando a fonte
        ConfiguraFonteControles GetValue(rsMdl, "FonteNome", "Arial"), CSng(GetValue(rsMdl, "FonteSize", 9)), GetValue(rsMdl, "FonteTipo", 0)

        '
        'posicionando os campos
        '
        PosicionaCampo rsMdl, rfdInfBanco, "InfPos", , "VlrPos"
        PosicionaCampo rsMdl, rfdValor, "VlrPos"
        PosicionaCampo rsMdl, rfdExtensoA, "ExtAPos"
        PosicionaCampo rsMdl, rfdExtensoB, "ExtBPos"
        PosicionaCampo rsMdl, rfdNominal, "NomPos"
        PosicionaCampo rsMdl, rfdLocal, "LocPos"
        PosicionaCampo rsMdl, rfdMes, "MesPos", , "LocPos"
        PosicionaCampo rsMdl, rfdAno, "AnoPos", , "LocPos"
        PosicionaCampo rsMdl, rfdBancoCheque, "NumBan"
        
'        '
'        'carrega configuracoes
'        '
'        mComplementoValor = GetValue(rsMdl, "CaracterSeguranca", Empty)
'        mValorEntreParenteses = GetValue(rsMdl, "FecharValor", False)
'
'        mComplementoExtenso = GetValue(rsMdl, "CaracterComplemento", Empty)
'        mExtensoEntreParentes = GetValue(rsMdl, "FecharExtenso", False)
'
'        mMesCompleto = GetValue(rsMdl, "MesCompleto", True)
'        mAnoCompleto = GetValue(rsMdl, "AnoCompleto", True)
        
    End If
    
    FechaRecordset rsMdl


End Sub

Private Sub PosicionaCampo(pRsMdlChq As Object, pRfd As ReportField, pPrefixoPosWidth As String, Optional pPrefixoPosLeft As String = Empty, Optional pPrefixoPosBase As String = Empty)
       
    If pPrefixoPosLeft = Empty Then
        pPrefixoPosLeft = pPrefixoPosWidth
    End If
    
    If pPrefixoPosBase = Empty Then
        pPrefixoPosBase = pPrefixoPosWidth
    End If
    
    pRfd.Width = Me.ScaleX(GetValue(pRsMdlChq, pPrefixoPosWidth & "Width"), vbMillimeters, vbTwips)
    pRfd.Left = Me.ScaleX(GetValue(pRsMdlChq, pPrefixoPosLeft & "Left"), vbMillimeters, vbTwips)
    pRfd.Top = Me.ScaleY(GetValue(pRsMdlChq, pPrefixoPosBase & "Base"), vbMillimeters, vbTwips) - pRfd.Height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
  Set fimpCheque = Nothing
  
End Sub

Private Sub Imprimir_Erro(ByVal Numero As Long)
  
  RpxMsgErro Numero
  
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
