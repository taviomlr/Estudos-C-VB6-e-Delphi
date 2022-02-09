VERSION 5.00
Begin VB.Form Cronometro2 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnParar 
      Caption         =   "Parar"
      Height          =   450
      Left            =   375
      TabIndex        =   1
      Top             =   1725
      Width           =   1215
   End
   Begin VB.CommandButton btnIniciar 
      Caption         =   "Iniciar"
      Height          =   495
      Left            =   375
      TabIndex        =   0
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblDecorrido 
      BackColor       =   &H8000000E&
      Height          =   300
      Left            =   3435
      TabIndex        =   7
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label lblTempoDecorrido 
      Caption         =   "Tempo Decorrido:"
      Height          =   330
      Left            =   2055
      TabIndex        =   6
      Top             =   2205
      Width           =   1425
   End
   Begin VB.Label lblFinal 
      BackColor       =   &H8000000E&
      Height          =   300
      Left            =   3435
      TabIndex        =   5
      Top             =   1275
      Width           =   900
   End
   Begin VB.Label lblTempoFinal 
      Caption         =   "Tempo Final:"
      Height          =   330
      Left            =   2055
      TabIndex        =   4
      Top             =   1305
      Width           =   1170
   End
   Begin VB.Label lblInicial 
      BackColor       =   &H8000000E&
      Height          =   300
      Left            =   3420
      TabIndex        =   3
      Top             =   375
      Width           =   900
   End
   Begin VB.Label lblTempoInicial 
      Caption         =   "Tempo Inicial:"
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   420
      Width           =   1170
   End
End
Attribute VB_Name = "Cronometro2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempoInicial As Variant
Dim TempoFinal As Variant
Dim TempoDecorrido As Variant

Private Sub btnIniciar_Click()
  TempoInicial = Now
  lblInicial.Caption = Format(TempoInicial, "hh:mm:ss")
  lblFinal.Caption = ""
  lblDecorrido.Caption = ""
  btnIniciar.Enabled = False
  btnParar.Enabled = True
End Sub

Private Sub btnParar_Click()
  TempoFinal = Now
  TempoDecorrido = TempoFinal - TempoInicial
  lblFinal.Caption = Format(TempoFinal, "hh:mm:ss")
  lblDecorrido.Caption = Format(TempoDecorrido, "hh:mm:ss")
  btnParar.Enabled = False
  btnIniciar.Enabled = True
End Sub

