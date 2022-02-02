VERSION 5.00
Begin VB.Form Cronometro 
   Caption         =   "Cronometro"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDecorrido 
      Height          =   495
      Left            =   3075
      TabIndex        =   4
      Top             =   2220
      Width           =   1215
   End
   Begin VB.TextBox txtFinal 
      Height          =   495
      Left            =   3105
      TabIndex        =   3
      Top             =   1335
      Width           =   1215
   End
   Begin VB.TextBox txtInicial 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   450
      Width           =   1215
   End
   Begin VB.CommandButton btnParar 
      Caption         =   "Parar"
      Height          =   495
      Left            =   255
      TabIndex        =   1
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton btnIniciar 
      Caption         =   "Iniciar"
      Height          =   495
      Left            =   270
      TabIndex        =   0
      Top             =   705
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Tempo Decorrido:"
      Height          =   210
      Left            =   1800
      TabIndex        =   7
      Top             =   2355
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "Tenpo Final:"
      Height          =   210
      Left            =   1815
      TabIndex        =   6
      Top             =   1470
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Tenpo Inicial:"
      Height          =   210
      Left            =   1815
      TabIndex        =   5
      Top             =   600
      Width           =   1290
   End
End
Attribute VB_Name = "Cronometro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempoInicial As Variant
Dim TempoFinal As Variant
Dim TempoDecorrido As Variant

Private Sub btnIniciar_Click()
  TempoInicial = Now
  txtInicial.Text = Format(TempoInicial, "hh:mm:ss")
  txtFinal.Text = ""
  txtDecorrido.Text = ""
  btnParar.Enabled = True
  btnIniciar.Enabled = False
End Sub

Private Sub btnParar_Click()
  TempoFinal = Now
  TempoDecorrido = TempoFinal - TempoInicial
  txtFinal.Text = Format(TempoFinal, "hh:mm:ss")
  txtDecorrido.Text = Format(TempoDecorrido, "hh:mm:ss")
  btnParar.Enabled = False
  btnIniciar.Enabled = True
End Sub
