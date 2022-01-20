VERSION 5.00
Begin VB.Form MeuForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cronômetro"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDecorrido 
      Height          =   435
      Left            =   3135
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1860
      Width           =   1080
   End
   Begin VB.TextBox txtFinal 
      Height          =   435
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1095
      Width           =   1080
   End
   Begin VB.TextBox txtInicial 
      Height          =   435
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   315
      Width           =   1080
   End
   Begin VB.Timer tmrCorrido 
      Interval        =   1000
      Left            =   375
      Top             =   2040
   End
   Begin VB.CommandButton BtnParar 
      Caption         =   "Parar"
      Enabled         =   0   'False
      Height          =   540
      Left            =   270
      TabIndex        =   1
      Top             =   1305
      Width           =   1110
   End
   Begin VB.CommandButton Btnlniciar 
      Caption         =   "Iniciar"
      Height          =   540
      Left            =   240
      TabIndex        =   0
      Top             =   405
      Width           =   1110
   End
   Begin VB.Label Label3 
      Caption         =   "Tempo Decorrido:"
      Height          =   330
      Left            =   1725
      TabIndex        =   4
      Top             =   1980
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Hora Final:"
      Height          =   330
      Left            =   2265
      TabIndex        =   3
      Top             =   1215
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Hora Inicial:"
      Height          =   330
      Left            =   2175
      TabIndex        =   2
      Top             =   420
      Width           =   1140
   End
End
Attribute VB_Name = "MeuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempoInicial As Variant
Dim TempoFinal As Variant
Dim TempoDecorrido As Variant

Private Sub Btnlniciar_Click()
  TempoInicial = Now
  txtInicial.Text = Format(TempoInicial, "hh:mm:ss")
  txtFinal = ""
  txtDecorrido = ""
  Btnlniciar.Enabled = False
  BtnParar.Enabled = True
End Sub

Private Sub BtnParar_Click()
  TempoFinal = Now
  TempoDecorrido = TempoFinal - TempoInicial
  txtFinal.Text = Format(TempoFinal, "hh:mm:ss")
  txtDecorrido.Text = Format(TempoDecorrido, "hh:mm:ss")
  Btnlniciar.Enabled = True
  BtnParar.Enabled = False
End Sub
