VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboRefeição 
      Height          =   315
      Left            =   2970
      TabIndex        =   7
      Top             =   525
      Width           =   1065
   End
   Begin VB.ComboBox cboAssento 
      Height          =   315
      Left            =   1545
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   525
      Width           =   1065
   End
   Begin VB.ComboBox cboDestino 
      Height          =   1155
      Left            =   165
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   525
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2685
      TabIndex        =   4
      Top             =   2415
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   525
      TabIndex        =   3
      Top             =   2415
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "Refeição"
      Height          =   345
      Left            =   3255
      TabIndex        =   2
      Top             =   165
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Assento"
      Height          =   330
      Left            =   1845
      TabIndex        =   1
      Top             =   165
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Destino"
      Height          =   330
      Left            =   405
      TabIndex        =   0
      Top             =   180
      Width           =   900
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Debug.Print ""
  Debug.Print cboDestino.Text
  Debug.Print cboAssento.Text
  Debug.Print cboRefeição.Text
End Sub

Private Sub Command2_Click()
  Debug.Print ""
  End
End Sub

Private Sub Form_Load()
  cboDestino.Text = ""
  cboDestino.AddItem "Paris"
  cboDestino.AddItem "Moscow"
  cboDestino.AddItem "Roma"
  cboDestino.AddItem "Nova York"
  cboDestino.AddItem "Rio de Janeiro"
  cboDestino.AddItem "Tokyo"
  cboDestino.AddItem "Londre"
  
  cboAssento.AddItem "Corredor"
  cboAssento.AddItem "Centro"
  cboAssento.AddItem "Janela"
  cboAssento.ListIndex = 1
  
  cboRefeição.AddItem "Frango"
  cboRefeição.AddItem "Lasanha"
  cboRefeição.AddItem "Vegetariano"
  cboRefeição.Text = "Sem preferência"
End Sub
