VERSION 5.00
Begin VB.Form Banco 
   Caption         =   "Banco"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnMedia 
      Caption         =   "Média"
      Height          =   645
      Left            =   330
      TabIndex        =   1
      Top             =   2340
      Width           =   1920
   End
   Begin VB.ComboBox cboEntrada 
      Height          =   1935
      Left            =   330
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   255
      Width           =   4020
   End
   Begin VB.Label lblMedia 
      BorderStyle     =   1  'Fixed Single
      Height          =   690
      Left            =   2490
      TabIndex        =   2
      Top             =   2355
      Width           =   1830
   End
End
Attribute VB_Name = "Banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnMedia_Click()
  If cboEntrada.ListCount = 0 Then
    MsgBox "Não há valor informado"
    Exit Sub
  End If
  
  Dim Atual As Integer, Total As Single
  
  Total = 0
  Atual = 0
  
  Do While Atual < cboEntrada.ListCount
    Total = Total + Val(cboEntrada.List(Atual))
    Atual = Atual + 1
  Loop
  lblMedia.Caption = Str$(Total / cboEntrada.ListCount)
  
End Sub

Private Sub cboEntrada_KeyPress(KeyAscii As Integer)
  'Se a tecla for Enter
  If KeyAscii = 13 Then
    'Inserir nova entrada
    cboEntrada.AddItem cboEntrada.Text
    'Limpar a parte do texto
    cboEntrada.Text = ""
    'Descartar a tecla pressionada
    KeyAscii = 0
  End If
End Sub

'Private Sub cboEntrada_KeyPress(KeyAscii As Integer)
' If KeyAscii = 83 Or KeyAscii = 84 Then
 '  KeyAscii = 42
 ' End If
'End Sub



