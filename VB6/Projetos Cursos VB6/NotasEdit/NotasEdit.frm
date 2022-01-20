VERSION 5.00
Begin VB.Form NotasEdit 
   Caption         =   "NotasEdit"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtBox 
      Height          =   2985
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   4440
   End
End
Attribute VB_Name = "NotasEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
  TxtBox.Top = 0
  TxtBox.Left = 0
  TxtBox.Width = ScaleWidth
  TxtBox.Height = ScaleHeight
End Sub
