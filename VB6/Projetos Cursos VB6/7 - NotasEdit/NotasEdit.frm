VERSION 5.00
Begin VB.Form NotasEdit 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBox 
      Height          =   2475
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   75
      Width           =   3300
   End
End
Attribute VB_Name = "NotasEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
  txtBox.Top = 0
  txtBox.Left = 0
  txtBox.Width = ScaleWidth
  txtBox.Height = ScaleHeight
End Sub

