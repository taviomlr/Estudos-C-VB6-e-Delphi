VERSION 5.00
Begin VB.Form Sorriso 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   Picture         =   "Sorriso.frx":0000
   ScaleHeight     =   4680
   ScaleWidth      =   4140
End
Attribute VB_Name = "Sorriso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Form_Click()
'  Sorriso.Print "Tenha um ";
'  Sorriso.FontBold = True
'  Sorriso.Print "bom";
'  Sorriso.FontBold = False
'  Sorriso.Print " dia "
'End Sub
'
'Private Sub Form_DblClick()
'  Sorriso.Cls
'End Sub

'opção sem o nome do formulário, facilita o reaproveitamento de métodos
Private Sub Form_Click()
  Print "Tenha um ";
  FontBold = True
  Print "bom";
  FontBold = False
  Print " dia "
End Sub

Private Sub Form_DblClick()
  Cls
End Sub
