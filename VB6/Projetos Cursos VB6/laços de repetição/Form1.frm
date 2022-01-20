VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_Load()
 ' Dim Contador
'  Contador = 1
  
 ' Do While Contador <= 199
'  Debug.Print Contador
'  Contador = Contador + 1
'  Loop
  
  
'End Sub


'Private Sub Form_Load()
 ' Dim Contador
 ' Contador = 1
'  Do Until Contador > 100
 ' Debug.Print Contador
 ' Contador = Contador + 1
 ' Loop
'End Sub

'Private Sub Form_Load()
 ' Dim Password
 ' Do
 ' Password = InputBox("Digite a senha")
 ' Loop While Password <> "Mussum"
'End Sub

Private Sub Form_Load()
  Dim Password
  Do
  Password = InputBox("Digite a senha")
  Loop Until Password = "Mussum"
End Sub
