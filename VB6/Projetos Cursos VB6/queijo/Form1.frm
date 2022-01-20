VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7170
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   3375
      Left            =   3585
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3435
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   135
      TabIndex        =   0
      Top             =   360
      Width           =   3240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  List1.AddItem "Provolone"
  List1.AddItem "Camembert"
  List1.AddItem "Cheddar"
  List1.AddItem "Brie"
  List1.AddItem "Suiço"
  List1.AddItem "Roquefort"
End Sub

Private Sub List1_DblClick()
  List2.AddItem List1.Text
  List1.RemoveItem List1.ListIndex
End Sub

Private Sub List2_DblClick()
  List1.AddItem List2.Text
  List2.RemoveItem List2.ListIndex
End Sub
