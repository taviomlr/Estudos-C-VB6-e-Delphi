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
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   1845
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   270
      TabIndex        =   0
      Top             =   630
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  List1.AddItem "Provolone"
  List1.AddItem "Camembert"
  List1.AddItem "Cheddar"
  List1.AddItem "Brie"
  List1.AddItem "Suíço"
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
