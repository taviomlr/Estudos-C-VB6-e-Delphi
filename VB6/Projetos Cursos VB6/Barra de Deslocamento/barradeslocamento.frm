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
   Begin VB.HScrollBar HScroll1 
      Height          =   390
      LargeChange     =   10
      Left            =   375
      Max             =   100
      SmallChange     =   2
      TabIndex        =   1
      Top             =   1875
      Width           =   3750
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   765
      TabIndex        =   0
      Top             =   405
      Width           =   2445
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HScroll1_Change()
  Label1.Caption = Str$(HScroll1.Value)
End Sub


Private Sub HScroll1_Scroll()
  Label1.Caption = "Movendo para " & Str$(HScroll1.Value)
End Sub
