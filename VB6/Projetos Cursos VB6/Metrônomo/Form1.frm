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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   570
      Left            =   1440
      TabIndex        =   3
      Top             =   2325
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   435
      LargeChange     =   10
      Left            =   225
      Max             =   300
      Min             =   1500
      SmallChange     =   25
      TabIndex        =   0
      Top             =   1455
      Value           =   1000
      Width           =   4005
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1860
      Top             =   330
   End
   Begin VB.Label Label2 
      Caption         =   "Rápido"
      Height          =   315
      Left            =   3705
      TabIndex        =   2
      Top             =   2010
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Lento"
      Height          =   315
      Left            =   210
      TabIndex        =   1
      Top             =   2055
      Width           =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Timer1.Enabled = False
  End
End Sub

Private Sub HScroll1_Change()
  Timer1.Interval = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
  Beep
End Sub
