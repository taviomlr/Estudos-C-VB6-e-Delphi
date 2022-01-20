VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      FillStyle       =   0  'Solid
      Height          =   1020
      Left            =   1260
      Shape           =   3  'Circle
      Top             =   705
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   1095
      Top             =   570
      Width           =   1500
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArqOn 
         Caption         =   "&On"
         Shortcut        =   ^O
      End
      Begin VB.Menu hifen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqOff 
         Caption         =   "Of&f"
         Shortcut        =   ^F
      End
      Begin VB.Menu hifen 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArqExit 
         Caption         =   "E&xit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuCanal 
      Caption         =   "&Canal"
      Begin VB.Menu mnuSelecionar 
         Caption         =   "&Selecionar ..."
      End
      Begin VB.Menu hifen3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUp 
         Caption         =   "&Up"
      End
      Begin VB.Menu hifen4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDown 
         Caption         =   "&Down"
      End
   End
   Begin VB.Menu mnuSom 
      Caption         =   "&Som"
      Visible         =   0   'False
      Begin VB.Menu mnuMute 
         Caption         =   "&Mute"
      End
      Begin VB.Menu hifen5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreset 
         Caption         =   "&Preset"
         Begin VB.Menu mnuSomPreSoft 
            Caption         =   "&Soft"
         End
         Begin VB.Menu hifen6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSomPreModerate 
            Caption         =   "&Moderate"
         End
         Begin VB.Menu hifen8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSomPreLoud 
            Caption         =   "&Loud"
         End
      End
      Begin VB.Menu hifen9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSomLouder 
         Caption         =   "&Louder"
      End
      Begin VB.Menu hifen7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSomSofter 
         Caption         =   "&Softer"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  PopupMenu mnuSom
End Sub

Private Sub mnuArqExit_Click()
  End
End Sub

Private Sub mnuArqOff_Click()
  Debug.Print ""
  Debug.Print "Som Off"
End Sub

Private Sub mnuArqOn_Click()
  Debug.Print ""
  Debug.Print "Som Ligado"
End Sub

Private Sub mnuDown_Click()
  Debug.Print ""
  Debug.Print "Canal Down"
End Sub

Private Sub mnuMute_Click()
  Debug.Print ""
  Debug.Print "Som Mute"
End Sub

Private Sub mnuSelecionar_Click()
  Debug.Print ""
  Debug.Print "Canal Selecionar"
End Sub

Private Sub mnuSomLouder_Click()
  Debug.Print ""
  Debug.Print "Som Louder"
End Sub

Private Sub mnuSomPreLoud_Click()
  Debug.Print ""
  Debug.Print "Som Preset Loud"
End Sub

Private Sub mnuSomPreModerate_Click()
  Debug.Print ""
  Debug.Print "Som Preset Moderate"
End Sub

Private Sub mnuSomPreSoft_Click()
  Debug.Print ""
  Debug.Print "Som Preset Soft"
End Sub

Private Sub mnuSomSofter_Click()
  Debug.Print ""
  Debug.Print "Som Softer"
End Sub

Private Sub mnuUp_Click()
  Debug.Print ""
  Debug.Print "Canal Up"
End Sub
