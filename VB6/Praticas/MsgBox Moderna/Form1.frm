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
   Begin VB.CheckBox lngRegistrosSel 
      Caption         =   "Check1"
      Height          =   315
      Left            =   1005
      TabIndex        =   1
      Top             =   435
      Width           =   270
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1155
      Left            =   945
      TabIndex        =   0
      Top             =   1110
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If lngRegistrosSel = 0 Then
    MsgBox "IMPORTANTE! A VPN tem que está conectada.", vbCritical, "Atenção"
    If MsgBox("A VPN está conectada?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
         Form2.Show
  Else
      MsgBox "Conecte na VPN e depois execute esta operação novamente.", vbInformation
  End If
  End If
End Sub



