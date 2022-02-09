VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
<<<<<<< HEAD
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   1815
=======
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   1290
      TabIndex        =   1
      Top             =   1755
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "oi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   1980
      TabIndex        =   0
      Top             =   840
      Width           =   870
>>>>>>> 83e8114b2c32a0685b9c4d4202c2ada36b75d763
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
<<<<<<< HEAD
    Form3.Show vbModal
=======
  Form3.Show
  
>>>>>>> 83e8114b2c32a0685b9c4d4202c2ada36b75d763
End Sub
