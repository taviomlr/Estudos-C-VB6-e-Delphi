VERSION 5.00
Begin VB.Form Simbolos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   1050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   1050
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   -15
      Picture         =   "Simbolos.frx":0000
      Top             =   540
      Width           =   540
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   525
      Picture         =   "Simbolos.frx":0442
      Top             =   30
      Width           =   540
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   525
      Picture         =   "Simbolos.frx":0884
      Top             =   525
      Width           =   540
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   -15
      Picture         =   "Simbolos.frx":0CC6
      Top             =   30
      Width           =   540
   End
End
Attribute VB_Name = "Simbolos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image1_Click()
  Debug.Print "Selecionado Paus"
End Sub

Private Sub Image2_Click()
  Debug.Print "Selecionado Ouro"
End Sub

Private Sub Image3_Click()
  Debug.Print "Selecionado Copas"
End Sub

Private Sub Image4_Click()
  Debug.Print "Selecionado Espada"
End Sub
