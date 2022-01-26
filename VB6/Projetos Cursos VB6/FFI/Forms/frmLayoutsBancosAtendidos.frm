VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHflxgd.ocx"
Begin VB.Form frmLayoutsBancosAtendidos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Layouts de Bancos Atendidos"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9195
   Icon            =   "frmLayoutsBancosAtendidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   5775
      Left            =   0
      TabIndex        =   2
      Top             =   -60
      Width           =   7755
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBancosAtendidos 
         Height          =   5565
         Left            =   60
         TabIndex        =   3
         Top             =   150
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   9816
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraBotoes 
      Height          =   5775
      Left            =   7770
      TabIndex        =   0
      Top             =   -60
      Width           =   1395
      Begin VB.CommandButton cmdSair 
         Caption         =   "&Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   1
         Top             =   150
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmLayoutsBancosAtendidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---|---------------------------------------------------------------------------------------------------------------------------
'---|   Ueder Budni (14/12/2017)
'---|---------------------------------------------------------------------------------------------------------------------------

Private strListaBancos(3)    As String

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    strListaBancos(0) = "Bradesco" & Chr(vbKeyTab) & "N/A - Específico" & Chr(vbKeyTab) & "Homologado"
    strListaBancos(1) = "Itaú" & Chr(vbKeyTab) & "240" & Chr(vbKeyTab) & "Em Homologação"
    strListaBancos(2) = "Banco do Brasil" & Chr(vbKeyTab) & "240" & Chr(vbKeyTab) & "Homologado"

    Call CarregaCabecalho
    Call CarregaGrid
End Sub

Private Sub CarregaCabecalho()
    With grdBancosAtendidos
        .Cols = 3
        .FixedCols = 0
        .FixedRows = 1
        
        .Clear
        
        .TextMatrix(0, 0) = "Bancos"
        .ColWidth(0) = 3100
        .ColAlignment(0) = flexAlignLeftCenter
        
        .TextMatrix(0, 1) = "CNAB"
        .ColWidth(1) = 1500
        .ColAlignment(1) = flexAlignRightCenter
                
        .TextMatrix(0, 2) = "Situação"
        .ColWidth(2) = 2700
        .ColAlignment(2) = flexAlignLeftCenter
    End With
End Sub

Private Sub CarregaGrid()
    Dim i As Integer

    With grdBancosAtendidos
        For i = 0 To SizeOf(strListaBancos) - 1
            Call .AddItem(strListaBancos(i))
        Next
        Call .RemoveItem(1)
    End With
    
End Sub
