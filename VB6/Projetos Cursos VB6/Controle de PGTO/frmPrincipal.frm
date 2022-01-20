VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "CHAMALEONBUTTON.OCX"
Begin VB.Form frmPrincipal 
   BackColor       =   &H80000002&
   Caption         =   "Controle de Pagamento"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15630
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   15630
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   63
      Top             =   8790
      Width           =   15630
      _ExtentX        =   27570
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10760
            MinWidth        =   10760
            Picture         =   "frmPrincipal.frx":08CA
            Text            =   "  TOTAL DE CLIENTES: 00"
            TextSave        =   "  TOTAL DE CLIENTES: 00"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "28/12/2021"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "23:17"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11360
            MinWidth        =   11360
            Text            =   "Curso Visual Basic"
            TextSave        =   "Curso Visual Basic"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6930
      Top             =   705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":15F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":2688
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":371A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":3FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":48CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":51A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":5A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":6B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":7BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":8C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":9CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":A5A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":AE7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":BF10
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal.frx":CFA2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   15630
      _ExtentX        =   27570
      _ExtentY        =   1217
      ButtonWidth     =   1693
      ButtonHeight    =   1217
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   29
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NOVO"
            Object.ToolTipText     =   "Cadastrar um novo cliente."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "GRAVAR"
            Object.ToolTipText     =   "Gravar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ATIVOS"
            Object.ToolTipText     =   "Exibe clientes ativos"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SEM PGTO"
            Object.ToolTipText     =   "Exibe clientes ativos e sem pgto"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "VENCIDO"
            Object.ToolTipText     =   "Exibe clientes ativos e em atraso"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "INATIVO"
            Object.ToolTipText     =   "Exibe clientes inativos"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "INF. PGTO"
            Object.ToolTipText     =   "Informar Pgto"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "HISTOR."
            Object.ToolTipText     =   "Exibe movimentação financeira"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ALTERAR"
            Object.ToolTipText     =   "Altera o registro selecionado"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CANCEL"
            Object.ToolTipText     =   "Cancela ação atual"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "IMPRIMIR"
            Object.ToolTipText     =   "Imprime a lista atual"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ARQUIVAR"
            Object.ToolTipText     =   "Começar um novo mês"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EXCLUIR"
            Object.ToolTipText     =   "Exclue o registro atual"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ATUALIZA"
            Object.ToolTipText     =   "Atualizar banco de dados (Rede)"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab shprincipal 
      Height          =   7140
      Left            =   240
      TabIndex        =   0
      Top             =   930
      Width           =   15350
      _ExtentX        =   27067
      _ExtentY        =   12594
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmPrincipal.frx":E034
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmPrincipal.frx":E050
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmPrincipal.frx":E06C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmPrincipal.frx":E088
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "frmPrincipal.frx":E0A4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.Frame Frame6 
         Height          =   6210
         Left            =   180
         TabIndex        =   34
         Top             =   585
         Width           =   15000
         Begin VB.Frame Frame7 
            Height          =   2600
            Left            =   180
            TabIndex        =   36
            Top             =   195
            Width           =   14650
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4890
               TabIndex        =   53
               Top             =   1860
               Width           =   1485
            End
            Begin MSComCtl2.DTPicker dtpAnoRef 
               Height          =   300
               Left            =   4890
               TabIndex        =   51
               Top             =   840
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   529
               _Version        =   393216
               Format          =   81920001
               CurrentDate     =   44552
            End
            Begin VB.TextBox txtHVrMensal 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4890
               TabIndex        =   46
               Top             =   1320
               Width           =   1485
            End
            Begin VB.TextBox txtHVenc 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1785
               TabIndex        =   45
               Top             =   840
               Width           =   1185
            End
            Begin VB.CheckBox chkHAtivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Ativo.............................:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   180
               TabIndex        =   44
               Top             =   1830
               Value           =   1  'Checked
               Width           =   2595
            End
            Begin VB.CheckBox chkHPago 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Pago..............................:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   43
               Top             =   1335
               Width           =   2640
            End
            Begin VB.TextBox txtHNome 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1755
               TabIndex        =   42
               Top             =   390
               Width           =   4665
            End
            Begin VB.Frame Frame8 
               Caption         =   "Resumo"
               Enabled         =   0   'False
               Height          =   1710
               Left            =   6990
               TabIndex        =   37
               Top             =   300
               Width           =   1755
               Begin VB.TextBox txtHRecebido 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   180
                  TabIndex        =   39
                  Top             =   525
                  Width           =   1425
               End
               Begin VB.TextBox txtHReceber 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   180
                  TabIndex        =   38
                  Top             =   1260
                  Width           =   1425
               End
               Begin VB.Label Label14 
                  Caption         =   "Recebido:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   165
                  TabIndex        =   41
                  Top             =   255
                  Width           =   1170
               End
               Begin VB.Label Label13 
                  Caption         =   "À Receber:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   165
                  TabIndex        =   40
                  Top             =   990
                  Width           =   1170
               End
            End
            Begin ChamaleonButton.ChameleonBtn cmdHNpme 
               Height          =   345
               Left            =   6465
               TabIndex        =   56
               Top             =   360
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "ChameleonBtn1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPrincipal.frx":E0C0
               PICN            =   "frmPrincipal.frx":E0DC
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn cmdHMesRef 
               Height          =   345
               Left            =   6465
               TabIndex        =   57
               Top             =   1275
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "ChameleonBtn1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPrincipal.frx":E9B6
               PICN            =   "frmPrincipal.frx":E9D2
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn cmdHAnoRef 
               Height          =   345
               Left            =   6450
               TabIndex        =   58
               Top             =   1815
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "ChameleonBtn1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPrincipal.frx":F2AC
               PICN            =   "frmPrincipal.frx":F2C8
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn cmdHVenc 
               Height          =   345
               Left            =   3015
               TabIndex        =   59
               Top             =   810
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "ChameleonBtn1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPrincipal.frx":FBA2
               PICN            =   "frmPrincipal.frx":FBBE
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn cmdHPago 
               Height          =   345
               Left            =   3000
               TabIndex        =   60
               Top             =   1305
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "ChameleonBtn1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPrincipal.frx":10498
               PICN            =   "frmPrincipal.frx":104B4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ChamaleonButton.ChameleonBtn cmdHAtivo 
               Height          =   345
               Left            =   3000
               TabIndex        =   61
               Top             =   1770
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "ChameleonBtn1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPrincipal.frx":10D8E
               PICN            =   "frmPrincipal.frx":10DAA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "Ano Referência:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3405
               TabIndex        =   54
               Top             =   1890
               Width           =   1410
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Referência......:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3465
               TabIndex        =   52
               Top             =   855
               Width           =   1425
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Mês Referência:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3420
               TabIndex        =   49
               Top             =   1335
               Width           =   1650
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Data Vencimento:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   48
               Top             =   885
               Width           =   1650
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Nome................:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   47
               Top             =   450
               Width           =   1560
            End
            Begin VB.Shape shHistorico 
               BackColor       =   &H00FFFFC0&
               BackStyle       =   1  'Opaque
               Height          =   975
               Left            =   8820
               Shape           =   4  'Rounded Rectangle
               Top             =   585
               Width           =   645
            End
         End
         Begin MSComctlLib.ListView lstHistorico 
            Height          =   2685
            Left            =   90
            TabIndex        =   35
            Top             =   2895
            Width           =   14750
            _ExtentX        =   26009
            _ExtentY        =   4736
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   11
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "COD"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nome do Cliente"
               Object.Width           =   4304
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Dia Venc"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "R$ Mensal"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Pago"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Vencido"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Ativo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Ultimo Reg"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Mês"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Ano"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "Observações"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6195
         Left            =   -74820
         TabIndex        =   14
         Top             =   600
         Width           =   15000
         Begin MSComctlLib.ListView lstClientes 
            Height          =   3195
            Left            =   180
            TabIndex        =   27
            Top             =   2835
            Width           =   14650
            _ExtentX        =   25850
            _ExtentY        =   5636
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "COD"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nome do Cliente"
               Object.Width           =   4304
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Dia Venc"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "R$ Mensal"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Pago"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Vencido"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Ativo"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Ultimo Reg"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Obsevações"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Frame Frame4 
            Height          =   2600
            Left            =   180
            TabIndex        =   15
            Top             =   195
            Width           =   14650
            Begin ChamaleonButton.ChameleonBtn cmdCNome 
               Height          =   345
               Left            =   5985
               TabIndex        =   50
               Top             =   195
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "ChameleonBtn1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPrincipal.frx":11684
               PICN            =   "frmPrincipal.frx":116A0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Frame Frame5 
               Caption         =   "Resumo"
               Enabled         =   0   'False
               Height          =   2250
               Left            =   7185
               TabIndex        =   26
               Top             =   135
               Width           =   2010
               Begin VB.TextBox txtCVrVencido 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   180
                  TabIndex        =   32
                  Top             =   1785
                  Width           =   1425
               End
               Begin VB.TextBox txtCVrReceber 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   180
                  TabIndex        =   30
                  Top             =   1170
                  Width           =   1425
               End
               Begin VB.TextBox txtCVrRecebido 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   180
                  TabIndex        =   28
                  Top             =   525
                  Width           =   1425
               End
               Begin VB.Label Label11 
                  Caption         =   "Vencido:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   165
                  TabIndex        =   33
                  Top             =   1515
                  Width           =   1170
               End
               Begin VB.Label Label10 
                  Caption         =   "À Receber:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   165
                  TabIndex        =   31
                  Top             =   900
                  Width           =   1170
               End
               Begin VB.Label Label9 
                  Caption         =   "Recebido:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   165
                  TabIndex        =   29
                  Top             =   255
                  Width           =   1170
               End
            End
            Begin VB.TextBox txtCNome 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1755
               TabIndex        =   21
               Top             =   240
               Width           =   4140
            End
            Begin VB.CheckBox chkCPago 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Pago.................:"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   20
               Top             =   1020
               Width           =   1920
            End
            Begin VB.CheckBox chqCAtivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Ativo................................:"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3135
               TabIndex        =   19
               Top             =   1020
               Value           =   1  'Checked
               Width           =   2760
            End
            Begin VB.TextBox txtCVenc 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4710
               TabIndex        =   18
               Top             =   660
               Width           =   1185
            End
            Begin VB.TextBox txtCMensal 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1770
               TabIndex        =   17
               Top             =   645
               Width           =   1095
            End
            Begin VB.TextBox txtCObs 
               Appearance      =   0  'Flat
               Height          =   990
               Left            =   1740
               TabIndex        =   16
               Top             =   1380
               Width           =   4215
            End
            Begin ChamaleonButton.ChameleonBtn cmdCVenc 
               Height          =   345
               Left            =   5985
               TabIndex        =   55
               Top             =   600
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   609
               BTYPE           =   3
               TX              =   "ChameleonBtn1"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPrincipal.frx":11F7A
               PICN            =   "frmPrincipal.frx":11F96
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label8 
               BackColor       =   &H80000004&
               BackStyle       =   0  'Transparent
               Caption         =   "Nome................:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   25
               Top             =   300
               Width           =   1560
            End
            Begin VB.Label Label7 
               BackColor       =   &H80000004&
               BackStyle       =   0  'Transparent
               Caption         =   "Data Vencimento:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3075
               TabIndex        =   24
               Top             =   705
               Width           =   1650
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000004&
               BackStyle       =   0  'Transparent
               Caption         =   "Observações......:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   23
               Top             =   1335
               Width           =   1560
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000004&
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Mensal......:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   120
               TabIndex        =   22
               Top             =   675
               Width           =   2145
            End
            Begin VB.Shape shClientes 
               BackColor       =   &H00FFFFC0&
               BackStyle       =   1  'Opaque
               Height          =   975
               Left            =   9255
               Shape           =   4  'Rounded Rectangle
               Top             =   735
               Width           =   645
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6495
         Left            =   -74820
         TabIndex        =   1
         Top             =   630
         Width           =   15000
         Begin VB.Frame Frame2 
            Height          =   2600
            Left            =   180
            TabIndex        =   2
            Top             =   195
            Width           =   14650
            Begin VB.TextBox txtObs 
               Appearance      =   0  'Flat
               Height          =   1050
               Left            =   1785
               TabIndex        =   12
               Top             =   1305
               Width           =   4590
            End
            Begin VB.TextBox txtVrMensal 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5355
               TabIndex        =   11
               Top             =   600
               Width           =   1035
            End
            Begin VB.TextBox txtVenc 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1785
               TabIndex        =   10
               Top             =   600
               Width           =   1185
            End
            Begin VB.CheckBox chkAtivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Ativo"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   5295
               TabIndex        =   7
               Top             =   945
               Value           =   1  'Checked
               Width           =   1080
            End
            Begin VB.CheckBox chkPago 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Pago.................:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   6
               Top             =   960
               Width           =   1920
            End
            Begin VB.TextBox txtNome 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1770
               TabIndex        =   4
               Top             =   205
               Width           =   4665
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Valor da Mensalidade...:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3135
               TabIndex        =   13
               Top             =   615
               Width           =   2145
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Observações......:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   9
               Top             =   1275
               Width           =   1560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Data Vencimento:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   8
               Top             =   645
               Width           =   1650
            End
            Begin VB.Label Nome 
               BackStyle       =   0  'Transparent
               Caption         =   "Nome................:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   150
               TabIndex        =   3
               Top             =   275
               Width           =   1560
            End
            Begin VB.Shape shCadastro 
               BackColor       =   &H00FFFFC0&
               BackStyle       =   1  'Opaque
               Height          =   975
               Left            =   6615
               Shape           =   4  'Rounded Rectangle
               Top             =   780
               Width           =   645
            End
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   570
      TabIndex        =   5
      Top             =   2925
      Width           =   690
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChameleonBtn2_Click()

End Sub

Private Sub ChameleonBtn8_Click()

End Sub

Private Sub cmdCNome_Click()
  If txtCNome = "" Then
    txtCNome.BackColor = 5555
    txtCNome.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cmdCVenc_Click()
  If txtCVenc = "" Then
    txtCVenc.BackColor = 5555
    txtCVenc.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cmdHNpme_Click()
  If txtHNome.Text = "" Then
    txtHNome.BackColor = 5555
    txtCNome.SetFocus
    Exit Sub
  End If
End Sub

Private Sub cmdHVenc_Click()
 If txtHVenc.Text = "" Then
    txtHVenc.BackColor = 5555
    txtCVenc.SetFocus
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
  Frame2.BorderStyle = 0
  Frame4.BorderStyle = 0
  Frame7.BorderStyle = 0
End Sub

Private Sub Form_Resize()

'On Error Resume Next
  shprincipal.Width = Me.Width - 400
  Frame6.Width = shprincipal.Width - 350
  Frame7.Width = Frame6.Width - 350
  lstHistorico.Width = Frame6.Width - 250
  With Frame4
    shClientes.Top = .Top - 160
    shClientes.Left = .Left - 100
    shClientes.Height = .Height - 110
    shClientes.Width = .Width
    
    shHistorico.Top = .Top - 160
    shHistorico.Left = .Left - 100
    shHistorico.Height = .Height - 110
    shHistorico.Width = .Width
    
   shCadastro.Top = .Top - 160
   shCadastro.Left = .Left - 100
   shCadastro.Height = .Height - 110
   shCadastro.Width = .Width
  End With

  
    
End Sub



Private Sub txtCMensal_Change()
  txtCMensal.BackColor = vbWhite
End Sub

Private Sub txtCNome_Change()
  txtCNome.BackColor = vbWhite
  Dim Pos As Integer
      Pos = txtCNome.SelStart
      txtCNome.Text = VBA.UCase(txtCNome.Text)
      txtCNome.SelStart = Pos
End Sub

Private Sub txtCObs_Change()
  Dim Pos As Integer
      Pos = txtCObs.SelStart
      txtCObs.Text = VBA.UCase(txtCObs.Text)
      txtCObs.SelStart = Pos
End Sub

Private Sub txtCVenc_Change()
  txtCVenc.BackColor = vbWhite
End Sub

Private Sub txtHNome_Change()
  txtHNome.BackColor = vbWhite
  Dim Pos As Integer
      Pos = txtHNome.SelStart
      txtHNome.Text = VBA.UCase(txtHNome.Text)
      txtHNome.SelStart = Pos
End Sub

Private Sub txtHVenc_Change()
  txtHVenc.BackColor = vbWhite
End Sub

Private Sub txtNome_Change()
  txtNome.BackColor = vbWhite
  Dim Pos As Integer
      Pos = txtNome.SelStart
      txtNome.Text = VBA.UCase(txtNome.Text)
      txtNome.SelStart = Pos
End Sub

Private Sub txtObs_Change()
  Dim Pos As Integer
      Pos = txtObs.SelStart
      txtObs.Text = VBA.UCase(txtObs.Text)
      txtObs.SelStart = Pos
End Sub

Private Sub txtVenc_Change()
  txtVenc.BackColor = vbWhite
End Sub

Private Sub txtVrMensal_Change()
  txtVrMensal.BackColor = vbWhite
End Sub
