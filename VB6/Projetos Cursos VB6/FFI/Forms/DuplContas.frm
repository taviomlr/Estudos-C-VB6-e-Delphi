VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSComctl.ocx"
Begin VB.Form frmDuplContas 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Integra��o Cont�bil"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11955
   Icon            =   "DuplContas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10440
      TabIndex        =   110
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   375
      Left            =   8880
      TabIndex        =   109
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Frame FraPrincial 
      Caption         =   "Dados da Conta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   7215
      Index           =   1
      Left            =   120
      TabIndex        =   108
      Top             =   120
      Width           =   11775
      Begin VB.Frame fraTab 
         Height          =   5895
         Index           =   1
         Left            =   240
         TabIndex        =   112
         Top             =   1200
         Width           =   11415
         Begin VB.Frame Frame2 
            Caption         =   "Fato Gerador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   5895
            Left            =   0
            TabIndex        =   116
            Top             =   0
            Width           =   11655
            Begin VB.CheckBox chkFatoGerador 
               Caption         =   "Fato Gerador"
               DataField       =   "Fato Gerador"
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   240
               TabIndex        =   81
               Top             =   360
               Width           =   1335
            End
            Begin VB.Frame fraFatoGerador 
               Caption         =   "Fato Gerador"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   5175
               Left            =   0
               TabIndex        =   117
               Top             =   720
               Width           =   11415
               Begin VB.TextBox txtContas 
                  DataField       =   "Complemento - Fato Gerador - Nota - Cr�dito"
                  Height          =   315
                  Index           =   46
                  Left            =   8640
                  MaxLength       =   1
                  TabIndex        =   100
                  Tag             =   "Contas"
                  Top             =   840
                  Width           =   495
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "Complemento - Fato Gerador - Empresa - Cr�dito"
                  Height          =   315
                  Index           =   45
                  Left            =   8640
                  MaxLength       =   1
                  TabIndex        =   102
                  Tag             =   "Contas"
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "Complemento - Fato Gerador - Data - Cr�dito"
                  Height          =   315
                  Index           =   44
                  Left            =   8640
                  MaxLength       =   1
                  TabIndex        =   104
                  Tag             =   "Contas"
                  Top             =   1560
                  Width           =   495
               End
               Begin VB.CheckBox chkComplDescrFatoGerador 
                  Alignment       =   1  'Right Justify
                  Caption         =   "&Descri��o"
                  DataField       =   "Complemento - Fato Gerador - Descri��o"
                  ForeColor       =   &H8000000D&
                  Height          =   255
                  Index           =   0
                  Left            =   4920
                  TabIndex        =   98
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "Complemento - Fato Gerador - Data - D�bito"
                  Height          =   315
                  Index           =   36
                  Left            =   5760
                  MaxLength       =   1
                  TabIndex        =   97
                  Tag             =   "Contas"
                  Top             =   1560
                  Width           =   495
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "Complemento - Fato Gerador - Empresa - D�bito"
                  Height          =   315
                  Index           =   35
                  Left            =   5760
                  MaxLength       =   1
                  TabIndex        =   95
                  Tag             =   "Contas"
                  Top             =   1200
                  Width           =   495
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "Complemento - Fato Gerador - Nota - D�bito"
                  Height          =   315
                  Index           =   33
                  Left            =   5760
                  MaxLength       =   1
                  TabIndex        =   93
                  Tag             =   "Contas"
                  Top             =   840
                  Width           =   495
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "C�digo do Hist�rico 2 - Fato Gerador"
                  Height          =   315
                  Index           =   16
                  Left            =   1320
                  MaxLength       =   3
                  TabIndex        =   91
                  Tag             =   "Contas"
                  Top             =   1440
                  Width           =   2055
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "C�digo do Hist�rico 1 - Fato Gerador"
                  Height          =   315
                  Index           =   15
                  Left            =   1320
                  MaxLength       =   3
                  TabIndex        =   89
                  Tag             =   "Contas"
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.ComboBox cboContas 
                  DataField       =   "Conta a Cr�dito - Fato Gerador"
                  Height          =   315
                  Index           =   10
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   86
                  Tag             =   "Contas"
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "Conta a Cr�dito Outros - Fato Gerador"
                  Height          =   315
                  Index           =   18
                  Left            =   3480
                  MaxLength       =   9
                  TabIndex        =   87
                  Tag             =   "Contas"
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.ComboBox cboContas 
                  DataField       =   "Conta a D�bito - Fato Gerador"
                  Height          =   315
                  Index           =   9
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   83
                  Tag             =   "Contas"
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.TextBox txtContas 
                  DataField       =   "Conta a D�bito Outros - Fato Gerador"
                  Height          =   315
                  Index           =   17
                  Left            =   3480
                  MaxLength       =   9
                  TabIndex        =   84
                  Tag             =   "Contas"
                  Top             =   360
                  Width           =   1215
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "&Nota:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   54
                  Left            =   7800
                  TabIndex        =   99
                  Top             =   840
                  Width           =   390
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "&Empresa:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   53
                  Left            =   7800
                  TabIndex        =   101
                  Top             =   1200
                  Width           =   660
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "&Data:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   52
                  Left            =   7800
                  TabIndex        =   103
                  Top             =   1560
                  Width           =   390
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "(Ordens dos Complementos - Cr�dito)"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   51
                  Left            =   7680
                  TabIndex        =   125
                  Top             =   480
                  Width           =   2610
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "(Ordens dos Complementos - D�bito)"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   39
                  Left            =   4800
                  TabIndex        =   121
                  Top             =   480
                  Width           =   2580
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "&Data:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   37
                  Left            =   4920
                  TabIndex        =   96
                  Top             =   1560
                  Width           =   390
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "&Empresa:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   36
                  Left            =   4920
                  TabIndex        =   94
                  Top             =   1200
                  Width           =   660
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "&Nota:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   35
                  Left            =   4920
                  TabIndex        =   92
                  Top             =   840
                  Width           =   390
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "C�d. Hist�rico 2:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   20
                  Left            =   120
                  TabIndex        =   90
                  Top             =   1440
                  Width           =   1170
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "Cod. Hist�rico 1:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   18
                  Left            =   120
                  TabIndex        =   88
                  Top             =   1080
                  Width           =   1170
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "Conta a Cr�dito:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   85
                  Top             =   720
                  Width           =   1140
               End
               Begin VB.Label lblContas 
                  AutoSize        =   -1  'True
                  Caption         =   "Conta a D�bito:"
                  ForeColor       =   &H80000002&
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   82
                  Top             =   360
                  Width           =   1110
               End
            End
         End
      End
      Begin VB.Frame fraTab 
         Height          =   5895
         Index           =   0
         Left            =   240
         TabIndex        =   111
         Top             =   1200
         Width           =   11415
         Begin VB.Frame Frame1 
            Caption         =   "Fato Pagamento - Acr�scimo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   2055
            Index           =   2
            Left            =   0
            TabIndex        =   113
            Top             =   3840
            Width           =   11415
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Acr�scimo - Nota - Cr�dito"
               Height          =   315
               Index           =   43
               Left            =   9120
               MaxLength       =   1
               TabIndex        =   74
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Acr�scimo - Cheque - Cr�dito"
               Height          =   315
               Index           =   42
               Left            =   10560
               MaxLength       =   1
               TabIndex        =   80
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Acr�scimo - Empresa - Cr�dito"
               Height          =   315
               Index           =   41
               Left            =   9120
               MaxLength       =   1
               TabIndex        =   76
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Acr�scimo - Data - Cr�dito"
               Height          =   315
               Index           =   40
               Left            =   10560
               MaxLength       =   1
               TabIndex        =   78
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox chkComplDescrAcres 
               Alignment       =   1  'Right Justify
               Caption         =   "&Descri��o"
               DataField       =   "Complemento - Acr�scimo - Descri��o"
               ForeColor       =   &H8000000D&
               Height          =   255
               Index           =   0
               Left            =   5040
               TabIndex        =   72
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Acr�scimo - Data - D�bito"
               Height          =   315
               Index           =   31
               Left            =   7320
               MaxLength       =   1
               TabIndex        =   69
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Acr�scimo - Empresa - D�bito"
               Height          =   315
               Index           =   30
               Left            =   5880
               MaxLength       =   1
               TabIndex        =   67
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Acr�scimo - Cheque - D�bito"
               Height          =   315
               Index           =   29
               Left            =   7320
               MaxLength       =   1
               TabIndex        =   71
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Acr�scimo - Nota - D�bito"
               Height          =   315
               Index           =   28
               Left            =   5880
               MaxLength       =   1
               TabIndex        =   65
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "C�digo do Hist�rico 2 - Acr�scimo"
               Height          =   315
               Index           =   14
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   63
               Tag             =   "Contas"
               Top             =   1440
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "C�digo do Hist�rico 1 - Acr�scimo"
               Height          =   315
               Index           =   13
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   61
               Tag             =   "Contas"
               Top             =   1080
               Width           =   2055
            End
            Begin VB.ComboBox cboContas 
               DataField       =   "Conta a Cr�dito - Acr�scimo"
               Height          =   315
               Index           =   7
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Tag             =   "Contas"
               Top             =   720
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Conta a Cr�dito Outros - Acr�scimo"
               Height          =   315
               Index           =   12
               Left            =   3480
               MaxLength       =   9
               TabIndex        =   59
               Tag             =   "Contas"
               Top             =   720
               Width           =   1215
            End
            Begin VB.ComboBox cboContas 
               DataField       =   "Conta a D�bito - Acr�scimo"
               Height          =   315
               Index           =   6
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Tag             =   "Contas"
               Top             =   360
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Conta a D�bito Outros - Acr�scimo"
               Height          =   315
               Index           =   11
               Left            =   3480
               MaxLength       =   9
               TabIndex        =   56
               Tag             =   "Contas"
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Nota:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   50
               Left            =   8280
               TabIndex        =   73
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Cheque:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   49
               Left            =   9840
               TabIndex        =   79
               Top             =   1080
               Width           =   600
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "(Ordens dos Complementos - Cr�dito)"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   48
               Left            =   8280
               TabIndex        =   124
               Top             =   360
               Width           =   2610
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Data:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   47
               Left            =   9840
               TabIndex        =   77
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Empresa:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   46
               Left            =   8280
               TabIndex        =   75
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Empresa:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   38
               Left            =   5040
               TabIndex        =   66
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Data:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   19
               Left            =   6600
               TabIndex        =   68
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "(Ordens dos Complementos - D�bito)"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   34
               Left            =   5040
               TabIndex        =   120
               Top             =   360
               Width           =   2580
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Cheque:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   33
               Left            =   6600
               TabIndex        =   70
               Top             =   1080
               Width           =   600
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Nota:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   32
               Left            =   5040
               TabIndex        =   64
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "C�d. Hist�rico 2:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   62
               Top             =   1440
               Width           =   1170
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Cod. Hist�rico 1:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   15
               Left            =   120
               TabIndex        =   60
               Top             =   1080
               Width           =   1170
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Conta a Cr�dito:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   57
               Top             =   720
               Width           =   1140
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Conta a D�bito:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   13
               Left            =   120
               TabIndex        =   54
               Top             =   360
               Width           =   1110
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Fato Pagamento - Abatimento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   1935
            Index           =   1
            Left            =   0
            TabIndex        =   114
            Top             =   1920
            Width           =   11415
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Abatimento - Nota - Cr�dito"
               Height          =   315
               Index           =   39
               Left            =   9120
               MaxLength       =   1
               TabIndex        =   47
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Abatimento - Cheque - Cr�dito"
               Height          =   315
               Index           =   37
               Left            =   10560
               MaxLength       =   1
               TabIndex        =   53
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Abatimento - Empresa - Cr�dito"
               Height          =   315
               Index           =   34
               Left            =   9120
               MaxLength       =   1
               TabIndex        =   49
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Abatimento - Data - Cr�dito"
               Height          =   315
               Index           =   32
               Left            =   10560
               MaxLength       =   1
               TabIndex        =   51
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox chkComplDescrAbat 
               Alignment       =   1  'Right Justify
               Caption         =   "&Descri��o"
               DataField       =   "Complemento - Abatimento - Descri��o"
               ForeColor       =   &H8000000D&
               Height          =   255
               Index           =   0
               Left            =   5040
               TabIndex        =   41
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Abatimento - Data - D�bito"
               Height          =   315
               Index           =   26
               Left            =   7320
               MaxLength       =   1
               TabIndex        =   43
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Abatimento - Empresa - D�bito"
               Height          =   315
               Index           =   25
               Left            =   5880
               MaxLength       =   1
               TabIndex        =   40
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Abatimento - Cheque - D�bito"
               Height          =   315
               Index           =   24
               Left            =   7320
               MaxLength       =   1
               TabIndex        =   45
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Abatimento - Nota - D�bito"
               Height          =   315
               Index           =   23
               Left            =   5880
               MaxLength       =   1
               TabIndex        =   38
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Conta a D�bito Outros - Abatimento"
               Height          =   315
               Index           =   9
               Left            =   3480
               MaxLength       =   9
               TabIndex        =   29
               Tag             =   "Contas"
               Top             =   360
               Width           =   1215
            End
            Begin VB.ComboBox cboContas 
               DataField       =   "Conta a D�bito - Abatimento"
               Height          =   315
               Index           =   3
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Tag             =   "Contas"
               Top             =   360
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Conta a Cr�dito Outros - Abatimento"
               Height          =   315
               Index           =   10
               Left            =   3480
               MaxLength       =   9
               TabIndex        =   32
               Tag             =   "Contas"
               Top             =   720
               Width           =   1215
            End
            Begin VB.ComboBox cboContas 
               DataField       =   "Conta a Cr�dito - Abatimento"
               Height          =   315
               Index           =   4
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Tag             =   "Contas"
               Top             =   720
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "C�digo do Hist�rico 1 - Abatimento"
               Height          =   315
               Index           =   8
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   34
               Tag             =   "Contas"
               Top             =   1080
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "C�digo do Hist�rico 2 - Abatimento"
               Height          =   315
               Index           =   7
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   36
               Tag             =   "Contas"
               Top             =   1440
               Width           =   2055
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Nota:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   45
               Left            =   8280
               TabIndex        =   46
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Empresa:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   44
               Left            =   8280
               TabIndex        =   48
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Data:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   43
               Left            =   9840
               TabIndex        =   50
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Cheque:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   42
               Left            =   9840
               TabIndex        =   52
               Top             =   1080
               Width           =   600
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "(Ordens dos Complementos - Cr�dito)"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   41
               Left            =   8280
               TabIndex        =   123
               Top             =   360
               Width           =   2610
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "(Ordens dos Complementos - D�bito)"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   31
               Left            =   5040
               TabIndex        =   119
               Top             =   360
               Width           =   2580
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Cheque:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   30
               Left            =   6600
               TabIndex        =   44
               Top             =   1080
               Width           =   600
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Data:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   29
               Left            =   6600
               TabIndex        =   42
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Empresa:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   28
               Left            =   5040
               TabIndex        =   39
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Nota:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   27
               Left            =   5040
               TabIndex        =   37
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Conta a D�bito:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   12
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Width           =   1110
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Conta a Cr�dito:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   11
               Left            =   120
               TabIndex        =   30
               Top             =   720
               Width           =   1140
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Cod. Hist�rico 1:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   33
               Top             =   1080
               Width           =   1170
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "C�d. Hist�rico 2:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   35
               Top             =   1440
               Width           =   1170
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Fato Pagamento - Valor Original"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   1935
            Index           =   0
            Left            =   0
            TabIndex        =   115
            Top             =   0
            Width           =   11415
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Valor Original - Nota - Cr�dito"
               Height          =   315
               Index           =   27
               Left            =   9120
               MaxLength       =   1
               TabIndex        =   20
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Valor Original - Cheque - Cr�dito"
               Height          =   315
               Index           =   22
               Left            =   10560
               MaxLength       =   1
               TabIndex        =   26
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Valor Original - Empresa - Cr�dito"
               Height          =   315
               Index           =   47
               Left            =   9120
               MaxLength       =   1
               TabIndex        =   22
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Valor Original - Data - Cr�dito"
               Height          =   315
               Index           =   48
               Left            =   10560
               MaxLength       =   1
               TabIndex        =   24
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.CheckBox chkComplDescrVO 
               Alignment       =   1  'Right Justify
               Caption         =   "&Descri��o"
               DataField       =   "Complemento - Valor Original - Descri��o"
               ForeColor       =   &H8000000D&
               Height          =   255
               Index           =   0
               Left            =   5040
               TabIndex        =   14
               Top             =   1440
               Width           =   1335
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Valor Original - Data - D�bito"
               Height          =   315
               Index           =   21
               Left            =   7320
               MaxLength       =   1
               TabIndex        =   16
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Valor Original - Empresa - D�bito"
               Height          =   315
               Index           =   20
               Left            =   5880
               MaxLength       =   1
               TabIndex        =   13
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Valor Original - Cheque - D�bito"
               Height          =   315
               Index           =   19
               Left            =   7320
               MaxLength       =   1
               TabIndex        =   18
               Tag             =   "Contas"
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Complemento - Valor Original - Nota - D�bito"
               Height          =   315
               Index           =   38
               Left            =   5880
               MaxLength       =   1
               TabIndex        =   11
               Tag             =   "Contas"
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtContas 
               DataField       =   "C�digo do Hist�rico 2 - Valor Original"
               Height          =   315
               Index           =   6
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   9
               Tag             =   "Contas"
               Top             =   1440
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "C�digo do Hist�rico 1 - Valor Original"
               Height          =   315
               Index           =   5
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   7
               Tag             =   "Contas"
               Top             =   1080
               Width           =   2055
            End
            Begin VB.ComboBox cboContas 
               DataField       =   "Conta a Cr�dito - Valor Original"
               Height          =   315
               Index           =   1
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Tag             =   "Contas"
               Top             =   720
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Conta a Cr�dito Outros - Valor Original"
               Height          =   315
               Index           =   4
               Left            =   3480
               MaxLength       =   9
               TabIndex        =   5
               Tag             =   "Contas"
               Top             =   720
               Width           =   1215
            End
            Begin VB.ComboBox cboContas 
               DataField       =   "Conta a D�bito - Valor Original"
               Height          =   315
               Index           =   0
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Tag             =   "Contas"
               Top             =   360
               Width           =   2055
            End
            Begin VB.TextBox txtContas 
               DataField       =   "Conta a D�bito Outros - Valor Original"
               Height          =   315
               Index           =   3
               Left            =   3480
               MaxLength       =   9
               TabIndex        =   2
               Tag             =   "Contas"
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Nota:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   40
               Left            =   8280
               TabIndex        =   19
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Empresa:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   21
               Left            =   8280
               TabIndex        =   21
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Data:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   17
               Left            =   9840
               TabIndex        =   23
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Cheque:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   8
               Left            =   9840
               TabIndex        =   25
               Top             =   1080
               Width           =   600
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "(Ordens dos Complementos - Cr�dito)"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   7
               Left            =   8280
               TabIndex        =   122
               Top             =   360
               Width           =   2610
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "(Ordens dos Complementos - D�bito)"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   26
               Left            =   5040
               TabIndex        =   118
               Top             =   360
               Width           =   2580
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Cheque:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   25
               Left            =   6600
               TabIndex        =   17
               Top             =   1080
               Width           =   600
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Data:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   24
               Left            =   6600
               TabIndex        =   15
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Empresa:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   23
               Left            =   5040
               TabIndex        =   12
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "&Nota:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   22
               Left            =   5040
               TabIndex        =   10
               Top             =   720
               Width           =   390
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "C�d. Hist�rico 2:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   8
               Top             =   1440
               Width           =   1170
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Cod. Hist�rico 1:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   6
               Top             =   1080
               Width           =   1170
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Conta a Cr�dito:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   3
               Top             =   720
               Width           =   1140
            End
            Begin VB.Label lblContas 
               AutoSize        =   -1  'True
               Caption         =   "Conta a D�bito:"
               ForeColor       =   &H80000002&
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   0
               Top             =   360
               Width           =   1110
            End
         End
      End
      Begin MSComctlLib.TabStrip TabIntegracao 
         Height          =   6375
         Left            =   120
         TabIndex        =   107
         Top             =   840
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   11245
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fato Gerador"
               Key             =   "FatoGerador"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Fato Pagamento"
               Key             =   "FatoPagamento"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtContas 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   106
         Tag             =   "Contas"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblContas 
         AutoSize        =   -1  'True
         Caption         =   "Conta:"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   105
         Top             =   360
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmDuplContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Conta        As Long
Public Registro     As String
Public PagRec       As String
Public Tipo         As String
Public Numero       As Long
Public Parcela      As Long
Public Empresa      As String

Public Abatimento      As Double
Public Acrescimo       As Double


Private Sub chkFatoGerador_Click()

  If chkFatoGerador.Value = vbChecked Then
     fraFatoGerador.Visible = True
  Else
     fraFatoGerador.Visible = False
     'cboContas(9).Text = "Conta"
     'cboContas(10).Text = "Conta"
     'txtContas(17).Text = ""
     'txtContas(18).Text = ""
     'txtContas(15).Text = ""
     'txtContas(16).Text = ""
     'chkComplFatoGerador(0).Value = False
     'chkComplFatoGerador(1).Value = False
     'chkComplFatoGerador(2).Value = False
     'chkComplFatoGerador(3).Value = False
  End If

End Sub

Private Sub cmdCancelar_Click()
  Unload Me
  Exit Sub
End Sub

Private Sub cmdGravar_Click()
  If MsgFunc("Deseja gravar dados espec�ficos para " & IIf(Registro = "Lan�amentos", "este Lan�amento?", "esta Duplicata?"), vbQuestion + vbYesNo) = vbYes Then
    
    If ValidarCampos Then
    
      Dim rstIntegracao      As Object
      
      If AbreRecordset(rstIntegracao, "Select * from [Integra��o Cont�bil] where Registro = '" & Registro & "' " & _
                        "and PagRec = '" & PagRec & "' and N�mero = " & Numero & " and Empresa = '" & Empresa & "' " & _
                        "and Tipo = '" & Tipo & "' and Parcela = " & Parcela, dbOpenDynaset) <> WL_ERRO Then
        If rstIntegracao.EOF Then
          rstIntegracao.AddNew
        Else
          If TypeOf rstIntegracao Is dao.Recordset Then rstIntegracao.Edit
        End If
        rstIntegracao("Registro").Value = Registro
        rstIntegracao("PagRec").Value = PagRec
        rstIntegracao("N�mero").Value = Numero
        rstIntegracao("Empresa").Value = Empresa
        rstIntegracao("Tipo").Value = Tipo
        rstIntegracao("Parcela").Value = Parcela
        'Valor Original"
        rstIntegracao("Conta a D�bito - Valor Original").Value = cboContas(0).Text
        rstIntegracao("Conta a D�bito Outros - Valor Original").Value = CLngDef(txtContas(3).Text)
        rstIntegracao("Conta a Cr�dito - Valor Original").Value = cboContas(1).Text
        rstIntegracao("Conta a Cr�dito Outros - Valor Original").Value = CLngDef(txtContas(4).Text)
        rstIntegracao("C�digo do Hist�rico 1 - Valor Original").Value = CLngDef(txtContas(5).Text)
        rstIntegracao("C�digo do Hist�rico 2 - Valor Original").Value = CLngDef(txtContas(6).Text)
        rstIntegracao("Complemento - Valor Original - Descri��o").Value = IIf(chkComplDescrVO(0).Value = vbChecked, True, False)
        rstIntegracao("Complemento - Valor Original - Empresa - D�bito").Value = CIntDef(txtContas(20).Text)
        rstIntegracao("Complemento - Valor Original - Nota - D�bito").Value = CIntDef(txtContas(38).Text)
        rstIntegracao("Complemento - Valor Original - Data - D�bito").Value = CIntDef(txtContas(21).Text)
        rstIntegracao("Complemento - Valor Original - Cheque - D�bito").Value = CIntDef(txtContas(19).Text)
        rstIntegracao("Complemento - Valor Original - Empresa - Cr�dito").Value = CIntDef(txtContas(47).Text)
        rstIntegracao("Complemento - Valor Original - Nota - Cr�dito").Value = CIntDef(txtContas(27).Text)
        rstIntegracao("Complemento - Valor Original - Data - Cr�dito").Value = CIntDef(txtContas(48).Text)
        rstIntegracao("Complemento - Valor Original - Cheque - Cr�dito").Value = CIntDef(txtContas(22).Text)
        'Abatimento")
        rstIntegracao("Conta a D�bito - Abatimento").Value = cboContas(3).Text
        rstIntegracao("Conta a D�bito Outros - Abatimento").Value = CLngDef(txtContas(9).Text)
        rstIntegracao("Conta a Cr�dito - Abatimento").Value = cboContas(4).Text
        rstIntegracao("Conta a Cr�dito Outros - Abatimento").Value = CLngDef(txtContas(10).Text)
        rstIntegracao("C�digo do Hist�rico 1 - Abatimento").Value = CLngDef(txtContas(8).Text)
        rstIntegracao("C�digo do Hist�rico 2 - Abatimento").Value = CLngDef(txtContas(7).Text)
        rstIntegracao("Complemento - Abatimento - Descri��o").Value = IIf(chkComplDescrAbat(0).Value = vbChecked, True, False)
        rstIntegracao("Complemento - Abatimento - Empresa - D�bito").Value = CIntDef(txtContas(25).Text)
        rstIntegracao("Complemento - Abatimento - Nota - D�bito").Value = CIntDef(txtContas(23).Text)
        rstIntegracao("Complemento - Abatimento - Data - D�bito").Value = CIntDef(txtContas(26).Text)
        rstIntegracao("Complemento - Abatimento - Cheque - D�bito").Value = CIntDef(txtContas(24).Text)
        rstIntegracao("Complemento - Abatimento - Empresa - Cr�dito").Value = CIntDef(txtContas(34).Text)
        rstIntegracao("Complemento - Abatimento - Nota - Cr�dito").Value = CIntDef(txtContas(39).Text)
        rstIntegracao("Complemento - Abatimento - Data - Cr�dito").Value = CIntDef(txtContas(32).Text)
        rstIntegracao("Complemento - Abatimento - Cheque - Cr�dito").Value = CIntDef(txtContas(37).Text)
        'Acr�scimo")
        rstIntegracao("Conta a D�bito - Acr�scimo").Value = cboContas(6).Text
        rstIntegracao("Conta a D�bito Outros - Acr�scimo").Value = CLngDef(txtContas(11).Text)
        rstIntegracao("Conta a Cr�dito - Acr�scimo").Value = cboContas(7).Text
        rstIntegracao("Conta a Cr�dito Outros - Acr�scimo").Value = CLngDef(txtContas(12).Text)
        rstIntegracao("C�digo do Hist�rico 1 - Acr�scimo").Value = CLngDef(txtContas(13).Text)
        rstIntegracao("C�digo do Hist�rico 2 - Acr�scimo").Value = CLngDef(txtContas(14).Text)
        rstIntegracao("Complemento - Acr�scimo - Descri��o").Value = IIf(chkComplDescrAcres(0).Value = vbChecked, True, False)
        rstIntegracao("Complemento - Acr�scimo - Empresa - D�bito").Value = CIntDef(txtContas(30).Text)
        rstIntegracao("Complemento - Acr�scimo - Nota - D�bito").Value = CIntDef(txtContas(28).Text)
        rstIntegracao("Complemento - Acr�scimo - Data - D�bito").Value = CIntDef(txtContas(31).Text)
        rstIntegracao("Complemento - Acr�scimo - Cheque - D�bito").Value = CIntDef(txtContas(29).Text)
        rstIntegracao("Complemento - Acr�scimo - Empresa - Cr�dito").Value = CIntDef(txtContas(41).Text)
        rstIntegracao("Complemento - Acr�scimo - Nota - Cr�dito").Value = CIntDef(txtContas(43).Text)
        rstIntegracao("Complemento - Acr�scimo - Data - Cr�dito").Value = CIntDef(txtContas(40).Text)
        rstIntegracao("Complemento - Acr�scimo - Cheque - Cr�dito").Value = CIntDef(txtContas(42).Text)
        'Fato Gerador)
        rstIntegracao("Conta a D�bito - Fato Gerador").Value = cboContas(9).Text
        rstIntegracao("Conta a D�bito Outros - Fato Gerador").Value = CLngDef(txtContas(17).Text)
        rstIntegracao("Conta a Cr�dito - Fato Gerador").Value = cboContas(10).Text
        rstIntegracao("Conta a Cr�dito Outros - Fato Gerador").Value = CLngDef(txtContas(18).Text)
        rstIntegracao("C�digo do Hist�rico 1 - Fato Gerador").Value = CLngDef(txtContas(15).Text)
        rstIntegracao("C�digo do Hist�rico 2 - Fato Gerador").Value = CLngDef(txtContas(16).Text)
        rstIntegracao("Complemento - Fato Gerador - Descri��o").Value = IIf(chkComplDescrFatoGerador(0).Value = vbChecked, True, False)
        rstIntegracao("Complemento - Fato Gerador - Empresa - D�bito").Value = CIntDef(txtContas(35).Text)
        rstIntegracao("Complemento - Fato Gerador - Nota - D�bito").Value = CIntDef(txtContas(33).Text)
        rstIntegracao("Complemento - Fato Gerador - Data - D�bito").Value = CIntDef(txtContas(36).Text)
        rstIntegracao("Complemento - Fato Gerador - Empresa - Cr�dito").Value = CIntDef(txtContas(45).Text)
        rstIntegracao("Complemento - Fato Gerador - Nota - Cr�dito").Value = CIntDef(txtContas(46).Text)
        rstIntegracao("Complemento - Fato Gerador - Data - Cr�dito").Value = CIntDef(txtContas(44).Text)
        rstIntegracao("Fato Gerador").Value = IIf(chkFatoGerador.Value = vbChecked, True, False)
  
        rstIntegracao.update
      End If
      FechaRecordset rstIntegracao
      
      Unload Me
      Exit Sub
      
    End If
  End If
End Sub

Private Sub Form_Load()
 
  Dim X          As Integer
  Dim strSql     As String
  
  For X = 0 To 10
    cboContas(X).AddItem "Conta"
    cboContas(X).AddItem "Banco"
    cboContas(X).AddItem "Empresa"
    cboContas(X).AddItem "Conta Cont�bil Banco"
    cboContas(X).AddItem "Outros"
  
    cboContas(X + 1).AddItem "Conta"
    cboContas(X + 1).AddItem "Banco"
    cboContas(X + 1).AddItem "Empresa"
    cboContas(X + 1).AddItem "Conta Cont�bil Banco"
    cboContas(X + 1).AddItem "Outros"
    
    X = X + 2
  Next

  txtContas(0).Text = Conta

  If Recordcount("Select * from [Integra��o Cont�bil] where Registro = '" & Registro & "' " & _
                      "and PagRec = '" & PagRec & "' and N�mero = " & Numero & " and Empresa = '" & Empresa & "' " & _
                      "and Tipo = '" & Tipo & "' and Parcela = " & Parcela) = 0 Then

    strSql = "SELECT Contas.[Conta a D�bito - Valor Original], Contas.[Conta a D�bito Outros - Valor Original], "
    Concat strSql, "Contas.[Conta a Cr�dito - Valor Original], Contas.[Conta a Cr�dito Outros - Valor Original], "
    Concat strSql, "Contas.[C�digo do Hist�rico 1 - Valor Original], Contas.[C�digo do Hist�rico 2 - Valor Original], "
    Concat strSql, "Contas.[Complemento - Valor Original - Descri��o], "
    Concat strSql, "Contas.[Complemento - Valor Original - Empresa - D�bito], "
    Concat strSql, "Contas.[Complemento - Valor Original - Nota - D�bito], "
    Concat strSql, "Contas.[Complemento - Valor Original - Data - D�bito], "
    Concat strSql, "Contas.[Complemento - Valor Original - Cheque - D�bito], "
    Concat strSql, "Contas.[Complemento - Valor Original - Empresa - Cr�dito], "
    Concat strSql, "Contas.[Complemento - Valor Original - Nota - Cr�dito], "
    Concat strSql, "Contas.[Complemento - Valor Original - Data - Cr�dito], "
    Concat strSql, "Contas.[Complemento - Valor Original - Cheque - Cr�dito], "
    Concat strSql, "Contas.[Conta a D�bito - Abatimento], "
    Concat strSql, "Contas.[Conta a D�bito Outros - Abatimento], Contas.[Conta a Cr�dito - Abatimento], "
    Concat strSql, "Contas.[Conta a Cr�dito Outros - Abatimento], Contas.[C�digo do Hist�rico 1 - Abatimento], "
    Concat strSql, "Contas.[C�digo do Hist�rico 2 - Abatimento],"
    Concat strSql, "Contas.[Complemento - Abatimento - Descri��o],"
    Concat strSql, "Contas.[Complemento - Abatimento - Empresa - D�bito],"
    Concat strSql, "Contas.[Complemento - Abatimento - Nota - D�bito],"
    Concat strSql, "Contas.[Complemento - Abatimento - Data - D�bito],"
    Concat strSql, "Contas.[Complemento - Abatimento - Cheque - D�bito],"
    Concat strSql, "Contas.[Complemento - Abatimento - Empresa - Cr�dito],"
    Concat strSql, "Contas.[Complemento - Abatimento - Nota - Cr�dito],"
    Concat strSql, "Contas.[Complemento - Abatimento - Data - Cr�dito],"
    Concat strSql, "Contas.[Complemento - Abatimento - Cheque - Cr�dito],"
    Concat strSql, "Contas.[Conta a D�bito - Acr�scimo], Contas.[Conta a D�bito Outros - Acr�scimo], "
    Concat strSql, "Contas.[Conta a Cr�dito - Acr�scimo], Contas.[Conta a Cr�dito Outros - Acr�scimo], "
    Concat strSql, "Contas.[C�digo do Hist�rico 1 - Acr�scimo], Contas.[C�digo do Hist�rico 2 - Acr�scimo], "
    Concat strSql, "Contas.[Complemento - Acr�scimo - Descri��o],"
    Concat strSql, "Contas.[Complemento - Acr�scimo - Empresa - D�bito],"
    Concat strSql, "Contas.[Complemento - Acr�scimo - Nota - D�bito],"
    Concat strSql, "Contas.[Complemento - Acr�scimo - Data - D�bito],"
    Concat strSql, "Contas.[Complemento - Acr�scimo - Cheque - D�bito],"
    Concat strSql, "Contas.[Complemento - Acr�scimo - Empresa - Cr�dito],"
    Concat strSql, "Contas.[Complemento - Acr�scimo - Nota - Cr�dito],"
    Concat strSql, "Contas.[Complemento - Acr�scimo - Data - Cr�dito],"
    Concat strSql, "Contas.[Complemento - Acr�scimo - Cheque - Cr�dito],"
    Concat strSql, "Contas.[Conta a D�bito - Fato Gerador], Contas.[Conta a D�bito Outros - Fato Gerador], "
    Concat strSql, "Contas.[Conta a Cr�dito - Fato Gerador], Contas.[Conta a Cr�dito Outros - Fato Gerador], "
    Concat strSql, "Contas.[C�digo do Hist�rico 1 - Fato Gerador], Contas.[C�digo do Hist�rico 2 - Fato Gerador], "
    Concat strSql, "Contas.[Complemento - Fato Gerador - Descri��o],"
    Concat strSql, "Contas.[Complemento - Fato Gerador - Empresa - D�bito],"
    Concat strSql, "Contas.[Complemento - Fato Gerador - Nota - D�bito],"
    Concat strSql, "Contas.[Complemento - Fato Gerador - Data - D�bito], "
    Concat strSql, "Contas.[Complemento - Fato Gerador - Empresa - Cr�dito],"
    Concat strSql, "Contas.[Complemento - Fato Gerador - Nota - Cr�dito],"
    Concat strSql, "Contas.[Complemento - Fato Gerador - Data - Cr�dito], "
    Concat strSql, "Contas.[Fato Gerador] "
    Concat strSql, "FROM Contas "
    Concat strSql, "WHERE C�digo = " & Conta
  
  
    GetAssocValue strSql, cboContas(0), txtContas(3), cboContas(1), txtContas(4), txtContas(5), txtContas(6), _
                          chkComplDescrVO(0), txtContas(20), txtContas(38), txtContas(21), txtContas(19), _
                          txtContas(47), txtContas(27), txtContas(48), txtContas(22), _
                          cboContas(3), txtContas(9), cboContas(4), txtContas(10), txtContas(8), txtContas(7), _
                          chkComplDescrAbat(0), txtContas(25), txtContas(23), txtContas(26), txtContas(24), _
                          txtContas(34), txtContas(39), txtContas(32), txtContas(37), _
                          cboContas(6), txtContas(11), cboContas(7), txtContas(12), txtContas(13), txtContas(14), _
                          chkComplDescrAcres(0), txtContas(30), txtContas(28), txtContas(31), txtContas(29), _
                          txtContas(41), txtContas(43), txtContas(40), txtContas(42), _
                          cboContas(9), txtContas(17), cboContas(10), txtContas(18), txtContas(15), txtContas(16), _
                          chkComplDescrFatoGerador(0), txtContas(35), txtContas(33), txtContas(36), _
                          txtContas(45), txtContas(46), txtContas(44), _
                          chkFatoGerador
                          
  Else
    strSql = "Select [Integra��o Cont�bil].[Conta a D�bito - Valor Original], [Integra��o Cont�bil].[Conta a D�bito Outros - Valor Original], "
    Concat strSql, "[Integra��o Cont�bil].[Conta a Cr�dito - Valor Original], [Integra��o Cont�bil].[Conta a Cr�dito Outros - Valor Original], "
    Concat strSql, "[Integra��o Cont�bil].[C�digo do Hist�rico 1 - Valor Original], [Integra��o Cont�bil].[C�digo do Hist�rico 2 - Valor Original], "
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Descri��o],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Empresa - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Nota - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Data - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Cheque - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Empresa - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Nota - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Data - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Valor Original - Cheque - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Conta a D�bito - Abatimento],"
    Concat strSql, "[Integra��o Cont�bil].[Conta a D�bito Outros - Abatimento], [Integra��o Cont�bil].[Conta a Cr�dito - Abatimento], "
    Concat strSql, "[Integra��o Cont�bil].[Conta a Cr�dito Outros - Abatimento], [Integra��o Cont�bil].[C�digo do Hist�rico 1 - Abatimento], "
    Concat strSql, "[Integra��o Cont�bil].[C�digo do Hist�rico 2 - Abatimento],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Descri��o],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Empresa - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Nota - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Data - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Cheque - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Empresa - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Nota - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Data - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Abatimento - Cheque - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Conta a D�bito - Acr�scimo], [Integra��o Cont�bil].[Conta a D�bito Outros - Acr�scimo], "
    Concat strSql, "[Integra��o Cont�bil].[Conta a Cr�dito - Acr�scimo], [Integra��o Cont�bil].[Conta a Cr�dito Outros - Acr�scimo], "
    Concat strSql, "[Integra��o Cont�bil].[C�digo do Hist�rico 1 - Acr�scimo], [Integra��o Cont�bil].[C�digo do Hist�rico 2 - Acr�scimo], "
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Descri��o],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Empresa - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Nota - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Data - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Cheque - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Empresa - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Nota - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Data - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Acr�scimo - Cheque - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Conta a D�bito - Fato Gerador], [Integra��o Cont�bil].[Conta a D�bito Outros - Fato Gerador], "
    Concat strSql, "[Integra��o Cont�bil].[Conta a Cr�dito - Fato Gerador], [Integra��o Cont�bil].[Conta a Cr�dito Outros - Fato Gerador], "
    Concat strSql, "[Integra��o Cont�bil].[C�digo do Hist�rico 1 - Fato Gerador], [Integra��o Cont�bil].[C�digo do Hist�rico 2 - Fato Gerador], "
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Fato Gerador - Descri��o],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Fato Gerador - Empresa - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Fato Gerador - Nota - D�bito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Fato Gerador - Data - D�bito], "
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Fato Gerador - Empresa - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Fato Gerador - Nota - Cr�dito],"
    Concat strSql, "[Integra��o Cont�bil].[Complemento - Fato Gerador - Data - Cr�dito], "
    Concat strSql, "[Integra��o Cont�bil].[Fato Gerador] "
    Concat strSql, "FROM [Integra��o Cont�bil] "
    Concat strSql, "WHERE Registro = '" & Registro & "' " & "and PagRec = '" & PagRec & "' and N�mero = " & Numero
    Concat strSql, " and Empresa = '" & Empresa & "' " & "and Tipo = '" & Tipo & "' and Parcela = " & Parcela
    
    GetAssocValue strSql, cboContas(0), txtContas(3), cboContas(1), txtContas(4), txtContas(5), txtContas(6), _
                          chkComplDescrVO(0), txtContas(20), txtContas(38), txtContas(21), txtContas(19), _
                          txtContas(47), txtContas(27), txtContas(48), txtContas(22), _
                          cboContas(3), txtContas(9), cboContas(4), txtContas(10), txtContas(8), txtContas(7), _
                          chkComplDescrAbat(0), txtContas(25), txtContas(23), txtContas(26), txtContas(24), _
                          txtContas(34), txtContas(39), txtContas(32), txtContas(37), _
                          cboContas(6), txtContas(11), cboContas(7), txtContas(12), txtContas(13), txtContas(14), _
                          chkComplDescrAcres(0), txtContas(30), txtContas(28), txtContas(31), txtContas(29), _
                          txtContas(41), txtContas(43), txtContas(40), txtContas(42), _
                          cboContas(9), txtContas(17), cboContas(10), txtContas(18), txtContas(15), txtContas(16), _
                          chkComplDescrFatoGerador(0), txtContas(35), txtContas(33), txtContas(36), _
                          txtContas(45), txtContas(46), txtContas(44), _
                          chkFatoGerador
  End If
  
  If chkFatoGerador.Value = vbChecked Then
     fraFatoGerador.Visible = True
  Else
     fraFatoGerador.Visible = False
  End If

End Sub
Private Sub cboContas_Click(Index As Integer)
  
  Select Case Index
    Case 0
      If cboContas(Index).Text <> "Outros" Then txtContas(3).Text = NUL
      txtContas(3).Visible = IIf(cboContas(Index).Text = "Outros", True, False)
    Case 1
      If cboContas(Index).Text <> "Outros" Then txtContas(4).Text = NUL
      txtContas(4).Visible = IIf(cboContas(Index).Text = "Outros", True, False)
    Case 3
      If cboContas(Index).Text <> "Outros" Then txtContas(9).Text = NUL
      txtContas(9).Visible = IIf(cboContas(Index).Text = "Outros", True, False)
    Case 4
      If cboContas(Index).Text <> "Outros" Then txtContas(10).Text = NUL
      txtContas(10).Visible = IIf(cboContas(Index).Text = "Outros", True, False)
    Case 6
      If cboContas(Index).Text <> "Outros" Then txtContas(11).Text = NUL
      txtContas(11).Visible = IIf(cboContas(Index).Text = "Outros", True, False)
    Case 7
      If cboContas(Index).Text <> "Outros" Then txtContas(12).Text = NUL
      txtContas(12).Visible = IIf(cboContas(Index).Text = "Outros", True, False)
    Case 9
      If cboContas(Index).Text <> "Outros" Then txtContas(17).Text = NUL
      txtContas(17).Visible = IIf(cboContas(Index).Text = "Outros", True, False)
    Case 10
      If cboContas(Index).Text <> "Outros" Then txtContas(18).Text = NUL
      txtContas(18).Visible = IIf(cboContas(Index).Text = "Outros", True, False)
  End Select
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmDuplContas = Nothing
End Sub

Private Sub TabStrip1_Click()

End Sub

Private Sub TabIntegracao_Click()
  fraTab(0).Visible = (TabIntegracao.SelectedItem.Key = "FatoPagamento")
  fraTab(1).Visible = (TabIntegracao.SelectedItem.Key = "FatoGerador")
End Sub

Private Sub txtContas_KeyPress(Index As Integer, KeyAscii As Integer)
  DValor KeyAscii
End Sub
Private Function ValidarCampos() As Boolean

  ValidarCampos = False

  If cboContas(0).Text = "Outros" Then
    If CLngDef(txtContas(3).Text) = 0 Then
      MsgFunc "C�digo Conta a D�bito - Valor Original n�o informado."
      Exit Function
    End If
  End If
  If cboContas(1).Text = "Outros" Then
    If CLngDef(txtContas(4).Text) = 0 Then
      MsgFunc "C�digo de Conta a Cr�dito - Valor Original n�o informado."
      Exit Function
    End If
  End If
  If CLngDef(txtContas(5).Text) = 0 Then
    MsgFunc "C�digo do Hist�rico 1 - Valor Original n�o informado."
    Exit Function
  End If

  If Abatimento > 0 Then
    If cboContas(3).Text = "Outros" Then
      If CLngDef(txtContas(9).Text) = 0 Then
        MsgFunc "C�digo Conta a D�bito - Abatimento n�o informado."
        Exit Function
      End If
    End If
    If cboContas(4).Text = "Outros" Then
      If CLngDef(txtContas(10).Text) = 0 Then
        MsgFunc "C�digo de Conta a Cr�dito - Abatimento n�o informado."
        Exit Function
      End If
    End If
    If CLngDef(txtContas(8).Text) = 0 Then
      MsgFunc "C�digo do Hist�rico 1 - Abatimento n�o informado."
      Exit Function
    End If
  End If
  
  If Acrescimo > 0 Then
    If cboContas(6).Text = "Outros" Then
      If CLngDef(txtContas(11).Text) = 0 Then
        MsgFunc "C�digo Conta a D�bito - Acr�scimo n�o informado."
        Exit Function
      End If
    End If
    If cboContas(7).Text = "Outros" Then
      If CLngDef(txtContas(12).Text) = 0 Then
        MsgFunc "C�digo de Conta a Cr�dito - Acr�scimo n�o informado."
        Exit Function
      End If
    End If
    If CLngDef(txtContas(13).Text) = 0 Then
      MsgFunc "C�digo do Hist�rico 1 - Acr�scimo n�o informado."
      Exit Function
    End If
  End If
  
  If chkFatoGerador.Value = vbChecked Then
    
    If cboContas(9).Text = "Outros" Then
      If CLngDef(txtContas(17).Text) = 0 Then
        MsgFunc "C�digo Conta a D�bito - Fato Gerador - N�o informado."
        Exit Function
       End If
    End If
    
    If cboContas(10).Text = "Outros" Then
      If CLngDef(txtContas(18).Text) = 0 Then
        MsgFunc "C�digo Conta a Cr�dito - Fato Gerador - N�o informado."
        Exit Function
      End If
    End If
    
    If CLngDef(txtContas(15).Text) = 0 Then
        MsgFunc "C�digo do Hist�rico 1 - Fato Gerador - N�o informado."
        Exit Function
    End If
    
  End If
  
  ValidarCampos = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oHelpHtml As New clsHelp
    If KeyCode = vbKeyF1 Then
        oHelpHtml.Origem = 0
        oHelpHtml.hWnd = Me.hWnd
        oHelpHtml.HelpContext = Me.HelpContextID
        Call oHelpHtml.ShowHelp
        Set oHelpHtml = Nothing
    End If
End Sub
