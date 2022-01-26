VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "ComCtl32.ocx"
Begin VB.Form frmImportarDuplicata 
   KeyPreview      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de Duplicatas"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   Icon            =   "frmImportarDuplicata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   6915
   Begin VB.Frame frameBtns 
      Height          =   1455
      Left            =   5460
      TabIndex        =   3
      Top             =   -15
      Width           =   1455
      Begin VB.CommandButton btnAjuda 
         Caption         =   "Ajuda"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton btnImportar 
         Caption         =   "Importar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton btnSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   1215
      End
   End
   Begin VB.Frame frameValores 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   5415
      Begin ComctlLib.ProgressBar pgrBar 
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   900
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   1
      End
      Begin Fox.EBSArquivo ebsImportarArquivo 
         Height          =   525
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   5175
         _ExtentX        =   25638
         _ExtentY        =   926
         Caption         =   "Selecionar o arquivo de importação:"
         PosicaoCaption  =   2
         TipoTratamento  =   2
         Filter          =   "*.txt"
      End
   End
End
Attribute VB_Name = "frmImportarDuplicata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Autor.......: Gustavo Cuman
'Data........: 09/12/2008
'Descrição...:

Private Sub btnAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub btnImportar_Click()
    Dim objArquivoDuplicata As clsArquivoDuplicata
    Dim objLancamento       As clsLancamento
    Dim objHeader           As clsHeader
    Dim blnValido           As Boolean
    Dim blnGravou           As Boolean
    Dim oArqDup As FoxArquivos.ImportDup
    Dim oLeitor As FoxArquivos.ImportacaoArquivo
    Dim oLancamento As FoxArquivos.ImportDupDet
    Dim oTipoGlobal As New CTiposGlobais
    Dim sMsg, sTipoGlobal As String
    Dim iCont As Integer
    
On Error GoTo Error_Handler

    ' 23/04/2019 - FBMI:618 - Yuji F. - Incluído novo processo de leitura dos arquivos,
    'utiliza versão antiga em caso de erro
    Set oArqDup = New FoxArquivos.ImportDup
    Set oLeitor = New FoxArquivos.ImportacaoArquivo
    
    If Dir(ebsImportarArquivo.Valor, vbArchive) <> "" Then
        Call oLeitor.ReadFile(ebsImportarArquivo.Valor, TipoArquivoEnum_Duplicata, oArqDup)
    Else
        MsgBox "O arquivo não existe para ser carregado", vbExclamation, NomeModulo
        GoTo arquivoInvalido
    End If
    
    If Not oArqDup Is Nothing Then
        If (limpaDOC(EmpresaUsuaria.GetCadastroEmpresa.CNPJ_CPF)) <> oArqDup.cnpjEmpresa Then
            sMsg = "CNPJ/CPF da empresa não corresponde com o CNPJ/CPF do arquivo de importação."
            GoTo finalizarRotina
        End If
        pgrBar.Max = oArqDup.lancamentos.Count
        pgrBar.value = 0
        BeginTrans
        
        For iCont = 0 To oArqDup.lancamentos.Count - 1
            Set oLancamento = oArqDup.lancamentos(iCont)
            If Not (CDate(oLancamento.dtEmissao) >= CDate(oArqDup.dtInicial) And _
            CDate(oLancamento.dtEmissao) <= CDate(oArqDup.dtFinal)) Then
                Rollback
                sMsg = "Data de Emissão fora do intervalo."
                GoTo finalizarRotina
            End If
        
            sTipoGlobal = oTipoGlobal.RetornaTipo(oLancamento.TpGlobal)
            If Len(sTipoGlobal) = 0 Then
                Rollback
                sMsg = "Tipo Global " & oLancamento.TpGlobal & " não localizado no cadastro."
                GoTo finalizarRotina
            Else
                oLancamento.TpGlobal = sTipoGlobal
            End If
            
            Set objLancamento = New clsLancamento
            If Not objLancamento.GravarDuplicata(oLancamento) Then
                Rollback
                sMsg = "Não foi possível importar a(s) Duplicata(s)."
                GoTo finalizarRotina
            End If
            
            pgrBar.value = pgrBar + 1
        Next
        
        CommitTrans
        Call MsgBox("Importação da(s) duplicata(s) realizada com sucesso.", vbOKOnly + vbInformation, NomeModulo)
        pgrBar.value = 0
   
    Else
    
        Set objArquivoDuplicata = New clsArquivoDuplicata
        With objArquivoDuplicata
            If .Carregar(ebsImportarArquivo.Valor, IIf(LCase(oLeitor.encoding) <> "utf-8", "x-ansi", "utf-8")) Then
                .PrimeiraLinha
                Set objHeader = .getHeader
                If Not objHeader Is Nothing Then
                    If .getHeader.Validar(limpaDOC(EmpresaUsuaria.GetCadastroEmpresa.CNPJ_CPF)) Then
                        .ProximaLinha
                        BeginTrans
                        pgrBar.Max = .TotalLinhas
                        pgrBar.value = 0
                        blnValido = True
                        blnGravou = True
                        While Not .UltimaLinha And blnValido
                            Set objLancamento = .getLinha
                            sTipoGlobal = oTipoGlobal.RetornaTipo(objLancamento.TipoGlobal)
                            If objLancamento.Validar(.getHeader) And (Len(sTipoGlobal) > 0) Then
                                objLancamento.TipoGlobal = sTipoGlobal
                                If objLancamento.Gravar Then
                                    pgrBar.value = pgrBar + 1
                                    .ProximaLinha
                                Else
                                    blnGravou = False
                                End If
                            Else
                                blnValido = False
                            End If
                        Wend
                        If blnValido And blnGravou Then
                            CommitTrans
                            Call MsgBox("Importação da(s) duplicata(s) realizada com sucesso.", vbOKOnly + vbInformation, NomeModulo)
                            pgrBar.value = 0
                        Else
                            'Correção Importação Duplicata - PT 92531 - Gustavo
                            Rollback
                            If Not blnGravou Then
                                Call MsgBox("Não foi possível importar a(s) Duplicata(s).", vbInformation, NomeModulo)
                            Else
                                Call MsgBox("Data de Emissão fora do intervalo " & _
                                "ou Tipo Global não localizado no cadastro.", vbInformation, NomeModulo)
                            End If
                        End If
                    Else
                        Call MsgBox("CNPJ/CPF da empresa não corresponde com o CNPJ/CPF do arquivo de importação.", vbInformation, NomeModulo)
                    End If
                Else
                    Call MsgBox("Arquivo de importação informado é inválido.", vbInformation, NomeModulo)
                End If
            Else
                Call MsgBox("Não foi possivel carregar o arquivo informado", vbInformation, NomeModulo)
            End If
        End With
        
    End If
    
finalizarRotina:
    If Len(sMsg) > 0 Then Call MsgBox(sMsg, vbInformation, NomeModulo)
arquivoInvalido:
    Set oTipoGlobal = Nothing
    Set objLancamento = Nothing
    Set oLancamento = Nothing
    Set oLeitor = Nothing
    Set oArqDup = Nothing
    Exit Sub
    
Error_Handler:
    
    Call Err.Raise(Err.Number, TypeName(Me), Err.Description)
    Rollback
End Sub

Private Sub btnSair_Click()
    Unload Me
End Sub


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
