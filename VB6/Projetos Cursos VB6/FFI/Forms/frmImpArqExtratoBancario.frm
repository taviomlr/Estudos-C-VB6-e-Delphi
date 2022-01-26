VERSION 5.00
Begin VB.Form frmImpArqExtratoBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de Extrato Bancário"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmImpArqExtratoBancario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   8010
   Begin VB.Frame frameBtns 
      Height          =   2115
      Left            =   6570
      TabIndex        =   2
      Top             =   -45
      Width           =   1395
      Begin VB.CommandButton cmdAjuda 
         Caption         =   "Ajuda"
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton cmdImportar 
         Caption         =   "Importar"
         Height          =   375
         Left            =   90
         TabIndex        =   4
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdSair 
         Caption         =   "Sair"
         Height          =   375
         Left            =   90
         TabIndex        =   3
         Top             =   990
         Width           =   1215
      End
   End
   Begin VB.Frame frameValores 
      Height          =   2115
      Left            =   40
      TabIndex        =   0
      Top             =   -40
      Width           =   6525
      Begin Fox.EBSArquivo ebsImportarArquivo 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   6240
         _ExtentX        =   11165
         _ExtentY        =   582
         Caption         =   "Arquivo"
         TipoTratamento  =   2
         Filter          =   "*.OFC; *.OFX"
      End
      Begin VB.Image imgInformativa 
         Height          =   480
         Left            =   45
         Picture         =   "frmImpArqExtratoBancario.frx":038A
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   $"frmImpArqExtratoBancario.frx":0FCC
         Height          =   1215
         Left            =   30
         TabIndex        =   6
         Top             =   840
         Width           =   6435
      End
   End
End
Attribute VB_Name = "frmImpArqExtratoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAjuda_Click()
    Dim oHelpHtml As New clsHelp
    
    oHelpHtml.Origem = 0
    oHelpHtml.hWnd = Me.hWnd
    oHelpHtml.HelpContext = Me.HelpContextID
    Call oHelpHtml.ShowHelp
    Set oHelpHtml = Nothing
End Sub

Private Sub cmdImportar_Click()
    Dim objImpExtrato           As New ImportacaoExtrato
    Dim objListBancos           As ListaBancos
    Dim objLancamentoExtrato    As New VoImpDigExtratoBancario
    Dim intCont                 As Integer
    Dim intContLanc             As Integer
    Dim objDAO                  As New DaoImpDigExtratoBancario
    Dim objBiz                  As New BizImpDigExtratoBancario
    Dim objBanco                As New DaoBanco
    Dim strDescricaoHistorico   As String
    Dim strDataInicial          As String
    Dim strCodigoBanco          As String
    Dim lngNumExtrato           As Long
    Dim lngHistorico            As Long
    Dim strDataInicialArq       As String
    Dim strDataFinalArq         As String
    Dim lngQtdConciliado        As Long
    Dim lngQtdNaoConciliado     As Long
    Dim FSO                     As New FileSystemObject
    Dim arqTexto                As TextStream
    Dim strTexto                As String
    Dim blnItau                 As Boolean
    
On Error GoTo err_Handler

    If Trim(ebsImportarArquivo.Valor) <> Empty Then
        Aplicacao.Connect
        Aplicacao.BeginTransaction
        'Right(ebsImportarArquivo.Valor, InStr(1,StrReverse(ebsimportararquivo.valor),"\")-1)
        
        Set objListBancos = objImpExtrato.ObterInformacoesExtratoBancos(ebsImportarArquivo.Valor)
        'Set objListBancos = objImpExtrato.ObterInformacoesExtratoBancos(ebsImportarArquivo.NomeArquivo)
        
        If MsgBox("Confirma importação de extrato para o banco?" & Chr(13) & "(" & frmImpDigExtratoBancario.etxBancoInicial.valorInteiro & _
                  " - " & frmImpDigExtratoBancario.etxBancoInicial.ValorDescricao & ")", vbQuestion + vbYesNo) = vbYes Then
            For intCont = 0 To objListBancos.Quantidade - 1
                For intContLanc = 0 To objListBancos.Banco(intCont).lancamentos.Quantidade - 1
                    With objLancamentoExtrato
                        'Guarda a data inicial
                        If intCont = 0 And intContLanc = 0 Then
                            ' 27/02/2019 - FBMI:88 - Yuji F. - Ajuste na validação da data inicial, deve ser o dia seguinte
                            strDataInicialArq = objListBancos.Banco(intCont).lancamentos.PeriodoInicial + 1
                            strDataFinalArq = objListBancos.Banco(intCont).lancamentos.PeriodoFinal
                            'Verifica se já existe algum extrato conciliado neste período
                            lngQtdConciliado = ExisteRegistroConciliado(strDataInicialArq, strDataFinalArq, True, frmImpDigExtratoBancario.etxBancoInicial.valorInteiro)
                            If lngQtdConciliado > 0 Then
                               MsgBox "Não foi possível importar o extrato bancário." & vbNewLine & _
                                      "Existe(m) " & lngQtdConciliado & " lançamento(s) de extrato(s) conciliado(s) para o período de " & Format(strDataInicialArq, "dd/mm/yyyy") & " a " & Format(strDataFinalArq, "dd/mm/yyyy") & ". " & _
                                      "Para incluir novamente o(s) lançamento(s) é necessário desfazer a conciliação do(s) mesmo(s).", vbExclamation, "Atenção"
                                Aplicacao.RollbackTransaction
                                Aplicacao.Disconnect
                                Exit Sub
                            Else
                                lngQtdNaoConciliado = ExisteRegistroConciliado(strDataInicialArq, strDataFinalArq, False, frmImpDigExtratoBancario.etxBancoInicial.valorInteiro)
                                If lngQtdNaoConciliado > 0 Then
                                    If MsgBox("Já existe(m) " & lngQtdNaoConciliado & " lançamento(s) de extrato para o período " & Format(strDataInicialArq, "dd/mm/yyyy") & " a " & Format(strDataFinalArq, "dd/mm/yyyy") & ". " & _
                                       "Deseja sobrepor o(s) mesmo(s)?", vbQuestion + vbYesNo) = vbYes Then
                                       With frmImpDigExtratoBancario
                                            Call objDAO.ExcluirRegistros(.etxBancoInicial.valorInteiro, strDataInicialArq, strDataFinalArq)
                                       End With
                                    Else
                                        Aplicacao.RollbackTransaction
                                        Aplicacao.Disconnect
                                        Exit Sub
                                    End If
                                End If
                            End If
                            strDataInicial = Mid(objListBancos.Banco(intCont).lancamentos.Lancamento(intContLanc).Data, 4, 7)
                            strCodigoBanco = Trim(frmImpDigExtratoBancario.etxBancoInicial.valorInteiro)
                            lngNumExtrato = ProximoCodigoExtrato(strCodigoBanco)
                        End If
                        .CdExtrato = lngNumExtrato
                        .SeqLancExtrato = intContLanc + 1
                        .DataExtrato = objListBancos.Banco(intCont).lancamentos.Lancamento(intContLanc).Data
                        .documento = objListBancos.Banco(intCont).lancamentos.Lancamento(intContLanc).documento
                        .Valor = objListBancos.Banco(intCont).lancamentos.Lancamento(intContLanc).Valor
                        .TipoOperacao = objListBancos.Banco(intCont).lancamentos.Lancamento(intContLanc).natureza
                        .CdBanco = Trim(strCodigoBanco)
                        .Descricao = "" 'Descrição não é definida pois não existe no extrato
                        .ValorInterno = objListBancos.Banco(intCont).lancamentos.Lancamento(intContLanc).Valor 'Verificar com Ivo
                        .Conciliado = False
                        .DataConciliacao = "00:00:00"
                    End With
                    'Verifica se existe o banco e insere
                    If Not objBanco.Existe(Trim(strCodigoBanco)) Then
                        MsgBox ("O Banco com código " & strCodigoBanco & " não existe. Favor cadastrá-lo e fazer a importação novamente (Cadastro > Banco)."), vbInformation, "Importação de Extrato Bancário"
                        Aplicacao.RollbackTransaction
                        Aplicacao.Disconnect
                        Exit Sub
                    End If
                    strDescricaoHistorico = objListBancos.Banco(intCont).lancamentos.Lancamento(intContLanc).Descricao
                    'Verifica se já existe algum histórico com os mesmos parametros para reutilizar
                    lngHistorico = objDAO.ExisteHistorico(strCodigoBanco, strDescricaoHistorico, objListBancos.Banco(intCont).lancamentos.Lancamento(intContLanc).natureza)
                    objLancamentoExtrato.CdHistorico = IIf(lngHistorico > 0, lngHistorico, ProximoCodigoHistorico(strCodigoBanco))
                    If lngHistorico = 0 Then
                        Call objDAO.GravarVOHistorico(objLancamentoExtrato, strDescricaoHistorico, Aplicacao)
                    End If
                    Call objDAO.GravarVO(objLancamentoExtrato, Aplicacao, lngNumExtrato)
                    DoEvents
                Next
                DoEvents
            Next
            
            If frmImpDigExtratoBancario.lblOrigemConciliacao.Caption = "1" Then
                frmConciliacaoTitulosAutomatica.etxExtratoBancario.valorInteiro = lngNumExtrato
                frmConciliacaoTitulosAutomatica.CarregaGridExtrato
            End If
            frmImpDigExtratoBancario.edtEmissaoInicial.MesAno = Format(strDataInicial, "mm/yyyy")
            frmImpDigExtratoBancario.etxExtrato.valorInteiro = lngNumExtrato
            frmImpDigExtratoBancario.CarregaGridLancamentos
            frmImpDigExtratoBancario.fraLanc.Enabled = True
            
            MsgBox "Importação feita com sucesso. Extrato gerado com o número: " & lngNumExtrato & ".", vbInformation, "Importação de Extrato Bancário"
            Aplicacao.CommitTransaction
            Aplicacao.Disconnect
            Unload Me
        End If
    End If
    Exit Sub
err_Handler:
    MsgBox "Não foi possível importar o arquivo. " & vbNewLine & "Certifique-se que as regras em destaque na tela se aplicam ao mesmo."
    'MsgBox err.Description
    Aplicacao.RollbackTransaction
    Aplicacao.Disconnect
End Sub

Private Sub cmdSair_Click()
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

Private Function ProximoCodigoHistorico(ByVal lngBanco As Long) As Long
    Dim strSql    As String
    Dim rstResult As Object
    
    strSql = "SELECT MAX(cd_historico) as ultimoCodigo FROM FFIExtratoBancarioHistorico Where cd_banco = " & lngBanco
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        ProximoCodigoHistorico = strToLng(rstResult.Fields("ultimoCodigo").value & "") + 1
    Else
        ProximoCodigoHistorico = 1
    End If
    Call FechaRecordset(rstResult)
End Function

Public Function ProximoCodigoExtrato(ByVal lngBanco As Long) As Long
    Dim strSql    As String
    Dim rstResult As Object
    
    strSql = "SELECT MAX(cd_extrato) as ultimoCodigo FROM FFIExtratoBancario "
    
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        ProximoCodigoExtrato = strToLng(rstResult.Fields("ultimoCodigo").value & "") + 1
    Else
        ProximoCodigoExtrato = 1
    End If
    Call FechaRecordset(rstResult)
End Function

Private Function ExisteRegistroConciliado(ByVal strDataInicialArq As String, ByVal strDataFinalArq As String, blnConciliado As Boolean, ByVal strBanco As String) As Long
    Dim strSql                  As String
    Dim rstResult               As Object
    
    On Error GoTo err_Handler
    
    strSql = ""
    strSql = strSql & "SELECT COUNT(cd_extrato) as contConciliado "
    strSql = strSql & "FROM FFIExtratoBancario "
    strSql = strSql & "WHERE  CONVERT(varchar(10),data_extrato,120) >= " & InverteData(strDataInicialArq, True) & " and "
    strSql = strSql & "       CONVERT(varchar(10),data_extrato,120) <= " & InverteData(strDataFinalArq, True) & " and conciliado = " & blnConciliado & " and "
    strSql = strSql & "       cd_banco = " & strBanco
                                
    If AbreRecordset(rstResult, strSql) = WL_OK Then
        ExisteRegistroConciliado = rstResult.Fields("contConciliado").value
    Else
        ExisteRegistroConciliado = 0
    End If
    
err_Handler:
    Call FechaRecordset(rstResult)
End Function
