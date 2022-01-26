VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImpressaoBoleto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngEnterprise_id                       As Long
Private mlngCd_estabelecimento                  As Long
Private mlngBanco                               As Long
Private mlngId_carteira                         As Long
Private mstrTipofiltro                          As String
Private mstrOrigem                              As String
Private mdteData_inicial                        As Date
Private mdteData_final                          As Date
Private mlngNumero_inicial                      As Long
Private mlngParcela_inicial                     As Long
Private mlngNumero_final                        As Long
Private mlngParcela_final                       As Long
Private mstrEmpresa                             As String
Private mstrMensagemValidacao                   As String
Private mbolReimpressao                         As Boolean
Private mintStatusEmail                         As ENUMEmail
Private mApp                                    As IApplication
Private mcolTitulo                              As clsColTituloCobrebem

Public Enum ENUMEmail
    e_Enviados = 1
    e_Pendentes = 2
    e_Todos = 3
End Enum

Public Sub init(ByVal objApp As IApplication)
    Set mApp = objApp
End Sub

Public Property Let Enterprise_id(Valor As Long)
   mlngEnterprise_id = Valor
End Property

Public Property Get Enterprise_id() As Long
   Enterprise_id = mlngEnterprise_id
End Property

Public Property Let Cd_estabelecimento(Valor As Long)
   mlngCd_estabelecimento = Valor
End Property

Public Property Get Cd_estabelecimento() As Long
   Cd_estabelecimento = mlngCd_estabelecimento
End Property

Public Property Let Banco(Valor As Long)
   mlngBanco = Valor
End Property

Public Property Get Banco() As Long
   Banco = mlngBanco
End Property

Public Property Let Id_carteira(Valor As Long)
   mlngId_carteira = Valor
End Property

Public Property Get Id_carteira() As Long
   Id_carteira = mlngId_carteira
End Property


Public Property Let Tipofiltro(Valor As String)
   mstrTipofiltro = Valor
End Property

Public Property Get Tipofiltro() As String
   Tipofiltro = mstrTipofiltro
End Property

Public Property Let Origem(Valor As String)
   mstrOrigem = Valor
End Property

Public Property Get Origem() As String
   Origem = mstrOrigem
End Property

Public Property Let Empresa(Valor As String)
   mstrEmpresa = Valor
End Property

Public Property Get Empresa() As String
   Empresa = mstrEmpresa
End Property

Public Property Let Data_inicial(Valor As Date)
   mdteData_inicial = Valor
End Property

Public Property Get Data_inicial() As Date
   Data_inicial = mdteData_inicial
End Property


Public Property Let Data_final(Valor As Date)
   mdteData_final = Valor
End Property

Public Property Get Data_final() As Date
   Data_final = mdteData_final
End Property


Public Property Let Numero_inicial(Valor As Long)
   mlngNumero_inicial = Valor
End Property

Public Property Get Numero_inicial() As Long
   Numero_inicial = mlngNumero_inicial
End Property


Public Property Let Parcela_inicial(Valor As Long)
   mlngParcela_inicial = Valor
End Property

Public Property Get Parcela_inicial() As Long
   Parcela_inicial = mlngParcela_inicial
End Property


Public Property Let Numero_final(Valor As Long)
   mlngNumero_final = Valor
End Property

Public Property Get Numero_final() As Long
   Numero_final = mlngNumero_final
End Property


Public Property Let Parcela_final(Valor As Long)
   mlngParcela_final = Valor
End Property

Public Property Get Parcela_final() As Long
   Parcela_final = mlngParcela_final
End Property


Public Property Let Reimpressao(Valor As Boolean)
   mbolReimpressao = Valor
End Property

Public Property Get Reimpressao() As Boolean
   Reimpressao = mbolReimpressao
End Property

Public Property Let StatusEmail(Valor As ENUMEmail)
   mintStatusEmail = Valor
End Property

Public Property Get StatusEmail() As ENUMEmail
   StatusEmail = mintStatusEmail
End Property

Public Property Get MensagemValidacao() As String
   MensagemValidacao = mstrMensagemValidacao
End Property

Public Property Get colTitulo() As clsColTituloCobrebem
    Set colTitulo = mcolTitulo
End Property

Public Property Let colTitulo(ByVal NewVal As clsColTituloCobrebem)
    Set mcolTitulo = NewVal
End Property

Public Function carregarConsulta() As clsColTituloCobrebem
    Dim cmd             As IDBSelectCommand
    Dim rdResult        As IDBReader
    Dim result          As clsTituloCobrebem

On Error GoTo erro_carregando
    Set carregarConsulta = New clsColTituloCobrebem
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = "[" & mstrOrigem & "]"
                           
        Call .Filter.Append("[PagRec] = @pPagRec")
        Call .Parameters.add(.CreateParameter("@pPagRec", "R", dbFieldTypeString))
        
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", mlngBanco, dbFieldTypeLong))
        
        'Obrigatorio
        If mstrTipofiltro = "Emissão" Then
            Call .Filter.Append("[Emissão] >= @pEmissao_inicial")
            Call .Parameters.add(.CreateParameter("@pEmissao_inicial", mdteData_inicial, dbFieldTypeDate))
            
            Call .Filter.Append("[Emissão] <= @pEmissao_final")
            Call .Parameters.add(.CreateParameter("@pEmissao_final", mdteData_final, dbFieldTypeDate))
        ElseIf mstrTipofiltro = "Vencimento" Then
            Call .Filter.Append("[Vencimento] >= @pVencimento_inicial")
            Call .Parameters.add(.CreateParameter("@pVencimento_inicial", mdteData_inicial, dbFieldTypeDate))
            
            Call .Filter.Append("[Vencimento] <= @pVencimento_final")
            Call .Parameters.add(.CreateParameter("@pVencimento_final", mdteData_final, dbFieldTypeDate))
        ElseIf mstrTipofiltro = "Liberação" Then
            Call .Filter.Append("[Liberação] >= @pLiberacao_inicial")
            Call .Parameters.add(.CreateParameter("@pLiberacao_inicial", mdteData_inicial, dbFieldTypeDate))
            
            Call .Filter.Append("[Liberação] <= @pLiberacao_final")
            Call .Parameters.add(.CreateParameter("@pLiberacao_final", mdteData_final, dbFieldTypeDate))
        End If
        
        'Não obrigatorio
        If mlngNumero_inicial > 0 Then
            If mstrOrigem = "Duplicatas" Then
                Call .Filter.Append("[Nota] >= @pCodigoInicial")
            ElseIf mstrOrigem = "Lançamentos" Then
                Call .Filter.Append("[Código] >= @pCodigoInicial")
            End If
            Call .Parameters.add(.CreateParameter("@pCodigoInicial", mlngNumero_inicial, dbFieldTypeLong))
            If mlngParcela_inicial > 0 Then
                Call .Filter.Append("[Parcela] >= @pParcelaInicial")
                Call .Parameters.add(.CreateParameter("@pParcelaInicial", mlngParcela_inicial, dbFieldTypeLong))
            End If
        End If
        
        'Não obrigatorio
        If mlngNumero_final > 0 Then
            If mstrOrigem = "Duplicatas" Then
                Call .Filter.Append("[Nota] <= @pCodigoFinal")
            ElseIf mstrOrigem = "Lançamentos" Then
                Call .Filter.Append("[Código] <= @pCodigoFinal")
            End If
            Call .Parameters.add(.CreateParameter("@pCodigoFinal", mlngNumero_final, dbFieldTypeLong))
            If mlngParcela_final > 0 Then
                Call .Filter.Append("[Parcela] <= @pParcelafinal")
                Call .Parameters.add(.CreateParameter("@pParcelafinal", mlngParcela_final, dbFieldTypeLong))
            End If
        End If
        
        'Não obrigatorio
        If Len(Trim(mstrEmpresa)) > 0 Then
            Call .Filter.Append("[Empresa] = @pEmpresa")
            Call .Parameters.add(.CreateParameter("@pEmpresa", mstrEmpresa, dbFieldTypeString, 15))
        End If
                        
        If mintStatusEmail <> ENUMEmail.e_Todos Then
            If mintStatusEmail = ENUMEmail.e_Enviados Then
                Call .Filter.Append("[emailboleto_enviado] = @pEmailBoleto_Enviado")
                Call .Parameters.add(.CreateParameter("@pEmailBoleto_Enviado", True, dbFieldTypeBool))
            ElseIf mintStatusEmail = ENUMEmail.e_Pendentes Then
                Call .Filter.Append("[emailboleto_enviado] = @pEmailBoleto_Enviado")
                Call .Parameters.add(.CreateParameter("@pEmailBoleto_Enviado", False, dbFieldTypeBool))
            End If
        End If
                   
        'Não obrigatorio
        'Projeto: #1203 - História: #10582 - Desenvolvimento#12134 - João Henrique(19/04/2012)
        If mbolReimpressao = True Then
            'Vinicius Elyseu(01/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
            Call .Filter.Append("boleto > 0")
            'Call .Filter.Append("NOT ([NOSNUM]IS NULL AND Id_carteira>0)")
        Else
            'Vinicius Elyseu(01/03/2016) - Projeto: #0 - História: #0 - Desenv: #0
            Call .Filter.Append("boleto = 0")
            'Call .Filter.Append("([NOSNUM] IS NULL OR [NOSNUM]='') AND (Id_carteira=0 OR Id_carteira IS NULL)")
        End If
        
        Call .Filter.Append("[Pagamento] IS NULL")
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        rdResult.MoveFirst
        mstrMensagemValidacao = ""
        While Not rdResult.EOF
            Set result = New clsTituloCobrebem
            result.Origem = mstrOrigem
            Call result.CriarObjeto(rdResult)
            'Validação boleto
            If fValidaBoleto(result) Then
                carregarConsulta.add result
            End If
            Set result = Nothing
            rdResult.MoveNext
        Wend
'    Else
'        Set carregarConsulta = Nothing
    End If
    rdResult.CloseReader
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function

erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    Set carregarConsulta = Nothing
    'Resume
End Function

Private Function fValidaBoleto(ByVal obj As clsTituloCobrebem) As Boolean
    Dim strMensagem             As String
    
On Error GoTo err
    strMensagem = ""
    If Trim(obj.Cnpjsacado) = "" And Trim(obj.Cpfsacado) = "" Then
        strMensagem = strMensagem & " Para o documento: " & Trim(obj.NumeroDocumento) & " e parcela: " & Trim(obj.Parcela) & " Empresa: " & Trim(obj.Empresa) & Space(15 - Len(Trim(obj.Empresa))) & " , CNPJ/CPF é de preenchimento obrigatório." & vbCrLf
    End If
   
    If Trim(obj.Enderecosacado) = "" Then
        strMensagem = strMensagem & " Para o documento: " & Trim(obj.NumeroDocumento) & " e parcela: " & Trim(obj.Parcela) & " Empresa: " & Trim(obj.Empresa) & Space(15 - Len(Trim(obj.Empresa))) & " , Endereço do sacado é de preenchimento obrigatório." & vbCrLf
    End If
   
    If Trim(obj.BairroSacado) = "" Then
        strMensagem = strMensagem & " Para o documento: " & Trim(obj.NumeroDocumento) & " e parcela: " & Trim(obj.Parcela) & " Empresa: " & Trim(obj.Empresa) & Space(15 - Len(Trim(obj.Empresa))) & " , Bairro do sacado é de preenchimento obrigatório." & vbCrLf
    End If
    
    If Trim(obj.Cidadesacado) = "" Then
        strMensagem = strMensagem & " Para o documento: " & Trim(obj.NumeroDocumento) & " e parcela: " & Trim(obj.Parcela) & " Empresa: " & Trim(obj.Empresa) & Space(15 - Len(Trim(obj.Empresa))) & " , Cidade do sacado é de preenchimento obrigatório." & vbCrLf
    End If
    
    If Trim(obj.Estadosacado) = "" Then
        strMensagem = strMensagem & " Para o documento: " & Trim(obj.NumeroDocumento) & " e parcela: " & Trim(obj.Parcela) & " Empresa: " & Trim(obj.Empresa) & Space(15 - Len(Trim(obj.Empresa))) & " , Estado do sacado é de preenchimento obrigatório." & vbCrLf
    End If
   
    If Trim(obj.Cepsacado) = "" Then
        strMensagem = strMensagem & " Para o documento: " & Trim(obj.NumeroDocumento) & " e parcela: " & Trim(obj.Parcela) & " Empresa: " & Trim(obj.Empresa) & Space(15 - Len(Trim(obj.Empresa))) & " , CEP do sacado é de preenchimento obrigatório." & vbCrLf
    End If
   
    If strMensagem = "" Then
        fValidaBoleto = True
    Else
        mstrMensagemValidacao = mstrMensagemValidacao & strMensagem
    End If
    
    Exit Function
err:
    fValidaBoleto = False

End Function

Public Function ExisteCarteira(ByVal lngBanco As Long, ByVal lngId_carteira As Long) As Boolean
    Dim cmd             As IDBSelectCommand
    Dim rdResult        As IDBReader
    Dim result          As clsCarteira

On Error GoTo erro_carregando
    Set cmd = mApp.CreateSelectCommand
    With cmd
        .Table.TableName = "[FFIBanco_carteira]"
        Call .Filter.Append("[Banco] = @pBanco")
        Call .Parameters.add(.CreateParameter("@pBanco", lngBanco, dbFieldTypeLong))
        
        Call .Filter.Append("[id_carteira] = @pId_carteira")
        Call .Parameters.add(.CreateParameter("@pId_carteira", lngId_carteira, dbFieldTypeLong))
    End With
    Set rdResult = mApp.ExecuteReader(mApp.GetInternalAuthorization, cmd)
    If Not rdResult.EOF Then
        ExisteCarteira = True
    Else
        ExisteCarteira = False
    End If
    Set cmd = Nothing
    Set rdResult = Nothing
    Exit Function

erro_carregando:
    Set cmd = Nothing
    Set rdResult = Nothing
    ExisteCarteira = False
End Function

Private Sub Class_Initialize()
    Set mcolTitulo = New clsColTituloCobrebem
End Sub

Private Sub Class_Terminate()
    Set mcolTitulo = Nothing
End Sub
