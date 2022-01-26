VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeracaoTituloReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngCd_titulo               As Long
Private mstrDescricao               As String
Private mlngNumero_nota             As Long
Private mstrTipo_registro           As String
Private mstrEmpresa                 As String
Private mdblVl_valor_nota           As Double
Private mintIntervalo_vencimento    As Integer
Private mdtaDt_data_emissao         As Date
Private mlngCd_banco                As Long
Private mlngCd_conta                As Long
Private mlngCd_centro_custo         As Long
Private mstrCd_moeda                As String
Private mlngCd_operacao_contabil    As Long
Private mintNr_parcela              As Integer
Private mstrStatus                  As String
Private mbooPermiteAlteracao        As Boolean
'Parcelas.
'Private mstrP_PagRec                As String
'Private mlngP_Nota                  As Long
'Private mstrP_Empresa               As String
'Private mstrP_Tipo                  As String
'Private mintP_Parcela               As Integer
'Private mstrP_Descricao             As String
'Private mdtaP_Emissao               As Date
'Private mdtaP_Vencimento            As Date
'Private mdtaP_Pagamento             As String
'Private mdblP_Valor_Original        As Double
'Private mlngP_Banco                 As Long
'Private mlngP_Conta                 As Long
'Private mlngP_Centro                As Long
'Private mstrP_Moeda                 As String
'Private mlngP_cd_operacao_contabil  As Long
'Private mobjParcelas                As cColecaoGeracaoTituloReceber
'Rateio
Private mlngR_Cd_titulo             As Long
Private mlngR_Cd_centro_custo       As Long
Private mlngR_Cd_conta              As Long
Private mdblR_Percentual            As Double
Private mobjRateio                  As cColecaoRateioTituloReceber
'pt. 85684 - Moacir Pfau(02/07/2008) - GERAÇÃO DE TITULOS

Public Property Let Cd_Titulo(ByVal NewVal As Long)
    mlngCd_titulo = NewVal
End Property

Public Property Get Cd_Titulo() As Long
    Cd_Titulo = mlngCd_titulo
End Property

Public Property Let Descricao(ByVal NewVal As String)
    mstrDescricao = NewVal
End Property

Public Property Get Descricao() As String
    Descricao = mstrDescricao
End Property

Public Property Let Numero_nota(ByVal NewVal As Long)
    mlngNumero_nota = NewVal
End Property

Public Property Get Numero_nota() As Long
    Numero_nota = mlngNumero_nota
End Property

Public Property Let Tipo_registro(ByVal NewVal As String)
    mstrTipo_registro = NewVal
End Property

Public Property Get Tipo_registro() As String
    Tipo_registro = mstrTipo_registro
End Property

Public Property Let Empresa(ByVal NewVal As String)
    mstrEmpresa = NewVal
End Property

Public Property Get Empresa() As String
    Empresa = mstrEmpresa
End Property

Public Property Let Vl_valor_nota(ByVal NewVal As Double)
    mdblVl_valor_nota = NewVal
End Property

Public Property Get Vl_valor_nota() As Double
    Vl_valor_nota = mdblVl_valor_nota
End Property

Public Property Let Intervalo_vencimento(ByVal NewVal As Integer)
    mintIntervalo_vencimento = NewVal
End Property

Public Property Get Intervalo_vencimento() As Integer
    Intervalo_vencimento = mintIntervalo_vencimento
End Property

Public Property Let Dt_data_emissao(ByVal NewVal As Date)
    mdtaDt_data_emissao = NewVal
End Property

Public Property Get Dt_data_emissao() As Date
    Dt_data_emissao = mdtaDt_data_emissao
End Property

Public Property Let Cd_banco(ByVal NewVal As Long)
    mlngCd_banco = NewVal
End Property

Public Property Get Cd_banco() As Long
    Cd_banco = mlngCd_banco
End Property

Public Property Let Cd_conta(ByVal NewVal As Long)
    mlngCd_conta = NewVal
End Property

Public Property Get Cd_conta() As Long
    Cd_conta = mlngCd_conta
End Property

Public Property Let Cd_centro_custo(ByVal NewVal As Long)
    mlngCd_centro_custo = NewVal
End Property

Public Property Get Cd_centro_custo() As Long
    Cd_centro_custo = mlngCd_centro_custo
End Property

Public Property Let Cd_moeda(ByVal NewVal As String)
    mstrCd_moeda = NewVal
End Property

Public Property Get Cd_moeda() As String
    Cd_moeda = mstrCd_moeda
End Property

Public Property Let cd_operacao_contabil(ByVal NewVal As Long)
    mlngCd_operacao_contabil = NewVal
End Property

Public Property Get cd_operacao_contabil() As Long
    cd_operacao_contabil = mlngCd_operacao_contabil
End Property

Public Property Let nr_parcela(ByVal NewVal As Integer)
    mintNr_parcela = NewVal
End Property

Public Property Get nr_parcela() As Integer
    nr_parcela = mintNr_parcela
End Property

Public Property Let status(ByVal NewVal As String)
    mstrStatus = NewVal
End Property

Public Property Get status() As String
    status = mstrStatus
End Property

'Public Property Let P_PagRec(ByVal NewVal As String)
'    mstrP_PagRec = NewVal
'End Property
'
'Public Property Get P_PagRec() As String
'    P_PagRec = mstrP_PagRec
'End Property
'
'Public Property Let P_Nota(ByVal NewVal As Long)
'    mlngP_Nota = NewVal
'End Property
'
'Public Property Get P_Nota() As Long
'    P_Nota = mlngP_Nota
'End Property
'
'Public Property Let P_Empresa(ByVal NewVal As String)
'    mstrP_Empresa = NewVal
'End Property
'
'Public Property Get P_Empresa() As String
'    P_Empresa = mstrP_Empresa
'End Property
'
'Public Property Let P_Tipo(ByVal NewVal As String)
'    mstrP_Tipo = NewVal
'End Property
'
'Public Property Get P_Tipo() As String
'    P_Tipo = mstrP_Tipo
'End Property
'
'Public Property Let P_Parcela(ByVal NewVal As Integer)
'    mintP_Parcela = NewVal
'End Property
'
'Public Property Get P_Parcela() As Integer
'    P_Parcela = mintP_Parcela
'End Property
'
'Public Property Let P_Descricao(ByVal NewVal As String)
'    mstrP_Descricao = NewVal
'End Property
'
'Public Property Get P_Descricao() As String
'    P_Descricao = mstrP_Descricao
'End Property
'
'Public Property Let P_Emissao(ByVal NewVal As Date)
'    mdtaP_Emissao = NewVal
'End Property
'
'Public Property Get P_Emissao() As Date
'    P_Emissao = mdtaP_Emissao
'End Property
'
'Public Property Let P_Vencimento(ByVal NewVal As Date)
'    mdtaP_Vencimento = NewVal
'End Property
'
'Public Property Get P_Vencimento() As Date
'    P_Vencimento = mdtaP_Vencimento
'End Property
'
'Public Property Let P_Pagamento(ByVal NewVal As String)
'    mdtaP_Pagamento = NewVal
'End Property
'
'Public Property Get P_Pagamento() As String
'    P_Pagamento = mdtaP_Pagamento
'End Property
'
'Public Property Let P_Valor_Original(ByVal NewVal As Double)
'    mdblP_Valor_Original = NewVal
'End Property
'
'Public Property Get P_Valor_Original() As Double
'    P_Valor_Original = mdblP_Valor_Original
'End Property
'
'Public Property Let P_Banco(ByVal NewVal As Long)
'    mlngP_Banco = NewVal
'End Property
'
'Public Property Get P_Banco() As Long
'    P_Banco = mlngP_Banco
'End Property
'
'Public Property Let P_Conta(ByVal NewVal As Long)
'    mlngP_Conta = NewVal
'End Property
'
'Public Property Get P_Conta() As Long
'    P_Conta = mlngP_Conta
'End Property
'
'Public Property Let P_Centro(ByVal NewVal As Long)
'    mlngP_Centro = NewVal
'End Property
'
'Public Property Get P_Centro() As Long
'    P_Centro = mlngP_Centro
'End Property
'
'Public Property Let P_Moeda(ByVal NewVal As String)
'    mstrP_Moeda = NewVal
'End Property
'
'Public Property Get P_Moeda() As String
'    P_Moeda = mstrP_Moeda
'End Property
'
'Public Property Let P_cd_operacao_contabil(ByVal NewVal As Long)
'    mlngP_cd_operacao_contabil = NewVal
'End Property
'
'Public Property Get P_cd_operacao_contabil() As Long
'    P_cd_operacao_contabil = mlngP_cd_operacao_contabil
'End Property

Public Property Get PermiteAlteracao() As Boolean
    PermiteAlteracao = mbooPermiteAlteracao
End Property

Public Property Let PermiteAlteracao(ByVal NewVal As Boolean)
    mbooPermiteAlteracao = NewVal
End Property

Private Sub Class_Initialize()
    mbooPermiteAlteracao = True
    'Set mobjParcelas = New cColecaoGeracaoTituloReceber
    Set mobjRateio = New cColecaoRateioTituloReceber
End Sub

Public Function CriarObjeto(rd As IDBReader) As cGeracaoTituloReceber
    With rd
        mlngCd_titulo = .GetLong("Cd_titulo")
        mstrDescricao = .GetString("Descricao")
        mlngNumero_nota = .GetLong("Numero_nota")
        mstrTipo_registro = .GetString("Tipo_registro")
        mstrEmpresa = .GetString("Empresa")
        mdblVl_valor_nota = .GetDouble("Vl_valor_nota")
        mintIntervalo_vencimento = .GetInteger("Intervalo_vencimento")
        mdtaDt_data_emissao = .GetDate("Dt_data_emissao")
        mlngCd_banco = .GetLong("Cd_banco")
        mlngCd_conta = .GetLong("Cd_conta")
        mlngCd_centro_custo = .GetLong("Cd_centro_custo")
        mstrCd_moeda = .GetString("Cd_moeda")
        mlngCd_operacao_contabil = .GetLong("Cd_operacao_contabil")
        mintNr_parcela = .GetInteger("Nr_parcela")
        mstrStatus = .GetString("Status")
        Set CriarObjeto = Me
    End With
End Function

''Parcelas
'Public Property Get parcelas() As cColecaoGeracaoTituloReceber
'    Set parcelas = mobjParcelas
'End Property
'
'Public Property Let parcelas(ByVal NewVal As cColecaoGeracaoTituloReceber)
'    Set mobjParcelas = NewVal
'End Property

Private Sub Class_Terminate()
    'Set mobjParcelas = Nothing
    Set mobjRateio = Nothing
End Sub

'Rateio
Public Property Let R_Cd_titulo(NewVal As Long)
    mlngR_Cd_titulo = NewVal
End Property

Public Property Get R_Cd_titulo() As Long
    R_Cd_titulo = mlngR_Cd_titulo
End Property

Public Property Let R_Cd_centro_custo(ByVal NewVal As Long)
    mlngR_Cd_centro_custo = NewVal
End Property

Public Property Get R_Cd_centro_custo() As Long
    R_Cd_centro_custo = mlngR_Cd_centro_custo
End Property

Public Property Let R_Cd_conta(ByVal NewVal As Long)
    mlngR_Cd_conta = NewVal
End Property

Public Property Get R_Cd_conta() As Long
    R_Cd_conta = mlngR_Cd_conta
End Property

Public Property Let R_Percentual(ByVal NewVal As Double)
    mdblR_Percentual = NewVal
End Property

Public Property Get R_Percentual() As Double
    R_Percentual = mdblR_Percentual
End Property

Public Property Get Rateio() As cColecaoRateioTituloReceber
    Set Rateio = mobjRateio
End Property

Public Property Let Rateio(ByVal NewVal As cColecaoRateioTituloReceber)
    Set mobjRateio = NewVal
End Property

Public Function equals(obj As Object) As Boolean
    Dim booIgual As Boolean
    Dim newObj As cGeracaoTituloReceber
    
    Set newObj = obj
    booIgual = (newObj.R_Cd_titulo = mlngR_Cd_titulo)
    booIgual = booIgual And (newObj.R_Cd_centro_custo = mlngR_Cd_centro_custo)
    booIgual = booIgual And (newObj.R_Cd_conta = mlngR_Cd_conta)
    equals = booIgual
End Function
