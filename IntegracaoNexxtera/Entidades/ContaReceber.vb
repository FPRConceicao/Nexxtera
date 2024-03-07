Imports Teleatlantic.TLS.Common

Public Class ContaReceber : Inherits Retorno
    Public Property TituloSelecionado() As Byte
        Get
            Return m_TituloSelecionado
        End Get
        Set(ByVal value As Byte)
            m_TituloSelecionado = value
        End Set
    End Property
    Private m_TituloSelecionado As Byte

    Public Property NumTit() As String
        Get
            Return m_NumTit
        End Get
        Set(ByVal value As String)
            m_NumTit = value
        End Set
    End Property
    Private m_NumTit As String

    Public Property DescFilial() As String
        Get
            Return m_DescFilial
        End Get
        Set(ByVal value As String)
            m_DescFilial = value
        End Set
    End Property
    Private m_DescFilial As String

    Public Property SeqTit() As String
        Get
            Return m_SeqTit
        End Get
        Set(ByVal value As String)
            m_SeqTit = value
        End Set
    End Property
    Private m_SeqTit As String

    Public Property DtEmissao() As Nullable(Of DateTime)
        Get
            Return m_DtEmissao
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtEmissao = value
        End Set
    End Property
    Private m_DtEmissao As Nullable(Of DateTime)


    Public Property TipoDup() As String
        Get
            Return m_TipoDup
        End Get
        Set(ByVal value As String)
            m_TipoDup = value
        End Set
    End Property
    Private m_TipoDup As String

    Public Property DtVcto() As Nullable(Of DateTime)
        Get
            Return m_DtVcto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtVcto = value
        End Set
    End Property
    Private m_DtVcto As Nullable(Of DateTime)

    Public Property DtVctoDMaisUm() As Nullable(Of DateTime)
        Get
            Return m_DtVctoDMaisUm
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtVctoDMaisUm = value
        End Set
    End Property
    Private m_DtVctoDMaisUm As Nullable(Of DateTime)

    Public Property CodClie() As String
        Get
            Return m_CodClie
        End Get
        Set(ByVal value As String)
            m_CodClie = value
        End Set
    End Property
    Private m_CodClie As String

    Public Property CodInd() As String
        Get
            Return m_CodInd
        End Get
        Set(ByVal value As String)
            m_CodInd = value
        End Set
    End Property
    Private m_CodInd As String

    Public Property VlrInd() As Double
        Get
            Return m_VlrInd
        End Get
        Set(ByVal value As Double)
            m_VlrInd = value
        End Set
    End Property
    Private m_VlrInd As Double

    Public Property Situacao() As String
        Get
            Return m_Situacao
        End Get
        Set(ByVal value As String)
            m_Situacao = value
        End Set
    End Property
    Private m_Situacao As String

    Public Property NumPortador() As String
        Get
            Return m_NumPortador
        End Get
        Set(ByVal value As String)
            m_NumPortador = value
        End Set
    End Property
    Private m_NumPortador As String

    Public Property ObsTit() As String
        Get
            Return m_ObsTit
        End Get
        Set(ByVal value As String)
            m_ObsTit = value
        End Set
    End Property
    Private m_ObsTit As String

    Public Property DtCad() As Nullable(Of DateTime)
        Get
            Return m_DtCad
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtCad = value
        End Set
    End Property
    Private m_DtCad As Nullable(Of DateTime)

    Public Property UsrCad() As String
        Get
            Return m_UsrCad
        End Get
        Set(ByVal value As String)
            m_UsrCad = value
        End Set
    End Property
    Private m_UsrCad As String

    Public Property VlrPrevPgDI() As String
        Get
            Return m_VlrPrevPgDI
        End Get
        Set(ByVal value As String)
            m_VlrPrevPgDI = value
        End Set
    End Property
    Private m_VlrPrevPgDI As String

    Public Property StatusBco() As String
        Get
            Return m_StatusBco
        End Get
        Set(ByVal value As String)
            m_StatusBco = value
        End Set
    End Property
    Private m_StatusBco As String

    Public Property CodCCusto() As String
        Get
            Return m_CodCCusto
        End Get
        Set(ByVal value As String)
            m_CodCCusto = value
        End Set
    End Property
    Private m_CodCCusto As String

    Public Property CodBanco() As String
        Get
            Return m_CodBanco
        End Get
        Set(ByVal value As String)
            m_CodBanco = value
        End Set
    End Property
    Private m_CodBanco As String

    Public Property CodAgen() As String
        Get
            Return m_CodAgen
        End Get
        Set(ByVal value As String)
            m_CodAgen = value
        End Set
    End Property
    Private m_CodAgen As String

    Public Property NumCta() As String
        Get
            Return m_NumCta
        End Get
        Set(ByVal value As String)
            m_NumCta = value
        End Set
    End Property
    Private m_NumCta As String

    Public Property TipoPgtDI() As String
        Get
            Return m_TipoPgtDI
        End Get
        Set(ByVal value As String)
            m_TipoPgtDI = value
        End Set
    End Property
    Private m_TipoPgtDI As String

    Public Property MenBco() As String
        Get
            Return m_MenBco
        End Get
        Set(ByVal value As String)
            m_MenBco = value
        End Set
    End Property
    Private m_MenBco As String

    Public Property TipoPagto() As String
        Get
            Return m_TipoPagto
        End Get
        Set(ByVal value As String)
            m_TipoPagto = value
        End Set
    End Property
    Private m_TipoPagto As String

    Public Property MsgBole() As String
        Get
            Return m_MsgBole
        End Get
        Set(ByVal value As String)
            m_MsgBole = value
        End Set
    End Property
    Private m_MsgBole As String

    Public Property TitDesc() As String
        Get
            Return m_TitDesc
        End Get
        Set(ByVal value As String)
            m_TitDesc = value
        End Set
    End Property
    Private m_TitDesc As String

    Public Property IsTitNegoc() As String
        Get
            Return m_IsTitNegoc
        End Get
        Set(ByVal value As String)
            m_IsTitNegoc = value
        End Set
    End Property
    Private m_IsTitNegoc As String

    Public Property TipoCartaoCred() As String
        Get
            Return m_TipoCartaoCred
        End Get
        Set(ByVal value As String)
            m_TipoCartaoCred = value
        End Set
    End Property
    Private m_TipoCartaoCred As String

    Public Property DtEnvCobExt() As Nullable(Of DateTime)
        Get
            Return m_DtEnvCobExt
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtEnvCobExt = value
        End Set
    End Property
    Private m_DtEnvCobExt As Nullable(Of DateTime)

    Public Property DtTitDesc() As Nullable(Of DateTime)
        Get
            Return m_DtTitDesc
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtTitDesc = value
        End Set
    End Property
    Private m_DtTitDesc As Nullable(Of DateTime)

    Public Property IdFilial() As String
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As String)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As String

    Public Property DtIncSerasa() As Nullable(Of DateTime)
        Get
            Return m_DtIncSerasa
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtIncSerasa = value
        End Set
    End Property
    Private m_DtIncSerasa As Nullable(Of DateTime)

    Public Property DtExcSerasa() As Nullable(Of DateTime)
        Get
            Return m_DtExcSerasa
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtExcSerasa = value
        End Set
    End Property
    Private m_DtExcSerasa As Nullable(Of DateTime) '

    Public Property CodMotExc() As String
        Get
            Return m_CodMotExc
        End Get
        Set(ByVal value As String)
            m_CodMotExc = value
        End Set
    End Property
    Private m_CodMotExc As String

    Public Property VlrProRata() As Double
        Get
            Return m_VlrProRata
        End Get
        Set(ByVal value As Double)
            m_VlrProRata = value
        End Set
    End Property
    Private m_VlrProRata As Double

    Public Property CodDI() As String
        Get
            Return m_CodDI
        End Get
        Set(ByVal value As String)
            m_CodDI = value
        End Set
    End Property
    Private m_CodDI As String

    Public Property NossoNumeroBco() As String
        Get
            Return m_NossoNumeroBco
        End Get
        Set(ByVal value As String)
            m_NossoNumeroBco = value
        End Set
    End Property
    Private m_NossoNumeroBco As String

    Public Property DtPrevPgDI() As Nullable(Of DateTime)
        Get
            Return m_DtPrevPgDI
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtPrevPgDI = value
        End Set
    End Property
    Private m_DtPrevPgDI As Nullable(Of DateTime)

    Public Property CodBancoDeb() As String
        Get
            Return m_CodBancoDeb
        End Get
        Set(ByVal value As String)
            m_CodBancoDeb = value
        End Set
    End Property
    Private m_CodBancoDeb As String

    Public Property CodAgenDeb() As String
        Get
            Return m_CodAgenDeb
        End Get
        Set(ByVal value As String)
            m_CodAgenDeb = value
        End Set
    End Property
    Private m_CodAgenDeb As String

    Public Property NumCtaDeb() As String
        Get
            Return m_NumCtaDeb
        End Get
        Set(ByVal value As String)
            m_NumCtaDeb = value
        End Set
    End Property
    Private m_NumCtaDeb As String

    Public Property StatusCob() As String
        Get
            Return m_StatusCob
        End Get
        Set(ByVal value As String)
            m_StatusCob = value
        End Set
    End Property
    Private m_StatusCob As String

    Public Property DtUltCobranca() As Nullable(Of DateTime)
        Get
            Return m_DtUltCobranca
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtUltCobranca = value
        End Set
    End Property
    Private m_DtUltCobranca As Nullable(Of DateTime)

    Public Property NotaFiscal() As NotaFiscal
        Get
            Return m_NotaFiscal
        End Get
        Set(ByVal value As NotaFiscal)
            m_NotaFiscal = value
        End Set
    End Property
    Private m_NotaFiscal As NotaFiscal

    Public Property BaixaContaReceber() As BaixaContaReceber
        Get
            Return m_BaixaContaReceber
        End Get
        Set(ByVal value As BaixaContaReceber)
            m_BaixaContaReceber = value
        End Set
    End Property
    Private m_BaixaContaReceber As BaixaContaReceber

    Public Property CategoriaPedido() As CategoriaPedido
        Get
            Return m_CategoriaPedido
        End Get
        Set(ByVal value As CategoriaPedido)
            m_CategoriaPedido = value
        End Set
    End Property
    Private m_CategoriaPedido As CategoriaPedido

    Public Property DiasAtraso() As Integer
        Get
            Return m_DiasAtraso
        End Get
        Set(ByVal value As Integer)
            m_DiasAtraso = value
        End Set
    End Property
    Private m_DiasAtraso As Integer


    Public Property VlrCorr() As Double
        Get
            Return m_VlrCorr
        End Get
        Set(ByVal value As Double)
            m_VlrCorr = value
        End Set
    End Property
    Private m_VlrCorr As Double


    Public Property VlrDesconto() As Double
        Get
            Return m_VlrDesconto
        End Get
        Set(ByVal value As Double)
            m_VlrDesconto = value
        End Set
    End Property
    Private m_VlrDesconto As Double


    Public Property DtLimiteDesconto() As Nullable(Of DateTime)
        Get
            Return m_DtLimiteDesconto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtLimiteDesconto = value
        End Set
    End Property
    Private m_DtLimiteDesconto As Nullable(Of DateTime)


    Public Property NFe() As String
        Get
            Return m_NFe
        End Get
        Set(ByVal value As String)
            m_NFe = value
        End Set
    End Property
    Private m_NFe As String


    Public Property NFMonit() As String
        Get
            Return m_NFMonit
        End Get
        Set(ByVal value As String)
            m_NFMonit = value
        End Set
    End Property
    Private m_NFMonit As String


    Public Property UsrAlt() As String
        Get
            Return m_UsrAlt
        End Get
        Set(ByVal value As String)
            m_UsrAlt = value
        End Set
    End Property
    Private m_UsrAlt As String


    Public Property CodIntClie() As String
        Get
            Return m_CodIntClie
        End Get
        Set(ByVal value As String)
            m_CodIntClie = value
        End Set
    End Property
    Private m_CodIntClie As String

    Public Property TaxaMes() As Double
        Get
            Return m_TaxaMes
        End Get
        Set(ByVal value As Double)
            m_TaxaMes = value
        End Set
    End Property
    Private m_TaxaMes As Double

    Public Property TaxaDia() As Double
        Get
            Return m_TaxaDia
        End Get
        Set(ByVal value As Double)
            m_TaxaDia = value
        End Set
    End Property
    Private m_TaxaDia As Double

    Public Property DtAlt() As Nullable(Of DateTime)
        Get
            Return m_DtAlt
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtAlt = value
        End Set
    End Property
    Private m_DtAlt As Nullable(Of DateTime)
    Public Property FisiJuri As String
        Get
            Return m_FisiJuri
        End Get
        Set(ByVal value As String)
            m_FisiJuri = value
        End Set
    End Property
    Private m_FisiJuri As String
    Public Property RazaoSocial As String
        Get
            Return m_RazaoSocial
        End Get
        Set(ByVal value As String)
            m_RazaoSocial = value
        End Set
    End Property
    Private m_RazaoSocial As String

    Public Property Atendente As String
    Public Property UsrAtendente As String
    Public Property idDeposito As Integer
    Public Property NomeFantasia As String
    Public Property Saldo As Double
    Public Property vlrPago As Double
    Public Property QtdTit As Integer
    Public Property TitMesAnt As String
    Public Property ExcecaoEnvioNumNfe As String


    Public Property DtPgto As Nullable(Of DateTime)
    'DtPgto do where na procedure P_AlterarDepositoNaoIdentificadoContasAReceberReintegracao
    Public Property DtPagamento As Nullable(Of DateTime)

    Public Property CondicaoPagamentoVenda As CondicaoPagamentoVenda
    Public Property EnviadoRentSoft As Integer
    Public Property IdBU As Integer
    Public Property SituacaoDesc As String
    Public Property TitulosBU As String
    Public Property linkNFe As String

    Public Property Banco As String
    Public Property NumTitFinanciado As String
    Public Property isTituloFinanciado As Integer
    Public Property DtEmissaoFinanciado As Nullable(Of DateTime)
    Public Property vlrFinanciamento As Double
    Public Property IdFilialFinanciado As String
    Public Property vencido As String
    Public Property descTipoDup As String
    Public Property DtReenvio As String
    Public Property CodOperacaoConta As String
    Public Property QtdFinanciadosAberto As Integer?
    Public Property BancoAgenNumCta As String
    Public Property TituloFinanciado As String
    Public Property VlrAcordo As Double
    Public Property PagSeguro As String
    Public Property ObsBoleto() As String
        Get
            Return m_ObsBoleto
        End Get
        Set(ByVal value As String)
            m_ObsBoleto = value
        End Set
    End Property
    Private m_ObsBoleto As String
    Public Property TextVarEmail() As String
        Get
            Return m_TextVarEmail
        End Get
        Set(ByVal value As String)
            m_TextVarEmail = value
        End Set
    End Property

    Private m_TextVarEmail As String

    Public Property CodReclamacaoJustificativaDesconto As String
        Get
            Return m_CodReclamacaoJustificativaDesconto
        End Get
        Set(value As String)
            m_CodReclamacaoJustificativaDesconto = value
        End Set
    End Property

    Public Property PercBU As Double
        Get
            Return m_PercBU
        End Get
        Set(value As Double)
            m_PercBU = value
        End Set
    End Property

    Private m_CodReclamacaoJustificativaDesconto As String

    Private m_MotivoReversao As String
    Public Property MotivoReversao As String
        Get
            Return m_MotivoReversao
        End Get
        Set(value As String)
            m_MotivoReversao = value
        End Set
    End Property

    Private m_PercBU As Double
    Public Property IsBoletoAtualizado As Integer
    Public Property descCodReclamacaoJustificativaDesconto As String
    Public Property idAgendamento As Integer
    Private m_PercJuros As Double
    Public Property PercJuros As Double
        Get
            Return m_PercJuros
        End Get
        Set(value As Double)
            m_PercJuros = value
        End Set
    End Property
    Public Property VlrJuros As Double

    Private m_Carteira As String
    Public Property Carteira As String
        Get
            Return m_Carteira
        End Get
        Set(value As String)
            m_Carteira = value
        End Set
    End Property

    Private m_CodigoBarra As String
    Public Property CodigoBarra As String
        Get
            Return m_CodigoBarra
        End Get
        Set(value As String)
            m_CodigoBarra = value
        End Set
    End Property
End Class
