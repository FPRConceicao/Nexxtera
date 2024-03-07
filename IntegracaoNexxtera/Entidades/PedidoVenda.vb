Imports Teleatlantic.TLS.Common

Public Class PedidoVenda : Inherits Retorno

    Public Property NumPedVda() As String
        Get
            Return m_NumPedVda
        End Get
        Set(ByVal value As String)
            m_NumPedVda = value
        End Set
    End Property
    Private m_NumPedVda As String

    Public Property StatusFicha() As String
        Get
            Return m_StatusFicha
        End Get
        Set(ByVal value As String)
            m_StatusFicha = value
        End Set
    End Property
    Private m_StatusFicha As String
    Public Property CodEmpresa() As String
        Get
            Return m_CodEmpresa
        End Get
        Set(ByVal value As String)
            m_CodEmpresa = value
        End Set
    End Property
    Private m_CodEmpresa As String

    Public Property TipoPedido() As String
        Get
            Return m_TipoPedido
        End Get
        Set(ByVal value As String)
            m_TipoPedido = value
        End Set
    End Property
    Private m_TipoPedido As String

    Public Property TipoOrcamento() As String
        Get
            Return m_TipoOrcamento
        End Get
        Set(ByVal value As String)
            m_TipoOrcamento = value
        End Set
    End Property
    Private m_TipoOrcamento As String

    Public Property Orcamento() As Orcamento
        Get
            Return m_Orcamento
        End Get
        Set(ByVal value As Orcamento)
            m_Orcamento = value
        End Set
    End Property
    Private m_Orcamento As Orcamento

    Public Property CodVend() As String
        Get
            Return m_CodVend
        End Get
        Set(ByVal value As String)
            m_CodVend = value
        End Set
    End Property
    Private m_CodVend As String
    Public Property codClie() As String
        Get
            Return m_codClie
        End Get
        Set(value As String)
            m_codClie = value
        End Set
    End Property
    Private m_codClie As String
    Public Property CodCpgt() As String
        Get
            Return m_CodCpgt
        End Get
        Set(ByVal value As String)
            m_CodCpgt = value
        End Set
    End Property
    Private m_CodCpgt As String

    Public Property VlrDescVenda() As Double
        Get
            Return m_VlrDescVenda
        End Get
        Set(ByVal value As Double)
            m_VlrDescVenda = value
        End Set
    End Property
    Private m_VlrDescVenda As Double

    Public Property DescProd() As String
        Get
            Return m_DescProd
        End Get
        Set(ByVal value As String)
            m_DescProd = value
        End Set
    End Property
    Private m_DescProd As String

    Public Property TemIcms() As String
        Get
            Return m_TemIcms
        End Get
        Set(ByVal value As String)
            m_TemIcms = value
        End Set
    End Property
    Private m_TemIcms As String

    Public Property StatusNF() As String
        Get
            Return m_StatusNF
        End Get
        Set(ByVal value As String)
            m_StatusNF = value
        End Set
    End Property
    Private m_StatusNF As String


    Public Property ObsInstal() As String
        Get
            Return m_ObsInstal
        End Get
        Set(ByVal value As String)
            m_ObsInstal = value
        End Set
    End Property
    Private m_ObsInstal As String


    Public Property CodInstal() As String
        Get
            Return m_CodInstal
        End Get
        Set(ByVal value As String)
            m_CodInstal = value
        End Set
    End Property
    Private m_CodInstal As String


    Public Property CodStInstal() As String
        Get
            Return m_CodStInstal
        End Get
        Set(ByVal value As String)
            m_CodStInstal = value
        End Set
    End Property
    Private m_CodStInstal As String


    Public Property DtPrevInstal() As Nullable(Of DateTime)
        Get
            Return m_DtPrevInstal
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtPrevInstal = value
        End Set
    End Property
    Private m_DtPrevInstal As Nullable(Of DateTime)


    Public Property Pendencia() As String
        Get
            Return m_Pendencia
        End Get
        Set(ByVal value As String)
            m_Pendencia = value
        End Set
    End Property
    Private m_Pendencia As String


    Public Property Analise() As String
        Get
            Return m_Analise
        End Get
        Set(ByVal value As String)
            m_Analise = value
        End Set
    End Property
    Private m_Analise As String


    Public Property StatusNFTaxaAdesao() As String
        Get
            Return m_StatusNFTaxaAdesao
        End Get
        Set(ByVal value As String)
            m_StatusNFTaxaAdesao = value
        End Set
    End Property
    Private m_StatusNFTaxaAdesao As String


    Public Property StatusNFInstalacao() As String
        Get
            Return m_StatusNFInstalacao
        End Get
        Set(ByVal value As String)
            m_StatusNFInstalacao = value
        End Set
    End Property
    Private m_StatusNFInstalacao As String


    Public Property StatusNFEmpComodato() As String
        Get
            Return m_StatusNFEmpComodato
        End Get
        Set(ByVal value As String)
            m_StatusNFEmpComodato = value
        End Set
    End Property
    Private m_StatusNFEmpComodato As String


    Public Property HouveEntradaNF() As String
        Get
            Return m_HouveEntradaNF
        End Get
        Set(ByVal value As String)
            m_HouveEntradaNF = value
        End Set
    End Property
    Private m_HouveEntradaNF As String

    Public Property PedNovo As String
        Get
            Return m_PedNovo
        End Get
        Set(ByVal value As String)
            m_PedNovo = value
        End Set
    End Property
    Private m_PedNovo As String


    ''' <summary>
    ''' Variável utilizada na busca de pedido de venda por id de indicação, execução de indicação
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IdIndicacao As String
        Get
            Return m_IdIndicacao
        End Get
        Set(ByVal value As String)
            m_IdIndicacao = value
        End Set
    End Property
    Private m_IdIndicacao As String

    Public Property IdFilial As String
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As String)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As String


    Public Property StatusPed() As String
        Get
            Return m_StatusPed
        End Get
        Set(ByVal value As String)
            m_StatusPed = value
        End Set
    End Property
    Private m_StatusPed As String


    Public Property VarCalcDifPgtoInst() As Double
        Get
            Return m_VarCalcDifPgtoInst
        End Get
        Set(ByVal value As Double)
            m_VarCalcDifPgtoInst = value
        End Set
    End Property
    Private m_VarCalcDifPgtoInst As Double

    Public Property ServExecRepr() As String
        Get
            Return m_ServExecRepr
        End Get
        Set(ByVal value As String)
            m_ServExecRepr = value
        End Set
    End Property
    Private m_ServExecRepr As String



    Public Property CatPed() As String
        Get
            Return m_CatPed
        End Get
        Set(ByVal value As String)
            m_CatPed = value
        End Set
    End Property
    Private m_CatPed As String

    Public Property CodTabComiss() As String
        Get
            Return m_CodTabComiss
        End Get
        Set(ByVal value As String)
            m_CodTabComiss = value
        End Set
    End Property
    Private m_CodTabComiss As String

    Public Property TipoPgtoComiss() As String
        Get
            Return m_TipoPgtoComiss
        End Get
        Set(ByVal value As String)
            m_TipoPgtoComiss = value
        End Set
    End Property
    Private m_TipoPgtoComiss As String

    Public Property StatusNFDoacao() As String
        Get
            Return m_StatusNFDoacao
        End Get
        Set(ByVal value As String)
            m_StatusNFDoacao = value
        End Set
    End Property
    Private m_StatusNFDoacao As String

    Public Property StatusNFMonitoria() As String
        Get
            Return m_StatusNFMonitoria
        End Get
        Set(ByVal value As String)
            m_StatusNFMonitoria = value
        End Set
    End Property
    Private m_StatusNFMonitoria As String



    'ATRIBUTOS NOVOS PARA O RANKING DE VENDAS GERENCIAL
    Private m_PedCad As Integer
    Public Property PedCad() As Integer
        Get
            Return m_PedCad
        End Get
        Set(ByVal value As Integer)
            m_PedCad = value
        End Set
    End Property

    Private m_PedFinal As Integer
    Public Property PedFinal() As Integer
        Get
            Return m_PedFinal
        End Get
        Set(ByVal value As Integer)
            m_PedFinal = value
        End Set
    End Property


    Private m_Indicacoes As Integer
    Public Property Indicacoes() As Integer
        Get
            Return m_Indicacoes
        End Get
        Set(ByVal value As Integer)
            m_Indicacoes = value
        End Set
    End Property

    Private m_PedCadTeleview As Integer
    Public Property PedCadTeleview() As Integer
        Get
            Return m_PedCadTeleview
        End Get
        Set(ByVal value As Integer)
            m_PedCadTeleview = value
        End Set
    End Property

    Private m_PedFinalTeleviw As Integer
    Public Property PedFinalTeleview() As Integer
        Get
            Return m_PedFinalTeleviw
        End Get
        Set(ByVal value As Integer)
            m_PedFinalTeleviw = value
        End Set
    End Property

    Private m_IndicacoesTeleview As Integer
    Public Property IndicacoesTeleview() As Integer
        Get
            Return m_IndicacoesTeleview
        End Get
        Set(ByVal value As Integer)
            m_IndicacoesTeleview = value
        End Set
    End Property


    Private m_strPedCadTeleView As String
    Public Property strPedCadTeleView() As String
        Get
            Return m_strPedCadTeleView
        End Get
        Set(ByVal value As String)
            m_strPedCadTeleView = value
        End Set
    End Property

    Private m_strPedFinalTeleview As String
    Public Property strPedFinalTeleview() As String
        Get
            Return m_strPedFinalTeleview
        End Get
        Set(ByVal value As String)
            m_strPedFinalTeleview = value
        End Set
    End Property

    Private m_strIndicacoesTeleview As String
    Public Property strIndicacoesTeleview() As String
        Get
            Return m_strIndicacoesTeleview
        End Get
        Set(ByVal value As String)
            m_strIndicacoesTeleview = value
        End Set
    End Property

    Private m_MesAno As String
    Public Property MesAno As String
        Get
            Return m_MesAno
        End Get
        Set(ByVal value As String)
            m_MesAno = value
        End Set
    End Property

    Private m_MediaTSM As Double
    Public Property MediaTSM() As Double
        Get
            Return m_MediaTSM
        End Get
        Set(ByVal value As Double)
            m_MediaTSM = value
        End Set
    End Property

    Private m_MediaTSMStr As String
    Public Property MediaTSMStr() As String
        Get
            Return m_MediaTSMStr
        End Get
        Set(ByVal value As String)
            m_MediaTSMStr = value
        End Set
    End Property

    Public Property StatusNFCFTV() As String
        Get
            Return m_StatusNFCFTV
        End Get
        Set(ByVal value As String)
            m_StatusNFCFTV = value
        End Set
    End Property
    Private m_StatusNFCFTV As String

    Public Property StatusNFInstalacaoCFTV() As String
        Get
            Return m_StatusNFInstalacaoCFTV
        End Get
        Set(ByVal value As String)
            m_StatusNFInstalacaoCFTV = value
        End Set
    End Property
    Private m_StatusNFInstalacaoCFTV As String

    'atributos para aprovação financeira
    Public Property RazaoSocial() As String
        Get
            Return m_RazaoSocial
        End Get
        Set(ByVal value As String)
            m_RazaoSocial = value
        End Set
    End Property
    Private m_RazaoSocial As String

    Public Property Estabelecimento() As String
        Get
            Return m_Estabelecimento
        End Get
        Set(ByVal value As String)
            m_Estabelecimento = value
        End Set
    End Property
    Private m_Estabelecimento As String

    Public Property DtPedido() As Nullable(Of DateTime)
        Get
            Return m_DtPedido
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtPedido = value
        End Set
    End Property
    Private m_DtPedido As Nullable(Of DateTime)

    Public Property DtFinalInstal() As Nullable(Of DateTime)
        Get
            Return m_DtFinalInstal
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtFinalInstal = value
        End Set
    End Property
    Private m_DtFinalInstal As Nullable(Of DateTime)

    Public Property Status() As String
        Get
            Return m_Status
        End Get
        Set(ByVal value As String)
            m_Status = value
        End Set
    End Property
    Private m_Status As String

    Public Property TipoPed() As String
        Get
            Return m_TipoPed
        End Get
        Set(ByVal value As String)
            m_TipoPed = value
        End Set
    End Property
    Private m_TipoPed As String

    Public Property VdaExpress() As String
        Get
            Return m_VdaExpress
        End Get
        Set(value As String)
            m_VdaExpress = value
        End Set
    End Property
    Private m_VdaExpress As String

    Public Property VlrTotalPedido As Double
    Public Property VlrTotInst As Double
    Public Property VlrTotVenda As Double
    Public Property vlrMaoObraVerisurePRO As Double
    Public Property VlrTotAdes As Double
    Public Property EmailRepr As String
    Public Property CodIntClie As String
    Public Property ClienteDe As String
    Public Property MonitoradoPor As String
    Public Property Vendas As Integer
    Public Property strVendas As String
    Public Property IngressoMedio As Integer
    Public Property strIngressoMedio As String
    Public Property NumPedVdaFisico As String
    Public Property CondicaoPgto As String
    Public Property Situacao As String
    Public Property Demo As Integer
    Public Property MatVend As String
    Public Property CodAreaVenda As String
    Public Property DtAprovacao As String
    Public Property TpPedido As String
    Public Property MotivoRepr As String
    Public Property codRepr As String
    Public Property usrGerente As String
    Public Property usrChefeLider As String
    Public Property usrDiretor As String
    Public Property isVerisurePRO As String
End Class
