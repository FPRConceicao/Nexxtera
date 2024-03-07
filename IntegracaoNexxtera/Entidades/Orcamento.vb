Imports Teleatlantic.TLS.Common

Public Class Orcamento : Inherits Retorno

    Public Property CodOrc() As String
        Get
            Return m_CodOrc
        End Get
        Set(ByVal value As String)
            m_CodOrc = value
        End Set
    End Property
    Private m_CodOrc As String


    Public Property DtOrcam() As Nullable(Of DateTime)
        Get
            Return m_DtOrcam
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtOrcam = value
        End Set
    End Property
    Private m_DtOrcam As Nullable(Of DateTime)


    Public Property OrigemOrc() As String
        Get
            Return m_OrigemOrc
        End Get
        Set(ByVal value As String)
            m_OrigemOrc = value
        End Set
    End Property
    Private m_OrigemOrc As String


    Public Property QtdePontos() As String
        Get
            Return m_QtdePontos
        End Get
        Set(ByVal value As String)
            m_QtdePontos = value
        End Set
    End Property
    Private m_QtdePontos As String


    Public Property Protocolo() As String
        Get
            Return m_Protocolo
        End Get
        Set(ByVal value As String)
            m_Protocolo = value
        End Set
    End Property
    Private m_Protocolo As String


    Public Property NomeOrc() As String
        Get
            Return m_NomeOrc
        End Get
        Set(ByVal value As String)
            m_NomeOrc = value
        End Set
    End Property
    Private m_NomeOrc As String


    Public Property CodClie() As String
        Get
            Return m_CodClie
        End Get
        Set(ByVal value As String)
            m_CodClie = value
        End Set
    End Property
    Private m_CodClie As String

    Public Property StatusOrc() As String
        Get
            Return m_StatusOrc
        End Get
        Set(ByVal value As String)
            m_StatusOrc = value
        End Set
    End Property
    Private m_StatusOrc As String

    Public Property Representante() As Representante
        Get
            Return m_Representante
        End Get
        Set(ByVal value As Representante)
            m_Representante = value
        End Set
    End Property
    Private m_Representante As Representante


    Public Property VlrMonit() As Double
        Get
            Return m_VlrMonit
        End Get
        Set(ByVal value As Double)
            m_VlrMonit = value
        End Set
    End Property
    Private m_VlrMonit As Double

    Public Property VlrSuperMotor() As Double
        Get
            Return m_VlrSuperMotor
        End Get
        Set(ByVal value As Double)
            m_VlrSuperMotor = value
        End Set
    End Property
    Private m_VlrSuperMotor As Double

    Public Property VlrSeguro() As Double
        Get
            Return m_VlrSeguro
        End Get
        Set(ByVal value As Double)
            m_VlrSeguro = value
        End Set
    End Property
    Private m_VlrSeguro As Double

    Public Property VlrRadio() As Double
        Get
            Return m_VlrRadio
        End Get
        Set(ByVal value As Double)
            m_VlrRadio = value
        End Set
    End Property
    Private m_VlrRadio As Double

    Public Property TeleEmerg() As Double
        Get
            Return m_TeleEmerg
        End Get
        Set(ByVal value As Double)
            m_TeleEmerg = value
        End Set
    End Property
    Private m_TeleEmerg As Double

    Public Property VlrLocacao() As Double
        Get
            Return m_VlrLocacao
        End Get
        Set(ByVal value As Double)
            m_VlrLocacao = value
        End Set
    End Property
    Private m_VlrLocacao As Double

    Public Property VlrRelat() As Double
        Get
            Return m_VlrRelat
        End Get
        Set(ByVal value As Double)
            m_VlrRelat = value
        End Set
    End Property
    Private m_VlrRelat As Double

    Public Property VlrRonda() As Double
        Get
            Return m_VlrRonda
        End Get
        Set(ByVal value As Double)
            m_VlrRonda = value
        End Set
    End Property
    Private m_VlrRonda As Double

    Public Property VlrAviso() As Double
        Get
            Return m_VlrAviso
        End Get
        Set(ByVal value As Double)
            m_VlrAviso = value
        End Set
    End Property
    Private m_VlrAviso As Double

    Public Property VlrArme() As Double
        Get
            Return m_VlrArme
        End Get
        Set(ByVal value As Double)
            m_VlrArme = value
        End Set
    End Property
    Private m_VlrArme As Double

    Public Property VlrTxManMon() As Double
        Get
            Return m_VlrTxManMon
        End Get
        Set(ByVal value As Double)
            m_VlrTxManMon = value
        End Set
    End Property
    Private m_VlrTxManMon As Double

    Public Property VlrCercaElet() As Double
        Get
            Return m_VlrCercaElet
        End Get
        Set(ByVal value As Double)
            m_VlrCercaElet = value
        End Set
    End Property
    Private m_VlrCercaElet As Double

    Public Property VlrGPRS() As Double
        Get
            Return m_VlrGPRS
        End Get
        Set(ByVal value As Double)
            m_VlrGPRS = value
        End Set
    End Property
    Private m_VlrGPRS As Double

    Public Property Manutencao() As Double
        Get
            Return m_Manutencao
        End Get
        Set(ByVal value As Double)
            m_Manutencao = value
        End Set
    End Property
    Private m_Manutencao As Double

    Public Property VlrMonitAdic() As Double
        Get
            Return m_VlrMonitAdic
        End Get
        Set(ByVal value As Double)
            m_VlrMonitAdic = value
        End Set
    End Property
    Private m_VlrMonitAdic As Double

    Public Property VlrTeleVision() As Double
        Get
            Return m_VlrTeleVision
        End Get
        Set(ByVal value As Double)
            m_VlrTeleVision = value
        End Set
    End Property
    Private m_VlrTeleVision As Double

    Public Property TipoTeleVideo() As String
        Get
            Return m_TipoTeleVideo
        End Get
        Set(ByVal value As String)
            m_TipoTeleVideo = value
        End Set
    End Property
    Private m_TipoTeleVideo As String

    Public Property VlrTarBanc() As Double
        Get
            Return m_VlrTarBanc
        End Get
        Set(ByVal value As Double)
            m_VlrTarBanc = value
        End Set
    End Property
    Private m_VlrTarBanc As Double

    Public Property Total() As Double
        Get
            Return m_Total
        End Get
        Set(ByVal value As Double)
            m_Total = value
        End Set
    End Property
    Private m_Total As Double

    Public Property TipoPagto() As String
        Get
            Return m_TipoPagto
        End Get
        Set(ByVal value As String)
            m_TipoPagto = value
        End Set
    End Property
    Private m_TipoPagto As String

    Public Property TipoCartaoCred() As String
        Get
            Return m_TipoCartaoCred
        End Get
        Set(ByVal value As String)
            m_TipoCartaoCred = value
        End Set
    End Property
    Private m_TipoCartaoCred As String

    Public Property CodBancoDeb() As String
        Get
            Return m_CodBancoDeb
        End Get
        Set(ByVal value As String)
            m_CodBancoDeb = value
        End Set
    End Property
    Private m_CodBancoDeb As String

    Public Property GPRS() As String
        Get
            Return m_GPRS
        End Get
        Set(ByVal value As String)
            m_GPRS = value
        End Set
    End Property
    Private m_GPRS As String


    Public Property DescProdInst() As Double
        Get
            Return m_DescProdInst
        End Get
        Set(ByVal value As Double)
            m_DescProdInst = value
        End Set
    End Property
    Private m_DescProdInst As Double


    Public Property DescProd() As Double
        Get
            Return m_DescProd
        End Get
        Set(ByVal value As Double)
            m_DescProd = value
        End Set
    End Property
    Private m_DescProd As Double


    Public Property DescTotal() As Double
        Get
            Return m_DescTotal
        End Get
        Set(ByVal value As Double)
            m_DescTotal = value
        End Set
    End Property
    Private m_DescTotal As Double


    Public Property DescMObra() As Double
        Get
            Return m_DescMObra
        End Get
        Set(ByVal value As Double)
            m_DescMObra = value
        End Set
    End Property
    Private m_DescMObra As Double


    Public Property Cliente As Cliente
        Get
            Return m_Cliente
        End Get
        Set(ByVal value As Cliente)
            m_Cliente = value
        End Set
    End Property
    Private m_Cliente As New Cliente


    Public Property PedidoVenda As PedidoVenda
        Get
            Return m_PedidoVenda
        End Get
        Set(ByVal value As PedidoVenda)
            m_PedidoVenda = value
        End Set
    End Property
    Private m_PedidoVenda As New PedidoVenda


    Public Property Orcamento As String
        Get
            Return m_Orcamento
        End Get
        Set(ByVal value As String)
            m_Orcamento = value
        End Set
    End Property
    Private m_Orcamento As String


    Public Property DtAlt As Nullable(Of DateTime)
        Get
            Return m_DtAlt
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtAlt = value
        End Set
    End Property
    Private m_DtAlt As Nullable(Of DateTime)


    Public Property DtEnvEmail As Nullable(Of DateTime)
        Get
            Return m_DtEnvEmail
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtEnvEmail = value
        End Set
    End Property
    Private m_DtEnvEmail As Nullable(Of DateTime)


    Public Property IsDescTotal As String
        Get
            Return m_IsDescTotal
        End Get
        Set(ByVal value As String)
            m_IsDescTotal = value
        End Set
    End Property
    Private m_IsDescTotal As String


    Public Property CodFVisita As String
        Get
            Return m_CodFVisita
        End Get
        Set(ByVal value As String)
            m_CodFVisita = value
        End Set
    End Property
    Private m_CodFVisita As String


    Public Property TipoOrc As String
        Get
            Return m_TipoOrc
        End Get
        Set(ByVal value As String)
            m_TipoOrc = value
        End Set
    End Property
    Private m_TipoOrc As String

    Private m_CodCpgt As String
    Public Property CodCpgt() As String
        Get
            Return m_CodCpgt
        End Get
        Set(ByVal value As String)
            m_CodCpgt = value
        End Set
    End Property


    Public Property Obs As String
        Get
            Return m_Obs
        End Get
        Set(ByVal value As String)
            m_Obs = value
        End Set
    End Property
    Private m_Obs As String


    Public Property Contato As String
        Get
            Return m_Contato
        End Get
        Set(ByVal value As String)
            m_Contato = value
        End Set
    End Property
    Private m_Contato As String


    Public Property Dealer As String
        Get
            Return m_Dealer
        End Get
        Set(ByVal value As String)
            m_Dealer = value
        End Set
    End Property
    Private m_Dealer As String


    Public Property TxAdesao As Double
        Get
            Return m_TxAdesao
        End Get
        Set(ByVal value As Double)
            m_TxAdesao = value
        End Set
    End Property
    Private m_TxAdesao As Double


    Public Property FmPgtEntr As Double
        Get
            Return m_FmPgtEntr
        End Get
        Set(ByVal value As Double)
            m_FmPgtEntr = value
        End Set
    End Property
    Private m_FmPgtEntr As Double


    Public Property FmPgtQtde As Double
        Get
            Return m_FmPgtQtde
        End Get
        Set(ByVal value As Double)
            m_FmPgtQtde = value
        End Set
    End Property
    Private m_FmPgtQtde As Double


    Public Property CtEqtKit As Double
        Get
            Return m_CtEqtKit
        End Get
        Set(ByVal value As Double)
            m_CtEqtKit = value
        End Set
    End Property
    Private m_CtEqtKit As Double


    Public Property CtInstKit As Double
        Get
            Return m_CtInstKit
        End Get
        Set(ByVal value As Double)
            m_CtInstKit = value
        End Set
    End Property
    Private m_CtInstKit As Double


    Public Property CtInstExtra As Double
        Get
            Return m_CtInstExtra
        End Get
        Set(ByVal value As Double)
            m_CtInstExtra = value
        End Set
    End Property
    Private m_CtInstExtra As Double


    Public Property CatPed As String
        Get
            Return m_CatPed
        End Get
        Set(ByVal value As String)
            m_CatPed = value
        End Set
    End Property
    Private m_CatPed As String


    Public Property ComICMS As String
        Get
            Return m_ComICMS
        End Get
        Set(ByVal value As String)
            m_ComICMS = value
        End Set
    End Property
    Private m_ComICMS As String


    Public Property CustoTele As String
        Get
            Return m_CustoTele
        End Get
        Set(ByVal value As String)
            m_CustoTele = value
        End Set
    End Property
    Private m_CustoTele As String


    Public Property ValorIcms As Double
        Get
            Return m_ValorIcms
        End Get
        Set(ByVal value As Double)
            m_ValorIcms = value
        End Set
    End Property
    Private m_ValorIcms As Double


    Public Property ECAdesao As Double
        Get
            Return m_ECAdesao
        End Get
        Set(ByVal value As Double)
            m_ECAdesao = value
        End Set
    End Property
    Private m_ECAdesao As Double


    Public Property ValPGIns As Double
        Get
            Return m_ValPGIns
        End Get
        Set(ByVal value As Double)
            m_ValPGIns = value
        End Set
    End Property
    Private m_ValPGIns As Double


    Public Property DescCVen As Double
        Get
            Return m_DescCVen
        End Get
        Set(ByVal value As Double)
            m_DescCVen = value
        End Set
    End Property
    Private m_DescCVen As Double


    Public Property IntaEqui As Double
        Get
            Return m_IntaEqui
        End Get
        Set(ByVal value As Double)
            m_IntaEqui = value
        End Set
    End Property
    Private m_IntaEqui As Double


    Public Property TipoComMon As String
        Get
            Return m_TipoComMon
        End Get
        Set(ByVal value As String)
            m_TipoComMon = value
        End Set
    End Property
    Private m_TipoComMon As String


    Public Property VlrInstal As Double
        Get
            Return m_VlrInstal
        End Get
        Set(ByVal value As Double)
            m_VlrInstal = value
        End Set
    End Property
    Private m_VlrInstal As Double


    Public Property VlrTxVisita As Double
        Get
            Return m_VlrTxVisita
        End Get
        Set(ByVal value As Double)
            m_VlrTxVisita = value
        End Set
    End Property
    Private m_VlrTxVisita As Double


    Public Property Indicador As String
        Get
            Return m_Indicador
        End Get
        Set(ByVal value As String)
            m_Indicador = value
        End Set
    End Property
    Private m_Indicador As String


    Public Property Botom As String
        Get
            Return m_Botom
        End Get
        Set(ByVal value As String)
            m_Botom = value
        End Set
    End Property
    Private m_Botom As String


    Public Property Radio As String
        Get
            Return m_Radio
        End Get
        Set(ByVal value As String)
            m_Radio = value
        End Set
    End Property
    Private m_Radio As String


    Public Property UsrIns As String
        Get
            Return m_UsrIns
        End Get
        Set(ByVal value As String)
            m_UsrIns = value
        End Set
    End Property
    Private m_UsrIns As String


    Public Property VlrBkpGPRS As Double
        Get
            Return m_VlrBkpGPRS
        End Get
        Set(ByVal value As Double)
            m_VlrBkpGPRS = value
        End Set
    End Property
    Private m_VlrBkpGPRS As Double


    Public Property VlrManut As Double
        Get
            Return m_VlrManut
        End Get
        Set(ByVal value As Double)
            m_VlrManut = value
        End Set
    End Property
    Private m_VlrManut As Double


    Public Property TipoManut As String
        Get
            Return m_TipoManut
        End Get
        Set(ByVal value As String)
            m_TipoManut = value
        End Set
    End Property
    Private m_TipoManut As String


    Public Property CodDeptOrigem As String
        Get
            Return m_CodDeptOrigem
        End Get
        Set(ByVal value As String)
            m_CodDeptOrigem = value
        End Set
    End Property
    Private m_CodDeptOrigem As String


    Public Property UsrAlt As String
        Get
            Return m_UsrAlt
        End Get
        Set(ByVal value As String)
            m_UsrAlt = value
        End Set
    End Property
    Private m_UsrAlt As String


    Public Property VlrInstalPgClie As Double
        Get
            Return m_VlrInstalPgClie
        End Get
        Set(ByVal value As Double)
            m_VlrInstalPgClie = value
        End Set
    End Property
    Private m_VlrInstalPgClie As Double


    Public Property VlrInstalAumPtoPgClie As Double
        Get
            Return m_VlrInstalAumPtoPgClie
        End Get
        Set(ByVal value As Double)
            m_VlrInstalAumPtoPgClie = value
        End Set
    End Property
    Private m_VlrInstalAumPtoPgClie As Double


    Public Property TipoEmerg As String
        Get
            Return m_TipoEmerg
        End Get
        Set(ByVal value As String)
            m_TipoEmerg = value
        End Set
    End Property
    Private m_TipoEmerg As String


    Public Property VlrInstComodato As Double
        Get
            Return m_VlrInstComodato
        End Get
        Set(ByVal value As Double)
            m_VlrInstComodato = value
        End Set
    End Property
    Private m_VlrInstComodato As Double


    Public Property DtConferencia As Nullable(Of DateTime)
        Get
            Return m_DtConferencia
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtConferencia = value
        End Set
    End Property
    Private m_DtConferencia As Nullable(Of DateTime)


    Public Property TipoService As String
        Get
            Return m_TipoService
        End Get
        Set(ByVal value As String)
            m_TipoService = value
        End Set
    End Property
    Private m_TipoService As String


    Public Property QtdeCamTV As String
        Get
            Return m_QtdeCamTV
        End Get
        Set(ByVal value As String)
            m_QtdeCamTV = value
        End Set
    End Property
    Private m_QtdeCamTV As String


    Public Property SolDesc As String
        Get
            Return m_SolDesc
        End Get
        Set(ByVal value As String)
            m_SolDesc = value
        End Set
    End Property
    Private m_SolDesc As String


    Public Property VlrPagoInstaladorComodato As Double
        Get
            Return m_VlrPagoInstaladorComodato
        End Get
        Set(ByVal value As Double)
            m_VlrPagoInstaladorComodato = value
        End Set
    End Property
    Private m_VlrPagoInstaladorComodato As Double


    Public Property VlrManutencaoCFTV As Double
        Get
            Return m_VlrManutencaoCFTV
        End Get
        Set(ByVal value As Double)
            m_VlrManutencaoCFTV = value
        End Set
    End Property
    Private m_VlrManutencaoCFTV As Double


    Public Property TipoManutencaoCFTV As String
        Get
            Return m_TipoManutencaoCFTV
        End Get
        Set(ByVal value As String)
            m_TipoManutencaoCFTV = value
        End Set
    End Property
    Private m_TipoManutencaoCFTV As String


    Public Property QtdeCamManutencaoCFTV As String
        Get
            Return m_QtdeCamManutencaoCFTV
        End Get
        Set(ByVal value As String)
            m_QtdeCamManutencaoCFTV = value
        End Set
    End Property
    Private m_QtdeCamManutencaoCFTV As String


    Public Property CodIntClie As String
        Get
            Return m_CodIntClie
        End Get
        Set(ByVal value As String)
            m_CodIntClie = value
        End Set
    End Property
    Private m_CodIntClie As String


    Public Property Operadora As String
        Get
            Return m_Operadora
        End Get
        Set(ByVal value As String)
            m_Operadora = value
        End Set
    End Property
    Private m_Operadora As String


    Public Property NumPedVdaTeleVideo As String
        Get
            Return m_NumPedVdaTeleVideo
        End Get
        Set(ByVal value As String)
            m_NumPedVdaTeleVideo = value
        End Set
    End Property
    Private m_NumPedVdaTeleVideo As String


    Public Property RepasseTE As Double
        Get
            Return m_RepasseTE
        End Get
        Set(ByVal value As Double)
            m_RepasseTE = value
        End Set
    End Property
    Private m_RepasseTE As Double

    Private _VlrTeleGPS As Double
    Public Property VlrTeleGPS() As Double
        Get
            Return _VlrTeleGPS
        End Get
        Set(ByVal value As Double)
            _VlrTeleGPS = value
        End Set
    End Property

    Public Property VlrInstalRepasse() As Double
        Get
            Return m_VlrInstalRepasse
        End Get
        Set(ByVal value As Double)
            m_VlrInstalRepasse = value
        End Set
    End Property

    Public Property VlrTotalServico() As Double
        Get
            Return m_VlrTotalServico
        End Get
        Set(ByVal value As Double)
            m_VlrTotalServico = value
        End Set
    End Property
    Private m_VlrTotalServico As Double

    Public Property DtFinalizacaoOrcamento() As Nullable(Of DateTime)
        Get
            Return m_DtFinOrc
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtFinOrc = value
        End Set
    End Property
    Private m_DtFinOrc As Nullable(Of DateTime)

    Public Property UsuarioAprov As String
        Get
            Return m_UsuarioAprov
        End Get
        Set(ByVal value As String)
            m_UsuarioAprov = value
        End Set
    End Property
    Private m_UsuarioAprov As String

    Public Property TipoClassificacao() As TipoClassificacao
        Get
            Return m_TipoClassificacao
        End Get
        Set(ByVal value As TipoClassificacao)
            m_TipoClassificacao = value
        End Set
    End Property
    Private m_TipoClassificacao As New TipoClassificacao

    Private m_VlrInstalRepasse As Double
    Public Property CodReclamacaoDesconto As String
    Public Property QtdeParticoes As Integer
    Public Property ValidCCred As String
    Public Property NomePortadorCC As String
    Public Property CodSegurancaCC As String
    Public Property CartaoCred As String
    Public Property CodAgenDeb As String
    Public Property NumCtaDeb As String
    Public Property Prazo As Integer
    Public Property UsrPendencia As String
    Public Property NumCta As String
    Public Property CodAgen As String
    Public Property CodBanco As String
    Public Property CodEmpManut As String
    Public Property EmpManut As String
    Public Property CodTecnico As String
    Public Property NomeTecnico As String
    Public Property IdTipoFechamento As Integer?
    Public Property IdAprovador As Integer?
    Public Property vlrTotalOrcamento As String
    Public Property CodOperacaoConta As String
    Public Property vlrMaoObraAlarme As Double
    Public Property vlrMaoObraMelhoria As Double

    Public Property UsrAltOrc As String
        Get
            Return m_UsrAltOrc
        End Get
        Set(ByVal value As String)
            m_UsrAltOrc = value
        End Set
    End Property
    Private m_UsrAltOrc As String
End Class
