Imports Teleatlantic.TLS.Common

Public Class CondicaoPagamentoVenda : Inherits Retorno
    ' Entidade de Codição de Pagamento de Vendas

    Public Property Codigo() As String
        Get
            Return m_codigo
        End Get
        Set(ByVal value As String)
            m_codigo = value
        End Set
    End Property
    Private m_codigo As String

    Public Property Descricao() As String
        Get
            Return m_Descricao
        End Get
        Set(ByVal valeu As String)
            m_Descricao = valeu
        End Set
    End Property
    Private m_Descricao As String

    Public Property UsrAlt() As String
        Get
            Return m_UsrAlt
        End Get
        Set(ByVal value As String)
            m_UsrAlt = value
        End Set
    End Property
    Private m_UsrAlt As String

    Public Property TipoVecto() As String
        Get
            Return m_TipoVecto
        End Get
        Set(ByVal value As String)
            m_TipoVecto = value
        End Set
    End Property
    Private m_TipoVecto As String

    Public Property Status() As String
        Get
            Return m_Status
        End Get
        Set(ByVal valeu As String)
            m_Status = valeu
        End Set
    End Property
    Private m_Status As String

    Public Property QtdeParcelas() As String
        Get
            Return m_QtdeParcelas
        End Get
        Set(ByVal value As String)
            m_QtdeParcelas = value
        End Set
    End Property
    Private m_QtdeParcelas As String

    Public Property Juros() As String
        Get
            Return m_Juros
        End Get
        Set(ByVal value As String)
            m_Juros = value
        End Set
    End Property
    Private m_Juros As String

    Public Property UtilizandoSite() As String
        Get
            Return m_UtilizandoSite
        End Get
        Set(ByVal value As String)
            m_UtilizandoSite = value
        End Set
    End Property
    Private m_UtilizandoSite As String

    Public Property DescricaoSite() As String
        Get
            Return m_DescricaoSite
        End Get
        Set(ByVal value As String)
            m_DescricaoSite = value
        End Set
    End Property
    Private Property m_DescricaoSite As String

    Public Property QtDias As Integer
        Get
            Return m_QtDias
        End Get
        Set(ByVal value As Integer)
            m_QtDias = value
        End Set
    End Property
    Private m_QtDias As Integer

    Public Property Perc As Double
        Get
            Return m_Perc
        End Get
        Set(ByVal value As Double)
            m_Perc = value
        End Set
    End Property
    Private m_Perc As Double

    Private _TipoPgto As String
    Public Property TipoPgto() As String
        Get
            Return _TipoPgto
        End Get
        Set(ByVal value As String)
            _TipoPgto = value
        End Set
    End Property

    Public Property IsGerarParcelaUnica As Boolean
    Public Property NumParc As Integer
    Public Property CodNavision As String
    Public Property verisurePRO As Integer


End Class
