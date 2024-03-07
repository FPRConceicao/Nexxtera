Imports Teleatlantic.TLS.Common

Public Class ClienteFinanceiro : Inherits Retorno


    Public Property LibCobrMonit() As String
        Get
            Return m_LibCobrMonit
        End Get
        Set(ByVal value As String)
            m_LibCobrMonit = value
        End Set
    End Property
    Private m_LibCobrMonit As String


    Public Property CheqSemFdo() As String
        Get
            Return m_CheqSemFdo
        End Get
        Set(ByVal value As String)
            m_CheqSemFdo = value
        End Set
    End Property
    Private m_CheqSemFdo As String


    Public Property MesAnoEnvBco() As String
        Get
            Return m_MesAnoEnvBco
        End Get
        Set(ByVal value As String)
            m_MesAnoEnvBco = value
        End Set
    End Property
    Private m_MesAnoEnvBco As String


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

    Public Property ObsNotaFiscal() As String
        Get
            Return m_ObsNotaFiscal
        End Get
        Set(ByVal value As String)
            m_ObsNotaFiscal = value
        End Set
    End Property
    Private m_ObsNotaFiscal As String

    Public Property TipoPagto() As String
        Get
            Return m_TipoPagto
        End Get
        Set(ByVal value As String)
            m_TipoPagto = value
        End Set
    End Property
    Private m_TipoPagto As String

    Public Property CartaoCred() As String
        Get
            Return m_CartaoCred
        End Get
        Set(ByVal value As String)
            m_CartaoCred = value
        End Set
    End Property
    Private m_CartaoCred As String

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

    Public Property TipoCartaoCred() As String
        Get
            Return m_TipoCartaoCred
        End Get
        Set(ByVal value As String)
            m_TipoCartaoCred = value
        End Set
    End Property
    Private m_TipoCartaoCred As String

    Public Property ValidCCred() As String
        Get
            Return m_ValidCCred
        End Get
        Set(ByVal value As String)
            m_ValidCCred = value
        End Set
    End Property
    Private m_ValidCCred As String

    Public Property IsInadimp() As Integer
        Get
            Return m_IsInadimp
        End Get
        Set(ByVal value As Integer)
            m_IsInadimp = value
        End Set
    End Property
    Private m_IsInadimp As Integer

    Public Property DtRegInadimp() As DateTime
        Get
            Return m_DtRegInadimp
        End Get
        Set(ByVal value As DateTime)
            m_DtRegInadimp = value
        End Set
    End Property
    Private m_DtRegInadimp As DateTime

    Public Property AtendCobr() As String
        Get
            Return m_AtendCobr
        End Get
        Set(ByVal value As String)
            m_AtendCobr = value
        End Set
    End Property
    Private m_AtendCobr As String

    Public Property ObsNF() As String
        Get
            Return m_ObsNF
        End Get
        Set(ByVal value As String)
            m_ObsNF = value
        End Set
    End Property
    Private m_ObsNF As String
    Public Property CodSegurancaCC As String
    Public Property NomePortadorCC As String
    Public Property CodOperacaoConta As String
    Public Property OptDebAutorizado As Integer
    Public Property shopperReference As String
    Public Property TipoPagtoDesc() As String
        Get
            Return m_TipoPagtoDesc
        End Get
        Set(ByVal value As String)
            m_TipoPagtoDesc = value
        End Set
    End Property
    Private m_TipoPagtoDesc As String
    Public Property pspReference As String
    Public Property ObsBoletoUnificado As String
End Class
