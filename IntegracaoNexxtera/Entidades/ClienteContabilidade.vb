Imports Teleatlantic.TLS.Common

Public Class ClienteContabilidade : Inherits Retorno


    Public Property EnvioBanco() As String
        Get
            Return m_EnvioBanco
        End Get
        Set(ByVal value As String)
            m_EnvioBanco = value
        End Set
    End Property
    Private m_EnvioBanco As String

    Public Property PercRetIssMonitoria() As Double
        Get
            Return m_PercRetIssMonitoria
        End Get
        Set(ByVal value As Double)
            m_PercRetIssMonitoria = value
        End Set
    End Property
    Private m_PercRetIssMonitoria As Double

    Public Property PercRetIssManutencao() As Double
        Get
            Return m_PercRetIssManutencao
        End Get
        Set(ByVal value As Double)
            m_PercRetIssManutencao = value
        End Set
    End Property
    Private m_PercRetIssManutencao As Double

    Public Property PercRetIssInstalacao() As Double
        Get
            Return m_PercRetIssInstalacao
        End Get
        Set(ByVal value As Double)
            m_PercRetIssInstalacao = value
        End Set
    End Property
    Private m_PercRetIssInstalacao As Double

    Public Property IsRetCsll() As Integer
        Get
            Return m_IsRetCsll
        End Get
        Set(ByVal value As Integer)
            m_IsRetCsll = value
        End Set
    End Property
    Private m_IsRetCsll As Integer

    Public Property IsRetCofins() As Integer
        Get
            Return m_IsRetCofins
        End Get
        Set(ByVal value As Integer)
            m_IsRetCofins = value
        End Set
    End Property
    Private m_IsRetCofins As Integer

    Public Property IsRetPis() As Integer
        Get
            Return m_IsRetPis
        End Get
        Set(ByVal value As Integer)
            m_IsRetPis = value
        End Set
    End Property
    Private m_IsRetPis As Integer

    Public Property IsRetIR() As Integer
        Get
            Return m_IsRetIR
        End Get
        Set(ByVal value As Integer)
            m_IsRetIR = value
        End Set
    End Property
    Private m_IsRetIR As Integer

    Public Property IsRetINSS() As Integer
        Get
            Return m_IsRetINSS
        End Get
        Set(ByVal value As Integer)
            m_IsRetINSS = value
        End Set
    End Property
    Private m_IsRetINSS As Integer

    Public Property IsNFe As Integer
        Get
            Return m_IsNFe
        End Get
        Set(ByVal value As Integer)
            m_IsNFe = value
        End Set
    End Property
    Private m_IsNFe As Integer

    Public Property IsNNC As Integer
        Get
            Return m_IsNNC
        End Get
        Set(ByVal value As Integer)
            m_IsNNC = value
        End Set
    End Property
    Private m_IsNNC As Integer

    Public Property IsOptSimples As Integer
        Get
            Return m_IsOptSimples
        End Get
        Set(ByVal value As Integer)
            m_IsOptSimples = value
        End Set
    End Property
    Private m_IsOptSimples As Integer

End Class
