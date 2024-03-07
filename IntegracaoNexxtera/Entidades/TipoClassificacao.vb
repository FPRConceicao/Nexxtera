Imports Teleatlantic.TLS.Common

Public Class TipoClassificacao : Inherits Retorno

    Public Property Protocolo() As String
        Get
            Return m_Protocolo
        End Get
        Set(ByVal value As String)
            m_Protocolo = value
        End Set
    End Property
    Private m_Protocolo As String

    Public Property UsrCad() As String
        Get
            Return m_UsrCad
        End Get
        Set(ByVal value As String)
            m_UsrCad = value
        End Set
    End Property
    Private m_UsrCad As String

    Public Property Obs() As String
        Get
            Return m_Obs
        End Get
        Set(ByVal value As String)
            m_Obs = value
        End Set
    End Property
    Private m_Obs As String

    Public Property CodTipoClassificacao As Integer
        Get
            Return m_CodTipoClassificacao
        End Get
        Set(ByVal value As Integer)
            m_CodTipoClassificacao = value
        End Set
    End Property
    Private m_CodTipoClassificacao As Integer

    Public Property Classificacao() As String
        Get
            Return m_Classificacao
        End Get
        Set(ByVal value As String)
            m_Classificacao = value
        End Set
    End Property
    Private m_Classificacao As String

    Public Property Orcamento() As String
        Get
            Return m_Orcamento
        End Get
        Set(ByVal value As String)
            m_Orcamento = value
        End Set
    End Property
    Private m_Orcamento As String

End Class
