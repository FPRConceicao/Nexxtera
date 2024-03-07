Imports Teleatlantic.TLS.Common

Public Class EventoCtaReceber : Inherits Retorno


    Public Property CodEvento() As String
        Get
            Return m_CodEvento
        End Get
        Set(ByVal value As String)
            m_CodEvento = value
        End Set
    End Property
    Private m_CodEvento As String

    Public Property Descricao() As String
        Get
            Return m_Descricao
        End Get
        Set(ByVal value As String)
            m_Descricao = value
        End Set
    End Property
    Private m_Descricao As String


    Public Property SitTit() As String
        Get
            Return m_SitTit
        End Get
        Set(ByVal value As String)
            m_SitTit = value
        End Set
    End Property
    Private m_SitTit As String

End Class
