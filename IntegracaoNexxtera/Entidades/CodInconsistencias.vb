Imports Teleatlantic.TLS.Common

Public Class CodInconsistencias : Inherits Retorno
    Public Property CodIncons() As String
        Get
            Return m_CodIncons
        End Get
        Set(ByVal value As String)
            m_CodIncons = value
        End Set
    End Property
    Private m_CodIncons As String

    Public Property Mensagem() As String
        Get
            Return m_Mensagem
        End Get
        Set(ByVal value As String)
            m_Mensagem = value
        End Set
    End Property
    Private m_Mensagem As String

    Public Property CodOcor() As String
        Get
            Return m_CodOcor
        End Get
        Set(ByVal value As String)
            m_CodOcor = value
        End Set
    End Property
    Private m_CodOcor As String

    Public Property CodBanco() As String
        Get
            Return m_CodBanco
        End Get
        Set(ByVal value As String)
            m_CodBanco = value
        End Set
    End Property
    Private m_CodBanco As String

End Class
