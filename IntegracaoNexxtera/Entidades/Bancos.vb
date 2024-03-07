Imports Teleatlantic.TLS.Common

Public Class Bancos : Inherits Retorno
    Public Property CodBanco() As String
        Get
            Return m_CodBanco
        End Get
        Set(ByVal value As String)
            m_CodBanco = value
        End Set
    End Property
    Private m_CodBanco As String

    Public Property NomeBanco() As String
        Get
            Return m_NomeBanco
        End Get
        Set(ByVal value As String)
            m_NomeBanco = value
        End Set
    End Property
    Private m_NomeBanco As String

    Public Property NomeFantasia() As String
        Get
            Return m_NomeFantasia
        End Get
        Set(ByVal value As String)
            m_NomeFantasia = value
        End Set
    End Property
    Private m_NomeFantasia As String
    Private m_CodBancoNomeBanco As String
    Public Property CodBancoNomeBanco As String
        Get
            Return m_CodBancoNomeBanco
        End Get
        Set(ByVal value As String)
            m_CodBancoNomeBanco = value
        End Set
    End Property
End Class
