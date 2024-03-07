Imports Teleatlantic.TLS.Common

Public Class ClienteVenda : Inherits Retorno

    Public Property CodCpgt() As String
        Get
            Return m_CodCpgt
        End Get
        Set(ByVal value As String)
            m_CodCpgt = value
        End Set
    End Property
    Private m_CodCpgt As String

    Public Property CodCpgtCre() As String
        Get
            Return m_CodCpgtCre
        End Get
        Set(ByVal value As String)
            m_CodCpgtCre = value
        End Set
    End Property
    Private m_CodCpgtCre As String

    Public Property ObsVendas() As String
        Get
            Return m_ObsVendas
        End Get
        Set(ByVal value As String)
            m_ObsVendas = value
        End Set
    End Property
    Private m_ObsVendas As String

    Public Property VlrTarBanc() As Double
        Get
            Return m_VlrTarBanc
        End Get
        Set(ByVal value As Double)
            m_VlrTarBanc = value
        End Set
    End Property
    Private m_VlrTarBanc As Double

    Public Property PlVda() As String
        Get
            Return m_PlVda
        End Get
        Set(ByVal value As String)
            m_PlVda = value
        End Set
    End Property
    Private m_PlVda As String

    Public Property CatPed() As String
        Get
            Return m_CatPed
        End Get
        Set(ByVal value As String)
            m_CatPed = value
        End Set
    End Property
    Private m_CatPed As String

End Class
