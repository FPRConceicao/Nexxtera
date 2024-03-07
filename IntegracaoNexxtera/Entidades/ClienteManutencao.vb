Imports Teleatlantic.TLS.Common

Public Class ClienteManutencao : Inherits Retorno

   

    Public Property DtUltManutPrev() As DateTime
        Get
            Return m_DtUltManutPrev
        End Get
        Set(ByVal value As DateTime)
            m_DtUltManutPrev = value
        End Set
    End Property
    Private m_DtUltManutPrev As DateTime

    Public Property CodEmpManut() As String
        Get
            Return m_CodEmpManut
        End Get
        Set(ByVal value As String)
            m_CodEmpManut = value
        End Set
    End Property
    Private m_CodEmpManut As String

    Public Property CodEmpManutTV() As String
        Get
            Return m_CodEmpManutTV
        End Get
        Set(ByVal value As String)
            m_CodEmpManutTV = value
        End Set
    End Property
    Private m_CodEmpManutTV As String
End Class
