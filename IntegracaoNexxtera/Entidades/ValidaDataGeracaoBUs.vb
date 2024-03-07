Imports Teleatlantic.TLS.Common

Public Class ValidaDataGeracaoBUs : Inherits Retorno
    Public Property DtResult() As Nullable(Of DateTime)
        Get
            Return m_DtResult
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtResult = value
        End Set
    End Property
    Private m_DtResult As Nullable(Of DateTime)

End Class
