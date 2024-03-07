Imports Teleatlantic.TLS.Common

Public Class ClienteRonda : Inherits Retorno

    Public Property Regiao() As Regiao
        Get
            Return m_Regiao
        End Get
        Set(ByVal value As Regiao)
            m_Regiao = value
        End Set
    End Property
    Private m_Regiao As Regiao

    Public Property IsTelerondaPrior() As String
        Get
            Return m_IsTelerondaPrior
        End Get
        Set(ByVal value As String)
            m_IsTelerondaPrior = value
        End Set
    End Property
    Private m_IsTelerondaPrior As String

    Public Property SeqTeleRonda() As String
        Get
            Return m_SeqTeleRonda
        End Get
        Set(ByVal value As String)
            m_SeqTeleRonda = value
        End Set
    End Property
    Private m_SeqTeleRonda As String

End Class
