''' <summary>
''' Entidade de código do pais e código de área
''' </summary>
''' <remarks>
''' 
''' Data Criação:     12/04/2011
''' Auttor:           Wolney Alexandre Fernandes
''' 
''' </remarks>

Public Class DDD_DDI : Inherits TelefoneEmailComum
    Public Property Local() As String
        Get
            Return m_Local
        End Get
        Set(ByVal value As String)
            m_Local = value
        End Set
    End Property
    Private m_Local As String

    Public Property DDDClie() As String
        Get
            Return m_DDDClie
        End Get
        Set(ByVal value As String)
            m_DDDClie = value
        End Set
    End Property
    Private m_DDDClie As String

    Public Property DDDClie2() As String
        Get
            Return m_DDDClie2
        End Get
        Set(ByVal value As String)
            m_DDDClie2 = value
        End Set
    End Property
    Private m_DDDClie2 As String
End Class
