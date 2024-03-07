''' <summary>
''' Entidade de numero de telefone ou fax.
''' </summary>
''' <remarks>
''' 
''' Data Criação:     12/04/2011
''' Auttor:           Wolney Alexandre Fernandes
''' 
''' </remarks>

Public Class Telefone : Inherits DDD_DDI

    Public Property Numero() As String
        Get
            Return m_Numero
        End Get
        Set(ByVal value As String)
            m_Numero = value
        End Set
    End Property
    Private m_Numero As String

    Public Property Contato As String
End Class
