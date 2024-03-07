
''' <summary>
''' Esta classe herda das classes DDD_DDI e Telefone
''' </summary>
''' <remarks></remarks>
Public Class Email : Inherits TelefoneEmailComum
    Public Property Email() As String
        Get
            Return m_Email
        End Get
        Set(ByVal value As String)
            m_Email = value
        End Set
    End Property
    Private m_Email As String
End Class
