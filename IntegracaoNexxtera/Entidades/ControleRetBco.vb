
Public Class ControleRetBco : Inherits Retorno

    Public Property CodBanco() As String
        Get
            Return m_CodBanco
        End Get
        Set(ByVal value As String)
            m_CodBanco = value
        End Set
    End Property
    Private m_CodBanco As String

    Public Property NumAviso() As String
        Get
            Return m_NumAviso
        End Get
        Set(ByVal value As String)
            m_NumAviso = value
        End Set
    End Property
    Private m_NumAviso As String

    Public Property DtArq() As Nullable(Of DateTime)
        Get
            Return m_DtArq
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtArq = value
        End Set
    End Property
    Private m_DtArq As Nullable(Of DateTime)

    Public Property NumCta() As String
        Get
            Return m_NumCta
        End Get
        Set(ByVal value As String)
            m_NumCta = value
        End Set
    End Property
    Private m_NumCta As String
    Public Property numTit As String

End Class
