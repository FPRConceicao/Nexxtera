Imports Teleatlantic.TLS.Common

Public Class BxAutosErros : Inherits Retorno

    Public Property NumAviso() As String
        Get
            Return m_NumAviso
        End Get
        Set(ByVal value As String)
            m_NumAviso = value
        End Set
    End Property
    Private m_NumAviso As String

    Public Property NumTit() As String
        Get
            Return m_NumTit
        End Get
        Set(ByVal value As String)
            m_NumTit = value
        End Set
    End Property
    Private m_NumTit As String

    Public Property SeqTit() As String
        Get
            Return m_SeqTit
        End Get
        Set(ByVal value As String)
            m_SeqTit = value
        End Set
    End Property
    Private m_SeqTit As String

    Public Property DtEmissao() As Nullable(Of DateTime)
        Get
            Return m_DtEmissao
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtEmissao = value
        End Set
    End Property
    Private m_DtEmissao As Nullable(Of DateTime)

End Class
