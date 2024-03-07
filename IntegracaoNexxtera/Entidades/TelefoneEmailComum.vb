Imports Teleatlantic.TLS.Common

Public Class TelefoneEmailComum : Inherits Retorno
    Public Property CodClie() As String
        Get
            Return m_CodClie
        End Get
        Set(ByVal value As String)
            m_CodClie = value
        End Set
    End Property
    Private m_CodClie As String


    Public Property CodFone() As String
        Get
            Return m_CodFone
        End Get
        Set(ByVal value As String)
            m_CodFone = value
        End Set
    End Property
    Private m_CodFone As String

    Public Property Tipo() As String
        Get
            Return m_Tipo
        End Get
        Set(ByVal value As String)
            m_Tipo = value
        End Set
    End Property
    Private m_Tipo As String

    Public Property TipoTel() As String
        Get
            Return m_TipoTel
        End Get
        Set(ByVal value As String)
            m_TipoTel = value
        End Set
    End Property
    Private m_TipoTel As String

    Public Property DataAlteracao() As String
        Get
            Return m_DataAlteracao
        End Get
        Set(ByVal value As String)
            m_DataAlteracao = value
        End Set
    End Property
    Private m_DataAlteracao As String


    Public Property UserAlteracao() As String
        Get
            Return m_UserAlteracao
        End Get
        Set(ByVal value As String)
            m_UserAlteracao = value
        End Set
    End Property
    Private m_UserAlteracao As String

    Public Property RecebeMkt() As String
        Get
            Return m_RecebeMkt
        End Get
        Set(ByVal value As String)
            m_RecebeMkt = value
        End Set
    End Property
    Private m_RecebeMkt As String

    Public Property CodIntClie As String
    Public Property Local As String
    Public Property DDDClie As String
    Public Property Contato As String


End Class
