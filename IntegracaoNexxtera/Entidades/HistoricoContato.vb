Imports Teleatlantic.TLS.Common

Public Class HistoricoContato : Inherits Retorno

    Public Property CodIntClie() As String
        Get
            Return m_CodIntClie
        End Get
        Set(ByVal value As String)
            m_CodIntClie = value
        End Set
    End Property
    Private m_CodIntClie As String

    Public Property Data() As String
        Get
            Return m_Data
        End Get
        Set(ByVal value As String)
            m_Data = value
        End Set
    End Property
    Private m_Data As String

    Public Property DataHora() As DateTime

        Get
            Return m_DataHora
        End Get
        Set(ByVal value As DateTime)
            m_DataHora = value
        End Set
    End Property
    Private m_DataHora As DateTime

    Public Property Usuario() As String
        Get
            Return m_Usuario
        End Get
        Set(ByVal value As String)
            m_Usuario = value
        End Set
    End Property
    Private m_Usuario As String

    Public Property Mensagem() As String
        Get
            Return m_Mensagem
        End Get
        Set(ByVal value As String)
            m_Mensagem = value
        End Set
    End Property
    Private m_Mensagem As String

    Public Property RazaoSocial() As String
        Get
            Return m_RazaoSocial
        End Get
        Set(ByVal value As String)
            m_RazaoSocial = value
        End Set
    End Property
    Private m_RazaoSocial As String

    'Novo campo "Setor" do usuário que registrou o chamado - Lucas 30/09/2014
    Private m_DescSetor As String
    Public Property DescSetor() As String
        Get
            Return m_DescSetor
        End Get
        Set(ByVal value As String)
            m_DescSetor = value
        End Set
    End Property

    Public Property TipoContato As String
    Public Property IdReclamacao As String
    Public Property IdReclamacaoSecundaria As String
    Public Property Protocolo As String
    Public Property IdReclamacaoTerciaria As String
    Public Property Criticidade As String
    Public Property ConfirmacoesNecessarias As String

End Class
