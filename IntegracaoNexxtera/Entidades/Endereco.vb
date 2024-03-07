Imports Teleatlantic.TLS.Common


''' <summary>
''' Endidade de Endereço.
''' </summary>
''' <remarks>
''' 
''' 
''' 
''' </remarks>
Public Class Endereco : Inherits Retorno


    Public Property CodigoCliente() As String
        Get
            Return m_CodigoCliente
        End Get
        Set(ByVal value As String)
            m_CodigoCliente = value
        End Set
    End Property
    Private m_CodigoCliente As String

    Public Property CodigoEndereco() As Long
        Get
            Return m_CodigoEndereco
        End Get
        Set(ByVal value As Long)
            m_CodigoEndereco = value
        End Set
    End Property
    Private m_CodigoEndereco As Long

    Public Property Tipo() As String
        Get
            Return m_Tipo
        End Get
        Set(ByVal value As String)
            m_Tipo = value
        End Set
    End Property
    Private m_Tipo As String

    Public Property CodigoPais() As String
        Get
            Return m_CodigoPais
        End Get
        Set(ByVal value As String)
            m_CodigoPais = value
        End Set
    End Property
    Private m_CodigoPais As String

    Public Property Endereco() As String
        Get
            Return m_Endereco
        End Get
        Set(ByVal value As String)
            m_Endereco = value
        End Set
    End Property
    Private m_Endereco As String

    Public Property NumeroEndereco() As String
        Get
            Return m_NumeroEndereco
        End Get
        Set(ByVal value As String)
            m_NumeroEndereco = value
        End Set
    End Property
    Private m_NumeroEndereco As String

    Public Property Bairro() As String
        Get
            Return m_Bairro
        End Get
        Set(ByVal value As String)
            m_Bairro = value
        End Set
    End Property
    Private m_Bairro As String

    Public Property Complemento() As String
        Get
            Return m_Complemento
        End Get
        Set(ByVal value As String)
            m_Complemento = value
        End Set
    End Property
    Private m_Complemento As String

    Public Property Cep() As String
        Get
            Return m_Cep
        End Get
        Set(ByVal value As String)
            m_Cep = value
        End Set
    End Property
    Private m_Cep As String

    Public Property Cidade() As String
        Get
            Return m_Cidade
        End Get
        Set(ByVal value As String)
            m_Cidade = value
        End Set
    End Property
    Private m_Cidade As String

    Public Property UF() As String
        Get
            Return m_UF
        End Get
        Set(ByVal value As String)
            m_UF = value
        End Set
    End Property
    Private m_UF As String

    Public Property Contato() As String
        Get
            Return m_Contato
        End Get
        Set(ByVal value As String)
            m_Contato = value
        End Set
    End Property
    Private m_Contato As String

    Public Property Cargo() As String
        Get
            Return m_Cargo
        End Get
        Set(ByVal value As String)
            m_Cargo = value
        End Set
    End Property
    Private m_Cargo As String


    Public Property DataAlteracao() As DateTime
        Get
            Return m_DataAlteracao
        End Get
        Set(ByVal value As DateTime)
            m_DataAlteracao = value
        End Set
    End Property
    Private m_DataAlteracao As DateTime

    Public Property UsuarioAlteracao() As String
        Get
            Return m_UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            m_UsuarioAlteracao = value
        End Set
    End Property
    Private m_UsuarioAlteracao As String

    Public Property CodOrc() As String
        Get
            Return m_CodOrc
        End Get
        Set(ByVal value As String)
            m_CodOrc = value
        End Set
    End Property
    Private m_CodOrc As String

    Public Property TipoLogradouro As String
    Public Property AbreviacaoLogradouro As String

End Class
