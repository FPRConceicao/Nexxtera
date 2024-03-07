Imports Teleatlantic.TLS.Common
Imports IntegracaoNexxtera

Public Class Exportacao : Inherits Retorno

    Public Property Cliente() As Cliente
        Get
            Return m_Cliente
        End Get
        Set(ByVal value As Cliente)
            m_Cliente = value
        End Set
    End Property
    Private m_Cliente As Cliente


    Public Property NotaFiscal() As NotaFiscal
        Get
            Return m_NotaFiscal
        End Get
        Set(ByVal value As NotaFiscal)
            m_NotaFiscal = value
        End Set
    End Property
    Private m_NotaFiscal As NotaFiscal


    Public Property DetNotaFiscal() As DetNotaFiscal
        Get
            Return m_DetNotaFiscal
        End Get
        Set(ByVal value As DetNotaFiscal)
            m_DetNotaFiscal = value
        End Set
    End Property
    Private m_DetNotaFiscal As DetNotaFiscal


    Public Property Endereco() As Endereco
        Get
            Return m_Endereco
        End Get
        Set(ByVal value As Endereco)
            m_Endereco = value
        End Set
    End Property
    Private m_Endereco As Endereco


    Public Property Email() As Email
        Get
            Return m_Email
        End Get
        Set(ByVal value As Email)
            m_Email = value
        End Set
    End Property
    Private m_Email As Email


    Public Property ContaReceber() As ContaReceber
        Get
            Return m_ContaReceber
        End Get
        Set(ByVal value As ContaReceber)
            m_ContaReceber = value
        End Set
    End Property
    Private m_ContaReceber As ContaReceber

    Public Property AreaFilial() As AreaFilial

    Public Property Parametro() As Parametros

End Class
