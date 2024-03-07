Imports Teleatlantic.TLS.Common

''' <summary>
''' 
''' </summary>
''' <remarks>
''' 
''' </remarks>
Public Class CategoriaPedido : Inherits Retorno
    Public Property Codigo() As String
        Get
            Return m_Codigo
        End Get
        Set(ByVal value As String)
            m_Codigo = value
        End Set
    End Property
    Private m_Codigo As String

    Public Property Descricao() As String
        Get
            Return m_Descricao
        End Get
        Set(ByVal value As String)
            m_Descricao = value
        End Set
    End Property
    Private m_Descricao As String

    Public Property Situacao() As String
        Get
            Return m_Situacao
        End Get
        Set(ByVal value As String)
            m_Situacao = value
        End Set
    End Property
    Private m_Situacao As String

    Public Property PedidoVenda() As String
        Get
            Return m_PedidoVenda
        End Get
        Set(ByVal value As String)
            m_PedidoVenda = value
        End Set
    End Property
    Private m_PedidoVenda As String

    Public Property Financeiro() As String
        Get
            Return m_Financeiro
        End Get
        Set(ByVal value As String)
            m_Financeiro = value
        End Set
    End Property
    Private m_Financeiro As String

    Public Property Gerencial() As String
        Get
            Return m_Gerencial
        End Get
        Set(ByVal value As String)
            m_Gerencial = value
        End Set
    End Property
    Private m_Gerencial As String

    Public Property DescDC() As Double
        Get
            Return m_DescDC
        End Get
        Set(ByVal value As Double)
            m_DescDC = value
        End Set
    End Property
    Private m_DescDC As Double

    Public Property SubCatGerencial()
        Get
            Return m_SubCatGerencial
        End Get
        Set(ByVal value)
            m_SubCatGerencial = value
        End Set
    End Property
    Private m_SubCatGerencial As String

    Public Property ClassificacaoCtbl
        Get
            Return m_ClassificaoCtbl
        End Get
        Set(ByVal value)
            m_ClassificaoCtbl = value
        End Set
    End Property
    Private m_ClassificaoCtbl As String

    Public Property ContaCtblCliente
        Get
            Return m_ContaCtblCliente
        End Get
        Set(ByVal value)
            m_ContaCtblCliente = value
        End Set
    End Property
    Private m_ContaCtblCliente As String
    Public Property PesoScorePosVendas As Integer
End Class
