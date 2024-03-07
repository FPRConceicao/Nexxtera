Imports Teleatlantic.TLS.Common


Public Class TempCheckArqRetorno : Inherits Retorno

    Public Property NumTit() As String
        Get
            Return m_NumTit
        End Get
        Set(ByVal value As String)
            m_NumTit = value
        End Set
    End Property
    Private m_NumTit As String

    Public Property CodBanco() As String
        Get
            Return m_CodBanco
        End Get
        Set(ByVal value As String)
            m_CodBanco = value
        End Set
    End Property
    Private m_CodBanco As String

    Public Property NomeBanco() As String
        Get
            Return m_NomeBanco
        End Get
        Set(ByVal value As String)
            m_NomeBanco = value
        End Set
    End Property
    Private m_NomeBanco As String

    Public Property NumAviso() As String
        Get
            Return m_NumAviso
        End Get
        Set(ByVal value As String)
            m_NumAviso = value
        End Set
    End Property
    Private m_NumAviso As String

    Public Property NumCta() As String
        Get
            Return m_NumCta
        End Get
        Set(ByVal value As String)
            m_NumCta = value
        End Set
    End Property
    Private m_NumCta As String

    Public Property DtArq() As Nullable(Of DateTime)
        Get
            Return m_DtArq
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtArq = value
        End Set
    End Property
    Private m_DtArq As Nullable(Of DateTime)

    Public Property CodAgen() As String
        Get
            Return m_CodAgen
        End Get
        Set(ByVal value As String)
            m_CodAgen = value
        End Set
    End Property
    Private m_CodAgen As String

    Public Property DtPagto() As Nullable(Of DateTime)
        Get
            Return m_DtPagto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtPagto = value
        End Set
    End Property
    Private m_DtPagto As Nullable(Of DateTime)

    Public Property NomeCliente() As String
        Get
            Return m_NomeClientem
        End Get
        Set(ByVal value As String)
            m_NomeClientem = value
        End Set
    End Property
    Private m_NomeClientem As String

    Public Property CodOcor() As String
        Get
            Return m_CodOcor
        End Get
        Set(ByVal value As String)
            m_CodOcor = value
        End Set
    End Property
    Private m_CodOcor As String

    Public Property CodIncons() As String
        Get
            Return m_CodIncons
        End Get
        Set(ByVal value As String)
            m_CodIncons = value
        End Set
    End Property
    Private m_CodIncons As String

    Public Property CodBancoDeb() As String
        Get
            Return m_CodBancoDeb
        End Get
        Set(ByVal value As String)
            m_CodBancoDeb = value
        End Set
    End Property
    Private m_CodBancoDeb As String

    Public Property CodAgenDeb() As String
        Get
            Return m_CodAgenDeb
        End Get
        Set(ByVal value As String)
            m_CodAgenDeb = value
        End Set
    End Property
    Private m_CodAgenDeb As String


    Public Property VlrMulta() As Double
        Get
            Return If(m_VlrMulta <> 0, m_VlrMulta, 0)
        End Get
        Set(ByVal value As Double)
            m_VlrMulta = value
        End Set
    End Property

    Private m_VlrMulta As Double


    Public Property VlrJuros() As Double
        Get
            Return If(m_VlrJuros <> 0, m_VlrJuros, 0)
        End Get
        Set(ByVal value As Double)
            m_VlrJuros = value
        End Set
    End Property

    Private m_VlrJuros As Double

    Public Property TipoDup() As String
        Get
            Return m_TipoDup
        End Get
        Set(ByVal value As String)
            m_TipoDup = value
        End Set
    End Property
    Private m_TipoDup As String

    Public Property Situacao() As String
        Get
            Return m_Situacao
        End Get
        Set(ByVal value As String)
            m_Situacao = value
        End Set
    End Property
    Private m_Situacao As String

    Public Property CodNumCtaDeb() As String
        Get
            Return m_CodNumCtaDeb
        End Get
        Set(ByVal value As String)
            m_CodNumCtaDeb = value
        End Set
    End Property
    Private m_CodNumCtaDeb As String

    Public Property Valor() As Double
        Get
            Return m_Valor
        End Get
        Set(ByVal value As Double)
            m_Valor = value
        End Set
    End Property
    Private m_Valor As Double

    Public Property Endereco() As String
        Get
            Return m_Endereco
        End Get
        Set(ByVal value As String)
            m_Endereco = value
        End Set
    End Property
    Private m_Endereco As String

    Public Property UF() As String
        Get
            Return m_UF
        End Get
        Set(ByVal value As String)
            m_UF = value
        End Set
    End Property
    Private m_UF As String

    Public Property Cidade() As String
        Get
            Return m_Cidade
        End Get
        Set(ByVal value As String)
            m_Cidade = value
        End Set
    End Property
    Private m_Cidade As String

    Public Property Cep() As String
        Get
            Return m_Cep
        End Get
        Set(ByVal value As String)
            m_Cep = value
        End Set
    End Property
    Private m_Cep As String

    Public Property Arquivo() As String
        Get
            Return m_Arquivo
        End Get
        Set(ByVal value As String)
            m_Arquivo = value
        End Set
    End Property
    Private m_Arquivo As String

    Public Property Mensagem() As String
        Get
            Return m_Mensagem
        End Get
        Set(ByVal value As String)
            m_Mensagem = value
        End Set
    End Property
    Private m_Mensagem As String

    Public Property SeqTit() As String
        Get
            Return m_SeqTit
        End Get
        Set(ByVal value As String)
            m_SeqTit = value
        End Set
    End Property
    Private m_SeqTit As String

    Public Property DtVcto() As Nullable(Of DateTime)
        Get
            Return m_DtVcto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtVcto = value
        End Set
    End Property
    Private m_DtVcto As Nullable(Of DateTime)

    Public Property NossoNumeroBco() As String
        Get
            Return m_NossoNumeroBco
        End Get
        Set(ByVal value As String)
            m_NossoNumeroBco = value
        End Set
    End Property
    Private m_NossoNumeroBco As String
    Public Property CGC_CPF As String
    Public Property Ocorrencia As String
    Public Property Inconsistencia As String
    Public Property Mensalidade As Double?
    Public Property DiaPgto As String
    Public Property DDD As String
    Public Property Telefone As String
    Public Property Email As String
End Class
