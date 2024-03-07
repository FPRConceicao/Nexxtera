Imports Teleatlantic.TLS.Common

Public Class RetornoBco : Inherits Retorno

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


    Public Property CodAgen() As String
        Get
            Return m_CodAgen
        End Get
        Set(ByVal value As String)
            m_CodAgen = value
        End Set
    End Property
    Private m_CodAgen As String


    Public Property NumCta() As String
        Get
            Return m_NumCta
        End Get
        Set(ByVal value As String)
            m_NumCta = value
        End Set
    End Property
    Private m_NumCta As String


    Public Property VlrPago() As Double
        Get
            Return m_VlrPago
        End Get
        Set(ByVal value As Double)
            m_VlrPago = value
        End Set
    End Property
    Private m_VlrPago As Double


    Public Property VlrJuros() As Double
        Get
            Return m_VlrJuros
        End Get
        Set(ByVal value As Double)
            m_VlrJuros = value
        End Set
    End Property
    Private m_VlrJuros As Double



    Public Property VlrMulta() As Double
        Get
            Return If(m_VlrMulta <> 0, m_VlrMulta, 0)
        End Get
        Set(ByVal value As Double)
            m_VlrMulta = value
        End Set
    End Property

    Private m_VlrMulta As Double



    Public Property VlrDesc() As Double
        Get
            Return m_VlrDesc
        End Get
        Set(ByVal value As Double)
            m_VlrDesc = value
        End Set
    End Property
    Private m_VlrDesc As Double


    Public Property VlrIOF() As Double
        Get
            Return m_VlrIOF
        End Get
        Set(ByVal value As Double)
            m_VlrIOF = value
        End Set
    End Property
    Private m_VlrIOF As Double


    Public Property VlrAbat() As Double
        Get
            Return m_VlrAbat
        End Get
        Set(ByVal value As Double)
            m_VlrAbat = value
        End Set
    End Property
    Private m_VlrAbat As Double


    Public Property Processado() As String
        Get
            Return m_Processado
        End Get
        Set(ByVal value As String)
            m_Processado = value
        End Set
    End Property
    Private m_Processado As String


    Public Property DtVcto() As Nullable(Of DateTime)
        Get
            Return m_DtVcto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtVcto = value
        End Set
    End Property
    Private m_DtVcto As Nullable(Of DateTime)


    Public Property DtPagto() As Nullable(Of DateTime)
        Get
            Return m_DtPagto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtPagto = value
        End Set
    End Property
    Private m_DtPagto As Nullable(Of DateTime)


    Public Property DtArq() As Nullable(Of DateTime)
        Get
            Return m_DtArq
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtArq = value
        End Set
    End Property
    Private m_DtArq As Nullable(Of DateTime)


    Public Property VlrInd() As Double
        Get
            Return m_VlrInd
        End Get
        Set(ByVal value As Double)
            m_VlrInd = value
        End Set
    End Property
    Private m_VlrInd As Double

End Class
