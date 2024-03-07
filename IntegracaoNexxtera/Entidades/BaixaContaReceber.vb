
Imports Teleatlantic.TLS.Common

Public Class BaixaContaReceber : Inherits Retorno

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

    Public Property DtPgto() As Nullable(Of DateTime)
        Get
            Return m_DtPgto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtPgto = value
        End Set
    End Property
    Private m_DtPgto As Nullable(Of DateTime)

    'Public Property VlrPgto() As String
    '    Get
    '        Return m_VlrPgto
    '    End Get
    '    Set(ByVal value As String)
    '        m_VlrPgto = value
    '    End Set
    'End Property
    'Private m_VlrPgto As String

    Public Property CodEvento() As String
        Get
            Return m_CodEvento
        End Get
        Set(ByVal value As String)
            m_CodEvento = value
        End Set
    End Property
    Private m_CodEvento As String

    Public Property VlrMulta() As Double
        Get
            Return m_VlrMulta
        End Get
        Set(ByVal value As Double)
            m_VlrMulta = value
        End Set
    End Property
    Private m_VlrMulta As Double

    Public Property VlrJuros() As Double
        Get
            Return m_VlrJuros
        End Get
        Set(ByVal value As Double)
            m_VlrJuros = value
        End Set
    End Property
    Private m_VlrJuros As Double

    Public Property VlrDesc() As Double
        Get
            Return m_VlrDesc
        End Get
        Set(ByVal value As Double)
            m_VlrDesc = value
        End Set
    End Property
    Private m_VlrDesc As Double

    Public Property VlrDevol() As Double
        Get
            Return m_VlrDevol
        End Get
        Set(ByVal value As Double)
            m_VlrDevol = value
        End Set
    End Property
    Private m_VlrDevol As Double

    Public Property VlrDescAbat() As Double
        Get
            Return m_VlrDescAbat
        End Get
        Set(ByVal value As Double)
            m_VlrDescAbat = value
        End Set
    End Property
    Private m_VlrDescAbat As Double

    Public Property VlrVarCamb() As Double
        Get
            Return m_VlrVarCamb
        End Get
        Set(ByVal value As Double)
            m_VlrVarCamb = value
        End Set
    End Property
    Private m_VlrVarCamb As Double

    Public Property VlrAbat() As Double
        Get
            Return m_VlrAbat
        End Get
        Set(ByVal value As Double)
            m_VlrAbat = value
        End Set
    End Property
    Private m_VlrAbat As Double

    Public Property VlrAcresc() As Double
        Get
            Return m_VlrAcresc
        End Get
        Set(ByVal value As Double)
            m_VlrAcresc = value
        End Set
    End Property
    Private m_VlrAcresc As Double

    Public Property VlrPago() As Double
        Get
            Return m_VlrPago
        End Get
        Set(ByVal value As Double)
            m_VlrPago = value
        End Set
    End Property
    Private m_VlrPago As Double


    Public Property ObsBaixa() As String
        Get
            Return m_ObsBaixa
        End Get
        Set(ByVal value As String)
            m_ObsBaixa = value
        End Set
    End Property
    Private m_ObsBaixa As String

    Public Property OrigBaixa() As String
        Get
            Return m_OrigBaixa
        End Get
        Set(ByVal value As String)
            m_OrigBaixa = value
        End Set
    End Property
    Private m_OrigBaixa As String

    Public Property DtInc() As DateTime
        Get
            Return m_DtInc
        End Get
        Set(ByVal value As DateTime)
            m_DtInc = value
        End Set
    End Property
    Private m_DtInc As DateTime

    Public Property UsrInc() As String
        Get
            Return m_UsrInc
        End Get
        Set(ByVal value As String)
            m_UsrInc = value
        End Set
    End Property
    Private m_UsrInc As String

    Public Property NumLote() As String
        Get
            Return m_NumLote
        End Get
        Set(ByVal value As String)
            m_NumLote = value
        End Set
    End Property
    Private m_NumLote As String

    Public Property MesAnoLote() As String
        Get
            Return m_MesAnoLote
        End Get
        Set(ByVal value As String)
            m_MesAnoLote = value
        End Set
    End Property
    Private m_MesAnoLote As String

    Public Property NumLcto() As String
        Get
            Return m_NumLcto
        End Get
        Set(ByVal value As String)
            m_NumLcto = value
        End Set
    End Property
    Private m_NumLcto As String

    Public Property SeqLctoLote() As Integer
        Get
            Return m_SeqLctoLote
        End Get
        Set(ByVal value As Integer)
            m_SeqLctoLote = value
        End Set
    End Property
    Private m_SeqLctoLote As Integer

    Public Property NumUltLcto() As String
        Get
            Return m_NumUltLcto
        End Get
        Set(ByVal value As String)
            m_NumUltLcto = value
        End Set
    End Property
    Private m_NumUltLcto As String

    Public Property CodEmpCtbl() As String
        Get
            Return m_CodEmpCtbl
        End Get
        Set(ByVal value As String)
            m_CodEmpCtbl = value
        End Set
    End Property
    Private m_CodEmpCtbl As String

    Public Property TipoDup() As String
        Get
            Return m_TipoDup
        End Get
        Set(ByVal value As String)
            m_TipoDup = value
        End Set
    End Property
    Private m_TipoDup As String

    Public Property CodBanco() As String
        Get
            Return m_CodBanco
        End Get
        Set(ByVal value As String)
            m_CodBanco = value
        End Set
    End Property
    Private m_CodBanco As String

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

    Public Property DtCredito() As DateTime
        Get
            Return m_DtCredito
        End Get
        Set(ByVal value As DateTime)
            m_DtCredito = value
        End Set
    End Property
    Private m_DtCredito As DateTime

    Public Property IdFilial() As String
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As String)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As String

    Public Property TipoPagamento() As String
        Get
            Return m_TipoPagamento
        End Get
        Set(ByVal value As String)
            m_TipoPagamento = value
        End Set
    End Property
    Private m_TipoPagamento As String

    Public Property VlrPrevPgAcordo() As Double
        Get
            Return m_VlrPrevPgAcordo
        End Get
        Set(ByVal value As Double)
            m_VlrPrevPgAcordo = value
        End Set
    End Property
    Private m_VlrPrevPgAcordo As Double
    ''' <summary>
    ''' Utilizada na exclusão de lançamento de cta corrente
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Total As Integer
        Get
            Return m_Total
        End Get
        Set(ByVal value As Integer)
            m_Total = value
        End Set
    End Property
    Private m_Total As Integer
    Public Property Saldo As Double
        Get
            Return m_Saldo
        End Get
        Set(ByVal value As Double)
            m_Saldo = value
        End Set
    End Property
    Private m_Saldo As Double
    Public Property TipoPgDI As String
    Public Property BancoBaixa As String
End Class
