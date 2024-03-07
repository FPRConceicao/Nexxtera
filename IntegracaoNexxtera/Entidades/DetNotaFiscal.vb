Imports Teleatlantic.TLS.Common


Public Class DetNotaFiscal : Inherits Retorno

    Public Property CodProd() As String
        Get
            Return m_CodProd
        End Get
        Set(ByVal value As String)
            m_CodProd = value
        End Set
    End Property
    Private m_CodProd As String


    Public Property NumNF() As String
        Get
            Return m_NumNF
        End Get
        Set(ByVal value As String)
            m_NumNF = value
        End Set
    End Property
    Private m_NumNF As String


    Public Property SerieNF() As String
        Get
            Return m_SerieNF
        End Get
        Set(ByVal value As String)
            m_SerieNF = value
        End Set
    End Property
    Private m_SerieNF As String


    Public Property DescrItem() As String
        Get
            Return m_DescrItem
        End Get
        Set(ByVal value As String)
            m_DescrItem = value
        End Set
    End Property
    Private m_DescrItem As String


    Public Property VlrUnit() As Double
        Get
            Return m_VlrUnit
        End Get
        Set(ByVal value As Double)
            m_VlrUnit = value
        End Set
    End Property
    Private m_VlrUnit As Double


    Public Property BaseIss() As Double
        Get
            Return m_BaseIss
        End Get
        Set(ByVal value As Double)
            m_BaseIss = value
        End Set
    End Property
    Private m_BaseIss As Double


    Public Property BaseICMS() As Double
        Get
            Return m_BaseICMS
        End Get
        Set(ByVal value As Double)
            m_BaseICMS = value
        End Set
    End Property
    Private m_BaseICMS As Double


    Public Property AliqICMS() As Double
        Get
            Return m_AliqICMS
        End Get
        Set(ByVal value As Double)
            m_AliqICMS = value
        End Set
    End Property
    Private m_AliqICMS As Double


    Public Property BaseIRRF() As Double
        Get
            Return m_BaseIRRF
        End Get
        Set(ByVal value As Double)
            m_BaseIRRF = value
        End Set
    End Property
    Private m_BaseIRRF As Double


    Public Property AliqISS() As Double
        Get
            Return m_AliqISS
        End Get
        Set(ByVal value As Double)
            m_AliqISS = value
        End Set
    End Property
    Private m_AliqISS As Double


    Public Property AliqISSRet() As Double
        Get
            Return m_AliqISSRet
        End Get
        Set(ByVal value As Double)
            m_AliqISSRet = value
        End Set
    End Property
    Private m_AliqISSRet As Double


    Public Property VlrICMS() As Double
        Get
            Return m_VlrICMS
        End Get
        Set(ByVal value As Double)
            m_VlrICMS = value
        End Set
    End Property
    Private m_VlrICMS As Double


    Public Property DtEmissao() As DateTime
        Get
            Return m_DtEmissao
        End Get
        Set(ByVal value As DateTime)
            m_DtEmissao = value
        End Set
    End Property
    Private m_DtEmissao As DateTime


    Public Property SubTotal() As Double
        Get
            Return m_SubTotal
        End Get
        Set(ByVal value As Double)
            m_SubTotal = value
        End Set
    End Property
    Private m_SubTotal As Double


    Public Property Qtde() As Integer
        Get
            Return m_Qtde
        End Get
        Set(ByVal value As Integer)
            m_Qtde = value
        End Set
    End Property
    Private m_Qtde As Integer


    Public Property ItemNF() As Integer
        Get
            Return m_ItemNF
        End Get
        Set(ByVal value As Integer)
            m_ItemNF = value
        End Set
    End Property
    Private m_ItemNF As Integer


    Public Property VlrIss() As Double
        Get
            Return m_VlrIss
        End Get
        Set(ByVal value As Double)
            m_VlrIss = value
        End Set
    End Property
    Private m_VlrIss As Double


    Public Property VlrIRRF() As Double
        Get
            Return m_VlrIRRF
        End Get
        Set(ByVal value As Double)
            m_VlrIRRF = value
        End Set
    End Property
    Private m_VlrIRRF As Double


    Public Property AliqIRRF() As Double
        Get
            Return m_AliqIRRF
        End Get
        Set(ByVal value As Double)
            m_AliqIRRF = value
        End Set
    End Property
    Private m_AliqIRRF As Double


    Public Property IdFilial() As String
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As String)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As String


    Public Property AliqNovaCOFINS() As Double
        Get
            Return m_AliqNovaCOFINS
        End Get
        Set(ByVal value As Double)
            m_AliqNovaCOFINS = value
        End Set
    End Property
    Private m_AliqNovaCOFINS As Double


    Public Property AliqRetIR() As Double
        Get
            Return m_AliqRetIR
        End Get
        Set(ByVal value As Double)
            m_AliqRetIR = value
        End Set
    End Property
    Private m_AliqRetIR As Double


    Public Property AliqRetINSS() As Double
        Get
            Return m_AliqRetINSS
        End Get
        Set(ByVal value As Double)
            m_AliqRetINSS = value
        End Set
    End Property
    Private m_AliqRetINSS As Double


    Public Property PrecoMedio() As Double
        Get
            Return m_PrecoMedio
        End Get
        Set(ByVal value As Double)
            m_PrecoMedio = value
        End Set
    End Property
    Private m_PrecoMedio As Double


    Public Property VlrIPI() As Double
        Get
            Return m_VlrIPI
        End Get
        Set(ByVal value As Double)
            m_VlrIPI = value
        End Set
    End Property
    Private m_VlrIPI As Double


    Public Property CFOP() As String
        Get
            Return m_CFOP
        End Get
        Set(ByVal value As String)
            m_CFOP = value
        End Set
    End Property
    Private m_CFOP As String


    Public Property AliqIPI() As Double
        Get
            Return m_AliqIPI
        End Get
        Set(ByVal value As Double)
            m_AliqIPI = value
        End Set
    End Property
    Private m_AliqIPI As Double


    Public Property VlrDescVenda() As Double
        Get
            Return m_VlrDescVenda
        End Get
        Set(ByVal value As Double)
            m_VlrDescVenda = value
        End Set
    End Property
    Private m_VlrDescVenda As Double

    Public Property NCM() As String
        Get
            Return m_NCM
        End Get
        Set(ByVal value As String)
            m_NCM = value
        End Set
    End Property
    Private m_NCM As String

    Public Property IdNFEntrada As String
        Get
            Return m_IdNFEntrada
        End Get
        Set(ByVal value As String)
            m_IdNFEntrada = value
        End Set
    End Property
    Private m_IdNFEntrada As String

    Public Property TipoEquip As String
        Get
            Return m_TipoEquip
        End Get
        Set(ByVal value As String)
            m_TipoEquip = value
        End Set
    End Property
    Private m_TipoEquip As String

    Private _UniMed As String
    Public Property UniMed() As String
        Get
            Return _UniMed
        End Get
        Set(ByVal value As String)
            _UniMed = value
        End Set
    End Property

End Class
