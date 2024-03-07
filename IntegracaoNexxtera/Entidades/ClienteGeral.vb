Imports Teleatlantic.TLS.Common

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class ClienteGeral : Inherits Retorno

    Public Property CodClie() As String
        Get
            Return m_CodClie
        End Get
        Set(ByVal value As String)
            m_CodClie = value
        End Set
    End Property
    Private m_CodClie As String


    Public Property CodIntClie() As String
        Get
            Return m_CodIntClie
        End Get
        Set(ByVal value As String)
            m_CodIntClie = value
        End Set
    End Property
    Private m_CodIntClie As String


    Public Property FisiJuri() As String
        Get
            Return m_FisiJuri
        End Get
        Set(ByVal value As String)
            m_FisiJuri = value
        End Set
    End Property
    Private m_FisiJuri As String


    Public Property RazaoSocial() As String
        Get
            Return m_RazaoSocial
        End Get
        Set(ByVal value As String)
            m_RazaoSocial = value
        End Set
    End Property
    Private m_RazaoSocial As String


    Public Property NomeFantasia() As String
        Get
            Return m_NomeFantasia
        End Get
        Set(ByVal value As String)
            m_NomeFantasia = value
        End Set
    End Property
    Private m_NomeFantasia As String


    Public Property CGC_CPF() As String
        Get
            Return m_CGC_CPF
        End Get
        Set(ByVal value As String)
            m_CGC_CPF = value
        End Set
    End Property
    Private m_CGC_CPF As String


    Public Property Estabelecimento() As String
        Get
            Return m_Estabelecimento
        End Get
        Set(ByVal value As String)
            m_Estabelecimento = value
        End Set
    End Property
    Private m_Estabelecimento As String

    Public Property TipoEstabelecimento() As String
        Get
            Return m_TipoEstabelecimento
        End Get
        Set(ByVal value As String)
            m_TipoEstabelecimento = value
        End Set
    End Property
    Private m_TipoEstabelecimento As String


    Public Property CodStatus() As String
        Get
            Return m_CodStatus
        End Get
        Set(ByVal value As String)
            m_CodStatus = value
        End Set
    End Property
    Private m_CodStatus As String

    Public Property TipoAliquota() As String
        Get
            Return m_TipoAliquota
        End Get
        Set(ByVal value As String)
            m_TipoAliquota = value
        End Set
    End Property
    Private m_TipoAliquota As String

    Public Property TipoStatus() As String
        Get
            Return m_TipoStatus
        End Get
        Set(ByVal value As String)
            m_TipoStatus = value
        End Set
    End Property
    Private m_TipoStatus As String

    Public Property PermiteOsCanc() As Boolean
        Get
            Return m_PermiteOsCanc
        End Get
        Set(ByVal value As Boolean)
            m_PermiteOsCanc = value
        End Set
    End Property
    Private m_PermiteOsCanc As String

    Public Property DtIniPgto() As Nullable(Of DateTime)
        Get
            Return m_DtIniPgto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtIniPgto = value
        End Set
    End Property
    Private m_DtIniPgto As Nullable(Of DateTime)

    Public Property DtFimMonit() As Nullable(Of DateTime)
        Get
            Return m_DtFimMonit
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtFimMonit = value
        End Set
    End Property
    Private m_DtFimMonit As Nullable(Of DateTime)

    Public Property UnidadeAntiga() As String
        Get
            Return m_UnidadeAntiga
        End Get
        Set(ByVal value As String)
            m_UnidadeAntiga = value
        End Set
    End Property
    Private m_UnidadeAntiga As String

    Public Property IdFilial() As String
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As String)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As String

    Public Property OmbudsmanConf() As String
        Get
            Return m_OmbudsmanConf
        End Get
        Set(ByVal value As String)
            m_OmbudsmanConf = value
        End Set
    End Property
    Private m_OmbudsmanConf As String

    Public Property Estrelas() As String
        Get
            Return m_Estrelas
        End Get
        Set(ByVal value As String)
            m_Estrelas = value
        End Set
    End Property
    Private m_Estrelas As String

    Public Property DtLibCobranca() As Nullable(Of DateTime)
        Get
            Return m_DtLibCobranca
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtLibCobranca = value
        End Set
    End Property
    Private m_DtLibCobranca As Nullable(Of DateTime)

    Public Property QtdeMesAtivo() As Integer
        Get
            Return m_QtdeMesAtivo
        End Get
        Set(ByVal value As Integer)
            m_QtdeMesAtivo = value
        End Set
    End Property
    Private m_QtdeMesAtivo As Integer

    Public Property VlrTSMAtual() As Integer
        Get
            Return m_VlrTSMAtual
        End Get
        Set(ByVal value As Integer)
            m_VlrTSMAtual = value
        End Set
    End Property
    Private m_VlrTSMAtual As Integer

    Public Property ESenderEmailManut() As String
        Get
            Return m_ESenderEmailManut
        End Get
        Set(ByVal value As String)
            m_ESenderEmailManut = value
        End Set
    End Property
    Private m_ESenderEmailManut As String

    Public Property ESenderNameManut() As String
        Get
            Return m_ESenderNameManut
        End Get
        Set(ByVal value As String)
            m_ESenderNameManut = value
        End Set
    End Property
    Private m_ESenderNameManut As String

    Public Property ESenderEmailInst() As String
        Get
            Return m_ESenderEmailInst
        End Get
        Set(ByVal value As String)
            m_ESenderEmailInst = value
        End Set
    End Property
    Private m_ESenderEmailInst As String

    Public Property ESenderNameInst() As String
        Get
            Return m_ESenderNameInst
        End Get
        Set(ByVal value As String)
            m_ESenderNameInst = value
        End Set
    End Property
    Private m_ESenderNameInst As String

    Public Property UsrCad() As String
        Get
            Return m_UsrCad
        End Get
        Set(ByVal value As String)
            m_UsrCad = value
        End Set
    End Property
    Private m_UsrCad As String

    Public Property DtCad() As Nullable(Of DateTime)
        Get
            Return m_DtCad
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtCad = value
        End Set
    End Property
    Private m_DtCad As Nullable(Of DateTime)

    Public Property ClienteDe As String
    Public Property MonitoradoPor As String
    Public Property CGC_CPFBilling As String
    Public Property RazaoSocialBilling As String
    Public Property CodIntClie_IBS As String
    Public Property isVerisurePRO As String
End Class
