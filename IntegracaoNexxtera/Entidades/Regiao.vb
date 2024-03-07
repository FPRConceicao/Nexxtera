Imports Teleatlantic.TLS.Common

Public Class Regiao : Inherits Retorno

    Private m_RotaUF As String
    Public Property RotaUF() As String
        Get
            Return m_RotaUF
        End Get
        Set(value As String)
            m_RotaUF = value
        End Set
    End Property

    Public Property CodRegiao() As String
        Get
            Return m_CodRegiao
        End Get
        Set(ByVal value As String)
            m_CodRegiao = value
        End Set
    End Property
    Private m_CodRegiao As String


    Public Property NomeRegiao() As String
        Get
            Return m_NomeRegiao
        End Get
        Set(ByVal value As String)
            m_NomeRegiao = value
        End Set
    End Property
    Private m_NomeRegiao As String


    Public Property Status() As String
        Get
            Return m_Status
        End Get
        Set(ByVal value As String)
            m_Status = value
        End Set
    End Property
    Private m_Status As String


    Public Property Filial() As String
        Get
            Return m_Filial
        End Get
        Set(ByVal value As String)
            m_Filial = value
        End Set
    End Property
    Private m_Filial As String


    Public Property VlrPgCliente() As Double
        Get
            Return m_VlrPgCliente
        End Get
        Set(ByVal value As Double)
            m_VlrPgCliente = value
        End Set
    End Property
    Private m_VlrPgCliente As Double


    Public Property VlrPgClienteSM2() As Double
        Get
            Return m_VlrPgClienteSM2
        End Get
        Set(ByVal value As Double)
            m_VlrPgClienteSM2 = value
        End Set
    End Property
    Private m_VlrPgClienteSM2 As Double


    Public Property VlrPgClienteSM3() As Double
        Get
            Return m_VlrPgClienteSM3
        End Get
        Set(ByVal value As Double)
            m_VlrPgClienteSM3 = value
        End Set
    End Property
    Private m_VlrPgClienteSM3 As Double


    Public Property VlrPgClienteSM4() As Double
        Get
            Return m_VlrPgClienteSM4
        End Get
        Set(ByVal value As Double)
            m_VlrPgClienteSM4 = value
        End Set
    End Property
    Private m_VlrPgClienteSM4 As Double

    Public Property DescricaoFilial() As String
        Get
            Return m_DescricaoFilial
        End Get
        Set(ByVal value As String)
            m_DescricaoFilial = value
        End Set
    End Property
    Private m_DescricaoFilial As String

    Public Property DtCad() As Nullable(Of DateTime)
        Get
            Return m_DtCad
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtCad = value
        End Set
    End Property
    Private m_DtCad As Nullable(Of DateTime)

    Public Property UsrCad() As String
        Get
            Return m_UsrCad
        End Get
        Set(ByVal value As String)
            m_UsrCad = value
        End Set
    End Property
    Private m_UsrCad As String

    Public Property AtendeSomenteSM1() As String
        Get
            Return m_AtendeSomenteSM1
        End Get
        Set(ByVal value As String)
            m_AtendeSomenteSM1 = value
        End Set
    End Property
    Private m_AtendeSomenteSM1 As String

    Public Property VlrPgPorCamTeleVideo() As Double

        Get
            Return m_VlrPgPorCamTeleVideo
        End Get
        Set(ByVal value As Double)
            m_VlrPgPorCamTeleVideo = value
        End Set
    End Property
    Private m_VlrPgPorCamTeleVideo As Double

    Public Property PertenceAreaCobertura() As String
        Get
            Return m_PertenceAreaCobertura
        End Get
        Set(ByVal value As String)
            m_PertenceAreaCobertura = value
        End Set
    End Property
    Private m_PertenceAreaCobertura As String

    Public Property CodDescrRegiao() As String
        Get
            Return m_CodDescrRegiao
        End Get
        Set(ByVal value As String)
            m_CodDescrRegiao = value
        End Set
    End Property
    Private m_CodDescrRegiao As String

    Private m_MicroArea As String
    Public Property MicroArea() As String
        Get
            Return m_MicroArea
        End Get
        Set(ByVal value As String)
            m_MicroArea = value
        End Set
    End Property

    Private m_VlrPgPorCamManutCFTV As Double
    Public Property VlrPgPorCamManutCFTV() As Double
        Get
            Return m_VlrPgPorCamManutCFTV
        End Get
        Set(ByVal value As Double)
            m_VlrPgPorCamManutCFTV = value
        End Set
    End Property

    Private m_UF As String
    Public Property UF() As String
        Get
            Return m_UF
        End Get
        Set(value As String)
            m_UF = value
        End Set
    End Property

    Private m_DescEstado As String
    Public Property DescEstado As String
        Get
            Return m_DescEstado
        End Get
        Set(value As String)
            m_DescEstado = value
        End Set
    End Property





End Class
