Imports Teleatlantic.TLS.Common
''' <summary>
'''  Entidade de monitoriamento contendo as quantidades máxima de envio de Tele Online(ToQdeMax) e a quantidade enviadas mês(ToQdeEnv).
''' </summary>
''' <remarks>
''' 
''' 
''' </remarks>
Public Class ClienteMonitoria : Inherits Retorno


    Public Property DiaPgtoMonit() As String
        Get
            Return m_DiaPgtoMonit
        End Get
        Set(ByVal value As String)
            m_DiaPgtoMonit = value
        End Set
    End Property
    Private m_DiaPgtoMonit As String

    Public Property VlrTarBanc() As Double
        Get
            Return m_VlrTarBanc
        End Get
        Set(ByVal value As Double)
            m_VlrTarBanc = value
        End Set
    End Property
    Private m_VlrTarBanc As Double

    Public Property VlrMonitoria() As Double
        Get
            Return m_VlrMonitoria
        End Get
        Set(ByVal value As Double)
            m_VlrMonitoria = value
        End Set
    End Property
    Private m_VlrMonitoria As Double

    Public Property VlrSeguro() As Double
        Get
            Return m_VlrSeguro
        End Get
        Set(ByVal value As Double)
            m_VlrSeguro = value
        End Set
    End Property
    Private m_VlrSeguro As Double

    Public Property VlrRelat() As Double
        Get
            Return m_VlrRelat
        End Get
        Set(ByVal value As Double)
            m_VlrRelat = value
        End Set
    End Property
    Private m_VlrRelat As Double

    Public Property VlrSuperMotor() As Double
        Get
            Return m_VlrSuperMotor
        End Get
        Set(ByVal value As Double)
            m_VlrSuperMotor = value
        End Set
    End Property
    Private m_VlrSuperMotor As Double

    Public Property VlrLocacao() As Double
        Get
            Return m_VlrLocacao
        End Get
        Set(ByVal value As Double)
            m_VlrLocacao = value
        End Set
    End Property
    Private m_VlrLocacao As Double

    Public Property DtUltReajuste() As Nullable(Of DateTime)
        Get
            Return m_DtUltReajuste
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtUltReajuste = value
        End Set
    End Property
    Private m_DtUltReajuste As Nullable(Of DateTime)

    Public Property PrecReajustes() As Double
        Get
            Return m_PrecReajustes
        End Get
        Set(ByVal value As Double)
            m_PrecReajustes = value
        End Set
    End Property
    Private m_PrecReajustes As Double

    Public Property vlrMonitoriaAnt() As Double
        Get
            Return m_vlrMonitoriaAnt
        End Get
        Set(ByVal value As Double)
            m_vlrMonitoriaAnt = value
        End Set
    End Property
    Private m_vlrMonitoriaAnt As Double

    Public Property TeleEmerg() As Double
        Get
            Return m_TeleEmerg
        End Get
        Set(ByVal value As Double)
            m_TeleEmerg = value
        End Set
    End Property
    Private m_TeleEmerg As Double

    Public Property VlrTxtManMon() As Double
        Get
            Return m_VlrTxtManMon
        End Get
        Set(ByVal value As Double)
            m_VlrTxtManMon = value
        End Set
    End Property
    Private m_VlrTxtManMon As Double

    Public Property VlrRonda() As Double
        Get
            Return m_VlrRonda
        End Get
        Set(ByVal value As Double)
            m_VlrRonda = value
        End Set
    End Property
    Private m_VlrRonda As Double

    Public Property VlrCercaElet() As Double
        Get
            Return m_VlrCercaElet
        End Get
        Set(ByVal value As Double)
            m_VlrCercaElet = value
        End Set
    End Property
    Private m_VlrCercaElet As Double

    Public Property VlrRadio() As Double
        Get
            Return m_VlrRadio
        End Get
        Set(ByVal value As Double)
            m_VlrRadio = value
        End Set
    End Property
    Private m_VlrRadio As Double

    Public Property VlrAviso() As Double
        Get
            Return m_VlrAviso
        End Get
        Set(ByVal value As Double)
            m_VlrAviso = value
        End Set
    End Property
    Private m_VlrAviso As Double

    Public Property VlrArme() As Double
        Get
            Return m_VlrArme
        End Get
        Set(ByVal value As Double)
            m_VlrArme = value
        End Set
    End Property
    Private m_VlrArme As Double

    Public Property VlrBkpGPRS() As Double
        Get
            Return m_VlrBkpGPRS
        End Get
        Set(ByVal value As Double)
            m_VlrBkpGPRS = value
        End Set
    End Property
    Private m_VlrBkpGPRS As Double

    Public Property VlrManut() As Double
        Get
            Return m_VlrManut
        End Get
        Set(ByVal value As Double)
            m_VlrManut = value
        End Set
    End Property
    Private m_VlrManut As Double

    Public Property TipoManut() As String
        Get
            Return m_TipoManut
        End Get
        Set(ByVal value As String)
            m_TipoManut = value
        End Set
    End Property
    Private m_TipoManut As String

    Public Property VlrMonitAdic() As Double
        Get
            Return m_VlrMonitAdic
        End Get
        Set(ByVal value As Double)
            m_VlrMonitAdic = value
        End Set
    End Property
    Private m_VlrMonitAdic As Double

    Public Property VlrTeleVision() As Double
        Get
            Return m_VlrTeleVision
        End Get
        Set(ByVal value As Double)
            m_VlrTeleVision = value
        End Set
    End Property
    Private m_VlrTeleVision As Double

    Public Property Botom() As String
        Get
            Return m_Botom
        End Get
        Set(ByVal value As String)
            m_Botom = value
        End Set
    End Property
    Private m_Botom As String

    Public Property Radio() As String
        Get
            Return m_Radio
        End Get
        Set(ByVal value As String)
            m_Radio = value
        End Set
    End Property
    Private m_Radio As String

    Public Property Gprs() As String
        Get
            Return m_Gprs
        End Get
        Set(ByVal value As String)
            m_Gprs = value
        End Set
    End Property
    Private m_Gprs As String

    Public Property VlrRepasseTeleEmerg() As Double
        Get
            Return m_VlrRepasseTeleEmerg
        End Get
        Set(ByVal value As Double)
            m_VlrRepasseTeleEmerg = value
        End Set
    End Property
    Private m_VlrRepasseTeleEmerg As Double

    Public Property TipoService() As String
        Get
            Return m_TipoService
        End Get
        Set(ByVal value As String)
            m_TipoService = value
        End Set
    End Property
    Private m_TipoService As String

    Public Property TipoTeleVideo() As String
        Get
            Return m_TipoTeleVideo
        End Get
        Set(ByVal value As String)
            m_TipoTeleVideo = value
        End Set
    End Property
    Private m_TipoTeleVideo As String

    Public Property TipoEmerg() As String
        Get
            Return m_TipoEmerg
        End Get
        Set(ByVal value As String)
            m_TipoEmerg = value
        End Set
    End Property
    Private m_TipoEmerg As String


    Public Property ToQtdeMax() As Integer
        Get
            Return m_ToQtdeMax
        End Get
        Set(ByVal value As Integer)
            m_ToQtdeMax = value
        End Set
    End Property
    Private m_ToQtdeMax As Integer

    Public Property ToQtdeEnv() As Integer
        Get
            Return m_ToQtdeEnv
        End Get
        Set(ByVal value As Integer)
            m_ToQtdeEnv = value
        End Set
    End Property
    Private m_ToQtdeEnv As Integer

    Public Property TotalServico() As Double
        Get
            Return m_TotalServico
        End Get
        Set(ByVal value As Double)
            m_TotalServico = value
        End Set
    End Property
    Private m_TotalServico As Double

    Public Property QtdeCamTV() As String
        Get
            Return m_QtdeCamTV
        End Get
        Set(ByVal value As String)
            m_QtdeCamTV = value
        End Set
    End Property
    Private m_QtdeCamTV As String

    Public Property Mensalidade() As Double
        Get
            Return m_Mensalidade
        End Get
        Set(ByVal value As Double)
            m_Mensalidade = value
        End Set
    End Property
    Private m_Mensalidade As Double

    Public Property VlrManutencaoCFTV() As Double
        Get
            Return m_VlrManutencaoCFTV
        End Get
        Set(ByVal value As Double)
            m_VlrManutencaoCFTV = value
        End Set
    End Property
    Private m_VlrManutencaoCFTV As Double

    Public Property TipoManutencaoCFTV() As String
        Get
            Return m_TipoManutencaoCFTV
        End Get
        Set(ByVal value As String)
            m_TipoManutencaoCFTV = value
        End Set
    End Property
    Private m_TipoManutencaoCFTV As String

    Public Property QtdeCamManutencaoCFTV() As Integer
        Get
            Return m_QtdeCamManutencaoCFTV
        End Get
        Set(ByVal value As Integer)
            m_QtdeCamManutencaoCFTV = value
        End Set
    End Property
    Private m_QtdeCamManutencaoCFTV As Integer

    Private _VlrTeleGPS As Double
    Public Property VlrTeleGPS() As Double
        Get
            Return _VlrTeleGPS
        End Get
        Set(ByVal value As Double)
            _VlrTeleGPS = value
        End Set
    End Property

    Public Property QtdeMesAtivo As Integer
    Public Property QtdeVeiculos As Integer
    Public Property Pago As Integer
    Public Property TipoVenda As String

    Public Property VlrTeleViewFaturado As Double
    Public Property VlrTeleGPSFaturado As Double
    Public Property VlrTeleGPSPago As Double
    Public Property VlrTeleViewPago As Double
    Public Property VlrPagoFornecedorTV As Double
    Public Property VlrPagoFornecedorTeleGPS As Double

    Public Property StatusServico As String

    Public Property vlrMargemAlianca As Integer
    Public Property vlrRepasseAlianca As Integer
    Public Property isFinanciamento As Integer
    Public Property vlrParcelaFinanciamento As Double
    Public Property qtdeParcelasFinanciamento As Integer
    Public Property parcelaGeradaFinanciamento As Integer
End Class

