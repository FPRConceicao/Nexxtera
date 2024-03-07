Imports Teleatlantic.TLS.Common

Public Class AreaFilial : Inherits Retorno
    Public Property IdFilial() As Long
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As Long)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As Long


    Public Property Cidade() As String
        Get
            Return m_Cidade
        End Get
        Set(ByVal value As String)
            m_Cidade = value
        End Set
    End Property
    Private m_Cidade As String


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


    Public Property CepIni() As String
        Get
            Return m_CepIni
        End Get
        Set(ByVal value As String)
            m_CepIni = value
        End Set
    End Property
    Private m_CepIni As String

    Public Property CepFim() As String
        Get
            Return m_CepFim
        End Get
        Set(ByVal value As String)
            m_CepFim = value
        End Set
    End Property
    Private m_CepFim As String


    Public Property PercRetISSMonitoria() As Double
        Get
            Return m_PercRetISSMonitoria
        End Get
        Set(ByVal value As Double)
            m_PercRetISSMonitoria = value
        End Set
    End Property
    Private m_PercRetISSMonitoria As Double

    Public Property PercRetISSInstalacao() As Double
        Get
            Return m_PercRetISSInstalacao
        End Get
        Set(ByVal value As Double)
            m_PercRetISSInstalacao = value
        End Set
    End Property
    Private m_PercRetISSInstalacao As Double

    Public Property PercRetISSManutencao() As Double
        Get
            Return m_PercRetISSManutencao
        End Get
        Set(ByVal value As Double)
            m_PercRetISSManutencao = value
        End Set
    End Property
    Private m_PercRetISSManutencao As Double

    ''' <summary>
    ''' As propriedades iniciadas com str são utilizadas no Cadastro de Cidades para Filiais para corrigir o problema de conversão do tipo double para string
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Property strPercRetISSManutencao As String
        Get
            Return m_strPercRetISSManutencao
        End Get
        Set(ByVal value As String)
            m_strPercRetISSManutencao = value
        End Set
    End Property
    Private m_strPercRetISSManutencao As String

    Public Property strPercRetISSInstalacao As String
        Get
            Return m_strPercRetISSInstalacao
        End Get
        Set(ByVal value As String)
            m_strPercRetISSInstalacao = value
        End Set
    End Property
    Private m_strPercRetISSInstalacao As String

    Public Property strPercRetISSMonitoria As String
        Get
            Return m_strPercRetISSMonitoria
        End Get
        Set(ByVal value As String)
            m_strPercRetISSMonitoria = value
        End Set
    End Property
    Private m_strPercRetISSMonitoria As String
    Public Property PesoScorePosVendas As Integer
    Public Property IsConurbada As Integer

    Public Property TribComISSSemPCCSemIR As String
    Public Property TribComISSComPCCComIR As String
    Public Property TribComISSComPCCSemIR As String
    Public Property TribComISSSemPCCComIR As String
    Public Property TribSemISSSemPCCSemIR As String
    Public Property TribSemISSComPCCComIR As String
    Public Property TribSemISSComPCCSemIR As String
    Public Property TribSemISSSemPCCComIR As String

    Public Property TribComISSSemPCCSemIRProj As String
    Public Property TribComISSComPCCComIRProj As String
    Public Property TribComISSComPCCSemIRProj As String
    Public Property TribComISSSemPCCComIRProj As String
    Public Property TribSemISSSemPCCSemIRProj As String
    Public Property TribSemISSComPCCComIRProj As String
    Public Property TribSemISSComPCCSemIRProj As String
    Public Property TribSemISSSemPCCComIRProj As String
    Public Property UF As String
    Public Property DescUF As String
End Class
