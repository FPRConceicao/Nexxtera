Imports Teleatlantic.TLS.Common

Public Class tNumeroGPRS : Inherits Retorno
    Public Property Numero As String
        Get
            Return m_Numero
        End Get
        Set(ByVal value As String)
            m_Numero = value
        End Set
    End Property
    Private m_Numero As String
    Public Property SIMCARD As String
        Get
            Return m_SIMCARD
        End Get
        Set(ByVal value As String)
            m_SIMCARD = value
        End Set
    End Property
    Private m_SIMCARD As String
    Public Property Operadora As String
        Get
            Return m_Operadora
        End Get
        Set(ByVal value As String)
            m_Operadora = value
        End Set
    End Property
    Private m_Operadora As String
    Public Property Status As String
        Get
            Return m_Status
        End Get
        Set(ByVal value As String)
            m_Status = value
        End Set
    End Property
    Private m_Status As String
    Public Property DtUltAlt As Nullable(Of DateTime)
        Get
            Return m_DtUltAlt
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtUltAlt = value
        End Set
    End Property
    Private m_DtUltAlt As Nullable(Of DateTime)
    Public Property UsrUltAlt As String
        Get
            Return m_UsrUltAlt
        End Get
        Set(ByVal value As String)
            m_UsrUltAlt = value
        End Set
    End Property
    Private m_UsrUltAlt As String
    Public Property IdFilial As String
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As String)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As String
    Public Property DtCad As Nullable(Of DateTime)
        Get
            Return m_DtCad
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtCad = value
        End Set
    End Property
    Private m_DtCad As Nullable(Of DateTime)
    Public Property UsrCad As String
        Get
            Return m_UsrCad
        End Get
        Set(ByVal value As String)
            m_UsrCad = value
        End Set
    End Property
    Private m_UsrCad As String

    Private _Tipo As String
    Public Property Tipo() As String
        Get
            Return _Tipo
        End Get
        Set(ByVal value As String)
            _Tipo = value
        End Set
    End Property

    Private _DescricaoFilial As String
    Public Property DescricaoFilial() As String
        Get
            Return _DescricaoFilial
        End Get
        Set(ByVal value As String)
            _DescricaoFilial = value
        End Set
    End Property

    Private _CodIntClie As String
    Public Property CodIntClie() As String
        Get
            Return _CodIntClie
        End Get
        Set(ByVal value As String)
            _CodIntClie = value
        End Set
    End Property

    Public Property EmpMonit() As String
    Public Property CodEmpresa As String
    Public Property CodInstal As String
    Public Property NomeEmpresa As String
    Public Property NomeInstal As String
End Class
