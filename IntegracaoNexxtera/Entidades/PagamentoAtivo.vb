Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Entidades

Public Class PagamentoAtivo : Inherits Retorno

    Private _NumParcela As Integer
    Public Property NumParcelas() As Integer
        Get
            Return _NumParcela
        End Get
        Set(ByVal value As Integer)
            _NumParcela = value
        End Set
    End Property

    Public Sub New()
        _ContRec = New ContaReceber()
        _ComplHistDeb = ""
        _ComplHistDeb = ""
        _ObsDeb = ""
    End Sub

    Private _ContRec As ContaReceber
    Public Property ContaReceber() As ContaReceber
        Get
            Return _ContRec
        End Get
        Set(ByVal value As ContaReceber)
            _ContRec = value
        End Set
    End Property

    Private _CtaCtblCred As String
    Public Property CtaCtblCred() As String
        Get
            Return _CtaCtblCred
        End Get
        Set(ByVal value As String)
            _CtaCtblCred = value
        End Set
    End Property

    Private _CtaContabilDeb As String
    Public Property ContaContabilDeb() As String
        Get
            Return _CtaContabilDeb
        End Get
        Set(ByVal value As String)
            _CtaContabilDeb = value
        End Set
    End Property

    Private _CodHistCred As String
    Public Property CodHistCred() As String
        Get
            Return _CodHistCred
        End Get
        Set(ByVal value As String)
            _CodHistCred = value
        End Set
    End Property

    Private _ComplHistCred As String
    Public Property ComplHistCred() As String
        Get
            Return _ComplHistCred
        End Get
        Set(ByVal value As String)
            _ComplHistCred = value
        End Set
    End Property

    Private _CodHistDeb As String
    Public Property CodHistDeb() As String
        Get
            Return _CodHistDeb
        End Get
        Set(ByVal value As String)
            _CodHistDeb = value
        End Set
    End Property

    Private _ComplHistDeb As String
    Public Property ComplHistDeb() As String
        Get
            Return IIf(_ComplHistDeb = "", "", _ComplHistDeb)
        End Get
        Set(ByVal value As String)
            _ComplHistDeb = value
        End Set
    End Property

    Private _VlrLcto As Double
    Public Property VlrLcto() As Double
        Get
            Return _VlrLcto
        End Get
        Set(ByVal value As Double)
            _VlrLcto = value
        End Set
    End Property

    Private _VlrLctoJuros As Double
    Public Property VlrLctoJuros() As Double
        Get
            Return _VlrLctoJuros
        End Get
        Set(ByVal value As Double)
            _VlrLctoJuros = value
        End Set
    End Property

    Private _CtaCtblCredJuros As Double
    Public Property CtaCtblCredJuros() As Double
        Get
            Return _CtaCtblCredJuros
        End Get
        Set(ByVal value As Double)
            _CtaCtblCredJuros = value
        End Set
    End Property

    Private _CodHistCredJuros As String
    Public Property CodHistCredJuros() As String
        Get
            Return _CodHistCredJuros
        End Get
        Set(ByVal value As String)
            _CodHistCredJuros = value
        End Set
    End Property


    Private _ObsJuros As String
    Public Property ObsJuros() As String
        Get
            Return _ObsJuros
        End Get
        Set(ByVal value As String)
            _ObsJuros = value
        End Set
    End Property

    Private _VlrLctoDesc As String
    Public Property VlrLctoDesc() As String
        Get
            Return _VlrLctoDesc
        End Get
        Set(ByVal value As String)
            _VlrLctoDesc = value
        End Set
    End Property

    Private _CtaCtblDebDesc As String
    Public Property CtaCtblDebDesc() As String
        Get
            Return _CtaCtblDebDesc
        End Get
        Set(ByVal value As String)
            _CtaCtblDebDesc = value
        End Set
    End Property

    Private _CodHistDebDesc As String
    Public Property CodHistDebDesc() As String
        Get
            Return _CodHistDebDesc
        End Get
        Set(ByVal value As String)
            _CodHistDebDesc = value
        End Set
    End Property


    Private _ObsDesc As String
    Public Property ObsDesc() As String
        Get
            Return _ObsDesc
        End Get
        Set(ByVal value As String)
            _ObsDesc = value
        End Set
    End Property

    Private _CtaCtblDebEncCA As String
    Public Property CtaCtblDebEncCA() As String
        Get
            Return _CtaCtblDebEncCA
        End Get
        Set(ByVal value As String)
            _CtaCtblDebEncCA = value
        End Set
    End Property

    Private _CodHistDebEncCA As String
    Public Property CodHistDebEncCA() As String
        Get
            Return _CodHistDebEncCA
        End Get
        Set(ByVal value As String)
            _CodHistDebEncCA = value
        End Set
    End Property

    Private _ObsEncCA As String
    Public Property ObsEncCA() As String
        Get
            Return _ObsEncCA
        End Get
        Set(ByVal value As String)
            _ObsEncCA = value
        End Set
    End Property

    Private _ObsCred As String
    Public Property ObsCred() As String
        Get
            Return _ObsCred
        End Get
        Set(ByVal value As String)
            _ObsCred = value
        End Set
    End Property

    Private _ObsDeb As String
    Public Property ObsDeb() As String
        Get
            Return _ObsDeb
        End Get
        Set(ByVal value As String)
            _ObsDeb = value
        End Set
    End Property





End Class
