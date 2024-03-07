Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Entidades

Public Class FilaCobranca : Inherits Retorno
    Private _codFilaOrigem As String
    Private _descricao As String
    Private _dtCad As Nullable(Of DateTime)
    Private _isFilaManual As Boolean
    Private _codintclie As String
    Private _razaoSocial As String
    Private _tipoEstabelecimento As String
    Private _usrAtendente As String
    Private _diffDate As Integer
    Private _status As String
    Private _selecionadoPor As String
    Private _vlrInadimp As Boolean
    Private _qtdeTitInadimp As Integer
    Private _dtRetorno As Nullable(Of DateTime)
    Private _ordemFila As Integer
    Private _atendente As String

    Public Property CodFilaOrigem As String
        Get
            Return _codFilaOrigem
        End Get
        Set(value As String)
            _codFilaOrigem = value
        End Set
    End Property

    Public Property Descricao As String
        Get
            Return _descricao
        End Get
        Set(value As String)
            _descricao = value
        End Set
    End Property

    Public Property DtCad As Date?
        Get
            Return _dtCad
        End Get
        Set(value As Date?)
            _dtCad = value
        End Set
    End Property

    Public Property IsFilaManual As Boolean
        Get
            Return _isFilaManual
        End Get
        Set(value As Boolean)
            _isFilaManual = value
        End Set
    End Property

    Public Property Codintclie As String
        Get
            Return _codintclie
        End Get
        Set(value As String)
            _codintclie = value
        End Set
    End Property

    Public Property RazaoSocial As String
        Get
            Return _razaoSocial
        End Get
        Set(value As String)
            _razaoSocial = value
        End Set
    End Property

    Public Property TipoEstabelecimento As String
        Get
            Return _tipoEstabelecimento
        End Get
        Set(value As String)
            _tipoEstabelecimento = value
        End Set
    End Property

    Public Property UsrAtendente As String
        Get
            Return _usrAtendente
        End Get
        Set(value As String)
            _usrAtendente = value
        End Set
    End Property

    Public Property DiffDate As Integer
        Get
            Return _diffDate
        End Get
        Set(value As Integer)
            _diffDate = value
        End Set
    End Property

    Public Property Status As String
        Get
            Return _status
        End Get
        Set(value As String)
            _status = value
        End Set
    End Property

    Public Property SelecionadoPor As String
        Get
            Return _selecionadoPor
        End Get
        Set(value As String)
            _selecionadoPor = value
        End Set
    End Property

    Public Property VlrInadimp As Boolean
        Get
            Return _vlrInadimp
        End Get
        Set(value As Boolean)
            _vlrInadimp = value
        End Set
    End Property

    Public Property QtdeTitInadimp As Integer
        Get
            Return _qtdeTitInadimp
        End Get
        Set(value As Integer)
            _qtdeTitInadimp = value
        End Set
    End Property

    Public Property DtRetorno As Date?
        Get
            Return _dtRetorno
        End Get
        Set(value As Date?)
            _dtRetorno = value
        End Set
    End Property

    Public Property OrdemFila As Integer
        Get
            Return _ordemFila
        End Get
        Set(value As Integer)
            _ordemFila = value
        End Set
    End Property

    Public Property Atendente As String
        Get
            Return _atendente
        End Get
        Set(value As String)
            _atendente = value
        End Set
    End Property
End Class
