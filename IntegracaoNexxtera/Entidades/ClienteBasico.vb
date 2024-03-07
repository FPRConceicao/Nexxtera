''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class ClienteBasico : Inherits ClienteGeral 

    Public Property MicroEmpresa() As String
        Get
            Return m_MicroEmpresa
        End Get
        Set(ByVal value As String)
            m_MicroEmpresa = value
        End Set
    End Property
    Private m_MicroEmpresa As String

    Public Property InscEstad() As String
        Get
            Return m_InscEstad
        End Get
        Set(ByVal value As String)
            m_InscEstad = value
        End Set
    End Property
    Private m_InscEstad As String

    Public Property InscMunic() As String
        Get
            Return m_InscMunic
        End Get
        Set(ByVal value As String)
            m_InscMunic = value
        End Set
    End Property
    Private m_InscMunic As String

    Public Property CodOrigem() As String
        Get
            Return m_CodOrigem
        End Get
        Set(ByVal value As String)
            m_CodOrigem = value
        End Set
    End Property
    Private m_CodOrigem As String

    Public Property ObsClie() As String
        Get
            Return m_ObsClie
        End Get
        Set(ByVal value As String)
            m_ObsClie = value
        End Set
    End Property
    Private m_ObsClie As String

    Public Property IsPreCliente() As Integer
        Get
            Return m_IsPreCliente
        End Get
        Set(ByVal value As Integer)
            m_IsPreCliente = value
        End Set
    End Property
    Private m_IsPreCliente As Integer

    Public Property IsVip() As Integer
        Get
            Return m_IsVip
        End Get
        Set(ByVal value As Integer)
            m_IsVip = value
        End Set
    End Property
    Private m_IsVip As Integer

    Public Property IsAdContratual() As String
        Get
            Return m_IsAdContratual
        End Get
        Set(ByVal value As String)
            m_IsAdContratual = value
        End Set
    End Property
    Private m_IsAdContratual As String

    Public Property MsgPadrao() As String
        Get
            Return m_MsgPadrao
        End Get
        Set(ByVal value As String)
            m_MsgPadrao = value
        End Set
    End Property
    Private m_MsgPadrao As String

    Public Property Carteira() As String
        Get
            Return m_Carteira
        End Get
        Set(ByVal value As String)
            m_Carteira = value
        End Set
    End Property
    Private m_Carteira As String

    Public Property DtLibBloq() As DateTime
        Get
            Return m_DtLibBloq
        End Get
        Set(ByVal value As DateTime)
            m_DtLibBloq = value
        End Set
    End Property
    Private m_DtLibBloq As DateTime

    Public Property CodCCusto() As String
        Get
            Return m_CodCCusto
        End Get
        Set(ByVal value As String)
            m_CodCCusto = value
        End Set
    End Property
    Private m_CodCCusto As String


    Public Property FalConc() As String
        Get
            Return m_FalConc
        End Get
        Set(ByVal value As String)
            m_FalConc = value
        End Set
    End Property
    Private m_FalConc As String


    Public Property CodSeq() As String
        Get
            Return m_CodSeq
        End Get
        Set(ByVal value As String)
            m_CodSeq = value
        End Set
    End Property
    Private m_CodSeq As String

    Public Property LibCobr() As String
        Get
            Return m_LibCobr
        End Get
        Set(ByVal value As String)
            m_LibCobr = value
        End Set
    End Property
    Private m_LibCobr As String
    Public Property CatPed As String
        Get
            Return m_CatPed
        End Get
        Set(ByVal value As String)
            m_CatPed = value
        End Set
    End Property
    Private m_CatPed As String
    Public Property IsComodato As String
        Get
            Return m_IsComodato
        End Get
        Set(ByVal value As String)
            m_IsComodato = value
        End Set
    End Property
    Private m_IsComodato As String

    Private m_ImgVip As System.Drawing.Bitmap
    Public Property ImgVip() As System.Drawing.Bitmap
        Get
            Return m_ImgVip
        End Get
        Set(ByVal value As System.Drawing.Bitmap)
            m_ImgVip = value
        End Set
    End Property

    Public Property UnidadeMonitorada As String
        Get
            Return m_UnidadeMonitorada
        End Get
        Set(ByVal value As String)
            m_UnidadeMonitorada = value
        End Set
    End Property
    Private m_UnidadeMonitorada As String
    Public Property QtdeParticoes As Integer

    Public Property alianca As Alianca
End Class
