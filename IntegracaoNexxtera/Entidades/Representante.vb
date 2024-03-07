Imports Teleatlantic.TLS.Common
''' <summary>
''' Entidade de numero de telefone ou fax.
''' </summary>
''' <remarks>
''' 
''' Data Criação:     12/04/2011
''' Auttor:           Wolney Alexandre Fernandes
''' 
''' </remarks>
Public Class Representante : Inherits Retorno

    Public Property MatRepr As String

    Public Property MinDuracaoVisita As Integer
    Public Property MinIntervaloVisita As Integer
    Public Property DealerProgram As Integer
    Public Property EmailGerenteAreaVenda As String
    Public Property NomeGerenteAreaVenda As String
    Public Property TempoMonitoria As Integer

    Public Property CodRepr() As String
        Get
            Return m_CodRepr
        End Get
        Set(ByVal value As String)
            m_CodRepr = value
        End Set
    End Property
    Private m_CodRepr As String

    Public Property Nome() As String
        Get
            Return m_Nome
        End Get
        Set(ByVal value As String)
            m_Nome = value
        End Set
    End Property
    Private m_Nome As String

    Public Property NomeRedu() As String
        Get
            Return m_NomeRedu
        End Get
        Set(ByVal value As String)
            m_NomeRedu = value
        End Set
    End Property
    Private m_NomeRedu As String

    Public Property Situacao() As String
        Get
            Return m_Situacao
        End Get
        Set(ByVal value As String)
            m_Situacao = value
        End Set
    End Property
    Private m_Situacao As String

    Public Property DtAlt() As DateTime
        Get
            Return m_DtAlt
        End Get
        Set(ByVal value As DateTime)
            m_DtAlt = value
        End Set
    End Property
    Private m_DtAlt As DateTime

    Public Property UsrAlt() As String
        Get
            Return m_UsrAlt
        End Get
        Set(ByVal value As String)
            m_UsrAlt = value
        End Set
    End Property
    Private m_UsrAlt As String

    Public Property CodTabComiss() As String
        Get
            Return m_CodTabComiss
        End Get
        Set(ByVal value As String)
            m_CodTabComiss = value
        End Set
    End Property
    Private m_CodTabComiss As String

    Public Property CodDepto() As String
        Get
            Return m_CodDepto
        End Get
        Set(ByVal value As String)
            m_CodDepto = value
        End Set
    End Property
    Private m_CodDepto As String

    Public Property Email() As String
        Get
            Return m_Email
        End Get
        Set(ByVal value As String)
            m_Email = value
        End Set
    End Property
    Private m_Email As String

    Public Property EmailSup() As String
        Get
            Return m_EmailSup
        End Get
        Set(ByVal value As String)
            m_EmailSup = value
        End Set
    End Property
    Private m_EmailSup As String

    Public Property CNPJ() As String
        Get
            Return m_CNPJ
        End Get
        Set(ByVal value As String)
            m_CNPJ = value
        End Set
    End Property
    Private m_CNPJ As String

    Public Property Telefone() As String
        Get
            Return m_Telefone
        End Get
        Set(ByVal value As String)
            m_Telefone = value
        End Set
    End Property
    Private m_Telefone As String

    Public Property Cep() As String
        Get
            Return m_Cep
        End Get
        Set(ByVal value As String)
            m_Cep = value
        End Set
    End Property
    Private m_Cep As String

    Public Property Endereco() As String
        Get
            Return m_Endereco
        End Get
        Set(ByVal value As String)
            m_Endereco = value
        End Set
    End Property
    Private m_Endereco As String

    Public Property Bairro() As String
        Get
            Return m_Bairro
        End Get
        Set(ByVal value As String)
            m_Bairro = value
        End Set
    End Property
    Private m_Bairro As String

    Public Property Cidade() As String
        Get
            Return m_Cidade
        End Get
        Set(ByVal value As String)
            m_Cidade = value
        End Set
    End Property
    Private m_Cidade As String

    Public Property UF() As String
        Get
            Return m_UF
        End Get
        Set(ByVal value As String)
            m_UF = value
        End Set
    End Property
    Private m_UF As String

    Public Property Usr() As String
        Get
            Return m_Usr
        End Get
        Set(ByVal value As String)
            m_Usr = value
        End Set
    End Property
    Private m_Usr As String

    Public Property Contato() As String
        Get
            Return m_Contato
        End Get
        Set(ByVal value As String)
            m_Contato = value
        End Set
    End Property
    Private m_Contato As String

    Public Property IdFilial() As String
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As String)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As String

    Public Property IsCadPromo() As String
        Get
            Return m_IsCadPromo
        End Get
        Set(ByVal value As String)
            m_IsCadPromo = value
        End Set
    End Property
    Private m_IsCadPromo As String

    Public Property Canal() As String
        Get
            Return m_Canal
        End Get
        Set(ByVal value As String)
            m_Canal = value
        End Set
    End Property
    Private m_Canal As String

    Public Property CodAreaVenda() As String
        Get
            Return m_CodAreaVenda
        End Get
        Set(ByVal value As String)
            m_CodAreaVenda = value
        End Set
    End Property
    Private m_CodAreaVenda As String

    Public Property CodVar() As String
        Get
            Return m_CodVar
        End Get
        Set(ByVal value As String)
            m_CodVar = value
        End Set
    End Property
    Private m_CodVar As String

    Public Property DtLibVda() As DateTime
        Get
            Return m_DtLibVda
        End Get
        Set(ByVal value As DateTime)
            m_DtLibVda = value
        End Set
    End Property
    Private m_DtLibVda As DateTime

    Public Property DescricaoDepto() As String
        Get
            Return m_DescricaoDepto
        End Get
        Set(ByVal value As String)
            m_DescricaoDepto = value
        End Set
    End Property
    Private m_DescricaoDepto As String

    Public Property DescricaoFilial() As String
        Get
            Return m_DescricaoFilial
        End Get
        Set(ByVal value As String)
            m_DescricaoFilial = value
        End Set
    End Property
    Private m_DescricaoFilial As String

    Public Property DescricaoTabComissao() As String
        Get
            Return m_DescricaoTabComissao
        End Get
        Set(ByVal value As String)
            m_DescricaoTabComissao = value
        End Set
    End Property
    Private m_DescricaoTabComissao As String

    Public Property PermiteCFTV() As String
        Get
            Return m_PermiteCFTV
        End Get
        Set(ByVal value As String)
            m_PermiteCFTV = value
        End Set
    End Property
    Private m_PermiteCFTV As String
    Public Property DDDTel As String
        Get
            Return m_DDDTel
        End Get
        Set(ByVal value As String)
            m_DDDTel = value
        End Set
    End Property
    Private m_DDDTel As String

    Public Property LimiteIndicacoes As Integer
    Public Property LimiteIndicacoesAumPto As Integer
    Public Property EmpresaContratante As String
    Public Property UsrChefe As String
    Public Property DtDesligamento As Nullable(Of DateTime)
    Public Property Usuario As String
    Public Property QtdeDemosPermitidas As Integer
    Public Property isLiderVenda As String
    Public Property CodEmpContratada As String

    Public Property idClassificacaoEspecialista As Integer
    Public Property idCargoEspecilaista As Integer
    Public Property NomeRepr As String
    Public Property isAssociadoRaking As Boolean
    Public Property id_oneSignal As String
    Public Property nomeLider As String

End Class
