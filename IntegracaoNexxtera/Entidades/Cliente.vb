''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class Cliente : Inherits ClienteBasico

    ''' <summary>
    ''' Descricao do STATUS do cliente ex: Ativo, Cancelado, Ativo - Bonificado, etc.
    ''' </summary>
    ''' <value>seta a descrição do tipo string</value>
    ''' <returns>Retorna a descrição do tipo string</returns>
    ''' <remarks></remarks>
    ''' 
    Private Property m_VlrSuperMotor As String
    Public Property VlrSuperMotor() As String
        Get
            Return m_VlrSuperMotor
        End Get
        Set(value As String)
            m_VlrSuperMotor = value
        End Set
    End Property

    Private Property m_RotaUF As String
    Public Property RotaUF() As String
        Get
            Return m_RotaUF
        End Get
        Set(value As String)
            m_RotaUF = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return m_Descricao
        End Get
        Set(ByVal value As String)
            m_Descricao = value
        End Set
    End Property
    Private m_Descricao As String

    Public Property Endereco() As Endereco
        Get
            Return m_Endereco
        End Get
        Set(ByVal value As Endereco)
            m_Endereco = value
        End Set
    End Property
    Private m_Endereco As Endereco

    Public Property Financeiro As ClienteFinanceiro
        Get
            Return m_Financeiro
        End Get
        Set(ByVal value As ClienteFinanceiro)
            m_Financeiro = value
        End Set
    End Property
    Private m_Financeiro As ClienteFinanceiro

    Public Property Contabilidade As ClienteContabilidade
        Get
            Return m_Contabilidade
        End Get
        Set(ByVal value As ClienteContabilidade)
            m_Contabilidade = value
        End Set
    End Property
    Private m_Contabilidade As ClienteContabilidade

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

    Public Property CGC_CPF() As String
        Get
            Return m_CGC_CPF
        End Get
        Set(ByVal value As String)
            m_CGC_CPF = value
        End Set
    End Property
    Private m_CGC_CPF As String

    Public Property RazaoSocial() As String
        Get
            Return m_RazaoSocial
        End Get
        Set(ByVal value As String)
            m_RazaoSocial = value
        End Set
    End Property
    Private m_RazaoSocial As String

    Public Property Monitoria As ClienteMonitoria
        Get
            Return m_Monitoria
        End Get
        Set(ByVal value As ClienteMonitoria)
            m_Monitoria = value
        End Set
    End Property
    Private m_Monitoria As ClienteMonitoria

    Public Property Venda As ClienteVenda
        Get
            Return m_Venda
        End Get
        Set(ByVal value As ClienteVenda)
            m_Venda = value
        End Set
    End Property
    Private m_Venda As ClienteVenda

    Public Property Ronda As ClienteRonda
        Get
            Return m_Ronda
        End Get
        Set(ByVal value As ClienteRonda)
            m_Ronda = value
        End Set
    End Property
    Private m_Ronda As ClienteRonda

    Public Property Manutencao() As ClienteManutencao
        Get
            Return m_Manutencao
        End Get
        Set(ByVal value As ClienteManutencao)
            m_Manutencao = value
        End Set
    End Property
    Private m_Manutencao As ClienteManutencao

    Public Property TelefoneEmail() As Email
        Get
            Return m_TelefoneEmail
        End Get
        Set(ByVal value As Email)
            m_TelefoneEmail = value
        End Set
    End Property
    Private m_TelefoneEmail As Email

    Public Property PedidoVenda() As PedidoVenda
        Get
            Return m_PedidoVenda
        End Get
        Set(ByVal value As PedidoVenda)
            m_PedidoVenda = value
        End Set
    End Property
    Private m_PedidoVenda As PedidoVenda

    Public Property Quest() As String
        Get
            Return m_Quest
        End Get
        Set(ByVal value As String)
            m_Quest = value
        End Set
    End Property
    Private m_Quest As String
    Public Property NumeroGprs As tNumeroGPRS
        Get
            Return m_NumeroGprs
        End Get
        Set(ByVal value As tNumeroGPRS)
            m_NumeroGprs = value
        End Set
    End Property
    Private m_NumeroGprs As tNumeroGPRS
    Private m_NumeroEmail As TelefoneEmail
    Public Property NumeroEmail() As TelefoneEmail
        Get
            Return m_NumeroEmail
        End Get
        Set(ByVal value As TelefoneEmail)
            m_NumeroEmail = value
        End Set
    End Property

    Private m_Parametro As Parametros
    Public Property Parametro() As Parametros
        Get
            Return m_Parametro
        End Get
        Set(ByVal value As Parametros)
            m_Parametro = value
        End Set
    End Property

    Public Property CodVerNfe As String
    Public Property isFinanciadoAntigo As String
    Public Property isErroDadosPgto As String

    Public Property DescTpPgto As String
    Public Property CodTransmissor As String
    Public Property DescChave As String

    Public Property Rota1 As String
    Public Property Rota2 As String
    Public Property RotaDesc1 As String
    Public Property RotaDesc2 As String
    Public Property NossoNumero As Integer
    Public Property NumTit As String
    Public Property SeqTit As String
	Public Property isClienteProjetos As String
	Public Property Carteira As String

    Public Property DtFimVigenciaContrato As DateTime
    Public Property DtSupensaoIda As DateTime
    Public Property DtvoltaSupensao As DateTime
    Public Property ContatoImprodutivo As String
    Public Property ObsMsgVeriOnlineProjetos As String
    Public Property CodPreCliente As String
    Public Property IsAutarquia As String
    Public Property IsInfluenciador As String
    Public Property DtUltimaAlteracao As DateTime
    Public Property protocoloOS As String
    Public Property urlPDFAdyen As String
    Public Property urlLinkAdyen As String

End Class
