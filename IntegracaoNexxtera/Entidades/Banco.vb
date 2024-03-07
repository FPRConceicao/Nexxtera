Imports Teleatlantic.TLS.Common

Public Class Banco : Inherits Retorno


    Public Property CodBancoCodAgenciaNomeBanco() As String
        Get
            Return m_CodBancoCodAgenciaNomeBanco
        End Get
        Set(ByVal value As String)
            m_CodBancoCodAgenciaNomeBanco = value
        End Set
    End Property
    Private m_CodBancoCodAgenciaNomeBanco As String


    Public Property Status() As String
        Get
            Return m_Status
        End Get
        Set(ByVal value As String)
            m_Status = value
        End Set
    End Property
    Private m_Status As String


    Public Property CodBanco() As String
        Get
            Return m_CodBanco
        End Get
        Set(ByVal value As String)
            m_CodBanco = value
        End Set
    End Property
    Private m_CodBanco As String


    Public Property CodAgen() As String
        Get
            Return m_CodAgen
        End Get
        Set(ByVal value As String)
            m_CodAgen = value
        End Set
    End Property
    Private m_CodAgen As String


    Public Property Numcta() As String
        Get
            Return m_numcta
        End Get
        Set(ByVal value As String)
            m_numcta = value
        End Set
    End Property
    Private m_numcta As String


    Public Property NomeBanco() As String
        Get
            Return m_NomeBanco
        End Get
        Set(ByVal value As String)
            m_NomeBanco = value
        End Set
    End Property
    Private m_NomeBanco As String


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


    Public Property Cep() As String
        Get
            Return m_Cep
        End Get
        Set(ByVal value As String)
            m_Cep = value
        End Set
    End Property
    Private m_Cep As String


    Public Property DDD() As String
        Get
            Return m_DDD
        End Get
        Set(ByVal value As String)
            m_DDD = value
        End Set
    End Property
    Private m_DDD As String


    Public Property Fone1() As String
        Get
            Return m_Fone1
        End Get
        Set(ByVal value As String)
            m_Fone1 = value
        End Set
    End Property
    Private m_Fone1 As String


    Public Property Fone2() As String
        Get
            Return m_Fone2
        End Get
        Set(ByVal value As String)
            m_Fone2 = value
        End Set
    End Property
    Private m_Fone2 As String


    Public Property Fax() As String
        Get
            Return m_Fax
        End Get
        Set(ByVal value As String)
            m_Fax = value
        End Set
    End Property
    Private m_Fax As String


    Public Property Gerente1() As String
        Get
            Return m_Gerente1
        End Get
        Set(ByVal value As String)
            m_Gerente1 = value
        End Set
    End Property
    Private m_Gerente1 As String


    Public Property Gerente2() As String
        Get
            Return m_Gerente2
        End Get
        Set(ByVal value As String)
            m_Gerente2 = value
        End Set
    End Property
    Private m_Gerente2 As String


    Public Property Defasagem() As String
        Get
            Return m_Defasagem
        End Get
        Set(ByVal value As String)
            m_Defasagem = value
        End Set
    End Property
    Private m_Defasagem As String


    Public Property TaxaDia() As Double
        Get
            Return m_TaxaDia
        End Get
        Set(ByVal value As Double)
            m_TaxaDia = value
        End Set
    End Property
    Private m_TaxaDia As Double


    Public Property TaxaMulta() As Double
        Get
            Return m_TaxaMulta
        End Get
        Set(ByVal value As Double)
            m_TaxaMulta = value
        End Set
    End Property
    Private m_TaxaMulta As Double


    Public Property TaxaMes() As Double
        Get
            Return m_TaxaMes
        End Get
        Set(ByVal value As Double)
            m_TaxaMes = value
        End Set
    End Property
    Private m_TaxaMes As Double


    Public Property CodPRemes() As String
        Get
            Return m_CodPRemes
        End Get
        Set(ByVal value As String)
            m_CodPRemes = value
        End Set
    End Property
    Private m_CodPRemes As String


    Public Property SeqRemesUnico() As String
        Get
            Return m_SeqRemesUnico
        End Get
        Set(ByVal value As String)
            m_SeqRemesUnico = value
        End Set
    End Property
    Private m_SeqRemesUnico As String


    Public Property QtdeMinDiasTitDesc() As String
        Get
            Return m_QtdeMinDiasTitDesc
        End Get
        Set(ByVal value As String)
            m_QtdeMinDiasTitDesc = value
        End Set
    End Property
    Private m_QtdeMinDiasTitDesc As String


    Public Property QtdeMaxDiasTitDesc() As String
        Get
            Return m_QtdeMaxDiasTitDesc
        End Get
        Set(ByVal value As String)
            m_QtdeMaxDiasTitDesc = value
        End Set
    End Property
    Private m_QtdeMaxDiasTitDesc As String


    Public Property TipoConta() As String
        Get
            Return m_TipoConta
        End Get
        Set(ByVal value As String)
            m_TipoConta = value
        End Set
    End Property
    Private m_TipoConta As String


    Public Property NomeAgen() As String
        Get
            Return m_NomeAgen
        End Get
        Set(ByVal value As String)
            m_NomeAgen = value
        End Set
    End Property
    Private m_NomeAgen As String


    Public Property ContaCorrente() As ContaCorrente
        Get
            Return m_ContaCorrente
        End Get
        Set(ByVal value As ContaCorrente)
            m_ContaCorrente = value
        End Set
    End Property
    Private m_ContaCorrente As ContaCorrente

End Class
