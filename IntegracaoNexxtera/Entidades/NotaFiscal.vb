Imports Teleatlantic.TLS.Common

Public Class NotaFiscal : Inherits Retorno

    Public Property NumNF() As String
        Get
            Return m_NumNF
        End Get
        Set(ByVal value As String)
            m_NumNF = value
        End Set
    End Property
    Private m_NumNF As String

    Public Property SerieNF() As String
        Get
            Return m_SerieNF
        End Get
        Set(ByVal value As String)
            m_SerieNF = value
        End Set
    End Property
    Private m_SerieNF As String

    Public Property DtEmissao() As DateTime
        Get
            Return m_DtEmissao
        End Get
        Set(ByVal value As DateTime)
            m_DtEmissao = value
        End Set
    End Property
    Private m_DtEmissao As DateTime

    Public Property TipoNF() As String
        Get
            Return m_TipoNF
        End Get
        Set(ByVal value As String)
            m_TipoNF = value
        End Set
    End Property
    Private m_TipoNF As String

    Public Property CodClie() As String
        Get
            Return m_CodClie
        End Get
        Set(ByVal value As String)
            m_CodClie = value
        End Set
    End Property
    Private m_CodClie As String

    Public Property FisiJuri() As String
        Get
            Return m_FisiJuri
        End Get
        Set(ByVal value As String)
            m_FisiJuri = value
        End Set
    End Property
    Private m_FisiJuri As String

    Public Property CodEnde() As Integer
        Get
            Return m_CodEnde
        End Get
        Set(ByVal value As Integer)
            m_CodEnde = value
        End Set
    End Property
    Private m_CodEnde As Integer

    Public Property QtdeVol() As Double
        Get
            Return m_QtdeVol
        End Get
        Set(ByVal value As Double)
            m_QtdeVol = value
        End Set
    End Property
    Private m_QtdeVol As Double

    Public Property EspcVol() As String
        Get
            Return m_EspcVol
        End Get
        Set(ByVal value As String)
            m_EspcVol = value
        End Set
    End Property
    Private m_EspcVol As String

    Public Property MarcaVol() As String
        Get
            Return m_MarcaVol
        End Get
        Set(ByVal value As String)
            m_MarcaVol = value
        End Set
    End Property
    Private m_MarcaVol As String

    Public Property CFOP() As String
        Get
            Return m_CFOP
        End Get
        Set(ByVal value As String)
            m_CFOP = value
        End Set
    End Property
    Private m_CFOP As String

    Public Property NumVol() As String
        Get
            Return m_NumVol
        End Get
        Set(ByVal value As String)
            m_NumVol = value
        End Set
    End Property
    Private m_NumVol As String

    Public Property PesoLiq() As Double
        Get
            Return m_PesoLiq
        End Get
        Set(ByVal value As Double)
            m_PesoLiq = value
        End Set
    End Property
    Private m_PesoLiq As Double

    Public Property PesoBruto() As Double
        Get
            Return m_PesoBruto
        End Get
        Set(ByVal value As Double)
            m_PesoBruto = value
        End Set
    End Property
    Private m_PesoBruto As Double

    Public Property NumLote() As String
        Get
            Return m_NumLote
        End Get
        Set(ByVal value As String)
            m_NumLote = value
        End Set
    End Property
    Private m_NumLote As String

    Public Property MesAnoLote() As String
        Get
            Return m_MesAnoLote
        End Get
        Set(ByVal value As String)
            m_MesAnoLote = value
        End Set
    End Property
    Private m_MesAnoLote As String

    Public Property Impressa() As String
        Get
            Return m_Impressa
        End Get
        Set(ByVal value As String)
            m_Impressa = value
        End Set
    End Property
    Private m_Impressa As String

    Public Property NumLcto() As String
        Get
            Return m_NumLcto
        End Get
        Set(ByVal value As String)
            m_NumLcto = value
        End Set
    End Property
    Private m_NumLcto As String

    Public Property Situacao() As String
        Get
            Return m_Situacao
        End Get
        Set(ByVal value As String)
            m_Situacao = value
        End Set
    End Property
    Private m_Situacao As String

    Public Property Observacao() As String
        Get
            Return m_Observacao
        End Get
        Set(ByVal value As String)
            m_Observacao = value
        End Set
    End Property
    Private m_Observacao As String

    Public Property VlrNF() As Double
        Get
            Return m_VlrNF
        End Get
        Set(ByVal value As Double)
            m_VlrNF = value
        End Set
    End Property
    Private m_VlrNF As Double

    Public Property VlrICMS() As Double
        Get
            Return m_VlrICMS
        End Get
        Set(ByVal value As Double)
            m_VlrICMS = value
        End Set
    End Property
    Private m_VlrICMS As Double

    Public Property VlrIPI() As Double
        Get
            Return m_VlrIPI
        End Get
        Set(ByVal value As Double)
            m_VlrIPI = value
        End Set
    End Property
    Private m_VlrIPI As Double

    Public Property VlrISS() As Double
        Get
            Return m_VlrISS
        End Get
        Set(ByVal value As Double)
            m_VlrISS = value
        End Set
    End Property
    Private m_VlrISS As Double

    Public Property VlrIRRF() As Double
        Get
            Return m_VlrIRRF
        End Get
        Set(ByVal value As Double)
            m_VlrIRRF = value
        End Set
    End Property
    Private m_VlrIRRF As Double

    Public Property StatusContab() As String
        Get
            Return m_StatusContab
        End Get
        Set(ByVal value As String)
            m_StatusContab = value
        End Set
    End Property
    Private m_StatusContab As String

    Public Property DtVcto() As DateTime
        Get
            Return m_DtVcto
        End Get
        Set(ByVal value As DateTime)
            m_DtVcto = value
        End Set
    End Property
    Private m_DtVcto As DateTime

    Public Property CodPredFat() As String
        Get
            Return m_CodPredFat
        End Get
        Set(ByVal value As String)
            m_CodPredFat = value
        End Set
    End Property
    Private m_CodPredFat As String

    Public Property NumPedVda() As String
        Get
            Return m_NumPedVda
        End Get
        Set(ByVal value As String)
            m_NumPedVda = value
        End Set
    End Property
    Private m_NumPedVda As String

    Public Property Origem() As String
        Get
            Return m_Origem
        End Get
        Set(ByVal value As String)
            m_Origem = value
        End Set
    End Property
    Private m_Origem As String

    Public Property ProtManut() As String
        Get
            Return m_ProtManut
        End Get
        Set(ByVal value As String)
            m_ProtManut = value
        End Set
    End Property
    Private m_ProtManut As String

    Public Property DescProd() As Double
        Get
            Return m_DescProd
        End Get
        Set(ByVal value As Double)
            m_DescProd = value
        End Set
    End Property
    Private m_DescProd As Double

    Public Property VlrISSRet() As Double
        Get
            Return m_VlrISSRet
        End Get
        Set(ByVal value As Double)
            m_VlrISSRet = value
        End Set
    End Property
    Private m_VlrISSRet As Double

    Public Property VlrInstal() As Double
        Get
            Return m_VlrInstal
        End Get
        Set(ByVal value As Double)
            m_VlrInstal = value
        End Set
    End Property
    Private m_VlrInstal As Double

    Public Property VlrTxVisita() As Double
        Get
            Return m_VlrTxVisita
        End Get
        Set(ByVal value As Double)
            m_VlrTxVisita = value
        End Set
    End Property
    Private m_VlrTxVisita As Double

    Public Property VlrRetNovaCOFINS() As Double
        Get
            Return m_VlrRetNovaCOFINS
        End Get
        Set(ByVal value As Double)
            m_VlrRetNovaCOFINS = value
        End Set
    End Property
    Private m_VlrRetNovaCOFINS As Double

    Public Property VlrRetIR() As Double
        Get
            Return m_VlrRetIR
        End Get
        Set(ByVal value As Double)
            m_VlrRetIR = value
        End Set
    End Property
    Private m_VlrRetIR As Double

    Public Property VlrRetINSS() As Double
        Get
            Return m_VlrRetINSS
        End Get
        Set(ByVal value As Double)
            m_VlrRetINSS = value
        End Set
    End Property
    Private m_VlrRetINSS As Double

    Public Property StatusPref() As String
        Get
            Return m_StatusPref
        End Get
        Set(ByVal value As String)
            m_StatusPref = value
        End Set
    End Property
    Private m_StatusPref As String

    Public Property NumNFe() As String
        Get
            Return m_NumNFe
        End Get
        Set(ByVal value As String)
            m_NumNFe = value
        End Set
    End Property
    Private m_NumNFe As String

    Public Property CodVerNFe() As String
        Get
            Return m_CodVerNFe
        End Get
        Set(ByVal value As String)
            m_CodVerNFe = value
        End Set
    End Property
    Private m_CodVerNFe As String

    Public Property CodForn() As String
        Get
            Return m_CodForn
        End Get
        Set(ByVal value As String)
            m_CodForn = value
        End Set
    End Property
    Private m_CodForn As String

    Public Property StatusNFP() As String
        Get
            Return m_StatusNFP
        End Get
        Set(ByVal value As String)
            m_StatusNFP = value
        End Set
    End Property
    Private m_StatusNFP As String

    Public Property DtCanc() As DateTime
        Get
            Return m_DtCanc
        End Get
        Set(ByVal value As DateTime)
            m_DtCanc = value
        End Set
    End Property
    Private m_DtCanc As DateTime

    Public Property IdFilial() As String
        Get
            Return m_IdFilial
        End Get
        Set(ByVal value As String)
            m_IdFilial = value
        End Set
    End Property
    Private m_IdFilial As String

    Public Property StatusNFeEstadual() As String
        Get
            Return m_StatusNFeEstadual
        End Get
        Set(ByVal value As String)
            m_StatusNFeEstadual = value
        End Set
    End Property
    Private m_StatusNFeEstadual As String

    Public Property RazaoSocial() As String
        Get
            Return m_RazaoSocial
        End Get
        Set(ByVal value As String)
            m_RazaoSocial = value
        End Set
    End Property
    Private m_RazaoSocial As String
    Public Property Categoria() As String
        Get
            Return m_Categoria
        End Get
        Set(ByVal value As String)
            m_Categoria = value
        End Set
    End Property
    Private m_Categoria As String
    Public Property Municipio() As String
        Get
            Return m_Municipio
        End Get
        Set(ByVal value As String)
            m_Municipio = value
        End Set
    End Property
    Private m_Municipio As String

    Public Property Uf() As String
        Get
            Return m_Uf
        End Get
        Set(ByVal value As String)
            m_Uf = value
        End Set
    End Property
    Private m_Uf As String

    Public Property TipoNumNF() As String
        Get
            Return m_TipoNumNF
        End Get
        Set(ByVal value As String)
            m_TipoNumNF = value
        End Set
    End Property
    Private m_TipoNumNF As String

    'Marco 31/05/2023'
    Public Property TituloCategoria() As String
        Get
            Return m_TituloCategoria
        End Get
        Set(ByVal value As String)
            m_TituloCategoria = value
        End Set
    End Property
    Private m_TituloCategoria As String

    'Marco 05/06/2023'
    Public Property Filial() As String
        Get
            Return m_Filial
        End Get
        Set(ByVal value As String)
            m_Filial = value
        End Set
    End Property
    Private m_Filial As String

    'Marco 05/06/2023'
    Public Property Cnpj() As String
        Get
            Return m_Cnpj
        End Get
        Set(ByVal value As String)
            m_Cnpj = value
        End Set
    End Property
    Private m_Cnpj As String

    'VARIÁVEIS UTILIZADAS NA MANUTENÇÃO DE NF DE ENTRADA

    Public Property CodCCusto As String
        Get
            Return m_CodCCusto
        End Get
        Set(ByVal value As String)
            m_CodCCusto = value
        End Set
    End Property
    Private m_CodCCusto As String
    Public Property Historico As String
        Get
            Return m_Historico
        End Get
        Set(ByVal value As String)
            m_Historico = value
        End Set
    End Property
    Private m_Historico As String
    Public Property VlrTotal As Double
        Get
            Return m_VlrTotal
        End Get
        Set(ByVal value As Double)
            m_VlrTotal = value
        End Set
    End Property
    Private m_VlrTotal As Double
    Public Property VlrDesconto As Double
        Get
            Return m_VlrDesconto
        End Get
        Set(ByVal value As Double)
            m_VlrDesconto = value
        End Set
    End Property
    Private m_VlrDesconto As Double
    Public Property Parcelas As Integer
        Get
            Return m_Parcelas
        End Get
        Set(ByVal value As Integer)
            m_Parcelas = value
        End Set
    End Property
    Private m_Parcelas As Integer
    Public Property DtLcto As Nullable(Of DateTime)
        Get
            Return m_DtLcto
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtLcto = value
        End Set
    End Property
    Private m_DtLcto As Nullable(Of DateTime)
    Public Property VlrBaseICMS
        Get
            Return m_VlrBaseICMS
        End Get
        Set(ByVal value)
            m_VlrBaseICMS = value
        End Set
    End Property
    Private m_VlrBaseICMS As Double
    Public Property IdNFEntrada As String
        Get
            Return m_IdNFEntrada
        End Get
        Set(ByVal value As String)
            m_IdNFEntrada = value
        End Set
    End Property
    Private m_IdNFEntrada As String

    Public Property NotaEquip As String
        Get
            Return m_NotaEquip
        End Get
        Set(ByVal value As String)
            m_NotaEquip = value
        End Set
    End Property
    Private m_NotaEquip As String
    Public Property Usr As String
        Get
            Return m_Usr
        End Get
        Set(ByVal value As String)
            m_Usr = value
        End Set
    End Property
    Private m_Usr As String

    Public Property Valor As Double
        Get
            Return m_Valor
        End Get
        Set(ByVal value As Double)
            m_Valor = value
        End Set
    End Property
    Private m_Valor As Double
    Public Property Qtde As Integer
        Get
            Return m_Qtde
        End Get
        Set(ByVal value As Integer)
            m_Qtde = value
        End Set
    End Property
    Private m_Qtde As Integer
    Public Property VlrMercadoria As Double
        Get
            Return m_VlrMercadoria
        End Get
        Set(ByVal value As Double)
            m_VlrMercadoria = value
        End Set
    End Property
    Private m_VlrMercadoria As Double
    Public Property ItemNF As Integer
        Get
            Return m_ItemNF
        End Get
        Set(ByVal value As Integer)
            m_ItemNF = value
        End Set
    End Property
    Private m_ItemNF As Integer
    Public Property CodProd As String
        Get
            Return m_CodProd
        End Get
        Set(ByVal value As String)
            m_CodProd = value
        End Set
    End Property
    Private m_CodProd As String

   
    Public Property CodVerNFe2 As String
    Public Property NumNFe2 As String
    Public Property ObsNFAvulsaPrioritaria As String
    Public Property CondPgtoVda As String
End Class
