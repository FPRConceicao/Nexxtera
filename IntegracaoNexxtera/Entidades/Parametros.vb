Imports Teleatlantic.TLS.Common

Public Class Parametros : Inherits Retorno


    Public Property EAprovSolicHE() As String
        Get
            Return m_EAprovSolicHE
        End Get
        Set(ByVal value As String)
            m_EAprovSolicHE = value
        End Set
    End Property
    Private m_EAprovSolicHE As String

    Public Property ESenderNameManut() As String
        Get
            Return m_ESenderNameManut
        End Get
        Set(ByVal value As String)
            m_ESenderNameManut = value
        End Set
    End Property
    Private m_ESenderNameManut As String

    Public Property ESenderEmailManut() As String
        Get
            Return m_ESenderEmailManut
        End Get
        Set(ByVal value As String)
            m_ESenderEmailManut = value
        End Set
    End Property
    Private m_ESenderEmailManut As String

    Public Property ESubjectManut() As String
        Get
            Return m_ESubjectManut
        End Get
        Set(ByVal value As String)
            m_ESubjectManut = value
        End Set
    End Property
    Private m_ESubjectManut As String

    Public Property PermAumPtoInadimp() As String
        Get
            Return m_PermAumPtoInadimp
        End Get
        Set(ByVal value As String)
            m_PermAumPtoInadimp = value
        End Set
    End Property
    Private m_PermAumPtoInadimp As String


    Public Property PermMRSInadimp() As String
        Get
            Return m_PermMRSInadimp
        End Get
        Set(ByVal value As String)
            m_PermMRSInadimp = value
        End Set
    End Property
    Private m_PermMRSInadimp As String


    Public Property SeqDI() As String
        Get
            Return m_SeqDI
        End Get
        Set(ByVal value As String)
            m_SeqDI = value
        End Set
    End Property
    Private m_SeqDI As String


    Public Property VenUNSC() As String
        Get
            Return m_VenUNSC
        End Get
        Set(ByVal value As String)
            m_VenUNSC = value
        End Set
    End Property
    Private m_VenUNSC As String


    Public Property ESenderName() As String
        Get
            Return m_ESenderName
        End Get
        Set(ByVal value As String)
            m_ESenderName = value
        End Set
    End Property
    Private m_ESenderName As String


    Public Property EmailConfPath() As String
        Get
            Return m_EmailConfPath
        End Get
        Set(ByVal value As String)
            m_EmailConfPath = value
        End Set
    End Property
    Private m_EmailConfPath As String


    Public Property ESMTPServer() As String
        Get
            Return m_ESMTPServer
        End Get
        Set(ByVal value As String)
            m_ESMTPServer = value
        End Set
    End Property
    Private m_ESMTPServer As String


    Public Property ContratoPath() As String
        Get
            Return m_ContratoPath
        End Get
        Set(ByVal value As String)
            m_ContratoPath = value
        End Set
    End Property
    Private m_ContratoPath As String

    Public Property URLClaro() As String
        Get
            Return m_URLClaro
        End Get
        Set(ByVal value As String)
            m_URLClaro = value
        End Set
    End Property
    Private m_URLClaro As String

    Public Property ProfileClaro() As String
        Get
            Return m_ProfileClaro
        End Get
        Set(ByVal value As String)
            m_ProfileClaro = value
        End Set
    End Property
    Private m_ProfileClaro As String

    Public Property PwdClaro() As String
        Get
            Return m_PwdClaro
        End Get
        Set(ByVal value As String)
            m_PwdClaro = value
        End Set
    End Property
    Private m_PwdClaro As String

    Public Property ModeClaro() As String
        Get
            Return m_ModeClaro
        End Get
        Set(ByVal value As String)
            m_ModeClaro = value
        End Set
    End Property
    Private m_ModeClaro As String

    Public Property Mode_MB() As String
        Get
            Return m_Mode_MB
        End Get
        Set(ByVal value As String)
            m_Mode_MB = value
        End Set
    End Property
    Private m_Mode_MB As String

    Public Property User_MB() As String
        Get
            Return m_User_MB
        End Get
        Set(ByVal value As String)
            m_User_MB = value
        End Set
    End Property
    Private m_User_MB As String

    Public Property Credential_MB() As String
        Get
            Return m_Credential_MB
        End Get
        Set(ByVal value As String)
            m_Credential_MB = value
        End Set
    End Property
    Private m_Credential_MB As String

    Public Property URL_MB() As String
        Get
            Return m_URL_MB
        End Get
        Set(ByVal value As String)
            m_URL_MB = value
        End Set
    End Property
    Private m_URL_MB As String

    Public Property ESenderEmailInst() As String
        Get
            Return m_ESenderEmailInst
        End Get
        Set(ByVal value As String)
            m_ESenderEmailInst = value
        End Set
    End Property
    Private m_ESenderEmailInst As String

    Public Property ESenderNameInst() As String
        Get
            Return m_ESenderNameInst
        End Get
        Set(ByVal value As String)
            m_ESenderNameInst = value
        End Set
    End Property
    Private m_ESenderNameInst As String

    Public Property UF() As String
        Get
            Return m_UF
        End Get
        Set(ByVal value As String)
            m_UF = value
        End Set
    End Property
    Private m_UF As String

    Public Property UpdatePath() As String
        Get
            Return m_UpdatePath
        End Get
        Set(ByVal value As String)
            m_UpdatePath = value
        End Set
    End Property
    Private m_UpdatePath As String

    Public Property ESenderNameInd() As String
        Get
            Return m_ESenderNameInd
        End Get
        Set(ByVal value As String)
            m_ESenderNameInd = value
        End Set
    End Property
    Private m_ESenderNameInd As String

    Public Property ESenderEmailInd() As String
        Get
            Return m_ESenderEmailInd
        End Get
        Set(ByVal value As String)
            m_ESenderEmailInd = value
        End Set
    End Property
    Private m_ESenderEmailInd As String

    Public Property ESubjectInd() As String
        Get
            Return m_ESubjectInd
        End Get
        Set(ByVal value As String)
            m_ESubjectInd = value
        End Set
    End Property
    Private m_ESubjectInd As String

    Public Property RazaoSocial() As String
        Get
            Return m_RazaoSocial
        End Get
        Set(ByVal value As String)
            m_RazaoSocial = value
        End Set
    End Property
    Private m_RazaoSocial As String

    Public Property Fone() As String
        Get
            Return m_Fone
        End Get
        Set(ByVal value As String)
            m_Fone = value
        End Set
    End Property
    Private m_Fone As String

    Public Property Fax() As String
        Get
            Return m_Fax
        End Get
        Set(ByVal value As String)
            m_Fax = value
        End Set
    End Property
    Private m_Fax As String

    Public Property ESubjectCongratulation() As String
        Get
            Return m_ESubjectCongratulation
        End Get
        Set(ByVal value As String)
            m_ESubjectCongratulation = value
        End Set
    End Property
    Private m_ESubjectCongratulation As String


    Public Property SinalTesteF() As String
        Get
            Return m_SinalTesteF
        End Get
        Set(ByVal value As String)
            m_SinalTesteF = value
        End Set
    End Property
    Private m_SinalTesteF As String


    Public Property SinalTesteJ() As String
        Get
            Return m_SinalTesteJ
        End Get
        Set(ByVal value As String)
            m_SinalTesteJ = value
        End Set
    End Property
    Private m_SinalTesteJ As String


    Public Property DiaIniMonit() As String
        Get
            Return m_DiaIniMonit
        End Get
        Set(ByVal value As String)
            m_DiaIniMonit = value
        End Set
    End Property
    Private m_DiaIniMonit As String


    Public Property DiaFimMonit() As String
        Get
            Return m_DiaFimMonit
        End Get
        Set(ByVal value As String)
            m_DiaFimMonit = value
        End Set
    End Property
    Private m_DiaFimMonit As String


    Public Property VctoMonitClieNovos() As String
        Get
            Return m_VctoMonitClieNovos
        End Get
        Set(ByVal value As String)
            m_VctoMonitClieNovos = value
        End Set
    End Property
    Private m_VctoMonitClieNovos As String


    Public Property DtEncerrContabil() As Nullable(Of DateTime)
        Get
            Return m_DtEncerrContabil
        End Get
        Set(ByVal value As Nullable(Of DateTime))
            m_DtEncerrContabil = value
        End Set
    End Property
    Private m_DtEncerrContabil As Nullable(Of DateTime)

    Public Property NumNF() As String
        Get
            Return m_NumNF
        End Get
        Set(ByVal value As String)
            m_NumNF = value
        End Set
    End Property
    Private m_NumNF As String

    Public Property NumNFe() As String
        Get
            Return m_NumNFe
        End Get
        Set(ByVal value As String)
            m_NumNFe = value
        End Set
    End Property
    Private m_NumNFe As String

    Public Property Serie() As String
        Get
            Return m_Serie
        End Get
        Set(ByVal value As String)
            m_Serie = value
        End Set
    End Property
    Private m_Serie As String

    Public Property SerieNFServ() As String
        Get
            Return m_SerieNFServ
        End Get
        Set(ByVal value As String)
            m_SerieNFServ = value
        End Set
    End Property
    Private m_SerieNFServ As String

    Public Property URL_MB_SUP() As String
        Get
            Return m_URL_MB_SUP
        End Get
        Set(ByVal value As String)
            m_URL_MB_SUP = value
        End Set
    End Property
    Private m_URL_MB_SUP As String

    Public Property Credential_MB_SUP() As String
        Get
            Return m_Credential_MB_SUP
        End Get
        Set(ByVal value As String)
            m_Credential_MB_SUP = value
        End Set
    End Property
    Private m_Credential_MB_SUP As String

    Public Property User_MB_SUP() As String
        Get
            Return m_User_MB_SUP
        End Get
        Set(ByVal value As String)
            m_User_MB_SUP = value
        End Set
    End Property
    Private m_User_MB_SUP As String

    Public Property Mode_MB_SUP() As String
        Get
            Return m_Mode_MB_SUP
        End Get
        Set(ByVal value As String)
            m_Mode_MB_SUP = value
        End Set
    End Property
    Private m_Mode_MB_SUP As String

    Public Property PrepFatVen() As String
        Get
            Return m_PrepFatVen
        End Get
        Set(ByVal value As String)
            m_PrepFatVen = value
        End Set
    End Property
    Private m_PrepFatVen As String

    Public Property PrepFatEmp() As String
        Get
            Return m_PrepFatEmp
        End Get
        Set(ByVal value As String)
            m_PrepFatEmp = value
        End Set
    End Property
    Private m_PrepFatEmp As String

    Public Property PrepFatIns() As String
        Get
            Return m_PrepFatIns
        End Get
        Set(ByVal value As String)
            m_PrepFatIns = value
        End Set
    End Property
    Private m_PrepFatIns As String

    Public Property PrepFatAde() As String
        Get
            Return m_PrepFatAde
        End Get
        Set(ByVal value As String)
            m_PrepFatAde = value
        End Set
    End Property
    Private m_PrepFatAde As String

    Public Property TravaEstoque() As String
        Get
            Return m_TravaEstoque
        End Get
        Set(ByVal value As String)
            m_TravaEstoque = value
        End Set
    End Property
    Private m_TravaEstoque As String

    Public Property PrepFatMon() As String
        Get
            Return m_PrepFatMon
        End Get
        Set(ByVal value As String)
            m_PrepFatMon = value
        End Set
    End Property
    Private m_PrepFatMon As String


    Public Property PrepFatDoa() As String
        Get
            Return m_PrepFatDoa
        End Get
        Set(ByVal value As String)
            m_PrepFatDoa = value
        End Set
    End Property
    Private m_PrepFatDoa As String


    Public Property PrepFatVenST() As String
        Get
            Return m_PrepFatVenST
        End Get
        Set(ByVal value As String)
            m_PrepFatVenST = value
        End Set
    End Property
    Private m_PrepFatVenST As String


    Public Property NumNFServ() As String
        Get
            Return m_NumNFServ
        End Get
        Set(ByVal value As String)
            m_NumNFServ = value
        End Set
    End Property
    Private m_NumNFServ As String


    Public Property SerieNF() As String
        Get
            Return m_SerieNF
        End Get
        Set(ByVal value As String)
            m_SerieNF = value
        End Set
    End Property
    Private m_SerieNF As String


    Public Property AliqIPI() As Double
        Get
            Return m_AliqIPI
        End Get
        Set(ByVal value As Double)
            m_AliqIPI = value
        End Set
    End Property
    Private m_AliqIPI As Double


    Public Property InscMunic() As String
        Get
            Return m_InscMunic
        End Get
        Set(ByVal value As String)
            m_InscMunic = value
        End Set
    End Property
    Private m_InscMunic As String


    Public Property CGC() As String
        Get
            Return m_CGC
        End Get
        Set(ByVal value As String)
            m_CGC = value
        End Set
    End Property
    Private m_CGC As String


    Public Property MemBoleto() As String
        Get
            Return m_MemBoleto
        End Get
        Set(ByVal value As String)
            m_MemBoleto = value
        End Set
    End Property
    Private m_MemBoleto As String


    Public Property Endereco() As String
        Get
            Return m_Endereco
        End Get
        Set(ByVal value As String)
            m_Endereco = value
        End Set
    End Property
    Private m_Endereco As String


    Public Property Numero() As String
        Get
            Return m_Numero
        End Get
        Set(ByVal value As String)
            m_Numero = value
        End Set
    End Property
    Private m_Numero As String


    Public Property Bairro() As String
        Get
            Return m_Bairro
        End Get
        Set(ByVal value As String)
            m_Bairro = value
        End Set
    End Property
    Private m_Bairro As String


    Public Property Complemento() As String
        Get
            Return m_Complemento
        End Get
        Set(ByVal value As String)
            m_Complemento = value
        End Set
    End Property
    Private m_Complemento As String

    Public Property Cidade() As String
        Get
            Return m_Cidade
        End Get
        Set(ByVal value As String)
            m_Cidade = value
        End Set
    End Property
    Private m_Cidade As String

    Public Property Cep() As String
        Get
            Return m_Cep
        End Get
        Set(ByVal value As String)
            m_Cep = value
        End Set
    End Property
    Private m_Cep As String

    Public Property NumArqRemREDECARD() As Integer
        Get
            Return m_NumArqRemREDECARD
        End Get
        Set(ByVal value As Integer)
            m_NumArqRemREDECARD = value
        End Set
    End Property
    Private m_NumArqRemREDECARD As Integer

    Public Property NumArqRemVISA() As Integer
        Get
            Return m_NumArqRemVISA
        End Get
        Set(ByVal value As Integer)
            m_NumArqRemVISA = value
        End Set
    End Property
    Private m_NumArqRemVISA As Integer


    Public Property NumArqRemAMEX() As Integer
        Get
            Return m_NumArqRemAMEX
        End Get
        Set(ByVal value As Integer)
            m_NumArqRemAMEX = value
        End Set
    End Property
    Private m_NumArqRemAMEX As Integer


    Public Property CCorAbase() As String
        Get
            Return m_CCorAbase
        End Get
        Set(ByVal value As String)
            m_CCorAbase = value
        End Set
    End Property
    Private m_CCorAbase As String


    Public Property CCorMbase() As String
        Get
            Return m_CCorMbase
        End Get
        Set(ByVal value As String)
            m_CCorMbase = value
        End Set
    End Property
    Private m_CCorMbase As String


    Public Property IndiceSug() As String
        Get
            Return m_IndiceSug
        End Get
        Set(ByVal value As String)
            m_IndiceSug = value
        End Set
    End Property
    Private m_IndiceSug As String

    Public Property CodEmprCtbl() As String
        Get
            Return m_CodEmprCtbl
        End Get
        Set(ByVal value As String)
            m_CodEmprCtbl = value
        End Set
    End Property
    Private m_CodEmprCtbl As String

    Public Property HistBalan() As String
        Get
            Return m_HistBalan
        End Get
        Set(ByVal value As String)
            m_HistBalan = value
        End Set
    End Property
    Private m_HistBalan As String


    Public Property CtaCtblBalan() As String
        Get
            Return m_CtaCtblBalan
        End Get
        Set(ByVal value As String)
            m_CtaCtblBalan = value
        End Set
    End Property
    Private m_CtaCtblBalan As String

    Public Property CtbMbase() As String
        Get
            Return m_CtbMbase
        End Get
        Set(ByVal value As String)
            m_CtbMbase = value
        End Set
    End Property
    Private m_CtbMbase As String

    Public Property UltCodOrc() As String
        Get
            Return m_UltCodOrc
        End Get
        Set(ByVal value As String)
            m_UltCodOrc = value
        End Set
    End Property
    Private m_UltCodOrc As String


    Public Property CNAE() As String
        Get
            Return m_CNAE
        End Get
        Set(ByVal value As String)
            m_CNAE = value
        End Set
    End Property
    Private m_CNAE As String



    Public Property InscrEstad() As String
        Get
            Return m_InscrEstad
        End Get
        Set(ByVal value As String)
            m_InscrEstad = value
        End Set
    End Property
    Private m_InscrEstad As String


    Public Property DDD() As String
        Get
            Return m_DDD
        End Get
        Set(ByVal value As String)
            m_DDD = value
        End Set
    End Property
    Private m_DDD As String

    Public Property NomeFantasia() As String
        Get
            Return m_NomeFantasia
        End Get
        Set(ByVal value As String)
            m_NomeFantasia = value
        End Set
    End Property
    Private m_NomeFantasia As String


    Public Property EMail() As String
        Get
            Return m_EMail
        End Get
        Set(ByVal value As String)
            m_EMail = value
        End Set
    End Property
    Private m_EMail As String

    Public Property HomePage() As String
        Get
            Return m_HomePage
        End Get
        Set(ByVal value As String)
            m_HomePage = value
        End Set
    End Property
    Private m_HomePage As String

    Public Property RamalFax() As String
        Get
            Return m_RamalFax
        End Get
        Set(ByVal value As String)
            m_RamalFax = value
        End Set
    End Property
    Private m_RamalFax As String

    Public Property AtivEcon As String
        Get
            Return m_AtivEcon
        End Get
        Set(ByVal value As String)
            m_AtivEcon = value
        End Set
    End Property
    Private m_AtivEcon As String

    Public Property UltNum() As String
        Get
            Return m_UltNum
        End Get
        Set(ByVal value As String)
            m_UltNum = value
        End Set
    End Property
    Private m_UltNum As String

    Public Property NumArqRemSerasa() As String
        Get
            Return m_NumArqRemSerasa
        End Get
        Set(ByVal value As String)
            m_NumArqRemSerasa = value
        End Set
    End Property
    Private m_NumArqRemSerasa As String

    Public Property EvCpSug() As String
        Get
            Return m_EvCpSug
        End Get
        Set(ByVal value As String)
            m_EvCpSug = value
        End Set
    End Property
    Private m_EvCpSug As String

    Public Property EvCrSug() As String
        Get
            Return m_EvCrSug
        End Get
        Set(ByVal value As String)
            m_EvCrSug = value
        End Set
    End Property
    Private m_EvCrSug As String

    Public Property NumDiasCCor() As String
        Get
            Return m_NumDiasCCor
        End Get
        Set(ByVal value As String)
            m_NumDiasCCor = value
        End Set
    End Property
    Private m_NumDiasCCor As String


    Public Property BloqVctoSolCpa() As String
        Get
            Return m_BloqVctoSolCpa
        End Get
        Set(ByVal value As String)
            m_BloqVctoSolCpa = value
        End Set
    End Property
    Private m_BloqVctoSolCpa As String


    Public Property BloqHoraSolCpa() As String
        Get
            Return m_BloqHoraSolCpa
        End Get
        Set(ByVal value As String)
            m_BloqHoraSolCpa = value
        End Set
    End Property
    Private m_BloqHoraSolCpa As String

    Public Property CtaCtblVendas() As String
        Get
            Return m_CtaCtblVendas
        End Get
        Set(ByVal value As String)
            m_CtaCtblVendas = value
        End Set
    End Property
    Private m_CtaCtblVendas As String

    Public Property CtaCtblVendasD() As String
        Get
            Return m_CtaCtblVendasD
        End Get
        Set(ByVal value As String)
            m_CtaCtblVendasD = value
        End Set
    End Property
    Private m_CtaCtblVendasD As String

    Public Property HistVendas() As String
        Get
            Return m_HistVendas
        End Get
        Set(ByVal value As String)
            m_HistVendas = value
        End Set
    End Property
    Private m_HistVendas As String

    Public Property CtaCtblICMSFat() As String
        Get
            Return m_CtaCtblICMSFat
        End Get
        Set(ByVal value As String)
            m_CtaCtblICMSFat = value
        End Set
    End Property
    Private m_CtaCtblICMSFat As String

    Public Property CtaCtblICMSaRec() As String
        Get
            Return m_CtaCtblICMSaRec
        End Get
        Set(ByVal value As String)
            m_CtaCtblICMSaRec = value
        End Set
    End Property
    Private m_CtaCtblICMSaRec As String

    Public Property HistICMSFat() As String
        Get
            Return m_HistICMSFat
        End Get
        Set(ByVal value As String)
            m_HistICMSFat = value
        End Set
    End Property
    Private m_HistICMSFat As String

    Public Property CtaCtblMonitCred() As String
        Get
            Return m_CtaCtblMonitCred
        End Get
        Set(ByVal value As String)
            m_CtaCtblMonitCred = value
        End Set
    End Property
    Private m_CtaCtblMonitCred As String

    Public Property CtaCtblMonitDebi() As String
        Get
            Return m_CtaCtblMonitDebi
        End Get
        Set(ByVal value As String)
            m_CtaCtblMonitDebi = value
        End Set
    End Property
    Private m_CtaCtblMonitDebi As String

    Public Property HistMonit() As String
        Get
            Return m_HistMonit
        End Get
        Set(ByVal value As String)
            m_HistMonit = value
        End Set
    End Property
    Private m_HistMonit As String

    Public Property NumNFRJ() As String
        Get
            Return m_NumNFRJ
        End Get
        Set(ByVal value As String)
            m_NumNFRJ = value
        End Set
    End Property
    Private m_NumNFRJ As String

    Public Property SerieNFRJ() As String
        Get
            Return m_SerieNFRJ
        End Get
        Set(ByVal value As String)
            m_SerieNFRJ = value
        End Set
    End Property
    Private m_SerieNFRJ As String

    Public Property NumNFServRJ() As String
        Get
            Return m_NumNFServRJ
        End Get
        Set(ByVal value As String)
            m_NumNFServRJ = value
        End Set
    End Property
    Private m_NumNFServRJ As String

    Public Property SerieNFServRJ() As String
        Get
            Return m_SerieNFServRJ
        End Get
        Set(ByVal value As String)
            m_SerieNFServRJ = value
        End Set
    End Property
    Private m_SerieNFServRJ As String

    Public Property NumNFServCamp() As String
        Get
            Return m_NumNFServCamp
        End Get
        Set(ByVal value As String)
            m_NumNFServCamp = value
        End Set
    End Property
    Private m_NumNFServCamp As String

    Public Property SerieNFServCamp() As String
        Get
            Return m_SerieNFServCamp
        End Get
        Set(ByVal value As String)
            m_SerieNFServCamp = value
        End Set
    End Property
    Private m_SerieNFServCamp As String

    Public Property NumNFCamp() As String
        Get
            Return m_NumNFCamp
        End Get
        Set(ByVal value As String)
            m_NumNFCamp = value
        End Set
    End Property
    Private m_NumNFCamp As String

    Public Property SerieNFCamp() As String
        Get
            Return m_SerieNFCamp
        End Get
        Set(ByVal value As String)
            m_SerieNFCamp = value
        End Set
    End Property
    Private m_SerieNFCamp As String

    Public Property NNC() As String
        Get
            Return m_NNC
        End Get
        Set(ByVal value As String)
            m_NNC = value
        End Set
    End Property
    Private m_NNC As String

    Public Property IndiceRel() As String
        Get
            Return m_IndiceRel
        End Get
        Set(ByVal value As String)
            m_IndiceRel = value
        End Set
    End Property
    Private m_IndiceRel As String

    Public Property CCustoVenda() As String
        Get
            Return m_CCustoVenda
        End Get
        Set(ByVal value As String)
            m_CCustoVenda = value
        End Set
    End Property
    Private m_CCustoVenda As String


    Public Property CCustoapto() As String
        Get
            Return m_CCustoapto
        End Get
        Set(ByVal value As String)
            m_CCustoapto = value
        End Set
    End Property
    Private m_CCustoapto As String

    Public Property PrepFatT() As String
        Get
            Return m_PrepFatT
        End Get
        Set(ByVal value As String)
            m_PrepFatT = value
        End Set
    End Property
    Private m_PrepFatT As String

    Public Property PComissaoMonitDealer As String
        Get
            Return m_PComissaoMonitDealer
        End Get
        Set(ByVal value As String)
            m_PComissaoMonitDealer = value
        End Set
    End Property
    Private m_PComissaoMonitDealer As String

    Public Property PComissaoMonit As String
        Get
            Return m_PComissaoMonit
        End Get
        Set(ByVal value As String)
            m_PComissaoMonit = value
        End Set
    End Property
    Private m_PComissaoMonit As String

    Public Property PercComisEquip As String
        Get
            Return m_PercComisEquip
        End Get
        Set(ByVal value As String)
            m_PercComisEquip = value
        End Set
    End Property
    Private m_PercComisEquip As String


    Public Property vlrMonBol As String
        Get
            Return m_vlrMonBol
        End Get
        Set(ByVal value As String)
            m_vlrMonBol = value
        End Set
    End Property
    Private m_vlrMonBol As String

    Public Property PrepFatCartaoVA() As String
        Get
            Return m_PrepFatCartaoVA
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVA = value
        End Set
    End Property
    Private m_PrepFatCartaoVA As String


    Public Property PrepFatCartaoVR() As String
        Get
            Return m_PrepFatCartaoVR
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVR = value
        End Set
    End Property
    Private m_PrepFatCartaoVR As String

    Public Property PrepFatCartaoVV() As String
        Get
            Return m_PrepFatCartaoVV
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVV = value
        End Set
    End Property
    Private m_PrepFatCartaoVV As String

    Public Property PrepFatCartaoMA() As String
        Get
            Return m_PrepFatCartaoMA
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMA = value
        End Set
    End Property
    Private m_PrepFatCartaoMA As String

    Public Property PrepFatCartaoMR() As String
        Get
            Return m_PrepFatCartaoMR
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMR = value
        End Set
    End Property
    Private m_PrepFatCartaoMR As String

    Public Property PrepFatCartaoMV() As String
        Get
            Return m_PrepFatCartaoMV
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMV = value
        End Set
    End Property
    Private m_PrepFatCartaoMV As String

    Public Property QtdeManutDia As String
        Get
            Return m_QtdeManutDia
        End Get
        Set(ByVal value As String)
            m_QtdeManutDia = value
        End Set
    End Property
    Private m_QtdeManutDia As String

    Public Property VlrRepPlantao() As String
        Get
            Return m_VlrRepPlantao
        End Get
        Set(ByVal value As String)
            m_VlrRepPlantao = value
        End Set
    End Property
    Private m_VlrRepPlantao As String

    Public Property CodComExpPag237() As String
        Get
            Return m_CodComExpPag237
        End Get
        Set(ByVal value As String)
            m_CodComExpPag237 = value
        End Set
    End Property
    Private m_CodComExpPag237 As String


    Public Property NumArqRemExpPag237() As String
        Get
            Return m_NumArqRemExpPag237
        End Get
        Set(ByVal value As String)
            m_NumArqRemExpPag237 = value
        End Set
    End Property
    Private m_NumArqRemExpPag237 As String


    Public Property NumSeqExpPag237() As String
        Get
            Return m_NumSeqExpPag237
        End Get
        Set(ByVal value As String)
            m_NumSeqExpPag237 = value
        End Set
    End Property
    Private m_NumSeqExpPag237 As String

    Public Property NumArqRemExpPag341() As String
        Get
            Return m_NumArqRemExpPag341
        End Get
        Set(ByVal value As String)
            m_NumArqRemExpPag341 = value
        End Set
    End Property
    Private m_NumArqRemExpPag341 As String



    Public Property CodComExpPag745() As String
        Get
            Return m_CodComExpPag745
        End Get
        Set(ByVal value As String)
            m_CodComExpPag745 = value
        End Set
    End Property
    Private m_CodComExpPag745 As String


    Public Property NumArqRemExpPag745() As String
        Get
            Return m_NumArqRemExpPag745
        End Get
        Set(ByVal value As String)
            m_NumArqRemExpPag745 = value
        End Set
    End Property
    Private m_NumArqRemExpPag745 As String


    Public Property NumSeqExpPag745() As String
        Get
            Return m_NumSeqExpPag745
        End Get
        Set(ByVal value As String)
            m_NumSeqExpPag745 = value
        End Set
    End Property
    Private m_NumSeqExpPag745 As String


    Public Property DirSispag() As String
        Get
            Return m_DirSispag
        End Get
        Set(ByVal value As String)
            m_DirSispag = value
        End Set
    End Property
    Private m_DirSispag As String


    Public Property CodAgeSisPag() As String
        Get
            Return m_CodAgeSisPag
        End Get
        Set(ByVal value As String)
            m_CodAgeSisPag = value
        End Set
    End Property
    Private m_CodAgeSisPag As String






    Private m_PrepFatAB As String
    Private m_PrepFatMB As String
    Private m_PrepFatTB As String
    Private m_PrepFatVBD As String
    Private m_PrepFatIBD As String
    Private m_PrepFatABD As String
    Private m_PrepFatMBD As String
    Private m_PrepFatTBD As String
    Private m_PrepFatCartaoMVB As String
    Private m_PrepFatCartaoMRB As String
    Private m_PrepFatCartaoMAB As String
    Private m_PrepFatCartaoMVBD As String
    Private m_PrepFatCartaoMRBD As String
    Private m_PrepFatCartaoMABD As String
    Private m_PrepFatCartaoVVB As String
    Private m_PrepFatCartaoVRB As String
    Private m_PrepFatCartaoVAB As String
    Private m_PrepFatCartaoVVBD As String
    Private m_PrepFatCartaoVRBD As String
    Private m_PrepFatCartaoVABD As String
    Private m_CodBcoSisPag As String
    Private m_PrepFatVC As String
    Private m_PrepFatIC As String
    Private m_PrepFatAC As String
    Private m_PrepFatMC As String
    Private m_PrepFatTC As String
    Private m_PrepFatVCD As String
    Private m_PrepFatICD As String
    Private m_PrepFatACD As String
    Private m_PrepFatMCD As String
    Private m_PrepFatTCD As String
    Private m_PrepFatCartaoMVC As String
    Private m_PrepFatCartaoMRC As String
    Private m_PrepFatCartaoMAC As String
    Private m_PrepFatCartaoMVCD As String
    Private m_PrepFatCartaoMRCD As String
    Private m_PrepFatCartaoMACD As String
    Private m_PrepFatCartaoVVC As String
    Private m_PrepFatCartaoVRC As String
    Private m_PrepFatCartaoVAC As String
    Private m_PrepFatCartaoVVCD As String
    Private m_PrepFatCartaoVRCD As String
    Private m_PrepFatCartaoVACD As String
    Private m_NumCtaSisPag As String
    Private m_PrepFatVA As String
    Private m_PrepFatIA As String
    Private m_PrepFatAA As String
    Private m_PrepFatMA As String
    Private m_PrepFatTA As String
    Private m_PrepFatVAD As String
    Private m_PrepFatIAD As String
    Private m_PrepFatAAD As String
    Private m_PrepFatMAD As String
    Private m_PrepFatTAD As String
    Private m_PrepFatCartaoMVA As String
    Private m_PrepFatCartaoMRA As String
    Private m_PrepFatCartaoMAA As String
    Private m_PrepFatCartaoMVAD As String
    Private m_PrepFatCartaoMRAD As String
    Private m_PrepFatCartaoMAAD As String
    Private m_PrepFatCartaoVVA As String
    Private m_PrepFatCartaoVRA As String
    Private m_PrepFatCartaoVAA As String
    Private m_PrepFatCartaoVVAD As String
    Private m_PrepFatCartaoVRAD As String
    Private m_PrepFatCartaoVAAD As String



    Public Property PrepFatVB() As String
        Get
            Return m_PrepFatVB
        End Get
        Set(ByVal value As String)
            m_PrepFatVB = value
        End Set
    End Property
    Private m_PrepFatVB As String

    Public Property PrepFatIB() As String
        Get
            Return m_PrepFatIB
        End Get
        Set(ByVal value As String)
            m_PrepFatIB = value
        End Set
    End Property
    Private m_PrepFatIB As String

    Public Property PrepFatAB() As String
        Get
            Return m_PrepFatAB
        End Get
        Set(ByVal value As String)
            m_PrepFatAB = value
        End Set
    End Property
    Public Property PrepFatMB() As String
        Get
            Return m_PrepFatMB
        End Get
        Set(ByVal value As String)
            m_PrepFatMB = value
        End Set
    End Property
    Public Property PrepFatTB() As String
        Get
            Return m_PrepFatTB
        End Get
        Set(ByVal value As String)
            m_PrepFatTB = value
        End Set
    End Property
    Public Property PrepFatVBD() As String
        Get
            Return m_PrepFatVBD
        End Get
        Set(ByVal value As String)
            m_PrepFatVBD = value
        End Set
    End Property
    Public Property PrepFatIBD() As String
        Get
            Return m_PrepFatIBD
        End Get
        Set(ByVal value As String)
            m_PrepFatIBD = value
        End Set
    End Property
    Public Property PrepFatABD() As String
        Get
            Return m_PrepFatABD
        End Get
        Set(ByVal value As String)
            m_PrepFatABD = value
        End Set
    End Property
    Public Property PrepFatMBD() As String
        Get
            Return m_PrepFatMBD
        End Get
        Set(ByVal value As String)
            m_PrepFatMBD = value
        End Set
    End Property
    Public Property PrepFatTBD() As String
        Get
            Return m_PrepFatTBD
        End Get
        Set(ByVal value As String)
            m_PrepFatTBD = value
        End Set
    End Property
    Public Property PrepFatCartaoMVB() As String
        Get
            Return m_PrepFatCartaoMVB
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMVB = value
        End Set
    End Property
    Public Property PrepFatCartaoMRB() As String
        Get
            Return m_PrepFatCartaoMRB
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMRB = value
        End Set
    End Property
    Public Property PrepFatCartaoMAB() As String
        Get
            Return m_PrepFatCartaoMAB
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMAB = value
        End Set
    End Property
    Public Property PrepFatCartaoMVBD() As String
        Get
            Return m_PrepFatCartaoMVBD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMVBD = value
        End Set
    End Property
    Public Property PrepFatCartaoMRBD() As String
        Get
            Return m_PrepFatCartaoMRBD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMRBD = value
        End Set
    End Property
    Public Property PrepFatCartaoMABD() As String
        Get
            Return m_PrepFatCartaoMABD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMABD = value
        End Set
    End Property
    Public Property PrepFatCartaoVVB() As String
        Get
            Return m_PrepFatCartaoVVB
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVVB = value
        End Set
    End Property
    Public Property PrepFatCartaoVRB() As String
        Get
            Return m_PrepFatCartaoVRB
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVRB = value
        End Set
    End Property
    Public Property PrepFatCartaoVAB() As String
        Get
            Return m_PrepFatCartaoVAB
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVAB = value
        End Set
    End Property
    Public Property PrepFatCartaoVVBD() As String
        Get
            Return m_PrepFatCartaoVVBD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVVBD = value
        End Set
    End Property
    Public Property PrepFatCartaoVRBD() As String
        Get
            Return m_PrepFatCartaoVRBD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVRBD = value
        End Set
    End Property
    Public Property PrepFatCartaoVABD() As String
        Get
            Return m_PrepFatCartaoVABD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVABD = value
        End Set
    End Property
    Public Property CodBcoSisPag() As String
        Get
            Return m_CodBcoSisPag
        End Get
        Set(ByVal value As String)
            m_CodBcoSisPag = value
        End Set
    End Property
    Public Property PrepFatVC() As String
        Get
            Return m_PrepFatVC
        End Get
        Set(ByVal value As String)
            m_PrepFatVC = value
        End Set
    End Property
    Public Property PrepFatIC() As String
        Get
            Return m_PrepFatIC
        End Get
        Set(ByVal value As String)
            m_PrepFatIC = value
        End Set
    End Property
    Public Property PrepFatAC() As String
        Get
            Return m_PrepFatAC
        End Get
        Set(ByVal value As String)
            m_PrepFatAC = value
        End Set
    End Property
    Public Property PrepFatMC() As String
        Get
            Return m_PrepFatMC
        End Get
        Set(ByVal value As String)
            m_PrepFatMC = value
        End Set
    End Property
    Public Property PrepFatTC() As String
        Get
            Return m_PrepFatTC
        End Get
        Set(ByVal value As String)
            m_PrepFatTC = value
        End Set
    End Property
    Public Property PrepFatVCD() As String
        Get
            Return m_PrepFatVCD
        End Get
        Set(ByVal value As String)
            m_PrepFatVCD = value
        End Set
    End Property
    Public Property PrepFatICD() As String
        Get
            Return m_PrepFatICD
        End Get
        Set(ByVal value As String)
            m_PrepFatICD = value
        End Set
    End Property
    Public Property PrepFatACD() As String
        Get
            Return m_PrepFatACD
        End Get
        Set(ByVal value As String)
            m_PrepFatACD = value
        End Set
    End Property
    Public Property PrepFatMCD() As String
        Get
            Return m_PrepFatMCD
        End Get
        Set(ByVal value As String)
            m_PrepFatMCD = value
        End Set
    End Property
    Public Property PrepFatTCD() As String
        Get
            Return m_PrepFatTCD
        End Get
        Set(ByVal value As String)
            m_PrepFatTCD = value
        End Set
    End Property
    Public Property PrepFatCartaoMVC() As String
        Get
            Return m_PrepFatCartaoMVC
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMVC = value
        End Set
    End Property
    Public Property PrepFatCartaoMRC() As String
        Get
            Return m_PrepFatCartaoMRC
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMRC = value
        End Set
    End Property
    Public Property PrepFatCartaoMAC() As String
        Get
            Return m_PrepFatCartaoMAC
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMAC = value
        End Set
    End Property
    Public Property PrepFatCartaoMVCD() As String
        Get
            Return m_PrepFatCartaoMVCD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMVCD = value
        End Set
    End Property
    Public Property PrepFatCartaoMRCD() As String
        Get
            Return m_PrepFatCartaoMRCD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMRCD = value
        End Set
    End Property
    Public Property PrepFatCartaoMACD() As String
        Get
            Return m_PrepFatCartaoMACD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMACD = value
        End Set
    End Property
    Public Property PrepFatCartaoVVC() As String
        Get
            Return m_PrepFatCartaoVVC
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVVC = value
        End Set
    End Property
    Public Property PrepFatCartaoVRC() As String
        Get
            Return m_PrepFatCartaoVRC
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVRC = value
        End Set
    End Property
    Public Property PrepFatCartaoVAC() As String
        Get
            Return m_PrepFatCartaoVAC
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVAC = value
        End Set
    End Property
    Public Property PrepFatCartaoVVCD() As String
        Get
            Return m_PrepFatCartaoVVCD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVVCD = value
        End Set
    End Property
    Public Property PrepFatCartaoVRCD() As String
        Get
            Return m_PrepFatCartaoVRCD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVRCD = value
        End Set
    End Property
    Public Property PrepFatCartaoVACD() As String
        Get
            Return m_PrepFatCartaoVACD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVACD = value
        End Set
    End Property
    Public Property NumCtaSisPag() As String
        Get
            Return m_NumCtaSisPag
        End Get
        Set(ByVal value As String)
            m_NumCtaSisPag = value
        End Set
    End Property
    Public Property PrepFatVA() As String
        Get
            Return m_PrepFatVA
        End Get
        Set(ByVal value As String)
            m_PrepFatVA = value
        End Set
    End Property
    Public Property PrepFatIA() As String
        Get
            Return m_PrepFatIA
        End Get
        Set(ByVal value As String)
            m_PrepFatIA = value
        End Set
    End Property
    Public Property PrepFatAA() As String
        Get
            Return m_PrepFatAA
        End Get
        Set(ByVal value As String)
            m_PrepFatAA = value
        End Set
    End Property
    Public Property PrepFatMA() As String
        Get
            Return m_PrepFatMA
        End Get
        Set(ByVal value As String)
            m_PrepFatMA = value
        End Set
    End Property
    Public Property PrepFatTA() As String
        Get
            Return m_PrepFatTA
        End Get
        Set(ByVal value As String)
            m_PrepFatTA = value
        End Set
    End Property
    Public Property PrepFatVAD() As String
        Get
            Return m_PrepFatVAD
        End Get
        Set(ByVal value As String)
            m_PrepFatVAD = value
        End Set
    End Property
    Public Property PrepFatIAD() As String
        Get
            Return m_PrepFatIAD
        End Get
        Set(ByVal value As String)
            m_PrepFatIAD = value
        End Set
    End Property
    Public Property PrepFatAAD() As String
        Get
            Return m_PrepFatAAD
        End Get
        Set(ByVal value As String)
            m_PrepFatAAD = value
        End Set
    End Property
    Public Property PrepFatMAD() As String
        Get
            Return m_PrepFatMAD
        End Get
        Set(ByVal value As String)
            m_PrepFatMAD = value
        End Set
    End Property
    Public Property PrepFatTAD() As String
        Get
            Return m_PrepFatTAD
        End Get
        Set(ByVal value As String)
            m_PrepFatTAD = value
        End Set
    End Property
    Public Property PrepFatCartaoMVA() As String
        Get
            Return m_PrepFatCartaoMVA
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMVA = value
        End Set
    End Property
    Public Property PrepFatCartaoMRA() As String
        Get
            Return m_PrepFatCartaoMRA
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMRA = value
        End Set
    End Property
    Public Property PrepFatCartaoMAA() As String
        Get
            Return m_PrepFatCartaoMAA
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMAA = value
        End Set
    End Property
    Public Property PrepFatCartaoMVAD() As String
        Get
            Return m_PrepFatCartaoMVAD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMVAD = value
        End Set
    End Property
    Public Property PrepFatCartaoMRAD() As String
        Get
            Return m_PrepFatCartaoMRAD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMRAD = value
        End Set
    End Property
    Public Property PrepFatCartaoMAAD() As String
        Get
            Return m_PrepFatCartaoMAAD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoMAAD = value
        End Set
    End Property
    Public Property PrepFatCartaoVVA() As String
        Get
            Return m_PrepFatCartaoVVA
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVVA = value
        End Set
    End Property
    Public Property PrepFatCartaoVRA() As String
        Get
            Return m_PrepFatCartaoVRA
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVRA = value
        End Set
    End Property
    Public Property PrepFatCartaoVAA() As String
        Get
            Return m_PrepFatCartaoVAA
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVAA = value
        End Set
    End Property
    Public Property PrepFatCartaoVVAD() As String
        Get
            Return m_PrepFatCartaoVVAD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVVAD = value
        End Set
    End Property
    Public Property PrepFatCartaoVRAD() As String
        Get
            Return m_PrepFatCartaoVRAD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVRAD = value
        End Set
    End Property
    Public Property PrepFatCartaoVAAD() As String
        Get
            Return m_PrepFatCartaoVAAD
        End Get
        Set(ByVal value As String)
            m_PrepFatCartaoVAAD = value
        End Set
    End Property


    Public Property NumArqRemCobExt() As String
        Get
            Return m_NumArqRemCobExt
        End Get
        Set(ByVal value As String)
            m_NumArqRemCobExt = value
        End Set
    End Property
    Private m_NumArqRemCobExt As String


    Public Property ESenderNameCobExt() As String
        Get
            Return m_ESenderNameCobExt
        End Get
        Set(ByVal value As String)
            m_ESenderNameCobExt = value
        End Set
    End Property
    Private m_ESenderNameCobExt As String


    Public Property ESenderEmailCobExt() As String
        Get
            Return m_ESenderEmailCobExt
        End Get
        Set(ByVal value As String)
            m_ESenderEmailCobExt = value
        End Set
    End Property
    Private m_ESenderEmailCobExt As String


    Public Property ESubjectCobExt() As String
        Get
            Return m_ESubjectCobExt
        End Get
        Set(ByVal value As String)
            m_ESubjectCobExt = value
        End Set
    End Property
    Private m_ESubjectCobExt As String


    Public Property CodPreClie() As String
        Get
            Return m_CodPreClie
        End Get
        Set(ByVal value As String)
            m_CodPreClie = value
        End Set
    End Property
    Private m_CodPreClie As String




    Public Property VlrTxVisita() As Double
        Get
            Return m_VlrTxVisita
        End Get
        Set(ByVal value As Double)
            m_VlrTxVisita = value
        End Set
    End Property
    Private m_VlrTxVisita As Double

    Public Property VlrRepasseInst() As Double
        Get
            Return m_VlrRepasseInst
        End Get
        Set(ByVal value As Double)
            m_VlrRepasseInst = value
        End Set
    End Property
    Private m_VlrRepasseInst As Double

    Public Property VlrConexao() As Double
        Get
            Return m_VlrConexao
        End Get
        Set(ByVal value As Double)
            m_VlrConexao = value
        End Set
    End Property
    Private m_VlrConexao As Double

    Public Property VlrVistoria() As Double
        Get
            Return m_VlrVistoria
        End Get
        Set(ByVal value As Double)
            m_VlrVistoria = value
        End Set
    End Property
    Private m_VlrVistoria As Double

    Public Property VlrRepasseCon() As Double
        Get
            Return m_VlrRepasseCon
        End Get
        Set(ByVal value As Double)
            m_VlrRepasseCon = value
        End Set
    End Property
    Private m_VlrRepasseCon As Double

    Public Property VlrMinEquip() As Double
        Get
            Return m_VlrMinEquip
        End Get
        Set(ByVal value As Double)
            m_VlrMinEquip = value
        End Set
    End Property
    Private m_VlrMinEquip As Double

    Public Property VlrMaxDescMObra() As Double
        Get
            Return m_VlrMaxDescMObra
        End Get
        Set(ByVal value As Double)
            m_VlrMaxDescMObra = value
        End Set
    End Property
    Private m_VlrMaxDescMObra As Double

    Public Property VlrvistoriaMoto() As Double
        Get
            Return m_VlrvistoriaMoto
        End Get
        Set(ByVal value As Double)
            m_VlrvistoriaMoto = value
        End Set
    End Property
    Private m_VlrvistoriaMoto As Double

    Public Property CodCpgt() As String
        Get
            Return m_CodCpgt
        End Get
        Set(ByVal value As String)
            m_CodCpgt = value
        End Set
    End Property
    Private m_CodCpgt As String

    Public Property VlrDescTxCon As Double
        Get
            Return m_VlrDescTxCon
        End Get
        Set(ByVal value As Double)
            m_VlrDescTxCon = value
        End Set
    End Property
    Private m_VlrDescTxCon As Double

    Public Property BDLiberado() As String
        Get
            Return m_BDLiberado
        End Get
        Set(ByVal value As String)
            m_BDLiberado = value
        End Set
    End Property
    Private m_BDLiberado As String

    'Criada para complementa a busca da data de fechamento teleservec
    'DataFechamento
    Public Property DataFechamento() As String
        Get
            Return m_DataFechamento
        End Get
        Set(ByVal value As String)
            m_DataFechamento = value
        End Set
    End Property
    Private m_DataFechamento As String

    Public Property EventosCentralConfPath() As String
        Get
            Return m_EventosCentralConfPath
        End Get
        Set(ByVal value As String)
            m_EventosCentralConfPath = value
        End Set
    End Property
    Private m_EventosCentralConfPath As String

    Private _PercRetCSLL As Single
    Public Property PercRetCSLL() As Single
        Get
            Return _PercRetCSLL
        End Get
        Set(ByVal value As Single)
            _PercRetCSLL = value
        End Set
    End Property

    Private _PercRetCofins As Decimal
    Public Property PercRetCofins() As Single
        Get
            Return _PercRetCofins
        End Get
        Set(ByVal value As Single)
            _PercRetCofins = value
        End Set
    End Property

    Private _PercRetPIS As Single
    Public Property PercRetPis() As Single
        Get
            Return _PercRetPIS
        End Get
        Set(ByVal value As Single)
            _PercRetPIS = value
        End Set
    End Property

    Private _VlrPisoCOFINS As Single
    Public Property VlrPisoCOFINS() As Single
        Get
            Return _VlrPisoCOFINS
        End Get
        Set(ByVal value As Single)
            _VlrPisoCOFINS = value
        End Set
    End Property

    Private _AliqRetIr As Single
    Public Property AliqRetIr() As Single
        Get
            Return _AliqRetIr
        End Get
        Set(ByVal value As Single)
            _AliqRetIr = value
        End Set
    End Property

    Private _VlrMinRetIr As Single
    Public Property VlrMinRetIr() As Single
        Get
            Return _VlrMinRetIr
        End Get
        Set(ByVal value As Single)
            _VlrMinRetIr = value
        End Set
    End Property

    Private _AliqRetINSS As Single
    Public Property AliqRetINSS() As Single
        Get
            Return _AliqRetINSS
        End Get
        Set(ByVal value As Single)
            _AliqRetINSS = value
        End Set
    End Property

    Public Property ESenderNameConfVda() As String
        Get
            Return m_ESenderNameConfVda
        End Get
        Set(ByVal value As String)
            m_ESenderNameConfVda = value
        End Set
    End Property
    Private m_ESenderNameConfVda As String

    Public Property ESenderEmailConfVda() As String
        Get
            Return m_ESenderEmailConfVda
        End Get
        Set(ByVal value As String)
            m_ESenderEmailConfVda = value
        End Set
    End Property
    Private m_ESenderEmailConfVda As String

    Public Property ESubjectConfVda() As String
        Get
            Return m_ESubjectConfVda
        End Get
        Set(ByVal value As String)
            m_ESubjectConfVda = value
        End Set
    End Property
    Private m_ESubjectConfVda As String
    Private m_ESenderSuporte As String
    Public Property ESenderSuporte As String
        Get
            Return m_ESenderSuporte
        End Get
        Set(ByVal value As String)
            m_ESenderSuporte = value
        End Set
    End Property


    Private m_ESenderEmail As String
    Public Property ESenderEmail() As String
        Get
            Return m_ESenderEmail
        End Get
        Set(ByVal value As String)
            m_ESenderEmail = value
        End Set
    End Property

    Private m_ESubject As String
    Public Property ESubject() As String
        Get
            Return m_ESubject
        End Get
        Set(ByVal value As String)
            m_ESubject = value
        End Set
    End Property

    Private m_EmailOSVistRJ As String
    Public Property EmailOSVistRJ() As String
        Get
            Return m_EmailOSVistRJ
        End Get
        Set(ByVal value As String)
            m_EmailOSVistRJ = value
        End Set
    End Property

    Private m_ESubjectInst As String
    Public Property ESubjectInst() As String
        Get
            Return m_ESubjectInst
        End Get
        Set(ByVal value As String)
            m_ESubjectInst = value
        End Set
    End Property

    Private m_ESubjectWorking As String
    Public Property ESubjectWorking() As String
        Get
            Return m_ESubjectWorking
        End Get
        Set(ByVal value As String)
            m_ESubjectWorking = value
        End Set
    End Property

    Private m_ESubjectWelcome As String
    Public Property ESubjectWelcome() As String
        Get
            Return m_ESubjectWelcome
        End Get
        Set(ByVal value As String)
            m_ESubjectWelcome = value
        End Set
    End Property

    Private m_ESenderNameLV As String
    Public Property ESenderNameLV() As String
        Get
            Return m_ESenderNameLV
        End Get
        Set(ByVal value As String)
            m_ESenderNameLV = value
        End Set
    End Property

    Private m_ESenderEmailLV As String
    Public Property ESenderEmailLV() As String
        Get
            Return m_ESenderEmailLV
        End Get
        Set(ByVal value As String)
            m_ESenderEmailLV = value
        End Set
    End Property

    Private m_ESubjectLV As String
    Public Property ESubjectLV() As String
        Get
            Return m_ESubjectLV
        End Get
        Set(ByVal value As String)
            m_ESubjectLV = value
        End Set
    End Property

    Private m_UltCodLote As Integer
    Public Property UltCodLote() As Integer
        Get
            Return m_UltCodLote
        End Get
        Set(ByVal value As Integer)
            m_UltCodLote = value
        End Set
    End Property

    Private m_PrepFatVBCFTV As String
    Public Property PrepFatVBCFTV() As String
        Get
            Return m_PrepFatVBCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatVBCFTV = value
        End Set
    End Property

    Private m_PrepFatVACFTV As String
    Public Property PrepFatVACFTV() As String
        Get
            Return m_PrepFatVACFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatVACFTV = value
        End Set
    End Property

    Private m_PrepFatVCCFTV As String
    Public Property PrepFatVCCFTV() As String
        Get
            Return m_PrepFatVCCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatVCCFTV = value
        End Set
    End Property

    Private m_PrepFatVBDCFTV As String
    Public Property PrepFatVBDCFTV() As String
        Get
            Return m_PrepFatVBDCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatVBDCFTV = value
        End Set
    End Property

    Private m_PrepFatVADCFTV As String
    Public Property PrepFatVADCFTV() As String
        Get
            Return m_PrepFatVADCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatVADCFTV = value
        End Set
    End Property

    Private m_PrepFatVCDCFTV As String
    Public Property PrepFatVCDCFTV() As String
        Get
            Return m_PrepFatVCDCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatVCDCFTV = value
        End Set
    End Property

    Private m_PrepFatIBCFTV As String
    Public Property PrepFatIBCFTV() As String
        Get
            Return m_PrepFatIBCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatIBCFTV = value
        End Set
    End Property

    Private m_PrepFatIACFTV As String
    Public Property PrepFatIACFTV() As String
        Get
            Return m_PrepFatIACFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatIACFTV = value
        End Set
    End Property

    Private m_PrepFatICCFTV As String
    Public Property PrepFatICCFTV() As String
        Get
            Return m_PrepFatICCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatICCFTV = value
        End Set
    End Property

    Private m_PrepFatIBDCFTV As String
    Public Property PrepFatIBDCFTV() As String
        Get
            Return m_PrepFatIBDCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatIBDCFTV = value
        End Set
    End Property

    Private m_PrepFatIADCFTV As String
    Public Property PrepFatIADCFTV() As String
        Get
            Return m_PrepFatIADCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatIADCFTV = value
        End Set
    End Property

    Private m_PrepFatICDCFTV As String
    Public Property PrepFatICDCFTV() As String
        Get
            Return m_PrepFatICDCFTV
        End Get
        Set(ByVal value As String)
            m_PrepFatICDCFTV = value
        End Set
    End Property


    Private m_HelpDeskPath As String
    Public Property HelpDeskPath() As String
        Get
            Return m_HelpDeskPath
        End Get
        Set(ByVal value As String)
            m_HelpDeskPath = value
        End Set
    End Property

    Private m_PrepFatAdiantamento As String
    Public Property PrepFatAdiantamento() As String
        Get
            Return m_PrepFatAdiantamento
        End Get
        Set(ByVal value As String)
            m_PrepFatAdiantamento = value
        End Set
    End Property

    Private m_PrepFatBAdiantamento As String
    Public Property PrepFatBAdiantamento() As String
        Get
            Return m_PrepFatBAdiantamento
        End Get
        Set(ByVal value As String)
            m_PrepFatBAdiantamento = value
        End Set
    End Property

    Private m_PrepFatAAdiantamento As String
    Public Property PrepFatAAdiantamento() As String
        Get
            Return m_PrepFatAAdiantamento
        End Get
        Set(ByVal value As String)
            m_PrepFatAAdiantamento = value
        End Set
    End Property

    Private m_PrepFatCAdiantamento As String
    Public Property PrepFatCAdiantamento() As String
        Get
            Return m_PrepFatCAdiantamento
        End Get
        Set(ByVal value As String)
            m_PrepFatCAdiantamento = value
        End Set
    End Property

    Private m_PrepFatBDAdiantamento As String
    Public Property PrepFatBDAdiantamento() As String
        Get
            Return m_PrepFatBDAdiantamento
        End Get
        Set(ByVal value As String)
            m_PrepFatBDAdiantamento = value
        End Set
    End Property

    Private m_PrepFatADAdiantamento As String
    Public Property PrepFatADAdiantamento() As String
        Get
            Return m_PrepFatADAdiantamento
        End Get
        Set(ByVal value As String)
            m_PrepFatADAdiantamento = value
        End Set
    End Property

    Private m_PrepFatCDAdiantamento As String
    Public Property PrepFatCDAdiantamento() As String
        Get
            Return m_PrepFatCDAdiantamento
        End Get
        Set(ByVal value As String)
            m_PrepFatCDAdiantamento = value
        End Set
    End Property

    Private m_LimiteCompra As Double
    Public Property LimiteCompra() As Double
        Get
            Return m_LimiteCompra
        End Get
        Set(ByVal value As Double)
            m_LimiteCompra = value
        End Set
    End Property

    Public Property UpdatePathVerisure() As String
        Get
            Return m_UpdatePathVerisure
        End Get
        Set(ByVal value As String)
            m_UpdatePathVerisure = value
        End Set
    End Property
    Private m_UpdatePathVerisure As String

    Public Property CREDENTIAL_MB_IND As String
    Public Property CaminhoFotoVendedor As String

    Public Property PercRetPISVenda As Single
    Public Property PercRetCOFINSVenda As Single

    Public Property CaminhoFotoVendedorVerisure As String

    'Novos campos para conexão promocional
    Public Property vlrConPromo As Double
    Public Property vlrRepasseInstConPromo As Double
    Public Property vlrDescConPromo As Double

    Public Property LoteRecebimentoNAV As String

    Public Property NumArqRemBVS As String
    Public Property NumArqRemSerasaTELE As String
    Public Property UrlEnvioBoleto As String
    Public Property ContratoPathVERI As String
    Public Property isFaturamentoEmProgresso As Integer
    Public Property VlrPlantaoManut As Double
    Public Property Versao As String
    Public Property ProtocoloHistContato As String

End Class