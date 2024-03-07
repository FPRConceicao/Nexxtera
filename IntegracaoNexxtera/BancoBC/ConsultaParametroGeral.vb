Imports System.Data.SqlClient
Imports System.Data
Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.ExceptionsPersonalizadas
Imports System.Reflection


Public Class ConsultaParametroGeral
    ''' <summary>
    ''' Retorna os dados da tabela parametro
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PesquisaParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_PesquisaParametros", connection)
        Try
            ''Informa a procedure

            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            connection.Open()

            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _retParametros.PermAumPtoInadimp = rdr("PermAumPtoInadimp").ToString()
                _retParametros.PermMRSInadimp = rdr("PermMRSInadimp").ToString()
                _retParametros.SeqDI = rdr("SeqDI").ToString()
                _retParametros.VenUNSC = rdr("VenUNSC").ToString()
                _retParametros.EmailConfPath = rdr("EmailConfPath").ToString()
                _retParametros.ESMTPServer = rdr("ESMTPServer").ToString()
                _retParametros.ContratoPath = rdr("ContratoPath").ToString()
                _retParametros.URLClaro = rdr("URLClaro").ToString()
                _retParametros.ProfileClaro = rdr("ProfileClaro").ToString()
                _retParametros.PwdClaro = rdr("PwdClaro").ToString()
                _retParametros.ModeClaro = rdr("ModeClaro").ToString()
                _retParametros.Mode_MB = rdr("Mode_MB").ToString()
                _retParametros.User_MB = rdr("User_MB").ToString()
                _retParametros.Credential_MB = rdr("Credential_MB").ToString()
                _retParametros.URL_MB = rdr("URL_MB").ToString()
                _retParametros.ESenderName = Trim(rdr("ESenderName").ToString())
                _retParametros.UF = Trim(rdr("UF").ToString())
                _retParametros.UpdatePath = Trim(rdr("UpdatePath").ToString())
                _retParametros.ESenderNameInst = rdr("ESenderNameInst").ToString()
                _retParametros.ESenderEmailInst = rdr("ESenderEmailInst").ToString()
                _retParametros.ESenderNameManut = rdr("ESenderNameManut").ToString()
                _retParametros.ESenderEmailManut = rdr("ESenderEmailManut").ToString()
                _retParametros.ESubjectManut = rdr("ESubjectManut").ToString()
                _retParametros.AliqIPI = rdr("AliqIPI")
                _retParametros.ESenderNameInd = rdr("ESenderNameInd").ToString()
                _retParametros.ESenderEmailInd = rdr("ESenderEmailInd").ToString()
                _retParametros.ESubjectInd = rdr("ESubjectInd").ToString()
                _retParametros.RazaoSocial = rdr("RazaoSocial").ToString()
                _retParametros.DDD = rdr("DDD").ToString()
                _retParametros.Fone = rdr("Fone").ToString()
                _retParametros.Fax = rdr("Fax").ToString()
                _retParametros.ESubjectCongratulation = rdr("ESubjectCongratulation").ToString()
                _retParametros.CNAE = rdr("CNAE").ToString()

                _retParametros.DiaIniMonit = rdr("DiaIniMonit").ToString()
                _retParametros.DiaFimMonit = rdr("DiaFimMonit").ToString()
                _retParametros.VctoMonitClieNovos = rdr("VctoMonitClieNovos").ToString()

                _retParametros.Credential_MB_SUP = rdr("Credential_MB_SUP").ToString()
                _retParametros.Mode_MB_SUP = rdr("Mode_MB_SUP").ToString()
                _retParametros.URL_MB_SUP = rdr("URL_MB_SUP").ToString()
                _retParametros.User_MB_SUP = rdr("User_MB_SUP").ToString()
                _retParametros.InscMunic = rdr("InscMunic").ToString()
                _retParametros.InscrEstad = rdr("InscrEstad").ToString()
                _retParametros.CGC = rdr("CGC").ToString()
                _retParametros.MemBoleto = rdr("MemBoleto").ToString()

                _retParametros.Endereco = rdr("Endereco").ToString()
                _retParametros.Numero = rdr("Numero").ToString()
                _retParametros.Complemento = rdr("Complemento").ToString()
                _retParametros.Bairro = rdr("Bairro").ToString()
                _retParametros.Cidade = rdr("Cidade").ToString()
                _retParametros.Cep = rdr("Cep").ToString()
                _retParametros.NumArqRemREDECARD = rdr("NumArqRemREDECARD").ToString
                _retParametros.NumArqRemVISA = rdr("NumArqRemVISA").ToString
                _retParametros.NumArqRemAMEX = rdr("NumArqRemAMEX").ToString
                _retParametros.CCorAbase = rdr("CCorAbase").ToString
                _retParametros.CCorMbase = rdr("CCorMbase").ToString
                _retParametros.IndiceSug = rdr("IndiceSug").ToString
                _retParametros.UltCodOrc = rdr("UltCodOrc").ToString
                _retParametros.CtbMbase = rdr("CtbMbase").ToString
                _retParametros.CtaCtblBalan = rdr("CtaCtblBalan").ToString
                _retParametros.HistBalan = rdr("HistBalan").ToString
                _retParametros.CodEmprCtbl = rdr("CodEmprCtbl").ToString
                _retParametros.NomeFantasia = IIf(IsDBNull(rdr("NomeFantasia")), "", rdr("NomeFantasia"))
                _retParametros.EMail = IIf(IsDBNull(rdr("EMail")), "", rdr("EMail"))
                _retParametros.HomePage = IIf(IsDBNull(rdr("HomePage")), "", rdr("HomePage"))
                _retParametros.RamalFax = IIf(IsDBNull(rdr("RamalFax")), "", rdr("RamalFax"))
                _retParametros.AtivEcon = IIf(IsDBNull(rdr("AtivEcon")), "", rdr("AtivEcon"))
                _retParametros.UltNum = IIf(IsDBNull(rdr("UltNum")), "", rdr("UltNum"))
                _retParametros.NumArqRemSerasa = IIf(IsDBNull(rdr("NumArqRemSerasa")), "", rdr("NumArqRemSerasa"))
                _retParametros.EvCpSug = IIf(IsDBNull(rdr("EvCpSug")), "", rdr("EvCpSug"))
                _retParametros.EvCrSug = IIf(IsDBNull(rdr("EvCrSug")), "", rdr("EvCrSug"))
                _retParametros.NumDiasCCor = IIf(IsDBNull(rdr("NumDiasCCor")), "", rdr("NumDiasCCor"))
                _retParametros.BloqVctoSolCpa = IIf(IsDBNull(rdr("BloqVctoSolCpa")), "", rdr("BloqVctoSolCpa"))
                _retParametros.BloqHoraSolCpa = IIf(IsDBNull(rdr("BloqHoraSolCpa")), "", rdr("BloqHoraSolCpa"))
                _retParametros.CtaCtblVendas = IIf(IsDBNull(rdr("CtaCtblVendas")), "", rdr("CtaCtblVendas"))
                _retParametros.CtaCtblVendasD = IIf(IsDBNull(rdr("CtaCtblVendasD")), "", rdr("CtaCtblVendasD"))
                _retParametros.HistVendas = IIf(IsDBNull(rdr("HistVendas")), "", rdr("HistVendas"))
                _retParametros.CtaCtblICMSFat = IIf(IsDBNull(rdr("CtaCtblICMSFat")), "", rdr("CtaCtblICMSFat"))
                _retParametros.CtaCtblICMSaRec = IIf(IsDBNull(rdr("CtaCtblICMSaRec")), "", rdr("CtaCtblICMSaRec"))
                _retParametros.HistICMSFat = IIf(IsDBNull(rdr("HistICMSFat")), "", rdr("HistICMSFat"))
                _retParametros.CtaCtblMonitCred = IIf(IsDBNull(rdr("CtaCtblMonitCred")), "", rdr("CtaCtblMonitCred"))
                _retParametros.CtaCtblMonitDebi = IIf(IsDBNull(rdr("CtaCtblMonitDebi")), "", rdr("CtaCtblMonitDebi"))
                _retParametros.HistMonit = IIf(IsDBNull(rdr("HistMonit")), "", rdr("HistMonit"))
                _retParametros.NumNFRJ = IIf(IsDBNull(rdr("NumNFRJ")), "", rdr("NumNFRJ"))
                _retParametros.NumNFServ = rdr("NumNFServ").ToString()
                _retParametros.NumNF = IIf(IsDBNull(rdr("NumNF")), "", rdr("NumNF"))
                _retParametros.SerieNFRJ = IIf(IsDBNull(rdr("SerieNFRJ")), "", rdr("SerieNFRJ"))
                _retParametros.SerieNFServ = IIf(IsDBNull(rdr("SerieNFServ")), "", rdr("SerieNFServ"))
                _retParametros.NumNFServRJ = IIf(IsDBNull(rdr("NumNFServRJ")), "", rdr("NumNFServRJ"))
                _retParametros.SerieNFServRJ = IIf(IsDBNull(rdr("SerieNFServRJ")), "", rdr("SerieNFServRJ"))
                _retParametros.NumNFServCamp = IIf(IsDBNull(rdr("NumNFServCamp")), "", rdr("NumNFServCamp"))
                _retParametros.SerieNFServCamp = IIf(IsDBNull(rdr("SerieNFServCamp")), "", rdr("SerieNFServCamp"))
                _retParametros.NumNFCamp = IIf(IsDBNull(rdr("NumNFCamp")), "", rdr("NumNFCamp"))
                _retParametros.SerieNFCamp = IIf(IsDBNull(rdr("SerieNFCamp")), "", rdr("SerieNFCamp"))
                _retParametros.NNC = IIf(IsDBNull(rdr("NNC")), "", rdr("NNC"))
                _retParametros.SerieNF = IIf(IsDBNull(rdr("SerieNF")), "", rdr("SerieNF"))
                _retParametros.NumNFe = IIf(IsDBNull(rdr("NumNFe")), "", rdr("NumNFe"))
                _retParametros.IndiceRel = IIf(IsDBNull(rdr("IndiceRel")), "", rdr("IndiceRel"))
                _retParametros.CCustoVenda = IIf(IsDBNull(rdr("CCustoVenda")), "", rdr("CCustoVenda"))
                _retParametros.CCustoapto = IIf(IsDBNull(rdr("CCustoapto")), "", rdr("CCustoapto"))
                _retParametros.PrepFatT = IIf(IsDBNull(rdr("PrepFatT")), "", rdr("PrepFatT"))
                _retParametros.PComissaoMonitDealer = IIf(IsDBNull(rdr("PComissaoMonitDealer")), "", rdr("PComissaoMonitDealer"))
                _retParametros.PComissaoMonit = IIf(IsDBNull(rdr("PComissaoMonit")), "", rdr("PComissaoMonit"))
                _retParametros.PercComisEquip = IIf(IsDBNull(rdr("PercComisEquip")), "", rdr("PercComisEquip"))
                _retParametros.vlrMonBol = IIf(IsDBNull(rdr("vlrMonBol")), "", rdr("vlrMonBol"))
                _retParametros.PrepFatCartaoVA = IIf(IsDBNull(rdr("PrepFatCartaoVA")), "", rdr("PrepFatCartaoVA"))
                _retParametros.PrepFatCartaoVR = IIf(IsDBNull(rdr("PrepFatCartaoVR")), "", rdr("PrepFatCartaoVR"))
                _retParametros.PrepFatCartaoVV = IIf(IsDBNull(rdr("PrepFatCartaoVV")), "", rdr("PrepFatCartaoVV"))
                _retParametros.PrepFatCartaoMA = IIf(IsDBNull(rdr("PrepFatCartaoMA")), "", rdr("PrepFatCartaoMA"))
                _retParametros.PrepFatCartaoMR = IIf(IsDBNull(rdr("PrepFatCartaoMR")), "", rdr("PrepFatCartaoMR"))
                _retParametros.PrepFatCartaoMV = IIf(IsDBNull(rdr("PrepFatCartaoMV")), "", rdr("PrepFatCartaoMV"))
                _retParametros.QtdeManutDia = IIf(IsDBNull(rdr("QtdeManutDia")), "", rdr("QtdeManutDia"))
                _retParametros.VlrRepPlantao = IIf(IsDBNull(rdr("VlrRepPlantao")), "", rdr("VlrRepPlantao"))
                _retParametros.CodComExpPag237 = IIf(IsDBNull(rdr("CodComExpPag237")), "", rdr("CodComExpPag237"))
                _retParametros.NumArqRemExpPag237 = IIf(IsDBNull(rdr("NumArqRemExpPag237")), "", rdr("NumArqRemExpPag237"))
                _retParametros.NumSeqExpPag237 = IIf(IsDBNull(rdr("NumSeqExpPag237")), "", rdr("NumSeqExpPag237"))
                _retParametros.NumArqRemExpPag341 = IIf(IsDBNull(rdr("NumArqRemExpPag341")), "", rdr("NumArqRemExpPag341"))
                _retParametros.DirSispag = IIf(IsDBNull(rdr("DirSispag")), "", rdr("DirSispag"))
                _retParametros.PrepFatT = IIf(IsDBNull(rdr("PrepFatT")), "", rdr("PrepFatT"))
                _retParametros.NumCtaSisPag = IIf(IsDBNull(rdr("NumCtaSisPag")), "", rdr("NumCtaSisPag"))
                _retParametros.CodAgeSisPag = IIf(IsDBNull(rdr("CodAgeSisPag")), "", rdr("CodAgeSisPag"))
                _retParametros.CodBcoSisPag = IIf(IsDBNull(rdr("CodBcoSisPag")), "", rdr("CodBcoSisPag"))
                _retParametros.PrepFatIB = IIf(IsDBNull(rdr("PrepFatIB")), "", rdr("PrepFatIB"))
                _retParametros.PrepFatVB = IIf(IsDBNull(rdr("PrepFatVB")), "", rdr("PrepFatVB"))
                _retParametros.PrepFatAB = IIf(IsDBNull(rdr("PrepFatAB")), "", rdr("PrepFatAB"))
                _retParametros.PrepFatMB = IIf(IsDBNull(rdr("PrepFatMB")), "", rdr("PrepFatMB"))
                _retParametros.PrepFatTB = IIf(IsDBNull(rdr("PrepFatTB")), "", rdr("PrepFatTB"))
                _retParametros.PrepFatVBD = IIf(IsDBNull(rdr("PrepFatVBD")), "", rdr("PrepFatVBD"))
                _retParametros.PrepFatIBD = IIf(IsDBNull(rdr("PrepFatIBD")), "", rdr("PrepFatIBD"))
                _retParametros.PrepFatABD = IIf(IsDBNull(rdr("PrepFatABD")), "", rdr("PrepFatABD"))
                _retParametros.PrepFatMBD = IIf(IsDBNull(rdr("PrepFatMBD")), "", rdr("PrepFatMBD"))
                _retParametros.PrepFatTBD = IIf(IsDBNull(rdr("PrepFatTBD")), "", rdr("PrepFatTBD"))
                _retParametros.PrepFatCartaoMVB = IIf(IsDBNull(rdr("PrepFatCartaoMVB")), "", rdr("PrepFatCartaoMVB"))
                _retParametros.PrepFatCartaoMRB = IIf(IsDBNull(rdr("PrepFatCartaoMRB")), "", rdr("PrepFatCartaoMRB"))
                _retParametros.PrepFatCartaoMAB = IIf(IsDBNull(rdr("PrepFatCartaoMAB")), "", rdr("PrepFatCartaoMAB"))
                _retParametros.PrepFatCartaoMVBD = IIf(IsDBNull(rdr("PrepFatCartaoMVBD")), "", rdr("PrepFatCartaoMVBD"))
                _retParametros.PrepFatCartaoMRBD = IIf(IsDBNull(rdr("PrepFatCartaoMRBD")), "", rdr("PrepFatCartaoMRBD"))
                _retParametros.PrepFatCartaoMABD = IIf(IsDBNull(rdr("PrepFatCartaoMABD")), "", rdr("PrepFatCartaoMABD"))
                _retParametros.PrepFatCartaoVVB = IIf(IsDBNull(rdr("PrepFatCartaoVVB")), "", rdr("PrepFatCartaoVVB"))
                _retParametros.PrepFatCartaoVRB = IIf(IsDBNull(rdr("PrepFatCartaoVRB")), "", rdr("PrepFatCartaoVRB"))
                _retParametros.PrepFatCartaoVAB = IIf(IsDBNull(rdr("PrepFatCartaoVAB")), "", rdr("PrepFatCartaoVAB"))
                _retParametros.PrepFatCartaoVVBD = IIf(IsDBNull(rdr("PrepFatCartaoVVBD")), "", rdr("PrepFatCartaoVVBD"))
                _retParametros.PrepFatCartaoVRBD = IIf(IsDBNull(rdr("PrepFatCartaoVRBD")), "", rdr("PrepFatCartaoVRBD"))
                _retParametros.PrepFatCartaoVABD = IIf(IsDBNull(rdr("PrepFatCartaoVABD")), "", rdr("PrepFatCartaoVABD"))
                _retParametros.CodBcoSisPag = IIf(IsDBNull(rdr("CodBcoSisPag")), "", rdr("CodBcoSisPag"))
                _retParametros.PrepFatVC = IIf(IsDBNull(rdr("PrepFatVC")), "", rdr("PrepFatVC"))
                _retParametros.PrepFatIC = IIf(IsDBNull(rdr("PrepFatIC")), "", rdr("PrepFatIC"))
                _retParametros.PrepFatAC = IIf(IsDBNull(rdr("PrepFatAC")), "", rdr("PrepFatAC"))
                _retParametros.PrepFatMC = IIf(IsDBNull(rdr("PrepFatMC")), "", rdr("PrepFatMC"))
                _retParametros.PrepFatTC = IIf(IsDBNull(rdr("PrepFatTC")), "", rdr("PrepFatTC"))
                _retParametros.PrepFatVCD = IIf(IsDBNull(rdr("PrepFatVCD")), "", rdr("PrepFatVCD"))
                _retParametros.PrepFatICD = IIf(IsDBNull(rdr("PrepFatICD")), "", rdr("PrepFatICD"))
                _retParametros.PrepFatACD = IIf(IsDBNull(rdr("PrepFatACD")), "", rdr("PrepFatACD"))
                _retParametros.PrepFatMCD = IIf(IsDBNull(rdr("PrepFatMCD")), "", rdr("PrepFatMCD"))
                _retParametros.PrepFatTCD = IIf(IsDBNull(rdr("PrepFatTCD")), "", rdr("PrepFatTCD"))
                _retParametros.PrepFatCartaoMVC = IIf(IsDBNull(rdr("PrepFatCartaoMVC")), "", rdr("PrepFatCartaoMVC"))
                _retParametros.PrepFatCartaoMRC = IIf(IsDBNull(rdr("PrepFatCartaoMRC")), "", rdr("PrepFatCartaoMRC"))
                _retParametros.PrepFatCartaoMAC = IIf(IsDBNull(rdr("PrepFatCartaoMAC")), "", rdr("PrepFatCartaoMAC"))
                _retParametros.PrepFatCartaoMVCD = IIf(IsDBNull(rdr("PrepFatCartaoMVCD")), "", rdr("PrepFatCartaoMVCD"))
                _retParametros.PrepFatCartaoMRCD = IIf(IsDBNull(rdr("PrepFatCartaoMRCD")), "", rdr("PrepFatCartaoMRCD"))
                _retParametros.PrepFatCartaoMACD = IIf(IsDBNull(rdr("PrepFatCartaoMACD")), "", rdr("PrepFatCartaoMACD"))
                _retParametros.PrepFatCartaoVVC = IIf(IsDBNull(rdr("PrepFatCartaoVVC")), "", rdr("PrepFatCartaoVVC"))
                _retParametros.PrepFatCartaoVRC = IIf(IsDBNull(rdr("PrepFatCartaoVRC")), "", rdr("PrepFatCartaoVRC"))
                _retParametros.PrepFatCartaoVAC = IIf(IsDBNull(rdr("PrepFatCartaoVAC")), "", rdr("PrepFatCartaoVAC"))
                _retParametros.PrepFatCartaoVVCD = IIf(IsDBNull(rdr("PrepFatCartaoVVCD")), "", rdr("PrepFatCartaoVVCD"))
                _retParametros.PrepFatCartaoVRCD = IIf(IsDBNull(rdr("PrepFatCartaoVRCD")), "", rdr("PrepFatCartaoVRCD"))
                _retParametros.PrepFatCartaoVACD = IIf(IsDBNull(rdr("PrepFatCartaoVACD")), "", rdr("PrepFatCartaoVACD"))
                _retParametros.NumCtaSisPag = IIf(IsDBNull(rdr("NumCtaSisPag")), "", rdr("NumCtaSisPag"))
                _retParametros.PrepFatVA = IIf(IsDBNull(rdr("PrepFatVA")), "", rdr("PrepFatVA"))
                _retParametros.PrepFatIA = IIf(IsDBNull(rdr("PrepFatIA")), "", rdr("PrepFatIA"))
                _retParametros.PrepFatAA = IIf(IsDBNull(rdr("PrepFatAA")), "", rdr("PrepFatAA"))
                _retParametros.PrepFatMA = IIf(IsDBNull(rdr("PrepFatMA")), "", rdr("PrepFatMA"))
                _retParametros.PrepFatTA = IIf(IsDBNull(rdr("PrepFatTA")), "", rdr("PrepFatTA"))
                _retParametros.PrepFatVAD = IIf(IsDBNull(rdr("PrepFatVAD")), "", rdr("PrepFatVAD"))
                _retParametros.PrepFatIAD = IIf(IsDBNull(rdr("PrepFatIAD")), "", rdr("PrepFatIAD"))
                _retParametros.PrepFatAAD = IIf(IsDBNull(rdr("PrepFatAAD")), "", rdr("PrepFatAAD"))
                _retParametros.PrepFatMAD = IIf(IsDBNull(rdr("PrepFatMAD")), "", rdr("PrepFatMAD"))
                _retParametros.PrepFatTAD = IIf(IsDBNull(rdr("PrepFatTAD")), "", rdr("PrepFatTAD"))
                _retParametros.PrepFatCartaoMVA = IIf(IsDBNull(rdr("PrepFatCartaoMVA")), "", rdr("PrepFatCartaoMVA"))
                _retParametros.PrepFatCartaoMRA = IIf(IsDBNull(rdr("PrepFatCartaoMRA")), "", rdr("PrepFatCartaoMRA"))
                _retParametros.PrepFatCartaoMAA = IIf(IsDBNull(rdr("PrepFatCartaoMAA")), "", rdr("PrepFatCartaoMAA"))
                _retParametros.PrepFatCartaoMVAD = IIf(IsDBNull(rdr("PrepFatCartaoMVAD")), "", rdr("PrepFatCartaoMVAD"))
                _retParametros.PrepFatCartaoMRAD = IIf(IsDBNull(rdr("PrepFatCartaoMRAD")), "", rdr("PrepFatCartaoMRAD"))
                _retParametros.PrepFatCartaoMAAD = IIf(IsDBNull(rdr("PrepFatCartaoMAAD")), "", rdr("PrepFatCartaoMAAD"))
                _retParametros.PrepFatCartaoVVA = IIf(IsDBNull(rdr("PrepFatCartaoVVA")), "", rdr("PrepFatCartaoVVA"))
                _retParametros.PrepFatCartaoVRA = IIf(IsDBNull(rdr("PrepFatCartaoVRA")), "", rdr("PrepFatCartaoVRA"))
                _retParametros.PrepFatCartaoVAA = IIf(IsDBNull(rdr("PrepFatCartaoVAA")), "", rdr("PrepFatCartaoVAA"))
                _retParametros.PrepFatCartaoVVAD = IIf(IsDBNull(rdr("PrepFatCartaoVVAD")), "", rdr("PrepFatCartaoVVAD"))
                _retParametros.PrepFatCartaoVRAD = IIf(IsDBNull(rdr("PrepFatCartaoVRAD")), "", rdr("PrepFatCartaoVRAD"))
                _retParametros.PrepFatCartaoVAAD = IIf(IsDBNull(rdr("PrepFatCartaoVAAD")), "", rdr("PrepFatCartaoVAAD"))
                _retParametros.DtEncerrContabil = IIf(IsDBNull(rdr("DtEncerrContabil")), Nothing, rdr("DtEncerrContabil"))
                _retParametros.PrepFatEmp = rdr("PrepFatEmp").ToString()
                _retParametros.PrepFatMon = rdr("PrepFatMon").ToString()
                _retParametros.PrepFatDoa = rdr("PrepFatDoa").ToString()
                _retParametros.PrepFatVenST = rdr("PrepFatVenST").ToString()
                _retParametros.PrepFatVen = rdr("PrepFatVen").ToString()
                _retParametros.PrepFatAde = rdr("PrepFatAde").ToString()
                _retParametros.PrepFatIns = rdr("PrepFatIns").ToString()

                _retParametros.NumArqRemCobExt = rdr("NumArqRemCobExt").ToString()
                _retParametros.ESenderNameCobExt = rdr("ESenderNameCobExt").ToString()
                _retParametros.ESenderEmailCobExt = rdr("ESenderEmailCobExt").ToString()
                _retParametros.ESubjectCobExt = rdr("ESubjectCobExt").ToString()

                _retParametros.VlrTxVisita = rdr("VlrTxVisita").ToString()
                _retParametros.VlrRepasseInst = rdr("VlrRepasseInst").ToString()
                _retParametros.VlrConexao = rdr("VlrConexao").ToString()
                _retParametros.VlrVistoria = rdr("VlrVistoria").ToString()
                _retParametros.VlrRepasseCon = rdr("VlrRepasseCon").ToString()
                _retParametros.VlrMinEquip = rdr("VlrMinEquip").ToString()
                _retParametros.VlrMaxDescMObra = rdr("VlrMaxDescMObra").ToString()
                _retParametros.VlrvistoriaMoto = rdr("VlrvistoriaMoto").ToString()
                _retParametros.CodCpgt = rdr("CodCpgt").ToString()
                _retParametros.VlrDescTxCon = rdr("VlrDescTxCon").ToString()
                _retParametros.BDLiberado = rdr("BDLiberado").ToString()

                _retParametros.EventosCentralConfPath = IIf(IsDBNull(rdr("EventosCentralConfPath")), Nothing, rdr("EventosCentralConfPath"))

                _retParametros.ESenderNameConfVda = rdr("ESenderNameConfVda").ToString()
                _retParametros.ESenderEmailConfVda = rdr("ESenderEmailConfVda").ToString()
                _retParametros.ESubjectConfVda = rdr("ESubjectConfVda").ToString()

                'EMAIL DO SUPORTE 11/04/2013 - FERNANDO
                _retParametros.ESenderSuporte = rdr("ESenderSuporte").ToString()

                'TELA DE CONFIGURAÇÃO DE E-MAILS - FERNANDO - 22/04/2013
                _retParametros.ESenderEmail = rdr("ESenderEmail").ToString()
                _retParametros.ESubject = rdr("ESubject").ToString()
                _retParametros.EmailOSVistRJ = rdr("EmailOSVistRJ").ToString()
                _retParametros.ESubjectInst = rdr("ESubjectInst").ToString()
                _retParametros.ESubjectWorking = rdr("ESubjectWorking").ToString()
                _retParametros.ESubjectWelcome = rdr("ESubjectWelcome").ToString()
                _retParametros.ESenderNameLV = rdr("ESenderNameLV").ToString()
                _retParametros.ESenderEmailLV = rdr("ESenderEmailLV").ToString()
                _retParametros.ESubjectLV = rdr("ESubjectLV").ToString()

                'Email do aprovador de solicitações de hora extra - Renato - 20/05/2014
                _retParametros.EAprovSolicHE = rdr("EAprovSolicHE").ToString()

                _retParametros.PrepFatVBCFTV = IIf(IsDBNull(rdr("PrepFatVBCFTV")), "", rdr("PrepFatVBCFTV"))
                _retParametros.PrepFatVACFTV = IIf(IsDBNull(rdr("PrepFatVACFTV")), "", rdr("PrepFatVACFTV"))
                _retParametros.PrepFatVCCFTV = IIf(IsDBNull(rdr("PrepFatVCCFTV")), "", rdr("PrepFatVCCFTV"))

                _retParametros.PrepFatVBDCFTV = IIf(IsDBNull(rdr("PrepFatVBDCFTV")), "", rdr("PrepFatVBDCFTV"))
                _retParametros.PrepFatVADCFTV = IIf(IsDBNull(rdr("PrepFatVADCFTV")), "", rdr("PrepFatVADCFTV"))
                _retParametros.PrepFatVCDCFTV = IIf(IsDBNull(rdr("PrepFatVCDCFTV")), "", rdr("PrepFatVCDCFTV"))

                _retParametros.PrepFatIBCFTV = IIf(IsDBNull(rdr("PrepFatIBCFTV")), "", rdr("PrepFatIBCFTV"))
                _retParametros.PrepFatIACFTV = IIf(IsDBNull(rdr("PrepFatIACFTV")), "", rdr("PrepFatIACFTV"))
                _retParametros.PrepFatICCFTV = IIf(IsDBNull(rdr("PrepFatICCFTV")), "", rdr("PrepFatICCFTV"))

                _retParametros.PrepFatIBDCFTV = IIf(IsDBNull(rdr("PrepFatIBDCFTV")), "", rdr("PrepFatIBDCFTV"))
                _retParametros.PrepFatIADCFTV = IIf(IsDBNull(rdr("PrepFatIADCFTV")), "", rdr("PrepFatIADCFTV"))
                _retParametros.PrepFatICDCFTV = IIf(IsDBNull(rdr("PrepFatICDCFTV")), "", rdr("PrepFatICDCFTV"))

                _retParametros.HelpDeskPath = IIf(IsDBNull(rdr("HelpDeskPath")), "", rdr("HelpDeskPath"))

                _retParametros.PrepFatAdiantamento = rdr("PrepFatAdiantamento").ToString()

                _retParametros.PrepFatBAdiantamento = IIf(IsDBNull(rdr("PrepFatBAdiantamento")), "", rdr("PrepFatBAdiantamento"))
                _retParametros.PrepFatAAdiantamento = IIf(IsDBNull(rdr("PrepFatAAdiantamento")), "", rdr("PrepFatAAdiantamento"))
                _retParametros.PrepFatCAdiantamento = IIf(IsDBNull(rdr("PrepFatCAdiantamento")), "", rdr("PrepFatCAdiantamento"))

                _retParametros.PrepFatBDAdiantamento = IIf(IsDBNull(rdr("PrepFatBDAdiantamento")), "", rdr("PrepFatBDAdiantamento"))
                _retParametros.PrepFatADAdiantamento = IIf(IsDBNull(rdr("PrepFatADAdiantamento")), "", rdr("PrepFatADAdiantamento"))
                _retParametros.PrepFatCDAdiantamento = IIf(IsDBNull(rdr("PrepFatCDAdiantamento")), "", rdr("PrepFatCDAdiantamento"))

                _retParametros.CREDENTIAL_MB_IND = IIf(IsDBNull(rdr("CREDENTIAL_MB_IND")), "", rdr("CREDENTIAL_MB_IND"))

                _retParametros.CaminhoFotoVendedor = IIf(IsDBNull(rdr("CaminhoFotoVendedor")), "", rdr("CaminhoFotoVendedor"))

                _retParametros.PercRetPISVenda = CDec(IIf(IsDBNull(rdr("PercRetPISVenda")), "", rdr("PercRetPISVenda")).ToString())
                _retParametros.PercRetCOFINSVenda = CDec(IIf(IsDBNull(rdr("PercRetCOFINSVenda")), "", rdr("PercRetCOFINSVenda")).ToString())
                '_retParametros.LimiteCompra = CDbl(rdr("LimiteCompra"))
                _retParametros.UpdatePathVerisure = IIf(IsDBNull(rdr("UpdatePathVerisure")), "", rdr("UpdatePathVerisure"))

                _retParametros.CaminhoFotoVendedorVerisure = IIf(IsDBNull(rdr("CaminhoFotoVendedorVerisure")), "", rdr("CaminhoFotoVendedorVerisure"))

                'Novos campos para conexão promocional
                _retParametros.vlrConPromo = (rdr("VlrConPromo"))
                _retParametros.vlrRepasseInstConPromo = (rdr("VlrRepasseInstConPromo"))
                _retParametros.vlrDescConPromo = (rdr("VlrDescConPromo"))
                _retParametros.LoteRecebimentoNAV = IIf(IsDBNull(rdr("LoteRecebimentoNAV")), "", rdr("LoteRecebimentoNAV"))

                _retParametros.NumArqRemBVS = IIf(IsDBNull(rdr("NumArqRemBVS")), "", rdr("NumArqRemBVS"))
                _retParametros.NumArqRemSerasaTELE = IIf(IsDBNull(rdr("NumArqRemSerasaTELE")), "", rdr("NumArqRemSerasaTELE"))

                _retParametros.UrlEnvioBoleto = IIf(IsDBNull(rdr("UrlEnvioBoleto")), "", rdr("UrlEnvioBoleto"))

                _retParametros.ContratoPathVERI = IIf(IsDBNull(rdr("ContratoPathVERI")), "", rdr("ContratoPathVERI"))

                _retParametros.isFaturamentoEmProgresso = IIf(IsDBNull(rdr("isFaturamentoEmProgresso")), "", rdr("isFaturamentoEmProgresso"))

                _retParametros.Versao = IIf(IsDBNull(rdr("Versao")), "", rdr("Versao"))

                _retParametros.Sucesso = True
                _retParametros.TipoErro = DadosGenericos.TipoErro.None
            Else
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_PESQUISAPARAMETRO.Descricao
            End If


            rdr.Close()

        Catch ex As DadosInexistentesTabelaParametroException
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_PESQUISAPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_PESQUISAPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: PesquisaParametro(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Catch ex As Exception
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_PESQUISAPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_PESQUISAPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: PesquisaParametro(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            Command.Dispose()
            connection.Close()
            connection.Dispose()
        End Try

        Return _retParametros

    End Function


    Public Function BuscaHorarioSinalTeste() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_BuscaHorarioSinalTeste", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()

                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.SinalTesteF = rdr("SinalTesteF").ToString()
                        _retParametros.SinalTesteJ = rdr("SinalTesteJ").ToString()

                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORARIOSINALTESTE.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORARIOSINALTESTE.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaHorarioSinalTeste(2)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORARIOSINALTESTE.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORARIOSINALTESTE.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaHorarioSinalTeste(2)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using
    End Function


    Public Function PesquisaDtEncContabilParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_PesquisaDtEncContabilParametro", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()

                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.DtEncerrContabil = IIf(rdr("DtEncerrContabil").ToString() = "", Nothing, rdr("DtEncerrContabil"))

                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORARIOSINALTESTE.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORARIOSINALTESTE.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: PesquisaDtEncContabilParametro(3)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORARIOSINALTESTE.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORARIOSINALTESTE.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: PesquisaDtEncContabilParametro(3)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using
    End Function



    Public Function PesquisaNumNFNumNFeParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_BuscaNumNFNumNFeParametro", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()

                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.NumNF = rdr("NumNFServ").ToString()
                        _retParametros.NumNFe = rdr("NumNFe").ToString()


                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_PESQUISANUMNFNUMNFEPARAMETRO.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_PESQUISANUMNFNUMNFEPARAMETRO.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: PesquisaNumNFNumNFeParametro(4)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_PESQUISANUMNFNUMNFEPARAMETRO.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_PESQUISANUMNFNUMNFEPARAMETRO.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: PesquisaNumNFNumNFeParametro(4)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using
    End Function


    Public Function ConsultaSerieNFServUFParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_ConsultaSerieNFServUFParametro", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()

                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.SerieNFServ = rdr("SerieNFServ").ToString()
                        _retParametros.UF = rdr("UF").ToString()


                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_CONSULTASERIENFSERVUFPARAMETRO.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_CONSULTASERIENFSERVUFPARAMETRO.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: ConsultaSerieNFServUFParametro(5)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_CONSULTASERIENFSERVUFPARAMETRO.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_CONSULTASERIENFSERVUFPARAMETRO.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: ConsultaSerieNFServUFParametro(5)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using
    End Function



    Public Function BuscaDadosFatParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_BuscaDadosFatParametro", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()

                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.PrepFatVen = rdr("PrepFatVen").ToString()
                        _retParametros.NumNF = rdr("NumNF").ToString()
                        _retParametros.NumNFServ = rdr("NumNFServ").ToString()
                        _retParametros.PrepFatEmp = rdr("PrepFatEmp").ToString()
                        _retParametros.PrepFatIns = rdr("PrepFatIns").ToString()
                        _retParametros.PrepFatAde = rdr("PrepFatAde").ToString()
                        _retParametros.NumNFe = rdr("NumNFe").ToString()
                        _retParametros.TravaEstoque = rdr("TravaEstoque").ToString()
                        _retParametros.PrepFatMon = rdr("PrepFatMon").ToString()
                        _retParametros.PrepFatDoa = rdr("PrepFatDoa").ToString()
                        _retParametros.PrepFatVenST = rdr("PrepFatVenST").ToString()


                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCADADOSFATPARAMETRO.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCADADOSFATPARAMETRO.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaDadosFatParametro(6)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCADADOSFATPARAMETRO.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCADADOSFATPARAMETRO.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaDadosFatParametro(6)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using
    End Function

    Public Function BuscaNfParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_BuscaNumNfParametro", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()
                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.NumNF = rdr("NumNF").ToString()

                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCADADOSFATPARAMETRO.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCADADOSFATPARAMETRO.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaDadosFatParametro(6)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCADADOSFATPARAMETRO.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCADADOSFATPARAMETRO.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaDadosFatParametro(6)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using

    End Function


    Public Function BuscaNumNFeSerieParametro(ByVal Connection As SqlConnection,
                                              ByVal Transaction As SqlTransaction) As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()


        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaNumNFeSerieParametro", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _retParametros.NumNF = rdr("NumNF").ToString()
                _retParametros.SerieNF = rdr("SerieNF").ToString()

                _retParametros.Sucesso = True
                _retParametros.TipoErro = DadosGenericos.TipoErro.None

            Else
                Throw New DadosInexistentesTabelaParametroException()
            End If
            rdr.Close()
        Catch ex As DadosInexistentesTabelaParametroException
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCANUMNFESERIEPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCANUMNFESERIEPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaNumNFeSerieParametro(7)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Catch ex As Exception
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCANUMNFESERIEPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCANUMNFESERIEPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaNumNFeSerieParametro(7)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try
        Return _retParametros
    End Function



    Public Function BuscaNumNFNumNFServNumNFeParametro(ByVal Connection As SqlConnection,
                ByVal Transaction As SqlTransaction) As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()


        Try
            ''Informa a procedure
            Using Command As SqlCommand = New SqlCommand("P_BuscaNumNFNumNFServNumNFeParametro", Connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Query
                Command.Transaction = Transaction

                rdr = Command.ExecuteReader()

                If rdr.HasRows Then
                    rdr.Read()

                    _retParametros.NumNF = rdr("NumNF").ToString()
                    _retParametros.NumNFe = rdr("NumNFe").ToString()
                    _retParametros.NumNFServ = rdr("NumNFServ").ToString()

                    _retParametros.Sucesso = True
                    _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    rdr.Close()
                Else
                    Throw New DadosInexistentesTabelaParametroException()
                End If

                rdr.Close()
            End Using

        Catch ex As DadosInexistentesTabelaParametroException
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaNumNFNumNFServNumNFeParametro(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Catch ex As Exception
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaNumNFNumNFServNumNFeParametro(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try
        Return _retParametros
    End Function



    Public Function BuscaAliqIPIParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()
        Dim Connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Try
            ''Informa a procedure
            Using Command As SqlCommand = New SqlCommand("P_BuscaAliqIPIParametro", Connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Query

                Connection.Open()

                rdr = Command.ExecuteReader()

                If rdr.HasRows Then
                    rdr.Read()

                    _retParametros.AliqIPI = rdr("AliqIPI").ToString()

                    _retParametros.Sucesso = True
                    _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    rdr.Close()
                Else
                    Throw New DadosInexistentesTabelaParametroException()
                End If


            End Using

        Catch ex As DadosInexistentesTabelaParametroException
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCAALIQIPIPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCAALIQIPIPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
        Catch ex As Exception
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCAALIQIPIPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCAALIQIPIPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaAliqIPIParametro(9)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            Connection.Close()
        End Try
        Return _retParametros
    End Function


    Public Function BuscaNumNFNumNFServNumNFeParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Try
            ''Informa a procedure
            Using Command As SqlCommand = New SqlCommand("P_BuscaNumNFNumNFServNumNFeParametro", Connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Query

                Connection.Open()

                rdr = Command.ExecuteReader()

                If rdr.HasRows Then
                    rdr.Read()

                    _retParametros.NumNF = rdr("NumNF").ToString()
                    _retParametros.NumNFe = rdr("NumNFe").ToString()
                    _retParametros.NumNFServ = rdr("NumNFServ").ToString()

                    _retParametros.Sucesso = True
                    _retParametros.TipoErro = DadosGenericos.TipoErro.None
                Else
                    Throw New DadosInexistentesTabelaParametroException()
                End If

                rdr.Close()
            End Using

        Catch ex As DadosInexistentesTabelaParametroException
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaNumNFNumNFServNumNFeParametro(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Catch ex As Exception
            _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Descricao + ex.Message
            _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Id
            _retParametros.Sucesso = False
            _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaNumNFNumNFServNumNFeParametro(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            Connection.Close()
        End Try
        Return _retParametros
    End Function

    Public Function GeraUltNumParametro() As Parametros
        Dim rdr As SqlDataReader
        Dim _Parametros As New Parametros()
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Try
            ''Informa a procedure
            Using Command As SqlCommand = New SqlCommand("P_GeraUltNumParametro", Connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Query

                Connection.Open()

                rdr = Command.ExecuteReader()

                If rdr.HasRows Then
                    rdr.Read()

                    _Parametros.UltNum = rdr("UltNum").ToString()

                    _Parametros.Sucesso = True
                    _Parametros.TipoErro = DadosGenericos.TipoErro.None
                Else
                    Throw New DadosInexistentesTabelaParametroException()
                End If

                rdr.Close()
            End Using

        Catch ex As DadosInexistentesTabelaParametroException
            _Parametros.MsgErro = ErrorConstants.EXCEPTION_METODO_GERAULTNUMPARAMETRO.Descricao + ex.Message
            _Parametros.NumErro = ErrorConstants.EXCEPTION_METODO_GERAULTNUMPARAMETRO.Id
            _Parametros.Sucesso = False
            _Parametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Parametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Parametros.NumErro, _Parametros.MsgErro, _Parametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: GeraUltNumParametro(9)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Catch ex As Exception
            _Parametros.MsgErro = ErrorConstants.EXCEPTION_METODO_GERAULTNUMPARAMETRO.Descricao + ex.Message
            _Parametros.NumErro = ErrorConstants.EXCEPTION_METODO_GERAULTNUMPARAMETRO.Id
            _Parametros.Sucesso = False
            _Parametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Parametros.NumErro, _Parametros.MsgErro, _Parametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: GeraUltNumParametro(9)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            Connection.Close()
        End Try
        Return _Parametros
    End Function

    'Public Function BuscaNumNFNumNFServNumNFeParametro() As Parametros
    '    Dim rdr As SqlDataReader
    '    Dim _retParametros As New Parametros()
    '    Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

    '    Try
    '        ''Informa a procedure
    '        Using Command As SqlCommand = New SqlCommand("P_BuscaNumNFNumNFServNumNFeParametro", Connection)
    '            command.CommandType = CommandType.StoredProcedure
    '            command.CommandTimeout = DadosGenericos.Timeout.Query

    '            Connection.Open()

    '            rdr = Command.ExecuteReader()

    '            If rdr.HasRows Then
    '                rdr.Read()

    '                _retParametros.NumNF = rdr("NumNF").ToString()
    '                _retParametros.NumNFe = rdr("NumNFe").ToString()
    '                _retParametros.NumNFServ = rdr("NumNFServ").ToString()

    '                _retParametros.Sucesso = True
    '                _retParametros.TipoErro = DadosGenericos.TipoErro.None
    '            Else
    '                Throw New DadosInexistentesTabelaParametroException()
    '            End If

    '            rdr.Close()
    '        End Using

    '    Catch ex As DadosInexistentesTabelaParametroException
    '        _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Descricao + ex.Message
    '        _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Id
    '        _retParametros.Sucesso = False
    '        _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaNumNFNumNFServNumNFeParametro(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    Catch ex As Exception
    '        _retParametros.MsgErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Descricao + ex.Message
    '        _retParametros.NumErro = ErrorConstants.EXCEPTION_BUSCANUMNFNUMNFSERVNUMNFEPARAMETRO.Id
    '        _retParametros.Sucesso = False
    '        _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, "Projeto: ParametroBC - Classe: ConsultaParametroGeral - Função: BuscaNumNFNumNFServNumNFeParametro(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    Finally
    '        Connection.Close()
    '    End Try
    '    Return _retParametros
    'End Function

    ''' <summary>
    ''' Busca os valores referentes as retenções da nova CONFINS, como CSLL (contribuição social sobre lucro líquido, PIS/PASEP, etc)
    ''' </summary>
    ''' <returns>Retorna um objeto do tipo Parametros com apenas as alíquotas referentes instanciadas.</returns>
    ''' <remarks></remarks>
    Public Function BuscaAliquotasNovaCofins() As Parametros

        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_BuscaAliquotasNovaCofins", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()

                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.PercRetCSLL = CDec(rdr("PercRetCSLL").ToString())
                        _retParametros.PercRetCofins = CDec(rdr("PercRetCOFINS").ToString())
                        _retParametros.PercRetPis = CDec(rdr("PercRetPIS").ToString())
                        _retParametros.VlrPisoCOFINS = CDec(rdr("VlrPisoCOFINS").ToString())
                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASNOVACOFINS.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASNOVACOFINS.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASNOVACOFINS.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASNOVACOFINS.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using


    End Function


    ''' <summary>
    ''' Busca os valores referentes as retenções de IR
    ''' </summary>
    ''' <returns>Retorna um objeto do tipo Parametros com apenas as alíquotas referentes instanciadas.</returns>
    ''' <remarks></remarks>
    Public Function BuscaAliquotasRetencaoIr() As Parametros

        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_BuscaAliquotasRetIr", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()

                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.AliqRetIr = CDec(rdr("AliqRetIr").ToString())
                        _retParametros.VlrMinRetIr = CDec(rdr("VlrMinRetIr").ToString())
                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASNOVACOFINS.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASNOVACOFINS.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASNOVACOFINS.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASNOVACOFINS.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using


    End Function

    ''' <summary>
    ''' Busca os valores referentes as retenções de IR
    ''' </summary>
    ''' <returns>Retorna um objeto do tipo Parametros com apenas as alíquotas referentes instanciadas.</returns>
    ''' <remarks></remarks>
    Public Function BuscaAliquotasRetencaoINSS() As Parametros

        Dim rdr As SqlDataReader
        Dim _retParametros As New Parametros()

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                ''Informa a procedure
                Using Command As SqlCommand = New SqlCommand("P_BuscaAliquotasRetINSS", connection)
                    Command.CommandType = CommandType.StoredProcedure
                    Command.CommandTimeout = DadosGenericos.Timeout.Query

                    connection.Open()

                    rdr = Command.ExecuteReader()

                    If rdr.HasRows Then
                        rdr.Read()
                        _retParametros.AliqRetINSS = CDec(rdr("AliqRetINSS").ToString())
                        _retParametros.Sucesso = True
                        _retParametros.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        Throw New DadosInexistentesTabelaParametroException()
                    End If


                End Using

            Catch ex As DadosInexistentesTabelaParametroException
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASRETINSS.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASRETINSS.Id
                _retParametros.Sucesso = False
                _retParametros.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
            Catch ex As Exception
                _retParametros.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASRETINSS.Descricao + ex.Message
                _retParametros.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAALIQUOTASRETINSS.Id
                _retParametros.Sucesso = False
                _retParametros.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retParametros.NumErro, _retParametros.MsgErro, _retParametros.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            End Try
            Return _retParametros
        End Using


    End Function

    ''' <summary>
    ''' Porcedure que gera o ultimo cod orc buscando o max + 1 da parametro e fazendo um update para guardar o valor gerado.
    ''' </summary>
    ''' <param name="connection"></param>
    ''' <param name="trans"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GeraUltCodOrc(ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Parametros

        Dim _Parametro As New Parametros
        Dim command As New SqlCommand("P_GeraUltCodOrc", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim rdr As SqlDataReader

        Try
            command.Transaction = trans

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                rdr.Read()

                _Parametro.UltCodOrc = IIf(IsDBNull(rdr("UltCodOrc")), Nothing, rdr("UltCodOrc"))

                _Parametro.Sucesso = True
                _Parametro.TipoErro = DadosGenericos.TipoErro.None


            Else

                _Parametro.Sucesso = False
                _Parametro.TipoErro = DadosGenericos.TipoErro.Funcional
                _Parametro.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Parametro.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If
            rdr.Close()
        Catch ex As Exception

            _Parametro.Sucesso = False
            _Parametro.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Parametro.MsgErro = ErrorConstants.EXCEPTION_METODO_GERAULTCODORC.Descricao & ex.Message
            _Parametro.NumErro = ErrorConstants.EXCEPTION_METODO_GERAULTCODORC.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Parametro.NumErro, _Parametro.MsgErro, _Parametro.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            command.Dispose()
        End Try

        Return _Parametro
    End Function


    Public Function BuscaProximoUltCodLoteParametro(ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Parametros

        Dim _Parametro As New Parametros
        Dim command As New SqlCommand("P_BuscaProximoUltCodLoteParametro", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim rdr As SqlDataReader

        Try
            command.Transaction = trans

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                rdr.Read()

                _Parametro.UltCodLote = IIf(IsDBNull(rdr("UltCodLote")), Nothing, rdr("UltCodLote"))

                _Parametro.Sucesso = True
                _Parametro.TipoErro = DadosGenericos.TipoErro.None
            Else
                _Parametro.Sucesso = False
                _Parametro.TipoErro = DadosGenericos.TipoErro.Funcional
                _Parametro.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Parametro.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If
            rdr.Close()
        Catch ex As Exception
            _Parametro.Sucesso = False
            _Parametro.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Parametro.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAPROXIMOULTCODLOTEPARAMETRO.Descricao & ex.Message
            _Parametro.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAPROXIMOULTCODLOTEPARAMETRO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Parametro.NumErro, _Parametro.MsgErro, _Parametro.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            command.Dispose()
        End Try

        Return _Parametro
    End Function

    Public Function BuscaHoraBloqueioSolicitacaoCompra() As Parametros

        Dim _Parametro As New Parametros()

        Try
            Using con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
                con.Open()

                Dim cmd As New SqlCommand("P_BuscaHoraBloqueioSolicitacaoCompra", con)
                cmd.CommandType = CommandType.StoredProcedure

                Dim rdr As SqlDataReader = cmd.ExecuteReader()

                If (rdr.HasRows) Then
                    While (rdr.Read())
                        _Parametro.BloqHoraSolCpa = rdr("BloqHoraSolCpa")
                        _Parametro.Sucesso = True
                        _Parametro.TipoErro = DadosGenericos.TipoErro.None
                    End While
                Else
                    _Parametro.Sucesso = False
                    _Parametro.TipoErro = DadosGenericos.TipoErro.Funcional
                    _Parametro.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                    _Parametro.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                    _Parametro.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                End If

                con.Close()
            End Using
        Catch ex As Exception
            Dim _Retorno As New Parametros()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORABLOQUEIOSOLICITACAOCOMPRA.Id
                .MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAHORABLOQUEIOSOLICITACAOCOMPRA.Descricao + ex.Message
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Parametro
    End Function


    Public Function BuscaBuscaValorPlantao() As Parametros

        Dim _Parametro As New Parametros()

        Try
            Using con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
                con.Open()

                Dim cmd As New SqlCommand("P_BuscaValorPlantao", con)
                cmd.CommandType = CommandType.StoredProcedure

                Dim rdr As SqlDataReader = cmd.ExecuteReader()

                If (rdr.HasRows) Then
                    While (rdr.Read())
                        _Parametro.VlrPlantaoManut = rdr("VlrPlantaoManut")
                        _Parametro.Sucesso = True
                        _Parametro.TipoErro = DadosGenericos.TipoErro.None
                    End While
                Else
                    _Parametro.Sucesso = False
                    _Parametro.TipoErro = DadosGenericos.TipoErro.Funcional
                    _Parametro.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                    _Parametro.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                    _Parametro.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                End If

                con.Close()
            End Using
        Catch ex As Exception
            Dim _Retorno As New Parametros()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
                .MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Descricao + ex.Message
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Parametro
    End Function

    Public Function BuscaProtocoloHistContato(ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Parametros

        Dim _Parametro As New Parametros()

        Try

            Dim cmd As New SqlCommand("P_BuscaProtocoloHistContato", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Transaction = trans

            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            If (rdr.HasRows) Then
                While (rdr.Read())
                    _Parametro.ProtocoloHistContato = rdr("Protocolo")
                    _Parametro.Sucesso = True
                    _Parametro.TipoErro = DadosGenericos.TipoErro.None
                End While
            Else
                _Parametro.Sucesso = False
                _Parametro.TipoErro = DadosGenericos.TipoErro.Funcional
                _Parametro.MsgErro = "Nenhum protocolo encontrado."
                _Parametro.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Parametro.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception
            Dim _Retorno As New Parametros()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
                .MsgErro = ex.Message
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Parametro
    End Function

    'Public Function BuscaConfigDebCredMonit(ByVal CodDepto As String, ByVal TpCliente As String, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As ConfigDebCredMonit

    '    Dim Configuracao As New ConfigDebCredMonit()

    '    Try

    '        Dim cmd As New SqlCommand("P_BuscaConfigDebCredMonit", conn)
    '        cmd.CommandType = CommandType.StoredProcedure

    '        cmd.Transaction = trans

    '        cmd.Parameters.Add(New SqlParameter("@CodDepto", CodDepto))
    '        cmd.Parameters.Add(New SqlParameter("@TpCliente", TpCliente))

    '        Dim rdr As SqlDataReader = cmd.ExecuteReader()

    '        If (rdr.HasRows) Then
    '            While (rdr.Read())
    '                Configuracao.QtdMeses1Est1Solic = rdr("QtdMeses1Est1Solic")
    '                Configuracao.QtdMeses1Est2Solic = rdr("QtdMeses1Est2Solic")
    '                Configuracao.QtdMeses2Est1Solic = rdr("QtdMeses2Est1Solic")
    '                Configuracao.QtdMeses2Est2Solic = rdr("QtdMeses2Est2Solic")
    '                Configuracao.QtdMeses3Est1Solic = rdr("QtdMeses3Est1Solic")
    '                Configuracao.QtdMeses3Est2Solic = rdr("QtdMeses3Est2Solic")
    '                Configuracao.QtdMeses4Est1Solic = rdr("QtdMeses4Est1Solic")
    '                Configuracao.QtdMeses4Est2Solic = rdr("QtdMeses4Est2Solic")
    '                Configuracao.QtdMeses5Est1Solic = rdr("QtdMeses5Est1Solic")
    '                Configuracao.QtdMeses5Est2Solic = rdr("QtdMeses5Est2Solic")
    '                Configuracao.PorcentagemMax = rdr("PorcentagemMax")
    '                Configuracao.Sucesso = True
    '                Configuracao.TipoErro = DadosGenericos.TipoErro.None
    '            End While
    '        Else
    '            Configuracao.Sucesso = False
    '            Configuracao.TipoErro = DadosGenericos.TipoErro.Funcional
    '            Configuracao.MsgErro = "Nenhuma configuração encontrada."
    '            Configuracao.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
    '            Configuracao.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
    '        End If
    '        rdr.Close()
    '    Catch ex As Exception
    '        Dim _Retorno As New Parametros()
    '        With _Retorno
    '            .Sucesso = False
    '            .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
    '            .MsgErro = ex.Message
    '            .TipoErro = DadosGenericos.TipoErro.Arquitetura
    '            .ImagemErro = DadosGenericos.ImagemRetorno.Erro
    '        End With

    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

    '    End Try

    '    Return Configuracao
    'End Function

    'Public Function BuscaConfigDebCredMonit(ByVal CodDepto As String, ByVal TpCliente As String) As ConfigDebCredMonit

    '    Dim Configuracao As New ConfigDebCredMonit

    '    Try
    '        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '        Dim cmd As New SqlCommand("P_BuscaConfigDebCredMonit", connection)
    '        cmd.CommandType = CommandType.StoredProcedure

    '        cmd.Parameters.Add(New SqlParameter("@CodDepto", CodDepto))
    '        cmd.Parameters.Add(New SqlParameter("@TpCliente", TpCliente))

    '        connection.Open()

    '        Dim rdr As SqlDataReader = cmd.ExecuteReader()

    '        If (rdr.HasRows) Then
    '            While (rdr.Read())
    '                Configuracao.QtdMeses1Est1Solic = rdr("QtdMeses1Est1Solic")
    '                Configuracao.QtdMeses1Est2Solic = rdr("QtdMeses1Est2Solic")
    '                Configuracao.QtdMeses2Est1Solic = rdr("QtdMeses2Est1Solic")
    '                Configuracao.QtdMeses2Est2Solic = rdr("QtdMeses2Est2Solic")
    '                Configuracao.QtdMeses3Est1Solic = rdr("QtdMeses3Est1Solic")
    '                Configuracao.QtdMeses3Est2Solic = rdr("QtdMeses3Est2Solic")
    '                Configuracao.QtdMeses4Est1Solic = rdr("QtdMeses4Est1Solic")
    '                Configuracao.QtdMeses4Est2Solic = rdr("QtdMeses4Est2Solic")
    '                Configuracao.QtdMeses5Est1Solic = rdr("QtdMeses5Est1Solic")
    '                Configuracao.QtdMeses5Est2Solic = rdr("QtdMeses5Est2Solic")
    '                Configuracao.PorcentagemMax = rdr("PorcentagemMax")
    '                Configuracao.Sucesso = True
    '                Configuracao.TipoErro = DadosGenericos.TipoErro.None
    '            End While
    '        Else
    '            Configuracao.Sucesso = False
    '            Configuracao.TipoErro = DadosGenericos.TipoErro.Funcional
    '            Configuracao.MsgErro = "Nenhuma configuração encontrada."
    '            Configuracao.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
    '            Configuracao.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
    '        End If
    '        rdr.Close()
    '    Catch ex As Exception
    '        Dim _Retorno As New Parametros()
    '        With _Retorno
    '            .Sucesso = False
    '            .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
    '            .MsgErro = ex.Message
    '            .TipoErro = DadosGenericos.TipoErro.Arquitetura
    '            .ImagemErro = DadosGenericos.ImagemRetorno.Erro
    '        End With

    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    End Try

    '    Return Configuracao
    'End Function


    'Public Function BuscaHTMLVerisystem(ByVal CodIntClie As String, ByVal Envio As String) As HTMLVerisystem

    '    Dim HTMLVerisystem As New HTMLVerisystem

    '    Try
    '        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '        Dim cmd As New SqlCommand("P_BuscaHTMLVerisystem", connection)
    '        cmd.CommandType = CommandType.StoredProcedure

    '        cmd.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
    '        cmd.Parameters.Add(New SqlParameter("@Envio", Envio))

    '        connection.Open()

    '        Dim rdr As SqlDataReader = cmd.ExecuteReader()

    '        If (rdr.HasRows) Then
    '            While (rdr.Read())
    '                HTMLVerisystem.ID = rdr("ID")
    '                HTMLVerisystem.Carteira = rdr("Carteira")
    '                HTMLVerisystem.Envio = rdr("Envio")
    '                HTMLVerisystem.Titulo = rdr("Titulo")
    '                HTMLVerisystem.NomeRemetente = rdr("NomeRemetente")
    '                HTMLVerisystem.HTML = rdr("HTML")
    '                HTMLVerisystem.Status = rdr("Status")
    '                'HTMLVerisystem.textoHTML = rdr("TEXTO") -- Modif. Felipe: 2021-09-23
    '                HTMLVerisystem.textoHTML = rdr("TEXTO").ToString()
    '                HTMLVerisystem.Sucesso = True
    '                HTMLVerisystem.TipoErro = DadosGenericos.TipoErro.None
    '            End While
    '        Else
    '            HTMLVerisystem.Sucesso = False
    '            HTMLVerisystem.TipoErro = DadosGenericos.TipoErro.Funcional
    '            HTMLVerisystem.MsgErro = "Nenhum HTML encontrado."
    '            HTMLVerisystem.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
    '            HTMLVerisystem.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
    '        End If
    '        rdr.Close()
    '    Catch ex As Exception
    '        Dim _Retorno As New Parametros()
    '        With _Retorno
    '            .Sucesso = False
    '            .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
    '            .MsgErro = ex.Message
    '            .TipoErro = DadosGenericos.TipoErro.Arquitetura
    '            .ImagemErro = DadosGenericos.ImagemRetorno.Erro
    '        End With

    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    End Try

    '    Return HTMLVerisystem
    'End Function

    'Public Function BuscaHTMLNotaFiscal(ByVal Id As Integer) As HTMLVerisystem

    '    Dim HTMLVerisystem As New HTMLVerisystem


    '    Try
    '        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '        Dim cmd As New SqlCommand("P_BuscaHTMLReenvioNotaFiscal", connection)
    '        cmd.CommandType = CommandType.StoredProcedure

    '        cmd.Parameters.Add(New SqlParameter("@IdHtml", Id))

    '        connection.Open()

    '        Dim rdr As SqlDataReader = cmd.ExecuteReader()

    '        If (rdr.HasRows) Then
    '            While (rdr.Read())
    '                HTMLVerisystem.ID = rdr("ID")
    '                HTMLVerisystem.Envio = rdr("Envio")
    '                HTMLVerisystem.Titulo = rdr("Titulo")
    '                HTMLVerisystem.NomeRemetente = rdr("NomeRemetente")
    '                HTMLVerisystem.HTML = rdr("HTML")
    '                HTMLVerisystem.Status = rdr("Status")
    '                HTMLVerisystem.Sucesso = True
    '                HTMLVerisystem.TipoErro = DadosGenericos.TipoErro.None
    '            End While
    '        Else
    '            HTMLVerisystem.Sucesso = False
    '            HTMLVerisystem.TipoErro = DadosGenericos.TipoErro.Funcional
    '            HTMLVerisystem.MsgErro = "Nenhum HTML encontrado."
    '            HTMLVerisystem.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
    '            HTMLVerisystem.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
    '        End If
    '        rdr.Close()
    '    Catch ex As Exception
    '        Dim _Retorno As New Parametros()
    '        With _Retorno
    '            .Sucesso = False
    '            .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
    '            .MsgErro = ex.Message
    '            .TipoErro = DadosGenericos.TipoErro.Arquitetura
    '            .ImagemErro = DadosGenericos.ImagemRetorno.Erro
    '        End With

    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    End Try

    '    Return HTMLVerisystem
    'End Function

    Public Function BuscaHTMLManual(ByVal Envio As String) As HTMLVerisystem

        Dim HTMLVerisystem As New HTMLVerisystem

        Try
            Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            Dim cmd As New SqlCommand("P_BuscaHTMLManual", connection)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.Add(New SqlParameter("@NomeArquivo", Envio))

            connection.Open()

            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            If (rdr.HasRows) Then
                While (rdr.Read())
                    HTMLVerisystem.ID = rdr("ID")
                    HTMLVerisystem.Envio = rdr("Envio")
                    HTMLVerisystem.Titulo = rdr("Titulo")
                    HTMLVerisystem.NomeRemetente = rdr("NomeRemetente")
                    HTMLVerisystem.HTML = rdr("HTML")
                    HTMLVerisystem.Status = rdr("Status")
                    HTMLVerisystem.Sucesso = True
                    HTMLVerisystem.TipoErro = DadosGenericos.TipoErro.None
                End While
            Else
                HTMLVerisystem.Sucesso = False
                HTMLVerisystem.TipoErro = DadosGenericos.TipoErro.Funcional
                HTMLVerisystem.MsgErro = "Nenhum HTML encontrado."
                HTMLVerisystem.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                HTMLVerisystem.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception
            Dim _Retorno As New Parametros()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
                .MsgErro = ex.Message
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return HTMLVerisystem
    End Function
    'Public Function BuscaHTMLVerisystemCodOrc(ByVal CodOrc As String, ByVal Envio As String) As HTMLVerisystem

    '    Dim HTMLVerisystem As New HTMLVerisystem

    '    Try
    '        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '        Dim cmd As New SqlCommand("P_BuscaHTMLVerisystemCodOrc", connection)
    '        cmd.CommandType = CommandType.StoredProcedure

    '        cmd.Parameters.Add(New SqlParameter("@CodOrc", CodOrc))
    '        cmd.Parameters.Add(New SqlParameter("@Envio", Envio))

    '        connection.Open()

    '        Dim rdr As SqlDataReader = cmd.ExecuteReader()

    '        If (rdr.HasRows) Then
    '            While (rdr.Read())
    '                HTMLVerisystem.ID = rdr("ID")
    '                HTMLVerisystem.Carteira = rdr("Carteira")
    '                HTMLVerisystem.Envio = rdr("Envio")
    '                HTMLVerisystem.Titulo = rdr("Titulo")
    '                HTMLVerisystem.NomeRemetente = rdr("NomeRemetente")
    '                HTMLVerisystem.HTML = rdr("HTML")
    '                HTMLVerisystem.Status = rdr("Status")
    '                HTMLVerisystem.Sucesso = True
    '                HTMLVerisystem.TipoErro = DadosGenericos.TipoErro.None
    '            End While
    '        Else
    '            HTMLVerisystem.Sucesso = False
    '            HTMLVerisystem.TipoErro = DadosGenericos.TipoErro.Funcional
    '            HTMLVerisystem.MsgErro = "Nenhum HTML encontrado."
    '            HTMLVerisystem.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
    '            HTMLVerisystem.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
    '        End If
    '        rdr.Close()
    '    Catch ex As Exception
    '        Dim _Retorno As New Parametros()
    '        With _Retorno
    '            .Sucesso = False
    '            .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
    '            .MsgErro = ex.Message
    '            .TipoErro = DadosGenericos.TipoErro.Arquitetura
    '            .ImagemErro = DadosGenericos.ImagemRetorno.Erro
    '        End With

    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    End Try

    '    Return HTMLVerisystem
    'End Function

    'Public Function VerificarVersaoVerisystem() As EntidadeGenerica

    '    Dim Configuracao As New EntidadeGenerica

    '    Try
    '        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '        Dim cmd As New SqlCommand("P_VerificaVersaoVerisystem", connection)
    '        cmd.CommandType = CommandType.StoredProcedure

    '        connection.Open()

    '        Dim rdr As SqlDataReader = cmd.ExecuteReader()

    '        If (rdr.HasRows) Then
    '            rdr.Read()

    '            Configuracao.Descricao = rdr("Versao")
    '            Configuracao.Sucesso = True
    '            Configuracao.TipoErro = DadosGenericos.TipoErro.None

    '        Else
    '            Configuracao.Sucesso = False
    '            Configuracao.TipoErro = DadosGenericos.TipoErro.Funcional
    '            Configuracao.MsgErro = "Nenhuma configuração encontrada."
    '            Configuracao.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
    '            Configuracao.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
    '        End If
    '        rdr.Close()
    '    Catch ex As Exception
    '        Dim _Retorno As New Parametros()
    '        With _Retorno
    '            .Sucesso = False
    '            .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVALORPLANTAO.Id
    '            .MsgErro = ex.Message
    '            .TipoErro = DadosGenericos.TipoErro.Arquitetura
    '            .ImagemErro = DadosGenericos.ImagemRetorno.Erro
    '        End With

    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    End Try

    '    Return Configuracao
    'End Function


End Class
