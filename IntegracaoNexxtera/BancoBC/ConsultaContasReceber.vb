Imports System.Data.SqlClient
Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Entidades
Imports System.Reflection

Public Class ConsultaContasReceber


    Public Function ConsultarValorTituloContaReceber(ByVal sNumtit As String, ByVal sSeqTit As String) As Decimal

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_ConsultaContaReceberValorTitulo", connection)
        Dim vlrInd As Decimal = 0

        Try
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New SqlParameter("@NumTit", sNumtit))
            command.Parameters.Add(New SqlParameter("@SeqTit", sSeqTit))

            connection.Open()
            Dim rdr As SqlDataReader = command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                vlrInd = Convert.ToDecimal(rdr("VlrInd"))
            End If

            rdr.Close()
        Catch ex As Exception
            'Trate aqui os casos em que a consulta não foi bem-sucedida
            '...
        Finally
            connection.Close()
        End Try

        Return vlrInd
    End Function

    Public Function ConsultaContasReceber(ByVal strNumTit As String,
                                          ByVal strSeqTit As String,
                                          ByVal dtEmissao As DateTime,
                                          ByVal strIdFilial As String) As ContaReceber


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaReceber
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_ConsultaContaReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", strIdFilial))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                contaReceber = New ContaReceber

                contaReceber.Situacao = rdr("Situacao").ToString
                contaReceber.NumTit = rdr("NumTit")
                contaReceber.CodClie = rdr("CodClie")
                contaReceber.SeqTit = rdr("Seqtit")
                contaReceber.DtEmissao = rdr("dtEmissao")
                contaReceber.IdFilial = rdr("IdFilial")


                contaReceber.Sucesso = True
                contaReceber.TipoErro = DadosGenericos.TipoErro.None



            Else
                contaReceber = New ContaReceber
                contaReceber.Sucesso = False
                contaReceber.TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            contaReceber = New ContaReceber
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_CONSULTACONTASRECEBER.Descricao & ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_CONSULTACONTASRECEBER.Id
            'lista(0).NumErro = EXCEPTION_METODO_CONSULTATITULOS.Id
            'lista(0).MsgErro = EXCEPTION_METODO_CONSULTATITULOS.Descricao & ex.Message
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, "Classe: ConsultaContasReceber - Função: ConsultaContasReceber(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return contaReceber

    End Function



    Public Function ConsultaContasReceber(ByVal strNumTit As String,
                                          ByVal strSeqTit As String,
                                          ByVal dtEmissao As DateTime,
                                          ByVal strIdFilial As String,
                                          ByVal connection As SqlConnection,
                                          ByVal Transacao As SqlTransaction) As ContaReceber


        Dim _Retorno = New Retorno
        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaReceber
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_ConsultaContaReceber", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure

            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", strIdFilial))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                contaReceber = New ContaReceber

                contaReceber.Situacao = rdr("Situacao").ToString
                contaReceber.NumTit = rdr("NumTit")
                contaReceber.CodClie = rdr("CodClie")
                contaReceber.SeqTit = rdr("Seqtit")
                contaReceber.DtEmissao = rdr("dtEmissao")
                contaReceber.IdFilial = rdr("IdFilial")
                contaReceber.ObsTit = rdr("ObsTit").ToString
                contaReceber.isTituloFinanciado = IIf(IsDBNull(rdr("isTituloFinanciado")), 0, rdr("isTituloFinanciado"))


                contaReceber.Sucesso = True
                contaReceber.TipoErro = DadosGenericos.TipoErro.None



            Else
                contaReceber = New ContaReceber
                contaReceber.Sucesso = False
                contaReceber.TipoErro = DadosGenericos.TipoErro.None
            End If

        Catch ex As Exception

            contaReceber = New ContaReceber
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_CONSULTACONTASRECEBER.Descricao & ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_CONSULTACONTASRECEBER.Id
            'lista(0).NumErro = EXCEPTION_METODO_CONSULTATITULOS.Id
            'lista(0).MsgErro = EXCEPTION_METODO_CONSULTATITULOS.Descricao & ex.Message
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: ConsultaContasReceber(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        rdr.Close()
        Return contaReceber
    End Function



    Public Function BuscaContaAReceberPorNumNFDtEmissao(ByVal strNumTit As String,
                                                        ByVal dtEmissao As DateTime,
                                                        ByVal connection As SqlConnection,
                                                        ByVal Transation As SqlTransaction) As List(Of ContaReceber)

        Dim _Retorno = New Retorno
        Dim rdr As SqlDataReader
        Dim contaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaAReceberPorNumNFDtEmissao", connection)
            Command.Transaction = Transation
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure

            Command.Parameters.Add(New SqlParameter("@NumNF", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    contaReceber.Add(New ContaReceber)

                    contaReceber(i).NumTit = rdr("NumTit")
                    contaReceber(i).SeqTit = rdr("Seqtit")
                    contaReceber(i).DtEmissao = rdr("dtEmissao")
                    contaReceber(i).TipoDup = rdr("TipoDup")
                    contaReceber(i).DtVcto = rdr("dtVcto")
                    contaReceber(i).CodClie = rdr("CodClie")
                    contaReceber(i).CodInd = rdr("CodInd")
                    contaReceber(i).VlrInd = rdr("VlrInd")
                    contaReceber(i).Situacao = rdr("Situacao").ToString
                    contaReceber(i).NumPortador = rdr("NumPortador").ToString
                    contaReceber(i).ObsTit = rdr("ObsTit")
                    contaReceber(i).DtCad = rdr("dtCad")
                    contaReceber(i).UsrCad = rdr("UsrCad")
                    contaReceber(i).StatusBco = rdr("StatusBco")
                    contaReceber(i).CodCCusto = rdr("CodCCusto")
                    contaReceber(i).CodBanco = rdr("CodBanco")
                    contaReceber(i).CodAgen = rdr("CodAgen")
                    contaReceber(i).NumCta = rdr("NumCta")
                    contaReceber(i).MenBco = rdr("MenBco")
                    contaReceber(i).TipoPagto = rdr("TipoPagto")
                    contaReceber(i).MsgBole = rdr("MsgBole").ToString
                    contaReceber(i).TitDesc = rdr("TitDesc")
                    contaReceber(i).IsTitNegoc = rdr("IsTitNegoc")
                    contaReceber(i).TipoCartaoCred = rdr("TipoCartaoCred").ToString
                    contaReceber(i).DtEnvCobExt = IIf(String.IsNullOrEmpty(rdr("DtEnvCobExt").ToString), Nothing, rdr("DtEnvCobExt"))
                    contaReceber(i).DtTitDesc = IIf(String.IsNullOrEmpty(rdr("DtTitDesc").ToString), Nothing, rdr("DtTitDesc"))
                    contaReceber(i).IdFilial = rdr("IdFilial").ToString
                    contaReceber(i).DtIncSerasa = IIf(String.IsNullOrEmpty(rdr("DtIncSerasa").ToString), Nothing, rdr("DtIncSerasa"))
                    contaReceber(i).DtExcSerasa = IIf(String.IsNullOrEmpty(rdr("DtExcSerasa").ToString), Nothing, rdr("DtExcSerasa"))
                    contaReceber(i).CodMotExc = rdr("CodMotExc").ToString
                    contaReceber(i).VlrProRata = IIf(String.IsNullOrEmpty(rdr("VlrProRata").ToString), Nothing, rdr("VlrProRata"))
                    contaReceber(i).NossoNumeroBco = rdr("NossoNumeroBco").ToString
                    contaReceber(i).DtPrevPgDI = IIf(String.IsNullOrEmpty(rdr("DtPrevPgDI").ToString), Nothing, rdr("DtPrevPgDI"))
                    contaReceber(i).CodBancoDeb = rdr("CodBancoDeb").ToString
                    contaReceber(i).CodAgenDeb = rdr("CodAgenDeb").ToString
                    contaReceber(i).NumCtaDeb = rdr("NumCtaDeb").ToString
                    contaReceber(i).StatusCob = rdr("StatusCob").ToString
                    contaReceber(i).DtUltCobranca = IIf(String.IsNullOrEmpty(rdr("DtUltCobranca").ToString), Nothing, rdr("DtUltCobranca"))
                    contaReceber(i).VlrPrevPgDI = IIf(String.IsNullOrEmpty(rdr("VlrPrevPgDI").ToString), Nothing, rdr("VlrPrevPgDI"))
                    contaReceber(i).TipoPgtDI = rdr("TipoPgDI").ToString


                    contaReceber(i).Sucesso = True
                    contaReceber(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

                'rdr.Close()
            Else
                contaReceber.Add(New ContaReceber)
                contaReceber(0).Sucesso = True
                contaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            contaReceber.Add(New ContaReceber)
            contaReceber(0).Sucesso = False
            contaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTAARECEBERPORNUMNFDTEMISSAO.Descricao & ex.Message
            contaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTAARECEBERPORNUMNFDTEMISSAO.Id
            contaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber(0).NumErro, contaReceber(0).MsgErro, contaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: ConsultaContasReceber(3)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return contaReceber
    End Function
    Public Function BuscaContaAReceberPorNumNFDtEmissaoMonitoriaNaoCancelado(ByVal strNumTit As String,
                                                                             ByVal strCodClie As String,
                                                                             ByVal dtEmissao As DateTime) As List(Of ContaReceber)

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Dim _Retorno = New Retorno
        Dim rdr As SqlDataReader
        Dim contaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaAReceberPorNumNFDtEmissaoTipoDupSituacao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure

            Command.Parameters.Add(New SqlParameter("@NumNF", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))
            Command.Parameters.Add(New SqlParameter("@CodClie", strCodClie))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    contaReceber.Add(New ContaReceber)

                    contaReceber(i).NumTit = rdr("NumTit")
                    contaReceber(i).SeqTit = rdr("Seqtit")
                    contaReceber(i).DtEmissao = rdr("dtEmissao")
                    contaReceber(i).TipoDup = rdr("TipoDup")
                    contaReceber(i).DtVcto = rdr("dtVcto")
                    contaReceber(i).CodClie = rdr("CodClie")
                    contaReceber(i).CodInd = rdr("CodInd")
                    contaReceber(i).VlrInd = rdr("VlrInd")
                    contaReceber(i).Situacao = rdr("Situacao").ToString
                    contaReceber(i).NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    contaReceber(i).ObsTit = IIf(IsDBNull(rdr("ObsTit")), Nothing, rdr("ObsTit"))
                    contaReceber(i).DtCad = IIf(IsDBNull(rdr("dtCad")), Nothing, rdr("dtCad"))
                    contaReceber(i).UsrCad = IIf(IsDBNull(rdr("UsrCad")), Nothing, rdr("UsrCad"))
                    contaReceber(i).StatusBco = IIf(IsDBNull(rdr("StatusBco")), Nothing, rdr("StatusBco"))
                    contaReceber(i).CodCCusto = IIf(IsDBNull(rdr("CodCCusto")), Nothing, rdr("CodCCusto"))
                    contaReceber(i).CodBanco = IIf(IsDBNull(rdr("CodBanco")), Nothing, rdr("CodBanco"))
                    contaReceber(i).CodAgen = IIf(IsDBNull(rdr("CodAgen")), Nothing, rdr("CodAgen"))
                    contaReceber(i).NumCta = IIf(IsDBNull(rdr("NumCta")), Nothing, rdr("NumCta"))
                    contaReceber(i).MenBco = IIf(IsDBNull(rdr("MenBco")), Nothing, rdr("MenBco"))
                    contaReceber(i).TipoPagto = IIf(IsDBNull(rdr("TipoPagto")), Nothing, rdr("TipoPagto"))
                    contaReceber(i).MsgBole = IIf(IsDBNull(rdr("MsgBole")), Nothing, rdr("MsgBole"))
                    contaReceber(i).TitDesc = IIf(IsDBNull(rdr("TitDesc")), Nothing, rdr("TitDesc"))
                    contaReceber(i).IsTitNegoc = IIf(IsDBNull(rdr("IsTitNegoc")), Nothing, rdr("IsTitNegoc"))
                    contaReceber(i).TipoCartaoCred = IIf(IsDBNull(rdr("TipoCartaoCred")), Nothing, rdr("TipoCartaoCred"))
                    contaReceber(i).DtEnvCobExt = IIf(IsDBNull(rdr("DtEnvCobExt")), Nothing, rdr("DtEnvCobExt"))
                    contaReceber(i).DtTitDesc = IIf(IsDBNull(rdr("DtTitDesc")), Nothing, rdr("DtTitDesc"))
                    contaReceber(i).IdFilial = rdr("IdFilial").ToString
                    contaReceber(i).DtIncSerasa = IIf(String.IsNullOrEmpty(rdr("DtIncSerasa").ToString), Nothing, rdr("DtIncSerasa"))
                    contaReceber(i).DtExcSerasa = IIf(String.IsNullOrEmpty(rdr("DtExcSerasa").ToString), Nothing, rdr("DtExcSerasa"))
                    contaReceber(i).CodMotExc = rdr("CodMotExc").ToString
                    contaReceber(i).VlrProRata = IIf(String.IsNullOrEmpty(rdr("VlrProRata").ToString), Nothing, rdr("VlrProRata"))
                    contaReceber(i).NossoNumeroBco = rdr("NossoNumeroBco").ToString
                    contaReceber(i).DtPrevPgDI = IIf(String.IsNullOrEmpty(rdr("DtPrevPgDI").ToString), Nothing, rdr("DtPrevPgDI"))
                    contaReceber(i).CodBancoDeb = IIf(IsDBNull(rdr("CodBancoDeb")), Nothing, rdr("CodBancoDeb"))
                    contaReceber(i).CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    contaReceber(i).NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    contaReceber(i).StatusCob = IIf(IsDBNull(rdr("StatusCob")), Nothing, rdr("StatusCob"))
                    contaReceber(i).DtUltCobranca = IIf(String.IsNullOrEmpty(rdr("DtUltCobranca").ToString), Nothing, rdr("DtUltCobranca"))
                    contaReceber(i).VlrPrevPgDI = IIf(String.IsNullOrEmpty(rdr("VlrPrevPgDI").ToString), Nothing, rdr("VlrPrevPgDI"))
                    contaReceber(i).TipoPgtDI = IIf(IsDBNull(rdr("TipoPgDI")), Nothing, rdr("TipoPgDI"))


                    contaReceber(i).Sucesso = True
                    contaReceber(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                contaReceber.Add(New ContaReceber)
                contaReceber(0).Sucesso = True
                contaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            contaReceber.Add(New ContaReceber)
            contaReceber(0).Sucesso = False
            contaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTAARECEBERPORNUMNFDTEMISSAO.Descricao & ex.Message
            contaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTAARECEBERPORNUMNFDTEMISSAO.Id
            contaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber(0).NumErro, contaReceber(0).MsgErro, contaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaAReceberPorNumNFDtEmissaoTipoDupSituacao(4)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return contaReceber
    End Function

    Public Function BuscaContaAReceberPorNumNFDtEmissaoTipoDupSituacao(ByVal strNumTit As String,
                                                                       ByVal strCodClie As String,
                                                                       ByVal dtEmissao As DateTime,
                                                                       ByVal connection As SqlConnection,
                                                                       ByVal Transacao As SqlTransaction) As List(Of ContaReceber)
        Dim _Retorno = New Retorno
        Dim rdr As SqlDataReader
        Dim contaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaAReceberPorNumNFDtEmissaoTipoDupSituacao", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure

            Command.Parameters.Add(New SqlParameter("@NumNF", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))
            Command.Parameters.Add(New SqlParameter("@CodClie", strCodClie))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    contaReceber.Add(New ContaReceber)

                    contaReceber(i).NumTit = rdr("NumTit")
                    contaReceber(i).SeqTit = rdr("Seqtit")
                    contaReceber(i).DtEmissao = rdr("dtEmissao")
                    contaReceber(i).TipoDup = rdr("TipoDup")
                    contaReceber(i).DtVcto = rdr("dtVcto")
                    contaReceber(i).CodClie = rdr("CodClie")
                    contaReceber(i).CodInd = rdr("CodInd")
                    contaReceber(i).VlrInd = rdr("VlrInd")
                    contaReceber(i).Situacao = rdr("Situacao").ToString
                    contaReceber(i).NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    contaReceber(i).ObsTit = IIf(IsDBNull(rdr("ObsTit")), Nothing, rdr("ObsTit"))
                    contaReceber(i).DtCad = IIf(IsDBNull(rdr("dtCad")), Nothing, rdr("dtCad"))
                    contaReceber(i).UsrCad = IIf(IsDBNull(rdr("UsrCad")), Nothing, rdr("UsrCad"))
                    contaReceber(i).StatusBco = IIf(IsDBNull(rdr("StatusBco")), Nothing, rdr("StatusBco"))
                    contaReceber(i).CodCCusto = IIf(IsDBNull(rdr("CodCCusto")), Nothing, rdr("CodCCusto"))
                    contaReceber(i).CodBanco = IIf(IsDBNull(rdr("CodBanco")), Nothing, rdr("CodBanco"))
                    contaReceber(i).CodAgen = IIf(IsDBNull(rdr("CodAgen")), Nothing, rdr("CodAgen"))
                    contaReceber(i).NumCta = IIf(IsDBNull(rdr("NumCta")), Nothing, rdr("NumCta"))
                    contaReceber(i).MenBco = IIf(IsDBNull(rdr("MenBco")), Nothing, rdr("MenBco"))
                    contaReceber(i).TipoPagto = IIf(IsDBNull(rdr("TipoPagto")), Nothing, rdr("TipoPagto"))
                    contaReceber(i).MsgBole = IIf(IsDBNull(rdr("MsgBole")), Nothing, rdr("MsgBole"))
                    contaReceber(i).TitDesc = IIf(IsDBNull(rdr("TitDesc")), Nothing, rdr("TitDesc"))
                    contaReceber(i).IsTitNegoc = IIf(IsDBNull(rdr("IsTitNegoc")), Nothing, rdr("IsTitNegoc"))
                    contaReceber(i).TipoCartaoCred = IIf(IsDBNull(rdr("TipoCartaoCred")), Nothing, rdr("TipoCartaoCred"))
                    contaReceber(i).DtEnvCobExt = IIf(IsDBNull(rdr("DtEnvCobExt")), Nothing, rdr("DtEnvCobExt"))
                    contaReceber(i).DtTitDesc = IIf(IsDBNull(rdr("DtTitDesc")), Nothing, rdr("DtTitDesc"))
                    contaReceber(i).IdFilial = rdr("IdFilial").ToString
                    contaReceber(i).DtIncSerasa = IIf(String.IsNullOrEmpty(rdr("DtIncSerasa").ToString), Nothing, rdr("DtIncSerasa"))
                    contaReceber(i).DtExcSerasa = IIf(String.IsNullOrEmpty(rdr("DtExcSerasa").ToString), Nothing, rdr("DtExcSerasa"))
                    contaReceber(i).CodMotExc = rdr("CodMotExc").ToString
                    contaReceber(i).VlrProRata = IIf(String.IsNullOrEmpty(rdr("VlrProRata").ToString), Nothing, rdr("VlrProRata"))
                    contaReceber(i).NossoNumeroBco = rdr("NossoNumeroBco").ToString
                    contaReceber(i).DtPrevPgDI = IIf(String.IsNullOrEmpty(rdr("DtPrevPgDI").ToString), Nothing, rdr("DtPrevPgDI"))
                    contaReceber(i).CodBancoDeb = IIf(IsDBNull(rdr("CodBancoDeb")), Nothing, rdr("CodBancoDeb"))
                    contaReceber(i).CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    contaReceber(i).NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    contaReceber(i).StatusCob = IIf(IsDBNull(rdr("StatusCob")), Nothing, rdr("StatusCob"))
                    contaReceber(i).DtUltCobranca = IIf(String.IsNullOrEmpty(rdr("DtUltCobranca").ToString), Nothing, rdr("DtUltCobranca"))
                    contaReceber(i).VlrPrevPgDI = IIf(String.IsNullOrEmpty(rdr("VlrPrevPgDI").ToString), Nothing, rdr("VlrPrevPgDI"))
                    contaReceber(i).TipoPgtDI = IIf(IsDBNull(rdr("TipoPgDI")), Nothing, rdr("TipoPgDI"))


                    contaReceber(i).Sucesso = True
                    contaReceber(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                contaReceber.Add(New ContaReceber)
                contaReceber(0).Sucesso = True
                contaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            contaReceber.Add(New ContaReceber)
            contaReceber(0).Sucesso = False
            contaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTAARECEBERPORNUMNFDTEMISSAO.Descricao & ex.Message
            contaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTAARECEBERPORNUMNFDTEMISSAO.Id
            contaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber(0).NumErro, contaReceber(0).MsgErro, contaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaAReceberPorNumNFDtEmissaoTipoDupSituacao(4)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return contaReceber
    End Function

    Public Function BuscaContaAReceberPorNumNFDtEmissao(ByVal strNumTit As String,
                                                        ByVal dtEmissao As DateTime) As List(Of ContaReceber)

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Dim rdr As SqlDataReader
        Dim contaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaAReceberPorNumNFDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure

            Command.Parameters.Add(New SqlParameter("@NumNF", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    contaReceber.Add(New ContaReceber)

                    contaReceber(i).NumTit = rdr("NumTit")
                    contaReceber(i).SeqTit = rdr("Seqtit")
                    contaReceber(i).DtEmissao = rdr("dtEmissao")
                    contaReceber(i).TipoDup = rdr("TipoDup")
                    contaReceber(i).DtVcto = rdr("dtVcto")
                    contaReceber(i).CodClie = rdr("CodClie")
                    contaReceber(i).CodInd = rdr("CodInd")
                    contaReceber(i).VlrInd = rdr("VlrInd")
                    contaReceber(i).Situacao = rdr("Situacao").ToString
                    contaReceber(i).NumPortador = rdr("NumPortador").ToString
                    contaReceber(i).ObsTit = rdr("ObsTit")
                    contaReceber(i).DtCad = rdr("dtCad")
                    contaReceber(i).UsrCad = rdr("UsrCad")
                    contaReceber(i).StatusBco = rdr("StatusBco")
                    contaReceber(i).CodCCusto = rdr("CodCCusto")
                    contaReceber(i).CodBanco = rdr("CodBanco")
                    contaReceber(i).CodAgen = rdr("CodAgen")
                    contaReceber(i).NumCta = rdr("NumCta")
                    contaReceber(i).MenBco = rdr("MenBco")
                    contaReceber(i).TipoPagto = rdr("TipoPagto")
                    contaReceber(i).MsgBole = rdr("MsgBole").ToString
                    contaReceber(i).TitDesc = rdr("TitDesc")
                    contaReceber(i).IsTitNegoc = rdr("IsTitNegoc")
                    contaReceber(i).TipoCartaoCred = rdr("TipoCartaoCred").ToString
                    contaReceber(i).DtEnvCobExt = IIf(String.IsNullOrEmpty(rdr("DtEnvCobExt").ToString), Nothing, rdr("DtEnvCobExt"))
                    contaReceber(i).DtTitDesc = IIf(String.IsNullOrEmpty(rdr("DtTitDesc").ToString), Nothing, rdr("DtTitDesc"))
                    contaReceber(i).IdFilial = rdr("IdFilial").ToString
                    contaReceber(i).DtIncSerasa = IIf(String.IsNullOrEmpty(rdr("DtIncSerasa").ToString), Nothing, rdr("DtIncSerasa"))
                    contaReceber(i).DtExcSerasa = IIf(String.IsNullOrEmpty(rdr("DtExcSerasa").ToString), Nothing, rdr("DtExcSerasa"))
                    contaReceber(i).CodMotExc = rdr("CodMotExc").ToString
                    contaReceber(i).VlrProRata = IIf(String.IsNullOrEmpty(rdr("VlrProRata").ToString), Nothing, rdr("VlrProRata"))
                    contaReceber(i).NossoNumeroBco = rdr("NossoNumeroBco").ToString
                    contaReceber(i).DtPrevPgDI = IIf(String.IsNullOrEmpty(rdr("DtPrevPgDI").ToString), Nothing, rdr("DtPrevPgDI"))
                    contaReceber(i).CodBancoDeb = rdr("CodBancoDeb").ToString
                    contaReceber(i).CodAgenDeb = rdr("CodAgenDeb").ToString
                    contaReceber(i).NumCtaDeb = rdr("NumCtaDeb").ToString
                    contaReceber(i).StatusCob = rdr("StatusCob").ToString
                    contaReceber(i).DtUltCobranca = IIf(String.IsNullOrEmpty(rdr("DtUltCobranca").ToString), Nothing, rdr("DtUltCobranca"))
                    contaReceber(i).VlrPrevPgDI = IIf(String.IsNullOrEmpty(rdr("VlrPrevPgDI").ToString), Nothing, rdr("VlrPrevPgDI"))
                    contaReceber(i).TipoPgtDI = rdr("TipoPgDI").ToString


                    contaReceber(i).Sucesso = True
                    contaReceber(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

                rdr.Close()
            Else
                contaReceber.Add(New ContaReceber)
                contaReceber(0).Sucesso = True
                contaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If

        Catch ex As Exception

            contaReceber.Add(New ContaReceber)
            contaReceber(0).Sucesso = False
            contaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTAARECEBERPORNUMNFDTEMISSAO.Descricao & ex.Message
            contaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTAARECEBERPORNUMNFDTEMISSAO.Id
            contaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber(0).NumErro, contaReceber(0).MsgErro, contaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: ConsultaContasReceber(5)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return contaReceber
    End Function



    Public Function BuscaContaReceberGeraRemessa(ByVal strCodBanco As String,
                                                 ByVal strCodAgencia As String,
                                                 ByVal strNumCta As String,
                                                 ByVal CodIntClie As String,
                                                 ByVal DtInicio As String,
                                                 ByVal DtFim As String,
                                                 ByVal TitInicio As String,
                                                 ByVal TitFim As String) As List(Of Exportacao)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = rdr("NumPortador")
                    lstExportar(i).ContaReceber.ObsTit = rdr("ObsTit")
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.NumPortador = rdr("TipoDup")
                    lstExportar(i).ContaReceber.ObsTit = rdr("MenBco")
                    lstExportar(i).ContaReceber.StatusBco = rdr("CodAgenDeb")
                    lstExportar(i).ContaReceber.NumPortador = rdr("NumCtaDeb")
                    lstExportar(i).ContaReceber.ObsTit = rdr("TipoPagto")


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Cep = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = rdr("NumNFe")
                    lstExportar(i).NotaFiscal.CodVerNFe = rdr("CodVerNFe")


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa(6)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return lstExportar

    End Function



    Public Function BuscaContaReceberGeraRemessa001(ByVal strCodBanco As String,
                                                    ByVal strCodAgencia As String,
                                                    ByVal strNumCta As String,
                                                    ByVal strTipoPagto As String,
                                                    ByVal CodIntClie As String,
                                                    ByVal DtInicio As String,
                                                    ByVal DtFim As String,
                                                    ByVal TitInicio As String,
                                                    ByVal TitFim As String,
                                                    ByVal Connection As SqlConnection,
                                                    ByVal Transaction As SqlTransaction,
                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)

        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa001", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA001.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA001.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa001_341(7)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function


    Public Function BuscaContaReceberGeraRemessa237(ByVal strCodBanco As String,
                                                    ByVal strCodAgencia As String,
                                                    ByVal strNumCta As String,
                                                    ByVal CodIntClie As String,
                                                    ByVal DtInicio As String,
                                                    ByVal DtFim As String,
                                                    ByVal TitInicio As String,
                                                    ByVal TitFim As String,
                                                    ByVal Connection As SqlConnection,
                                                    ByVal Transaction As SqlTransaction,
                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa237", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = IIf(IsDBNull(rdr("NumTit")), Nothing, rdr("NumTit"))
                    lstExportar(i).ContaReceber.SeqTit = IIf(IsDBNull(rdr("SeqTit")), Nothing, rdr("SeqTit"))
                    lstExportar(i).ContaReceber.DtEmissao = IIf(IsDBNull(rdr("DtEmissao")), Nothing, rdr("DtEmissao"))
                    lstExportar(i).ContaReceber.DtVcto = IIf(IsDBNull(rdr("DtVcto")), Nothing, rdr("DtVcto"))
                    lstExportar(i).ContaReceber.CodClie = IIf(IsDBNull(rdr("CodClie")), Nothing, rdr("CodClie"))
                    lstExportar(i).ContaReceber.CodInd = IIf(IsDBNull(rdr("CodInd")), Nothing, rdr("CodInd"))
                    lstExportar(i).ContaReceber.VlrInd = IIf(IsDBNull(rdr("VlrInd")), Nothing, rdr("VlrInd"))
                    lstExportar(i).ContaReceber.Situacao = IIf(IsDBNull(rdr("Situacao")), Nothing, rdr("Situacao"))
                    lstExportar(i).ContaReceber.CodBanco = IIf(IsDBNull(rdr("CodBanco")), Nothing, rdr("CodBanco"))
                    lstExportar(i).ContaReceber.CodAgen = IIf(IsDBNull(rdr("CodAgen")), Nothing, rdr("CodAgen"))
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), "", rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(IsDBNull(rdr("ObsTit")), Nothing, rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = IIf(IsDBNull(rdr("StatusBco")), Nothing, rdr("StatusBco"))
                    lstExportar(i).ContaReceber.TipoDup = IIf(IsDBNull(rdr("TipoDup")), Nothing, rdr("TipoDup"))
                    lstExportar(i).ContaReceber.MenBco = IIf(IsDBNull(rdr("MenBco")), Nothing, rdr("MenBco"))
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = IIf(IsDBNull(rdr("TipoPagto")), Nothing, rdr("TipoPagto"))
                    lstExportar(i).ContaReceber.VlrDesconto = IIf(IsDBNull(rdr("VlrDesconto")), Nothing, rdr("VlrDesconto"))
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = IIf(IsDBNull(rdr("RazaoSocial")), Nothing, rdr("RazaoSocial"))
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = IIf(IsDBNull(rdr("EnvioBanco")), Nothing, rdr("EnvioBanco"))
                    lstExportar(i).Cliente.Contabilidade.IsNNC = IIf(IsDBNull(rdr("IsNNC")), Nothing, rdr("IsNNC"))


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = IIf(IsDBNull(rdr("Endereco")), Nothing, rdr("Endereco"))
                    lstExportar(i).Cliente.Endereco.Bairro = IIf(IsDBNull(rdr("Bairro")), Nothing, rdr("Bairro"))
                    lstExportar(i).Cliente.Endereco.Cidade = IIf(IsDBNull(rdr("Cidade")), Nothing, rdr("Cidade"))
                    lstExportar(i).Cliente.Endereco.Cep = IIf(IsDBNull(rdr("CEP")), Nothing, rdr("CEP"))
                    lstExportar(i).Cliente.Endereco.UF = IIf(IsDBNull(rdr("UF")), Nothing, rdr("UF"))


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(IsDBNull(rdr("NumNFe")), Nothing, rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(IsDBNull(rdr("CodVerNFe")), Nothing, rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = IIf(IsDBNull(rdr("NFe")), Nothing, rdr("NFe"))
                    lstExportar(i).ContaReceber.NFMonit = IIf(IsDBNull(rdr("NFMonit")), Nothing, rdr("NFMonit"))

                    lstExportar(i).ContaReceber.CodIntClie = IIf(IsDBNull(rdr("CodIntClie")), Nothing, rdr("CodIntClie"))

                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()

        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA237.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA237.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa237(8)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function


    Public Function BuscaContaReceberGeraRemessa347(ByVal strCodBanco As String,
                                                    ByVal strCodAgencia As String,
                                                    ByVal strNumCta As String,
                                                    ByVal CodIntClie As String,
                                                    ByVal DtInicio As String,
                                                    ByVal DtFim As String,
                                                    ByVal TitInicio As String,
                                                    ByVal TitFim As String,
                                                    ByVal Connection As SqlConnection,
                                                    ByVal Transaction As SqlTransaction) As List(Of Exportacao)

        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa347", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = rdr("NumPortador")
                    lstExportar(i).ContaReceber.ObsTit = rdr("ObsTit")
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    'lstExportar(i).ContaReceber.CodAgenDeb = rdr("CodAgenDeb")
                    'lstExportar(i).ContaReceber.NumCtaDeb = rdr("NumCtaDeb")
                    'lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    'lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    'lstExportar(i).ContaReceber.DtLimiteDesconto = rdr("DtLimiteDesconto")


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    'lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    'lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    'lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    'lstExportar(i).NotaFiscal = New NotaFiscal
                    'lstExportar(i).NotaFiscal.NumNFe = rdr("NumNFe")
                    'lstExportar(i).NotaFiscal.CodVerNFe = rdr("CodVerNFe")


                    'lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    'lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA347.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA347.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa347(9)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function



    Public Function BuscaContaReceberGeraRemessa291(ByVal strCodBanco As String,
                                                    ByVal strCodAgencia As String,
                                                    ByVal strNumCta As String,
                                                    ByVal CodIntClie As String,
                                                    ByVal DtInicio As String,
                                                    ByVal DtFim As String,
                                                    ByVal TitInicio As String,
                                                    ByVal TitFim As String,
                                                    ByVal Connection As SqlConnection,
                                                    ByVal Transaction As SqlTransaction) As List(Of Exportacao)

        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa291", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = rdr("NumPortador")
                    lstExportar(i).ContaReceber.ObsTit = rdr("ObsTit")
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    'lstExportar(i).ContaReceber.CodAgenDeb = rdr("CodAgenDeb")
                    'lstExportar(i).ContaReceber.NumCtaDeb = rdr("NumCtaDeb")
                    'lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    'lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    'lstExportar(i).ContaReceber.DtLimiteDesconto = rdr("DtLimiteDesconto")


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    'lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    'lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    'lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    'lstExportar(i).NotaFiscal = New NotaFiscal
                    'lstExportar(i).NotaFiscal.NumNFe = rdr("NumNFe")
                    'lstExportar(i).NotaFiscal.CodVerNFe = rdr("CodVerNFe")


                    'lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    'lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA291.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA291.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa291(10)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function




    Public Function BuscaContaReceberGeraRemessa341_033_399_356(ByVal strCodBanco As String,
                                                            ByVal strCodAgencia As String,
                                                            ByVal strNumCta As String,
                                                            ByVal strTipoPagto As String,
                                                            ByVal CodIntClie As String,
                                                            ByVal DtInicio As String,
                                                            ByVal DtFim As String,
                                                            ByVal TitInicio As String,
                                                            ByVal TitFim As String,
                                                            ByVal Connection As SqlConnection,
                                                            ByVal Transaction As SqlTransaction,
                                                            Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa341", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")

                    lstExportar(i).ContaReceber.CodOperacaoConta = rdr("CodOperacaoConta")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa341(11)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function

    Public Function BuscaContaReceberGeraRemessaBoletoUnificado033(ByVal strCodBanco As String,
                                                            ByVal strCodAgencia As String,
                                                            ByVal strNumCta As String,
                                                            ByVal strTipoPagto As String,
                                                            ByVal CodIntClie As String,
                                                            ByVal DtInicio As String,
                                                            ByVal DtFim As String,
                                                            ByVal TitInicio As String,
                                                            ByVal TitFim As String,
                                                            ByVal Connection As SqlConnection,
                                                            ByVal Transaction As SqlTransaction) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa341_BoletoUnificado", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")

                    lstExportar(i).ContaReceber.CodOperacaoConta = rdr("CodOperacaoConta")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa341(11)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function


    Public Function BuscaContaReceberGeraRemessa_CartaoCredito(ByVal CodIntClie As String,
                                                               ByVal DtInicio As String,
                                                               ByVal DtFim As String,
                                                               ByVal TitInicio As String,
                                                               ByVal TitFim As String,
                                                               ByVal Connection As SqlConnection,
                                                               ByVal Transaction As SqlTransaction,
                                                               Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa_CartaoCredito", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(String.IsNullOrEmpty(rdr("ObsTit")), "", rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.MsgBole = IIf(IsDBNull(rdr("MsgBole")), "", rdr("MsgBole"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Financeiro = New ClienteFinanceiro
                    lstExportar(i).Cliente.Financeiro.TipoCartaoCred = rdr("TipoCartaoCred")
                    lstExportar(i).Cliente.Financeiro.CartaoCred = rdr("CartaoCred")
                    lstExportar(i).Cliente.Financeiro.ValidCCred = rdr("ValidCCred")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa_CartaoCredito(12)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function



    Public Function BuscaContaReceberGeraRemessa_CartaoCredito_Visa(ByVal CodIntClie As String,
                                                                    ByVal DtInicio As String,
                                                                    ByVal DtFim As String,
                                                                    ByVal TitInicio As String,
                                                                    ByVal TitFim As String,
                                                                    ByVal Connection As SqlConnection,
                                                                    ByVal Transaction As SqlTransaction,
                                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa_CartaoCredito_Visa", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(IsDBNull(rdr("ObsTit")), "", rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.MsgBole = IIf(IsDBNull(rdr("MsgBole")), "", rdr("MsgBole"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Financeiro = New ClienteFinanceiro
                    lstExportar(i).Cliente.Financeiro.TipoCartaoCred = rdr("TipoCartaoCred")
                    lstExportar(i).Cliente.Financeiro.CartaoCred = rdr("CartaoCred")
                    lstExportar(i).Cliente.Financeiro.ValidCCred = rdr("ValidCCred")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa_CartaoCredito_Visa(13)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function



    Public Function BuscaContaReceberGeraRemessa_CartaoCredito_Amex(ByVal CodIntClie As String,
                                                                    ByVal DtInicio As String,
                                                                    ByVal DtFim As String,
                                                                    ByVal TitInicio As String,
                                                                    ByVal TitFim As String, ByVal Connection As SqlConnection,
                                                                    ByVal Transaction As SqlTransaction,
                                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)

        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa_CartaoCredito_Amex", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(IsDBNull(rdr("ObsTit")), "", rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.MsgBole = IIf(IsDBNull(rdr("MsgBole")), "", rdr("MsgBole"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Financeiro = New ClienteFinanceiro
                    lstExportar(i).Cliente.Financeiro.TipoCartaoCred = rdr("TipoCartaoCred")
                    lstExportar(i).Cliente.Financeiro.CartaoCred = rdr("CartaoCred")
                    lstExportar(i).Cliente.Financeiro.ValidCCred = rdr("ValidCCred")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_AMEX.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_AMEX.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa_CartaoCredito_Amex(14)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function


    Public Function BuscaDuplicataContasReceber(ByVal strWhere As String,
                                                ByRef otable As DataTable, Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As Retorno


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno

        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaDuplicataContasReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@sWhere", strWhere))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure

            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                otable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Else

                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Dispose()
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCADUPLICATACONTASRECEBER.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCADUPLICATACONTASRECEBER.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaDuplicataContasReceber(15)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()

        End Try


        Return _Retorno

    End Function


    Public Function BuscaContasReceberPorNumTitSeqTitDtEmissao(ByVal strNumTit As String,
                                                               ByVal strSeqTit As String,
                                                               ByVal dtEmissao As DateTime) As ContaReceber

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaReceber
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContasReceberPorNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                contaReceber = New ContaReceber

                contaReceber.NumTit = rdr("NumTit")
                contaReceber.SeqTit = rdr("Seqtit")
                contaReceber.DtEmissao = rdr("dtEmissao")
                contaReceber.Situacao = rdr("Situacao").ToString
                contaReceber.IdFilial = rdr("IdFilial")
                contaReceber.CodClie = rdr("CodClie")
                contaReceber.TipoDup = rdr("TipoDup")
                If Not IsDBNull(rdr("IdFilialFinanciado")) Then
                    contaReceber.IdFilialFinanciado = rdr("IdFilialFinanciado")
                End If


                contaReceber.isTituloFinanciado = IIf(IsDBNull(rdr("isTituloFinanciado")), 0, rdr("isTituloFinanciado"))
                contaReceber.NumTitFinanciado = rdr("NumTitFinanciado").ToString()
                contaReceber.vlrFinanciamento = IIf(IsDBNull(rdr("vlrFinanciamento")), 0, rdr("vlrFinanciamento"))
                If Not IsDBNull(rdr("DtEmissaoFinanciado")) Then
                    contaReceber.DtEmissaoFinanciado = rdr("DtEmissaoFinanciado")
                End If


                contaReceber.Sucesso = True
                contaReceber.TipoErro = DadosGenericos.TipoErro.None

            Else
                contaReceber = New ContaReceber
                contaReceber.Sucesso = False
                contaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
        Catch ex As Exception

            contaReceber = New ContaReceber
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASRECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASRECEBERPORNUMTITSEQTITDTEMISSAO.Id
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContasReceberPorNumTitSeqTitDtEmissaO(16)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return contaReceber

    End Function

    Public Function BuscaContasReceberPorNumTitSeqTitDtEmissao(ByVal strNumTit As String,
                                                               ByVal strSeqTit As String,
                                                               ByVal dtEmissao As DateTime,
                                                               ByVal conn As SqlConnection,
                                                               ByVal trans As SqlTransaction) As ContaReceber

        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaReceber
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContasReceberPorNumTitSeqTitDtEmissao", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))


            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                contaReceber = New ContaReceber

                contaReceber.NumTit = rdr("NumTit")
                contaReceber.SeqTit = rdr("Seqtit")
                contaReceber.DtEmissao = rdr("dtEmissao")
                contaReceber.Situacao = rdr("Situacao").ToString
                contaReceber.IdFilial = rdr("IdFilial")
                contaReceber.CodClie = rdr("CodClie")
                contaReceber.TipoDup = rdr("TipoDup")
                If Not IsDBNull(rdr("IdFilialFinanciado")) Then
                    contaReceber.IdFilialFinanciado = rdr("IdFilialFinanciado")
                End If


                contaReceber.isTituloFinanciado = IIf(IsDBNull(rdr("isTituloFinanciado")), 0, rdr("isTituloFinanciado"))
                contaReceber.NumTitFinanciado = rdr("NumTitFinanciado").ToString()
                contaReceber.vlrFinanciamento = IIf(IsDBNull(rdr("vlrFinanciamento")), 0, rdr("vlrFinanciamento"))
                If Not IsDBNull(rdr("DtEmissaoFinanciado")) Then
                    contaReceber.DtEmissaoFinanciado = rdr("DtEmissaoFinanciado")
                End If


                contaReceber.Sucesso = True
                contaReceber.TipoErro = DadosGenericos.TipoErro.None

                rdr.Close()

            Else
                contaReceber = New ContaReceber
                contaReceber.Sucesso = False
                contaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
        Catch ex As Exception

            contaReceber = New ContaReceber
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASRECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASRECEBERPORNUMTITSEQTITDTEMISSAO.Id
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContasReceberPorNumTitSeqTitDtEmissaO(16)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'connection.Close()
        End Try


        Return contaReceber

    End Function

    Public Function PesqusiaDuplicataContasReceberPorNumTitSeqTitDtEmissao(ByVal strNumTit As String,
                                                                           ByVal strSeqTit As String,
                                                                           ByVal DtEmissao As DateTime,
                                                                           ByRef otable As DataTable,
                                                                           Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As Retorno

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno

        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_PesqusiaDuplicataContasReceberPorNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                otable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Else

                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_PESQUISADUPLICATACONTASRECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_PESQUISADUPLICATACONTASRECEBERPORNUMTITSEQTITDTEMISSAO.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return _Retorno

    End Function



    Public Function BuscaContaReceberPorSituacaoCodClieDtEmissaoTipoDup(ByVal strCodClie As String,
                                                                        ByVal DtEmissao As DateTime) As List(Of ContaReceber)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberPorSituacaoCodClieDtEmissaoTipoDup", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codclie", strCodClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).NumTit = rdr("NumTit")
                    lstContaReceber(i).SeqTit = rdr("SeqTit")
                    lstContaReceber(i).DtEmissao = rdr("DtEmissao")
                    lstContaReceber(i).DtVcto = rdr("DtVcto")
                    lstContaReceber(i).CodClie = rdr("CodClie")
                    lstContaReceber(i).Situacao = rdr("Situacao")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERPORSITUACAOCODCLIEDTEMISSAOTIPODUP.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERPORSITUACAOCODCLIEDTEMISSAOTIPODUP.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberPorSituacaoCodClieDtEmissaoTipoDup(18)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return lstContaReceber

    End Function



    Public Function BuscaDtEmissaoTitDescDtVctoSeqTitNumTitContasReceber(ByVal strNumTit As String,
                                                                         ByVal strSeqTit As String,
                                                                         ByVal DtVcto As DateTime,
                                                                         ByVal connection As SqlConnection,
                                                                         ByVal Transaction As SqlTransaction) As List(Of ContaReceber)

        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaDtEmissaoTitDescDtVctoSeqTitNumTitContasReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtVcto", DtVcto))
            Command.Transaction = Transaction
            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).NumTit = rdr("NumTit")
                    lstContaReceber(i).SeqTit = rdr("SeqTit")
                    lstContaReceber(i).DtEmissao = rdr("DtEmissao")
                    lstContaReceber(i).DtVcto = rdr("DtVcto")
                    lstContaReceber(i).TitDesc = rdr("TitDesc")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCADTEMISSAOTITDESCDTVCTOSEQTITNUMTITCONTASRECEBER.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCADTEMISSAOTITDESCDTVCTOSEQTITNUMTITCONTASRECEBER.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaDtEmissaoTitDescDtVctoSeqTitNumTitContasReceber(19)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'connection.Close()
        End Try

        Return lstContaReceber

    End Function

    Public Function BuscaDtEmissaoTitDescDtVctoSeqTitNumTitContasReceber(ByVal strNumTit As String, ByVal strSeqTit As String, ByVal DtVcto As DateTime) As List(Of ContaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaDtEmissaoTitDescDtVctoSeqTitNumTitContasReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtVcto", DtVcto))
            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).NumTit = rdr("NumTit")
                    lstContaReceber(i).SeqTit = rdr("SeqTit")
                    lstContaReceber(i).DtEmissao = rdr("DtEmissao")
                    lstContaReceber(i).DtVcto = rdr("DtVcto")
                    lstContaReceber(i).TitDesc = rdr("TitDesc")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCADTEMISSAOTITDESCDTVCTOSEQTITNUMTITCONTASRECEBER.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCADTEMISSAOTITDESCDTVCTOSEQTITNUMTITCONTASRECEBER.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaDtEmissaoTitDescDtVctoSeqTitNumTitContasReceber(19)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return lstContaReceber

    End Function

    Public Function BuscaEventosContasAReceber(Optional ByVal strCodEnvento As String = "") As List(Of EventoCtaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstEventoCtaReceber As New List(Of EventoCtaReceber)
        Dim i As Integer = 0
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaEventosContasAReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodEvento", strCodEnvento))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstEventoCtaReceber.Add(New EventoCtaReceber)
                    lstEventoCtaReceber(i).CodEvento = rdr("CodEvento")
                    lstEventoCtaReceber(i).Descricao = rdr("Descricao")

                    lstEventoCtaReceber(0).Sucesso = True
                    lstEventoCtaReceber(0).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstEventoCtaReceber.Add(New EventoCtaReceber)
                lstEventoCtaReceber(0).Sucesso = False
                lstEventoCtaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstEventoCtaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstEventoCtaReceber(0).TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()


        Catch ex As Exception
            lstEventoCtaReceber.Add(New EventoCtaReceber)
            lstEventoCtaReceber(0).Sucesso = False
            lstEventoCtaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAEVENTOSCTARECEBER.Descricao & ex.Message
            lstEventoCtaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAEVENTOSCTARECEBER.Id
            lstEventoCtaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstEventoCtaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstEventoCtaReceber(0).NumErro, lstEventoCtaReceber(0).MsgErro, lstEventoCtaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaEventosContasAReceber(20)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return lstEventoCtaReceber

    End Function




    Public Function BuscaTitulosContasAReceberCobranca(ByVal strCodClie As String,
                                                       ByVal strQuery As String,
                                                       ByVal oTable As DataTable) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno
        Dim i As Integer = 0
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaTitulosContasAReceberCobranca", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodClie", strCodClie))
            Command.Parameters.Add(New SqlParameter("@sQuery", strQuery))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                oTable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None

            Else
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSCONTASARECEBERCOBRANCA.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSCONTASARECEBERCOBRANCA.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaTitulosContasAReceberCobranca(21)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return _Retorno

    End Function




    Public Function BuscaContaReceberCobrancaExterna(ByVal strQuery As String,
                                                     ByVal oTable As DataTable) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberCobrancaExterna", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@strCondicao", strQuery))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                oTable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None

            Else
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERCOBRANCAEXTERNA.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERCOBRANCAEXTERNA.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberCobrancaExterna(22)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return _Retorno

    End Function



    Public Function BuscaContasAReceberInclusaoSerasa(ByVal strQuery As String,
                                                      ByVal intQtdeTitulos As Integer,
                                                      ByVal strTipoTitulo As String,
                                                      ByVal oTable As DataTable) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno
        Dim i As Integer = 0
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContasAReceberInclusaoSerasa", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@strCondicao", strQuery))
            Command.Parameters.Add(New SqlParameter("@QtdeTitulos", intQtdeTitulos))
            Command.Parameters.Add(New SqlParameter("@TpTitulo", strTipoTitulo))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                oTable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None

            Else
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASARECEBERINCLUSAOSERASA.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASARECEBERINCLUSAOSERASA.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContasAReceberInclusaoSerasa(23)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return _Retorno

    End Function



    Public Function BuscaContasAReceberExclusaoSerasa(ByVal strQuery As String,
                                                      ByVal strEmpresaNegativacao As String,
                                                      ByVal oTable As DataTable) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno
        Dim i As Integer = 0
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContasAReceberExclusaoSerasa", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@strCondicao", strQuery))
            Command.Parameters.Add(New SqlParameter("@EmpresaNegativacao", strEmpresaNegativacao))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                oTable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None

            Else
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASARECEBEREXCLUSAOSERASA.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASARECEBEREXCLUSAOSERASA.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContasAReceberExclusaoSerasa(24)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return _Retorno

    End Function




    Public Function BuscaContasAReceberPorRazaoSNumTitSeqTitDtEmissao(ByVal RazaoSocial As String,
                                                                      ByVal NumTit As String,
                                                                      ByVal SeqTit As String,
                                                                      ByVal DtEmissao As DateTime,
                                                                      ByRef connection As SqlConnection,
                                                                      ByRef Transacao As SqlTransaction) As List(Of ContaReceber)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContasAReceberPorRazaoSNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transacao

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@razaosocial", RazaoSocial))
            Command.Parameters.Add(New SqlParameter("@NumTit", NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", SeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).CodIntClie = rdr("CodIntClie")
                    lstContaReceber(i).DtIncSerasa = rdr("DtIncSerasa")
                    lstContaReceber(i).DtExcSerasa = rdr("DtExcSerasa")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASARECEBERPORRAZAOSNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASARECEBERPORRAZAOSNUMTITSEQTITDTEMISSAO.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContasAReceberPorRazaoSNumTitSeqTitDtEmissao(25)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return lstContaReceber

    End Function



    Public Function BuscaBcoPortContaAReceberPorNumTitSeqTitDataEmissaoCodClie(ByVal strCodClie As String,
                                                                               ByVal NumTit As String,
                                                                               ByVal SeqTit As String,
                                                                               ByVal DtEmissao As DateTime) As List(Of ContaReceber)

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaBcoPortContaAReceberPorNumTitSeqTitDataEmissaoCodClie", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query


            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodClie", strCodClie))
            Command.Parameters.Add(New SqlParameter("@NumTit", NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", SeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).CodBanco = rdr("CodBanco")
                    lstContaReceber(i).CodAgen = rdr("CodAgen")
                    lstContaReceber(i).NumPortador = rdr("NumPortador")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASARECEBERPORRAZAOSNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASARECEBERPORRAZAOSNUMTITSEQTITDTEMISSAO.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaBcoPortContaAReceberPorNumTitSeqTitDataEmissaoCodClie(26)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return lstContaReceber

    End Function



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' 
    ''' SELECT 
    '''    Contas_a_Receber.NumTit,
    '''    Contas_a_Receber.SeqTit, 
    '''    Contas_a_Receber.DtEmissao,
    '''    Cliente.CodClie
    ''' FROM Contas_a_Receber
    '''	    left join Cliente on Contas_a_Receber.CodClie=Cliente.CodClie
    ''' WHERE Contas_a_Receber.Situacao != 'Q'
    ''' ORDER BY Contas_a_Receber.CodClie
    ''' 
    ''' </remarks>
    Public Function BuscaNumTitSeqTitDtEmissaoContasAReceberPorSituacaoQ() As List(Of ContaReceber)

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaNumTitSeqTitDtEmissaoContasAReceberPorSituacaoQ", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            connection.Open()

            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).NumTit = rdr("NumTit")
                    lstContaReceber(i).SeqTit = rdr("SeqTit")
                    lstContaReceber(i).DtEmissao = rdr("DtEmissao")
                    lstContaReceber(i).CodClie = rdr("CodClie")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCANUMTITSEQTITDTEMISSAOCONTASARECEBERPORSITUACAIQ.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCANUMTITSEQTITDTEMISSAOCONTASARECEBERPORSITUACAIQ.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaNumTitSeqTitDtEmissaoContasAReceberPorSituacaoQ(27)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return lstContaReceber

    End Function




    Public Function BuscaTituloBaixaManual(ByVal Tipo As String, ByVal DtIni As DateTime, ByVal DtFim As DateTime, ByRef oTable As DataTable) As Retorno

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno
        Dim dtTable As New DataTable
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaTituloBaixaManual_201309", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            'Command.Parameters.Add(New SqlParameter("@sWhere", strWhere))
            Command.Parameters.Add(New SqlParameter("@Tipo", Tipo))
            Command.Parameters.Add(New SqlParameter("@DtIni", DtIni))
            Command.Parameters.Add(New SqlParameter("@DtFim", DtFim))

            connection.Open()

            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                oTable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None

            Else
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOBAIXAMANUAL.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOBAIXAMANUAL.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaTituloBaixaManual(28)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return _Retorno

    End Function




    Public Function BuscaVlrIndContasAReceberPorNumTitSeqTitDtEmissaoIdFilial(ByVal strIdFilial As String,
                                                                              ByVal strNumTit As String,
                                                                              ByVal strSeqTit As String,
                                                                              ByVal strDtEmissao As DateTime) As List(Of ContaReceber)

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaVlrIndContasAReceberPorNumTitSeqTitDtEmissaoIdFilial", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query


            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@IdFilial", strIdFilial))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", strDtEmissao))


            connection.Open()

            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).NumTit = rdr("NumTit")
                    lstContaReceber(i).SeqTit = rdr("SeqTit")
                    lstContaReceber(i).DtEmissao = rdr("DtEmissao")
                    lstContaReceber(i).VlrInd = rdr("VlrInd")
                    lstContaReceber(i).CodClie = rdr("CodClie")
                    lstContaReceber(i).IdFilial = rdr("IdFilial")
                    lstContaReceber(i).IdBU = rdr("IdBU")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAVLRINDCONTASARECEBERPORNUMTITSEQTITDTEMISSAOIDFILIAL.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVLRINDCONTASARECEBERPORNUMTITSEQTITDTEMISSAOIDFILIAL.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaVlrIndContasAReceberPorNumTitSeqTitDtEmissaoIdFilial(29)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return lstContaReceber

    End Function


    Public Function BuscaVlrIndContasAReceberPorNumTitSeqTitDtEmissaoIdFilial(ByVal strIdFilial As String,
                                                                              ByVal strNumTit As String,
                                                                              ByVal strSeqTit As String,
                                                                              ByVal strDtEmissao As DateTime,
                                                                              ByVal Connection As SqlConnection,
                                                                              ByVal Transaction As SqlTransaction) As List(Of ContaReceber)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaVlrIndContasAReceberPorNumTitSeqTitDtEmissaoIdFilial", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@IdFilial", strIdFilial))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", strDtEmissao))

            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).NumTit = rdr("NumTit")
                    lstContaReceber(i).SeqTit = rdr("SeqTit")
                    lstContaReceber(i).DtEmissao = rdr("DtEmissao")
                    lstContaReceber(i).VlrInd = IIf(IsDBNull(rdr("VlrInd")), 0, rdr("VlrInd"))
                    lstContaReceber(i).CodClie = rdr("CodClie")
                    lstContaReceber(i).IdFilial = rdr("IdFilial")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAVLRINDCONTASARECEBERPORNUMTITSEQTITDTEMISSAOIDFILIAL.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAVLRINDCONTASARECEBERPORNUMTITSEQTITDTEMISSAOIDFILIAL.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaVlrIndContasAReceberPorNumTitSeqTitDtEmissaoIdFilial(30)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return lstContaReceber

    End Function


    Public Function BuscaTitReintegracao(ByVal Dtinicial As String, ByVal DtFinal As String, ByRef oTable As DataTable, Optional NumTit As String = "") As Retorno

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        'Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno
        Dim i As Integer = 0
        'Dim dta As SqlDataAdapter
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaTitReintegracao_DEV", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@DtInicial", Dtinicial))
            Command.Parameters.Add(New SqlParameter("@DtFinal", DtFinal))
            Command.Parameters.Add(New SqlParameter("@NumTit", NumTit))

            connection.Open()

            'dta = New SqlDataAdapter(Command)
            'rdr = Command.ExecuteReader()

            'dta.Fill(oTable)

            oTable.Load(Command.ExecuteReader(CommandBehavior.CloseConnection))

            If oTable.Rows.Count > 0 Then
                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None

            Else
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            'dta.Dispose()
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCATITREINTEGRACAO.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCATITREINTEGRACAO.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaTitReintegracao(31)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return _Retorno

    End Function



    Public Function BuscaContaReceberPorNumTitSeqTitCodClieDtEmissaoIdFilial(ByVal strNumTit As String,
                                                                             ByVal strSeqTit As String,
                                                                             ByVal strCodClie As String,
                                                                             ByVal strIdFilial As String,
                                                                             ByVal dtEmissao As DateTime,
                                                                             ByVal connection As SqlConnection,
                                                                             ByVal Transation As SqlTransaction) As List(Of ContaReceber)

        Dim rdr As SqlDataReader
        Dim contaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberPorNumTitSeqTitCodClieDtEmissaoIdFilial", connection)
            Command.Transaction = Transation
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure

            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@CodClie", strCodClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", strIdFilial))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    contaReceber.Add(New ContaReceber)

                    contaReceber(i).NumTit = rdr("NumTit")
                    contaReceber(i).SeqTit = rdr("Seqtit")
                    contaReceber(i).DtEmissao = rdr("dtEmissao")
                    contaReceber(i).TipoDup = rdr("TipoDup")
                    contaReceber(i).DtVcto = rdr("dtVcto")
                    contaReceber(i).CodClie = rdr("CodClie")
                    contaReceber(i).CodInd = rdr("CodInd")
                    contaReceber(i).VlrInd = rdr("VlrInd")
                    contaReceber(i).Situacao = rdr("Situacao").ToString
                    contaReceber(i).NumPortador = rdr("NumPortador").ToString
                    contaReceber(i).ObsTit = rdr("ObsTit")
                    contaReceber(i).DtCad = rdr("dtCad")
                    contaReceber(i).UsrCad = rdr("UsrCad")
                    contaReceber(i).StatusBco = rdr("StatusBco")
                    contaReceber(i).CodCCusto = rdr("CodCCusto")
                    contaReceber(i).CodBanco = rdr("CodBanco")
                    contaReceber(i).CodAgen = rdr("CodAgen")
                    contaReceber(i).NumCta = rdr("NumCta")
                    contaReceber(i).MenBco = rdr("MenBco")
                    contaReceber(i).TipoPagto = rdr("TipoPagto")
                    contaReceber(i).MsgBole = rdr("MsgBole").ToString
                    contaReceber(i).TitDesc = rdr("TitDesc")
                    contaReceber(i).IsTitNegoc = rdr("IsTitNegoc")
                    contaReceber(i).TipoCartaoCred = rdr("TipoCartaoCred").ToString
                    contaReceber(i).DtEnvCobExt = IIf(String.IsNullOrEmpty(rdr("DtEnvCobExt").ToString), Nothing, rdr("DtEnvCobExt"))
                    contaReceber(i).DtTitDesc = IIf(String.IsNullOrEmpty(rdr("DtTitDesc").ToString), Nothing, rdr("DtTitDesc"))
                    contaReceber(i).IdFilial = rdr("IdFilial").ToString
                    contaReceber(i).DtIncSerasa = IIf(String.IsNullOrEmpty(rdr("DtIncSerasa").ToString), Nothing, rdr("DtIncSerasa"))
                    contaReceber(i).DtExcSerasa = IIf(String.IsNullOrEmpty(rdr("DtExcSerasa").ToString), Nothing, rdr("DtExcSerasa"))
                    contaReceber(i).CodMotExc = rdr("CodMotExc").ToString
                    contaReceber(i).VlrProRata = IIf(String.IsNullOrEmpty(rdr("VlrProRata").ToString), Nothing, rdr("VlrProRata"))
                    contaReceber(i).NossoNumeroBco = rdr("NossoNumeroBco").ToString
                    contaReceber(i).DtPrevPgDI = IIf(String.IsNullOrEmpty(rdr("DtPrevPgDI").ToString), Nothing, rdr("DtPrevPgDI"))
                    contaReceber(i).CodBancoDeb = rdr("CodBancoDeb").ToString
                    contaReceber(i).CodAgenDeb = rdr("CodAgenDeb").ToString
                    contaReceber(i).NumCtaDeb = rdr("NumCtaDeb").ToString
                    contaReceber(i).StatusCob = rdr("StatusCob").ToString
                    contaReceber(i).DtUltCobranca = IIf(String.IsNullOrEmpty(rdr("DtUltCobranca").ToString), Nothing, rdr("DtUltCobranca"))
                    contaReceber(i).VlrPrevPgDI = IIf(String.IsNullOrEmpty(rdr("VlrPrevPgDI").ToString), Nothing, rdr("VlrPrevPgDI"))
                    contaReceber(i).TipoPgtDI = rdr("TipoPgDI").ToString


                    contaReceber(i).Sucesso = True
                    contaReceber(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

                'rdr.Close()    
            Else
                contaReceber.Add(New ContaReceber)
                contaReceber(0).Sucesso = True
                contaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            contaReceber.Add(New ContaReceber)
            contaReceber(0).Sucesso = False
            contaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERPORNUMTITSEQTITCODCLIEDTEMISSAOIDFILIAL.Descricao & ex.Message
            contaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERPORNUMTITSEQTITCODCLIEDTEMISSAOIDFILIAL.Id
            contaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber(0).NumErro, contaReceber(0).MsgErro, contaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberPorNumTitSeqTitCodClieDtEmissaoIdFilial(32)", "08", "Verisure", Environment.MachineName, "2.0", "13")
            rdr.Close()
        End Try

        Return contaReceber
    End Function



    Public Function BuscaNumTitSeqTitDtVctoCodAgenNumCtaVlrIndContasAReceberPorCodDI(ByVal strCodDI As String,
                                                                                     ByVal Connection As SqlConnection,
                                                                                     ByVal Transaction As SqlTransaction) As List(Of ContaReceber)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaNumTitSeqTitDtVctoCodAgenNumCtaVlrIndContasAReceberPorCodDI", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodDI", strCodDI))


            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    lstContaReceber(i).NumTit = rdr("NumTit")
                    lstContaReceber(i).SeqTit = rdr("SeqTit")
                    lstContaReceber(i).DtVcto = rdr("DtVcto")
                    lstContaReceber(i).CodAgen = rdr("CodAgen")
                    lstContaReceber(i).NumCta = rdr("NumCta")
                    lstContaReceber(i).VlrInd = rdr("VlrInd")

                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                Loop

            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCANUMTITSEQTITDTVCTOCODAGENNUMCTAVLRINDCONTASARECEBERPORCODDI.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCANUMTITSEQTITDTVCTOCODAGENNUMCTAVLRINDCONTASARECEBERPORCODDI.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaNumTitSeqTitDtVctoCodAgenNumCtaVlrIndContasAReceberPorCodDI(33)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try
        Return lstContaReceber

    End Function



    Public Function BuscaTitulosContasAReceberCobranca(ByVal strCodClie As String,
                                                      ByVal strQuery As String,
                                                      ByVal oTable As DataTable,
                                                      ByVal connection As SqlConnection,
                                                      ByVal Transaction As SqlTransaction) As Retorno

        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaTitulosContasAReceberCobranca", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodClie", strCodClie))
            Command.Parameters.Add(New SqlParameter("@sQuery", strQuery))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                oTable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None

            Else
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSCONTASARECEBERCOBRANCA.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSCONTASARECEBERCOBRANCA.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaTitulosContasAReceberCobranca(34)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function


    Public Function BuscaContasREceberTitulosDescontados(ByVal strQuery As String,
                                                         ByVal oTable As DataTable) As Retorno

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Retorno As New Retorno
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContasREceberTitulosDescontados", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@sWhere", strQuery))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                oTable.Load(rdr)

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None

            Else
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASRECEBERTITULOSDESCONTADOS.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASRECEBERTITULOSDESCONTADOS.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContasREceberTitulosDescontados(35)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return _Retorno

    End Function

    Public Function BuscaTitulosContasAReceberPorDtVctoCodCCusto(ByVal dtDe As DateTime, ByVal dtAte As DateTime, ByVal strCodCCusto As String) As List(Of ContaReceber)

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstCtaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaTitulosContasAReceberPorDtVctoCodCCusto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@DTDE", (dtDe.Date)))
            Command.Parameters.Add(New SqlParameter("@DTATE", (dtAte.Date)))
            Command.Parameters.Add(New SqlParameter("@CODCCUSTO", strCodCCusto))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstCtaReceber.Add(New ContaReceber)
                    lstCtaReceber(i).NumTit = rdr("NumTit").ToString
                    lstCtaReceber(i).SeqTit = rdr("SeqTit").ToString
                    lstCtaReceber(i).RazaoSocial = rdr("RazaoSocial").ToString
                    lstCtaReceber(i).VlrInd = Convert.ToDouble(rdr("VlrInd"))
                    lstCtaReceber(i).DtVcto = Convert.ToDateTime(rdr("DtVcto"))
                    lstCtaReceber(i).DtEmissao = Convert.ToDateTime(rdr("DtEmissao"))
                    lstCtaReceber(i).FisiJuri = rdr("FisiJuri")
                    lstCtaReceber(i).CodInd = rdr("CodInd")
                    lstCtaReceber(i).ObsTit = rdr("ObsTit")


                    lstCtaReceber(i).Sucesso = True
                    lstCtaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop




            Else
                lstCtaReceber.Add(New ContaReceber)
                lstCtaReceber(0).Sucesso = False
                lstCtaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstCtaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstCtaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception
            lstCtaReceber.Add(New ContaReceber)
            lstCtaReceber(0).Sucesso = False
            lstCtaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSCONTASARECEBERPORDTVCTOCODCCUSTO.Descricao & ex.Message
            lstCtaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSCONTASARECEBERPORDTVCTOCODCCUSTO.Id
            lstCtaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstCtaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstCtaReceber(0).NumErro, lstCtaReceber(0).MsgErro, lstCtaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaTitulosContasAReceberPorDtVctoCodCCusto(36)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return lstCtaReceber

    End Function


    Public Function BuscaTodosEventosCtaReceber() As List(Of EventoCtaReceber)

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstCtaReceber As New List(Of EventoCtaReceber)
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaTodosEventosCtaReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query


            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstCtaReceber.Add(New EventoCtaReceber)
                    lstCtaReceber(i).CodEvento = rdr("CodEvento").ToString
                    lstCtaReceber(i).Descricao = rdr("Descricao").ToString


                    lstCtaReceber(i).Sucesso = True
                    lstCtaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop




            Else
                lstCtaReceber.Add(New EventoCtaReceber)
                lstCtaReceber(0).Sucesso = False
                lstCtaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstCtaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstCtaReceber(0).TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception
            lstCtaReceber.Add(New EventoCtaReceber)
            lstCtaReceber(0).Sucesso = False
            lstCtaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSCONTASARECEBERPORDTVCTOCODCCUSTO.Descricao & ex.Message
            lstCtaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSCONTASARECEBERPORDTVCTOCODCCUSTO.Id
            lstCtaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstCtaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstCtaReceber(0).NumErro, lstCtaReceber(0).MsgErro, lstCtaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaTodosEventosCtaReceber(37)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return lstCtaReceber

    End Function


    Public Function BuscaCodIndPorNumTitSeqTitDtEmissao(ByVal dtEmissao As DateTime, ByVal strNumTit As String, ByVal strSeqTit As String) As ContaReceber

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _CtaReceber As New ContaReceber
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodIndPorNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@SEQTIT", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@NUMTIT", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DTEMISSAO", Convert.ToDateTime(dtEmissao)))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read

                    _CtaReceber.CodInd = rdr("CodInd")
                    _CtaReceber.IdFilial = rdr("IdFilial")

                    _CtaReceber.Sucesso = True
                    _CtaReceber.TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop




            Else
                'lstCtaReceber.Add(New ContaReceber)
                _CtaReceber.Sucesso = False
                _CtaReceber.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _CtaReceber.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _CtaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception
            'lstCtaReceber.Add(New ContaReceber)
            _CtaReceber.Sucesso = False
            _CtaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINDPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _CtaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINDPORNUMTITSEQTITDTEMISSAO.Id
            _CtaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _CtaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_CtaReceber.NumErro, _CtaReceber.MsgErro, _CtaReceber.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaCodIndPorNumTitSeqTitDtEmissao(38)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return _CtaReceber

    End Function

    Public Function BuscarContasAReceberPorNumTitSeqTitDtEmissao(ByVal dtEmissao As DateTime, ByVal strNumTit As String, ByVal strSeqTit As String) As ContaReceber

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _CtaReceber As New ContaReceber
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarContasAReceberPorNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@SEQTIT", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@NUMTIT", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DTEMISSAO", Convert.ToDateTime(dtEmissao)))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read

                    _CtaReceber.NumTit = rdr("NumTit").ToString()
                    _CtaReceber.SeqTit = rdr("SeqTit").ToString()
                    _CtaReceber.CodClie = rdr("CodClie").ToString()
                    _CtaReceber.RazaoSocial = rdr("RazaoSocial").ToString()
                    _CtaReceber.DtEmissao = Convert.ToDateTime(rdr("DtEmissao"))
                    _CtaReceber.DtVcto = Convert.ToDateTime(rdr("DtVcto"))
                    _CtaReceber.VlrInd = rdr("VlrInd")
                    _CtaReceber.FisiJuri = rdr("FisiJuri")
                    _CtaReceber.CodInd = rdr("CodInd").ToString()
                    _CtaReceber.CodCCusto = rdr("CCusto").ToString()
                    _CtaReceber.ObsTit = rdr("Obstit").ToString()



                    _CtaReceber.Sucesso = True
                    _CtaReceber.TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop




            Else
                'lstCtaReceber.Add(New ContaReceber)
                _CtaReceber.Sucesso = False
                _CtaReceber.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _CtaReceber.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _CtaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception
            'lstCtaReceber.Add(New ContaReceber)
            _CtaReceber.Sucesso = False
            _CtaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARCONTASARECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _CtaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARCONTASARECEBERPORNUMTITSEQTITDTEMISSAO.Id
            _CtaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _CtaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_CtaReceber.NumErro, _CtaReceber.MsgErro, _CtaReceber.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscarContasAReceberPorNumTitSeqTitDtEmissao(39)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return _CtaReceber

    End Function

    Public Function BuscarIsnullVlrIndCtaReceber(ByVal dtEmissao As DateTime, ByVal strNumTit As String, ByVal strSeqTit As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As ContaReceber

        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _CtaReceber As New ContaReceber
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarIsnullVlrIndCtaReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans
            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@SEQTIT", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@NUMTIT", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DTEMISSAO", Convert.ToDateTime(dtEmissao)))

            ' connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()


                _CtaReceber.VlrInd = rdr("VlrInd")




                _CtaReceber.Sucesso = True
                _CtaReceber.TipoErro = DadosGenericos.TipoErro.None





            Else
                'lstCtaReceber.Add(New ContaReceber)
                _CtaReceber.Sucesso = False
                _CtaReceber.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _CtaReceber.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _CtaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception
            'lstCtaReceber.Add(New ContaReceber)
            _CtaReceber.Sucesso = False
            _CtaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARCONTASARECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _CtaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARCONTASARECEBERPORNUMTITSEQTITDTEMISSAO.Id
            _CtaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _CtaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_CtaReceber.NumErro, _CtaReceber.MsgErro, _CtaReceber.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscarIsnullVlrIndCtaReceber(40)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'connection.Close()
        End Try

        Return _CtaReceber

    End Function

    Public Function BuscaCodIndPorNumTitSeqTitDtEmissao(ByVal dtEmissao As DateTime, ByVal strNumTit As String, ByVal strSeqTit As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As ContaReceber

        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _CtaReceber As New ContaReceber
        Dim i As Integer = 0
        Try


            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodIndPorNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans
            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@SEQTIT", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@NUMTIT", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DTEMISSAO", Convert.ToDateTime(dtEmissao)))

            ' connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read

                    _CtaReceber.CodInd = rdr("CodInd")
                    _CtaReceber.IdFilial = rdr("IdFilial")
                    _CtaReceber.DtEmissao = Convert.ToDateTime(rdr("DtEmissao"))

                    _CtaReceber.Sucesso = True
                    _CtaReceber.TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop




            Else
                'lstCtaReceber.Add(New ContaReceber)
                _CtaReceber.Sucesso = False
                _CtaReceber.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _CtaReceber.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _CtaReceber.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception
            'lstCtaReceber.Add(New ContaReceber)
            _CtaReceber.Sucesso = False
            _CtaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINDPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _CtaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINDPORNUMTITSEQTITDTEMISSAO.Id
            _CtaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _CtaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_CtaReceber.NumErro, _CtaReceber.MsgErro, _CtaReceber.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaCodIndPorNumTitSeqTitDtEmissao(41)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            ' connection.Close()
        End Try

        Return _CtaReceber

    End Function

    Public Function BuscaExportPgtosAtivo(ByVal data As Date, ByVal empresa As String) As List(Of PagamentoAtivo)

        Dim lstContas As New List(Of PagamentoAtivo)
        Try

            Using con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
                con.Open()

                Dim cmd As New SqlCommand
                With cmd
                    .Connection = con
                    .CommandTimeout = DadosGenericos.Timeout.Query
                    .CommandText = "P_ExpPgtosAtivo"
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add(New SqlParameter("@AnoMesDia", Funcoes.FormatDate(data, 0).Replace("-", "")))
                    .Parameters.Add(New SqlParameter("@CodEmpresa", empresa))
                End With

                Dim reader As SqlDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                lstContas.Clear()
                While (reader.Read())
                    Dim _PagAtivo As New PagamentoAtivo()
                    With _PagAtivo
                        .ContaReceber.NumTit = IIf(IsDBNull(reader("NumTit")), "", reader("NumTit"))
                        .ContaReceber.SeqTit = IIf(IsDBNull(reader("SeqTit")), "", reader("SeqTit"))
                        .ContaReceber.DtEmissao = IIf(IsDBNull(reader("DtEmissao")), "", reader("DtEmissao"))
                        .ContaReceber.TipoPagto = IIf(IsDBNull(reader("TipoPagto")), "", reader("TipoPagto"))
                        .ContaReceber.TipoCartaoCred = IIf(IsDBNull(reader("TipoCartaoCred")), "", reader("TipoCartaoCred"))
                        .NumParcelas = IIf(IsDBNull(reader("NumParc")), "", reader("NumParc"))
                        .CtaCtblCred = IIf(IsDBNull(reader("CtaCtblCred")), "", reader("CtaCtblCred"))
                        .ContaContabilDeb = IIf(IsDBNull(reader("CtaCtblDeb")), "", reader("CtaCtblDeb"))
                        .CodHistCred = IIf(IsDBNull(reader("CodHistCred")), "", reader("CodHistCred"))
                        .ComplHistCred = IIf(IsDBNull(reader("ComplHistCred")), "", reader("ComplHistCred"))
                        .CodHistDeb = IIf(IsDBNull(reader("CodHistDeb")), "", reader("CodHistDeb"))
                        .ComplHistDeb = IIf(IsDBNull(reader("ComplHistDeb")), "", reader("ComplHistDeb"))
                        .VlrLcto = IIf(IsDBNull(reader("VlrLcto")), "", reader("VlrLcto"))
                        .VlrLctoJuros = IIf(IsDBNull(reader("VlrLctoJuros")), "", reader("VlrLctoJuros"))
                        .CtaCtblCredJuros = IIf(IsDBNull(reader("CtaCtblCredJuros")), "", reader("CtaCtblCredJuros"))
                        .CodHistCredJuros = IIf(IsDBNull(reader("CodHistCredJuros")), "", reader("CodHistCredJuros"))
                        .ObsJuros = IIf(IsDBNull(reader("ObsJuros")), "", reader("ObsJuros"))
                        .VlrLctoDesc = IIf(IsDBNull(reader("VlrLctoDesc")), "", reader("VlrLctoDesc"))
                        .CtaCtblDebDesc = IIf(IsDBNull(reader("CtaCtblDebDesc")), "", reader("CtaCtblDebDesc"))
                        .ObsDesc = IIf(IsDBNull(reader("ObsDesc")), "", reader("ObsDesc"))
                        .CodHistDebDesc = IIf(IsDBNull(reader("CodHistDebDesc")), "", reader("CodHistDebDesc"))
                        .CtaCtblDebEncCA = IIf(IsDBNull(reader("CtaCtblDebEncCA")), "", reader("CtaCtblDebEncCA"))
                        .CodHistDebEncCA = IIf(IsDBNull(reader("CodHistDebEncCA")), "", reader("CodHistDebEncCA"))
                        .ObsEncCA = IIf(IsDBNull(reader("ObsEncCA")), "", reader("ObsEncCA"))
                        .ObsCred = IIf(IsDBNull(reader("ObsCred")), "", reader("ObsCred"))
                        .ObsDeb = IIf(IsDBNull(reader("ObsDeb")), "", reader("ObsDeb"))


                        .Sucesso = True
                        .TipoErro = DadosGenericos.TipoErro.None
                    End With

                    lstContas.Add(_PagAtivo)
                End While

                con.Close()
            End Using

        Catch ex As Exception
            Dim _Retorno As New Retorno()
            With _Retorno
                .Sucesso = False
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
                .MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAEXPORTPGTOSATIVO.Descricao & ex.Message
                .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAEXPORTPGTOSATIVO.Id
                'CRIAR LOG NO WINDOWS
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")

        End Try

        Return lstContas
    End Function


    Public Function PesquisaContasAReceberPorNumTitSeqTitDtEmissao(ByVal strNumTit As String, ByVal strSeqTit As String) As ContaReceber

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _CtaReceber As New ContaReceber
        Dim i As Integer = 0
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_PesquisaContasAReceberPorNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            'Command.Parameters.Add(New SqlParameter("@DtEmissao", Convert.ToDateTime(dtEmissao)))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read

                    _CtaReceber.NumTit = rdr("NUMTIT")
                    _CtaReceber.SeqTit = rdr("SEQTIT")
                    _CtaReceber.DtEmissao = Convert.ToDateTime(rdr("DtEmissao"))
                    _CtaReceber.TipoDup = rdr("TipoDup")
                    _CtaReceber.IdFilial = rdr("IdFilial")
                    If IsDBNull(rdr("IdBU")) Then
                        _CtaReceber.IdBU = 0
                    Else
                        _CtaReceber.IdBU = rdr("idBU")
                    End If

                    _CtaReceber.Sucesso = True
                    _CtaReceber.TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop

            Else
                'lstCtaReceber.Add(New ContaReceber)
                _CtaReceber.Sucesso = False
                _CtaReceber.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _CtaReceber.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _CtaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception
            'lstCtaReceber.Add(New ContaReceber)
            _CtaReceber.Sucesso = False
            _CtaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_PESQUISACONTASARECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _CtaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_PESQUISACONTASARECEBERPORNUMTITSEQTITDTEMISSAO.Id
            _CtaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _CtaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_CtaReceber.NumErro, _CtaReceber.MsgErro, _CtaReceber.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return _CtaReceber

    End Function

    'Valida data de emissão do titulo com a data atual
    Public Function ValidaBxTit(ByVal strNumTit As String, ByVal strSeqTit As String) As Retorno

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim retorno As New Retorno

        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_ValidaBxTit", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()

                If rdr("VendServ").ToString().Equals("S") Then
                    retorno.Sucesso = True
                    retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                Else
                    retorno.Sucesso = True
                    retorno.TipoErro = DadosGenericos.TipoErro.None
                End If

            Else

                retorno.Sucesso = False
                retorno.TipoErro = DadosGenericos.TipoErro.Funcional

            End If

            rdr.Close()
            Command.Dispose()

        Catch ex As Exception

            retorno.Sucesso = False
            retorno.MsgErro = ex.Message
            retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(retorno.NumErro, retorno.MsgErro, retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")

        End Try

        Return retorno

    End Function


    Public Function Inadimplentes(ByVal strTipoDup As String, ByVal strgCodStatus As String, ByVal strPlVda As String,
                                  ByVal strIdFilial As String, ByVal strComodato As String, ByVal strChkTitNeg As String,
                                  ByVal strChkCobExt As String, ByVal strCanal As String, ByVal strTpEstabelecimento As String, ByVal strClienteDe As String) As List(Of ContaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstContas As New List(Of ContaReceber)
        Dim command As New SqlCommand("P_Inadimplentes", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim i As Integer = 0
        Dim rdr As SqlDataReader
        Try
            connection.Open()


            command.Parameters.Add(New SqlParameter("@TipoDup", IIf(String.IsNullOrEmpty(strTipoDup), DBNull.Value, strTipoDup)))
            command.Parameters.Add(New SqlParameter("@gCodStatus", IIf(String.IsNullOrEmpty(strgCodStatus), DBNull.Value, strgCodStatus)))
            command.Parameters.Add(New SqlParameter("@PlVda", IIf(String.IsNullOrEmpty(strPlVda), DBNull.Value, strPlVda)))
            command.Parameters.Add(New SqlParameter("@IdFilial", IIf(String.IsNullOrEmpty(strIdFilial), DBNull.Value, strIdFilial)))
            command.Parameters.Add(New SqlParameter("@Comodato", IIf(String.IsNullOrEmpty(strComodato), DBNull.Value, strComodato)))
            command.Parameters.Add(New SqlParameter("@chkTitNeg", IIf(String.IsNullOrEmpty(strChkTitNeg), DBNull.Value, strChkTitNeg)))
            command.Parameters.Add(New SqlParameter("@chkCobExt", IIf(String.IsNullOrEmpty(strChkCobExt), DBNull.Value, strChkCobExt)))
            command.Parameters.Add(New SqlParameter("@Canal", IIf(String.IsNullOrEmpty(strCanal), DBNull.Value, strCanal)))
            command.Parameters.Add(New SqlParameter("@TpEstabelecimento", IIf(String.IsNullOrEmpty(strTpEstabelecimento), DBNull.Value, strTpEstabelecimento)))
            command.Parameters.Add(New SqlParameter("@ClienteDe", IIf(String.IsNullOrEmpty(strClienteDe), DBNull.Value, strClienteDe)))

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                Do While (rdr.Read())
                    lstContas.Add(New ContaReceber)
                    lstContas(i).CodIntClie = IIf(IsDBNull(rdr("Unidade")), Nothing, rdr("Unidade"))
                    lstContas(i).RazaoSocial = IIf(IsDBNull(rdr("RazaoSocial")), Nothing, rdr("RazaoSocial"))

                    lstContas(i).Sucesso = True
                    lstContas(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            Else
                lstContas.Add(New ContaReceber)
                lstContas(0).Sucesso = False
                lstContas(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstContas(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContas(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If
            rdr.Close()
        Catch ex As Exception
            lstContas.Add(New ContaReceber)
            lstContas(0).Sucesso = False
            lstContas(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContas(0).MsgErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTES.Descricao & ex.Message
            lstContas(0).NumErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTES.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContas(0).NumErro, lstContas(0).MsgErro, lstContas(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstContas
    End Function


    Public Function InadimplentesAnalitico(ByVal strPeriodo As String, ByVal strAgrupamento As String, ByVal strTipoDup As String,
                                           ByVal strGCodStatus As String, ByVal strPlVda As String, ByVal dtDtInicial As String,
                                           ByVal dtDtFinal As String, ByVal strAnoMes As String, ByVal strIdFilial As String,
                                           ByVal strComodato As String, ByVal strChkTitNeg As String, ByVal strChkCobExt As String, ByVal strCanal As String, ByVal strTpEstabelecimento As String, ByVal strClienteDe As String) As List(Of ContaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstContas As New List(Of ContaReceber)
        Dim command As New SqlCommand("P_InadimplentesAnalitico", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim i As Integer = 0
        Dim rdr As SqlDataReader
        Try
            connection.Open()


            command.Parameters.Add(New SqlParameter("@Periodo", IIf(String.IsNullOrEmpty(strPeriodo), DBNull.Value, strPeriodo)))
            command.Parameters.Add(New SqlParameter("@Agrupamento", IIf(String.IsNullOrEmpty(strAgrupamento), DBNull.Value, strAgrupamento)))
            command.Parameters.Add(New SqlParameter("@TipoDup", IIf(String.IsNullOrEmpty(strTipoDup), DBNull.Value, strTipoDup)))
            command.Parameters.Add(New SqlParameter("@gCodStatus", IIf(String.IsNullOrEmpty(strGCodStatus), DBNull.Value, strGCodStatus)))
            command.Parameters.Add(New SqlParameter("@PlVda", IIf(String.IsNullOrEmpty(strPlVda), DBNull.Value, strPlVda)))
            command.Parameters.Add(New SqlParameter("@DtInicial", IIf(String.IsNullOrEmpty(dtDtInicial), DBNull.Value, dtDtInicial)))
            command.Parameters.Add(New SqlParameter("@DtFinal", IIf(String.IsNullOrEmpty(dtDtFinal), DBNull.Value, dtDtFinal)))
            command.Parameters.Add(New SqlParameter("@AnoMes", IIf(String.IsNullOrEmpty(strAnoMes), DBNull.Value, strAnoMes)))
            command.Parameters.Add(New SqlParameter("@IdFilial", IIf(String.IsNullOrEmpty(strIdFilial), DBNull.Value, strIdFilial)))
            command.Parameters.Add(New SqlParameter("@Comodato", IIf(String.IsNullOrEmpty(strComodato), DBNull.Value, strComodato)))
            command.Parameters.Add(New SqlParameter("@ChkTitNeg", IIf(String.IsNullOrEmpty(strChkTitNeg), DBNull.Value, strChkTitNeg)))
            command.Parameters.Add(New SqlParameter("@ChkCobExt", IIf(String.IsNullOrEmpty(strChkCobExt), DBNull.Value, strChkCobExt)))
            command.Parameters.Add(New SqlParameter("@Canal", IIf(String.IsNullOrEmpty(strCanal), DBNull.Value, strCanal)))
            command.Parameters.Add(New SqlParameter("@TpEstabelecimento", IIf(String.IsNullOrEmpty(strTpEstabelecimento), DBNull.Value, strTpEstabelecimento)))
            command.Parameters.Add(New SqlParameter("@ClienteDe", IIf(String.IsNullOrEmpty(strClienteDe), DBNull.Value, strClienteDe)))

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                Do While (rdr.Read())
                    lstContas.Add(New ContaReceber)
                    lstContas(i).CodIntClie = IIf(IsDBNull(rdr("CodIntClie")), Nothing, rdr("CodIntClie"))
                    lstContas(i).RazaoSocial = IIf(IsDBNull(rdr("RazaoSocial")), Nothing, rdr("RazaoSocial"))

                    lstContas(i).Sucesso = True
                    lstContas(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            Else
                lstContas.Add(New ContaReceber)
                lstContas(0).Sucesso = False
                lstContas(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstContas(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContas(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If
            rdr.Close()
        Catch ex As Exception
            lstContas.Add(New ContaReceber)
            lstContas(0).Sucesso = False
            lstContas(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContas(0).MsgErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTESANALITICO.Descricao & ex.Message
            lstContas(0).NumErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTESANALITICO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContas(0).NumErro, lstContas(0).MsgErro, lstContas(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstContas
    End Function


    Public Function BuscaTitulosPendentesUnidade(ByVal strCodIntClie As String) As List(Of ContaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstContas As New List(Of ContaReceber)
        Dim command As New SqlCommand("P_BuscaTitulosPendentesUnidade", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim i As Integer = 0
        Dim rdr As SqlDataReader
        Try
            connection.Open()


            command.Parameters.Add(New SqlParameter("@CODINTCLIE", IIf(String.IsNullOrEmpty(strCodIntClie), DBNull.Value, strCodIntClie)))

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                Do While (rdr.Read())
                    lstContas.Add(New ContaReceber)
                    lstContas(i).NumTit = IIf(IsDBNull(rdr("NumTit")), Nothing, rdr("NumTit"))
                    lstContas(i).SeqTit = IIf(IsDBNull(rdr("SeqTit")), Nothing, rdr("SeqTit"))
                    lstContas(i).DtEmissao = IIf(IsDBNull(rdr("DtEmissao")), Nothing, rdr("DtEmissao"))
                    lstContas(i).IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))
                    lstContas(i).DtVcto = IIf(IsDBNull(rdr("DtVcto")), Nothing, rdr("DtVcto"))
                    lstContas(i).VlrInd = IIf(IsDBNull(rdr("VlrInd")), Nothing, rdr("VlrInd"))
                    lstContas(i).CodInd = IIf(IsDBNull(rdr("CodInd")), Nothing, rdr("CodInd"))
                    lstContas(i).Saldo = IIf(IsDBNull(rdr("Saldo")), Nothing, rdr("Saldo"))
                    lstContas(i).TipoDup = IIf(IsDBNull(rdr("TipoDup")), Nothing, rdr("TipoDup"))

                    lstContas(i).Sucesso = True
                    lstContas(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            Else
                lstContas.Add(New ContaReceber)
                lstContas(0).Sucesso = False
                lstContas(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstContas(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContas(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If
            rdr.Close()
        Catch ex As Exception
            lstContas.Add(New ContaReceber)
            lstContas(0).Sucesso = False
            lstContas(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContas(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSPENDENTESUNIDADE.Descricao & ex.Message
            lstContas(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCATITULOSPENDENTESUNIDADE.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContas(0).NumErro, lstContas(0).MsgErro, lstContas(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstContas
    End Function


    Public Function BuscaContaReceberGeraRemessa_CartaoCredito_Cielo(ByVal CodIntClie As String,
                                                                    ByVal DtInicio As String,
                                                                    ByVal DtFim As String,
                                                                    ByVal TitInicio As String,
                                                                    ByVal TitFim As String,
                                                                    ByVal Connection As SqlConnection,
                                                                    ByVal Transaction As SqlTransaction,
                                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa_CartaoCredito_Cielo", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(IsDBNull(rdr("ObsTit")), "", rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.MsgBole = IIf(IsDBNull(rdr("MsgBole")), "", rdr("MsgBole"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Financeiro = New ClienteFinanceiro
                    lstExportar(i).Cliente.Financeiro.TipoCartaoCred = rdr("TipoCartaoCred")
                    lstExportar(i).Cliente.Financeiro.CartaoCred = rdr("CartaoCred")
                    lstExportar(i).Cliente.Financeiro.ValidCCred = rdr("ValidCCred")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function

    Public Function GeraArquivoNAVRecebimentosRelat(ByVal otb As DataTable,
                                                    ByVal DtLctoIni As DateTime,
                                                    ByVal DtLctoFim As DateTime,
                                                    ByVal NumNfIni As String,
                                                    ByVal NumNfFim As String,
                                                    ByVal NumCta As String
                                                   ) As Retorno
        Dim _Retorno As New Retorno
        Dim rdr As SqlDataReader = Nothing
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim cmd As New SqlCommand("P_GeraArquivoNAVRecebimentosRelat", Connection)

        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = DadosGenericos.Timeout.Query

            ' Parâmetros
            cmd.Parameters.Add(New SqlParameter("@DtLctoIni", DtLctoIni))
            cmd.Parameters.Add(New SqlParameter("@DtLctoFim", DtLctoFim))
            cmd.Parameters.Add(New SqlParameter("@NumNfIni", NumNfIni))
            cmd.Parameters.Add(New SqlParameter("@NumNfFim", NumNfFim))
            cmd.Parameters.Add(New SqlParameter("@NumCta", NumCta))

            Connection.Open()

            rdr = cmd.ExecuteReader()
            If rdr.HasRows Then
                otb.Load(rdr)
                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Else
                _Retorno.Sucesso = False
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            End If

            rdr.Close()
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = "GeraArquivoNAVRecebimentosRelat"
            _Retorno.MsgErro = "Erro ao buscar dados: " & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            Connection.Close()
            cmd.Dispose()
            Connection.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function AtualizaInadimplentesExcel() As Retorno '#16#
        Dim _Retorno As New Retorno
        Dim Transacao As SqlTransaction
        Using Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            Connection.Open()
            Transacao = Connection.BeginTransaction

            Try
                'Informa a procedure
                Dim Command As SqlCommand = New SqlCommand("P_InserirInadimplentes", Connection)
                Command.Transaction = Transacao
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Query
                'Executa a procedure
                Command.ExecuteNonQuery()
                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

            Catch ex As Exception
                _Retorno.Sucesso = False
                _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERIRCONTASRECEBER.Id
                _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERIRCONTASRECEBER.Descricao & ex.Message
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InserirContaReceber(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
                Transacao.Rollback()
                Connection.Close()

            End Try

        End Using
        Return _Retorno
    End Function

    Public Function BuscaInadimplentesExcel(ByVal oTable As DataTable, ByVal DataInicial As DateTime, ByVal DataFinal As DateTime) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_RelInadimplentesExcel", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Boletagem

                'parametro
                Command.Parameters.Add(New SqlParameter("@DataInicial", DataInicial))
                Command.Parameters.Add(New SqlParameter("@DataFinal", DataFinal))

                Using rdr As SqlDataReader = Command.ExecuteReader()

                    If rdr.HasRows Then

                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None

                    Else
                        _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                        _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                        _Retorno.Sucesso = False
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                    End If
                End Using

            Catch ex As Exception

                _Retorno.NumErro = "BuscaInadimplentesExcel"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: SatBC - Classe: ConsultaContasReceber - Função: BuscaInadimplentesExcel", "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function


    Public Function BuscarTitulosBoletoUnificado(ByVal idBU As Integer) As List(Of ContaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstTitulos As New List(Of ContaReceber)
        Dim command As New SqlCommand("P_BuscarTitulosBoletoUnificado", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim i As Integer = 0
        Dim rdr As SqlDataReader
        Try
            connection.Open()

            command.Parameters.Add(New SqlParameter("@IdBU", IIf(String.IsNullOrEmpty(idBU), DBNull.Value, idBU)))

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                Do While (rdr.Read())
                    lstTitulos.Add(New ContaReceber)
                    lstTitulos(i).NumTit = IIf(IsDBNull(rdr("NumTit")), Nothing, rdr("NumTit"))
                    lstTitulos(i).SeqTit = IIf(IsDBNull(rdr("SeqTit")), Nothing, rdr("SeqTit"))
                    lstTitulos(i).DtEmissao = IIf(IsDBNull(rdr("DtEmissao")), Date.MinValue, rdr("DtEmissao"))
                    lstTitulos(i).IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))
                    lstTitulos(i).VlrPrevPgDI = IIf(IsDBNull(rdr("VlrPrevPgDI")), Nothing, rdr("VlrPrevPgDI"))
                    lstTitulos(i).TipoDup = IIf(IsDBNull(rdr("TipoDup")), Nothing, rdr("TipoDup"))
                    lstTitulos(i).VlrInd = IIf(IsDBNull(rdr("VlrInd")), Nothing, rdr("VlrInd"))
                    lstTitulos(i).Situacao = IIf(IsDBNull(rdr("Situacao")), Nothing, rdr("Situacao"))
                    lstTitulos(i).DtVcto = IIf(IsDBNull(rdr("DtVcto")), Nothing, rdr("DtVcto"))
                    lstTitulos(i).DtPrevPgDI = IIf(IsDBNull(rdr("DtPrevPgDI")), Nothing, rdr("DtPrevPgDI"))
                    lstTitulos(i).SituacaoDesc = IIf(IsDBNull(rdr("SituacaoDesc")), Nothing, rdr("SituacaoDesc"))
                    lstTitulos(i).CodClie = IIf(IsDBNull(rdr("CodClie")), Nothing, rdr("CodClie"))
                    lstTitulos(i).CodReclamacaoJustificativaDesconto = IIf(IsDBNull(rdr("CodReclamacaoJustificativaDesconto")), Nothing, rdr("CodReclamacaoJustificativaDesconto"))

                    lstTitulos(i).Sucesso = True
                    lstTitulos(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            Else
                lstTitulos.Add(New ContaReceber)
                lstTitulos(0).Sucesso = False
                lstTitulos(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstTitulos(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstTitulos(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If
            rdr.Close()
        Catch ex As Exception
            lstTitulos.Add(New ContaReceber)
            lstTitulos(0).Sucesso = False
            lstTitulos(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstTitulos(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOSBOLETOUNIFICADO.Descricao & ex.Message
            lstTitulos(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOSBOLETOUNIFICADO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstTitulos(0).NumErro, lstTitulos(0).MsgErro, lstTitulos(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstTitulos
    End Function

    Public Function BuscarBoletoUnificado(ByVal idBu As Integer) As ContaReceber
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _boleto As New ContaReceber
        Dim command As New SqlCommand("P_BuscarBoletoUnificado", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim rdr As SqlDataReader

        Try
            connection.Open()

            command.Parameters.Add(New SqlParameter("@IdBU", IIf(String.IsNullOrEmpty(idBu), DBNull.Value, idBu)))

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                rdr.Read()

                _boleto.RazaoSocial = IIf(IsDBNull(rdr("RazaoSocial")), Nothing, rdr("RazaoSocial"))
                _boleto.NumTit = IIf(IsDBNull(rdr("NumTit")), Nothing, rdr("NumTit"))
                _boleto.SeqTit = IIf(IsDBNull(rdr("SeqTit")), Nothing, rdr("SeqTit"))
                _boleto.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))
                _boleto.VlrInd = IIf(IsDBNull(rdr("VlrInd")), Nothing, rdr("VlrInd"))
                _boleto.DtEmissao = IIf(IsDBNull(rdr("DtEmissao")), Nothing, rdr("DtEmissao"))
                _boleto.Situacao = IIf(IsDBNull(rdr("Situacao")), Nothing, rdr("Situacao"))
                _boleto.SituacaoDesc = IIf(IsDBNull(rdr("SituacaoDesc")), Nothing, rdr("SituacaoDesc"))
                _boleto.CodClie = IIf(IsDBNull(rdr("CodClie")), Nothing, rdr("CodClie"))
                _boleto.TipoDup = IIf(IsDBNull(rdr("TipoDup")), Nothing, rdr("TipoDup"))
                _boleto.CodReclamacaoJustificativaDesconto = IIf(IsDBNull(rdr("CodReclamacaoJustificativaDesconto")), Nothing, rdr("CodReclamacaoJustificativaDesconto"))

                _boleto.Sucesso = True
                _boleto.TipoErro = DadosGenericos.TipoErro.None


            Else

                _boleto.Sucesso = False
                _boleto.TipoErro = DadosGenericos.TipoErro.Funcional
                _boleto.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _boleto.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If
            rdr.Close()
        Catch ex As Exception

            _boleto.Sucesso = False
            _boleto.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _boleto.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARBOLETOUNIFICADO.Descricao & ex.Message
            _boleto.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARBOLETOUNIFICADO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_boleto.NumErro, _boleto.MsgErro, _boleto.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _boleto
    End Function

    Public Function BuscaClientesProdefence(ByVal oTable As DataTable, ByVal DtConexaoIni As String, ByVal DtConexaoFim As String, ByVal DtVctoIni As String, ByVal DtVctoFim As String) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_RelatClientesProdefence", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Query

                'parametro
                Command.Parameters.Add(New SqlParameter("@DtConexaoIni", DtConexaoIni))
                Command.Parameters.Add(New SqlParameter("@DtConexaoFim", DtConexaoFim))
                Command.Parameters.Add(New SqlParameter("@DtVctoIni", DtVctoIni))
                Command.Parameters.Add(New SqlParameter("@DtVctoFim", DtVctoFim))

                Using rdr As SqlDataReader = Command.ExecuteReader()

                    If rdr.HasRows Then

                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None

                    Else
                        _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                        _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                        _Retorno.Sucesso = False
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                    End If
                End Using

            Catch ex As Exception

                _Retorno.NumErro = "BuscaClientesProdefence"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function

    Public Function BuscaDadosContasAReceber(ByVal NumTit As String, ByVal SeqTit As String, ByVal connection As SqlConnection, ByVal Transation As SqlTransaction) As ContaReceber

        Dim _Retorno = New Retorno
        Dim rdr As SqlDataReader
        Dim contaReceber As New ContaReceber
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaDadosContasAReceber", connection)
            Command.Transaction = Transation
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure

            Command.Parameters.Add(New SqlParameter("@NumTit", NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", SeqTit))
            'Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    contaReceber.CodBanco = rdr("CodBanco")
                    contaReceber.CodAgen = rdr("CodAgen")
                    contaReceber.NumCta = rdr("NumCta")
                    contaReceber.DtVcto = rdr("DtVcto")
                    contaReceber.CodIntClie = rdr("CodIntClie")

                    contaReceber.Sucesso = True
                    contaReceber.TipoErro = DadosGenericos.TipoErro.None

                Loop

                rdr.Close()
            Else
                contaReceber.Sucesso = False
                contaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            contaReceber.Sucesso = False
            contaReceber.MsgErro = ex.Message
            contaReceber.NumErro = "BuscaDadosContasAReceber"
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaDadosContasAReceber", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return contaReceber
    End Function

    Public Function RelatTitAbertosExcel(ByVal filtro As Integer, ByRef dtData As DataTable) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_RelatTitAbertosExcel", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Faturamento
        Dim _Retorno As New Retorno
        Try
            'Abre a conexão
            connection.Open()

            'Adiciona os parâmetros
            command.Parameters.Add(New SqlParameter("@filtro", filtro))

            'Preenche o data table
            dtData.Load(command.ExecuteReader())

            'Define o sucesso
            _Retorno.Sucesso = True
        Catch ex As Exception
            _Retorno.MsgErro = ex.Message
            _Retorno.Sucesso = False
        Finally
            connection.Close()
        End Try
        Return _Retorno
    End Function

    Public Function BuscarTituloFinanciadoCliente(ByVal codClie As String, connection As SqlConnection, trans As SqlTransaction) As ContaReceber


        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaReceber
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarTituloFinanciadoCliente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodClie", codClie))

            Command.Transaction = trans

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                contaReceber = New ContaReceber

                contaReceber.NumTit = rdr("NumTit")
                contaReceber.SeqTit = rdr("Seqtit")
                contaReceber.DtEmissao = rdr("dtEmissao")
                contaReceber.IdFilial = rdr("IdFilial")
                contaReceber.CodClie = rdr("CodClie")
                contaReceber.TipoDup = rdr("TipoDup")
                contaReceber.QtdFinanciadosAberto = rdr("QtdFinanciadosAberto")


                contaReceber.Sucesso = True
                contaReceber.TipoErro = DadosGenericos.TipoErro.None

            Else
                contaReceber = New ContaReceber
                contaReceber.Sucesso = False
                contaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            contaReceber = New ContaReceber
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOFINANCIADOCLIENTE.Descricao & ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOFINANCIADOCLIENTE.Id
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return contaReceber

    End Function

    Public Function BuscaContaReceberGeraRemessa341BoletoUnificado(ByVal strCodBanco As String,
                                                                   ByVal strCodAgencia As String,
                                                                   ByVal strNumCta As String,
                                                                   ByVal strTipoPagto As String,
                                                                   ByVal CodIntClie As String,
                                                                   ByVal DtInicio As String,
                                                                   ByVal DtFim As String,
                                                                   ByVal TitInicio As String,
                                                                   ByVal TitFim As String,
                                                                   ByVal Connection As SqlConnection,
                                                                   ByVal Transaction As SqlTransaction) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa341BoletoUnificado", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA341BOLETOUNIFICADO.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA341BOLETOUNIFICADO.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function

    Public Function BuscarMultasAbertas(ByVal dtIni As DateTime, dtFim As DateTime, ByRef dtData As DataTable) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_BuscarMultasAbertas", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()
            command.Parameters.Add(New SqlParameter("@dtIni", IIf(String.IsNullOrEmpty(dtIni), DBNull.Value, dtIni)))
            command.Parameters.Add(New SqlParameter("@dtFim", IIf(String.IsNullOrEmpty(dtFim), DBNull.Value, dtFim)))

            'Preenche o data table
            dtData.Load(command.ExecuteReader())

            _Retorno.Sucesso = True
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARMULTASABERTAS.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARMULTASABERTAS.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function ConsultaContaReceberNumTitSeqTit(ByVal NumTit As String, ByVal SeqTit As String, connection As SqlConnection, trans As SqlTransaction) As ContaReceber


        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaReceber
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_ConsultaContaReceberNumTitSeqTit", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", SeqTit))

            Command.Transaction = trans

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                contaReceber = New ContaReceber

                contaReceber.CodIntClie = rdr("CodIntClie")

                contaReceber.Sucesso = True
                contaReceber.TipoErro = DadosGenericos.TipoErro.None

            Else
                contaReceber = New ContaReceber
                contaReceber.Sucesso = False
                contaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
                contaReceber.MsgErro = "Título não encontrado."
            End If
            rdr.Close()
        Catch ex As Exception

            contaReceber = New ContaReceber
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOFINANCIADOCLIENTE.Id
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return contaReceber

    End Function

    Public Function BuscaContasReceberRelatorioExcel(ByVal oTable As DataTable,
                                                     ByVal Periodo As String,
                                                     ByVal DtInicial As String,
                                                     ByVal DtFinal As String,
                                                     ByVal strSit1 As String,
                                                     ByVal strSit2 As String,
                                                     ByVal strSit3 As String,
                                                     ByVal strTipoPagto As String,
                                                     ByVal strPlanoVenda As String,
                                                     ByVal strRecebedor As String,
                                                     ByVal strFilial As String,
                                                     ByVal strSitTitEvento As String,
                                                     ByVal strTipoBaixa As String,
                                                     ByVal strCodClie As String,
                                                     ByVal strPortador As String,
                                                     ByVal strCategora As String,
                                                     ByVal strTitIni As String,
                                                     ByVal strTitFim As String,
                                                     ByVal strContaRecebedora As String) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_CRecRelNew", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = 7200


                'parametro
                Command.Parameters.Add(New SqlParameter("@Periodo", Periodo))
                Command.Parameters.Add(New SqlParameter("@DtInicial", DtInicial))
                Command.Parameters.Add(New SqlParameter("@DtFinal", DtFinal))
                Command.Parameters.Add(New SqlParameter("@Sit1", strSit1))
                Command.Parameters.Add(New SqlParameter("@Sit2", strSit2))
                Command.Parameters.Add(New SqlParameter("@Sit3", strSit3))
                Command.Parameters.Add(New SqlParameter("@TpPagto", strTipoPagto))
                Command.Parameters.Add(New SqlParameter("@PlVda", strPlanoVenda))
                Command.Parameters.Add(New SqlParameter("@Recebedor", strRecebedor))
                Command.Parameters.Add(New SqlParameter("@Filial", strFilial))
                Command.Parameters.Add(New SqlParameter("@SitTitEvento", strSitTitEvento))
                Command.Parameters.Add(New SqlParameter("@TipoBaixa", strTipoBaixa))
                Command.Parameters.Add(New SqlParameter("@CodClie", strCodClie))
                Command.Parameters.Add(New SqlParameter("@Portador", strPortador))
                Command.Parameters.Add(New SqlParameter("@Categoria", strCategora))
                Command.Parameters.Add(New SqlParameter("@TitIni", strTitIni))
                Command.Parameters.Add(New SqlParameter("@TitFim", strTitFim))
                Command.Parameters.Add(New SqlParameter("@ContaRecebedora", strContaRecebedora))


                Using rdr As SqlDataReader = Command.ExecuteReader()

                    If rdr.HasRows Then

                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None

                    Else
                        _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                        _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                        _Retorno.Sucesso = False
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                    End If
                End Using

            Catch ex As Exception

                _Retorno.NumErro = "BuscaContasReceberRelatorioExcel"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasReceber - Classe: ConsultaContasReceber - Função: BuscaContasReceberRelatorioExcel", "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function


    Public Function BuscaContasReceberDadosCliente(ByVal oTable As DataTable,
                                                     ByVal strMensagem As String,
                                                      ByVal strChkEmail As String) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_BuscaContasReceberDadosCLiente", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Query

                'parametro
                Command.Parameters.Add(New SqlParameter("@Mensagem", strMensagem))
                Command.Parameters.Add(New SqlParameter("@ChkEmail", strChkEmail))

                Using rdr As SqlDataReader = Command.ExecuteReader()

                    If rdr.HasRows Then

                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None

                    Else
                        _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                        _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                        _Retorno.Sucesso = False
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                    End If
                End Using

            Catch ex As Exception

                _Retorno.NumErro = "BuscaContasReceberDadosCliente"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasReceber - Classe: ConsultaContasReceber - Função: BuscaContasReceberDadosCliente", "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function

    Public Function BuscaAcordos(ByVal DtInicio As String, ByVal DtFinal As String, ByVal Operador As String, ByVal oTable As DataTable) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_BuscaAcordos", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.ContImprodutivos

                'parametro
                Command.Parameters.Add(New SqlParameter("@DtInicio", DtInicio))
                Command.Parameters.Add(New SqlParameter("@DtFim", DtFinal))
                Command.Parameters.Add(New SqlParameter("@Operador", Operador))

                Using rdr As SqlDataReader = Command.ExecuteReader()

                    If rdr.HasRows Then

                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None

                    Else
                        _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                        _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                        _Retorno.Sucesso = False
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                    End If
                End Using

            Catch ex As Exception

                _Retorno.NumErro = "BuscaAcordos"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasReceber - Classe: ConsultaContasReceber - Função: BuscaAcordos", "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function


    Public Function BuscaAcordosNew(ByVal DtInicio As String, ByVal DtFinal As String, ByVal Operador As String, ByVal oTable As DataTable) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_BuscaAcordosCompleto", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.ContImprodutivos

                'parametro
                Command.Parameters.Add(New SqlParameter("@DtInicio", DtInicio))
                Command.Parameters.Add(New SqlParameter("@DtFim", DtFinal))
                Command.Parameters.Add(New SqlParameter("@Operador", Operador))

                Using rdr As SqlDataReader = Command.ExecuteReader()

                    If rdr.HasRows Then

                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None

                    Else
                        _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                        _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                        _Retorno.Sucesso = False
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                    End If
                End Using

            Catch ex As Exception

                _Retorno.NumErro = "BuscaAcordos"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasReceber - Classe: ConsultaContasReceber - Função: BuscaAcordos", "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function

    Public Function BuscaContatosImprodutivos(ByVal DtInicio As String, ByVal DtFinal As String, ByVal Operador As String, ByVal oTable As DataTable) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_BuscaContatosImprodutivos", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.ContImprodutivos

                'parametro
                Command.Parameters.Add(New SqlParameter("@DtInicio", DtInicio))
                Command.Parameters.Add(New SqlParameter("@DtFim", DtFinal))
                Command.Parameters.Add(New SqlParameter("@Operador", Operador))

                Using rdr As SqlDataReader = Command.ExecuteReader()

                    If rdr.HasRows Then

                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None

                    Else
                        _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                        _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                        _Retorno.Sucesso = False
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                    End If
                End Using

            Catch ex As Exception

                _Retorno.NumErro = "BuscaContatosImprodutivos"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasReceber - Classe: ConsultaContasReceber - Função: BuscaContatosImprodutivos", "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function

    Public Function BuscaResultadoCobranca(ByVal dtInicial As String, ByVal dtFinal As String, ByVal oTable As DataTable) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()
            Try
                Dim Command As SqlCommand = New SqlCommand("P_BuscaResultadoCobranca", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.ContImprodutivos

                Command.Parameters.Add(New SqlParameter("@DtIni", Convert.ToDateTime(dtInicial)))
                Command.Parameters.Add(New SqlParameter("@DtFim", Convert.ToDateTime(dtFinal)))

                Using rdr As SqlDataReader = Command.ExecuteReader()
                    If rdr.HasRows Then
                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None
                    Else
                        _Retorno.Sucesso = False
                        _Retorno.MsgErro = ErrorConstants.EXCEPTION_DADOS_INEXISTENTES_NA_TABELA_PARAMETRO.Descricao
                        _Retorno.NumErro = ErrorConstants.EXCEPTION_DADOS_INEXISTENTES_NA_TABELA_PARAMETRO.Id
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                    End If
                End Using

            Catch ex As Exception
                _Retorno.NumErro = "BuscaReslutadoCobranca"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

                'CRIA LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasReceber - Classe: ConsultaContasReceber - Função: BuscaReslutadoCobranca", "08", "Verisure", Environment.MachineName, "2.0", "13")
            End Try
        End Using
        Return _Retorno
    End Function
    Public Function BuscaContasReceberDadosClientesRedes(ByVal oTable As DataTable) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_BuscaContasReceberDadosCLientesRedes", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Faturamento

                Using rdr As SqlDataReader = Command.ExecuteReader()

                    If rdr.HasRows Then

                        oTable.Load(rdr)
                        _Retorno.Sucesso = True
                        _Retorno.TipoErro = DadosGenericos.TipoErro.None

                    Else
                        _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                        _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                        _Retorno.Sucesso = False
                        _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                    End If
                End Using

            Catch ex As Exception

                _Retorno.NumErro = "BuscaContasReceberDadosCliente"
                _Retorno.MsgErro = ex.Message
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasReceber - Classe: ConsultaContasReceber - Função: BuscaContasReceberDadosCliente", "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function

    Public Function BuscarFilaCobrancaClientes(ByRef dt As DataTable) As List(Of FilaCobranca)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstFilaCobranca As New List(Of FilaCobranca)
        Dim fila As New FilaCobranca
        Dim command As New SqlCommand("P_BuscarFilaCobrancaClientes", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim i As Integer = 0
        Dim rdr As SqlDataReader
        Try
            connection.Open()

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                Do While rdr.Read()
                    fila = New FilaCobranca
                    fila.RazaoSocial = rdr("RazaoSocial")
                    fila.Codintclie = rdr("Codintclie")
                    fila.CodFilaOrigem = rdr("CodFilaOrigem")
                    fila.Sucesso = True
                    fila.TipoErro = DadosGenericos.TipoErro.None
                    lstFilaCobranca.Add(fila)
                Loop
                dt.Load(rdr)
            Else
                lstFilaCobranca.Add(New FilaCobranca)
                lstFilaCobranca(0).Sucesso = False
                lstFilaCobranca(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstFilaCobranca(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstFilaCobranca(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If

            rdr.Close()
        Catch ex As Exception
            lstFilaCobranca.Add(New FilaCobranca)
            lstFilaCobranca(0).Sucesso = False
            lstFilaCobranca(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstFilaCobranca(0).MsgErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTESANALITICO.Descricao & ex.Message
            lstFilaCobranca(0).NumErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTESANALITICO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstFilaCobranca(0).NumErro, lstFilaCobranca(0).MsgErro, lstFilaCobranca(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstFilaCobranca
    End Function

    Public Function BuscaContaReceberGeraRemessa041(ByVal strCodBanco As String,
                                                            ByVal strCodAgencia As String,
                                                            ByVal strNumCta As String,
                                                            ByVal strTipoPagto As String,
                                                            ByVal CodIntClie As String,
                                                            ByVal DtInicio As String,
                                                            ByVal DtFim As String,
                                                            ByVal TitInicio As String,
                                                            ByVal TitFim As String,
                                                            ByVal Connection As SqlConnection,
                                                            ByVal Transaction As SqlTransaction,
                                                            Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)

        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa041", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")

                    lstExportar(i).ContaReceber.CodOperacaoConta = rdr("CodOperacaoConta")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa341(11)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function

    Public Function BuscarTituloFF(ByVal codintclie As String) As List(Of ContaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstContasReceber As New List(Of ContaReceber)
        Dim contaReceber As New ContaReceber
        Dim command As New SqlCommand("P_BuscarTituloFF", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim i As Integer = 0
        Dim rdr As SqlDataReader
        Try
            connection.Open()

            command.Parameters.Add(New SqlParameter("@Codintclie", codintclie))

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                Do While rdr.Read()
                    contaReceber = New ContaReceber
                    contaReceber.NumTit = rdr("NumTit")
                    contaReceber.SeqTit = rdr("SeqTit")
                    contaReceber.DtEmissao = rdr("DtEmissao")
                    contaReceber.CodBanco = rdr("CodBanco")
                    contaReceber.CodAgen = rdr("CodAgen")
                    contaReceber.NumCta = rdr("NumCta")
                    contaReceber.VlrInd = rdr("VlrInd")
                    contaReceber.SituacaoDesc = rdr("Situacao")

                    contaReceber.Sucesso = True
                    contaReceber.TipoErro = DadosGenericos.TipoErro.None
                    lstContasReceber.Add(contaReceber)
                Loop
            Else
                lstContasReceber.Add(New ContaReceber)
                lstContasReceber(0).Sucesso = False
                lstContasReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstContasReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContasReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If

            rdr.Close()
        Catch ex As Exception
            lstContasReceber.Add(New ContaReceber)
            lstContasReceber(0).Sucesso = False
            lstContasReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContasReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTESANALITICO.Descricao & ex.Message
            lstContasReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTESANALITICO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContasReceber(0).NumErro, lstContasReceber(0).MsgErro, lstContasReceber(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstContasReceber
    End Function

    Public Function BuscarTitulosMonitoriaFinanciados(ByVal numtitFinanciado As String) As List(Of ContaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstContasReceber As New List(Of ContaReceber)
        Dim contaReceber As New ContaReceber
        Dim command As New SqlCommand("P_BuscarTitulosMonitoriaFinanciados", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim i As Integer = 0
        Dim rdr As SqlDataReader
        Try
            connection.Open()

            command.Parameters.Add(New SqlParameter("@Codintclie", numtitFinanciado))

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                Do While rdr.Read()
                    contaReceber = New ContaReceber
                    contaReceber.NumTit = rdr("NumTit")
                    contaReceber.SeqTit = rdr("SeqTit")
                    contaReceber.DtEmissao = rdr("DtEmissao")
                    contaReceber.CodBanco = rdr("CodBanco")
                    contaReceber.CodAgen = rdr("CodAgen")
                    contaReceber.NumCta = rdr("NumCta")
                    contaReceber.VlrInd = rdr("VlrInd")
                    contaReceber.SituacaoDesc = rdr("Situacao")

                    contaReceber.Sucesso = True
                    contaReceber.TipoErro = DadosGenericos.TipoErro.None
                    lstContasReceber.Add(contaReceber)
                Loop
            Else
                lstContasReceber.Add(New ContaReceber)
                lstContasReceber(0).Sucesso = False
                lstContasReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstContasReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContasReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If

            rdr.Close()
        Catch ex As Exception
            lstContasReceber.Add(New ContaReceber)
            lstContasReceber(0).Sucesso = False
            lstContasReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContasReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTESANALITICO.Descricao & ex.Message
            lstContasReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_INADIMPLENTESANALITICO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContasReceber(0).NumErro, lstContasReceber(0).MsgErro, lstContasReceber(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstContasReceber
    End Function

    Public Function BuscarTodosTitulosFinanciadoCliente(ByVal numTitFinanciado As String) As List(Of ContaReceber)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaReceber
        Dim i As Integer = 0
        Dim lstContasReceber As New List(Of ContaReceber)

        Try
            connection.Open()
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarTodosTitulosFinanciadoCliente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTitFinanciado", numTitFinanciado))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                While rdr.Read()
                    contaReceber = New ContaReceber

                    contaReceber.NumTit = rdr("NumTit")
                    contaReceber.SeqTit = rdr("Seqtit")
                    contaReceber.DtEmissao = rdr("dtEmissao")
                    contaReceber.IdFilial = rdr("IdFilial")
                    contaReceber.CodClie = rdr("CodClie")
                    contaReceber.TipoDup = rdr("TipoDup")
                    contaReceber.SituacaoDesc = rdr("SituacaoDesc")
                    contaReceber.VlrInd = rdr("VlrInd")
                    contaReceber.DtVcto = rdr("DtVcto")
                    contaReceber.Situacao = rdr("Situacao")

                    contaReceber.Sucesso = True
                    contaReceber.TipoErro = DadosGenericos.TipoErro.None

                    lstContasReceber.Add(contaReceber)
                End While



            Else
                contaReceber = New ContaReceber
                contaReceber.Sucesso = False
                contaReceber.TipoErro = DadosGenericos.TipoErro.Funcional
                lstContasReceber.Add(contaReceber)
            End If
            rdr.Close()
        Catch ex As Exception

            contaReceber = New ContaReceber
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOFINANCIADOCLIENTE.Descricao & ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOFINANCIADOCLIENTE.Id
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            lstContasReceber.Add(contaReceber)

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstContasReceber

    End Function

    Public Function BuscarLogEnvioArquivoRemessa(ByVal dtInicio As DateTime, ByVal dtFim As DateTime, ByVal isEnvioBanco As Boolean) As List(Of LogEnvioArquivoRemessa)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Dim rdr As SqlDataReader
        Dim log = New LogEnvioArquivoRemessa
        Dim i As Integer = 0
        Dim lstLogEnvioArquivoRemessa As New List(Of LogEnvioArquivoRemessa)

        Try
            connection.Open()
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarLogEnvioArquivoRemessa", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@DtInicio", dtInicio))
            Command.Parameters.Add(New SqlParameter("@DtFim", dtFim))
            Command.Parameters.Add(New SqlParameter("@IsEnvioBanco", IIf(isEnvioBanco, 1, 0)))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                While rdr.Read()
                    log = New LogEnvioArquivoRemessa

                    log.Id = rdr("Id")
                    log.EnviaBanco = IIf(IsDBNull(rdr("EnviaBanco")), False, rdr("EnviaBanco"))
                    log.CodBanco = IIf(String.IsNullOrEmpty(rdr("CodBanco")), "", rdr("CodBanco"))
                    log.NumCta = rdr("NumCta")
                    log.CodAgen = rdr("CodAgen")
                    log.Tipo = rdr("Tipo")
                    log.Cartao = rdr("Cartao")
                    log.Empresa = rdr("Empresa")
                    log.DtHoraGeracao = rdr("DtHoraGeracao")
                    log.UsrGeracao = rdr("UsrGeracao")
                    log.NumRemessa = rdr("NumRemessa")
                    log.IsOptante = rdr("IsOptante")
                    log.DtHoraGeracao = rdr("DtHoraGeracao")
                    log.UsrEnvioRemessa = rdr("UsrEnvioRemessa")
                    log.NomeUsr = rdr("NomeUsrGeracao")
                    log.NomeUsrEnvioRemessa = rdr("UsrEnvioRemessaDes")
                    log.CartaoDesc = rdr("CartaoDesc")
                    log.EnviaBancoDesc = rdr("EnviaBancoDesc")

                    log.Sucesso = True
                    lstLogEnvioArquivoRemessa.Add(log)

                End While

            Else
                log = New LogEnvioArquivoRemessa
                log.Sucesso = False
                log.TipoErro = DadosGenericos.TipoErro.Funcional
                log.MsgErro = "Nenhum dado encontrado."
                lstLogEnvioArquivoRemessa.Add(log)
            End If
            rdr.Close()
        Catch ex As Exception

            log = New LogEnvioArquivoRemessa
            log.Sucesso = False
            log.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOFINANCIADOCLIENTE.Descricao & ex.Message
            log.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOFINANCIADOCLIENTE.Id
            log.TipoErro = DadosGenericos.TipoErro.Arquitetura
            log.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            lstLogEnvioArquivoRemessa.Add(log)

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(log.NumErro, log.MsgErro, log.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstLogEnvioArquivoRemessa

    End Function

    Public Function SearchDataForEmail(DataForEmail As DataTable) As Retorno
        Dim retorno As New Retorno

        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscarDadosDaBoletagem", connection)

        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            DataForEmail.Load(command.ExecuteReader())

            If (DataForEmail.Rows.Count > 0) Then
                retorno.Sucesso = True
                retorno.TipoErro = DadosGenericos.TipoErro.None
            Else
                retorno.Sucesso = False
                retorno = Funcoes.RetornoFunc("Nenhum registro encontrado.")
            End If

        Catch ex As Exception

            retorno.Sucesso = False
            retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            retorno.MsgErro = ex.Message
            retorno.NumErro = ErrorConstants.EXCEPTION_METODO_SEARCHDATAFOREMAIL.Id

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(retorno.NumErro, retorno.MsgErro, retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            command.Dispose()
            connection.Close()
        End Try

        Return retorno
    End Function

    Public Function BuscaContaReceberGeraRemessa_CartaoCredito_Adyen(ByVal CodIntClie As String,
                                                                    ByVal DtInicio As String,
                                                                    ByVal DtFim As String,
                                                                    ByVal TitInicio As String,
                                                                    ByVal TitFim As String,
                                                                    ByVal Connection As SqlConnection,
                                                                    ByVal Transaction As SqlTransaction,
                                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa_CartaoCredito_Adyen", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(IsDBNull(rdr("ObsTit")), "", rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.MsgBole = IIf(IsDBNull(rdr("MsgBole")), "", rdr("MsgBole"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CodIntClie = rdr("Codintclie")
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Financeiro = New ClienteFinanceiro


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function

    Public Function RelaBuscarTitulosEnviadosAdyen(ByVal otb As DataTable, ByVal DataIni As String, ByVal DataFim As String) As Retorno

        Dim _Retorno As New Retorno
        Dim rdr As SqlDataReader = Nothing
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim cmd As New SqlCommand("P_RelBuscarTitulosEnviadosAdyen", Connection)

        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = DadosGenericos.Timeout.Query

            ' Parâmetros
            cmd.Parameters.Add(New SqlParameter("@DataIni", DataIni))
            cmd.Parameters.Add(New SqlParameter("@DataFim", DataFim))

            Connection.Open()

            rdr = cmd.ExecuteReader()
            If rdr.HasRows Then
                otb.Load(rdr)
                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Else
                _Retorno.Sucesso = False
                _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            End If

            rdr.Close()
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ""
            _Retorno.MsgErro = "Erro ao buscar dados: " & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            Connection.Close()
            cmd.Dispose()
            Connection.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function BuscaContasReceberAdyenAgendamento(ByVal CodIntClie As String,
                                                                    ByVal DtInicio As String,
                                                                    ByVal DtFim As String,
                                                                    ByVal TitInicio As String,
                                                                    ByVal TitFim As String,
                                                                    ByVal usr As String,
                                                                    ByVal dtAgendamento As DateTime,
                                                                    ByVal hrAgendamento As String,
                                                                    ByVal Connection As SqlConnection,
                                                                    ByVal Transaction As SqlTransaction,
                                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_GeraAgendamento_CartaoCredito_Adyen_2_0_1_161", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@Usr", usr))
            Command.Parameters.Add(New SqlParameter("@DtAgendamento", dtAgendamento))
            Command.Parameters.Add(New SqlParameter("@HrAgendamento", hrAgendamento))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")

                    lstExportar(i).ContaReceber.idAgendamento = rdr("IdAgendamento")



                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function

    Public Function BuscaAcordosCompletos(ByVal DtInicio As String, ByVal DtFinal As String, ByVal Operador As String, ByVal oTable As DataTable) As Retorno
        Dim _Retorno As New Retorno
        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()

            Try
                Dim Command As SqlCommand = New SqlCommand("P_RelAcordosCompletos", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.ContImprodutivos

                'parametro
                Command.Parameters.Add(New SqlParameter("@DtInicio", IIf(String.IsNullOrEmpty(DtInicio), DBNull.Value, DtInicio)))
                Command.Parameters.Add(New SqlParameter("@DtFim", IIf(String.IsNullOrEmpty(DtFinal), DBNull.Value, DtFinal)))
                Command.Parameters.Add(New SqlParameter("@Operador", IIf(String.IsNullOrEmpty(Operador), DBNull.Value, Operador)))

                oTable.Load(Command.ExecuteReader)

                If oTable.Rows.Count > 0 Then
                    _Retorno.Sucesso = True
                    _Retorno.TipoErro = DadosGenericos.TipoErro.None
                Else
                    _Retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                    _Retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                    _Retorno.Sucesso = False
                    _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                    _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                End If

            Catch ex As Exception
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAACORDOSCOMPLETOS.Descricao & ex.Message
                _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAACORDOSCOMPLETOS.Id
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")

            End Try

        End Using
        Return _Retorno
    End Function



    Public Function BuscaContaReceberGeraRemessaBU_DC(ByVal strCodBanco As String,
                                                 ByVal strCodAgencia As String,
                                                 ByVal strNumCta As String,
                                                 ByVal CodIntClie As String,
                                                 ByVal DtInicio As String,
                                                 ByVal DtFim As String,
                                                 ByVal TitInicio As String,
                                                 ByVal TitFim As String) As List(Of Exportacao)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa_BU_DC", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = rdr("NumPortador")
                    lstExportar(i).ContaReceber.ObsTit = rdr("ObsTit")
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.NumPortador = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.StatusBco = rdr("CodAgenDeb")
                    lstExportar(i).ContaReceber.NumPortador = rdr("NumCtaDeb")
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.IdFilial = rdr("IdFilial")

                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Cep = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = rdr("NumNFe")
                    lstExportar(i).NotaFiscal.CodVerNFe = rdr("CodVerNFe")


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa(6)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return lstExportar

    End Function


    Public Function BuscaContaReceberGeraRemessa237BuDc(ByVal strCodBanco As String,
                                                 ByVal strCodAgencia As String,
                                                 ByVal strNumCta As String,
                                                 ByVal CodIntClie As String,
                                                 ByVal DtInicio As String,
                                                 ByVal DtFim As String,
                                                 ByVal TitInicio As String,
                                                 ByVal TitFim As String,
                                                 ByVal Connection As SqlConnection,
                                                 ByVal Transaction As SqlTransaction,
                                                 Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)


        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaReceberGeraRemessa237_BU_DC", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            Command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            Command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            Command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            Command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = IIf(IsDBNull(rdr("NumTit")), Nothing, rdr("NumTit"))
                    lstExportar(i).ContaReceber.SeqTit = IIf(IsDBNull(rdr("SeqTit")), Nothing, rdr("SeqTit"))
                    lstExportar(i).ContaReceber.DtEmissao = IIf(IsDBNull(rdr("DtEmissao")), Nothing, rdr("DtEmissao"))
                    lstExportar(i).ContaReceber.DtVcto = IIf(IsDBNull(rdr("DtVcto")), Nothing, rdr("DtVcto"))
                    lstExportar(i).ContaReceber.CodClie = IIf(IsDBNull(rdr("CodClie")), Nothing, rdr("CodClie"))
                    lstExportar(i).ContaReceber.CodInd = IIf(IsDBNull(rdr("CodInd")), Nothing, rdr("CodInd"))
                    lstExportar(i).ContaReceber.VlrInd = IIf(IsDBNull(rdr("VlrInd")), Nothing, rdr("VlrInd"))
                    lstExportar(i).ContaReceber.Situacao = IIf(IsDBNull(rdr("Situacao")), Nothing, rdr("Situacao"))
                    lstExportar(i).ContaReceber.CodBanco = IIf(IsDBNull(rdr("CodBanco")), Nothing, rdr("CodBanco"))
                    lstExportar(i).ContaReceber.CodAgen = IIf(IsDBNull(rdr("CodAgen")), Nothing, rdr("CodAgen"))
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), "", rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(IsDBNull(rdr("ObsTit")), Nothing, rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = IIf(IsDBNull(rdr("StatusBco")), Nothing, rdr("StatusBco"))
                    lstExportar(i).ContaReceber.TipoDup = IIf(IsDBNull(rdr("TipoDup")), Nothing, rdr("TipoDup"))
                    lstExportar(i).ContaReceber.MenBco = IIf(IsDBNull(rdr("MenBco")), Nothing, rdr("MenBco"))
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = IIf(IsDBNull(rdr("TipoPagto")), Nothing, rdr("TipoPagto"))
                    lstExportar(i).ContaReceber.VlrDesconto = IIf(IsDBNull(rdr("VlrDesconto")), Nothing, rdr("VlrDesconto"))
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))

                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = IIf(IsDBNull(rdr("RazaoSocial")), Nothing, rdr("RazaoSocial"))
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = IIf(IsDBNull(rdr("EnvioBanco")), Nothing, rdr("EnvioBanco"))
                    lstExportar(i).Cliente.Contabilidade.IsNNC = IIf(IsDBNull(rdr("IsNNC")), Nothing, rdr("IsNNC"))


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = IIf(IsDBNull(rdr("Endereco")), Nothing, rdr("Endereco"))
                    lstExportar(i).Cliente.Endereco.Bairro = IIf(IsDBNull(rdr("Bairro")), Nothing, rdr("Bairro"))
                    lstExportar(i).Cliente.Endereco.Cidade = IIf(IsDBNull(rdr("Cidade")), Nothing, rdr("Cidade"))
                    lstExportar(i).Cliente.Endereco.Cep = IIf(IsDBNull(rdr("CEP")), Nothing, rdr("CEP"))
                    lstExportar(i).Cliente.Endereco.UF = IIf(IsDBNull(rdr("UF")), Nothing, rdr("UF"))


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(IsDBNull(rdr("NumNFe")), Nothing, rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(IsDBNull(rdr("CodVerNFe")), Nothing, rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = IIf(IsDBNull(rdr("NFe")), Nothing, rdr("NFe"))
                    lstExportar(i).ContaReceber.NFMonit = IIf(IsDBNull(rdr("NFMonit")), Nothing, rdr("NFMonit"))

                    lstExportar(i).ContaReceber.CodIntClie = IIf(IsDBNull(rdr("CodIntClie")), Nothing, rdr("CodIntClie"))

                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()

        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA237.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA237.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaContaReceberGeraRemessa237(8)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try


        Return lstExportar

    End Function

    Public Function BuscaContaReceberGeraRemessa237BuDc(ByVal strCodBanco As String, ByVal strCodAgencia As String, ByVal strNumCta As String, ByVal CodIntClie As String, ByVal DtInicio As String, ByVal DtFim As String, ByVal TitInicio As String, ByVal TitFim As String, Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaContaReceberGeraRemessa237_BU_DC", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = IIf(IsDBNull(rdr("NumTit")), Nothing, rdr("NumTit"))
                    lstExportar(i).ContaReceber.SeqTit = IIf(IsDBNull(rdr("SeqTit")), Nothing, rdr("SeqTit"))
                    lstExportar(i).ContaReceber.DtEmissao = IIf(IsDBNull(rdr("DtEmissao")), Nothing, rdr("DtEmissao"))
                    lstExportar(i).ContaReceber.DtVcto = IIf(IsDBNull(rdr("DtVcto")), Nothing, rdr("DtVcto"))
                    lstExportar(i).ContaReceber.CodClie = IIf(IsDBNull(rdr("CodClie")), Nothing, rdr("CodClie"))
                    lstExportar(i).ContaReceber.CodInd = IIf(IsDBNull(rdr("CodInd")), Nothing, rdr("CodInd"))
                    lstExportar(i).ContaReceber.VlrInd = IIf(IsDBNull(rdr("VlrInd")), Nothing, rdr("VlrInd"))
                    lstExportar(i).ContaReceber.Situacao = IIf(IsDBNull(rdr("Situacao")), Nothing, rdr("Situacao"))
                    lstExportar(i).ContaReceber.CodBanco = IIf(IsDBNull(rdr("CodBanco")), Nothing, rdr("CodBanco"))
                    lstExportar(i).ContaReceber.CodAgen = IIf(IsDBNull(rdr("CodAgen")), Nothing, rdr("CodAgen"))
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), "", rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(IsDBNull(rdr("ObsTit")), Nothing, rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = IIf(IsDBNull(rdr("StatusBco")), Nothing, rdr("StatusBco"))
                    lstExportar(i).ContaReceber.TipoDup = IIf(IsDBNull(rdr("TipoDup")), Nothing, rdr("TipoDup"))
                    lstExportar(i).ContaReceber.MenBco = IIf(IsDBNull(rdr("MenBco")), Nothing, rdr("MenBco"))
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = IIf(IsDBNull(rdr("TipoPagto")), Nothing, rdr("TipoPagto"))
                    lstExportar(i).ContaReceber.VlrDesconto = IIf(IsDBNull(rdr("VlrDesconto")), Nothing, rdr("VlrDesconto"))
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = IIf(IsDBNull(rdr("RazaoSocial")), Nothing, rdr("RazaoSocial"))
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = IIf(IsDBNull(rdr("EnvioBanco")), Nothing, rdr("EnvioBanco"))
                    lstExportar(i).Cliente.Contabilidade.IsNNC = IIf(IsDBNull(rdr("IsNNC")), Nothing, rdr("IsNNC"))


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = IIf(IsDBNull(rdr("Endereco")), Nothing, rdr("Endereco"))
                    lstExportar(i).Cliente.Endereco.Bairro = IIf(IsDBNull(rdr("Bairro")), Nothing, rdr("Bairro"))
                    lstExportar(i).Cliente.Endereco.Cidade = IIf(IsDBNull(rdr("Cidade")), Nothing, rdr("Cidade"))
                    lstExportar(i).Cliente.Endereco.Cep = IIf(IsDBNull(rdr("CEP")), Nothing, rdr("CEP"))
                    lstExportar(i).Cliente.Endereco.UF = IIf(IsDBNull(rdr("UF")), Nothing, rdr("UF"))


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(IsDBNull(rdr("NumNFe")), Nothing, rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(IsDBNull(rdr("CodVerNFe")), Nothing, rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = IIf(IsDBNull(rdr("NFe")), Nothing, rdr("NFe"))
                    lstExportar(i).ContaReceber.NFMonit = IIf(IsDBNull(rdr("NFMonit")), Nothing, rdr("NFMonit"))

                    lstExportar(i).ContaReceber.CodIntClie = IIf(IsDBNull(rdr("CodIntClie")), Nothing, rdr("CodIntClie"))

                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()

        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA237.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA237.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function

    Public Function BuscaContaReceberGeraRemessa237(ByVal strCodBanco As String,
                                                    ByVal strCodAgencia As String,
                                                    ByVal strNumCta As String,
                                                    ByVal CodIntClie As String,
                                                    ByVal DtInicio As String,
                                                    ByVal DtFim As String,
                                                    ByVal TitInicio As String,
                                                    ByVal TitFim As String,
                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaContaReceberGeraRemessa237", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = IIf(IsDBNull(rdr("NumTit")), Nothing, rdr("NumTit"))
                    lstExportar(i).ContaReceber.SeqTit = IIf(IsDBNull(rdr("SeqTit")), Nothing, rdr("SeqTit"))
                    lstExportar(i).ContaReceber.DtEmissao = IIf(IsDBNull(rdr("DtEmissao")), Nothing, rdr("DtEmissao"))
                    lstExportar(i).ContaReceber.DtVcto = IIf(IsDBNull(rdr("DtVcto")), Nothing, rdr("DtVcto"))
                    lstExportar(i).ContaReceber.CodClie = IIf(IsDBNull(rdr("CodClie")), Nothing, rdr("CodClie"))
                    lstExportar(i).ContaReceber.CodInd = IIf(IsDBNull(rdr("CodInd")), Nothing, rdr("CodInd"))
                    lstExportar(i).ContaReceber.VlrInd = IIf(IsDBNull(rdr("VlrInd")), Nothing, rdr("VlrInd"))
                    lstExportar(i).ContaReceber.Situacao = IIf(IsDBNull(rdr("Situacao")), Nothing, rdr("Situacao"))
                    lstExportar(i).ContaReceber.CodBanco = IIf(IsDBNull(rdr("CodBanco")), Nothing, rdr("CodBanco"))
                    lstExportar(i).ContaReceber.CodAgen = IIf(IsDBNull(rdr("CodAgen")), Nothing, rdr("CodAgen"))
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), "", rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = IIf(IsDBNull(rdr("ObsTit")), Nothing, rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = IIf(IsDBNull(rdr("StatusBco")), Nothing, rdr("StatusBco"))
                    lstExportar(i).ContaReceber.TipoDup = IIf(IsDBNull(rdr("TipoDup")), Nothing, rdr("TipoDup"))
                    lstExportar(i).ContaReceber.MenBco = IIf(IsDBNull(rdr("MenBco")), Nothing, rdr("MenBco"))
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = IIf(IsDBNull(rdr("TipoPagto")), Nothing, rdr("TipoPagto"))
                    lstExportar(i).ContaReceber.VlrDesconto = IIf(IsDBNull(rdr("VlrDesconto")), Nothing, rdr("VlrDesconto"))
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = IIf(IsDBNull(rdr("RazaoSocial")), Nothing, rdr("RazaoSocial"))
                    'lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    'lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = IIf(IsDBNull(rdr("EnvioBanco")), Nothing, rdr("EnvioBanco"))
                    lstExportar(i).Cliente.Contabilidade.IsNNC = IIf(IsDBNull(rdr("IsNNC")), Nothing, rdr("IsNNC"))


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = IIf(IsDBNull(rdr("Endereco")), Nothing, rdr("Endereco"))
                    lstExportar(i).Cliente.Endereco.Bairro = IIf(IsDBNull(rdr("Bairro")), Nothing, rdr("Bairro"))
                    lstExportar(i).Cliente.Endereco.Cidade = IIf(IsDBNull(rdr("Cidade")), Nothing, rdr("Cidade"))
                    lstExportar(i).Cliente.Endereco.Cep = IIf(IsDBNull(rdr("CEP")), Nothing, rdr("CEP"))
                    lstExportar(i).Cliente.Endereco.UF = IIf(IsDBNull(rdr("UF")), Nothing, rdr("UF"))


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(IsDBNull(rdr("NumNFe")), Nothing, rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(IsDBNull(rdr("CodVerNFe")), Nothing, rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = IIf(IsDBNull(rdr("NFe")), Nothing, rdr("NFe"))
                    lstExportar(i).ContaReceber.NFMonit = IIf(IsDBNull(rdr("NFMonit")), Nothing, rdr("NFMonit"))

                    lstExportar(i).ContaReceber.CodIntClie = IIf(IsDBNull(rdr("CodIntClie")), Nothing, rdr("CodIntClie"))

                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()

        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA237.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA237.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function

    Public Function BuscaContaReceberGeraRemessa341_033_399_356(ByVal strCodBanco As String,
                                                           ByVal strCodAgencia As String,
                                                           ByVal strNumCta As String,
                                                           ByVal strTipoPagto As String,
                                                           ByVal CodIntClie As String,
                                                           ByVal DtInicio As String,
                                                           ByVal DtFim As String,
                                                           ByVal TitInicio As String,
                                                           ByVal TitFim As String,
                                                           Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaContaReceberGeraRemessa341", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstExportar.Add(New Exportacao)
                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    If strTipoPagto = "BO" Then
                        lstExportar(i).ContaReceber.DtVctoDMaisUm = rdr("DtVctoDMaisUm")
                    End If
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))

                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")

                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")

                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")

                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))
                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")
                    lstExportar(i).ContaReceber.CodOperacaoConta = rdr("CodOperacaoConta")

                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)
                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If

            rdr.Close()
        Catch ex As Exception
            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function



    Public Function BuscaContaReceberGeraRemessaRT(ByVal strCodBanco As String,
                                                           ByVal strCodAgencia As String,
                                                           ByVal strNumCta As String,
                                                           ByVal strTipoPagto As String,
                                                           ByVal CodIntClie As String,
                                                           ByVal DtInicio As String,
                                                           ByVal DtFim As String,
                                                           ByVal TitInicio As String,
                                                           ByVal TitFim As String,
                                                           Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaContaReceberGeraRemessaRT", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstExportar.Add(New Exportacao)
                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    If strTipoPagto = "BO" Then
                        lstExportar(i).ContaReceber.DtVctoDMaisUm = rdr("DtVctoDMaisUm")
                    End If
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))

                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")

                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")

                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")

                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))
                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")
                    lstExportar(i).ContaReceber.CodOperacaoConta = rdr("CodOperacaoConta")

                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)
                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If

            rdr.Close()
        Catch ex As Exception
            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function




    Public Function BuscaContaReceberGeraRemessa341BoletoUnificado(ByVal strCodBanco As String,
                                                                   ByVal strCodAgencia As String,
                                                                   ByVal strNumCta As String,
                                                                   ByVal strTipoPagto As String,
                                                                   ByVal CodIntClie As String,
                                                                   ByVal DtInicio As String,
                                                                   ByVal DtFim As String,
                                                                   ByVal TitInicio As String,
                                                                   ByVal TitFim As String) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaContaReceberGeraRemessa341BoletoUnificado", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstExportar.Add(New Exportacao)
                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))

                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")

                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")

                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")

                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))
                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")
                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstExportar.Add(New Exportacao)
                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception
            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA341BOLETOUNIFICADO.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA341BOLETOUNIFICADO.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function

    Public Function BuscaContaReceberGeraRemessaBoletoUnificado033(ByVal strCodBanco As String,
                                                           ByVal strCodAgencia As String,
                                                           ByVal strNumCta As String,
                                                           ByVal strTipoPagto As String,
                                                           ByVal CodIntClie As String,
                                                           ByVal DtInicio As String,
                                                           ByVal DtFim As String,
                                                           ByVal TitInicio As String,
                                                           ByVal TitFim As String) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaContaReceberGeraRemessa341_BoletoUnificado", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))


            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")

                    lstExportar(i).ContaReceber.CodOperacaoConta = rdr("CodOperacaoConta")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function

    Public Function BuscaContaReceberGeraRemessa041(ByVal strCodBanco As String,
                                                            ByVal strCodAgencia As String,
                                                            ByVal strNumCta As String,
                                                            ByVal strTipoPagto As String,
                                                            ByVal CodIntClie As String,
                                                            ByVal DtInicio As String,
                                                            ByVal DtFim As String,
                                                            ByVal TitInicio As String,
                                                            ByVal TitFim As String,
                                                            Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaContaReceberGeraRemessa041", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()
                    lstExportar.Add(New Exportacao)
                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.ExcecaoEnvioNumNfe = IIf(IsDBNull(rdr("ExcecaoEnvioNumNfe")), "Não", rdr("ExcecaoEnvioNumNfe"))
                    lstExportar(i).ContaReceber.NossoNumeroBco = IIf(IsDBNull(rdr("NossoNumeroBco")), Nothing, rdr("NossoNumeroBco"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))

                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")

                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))
                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")
                    lstExportar(i).ContaReceber.CodOperacaoConta = rdr("CodOperacaoConta")
                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)
                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA341.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function

    Public Function BuscaContaReceberGeraRemessa001(ByVal strCodBanco As String,
                                                   ByVal strCodAgencia As String,
                                                   ByVal strNumCta As String,
                                                   ByVal strTipoPagto As String,
                                                   ByVal CodIntClie As String,
                                                   ByVal DtInicio As String,
                                                   ByVal DtFim As String,
                                                   ByVal TitInicio As String,
                                                   ByVal TitFim As String,
                                                   Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaContaReceberGeraRemessa001", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))
            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.DtVcto = rdr("DtVcto")
                    lstExportar(i).ContaReceber.CodClie = rdr("CodClie")
                    lstExportar(i).ContaReceber.CodInd = rdr("CodInd")
                    lstExportar(i).ContaReceber.VlrInd = rdr("VlrInd")
                    lstExportar(i).ContaReceber.Situacao = rdr("Situacao")
                    lstExportar(i).ContaReceber.CodBanco = rdr("CodBanco")
                    lstExportar(i).ContaReceber.CodAgen = rdr("CodAgen")
                    lstExportar(i).ContaReceber.NumPortador = IIf(IsDBNull(rdr("NumPortador")), Nothing, rdr("NumPortador"))
                    lstExportar(i).ContaReceber.ObsTit = String.IsNullOrEmpty(rdr("ObsTit"))
                    lstExportar(i).ContaReceber.StatusBco = rdr("StatusBco")
                    lstExportar(i).ContaReceber.TipoDup = rdr("TipoDup")
                    lstExportar(i).ContaReceber.MenBco = rdr("MenBco")
                    lstExportar(i).ContaReceber.CodAgenDeb = IIf(IsDBNull(rdr("CodAgenDeb")), Nothing, rdr("CodAgenDeb"))
                    lstExportar(i).ContaReceber.NumCtaDeb = IIf(IsDBNull(rdr("NumCtaDeb")), Nothing, rdr("NumCtaDeb"))
                    lstExportar(i).ContaReceber.TipoPagto = rdr("TipoPagto")
                    lstExportar(i).ContaReceber.VlrDesconto = rdr("VlrDesconto")
                    lstExportar(i).ContaReceber.DtLimiteDesconto = IIf(IsDBNull(rdr("DtLimiteDesconto")), Nothing, rdr("DtLimiteDesconto"))
                    lstExportar(i).ContaReceber.IdFilial = IIf(IsDBNull(rdr("IdFilial")), Nothing, rdr("IdFilial"))


                    lstExportar(i).Cliente = New Cliente
                    lstExportar(i).Cliente.RazaoSocial = rdr("RazaoSocial")
                    lstExportar(i).Cliente.CGC_CPF = rdr("CGC_CPF")
                    lstExportar(i).Cliente.FisiJuri = rdr("FisiJuri")
                    lstExportar(i).Cliente.Contabilidade = New ClienteContabilidade
                    lstExportar(i).Cliente.Contabilidade.EnvioBanco = rdr("EnvioBanco")
                    lstExportar(i).Cliente.Contabilidade.IsNNC = rdr("IsNNC")
                    lstExportar(i).Cliente.CodIntClie = rdr("CodIntClie")


                    lstExportar(i).Cliente.Endereco = New Endereco
                    lstExportar(i).Cliente.Endereco.Endereco = rdr("Endereco")
                    lstExportar(i).Cliente.Endereco.Bairro = rdr("Bairro")
                    lstExportar(i).Cliente.Endereco.Cidade = rdr("Cidade")
                    lstExportar(i).Cliente.Endereco.Cep = rdr("CEP")
                    lstExportar(i).Cliente.Endereco.UF = rdr("UF")


                    lstExportar(i).NotaFiscal = New NotaFiscal
                    lstExportar(i).NotaFiscal.NumNFe = IIf(String.IsNullOrEmpty(rdr("NumNFe")), "", rdr("NumNFe"))
                    lstExportar(i).NotaFiscal.CodVerNFe = IIf(String.IsNullOrEmpty(rdr("CodVerNFe")), "", rdr("CodVerNFe"))


                    lstExportar(i).ContaReceber.NFe = rdr("NFe")
                    lstExportar(i).ContaReceber.NFMonit = rdr("NFMonit")


                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA001.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_BUSCACONTARECEBERGERAREMESSA001.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function

    Public Function BuscaContasReceberAdyenAgendamento(ByVal CodIntClie As String,
                                                                    ByVal DtInicio As String,
                                                                    ByVal DtFim As String,
                                                                    ByVal TitInicio As String,
                                                                    ByVal TitFim As String,
                                                                    ByVal usr As String,
                                                                    ByVal dtAgendamento As DateTime,
                                                                    ByVal hrAgendamento As String,
                                                                    Optional ByVal IsMultaFormaPagtoCliente As Integer = 0) As List(Of Exportacao)
        Dim rdr As SqlDataReader
        Dim lstExportar = New List(Of Exportacao)
        Dim i As Integer = 0
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_GeraAgendamento_CartaoCredito_Adyen", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            command.Parameters.Add(New SqlParameter("@DtEmissaoInicio", DtInicio))
            command.Parameters.Add(New SqlParameter("@DtEmissaoFim", DtFim))
            command.Parameters.Add(New SqlParameter("@TitInicio", TitInicio))
            command.Parameters.Add(New SqlParameter("@TitFim", TitFim))
            command.Parameters.Add(New SqlParameter("@Usr", usr))
            command.Parameters.Add(New SqlParameter("@DtAgendamento", dtAgendamento))
            command.Parameters.Add(New SqlParameter("@HrAgendamento", hrAgendamento))
            command.Parameters.Add(New SqlParameter("@IsMultaFormaPagtoCliente", IsMultaFormaPagtoCliente))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then

                Do While rdr.Read()

                    lstExportar.Add(New Exportacao)

                    lstExportar(i).ContaReceber = New ContaReceber
                    lstExportar(i).ContaReceber.NumTit = rdr("NumTit")
                    lstExportar(i).ContaReceber.SeqTit = rdr("SeqTit")
                    lstExportar(i).ContaReceber.DtEmissao = rdr("DtEmissao")
                    lstExportar(i).ContaReceber.idAgendamento = rdr("IdAgendamento")
                    lstExportar(i).Sucesso = True
                    lstExportar(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop

            Else
                lstExportar.Add(New Exportacao)

                lstExportar(0).Sucesso = False
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.None
                lstExportar(0).MsgErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Descricao
                lstExportar(0).NumErro = ErrorConstants.NENHUMA_NOTA_FISCAL_ENCONTRADA.Id
                lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
                lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
            rdr.Close()
        Catch ex As Exception

            lstExportar.Add(New Exportacao)
            lstExportar(0).Sucesso = False
            lstExportar(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Descricao & ex.Message
            lstExportar(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTARECEBERGERAREMESSA_CARTAOCREDITO_VISA.Id
            lstExportar(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstExportar(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstExportar(0).NumErro, lstExportar(0).MsgErro, lstExportar(0).TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstExportar
    End Function


    Public Function BuscaTipoDup(ByVal sNumtit As String, ByVal sSeqTit As String)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaReceber As New List(Of ContaReceber)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarTipoDup", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", sNumtit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", sSeqTit))
            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstContaReceber.Add(New ContaReceber)
                    'lstContaReceber(i).TipoDup = rdr("TipoDup")

                    If Not rdr.IsDBNull(rdr.GetOrdinal("TipoDup")) Then
                        lstContaReceber(i).TipoDup = rdr("TipoDup")
                    Else
                        lstContaReceber(i).TipoDup = String.Empty ' Atribui um valor vazio se for nulo
                    End If

                    If Not rdr.IsDBNull(rdr.GetOrdinal("Situacao")) Then
                        lstContaReceber(i).Situacao = rdr("Situacao")
                    Else
                        lstContaReceber(i).Situacao = String.Empty ' Atribui um valor vazio se for nulo
                    End If


                    lstContaReceber(i).Sucesso = True
                    lstContaReceber(i).TipoErro = DadosGenericos.TipoErro.None
                    i += 1
                Loop
            Else
                lstContaReceber.Add(New ContaReceber)
                lstContaReceber(0).Sucesso = False
                lstContaReceber(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaReceber(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaReceber.Add(New ContaReceber)
            lstContaReceber(0).Sucesso = False
            lstContaReceber(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCADTEMISSAOTITDESCDTVCTOSEQTITNUMTITCONTASRECEBER.Descricao & ex.Message
            lstContaReceber(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCADTEMISSAOTITDESCDTVCTOSEQTITNUMTITCONTASRECEBER.Id
            lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaReceber(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaReceber(0).NumErro, lstContaReceber(0).MsgErro, lstContaReceber(0).TipoErro, "Projeto: ContasAReceberBC - Classe: ConsultaContasReceber - Função: BuscaDtEmissaoTitDescDtVctoSeqTitNumTitContasReceber(19)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return lstContaReceber

    End Function

    Public Function BuscarMultasAbertasPorCodIntClie(ByVal codIntClie As String) As DataTable
        Dim dt As New DataTable()
        Dim _Retorno As New Retorno

        Using connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            Try

                connection.Open()

                Dim Command As SqlCommand = New SqlCommand("P_BuscarMultasAbertasPorCodIntClie", connection)
                Command.Parameters.Add(New SqlParameter("@CodIntClie", codIntClie))
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.ContImprodutivos

                Dim adapter As New SqlDataAdapter(Command)
                adapter.Fill(dt)
            Catch ex As Exception
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
            End Try
        End Using
        Return dt
    End Function
End Class
