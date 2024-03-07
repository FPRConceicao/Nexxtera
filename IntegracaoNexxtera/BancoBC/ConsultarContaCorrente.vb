Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.Common

Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Data.Common


Public Class ConsultarContaCorrente

    Public Function BuscaNumCtaVincDescContaCorrente(ByVal strCodBanco As String,
                                                     ByVal strCodAgen As String,
                                                     ByVal strNumCta As String) As ContaCorrente

        'TODO VERIFICAR ONDE É UTILIZADO ESTA FUNCTION
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaNumCtaVincDescContaCorrente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                contaReceber = New ContaCorrente

                contaReceber.NumCtaVincDesc = rdr("NumCtaVincDesc").ToString

                contaReceber.Sucesso = True
                contaReceber.TipoErro = DadosGenericos.TipoErro.None
            Else
                contaReceber = New ContaCorrente
                contaReceber.Sucesso = False
                contaReceber.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                contaReceber.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                contaReceber.TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            contaReceber = New ContaCorrente
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_CONSULTACONTASRECEBER.Descricao & ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_CONSULTACONTASRECEBER.Id
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: ConsultaContasReceber(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return contaReceber

    End Function


    Public Function BuscaIsDebAutoContaCorrentePorBancoAgnConta(ByVal strCodBanco As String,
                                                                ByVal strCodAgen As String,
                                                                ByVal strNumCta As String) As ContaCorrente

        'TODO VERIFICAR ONDE É UTILIZADO ESTA FUNCTION
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim contaReceber = New ContaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaIsDebAutoContaCorrentePorBancoAgnConta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                contaReceber = New ContaCorrente

                contaReceber.IsDebAuto = rdr("IsDebAuto").ToString
                contaReceber.CodBanco = rdr("CodBanco").ToString
                contaReceber.CodAgen = rdr("CodAgen").ToString
                contaReceber.NumCta = rdr("NumCta").ToString

                contaReceber.Sucesso = True
                contaReceber.TipoErro = DadosGenericos.TipoErro.None
            Else
                contaReceber = New ContaCorrente
                contaReceber.Sucesso = False
                contaReceber.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                contaReceber.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                contaReceber.TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            contaReceber = New ContaCorrente
            contaReceber.Sucesso = False
            contaReceber.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAISDEBAUTOCONTACORRENTEPORBANCOAGNCONTA.Descricao & ex.Message
            contaReceber.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAISDEBAUTOCONTACORRENTEPORBANCOAGNCONTA.Id
            contaReceber.TipoErro = DadosGenericos.TipoErro.Arquitetura
            contaReceber.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(contaReceber.NumErro, contaReceber.MsgErro, contaReceber.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaIsDebAutoContaCorrentePorBancoAgnConta(2)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return contaReceber

    End Function

    Public Function BuscaTodasContasCorrentesGrid(Optional ByVal strCodBanco As String = "") As List(Of ContaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of ContaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaTodasContasCorrentesGrid", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", strCodBanco))



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New ContaCorrente)
                    lstContaCorrente(i).CodBanco = rdr("CodBanco").ToString
                    lstContaCorrente(i).NomeBanco = rdr("NomeBanco").ToString
                    lstContaCorrente(i).CodAgen = rdr("CodAgen").ToString
                    lstContaCorrente(i).NomeAgen = rdr("NomeAgen").ToString
                    lstContaCorrente(i).NumCta = rdr("NumCta").ToString
                    lstContaCorrente(i).Descricao = rdr("Descricao").ToString
                    lstContaCorrente(i).ContaFinan = rdr("ContaFinan").ToString
                    lstContaCorrente(i).LimiteCredito = rdr("LimiteCredito").ToString
                    lstContaCorrente(i).Fluxo = rdr("Fluxo").ToString
                    lstContaCorrente(i).Tipo = rdr("Tipo").ToString
                    lstContaCorrente(i).LayoutCheque = rdr("LayoutCheque").ToString
                    lstContaCorrente(i).TextoBordero = rdr("TextoBordero").ToString
                    lstContaCorrente(i).CodCCusto = rdr("CodCCusto").ToString
                    lstContaCorrente(i).ContaCtbl = rdr("ContaCtbl").ToString
                    lstContaCorrente(i).IsDebAuto = rdr("IsDebAuto").ToString
                    lstContaCorrente(i).FloatCredito = rdr("FloatCredito").ToString
                    lstContaCorrente(i).SeqRemes = rdr("SeqRemes").ToString
                    lstContaCorrente(i).TipoConta = rdr("TipoConta").ToString
                    lstContaCorrente(i).NumCtaVinc = rdr("NumCtaVinc").ToString
                    lstContaCorrente(i).Status = rdr("Status").ToString
                    lstContaCorrente(i).Convenio = rdr("Convenio").ToString
                    lstContaCorrente(i).CodCarteira = rdr("CodCarteira").ToString
                    lstContaCorrente(i).NumCtaVincDesc = rdr("NumCtaVincDesc").ToString
                    lstContaCorrente(i).PermiteDI = rdr("PermiteDI").ToString
                    lstContaCorrente(i).PagamentoEletronico = rdr("PagamentoEletronico").ToString
                    lstContaCorrente(i).CodCCustoDNI = rdr("CodCCustoDNI").ToString
                    lstContaCorrente(i).ContaPadraoDebAuto = rdr("ContaPadraoDebAuto")
                    lstContaCorrente(i).ContaPadraoDebAutoDesc = rdr("ContaPadraoDebAutoDesc").ToString

                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New ContaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstContaCorrente.Add(New ContaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCATODASCONTASCORRENTESGRID.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCATODASCONTASCORRENTESGRID.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaTodasContasCorrentesGrid(3)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscaContaContabilPorCodBancoCodAgenDifNumCta(ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String) As List(Of ContaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of ContaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaContabilPorCodBancoCodAgenDifNumCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", strNumCta))



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New ContaCorrente)
                    lstContaCorrente(i).CodBanco = rdr("CodBanco").ToString
                    'lstContaCorrente(i).NomeBanco = rdr("NomeBanco").ToString
                    lstContaCorrente(i).CodAgen = rdr("CodAgen").ToString
                    'lstContaCorrente(i).NomeAgen = rdr("NomeAgen").ToString
                    lstContaCorrente(i).NumCta = rdr("NumCta").ToString
                    lstContaCorrente(i).Descricao = rdr("Descricao").ToString
                    'lstContaCorrente(i).ContaFinan = rdr("ContaFinan").ToString
                    'lstContaCorrente(i).LimiteCredito = rdr("LimiteCredito").ToString
                    'lstContaCorrente(i).Fluxo = rdr("Fluxo").ToString
                    'lstContaCorrente(i).Tipo = rdr("Tipo").ToString
                    'lstContaCorrente(i).LayoutCheque = rdr("LayoutCheque").ToString
                    'lstContaCorrente(i).TextoBordero = rdr("TextoBordero").ToString
                    'lstContaCorrente(i).CodCCusto = rdr("CodCCusto").ToString
                    'lstContaCorrente(i).ContaCtbl = rdr("ContaCtbl").ToString
                    'lstContaCorrente(i).IsDebAuto = rdr("IsDebAuto").ToString
                    'lstContaCorrente(i).FloatCredito = rdr("FloatCredito").ToString
                    'lstContaCorrente(i).SeqRemes = rdr("SeqRemes").ToString
                    'lstContaCorrente(i).TipoConta = rdr("TipoConta").ToString
                    'lstContaCorrente(i).NumCtaVinc = rdr("NumCtaVinc").ToString
                    'lstContaCorrente(i).Status = rdr("Status").ToString
                    'lstContaCorrente(i).Convenio = rdr("Convenio").ToString
                    'lstContaCorrente(i).CodCarteira = rdr("CodCarteira").ToString
                    'lstContaCorrente(i).NumCtaVincDesc = rdr("NumCtaVincDesc").ToString
                    'lstContaCorrente(i).PermiteDI = rdr("PermiteDI").ToString
                    'lstContaCorrente(i).PagamentoEletronico = rdr("PagamentoEletronico").ToString

                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New ContaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstContaCorrente.Add(New ContaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTACONTABILPORCODBANCOCODAGENDIFNUMCTA.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTACONTABILPORCODBANCOCODAGENDIFNUMCTA.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaContaContabilPorCodBancoCodAgenDifNumCta(4)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscaMovimentacaoContaCorrente() As List(Of ContaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of ContaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaMovimentacaoContaCorrente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New ContaCorrente)
                    lstContaCorrente(i).BancoAgen = rdr("BancoAgen")
                    lstContaCorrente(i).NumCta = rdr("NumCta")
                    lstContaCorrente(i).SaldoTotal = rdr("SaldoTotal")
                    lstContaCorrente(i).LimiteCredito = rdr("LimiteCred")
                    lstContaCorrente(i).SaldoLiq = rdr("SaldoLiq")
                    lstContaCorrente(i).SaldoRes = rdr("SaldoRes")
                    lstContaCorrente(i).NomeAgen = rdr("NomeAgen").ToString
                    lstContaCorrente(i).Descricao = rdr("Descricao").ToString
                    lstContaCorrente(i).CodCCusto = rdr("CodCCusto")



                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New ContaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstContaCorrente.Add(New ContaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAMOVIMENTACAOCONTACORRENTE.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAMOVIMENTACAOCONTACORRENTE.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaMovimentacaoContaCorrente(5)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    'Public Function BuscaLancamentosContaCorrenteGrid(ByVal _Lcto As LctoCtaCorrente, ByVal dtDe As DateTime, ByVal dtAte As DateTime) As List(Of LctoCtaCorrente)


    '    Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '    Dim rdr As SqlDataReader
    '    Dim lstContaCorrente As New List(Of LctoCtaCorrente)
    '    Dim i As Integer = 0

    '    Try
    '        ''Informa a procedure
    '        Dim Command As SqlCommand = New SqlCommand("P_BuscaLancamentosContaCorrenteGrid", connection)
    '        Command.CommandType = CommandType.StoredProcedure
    '        Command.CommandTimeout = DadosGenericos.Timeout.Query
    '        Command.Parameters.Add(New SqlParameter("@CODBANCO", _Lcto.CodBanco))
    '        Command.Parameters.Add(New SqlParameter("@CODAGEN", _Lcto.CodAgen))
    '        Command.Parameters.Add(New SqlParameter("@NUMCTA", _Lcto.NumCta))
    '        Command.Parameters.Add(New SqlParameter("@CONCILIADO", _Lcto.Conciliado))
    '        Command.Parameters.Add(New SqlParameter("@DTDE", IIf(IsNothing(Convert.ToDateTime(dtDe, Funcoes.Cultura)) Or Convert.ToDateTime(dtDe, Funcoes.Cultura).Equals(Date.MinValue), DBNull.Value, Convert.ToDateTime(dtDe, Funcoes.Cultura))))
    '        Command.Parameters.Add(New SqlParameter("@DTATE", IIf(IsNothing(Convert.ToDateTime(dtAte, Funcoes.Cultura)) Or Convert.ToDateTime(dtAte, Funcoes.Cultura).Equals(Date.MinValue), DBNull.Value, Convert.ToDateTime(dtAte, Funcoes.Cultura))))

    '        ''Abre a conexao
    '        connection.Open()
    '        ''Executa a procedure
    '        rdr = Command.ExecuteReader()

    '        If rdr.HasRows Then

    '            Do While (rdr.Read)

    '                lstContaCorrente.Add(New LctoCtaCorrente)
    '                lstContaCorrente(i).DtLcto = Convert.ToDateTime(rdr("Data"), Funcoes.Cultura)
    '                lstContaCorrente(i).NatLcto = rdr("Nat")
    '                lstContaCorrente(i).DctoLcto = rdr("DctoLcto")
    '                lstContaCorrente(i).VlrLcto = CDbl(rdr("Vlr"))
    '                lstContaCorrente(i).NumLcto = rdr("NumLcto")
    '                lstContaCorrente(i).SitLcto = rdr("Situacao")
    '                lstContaCorrente(i).TipoLcto = rdr("Tipo")
    '                lstContaCorrente(i).Conciliado = rdr("Conciliado")
    '                lstContaCorrente(i).NumLote = IIf(IsDBNull(rdr("NumLote")), Nothing, rdr("NumLote"))
    '                lstContaCorrente(i).MesAnoLote = IIf(IsDBNull(rdr("MesAnoLote")), Nothing, rdr("MesAnoLote"))
    '                lstContaCorrente(i).CodEmpCtbl = IIf(IsDBNull(rdr("CodEmpCtbl")), Nothing, rdr("CodEmpCtbl"))


    '                lstContaCorrente(i).Sucesso = True
    '                lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
    '                i = i + 1

    '            Loop
    '        Else
    '            lstContaCorrente.Add(New LctoCtaCorrente)
    '            lstContaCorrente(0).Sucesso = False
    '            lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
    '            lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
    '            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
    '        End If
    '    Catch ex As Exception

    '        lstContaCorrente.Add(New LctoCtaCorrente)
    '        lstContaCorrente(0).Sucesso = False
    '        lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Descricao & ex.Message
    '        lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Id
    '        lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaLancamentosContaCorrenteGrid(6)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    Finally
    '        connection.Close()
    '    End Try


    '    Return lstContaCorrente

    'End Function

    Public Function BuscaDetLancamentosContaCorrentePorNumLcto(ByVal strNumLcto As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As List(Of DetLctoCtaCorrente)


        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of DetLctoCtaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaDetLancamentosContaCorrentePorNumLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumLcto))
            Command.Transaction = trans


            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New DetLctoCtaCorrente)
                    'lstContaCorrente(i).CodBanco = Convert.ToDateTime(rdr("CodBanco"))
                    lstContaCorrente(i).NumTit = rdr("NumTit")
					lstContaCorrente(i).DtEmissao = Date.Parse(rdr("DtEmissao"))
					lstContaCorrente(i).SeqTit = rdr("SeqTit")
					lstContaCorrente(i).NumLcto = rdr("NumLcto")

					If (IsDBNull(rdr("DtPgto"))) Then
						lstContaCorrente(i).DtPgto = Nothing
					Else
						lstContaCorrente(i).DtPgto = Date.Parse(rdr("DtPgto"))
					End If

					lstContaCorrente(i).CodClieForn = rdr("CodClieForn")
					lstContaCorrente(i).OrigTit = rdr("OrigTit")

					lstContaCorrente(i).Sucesso = True
					lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
					i = i + 1
				Loop
            Else
                lstContaCorrente.Add(New DetLctoCtaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception

            lstContaCorrente.Add(New DetLctoCtaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaDetLancamentosContaCorrentePorNumLcto(7)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            'connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscaDetLctoCtaCorrentePorNumTitSeqtitNumLcto(ByVal strNumTit As String, ByVal strSeqTit As String, ByVal iNumLcto As Integer) As List(Of DetLctoCtaCorrente)

        Dim con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstDetLcto As New List(Of DetLctoCtaCorrente)

        Try
            con.Open()

            Dim cmd As New SqlCommand("P_BuscaDetLctoCtaCorrente", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            cmd.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            cmd.Parameters.Add(New SqlParameter("@INumLcto", iNumLcto))

            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            If (rdr.HasRows) Then
                While (rdr.Read())
                    Dim _DetLcto As New DetLctoCtaCorrente
                    With _DetLcto
                        .NumLcto = Funcoes.VerificaDbNull(rdr("NumLcto"))
                        .OrigTit = Funcoes.VerificaDbNull(rdr("OrigTit"))
                        .NumTit = Funcoes.VerificaDbNull(rdr("NumTit"))
                        .SeqTit = Funcoes.VerificaDbNull(rdr("SeqTit"))
                        .CodClieForn = Funcoes.VerificaDbNull(rdr("CodClieForn"))
                        .DtEmissao = Funcoes.VerificaDbNull(rdr("DtEmissao"))
                        .VlrPago = Funcoes.VerificaDbNull(rdr("VlrPago"))
                        .DescrLcto = Funcoes.VerificaDbNull(rdr("DescrLcto"))
                        .VlrMulta = Funcoes.VerificaDbNull(rdr("VlrMulta"))
                        .VlrJuros = Funcoes.VerificaDbNull(rdr("VlrJuros"))
                        .VlrDesc = Funcoes.VerificaDbNull(rdr("VlrDesc"))
                        .VlrVarCamb = Funcoes.VerificaDbNull(rdr("VlrVarCamb"))
                        .VlrAbat = Funcoes.VerificaDbNull(rdr("VlrAbat"))
                        .VlrDevol = Funcoes.VerificaDbNull(rdr("VlrDevol"))
                        .ObsBaixa = Funcoes.VerificaDbNull(rdr("ObsBaixa"))
                        .DtUltAlt = Funcoes.VerificaDbNull(rdr("DtUltAlt"))
                        .UsrUltAlt = Funcoes.VerificaDbNull(rdr("UsrUltAlt"))
                        .DtPgto = Funcoes.VerificaDbNull(rdr("DtPgto"))
                        .NumLoteCtbl = Funcoes.VerificaDbNull(rdr("NumLoteCtbl"))
                        .MesAnoLoteCtbl = Funcoes.VerificaDbNull(rdr("MesAnoLoteCtbl"))
                        .NumLctoCtbl = Funcoes.VerificaDbNull(rdr("NumLctoCtbl"))
                        .SeqLctoLoteCtbl = Funcoes.VerificaDbNull(rdr("SeqLctoLoteCtbl"))
                        .CodEvento = Funcoes.VerificaDbNull(rdr("CodEvento"))
                        .VlrInd = Funcoes.VerificaDbNull(rdr("VlrInd"))
                        .CodBanco = Funcoes.VerificaDbNull(rdr("CodBanco"))
                        .CodEmpCtbl = Funcoes.VerificaDbNull(rdr("CodEmpCtbl"))
                        .CodCCusto = Funcoes.VerificaDbNull(rdr("CodCCusto"))
                        .Origem = Funcoes.VerificaDbNull(rdr("Origem"))
                        .CodHist = Funcoes.VerificaDbNull(rdr("CodHist"))
                        .Complemento = Funcoes.VerificaDbNull(rdr("Complemento"))
                        .IsCanc = Funcoes.VerificaDbNull(rdr("IsCanc"))
                    End With

                    lstDetLcto.Add(_DetLcto)
                End While
            Else
                Dim _Retorno As New DetLctoCtaCorrente()
                With _Retorno
                    .Sucesso = False
                    .NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                    .MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                    .TipoErro = DadosGenericos.TipoErro.Arquitetura
                    .ImagemErro = DadosGenericos.ImagemRetorno.Erro
                End With

                lstDetLcto.Add(_Retorno)
            End If
        Catch ex As Exception

            Dim _Retorno As New DetLctoCtaCorrente()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCADETLCTOCTACORRENTEPORNUMTITSEQTITNUMLCTO.Id
                .MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCADETLCTOCTACORRENTEPORNUMTITSEQTITNUMLCTO.Descricao
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            lstDetLcto.Clear()
            lstDetLcto.Add(_Retorno)

        Finally
            con.Close()
        End Try


        Return lstDetLcto
    End Function

    Public Function BuscarCountBaixaContaReceberPorNumTitSeqTitDtEmissaoDtPgto(ByVal _DetLctoCC As DetLctoCtaCorrente, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As BaixaContaReceber


        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente As New BaixaContaReceber


        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarCountBaixaContaReceberPorNumTitSeqTitDtEmissaoDtPgto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMTIT", _DetLctoCC.NumTit))
            Command.Parameters.Add(New SqlParameter("@SEQTIT", _DetLctoCC.SeqTit))
            Command.Parameters.Add(New SqlParameter("@DTEMISSAO", Convert.ToDateTime(_DetLctoCC.DtEmissao)))
            Command.Parameters.Add(New SqlParameter("@DTPGTO", Convert.ToDateTime(_DetLctoCC.DtPgto)))
            Command.Transaction = trans


            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                'lstContaCorrente.Add(New BaixaContaReceber)
                _ContaCorrente.Total = rdr("Total")



                _ContaCorrente.Sucesso = True
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
            Else
                'lstContaCorrente.Add(New BaixaContaReceber)
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception

            'lstContaCorrente.Add(New BaixaContaReceber)
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARCOUNTBAIXACONTARECEBERPORNUMTITSEQTITDTEMISSAODTPGTO.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARCOUNTBAIXACONTARECEBERPORNUMTITSEQTITDTEMISSAODTPGTO.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarCountBaixaContaReceberPorNumTitSeqTitDtEmissaoDtPgto(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            ' connection.Close()

        End Try


        Return _ContaCorrente

    End Function

    'Public Function BuscarItensLctoMovimentacaoCtaCorrentePorNumLcto(ByVal strNumLcto As String) As List(Of DetLctoCtaCorrente)


    '    Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '    Dim rdr As SqlDataReader
    '    Dim lstContaCorrente As New List(Of DetLctoCtaCorrente)
    '    Dim i As Integer = 0

    '    Try
    '        ''Informa a procedure
    '        Dim Command As SqlCommand = New SqlCommand("P_BuscarItensLctoMovimentacaoCtaCorrentePorNumLcto", connection)
    '        Command.CommandType = CommandType.StoredProcedure
    '        Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
    '        Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumLcto))



    '        ''Abre a conexao
    '        connection.Open()
    '        ''Executa a procedure
    '        rdr = Command.ExecuteReader()

    '        If rdr.HasRows Then

    '            Do While (rdr.Read)

    '                lstContaCorrente.Add(New DetLctoCtaCorrente)
    '                lstContaCorrente(i).NumLcto = rdr("NumLcto").ToString
    '                lstContaCorrente(i).OrigTit = rdr("OrigTit").ToString
    '                lstContaCorrente(i).NumTit = rdr("NumTit").ToString
    '                lstContaCorrente(i).SeqTit = rdr("SeqTit").ToString
    '                lstContaCorrente(i).ClieForn = rdr("ClieForn").ToString
    '                ' lstContaCorrente(i).DtEmi = IIf(IsDBNull(rdr("DtEmi")), Nothing, Convert.ToDateTime(rdr("DtEmi")))
    '                If IsDBNull(rdr("DtEmi")) Then
    '                    lstContaCorrente(i).DtEmi = Nothing
    '                Else
    '		lstContaCorrente(i).DtEmi = Convert.ToDateTime(rdr("DtEmi"), Funcoes.Cultura)
    '                End If
    '                lstContaCorrente(i).VlrTit = IIf(IsDBNull(rdr("VlrTit")), Nothing, rdr("VlrTit"))
    '                If IsDBNull(rdr("DtVcto")) Then
    '                    lstContaCorrente(i).DtVcto = Nothing
    '                Else
    '		lstContaCorrente(i).DtVcto = Convert.ToDateTime(rdr("DtVcto"), Funcoes.Cultura)
    '                End If
    '                If IsDBNull(rdr("DtPago")) Then
    '                    lstContaCorrente(i).DtPago = Nothing
    '                Else
    '		lstContaCorrente(i).DtPago = Convert.ToDateTime(rdr("DtPago"), Funcoes.Cultura)
    '	End If


    '	If IsDBNull(rdr("DtPgto")) Then
    '		lstContaCorrente(i).DtPgto = Nothing
    '	Else
    '		lstContaCorrente(i).DtPgto = Convert.ToDateTime(rdr("DtPgto"), Funcoes.Cultura)
    '	End If

    '                'lstContaCorrente(i).DtVcto = IIf(IsDBNull(rdr("DtVcto")), Nothing, Convert.ToDateTime(rdr("DtVcto")))
    '	'lstContaCorrente(i).DtPago = IIf(IsDBNull(rdr("DtPago")), Nothing, Convert.ToDateTime(rdr("DtPago")))
    '	'lstContaCorrente(i).DtPgto = IIf(IsDBNull(rdr("DtPgto")), Nothing, Convert.ToDateTime(rdr("DtPgto"), Funcoes.Cultura))
    '	lstContaCorrente(i).VlrPago = IIf(IsDBNull(rdr("VlrPago")), Nothing, Convert.ToDouble(rdr("VlrPago")))
    '                lstContaCorrente(i).VlrMulta = IIf(IsDBNull(rdr("VlrMulta")), Nothing, rdr("VlrMulta"))
    '                lstContaCorrente(i).VlrJuros = IIf(IsDBNull(rdr("VlrJuros")), Nothing, rdr("VlrJuros"))
    '                lstContaCorrente(i).VlrDesc = IIf(IsDBNull(rdr("VlrDesc")), Nothing, Convert.ToDouble(rdr("VlrDesc")))
    '                lstContaCorrente(i).VlrVarc = IIf(IsDBNull(rdr("VlrVarC")), Nothing, rdr("VlrVarC"))
    '                lstContaCorrente(i).ObsBaixa = rdr("ObsBaixa").ToString
    '                lstContaCorrente(i).CodEvento = rdr("CodEvento").ToString
    '                lstContaCorrente(i).VlrAbat = IIf(IsDBNull(rdr("VlrAbat")), Nothing, rdr("VlrAbat"))
    '                lstContaCorrente(i).VlrDevol = IIf(IsDBNull(rdr("VlrDevol")), Nothing, rdr("VlrDevol"))
    '                lstContaCorrente(i).CodCCusto = rdr("CodCCusto").ToString
    '                lstContaCorrente(i).Origem = rdr("Origem").ToString
    '                lstContaCorrente(i).CodHist = rdr("CodHist").ToString
    '                lstContaCorrente(i).Complemento = rdr("Complemento").ToString
    '                lstContaCorrente(i).CodClieForn = rdr("CodClieForn").ToString
    '                lstContaCorrente(i).CodBanco = rdr("NomeBanco").ToString
    '                lstContaCorrente(i).AliqRetIR = IIf(IsDBNull(rdr("AliqRetIR")), Nothing, rdr("AliqRetIR"))
    '                lstContaCorrente(i).AliqRetINSS = IIf(IsDBNull(rdr("AliqRetINSS")), Nothing, rdr("AliqRetINSS"))
    '                lstContaCorrente(i).AliqRetISS = IIf(IsDBNull(rdr("AliqRetISS")), Nothing, rdr("AliqRetISS"))
    '                lstContaCorrente(i).AliqNovaCOFINS = IIf(IsDBNull(rdr("AliqNovaCOFINS")), Nothing, rdr("AliqNovaCOFINS"))
    '                lstContaCorrente(i).VlrRetIR = IIf(IsDBNull(rdr("VlrRetIR")), Nothing, rdr("VlrRetIR"))
    '                lstContaCorrente(i).VlrRetINSS = IIf(IsDBNull(rdr("VlrRetINSS")), Nothing, rdr("VlrRetINSS"))
    '                lstContaCorrente(i).VlrRetISS = IIf(IsDBNull(rdr("VlrRetISS")), Nothing, rdr("VlrRetISS"))
    '                lstContaCorrente(i).VlrRetCOFINS = IIf(IsDBNull(rdr("VlrRetCOFINS")), Nothing, rdr("VlrRetCOFINS"))
    '                lstContaCorrente(i).CodTipoPgto = IIf(IsDBNull(rdr("CodTipoPgto")), Nothing, rdr("CodTipoPgto"))
    '                lstContaCorrente(i).TipoPgto = IIf(IsDBNull(rdr("TipoPgto")), Nothing, rdr("TipoPgto"))
    '                lstContaCorrente(i).CodLinhaDig = IIf(IsDBNull(rdr("CodLinhaDig")), Nothing, rdr("CodLinhaDig"))
    '                lstContaCorrente(i).CodHistBordero = rdr("CodHistBordero").ToString
    '                lstContaCorrente(i).CodEventoBordero = rdr("CodEventoBordero").ToString

    '                lstContaCorrente(i).Sucesso = True
    '                lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
    '                i = i + 1

    '            Loop
    '        Else
    '            lstContaCorrente.Add(New DetLctoCtaCorrente)
    '            lstContaCorrente(0).Sucesso = False
    '            lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
    '            lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
    '            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
    '        End If
    '        rdr.Close()
    '        Command.Dispose()
    '    Catch ex As Exception

    '        lstContaCorrente.Add(New DetLctoCtaCorrente)
    '        lstContaCorrente(0).Sucesso = False
    '        lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARITENSLCTOMOVIMENTACAOCTACORRENTEPORNUMLCTO.Descricao & ex.Message
    '        lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARITENSLCTOMOVIMENTACAOCTACORRENTEPORNUMLCTO.Id
    '        lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: P_BuscarItensLctoMovimentacaoCtaCorrentePorNumLcto(9)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    Finally
    '        connection.Close()
    '    End Try


    '    Return lstContaCorrente

    'End Function

    Public Function BuscarMaxNumLctoCtaCorrente(ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As LctoCtaCorrente


        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente As New LctoCtaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarMaxNumLctoCtaCorrente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans



            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                '_ContaCorrente.Add(New LctoCtaCorrente)
                _ContaCorrente.NumLcto = rdr("NumLcto")


                _ContaCorrente.Sucesso = True
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
            Else
                '_ContaCorrente.Add(New LctoCtaCorrente)
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            Command.Dispose()
            rdr.Close()
        Catch ex As Exception

            '_ContaCorrente.Add(New LctoCtaCorrente)
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARMAXNUMLCTOCTACORRENTE.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARMAXNUMLCTOCTACORRENTE.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarMaxNumLctoCtaCorrente(10)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            'connection.Close()
        End Try


        Return _ContaCorrente

    End Function

    Public Function BuscarSaldoResSaldoLiqLimiteCredContaCorrentePorCodBancoCodAgenNumCta(ByVal strCodBanco As String, ByVal strNumCta As String, ByVal strCodAgen As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As ContaCorrente


        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente As New ContaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarSaldoResSaldoLiqLimiteCredContaCorrentePorCodBancoCodAgenNumCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@CODBANCO", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", strNumCta))
            Command.Transaction = trans

            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                'lstContaCorrente.Add(New ContaCorrente)
                _ContaCorrente.LimiteCredito = rdr("LimiteCred")
                _ContaCorrente.SaldoLiq = rdr("SaldoLiq")
                _ContaCorrente.SaldoRes = rdr("SaldoRes")




                _ContaCorrente.Sucesso = True
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None

            Else
                'lstContaCorrente.Add(New ContaCorrente)
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
            End If
            Command.Dispose()
            rdr.Close()
        Catch ex As Exception

            'lstContaCorrente.Add(New ContaCorrente)
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARSALDORESSALDOLIQLIMITECREDCONTACORRENTEPORCODBANCOCODAGENNUMCTA.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARSALDORESSALDOLIQLIMITECREDCONTACORRENTEPORCODBANCOCODAGENNUMCTA.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarSaldoResSaldoLiqLimiteCredContaCorrentePorCodBancoCodAgenNumCta(11)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            'connection.Close()
        End Try


        Return _ContaCorrente

    End Function

    Public Function BuscarLctoCtaCorrentePorNumLcto(ByVal strNumlcto As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As LctoCtaCorrente


        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente As New LctoCtaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarLctoCtaCorrentePorNumLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumlcto))
            Command.Transaction = trans


            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                '_ContaCorrente.Add(New LctoCtaCorrente)
                _ContaCorrente.NumLcto = rdr("NumLcto")


                _ContaCorrente.Sucesso = True
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
            Else
                '_ContaCorrente.Add(New LctoCtaCorrente)
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            Command.Dispose()
            rdr.Close()
        Catch ex As Exception

            '_ContaCorrente.Add(New LctoCtaCorrente)
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARLCTOCTACORRENTEPORNUMLCTO.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARLCTOCTACORRENTEPORNUMLCTO.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarLctoCtaCorrentePorNumLcto(12)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            'connection.Close()
        End Try


        Return _ContaCorrente

    End Function

    Public Function BuscarSaldoCtaCorr(ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal strMesAno As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As SaldoCtaCorr


        ' Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _SaldoCtaCorr As New SaldoCtaCorr
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarSaldoCtaCorr", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@CODAGEN", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@CODBANCO", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", strNumCta))
            Command.Parameters.Add(New SqlParameter("@MESANO", strMesAno))
            Command.Transaction = trans


            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                '_ContaCorrente.Add(New LctoCtaCorrente)
                _SaldoCtaCorr.MesAnoBase = IIf(IsDBNull(rdr("MesAnoBase")), Nothing, rdr("MesAnoBase"))
                _SaldoCtaCorr.CodBanco = IIf(IsDBNull(rdr("CodBanco")), Nothing, rdr("CodBanco"))
                _SaldoCtaCorr.CodAgen = IIf(IsDBNull(rdr("CodAgen")), Nothing, rdr("CodAgen"))
                _SaldoCtaCorr.NumCta = IIf(IsDBNull(rdr("NumCta")), Nothing, rdr("NumCta"))
                _SaldoCtaCorr.SaldoInicial = IIf(IsDBNull(rdr("SaldoInicial")), Nothing, rdr("SaldoInicial"))
                _SaldoCtaCorr.TotVlrEntrada = IIf(IsDBNull(rdr("TotVlrEntrada")), Nothing, rdr("TotVlrEntrada"))
                _SaldoCtaCorr.TotVlrSaida = IIf(IsDBNull(rdr("TotVlrSaida")), Nothing, rdr("TotVlrSaida"))


                _SaldoCtaCorr.Sucesso = True
                _SaldoCtaCorr.TipoErro = DadosGenericos.TipoErro.None
            Else
                '_ContaCorrente.Add(New LctoCtaCorrente)
                _SaldoCtaCorr.Sucesso = False
                _SaldoCtaCorr.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _SaldoCtaCorr.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _SaldoCtaCorr.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            Command.Dispose()
            rdr.Close()
        Catch ex As Exception

            '_ContaCorrente.Add(New LctoCtaCorrente)
            _SaldoCtaCorr.Sucesso = False
            _SaldoCtaCorr.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARSALDOCTACORR.Descricao & ex.Message
            _SaldoCtaCorr.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARSALDOCTACORR.Id
            _SaldoCtaCorr.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _SaldoCtaCorr.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_SaldoCtaCorr.NumErro, _SaldoCtaCorr.MsgErro, _SaldoCtaCorr.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarSaldoCtaCorr(13)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            'connection.Close()
        End Try


        Return _SaldoCtaCorr

    End Function

    Public Function BuscarMaxNumLctoCtaCorrente() As LctoCtaCorrente


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente As New LctoCtaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarMaxNumLctoCtaCorrente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                '_ContaCorrente.Add(New LctoCtaCorrente)
                _ContaCorrente.NumLcto = rdr("NumLcto")


                _ContaCorrente.Sucesso = True
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
            Else
                '_ContaCorrente.Add(New LctoCtaCorrente)
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            Command.Dispose()
            rdr.Close()
        Catch ex As Exception

            '_ContaCorrente.Add(New LctoCtaCorrente)
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARMAXNUMLCTOCTACORRENTE.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARMAXNUMLCTOCTACORRENTE.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarMaxNumLctoCtaCorrente(14)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return _ContaCorrente

    End Function

    Public Function BuscarContaCorrentePorTipoCta(ByVal strTipoCta As String, Optional ByVal strStatus As String = "") As List(Of ContaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of ContaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarContaCorrentePorTipoCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@TIPOCTA", strTipoCta))
            Command.Parameters.Add(New SqlParameter("@Status", strStatus))



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New ContaCorrente)
                    lstContaCorrente(i).CodBanco = rdr("CodBanco")
                    lstContaCorrente(i).NumCta = rdr("NumCta")
                    lstContaCorrente(i).CodAgen = rdr("CodAgen")
                    lstContaCorrente(i).Descricao = rdr("Descricao")
                    lstContaCorrente(i).NomeAgen = rdr("NomeAgen").ToString




                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New ContaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstContaCorrente.Add(New ContaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARCONTACORRENTEPORTIPOCTA.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARCONTACORRENTEPORTIPOCTA.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarContaCorrentePorTipoCta(15)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscaContaCorrentePorTipoStatusJoinBanco() As List(Of ContaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of ContaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaContaCorrentePorTipoStatusJoinBanco", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New ContaCorrente)
                    lstContaCorrente(i).BancoAgen = rdr("BancoAgen").ToString
                    lstContaCorrente(i).NumCta = rdr("NumCta").ToString
                    lstContaCorrente(i).SaldoTotal = rdr("SaldoTotal")
                    lstContaCorrente(i).LimiteCredito = rdr("LimiteCred")
                    lstContaCorrente(i).SaldoLiq = rdr("SaldoLiq")
                    lstContaCorrente(i).SaldoRes = rdr("SaldoRes")
                    lstContaCorrente(i).NomeAgen = rdr("NomeAgen").ToString
                    lstContaCorrente(i).Descricao = rdr("Descricao").ToString
                    lstContaCorrente(i).CodCCusto = rdr("CodCCusto").ToString





                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New ContaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstContaCorrente.Add(New ContaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTACORRENTEPORTIPOSTATUSJOINBANCO.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTACORRENTEPORTIPOSTATUSJOINBANCO.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaContaCorrentePorTipoStatusJoinBanco(16)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscarChequesBorderoPorTipoLctoSitLctoCodBancoCodAgenNumCtaOrderByDtLcto(ByVal _Lcto As LctoCtaCorrente, ByVal intMeses As Integer, ByVal strSitLcto As String, ByVal strTipo As String) As List(Of LctoCtaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of LctoCtaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarChequesBorderoPorTipoLctoSitLctoCodBancoCodAgenNumCtaOrderByDtLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@CODBANCO", _Lcto.CodBanco))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", _Lcto.CodAgen))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", _Lcto.NumCta))
            Command.Parameters.Add(New SqlParameter("@MESES", intMeses))
            Command.Parameters.Add(New SqlParameter("@SITLCTO", strSitLcto))
            Command.Parameters.Add(New SqlParameter("@TIPO", strTipo))



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New LctoCtaCorrente)
					lstContaCorrente(i).DtLcto = Convert.ToDateTime(rdr("Data"), Funcoes.Cultura)
                    lstContaCorrente(i).DctoLcto = rdr("DctoLcto").ToString
                    lstContaCorrente(i).VlrLcto = CDbl(rdr("VlrLcto"))
                    lstContaCorrente(i).NumLcto = rdr("NumLcto")
                    lstContaCorrente(i).SitLcto = rdr("Situacao").ToString
                    lstContaCorrente(i).TipoLcto = rdr("Tipo").ToString
                    lstContaCorrente(i).Impresso = rdr("Impresso").ToString


                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New LctoCtaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
        Catch ex As Exception

            lstContaCorrente.Add(New LctoCtaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARCHEQUESBORDEROPORTIPOLCTOSITLCTOCODBANCOCODAGENNUMCTAORDERBYDTLCTO.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARCHEQUESBORDEROPORTIPOLCTOSITLCTOCODBANCOCODAGENNUMCTAORDERBYDTLCTO.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarChequesBorderoPorTipoLctoSitLctoCodBancoCodAgenNumCtaOrderByDtLcto(17)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscarChequesBorderoPorNumLcto(ByVal strNumLcto As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As List(Of DetLctoCtaCorrente)


        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of DetLctoCtaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarChequesBorderoPorNumLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumLcto))
            Command.Transaction = trans


            ''Abre a conexao
            'connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New DetLctoCtaCorrente)
                    'lstContaCorrente(i).CodBanco = Convert.ToDateTime(rdr("CodBanco"))
                    lstContaCorrente(i).Origem = rdr("Origem").ToString
                    lstContaCorrente(i).NumTit = rdr("NumTit").ToString
                    lstContaCorrente(i).SeqTit = rdr("SeqTit").ToString
                    lstContaCorrente(i).ClieForn = rdr("CodForn").ToString
					lstContaCorrente(i).CodClieForn = rdr("CodClieForn").ToString

					If (IsDBNull(rdr("DtEmi"))) Then
						lstContaCorrente(i).DtEmi = Convert.ToDateTime(rdr("DtEmi"), Funcoes.Cultura)
					Else
						lstContaCorrente(i).DtEmi = Nothing
					End If


					lstContaCorrente(i).VlrTit = IIf(IsDBNull(rdr("VlrTit")), 0.0, rdr("VlrTit"))

					If (Not IsDBNull(rdr("DtVcto"))) Then
						lstContaCorrente(i).DtVcto = Convert.ToDateTime(rdr("DtVcto"), Funcoes.Cultura)
					Else
						lstContaCorrente(i).DtVcto = Nothing
					End If

					'lstContaCorrente(i).DtVcto = IIf(IsDBNull(rdr("DtVcto")), Nothing, Convert.ToDateTime(rdr("DtVcto"), Funcoes.Cultura))
					'lstContaCorrente(i).DtPago = Convert.ToDateTime(rdr("DtPago"), Funcoes.Cultura)

					If (Not IsDBNull(rdr("DtPago"))) Then
						lstContaCorrente(i).DtPago = Convert.ToDateTime(rdr("DtPago"), Funcoes.Cultura)
					Else
						lstContaCorrente(i).DtPago = Nothing
					End If

					lstContaCorrente(i).VlrPago = IIf(IsDBNull(rdr("VlrPago")), 0, rdr("VlrPago"))
					lstContaCorrente(i).VlrMulta = IIf(IsDBNull(rdr("VlrMulta")), 0, rdr("VlrMulta"))
					lstContaCorrente(i).VlrJuros = IIf(IsDBNull(rdr("VlrJuros")), 0, rdr("VlrJuros"))
					lstContaCorrente(i).VlrDesc = IIf(IsDBNull(rdr("VlrDesc")), 0, rdr("VlrDesc"))
					lstContaCorrente(i).VlrVarCamb = IIf(IsDBNull(rdr("VlrVarCamb")), 0, rdr("VlrVarCamb"))
					lstContaCorrente(i).ObsBaixa = rdr("ObsBaixa").ToString()
					lstContaCorrente(i).CodEvento = rdr("CodEvento").ToString()
					lstContaCorrente(i).VlrAbat = IIf(IsDBNull(rdr("VlrAbat")), 0, rdr("VlrAbat"))

					lstContaCorrente(i).Sucesso = True
					lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
					i = i + 1

				Loop
            Else
                lstContaCorrente.Add(New DetLctoCtaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception

            lstContaCorrente.Add(New DetLctoCtaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARCHEQUESBORDEROPORNUMLCTO.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARCHEQUESBORDEROPORNUMLCTO.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
			Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            'connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscarChequesParaEstornoPorNumLcto(ByVal strNumLcto As String) As List(Of DetLctoCtaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of DetLctoCtaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarChequesParaEstornoPorNumLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumLcto))
            'Command.Transaction = trans


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New DetLctoCtaCorrente)
                    'lstContaCorrente(i).CodBanco = Convert.ToDateTime(rdr("CodBanco"))
                    'lstContaCorrente(i).Origem = rdr("Origem").ToString
                    lstContaCorrente(i).NumTit = rdr("NumTit").ToString
                    lstContaCorrente(i).SeqTit = rdr("SeqTit").ToString
                    lstContaCorrente(i).ClieForn = rdr("RazaoSocial").ToString
                    lstContaCorrente(i).DtEmissao = Convert.ToDateTime(rdr("DtEmissao"))
                    lstContaCorrente(i).VlrInd = rdr("VlrInd")
                    lstContaCorrente(i).DtVcto = Convert.ToDateTime(rdr("DtVcto"))
                    lstContaCorrente(i).ObsBaixa = rdr("ObsBaixa")
                    lstContaCorrente(i).CodCCusto = rdr("CodCCusto")
                    lstContaCorrente(i).CodClieForn = rdr("CodForn")





                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New DetLctoCtaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception

            lstContaCorrente.Add(New DetLctoCtaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARCHEQUESPARAESTORNOPORNUMLCTO.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARCHEQUESPARAESTORNOPORNUMLCTO.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarChequesParaEstornoPorNumLcto(19)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscarLctoCtaCorrentePorCodBancoCodAgenNumCtaNumLcto(ByVal _Lcto As LctoCtaCorrente) As LctoCtaCorrente


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente As New LctoCtaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarLctoCtaCorrentePorCodBancoCodAgenNumCtaNumLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@CODBANCO", _Lcto.CodBanco))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", _Lcto.CodAgen))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", _Lcto.NumCta))
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", _Lcto.NumLcto))




            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    '_ContaCorrente.Add(New LctoCtaCorrente)
                    _ContaCorrente.DtLcto = Convert.ToDateTime(rdr("DtLcto"))
                    _ContaCorrente.CodHist = IIf(IsDBNull(rdr("CodHist")), Nothing, rdr("CodHist"))
                    _ContaCorrente.Complemento = IIf(IsDBNull(rdr("Complemento")), Nothing, rdr("Complemento"))
                    _ContaCorrente.Nominal = IIf(IsDBNull(rdr("Nominal")), Nothing, rdr("Nominal"))
                    _ContaCorrente.VersoCheque = (IIf(IsDBNull(rdr("VersoCheque")), Nothing, rdr("VersoCheque")))
                    _ContaCorrente.Impresso = IIf(IsDBNull(rdr("Impresso")), Nothing, rdr("Impresso"))


                    _ContaCorrente.Sucesso = True
                    _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                '_ContaCorrente.Add(New LctoCtaCorrente)
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
        Catch ex As Exception

            '_ContaCorrente.Add(New LctoCtaCorrente)
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARLCTOCTACORRENTEPORCODBANCOCODAGENNUMCTANUMLCTO.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARLCTOCTACORRENTEPORCODBANCOCODAGENNUMCTANUMLCTO.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarLctoCtaCorrentePorCodBancoCodAgenNumCtaNumLcto(20)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return _ContaCorrente

    End Function

    Public Function BuscaDetLancamentosContaCorrentePorNumLcto(ByVal strNumLcto As String) As List(Of DetLctoCtaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of DetLctoCtaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaDetLancamentosContaCorrentePorNumLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumLcto))



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New DetLctoCtaCorrente)
                    'lstContaCorrente(i).CodBanco = Convert.ToDateTime(rdr("CodBanco"))
                    lstContaCorrente(i).NumTit = rdr("NumTit")
					lstContaCorrente(i).DtEmissao = Convert.ToDateTime(rdr("DtEmissao"), Funcoes.Cultura)
                    lstContaCorrente(i).SeqTit = rdr("SeqTit")
					lstContaCorrente(i).NumLcto = rdr("NumLcto")
					If (Not IsDBNull(rdr("DtPgto"))) Then
						lstContaCorrente(i).DtPgto = Convert.ToDateTime(rdr("DtPgto"), Funcoes.Cultura)
					End If
					lstContaCorrente(i).CodClieForn = rdr("CodClieForn")
					lstContaCorrente(i).OrigTit = rdr("OrigTit")


					lstContaCorrente(i).Sucesso = True
					lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
					i = i + 1

				Loop
            Else
                lstContaCorrente.Add(New DetLctoCtaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception

            lstContaCorrente.Add(New DetLctoCtaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaDetLancamentosContaCorrentePorNumLcto(21)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscarRazaoSocialPorNumLctoJoinForn(ByVal strNumLcto As String) As List(Of DetLctoCtaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of DetLctoCtaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarRazaoSocialPorNumLctoJoinForn", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumLcto))



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New DetLctoCtaCorrente)
                    'lstContaCorrente(i).CodBanco = Convert.ToDateTime(rdr("CodBanco"))
                    lstContaCorrente(i).ClieForn = rdr("RazaoSocial")



                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New DetLctoCtaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception

            lstContaCorrente.Add(New DetLctoCtaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarRazaoSocialPorNumLctoJoinForn(22)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscarContaCorrentePorCodBancoCodAgenNumCta(ByVal strCodBanco As String, ByVal strNumCta As String, ByVal strCodAgen As String) As ContaCorrente


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente As New ContaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarContaCorrentePorCodBancoCodAgenNumCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@CODBANCO", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", strNumCta))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                'lstContaCorrente.Add(New ContaCorrente)
				_ContaCorrente.TextoBordero = IIf(IsDBNull(rdr("TextoBordero")), "", rdr("TextoBordero"))




                _ContaCorrente.Sucesso = True
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None

            Else
                'lstContaCorrente.Add(New ContaCorrente)
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
            End If
            Command.Dispose()
            rdr.Close()
        Catch ex As Exception

            'lstContaCorrente.Add(New ContaCorrente)
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARCONTACORRENTEPORCODBANCOCODAGENNUMCTA.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARCONTACORRENTEPORCODBANCOCODAGENNUMCTA.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarContaCorrentePorCodBancoCodAgenNumCta(23)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return _ContaCorrente

    End Function

    Public Function BuscarQuantidadeTitulosValorPagoTotalPorNumLcto(ByVal strNumLcto As String) As DetLctoCtaCorrente


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente As New DetLctoCtaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarQuantidadeTitulosValorPagoTotalPorNumLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumLcto))



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    'lstContaCorrente.Add(New DetLctoCtaCorrente)
                    'lstContaCorrente(i).CodBanco = Convert.ToDateTime(rdr("CodBanco"))
                    _ContaCorrente.VlrPago = rdr("VlrPago")
                    _ContaCorrente.Qtde = rdr("Qtde")



                    _ContaCorrente.Sucesso = True
                    _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                'lstContaCorrente.Add(New DetLctoCtaCorrente)
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception

            'lstContaCorrente.Add(New DetLctoCtaCorrente)
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCALANCAMENTOSCONTACORRENTEGRID.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarQuantidadeTitulosValorPagoTotalPorNumLcto(24)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return _ContaCorrente

    End Function

    Public Function BuscarTitulosVlrDtLctoPorNumLcto(ByVal strNumLcto As String) As List(Of DetLctoCtaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of DetLctoCtaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarTitulosVlrDtLctoPorNumLcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", strNumLcto))



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New DetLctoCtaCorrente)
                    'lstContaCorrente(i).CodBanco = Convert.ToDateTime(rdr("CodBanco"))
                    lstContaCorrente(i).NumTit = rdr("NumTit").ToString
                    lstContaCorrente(i).CodClieForn = rdr("CodClieForn").ToString
                    lstContaCorrente(i).VlrTit = rdr("VlrTit")
                    lstContaCorrente(i).DtVcto = Convert.ToDateTime(rdr("DtVcto"))
                    lstContaCorrente(i).VlrPago = rdr("VlrPago")



                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New DetLctoCtaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
            Command.Dispose()
        Catch ex As Exception

            lstContaCorrente.Add(New DetLctoCtaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOSVLRDTLCTOPORNUMLCTO.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARTITULOSVLRDTLCTOPORNUMLCTO.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarTitulosVlrDtLctoPorNumLcto(25)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function

    Public Function BuscaIsnullContaCtblPorCodBancoCodAgenNumCta(ByVal strCodBanco As String,
                                                                ByVal strCodAgen As String,
                                                                ByVal strNumCta As String) As ContaCorrente

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _ContaCorrente = New ContaCorrente
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaIsnullContaCtblPorCodBancoCodAgenNumCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))


            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                rdr.Read()

                _ContaCorrente = New ContaCorrente

                _ContaCorrente.ContaCtbl = rdr("ContaCtbl").ToString
                _ContaCorrente.LayoutCheque = rdr("LayoutCheque").ToString
                _ContaCorrente.LayoutBordero = rdr("LayoutBordero").ToString

                _ContaCorrente.Sucesso = True
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
            Else
                _ContaCorrente = New ContaCorrente
                _ContaCorrente.Sucesso = False
                _ContaCorrente.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _ContaCorrente.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _ContaCorrente.TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            _ContaCorrente = New ContaCorrente
            _ContaCorrente.Sucesso = False
            _ContaCorrente.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAISNULLCONTACTBLPORCODBANCOCODAGENNUMCTA.Descricao & ex.Message
            _ContaCorrente.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAISNULLCONTACTBLPORCODBANCOCODAGENNUMCTA.Id
            _ContaCorrente.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _ContaCorrente.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_ContaCorrente.NumErro, _ContaCorrente.MsgErro, _ContaCorrente.TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscaIsnullContaCtblPorCodBancoCodAgenNumCta(26)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return _ContaCorrente

    End Function

    Public Function BuscarBancosGridRelSaldo() As List(Of ContaCorrente)


        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of ContaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarBancosGridRelSaldo", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query



            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                Do While (rdr.Read)

                    lstContaCorrente.Add(New ContaCorrente)
                    'lstContaCorrente(i).BancoAgen = rdr("BancoAgen").ToString
                    lstContaCorrente(i).CodBanco = rdr("CodBanco").ToString
                    lstContaCorrente(i).CodAgen = rdr("CodAgen").ToString
                    lstContaCorrente(i).NumCta = rdr("NumCta").ToString
                    lstContaCorrente(i).SaldoTotal = rdr("SaldoTotal")
                    lstContaCorrente(i).LimiteCredito = rdr("LimiteCred")
                    lstContaCorrente(i).SaldoLiq = rdr("SaldoLiq")
                    lstContaCorrente(i).SaldoRes = rdr("SaldoRes")
                    lstContaCorrente(i).NomeAgen = rdr("NomeAgen").ToString
                    lstContaCorrente(i).Descricao = rdr("Descricao").ToString
                    'lstContaCorrente(i).CodCCusto = rdr("CodCCusto").ToString





                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1

                Loop
            Else
                lstContaCorrente.Add(New ContaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            lstContaCorrente.Add(New ContaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARBANCOSGRIDRELSALDO.Descricao & ex.Message
            lstContaCorrente(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARBANCOSGRIDRELSALDO.Id
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: ContaCorrenteBC - Classe: ConsultaContasReceber - Função: BuscarBancosGridRelSaldo(27)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente
    End Function

    Public Function GerarExportacaoContasPagar(ByVal data As DateTime) As List(Of LctoCtaCorrente)

        Dim _LctoConta As New LctoCtaCorrente
        Dim lstContas As New List(Of LctoCtaCorrente)

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_GerarExportContPag", connection)
        Dim rdr As SqlDataReader
        Dim i As Integer = 0

        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            command.Parameters.Add(New SqlParameter("@DtLcto", Funcoes.FormatDate(data, 0).Replace("-", "")))

            connection.Open()


            rdr = command.ExecuteReader
            If rdr.HasRows Then
                While rdr.Read()
                    _LctoConta = New LctoCtaCorrente()
                    With _LctoConta
                        .NumLcto = rdr("NumLcto").ToString()
                        .DctoLcto = rdr("DctoLcto").ToString()
                        .DtLcto = rdr("DataLancamento").ToString()
                        .VlrLcto = rdr("VlrLcto").ToString()
                        .SitLcto = rdr("Situacao").ToString()
                        .CodBanco = rdr("CodBanco").ToString()
                        .CodAgen = rdr("CodAgen").ToString()
                        .NumCta = rdr("NumCta").ToString()
                        .TotalLctoContaCorrente = rdr("TotLctCC")
                        .TotalBaixaContasPagar = rdr("TotBaixaContPag")
                    End With


                    _LctoConta.Sucesso = True
                    _LctoConta.TipoErro = DadosGenericos.TipoErro.None

                    lstContas.Add(_LctoConta)
                End While

            Else
                _LctoConta.Sucesso = False
                _LctoConta.TipoErro = DadosGenericos.TipoErro.Funcional
                lstContas.Add(_LctoConta)
            End If

            rdr.Close()
            command.Dispose()
        Catch ex As Exception

            _LctoConta.Sucesso = False
            _LctoConta.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _LctoConta.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _LctoConta.MsgErro = ErrorConstants.EXCEPTION_METODO_GERAREXPORTACAOCONTASPAGAR.Descricao & ex.Message
            _LctoConta.NumErro = ErrorConstants.EXCEPTION_METODO_GERAREXPORTACAOCONTASPAGAR.Id

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_LctoConta.NumErro, _LctoConta.MsgErro, _LctoConta.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), _
                                          UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try


        Return lstContas


    End Function

    ''' <summary>
    '''  Gera as linhas para o arquivo de exportação bancária. 
    ''' </summary>
    ''' <param name="dtLcto">Data do lançamento para gerar o arquivo</param>
    ''' <returns>Retorna um DataSet contendo o retorno da procedure P_ExpTarifasBancarias, que contém duas colunas.</returns>
    ''' <remarks></remarks>
    Public Function GerarExportacaoTarifasBancarias(ByVal dtLcto As DateTime, ByVal intTipoExportacao As Integer) As DataSet

        Dim _Retorno As New Retorno()

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_ExpTarifasBancarias", connection)

        Dim i As Integer = 0
        Dim ds As DataSet = Nothing

        Try
            With command
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = DadosGenericos.Timeout.Query
                .Parameters.Add(New SqlParameter("@DtLcto", Funcoes.FormatDate(dtLcto, 0).Replace("-", "")))
                .Parameters.Add(New SqlParameter("@TipoExportacao", intTipoExportacao))
            End With

            connection.Open()

            Dim da As New SqlDataAdapter(command)
            ds = New DataSet()
            da.Fill(ds)

        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_GERAREXPORTACAOTARIFASBANCARIAS.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_GERAREXPORTACAOTARIFASBANCARIAS.Id

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), _
                                          UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            command.Dispose()
            connection.Close()
        End Try

        Return ds

    End Function


    Public Function BuscaContaCorrrentePorContaCtblCCusto(ByVal contaContabil As String, ByVal cCusto As String) As List(Of ContaCorrente)

        Dim con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstContCorrente As New List(Of ContaCorrente)

        Try
            con.Open()
            Dim cmd As New SqlCommand("P_BuscaCCorPorContaCtblCCusto", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@ContaCtbl", contaContabil))
            cmd.Parameters.Add(New SqlParameter("@CodCCusto", cCusto))

            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            If (rdr.HasRows) Then
                While (rdr.Read())
                    Dim _CC As New ContaCorrente()
                    _CC.Sucesso = True
                    _CC.TipoErro = DadosGenericos.TipoErro.None
                    _CC.CodBanco = rdr("CodBanco")
                    _CC.CodAgen = rdr("CodAgen")
                    _CC.NumCta = rdr("NumCta")
                    _CC.ContaCtbl = contaContabil
                    lstContCorrente.Add(_CC)
                End While
            Else
                Dim _CC As New ContaCorrente()
                _CC.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                _CC.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                _CC.TipoErro = DadosGenericos.TipoErro.Funcional
                _CC.Sucesso = False
                _CC.TipoErro = DadosGenericos.TipoErro.Funcional

                lstContCorrente.Add(_CC)
            End If

            rdr.Close()
        Catch ex As Exception

            Dim _Retorno As New ContaCorrente()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTACORRRENTEPORCONTACTBLCCUSTO.Id
                .MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTACORRRENTEPORCONTACTBLCCUSTO.Descricao
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            lstContCorrente.Add(_Retorno)
        Finally
            con.Close()
        End Try

        Return lstContCorrente

    End Function

    Public Function BuscaContaHistContEmCtaCorrente(ByVal codHist As String, ByVal cCustoEmpresa As String, ByVal tipoLcto As String) As Retorno

        Dim con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno()

        Try
            con.Open()

            Dim cmd As New SqlCommand("P_ConstaHistContEmCtaCorrente", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@CodHist", codHist))
            cmd.Parameters.Add(New SqlParameter("@Empresa", cCustoEmpresa))
            cmd.Parameters.Add(New SqlParameter("@TipoLcto", cCustoEmpresa))

            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            If (rdr.Read()) Then
                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Else
                _Retorno = New Retorno()
                With _Retorno
                    .Sucesso = False
                    .NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                    .MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                    .TipoErro = DadosGenericos.TipoErro.Arquitetura
                    .ImagemErro = DadosGenericos.ImagemRetorno.Erro
                End With

            End If
        Catch ex As Exception

            _Retorno = New Retorno()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_CONSULTAHISTCONTCONTPAG.Id
                .MsgErro = ErrorConstants.EXCEPTION_METODO_CONSULTAHISTCONTCONTPAG.Descricao
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            con.Close()
        End Try

        Return _Retorno

    End Function


    Public Function BuscaContasAtivas(ByVal strStatus As String) As List(Of ContaCorrente)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim lstContas As New List(Of ContaCorrente)
        Dim command As New SqlCommand("P_BuscaContasAtivas", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Dim i As Integer = 0
        Dim rdr As SqlDataReader
        Try
            connection.Open()


            command.Parameters.Add(New SqlParameter("@STATUS", IIf(String.IsNullOrEmpty(strStatus), DBNull.Value, strStatus)))

            rdr = command.ExecuteReader()
            If rdr.HasRows() Then
                Do While (rdr.Read())
                    lstContas.Add(New ContaCorrente)
                    lstContas(i).CodBanco = IIf(IsDBNull(rdr("CodBanco")), Nothing, rdr("CodBanco"))
                    lstContas(i).CodAgen = IIf(IsDBNull(rdr("CodAgen")), Nothing, rdr("CodAgen"))
                    lstContas(i).NumCta = IIf(IsDBNull(rdr("NumCta")), Nothing, rdr("NumCta"))
                    lstContas(i).Descricao = IIf(IsDBNull(rdr("Descricao")), Nothing, rdr("Descricao"))
                    lstContas(i).NomeBanco = IIf(IsDBNull(rdr("NomeBanco")), Nothing, rdr("NomeBanco"))
                    lstContas(i).CodCCustoDNI = IIf(IsDBNull(rdr("CodCCustoDNI")), Nothing, rdr("CodCCustoDNI"))

                    lstContas(i).Sucesso = True
                    lstContas(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            Else
                lstContas.Add(New ContaCorrente)
                lstContas(0).Sucesso = False
                lstContas(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstContas(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContas(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
            End If
            rdr.Close()
        Catch ex As Exception
            lstContas.Add(New ContaCorrente)
            lstContas(0).Sucesso = False
            lstContas(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContas(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASATIVAS.Descricao & ex.Message
            lstContas(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTASATIVAS.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContas(0).NumErro, lstContas(0).MsgErro, lstContas(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return lstContas
    End Function


End Class


