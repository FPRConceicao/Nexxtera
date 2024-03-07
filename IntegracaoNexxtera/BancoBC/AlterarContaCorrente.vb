Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Entidades
Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection


Public Class AlterarContaCorrente


    Public Function AlterarSeqRemesContaCorrente(ByVal strCodBanco As String,
                                                 ByVal strCodAgen As String,
                                                 ByVal strNumCta As String,
                                                 ByVal Connection As SqlConnection,
                                                 ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno


        Dim Command As SqlCommand = New SqlCommand("P_AlterarSeqRemesContaCorrente", Connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))


            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_ALTERARSEQREMESCONTACORRENTE.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_ALTERARSEQREMESCONTACORRENTE.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarSeqRemesContaCorrente(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno

    End Function


    Public Function AlterarContaCorrente(ByVal _ContaCorrente As ContaCorrente,
                                            ByVal connection As SqlConnection,
                                            ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno


        Dim Command As SqlCommand = New SqlCommand("P_AlterarContaCorrente", connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodBanco), DBNull.Value, _ContaCorrente.CodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodAgen), DBNull.Value, _ContaCorrente.CodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.NumCta), DBNull.Value, _ContaCorrente.NumCta)))
            Command.Parameters.Add(New SqlParameter("@DESCRICAO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.Descricao), DBNull.Value, _ContaCorrente.Descricao)))
            Command.Parameters.Add(New SqlParameter("@LIMITECREDITO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.LimiteCredito), DBNull.Value, _ContaCorrente.LimiteCredito)))
            Command.Parameters.Add(New SqlParameter("@FLUXOCAIXA", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.Fluxo), DBNull.Value, _ContaCorrente.Fluxo)))
            Command.Parameters.Add(New SqlParameter("@TIPOCTA", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.Tipo), DBNull.Value, _ContaCorrente.Tipo)))
            Command.Parameters.Add(New SqlParameter("@TEXTOBORDERO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.TextoBordero), DBNull.Value, _ContaCorrente.TextoBordero)))
            Command.Parameters.Add(New SqlParameter("@LAYOUTCHEQUE", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.LayoutCheque), DBNull.Value, _ContaCorrente.LayoutCheque)))
            Command.Parameters.Add(New SqlParameter("@CODCCUSTO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodCCusto), DBNull.Value, _ContaCorrente.CodCCusto)))
            Command.Parameters.Add(New SqlParameter("@CONTACTBL", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.ContaCtbl), DBNull.Value, _ContaCorrente.ContaCtbl)))
            Command.Parameters.Add(New SqlParameter("@ISDEBAUTO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.IsDebAuto), DBNull.Value, _ContaCorrente.IsDebAuto)))
            Command.Parameters.Add(New SqlParameter("@FLOATCREDITO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.FloatCredito), DBNull.Value, _ContaCorrente.FloatCredito)))
            Command.Parameters.Add(New SqlParameter("@SEQREMES", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.SeqRemes), DBNull.Value, _ContaCorrente.SeqRemes)))
            Command.Parameters.Add(New SqlParameter("@CODCARTEIRA", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodCarteira), DBNull.Value, _ContaCorrente.CodCarteira)))
            Command.Parameters.Add(New SqlParameter("@TIPOCONTA", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.TipoConta), DBNull.Value, _ContaCorrente.TipoConta)))
            Command.Parameters.Add(New SqlParameter("@NUMCTAVINC", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.NumCtaVinc), DBNull.Value, _ContaCorrente.NumCtaVinc)))
            Command.Parameters.Add(New SqlParameter("@NUMCTAVINCDESC", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.NumCtaVincDesc), DBNull.Value, _ContaCorrente.NumCtaVincDesc)))
            Command.Parameters.Add(New SqlParameter("@STATUS", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.Status), DBNull.Value, _ContaCorrente.Status)))
            Command.Parameters.Add(New SqlParameter("@CONVENIO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.Convenio), DBNull.Value, _ContaCorrente.Convenio)))
            'Command.Parameters.Add(New SqlParameter("@NOSSONUMERO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.NossoNumero), DBNull.Value, _ContaCorrente.NossoNumero)))
            Command.Parameters.Add(New SqlParameter("@PERMITEDI", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.PermiteDI), DBNull.Value, _ContaCorrente.PermiteDI)))
            Command.Parameters.Add(New SqlParameter("@PAGAMENTOELETRONICO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.PagamentoEletronico), DBNull.Value, _ContaCorrente.PagamentoEletronico)))
            Command.Parameters.Add(New SqlParameter("@CODCCUSTODNI", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodCCustoDNI), DBNull.Value, _ContaCorrente.CodCCustoDNI)))
            Command.Parameters.Add(New SqlParameter("@CONTAPADRAODEBAUTO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.ContaPadraoDebAuto), DBNull.Value, _ContaCorrente.ContaPadraoDebAuto)))





            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARCONTACORRENTE.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARCONTACORRENTE.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarContaCorrente(2)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno
    End Function

    ''' <summary>
    ''' Atualiza o valor de entrada ou saída do lançamento de acordo com o tipo
    ''' </summary>
    ''' <param name="_ContaCorrente"></param>
    ''' <param name="strMesAno"></param>
    ''' <param name="intTipo">1 para entradas e outros valores para saidas</param>
    ''' <param name="dblValor"></param>
    ''' <param name="connection"></param>
    ''' <param name="trans"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AtualizaValorEntraSaidaLancamentoCCorrente(ByVal _ContaCorrente As ContaCorrente, ByVal strMesAno As String, ByVal intTipo As Integer, ByVal dblValor As Double,
                                            ByVal connection As SqlConnection,
                                            ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno


        Dim Command As SqlCommand = New SqlCommand("P_AtualizaValorEntraSaidaLancamentoCCorrente", connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodBanco), DBNull.Value, _ContaCorrente.CodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodAgen), DBNull.Value, _ContaCorrente.CodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.NumCta), DBNull.Value, _ContaCorrente.NumCta)))
            Command.Parameters.Add(New SqlParameter("@MESANOBASE", IIf(String.IsNullOrWhiteSpace(strMesAno), DBNull.Value, strMesAno)))
            Command.Parameters.Add(New SqlParameter("@TIPO", IIf(String.IsNullOrWhiteSpace(intTipo), DBNull.Value, intTipo)))
            Command.Parameters.Add(New SqlParameter("@VALOR", IIf(String.IsNullOrWhiteSpace(dblValor), DBNull.Value, dblValor)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ATUALIZAVALORENTRASAIDALANCAMENTOCCORRENTE.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ATUALIZAVALORENTRASAIDALANCAMENTOCCORRENTE.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AtualizaValorEntraSaidaLancamentoCCorrente(3)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno
    End Function

    Public Function AtualizaSaldoCtaCorrentePorCodBancoCodAgenNumCtaDtLctoVlrLctoTipoLcto(ByVal _ContaCorrente As ContaCorrente, ByVal dtDtLcto As DateTime, ByVal dblVlrLcto As Double, ByVal strTipoLcto As String,
                                            ByVal connection As SqlConnection,
                                            ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno

        Dim Command As SqlCommand = New SqlCommand("P_AtualizaSaldoCtaCorrentePorCodBancoCodAgenNumCtaDtLctoVlrLctoTipoLcto", connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodBanco), DBNull.Value, _ContaCorrente.CodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.CodAgen), DBNull.Value, _ContaCorrente.CodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(_ContaCorrente.NumCta), DBNull.Value, _ContaCorrente.NumCta)))
            Command.Parameters.Add(New SqlParameter("@DTLCTO", Convert.ToDateTime(dtDtLcto)))
            Command.Parameters.Add(New SqlParameter("@VLRLCTO", IIf(String.IsNullOrWhiteSpace(dblVlrLcto), DBNull.Value, dblVlrLcto)))
            Command.Parameters.Add(New SqlParameter("@TIPOLCTO", IIf(String.IsNullOrWhiteSpace(strTipoLcto), DBNull.Value, strTipoLcto)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ATUALIZASALDOCTACORRENTEPORCODBANCOCODAGENNUMCTADTLCTOVLRLCTOTIPOLCTO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ATUALIZASALDOCTACORRENTEPORCODBANCOCODAGENNUMCTADTLCTOVLRLCTOTIPOLCTO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AtualizaSaldoCtaCorrentePorCodBancoCodAgenNumCtaDtLctoVlrLctoTipoLcto(4)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno
    End Function


    Public Function AtualizarSaldoCCorPorSaidaOuEntrada(ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal strDtLcto As String, ByVal dblVlrLcto As Double, ByVal strTipoLcto As String,
                                          ByVal connection As SqlConnection,
                                          ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno

        Dim Command As SqlCommand = New SqlCommand("P_AtualizarSaldoCCorPorSaidaOuEntrada", connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@MESANO", strDtLcto))
            Command.Parameters.Add(New SqlParameter("@VLRDCTO", IIf(String.IsNullOrWhiteSpace(dblVlrLcto), DBNull.Value, dblVlrLcto)))
            Command.Parameters.Add(New SqlParameter("@TIPO", IIf(String.IsNullOrWhiteSpace(strTipoLcto), DBNull.Value, strTipoLcto)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ATUALIZARSALDOCCORPORSAIDAOUENTRADA.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ATUALIZARSALDOCCORPORSAIDAOUENTRADA.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AtualizarSaldoCCorPorSaidaOuEntrada(4)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno
    End Function

    'Public Function AlterarDetLctoCtaCorrente(ByVal _DetLctoCtaCorrente As DetLctoCtaCorrente,
    '                                      ByVal connection As SqlConnection,
    '                                      ByVal trans As SqlTransaction) As Retorno
    '    Dim _Retorno As New Retorno

    '    Dim Command As SqlCommand = New SqlCommand("P_AlterarDetLctoCtaCorrente", connection)
    '    Try
    '        ''Informa a procedure
    '        Command.CommandType = CommandType.StoredProcedure
    '        Command.CommandTimeout = DadosGenericos.Timeout.Query
    '        Command.Transaction = trans

    '        ''define os parametros usados na stored procedure
    '        Command.Parameters.Add(New SqlParameter("@NUMLCTO", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.NumLcto), DBNull.Value, _DetLctoCtaCorrente.NumLcto)))
    '        Command.Parameters.Add(New SqlParameter("@ORIGTIT", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.OrigTit), DBNull.Value, _DetLctoCtaCorrente.OrigTit)))
    '        Command.Parameters.Add(New SqlParameter("@NUMTIT", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.NumTit), DBNull.Value, _DetLctoCtaCorrente.NumTit)))
    '        Command.Parameters.Add(New SqlParameter("@SEQTIT", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.SeqTit), DBNull.Value, _DetLctoCtaCorrente.SeqTit)))
    '        Command.Parameters.Add(New SqlParameter("@CODCLIEFORN", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.CodClieForn), DBNull.Value, _DetLctoCtaCorrente.CodClieForn)))
    '        Command.Parameters.Add(New SqlParameter("@DTEMISSAO", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.DtEmi), DBNull.Value, Convert.ToDateTime(_DetLctoCtaCorrente.DtEmi))))
    '        Command.Parameters.Add(New SqlParameter("@DTPGTO", IIf(IsDate(_DetLctoCtaCorrente.DtPgto), Convert.ToDateTime(_DetLctoCtaCorrente.DtPgto), DBNull.Value)))
    '        Command.Parameters.Add(New SqlParameter("@VLRIND", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.VlrInd), DBNull.Value, _DetLctoCtaCorrente.VlrInd)))
    '        Command.Parameters.Add(New SqlParameter("@VLRPAGO", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.VlrPago), DBNull.Value, _DetLctoCtaCorrente.VlrPago)))
    '        Command.Parameters.Add(New SqlParameter("@VLRMULTA", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.VlrMulta), DBNull.Value, _DetLctoCtaCorrente.VlrMulta)))
    '        Command.Parameters.Add(New SqlParameter("@VLRJUROS", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.VlrJuros), DBNull.Value, _DetLctoCtaCorrente.VlrJuros)))
    '        Command.Parameters.Add(New SqlParameter("@VLRDESC", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.VlrDesc), DBNull.Value, _DetLctoCtaCorrente.VlrDesc)))
    '        Command.Parameters.Add(New SqlParameter("@VLRVARCAMB", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.VlrVarc), DBNull.Value, _DetLctoCtaCorrente.VlrVarc)))
    '        Command.Parameters.Add(New SqlParameter("@OBSBAIXA", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.ObsBaixa), DBNull.Value, _DetLctoCtaCorrente.ObsBaixa)))
    '        Command.Parameters.Add(New SqlParameter("@CODEVENTO", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.CodEvento), DBNull.Value, _DetLctoCtaCorrente.CodEvento)))
    '        Command.Parameters.Add(New SqlParameter("@VLRABAT", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.VlrAbat), DBNull.Value, _DetLctoCtaCorrente.VlrAbat)))
    '        Command.Parameters.Add(New SqlParameter("@VLRDEVOL", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.VlrDevol), DBNull.Value, _DetLctoCtaCorrente.VlrDevol)))
    '        Command.Parameters.Add(New SqlParameter("@USRULTALT", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.UsrUltAlt), DBNull.Value, _DetLctoCtaCorrente.UsrUltAlt)))
    '        Command.Parameters.Add(New SqlParameter("@CODCCUSTO", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.CodCCusto), DBNull.Value, _DetLctoCtaCorrente.CodCCusto)))
    '        Command.Parameters.Add(New SqlParameter("@ORIGEM", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.Origem), DBNull.Value, _DetLctoCtaCorrente.Origem)))
    '        Command.Parameters.Add(New SqlParameter("@CODHIST", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.CodHist), DBNull.Value, _DetLctoCtaCorrente.CodHist)))
    '        Command.Parameters.Add(New SqlParameter("@COMPLEMENTO", IIf(String.IsNullOrWhiteSpace(_DetLctoCtaCorrente.Complemento), DBNull.Value, _DetLctoCtaCorrente.Complemento)))




    '        Command.ExecuteNonQuery()

    '        _Retorno.Sucesso = True
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.None
    '        Command.Dispose()
    '    Catch ex As Exception
    '        _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARDETLCTOCTACORRENTE.Descricao + ex.Message
    '        _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARDETLCTOCTACORRENTE.Id
    '        _Retorno.Sucesso = False
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarDetLctoCtaCorrente(5)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    End Try
    '    Return _Retorno
    'End Function

    Public Function AtualizarSldAtualCCorrentePorEntradaSaidaAlteracaoInsercao(ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal dblGValor As Double, ByVal dblVlrDcto As Double, ByVal strTipoLcto As String,
                                        ByVal connection As SqlConnection,
                                        ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno

        Dim Command As SqlCommand = New SqlCommand("P_AtualizarSldAtualCCorrentePorEntradaSaidaAlteracaoInsercao", connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@GVALOR", IIf(String.IsNullOrWhiteSpace(dblGValor), DBNull.Value, dblGValor)))
            Command.Parameters.Add(New SqlParameter("@VLRDCTO", IIf(String.IsNullOrWhiteSpace(dblVlrDcto), DBNull.Value, dblVlrDcto)))
            Command.Parameters.Add(New SqlParameter("@TIPO", IIf(String.IsNullOrWhiteSpace(strTipoLcto), DBNull.Value, strTipoLcto)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ATUALIZARSLDATUALCCORRENTEPORENTRADASAIDAALTERACAOINSERCAO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ATUALIZARSLDATUALCCORRENTEPORENTRADASAIDAALTERACAOINSERCAO.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AtualizarSldAtualCCorrentePorEntradaSaidaAlteracaoInsercao(6)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno
    End Function

    Public Function AlterarConciliadoDtConciliadoPorNumLcto(ByVal strStatus As String, ByVal intNumLcto As Integer, ByVal dtDtConciliacao As DateTime) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_AlterarConciliadoDtConciliadoPorNumLcto", connection)
        Try
            connection.Open()
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", IIf(String.IsNullOrWhiteSpace(intNumLcto), DBNull.Value, intNumLcto)))
            Command.Parameters.Add(New SqlParameter("@STATUS", IIf(String.IsNullOrWhiteSpace(strStatus), DBNull.Value, strStatus)))
            Command.Parameters.Add(New SqlParameter("@DTCONCILIACAO", IIf(String.IsNullOrWhiteSpace(dtDtConciliacao), DBNull.Value, Convert.ToDateTime(dtDtConciliacao))))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARCONCILIADODTCONCILIADOPORNUMLCTO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARCONCILIADODTCONCILIADOPORNUMLCTO.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarConciliadoDtConciliadoPorNumLcto(7)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try
        Return _Retorno
    End Function

    Public Function AlterarSitLctoLctoCtaCorrentePorNumLcto(ByVal strSitLcto As String, ByVal intNumLcto As Integer, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        ' Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_AlterarSitLctoLctoCtaCorrentePorNumLcto", connection)
        Try
            ' connection.Open()
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans
            Command.Transaction = trans
            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", IIf(String.IsNullOrWhiteSpace(intNumLcto), DBNull.Value, intNumLcto)))
            Command.Parameters.Add(New SqlParameter("@SITLCTO", IIf(String.IsNullOrWhiteSpace(strSitLcto), DBNull.Value, strSitLcto)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARSITLCTOLCTOCTACORRENTEPORNUMLCTO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARSITLCTOLCTOCTACORRENTEPORNUMLCTO.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarSitLctoLctoCtaCorrentePorNumLcto(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            ' connection.Close()
        End Try
        Return _Retorno
    End Function

    Public Function AlterarSldReservadoSldAtualPorCodBancoCodAgenNumCta(ByVal dblTotal As Double, ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        ' Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_AlterarSldReservadoSldAtualPorCodBancoCodAgenNumCta", connection)
        Try
            ' connection.Open()
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans
            Command.Transaction = trans
            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@TOTAL", IIf(String.IsNullOrWhiteSpace(dblTotal), DBNull.Value, dblTotal)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARSLDRESERVADOSLDATUALPORCODBANCOCODAGENNUMCTA.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARSLDRESERVADOSLDATUALPORCODBANCOCODAGENNUMCTA.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarSldReservadoSldAtualPorCodBancoCodAgenNumCta(9)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            ' connection.Close()
        End Try
        Return _Retorno
    End Function

    Public Function CancelarLancamentoChequeBordero(ByVal strSituacao As String, ByVal intNumlcto As Integer, ByVal dblValor As Double, ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_CancelarLancamentoChequeBordero", connection)
        Dim rdr As SqlDataReader
        Try
            connection.Open()
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@VALOR", IIf(String.IsNullOrWhiteSpace(dblValor), DBNull.Value, dblValor)))
            Command.Parameters.Add(New SqlParameter("@SITUACAO", IIf(String.IsNullOrWhiteSpace(strSituacao), DBNull.Value, strSituacao)))
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", IIf(String.IsNullOrWhiteSpace(intNumlcto), DBNull.Value, intNumlcto)))



            rdr = Command.ExecuteReader()
            If rdr.HasRows Then
                rdr.Read()
                If rdr("Sucesso") = "TRUE" Then
                    _Retorno.Sucesso = True
                    _Retorno.TipoErro = DadosGenericos.TipoErro.None
                ElseIf rdr("Sucesso") = "FALSE" Then
                    _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_CANCELARLANCAMENTOCHEQUEBORDERO.Descricao + rdr("ERRORMESSAGE").ToString
                    _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_CANCELARLANCAMENTOCHEQUEBORDERO.Id
                    _Retorno.Sucesso = False
                    _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                    _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
                    'CRIAR LOG NO WINDOWS
                    Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: CancelarLancamentoChequeBordero(10)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
                End If
            Else
                _Retorno.Sucesso = False
                _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                _Retorno.MsgErro = ErrorConstants.ERRO_AO_CANCELAR_CHEQUE.Descricao
                _Retorno.NumErro = ErrorConstants.ERRO_AO_CANCELAR_CHEQUE.Id
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If

            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_CANCELARLANCAMENTOCHEQUEBORDERO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_CANCELARLANCAMENTOCHEQUEBORDERO.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: CancelarLancamentoChequeBordero(10)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try
        Return _Retorno
    End Function

    'Public Function EstornarChequeBordero(ByVal _DetLcto As DetLctoCtaCorrente) As Retorno
    '    Dim _Retorno As New Retorno
    '    Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '    Dim Command As SqlCommand = New SqlCommand("P_EstornarChequeBordero", connection)
    '    Dim rdr As SqlDataReader
    '    Try
    '        connection.Open()
    '        ''Informa a procedure
    '        Command.CommandType = CommandType.StoredProcedure
    '        Command.CommandTimeout = DadosGenericos.Timeout.Query
    '        'Command.Transaction = trans

    '        ''define os parametros usados na stored procedure
    '        Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(_DetLcto.CodBanco), DBNull.Value, _DetLcto.CodBanco)))
    '        Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(_DetLcto.CodAgen), DBNull.Value, _DetLcto.CodAgen)))
    '        Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(_DetLcto.NumCta), DBNull.Value, _DetLcto.NumCta)))
    '        Command.Parameters.Add(New SqlParameter("@DCTOLCTO", IIf(String.IsNullOrWhiteSpace(_DetLcto.DctoLcto), DBNull.Value, _DetLcto.DctoLcto)))
    '        Command.Parameters.Add(New SqlParameter("@VALOR", IIf(String.IsNullOrWhiteSpace(_DetLcto.VlrInd), DBNull.Value, _DetLcto.VlrInd)))
    '        Command.Parameters.Add(New SqlParameter("@USR", IIf(String.IsNullOrWhiteSpace(_DetLcto.UsrUltAlt), DBNull.Value, _DetLcto.UsrUltAlt)))
    '        Command.Parameters.Add(New SqlParameter("@NUMTIT", IIf(String.IsNullOrWhiteSpace(_DetLcto.NumTit), DBNull.Value, _DetLcto.NumTit)))
    '        Command.Parameters.Add(New SqlParameter("@SEQTIT", IIf(String.IsNullOrWhiteSpace(_DetLcto.SeqTit), DBNull.Value, _DetLcto.SeqTit)))
    '        Command.Parameters.Add(New SqlParameter("@CODCLIEFORN", IIf(String.IsNullOrWhiteSpace(_DetLcto.CodClieForn), DBNull.Value, _DetLcto.CodClieForn)))
    '        Command.Parameters.Add(New SqlParameter("@DTEMISSAO", IIf(String.IsNullOrWhiteSpace(_DetLcto.DtEmissao), DBNull.Value, Convert.ToDateTime(_DetLcto.DtEmissao))))
    '        Command.Parameters.Add(New SqlParameter("@CHEQUE", IIf(String.IsNullOrWhiteSpace(_DetLcto.DctoLcto), DBNull.Value, _DetLcto.DctoLcto)))
    '        Command.Parameters.Add(New SqlParameter("@CODCCUSTO", IIf(String.IsNullOrWhiteSpace(_DetLcto.CodCCusto), DBNull.Value, _DetLcto.CodCCusto)))
    '        Command.Parameters.Add(New SqlParameter("@NUMLCTOE", IIf(String.IsNullOrWhiteSpace(_DetLcto.NumLcto), DBNull.Value, _DetLcto.NumLcto)))




    '        rdr = Command.ExecuteReader()
    '        If rdr.HasRows Then
    '            rdr.Read()
    '            If rdr("Sucesso") = "TRUE" Then
    '                _Retorno.Sucesso = True
    '                _Retorno.TipoErro = DadosGenericos.TipoErro.None
    '            ElseIf rdr("Sucesso") = "FALSE" Then
    '                _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ESTORNARCHEQUEBORDERO.Descricao + rdr("ERRORMESSAGE").ToString
    '                _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ESTORNARCHEQUEBORDERO.Id
    '                _Retorno.Sucesso = False
    '                _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
    '                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
    '                'CRIAR LOG NO WINDOWS
    '                Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: EstornarChequeBordero(11)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '            End If
    '        Else
    '            _Retorno.Sucesso = False
    '            _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
    '            _Retorno.MsgErro = ErrorConstants.ERRO_AO_ESTORNAR_CHEQUE.Descricao
    '            _Retorno.NumErro = ErrorConstants.ERRO_AO_ESTORNAR_CHEQUE.Id
    '            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
    '        End If

    '        Command.Dispose()
    '    Catch ex As Exception
    '        _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ESTORNARCHEQUEBORDERO.Descricao + ex.Message
    '        _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ESTORNARCHEQUEBORDERO.Id
    '        _Retorno.Sucesso = False
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: EstornarChequeBordero(11)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    Finally
    '        connection.Close()
    '    End Try
    '    Return _Retorno
    'End Function


    Public Function AlterarVersoComplementoCodHistLctoCtaCorrente(ByVal strComplemento As String, ByVal strCodHist As String, ByVal strVersoCheque As String, ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal intNumLcto As Integer) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_AlterarVersoComplementoCodHistLctoCtaCorrente", connection)
        Try
            connection.Open()
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@COMPLEMENTO", IIf(String.IsNullOrWhiteSpace(strComplemento), DBNull.Value, strComplemento)))
            Command.Parameters.Add(New SqlParameter("@CODHIST", IIf(String.IsNullOrWhiteSpace(strCodHist), DBNull.Value, strCodHist)))
            Command.Parameters.Add(New SqlParameter("@VERSOCHEQUE", IIf(String.IsNullOrWhiteSpace(strVersoCheque), DBNull.Value, strVersoCheque)))
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", IIf(String.IsNullOrWhiteSpace(intNumLcto), DBNull.Value, intNumLcto)))


            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARVERSOCOMPLEMENTOCODHISTLCTOCTACORRENTE.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARVERSOCOMPLEMENTOCODHISTLCTOCTACORRENTE.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarVersoComplementoCodHistLctoCtaCorrente(12)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try
        Return _Retorno
    End Function

    Public Function AlterarNominalImpressoLctoCtaCorrente(ByVal intNumLcto As Integer, ByVal strImpresso As String, ByVal strNominal As String, ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal codHist As String, ByVal compl As String) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_AlterarNominalImpressoLctoCtaCorrente", connection)
        Try
            connection.Open()
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", IIf(String.IsNullOrWhiteSpace(intNumLcto), DBNull.Value, intNumLcto)))
            Command.Parameters.Add(New SqlParameter("@IMPRESSO", IIf(String.IsNullOrWhiteSpace(strImpresso), DBNull.Value, strImpresso)))
            Command.Parameters.Add(New SqlParameter("@NOMINAL", IIf(String.IsNullOrWhiteSpace(strNominal), DBNull.Value, strNominal)))
            Command.Parameters.Add(New SqlParameter("@CODHIST", IIf(String.IsNullOrWhiteSpace(codHist), DBNull.Value, codHist)))
            Command.Parameters.Add(New SqlParameter("@COMPL", IIf(String.IsNullOrWhiteSpace(compl), DBNull.Value, compl)))


            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARNOMINALIMPRESSOLCTOCTACORRENTE.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARNOMINALIMPRESSOLCTOCTACORRENTE.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarNominalImpressoLctoCtaCorrente(13)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try
        Return _Retorno
    End Function

    Public Function Alterar(ByVal intNumLcto As Integer, ByVal strImpresso As String, ByVal strNominal As String, ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String) As Retorno

    End Function

    'Public Function AlterarLctoCtaCorrentePorNumLcto(ByVal _LctoContaCorrente As LctoCtaCorrente, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno

    '    Return AlterarLctoCtaCorrentePorNumLcto(_LctoContaCorrente.NumLcto, _LctoContaCorrente.DtLcto, _LctoContaCorrente.DctoLcto, _LctoContaCorrente.VlrLcto, connection, trans)

    'End Function


    Public Function AlterarLctoCtaCorrentePorNumLcto(ByVal intNumLcto As Integer, ByVal dtDtLcto As DateTime, ByVal strDctoLcto As String, ByVal dblVlrLcto As Double, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_AlterarLctoCtaCorrentePorNumLcto", connection)
        Try
            'connection.Open()
            ''Informa a procedure
            Command.Transaction = trans
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@DTLCTO", IIf(String.IsNullOrWhiteSpace(dtDtLcto), DBNull.Value, Convert.ToDateTime(dtDtLcto))))
            Command.Parameters.Add(New SqlParameter("@DCTOLCTO", IIf(String.IsNullOrWhiteSpace(strDctoLcto), DBNull.Value, strDctoLcto)))
            Command.Parameters.Add(New SqlParameter("@VLRLCTO", IIf(String.IsNullOrWhiteSpace(dblVlrLcto), DBNull.Value, dblVlrLcto)))
            Command.Parameters.Add(New SqlParameter("@NUMLCTO", IIf(String.IsNullOrWhiteSpace(intNumLcto), DBNull.Value, intNumLcto)))
            Command.Parameters.Add(New SqlParameter("@USR", IIf(String.IsNullOrWhiteSpace(UsuarioGlobal.Usuario), DBNull.Value, UsuarioGlobal.Usuario)))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARLCTOCTACORRENTEPORNUMLCTO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARLCTOCTACORRENTEPORNUMLCTO.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarLctoCtaCorrentePorNumLcto(14)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            'connection.Close()
        End Try
        Return _Retorno
    End Function

    Public Function AlterarTotVlrSaidaSaldoCtaCorr(ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal strDtLcto As String, ByVal dblVlrLcto As Double,
                                                   ByVal connection As SqlConnection,
                                                   ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno

        Dim Command As SqlCommand = New SqlCommand("P_AlterarTotVlrSaidaSaldoCtaCorr", connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@MESANOBASE", strDtLcto))
            Command.Parameters.Add(New SqlParameter("@VLR", IIf(String.IsNullOrWhiteSpace(dblVlrLcto), DBNull.Value, dblVlrLcto)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARTOTVLRSAIDASALDOCTACORR.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARTOTVLRSAIDASALDOCTACORR.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarTotVlrSaidaSaldoCtaCorr(15)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno
    End Function

    Public Function AlterarSldReservadoSldAtualContaCorrentePorTipoChequeBordero(ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal strDtLcto As String, ByVal dblVlrDcto As Double, ByVal dblIVlrDcto As Double,
                                                                                 ByVal strTipo As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno

        Dim Command As SqlCommand = New SqlCommand("P_AlterarSldReservadoSldAtualContaCorrentePorTipoChequeBordero", connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@DTLCTO", IIf(String.IsNullOrWhiteSpace(strDtLcto), DBNull.Value, strDtLcto)))
            Command.Parameters.Add(New SqlParameter("@VLRDCTO", IIf(String.IsNullOrWhiteSpace(dblVlrDcto), DBNull.Value, dblVlrDcto)))
            Command.Parameters.Add(New SqlParameter("@IVLRDCTO", IIf(String.IsNullOrWhiteSpace(dblIVlrDcto), DBNull.Value, dblIVlrDcto)))
            Command.Parameters.Add(New SqlParameter("@TIPO", IIf(String.IsNullOrWhiteSpace(strTipo), DBNull.Value, strTipo)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARSLDRESERVADOSLDATUALCONTACORRENTEPORTIPOCHEQUEBORDERO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARSLDRESERVADOSLDATUALCONTACORRENTEPORTIPOCHEQUEBORDERO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarSldReservadoSldAtualContaCorrentePorTipoChequeBordero(16)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno
    End Function

    Public Function AlterarTotVlrSaidaSaldoCtaCorrMenosVlr(ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal strDtLcto As String, ByVal dblVlrLcto As Double,
                                                  ByVal connection As SqlConnection,
                                                  ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno

        Dim Command As SqlCommand = New SqlCommand("P_AlterarTotVlrSaidaSaldoCtaCorrMenosVlr", connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@MESANOBASE", IIf(Format(Date.MinValue, "yyyyMM").Equals(strDtLcto), DBNull.Value, strDtLcto)))
            Command.Parameters.Add(New SqlParameter("@VLR", IIf(String.IsNullOrWhiteSpace(dblVlrLcto), DBNull.Value, dblVlrLcto)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARTOTVLRSAIDASALDOCTACORRMENOSVLR.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARTOTVLRSAIDASALDOCTACORRMENOSVLR.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarTotVlrSaidaSaldoCtaCorrMenosVlr(17)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno
    End Function

    Public Function AlterarSldReservadoSldAtualPorCodBancoCodAgenNumCtaTipoLcto(ByVal dblTotal As Double, ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction,
                                                                                ByVal strTipo As String) As Retorno
        Dim _Retorno As New Retorno
        ' Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim Command As SqlCommand = New SqlCommand("P_AlterarSldReservadoSldAtualPorCodBancoCodAgenNumCtaTipoLcto", connection)
        Try
            ' connection.Open()
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            'Command.Transaction = trans
            Command.Transaction = trans
            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CODBANCO", IIf(String.IsNullOrWhiteSpace(strCodBanco), DBNull.Value, strCodBanco)))
            Command.Parameters.Add(New SqlParameter("@CODAGEN", IIf(String.IsNullOrWhiteSpace(strCodAgen), DBNull.Value, strCodAgen)))
            Command.Parameters.Add(New SqlParameter("@NUMCTA", IIf(String.IsNullOrWhiteSpace(strNumCta), DBNull.Value, strNumCta)))
            Command.Parameters.Add(New SqlParameter("@TOTAL", IIf(String.IsNullOrWhiteSpace(dblTotal), DBNull.Value, dblTotal)))
            Command.Parameters.Add(New SqlParameter("@TIPO", IIf(String.IsNullOrWhiteSpace(strTipo), DBNull.Value, strTipo)))



            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARSLDRESERVADOSLDATUALPORCODBANCOCODAGENNUMCTATIPOLCTO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARSLDRESERVADOSLDATUALPORCODBANCOCODAGENNUMCTATIPOLCTO.Id
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaCorrenteBC - Classe: AlterarContaCorrente - Função: AlterarSldReservadoSldAtualPorCodBancoCodAgenNumCtaTipoLcto(18)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            ' connection.Close()
        End Try
        Return _Retorno
    End Function

    Public Function AtualizaVlrLctoContaCorrente(numLcto As Integer, somaValores As Double, con As SqlConnection, trans As SqlTransaction) As Retorno

        Dim _Retorno As New Retorno()

        Try

            Dim cmd As New SqlCommand("P_AtualizaVlrLcto", con, trans)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@NumLcto", numLcto))
            cmd.Parameters.Add(New SqlParameter("@SomaValores", somaValores))

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None


        Catch ex As Exception


            _Retorno = New Retorno()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_ATUALIZAVLRLCTOCONTACORRENTE.Id
                .MsgErro = ErrorConstants.EXCEPTION_METODO_ATUALIZAVLRLCTOCONTACORRENTE.Descricao + ex.Message
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno

    End Function

    Public Function VoltarLctoCtaCorrente(ByVal numTit As Integer, ByVal con As SqlConnection, ByVal trans As SqlTransaction) As Retorno

        Dim _Retorno As New Retorno()

        Try
            Dim cmd As New SqlCommand("P_VoltarLctoCtaCorrente", con, trans)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@Numtit", numTit))

            cmd.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception

            Dim _Retorno1 As New Retorno()
            With _Retorno1
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_BUSCALCTOCTACORRENTEPORTIPOLCTOS.Id
                .MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCALCTOCTACORRENTEPORTIPOLCTOS.Descricao + ex.Message
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            Funcoes.AtualizaApplEventLog(_Retorno1.NumErro, _Retorno1.MsgErro, _Retorno1.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

            _Retorno = _Retorno1
        End Try

        Return _Retorno
    End Function

    Public Function AlterarSeqRemesContaCorrente(ByVal strCodBanco As String, ByVal strCodAgen As String, ByVal strNumCta As String) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_AlterarSeqRemesContaCorrente", connection)
        Command.CommandType = CommandType.StoredProcedure
        Command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parametros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_ALTERARSEQREMESCONTACORRENTE.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_ALTERARSEQREMESCONTACORRENTE.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            Command.Dispose()
        End Try

        Return _Retorno
    End Function
End Class
