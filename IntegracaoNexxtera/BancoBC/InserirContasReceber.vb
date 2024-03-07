Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Common.ErrorConstants
Imports Teleatlantic.TLS.Entidades

Imports System.Data
Imports System.Data.SqlClient

Public Class InserirContasReceber

    Public Function InserirContaReceber(ByVal _ContasReceber As ContaReceber,
                                          ByVal connection As SqlConnection,
                                          ByVal Transacao As SqlTransaction) As Retorno '#16#
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InserirContaReceber", connection)
            Command.Transaction = Transacao
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit     ", _ContasReceber.NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit     ", _ContasReceber.SeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao  ", _ContasReceber.DtEmissao))
            Command.Parameters.Add(New SqlParameter("@Situacao   ", _ContasReceber.Situacao))
            Command.Parameters.Add(New SqlParameter("@CodClie    ", _ContasReceber.CodClie))
            Command.Parameters.Add(New SqlParameter("@DtVcto     ", _ContasReceber.DtVcto))
            Command.Parameters.Add(New SqlParameter("@CodInd     ", _ContasReceber.CodInd))
            Command.Parameters.Add(New SqlParameter("@VlrInd     ", _ContasReceber.VlrInd))
            Command.Parameters.Add(New SqlParameter("@CodBanco   ", _ContasReceber.CodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen    ", _ContasReceber.CodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta     ", _ContasReceber.NumCta))
            Command.Parameters.Add(New SqlParameter("@NumPortador", _ContasReceber.NumPortador))
            Command.Parameters.Add(New SqlParameter("@ObsTit     ", _ContasReceber.ObsTit))
            Command.Parameters.Add(New SqlParameter("@UsrCad     ", _ContasReceber.UsrCad))
            Command.Parameters.Add(New SqlParameter("@DtCad      ", _ContasReceber.DtCad))
            Command.Parameters.Add(New SqlParameter("@UsrAlt     ", _ContasReceber.UsrAlt))
            Command.Parameters.Add(New SqlParameter("@DtAlt      ", _ContasReceber.DtAlt))
            Command.Parameters.Add(New SqlParameter("@StatusBco  ", _ContasReceber.StatusBco))
            Command.Parameters.Add(New SqlParameter("@CodCCusto  ", _ContasReceber.CodCCusto))
            Command.Parameters.Add(New SqlParameter("@TipoDup    ", _ContasReceber.TipoDup))
            Command.Parameters.Add(New SqlParameter("@MenBco     ", _ContasReceber.MenBco))
            Command.Parameters.Add(New SqlParameter("@Tipopagto  ", _ContasReceber.TipoPagto))
            Command.Parameters.Add(New SqlParameter("@CodBancoDeb", IIf(IsNothing(_ContasReceber.CodBancoDeb), DBNull.Value, _ContasReceber.CodBancoDeb)))
            Command.Parameters.Add(New SqlParameter("@CodAgenDeb ", IIf(IsNothing(_ContasReceber.CodAgenDeb), DBNull.Value, _ContasReceber.CodAgenDeb)))
            Command.Parameters.Add(New SqlParameter("@NumCtaDeb  ", IIf(IsNothing(_ContasReceber.NumCtaDeb), DBNull.Value, _ContasReceber.NumCtaDeb)))
            Command.Parameters.Add(New SqlParameter("@MsgBole    ", _ContasReceber.MsgBole))
            Command.Parameters.Add(New SqlParameter("@TitDesc    ", _ContasReceber.TitDesc))
            Command.Parameters.Add(New SqlParameter("@IsTitNegoc ", _ContasReceber.IsTitNegoc))
            Command.Parameters.Add(New SqlParameter("@IdFilial   ", _ContasReceber.IdFilial))



            ''Executa a procedure
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
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InserirContaReceber(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function

    Public Function InserirRetornoCC(ByVal RetornoCC As RetornoCC, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InserirRetornoCC", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Tipo", RetornoCC.Tipo))
            Command.Parameters.Add(New SqlParameter("@NumAviso", RetornoCC.NumAviso))
            Command.Parameters.Add(New SqlParameter("@NumTit", RetornoCC.NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", RetornoCC.SeqTit))
            Command.Parameters.Add(New SqlParameter("@CodBanco", RetornoCC.CodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", RetornoCC.CodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta", RetornoCC.NumCta))
            Command.Parameters.Add(New SqlParameter("@VlrPago", RetornoCC.VlrPago))
            Command.Parameters.Add(New SqlParameter("@VlrJuros", RetornoCC.VlrJuros))
            Command.Parameters.Add(New SqlParameter("@VlrDesc", RetornoCC.VlrDesc))
            Command.Parameters.Add(New SqlParameter("@VlrIOF", RetornoCC.VlrIOF))
            Command.Parameters.Add(New SqlParameter("@VlrAbat", RetornoCC.VlrAbat))
            Command.Parameters.Add(New SqlParameter("@Processado", RetornoCC.Processado))
            Command.Parameters.Add(New SqlParameter("@DtVcto", RetornoCC.DtVcto))
            Command.Parameters.Add(New SqlParameter("@DtPagto", RetornoCC.DtPgto))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = "InserirRetornoCC"
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InserirRetornoCC", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function

    Public Function BaixaCRecCartao(ByVal NumAviso As String, ByVal Tipo As String, ByVal Usuario As String, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BaixaCRecCartao", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.FaturamentoCartaoCredito

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumAviso", NumAviso))
            Command.Parameters.Add(New SqlParameter("@Cartao", Tipo))
            Command.Parameters.Add(New SqlParameter("@Usuario", Usuario))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = "BaixaCRecCartao"
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: BaixaCRecCartao", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function

    Public Function InsereHistReenvioBoleto(ByVal CodIntClie As String, ByVal NumTit As String, ByVal SeqTit As String, ByVal DtEmissao As DateTime, ByVal IdFilial As String) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereHistReenvioBoleto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@NumTit", NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", SeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", IdFilial))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = "InsereHistReenvioBoleto"
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InsereHistReenvioBoleto", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try

        Return _Retorno
    End Function

    'Public Function InsereHistCadastroItau(idBU As Integer, Usr As String, RetornoItau As RetornoItau) As Retorno
    '    Dim _Retorno As New Retorno
    '    Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '    Try
    '        ''Informa a procedure
    '        Dim Command As SqlCommand = New SqlCommand("P_InsereLogRetornoItau_Sucesso", connection)
    '        Command.CommandType = CommandType.StoredProcedure
    '        Command.CommandTimeout = DadosGenericos.Timeout.Query
    '        connection.Open()

    '        ''define os parƒmetros usados na stored procedure

    '        'Beneficiário
    '        Command.Parameters.Add(New SqlParameter("@cpf_cnpj_beneficiario", RetornoItau.beneficiario.cpf_cnpj_beneficiario))
    '        Command.Parameters.Add(New SqlParameter("@agencia_beneficiario", RetornoItau.beneficiario.agencia_beneficiario))
    '        Command.Parameters.Add(New SqlParameter("@conta_beneficiario", RetornoItau.beneficiario.conta_beneficiario))
    '        Command.Parameters.Add(New SqlParameter("@digito_verificador_conta_beneficiario", RetornoItau.beneficiario.digito_verificador_conta_beneficiario))

    '        'Pagador
    '        Command.Parameters.Add(New SqlParameter("@cpf_cnpj_pagador", RetornoItau.pagador.cpf_cnpj_pagador))
    '        Command.Parameters.Add(New SqlParameter("@nome_pagador", RetornoItau.pagador.nome_razao_social_pagador))
    '        Command.Parameters.Add(New SqlParameter("@logradouro_pagador", RetornoItau.pagador.logradouro_pagador))
    '        Command.Parameters.Add(New SqlParameter("@cidade_pagador", RetornoItau.pagador.cidade_pagador))
    '        Command.Parameters.Add(New SqlParameter("@uf_pagador", RetornoItau.pagador.uf_pagador))
    '        Command.Parameters.Add(New SqlParameter("@cep_pagador", RetornoItau.pagador.cep_pagador))

    '        'Sacador Avalista
    '        Command.Parameters.Add(New SqlParameter("@cpf_cnpj_sacador_avalista", RetornoItau.sacador_avalista.cpf_cnpj_sacador_avalista))
    '        Command.Parameters.Add(New SqlParameter("@nome_razao_social_sacador_avalista", RetornoItau.sacador_avalista.nome_razao_social_sacador_avalista))

    '        'Moeda
    '        Command.Parameters.Add(New SqlParameter("@sigla_moeda", RetornoItau.moeda.sigla_moeda))
    '        Command.Parameters.Add(New SqlParameter("@quantidade_moeda", RetornoItau.moeda.quantidade_moeda))
    '        Command.Parameters.Add(New SqlParameter("@cotacao_moeda", RetornoItau.moeda.cotacao_moeda))

    '        'Dados
    '        Command.Parameters.Add(New SqlParameter("@vencimento_titulo", RetornoItau.vencimento_titulo))
    '        Command.Parameters.Add(New SqlParameter("@tipo_carteira_titulo", RetornoItau.tipo_carteira_titulo))
    '        Command.Parameters.Add(New SqlParameter("@nosso_numero", RetornoItau.nosso_numero))
    '        Command.Parameters.Add(New SqlParameter("@seu_numero", RetornoItau.seu_numero))
    '        Command.Parameters.Add(New SqlParameter("@especie_documento", RetornoItau.especie_documento))
    '        Command.Parameters.Add(New SqlParameter("@codigo_barras", RetornoItau.codigo_barras))
    '        Command.Parameters.Add(New SqlParameter("@numero_linha_digitavel", RetornoItau.numero_linha_digitavel))
    '        Command.Parameters.Add(New SqlParameter("@local_pagamento", RetornoItau.local_pagamento))
    '        Command.Parameters.Add(New SqlParameter("@data_processamento", RetornoItau.data_processamento))
    '        Command.Parameters.Add(New SqlParameter("@data_emissao", RetornoItau.data_emissao))
    '        Command.Parameters.Add(New SqlParameter("@uso_banco", RetornoItau.uso_banco))
    '        Command.Parameters.Add(New SqlParameter("@valor_titulo", RetornoItau.valor_titulo))
    '        Command.Parameters.Add(New SqlParameter("@valor_desconto", RetornoItau.valor_desconto))
    '        Command.Parameters.Add(New SqlParameter("@valor_outra_deducao", RetornoItau.valor_outra_deducao))
    '        Command.Parameters.Add(New SqlParameter("@valor_juro_multa", RetornoItau.valor_juro_multa))
    '        Command.Parameters.Add(New SqlParameter("@valor_outro_acrescimo", RetornoItau.valor_outro_acrescimo))
    '        Command.Parameters.Add(New SqlParameter("@valor_total_cobrado", RetornoItau.valor_total_cobrado))
    '        'Command.Parameters.Add(New SqlParameter("@texto_informacao_cliente_beneficiario", RetornoItau.texto_informacao_cliente_beneficiario))
    '        'Command.Parameters.Add(New SqlParameter("@codigo_mensagem_erro", RetornoItau.codigo_mensagem_erro))
    '        Command.Parameters.Add(New SqlParameter("@idBU", idBU))
    '        Command.Parameters.Add(New SqlParameter("@Usr", Usr))

    '        ''Executa a procedure
    '        Command.ExecuteNonQuery()

    '        _Retorno.Sucesso = True
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

    '    Catch ex As Exception
    '        _Retorno.Sucesso = False
    '        _Retorno.NumErro = "InsereHistCadastroItau"
    '        _Retorno.MsgErro = ex.Message
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InsereHistCadastroItau", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    Finally
    '        connection.Close()
    '    End Try

    '    Return _Retorno
    'End Function

    Public Function InsereHistCadastroItauErro(Codigo As String, Mensagem As String, Campo As String, MensagemErroCampo As String, Valor As String, idBU As Integer, Usr As String) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereLogRetornoItau_Erro", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codigo", Codigo))
            Command.Parameters.Add(New SqlParameter("@Mensagem", Mensagem))
            Command.Parameters.Add(New SqlParameter("@Campo", Campo))
            Command.Parameters.Add(New SqlParameter("@MensagemErroCampo", MensagemErroCampo))
            Command.Parameters.Add(New SqlParameter("@Valor", Valor))
            Command.Parameters.Add(New SqlParameter("@idBU", idBU))
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = "InsereHistCadastroItauErro"
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InsereHistCadastroItauErro", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try

        Return _Retorno
    End Function

    Public Function InsereTokenItau(Token As token) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTokenItau", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Token", Token.access_token))
            Command.Parameters.Add(New SqlParameter("@Usr", UsuarioGlobal.Usuario()))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = "InsereTokenItau"
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InsereTokenItau", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
        End Try

        Return _Retorno
    End Function

    'Public Function InsHistCadSucessoItauFIMU(idBU As Integer, Usr As String, RetornoItau As RetornoItau, ByVal Boleto As ContaReceber) As Retorno
    '    Dim _Retorno As New Retorno
    '    Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '    Try
    '        ''Informa a procedure
    '        Dim Command As SqlCommand = New SqlCommand("P_InsereLogRetornoItau_Sucesso_FI_MU", connection)
    '        Command.CommandType = CommandType.StoredProcedure
    '        Command.CommandTimeout = DadosGenericos.Timeout.Query
    '        connection.Open()

    '        ''define os parƒmetros usados na stored procedure

    '        'Beneficiário
    '        Command.Parameters.Add(New SqlParameter("@cpf_cnpj_beneficiario", RetornoItau.beneficiario.cpf_cnpj_beneficiario))
    '        Command.Parameters.Add(New SqlParameter("@agencia_beneficiario", RetornoItau.beneficiario.agencia_beneficiario))
    '        Command.Parameters.Add(New SqlParameter("@conta_beneficiario", RetornoItau.beneficiario.conta_beneficiario))
    '        Command.Parameters.Add(New SqlParameter("@digito_verificador_conta_beneficiario", RetornoItau.beneficiario.digito_verificador_conta_beneficiario))

    '        'Pagador
    '        Command.Parameters.Add(New SqlParameter("@cpf_cnpj_pagador", RetornoItau.pagador.cpf_cnpj_pagador))
    '        Command.Parameters.Add(New SqlParameter("@nome_pagador", RetornoItau.pagador.nome_razao_social_pagador))
    '        Command.Parameters.Add(New SqlParameter("@logradouro_pagador", RetornoItau.pagador.logradouro_pagador))
    '        Command.Parameters.Add(New SqlParameter("@cidade_pagador", RetornoItau.pagador.cidade_pagador))
    '        Command.Parameters.Add(New SqlParameter("@uf_pagador", RetornoItau.pagador.uf_pagador))
    '        Command.Parameters.Add(New SqlParameter("@cep_pagador", RetornoItau.pagador.cep_pagador))

    '        'Sacador Avalista
    '        Command.Parameters.Add(New SqlParameter("@cpf_cnpj_sacador_avalista", RetornoItau.sacador_avalista.cpf_cnpj_sacador_avalista))
    '        Command.Parameters.Add(New SqlParameter("@nome_razao_social_sacador_avalista", RetornoItau.sacador_avalista.nome_razao_social_sacador_avalista))

    '        'Moeda
    '        Command.Parameters.Add(New SqlParameter("@sigla_moeda", RetornoItau.moeda.sigla_moeda))
    '        Command.Parameters.Add(New SqlParameter("@quantidade_moeda", RetornoItau.moeda.quantidade_moeda))
    '        Command.Parameters.Add(New SqlParameter("@cotacao_moeda", RetornoItau.moeda.cotacao_moeda))

    '        'Dados
    '        Command.Parameters.Add(New SqlParameter("@vencimento_titulo", RetornoItau.vencimento_titulo))
    '        Command.Parameters.Add(New SqlParameter("@tipo_carteira_titulo", RetornoItau.tipo_carteira_titulo))
    '        Command.Parameters.Add(New SqlParameter("@nosso_numero", RetornoItau.nosso_numero))
    '        Command.Parameters.Add(New SqlParameter("@seu_numero", RetornoItau.seu_numero))
    '        Command.Parameters.Add(New SqlParameter("@especie_documento", RetornoItau.especie_documento))
    '        Command.Parameters.Add(New SqlParameter("@codigo_barras", RetornoItau.codigo_barras))
    '        Command.Parameters.Add(New SqlParameter("@numero_linha_digitavel", RetornoItau.numero_linha_digitavel))
    '        Command.Parameters.Add(New SqlParameter("@local_pagamento", RetornoItau.local_pagamento))
    '        Command.Parameters.Add(New SqlParameter("@data_processamento", RetornoItau.data_processamento))
    '        Command.Parameters.Add(New SqlParameter("@data_emissao", RetornoItau.data_emissao))
    '        Command.Parameters.Add(New SqlParameter("@uso_banco", RetornoItau.uso_banco))
    '        Command.Parameters.Add(New SqlParameter("@valor_titulo", RetornoItau.valor_titulo))
    '        Command.Parameters.Add(New SqlParameter("@valor_desconto", RetornoItau.valor_desconto))
    '        Command.Parameters.Add(New SqlParameter("@valor_outra_deducao", RetornoItau.valor_outra_deducao))
    '        Command.Parameters.Add(New SqlParameter("@valor_juro_multa", RetornoItau.valor_juro_multa))
    '        Command.Parameters.Add(New SqlParameter("@valor_outro_acrescimo", RetornoItau.valor_outro_acrescimo))
    '        Command.Parameters.Add(New SqlParameter("@valor_total_cobrado", RetornoItau.valor_total_cobrado))
    '        'Command.Parameters.Add(New SqlParameter("@texto_informacao_cliente_beneficiario", RetornoItau.texto_informacao_cliente_beneficiario))
    '        'Command.Parameters.Add(New SqlParameter("@codigo_mensagem_erro", RetornoItau.codigo_mensagem_erro))
    '        Command.Parameters.Add(New SqlParameter("@idBU", idBU))
    '        Command.Parameters.Add(New SqlParameter("@Usr", Usr))
    '        Command.Parameters.Add(New SqlParameter("@NumTit", Boleto.NumTit))
    '        Command.Parameters.Add(New SqlParameter("@SeqTit", Boleto.SeqTit))
    '        Command.Parameters.Add(New SqlParameter("@DtEmissao", Boleto.DtEmissao))
    '        Command.Parameters.Add(New SqlParameter("@IdFilial", Boleto.IdFilial))

    '        ''Executa a procedure
    '        Command.ExecuteNonQuery()

    '        _Retorno.Sucesso = True
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

    '    Catch ex As Exception
    '        _Retorno.Sucesso = False
    '        _Retorno.NumErro = "InsereHistCadastroItauFIMU"
    '        _Retorno.MsgErro = ex.Message
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
    '    Finally
    '        connection.Close()
    '    End Try

    '    Return _Retorno
    'End Function

    Public Function InsHistCadItauErroFIMU(Codigo As String, Mensagem As String, Campo As String, MensagemErroCampo As String, Valor As String, idBU As Integer, Usr As String, ByVal Boleto As ContaReceber) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereLogRetornoItau_Erro_FI_MU", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codigo", Codigo))
            Command.Parameters.Add(New SqlParameter("@Mensagem", Mensagem))
            Command.Parameters.Add(New SqlParameter("@Campo", Campo))
            Command.Parameters.Add(New SqlParameter("@MensagemErroCampo", MensagemErroCampo))
            Command.Parameters.Add(New SqlParameter("@Valor", Valor))
            Command.Parameters.Add(New SqlParameter("@idBU", idBU))
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@NumTit", Boleto.NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", Boleto.SeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", Boleto.DtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", Boleto.IdFilial))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = "InsereHistCadastroItauErroFIMU"
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        Finally
            connection.Close()
        End Try

        Return _Retorno
    End Function

    Public Function InserirObsBoletoUnificado(ByVal IdBU As Integer, ByVal ObsBU As String, ByVal TextVarEmail As String, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InserirObsBoletoUnificado", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@IdBU", IIf(String.IsNullOrEmpty(IdBU), DBNull.Value, IdBU)))
            Command.Parameters.Add(New SqlParameter("@ObsBU", IIf(String.IsNullOrEmpty(ObsBU), DBNull.Value, ObsBU)))
            Command.Parameters.Add(New SqlParameter("@TextVarEmail", IIf(String.IsNullOrEmpty(TextVarEmail), DBNull.Value, TextVarEmail)))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_INSERIROBSBOLETOUNIFICADO.Id
            _Retorno.MsgErro = EXCEPTION_METODO_INSERIROBSBOLETOUNIFICADO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno

    End Function

    Public Function QuitarFinanciamentoCliente(ByVal codintclie As String, ByVal isBaixaBonificada As Boolean,
                                          ByVal dtPgto As DateTime, usr As String,
                                          ByVal numCta As String, ByVal codAgen As String,
                                          ByVal codBanco As String,
                                          ByVal connection As SqlConnection,
                                          ByVal Transacao As SqlTransaction) As Retorno '#16#
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_QuitaFinanciamentoCliente", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codintclie", codintclie))
            Command.Parameters.Add(New SqlParameter("@isBaixaBonificada", IIf(isBaixaBonificada, 1, 0)))
            Command.Parameters.Add(New SqlParameter("@DtPgto", dtPgto))
            Command.Parameters.Add(New SqlParameter("@Usr", usr))
            Command.Parameters.Add(New SqlParameter("@NumCta", numCta))
            Command.Parameters.Add(New SqlParameter("@CodAgen", codAgen))
            Command.Parameters.Add(New SqlParameter("@Bco", codBanco))



            ''Executa a procedure
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
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InserirContaReceber(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function

    Public Function InsereLogEnvioArquivoRemessaBanco(ByVal enviadoBanco As Boolean, ByVal CodBanco As String,
                                                      ByVal NumCta As String, ByVal CodAgen As String,
                                                      ByVal Tipo As String, ByVal Cartao As Boolean,
                                                      ByVal Empresa As String, ByVal DtHora As DateTime,
                                                      ByVal usr As String, ByVal numRemessa As String,
                                                      ByVal isCadOptante As Boolean,
                                                      ByVal connection As SqlConnection,
                                                      ByVal Transacao As SqlTransaction) As Retorno '#16#
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereLogEnvioArquivoRemessaBanco", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@EnviadoBanco", enviadoBanco))
            Command.Parameters.Add(New SqlParameter("@CodBanco", IIf(String.IsNullOrEmpty(CodBanco), DBNull.Value, CodBanco)))
            Command.Parameters.Add(New SqlParameter("@NumCta", IIf(String.IsNullOrEmpty(NumCta), DBNull.Value, NumCta)))
            Command.Parameters.Add(New SqlParameter("@CodAgen", IIf(String.IsNullOrEmpty(CodAgen), DBNull.Value, CodAgen)))
            Command.Parameters.Add(New SqlParameter("@Tipo", IIf(String.IsNullOrEmpty(Tipo), DBNull.Value, Tipo)))
            Command.Parameters.Add(New SqlParameter("@Cartao", Cartao))
            Command.Parameters.Add(New SqlParameter("@Empresa", IIf(String.IsNullOrEmpty(Empresa), DBNull.Value, Empresa)))
            Command.Parameters.Add(New SqlParameter("@DtHora", DtHora))
            Command.Parameters.Add(New SqlParameter("@Usr", usr))
            Command.Parameters.Add(New SqlParameter("@NumRemessa", numRemessa))
            Command.Parameters.Add(New SqlParameter("@IsCadOptante", isCadOptante))



            ''Executa a procedure
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
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InserirContaReceber(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function

    Public Function AtualizaLogEnvioBanco(ByVal log As LogEnvioArquivoRemessa, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno '#16#
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AtualizaLogEnvioArquivoRemessa", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Id", log.Id))
            Command.Parameters.Add(New SqlParameter("@UsrEnvioRemessa", UsuarioGlobal.Usuario))


            ''Executa a procedure
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
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContaReceberBC - Classe: InserirContasReceber - Função: InserirContaReceber(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function

    Public Function InserirMaisDeUmTituloContaReceber(_ContasReceber As ContaReceber, QtdeTituloGerar As Integer, Obs As String, connection As SqlConnection, transacao As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InserirMaisDeUmTituloContaReceber", connection)
            Command.Transaction = transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", _ContasReceber.NumTit))
            Command.Parameters.Add(New SqlParameter("@VlrUltimaParc", _ContasReceber.VlrInd))
            Command.Parameters.Add(New SqlParameter("@QtdeTitulos", QtdeTituloGerar))
            Command.Parameters.Add(New SqlParameter("@Obs", Obs))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERIRMAISDEUMTITULOCONTARECEBER.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERIRCONTASRECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function
End Class
