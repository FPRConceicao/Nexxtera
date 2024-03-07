Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Imports IntegracaoNexxtera.ErrorConstants



Public Class ContasAReceberAlterar



    ''' <summary>
    ''' Alterar a observação e a situação do título por codclie, Numtit, Seqtit, dtEmissao e IdFilial.
    ''' </summary>
    ''' <param name="_ContaReceber">Passsa a Entidade preenchida.</param>
    ''' <param name="connection">Passa a conexao aberta.</param>
    ''' <param name="Transacao">Passa a transacao</param>
    ''' <returns>
    ''' 
    ''' CREATE PROCEDURE [dbo].[P_AlterarContasReceberDI]
    '''
    '''	@ObsTit varchar(500),
    '''	@Situacao varchar(5),
    '''	@codclie varchar(16),
    '''	@Numtit varchar(15),  
    '''	@Seqtit  varchar(20),
    '''	@dtEmissao DateTime,
    '''	@IdFilial varchar(3)
    '''AS
    '''	UPDATE Contas_a_Receber SET ObsTit = @ObsTit, Situacao = @Situacao
    '''	WHERE codclie = @codclie AND 
    '''		  Numtit = @Numtit AND 
    '''		  Seqtit = @Seqtit AND 
    '''		  dtEmissao = @dtEmissao AND 
    '''		  IdFilial = @IdFilial
    ''' 
    ''' </returns>
    ''' <remarks></remarks>
    Public Function AlterarContasReceberDI(ByVal _ContaReceber As ContaReceber,
                                           ByVal connection As SqlConnection,
                                           ByVal Transacao As SqlTransaction) As Retorno '#1#
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarContasReceberDI", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@ObsTit", _ContaReceber.ObsTit))
            Command.Parameters.Add(New SqlParameter("@Situacao", _ContaReceber.Situacao))
            Command.Parameters.Add(New SqlParameter("@CodClie", _ContaReceber.CodClie))
            Command.Parameters.Add(New SqlParameter("@NumTit", _ContaReceber.NumTit))
            Command.Parameters.Add(New SqlParameter("@Seqtit", _ContaReceber.SeqTit))
            Command.Parameters.Add(New SqlParameter("@dtEmissao", _ContaReceber.DtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", _ContaReceber.IdFilial))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARCONTASRECEBERDI.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARCONTASRECEBERDI.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarContasReceberDI(1)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function


    Public Function AlterarFormContasReceberDI(ByVal _ContaReceber As ContaReceber,
                                               ByVal connection As SqlConnection,
                                               ByVal Transacao As SqlTransaction) As Retorno '#2#
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarFormContasReceberDI", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure

            Command.Parameters.Add(New SqlParameter("@Seqtit", _ContaReceber.SeqTit))
            Command.Parameters.Add(New SqlParameter("@codclie", _ContaReceber.CodClie))
            Command.Parameters.Add(New SqlParameter("@Numtit", _ContaReceber.NumTit))
            Command.Parameters.Add(New SqlParameter("@IdFilial", _ContaReceber.IdFilial))
            Command.Parameters.Add(New SqlParameter("@CodDI", _ContaReceber.CodDI))
            Command.Parameters.Add(New SqlParameter("@DtPrevPgDI", _ContaReceber.DtPrevPgDI))
            Command.Parameters.Add(New SqlParameter("@dtEmissao", _ContaReceber.DtEmissao))
            Command.Parameters.Add(New SqlParameter("@VlrPrevPgDI", Convert.ToDouble(_ContaReceber.VlrPrevPgDI)))
            Command.Parameters.Add(New SqlParameter("@TipoPgDI", _ContaReceber.TipoPagto))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARFORMCONTASRECEBERDI.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARFORMCONTASRECEBERDI.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarContasReceberDI(2)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function


    Public Function CancelaTitulosContasAReceber(ByVal strUsr As String,
                                                 ByVal strNumTit As String,
                                                 ByVal DtEmissao As DateTime,
                                                 ByVal Quitado As String,
                                                 ByVal connection As SqlConnection,
                                                 ByVal Transacao As SqlTransaction) As Retorno '#3#'
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_CancelaTitulosContasAReceber", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Usr", strUsr))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@Quitado", Quitado))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_CANCELATITULOSCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_CANCELATITULOSCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: CancelaTitulosContasAReceber(3)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function

    Public Function AlterarStatusBcoContaReceber(ByVal strNumTit As String,
                                                 ByVal strSeqTit As String,
                                                 ByVal DtEmissao As DateTime,
                                                 ByVal Connection As SqlConnection,
                                                 ByVal Transaction As SqlTransaction) As Retorno '#4#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarStatusBcoContaReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_ALTERARSTATUSBCOCONTARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_ALTERARSTATUSBCOCONTARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarStatusBcoContaReceber(4)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function



    Public Function AlteraNossoNumeroContaCorrente(ByVal strNossoNumero As Integer,
                                                   ByVal strNumCta As String,
                                                   ByVal strCodBanco As String,
                                                   ByVal Connection As SqlConnection,
                                                   ByVal Transaction As SqlTransaction) As Retorno '#5#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlteraNossoNumeroContaCorrente", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NossoNumero", strNossoNumero))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_ALTERANOSSONUMEROCONTACORRENTE.Id
            _Retorno.MsgErro = EXCEPTION_ALTERANOSSONUMEROCONTACORRENTE.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlteraNossoNumeroContaCorrente(5)", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            Connection.Close()
        End Try

        Return _Retorno

    End Function


    Public Function AlterarContaReceber(ByVal _ContasReceber As ContaReceber,
                                          ByVal connection As SqlConnection,
                                          ByVal Transacao As SqlTransaction) As Retorno '#6#
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarContaReceber", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit     ", _ContasReceber.NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit     ", _ContasReceber.SeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao  ", _ContasReceber.DtEmissao))
            Command.Parameters.Add(New SqlParameter("@Situacao   ", _ContasReceber.Situacao))
            Command.Parameters.Add(New SqlParameter("@CodClie    ", _ContasReceber.CodClie))
            Command.Parameters.Add(New SqlParameter("@DtVcto     ", _ContasReceber.DtVcto))
            Command.Parameters.Add(New SqlParameter("@CodInd     ", _ContasReceber.CodInd))
            Command.Parameters.Add(New SqlParameter("@VlrInd     ", _ContasReceber.VlrInd))
            Command.Parameters.Add(New SqlParameter("@CodBanco   ", IIf(IsNothing(_ContasReceber.CodBanco), DBNull.Value, _ContasReceber.CodBanco)))
            Command.Parameters.Add(New SqlParameter("@CodAgen    ", IIf(IsNothing(_ContasReceber.CodAgen), DBNull.Value, _ContasReceber.CodAgen)))
            Command.Parameters.Add(New SqlParameter("@NumCta     ", IIf(IsNothing(_ContasReceber.NumCta), DBNull.Value, _ContasReceber.NumCta)))
            Command.Parameters.Add(New SqlParameter("@NumPortador", _ContasReceber.NumPortador))
            Command.Parameters.Add(New SqlParameter("@ObsTit     ", _ContasReceber.ObsTit))
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
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARCONTARECEBER.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARCONTARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarContaReceber(6)", "", "VERISURE", Environment.MachineName, "", "")

        End Try

        Return _Retorno
    End Function

    Public Function AlteraNossoNumeroBcoContasAReceber(ByVal GravaDDA As Boolean,
                                                       ByVal Connection As SqlConnection,
                                                       ByVal Transaction As SqlTransaction) As Retorno '#7#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlteraNossoNumeroBcoContasAReceber", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@GravaDDA", IIf(GravaDDA, 1, 0)))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERANOSSONUMEROBCOCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERANOSSONUMEROBCOCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlteraNossoNumeroBcoContasAReceber(7)", "", "VERISURE", Environment.MachineName, "", "")

        End Try

        Return _Retorno

    End Function



    Public Function AlterarDtTitDescTitDesContasReceber(ByVal strNumTit As String,
                                                        ByVal strSeqTit As String,
                                                        ByVal DtEmissao As DateTime,
                                                        ByVal strTitDesc As String,
                                                        ByVal Connection As SqlConnection,
                                                        ByVal Transaction As SqlTransaction) As Retorno '#8#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarDtTitDescTitDesContasReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit     ", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit     ", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao  ", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@TitDesc   ", strTitDesc))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARDTTITDESCTITDESCONTASRECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARDTTITDESCTITDESCONTASRECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarDtTitDescTitDesContasReceber(8)", "", "VERISURE", Environment.MachineName, "", "")

        End Try

        Return _Retorno

    End Function


    Public Function AlteraNossoNumeroBcoObsTitContasAReceber(ByVal Connection As SqlConnection,
                                                             ByVal Transaction As SqlTransaction) As Retorno '#9#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlteraNossoNumeroBcoObsTitContasAReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERANOSSONUMEROBCOCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERANOSSONUMEROBCOCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlteraNossoNumeroBcoContasAReceber(7)", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            Connection.Close()
        End Try

        Return _Retorno

    End Function



    Public Function AlterarDtUltCobrancaStatusCobContasaReceberPorNumTitSeqTitDtEmissao(ByVal strNumTit As String,
                                                                                        ByVal strSeqTit As String,
                                                                                        ByVal DtEmissao As DateTime,
                                                                                        ByVal strStatusCob As String,
                                                                                        ByVal Connection As SqlConnection,
                                                                                        ByVal Transaction As SqlTransaction) As Retorno '#10#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarDtUltCobrancaStatusCobContasaReceberPorNumTitSeqTitDtEmissao", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@StatusCob", strStatusCob))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARDTULTCOBRANCASTATUSCOBCONTASARECEBERPORNUMTITSEQTITDTEMISSAO.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARDTULTCOBRANCASTATUSCOBCONTASARECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarDtUltCobrancaStatusCobContasaReceberPorNumTitSeqTitDtEmissao(10)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function





    Public Function AlterarDtEnvCobExtContaAReceberPornumTitSeqTitDtEmissao(ByVal strNumTit As String,
                                                                            ByVal strSeqTit As String,
                                                                            ByVal DtEmissao As DateTime,
                                                                            ByVal Connection As SqlConnection,
                                                                            ByVal Transaction As SqlTransaction) As Retorno '#11#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarDtEnvCobExtContaAReceberPornumTitSeqTitDtEmissao", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARDTENVCOBEXTCONTAARECEBERPORNUMTITSEQTITDTEMISSAO.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARDTENVCOBEXTCONTAARECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarDtEnvCobExtContaAReceberPornumTitSeqTitDtEmissao(11)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function



    Public Function AlterarDtIncSerasaContasAReceber(ByVal strNumTit As String,
                                                     ByVal strSeqTit As String,
                                                     ByVal DtEmissao As DateTime,
                                                     ByVal strEmpresaNegativacao As String,
                                                     ByVal Connection As SqlConnection,
                                                     ByVal Transaction As SqlTransaction) As Retorno '#12#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarDtIncSerasaContasAReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@EmpresaNegativacao", strEmpresaNegativacao))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARDTINCSERASACONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARDTINCSERASACONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarDtIncSerasaContasAReceber(12)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function



    Public Function AlterarDtExcSerasaContasAReceber(ByVal strNumTit As String,
                                                     ByVal strSeqTit As String,
                                                     ByVal DtEmissao As DateTime,
                                                     ByVal Connection As SqlConnection,
                                                     ByVal Transaction As SqlTransaction) As Retorno '#13#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarDtExcSerasaContasAReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARDTINCSERASACONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARDTINCSERASACONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarDtIncSerasaContasAReceber(13)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function


    Public Function AlterarDtExcSerasaCodMotExcContasAReceber(ByVal strCodMotExc As String,
                                                              ByVal strNumTit As String,
                                                              ByVal strSeqTit As String,
                                                              ByVal DtEmissao As DateTime,
                                                              ByVal Connection As SqlConnection,
                                                              ByVal Transaction As SqlTransaction) As Retorno '#14#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarDtExcSerasaCodMotExcContasAReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodMotExc", strCodMotExc))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARDTEXCSERASACODMOTEXCCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARDTEXCSERASACODMOTEXCCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarDtExcSerasaCodMotExcContasAReceber(14)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function





    Public Function AlterarBancoPortadorContasAReceber(ByVal strCodBanco As String,
                                                         ByVal strCodAgen As String,
                                                         ByVal strNumCta As String,
                                                         ByVal strNumPortador As String,
                                                         ByVal strNumTit As String,
                                                         ByVal strSeqTit As String,
                                                         ByVal DtEmissao As DateTime) As Retorno '#15#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarBancoPortadorContasAReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@NumPortador", strNumPortador))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))

            connection.Open()

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARDTEXCSERASACODMOTEXCCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARDTEXCSERASACODMOTEXCCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarBancoPortadorContasAReceber(15)", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            connection.Close()
        End Try

        Return _Retorno

    End Function



    Public Function AlterarSituacaoContaAReceber(ByVal strSituacao As String,
                                                 ByVal strNumTit As String,
                                                 ByVal strSeqTit As String,
                                                 ByVal DtEmissao As DateTime,
                                                 ByVal Connection As SqlConnection,
                                                 ByVal Transaction As SqlTransaction) As Retorno '#16#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarSituacaoContaAReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Situacao", strSituacao))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARSITUACAOCONTAARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARSITUACAOCONTAARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarSituacaoContaAReceber(16)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function



    Public Function AlterarSitContasReceber(ByVal strSituacao As String,
                                            ByVal strNumTit As String,
                                            ByVal strSeqTit As String,
                                            ByVal DtEmissao As DateTime,
                                            ByVal strCodClie As String,
                                            ByVal strIdFilial As String,
                                            ByVal Connection As SqlConnection,
                                            ByVal Transaction As SqlTransaction) As Retorno '#17#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarSitContasReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Situacao", strSituacao))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@CodClie", strCodClie))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", strIdFilial))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARSITCONTASRECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARSITCONTASRECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarSitContasReceber(17)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function


    Public Function AlteraTitDescDtTitDescContasReceberPorNumTitSeqTitDtEmissao(ByVal strTitDesc As String,
                                                                                ByVal DtTitDesc As DateTime,
                                                                                ByVal strNumTit As String,
                                                                                ByVal strSeqTit As String,
                                                                                ByVal DtEmissao As DateTime,
                                                                                ByVal connection As SqlConnection,
                                                                                ByVal Transaction As SqlTransaction) As Retorno '#17#

        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlteraTitDescDtTitDescContasReceberPorNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@TitDesc", strTitDesc))
            Command.Parameters.Add(New SqlParameter("@DtTitDesc", IIf(DtTitDesc <= Date.MinValue, DBNull.Value, DtTitDesc)))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERATITDESCDTTITDESCCONTASRECEBERPORNUMTITSEQTITDTEMISSAO.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERATITDESCDTTITDESCCONTASRECEBERPORNUMTITSEQTITDTEMISSAO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlteraTitDescDtTitDescContasReceberPorNumTitSeqTitDtEmissao(18)", "", "VERISURE", Environment.MachineName, "", "")

        End Try

        Return _Retorno

    End Function


    Public Function AtualizarStatusSituacaoContasAReceber(ByVal strSituacao As String,
                                                          ByVal strNumTit As String,
                                                          ByVal strSeqTit As String,
                                                          ByVal DtEmissao As DateTime,
                                                          ByVal connection As SqlConnection,
                                                          ByVal Transaction As SqlTransaction) As Retorno '#17#

        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AtualizarStatusSituacaoContasAReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@SITUACAO", strSituacao))
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AtualizarStatusSituacaoContasAReceber(19)", "", "VERISURE", Environment.MachineName, "", "")

        End Try

        Return _Retorno

    End Function

    Public Function GerarArquivoBoletos(ByVal DtIni As DateTime,
                                        ByVal DtFim As DateTime,
                                        ByVal CodBanco As String,
                                        ByVal Usr As String,
                                        Optional ByVal idBu As Integer = 0,
                                        Optional ByVal email As String = "",
                                        Optional ByVal NumTitMultaComodato As String = "",
                                        Optional ByVal NumTitMultaFinanciamento As String = "") As DataTable
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim dtRetorno As New DataTable
        Try
            connection.Open()

            Dim command As SqlCommand = New SqlCommand("P_GerarArquivoBoletos_2", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Faturamento

            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New SqlParameter("@DtIni", DtIni))
            command.Parameters.Add(New SqlParameter("@DtFim", DtFim))
            command.Parameters.Add(New SqlParameter("@CodBanco", CodBanco))
            command.Parameters.Add(New SqlParameter("@Usr", Usr))
            command.Parameters.Add(New SqlParameter("@IdBU", idBu))
            If email.Trim() <> "" Then
                command.Parameters.Add(New SqlParameter("@Email", email))
            Else
                command.Parameters.Add(New SqlParameter("@Email", DBNull.Value))
            End If
            command.Parameters.Add(New SqlParameter("@NumTitMultaComodato", NumTitMultaComodato))
            command.Parameters.Add(New SqlParameter("@NumTitMultaFinanciamento", NumTitMultaFinanciamento))

            rdr = command.ExecuteReader

            If rdr.HasRows Then dtRetorno.Load(rdr)

            rdr.Close()
        Catch ex As Exception
            Throw ex
        Finally
            If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return dtRetorno
    End Function

    Public Function GerarArquivoBoletosRegistroItau(ByVal DtIni As DateTime,
                                                    ByVal DtFim As DateTime,
                                                    ByVal CodBanco As String,
                                                    ByVal Usr As String,
                                                    Optional ByVal idBu As Integer = 0,
                                                    Optional ByVal email As String = "",
                                                    Optional ByVal NumTitMultaComodato As String = "",
                                                    Optional ByVal NumTitMultaFinanciamento As String = "",
                                                    Optional ByVal IdLayout As Integer = 0) As DataTable
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim dtRetorno As New DataTable
        Try
            connection.Open()

            Dim command As SqlCommand = New SqlCommand("P_GerarArquivoBoletosRegItau", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Faturamento

            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New SqlParameter("@DtIni", DtIni))
            command.Parameters.Add(New SqlParameter("@DtFim", DtFim))
            command.Parameters.Add(New SqlParameter("@CodBanco", CodBanco))
            command.Parameters.Add(New SqlParameter("@Usr", Usr))
            command.Parameters.Add(New SqlParameter("@IdBU", idBu))
            If email.Trim() <> "" Then
                command.Parameters.Add(New SqlParameter("@Email", email))
            Else
                command.Parameters.Add(New SqlParameter("@Email", DBNull.Value))
            End If
            command.Parameters.Add(New SqlParameter("@NumTitMultaComodato", NumTitMultaComodato))
            command.Parameters.Add(New SqlParameter("@NumTitMultaFinanciamento", NumTitMultaFinanciamento))
            command.Parameters.Add(New SqlParameter("@IdLayout", IdLayout))

            rdr = command.ExecuteReader

            If rdr.HasRows Then dtRetorno.Load(rdr)

            rdr.Close()
        Catch ex As Exception
            Throw ex
        Finally
            If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return dtRetorno
    End Function

    Public Function GerarBoletagem(ByVal DtVencimento As DateTime,
                                   ByVal Usr As String,
                                   ByVal conn As SqlConnection,
                                   ByVal trans As SqlTransaction) As DataTable
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim dtRetorno As New DataTable
        Try

            Dim command As SqlCommand = New SqlCommand("P_GeraBoletagem", conn)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Boletagem
            command.Transaction = trans

            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add(New SqlParameter("@DtVencimento", DtVencimento))
            command.Parameters.Add(New SqlParameter("@Usr", Usr))

            rdr = command.ExecuteReader

            If rdr.HasRows Then dtRetorno.Load(rdr)

            rdr.Close()
        Catch ex As Exception
            Throw ex
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return dtRetorno
    End Function

    Public Function InsereClienteBoletagem(ByVal CodIntClie As String, ByVal Valor As Double, ByVal Desconto As Integer, ByVal Usr As String, ByVal DtVcto As Date, ByVal CodClie As String, ByVal CodReclamacao As String, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereClienteBoletagem", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@Desconto", Desconto))
            Command.Parameters.Add(New SqlParameter("@Valor", Valor))
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@DtVcto", DtVcto))
            Command.Parameters.Add(New SqlParameter("@CodClie", CodClie))
            Command.Parameters.Add(New SqlParameter("@CodReclamacao", IIf(String.IsNullOrEmpty(CodReclamacao), DBNull.Value, CodReclamacao)))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: InsereClienteBoletagem", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function


    Public Function InsereClienteBoletagemDev(ByVal CodIntClie As String, ByVal Valor As Double, ByVal Desconto As Integer, ByVal Usr As String, ByVal DtVcto As Date, ByVal CodClie As String, ByVal CodReclamacao As String, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereClienteBoletagemDev", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@Desconto", Desconto))
            Command.Parameters.Add(New SqlParameter("@Valor", Valor))
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@DtVcto", DtVcto))
            Command.Parameters.Add(New SqlParameter("@CodClie", CodClie))
            Command.Parameters.Add(New SqlParameter("@CodReclamacao", IIf(String.IsNullOrEmpty(CodReclamacao), DBNull.Value, CodReclamacao)))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: InsereClienteBoletagem", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function


    Public Function InsereClienteBoletagemAdyen(ByVal CodIntClie As String, ByVal Valor As Double, ByVal Desconto As Integer, ByVal Usr As String, ByVal DtVcto As Date, ByVal CodClie As String, ByVal CodReclamacao As String, ByVal tipoBoletagem As String, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereClienteBoletagemAdyen", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@Desconto", Desconto))
            Command.Parameters.Add(New SqlParameter("@Valor", Valor))
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@DtVcto", DtVcto))
            Command.Parameters.Add(New SqlParameter("@CodClie", CodClie))
            Command.Parameters.Add(New SqlParameter("@CodReclamacao", IIf(String.IsNullOrEmpty(CodReclamacao), DBNull.Value, CodReclamacao)))
            Command.Parameters.Add(New SqlParameter("@tipoBoletagem", tipoBoletagem))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: InsereClienteBoletagem", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function

    Public Function ClientesBoletagemReenvioRentSoft(ByVal CodIntClie As String, ByVal Valor As Double, ByVal Desconto As Integer, ByVal Usr As String, ByVal DtVcto As Date, ByVal CodClie As String, ByVal CodReclamacao As String, ByVal tipoBoletagem As String, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_ClientesBoletagemReenvioRentSoft", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@Desconto", Desconto))
            Command.Parameters.Add(New SqlParameter("@Valor", Valor))
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@DtVcto", DtVcto))
            Command.Parameters.Add(New SqlParameter("@CodClie", CodClie))
            Command.Parameters.Add(New SqlParameter("@CodReclamacao", IIf(String.IsNullOrEmpty(CodReclamacao), DBNull.Value, CodReclamacao)))
            Command.Parameters.Add(New SqlParameter("@tipoBoletagem", tipoBoletagem))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: InsereClienteBoletagem", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function



    Public Function RetiraClienteFilaCobrancaBU(ByVal Usr As String, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_RetiraClienteFilaCobrancaBU", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: RetiraClienteFilaCobrancaBU", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function


    Public Function RetiraClienteFilaCobrancaFila451833(ByVal Usr As String, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_RetiraClienteFilaCobrancaFila451833", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: RetiraClienteFilaCobrancaBU", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function




    Public Function DeletaClientesBoletagem() As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            connection.Open()

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_DeletaClientesBoletagem", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: DeletaClientesBoletagem", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function

    Public Function CancelarTituloContaReceber(ByVal numTit As String,
                                               ByVal seqTit As String,
                                               ByVal dtEmissao As DateTime,
                                               ByVal idFilial As String,
                                               ByVal usr As String) As Retorno '#3#'
        Dim _Retorno = New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Try
            connection.Open()
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_CancelarTituloContaReceber", connection)

            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", numTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", seqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", idFilial))
            Command.Parameters.Add(New SqlParameter("@Usr", usr))




            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_CANCELARTITULOCONTARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_CANCELARTITULOCONTARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            If connection.State <> ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno

    End Function

    Public Function CancelarBUTitulosBaixados(ByVal numTit As String,
                                              ByVal seqTit As String,
                                              ByVal connection As SqlConnection,
                                              ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_CancelarBUTitulosBaixados", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Try
            command.Transaction = trans
            command.Parameters.Add(New SqlParameter("@NumTit", IIf(String.IsNullOrEmpty(numTit), DBNull.Value, numTit)))
            command.Parameters.Add(New SqlParameter("@SeqTit", IIf(String.IsNullOrEmpty(seqTit), DBNull.Value, seqTit)))
            command.Parameters.Add(New SqlParameter("@Usuario", "VERISURE"))


            command.ExecuteNonQuery()
            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_CANCELARBUTITULOSBAIXADOS.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_CANCELARBUTITULOSBAIXADOS.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function AlterarSituacaoContasAReceber(ByVal strNumTit As String,
                                                  ByVal strSeqTit As String,
                                                  ByVal DtEmissao As DateTime,
                                                  ByVal IdFilial As String,
                                                  ByVal Connection As SqlConnection,
                                                  ByVal Transaction As SqlTransaction) As Retorno '#16#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarSituacaoContasAReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", IdFilial))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERARSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERARSITUACAOCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function

    Public Function CancelarTituloContaReceber(ByVal numTit As String,
                                              ByVal seqTit As String,
                                              ByVal dtEmissao As DateTime,
                                              ByVal idFilial As String,
                                              ByVal usr As String,
                                              ByVal connection As SqlConnection,
                                              ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_CancelarTituloContaReceber", connection)

            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", numTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", seqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))
            Command.Parameters.Add(New SqlParameter("@IdFilial", idFilial))
            Command.Parameters.Add(New SqlParameter("@Usr", usr))


            Command.Transaction = trans

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_CANCELARTITULOCONTARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_CANCELARTITULOCONTARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function

    Public Function RetirarValorFinanciadoMonitoriasAbertas(ByVal codIntClie As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_RetirarValorFinanciadoMonitoriasAbertas", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Try
            command.Transaction = trans
            command.Parameters.Add(New SqlParameter("@CodIntClie", IIf(String.IsNullOrEmpty(codIntClie), DBNull.Value, codIntClie)))


            command.ExecuteNonQuery()
            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_RETIRARVALORFINANCIADOMONITORIASABERTAS.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_RETIRARVALORFINANCIADOMONITORIASABERTAS.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function MarcarValorFinanciadoMonitoriasAbertas(ByVal codIntClie As String, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_MarcarValorFinanciadoMonitoriasAbertas", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Try
            command.Transaction = trans
            command.Parameters.Add(New SqlParameter("@CodIntClie", IIf(String.IsNullOrEmpty(codIntClie), DBNull.Value, codIntClie)))

            command.ExecuteNonQuery()
            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_MARCARVALORFINANCIADOMONITORIASABERTAS.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_MARCARVALORFINANCIADOMONITORIASABERTAS.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function CancelaBU(ByVal Usr As String, ByVal idBU As Integer, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_CancelaBU", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@idBU", idBU))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_CANCELATITULOSCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: CancelaBU", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function

    Public Function CancelaBU(ByVal Usr As String, ByVal idBU As Integer) As Retorno
        Dim _Retorno = New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_CancelaBU", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@idBU", idBU))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_CANCELATITULOSCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: CancelaBU", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            connection.Close()
        End Try

        Return _Retorno

    End Function

    Public Function AtualizaBoleto(ByVal idBU As Integer, ByVal NossoNumero As String, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AtualizaBoleto", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@idBU", idBU))
            Command.Parameters.Add(New SqlParameter("@NossoNumero", NossoNumero))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_CANCELATITULOSCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AtualizaBoleto", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function

    Public Function ValidaDataGeracaoBUs(ByVal DtBoletagem As Date) As ValidaDataGeracaoBUs
        Dim vdgbu_tb As New ValidaDataGeracaoBUs
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_ValidaDataGeracaoBUs", connection)
        command.CommandType = CommandType.StoredProcedure
        Dim rdr As SqlDataReader

        Try
            connection.Open()
            command.Parameters.Add(New SqlParameter("@DtBoletagem", DtBoletagem))
            'Os parâmetros abaixo são do tipo output, não preciso passar pois irão retornar via select no reader.
            'São utilizados como output apenas dentro do BD.
            command.Parameters.Add(New SqlParameter("@Result", DBNull.Value))

            rdr = command.ExecuteReader()
            If rdr.HasRows Then
                rdr.Read()

                vdgbu_tb.DtResult = Convert.ToDateTime(rdr("DtResult"))
            Else
                vdgbu_tb.Sucesso = False
                vdgbu_tb.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
        Catch ex As Exception
            vdgbu_tb.Sucesso = False
            vdgbu_tb.NumErro = ErrorConstants.EXCEPTION_METODO_VALIDADATAGERACAOBUS.Id
            vdgbu_tb.MsgErro = ErrorConstants.EXCEPTION_METODO_VALIDADATAGERACAOBUS.Descricao & ex.Message
            vdgbu_tb.TipoErro = DadosGenericos.TipoErro.Arquitetura
            vdgbu_tb.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(vdgbu_tb.NumErro, vdgbu_tb.MsgErro, vdgbu_tb.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            If connection.State <> ConnectionState.Closed Then connection.Close()
            connection.Dispose()
            If Not (IsNothing(rdr)) Then rdr.Close()
        End Try

        Return vdgbu_tb
    End Function

    Public Function CancelaMUFI(ByVal Usr As String, ByVal NumTit As String, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_Cancela_MU_FI", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@NumTit", NumTit))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_CANCELATITULOSCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function

    Public Function AtualizaBoletoMUFI(ByVal NumTit As String, ByVal NossoNumero As String, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AtualizaBoleto_MU_FI", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", NumTit))
            Command.Parameters.Add(New SqlParameter("@NossoNumero", NossoNumero))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_CANCELATITULOSCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function

    Public Function AtualizarObsBoletoUnificado(idBu As Integer, connection As SqlConnection, trans As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AtualizarObsBoletoUnificado", connection)
            Command.Transaction = trans
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@IdBU", IIf(String.IsNullOrEmpty(idBu), DBNull.Value, idBu)))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZAROBSBOLETOUNIFICADO.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ATUALIZAROBSBOLETOUNIFICADO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno
    End Function

    Public Function InsereClientesTituloBoletagem(ByVal isApenasMonitoria As Boolean, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereClientesTituloBoletagem", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            'Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@IsApenasMonitoria", IIf(isApenasMonitoria, 1, 0)))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: InsereClienteBoletagem", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function



    Public Function InsereClientesTituloBoletagemDevDetalhe(ByVal isApenasMonitoria As Boolean, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereClientesTituloBoletagemDev", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            'Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@IsApenasMonitoria", IIf(isApenasMonitoria, 1, 0)))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: InsereClienteBoletagem", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function





    Public Function AlterarCodReclamacaoJustificativaContasReceber(ByVal CodReclamacaoJustificativa As String, ByVal idBU As Integer, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try

            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarCodReclamacaoJustificativaContasReceber", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            'Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@CodReclamacaoJustificativa", IIf(String.IsNullOrEmpty(CodReclamacaoJustificativa), DBNull.Value, CodReclamacaoJustificativa)))
            Command.Parameters.Add(New SqlParameter("@IdBU", idBU))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ATUALIZARSTATUSSITUACAOCONTASARECEBER.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: InsereClienteBoletagem", "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function

    Public Function AtualizaNossoNumeroBco(ByVal nossoNumero As String,
                                           ByVal idBU As Integer,
                                           ByVal CodBanco As String,
                                           ByVal CodAgencia As String,
                                           ByVal NumCta As String,
                                                       ByVal Connection As SqlConnection,
                                                       ByVal Transaction As SqlTransaction) As Retorno '#7#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereNossoNumeroCalculadoDigitoBCO", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@idBU", idBU))
            Command.Parameters.Add(New SqlParameter("@CodBanco", CodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgencia", CodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", NumCta))
            Command.Parameters.Add(New SqlParameter("@NossoNumeroBCO", nossoNumero))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERANOSSONUMEROBCOCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERANOSSONUMEROBCOCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlteraNossoNumeroBcoContasAReceber(7)", "", "VERISURE", Environment.MachineName, "", "")

        End Try

        Return _Retorno

    End Function

    Public Function AtualizaNossoNumeroBco_Multas(ByVal nossoNumero As String,
                                           ByVal Numtit As String,
                                           ByVal CodBanco As String,
                                           ByVal CodAgencia As String,
                                           ByVal NumCta As String,
                                                       ByVal Connection As SqlConnection,
                                                       ByVal Transaction As SqlTransaction) As Retorno '#7#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereNossoNumeroCalculadoDigitoBCO_Multas", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Parameters.Add(New SqlParameter("@NumTit", Numtit))
            Command.Parameters.Add(New SqlParameter("@CodBanco", CodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgencia", CodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", NumCta))
            Command.Parameters.Add(New SqlParameter("@NossoNumeroBCO", nossoNumero))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_ALTERANOSSONUMEROBCOCONTASARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_METODO_ALTERANOSSONUMEROBCOCONTASARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlteraNossoNumeroBcoContasAReceber(7)", "", "VERISURE", Environment.MachineName, "", "")

        End Try

        Return _Retorno

    End Function

    Public Function InserirTitulosAgendadosAdyen(ByVal strNumTit As String,
                                                 ByVal strSeqTit As String,
                                                 ByVal DtEmissao As DateTime,
                                                 ByVal DtAgendamento As DateTime,
                                                 ByVal Usr As String,
                                                 ByVal idAgendamento As Integer,
                                                 ByVal Codintclie As String,
                                                 ByVal Connection As SqlConnection,
                                                 ByVal Transaction As SqlTransaction) As Retorno '#4#
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulosAgendadosAdyen", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", DtEmissao))
            Command.Parameters.Add(New SqlParameter("@DtAgendamento", DtAgendamento))
            Command.Parameters.Add(New SqlParameter("@Codintclie", Codintclie))
            Command.Parameters.Add(New SqlParameter("@Usr", Usr))
            Command.Parameters.Add(New SqlParameter("@IdAgendamento", idAgendamento))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_ALTERARSTATUSBCOCONTARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_ALTERARSTATUSBCOCONTARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarStatusBcoContaReceber(4)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function


    Public Function InserirAgendamentoAdyen(ByVal dtAgendamento As DateTime,
                                                 ByVal hrAgendamento As String,
                                                 ByVal usrAgendamento As String,
                                                 ByRef id As Integer,
                                                 ByVal Connection As SqlConnection,
                                                 ByVal Transaction As SqlTransaction) As Retorno '#4#
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereAgendamentoAdyen", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@DtAgendamento", dtAgendamento))
            Command.Parameters.Add(New SqlParameter("@HrAgendamento", hrAgendamento))
            Command.Parameters.Add(New SqlParameter("@Usr", usrAgendamento))

            ''Executa a procedure
            id = Convert.ToInt32(Command.ExecuteScalar())

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_ALTERARSTATUSBCOCONTARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_ALTERARSTATUSBCOCONTARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ContasAReceberBC - Classe: ContasAReceberAlterar - Função: AlterarStatusBcoContaReceber(4)", "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function

    Public Function InsertTempStatusBco(NumTit As String, SeqTit As String, DtEmissao As Date?, CodBco As String, CodAgen As String, NumCta As String, TipoPagto As String, BancoAgen As String, IdFilial As String) As Retorno '#4#
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_InsertTempStatusBco", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@NumTit", IIf(String.IsNullOrEmpty(NumTit), DBNull.Value, NumTit)))
            command.Parameters.Add(New SqlParameter("@SeqTit", IIf(String.IsNullOrEmpty(SeqTit), DBNull.Value, SeqTit)))
            command.Parameters.Add(New SqlParameter("@DtEmissao", IIf(String.IsNullOrEmpty(DtEmissao), DBNull.Value, DtEmissao)))
            command.Parameters.Add(New SqlParameter("@CodBco", IIf(String.IsNullOrEmpty(CodBco), DBNull.Value, CodBco)))
            command.Parameters.Add(New SqlParameter("@CodAgen", IIf(String.IsNullOrEmpty(CodAgen), DBNull.Value, CodAgen)))
            command.Parameters.Add(New SqlParameter("@NumCta", IIf(String.IsNullOrEmpty(NumCta), DBNull.Value, NumCta)))
            command.Parameters.Add(New SqlParameter("@TipoPagto", IIf(String.IsNullOrEmpty(TipoPagto), DBNull.Value, TipoPagto)))
            command.Parameters.Add(New SqlParameter("@BancoAgen", IIf(String.IsNullOrEmpty(BancoAgen), DBNull.Value, BancoAgen)))
            command.Parameters.Add(New SqlParameter("@IdFilial", IIf(String.IsNullOrEmpty(IdFilial), DBNull.Value, IdFilial)))

            ''Executa a procedure
            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_INSERTTEMPSTATUSBCO.Id
            _Retorno.MsgErro = EXCEPTION_INSERTTEMPSTATUSBCO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function UpdateStatusBcoContaReceber(ByVal Connection As SqlConnection, ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_UpdateStatusBcoContaReceber", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_ALTERARSTATUSBCOCONTARECEBER.Id
            _Retorno.MsgErro = EXCEPTION_ALTERARSTATUSBCOCONTARECEBER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno
    End Function

    Public Function AlteraNossoNumeroContaCorrente(ByVal strNossoNumero As Integer, ByVal strNumCta As String, ByVal strCodBanco As String) As Retorno '#5#
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_AlteraNossoNumeroContaCorrente", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@NossoNumero", strNossoNumero))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))

            ''Executa a procedure
            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_ALTERANOSSONUMEROCONTACORRENTE.Id
            _Retorno.MsgErro = EXCEPTION_ALTERANOSSONUMEROCONTACORRENTE.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function InsertTempTitulosEnvBco(NumTit As String, SeqTit As String, DtEmissao As Date?, CodBco As String, CodAgen As String, NumCta As String, TipoPagto As String, BancoAgen As String, IdFilial As String) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_InsertTempTitulosEnvBco", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@NumTit", IIf(String.IsNullOrEmpty(NumTit), DBNull.Value, NumTit)))
            command.Parameters.Add(New SqlParameter("@SeqTit", IIf(String.IsNullOrEmpty(SeqTit), DBNull.Value, SeqTit)))
            command.Parameters.Add(New SqlParameter("@DtEmissao", IIf(String.IsNullOrEmpty(DtEmissao), DBNull.Value, DtEmissao)))
            command.Parameters.Add(New SqlParameter("@CodBco", IIf(String.IsNullOrEmpty(CodBco), DBNull.Value, CodBco)))
            command.Parameters.Add(New SqlParameter("@CodAgen", IIf(String.IsNullOrEmpty(CodAgen), DBNull.Value, CodAgen)))
            command.Parameters.Add(New SqlParameter("@NumCta", IIf(String.IsNullOrEmpty(NumCta), DBNull.Value, NumCta)))
            command.Parameters.Add(New SqlParameter("@TipoPagto", IIf(String.IsNullOrEmpty(TipoPagto), DBNull.Value, TipoPagto)))
            command.Parameters.Add(New SqlParameter("@BancoAgen", IIf(String.IsNullOrEmpty(BancoAgen), DBNull.Value, BancoAgen)))
            command.Parameters.Add(New SqlParameter("@IdFilial", IIf(String.IsNullOrEmpty(IdFilial), DBNull.Value, IdFilial)))

            ''Executa a procedure
            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_ALTERANOSSONUMEROCONTACORRENTE.Id
            _Retorno.MsgErro = EXCEPTION_ALTERANOSSONUMEROCONTACORRENTE.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function InsertTitulosEnvBco(connection As SqlConnection, transaction As SqlTransaction) As Retorno
        Dim _Retorno = New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsertTitulosEnvBco", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = transaction

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_INSERTTITULOSENVBCO.Id
            _Retorno.MsgErro = EXCEPTION_INSERTTITULOSENVBCO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno
    End Function

    Public Function UpdateNewPaymentMethod(CodClie As String, conn As SqlConnection, trans As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarDadosBancariosNoTitulo", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodClie", CodClie))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_UPDATENEWPAYMENTMETHOD.Id
            _Retorno.MsgErro = EXCEPTION_UPDATENEWPAYMENTMETHOD.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try

        Return _Retorno
    End Function

    Public Function InsereClientesCanceladosClientesTituloBoletagem(isApenasMonitoria As Boolean, conn As SqlConnection, trans As SqlTransaction) As Retorno
        'Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno = New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereClientesCanceladosClientesTituloBoletagem", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Faturamento
            Command.Transaction = trans

            ''define os parƒmetros usados na stored procedure
            'Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Command.Parameters.Add(New SqlParameter("@IsApenasMonitoria", IIf(isApenasMonitoria, 1, 0)))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = EXCEPTION_METODO_INSERECLIENTESCANCELADOSCLIENTESTITULOBOLETAGEM.Id
            _Retorno.MsgErro = ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'If Not connection.State = ConnectionState.Closed Then connection.Close()
        End Try

        Return _Retorno
    End Function


    Public Function ConsultaseExiste(ByVal CodIntClie As String) As Retorno '#12#
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_VerificaUnidade", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))
            Dim existeParam As SqlParameter = Command.Parameters.Add("@Existe", SqlDbType.Bit)
            existeParam.Direction = ParameterDirection.Output
            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            Command.ExecuteNonQuery()

            Dim existe As Boolean = Convert.ToBoolean(existeParam.Value)

            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

            '_Retorno.Sucesso = True
            If Not existe Then
                _Retorno.Sucesso = False
            Else
                _Retorno.Sucesso = True
            End If

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return _Retorno
    End Function

End Class
