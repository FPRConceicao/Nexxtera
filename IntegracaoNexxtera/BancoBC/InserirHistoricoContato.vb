Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection
Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Entidades


Public Class InserirHistoricoContato
    Public Function IncluiHistoricoContato(strCodIntClie As String, strUsuario As String, strMensagem As String, connection As SqlConnection, Transacao As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InserirHistoricoContato", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", strCodIntClie))
            Command.Parameters.Add(New SqlParameter("@Usuario", strUsuario))
            Command.Parameters.Add(New SqlParameter("@Mensagem", strMensagem))
            Command.Parameters.Add(New SqlParameter("@TipoContato", ""))
            'Command.Parameters.Add(New SqlParameter("@DataHora", Funcoes.PegaData))
            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Id
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura


            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: HistoricoContatoBC - Classe: InsereHistoricoContato - Função: IncluiHistoricoContato(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function
    ''' <summary>
    ''' Insere dados na tabela tHist_Contato
    ''' </summary>
    ''' <param name="_historicoContato"> parametro do tipo HistoricoContato</param>
    ''' <returns>retorna uma varia vel do tipo retorno que ira informa r se a operação foi executado com sucesso</returns>
    ''' <remarks>
    ''' Data Criação:    14/08/2011
    ''' Autor:           Edson Ferreira
    ''' </remarks>
    Public Function InsereHistoricoContato(ByVal _historicoContato As HistoricoContato) As Retorno

        Dim _retorno As New Retorno

        Using connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

            Try
                Dim Command As SqlCommand = New SqlCommand("P_InserirHistoricoContato", connection)
                Command.CommandType = CommandType.StoredProcedure
                Command.CommandTimeout = DadosGenericos.Timeout.Query

                'define os parametros usados na stored procedure
                Command.Parameters.Add(New SqlParameter("@CodIntClie", _historicoContato.CodIntClie))
                'Command.Parameters.Add(New SqlParameter("@DataHora", _historicoContato.DataHora))
                Command.Parameters.Add(New SqlParameter("@Mensagem", _historicoContato.Mensagem))
                Command.Parameters.Add(New SqlParameter("@Usuario", _historicoContato.Usuario))
                Command.Parameters.Add(New SqlParameter("@TipoContato", _historicoContato.TipoContato))

                connection.Open()

                Command.ExecuteNonQuery()

                _retorno.Sucesso = True
                _retorno.TipoErro = DadosGenericos.TipoErro.None

                Return _retorno
            Catch ex As Exception

                _retorno.Sucesso = False
                _retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Id
                _retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Descricao & ex.Message
                _retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
                _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

                'CRIAR LOG NO WINDOWS
                Funcoes.AtualizaApplEventLog(_retorno.NumErro, _retorno.MsgErro, _retorno.TipoErro, "Projeto: HistoricoContatoBC - Classe: InsereHistoricoContato - Função: InsereHistoricoContato(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")

                Return _retorno
            End Try


        End Using

    End Function

    Public Function InsereHistoricoContatoMotivos(ByVal _historicoContato As HistoricoContato, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As Retorno

        Dim _retorno As New Retorno

        Try
            Dim Command As SqlCommand = New SqlCommand("P_InsereHistoricoContatoMotivos", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            Command.Transaction = trans

            'define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", _historicoContato.CodIntClie))
            'Command.Parameters.Add(New SqlParameter("@DataHora", _historicoContato.DataHora))
            Command.Parameters.Add(New SqlParameter("@Mensagem", _historicoContato.Mensagem))
            Command.Parameters.Add(New SqlParameter("@Usuario", _historicoContato.Usuario))
            Command.Parameters.Add(New SqlParameter("@TipoContato", IIf(String.IsNullOrEmpty(_historicoContato.TipoContato), DBNull.Value, _historicoContato.TipoContato)))
            Command.Parameters.Add(New SqlParameter("@IdReclamacao", IIf(String.IsNullOrEmpty(_historicoContato.IdReclamacao), DBNull.Value, _historicoContato.IdReclamacao)))
            Command.Parameters.Add(New SqlParameter("@IdReclamacaoSecundaria", IIf(String.IsNullOrEmpty(_historicoContato.IdReclamacaoSecundaria), DBNull.Value, _historicoContato.IdReclamacaoSecundaria)))
            Command.Parameters.Add(New SqlParameter("@Protocolo", IIf(String.IsNullOrEmpty(_historicoContato.Protocolo), DBNull.Value, _historicoContato.Protocolo)))
            Command.Parameters.Add(New SqlParameter("@IdReclamacaoTerciaria", IIf(String.IsNullOrEmpty(_historicoContato.IdReclamacaoTerciaria), DBNull.Value, _historicoContato.IdReclamacaoTerciaria)))
            Command.Parameters.Add(New SqlParameter("@Criticidade", IIf(String.IsNullOrEmpty(_historicoContato.Criticidade), DBNull.Value, _historicoContato.Criticidade)))
            Command.Parameters.Add(New SqlParameter("@ConfirmacoesNecessarias", IIf(String.IsNullOrEmpty(_historicoContato.ConfirmacoesNecessarias), DBNull.Value, _historicoContato.ConfirmacoesNecessarias)))

            Command.ExecuteNonQuery()

            _retorno.Sucesso = True
            _retorno.TipoErro = DadosGenericos.TipoErro.None

            Return _retorno
        Catch ex As Exception

            _retorno.Sucesso = False
            _retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Id
            _retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Descricao & ex.Message
            _retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retorno.NumErro, _retorno.MsgErro, _retorno.TipoErro, "Projeto: HistoricoContatoBC - Classe: InsereHistoricoContato - Função: InsereHistoricoContato(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")

            Return _retorno
        End Try

    End Function

    Public Function InsereHistoricoContato(ByVal _historicoContato As HistoricoContato, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno

        Dim _retorno As New Retorno

        'Using connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Try
            Dim Command As SqlCommand = New SqlCommand("P_InserirHistoricoContato", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            'define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", _historicoContato.CodIntClie))
            'Command.Parameters.Add(New SqlParameter("@DataHora", _historicoContato.DataHora))
            Command.Parameters.Add(New SqlParameter("@Mensagem", _historicoContato.Mensagem))
            Command.Parameters.Add(New SqlParameter("@Usuario", _historicoContato.Usuario))
            Command.Parameters.Add(New SqlParameter("@TipoContato", ""))

            'connection.Open()

            Command.ExecuteNonQuery()

            _retorno.Sucesso = True
            _retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()

            'Return _retorno
        Catch ex As Exception

            _retorno.Sucesso = False
            _retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Id
            _retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Descricao & ex.Message
            _retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retorno.NumErro, _retorno.MsgErro, _retorno.TipoErro, "Projeto: HistoricoContatoBC - Classe: InsereHistoricoContato - Função: InsereHistoricoContato(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")


        End Try

        Return _retorno
        'End Using

    End Function

    Public Function InsereHistoricoContatoOptante(ByVal _historicoContato As HistoricoContato, ByVal connection As SqlConnection, ByVal trans As SqlTransaction) As Retorno

        Dim _retorno As New Retorno

        'Using connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Try
            Dim Command As SqlCommand = New SqlCommand("P_InserirHistoricoContatoOptante", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            'define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", _historicoContato.CodIntClie))
            'Command.Parameters.Add(New SqlParameter("@DataHora", _historicoContato.DataHora))
            Command.Parameters.Add(New SqlParameter("@Mensagem", _historicoContato.Mensagem))
            Command.Parameters.Add(New SqlParameter("@Usuario", _historicoContato.Usuario))
            Command.Parameters.Add(New SqlParameter("@TipoContato", ""))

            'connection.Open()

            Command.ExecuteNonQuery()

            _retorno.Sucesso = True
            _retorno.TipoErro = DadosGenericos.TipoErro.None
            Command.Dispose()

            'Return _retorno
        Catch ex As Exception

            _retorno.Sucesso = False
            _retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Id
            _retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Descricao & ex.Message
            _retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retorno.NumErro, _retorno.MsgErro, _retorno.TipoErro, "Projeto: HistoricoContatoBC - Classe: InsereHistoricoContato - Função: InsereHistoricoContato(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")


        End Try

        Return _retorno
        'End Using

    End Function

    Public Function IncluiHistoricoContatoChecagem(ByVal strCodBanco As String, ByVal strNumAviso As String, ByVal strNumCta As String, ByVal dtDtArq As DateTime, ByVal strUsuario As String, ByVal FileName As String, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InserirHistoricoContatoChecagem_2_0_1_257", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso", strNumAviso))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@DtArq", dtDtArq))
            Command.Parameters.Add(New SqlParameter("@Usuario", strUsuario))
            Command.Parameters.Add(New SqlParameter("@NomeArquivo", FileName))
            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Id
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura


            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

    Public Function ExcluiPermissaoTipoContato(ByVal Usuario As String, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_DeletaPermissaoContato", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Usr", Usuario))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Id
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura


            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: HistoricoContatoBC - Classe: InsereHistoricoContato - Função: ExcluiPermissaoTipoContato()", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

    Public Function IncluiPermissaoTipoContato(ByVal Usuario As String, ByVal TipoContato As String, ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InserePermissaoTipoContato", connection)
            Command.Transaction = Transacao
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Usr", Usuario))
            Command.Parameters.Add(New SqlParameter("@TipoContato", TipoContato))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.MsgErro = ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSEREHISTORICOCONTATO.Id
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura


            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: HistoricoContatoBC - Classe: InsereHistoricoContato - Função: IncluiPermissaoTipoContato()", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

    Public Function InserirLogAcordoParcialDosTitulosEmAberto(UsrLibAcoPar As String, strCodIntClie As String, TitSelAcodoParcial As String) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_InserirLogAcordoParcialDosTitulosEmAberto", connection)
        Dim Retorno As Retorno = New Retorno()
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Faturamento

        Try
            connection.Open()
            command.Parameters.Add(New SqlParameter("@UsrLibAcoPar", IIf(String.IsNullOrEmpty(UsrLibAcoPar), DBNull.Value, UsrLibAcoPar)))
            command.Parameters.Add(New SqlParameter("@CodIntClie", IIf(String.IsNullOrEmpty(strCodIntClie), DBNull.Value, strCodIntClie)))
            command.Parameters.Add(New SqlParameter("@TitSelAcodoParcial", IIf(String.IsNullOrEmpty(TitSelAcodoParcial), DBNull.Value, TitSelAcodoParcial)))

            command.ExecuteNonQuery()
            Retorno.Sucesso = True
            Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
        Catch ex As Exception
            Retorno.Sucesso = False
            Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERIRLOGACORDOPARCIALDOSTITULOSEMABERTO.Descricao & ex.Message
            Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERIRLOGACORDOPARCIALDOSTITULOSEMABERTO.Id
            Funcoes.AtualizaApplEventLog(Retorno.NumErro, Retorno.MsgErro, Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
            command.Dispose()
        End Try

        Return Retorno
    End Function

    Public Function InserirLogPercentualDescontoAcordo(UsrLibDesc As String, gCodIntClie As String, gTitSelAcodoParcial As String, PercDesconto As Double) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_InserirLogPercentualDescontoAcordo", connection)
        Dim Retorno As Retorno = New Retorno()
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Faturamento

        Try
            connection.Open()
            command.Parameters.Add(New SqlParameter("@UsrLibDesc", IIf(String.IsNullOrEmpty(UsrLibDesc), DBNull.Value, UsrLibDesc)))
            command.Parameters.Add(New SqlParameter("@CodIntClie", IIf(String.IsNullOrEmpty(gCodIntClie), DBNull.Value, gCodIntClie)))
            command.Parameters.Add(New SqlParameter("@TitSelAcodoParcial", IIf(String.IsNullOrEmpty(gTitSelAcodoParcial), DBNull.Value, gTitSelAcodoParcial)))
            command.Parameters.Add(New SqlParameter("@PercDesconto", IIf(String.IsNullOrEmpty(PercDesconto), DBNull.Value, PercDesconto)))

            command.ExecuteNonQuery()
            Retorno.Sucesso = True
            Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
        Catch ex As Exception
            Retorno.Sucesso = False
            Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERIRLOGPERCENTUALDESCONTOACORDO.Descricao & ex.Message
            Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERIRLOGPERCENTUALDESCONTOACORDO.Id
            Funcoes.AtualizaApplEventLog(Retorno.NumErro, Retorno.MsgErro, Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            connection.Dispose()
            command.Dispose()
        End Try

        Return Retorno
    End Function
End Class
