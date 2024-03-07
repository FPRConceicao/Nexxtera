Imports Teleatlantic.TLS.Common
Imports System.Data
Imports System.Data.SqlClient
Imports Teleatlantic.TLS.Entidades
Imports System.Reflection

Public Class ApagarTitulosEnvBco

    Public Function ApagarTitulosEnvBco(ByVal strNumTit As String,
                                        ByVal strSeqTit As String,
                                        ByVal dtDataEmissao As DateTime,
                                        ByVal Connection As SqlConnection,
                                        ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_ApagarTItulosEnvBco", Connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtDataEmissao))


            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: ParametroBC - Classe: ApagarTitulosEnvBco - Função: ApagarTitulosEnvBco(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try
        Return _Retorno

    End Function

    Public Function ApagarTitulosEnvBco(ByVal strNumTit As String, ByVal strSeqTit As String, ByVal dtDataEmissao As DateTime) As Retorno
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_ApagarTItulosEnvBco", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parametros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@NumTit", strNumTit))
            command.Parameters.Add(New SqlParameter("@SeqTit", strSeqTit))
            command.Parameters.Add(New SqlParameter("@DtEmissao", dtDataEmissao))

            ''Executa a procedure
            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

End Class
