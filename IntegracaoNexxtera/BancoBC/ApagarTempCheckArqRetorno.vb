Imports Teleatlantic.TLS.Common
Imports System.Data
Imports System.Data.SqlClient
Imports Teleatlantic.TLS.Entidades


Public Class ApagarTempCheckArqRetorno

    Public Function ApagaTudoTempCheckArqRetorno(ByVal Connection As SqlConnection, ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_ApagaTudoTempCheckArqRetorno", Connection)

        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS  
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        End Try

        Return _Retorno

    End Function
    Public Function ApagaTudoTempCheckArqRetornoNexxtera() As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_ApagaTudoTempCheckArqRetorno", connection)

        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            connection.Open()
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS  
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        Finally
            connection.Close()
            Command.Dispose()
        End Try

        Return _Retorno

    End Function

    Public Function ApagaTudoTempCheckArqRetorno() As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_ApagaTudoTempCheckArqRetorno", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()
            command.ExecuteNonQuery()
            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function ApagaTudoTempArqRetornoOptante(ByVal Connection As SqlConnection, ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_ApagaTudoTempArqRetornoOptante", Connection)

        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        End Try
        Return _Retorno

    End Function
    Public Function ApagaTudoTempArqRetornoOptante() As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_ApagaTudoTempArqRetornoOptante", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()
            command.ExecuteNonQuery()
            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function ApagaTudoTempCheckArqRetCartao(ByVal Connection As SqlConnection, ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_ApagaTudoTempCheckArqRetCartao", Connection)

        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        End Try
        Return _Retorno

    End Function

    Public Function ApagaTudoTempCheckArqRetCartao() As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_ApagaTudoTempCheckArqRetCartao", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()
            command.ExecuteNonQuery()
            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_APAGATUDOTEMPCHECKARQRETORNO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

End Class
