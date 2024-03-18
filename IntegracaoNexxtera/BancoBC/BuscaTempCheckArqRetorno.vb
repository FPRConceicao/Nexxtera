Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.Common

Imports System.Data
Imports System.Data.SqlClient

Public Class BuscaTempCheckArqRetorno

    Public Function BuscaRelExcelCheckArqRetorno(ByVal otb As DataTable) As Retorno
        Dim _Retorno As New Retorno
        Dim rdr As SqlDataReader = Nothing
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim cmd As New SqlCommand("P_RelatChkRetornoExcel", Connection)

        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = DadosGenericos.Timeout.Query

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
            _Retorno.NumErro = "BuscaRelExcelCheckArqRetorno"
            _Retorno.MsgErro = "Erro ao buscar dados: " & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        Finally
            Connection.Close()
            cmd.Dispose()
            Connection.Dispose()
        End Try

        Return _Retorno
    End Function



    Public Function BuscaRelExcelCheckArqRetornoMultasJuros(ByVal otb As DataTable) As Retorno
        Dim _Retorno As New Retorno
        Dim rdr As SqlDataReader = Nothing
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        'Dim cmd As New SqlCommand("P_RelatChkRetornoExcelMultasJuros", Connection)
        Dim cmd As New SqlCommand("P_RelatChkRetornoExcelMultasJurosNexxtera", Connection)
        Try
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandTimeout = DadosGenericos.Timeout.Query

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
            _Retorno.NumErro = "BuscaRelExcelCheckArqRetorno"
            _Retorno.MsgErro = "Erro ao buscar dados: " & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "Verisure", Environment.MachineName, "", "")
        Finally
            Connection.Close()
            cmd.Dispose()
            Connection.Dispose()
        End Try

        Return _Retorno
    End Function



End Class
