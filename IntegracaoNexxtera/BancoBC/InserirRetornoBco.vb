Imports System.Data
Imports System.Data.SqlClient
Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Common.ErrorConstants

Public Class InserirRetornoBco

    Public Function InserirRetornoBco(ByVal _RetornoBco As RetornoBco,
                                      ByRef connection As SqlConnection,
                                      ByRef trans As SqlTransaction) As Retorno

        Dim _retorno As New Retorno

        Try
            Dim Command As SqlCommand = New SqlCommand("P_InserirRetornoBco", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            'define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", _RetornoBco.CodBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso  ", _RetornoBco.NumAviso))
            Command.Parameters.Add(New SqlParameter("@NumTit   ", _RetornoBco.NumTit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", _RetornoBco.SeqTit))
            Command.Parameters.Add(New SqlParameter("@CodAgen  ", _RetornoBco.CodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta   ", _RetornoBco.NumCta))
            Command.Parameters.Add(New SqlParameter("@VlrPago", _RetornoBco.VlrPago))
            Command.Parameters.Add(New SqlParameter("@VlrJuros  ", _RetornoBco.VlrJuros))
            Command.Parameters.Add(New SqlParameter("@VlrDesc   ", _RetornoBco.VlrDesc))
            Command.Parameters.Add(New SqlParameter("@VlrIOF", _RetornoBco.VlrIOF))
            Command.Parameters.Add(New SqlParameter("@VlrAbat  ", _RetornoBco.VlrAbat))
            Command.Parameters.Add(New SqlParameter("@Processado   ", _RetornoBco.Processado))
            Command.Parameters.Add(New SqlParameter("@DtVcto", _RetornoBco.DtVcto))
            Command.Parameters.Add(New SqlParameter("@DtPagto  ", _RetornoBco.DtPagto))
            Command.Parameters.Add(New SqlParameter("@DtArq", _RetornoBco.DtArq))
            Command.Parameters.Add(New SqlParameter("@VlrMulta  ", _RetornoBco.VlrMulta))


            Command.ExecuteNonQuery()

            _retorno.Sucesso = True
            _retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception

            _retorno.Sucesso = False
            _retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERIRRETORNOBCO.Id
            _retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERIRRETORNOBCO.Descricao & ex.Message
            _retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_retorno.NumErro, _retorno.MsgErro, _retorno.TipoErro, "Projeto: BancoBC - Classe: InserirRetornoBco - Função: InserirRetornoBco(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")

        End Try

        Return _retorno
    End Function

End Class
