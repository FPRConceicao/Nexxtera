Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Entidades

Imports System.Data
Imports System.Data.SqlClient

Public Class InsereCodInconsistencia

    Public Function InsereErroDadosPgtoCliente(ByVal CodIntClie As String, ByVal connection As SqlConnection, ByVal Transaction As SqlTransaction) As Retorno

        'TODO VERIFICAR ONDE É UTILIZADO ESTA FUNCTION
        Dim Retorno As New Retorno

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereErroDadosPgtoCliente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", CodIntClie))

            ''Executa a procedure
            Command.ExecuteNonQuery()
            Retorno.Sucesso = True
            Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            Retorno = Funcoes.RetornoArq(ex.Message(), "InsereErroDadosPgtoCliente")

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(Retorno.NumErro, Retorno.MsgErro, Retorno.TipoErro, "Projeto: CodInconsistenciasBC - Classe: InsereCodInconsistencia - Função: InsereErroDadosPgtoCliente", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return Retorno

    End Function

End Class
