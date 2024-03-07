Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.Common

Imports System.Data
Imports System.Data.SqlClient


Public Class InserirControleRetBco

    Public Function InserirControleRetBco(ByVal _ControleRetBco As ControleRetBco,
                                          ByRef connection As SqlConnection,
                                          ByRef Transaction As SqlTransaction) As Retorno

        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InserirControleRetBco", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", _ControleRetBco.CodBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso", _ControleRetBco.NumAviso))
            Command.Parameters.Add(New SqlParameter("@DtArq", _ControleRetBco.DtArq))
            Command.Parameters.Add(New SqlParameter("@NumCta", _ControleRetBco.NumCta))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERIRCONTROLERETBCO.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERIRCONTROLERETBCO.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BancoBC - Classe: InserirControleRetBco - Função: InserirControleRetBco(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

End Class
