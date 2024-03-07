Imports Teleatlantic.TLS.Common
Imports Teleatlantic.TLS.Entidades
Imports System.Data
Imports System.Data.SqlClient

Public Class ConsultaBxAutosErros

    Public Function BuscaBxAutoErros(ByRef connection As SqlConnection,
                                     ByRef Transaction As SqlTransaction) As List(Of BxAutosErros) '#1#

        Dim rdr As SqlDataReader
        Dim lstBxAutosErros As New List(Of BxAutosErros)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaBxAutoErros", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstBxAutosErros.Add(New BxAutosErros)
                    lstBxAutosErros(i).NumAviso = rdr("NumAviso").ToString()
                    lstBxAutosErros(i).NumTit = rdr("NumTit").ToString()
                    lstBxAutosErros(i).SeqTit = rdr("SeqTit").ToString()
                    lstBxAutosErros(i).DtEmissao = rdr("DtEmissao").ToString()

                    lstBxAutosErros(i).Sucesso = True
                    lstBxAutosErros(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstBxAutosErros.Add(New BxAutosErros)
                lstBxAutosErros(i).Sucesso = True
                lstBxAutosErros(i).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            lstBxAutosErros.Add(New BxAutosErros)
            lstBxAutosErros(0).Sucesso = False
            lstBxAutosErros(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCABXAUTOERROS.Id
            lstBxAutosErros(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCABXAUTOERROS.Descricao & ex.Message
            lstBxAutosErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstBxAutosErros(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstBxAutosErros(0).MsgErro, lstBxAutosErros(0).MsgErro, lstBxAutosErros(0).TipoErro, "Projeto: BancoBC - Classe: ConsultaBxAutosErros - Função: BuscaBxAutoErros(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")

        End Try

        Return lstBxAutosErros

    End Function


End Class
