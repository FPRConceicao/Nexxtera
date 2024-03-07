Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.Common

Imports System.Data
Imports System.Data.SqlClient

Public Class ConsultaControleRetBco


    Public Function BuscaControleRetBcoPorCodBancoNumAviso(ByVal strCodBanco As String,
                                                           ByVal strNumAviso As String,
                                                           ByVal NumCta As String,
                                                           ByRef connection As SqlConnection,
                                                           ByRef Transaction As SqlTransaction) As List(Of ControleRetBco) '#1#

        Dim rdr As SqlDataReader
        Dim lstControleRetBco As New List(Of ControleRetBco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaControleRetBcoPorCodBancoNumAviso", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso", strNumAviso))
            Command.Parameters.Add(New SqlParameter("@NumCta", NumCta))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstControleRetBco.Add(New ControleRetBco)
                    lstControleRetBco(i).CodBanco = rdr("CodBanco").ToString()
                    lstControleRetBco(i).NumAviso = rdr("NumAviso").ToString()
                    lstControleRetBco(i).DtArq = rdr("DtArq").ToString()
                    lstControleRetBco(i).NumCta = rdr("NumCta").ToString()

                    lstControleRetBco(i).Sucesso = True
                    lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstControleRetBco.Add(New ControleRetBco)
                lstControleRetBco(0).Sucesso = True
                lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            lstControleRetBco.Add(New ControleRetBco)
            lstControleRetBco(0).Sucesso = False
            lstControleRetBco(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISO.Id
            lstControleRetBco(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISO.Descricao & ex.Message
            lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstControleRetBco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstControleRetBco(0).NumErro, lstControleRetBco(0).MsgErro, lstControleRetBco(0).TipoErro, "Projeto: BancoBC - Classe: ConsultaControleRetBco - Função: BuscaControleRetBcoPorCodBancoNumAviso(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
            rdr.Close()
        End Try

        Return lstControleRetBco

    End Function


    Public Function BuscaControleRetBcoPorCodBancoNumAvisoNumCta(ByVal strCodBanco As String,
                                                                 ByVal strNumAviso As String,
                                                                 ByVal strNumCta As String,
                                                                 ByRef connection As SqlConnection,
                                                                 ByRef Transaction As SqlTransaction) As List(Of ControleRetBco) '#2#

        Dim rdr As SqlDataReader
        Dim lstControleRetBco As New List(Of ControleRetBco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaControleRetBcoPorCodBancoNumAvisoNumCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso", strNumAviso))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstControleRetBco.Add(New ControleRetBco)
                    lstControleRetBco(i).CodBanco = rdr("CodBanco").ToString()
                    lstControleRetBco(i).NumAviso = rdr("NumAviso").ToString()
                    lstControleRetBco(i).DtArq = rdr("DtArq").ToString()
                    lstControleRetBco(i).NumCta = rdr("NumCta").ToString()

                    lstControleRetBco(i).Sucesso = True
                    lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstControleRetBco.Add(New ControleRetBco)
                lstControleRetBco(i).Sucesso = True
                lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            lstControleRetBco.Add(New ControleRetBco)
            lstControleRetBco(0).Sucesso = False
            lstControleRetBco(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISONUMCTA.Id
            lstControleRetBco(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISONUMCTA.Descricao & ex.Message
            lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstControleRetBco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstControleRetBco(0).MsgErro, lstControleRetBco(0).MsgErro, lstControleRetBco(0).TipoErro, "Projeto: BancoBC - Classe: ConsultaControleRetBco - Função: BuscaControleRetBcoPorCodBancoNumAvisoNumCta(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return lstControleRetBco

    End Function



    Public Function BuscaControleRetBcoPorCodBcoNumAvisoNumCtaNumCta(ByVal strCodBanco As String,
                                                                     ByVal strNumAviso As String,
                                                                     ByVal strNumCta As String,
                                                                     ByVal DtArq As String,
                                                                     ByRef connection As SqlConnection,
                                                                     ByRef Transaction As SqlTransaction) As List(Of ControleRetBco) '#2#

        Dim rdr As SqlDataReader
        Dim lstControleRetBco As New List(Of ControleRetBco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaControleRetBcoPorCodBcoNumAvisoNumCtaNumCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso", strNumAviso))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@DtArq", CDate(DtArq)))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstControleRetBco.Add(New ControleRetBco)
                    lstControleRetBco(i).CodBanco = rdr("CodBanco").ToString()
                    lstControleRetBco(i).NumAviso = rdr("NumAviso").ToString()
                    lstControleRetBco(i).DtArq = rdr("DtArq").ToString()
                    lstControleRetBco(i).NumCta = rdr("NumCta").ToString()

                    lstControleRetBco(i).Sucesso = True
                    lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstControleRetBco.Add(New ControleRetBco)
                lstControleRetBco(i).Sucesso = True
                lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            lstControleRetBco.Add(New ControleRetBco)
            lstControleRetBco(0).Sucesso = False
            lstControleRetBco(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISONUMCTA.Id
            lstControleRetBco(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISONUMCTA.Descricao & ex.Message
            lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstControleRetBco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstControleRetBco(0).MsgErro, lstControleRetBco(0).MsgErro, lstControleRetBco(0).TipoErro, "Projeto: BancoBC - Classe: ConsultaControleRetBco - Função: BuscaControleRetBcoPorCodBcoNumAvisoNumCtaNumCta(3)", "08", "Verisure", Environment.MachineName, "2.0", "13")

        End Try

        Return lstControleRetBco

    End Function


    Public Function BuscaControleRetBcoChecagemPorCodBancoNumAvisoNumCta(ByVal strCodBanco As String, ByVal strNumAviso As String, ByVal strNumCta As String, ByRef connection As SqlConnection, ByRef Transaction As SqlTransaction) As List(Of ControleRetBco)
        Dim rdr As SqlDataReader
        Dim lstControleRetBco As New List(Of ControleRetBco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaControleRetBcoChecagemPorCodBancoNumAvisoNumCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso", strNumAviso))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstControleRetBco.Add(New ControleRetBco)
                    lstControleRetBco(i).CodBanco = rdr("CodBanco").ToString()
                    lstControleRetBco(i).NumAviso = rdr("NumAviso").ToString()
                    lstControleRetBco(i).DtArq = rdr("DtArq").ToString()
                    lstControleRetBco(i).NumCta = rdr("NumCta").ToString()

                    lstControleRetBco(i).Sucesso = True
                    lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstControleRetBco.Add(New ControleRetBco)
                lstControleRetBco(i).Sucesso = True
                lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            lstControleRetBco.Add(New ControleRetBco)
            lstControleRetBco(0).Sucesso = False
            lstControleRetBco(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISONUMCTA.Id
            lstControleRetBco(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISONUMCTA.Descricao & ex.Message
            lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstControleRetBco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstControleRetBco(0).MsgErro, lstControleRetBco(0).MsgErro, lstControleRetBco(0).TipoErro, "Projeto: BancoBC - Classe: ConsultaControleRetBco - Função: BuscaControleRetBcoPorCodBancoNumAvisoNumCta(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return lstControleRetBco

    End Function

    Public Function BuscaRelatorioClientesCobrancaBancaria(ByVal dtInicial As DateTime,
                                                                    ByVal dtFinal As DateTime,
                                                                    ByVal usr As String,
                                                                    ByVal ipMaquina As String,
                                                                    ByVal hostName As String,
                                                                    ByVal usrAD As String,
                                                                    ByRef dt As DataTable) As Retorno

        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim retorno As New Retorno
        Dim i As Integer = 0
        connection.Open()
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarRelChecagemClientesBancario", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query


            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@DtIni", IIf(IsNothing(dtInicial), DBNull.Value, dtInicial)))
            Command.Parameters.Add(New SqlParameter("@DtFim", IIf(IsNothing(dtFinal), DBNull.Value, dtFinal)))
            Command.Parameters.Add(New SqlParameter("@Usr", IIf(String.IsNullOrEmpty(usr), DBNull.Value, usr)))
            Command.Parameters.Add(New SqlParameter("@IpMaquina", IIf(String.IsNullOrEmpty(ipMaquina), DBNull.Value, ipMaquina)))
            Command.Parameters.Add(New SqlParameter("@HostName", IIf(String.IsNullOrEmpty(hostName), DBNull.Value, hostName)))
            Command.Parameters.Add(New SqlParameter("@UsrAD", IIf(String.IsNullOrEmpty(usrAD), DBNull.Value, usrAD)))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                dt.Load(rdr)
                retorno.Sucesso = True
                retorno.TipoErro = DadosGenericos.TipoErro.None
            Else
                retorno.Sucesso = False
                retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            retorno.Sucesso = False
            retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            retorno.MsgErro = ex.Message

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(retorno.MsgErro, retorno.MsgErro, retorno.TipoErro, "Projeto: BancoBC - Classe: ConsultaControleRetBco - Função: BuscaControleRetBcoPorCodBancoNumAvisoNumCta(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return retorno

    End Function

    Public Function BuscaRelatorioClientesCobrancaBancariaFiltros(ByVal dtInicial As DateTime,
                                                                    ByVal dtFinal As DateTime,
                                                                    ByVal tipoDt As String,
                                                                    ByVal isFinanciado As Boolean,
                                                                    ByVal formaPgto As String,
                                                                    ByVal bancos As String,
                                                                    ByVal usr As String,
                                                                    ByVal ipMaquina As String,
                                                                    ByVal hostName As String,
                                                                    ByVal usrAD As String,
                                                                    ByRef dt As DataTable) As Retorno

        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim retorno As New Retorno
        Dim i As Integer = 0
        connection.Open()
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarRelChecagemClientesBancario_2", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Boletagem

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@DtIni", IIf(IsNothing(dtInicial), DBNull.Value, dtInicial)))
            Command.Parameters.Add(New SqlParameter("@DtFim", IIf(IsNothing(dtFinal), DBNull.Value, dtFinal)))
            Command.Parameters.Add(New SqlParameter("@Tipo", IIf(String.IsNullOrEmpty(tipoDt), DBNull.Value, tipoDt)))
            Command.Parameters.Add(New SqlParameter("@Usr", IIf(String.IsNullOrEmpty(usr), DBNull.Value, usr)))
            Command.Parameters.Add(New SqlParameter("@IsFinanciado", IIf(isFinanciado, 1, 0)))
            Command.Parameters.Add(New SqlParameter("@FormaPgto", IIf(String.IsNullOrEmpty(formaPgto), DBNull.Value, formaPgto)))
            Command.Parameters.Add(New SqlParameter("@Bancos", IIf(String.IsNullOrEmpty(bancos), DBNull.Value, bancos)))
            Command.Parameters.Add(New SqlParameter("@IpMaquina", IIf(String.IsNullOrEmpty(ipMaquina), DBNull.Value, ipMaquina)))
            Command.Parameters.Add(New SqlParameter("@HostName", IIf(String.IsNullOrEmpty(hostName), DBNull.Value, hostName)))
            Command.Parameters.Add(New SqlParameter("@UsrAD", IIf(String.IsNullOrEmpty(usrAD), DBNull.Value, usrAD)))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                dt.Load(rdr)
                retorno.Sucesso = True
                retorno.TipoErro = DadosGenericos.TipoErro.None
            Else
                retorno.Sucesso = False
                retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
            End If
            rdr.Close()
        Catch ex As Exception

            retorno.Sucesso = False
            retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            retorno.MsgErro = ex.Message

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(retorno.MsgErro, retorno.MsgErro, retorno.TipoErro, "Projeto: BancoBC - Classe: ConsultaControleRetBco - Função: BuscaControleRetBcoPorCodBancoNumAvisoNumCta(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try

        Return retorno

    End Function

    Public Function BuscarArquivoJaProcessado(ByVal FileName As String, ByRef connection As SqlConnection, ByRef Transaction As SqlTransaction) As List(Of ControleRetBco)
        Dim rdr As SqlDataReader
        Dim lstControleRetBco As New List(Of ControleRetBco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarArquivoJaProcessado", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NomeArquivo", FileName))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstControleRetBco.Add(New ControleRetBco)
                    lstControleRetBco(i).CodBanco = rdr("CodBanco").ToString()
                    lstControleRetBco(i).NumAviso = rdr("NumAviso").ToString()
                    lstControleRetBco(i).DtArq = rdr("DtArq").ToString()
                    lstControleRetBco(i).NumCta = rdr("NumCta").ToString()
                    lstControleRetBco(i).NumCta = rdr("NomeArquivo").ToString()

                    lstControleRetBco(i).Sucesso = True
                    lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstControleRetBco.Add(New ControleRetBco)
                lstControleRetBco(i).Sucesso = False
                lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            lstControleRetBco.Add(New ControleRetBco)
            lstControleRetBco(0).Sucesso = False
            lstControleRetBco(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCARARQUIVOJAPROCESSADO.Id
            lstControleRetBco(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCARARQUIVOJAPROCESSADO.Descricao & ex.Message
            lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstControleRetBco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstControleRetBco(0).NumErro, lstControleRetBco(0).MsgErro, lstControleRetBco(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return lstControleRetBco

    End Function

End Class
