Imports System.Data
Imports System.Data.SqlClient
Imports IntegracaoNexxtera.ErrorConstants

Public Class BuscaBancoDebitoAutomatico

    ''' <summary>
    ''' Retorna uma lista com a relacao dos Banco para Debito Automatico.
    ''' </summary>
    ''' <returns>Uma lista com a relacao dos Banco para Debito Automatico.</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:     27/05/2011
    ''' Auttor:           Wolney Alexandre Fernandes
    ''' 
    ''' </remarks>
    Public Function BuscaBancoDebitoAutomatico() As List(Of Bancos) '#1#
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lista = New List(Of Bancos)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaBancoDebitoAutomatico", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Status", ""))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                While (rdr.Read)
                    lista.Add(New Bancos)
                    lista(i).CodBanco = rdr("CodBanco").ToString

                    lista(i).Sucesso = True
                    lista(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                End While
            End If
        Catch ex As Exception
            lista.Add(New Bancos)
            lista(0).Sucesso = False
            lista(0).NumErro = EXCEPTION_METODO_BUSCAADITIVOCONTRATUAL.Id
            lista(0).MsgErro = EXCEPTION_METODO_BUSCAADITIVOCONTRATUAL.Descricao & ex.Message
            lista(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lista(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lista(0).NumErro, lista(0).MsgErro, lista(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoDebitoAutomatico - Função: BuscaBancoDebitoAutomatico(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return lista

    End Function




    ''' <summary>
    ''' Retorna uma lista com a relação dos Banco para Debito Automatico filtrando pelo status.
    ''' </summary>
    ''' <param name="strStatus">informar o status desejado - campo obrigatório (A-I)</param>
    ''' <returns>Uma lista com a relacao dos Banco para Debito Automatico.</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:     15/06/2011
    ''' Auttor:           Wolney Alexandre Fernandes
    ''' 
    ''' </remarks>
    Public Function BuscaBancoDebitoAutomatico(ByVal strStatus As String) As List(Of Bancos)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lista = New List(Of Bancos)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaBancoDebitoAutomatico", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Status", strStatus))

            ''Abre a conexao
            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                While (rdr.Read)
                    lista.Add(New Bancos)
                    lista(i).CodBanco = rdr("CodBanco").ToString

                    lista(i).Sucesso = True
                    lista(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                End While
            End If
        Catch ex As Exception
            lista.Add(New Bancos)
            lista(0).Sucesso = False
            lista(0).NumErro = EXCEPTION_METODO_BUSCAADITIVOCONTRATUAL.Id
            lista(0).MsgErro = EXCEPTION_METODO_BUSCAADITIVOCONTRATUAL.Descricao & ex.Message
            lista(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lista(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lista(0).NumErro, lista(0).MsgErro, lista(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoDebitoAutomatico - Função: BuscaBancoDebitoAutomatico(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return lista

    End Function




    ''' <summary>
    ''' Busca o banco para debito Automatico 
    ''' </summary>
    ''' <param name="Banco">Variavel com o codigo do banco</param>
    ''' <returns>Um string contendo codigo do banco | agencia | conta | nome do banco</returns>
    ''' <remarks></remarks>
    Public Function PesquisaBancoParametro(ByVal Banco As String) As String '#3#
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim strBanco As String = ""
        Dim i As Integer = 0
        Dim _Retorno As New Retorno


        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_PesquisaContaDebAutomatico", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Banco", Banco))
            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                strBanco = rdr("Banco").ToString()

                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            Else
                _Retorno.Sucesso = True
                _Retorno.TipoErro = DadosGenericos.TipoErro.None
            End If
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_PESQUISABANCOPARAMETRO.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_PESQUISABANCOPARAMETRO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoDebitoAutomatico - Função: PesquisaBancoParametro(3)", "08", "Verisure", Environment.MachineName, "2.0", "13")

            Throw New Exception(ex.Message)
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return strBanco

    End Function



    Public Function BuscaBancoContaTipoContaDoBanco(ByVal strCodBanco As String,
                                                    ByVal strStatus As String) As List(Of Banco) '#4#
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstBanco As New List(Of Banco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaBancoContaTipoContaDoBanco", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@Status", strStatus))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstBanco.Add(New Banco)
                    lstBanco(i).CodBanco = rdr("CodBanco").ToString()
                    lstBanco(i).NomeBanco = rdr("NomeBanco").ToString()
                    lstBanco(i).CodAgen = rdr("CodAgen").ToString()
                    lstBanco(i).Numcta = rdr("numcta").ToString()
                    lstBanco(i).TipoConta = rdr("TipoConta").ToString()

                    lstBanco(i).Sucesso = True
                    lstBanco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
                rdr.Close()
            Else
                lstBanco.Add(New Banco)
                lstBanco(0).Sucesso = False
                lstBanco(0).NumErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Id
                lstBanco(0).MsgErro = "Nenhum Banco encontrado"
                lstBanco(0).TipoErro = DadosGenericos.TipoErro.None
                lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End If
        Catch ex As Exception

            lstBanco.Add(New Banco)
            lstBanco(0).Sucesso = True
            lstBanco(0).NumErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Id
            lstBanco(0).MsgErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Descricao & ex.Message
            lstBanco(0).TipoErro = DadosGenericos.TipoErro.None
            lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstBanco(0).NumErro, lstBanco(0).MsgErro, lstBanco(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoDebitoAutomatico - Função: BuscaBancoContaTipoContaDoBanco(4)", "08", "Verisure", Environment.MachineName, "2.0", "13")

        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return lstBanco

    End Function

    Public Function BuscarContaCorrenteCobranca(ByVal status As Integer) As List(Of Banco) '#4#
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstBanco As New List(Of Banco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarContaCorrenteCobranca_API", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Status", status))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstBanco.Add(New Banco)
                    lstBanco(i).CodBanco = rdr("CodBanco").ToString()
                    lstBanco(i).NomeBanco = rdr("NomeBanco").ToString()
                    lstBanco(i).CodAgen = rdr("CodAgen").ToString()
                    lstBanco(i).Numcta = rdr("numcta").ToString()

                    lstBanco(i).Sucesso = True
                    lstBanco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
                rdr.Close()
            Else
                lstBanco.Add(New Banco)
                lstBanco(0).Sucesso = False
                lstBanco(0).NumErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Id
                lstBanco(0).MsgErro = "Nenhum Banco encontrado"
                lstBanco(0).TipoErro = DadosGenericos.TipoErro.None
                lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End If
        Catch ex As Exception

            lstBanco.Add(New Banco)
            lstBanco(0).Sucesso = True
            lstBanco(0).NumErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Id
            lstBanco(0).MsgErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Descricao & ex.Message
            lstBanco(0).TipoErro = DadosGenericos.TipoErro.None
            lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstBanco(0).NumErro, lstBanco(0).MsgErro, lstBanco(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoDebitoAutomatico - Função: BuscaBancoContaTipoContaDoBanco(4)", "08", "Verisure", Environment.MachineName, "2.0", "13")

        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return lstBanco

    End Function


    Public Function BuscaBancoContaTipoContaDoBancoEStatus(ByVal strCodBanco As String,
                                                           ByVal strStatus As String) As List(Of Banco) '#4#
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstBanco As New List(Of Banco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaBancoContaTipoContaDoBanco", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@Status", strStatus))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstBanco.Add(New Banco)
                    lstBanco(i).CodBanco = rdr("CodBanco").ToString()
                    lstBanco(i).NomeBanco = rdr("NomeBanco").ToString()
                    lstBanco(i).CodAgen = rdr("CodAgen").ToString()
                    lstBanco(i).Numcta = rdr("numcta").ToString()
                    lstBanco(i).TipoConta = rdr("TipoConta").ToString()

                    lstBanco(i).Sucesso = True
                    lstBanco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
                rdr.Close()
            End If
        Catch ex As Exception

            lstBanco.Add(New Banco)
            lstBanco(0).Sucesso = True
            lstBanco(0).NumErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Id
            lstBanco(0).MsgErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Descricao & ex.Message
            lstBanco(0).TipoErro = DadosGenericos.TipoErro.None
            lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstBanco(0).NumErro, lstBanco(0).MsgErro, lstBanco(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoDebitoAutomatico - Função: BuscaBancoContaTipoContaDoBanco(4)", "08", "Verisure", Environment.MachineName, "2.0", "13")

        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return lstBanco

    End Function

    Public Function CodigoRetornoOptanteCaixa(ByVal codRetorno As String, ByRef msg As String) As Retorno

        Dim retono As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscarCodRetornoOptanteCaixa", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodRetorno", codRetorno))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                msg = rdr("Msg").ToString()
                rdr.Close()

                retono.Sucesso = True
            Else
                retono.Sucesso = True
                retono.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
        Catch ex As Exception

            retono = New Retorno
            retono.Sucesso = False
            retono.MsgErro = ErrorConstants.EXCEPTION_BUSCABANCOCONTATIPOCONTADOBANCO.Descricao & ex.Message
            retono.TipoErro = DadosGenericos.TipoErro.None
            retono.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(retono.NumErro, retono.MsgErro, retono.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoDebitoAutomatico - Função: BuscaBancoContaTipoContaDoBanco(4)", "08", "Verisure", Environment.MachineName, "2.0", "13")

        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return retono

    End Function


End Class
