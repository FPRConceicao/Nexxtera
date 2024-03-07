Imports System.Data
Imports System.Data.SqlClient
Imports IntegracaoNexxtera.ErrorConstants
Imports System.Reflection

Public Class BuscaBancoContaCorrente
    Public Function BuscaBancoContaCorrente() As List(Of Banco) '#1#
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lista = New List(Of Banco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaBancoContaCorrente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then

                While (rdr.Read)
                    lista.Add(New Banco)
                    lista(i).CodBancoCodAgenciaNomeBanco = rdr("Banco").ToString

                    lista(i).Sucesso = True
                    lista(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                End While
            End If
        Catch ex As Exception
            lista.Add(New Banco)
            lista(0).Sucesso = False
            lista(0).NumErro = EXCEPTION_METODO_BUSCAADITIVOCONTRATUAL.Id
            lista(0).MsgErro = EXCEPTION_METODO_BUSCAADITIVOCONTRATUAL.Descricao & ex.Message
            lista(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lista(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lista(0).NumErro, lista(0).MsgErro, lista(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaBancoContaCorrente(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return lista

    End Function


    Public Function BuscaTaxaJurosParametrizada(ByVal strCodBanco As String,
                                               ByVal sAgencia As String,
                                               ByVal sNumConta As String) As Banco '#2#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco341", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", sAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", sNumConta))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.Read() Then
                ' processa a linha do resultado
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.TaxaMulta = rdr("TaxaMulta").ToString()
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString()
                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None
            End If

            rdr.Close()

        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return _Banco

    End Function


    Public Function BuscaCodigoSeqRemessaBanco(ByVal strCodBanco As String,
                                               ByVal strCodAgencia As String,
                                               ByVal strNumCta As String) As Banco '#2#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.CodPRemes = rdr("CodPRemes").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return _Banco

    End Function



    Public Function BuscaCodigoSeqRemessaBanco001(ByVal strCodBanco As String,
                                                  ByVal strNumCta As String,
                                                  ByVal Connection As SqlConnection,
                                                  ByVal Transaction As SqlTransaction) As Banco '#2#

        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco001", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumConta", strNumCta))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.CodPRemes = rdr("Convenio").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO001.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO001.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco001(3)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Banco

    End Function



    Public Function BuscaCodigoSeqRemessaBanco237(ByVal strCodBanco As String,
                                                  ByVal strNumCta As String,
                                                  ByVal strCodAgencia As String) As Banco '#2#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco237", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgencia", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumConta", strNumCta))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.CodPRemes = rdr("CodPRemes").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco237(4)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return _Banco

    End Function


    Public Function BuscaCodigoSeqRemessaBanco237(ByVal strCodBanco As String,
                                                  ByVal strNumCta As String,
                                                  ByVal strCodAgencia As String,
                                                  ByVal Connection As SqlConnection,
                                                  ByVal Transaction As SqlTransaction) As Banco '#2#

        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco237", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgencia", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumConta", strNumCta))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.CodPRemes = rdr("CodPRemes").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco237(5)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Banco

    End Function


    Public Function BuscaCodigoSeqRemessaBanco347(ByVal strCodBanco As String,
                                                  ByVal strNumCta As String,
                                                  ByVal strCodAgencia As String,
                                                  ByVal Connection As SqlConnection,
                                                  ByVal Transaction As SqlTransaction) As Banco '#2#

        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco347", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.CodPRemes = rdr("Convenio").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                '_Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO347.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO347.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco347(6)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Banco

    End Function



    Public Function BuscaCodigoSeqRemessaBanco291(ByVal strCodBanco As String,
                                                  ByVal strNumCta As String,
                                                  ByVal strCodAgencia As String) As Banco '#2#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco237", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Abre a conexao
            connection.Open()
            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                '_Banco.CodPRemes = rdr("Convenio").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                '_Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco237(7)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            'Fecha a conexao
            connection.Close()
            connection.Dispose()
        End Try

        Return _Banco

    End Function


    Public Function BuscaCodigoSeqRemessaBanco341_033(ByVal strCodBanco As String,
                                                      ByVal strNumCta As String,
                                                      ByVal strCodAgencia As String,
                                                      ByVal Connection As SqlConnection,
                                                      ByVal Transaction As SqlTransaction) As Banco '#2#
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco341", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                '_Banco.CodPRemes = rdr("Convenio").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString
                _Banco.ContaCorrente.NossoNumero = rdr("NossoNumero").ToString
                _Banco.ContaCorrente.CodTransmissao = rdr("CodTransmissao").ToString
                _Banco.ContaCorrente.Convenio = rdr("Convenio").ToString
                _Banco.ContaCorrente.CodFlash = rdr("CodFlash").ToString
                _Banco.ContaCorrente.IsGerarNossoNumero = rdr("IsGerarNossoNumero")

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO341.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO341.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco341(8)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Banco

    End Function


    Public Function BuscaCodigoSeqRemessaBanco399(ByVal strCodBanco As String,
                                                  ByVal strNumCta As String,
                                                  ByVal strCodAgencia As String,
                                                  ByVal Connection As SqlConnection,
                                                  ByVal Transaction As SqlTransaction) As Banco '#2#

        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco399", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))


            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.ContaCorrente.Convenio = rdr("Convenio").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO347.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO347.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco347(9)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Banco

    End Function



    Public Function BuscabcoAgenCtaContaCorrente(Optional ByVal strStatus As String = "") As List(Of Banco) '#2#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstBanco As New List(Of Banco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscabcoAgenCtaContaCorrente", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Status", strStatus))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstBanco.Add(New Banco)
                    lstBanco(i).CodBancoCodAgenciaNomeBanco = rdr("BcoAgeCta")

                    lstBanco(i).Sucesso = True
                    lstBanco(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            End If
            rdr.Close()
        Catch ex As Exception

            lstBanco(0).Sucesso = False
            lstBanco(0).NumErro = EXCEPTION_METODO_BUSCABCOAGENCTACONTACORRENTE.Id
            lstBanco(0).MsgErro = EXCEPTION_METODO_BUSCABCOAGENCTACONTACORRENTE.Descricao & ex.Message
            lstBanco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstBanco(0).NumErro, lstBanco(0).MsgErro, lstBanco(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscabcoAgenCtaContaCorrente(10)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return lstBanco

    End Function



    Public Function BuscaBancoAgenciaNumCtaContaCorrentePorStatusA() As List(Of Banco) '#2#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstBanco As New List(Of Banco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaBancoAgenciaNumCtaContaCorrentePorStatusA", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstBanco.Add(New Banco)
                    lstBanco(i).CodBancoCodAgenciaNomeBanco = rdr("CodBancoCodAgenNumCta")

                    lstBanco(i).Sucesso = True
                    lstBanco(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            Else
                lstBanco(0).Sucesso = False
                lstBanco(0).NumErro = NENHUM_REGISTRO_ENCONTRADO.Id
                lstBanco(0).MsgErro = NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstBanco(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End If
            rdr.Close()
        Catch ex As Exception

            lstBanco(0).Sucesso = False
            lstBanco(0).NumErro = EXCEPTION_METODO_BUSCABANCOAGENCIANUMCTACONTACORRENTEPORSTATUSA.Id
            lstBanco(0).MsgErro = EXCEPTION_METODO_BUSCABANCOAGENCIANUMCTACONTACORRENTEPORSTATUSA.Descricao & ex.Message
            lstBanco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstBanco(0).NumErro, lstBanco(0).MsgErro, lstBanco(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaBancoAgenciaNumCtaContaCorrentePorStatusA(11)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return lstBanco

    End Function


    Public Function BuscaDescricaoContaCorrentePorCodBancoCodAgenNumCta(ByVal strCodBanco As String,
                                                                        ByVal strCodAgen As String,
                                                                        ByVal strNumCta As String) As List(Of ContaCorrente) '#2#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstContaCorrente As New List(Of ContaCorrente)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaDescricaoContaCorrentePorCodBancoCodAgenNumCta", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstContaCorrente.Add(New ContaCorrente)
                    lstContaCorrente(i).Descricao = rdr("Descricao")
                    lstContaCorrente(i).CodBanco = rdr("CodBanco")
                    lstContaCorrente(i).CodAgen = rdr("CodAgen")
                    lstContaCorrente(i).NumCta = rdr("NumCta")

                    lstContaCorrente(i).Sucesso = True
                    lstContaCorrente(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            Else
                lstContaCorrente.Add(New ContaCorrente)
                lstContaCorrente(0).Sucesso = False
                lstContaCorrente(0).NumErro = NENHUM_REGISTRO_ENCONTRADO.Id
                lstContaCorrente(0).MsgErro = NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End If
            rdr.Close()
        Catch ex As Exception
            lstContaCorrente.Add(New ContaCorrente)
            lstContaCorrente(0).Sucesso = False
            lstContaCorrente(0).NumErro = EXCEPTION_METODO_BUSCADESCRICAOCONTACORRENTEPORCODBANCOCODAGENNUMCTA.Id
            lstContaCorrente(0).MsgErro = EXCEPTION_METODO_BUSCADESCRICAOCONTACORRENTEPORCODBANCOCODAGENNUMCTA.Descricao & ex.Message
            lstContaCorrente(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstContaCorrente(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstContaCorrente(0).NumErro, lstContaCorrente(0).MsgErro, lstContaCorrente(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaDescricaoContaCorrentePorCodBancoCodAgenNumCta(12)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return lstContaCorrente

    End Function



    Public Function BuscaNomeAgenPorCodBancoCodAgenBanco(ByVal strCodBanco As String,
                                                         ByVal strCodAgen As String) As List(Of Banco) '#2#

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim rdr As SqlDataReader
        Dim lstBanco As New List(Of Banco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaNomeAgenPorCodBancoCodAgenBanco", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgen))

            connection.Open()

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstBanco.Add(New Banco)
                    lstBanco(i).NomeAgen = rdr("NomeAgen")
                    lstBanco(i).CodBanco = rdr("CodBanco")
                    lstBanco(i).CodAgen = rdr("CodAgen")

                    lstBanco(i).Sucesso = True
                    lstBanco(i).TipoErro = DadosGenericos.TipoErro.None
                    i = i + 1
                Loop
            Else
                lstBanco.Add(New Banco)
                lstBanco(0).Sucesso = False
                lstBanco(0).NumErro = NENHUM_REGISTRO_ENCONTRADO.Id
                lstBanco(0).MsgErro = NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstBanco(0).TipoErro = DadosGenericos.TipoErro.Funcional
                lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End If
            rdr.Close()
        Catch ex As Exception
            lstBanco.Add(New Banco)
            lstBanco(0).Sucesso = False
            lstBanco(0).NumErro = EXCEPTION_METODO_BUSCANOMEAGENPORCODBANCOCODAGENBANCO.Id
            lstBanco(0).MsgErro = EXCEPTION_METODO_BUSCANOMEAGENPORCODBANCOCODAGENBANCO.Descricao & ex.Message
            lstBanco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstBanco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstBanco(0).NumErro, lstBanco(0).MsgErro, lstBanco(0).TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaNomeAgenPorCodBancoCodAgenBanco(13)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
        End Try


        Return lstBanco

    End Function

    Public Function BuscaCodigoSeqRemessaBanco041(ByVal strCodBanco As String,
                                                  ByVal strNumCta As String,
                                                  ByVal strCodAgencia As String,
                                                  ByVal Connection As SqlConnection,
                                                  ByVal Transaction As SqlTransaction) As Banco '#2#

        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco041", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgencia", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumConta", strNumCta))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.CodPRemes = rdr("CodPRemes").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString
                _Banco.ContaCorrente.Convenio = rdr("Convenio").ToString()

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco237(5)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Banco

    End Function

    Public Function BuscaCodigoSeqRemessaBanco041(ByVal strCodBanco As String,
                                                  ByVal strNumCta As String,
                                                  ByVal strCodAgencia As String,
                                                  ByVal isContaPadrao As Boolean) As Banco '#3#

        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim i As Integer = 0
        Dim con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        con.Open()

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodigoSeqRemessaBanco104", con)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query


            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgencia", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumConta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@IsContaPadrao", isContaPadrao))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.ContaCorrente.Convenio = rdr("Convenio").ToString()

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, "Projeto: BancoBC - Classe: BuscaBancoContaCorrente - Função: BuscaCodigoSeqRemessaBanco237(5)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            con.Close()
            If Not IsNothing(rdr) Then rdr.Close()
        End Try

        Return _Banco

    End Function

    Public Function BuscaCodigoSeqRemessaBanco341_033(ByVal strCodBanco As String, ByVal strNumCta As String, ByVal strCodAgencia As String) As Banco '#2#
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaCodigoSeqRemessaBanco341", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.TaxaMulta = rdr("TaxaMulta").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString
                _Banco.ContaCorrente.NossoNumero = rdr("NossoNumero").ToString
                _Banco.ContaCorrente.CodTransmissao = rdr("CodTransmissao").ToString
                _Banco.ContaCorrente.Convenio = rdr("Convenio").ToString
                _Banco.ContaCorrente.CodFlash = rdr("CodFlash").ToString
                _Banco.ContaCorrente.IsGerarNossoNumero = rdr("IsGerarNossoNumero")

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO341.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO341.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Banco
    End Function

    Public Function BuscaCodigoSeqRemessaBanco041(ByVal strCodBanco As String, ByVal strNumCta As String, ByVal strCodAgencia As String) As Banco '#2#
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaCodigoSeqRemessaBanco041", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgencia", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumConta", strNumCta))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.CodPRemes = rdr("CodPRemes").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString
                _Banco.ContaCorrente.Convenio = rdr("Convenio").ToString()

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception
            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO237.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Banco

    End Function

    Public Function BuscaCodigoSeqRemessaBanco001(ByVal strCodBanco As String, ByVal strNumCta As String) As Banco '#2#
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaCodigoSeqRemessaBanco001", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@NumConta", strNumCta))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.CodPRemes = rdr("Convenio").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO001.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO001.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Banco
    End Function

    Public Function BuscaCodigoSeqRemessaBanco399(ByVal strCodBanco As String, ByVal strNumCta As String, ByVal strCodAgencia As String) As Banco '#2#
        Dim rdr As SqlDataReader
        Dim _Banco As New Banco
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaCodigoSeqRemessaBanco399", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))


            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                rdr.Read()
                _Banco.ContaCorrente = New ContaCorrente
                _Banco.ContaCorrente.SeqRemes = rdr("SeqRemes").ToString
                _Banco.ContaCorrente.Convenio = rdr("Convenio").ToString
                _Banco.TaxaDia = rdr("TaxaDia").ToString
                _Banco.TaxaMes = rdr("TaxaMes").ToString
                _Banco.ContaCorrente.CodCarteira = rdr("CodCarteira").ToString
                _Banco.SeqRemesUnico = rdr("SeqRemesUnico").ToString

                _Banco.Sucesso = True
                _Banco.TipoErro = DadosGenericos.TipoErro.None

            End If
            rdr.Close()
        Catch ex As Exception

            _Banco.Sucesso = False
            _Banco.NumErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO347.Id
            _Banco.MsgErro = EXCEPTION_BUSCACODIGOSEQREMESSABANCO347.Descricao & ex.Message
            _Banco.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Banco.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Banco.NumErro, _Banco.MsgErro, _Banco.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            Connection.Close()
            Command.Dispose()
        End Try

        Return _Banco
    End Function
End Class
