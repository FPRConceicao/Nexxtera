
Imports System.Data.SqlClient

Public Class ConsultaCodInconsistencias

    Public Function BuscaCodInconsistencias(ByVal strCodIncons As String, ByVal strCodBanco As String, ByVal strCodOcor As String, ByVal connection As SqlConnection, ByVal Transaction As SqlTransaction) As List(Of CodInconsistencias)

        'TODO VERIFICAR ONDE É UTILIZADO ESTA FUNCTION
        Dim rdr As SqlDataReader
        Dim lstCodInconsistencias As New List(Of CodInconsistencias)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodInconsistencias", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIncons", strCodIncons))
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodOcor", strCodOcor))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstCodInconsistencias.Add(New CodInconsistencias)
                    lstCodInconsistencias(i).CodIncons = rdr("CodIncons")
                    lstCodInconsistencias(i).Mensagem = rdr("Mensagem")
                    lstCodInconsistencias(i).CodOcor = rdr("CodOcor")
                    lstCodInconsistencias(i).CodBanco = rdr("CodBanco")

                    lstCodInconsistencias(i).Sucesso = True
                    lstCodInconsistencias(i).TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                lstCodInconsistencias.Add(New CodInconsistencias)
                lstCodInconsistencias(0).Sucesso = False
                lstCodInconsistencias(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstCodInconsistencias(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstCodInconsistencias(0).TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()
        Catch ex As Exception
            lstCodInconsistencias.Add(New CodInconsistencias)
            lstCodInconsistencias(0).Sucesso = False
            lstCodInconsistencias(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Descricao & ex.Message
            lstCodInconsistencias(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Id
            lstCodInconsistencias(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstCodInconsistencias(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstCodInconsistencias(0).NumErro, lstCodInconsistencias(0).MsgErro, lstCodInconsistencias(0).TipoErro, "Projeto: CodInconsistenciasBC - Classe: ConsultaCodInconsistencias - Função: BuscaCodInconsistencias(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return lstCodInconsistencias

    End Function

    Public Function BuscaCodInconsistencias(ByVal strCodIncons As String, ByVal strCodBanco As String, ByVal strCodOcor As String) As List(Of CodInconsistencias)
        'TODO VERIFICAR ONDE É UTILIZADO ESTA FUNCTION
        Dim rdr As SqlDataReader
        Dim lstCodInconsistencias As New List(Of CodInconsistencias)
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_BuscaCodInconsistencias", connection)
        Dim i As Integer = 0

        Try
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query
            connection.Open()

            ''define os parametros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@CodIncons", strCodIncons))
            command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodOcor", strCodOcor))

            ''Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    lstCodInconsistencias.Add(New CodInconsistencias)
                    lstCodInconsistencias(i).CodIncons = rdr("CodIncons")
                    lstCodInconsistencias(i).Mensagem = rdr("Mensagem")
                    lstCodInconsistencias(i).CodOcor = rdr("CodOcor")
                    lstCodInconsistencias(i).CodBanco = rdr("CodBanco")

                    lstCodInconsistencias(i).Sucesso = True
                    lstCodInconsistencias(i).TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                lstCodInconsistencias.Add(New CodInconsistencias)
                lstCodInconsistencias(0).Sucesso = False
                lstCodInconsistencias(0).MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                lstCodInconsistencias(0).NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                lstCodInconsistencias(0).TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()
        Catch ex As Exception
            lstCodInconsistencias.Add(New CodInconsistencias)
            lstCodInconsistencias(0).Sucesso = False
            lstCodInconsistencias(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Descricao & ex.Message
            lstCodInconsistencias(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Id
            lstCodInconsistencias(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstCodInconsistencias(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstCodInconsistencias(0).NumErro, lstCodInconsistencias(0).MsgErro, lstCodInconsistencias(0).TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "08", "Verisure", Environment.MachineName, "2.0", "13")
        Finally
            command.Dispose()
            connection.Close()
        End Try

        Return lstCodInconsistencias

    End Function

    Public Function BuscaCodInconsistenciasCartao(ByVal strCodIncons As String, ByVal TpCartao As String, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As CodInconsistencias

        'TODO VERIFICAR ONDE É UTILIZADO ESTA FUNCTION
        Dim rdr As SqlDataReader
        Dim CodInconsistencias As New CodInconsistencias
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodInconsistenciasCartao", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIncons", strCodIncons))
            Command.Parameters.Add(New SqlParameter("@TpCartao", TpCartao))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    CodInconsistencias.Mensagem = rdr("Mensagem")

                    CodInconsistencias.Sucesso = True
                    CodInconsistencias.TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                CodInconsistencias.Sucesso = False
                CodInconsistencias.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                CodInconsistencias.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                CodInconsistencias.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()
        Catch ex As Exception
            CodInconsistencias.Sucesso = False
            CodInconsistencias.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Descricao & ex.Message
            CodInconsistencias.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Id
            CodInconsistencias.TipoErro = DadosGenericos.TipoErro.Arquitetura
            CodInconsistencias.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(CodInconsistencias.NumErro, CodInconsistencias.MsgErro, CodInconsistencias.TipoErro, "Projeto: CodInconsistenciasBC - Classe: ConsultaCodInconsistencias - Função: BuscaCodInconsistenciasCC", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return CodInconsistencias

    End Function

    Public Function BuscaCodInconsistenciasDA(ByVal strCodIncons As String, ByVal CodBanco As String, ByRef isErro As Boolean, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As CodInconsistencias

        'TODO VERIFICAR ONDE É UTILIZADO ESTA FUNCTION
        Dim rdr As SqlDataReader
        Dim CodInconsistencias As New CodInconsistencias
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodInconsistenciasDA", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIncons", strCodIncons))
            Command.Parameters.Add(New SqlParameter("@CodBanco", CodBanco))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    CodInconsistencias.Mensagem = rdr("Mensagem")
                    isErro = IIf(rdr("isErro").Equals("S"), True, False)

                    CodInconsistencias.Sucesso = True
                    CodInconsistencias.TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                CodInconsistencias.Sucesso = False
                CodInconsistencias.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                CodInconsistencias.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                CodInconsistencias.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()
        Catch ex As Exception
            CodInconsistencias.Sucesso = False
            CodInconsistencias.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Descricao & ex.Message
            CodInconsistencias.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Id
            CodInconsistencias.TipoErro = DadosGenericos.TipoErro.Arquitetura
            CodInconsistencias.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(CodInconsistencias.NumErro, CodInconsistencias.MsgErro, CodInconsistencias.TipoErro, "Projeto: CodInconsistenciasBC - Classe: ConsultaCodInconsistencias - Função: BuscaCodInconsistenciasCC", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return CodInconsistencias

    End Function

    Public Function BuscaCodInconsistencias04(ByVal CodOcor As String, ByVal CodBanco As String, ByRef isErro As Boolean, ByVal conn As SqlConnection, ByVal trans As SqlTransaction) As CodInconsistencias

        'TODO VERIFICAR ONDE É UTILIZADO ESTA FUNCTION
        Dim rdr As SqlDataReader
        Dim CodInconsistencias As New CodInconsistencias
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BuscaCodInconsistencias04", conn)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = trans

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodOcor", CodOcor))
            Command.Parameters.Add(New SqlParameter("@CodBanco", CodBanco))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read
                    CodInconsistencias.Mensagem = rdr("Mensagem")
                    isErro = IIf(rdr("isErro").Equals("S"), True, False)

                    CodInconsistencias.Sucesso = True
                    CodInconsistencias.TipoErro = DadosGenericos.TipoErro.None
                Loop
            Else
                CodInconsistencias.Sucesso = False
                CodInconsistencias.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                CodInconsistencias.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id
                CodInconsistencias.TipoErro = DadosGenericos.TipoErro.None
            End If
            rdr.Close()
        Catch ex As Exception
            CodInconsistencias.Sucesso = False
            CodInconsistencias.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Descricao & ex.Message
            CodInconsistencias.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACODINCONSISTENCIAS.Id
            CodInconsistencias.TipoErro = DadosGenericos.TipoErro.Arquitetura
            CodInconsistencias.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(CodInconsistencias.NumErro, CodInconsistencias.MsgErro, CodInconsistencias.TipoErro, "Projeto: CodInconsistenciasBC - Classe: ConsultaCodInconsistencias - Função: BuscaCodInconsistenciasCC", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return CodInconsistencias

    End Function

End Class
