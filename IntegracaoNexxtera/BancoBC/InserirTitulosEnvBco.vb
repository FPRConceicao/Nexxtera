Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.Common

Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Public Class InserirTitulosEnvBco

    Public Function InsereTitulos_Env_Bco(ByVal strCodBanco As String,
                                          ByVal strCodAgencia As String,
                                          ByVal strNumCta As String,
                                          ByVal Connection As SqlConnection,
                                          ByVal Transaction As SqlTransaction) As Retorno  '#1#

        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco(1)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function


    Public Function InsereTitulos_Env_Bco001_341_033_399_356(ByVal strCodBanco As String,
                                                             ByVal strCodAgencia As String,
                                                             ByVal strNumCta As String,
                                                             ByVal strTipoPagto As String,
                                                             ByVal Connection As SqlConnection,
                                                             ByVal Transaction As SqlTransaction) As Retorno  '#1#


        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco001", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO001_341.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO001_341.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco001_341(2)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function



    Public Function InsereTitulos_Env_Bco409(ByVal strCodBanco As String,
                                             ByVal strCodAgencia As String,
                                             ByVal strNumCta As String,
                                             ByVal strTipoPagto As String,
                                             ByVal Connection As SqlConnection,
                                             ByVal Transaction As SqlTransaction) As Retorno  '#1#
        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco409", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO409.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO409.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco409(3)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function


    Public Function InsereTitulos_Env_Bco237(ByVal strCodBanco As String,
                                            ByVal strCodAgencia As String,
                                            ByVal strNumCta As String,
                                            ByVal Connection As SqlConnection,
                                            ByVal Transaction As SqlTransaction) As Retorno  '#1#


        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco237", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure 
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO237.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO237.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco237(4)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function


    Public Function InsereTitulos_Env_Bco347_291(ByVal strCodBanco As String,
                                                 ByVal strCodAgencia As String,
                                                 ByVal strNumCta As String,
                                                 ByVal Connection As SqlConnection,
                                                 ByVal Transaction As SqlTransaction) As Retorno  '#1#

        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco347", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure 
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO347_291.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO347_291.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco347_291(5)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function


    Public Function InsereTitulos_Env_Bco_CartaoCredito_Master(ByVal Connection As SqlConnection,
                                                               ByVal Transaction As SqlTransaction) As Retorno  '#1#

        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco_CartaoCredito", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_MASTER.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_MASTER.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco_CartaoCredito_Master(6)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function



    Public Function InsereTitulos_Env_Bco_CartaoCredito_Visa(ByVal Connection As SqlConnection,
                                                             ByVal Transaction As SqlTransaction) As Retorno  '#1#

        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco_CartaoCredito_V", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_VISA.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_VISA.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco_CartaoCredito_Visa(7)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_Bco_CartaoCredito_Amex(ByVal Connection As SqlConnection,
                                                             ByVal Transaction As SqlTransaction) As Retorno  '#1#

        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco_CartaoCredito_A", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_AMEX.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_AMEX.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco_CartaoCredito_Amex(8)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_Bco_CartaoCredito_Cielo(ByVal Connection As SqlConnection,
                                                            ByVal Transaction As SqlTransaction) As Retorno  '#1#

        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco_CartaoCredito_Cielo", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_VISA.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_VISA.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco_CartaoCredito_Cielo(9)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_BcoBoletoUnificado(ByVal strCodBanco As String,
                                                                 ByVal strCodAgencia As String,
                                                                 ByVal strNumCta As String,
                                                                 ByVal strTipoPagto As String,
                                                                 ByVal Connection As SqlConnection,
                                                                 ByVal Transaction As SqlTransaction) As Retorno  '#1#


        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_BcoBoletoUnificado", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCOBOLETOUNIFICADO.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCOBOLETOUNIFICADO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_Bco041(ByVal strCodBanco As String,
                                            ByVal strCodAgencia As String,
                                            ByVal strNumCta As String,
                                            ByVal Connection As SqlConnection,
                                            ByVal Transaction As SqlTransaction) As Retorno  '#1#


        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_Env_Bco041", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure 
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO237.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO237.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco237(4)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulosEnvBcoAdyen(ByVal Connection As SqlConnection,
                                                            ByVal Transaction As SqlTransaction) As Retorno  '#1#

        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulosEnvBcoAdyen", Connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_VISA.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCO_CARTAOCREDITO_VISA.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco_CartaoCredito_Cielo(9)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Bu_Dc(ByVal strCodBanco As String,
                                                         ByVal strCodAgencia As String,
                                                         ByVal strNumCta As String,
                                                         ByVal strTipoPagto As String,
                                                         ByVal Connection As SqlConnection,
                                                         ByVal Transaction As SqlTransaction) As Retorno  '#1#


        Dim _Retorno As New Retorno
        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_InsereTitulos_BU_DC", Connection)
            Command.Transaction = Transaction
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))


            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO001_341.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO001_341.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: TitulosEnvBcoBC - Classe: InserirTitulosEnvBco - Função: InsereTitulos_Env_Bco001_341(2)", UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)

        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Bu_Dc(ByVal strCodBanco As String, ByVal strCodAgencia As String, ByVal strNumCta As String, ByVal strTipoPagto As String) As Retorno  '#1#
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_InsereTitulos_BU_DC", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))

            ''Executa a procedure
            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO001_341.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO001_341.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_Bco237(ByVal strCodBanco As String, ByVal strCodAgencia As String, ByVal strNumCta As String) As Retorno  '#1#
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_InsereTitulos_Env_Bco237", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure 
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Executa a procedure
            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO237.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO237.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_Bco001_341_033_399_356(ByVal strCodBanco As String, ByVal strCodAgencia As String, ByVal strNumCta As String, ByVal strTipoPagto As String) As Retorno  '#1#
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_InsereTitulos_Env_Bco001", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))

            ''Executa a procedure
            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO001_341.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO001_341.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_BcoBoletoUnificado(ByVal strCodBanco As String, ByVal strCodAgencia As String, ByVal strNumCta As String, ByVal strTipoPagto As String) As Retorno  '#1#
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_InsereTitulos_Env_BcoBoletoUnificado", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))
            Command.Parameters.Add(New SqlParameter("@TipoPagto", strTipoPagto))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCOBOLETOUNIFICADO.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_INSERETITULOS_ENV_BCOBOLETOUNIFICADO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            Command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_Bco(ByVal strCodBanco As String, ByVal strCodAgencia As String, ByVal strNumCta As String) As Retorno  '#1#
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_InsereTitulos_Env_Bco", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            Command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function InsereTitulos_Env_Bco041(ByVal strCodBanco As String, ByVal strCodAgencia As String, ByVal strNumCta As String) As Retorno  '#1#
        Dim _Retorno As New Retorno
        Dim connection As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_InsereTitulos_Env_Bco041", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parƒmetros usados na stored procedure 
            command.Parameters.Add(New SqlParameter("@Codbanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@CodAgen", strCodAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCta", strNumCta))

            ''Executa a procedure
            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO237.Id
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERETITULOS_ENV_BCO237.Descricao & ex.Message
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(MethodBase.GetCurrentMethod()), UsuarioGlobal.CodSetor, UsuarioGlobal.Usuario, Environment.MachineName, VariavelGlobal.strVersao, UsuarioGlobal.CodDepartamento)
        Finally
            connection.Close()
            Command.Dispose()
        End Try

        Return _Retorno
    End Function
End Class
