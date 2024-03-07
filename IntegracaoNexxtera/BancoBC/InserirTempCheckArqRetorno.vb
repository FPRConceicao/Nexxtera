Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.Common

Imports System.Data
Imports System.Data.SqlClient


Public Class InserirTempCheckArqRetorno

    Public Function InserirTempCheckArqRetorno(ByVal _TempCheckArqRetorno As TempCheckArqRetorno, ByVal Connection As SqlConnection, ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_InserirTempCheckArqRetorno", Connection)

        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", _TempCheckArqRetorno.CodBanco))
            Command.Parameters.Add(New SqlParameter("@NomeBanco  ", _TempCheckArqRetorno.NomeBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso   ", _TempCheckArqRetorno.NumAviso))
            Command.Parameters.Add(New SqlParameter("@CodAgen    ", _TempCheckArqRetorno.CodAgen))
            Command.Parameters.Add(New SqlParameter("@NumCta     ", _TempCheckArqRetorno.NumCta))
            Command.Parameters.Add(New SqlParameter("@DtArq      ", _TempCheckArqRetorno.DtArq))
            Command.Parameters.Add(New SqlParameter("@NumTit     ", _TempCheckArqRetorno.NumTit))
            Command.Parameters.Add(New SqlParameter("@DtPagto    ", IIf(IsInvalidDate(_TempCheckArqRetorno.DtPagto), DBNull.Value, _TempCheckArqRetorno.DtPagto)))
            Command.Parameters.Add(New SqlParameter("@NomeCliente", _TempCheckArqRetorno.NomeCliente))
            Command.Parameters.Add(New SqlParameter("@CodOcor    ", _TempCheckArqRetorno.CodOcor))
            Command.Parameters.Add(New SqlParameter("@CodIncons  ", _TempCheckArqRetorno.CodIncons))
            Command.Parameters.Add(New SqlParameter("@Mensagem   ", IIf(IsNothing(_TempCheckArqRetorno.Mensagem), DBNull.Value, _TempCheckArqRetorno.Mensagem)))
            Command.Parameters.Add(New SqlParameter("@CodBancoDeb", IIf(IsNothing(_TempCheckArqRetorno.CodBancoDeb), DBNull.Value, _TempCheckArqRetorno.CodBancoDeb)))
            Command.Parameters.Add(New SqlParameter("@CodAgenDeb ", IIf(IsNothing(_TempCheckArqRetorno.CodAgenDeb), DBNull.Value, _TempCheckArqRetorno.CodAgenDeb)))
            Command.Parameters.Add(New SqlParameter("@CodNumCtaDeb", _TempCheckArqRetorno.CodNumCtaDeb))
            Command.Parameters.Add(New SqlParameter("@Valor      ", _TempCheckArqRetorno.Valor))
            Command.Parameters.Add(New SqlParameter("@Endereco   ", _TempCheckArqRetorno.Endereco))
            Command.Parameters.Add(New SqlParameter("@Cidade     ", _TempCheckArqRetorno.Cidade))
            Command.Parameters.Add(New SqlParameter("@UF         ", _TempCheckArqRetorno.UF))
            Command.Parameters.Add(New SqlParameter("@CEP        ", _TempCheckArqRetorno.Cep))
            Command.Parameters.Add(New SqlParameter("@Arquivo    ", _TempCheckArqRetorno.Arquivo))
            Command.Parameters.Add(New SqlParameter("@Seq        ", IIf(IsNothing(_TempCheckArqRetorno.SeqTit), DBNull.Value, _TempCheckArqRetorno.SeqTit)))
            Command.Parameters.Add(New SqlParameter("@DtVcto    ", IIf(IsInvalidDate(_TempCheckArqRetorno.DtVcto), DBNull.Value, IIf(_TempCheckArqRetorno.DtVcto.Equals(Date.MinValue), DBNull.Value, _TempCheckArqRetorno.DtVcto))))
            Command.Parameters.Add(New SqlParameter("@NossonumeroBco    ", IIf(IsNothing(_TempCheckArqRetorno.NossoNumeroBco), DBNull.Value, _TempCheckArqRetorno.NossoNumeroBco)))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try
        Return _Retorno

    End Function

    Public Function InserirTempCheckArqRetorno(ByVal _TempCheckArqRetorno As TempCheckArqRetorno) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_InserirTempCheckArqRetornoNexxtera", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parametros usados na stored procedure
            command.Parameters.Add(New SqlParameter("@CodBanco", _TempCheckArqRetorno.CodBanco))
            command.Parameters.Add(New SqlParameter("@NomeBanco  ", _TempCheckArqRetorno.NomeBanco))
            command.Parameters.Add(New SqlParameter("@NumAviso   ", _TempCheckArqRetorno.NumAviso))
            command.Parameters.Add(New SqlParameter("@CodAgen    ", _TempCheckArqRetorno.CodAgen))
            command.Parameters.Add(New SqlParameter("@NumCta     ", _TempCheckArqRetorno.NumCta))
            command.Parameters.Add(New SqlParameter("@DtArq      ", _TempCheckArqRetorno.DtArq))
            command.Parameters.Add(New SqlParameter("@NumTit     ", _TempCheckArqRetorno.NumTit))
            command.Parameters.Add(New SqlParameter("@DtPagto    ", IIf(IsInvalidDate(_TempCheckArqRetorno.DtPagto), DBNull.Value, _TempCheckArqRetorno.DtPagto)))
            command.Parameters.Add(New SqlParameter("@NomeCliente", _TempCheckArqRetorno.NomeCliente))
            command.Parameters.Add(New SqlParameter("@CodOcor    ", _TempCheckArqRetorno.CodOcor))
            command.Parameters.Add(New SqlParameter("@CodIncons  ", _TempCheckArqRetorno.CodIncons))
            command.Parameters.Add(New SqlParameter("@Mensagem   ", IIf(IsNothing(_TempCheckArqRetorno.Mensagem), DBNull.Value, _TempCheckArqRetorno.Mensagem)))
            command.Parameters.Add(New SqlParameter("@CodBancoDeb", IIf(IsNothing(_TempCheckArqRetorno.CodBancoDeb), DBNull.Value, _TempCheckArqRetorno.CodBancoDeb)))
            command.Parameters.Add(New SqlParameter("@CodAgenDeb ", IIf(IsNothing(_TempCheckArqRetorno.CodAgenDeb), DBNull.Value, _TempCheckArqRetorno.CodAgenDeb)))
            command.Parameters.Add(New SqlParameter("@CodNumCtaDeb", _TempCheckArqRetorno.CodNumCtaDeb))
            command.Parameters.Add(New SqlParameter("@Valor      ", _TempCheckArqRetorno.Valor))
            command.Parameters.Add(New SqlParameter("@Endereco   ", _TempCheckArqRetorno.Endereco))
            command.Parameters.Add(New SqlParameter("@Cidade     ", _TempCheckArqRetorno.Cidade))
            command.Parameters.Add(New SqlParameter("@UF         ", _TempCheckArqRetorno.UF))
            command.Parameters.Add(New SqlParameter("@CEP        ", _TempCheckArqRetorno.Cep))
            command.Parameters.Add(New SqlParameter("@Arquivo    ", _TempCheckArqRetorno.Arquivo))
            command.Parameters.Add(New SqlParameter("@Seq        ", IIf(IsNothing(_TempCheckArqRetorno.SeqTit), DBNull.Value, _TempCheckArqRetorno.SeqTit)))
            command.Parameters.Add(New SqlParameter("@DtVcto    ", IIf(IsInvalidDate(_TempCheckArqRetorno.DtVcto), DBNull.Value, IIf(_TempCheckArqRetorno.DtVcto.Equals(Date.MinValue), DBNull.Value, _TempCheckArqRetorno.DtVcto))))
            command.Parameters.Add(New SqlParameter("@NossonumeroBco    ", IIf(IsNothing(_TempCheckArqRetorno.NossoNumeroBco), DBNull.Value, _TempCheckArqRetorno.NossoNumeroBco)))
            command.Parameters.Add(New SqlParameter("@VlrJuros         ", _TempCheckArqRetorno.VlrJuros))
            command.Parameters.Add(New SqlParameter("@VlrMulta         ", _TempCheckArqRetorno.VlrMulta))
            command.Parameters.Add(New SqlParameter("@TipoDup ", IIf(IsNothing(_TempCheckArqRetorno.TipoDup), DBNull.Value, _TempCheckArqRetorno.TipoDup)))
            command.Parameters.Add(New SqlParameter("@Situacao ", IIf(IsNothing(_TempCheckArqRetorno.Situacao), DBNull.Value, _TempCheckArqRetorno.Situacao)))

            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function

    Public Function InsereTempRetornoOptante(ByVal TempOptante As Optante, ByVal Connection As SqlConnection, ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_InsereTempRetornoOptante", Connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodIntClie", IIf(String.IsNullOrEmpty(TempOptante.CodIntClie), DBNull.Value, TempOptante.CodIntClie)))
            Command.Parameters.Add(New SqlParameter("@CodMov", TempOptante.CodMov))
            Command.Parameters.Add(New SqlParameter("@LinhaErroTipoB", IIf(String.IsNullOrEmpty(TempOptante.LinhaErroTipoB), DBNull.Value, TempOptante.LinhaErroTipoB)))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try
        Return _Retorno

    End Function

    Private Function IsInvalidDate(ByVal data) As Boolean
        If IsNothing(data) Then
            Return True
        ElseIf data = Date.MinValue Then
            Return True
        End If
        Return False
    End Function

    Public Function InserirChecagemClientesBancariosLog(ByVal codBanco As String, ByVal numAviso As String, ByVal numCta As String, ByVal dtArq As String, ByVal usr As String, ByVal Connection As SqlConnection, ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_InserirChecagemClientesBancario", Connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", codBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso", numAviso))
            Command.Parameters.Add(New SqlParameter("@NumCta", numCta))
            Command.Parameters.Add(New SqlParameter("@DtArq", dtArq))
            Command.Parameters.Add(New SqlParameter("@Usr", usr))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try
        Return _Retorno

    End Function

    Public Function InserirChecagemClientesBancariosCartaoCreditoLog(ByVal numAviso As String, ByVal dtArq As String, ByVal usr As String,
                                        ByVal Connection As SqlConnection,
                                        ByVal Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim Command As SqlCommand = New SqlCommand("P_InserirChecagemClientesCartaoCreditoBancario", Connection)
        Try
            ''Informa a procedure
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumAviso", numAviso))
            Command.Parameters.Add(New SqlParameter("@DtArq", dtArq))
            Command.Parameters.Add(New SqlParameter("@Usr", usr))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            _Retorno.MsgErro = ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_APAGARTITULOSENVBCO.Id
            _Retorno.Sucesso = False
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        End Try
        Return _Retorno

    End Function

    Public Function InsertCodigoBarra(cr_tb As ContaReceber) As Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim _Retorno As New Retorno
        Dim command As New SqlCommand("P_InsertCodigoBarra", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            connection.Open()

            ''define os parametros usados na stored procedure
            'command.Parameters.Add(New SqlParameter("@CodClie", cr_tb.CodClie))
            command.Parameters.Add(New SqlParameter("@NumTit", cr_tb.NumTit))
            command.Parameters.Add(New SqlParameter("@SeqTit", cr_tb.SeqTit))
            command.Parameters.Add(New SqlParameter("@LinhaDigitavel", cr_tb.CodigoBarra))

            command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_INSERTCODIGOBARRA.Descricao + ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_INSERTCODIGOBARRA.Id
            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "VERISURE", Environment.MachineName, "", "")
        Finally
            connection.Close()
            command.Dispose()
        End Try

        Return _Retorno
    End Function
End Class
