
Imports System.Data.SqlClient


Public Class AlterarBaixaContasReceber

    Public Function AlteraNumLoteNumLctoMesAnoLoteSeqLctoLoteNumUltLctoBaixaReceberParaNull(ByVal strNumLote As String,
                                                                                             ByVal strMesAnoLote As String,
                                                                                             ByVal strNumLcto As String,
                                                                                             ByRef connection As SqlConnection,
                                                                                             ByRef Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlteraNumLoteNumLctoMesAnoLoteSeqLctoLoteNumUltLctoBaixaReceberParaNull", connection)
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumLote", strNumLote))
            Command.Parameters.Add(New SqlParameter("@MesAnoLote", strMesAnoLote))
            Command.Parameters.Add(New SqlParameter("@NumLcto", strNumLcto))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERANUMLOTENUMLCTOMESANOLOTESEQLCTOLOTENUMULTLCTOBAIXARECEBERPARANULL.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERANUMLOTENUMLCTOMESANOLOTESEQLCTOLOTENUMULTLCTOBAIXARECEBERPARANULL.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BaixaContasReceberBC - Classe: AlterarBaixaContasReceber - Função: AlteraNumLoteNumLctoMesAnoLoteSeqLctoLoteNumUltLctoBaixaReceberParaNull(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function


    'Public Function BaixaCRec(ByVal strNumAviso As String,
    '                          ByVal strBanco As String,
    '                          ByVal strUsuario As String,
    '                          ByVal DtArq As DateTime,
    '                          ByVal strContas As String,
    '                          ByRef connection As SqlConnection,
    '                          ByRef Transaction As SqlTransaction) As Retorno
    '    Dim _Retorno As New Retorno
    '    Dim i As Integer = 0

    '    Try
    '        ''Informa a procedure
    '        Dim Command As SqlCommand = New SqlCommand("P_BaixaCRec", connection)
    '        command.CommandType = CommandType.StoredProcedure
    '        command.CommandTimeout = DadosGenericos.Timeout.Query
    '        Command.Transaction = Transaction

    '        ''define os parametros usados na stored procedure
    '        Command.Parameters.Add(New SqlParameter("@NumAviso", strNumAviso))
    '        Command.Parameters.Add(New SqlParameter("@BancoArq", strBanco))
    '        Command.Parameters.Add(New SqlParameter("@Usuario", strUsuario))
    '        Command.Parameters.Add(New SqlParameter("@DataArq", Format(DtArq, "yyyyMMdd")))
    '        Command.Parameters.Add(New SqlParameter("@CodAgenArq", strContas))


    '        Command.ExecuteNonQuery()

    '        _Retorno.Sucesso = True
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.None
    '    Catch ex As Exception

    '        _Retorno.Sucesso = False
    '        _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BAIXACREC.Descricao & ex.Message
    '        _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BAIXACREC.Id
    '        _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

    '        'CRIAR LOG NO WINDOWS
    '        Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BaixaContasReceberBC - Classe: AlterarBaixaContasReceber - Função: BaixaCRec(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
    '    End Try

    '    Return _Retorno

    'End Function


    Public Function BaixaCRec(ByVal strNumAviso As String,
                              ByVal strBanco As String,
                              ByVal strUsuario As String,
                              ByVal DtArq As DateTime,
                              ByVal strContas As String,
                              ByVal strAgencia As String,
                              ByRef connection As SqlConnection,
                              ByRef Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_BaixaCRec", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction
            Command.CommandTimeout = 1500

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NumAviso", strNumAviso))
            Command.Parameters.Add(New SqlParameter("@BancoArq", strBanco))
            Command.Parameters.Add(New SqlParameter("@Usuario", strUsuario))
            Command.Parameters.Add(New SqlParameter("@DataArq", Format(DtArq, "yyyyMMdd")))
            Command.Parameters.Add(New SqlParameter("@CodAgenArq", strAgencia))
            Command.Parameters.Add(New SqlParameter("@NumCtaArq", strContas))


            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BAIXACREC.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BAIXACREC.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BaixaContasReceberBC - Classe: AlterarBaixaContasReceber - Função: BaixaCRec(2)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

    Public Function AlterarBaixaContaAReceber(ByVal _BxCtaReceber As BaixaContaReceber,
                          ByRef connection As SqlConnection,
                          ByRef Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarBaixaContaAReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NUMTIT", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.NumTit), DBNull.Value, _BxCtaReceber.NumTit)))
            Command.Parameters.Add(New SqlParameter("@SEQTIT", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.SeqTit), DBNull.Value, _BxCtaReceber.SeqTit)))
            Command.Parameters.Add(New SqlParameter("@DTEMISSAO", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.DtEmissao), DBNull.Value, _BxCtaReceber.DtEmissao)))
            Command.Parameters.Add(New SqlParameter("@DTPGTO", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.DtPgto), DBNull.Value, _BxCtaReceber.DtPgto)))
            Command.Parameters.Add(New SqlParameter("@VLRPAGO", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrPago), DBNull.Value, _BxCtaReceber.VlrPago)))
            Command.Parameters.Add(New SqlParameter("@VLRJUROS", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrJuros), DBNull.Value, _BxCtaReceber.VlrJuros)))
            Command.Parameters.Add(New SqlParameter("@VLRMULTA", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrMulta), DBNull.Value, _BxCtaReceber.VlrMulta)))
            Command.Parameters.Add(New SqlParameter("@VLRDESC", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrDesc), DBNull.Value, _BxCtaReceber.VlrDesc)))
            Command.Parameters.Add(New SqlParameter("@VLRVARCAMB", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrVarCamb), DBNull.Value, _BxCtaReceber.VlrVarCamb)))
            Command.Parameters.Add(New SqlParameter("@OBSBAIXA", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.ObsBaixa), DBNull.Value, _BxCtaReceber.ObsBaixa)))
            Command.Parameters.Add(New SqlParameter("@CODEVENTO", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.CodEvento), DBNull.Value, _BxCtaReceber.CodEvento)))
            Command.Parameters.Add(New SqlParameter("@VLRABAT", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrAbat), DBNull.Value, _BxCtaReceber.VlrAbat)))
            Command.Parameters.Add(New SqlParameter("@VLRDEVOL", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrDevol), DBNull.Value, _BxCtaReceber.VlrDevol)))
            Command.Parameters.Add(New SqlParameter("@USRINC", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.UsrInc), DBNull.Value, _BxCtaReceber.UsrInc)))
            Command.Parameters.Add(New SqlParameter("@IDFILIAL", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.IdFilial), DBNull.Value, _BxCtaReceber.IdFilial)))


            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARBAIXACONTAARECEBER.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARBAIXACONTAARECEBER.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BaixaContasReceberBC - Classe: AlterarBaixaContasReceber - Função: AlterarBaixaContaAReceber(3)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

    Public Function InsereTituloDivergente(ByVal _BxCtaReceber As BaixaContaReceber,
                          ByRef connection As SqlConnection,
                          ByRef Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_AlterarBaixaContaAReceber", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@NUMTIT", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.NumTit), DBNull.Value, _BxCtaReceber.NumTit)))
            Command.Parameters.Add(New SqlParameter("@SEQTIT", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.SeqTit), DBNull.Value, _BxCtaReceber.SeqTit)))
            Command.Parameters.Add(New SqlParameter("@DTEMISSAO", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.DtEmissao), DBNull.Value, _BxCtaReceber.DtEmissao)))
            Command.Parameters.Add(New SqlParameter("@DTPGTO", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.DtPgto), DBNull.Value, _BxCtaReceber.DtPgto)))
            Command.Parameters.Add(New SqlParameter("@VLRPAGO", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrPago), DBNull.Value, _BxCtaReceber.VlrPago)))
            Command.Parameters.Add(New SqlParameter("@VLRJUROS", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrJuros), DBNull.Value, _BxCtaReceber.VlrJuros)))
            Command.Parameters.Add(New SqlParameter("@VLRMULTA", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrMulta), DBNull.Value, _BxCtaReceber.VlrMulta)))
            Command.Parameters.Add(New SqlParameter("@VLRDESC", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrDesc), DBNull.Value, _BxCtaReceber.VlrDesc)))
            Command.Parameters.Add(New SqlParameter("@VLRVARCAMB", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrVarCamb), DBNull.Value, _BxCtaReceber.VlrVarCamb)))
            Command.Parameters.Add(New SqlParameter("@OBSBAIXA", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.ObsBaixa), DBNull.Value, _BxCtaReceber.ObsBaixa)))
            Command.Parameters.Add(New SqlParameter("@CODEVENTO", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.CodEvento), DBNull.Value, _BxCtaReceber.CodEvento)))
            Command.Parameters.Add(New SqlParameter("@VLRABAT", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrAbat), DBNull.Value, _BxCtaReceber.VlrAbat)))
            Command.Parameters.Add(New SqlParameter("@VLRDEVOL", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.VlrDevol), DBNull.Value, _BxCtaReceber.VlrDevol)))
            Command.Parameters.Add(New SqlParameter("@USRINC", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.UsrInc), DBNull.Value, _BxCtaReceber.UsrInc)))
            Command.Parameters.Add(New SqlParameter("@IDFILIAL", IIf(String.IsNullOrWhiteSpace(_BxCtaReceber.IdFilial), DBNull.Value, _BxCtaReceber.IdFilial)))


            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ALTERARBAIXACONTAARECEBER.Descricao & ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARBAIXACONTAARECEBER.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BaixaContasReceberBC - Classe: AlterarBaixaContasReceber - Função: AlterarBaixaContaAReceber(3)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

    Public Function RemoveTitulosAcordosSAT(ByVal protocolo As String,
                          ByRef connection As SqlConnection,
                          ByRef Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_RemoveTitulosAcordosSAT", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Protocolo", protocolo))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARBAIXACONTAARECEBER.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BaixaContasReceberBC - Classe: AlterarBaixaContasReceber - Função: RemoveTitulosAcordosSAT", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

    Public Function RemoveTitulosAcordosSATPorNumTit(ByVal numtit As String,
                                                     ByVal seqtit As String,
                                                     ByVal dtEmissao As DateTime,
                          ByRef connection As SqlConnection,
                          ByRef Transaction As SqlTransaction) As Retorno
        Dim _Retorno As New Retorno
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_RemoveTitulosAcordosSATPorNumTitSeqTitDtEmissao", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parametros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@Numtit", numtit))
            Command.Parameters.Add(New SqlParameter("@SeqTit", seqtit))
            Command.Parameters.Add(New SqlParameter("@DtEmissao", dtEmissao))

            Command.ExecuteNonQuery()

            _Retorno.Sucesso = True
            _Retorno.TipoErro = DadosGenericos.TipoErro.None
        Catch ex As Exception

            _Retorno.Sucesso = False
            _Retorno.MsgErro = ex.Message
            _Retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ALTERARBAIXACONTAARECEBER.Id
            _Retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(_Retorno.NumErro, _Retorno.MsgErro, _Retorno.TipoErro, "Projeto: BaixaContasReceberBC - Classe: AlterarBaixaContasReceber - Função: RemoveTitulosAcordosSAT", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return _Retorno

    End Function

End Class
