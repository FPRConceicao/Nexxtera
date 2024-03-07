Imports Teleatlantic.TLS.Entidades
Imports Teleatlantic.TLS.Common

Imports System.Data
Imports System.Data.SqlClient

Public Class ConsultaRetornoBco

    Public Function BuscaControleRetBcoPorCodBancoNumAviso(ByVal strCodBanco As String,
                                                           ByVal strNumAviso As String,
                                                           ByVal strNumtit As String,
                                                           ByVal strSeqTit As String,
                                                           ByVal DtDataVcto As DateTime,
                                                           ByRef connection As SqlConnection,
                                                           ByRef Transaction As SqlTransaction) As List(Of RetornoBco) '#4#

        Dim rdr As SqlDataReader
        Dim lstControleRetBco As New List(Of RetornoBco)
        Dim i As Integer = 0

        Try
            ''Informa a procedure
            Dim Command As SqlCommand = New SqlCommand("P_ConsultaRretornoBcoPorCodbcoNumTitSeqTitDeVcto", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Transaction = Transaction

            ''define os parƒmetros usados na stored procedure
            Command.Parameters.Add(New SqlParameter("@CodBanco", strCodBanco))
            Command.Parameters.Add(New SqlParameter("@NumAviso", strNumAviso))
            Command.Parameters.Add(New SqlParameter("@Numtit", strNumtit))
            Command.Parameters.Add(New SqlParameter("@seqtit", strSeqTit))
            Command.Parameters.Add(New SqlParameter("@dtvcto", DtDataVcto))

            ''Executa a procedure
            rdr = Command.ExecuteReader()

            If rdr.HasRows Then
                Do While rdr.Read()
                    lstControleRetBco.Add(New RetornoBco)
                    lstControleRetBco(i).CodBanco = rdr("CodBanco").ToString()
                    lstControleRetBco(i).NumAviso = rdr("NumAviso").ToString()
                    lstControleRetBco(i).NumTit = rdr("NumTit").ToString()
                    lstControleRetBco(i).SeqTit = rdr("SeqTit").ToString()
                    lstControleRetBco(i).CodAgen = rdr("CodAgen").ToString()
                    lstControleRetBco(i).NumCta = rdr("NumCta").ToString()
                    lstControleRetBco(i).VlrPago = rdr("VlrPago").ToString()
                    lstControleRetBco(i).VlrJuros = rdr("VlrJuros").ToString()
                    lstControleRetBco(i).VlrDesc = rdr("VlrDesc").ToString()
                    lstControleRetBco(i).VlrIOF = rdr("VlrIOF").ToString()
                    lstControleRetBco(i).VlrAbat = rdr("VlrAbat").ToString()
                    lstControleRetBco(i).Processado = rdr("Processado").ToString()
                    lstControleRetBco(i).DtVcto = rdr("DtVcto").ToString()
                    lstControleRetBco(i).DtPagto = rdr("DtPagto").ToString()
                    lstControleRetBco(i).DtArq = rdr("DtArq").ToString()
                    lstControleRetBco(i).VlrMulta = rdr("VlrMulta").ToString()





                    lstControleRetBco(i).Sucesso = True
                    lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.None

                    i = i + 1
                Loop
            Else
                lstControleRetBco.Add(New RetornoBco)
                lstControleRetBco(i).Sucesso = True
                lstControleRetBco(i).TipoErro = DadosGenericos.TipoErro.Funcional
            End If
            rdr.Close()
        Catch ex As Exception

            lstControleRetBco.Add(New RetornoBco)
            lstControleRetBco(0).Sucesso = False
            lstControleRetBco(0).NumErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISO.Id
            lstControleRetBco(0).MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCACONTROLERETBCOPORCODBANCONUMAVISO.Descricao & ex.Message
            lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura
            lstControleRetBco(0).ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(lstControleRetBco(0).NumErro, lstControleRetBco(0).MsgErro, lstControleRetBco(0).TipoErro, "Projeto: BancoBC - Classe: ConsultaControleRetBco - Função: BuscaControleRetBcoPorCodBancoNumAviso(1)", "08", "Verisure", Environment.MachineName, "2.0", "13")
        End Try

        Return lstControleRetBco

    End Function

End Class
