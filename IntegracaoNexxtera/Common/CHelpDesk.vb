Imports System.Data.SqlClient
Imports System.IO

Public Class CHelpDesk
    Public Shared Sub AbreOcorrência(ByVal strUsuario As String, ByVal strMensagem As String, ByVal strCaminhoPrint As String)
        'Dim blnResultado As Boolean = True
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand()
        Dim trans As SqlTransaction

        Try
            connection.Open()
            trans = connection.BeginTransaction

            'BUSCA OCORRÊNCIA REPETIDA QUE JÁ FOI GERADA AUTOMATICAMENTE, CASO NÃO EXISTA ABRE UMA NOVA
            If BuscarOcorrenciaAutomaticaRepetida(strUsuario, strMensagem, connection, command, trans) Then
                AbrirOcorrencia(strUsuario, strMensagem, connection, command, trans)
                CopiaAnexoEmail(strUsuario, strMensagem, strCaminhoPrint, connection, command, trans)
            End If

            trans.Commit()
            connection.Close()
        Catch ex As Exception
            If Not IsNothing(trans) Then trans.Rollback()
            'blnResultado = False
            Throw New Exception(ex.Message)
        Finally
            If connection.State = ConnectionState.Open Then connection.Close()
        End Try

        'Return blnResultado
    End Sub

    ''' <summary>
    ''' ABRE A OCORRENCIA
    ''' </summary>
    ''' <param name="strUsr"></param>
    ''' <param name="strObs"></param>
    ''' <param name="connection"></param>
    ''' <param name="command"></param>
    ''' <remarks></remarks>
    Private Shared Sub AbrirOcorrencia(ByVal strUsr As String, ByVal strObs As String, ByVal connection As SqlConnection, ByVal command As SqlCommand, ByVal trans As SqlTransaction)
        command = New SqlCommand("P_AbrirOcorrencia", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        command.Transaction = trans

        command.Parameters.Add(New SqlParameter("@USRCAD", IIf(String.IsNullOrEmpty(strUsr), DBNull.Value, strUsr)))
        command.Parameters.Add(New SqlParameter("@CODCHAMADO", "000605"))
        command.Parameters.Add(New SqlParameter("@OBS", IIf(String.IsNullOrEmpty(strObs), DBNull.Value, strObs)))
        command.Parameters.Add(New SqlParameter("@STATUS", "A"))
        command.Parameters.Add(New SqlParameter("@USRALT", IIf(String.IsNullOrEmpty(strUsr), DBNull.Value, strUsr)))
        command.Parameters.Add(New SqlParameter("@CODINTCLIE", DBNull.Value))
        command.Parameters.Add(New SqlParameter("@ENCAMINHADO", "N"))
        command.Parameters.Add(New SqlParameter("@ERROCONHECIDO", "N"))
        command.Parameters.Add(New SqlParameter("@REABERTURAINDEVIDA", "N"))
        command.Parameters.Add(New SqlParameter("@USRATEND", strUsr))
        command.Parameters.Add(New SqlParameter("@TIPO", "Correção"))

        command.ExecuteNonQuery()


    End Sub

    ''' <summary>
    ''' VERIFICA SE A OCORRENCIA JÁ EXISTE
    ''' </summary>
    ''' <param name="strUsr"></param>
    ''' <param name="strObs"></param>
    ''' <param name="connection"></param>
    ''' <param name="command"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function BuscarOcorrenciaAutomaticaRepetida(ByVal strUsr As String, ByVal strObs As String, ByVal connection As SqlConnection, ByVal command As SqlCommand, ByVal trans As SqlTransaction, Optional ByRef strProtocolo As String = "") As Boolean
        Dim blnResultado As Boolean = False

        Dim rdr As SqlDataReader

        command = New SqlCommand("P_BuscarOcorrenciaAutomaticaRepetida", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        command.Transaction = trans

        command.Parameters.Add(New SqlParameter("@USRCAD", IIf(String.IsNullOrEmpty(strUsr), DBNull.Value, strUsr)))
        command.Parameters.Add(New SqlParameter("@CODCHAMADO", "000605"))
        command.Parameters.Add(New SqlParameter("@OBS", IIf(String.IsNullOrEmpty(strObs), DBNull.Value, strObs)))
        command.Parameters.Add(New SqlParameter("@STATUS", "A"))
        command.Parameters.Add(New SqlParameter("@USRALT", IIf(String.IsNullOrEmpty(strUsr), DBNull.Value, strUsr)))

        rdr = command.ExecuteReader
        If rdr.HasRows Then
            While rdr.Read
                strProtocolo = rdr("Protocolo").ToString
                blnResultado = False
            End While
        Else
            blnResultado = True
        End If

        rdr.Close()

        Return blnResultado
    End Function

    ''' <summary>
    ''' COPIA O ANEXO QUE SERÁ ENVIADO COM O EMAIL PARA O SERVIDOR
    ''' </summary>
    ''' <param name="strCaminhoPrint"></param>
    ''' <param name="connection"></param>
    ''' <param name="command"></param>
    ''' <remarks></remarks>
    Private Shared Sub CopiaAnexoEmail(ByVal strUsr As String, ByVal strObs As String, ByVal strCaminhoPrint As String, ByVal connection As SqlConnection, ByVal command As SqlCommand, ByVal trans As SqlTransaction)
        'PEGA O CAMINHO PARA SALVAR O ANEXO DO ERRO
        Dim arrParametros As String() = PesquisaParametros(connection, command, trans)
        Dim strProtocolo As String = ""

        'CASO A OCORRÊNCIA TENHA SIDO ABERTA COM SUCESSO, COPIA OS ARQUIVOS
        If Not BuscarOcorrenciaAutomaticaRepetida(strUsr, strObs, connection, command, trans, strProtocolo) Then
            If UBound(arrParametros) > 0 Then
                If Not File.Exists(arrParametros(0) & strProtocolo & ".jpg") Then
                    'strAcompanhamento = vbCrLf & " Verifique se o arquivo está na pasta de origem ou remova-o da lista de anexos!"
                    File.Copy(strCaminhoPrint, arrParametros(0) & strProtocolo & ".jpg")
                Else
                    'strAcompanhamento = vbCrLf & " Falha ao salvar o anexo no servidor. Entre em contato com o Departamento de Sistemas ou remova o anexo da lista!"
                    File.SetAttributes(arrParametros(0) & strProtocolo & ".jpg", FileAttributes.Archive)
                    File.Delete(arrParametros(0) & strProtocolo & ".jpg")
                    File.Copy(strCaminhoPrint, arrParametros(0) & strProtocolo & ".jpg")
                End If

                InserirAnexo(strProtocolo, strUsr, strCaminhoPrint, arrParametros(0) & strProtocolo & ".jpg", connection, command, trans)
            End If
        End If


    End Sub

    ''' <summary>
    ''' SELECT DA PARAMETROS E ATRIBUI A UM ARRAY DE STRINGS
    ''' </summary>
    ''' <param name="connection"></param>
    ''' <param name="command"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function PesquisaParametros(ByVal connection As SqlConnection, ByVal command As SqlCommand, ByVal trans As SqlTransaction) As String()
        Dim rdr As SqlDataReader
        Dim arrParametros As String()

        command = New SqlCommand("P_PesquisaParametros", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        command.Transaction = trans

        rdr = command.ExecuteReader
        If rdr.HasRows Then
            While rdr.Read
                ReDim arrParametros(rdr.FieldCount)
                arrParametros(0) = rdr("HelpDeskPath").ToString
                arrParametros(1) = rdr("ESMTPServer").ToString
            End While
        End If

        rdr.Close()

        Return arrParametros
    End Function

    ''' <summary>
    ''' Insere os anexos na base
    ''' </summary>
    ''' <param name="connection"></param>
    ''' <param name="command"></param>
    ''' <remarks></remarks>
    Private Shared Sub InserirAnexo(ByVal strProtocolo As String, ByVal strUsr As String, ByVal strCaminhoDe As String, ByVal strCaminhoPara As String, ByVal connection As SqlConnection, ByVal command As SqlCommand, ByVal trans As SqlTransaction)
        command = New SqlCommand("P_InserirAnexo", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        command.Transaction = trans

        command.Parameters.Add(New SqlParameter("@PROTOCOLO", IIf(String.IsNullOrEmpty(strProtocolo), DBNull.Value, strProtocolo)))
        command.Parameters.Add(New SqlParameter("@CAMINHODE", IIf(String.IsNullOrEmpty(strCaminhoDe), DBNull.Value, strCaminhoDe)))
        command.Parameters.Add(New SqlParameter("@CAMINHOPARA", IIf(String.IsNullOrEmpty(strCaminhoPara), DBNull.Value, strCaminhoPara)))
        command.Parameters.Add(New SqlParameter("@USRANEXO", IIf(String.IsNullOrEmpty(strUsr), DBNull.Value, strUsr)))

        command.ExecuteNonQuery()

    End Sub

    Public Shared Sub LogarErroConhecido(ByVal usuario As String, mensagem As String, local As String, ErroConhecido As Integer, Tipo As Integer, Departamento As String, Setor As String, Computador As String, Versao As String, Base As String)
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As New SqlCommand("P_LogarErroConhecido", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query
        Try
            connection.Open()


            command.Parameters.Add(New SqlParameter("@usuario", IIf(String.IsNullOrEmpty(usuario), DBNull.Value, usuario)))
            command.Parameters.Add(New SqlParameter("@mensagem", IIf(String.IsNullOrEmpty(mensagem), DBNull.Value, mensagem)))
            command.Parameters.Add(New SqlParameter("@local", IIf(String.IsNullOrEmpty(local), DBNull.Value, local)))
            command.Parameters.Add(New SqlParameter("@ErroConhecido", IIf(String.IsNullOrEmpty(ErroConhecido), DBNull.Value, ErroConhecido)))
            command.Parameters.Add(New SqlParameter("@Tipo", IIf(String.IsNullOrEmpty(Tipo), DBNull.Value, Tipo)))
            command.Parameters.Add(New SqlParameter("@Departamento", IIf(String.IsNullOrEmpty(Departamento), DBNull.Value, Departamento)))
            command.Parameters.Add(New SqlParameter("@Setor", IIf(String.IsNullOrEmpty(Setor), DBNull.Value, Setor)))
            command.Parameters.Add(New SqlParameter("@Computador", IIf(String.IsNullOrEmpty(Computador), DBNull.Value, Computador)))
            command.Parameters.Add(New SqlParameter("@Versao", IIf(String.IsNullOrEmpty(Versao), DBNull.Value, Versao)))
            command.Parameters.Add(New SqlParameter("@Base", IIf(String.IsNullOrEmpty(Base), DBNull.Value, Base)))


            command.ExecuteNonQuery()
        Catch ex As Exception
            Throw ex
        Finally
            connection.Close()
            command.Dispose()
        End Try
    End Sub
End Class