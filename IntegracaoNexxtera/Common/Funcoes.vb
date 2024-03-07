Imports Microsoft.Win32
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Mail
Imports Telerik.WinControls
Imports System.Windows.Forms
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Reflection
Imports System.Globalization
Imports System.IO
Imports Telerik.WinControls.UI
Imports System.Data.OleDb
Imports System.Text
Imports OfficeOpenXml
Imports Renci.SshNet.Sftp
Imports Renci.SshNet
Imports ExcelApp = Microsoft.Office.Interop.Excel
Imports Teleatlantic.TLS.Common

''' <summary>
''' Classe de funções do projeto
''' </summary>
''' <remarks>
''' 
''' Data Criação:     08/04/2011
''' Auttor:           Edson Ferreira
''' 
''' Modificações: 
''' 08/04/2011
''' EDF - TL200001 - Classe de funções do projeto
''' Autor da Modificação: Edson Ferreira   
'''
''' </remarks>
Public Class Funcoes

#Region " Variáveis Globais "

    'Controladores de Formatação de Números
    Public gsSepMilhar As String = Mid(Format("1000", "#,##0.00"), 2, 1)
    Public gsSepDecimal As String = Mid(Format("10", "##0.00"), 3, 1)

    'Formatação da data
    Public Shared Property Cultura As IFormatProvider

    Public Shared gEmail As String = ""
    Public Shared gstrUnidadeDigitada As String = ""

#End Region

#Region " Funções "

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary> Valida se a procedure está sendo executada no momento da validação. </summary>
    '''
    ''' <remarks> Lucas dos S. Pieretti da Silva, 11/11/2014. </remarks>
    '''
    ''' <param name="query">Nome da query a ser verificada no Banco de Dados</param>
    ''' <param name="ativa">Boolean para indicar se a query está ativa ou não</param>
    ''' 
    ''' <returns> 
    ''' Retorna true para query ativa no momento da verificação ou false para inativa ou erro de acesso ao bd
    ''' </returns>
    ''' 
    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared Function verificaAcessoBD(ByVal query As String, ByRef ativa As Boolean) As Retorno

        Dim retorno As New Retorno
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_BloqueioRelatorio", connection)
        Dim rdr As SqlDataReader

        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        Try
            'Informa a procedure
            command.CommandType = CommandType.StoredProcedure
            command.CommandTimeout = DadosGenericos.Timeout.Query

            connection.Open()

            command.Parameters.Add(New SqlParameter("@Query", query))

            'Executa a procedure
            rdr = command.ExecuteReader()

            If rdr.HasRows Then
                ativa = True
                retorno.Sucesso = True
                retorno.TipoErro = DadosGenericos.TipoErro.None
            Else
                ativa = False
                retorno.Sucesso = True
                retorno.NumErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Id()
                retorno.MsgErro = ErrorConstants.NENHUM_REGISTRO_ENCONTRADO.Descricao
                retorno.TipoErro = DadosGenericos.TipoErro.Funcional
                retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            End If
        Catch ex As Exception

            ativa = False
            retorno.Sucesso = False
            retorno.NumErro = ErrorConstants.EXCEPTION_METODO_BUSCAESTOQUEEMPENHADO.Id
            retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_BUSCAESTOQUEEMPENHADO.Descricao & ex.Message
            retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro

            'CRIAR LOG NO WINDOWS
            Funcoes.AtualizaApplEventLog(retorno.NumErro, retorno.MsgErro, retorno.TipoErro, Funcoes.FormataLocalException(System.Reflection.MethodBase.GetCurrentMethod()), "", "", Environment.MachineName, "", "")
        Finally
            connection.Close()
        End Try

        Return retorno

    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary> Valida se a Exception foi acionada pelo servidor de e-mails. </summary>
    '''
    ''' <remarks> Renato G. Eraclide, 22/10/2014. </remarks>
    '''
    ''' <param name="ex">    Instância da Exception. </param>
    ''' <param name="email"> (Opcional) E-mail destinatário para exibir na mensagem. </param>
    ''' 
    ''' <returns> 
    ''' Retorna um objeto da classe Retorno que deve ser tratado como erro de e-mail caso retorne sucesso true. 
    ''' Conterá a mensagem de erro traduzida. 
    ''' </returns>
    ''' 
    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared Function IsErroEmail(ByVal ex As Exception,
                              Optional ByVal email As String = "") As Retorno

        Dim retorno As Retorno = RetornoFunc(ex.Message)

        retorno.Sucesso = True

        ' Adiciona espaços na variável para exibir mensagem corretamente.
        If email = "" Then
            email = " "
        Else
            email = " " & email & " "
        End If

        If (ex.Message.Contains("The specified string is not in the form required for an e-mail address")) Then
            retorno.MsgErro = "Endereço de e-mail" & email & "inválido."

        ElseIf (ex.Message.Contains("An invalid character was found in the mail header")) Then
            retorno.MsgErro = "Endereço de e-mail" & email & "inválido."

        ElseIf (ex.Message.Contains("Recipient address rejected: Domain not found")) Then
            retorno.MsgErro = "Endereço de e-mail" & email & "não existe ou está indisponível."

        ElseIf (ex.Message.Contains("Failure sending mail")) Then
            retorno.MsgErro = "Falha ao enviar e-mail."

        ElseIf (ex.Message.Contains("Unable to send to all recipients")) Then
            retorno.MsgErro = "Não foi possível enviar o e-mail para todos os endereços."

        ElseIf (ex.Message.Contains("Service not available") OrElse
                ex.Message.Contains("Error: timeout exceeded") OrElse
                ex.Message.Contains("The operation has timed out")) Then

            retorno.MsgErro = "Falha ao enviar e-mail, o tempo limite foi excedido. Tente novamente mais tarde."

        ElseIf (ex.Message.Contains("A recipient must be specified")) Then
            retorno.MsgErro = "Endereço de e-mail não informado."
        ElseIf (ex.Message.Contains("Recipient address rejected: need fully-qualified address")) Then
            retorno.MsgErro = "Endereço de e-mail" & email & "incompleto."
        Else
            ' Retorna erro de arquitetura
            retorno = RetornoFunc(ex.Message)
        End If

        Return retorno
    End Function

    'VALIDA HOME PAGE
    Public Shared Function RetornaDataFormatada(dataFormatada As String) As DateTime
        Try
            'dev:luiz gustavo de moura santos
            'tarefa:BU para títulos DC
            'data:13-03-2021
            'trecho:luiz_bu_titulo
            Dim ano = dataFormatada.Substring(6, 4)
            Dim mes = dataFormatada.Substring(3, 2)
            Dim dia = dataFormatada.Substring(0, 2)
            Dim dataFinal = ano + "-" + mes + "-" + dia + " 00:00:00"


            Dim dtFormata As DateTime = Convert.ToDateTime(dataFinal)

            Return dtFormata

        Catch ex As Exception
            Throw
        End Try

    End Function


    'Função que retorna data inicial para parâmetro de relatório
    Public Shared Function GetDtIni(ByVal dtpIni As RadDateTimePicker) As String
        Return dtpIni.Value.Year & "-" & dtpIni.Value.Month & "-" & dtpIni.Value.Day & " 00:00:00"
    End Function

    'Função que retorna data final para parâmetro de relatório
    Public Shared Function GetDtFim(ByVal dtpFim As RadDateTimePicker) As String
        Return dtpFim.Value.Year & "-" & dtpFim.Value.Month & "-" & dtpFim.Value.Day & " 23:58:50"
    End Function

    'Método que valida o Email
    Public Shared Function ValidaEmail(ByVal email As String) As Boolean
        Return System.Text.RegularExpressions.Regex.IsMatch(email, ("(?<user>[^@]+)@(?<host>.+)"))
    End Function

    Public Shared Function ExportarRelatorio(ByRef strPath As String, ByRef tipo As Integer, ByRef Report As Object) As String
        Dim retorno As String = ""

        Try
            'Verifica se existe uma instância do arquivo aberta e o fecha
            ApagaArquivo(strPath)

            'BUSCA O TIPO PARA GERAR A EXPORTACAO
            Report.ExportToDisk(tipo, strPath)


        Catch err As Exception
            retorno = err.Message
        End Try

        Return retorno
    End Function

    ''' <summary>
    ''' Valida se o usuario digitou palavra com acentos e com "ç". 
    ''' </summary>
    ''' <param name="tecla">informar o codigo ascii da letra.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function VerificaTecla(ByVal tecla As Integer) As Integer
        If Not (tecla = 59 Or tecla = 60 Or tecla = 62 Or tecla = 95 Or tecla = 64 Or tecla = 36 Or tecla = 41 Or tecla = 40 Or tecla = 47 Or tecla = 42 Or (tecla = 24 Or tecla = 22 Or tecla = 3 Or tecla = 1) Or tecla = 58 Or tecla = 45 Or tecla = 13 Or tecla = 8 Or tecla = 44 Or tecla = 46 Or tecla = 32 Or tecla = 38 Or (tecla >= 65 And tecla <= 90) Or (tecla >= 48 And tecla <= 57)) Then
            tecla = 0
        ElseIf tecla = 13 Then
            tecla = 13
        End If

        Return tecla
    End Function

    Public Shared Function ValidaTecla(ByVal tecla As Integer) As Boolean
        Dim blnR As Boolean = False
        If Not (tecla = 59 Or tecla = 60 Or tecla = 62 Or tecla = 95 Or tecla = 64 Or tecla = 36 Or tecla = 41 Or tecla = 40 Or tecla = 47 Or tecla = 42 Or (tecla = 24 Or tecla = 22 Or tecla = 3 Or tecla = 1) Or tecla = 58 Or tecla = 45 Or tecla = 13 Or tecla = 8 Or tecla = 44 Or tecla = 46 Or tecla = 32 Or tecla = 38 Or (tecla >= 65 And tecla <= 90) Or (tecla >= 48 And tecla <= 57)) Then
            blnR = True
        ElseIf tecla = 13 Then
            blnR = False
        End If

        Return blnR
    End Function

    ''' <summary>
    ''' Remove acentos e "ç"
    ''' </summary>
    ''' <param name="strTexto">Informar o texto.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function RetiraAcento(ByVal strTexto As String)
        Dim strLetras As String = "Á|À|Â|Ã|É|È|Ê|Í|Ì|Ó|Ò|Ô|Õ|Ú|Ù|Ü|á|à|â|ã|é|è|ê|í|ì|ó|ò|ô|õ|ú|ù|ü"
        Dim strLetraTroca As String = "A|A|A|A|E|E|E|I|I|O|O|O|O|U|U|U|a|a|a|a|e|e|e|i|i|o|o|o|o|u|u|u"
        Dim arrLetra = strLetras.Split("|")
        Dim arrLetraTroca = strLetraTroca.Split("|")
        Dim strTextoNovo As String = ""
        Dim blnEncontrado As Boolean = False

        For i As Integer = 0 To strTexto.Length - 1
            For j As Integer = 0 To UBound(arrLetra) - 1
                If strTexto(i) = arrLetra(j) Then
                    strTextoNovo = Strings.UCase(strTextoNovo & arrLetraTroca(j))
                    blnEncontrado = True
                End If
            Next
            If Not blnEncontrado Then
                strTextoNovo = Strings.UCase(strTextoNovo & strTexto(i))
                blnEncontrado = False
            Else
                blnEncontrado = False
            End If
        Next
        Return strTextoNovo
    End Function

    ''' <summary>
    ''' Retorno o nome do mes referente o numero correspondente.
    ''' 01-Janeiro
    ''' 02-Fevereiro
    ''' 03-Março
    ''' 04-Abril
    ''' 05-Maio
    ''' 06-Junho
    ''' 07-Julho
    ''' 08-Agosto
    ''' 09-Setembro
    ''' 10-Outubro
    ''' 11-Novembro
    ''' 12-Dezembro
    ''' </summary>
    ''' <param name="mes">Informar o numero do mes.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DescMes(ByVal mes As String) As String

        Select Case mes
            Case "01"
                DescMes = "Janeiro"
            Case "02"
                DescMes = "Fevereiro"
            Case "03"
                DescMes = "Março"
            Case "04"
                DescMes = "Abril"
            Case "05"
                DescMes = "Maio"
            Case "06"
                DescMes = "Junho"
            Case "07"
                DescMes = "Julho"
            Case "08"
                DescMes = "Agosto"
            Case "09"
                DescMes = "Setembro"
            Case "10"
                DescMes = "Outubro"
            Case "11"
                DescMes = "Novembro"
            Case "12"
                DescMes = "Dezembro"
            Case Else
                DescMes = ""
        End Select

    End Function

    ''' <summary>
    ''' Rotina para Envio de email.
    ''' </summary>
    ''' <param name="De">O email remetente do envio.</param>
    ''' <param name="Para">Endereço do email de envio.
    ''' No caso de envio para várias pessoas ao mesmo tempo é necessário separar por virgula.
    ''' </param>
    ''' <param name="BCC">Endereço de email para cópia oculta.
    ''' No caso de envio para várias pessoas ao mesmo tempo é necessário separar por virgula.
    ''' </param>
    ''' <param name="CC">Endereço de email para cópia.
    ''' No caso de envio para várias pessoas ao mesmo tempo é necessário separar por virgula.
    ''' </param>
    ''' <param name="Assunto">Assunto do email</param>
    ''' <param name="CorpoDoEmail">Conteúdo do email.
    ''' Permite texto.
    ''' Permite tag html.
    ''' </param>
    ''' <param name="IsHtml">Parametro opcional, de padrão sempre falso, para envio de mensagem em html.</param>
    ''' <param name="arrAnexo">Parametro opcional. Vetor de string com os endereços dos arquivos a ser anexados.</param>
    ''' <remarks></remarks>
    Public Shared Function EnviarMensagemEmail(ByRef strPath As [String],
                                                ByRef strSmtp As [String],
                                                ByRef De As [String],
                                                ByRef Para As [String],
                                                ByRef BCC As [String],
                                                ByRef CC As [String],
                                                ByRef Assunto As [String],
                                                ByRef CorpoDoEmail As [String],
                                                Optional ByRef IsHtml As Boolean = False,
                                                Optional ByRef arrAnexo() As String = Nothing,
                                                Optional ByVal strDisplayNameDE As String = "",
                                                Optional ByVal strDisplayNamePARA As String = "",
                                                Optional ByVal blnIsLog As Boolean = True) As String

        ' Evita travamento na rede Verisure.
        'If IsRedeVerisure() Then Return "true"

        Dim strMensagem As String = ""
        If strSmtp.ToUpper() = "WEBSERVICEMOBILEPRONTO" Then
            'FICOU DEFINIDO QUE PARA CADA "PARA", "CC" E "BCC" SERÁ ENVIADO UM E-MAIL SEPARADO UTILIZANDO A WEBSERVICE DA MOBILEPRONTO.
            Try
                'PARA
                EnvialEmailWebServiceMobilePronto("F41B5EA90553DB2FC61AE59A6DF108DE93D0EFD8", "83fE5d", Assunto, De, CorpoDoEmail, "HTML", "S", Para, "", "", "")
                'CC
                EnvialEmailWebServiceMobilePronto("F41B5EA90553DB2FC61AE59A6DF108DE93D0EFD8", "83fE5d", Assunto, De, CorpoDoEmail, "HTML", "S", CC, "", "", "")
                'BCC
                EnvialEmailWebServiceMobilePronto("F41B5EA90553DB2FC61AE59A6DF108DE93D0EFD8", "83fE5d", Assunto, De, CorpoDoEmail, "HTML", "S", BCC, "", "", "")

                If blnIsLog Then CriaLog(strPath & "\Emails.log", "Email enviado com sucesso.  (" & Format(PegaData, "dd/MM/yyyy hh:mm") & "); SMTP Server: " & strSmtp & "; Remetente: " & De & "; Destinatário: " & Para & "; Assunto: " & Assunto)

                strMensagem = "true"

            Catch ex As Exception
                strMensagem = "false"
                If blnIsLog Then CriaLog(strPath & "\Emails.log", "Falha no envio.  (" & Format(PegaData, "dd/MM/yyyy hh:mm") & "); SMTP Server: " & strSmtp & "; Remetente: " & De & "; Destinatário: " & Para & "; Assunto: " & Assunto & vbLf & ex.Message.ToString)
                'Throw ex
            End Try

        Else


            Dim mMailMessage As New MailMessage()
            Dim mItem As Attachment
            Dim limpaAnexo As Boolean = False

            'ENDERECO DE EMAIL PARA ENVIO
            If strDisplayNameDE = "" Then
                strDisplayNameDE = strDisplayNameDE
            End If

            If String.IsNullOrEmpty(De) Then
                mMailMessage.From = New MailAddress("verisure@verisure.com.br", strDisplayNameDE)
            Else
                mMailMessage.From = New MailAddress(De, strDisplayNameDE)
            End If
            'mMailMessage.From = New MailAddress(De, strDisplayNameDE)

            'REMETENTE DO EMAIL
            'If strDisplayNamePARA = "" Then
            '    strDisplayNamePARA = strDisplayNamePARA
            'End If

            'mMailMessage.[To].Add(New MailAddress(Para, strDisplayNamePARA))

            'ALTERAÇÃO NO ENVIO DE EMAIL COM SPLIT 11/04/2013 - FERNANDO
            If Para <> "" Then
                Dim strEmail() As String = Para.Split(";")
                Dim strPara() As String = strDisplayNamePARA.Split(";")
                If UBound(strEmail) = UBound(strPara) Then
                    For i As Integer = 0 To strEmail.Count - 1
                        mMailMessage.[To].Add(New MailAddress(strEmail(i), strPara(i)))
                    Next
                Else
                    For i As Integer = 0 To strEmail.Count - 1
                        mMailMessage.[To].Add(New MailAddress(strEmail(i)))
                    Next
                End If
            End If

            'ENDERECO DE EMAIL COM COPIA OCULTA
            If BCC <> "" Then
                Dim strEmail() As String = BCC.Split(";")
                For Each copia In strEmail
                    mMailMessage.Bcc.Add(New MailAddress(copia))
                Next
            End If

            'ENDERECO DE EMAIL COM COPIA
            If CC <> "" Then
                Dim strEmail() As String = CC.Split(";")
                For Each copia In strEmail
                    mMailMessage.CC.Add(New MailAddress(copia))
                Next
            End If

            'ASSUNDO DO EMAIL
            mMailMessage.Subject = Assunto

            'MENSAGEM DO CORPO DO ENVIO
            mMailMessage.Body = CorpoDoEmail

            'PERCORRE O VETOR ADICIONANDO OS ANEXOS
            If Not IsNothing(arrAnexo) Then
                For i As Integer = 0 To UBound(arrAnexo)
                    mItem = New Attachment(Trim(arrAnexo(i)))
                    mItem.ContentId = i
                    mMailMessage.Attachments.Add(mItem)
                    limpaAnexo = True
                Next
            End If

            'PERMITE HTML NO CORPO DO EMAIL
            If IsHtml Then
                mMailMessage.IsBodyHtml = True
            End If

            'PRIORIDADE DE ENVIO
            'mMailMessage.Priority = MailPriority.Normal

            'PASSA O SERVIDOR DE ENVIO
            Dim smtp As New SmtpClient(strSmtp)


            If (Trim(strSmtp) = "192.168.1.13") Then
                smtp.Credentials = New NetworkCredential("teleatlantic-sp", "tele@tele12", "teleatlantic")
            End If

            If (Trim(strSmtp) = "smtplw.com.br") Then
                smtp.Port = 587
                'smtp.EnableSsl = True
                'smtp.UseDefaultCredentials = False
                smtp.Credentials = New NetworkCredential("ricardoverisure", "verisure1smtp")
            End If

            If (Trim(strSmtp) = "email-smtp.sa-east-1.amazonaws.com") Then
                smtp.Port = 587
                smtp.EnableSsl = True
                smtp.Credentials = New System.Net.NetworkCredential("AKIAWRNT532ZBP2BKQGS", "BJ5w4jHFT6wliGsfeUYZ98N/7Qv2A6YyFrwmgw6cdTV/")
            End If




            If blnIsLog Then
                Try
                    smtp.Send(mMailMessage)
                    CriaLog(strPath & "\Emails.log", "Email enviado com sucesso.  (" & Format(PegaData, "dd/MM/yyyy hh:mm") & "); SMTP Server: " & strSmtp & "; Remetente: " & De & "; Destinatário: " & Para & "; Assunto: " & Assunto)
                    smtp.Dispose()
                    If limpaAnexo Then
                        mItem.Dispose()
                    End If

                    strMensagem = "true"

                Catch ex As Exception
                    strMensagem = "false"
                    CriaLog(strPath & "\Emails.log", "Falha no envio.  (" & Format(PegaData, "dd/MM/yyyy hh:mm") & "); SMTP Server: " & strSmtp & "; Remetente: " & De & "; Destinatário: " & Para & "; Assunto: " & Assunto & vbLf & ex.Message.ToString)
                    Throw New Exception(ex.Message.ToString())
                End Try
            Else
                Try
                    strMensagem = "true"
                    smtp.Send(mMailMessage)
                Catch ex As Exception
                    strMensagem = "false"
                    Throw ex
                End Try
            End If

            '07/08/2017 - Fernando
            'Dá um dispose nos attachments para evitar erros
            mMailMessage.Attachments.Dispose()
            mMailMessage.Dispose()

        End If
        Return strMensagem
    End Function

    Public Shared Function fValidaTel(ByVal pNumero As String,
                                      Optional ByVal Operadora As String = "") As Retorno
        Dim iCon As Integer
        Dim cValido As String
        Dim _Retorno As New Retorno


        pNumero = Trim(pNumero)

        _Retorno.Sucesso = True

        If Len(Trim(pNumero)) <> 10 Then
            If Len(Trim(pNumero)) = 11 And Operadora = "Nextel" Then
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.INFORME_TELEFONE.Descricao & "Operadora Nextel contem 8 digitos."
                _Retorno.NumErro = ErrorConstants.INFORME_TELEFONE.Id
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            ElseIf Len(Trim(pNumero)) < 10 Then
                _Retorno.Sucesso = False
                _Retorno.MsgErro = ErrorConstants.INFORME_TELEFONE.Descricao
                _Retorno.NumErro = ErrorConstants.INFORME_TELEFONE.Id
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
        End If

        If Mid(pNumero, 1, 1) = 0 Then
            _Retorno.Sucesso = False
            _Retorno.NumErro = ErrorConstants.INFORME_TELEFONE.Id
            _Retorno.MsgErro = ErrorConstants.INFORME_TELEFONE.Descricao
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
        End If

        For iCon = 1 To Len(Trim(pNumero))
            cValido = (Mid(pNumero, iCon, 1))
            If Not (LCase(cValido) Like "[0-9]") Then
                _Retorno.Sucesso = False
                _Retorno.NumErro = ErrorConstants.CAMPONUMERICO.Id
                _Retorno.MsgErro = ErrorConstants.CAMPONUMERICO.Descricao
                _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
                _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            End If
        Next

        Return _Retorno

    End Function

    ''' <summary>
    ''' Valida Operadora.
    ''' Valida se a operadora é nextel.
    ''' </summary>
    ''' <param name="strNumero"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ValidaTelefoneNextel(ByVal strNumero As String) As Boolean
        Dim strPrefixo As String
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Função que verifica se o número é Nextel
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ValidaTelefoneNextel = False
        strPrefixo = Mid(strNumero, 1, 4)

        Select Case strPrefixo
            'NEXTEL
            Case "7000", "7001", "7002", "7003", "7004", "7005", "7006", "7007", "7008", "7009", "7010", "7701", "7702", "7703", "7704", "7705", "7706", "7707", "7708", "7709", "7710", "7711", "7712", "7713", "7714", "7715", "7716", "7717", "7718", "7719", "7720", "7721", "7722", "7723", "7724", "7725", "7726", "7727", "7728", "7729", "7730", "7731", "7732", "7733", "7734", "7735", "7736", "7737"
                ValidaTelefoneNextel = True
                'NEXTEL
            Case "7738", "7739", "7740", "7741", "7742", "7743", "7744", "7745", "7746", "7747", "7748", "7749", "7750", "7751", "7752", "7753", "7754", "7755", "7756", "7757", "7758", "7759", "7760", "7761", "7762", "7763", "7764", "7765", "7766", "7767", "7768", "7769", "7770", "7771", "7772", "7773", "7774", "7775", "7776", "7777", "7778", "7779", "7780", "7781", "7782", "7783", "7784", "7785"
                ValidaTelefoneNextel = True
                'NEXTEL
            Case "7786", "7787", "7788", "7789", "7790", "7791", "7792", "7793", "7794", "7795", "7796", "7797", "7798", "7799", "7800", "7802", "7803", "7804", "7805", "7806", "7807", "7808", "7809", "7810", "7811", "7812", "7813", "7814", "7815", "7816", "7817", "7818", "7819", "7820", "7821", "7822", "7823", "7824", "7825", "7826", "7827", "7828", "7829", "7830", "7831", "7832", "7833", "7834"
                ValidaTelefoneNextel = True
                'NEXTEL
            Case "7835", "7836", "7837", "7838", "7839", "7840", "7841", "7842", "7843", "7844", "7845", "7846", "7847", "7848", "7849", "7850", "7851", "7852", "7853", "7854", "7855", "7856", "7857", "7858", "7859", "7860", "7861", "7862", "7863", "7864", "7865", "7866", "7867", "7868", "7869", "7870", "7871", "7872", "7873", "7874", "7875", "7876", "7877", "7878", "7879", "7880", "7881", "7882"
                ValidaTelefoneNextel = True
                'NEXTEL
            Case "7883", "7884", "7885", "7886", "7887", "7888", "7889", "7890", "7891", "7892", "7893", "7894", "7895", "7896", "7897", "7898", "7899", "7901", "7902", "7904", "7912", "7913", "7914", "7915", "7916", "7917", "7918", "7919", "7920", "7923", "7924", "7928", "7929", "7930", "7931", "7932", "7934", "7935", "7936", "7937", "7938", "7939", "7940", "7941", "7942", "7943", "7944", "7945", "7946", "7947", "7948", "7949"
                ValidaTelefoneNextel = True
                'NEXTEL (SEGUNDO Pedido do cadastro de um cliente específico - D66E)
            Case "7910"
                ValidaTelefoneNextel = True
                'OUTROS
            Case Else
                ValidaTelefoneNextel = False
        End Select
    End Function

    Public Shared Function IsCelular(ByVal strNumero As String, ByVal IsExcluiNextel As Boolean) As Boolean
        Dim strPrefixo As String
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Função que verifica se é celular
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        IsCelular = False
        strPrefixo = Mid(strNumero, 1, 4)

        Select Case Left(strNumero, 1)
            Case "5"
                Select Case Left(strNumero, 2)
                    Case "53", "54", "57"
                        IsCelular = True
                    Case Else
                        IsCelular = False
                End Select
            Case "6", "8", "9"
                IsCelular = True
            Case "7"
                If IsExcluiNextel = True Then
                    Select Case strPrefixo
                        'NEXTEL
                        Case "7000", "7001", "7002", "7003", "7004", "7005", "7006", "7007", "7008", "7009", "7010", "7701", "7702", "7703", "7704", "7705", "7706", "7707", "7708", "7709", "7710", "7711", "7712", "7713", "7714", "7715", "7716", "7717", "7718", "7719", "7720", "7721", "7722", "7723", "7724", "7725", "7726", "7727", "7728", "7729", "7730", "7731", "7732", "7733", "7734", "7735", "7736", "7737"
                            IsCelular = False
                            'NEXTEL
                        Case "7738", "7739", "7740", "7741", "7742", "7743", "7744", "7745", "7746", "7747", "7748", "7749", "7750", "7751", "7752", "7753", "7754", "7755", "7756", "7757", "7758", "7759", "7760", "7761", "7762", "7763", "7764", "7765", "7766", "7767", "7768", "7769", "7770", "7771", "7772", "7773", "7774", "7775", "7776", "7777", "7778", "7779", "7780", "7781", "7782", "7783", "7784", "7785"
                            IsCelular = False
                            'NEXTEL
                        Case "7786", "7787", "7788", "7789", "7790", "7791", "7792", "7793", "7794", "7795", "7796", "7797", "7798", "7799", "7800", "7802", "7803", "7804", "7805", "7806", "7807", "7808", "7809", "7810", "7811", "7812", "7813", "7814", "7815", "7816", "7817", "7818", "7819", "7820", "7821", "7822", "7823", "7824", "7825", "7826", "7827", "7828", "7829", "7830", "7831", "7832", "7833", "7834"
                            IsCelular = False
                            'NEXTEL
                        Case "7835", "7836", "7837", "7838", "7839", "7840", "7841", "7842", "7843", "7844", "7845", "7846", "7847", "7848", "7849", "7850", "7851", "7852", "7853", "7854", "7855", "7856", "7857", "7858", "7859", "7860", "7861", "7862", "7863", "7864", "7865", "7866", "7867", "7868", "7869", "7870", "7871", "7872", "7873", "7874", "7875", "7876", "7877", "7878", "7879", "7880", "7881", "7882"
                            IsCelular = False
                            'NEXTEL
                        Case "7883", "7884", "7885", "7886", "7887", "7888", "7889", "7890", "7891", "7892", "7893", "7894", "7895", "7896", "7897", "7898", "7899", "7901", "7902", "7904", "7912", "7913", "7914", "7915", "7916", "7917", "7918", "7919", "7920", "7923", "7924", "7928", "7929", "7930", "7931", "7932", "7934", "7935", "7936", "7937", "7938", "7939", "7940", "7941", "7942", "7943", "7944", "7945", "7946", "7947", "7948", "7949"
                            IsCelular = False
                            'NEXTEL (SEGUNDO Pedido do cadastro de um cliente específico - D66E)
                        Case "7910"
                            IsCelular = False
                            'OUTROS
                        Case Else
                            IsCelular = True
                    End Select
                Else
                    IsCelular = True
                End If
            Case Else
                IsCelular = False
        End Select

    End Function

    ''' <summary>
    ''' Verifica arquivo de atualizacao.
    ''' </summary>
    ''' <param name="strServer"></param>
    ''' <param name="strLocal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function VerificaArquivoAtualizacao(ByVal strServer As String, ByVal strLocal As String) As Boolean
        Dim blnAtualiza As Boolean = False

        Try
            'Endereço para teste do Atualiza 19/06/2020 - João
            'Dim strDiretorioServer As String = "\\br08vsysrep01v\TeleSystem2\Atualizacoes"
            'strDiretorioServer = strDiretorioServer & "\Telesystem2"

            'SELECIONA O DIRETORIO DO SERVIDOR
            Dim strDiretorioServer As String = strServer & "\Telesystem2"

            'ATUALIZA O PROGRAMA DE ATUALIZACAO
            If Not System.IO.File.Exists(strLocal & "\" & "TelesystemAtualiza.exe") Then
                System.IO.File.Copy(strDiretorioServer & "\" & "TelesystemAtualiza.exe", strLocal & "\" & "TelesystemAtualiza.exe")
                blnAtualiza = True
            Else
                If System.IO.File.GetLastWriteTime(strLocal & "\" & "TelesystemAtualiza.exe") < System.IO.File.GetLastWriteTime(strDiretorioServer & "\" & "TelesystemAtualiza.exe") Then
                    System.IO.File.Delete(strLocal & "\" & "TelesystemAtualiza.exe")
                    System.IO.File.Copy(strDiretorioServer & "\" & "TelesystemAtualiza.exe", strLocal & "\" & "TelesystemAtualiza.exe")
                    blnAtualiza = True
                End If
            End If
            '----------------------------------------------------------------------------------------------------------------------

            'Dim DiretorioServer As New System.IO.DirectoryInfo(strDiretorioServer)
            ''Executa função GetFile(Lista os arquivos desejados de acordo com o parametro)
            'Dim Arquivos As System.IO.FileInfo() = DiretorioServer.GetFiles("*.*")

            ''PERCORRE A LISTA DE ARQUIVO
            'For Each fileinfo As System.IO.FileInfo In Arquivos
            '    If fileinfo.Name.ToUpper() = "THUMBS.DB" Then
            '        Continue For
            '    End If

            '    'VERIFICA SE O ARQUIVO DO SERVIDOR É MAIS RECENTE QUE O DA PASTA LOCAL
            '    If Not System.IO.File.Exists(strLocal & "\" & fileinfo.Name) Then
            '        blnAtualiza = True
            '        Exit For
            '    Else
            '        If System.IO.File.GetLastWriteTime(strLocal & "\" & fileinfo.Name) < System.IO.File.GetLastWriteTime(strDiretorioServer & "\" & fileinfo.Name) Then
            '            blnAtualiza = True
            '            Exit For
            '        End If
            '    End If
            'Next

        Catch ex As Exception
            Throw ex
        End Try

        Return blnAtualiza
    End Function

    'Public Shared Function EnviarSMSClaro(ByVal strNumFone As String, ByVal strMsg As String, ByVal strProfile As String, ByVal strPwd As String, ByVal strMode As String, ByVal strURLClaro As String) As Retorno

    '    Dim myReg As Net.HttpWebRequest
    '    Dim myResp As Net.HttpWebResponse
    '    Dim _retorno As New Retorno

    '    Dim strReq As String = String.Empty

    '    'Monta Url para Envio URL_ENVIO_SMS_CLARO
    '    strReq = strURLClaro & "?" & _
    '    "profile=" & strProfile & _
    '    "&pwd=" & strPwd & _
    '    "&mode=" & strMode & _
    '    "&BNUM=" & strNumFone & _
    '    "&TEXT=" & strMsg
    '    Try
    '        myReg = DirectCast(Net.WebRequest.Create(strReq), Net.HttpWebRequest)
    '        myResp = DirectCast(myReg.GetResponse(), Net.HttpWebResponse)
    '        'Dim myStream As IO.Stream = myResp.GetResponseStream()
    '        'Dim myreader As New IO.StreamReader(myStream)

    '        _retorno.MsgErro = ErrorConstants.OK.Descricao
    '        _retorno.NumErro = ErrorConstants.OK.Id
    '        _retorno.Sucesso = True
    '        _retorno.TipoErro = DadosGenericos.TipoErro.None


    '        Return _retorno
    '    Catch ex As Exception
    '        _retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ENVIARSMSCLARO.Descricao & ex.Message
    '        _retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ENVIARSMSCLARO.Id
    '        _retorno.Sucesso = False
    '        _retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
    '        _retorno.ImagemErro = DadosGenericos.ImagemRetorno..Erro

    '        myResp.Close()


    '        Return _retorno
    '    End Try


    'End Function

    Public Shared Function EnviarSMS(ByVal strMobile55 As String,
                                     ByVal strCredencial As String,
                                     ByVal strPrincipal_User As String,
                                     ByVal Aux_User As String,
                                     ByVal strSend_Project As String,
                                     ByVal strURL_MB As String,
                                     ByVal strMessage As String,
                                     ByVal Token As String) As Retorno

        '' MODIF. Felipe SH
        strMobile55 = strMobile55.Replace(" ", "").Replace("-", "").Trim()
        Dim myReg As Net.HttpWebRequest = Nothing
        Dim myResp As Net.HttpWebResponse = Nothing
        Dim _retorno As New Retorno

        '' Dim strReq As String = String.Empty

        '' CallBackEnvioSMS.aspx?IDHASH=257805C7504349DB90F3503DCF14A382
        '' https://cliente.teleatlantic.com.br/TeleatlanticMP/CallBackEnvioSMS.aspx?IDHASH=28CB19A623964B8BA0B73D871FC4AACC

        'strReq = strURL_MB & "?" &
        '"Credencial=" & strCredencial &
        '"&Token=" & Token &
        '"&Principal_User=" & strPrincipal_User &
        '"&Aux_User=" & Aux_User &
        '"&Mobile=55" & strMobile55 &
        '"&Send_Project=" & strSend_Project &
        '"&Message=" & "https://cliente.teleatlantic.com.br/TeleatlanticMP/CallBackEnvioSMS.aspx?IDHASH=" & IdHash
        ''"&Message=" & strMessage

        Dim sbReq As New StringBuilder
        With sbReq
            .Append(strURL_MB & "?")
            .Append("Credencial=" & strCredencial)
            .Append("&Token=" & Token)
            .Append("&Principal_User=" & strPrincipal_User)
            .Append("&Aux_User=" & Aux_User)
            .Append("&Mobile=55" & strMobile55)
            .Append("&Send_Project=" & strSend_Project)
            .Append("&Message=" & strMessage)
        End With

        Try
            Dim HttpWReq As HttpWebRequest = CType(WebRequest.Create(sbReq.ToString()), HttpWebRequest)
            Dim HttpWResp As HttpWebResponse = CType(HttpWReq.GetResponse(), HttpWebResponse)
            ' Insert code that uses the response object.
            HttpWResp.Close()

            'myReg = DirectCast(Net.WebRequest.Create(strReq), Net.HttpWebRequest)
            'myResp = DirectCast(myReg.GetResponse(), Net.HttpWebResponse)
            'Dim myStream As IO.Stream = myResp.GetResponseStream()
            'Dim myreader As New IO.StreamReader(myStream)

            _retorno.MsgErro = ErrorConstants.OK.Descricao
            _retorno.NumErro = ErrorConstants.OK.Id
            _retorno.Sucesso = True
            _retorno.TipoErro = DadosGenericos.TipoErro.None

            Return _retorno

        Catch ex As Exception
            _retorno.MsgErro = ErrorConstants.EXCEPTION_METODO_ENVIARSMSTIMOINEXTEL.Descricao & ex.Message
            _retorno.NumErro = ErrorConstants.EXCEPTION_METODO_ENVIARSMSCLARO.Id
            _retorno.Sucesso = False
            _retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta

            'CRIAR LOG NO WINDOWS
            'Funcoes.AtualizaApplEventLog(_retorno.NumErro, _retorno.MsgErro, _retorno.TipoErro, "Projeto: Common - Classe: Funcoes - ConsultaDadosManutencao - Função: EnviarSMS()", "", "", Environment.MachineName, "", "")

            Return _retorno
        End Try

    End Function

    ''' <summary>
    ''' Lê um arquivo texto.
    ''' </summary>
    ''' <param name="StartupPath">Informar o endereço do arquivo.</param>
    ''' <returns>Uma string com os dados do arquivo.</returns>
    ''' <remarks></remarks>
    Public Shared Function LeArquivoTXT(ByVal StartupPath As String) As String
        Dim strTexto As String = ""
        'VERIFICA SE O ARQUIVO EXISTE
        If IO.File.Exists(StartupPath) Then

            'ABRE O ARQUIVO E CARREGA A VARIAVEL
            Dim fluxoTexto As IO.StreamReader = New IO.StreamReader(StartupPath, System.Text.Encoding.Default)
            strTexto = fluxoTexto.ReadToEnd
            fluxoTexto.Close()
        End If

        Return strTexto.ToString
    End Function

    ''' <summary>
    ''' Gera um log na maquina do remetente com os dados enviados por email
    ''' </summary>
    ''' <param name="StartupPath"></param>
    ''' <param name="strMensagem"></param>
    ''' <remarks></remarks>
    Public Shared Sub CriaLog(ByVal StartupPath As String, ByVal strMensagem As String)
        Try
            Dim strEndereco As String = StartupPath
            Dim strTexto As String = ""

            'VERIFICA SE O ARQUIVO EXISTE
            If IO.File.Exists(strEndereco) Then

                'ABRE O ARQUIVO E CARREGA A VARIAVEL
                Dim fluxoTexto As IO.StreamReader = New IO.StreamReader(strEndereco)
                strTexto = fluxoTexto.ReadToEnd
                fluxoTexto.Close()
            End If

            'CRIA INSTANCIA PRA ESCRITA
            Dim sw As IO.StreamWriter
            sw = New IO.StreamWriter(strEndereco)

            'ESCREVE O ARQUIVO
            sw.WriteLine(strMensagem)
            sw.WriteLine("****************************************************************************************************")
            sw.Write(strTexto)

            'FECHA O ARQUIVO 
            sw.Close()
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    ''' <summary>
    ''' Monta o codigo Clie com o numero de caracteres desejado para inserir.
    ''' </summary>
    ''' <param name="strCodClie">O codigo</param>
    ''' <param name="LengtnDesejado">O numero de caracter desejado</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AjustaCodClie(ByVal strCodClie As String, ByVal LengtnDesejado As Integer) As String
        Dim CodSeq As String = ""

        For i As Integer = 0 To (LengtnDesejado - strCodClie.Length) - 1
            CodSeq = CodSeq & "0"
        Next
        CodSeq = CodSeq & strCodClie

        Return CodSeq
    End Function

    ''' <summary>
    ''' Rotina de descriptografia.
    ''' </summary>
    ''' <param name="sStr">O valor que deseja descriptografar.</param>
    ''' <returns>Retorna um String com o valor criptografado.</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:     08/04/2011
    ''' Auttor:           Edson Ferreira
    ''' 
    ''' Modificações: 
    ''' 08/04/2011
    ''' WAF - TL200001 - Rotina de descriptografia.
    ''' Autor da Modificação: Edson Ferreira  
    ''' 
    ''' </remarks>
    Public Shared Function Decript(ByVal sStr As String) As String

        Dim iTam As Integer
        Dim iFor As Integer
        Dim sAux As String
        sAux = String.Empty

        If LTrim(sStr) = "" Then
            Decript = ""
            Exit Function
        End If

        iTam = Len(sStr)

        'Retira o número 133 da String criptografada.
        For iFor = 1 To iTam
            sAux = sAux & Format(Asc(Mid(sStr, iFor, 1)) - 133, "0#")
        Next

        'Se for par, retira o último dígito da string.
        Select Case Len(sAux)
            Case 4, 10, 16, 22, 28, 34, 40, 46
                sAux = Left(sAux, Len(sAux) - 1)
        End Select

        'Monta a string descriptografada com Chr() da criptografada
        'a cada 3 dígitos.
        iTam = Len(sAux)
        sStr = ""
        For iFor = 1 To iTam Step 3
            sStr = sStr & Chr(Mid(sAux, iFor, 3))
        Next

        Decript = sStr

    End Function

    ''' <summary>
    ''' Adiciona zero a direita do numero passado e a quantidade desejada
    ''' </summary>
    ''' <param name="sNumero">Informe o numero no formato string</param>
    ''' <param name="iRepeat">informe a quantidade de repetição no formato de inteiro - integer</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StrZero(ByVal sNumero As String, ByVal iRepeat As Integer)
        Dim iCount As Integer, sRetorno As String

        sRetorno = Trim(sNumero)

        For iCount = 1 To iRepeat - 1
            sRetorno = "0" & sRetorno
        Next

        StrZero = Right(sRetorno, iRepeat)

    End Function

    ''' <summary>
    ''' Rotina de criptografia.
    ''' </summary>
    ''' <param name="sStr">O valor que deseja criptografar.</param>
    ''' <returns>Retorna um String com o valor criptografado.</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:     08/04/2011
    ''' Auttor:           Edson Ferreira
    ''' 
    ''' Modificações: 
    ''' 08/04/2011
    ''' WAF - TL200001 - Rotina de criptografia.
    ''' Autor da Modificação: Edson Ferreira 
    ''' 
    ''' </remarks>
    Public Shared Function Cripto(ByVal sStr As String) As String

        Dim iTam As Integer
        Dim iFor As Integer
        Dim sAux As String
        Dim iMod As Integer
        sAux = String.Empty

        Cripto = sStr

        If LTrim(sStr) = "" Then
            Cripto = ""
            Exit Function
        End If

        iTam = Len(sStr)
        'Devolve em 3 dígitos, o Asc() de cada letra.
        For iFor = 1 To iTam
            sAux = sAux & Format(Asc(Mid(sStr, iFor, 1)), "0##")
        Next

        'Se o tamanho da string conseguida for ímpar, acrescenta
        'o número 3 à string.
        sStr = ""
        iMod = Len(sAux) Mod 2
        sAux = IIf(iMod = 0, sAux, sAux & "3")
        iTam = Len(sAux)

        'Monta uma string criptografada com os caracteres Chr() de cada
        '2 dígitos da String auxiliar somados ao número 133.
        For iFor = 1 To iTam Step 2
            sStr = sStr & Chr(Mid(sAux, iFor, 2) + 133)
        Next

        Cripto = sStr
    End Function

    ''' <summary>
    ''' Cria ou Altera o registro do windows
    ''' </summary>
    ''' <param name="Pasta">Nome da pasta que deseja ler o registro.</param>
    ''' <param name="Registro">Registro que deseja ler.</param>
    ''' <param name="valor">Valor padrão, caso o Registro não esteja criado.</param>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:     08/04/2011
    ''' Auttor:           Wolney Alexandre Fernandes
    ''' 
    ''' Modificações: 
    ''' 08/04/2011
    ''' WAF - TL200001 - Cria ou Altera o registro do windows 
    ''' Autor da Modificação: Wolney Alexandre Fernandes
    ''' 
    ''' </remarks>
    Public Shared Function CriarEAlterarRegistroDoWindows(ByVal pasta As String, ByVal Registro As String, ByVal Valor As String) As Boolean
        Try
            Dim rk As RegistryKey


            'VERIFICA SE A CRIACAO OU ALTERACAO DO REGISTRO, VAI SER DO DIRETORIO RAIZ DO TELESYSTEM 2 NO REGISTRO DO WINDOWS 
            'OU SERÁ EM ALGUMA PASTA DENTRO DO DIRETORIO RAIZ
            If pasta.ToUpper = "" And pasta.ToUpper = "TELESYSTEM 2" Then
                ' cria uma referêcnia para a chave de registro Software
                rk = Registry.CurrentUser.OpenSubKey("Software", True)

                ' cria um Subchave como o nome Telesystem 2
                rk = rk.CreateSubKey("Telesystem 2")

                'grava o caminho na SubChave Idioma
                rk.SetValue(Registro, Valor)

                ' fecha a Chave de Restistro registro
                rk.Close()
            Else

                ' cria uma referêcnia para a chave de registro Software
                rk = Registry.CurrentUser.OpenSubKey("Software\Telesystem 2", True)
                ' cria um Subchave como o nome Telesystem 2

                rk = rk.CreateSubKey(pasta)

                'grava o caminho na SubChave Idioma
                rk.SetValue(Registro, Valor)

                ' fecha a Chave de Restistro registro
                rk.Close()
            End If

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Cria o Diretório Telesystem 2
    ''' </summary>
    ''' <remarks>
    ''' 
    ''' Data Criação:     08/04/2011
    ''' Auttor:           Wolney Alexandre Fernandes
    ''' 
    ''' Modificações: 
    ''' 08/04/2011
    ''' WAF - TL200001 - Cria o Diretório Telesystem 2
    ''' Autor da Modificação: Wolney Alexandre Fernandes
    ''' 
    ''' </remarks>
    Public Shared Sub CriaDiretorioRaz()

        ' cria uma referêcnia para a chave de registro Software
        Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey("Software", True)
        ' cria um Subchave como o nome Telesystem 2

        rk = rk.CreateSubKey("Telesystem 2")

        ' fecha a Chave de Restistro registro
        rk.Close()
    End Sub

    ''' <summary>
    ''' Lê o registro do Windows
    ''' </summary>
    ''' <param name="Pasta">Nome da pasta que deseja ler o registro.</param>
    ''' <param name="Registro">Registro que deseja ler.</param>
    ''' <param name="valor">Valor padrão, caso o Registro não esteja criado.</param>
    ''' <returns>Retorna um String com o valor do Arquivo</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:     08/04/2011
    ''' Auttor:           Wolney Alexandre Fernandes
    ''' 
    ''' Modificações: 
    ''' 08/04/2011
    ''' WAF - TL200001 - Cria o Diretório Telesystem 2
    ''' Autor da Modificação: Wolney Alexandre Fernandes 
    ''' 
    ''' </remarks>
    Public Shared Function LerRegistroDoWindows(ByVal Pasta As String, ByVal Registro As String, ByVal valor As String) As String
        Dim strLerRegistro As String = ""
        Dim rk As RegistryKey

        Try

            'VERIFICA SE A LEITURA VAI SER RAIZ DO TELESYSTEM 2 OU SERÁ EM ALGUMA PASTA DENTRO DO DIRETORIO RAIZ
            If Pasta.ToUpper <> "" And Pasta.ToUpper <> "TELESYSTEM 2" Then

                ' cria uma referêcnia para a chave de registro Software
                rk = Registry.CurrentUser.OpenSubKey("Software\Telesystem 2", True)
                ' realiza a leitura do registro
                strLerRegistro = rk.OpenSubKey(Pasta, True).GetValue(Registro).ToString()
            Else

                ' cria uma referêcnia para a chave de registro Software
                rk = Registry.CurrentUser.OpenSubKey("Software", True)
                ' realiza a leitura do registro
                strLerRegistro = rk.OpenSubKey("Telesystem 2", True).GetValue(Registro).ToString()
            End If

        Catch
            'SE NÃO CONSEGUIR ENCONTRAR A PASTA, É CRIADO O REGISTRO COM O VALOR PADRÃO INFORMADO
            If CriarEAlterarRegistroDoWindows(Pasta, Registro, valor) Then

                'VERIFICA SE A LEITURA VAI SER RAIZ DO TELESYSTEM 2 OU SERÁ EM ALGUMA PASTA DENTRO DO DIRETORIO RAIZ
                If Pasta.ToUpper <> "" And Pasta.ToUpper <> "TELESYSTEM 2" Then

                    ' cria uma referêcnia para a chave de registro Software
                    rk = Registry.CurrentUser.OpenSubKey("Software\Telesystem 2", True)
                    ' realiza a leitura do registro
                    strLerRegistro = rk.OpenSubKey(Pasta, True).GetValue(Registro).ToString()
                Else

                    ' cria uma referêcnia para a chave de registro Software
                    rk = Registry.CurrentUser.OpenSubKey("Software", True)
                    ' realiza a leitura do registro
                    strLerRegistro = rk.OpenSubKey("Telesystem 2", True).GetValue(Registro).ToString()
                End If
            Else
                CriaDiretorioRaz()
                Return LerRegistroDoWindows(Pasta, Registro, valor)
            End If
        End Try
        Return strLerRegistro
    End Function

    ''' <summary>
    ''' Valida a digitatação de um campo  unidade
    ''' permitindo a digitação apenas de caracteres válidos 
    ''' </summary>
    ''' <remarks>
    ''' 
    ''' Data Criação:     27/04/2011
    ''' Auttor:           Edson Ferreira
    ''' 
    ''' Modificações: 
    ''' 
    ''' </remarks>
    Public Shared Function ValidaCampoUnidade(ByVal CodUnidade As String) As Boolean

        If (Microsoft.VisualBasic.Asc(CodUnidade) > 64 And Microsoft.VisualBasic.Asc(CodUnidade) < 91) Or (Microsoft.VisualBasic.Asc(CodUnidade) > 96 And Microsoft.VisualBasic.Asc(CodUnidade) < 123) Or (Microsoft.VisualBasic.Asc(CodUnidade) > 47 And Microsoft.VisualBasic.Asc(CodUnidade) < 58) Then
            Return True
        End If

        Return False

    End Function

    ''' <summary>
    ''' permitindo a digitação apenas de números
    ''' </summary>
    ''' <remarks>
    ''' 
    ''' Data Criação:     27/04/2011
    ''' Auttor:           Edson Ferreira
    ''' 
    ''' Modificações: 
    ''' 
    ''' </remarks>
    Public Shared Function ValidaCampoNumero(ByVal charDigito As Char) As Boolean

        If (Microsoft.VisualBasic.Asc(charDigito) > 47 And Microsoft.VisualBasic.Asc(charDigito) < 58) Or Microsoft.VisualBasic.Asc(charDigito) = 8 Or (Microsoft.VisualBasic.Asc(charDigito) > 95 And Microsoft.VisualBasic.Asc(charDigito) < 106) Then
            Return True
        End If

        Return False

    End Function

    ''' <summary>
    ''' Pega Data do Servidor
    ''' </summary>
    ''' <returns>tipo date</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:     04/05/2011
    ''' Auttor:           Edson Ferreira
    ''' 
    ''' </remarks>
    Public Shared Function PegaData() As DateTime
        Dim data As DateTime
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Try
            Dim Command As SqlCommand = New SqlCommand("Select getDate() As Data", connection)
            Command.CommandType = CommandType.Text

            connection.Open()
            Using rdr As SqlDataReader = Command.ExecuteReader()

                If rdr.HasRows Then

                    rdr.Read()
                    data = Convert.ToDateTime(rdr.Item("Data"), Funcoes.Cultura)
                End If
                rdr.Close()
            End Using
        Catch ex As Exception
            Throw
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return data
    End Function

    ''' <summary>Pega dia da semana</summary>
    ''' <returns>Dia da semana (Integer)</returns>
    ''' <remarks>
    ''' Data Criação:     24/06/2014
    ''' Autor:            Renato Eraclide
    ''' </remarks>
    Public Shared Function PegaDataSemana() As Integer
        Dim dia As Integer
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Try
            Dim Command As SqlCommand = New SqlCommand("Select DATEPART(dw, getDate()) As DiaSemana", connection)
            Command.CommandType = CommandType.Text

            connection.Open()
            Using rdr As SqlDataReader = Command.ExecuteReader()

                If rdr.HasRows Then

                    rdr.Read()
                    dia = Integer.Parse(rdr.Item("DiaSemana"))
                End If
                rdr.Close()
            End Using
        Catch ex As Exception
            Throw
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return dia
    End Function

    ''' <summary>
    ''' Pega dia da semana (data específica)
    ''' </summary>
    ''' <returns>tipo date</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:     24/06/2014
    ''' Auttor:           Renato Eraclide
    ''' 
    ''' </remarks>
    Public Shared Function PegaDataSemana(ByVal data As Date) As Integer

        Dim dia As Integer
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)

        Try

            Dim Command As SqlCommand = New SqlCommand("P_DiaDaSemana", connection)
            Command.CommandType = CommandType.StoredProcedure

            Command.Parameters.Add(New SqlParameter("@Data", data))

            connection.Open()

            Using rdr As SqlDataReader = Command.ExecuteReader()

                If rdr.HasRows Then

                    rdr.Read()
                    dia = Integer.Parse(rdr.Item("DiaSemana"))
                End If
                rdr.Close()
            End Using
        Catch ex As Exception
            Throw
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return dia
    End Function

    Public Shared Function CalcDiaUtilDoMes(ByVal dtMes As DateTime, ByVal intDias As Integer) As DateTime
        Dim dtUtil As DateTime
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("SELECT dbo.F_CalcDiaUtilDoMes ('" & Format(dtMes, "yyyy-MM-dd") &
                                                   "', " & intDias & ")", connection)
        command.CommandType = CommandType.Text
        Try
            connection.Open()
            Using rdr As SqlDataReader = command.ExecuteReader()
                If rdr.HasRows Then
                    rdr.Read()
                    dtUtil = Convert.ToDateTime(rdr(0))
                End If
                rdr.Close()
            End Using
        Catch ex As Exception
            Throw ex
        End Try


        Return dtUtil
    End Function

    Public Shared Function CalcDtUtilSabDomFer(ByVal dtDtInicial As DateTime, ByVal intDias As Integer,
                                               Optional ByVal intSab As Integer = 0, Optional ByVal intDom As Integer = 0,
                                               Optional ByVal intFer As Integer = 0) As DateTime
        Dim dtUtil As DateTime
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("SELECT dbo.F_CalcDtUtilSabDomFer ('" & Format(dtDtInicial, "yyyy-MM-dd") &
                                                   "', " & intDias & ", " & intSab & ", " & intDom & ", " & intFer & ")", connection)
        command.CommandType = CommandType.Text
        Try
            connection.Open()
            Using rdr As SqlDataReader = command.ExecuteReader()
                If rdr.HasRows Then
                    rdr.Read()
                    dtUtil = Convert.ToDateTime(rdr(0))
                End If
                rdr.Close()
            End Using
        Catch ex As Exception
            Throw ex
        End Try


        Return dtUtil
    End Function

    Public Shared Function Data_Util(ByVal dtDtInicial As DateTime) As DateTime
        Dim dtUtil As DateTime
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("SELECT dbo.data_Util ('" & Format(dtDtInicial, "yyyy-MM-dd HH:mm") & "')", connection)
        command.CommandType = CommandType.Text
        Try
            connection.Open()
            Using rdr As SqlDataReader = command.ExecuteReader()
                If rdr.HasRows Then
                    rdr.Read()
                    dtUtil = Convert.ToDateTime(rdr(0))
                End If
                rdr.Close()
            End Using
        Catch ex As Exception
            Throw ex
        End Try


        Return dtUtil
    End Function

    Public Shared Function Calendario(ByVal dtData As DateTime) As DateTime
        Dim data As DateTime
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Try
            Dim Command As SqlCommand = New SqlCommand("P_Calendario", connection)
            Command.CommandType = CommandType.StoredProcedure

            Command.Parameters.Add(New SqlParameter("@getdate", dtData))

            connection.Open()
            Using rdr As SqlDataReader = Command.ExecuteReader()

                If rdr.HasRows Then

                    rdr.Read()
                    data = rdr.Item(0)
                End If
                rdr.Close()
            End Using
        Catch ex As Exception
            Throw
        Finally
            connection.Close()
            connection.Dispose()
        End Try

        Return data
    End Function

    ''' <summary>
    ''' Valida o número do cartão de crédito
    ''' </summary>
    ''' <param name="bandeira">Passar um string com o nome da Bandeira</param>
    ''' <param name="Numero">Passar um string com o número do Cartão de Crédito</param>
    ''' <returns>Retorna uma Verdadeiro ou Falso</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:    04/05/2011
    ''' Auttor:           Wolney Alexandre Fernandes 
    ''' 
    ''' Modificações: 
    ''' 
    ''' </remarks>
    Public Shared Function ValidaCCredito(ByVal bandeira As String, ByVal Numero As String) As Boolean
        If Numero <> "" Then
            Select Case UCase(bandeira)
                Case "MASTERCARD", "DINERS"
                    Dim i As Integer, n As String, isPar As Boolean, p As Integer, R As Integer, dv As Integer
                    dv = 0
                    'Percorre o numero do cartao numero por numero menos o ultimo digito
                    For i = 1 To Len(Numero) - 1
                        'Pega o digito
                        n = Mid(Numero, i, 1)
                        'Zera o resultado
                        R = 0
                        'Verifica se posicao do digito eh par
                        isPar = IIf(i Mod 2 = 0, True, False)
                        'Se par multiplica por 1 senao 2
                        n = n * IIf(isPar, 1, 2)
                        'Soma parcelas do calculo do digito
                        For p = 1 To Len(CStr(n))
                            R = R + CInt(Mid(CStr(n), p, 1))
                        Next
                        'Acumula o resultado
                        dv = dv + R
                    Next

                    'DV eh 10 menos o resto do acumulado por 10
                    dv = 10 - (dv Mod 10)
                    'Caso DV seja 10 o DV eh zero
                    dv = IIf(dv = 10, 0, dv)

                    ValidaCCredito = IIf((dv = Mid(Numero, Len(Numero), 1)), True, False)

                Case "VISA"
                    ValidaCCredito = IIf(Mid(Numero, 1, 1) = "4", True, False)

                Case "AMEX"
                    ValidaCCredito = True

                Case "ELO"
                    ValidaCCredito = True

                Case Else
                    ValidaCCredito = False
            End Select
        Else
            ValidaCCredito = False
        End If
    End Function

    Private Shared Function InstanciaListaRetorno(ByVal lstRetorno As List(Of Retorno)) As List(Of Retorno)
        If (IsNothing(lstRetorno)) Then
            lstRetorno = New List(Of Retorno)
        End If

        Return lstRetorno
    End Function

    ''' <summary>
    ''' Monta uma lista contendo os erros enviado nos parametros.
    ''' </summary>
    ''' <param name="lstRetorno">Uma lista do tipo Retorno</param>
    ''' <param name="MsgErro">A mensagem do erro. (String) </param>
    ''' <param name="NumErro">O número do erro (String)</param>
    ''' <param name="Sucesso">O resulta se houve sucesso no procedimento (Boolean)</param>
    ''' <param name="TipoErro">O tipo de erro (String)</param>
    ''' <param name="Imagem">O nome da imagem (String)</param>
    ''' <returns>Uma lista de Retorno</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:    11/05/2011
    ''' Auttor:           Wolney Alexandre Fernandes  
    ''' 
    ''' </remarks>
    Public Shared Function CriaRetorno(ByRef lstRetorno As List(Of Retorno), ByVal MsgErro As String, ByVal NumErro As String, ByVal Sucesso As Boolean, ByVal TipoErro As String, ByVal Imagem As String) As List(Of Retorno)

        lstRetorno = Funcoes.InstanciaListaRetorno(lstRetorno)

        Dim _Retorno As New Retorno()
        _Retorno.MsgErro = MsgErro
        _Retorno.NumErro = NumErro
        _Retorno.Sucesso = Sucesso
        _Retorno.TipoErro = TipoErro
        _Retorno.ImagemErro = Imagem

        lstRetorno.Add(_Retorno)

        Return lstRetorno

    End Function

    ''' <summary>
    ''' Monta uma lista contendo os erros do objeto Retorno
    ''' </summary>
    ''' <param name="lstRetorno">Uma lista do tipo Retorno</param>
    ''' <param name="_Retorno">Objeto Retorno</param>
    ''' <returns>uma lista incluindo o objeto do tipo retorno </returns>
    ''' <remarks>
    '''       Data criação: 20/03/2013
    '''       Autor: Wolney Alexandre Fernandes / Eduardo Lacerda
    ''' </remarks>
    Public Shared Function CriaRetorno(ByRef lstRetorno As List(Of Retorno), ByVal _Retorno As Retorno) As List(Of Retorno)

        Return Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

    End Function

    ''' <summary>
    ''' Monta uma lista contendo os erros do objeto Retorno
    ''' </summary>
    ''' <param name="lstRetorno">Uma lista do tipo Retorno</param>
    ''' <param name="lstretorno">Objeto Retorno</param>
    ''' <returns>uma lista incluindo o objeto do tipo retorno </returns>
    ''' <remarks>
    '''       Data criação: 23/04/2013
    '''       Autor: Wolney Alexandre Fernandes / Fernando
    ''' </remarks>
    Public Shared Function CriaRetorno(ByRef lstRetorno As List(Of Retorno), ByVal _Erro As ErrorConstants, ByVal _Tipo As DadosGenericos.TipoErro) As List(Of Retorno)

        Dim _Retorno As New Retorno(_Erro, False, _Tipo, Retorno.SelecionaImagemRetorno(_Tipo))

        Return Funcoes.CriaRetorno(lstRetorno, _Retorno)

    End Function

    ''' <summary>
    ''' Rotina que valida CPF e CNPJ.
    ''' </summary>
    ''' <param name="Vl_CgcCpf">Recebe o CPF ou CNPJ. (String)</param>
    ''' <returns>Retorna um valor booliano.</returns>
    ''' <remarks>
    ''' 
    ''' Data Criação:    11/05/2011
    ''' Auttor:           Wolney Alexandre Fernandes   
    ''' 
    ''' </remarks>
    Public Shared Function FValidaCgcCpf(ByVal Vl_CgcCpf As String) As Boolean

        'Verifica se CGC ou CPF é Válido;

        Dim Soma As Integer
        Dim rEsto As Integer
        Dim i As Integer
        Dim a, j, D1, D2

        FValidaCgcCpf = False

        Vl_CgcCpf = Trim(Vl_CgcCpf)

        If Len(Vl_CgcCpf) > 11 Then

            'Validar CGC;

            If Len(Vl_CgcCpf) = 14 And Val(Vl_CgcCpf) > 0 Then
                a = 0
                i = 0
                D1 = 0
                D2 = 0
                j = 5
                For i = 1 To 12 Step 1
                    a = a + (Val(Mid(Vl_CgcCpf, i, 1)) * j)
                    j = IIf(j > 2, j - 1, 9)
                Next i
                a = a Mod 11
                D1 = IIf(a > 1, 11 - a, 0)
                a = 0
                i = 0
                j = 6
                For i = 1 To 13 Step 1
                    a = a + (Val(Mid(Vl_CgcCpf, i, 1)) * j)
                    j = IIf(j > 2, j - 1, 9)
                Next i
                a = a Mod 11
                D2 = IIf(a > 1, 11 - a, 0)
                If (D1 = Val(Mid(Vl_CgcCpf, 13, 1)) And D2 = Val(Mid(Vl_CgcCpf, 14, 1))) Then
                Else
                    Exit Function
                End If
            Else
                Exit Function
            End If


        Else

            'Validar CPF;
            If Vl_CgcCpf = "00000000000" Or Vl_CgcCpf = "11111111111" Then
                FValidaCgcCpf = False
                Exit Function
            End If

            'Valida argumento
            If Len(Vl_CgcCpf) <> 11 Then
                Exit Function
            End If
            Soma = 0
            For i = 1 To 9
                Soma = Soma + Val(Mid$(Vl_CgcCpf, i, 1)) * (11 - i)
            Next i

            rEsto = 11 - (Soma - (Int(Soma / 11) * 11))
            If rEsto = 10 Or rEsto = 11 Then
                rEsto = 0
            End If

            If rEsto <> Val(Mid$(Vl_CgcCpf, 10, 1)) Then
                Exit Function
            End If

            Soma = 0
            For i = 1 To 10
                Soma = Soma + Val(Mid$(Vl_CgcCpf, i, 1)) * (12 - i)
            Next i

            rEsto = 11 - (Soma - (Int(Soma / 11) * 11))
            If rEsto = 10 Or rEsto = 11 Then rEsto = 0
            If rEsto <> Val(Mid$(Vl_CgcCpf, 11, 1)) Then
                Exit Function
            End If
        End If

        FValidaCgcCpf = True

    End Function

    'Define a data de reajuste sempre um mes apos a lib do monit.
    Public Shared Function fDataReajuste(ByVal sDate As DateTime)

        On Error GoTo Erro

        Dim dDate As Date
        dDate = CDate(sDate)

        fDataReajuste = CStr(Format(DateAdd("m", 1, dDate), "yyyy-MM-dd"))

        Exit Function
Erro:
        MsgBox("Problemas na geração da data de reajuste !")

    End Function

    ''' <summary>
    ''' Valida IE- Inscrição Stadual
    ''' </summary>
    ''' <param name="pInscr">Informa a IE</param>
    ''' <param name="pUF">Informa o UF - Estado</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function FValidaInscEst(ByVal pInscr As String, ByVal pUF As String) As Boolean

        'Define variaveis
        Dim strBase As String
        Dim strBase2 As String
        Dim strOrigem As String
        Dim strDigito1 As String
        Dim strDigito2 As String
        Dim intPos As Integer
        Dim intValor As Integer
        Dim intSoma As Integer
        Dim intResto As Integer
        Dim intNumero As Integer
        Dim intPeso As Integer
        Dim intDig As Integer
        Dim d01 As Integer, d02 As Integer, d03 As Integer, d04 As Integer, d05 As Integer, d06 As Integer
        Dim d07 As Integer, d08 As Integer, d09 As Integer, d10 As Integer, d11 As Integer, d12 As Integer
        Dim d13 As Integer, dv01 As Integer, dv02 As Integer, ds As Integer, aux1 As Integer, aux2 As Integer
        Dim dfinal As Integer, digverificador As Integer, digverificador1 As Integer, digverificador2 As Integer
        Dim resto_do_calculo As Integer

        'Inicializa variaveis
        FValidaInscEst = False
        strBase = ""
        strBase2 = ""
        strOrigem = ""

        'Se for ISENTO ou EXTERIOR considera válido
        If Trim(pInscr.ToString.ToUpper) = "ISENTO" Or Trim(pInscr.ToString.ToUpper) = "ISENTA" Or Trim(pInscr.ToString.ToUpper) = "EX" Or Trim(pInscr.ToString.ToUpper) = "INDISPONIVEL" Then
            FValidaInscEst = True
            Exit Function
        End If

        'Limpa caracteres inválidos
        For intPos = 1 To Len(Trim(pInscr))
            If InStr(1, "0123456789P", Mid(pInscr, intPos, 1), vbTextCompare) > 0 Then
                strOrigem = strOrigem & Mid(pInscr, intPos, 1)
            End If
        Next

        'Busca regra por estado
        Select Case pUF
            Case "AC" 'Acre
                strBase = Left(Trim(strOrigem) & "000000000", 13)
                d01 = CInt(Mid(strBase, 1, 1))
                d02 = CInt(Mid(strBase, 2, 1))
                d03 = CInt(Mid(strBase, 3, 1))
                d04 = CInt(Mid(strBase, 4, 1))
                d05 = CInt(Mid(strBase, 5, 1))
                d06 = CInt(Mid(strBase, 6, 1))
                d07 = CInt(Mid(strBase, 7, 1))
                d08 = CInt(Mid(strBase, 8, 1))
                d09 = CInt(Mid(strBase, 9, 1))
                d10 = CInt(Mid(strBase, 10, 1))
                d11 = CInt(Mid(strBase, 11, 1))
                dv01 = CInt(Mid(strBase, 12, 1))
                dv02 = CInt(Mid(strBase, 13, 1))
                If d01 <> 0 Or d02 <> 1 Then
                    FValidaInscEst = False
                    Exit Function
                End If
                ds = 4 * d01 + 3 * d02 + 2 * d03 + 9 * d04 + 8 * d05 + 7 * d06 + 6 * d07 + 5 * d08 +
                                   4 * d09 + 3 * d10 + 2 * d11
                aux1 = Fix(ds / 11)
                aux1 = aux1 * 11
                aux2 = ds - aux1 ' aux2 é o resto, ou mod
                digverificador1 = 11 - aux2
                If digverificador1 = 10 Or digverificador1 = 11 Then
                    digverificador1 = 0 'primeiro digito
                End If
                ds = 5 * d01 + 4 * d02 + 3 * d03 + 2 * d04 + 9 * d05 + 8 * d06 + 7 * d07 + 6 * d08 +
                                   5 * d09 + 4 * d10 + 3 * d11 + 2 * digverificador1
                aux1 = Fix(ds / 11)
                aux1 = aux1 * 11
                aux2 = ds - aux1 ' aux2 é o resto, ou mod
                digverificador2 = 11 - aux2
                If digverificador2 = 10 Or digverificador2 = 11 Then
                    digverificador2 = 0 'primeiro digito
                End If
                If digverificador1 = dv01 And digverificador2 = dv02 Then
                    FValidaInscEst = True
                End If

            Case "AL" ' Alagoas
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                If Left(strBase, 2) = "24" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intSoma = intSoma * 10
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto = 10, "0", CStr(intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "AM" ' Amazonas
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                intSoma = 0
                For intPos = 1 To 8
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                If intSoma < 11 Then
                    strDigito1 = Right(CStr(11 - intSoma), 1)
                Else
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                End If
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "AP" ' Amapa
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                intPeso = 0
                intDig = 0
                If Left(strBase, 2) = "03" Then
                    intNumero = Fix(Left(strBase, 8))
                    If intNumero >= 3000001 And intNumero <= 3017000 Then
                        intPeso = 5
                        intDig = 0
                    ElseIf intNumero >= 3017001 And intNumero <= 3019022 Then
                        intPeso = 9
                        intDig = 1
                    ElseIf intNumero >= 3019023 Then
                        intPeso = 0
                        intDig = 0
                    End If
                    intSoma = intPeso
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    intValor = 11 - intResto
                    If intValor = 10 Then
                        intValor = 0
                    ElseIf intValor = 11 Then
                        intValor = intDig
                    End If
                    strDigito1 = Right(CStr(intValor), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "BA" ' Bahia
                'BAHIA - Se for com 9 digitos
                If strOrigem.Length > 8 And strOrigem.Length < 10 Then
                    Dim result As Integer

                    strBase = Left(Trim(strOrigem) & "000000000", 9)
                    If InStr(1, "0123458", Mid(strBase, 2, 1), vbTextCompare) > 0 Then
                        intSoma = 0
                        For intPos = 1 To 7
                            intValor = CInt(Mid(strBase, intPos, 1))
                            intValor = intValor * (9 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        '
                        intResto = intSoma Mod 10
                        strDigito2 = Right(IIf(intResto = 0, "0", CStr(10 - intResto)), 1)
                        strBase2 = Left(strBase, 7) & strDigito2
                        intSoma = 0
                        For intPos = 1 To 8
                            intValor = CInt(Mid(strBase2, intPos, 1))
                            intValor = intValor * (10 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 10
                        strDigito1 = Right(IIf(intResto = 0, "0", CStr(10 - intResto)), 1)
                    Else
                        intSoma = 0
                        For intPos = 1 To 7
                            intValor = CInt(Mid(strBase, intPos, 1))
                            intValor = intValor * (9 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 11
                        strDigito2 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                        strBase2 = Left(strBase, 7) & strDigito2
                        intSoma = 0
                        For intPos = 1 To 8
                            intValor = CInt(Mid(strBase2, intPos, 1))
                            intValor = intValor * (10 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 11
                        strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                    End If
                    If strOrigem.Length < 9 Then
                        strBase2 = Left(strBase, 6) & strDigito1 & strDigito2
                        If strBase2 = strOrigem Then
                            FValidaInscEst = True
                        End If

                    Else
                        strBase2 = Left(strBase2, 7) & strDigito1 & strDigito2
                        If strBase2 = strOrigem Then
                            FValidaInscEst = True
                        End If

                    End If

                    'BAHIA - caso o número seja = 8
                ElseIf strBase.Length = 8 Then
                    strBase = Left(Trim(strOrigem) & "00000000", 8)
                    If InStr(1, "0123458", Left(strBase, 1), vbTextCompare) > 0 Then
                        intSoma = 0
                        For intPos = 1 To 6
                            intValor = CInt(Mid(strBase, intPos, 1))
                            intValor = intValor * (8 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 10
                        strDigito2 = Right(IIf(intResto = 0, "0", CStr(10 - intResto)), 1)
                        strBase2 = Left(strBase, 6) & strDigito2
                        intSoma = 0
                        For intPos = 1 To 8
                            intValor = CInt(Mid(strBase2, intPos, 1))
                            intValor = intValor * (9 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 10
                        strDigito1 = Right(IIf(intResto = 0, "0", CStr(10 - intResto)), 1)
                    Else
                        intSoma = 0
                        For intPos = 1 To 6
                            intValor = CInt(Mid(strBase, intPos, 1))
                            intValor = intValor * (8 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 11
                        strDigito2 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                        strBase2 = Left(strBase, 6) & strDigito2
                        intSoma = 0
                        For intPos = 1 To 7
                            intValor = CInt(Mid(strBase2, intPos, 1))
                            intValor = intValor * (9 - intPos)
                            intSoma = intSoma + intValor
                        Next
                        intResto = intSoma Mod 11
                        strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                    End If
                    strBase2 = Left(strBase, 6) & strDigito1 & strDigito2
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If


            Case "CE" ' Ceara
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                intSoma = 0
                For intPos = 1 To 8
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                intValor = 11 - intResto
                If intValor > 9 Then
                    intValor = 0
                End If
                strDigito1 = Right(CStr(intValor), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "DF" ' Distrito Federal         
                strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                'If Left(strBase, 3) = "073" Then
                intSoma = 0
                intPeso = 2
                For intPos = 11 To 1 Step -1
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 9 Then
                        intPeso = 2
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = Left(strBase, 11) & strDigito1
                intSoma = 0
                intPeso = 2
                For intPos = 12 To 1 Step -1
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 9 Then
                        intPeso = 2
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito2 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = Left(strBase, 12) & strDigito2
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If
                'End If

            Case "ES" ' Espirito Santo
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                intSoma = 0
                For intPos = 1 To 8
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "GO" ' Goias
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                If InStr(1, "10,11,15", Left(strBase, 2), vbTextCompare) > 0 Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    If intResto = 0 Then
                        strDigito1 = "0"
                    ElseIf intResto = 1 Then
                        intNumero = CInt(Left(strBase, 8))
                        strDigito1 = Right(IIf(intNumero >= 10103105 And intNumero <= 10119997, "1", "0"), 1)
                    Else
                        strDigito1 = Right(CStr(11 - intResto), 1)
                    End If
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "MA" ' Maranhão
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                If Left(strBase, 2) = "12" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "MT" ' Mato Grosso
                While Len(strOrigem) < 11
                    strOrigem = "0" & strOrigem
                End While
                strBase = Left(Trim(strOrigem) & "000000000", 11)
                d01 = CInt(Mid(strBase, 1, 1))
                d02 = CInt(Mid(strBase, 2, 1))
                d03 = CInt(Mid(strBase, 3, 1))
                d04 = CInt(Mid(strBase, 4, 1))
                d05 = CInt(Mid(strBase, 5, 1))
                d06 = CInt(Mid(strBase, 6, 1))
                d07 = CInt(Mid(strBase, 7, 1))
                d08 = CInt(Mid(strBase, 8, 1))
                d09 = CInt(Mid(strBase, 9, 1))
                d10 = CInt(Mid(strBase, 10, 1))
                dfinal = CInt(Mid(strBase, 11, 1))
                ds = 3 * d01 + 2 * d02 + 9 * d03 + 8 * d04 + 7 * d05 + 6 * d06 + 5 * d07 + 4 * d08 +
                                   3 * d09 + 2 * d10
                aux1 = Fix(ds / 11)
                aux1 = aux1 * 11
                aux2 = ds - aux1
                If aux2 = 0 Or aux2 = 1 Then
                    digverificador = 0
                Else
                    digverificador = 11 - aux2
                End If
                If dfinal = digverificador Then
                    FValidaInscEst = True
                End If

            Case "MS" ' Mato Grosso do Sul
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                If Left(strBase, 2) = "28" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "MG" ' Minas Gerais, aki onde tem FormatNumber ante era format
                strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                strBase2 = Left(strBase, 3) & "0" & Mid(strBase, 4, 8)
                intNumero = 2
                For intPos = 1 To 12
                    intValor = CInt(Mid(strBase2, intPos, 1))
                    intNumero = IIf(intNumero = 2, 1, 2)
                    intValor = intValor * intNumero
                    If intValor > 9 Then
                        strDigito1 = FormatNumber(intValor, "00")
                        intValor = CInt(Left(strDigito1, 1)) + CInt(Right(strDigito1, 1))
                    End If
                    intSoma = intSoma + intValor
                Next
                intValor = intSoma
                While Right(FormatNumber(intValor, "000"), 1) <> "0"
                    intValor = intValor + 1
                End While
                strDigito1 = Right(FormatNumber(intValor - intSoma, "00"), 1)
                strBase2 = Left(strBase, 11) & strDigito1
                intSoma = 0
                intPeso = 2
                For intPos = 12 To 1 Step -1
                    intValor = CInt(Mid(strBase2, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 11 Then
                        intPeso = 2
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito2 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = strBase2 & strDigito2
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "PA" ' Para
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                If Left(strBase, 2) = "15" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "PB" ' Paraiba
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                intSoma = 0
                For intPos = 1 To 8
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                intValor = 11 - intResto
                If intValor > 9 Then
                    intValor = 0
                End If
                strDigito1 = Right(CStr(intValor), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "PE" ' Pernambuco
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                Dim isValidaZero As Boolean = IIf(Right(strOrigem, 1) = "0", True, False)
                intSoma = 0
                intPeso = 2
                For intPos = 8 To 1 Step -1
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 9 Then
                        intPeso = 1
                    End If
                Next
                intResto = intSoma Mod 11
                intValor = 11 - intResto
                If intValor > 9 Then
                    intValor = intValor - 10
                End If
                strDigito1 = Right(CStr(intValor), 1)

                If isValidaZero Then
                    strBase2 = Left(strBase, 8) & "0"
                Else
                    strBase2 = Left(strBase, 8) & strDigito1
                End If


                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "PI" ' Piaui
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                intSoma = 0
                For intPos = 1 To 8
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "PR" ' Parana
                strBase = Left(Trim(strOrigem) & "0000000000", 10)
                intSoma = 0
                intPeso = 2
                For intPos = 8 To 1 Step -1
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 7 Then
                        intPeso = 2
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                intSoma = 0
                intPeso = 2
                For intPos = 9 To 1 Step -1
                    intValor = CInt(Mid(strBase2, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 7 Then
                        intPeso = 2
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito2 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = strBase2 & strDigito2
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "RJ" ' Rio de Janeiro
                strBase = Left(Trim(strOrigem) & "00000000", 8)
                intSoma = 0
                intPeso = 2
                For intPos = 7 To 1 Step -1
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 7 Then
                        intPeso = 2
                    End If
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = Left(strBase, 7) & strDigito1
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "RN" ' Rio Grande do Norte
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                If Left(strBase, 2) = "20" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intSoma = intSoma * 10
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto > 9, "0", CStr(intResto)), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "RO" ' Rondonia, estado alterado
                strBase = Left(Trim(strOrigem) & "000000000", 14)
                d01 = CInt(Mid(strBase, 1, 1))
                d02 = CInt(Mid(strBase, 2, 1))
                d03 = CInt(Mid(strBase, 3, 1))
                d04 = CInt(Mid(strBase, 4, 1))
                d05 = CInt(Mid(strBase, 5, 1))
                d06 = CInt(Mid(strBase, 6, 1))
                d07 = CInt(Mid(strBase, 7, 1))
                d08 = CInt(Mid(strBase, 8, 1))
                d09 = CInt(Mid(strBase, 9, 1))
                d10 = CInt(Mid(strBase, 10, 1))
                d11 = CInt(Mid(strBase, 11, 1))
                d12 = CInt(Mid(strBase, 12, 1))
                d13 = CInt(Mid(strBase, 13, 1))
                dfinal = CInt(Mid(strBase, 14, 1))
                ds = 6 * d01 + 5 * d02 + 4 * d03 + 3 * d04 + 2 * d05 + 9 * d06 + 8 * d07 + 7 * d08 +
                                   6 * d09 + 5 * d10 + 4 * d11 + 3 * d12 + 2 * d13
                aux1 = Fix(ds / 11)
                aux1 = aux1 * 11
                aux2 = ds - aux1
                digverificador = 11 - aux2
                If digverificador > 9 Then
                    resto_do_calculo = digverificador - 10
                Else
                    resto_do_calculo = digverificador
                End If
                If dfinal <> resto_do_calculo Then
                    FValidaInscEst = False
                Else
                    FValidaInscEst = True
                End If

            Case "RR" ' Roraima
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                If Left(strBase, 2) = "24" Then
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * intPos
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 9
                    strDigito1 = Right(CStr(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "RS" ' Rio Grande do Sul
                strBase = Left(Trim(strOrigem) & "0000000000", 10)
                intNumero = CInt(Left(strBase, 3))
                If intNumero > 0 And intNumero < 468 Then
                    intSoma = 0
                    intPeso = 2
                    For intPos = 9 To 1 Step -1
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 9 Then
                            intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    intValor = 11 - intResto
                    If intValor > 9 Then
                        intValor = 0
                    End If
                    strDigito1 = Right(CStr(intValor), 1)
                    strBase2 = Left(strBase, 9) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

            Case "SC" ' Santa Catarina
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                intSoma = 0
                For intPos = 1 To 8
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "SE" ' Sergipe
                strBase = Left(Trim(strOrigem) & "000000000", 9)
                intSoma = 0
                For intPos = 1 To 8
                    intValor = CInt(Mid(strBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
                Next
                intResto = intSoma Mod 11
                intValor = 11 - intResto
                If intValor > 9 Then
                    intValor = 0
                End If
                strDigito1 = Right(CStr(intValor), 1)
                strBase2 = Left(strBase, 8) & strDigito1
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "SP" ' São Paulo
                If Left(strOrigem, 1) = "P" Then
                    strBase = Left(Trim(strOrigem) & "0000000000000", 13)
                    strBase2 = Mid(strBase, 2, 8)
                    intSoma = 0
                    intPeso = 1
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso = 2 Then
                            intPeso = 3
                        End If
                        If intPeso = 9 Then
                            intPeso = 10
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(CStr(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1 & Mid(strBase, 11, 3)
                Else
                    strBase = Left(Trim(strOrigem) & "000000000000", 12)
                    intSoma = 0
                    intPeso = 1
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso = 2 Then
                            intPeso = 3
                        End If
                        If intPeso = 9 Then
                            intPeso = 10
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(CStr(intResto), 1)
                    strBase2 = Left(strBase, 8) & strDigito1 & Mid(strBase, 10, 2)
                    intSoma = 0
                    intPeso = 2
                    For intPos = 11 To 1 Step -1
                        intValor = CInt(Mid(strBase, intPos, 1))
                        intValor = intValor * intPeso
                        intSoma = intSoma + intValor
                        intPeso = intPeso + 1
                        If intPeso > 10 Then
                            intPeso = 2
                        End If
                    Next
                    intResto = intSoma Mod 11
                    strDigito2 = Right(CStr(intResto), 1)
                    strBase2 = strBase2 & strDigito2
                End If
                If strBase2 = strOrigem Then
                    FValidaInscEst = True
                End If

            Case "TO" ' Tocantins
                strBase = Left(Trim(strOrigem) & "00000000000", 11)
                If InStr(1, "01,02,03,99", Mid(strBase, 3, 2), vbTextCompare) > 0 Then
                    strBase2 = Left(strBase, 2) & Mid(strBase, 5, 6)
                    intSoma = 0
                    For intPos = 1 To 8
                        intValor = CInt(Mid(strBase2, intPos, 1))
                        intValor = intValor * (10 - intPos)
                        intSoma = intSoma + intValor
                    Next
                    intResto = intSoma Mod 11
                    strDigito1 = Right(IIf(intResto < 2, "0", CStr(11 - intResto)), 1)
                    strBase2 = Left(strBase, 10) & strDigito1
                    If strBase2 = strOrigem Then
                        FValidaInscEst = True
                    End If
                End If

        End Select

    End Function

    Public Shared Function FPosInCombo(ByVal oCtr, ByVal sStr, ByVal iTam)

        Dim nc1 As Integer

        If IsNothing(sStr) = False Then
            For nc1 = 0 To oCtr.Items.Count - 1
                oCtr.SelectedIndex = nc1
                If Trim(Left(oCtr.Text, Len(sStr))) = Trim(Left(sStr, Len(sStr))) Then
                    FPosInCombo = nc1 : Exit Function
                End If
            Next
        End If

        FPosInCombo = -1

        oCtr.SelectedIndex = -1

    End Function

    Public Shared Function fValidaMail(ByVal pEmail As String, Optional ByVal Multiplo As Boolean = False) As Boolean
        'Verifica se o parametro passado e um mail valido
        Dim Conta As Integer, Flag As Integer, cValido As String

        fValidaMail = False
        pEmail = Trim(pEmail)

        If Len(pEmail) < 5 Then Exit Function

        'Verifica a existência de (@)
        If InStr(pEmail, "@") = 0 Then
            Exit Function
        Else
            Flag = 0

            If IsNothing(Multiplo) Or Multiplo = False Then
                For Conta = 1 To Len(pEmail)
                    If Mid(pEmail, Conta, 1) = "@" Then
                        Flag = Flag + 1
                    End If
                Next

                If Flag > 1 Then Exit Function
            End If
        End If

        If Left(pEmail, 1) = "@" Then
            Exit Function
        ElseIf Right(pEmail, 1) = "@" Then
            Exit Function
            'ElseIf InStr(pEmail, ".@") > 0 Then  'retirado 14-04-2004 - henrique - UOL permite email com .@
            '    Exit Function
        ElseIf InStr(pEmail, "@.") > 0 Then
            Exit Function
        ElseIf InStr(pEmail, ".b") > 0 Then
            If InStr(pEmail, ".b") + 1 = Len(pEmail) Then Exit Function
        End If

        'Verifica a existência de (.)
        If InStr(pEmail, ".") = 0 Then
            Exit Function
        ElseIf Left(pEmail, 1) = "." Then
            Exit Function
        ElseIf Right(pEmail, 1) = "." Then
            Exit Function
        ElseIf InStr(pEmail, "..") > 0 Then
            Exit Function
        End If

        'Verifica se existe caracter inválido
        For Conta = 1 To Len(pEmail)
            cValido = Mid(pEmail, Conta, 1)
            If IsNothing(Multiplo) Or Multiplo = False Then
                If Not (LCase(cValido) Like "[a-z]" Or cValido =
                 "@" Or cValido = "." Or cValido = "-" Or
                 cValido = "_" Or cValido Like "[0-9]") Then
                    Exit Function
                End If
            Else
                If Not (LCase(cValido) Like "[a-z]" Or cValido =
                 "@" Or cValido = "." Or cValido = "-" Or
                 cValido = "_" Or cValido = ";") Then
                    Exit Function
                End If
            End If
        Next

        fValidaMail = True
    End Function

    'VALIDA HOME PAGE
    Public Shared Function fValidaURL(ByVal strUrl) As Boolean
        Dim blnUrl As Boolean

        If Strings.InStr(strUrl, "www.") = 0 And Strings.InStr(strUrl, ".com") Then
            blnUrl = True
        Else
            blnUrl = False
        End If

        Return blnUrl
    End Function

    Public Function FDeci(ByVal sString As String)

        Dim sVol As String, sSum As String, iConta As Integer

        iConta = 1

        If sString = String.Empty Then
            sString = "0"
        End If

        Do Until False
            If Asc(Mid(sString, iConta, 1)) > 47 And Asc(Mid(sString, iConta, 1)) < 58 Then
                Exit Do
            End If
            iConta = iConta + 1
            If iConta > Len(sString) Then Exit Do
        Loop

        sSum = Mid(sString, iConta - IIf(iConta > 1, 1, 0), Len(sString) - iConta + 2)
        sVol = String.Empty

        If InStr(sSum, gsSepMilhar) > InStr(sSum, gsSepDecimal) Then

            For iConta = 1 To Len(sSum)

                If Mid(sSum, iConta, 1) = gsSepMilhar Then
                    sVol = sVol & gsSepDecimal
                Else
                    If Mid(sSum, iConta, 1) = gsSepDecimal Then
                        sVol = sVol & gsSepMilhar
                    Else
                        sVol = sVol & Mid(sSum, iConta, 1)
                    End If
                End If

            Next

            sSum = sVol
            sVol = String.Empty

        End If

        For iConta = 1 To Len(sSum)

            If Mid(sSum, iConta, 1) = gsSepDecimal Then
                sVol = sVol & "."
            Else
                If Mid(sSum, iConta, 1) <> gsSepMilhar And Mid(sSum, iConta, 1) <> "_" Then
                    sVol = sVol & Mid(sSum, iConta, 1)
                End If
            End If

        Next

        If sVol = String.Empty Or sVol = "." Then
            sVol = "0.00"
        End If

        If Left(sVol, 1) = gsSepMilhar Or Left(sVol, 1) = gsSepDecimal Then
            sVol = "0" & sVol
        End If

        FDeci = sVol

    End Function

    Public Function Fnd(ByVal sString As String)
        Dim sVol As String, iConta As Integer, iCon As Integer

        sVol = ""

        For iConta = 1 To Len(sString)

            If Mid(sString, iConta, 1) <> "_" And Mid(sString, iConta, 1) <> " " Then
                sVol = sVol & Mid(sString, iConta, 1)
            End If

        Next

        sVol = Trim(sVol)

        iCon = InStr(sVol, ".")
        Do Until iCon = 0
            sVol = Left(sVol, iCon - 1) & Right(sVol, Len(sVol) - iCon)
            iCon = InStr(iCon, sVol, ".")
        Loop

        If sVol = String.Empty Or sVol = "." Then
            sVol = String.Empty
        End If

        Do While True
            If Left(sVol, 1) = gsSepMilhar Then
                sVol = Right(sVol, Len(sVol) - 1)
            End If
            If Left(sVol, 1) <> gsSepMilhar Then
                Exit Do
            End If
            If Len(sVol) <= 1 Then
                Exit Do
            End If
        Loop

        Fnd = sVol

        If InStr(sVol, gsSepMilhar) > InStr(sVol, gsSepDecimal) And InStr(sVol, gsSepDecimal) > 0 Then

            Fnd = String.Empty

            For iConta = 1 To Len(sVol)

                If Mid(sVol, iConta, 1) = gsSepMilhar Then
                    Fnd = Fnd & gsSepDecimal
                Else
                    If Mid(sVol, iConta, 1) = gsSepDecimal Then
                        Fnd = Fnd & gsSepMilhar
                    Else
                        Fnd = Fnd & Mid(sVol, iConta, 1)
                    End If
                End If

            Next

        End If

        Do Until False
            If Left(Fnd, 1) = gsSepMilhar Then
                Fnd = Right(Fnd, Len(Fnd) - 1)
            Else
                Exit Do
            End If
            If Len(Fnd) = 0 Then Exit Do
        Loop


        For iCon = 1 To 2

            If Left(Fnd, 1) = gsSepDecimal Then
                If InStr(2, Fnd, gsSepDecimal) > 1 Then
                    Fnd = Right(Fnd, Len(Fnd) - 1)
                Else
                    Fnd = "0" & Fnd
                End If
            End If

        Next

    End Function

    ''' <summary>
    ''' FUNÇÃO QUE PRINTA A TELA
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Function PrintaTela() As String
        'Dim strPrintErro As String = ""
        'Dim bmpImg As System.Drawing.Bitmap = Nothing
        'Try
        '    bmpImg = New System.Drawing.Bitmap(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width, System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        '    Dim grpGraphic As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmpImg)

        '    grpGraphic.CopyFromScreen(New System.Drawing.Point(0, 0), New System.Drawing.Point(0, 0), New System.Drawing.Size(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width, System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height))

        '    bmpImg.Save(System.Windows.Forms.Application.StartupPath & "\ERRO.jpg")
        '    strPrintErro = (System.Windows.Forms.Application.StartupPath & "\ERRO.jpg")

        '    grpGraphic.Dispose()
        'Catch ex As Exception
        '    Throw New Exception(ex.Message, ex.InnerException)
        'Finally
        '    bmpImg.Dispose()
        'End Try
        'Return strPrintErro
    End Function

    ''' <summary>
    ''' Função que apaga a imagem gerada pelo print de erro
    ''' </summary>
    ''' <param name="strCaminho"></param>
    ''' <remarks></remarks>
    Private Shared Sub ApagaImagemErro(ByVal strCaminho As String)
        Try
            Kill(strCaminho)
        Catch ex As Exception
            Throw New Exception(ex.Message, ex.InnerException)
        End Try
    End Sub

    Public Shared Sub AtualizaApplEventLog(Optional ByVal strID As String = "",
                Optional ByVal strMenssagem As String = "",
                Optional ByVal strTipo As String = "",
                Optional ByVal strLocal As String = "",
                Optional ByVal strSetor As String = "",
                Optional ByVal strUsuario As String = "",
                Optional ByVal strComputador As String = "",
                Optional ByVal strVersao As String = "",
                Optional ByVal strDepto As String = "",
                Optional ByVal strPrintErro As String = "")

        '23/03/2017 - Fernando
        'Se estiver debugando não precisa executar a rotina
        If System.Diagnostics.Debugger.IsAttached Then Exit Sub

        Dim strMensagemHTMl As String = ""
        Dim strdetalhe As String = vbLf & "TIPO: " & IIf(strTipo = "1", "Arquitetura", "Funcional") & vbLf & "LOCAL:  " & strLocal & vbLf & "SETOR: " & strSetor & vbLf & vbLf & "USUARIO: " & strUsuario & vbLf & "COMPUTADOR: " & strComputador & vbLf & "VERSÃO: " & strVersao

        Dim sSource As String
        Dim sLog As String
        Dim sMachine As String


        Try

            Dim strCaminhoPrint As String = ""

            'If Not (strMenssagem.Contains("Timeout") OrElse strMenssagem.Contains("timed out") OrElse strMenssagem.Contains("deadlocked") OrElse strMenssagem.Contains("Domain not found")) OrElse strMenssagem.Contains("tempo limite") Then
            '07/08/2017 - Fernando
            'Valida se é um erro conhecido. Se for erro conhecido não é necessário o print.
            'Se não for erro conhecido apaga o print.
            If Not erroConhecido(strMenssagem) Then
                strCaminhoPrint = PrintaTela()
            End If

            If Trim(strPrintErro) = "" And Trim(strCaminhoPrint) <> "" Then
                strPrintErro = strCaminhoPrint
            ElseIf Trim(strPrintErro) <> "" And Trim(strCaminhoPrint) <> "" Then
                strPrintErro = strCaminhoPrint & "|" & strPrintErro
            End If

            Dim arrPrintErro() As String

            If Not Trim(strPrintErro) = "" Then
                'SPLIT DAS IMAGENS DE ERRO
                arrPrintErro = strPrintErro.Split("|")
            End If

            'ENVIA EMAIL PARA DESENVOLVEDOR
            If System.IO.File.Exists("Erro.html") Then
                strMensagemHTMl = LeArquivoTXT("Erro.html")
            Else
                strMensagemHTMl = "ID: <%ID%> <br> Mensagem: <%MENSAGEM%> <br> Tipo: <%TIPO%> <br> Local: <%LOCAL%> <br> Depto: <%DEPTO%> <br> Setor: <%SETOR%> <br> Usuario: <%USUARIO%> <br> Computador: <%COMPUTADOR%> <br> Versão: <%VERSAO%> <br> "
            End If

            strMensagemHTMl = strMensagemHTMl.Replace("<%ID%>", strID).Replace("<%MENSAGEM%>", strMenssagem).Replace("<%TIPO%>", IIf(strTipo = "1", "Arquitetura", "Funcional")).Replace("<%LOCAL%>", strLocal).Replace("<%SETOR%>", strSetor).Replace("<%USUARIO%>", strUsuario).Replace("<%COMPUTADOR%>", strComputador).Replace("<%VERSAO%>", strVersao).Replace("<%DEPTO%>", strDepto).Replace("<%BASE%>", Conection.ParametroConexao(0))

            'ANTES DE ENVIAR O E-MAIL PROCURA O ESMTP SERVER
            Dim arrParametros As String() = PesquisaParametros()

            'EMAIL DO ERRO  "Hercules@teleatlantic.com.br; sarah@teleatlantic.com.br; edson.ferreira@teleatlantic.com.br;
            'If Not ((strComputador = "D6JXB342" OrElse strComputador = "D40PM1P1" OrElse strComputador = "DCF0GXP1" OrElse strComputador = "D4DPM1P1") And Conection.ParametroConexao(0) = "AQIJ8IFASEF" Or Conection.ParametroConexao(0) = "D82SX861" Or Conection.ParametroConexao(0) = "D82SX861\D82SX861" Or Conection.ParametroConexao(0) = "POG-D82" Or Conection.ParametroConexao(0) = "192.168.1.6" Or Conection.ParametroConexao(0) = "D82SX61") Then
            '    If Not (strMenssagem.Contains("Failure sending mail") OrElse strMenssagem.Contains("Domain not found") OrElse strMenssagem.Contains("Timeout") OrElse strMenssagem.Contains("timed out") OrElse strMenssagem.Contains("tempo limite") OrElse strMenssagem.Contains("deadlocked") OrElse strMenssagem.Contains("Telesystem Desatualizado") OrElse strMenssagem.Contains("Inaccessible logs: Security") _
            '            OrElse strMenssagem.Contains("A transport-level error has occurred") OrElse strMenssagem.Contains("OutOfMemory") OrElse strMenssagem.Contains("Problemas cadastro do lead")) And Not (strComputador = "D6JXB342" OrElse strComputador = "D40PM1P1" OrElse strComputador = "DCF0GXP1" OrElse strComputador = "DDF0GXP1") And Trim(strUsuario) <> "" Then
            '        'EnviarMensagemEmail("", arrParametros(1), "no-reply@verisure.com.br", "", "", gEmail, Strings.Left(strMenssagem, 7) & " - Erro Verisystem", "Data Erro: " & PegaData() & "<Br/>" & strMensagemHTMl, True, arrPrintErro, "Verisystem", "", False)

            '        '27/07/2016 - Fernando
            '        'Todos e-mails de erro serão enviados pelo sharedrelay manualmente
            '        EnviarMensagemEmail("", arrParametros(2), "no-reply@verisure.com.br", "", "", gEmail, Strings.Left(strMenssagem, 7) & " - Erro Verisystem", "Data Erro: " & PegaData() & "<Br/>" & strMensagemHTMl, True, arrPrintErro, "Verisystem", "", False)
            '    Else
            '        'NASD - Não aberto Service Desk
            '        'EnviarMensagemEmail("", arrParametros(1), "no-reply@verisure.com.br", "", "", gEmail, "NASD - " & Strings.Left(strMenssagem, 7) & " - Erro Verisystem", "Data Erro: " & PegaData() & "<Br/>" & strMensagemHTMl, True, arrPrintErro, "Verisystem", "", False)

            '        '27/07/2016 - Fernando
            '        'Todos e-mails de erro serão enviados pelo sharedrelay manualmente
            '        EnviarMensagemEmail("", arrParametros(2), "no-reply@verisure.com.br", "", "", gEmail, "NASD - " & Strings.Left(strMenssagem, 7) & " - Erro Verisystem", "Data Erro: " & PegaData() & "<Br/>" & strMensagemHTMl, True, arrPrintErro, "Verisystem", "", False)
            '    End If
            'End If

            '07/08/2017 - Fernando
            'Valida se é um erro conhecido. Se for erro conhecido não é necessário envio de e-mail, apenas loga localmente.
            'Se não for erro conhecido efetua o disparo.
            If Not erroConhecido(strMenssagem) Then
                EnviarMensagemEmail("", arrParametros(2), "no-reply@verisure.com.br", "", "", gEmail, Strings.Left(strMenssagem, 7) & " - Erro Verisystem", "Data Erro: " & PegaData() & "<Br/>" & strMensagemHTMl, True, arrPrintErro, "Verisystem", "", False)
                CHelpDesk.LogarErroConhecido(strUsuario, strMenssagem, strLocal, 0, strTipo, strDepto, strSetor, strComputador, strVersao, Conection.ParametroConexao(0))
            Else
                'Se for erro conhecido insere em uma tabela de controle, para não ficar 100% no limbo
                CHelpDesk.LogarErroConhecido(strUsuario, strMenssagem, strLocal, 1, strTipo, strDepto, strSetor, strComputador, strVersao, Conection.ParametroConexao(0))
            End If


            'CRIA O LOG
            sSource = "Telesystem2"
            sLog = "Application"
            sMachine = "."

            If Not EventLog.SourceExists(sSource) Then
                EventLog.CreateEventSource(sSource, sLog)

            End If

            Dim ELog As New EventLog(sLog, sMachine, sSource)
            ELog.WriteEntry("ID: " & strID & vbLf & "MENSAGEM: " & strMenssagem & strdetalhe, EventLogEntryType.Error)


            '07/08/2017 - Fernando
            'Não há mais necessidade do bloco que abriria o chamado
            'ABRE OCORRÊNCIA PARA O HELPDESK
            'If Not (strMenssagem.Contains("Failure sending mail") OrElse strMenssagem.Contains("Domain not found") OrElse strMenssagem.Contains("Timeout") OrElse strMenssagem.Contains("timed out") OrElse strMenssagem.Contains("tempo limite") OrElse strMenssagem.Contains("deadlocked") OrElse strMenssagem.Contains("Telesystem Desatualizado") OrElse strMenssagem.Contains("Inaccessible logs: Security") _
            '  OrElse strMenssagem.Contains("A transport-level error has occurred") OrElse strMenssagem.Contains("OutOfMemory") OrElse strMenssagem.Contains("Problemas cadastro do lead")) And Not (strComputador = "D6JXB342" OrElse strComputador = "D40PM1P1" OrElse strComputador = "D4JXB342" OrElse strComputador = "DDF0GXP1" OrElse strComputador = "D4DPM1P1" OrElse strComputador = "DBJXB342") And Trim(strUsuario) <> "" Then
            '    CHelpDesk.AbreOcorrência(strUsuario, strMenssagem & vbCrLf & strLocal, strCaminhoPrint)
            'End If

            'If Not (strMenssagem.Contains("Domain not found") OrElse strMenssagem.Contains("Timeout") OrElse strMenssagem.Contains("timed out") OrElse strMenssagem.Contains("tempo limite") OrElse strMenssagem.Contains("deadlocked")) Then
            '07/08/2017 - Fernando
            'Apaga o print se houver dados de print
            If strCaminhoPrint <> "" Then
                ApagaImagemErro(strCaminhoPrint)
            End If

        Catch ex As Exception

            Try

                strMensagemHTMl = "ID: <%ID%> <br> Mensagem: <%MENSAGEM%> <br> Tipo: <%TIPO%> <br> Local: <%LOCAL%> <br> Depto: <%DEPTO%> <br> Setor: <%SETOR%> <br> Usuario: <%USUARIO%> <br> Computador: <%COMPUTADOR%> <br> Versão: <%VERSAO%> <br> "

                strMensagemHTMl = strMensagemHTMl.Replace("<%ID%>", "LOG").Replace("<%MENSAGEM%>", strMenssagem + " " + ex.Message) _
                 .Replace("<%TIPO%>", IIf(strTipo = "1", "Arquitetura", "Funcional")) _
                 .Replace("<%LOCAL%>", "Projeto: Common - Classe: Funcoes - Metodos: AtualizaApplEventLog") _
                 .Replace("<%SETOR%>", strSetor).Replace("<%USUARIO%>", strUsuario) _
                 .Replace("<%COMPUTADOR%>", Environment.MachineName).Replace("<%VERSAO%>", strVersao)

                'NASD - Não aberto Service Desk
                ''strMensagemHTMl = LeArquivoTXT("Erro.html")
                ''strMensagemHTMl = strMensagemHTMl.Replace("<%ID%>", "LOG").Replace("<%MENSAGEM%>", ex.Message).Replace("<%TIPO%>", IIf(strTipo = "1", "Arquitetura", "Funcional")).Replace("<%LOCAL%>", "Projeto: Common - Classe: Funcoes - Metodos: AtualizaApplEventLog").Replace("<%SETOR%>", strSetor).Replace("<%USUARIO%>", strUsuario).Replace("<%COMPUTADOR%>", strComputador).Replace("<%VERSAO%>", strVersao)
                ''EMAIL DO ERRO  "Hercules@teleatlantic.com.br; sarah@teleatlantic.com.br; edson.ferreira@teleatlantic.com.br; wolneyaf@gmail.com"

                'If (Not ex.Message.Contains("Inaccessible logs: Security")) AndAlso (Not ex.Message.Contains("Logs inacessíveis: Security")) Then
                '07/08/2017 - Fernando
                'Valida se é um erro conhecido. Se for erro conhecido não é necessário envio de e-mail, apenas loga localmente.
                'Se não for erro conhecido efetua o disparo.
                If Not erroConhecido(ex.Message) Then
                    'EnviarMensagemEmail("", "smtplw.com.br", "no-reply@verisure.com.br", "", "", gEmail, Strings.Left(strMenssagem, 7) & " - Erro Verisystem", "Data Erro: " & PegaData() & "<Br/>" & strMensagemHTMl, True, , "Verisystem", "", False)
                    EnviarMensagemEmail("", "email-smtp.sa-east-1.amazonaws.com", "no-reply@verisure.com.br", "", "", gEmail, Strings.Left(strMenssagem, 7) & " - Erro Verisystem", "Data Erro: " & PegaData() & "<Br/>" & strMensagemHTMl, True, , "Verisystem", "", False)
                    'EnviarMensagemEmail("", "192.168.1.13", "no-reply@teleatlantic.com.br", "", "", gEmail, "NASD - " & Strings.Left(strMenssagem, 7) & " - Erro Telesystem 2", "Data Erro: " & PegaData() & "<Br/>" & strMensagemHTMl, True, , "Telesystem 2", "", False)
                End If

            Catch innerEx As Exception
                CriaLog(Application.StartupPath & "\RelatorioErros.txt", ex.Message & vbCrLf & innerEx.Message)
            End Try

        End Try
    End Sub

    Private Shared Function erroConhecido(msgErro As String) As Boolean
        msgErro = msgErro.ToUpper()

        If msgErro.Contains("FAILURE SENDING MAIL") Or
           msgErro.Contains("TIMEOUT") Or
           msgErro.Contains("TIMED OUT") Or
           msgErro.Contains("TEMPO LIMITE") Or
           msgErro.Contains("DEADLOCKED") Or
           msgErro.Contains("INACCESSIBLE LOGS: SECURITY") Or
           msgErro.Contains("LOGS INACESSÍVEIS: SECURITY") Or
           msgErro.Contains("A TRANSPORT-LEVEL ERROR HAS OCCURRED") Or
           msgErro.Contains("OUTOFMEMORY") Or
           msgErro.Contains("PROBLEMAS CADASTRO DO LEAD") Or
           msgErro.Contains("A SEVERE ERROR OCCURRED ON THE CURRENT COMMAND") Or
           msgErro.Contains("ERRO GRAVE NO COMANDO ATUAL") Or
           msgErro.Contains("NENHUM ARQUIVO SELECIONADO") Or
           msgErro.Contains("UM ERRO SEVERO") Or
           msgErro.Contains("VALOR DO BOLETO INVÁLIDO") Or
           msgErro.Contains("NENHUM TÍTULO ENCONTRADO") Or
           msgErro.Contains("E-MAIL PARA CONTROLE.") Or
           msgErro.Contains("UTILIZE A BAIXA ESPECÍFICA DE BOLETOS UNIFICADOS") Or
           msgErro.Contains("SYSTEM.OUTOFMEMORYEXCEPTION") Then Return True

        'msgErro.Contains("TELESYSTEM DESATUALIZADO") Or
        'msgErro.Contains("DOMAIN NOT FOUND") Or

        Return False
    End Function

    ''' <summary>
    ''' Pesquisa os parametros
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function PesquisaParametros() As String()
        Dim rdr As SqlDataReader
        Dim arrParametros As String() = {"", "", ""}
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("P_PesquisaParametros", connection)
        command.CommandType = CommandType.StoredProcedure
        command.CommandTimeout = DadosGenericos.Timeout.Query

        connection.Open()

        rdr = command.ExecuteReader
        If rdr.HasRows Then
            While rdr.Read
                arrParametros = {"", "", ""}
                arrParametros(0) = rdr("HelpDeskPath").ToString
                arrParametros(1) = rdr("ESMTPServer").ToString
                arrParametros(2) = rdr("ESMTPServerServiceDesk").ToString
            End While
        End If

        connection.Close()
        rdr.Close()

        Return arrParametros
    End Function

    Public Shared Function ConsultaFeriado(ByVal intPrazo As Integer, ByVal dtData As DateTime) As DateTime

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim intQtdeDiasUteisOS As Integer = 0

        connection.Open()

        ''Informa a procedure
        Dim Command As SqlCommand
        Do While intQtdeDiasUteisOS < intPrazo

            dtData = DateAdd("d", 1, FormatDateTime(dtData, DateFormat.ShortDate))

            Command = New SqlCommand("P_ObterFeriado", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@DTFeriado", dtData))

            Using rdr As SqlDataReader = Command.ExecuteReader()

                If (Not rdr.HasRows) And DatePart("w", dtData) <> 1 And DatePart("w", dtData) <> 7 Then
                    rdr.Read()
                    intQtdeDiasUteisOS = intQtdeDiasUteisOS + 1
                End If
            End Using
        Loop

        Return dtData

    End Function

    Public Shared Function ConsultaDiaUteis(ByVal dtData As DateTime) As DateTime

        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim intQtdeDiasUteisOS As Integer = 0
        Dim otabels As New DataTable
        Dim strData As String = ""

        Try
            connection.Open()

            ''Informa a procedure
            Dim Command As SqlCommand



            Command = New SqlCommand("P_ObterFeriado", connection)
            Command.CommandType = CommandType.StoredProcedure
            Command.CommandTimeout = DadosGenericos.Timeout.Query
            Command.Parameters.Add(New SqlParameter("@DTFeriado", dtData))

            Using rdr As SqlDataReader = Command.ExecuteReader()

                If (Not rdr.HasRows) And (DatePart("w", dtData) <> 1 Or DatePart("w", dtData) <> 7) Then
                    connection.Close()
                    connection.Dispose()
                    rdr.Close()
                    Return dtData
                Else
                    rdr.Read()

                    If rdr.HasRows Then
                        strData = rdr(0)
                    End If

                    dtData = DateAdd("d", 1, FormatDateTime(CDate(IIf(strData = "", dtData, strData)), DateFormat.ShortDate))

                    rdr.Close()
                    dtData = ConsultaDiaUteis(dtData)
                End If
            End Using
        Catch ex As Exception
            Return Nothing
        Finally
            connection.Close()
            connection.Dispose()
        End Try
        Return dtData
    End Function

    'Public Sub ConectaRPT(ByVal report As Object,
    '                      ByVal crConnectionInfo As Object,
    '                      ByVal CrTables As Object,
    '                      ByVal crtableLogoninfo As Object)

    '    With crConnectionInfo
    '        .ServerName = Conection.ParametroConexao(0)
    '        .DatabaseName = Conection.ParametroConexao(1)
    '        .UserID = Conection.ParametroConexao(2)
    '        .Password = Conection.ParametroConexao(3)
    '    End With

    '    CrTables = report.Database.Tables
    '    For Each CrTable In CrTables
    '        crtableLogoninfo = CrTable.LogOnInfo
    '        crtableLogoninfo.ConnectionInfo = crConnectionInfo
    '        CrTable.ApplyLogOnInfo(crtableLogoninfo)
    '    Next
    'End Sub

    ''' <summary>
    ''' Carrega relatorio RPT.
    ''' </summary>
    ''' <param name="CrystalReportViewer">Passar o Objeto CrystalReportViewer</param>
    ''' <param name="lobjValorParametros">Um vetor contendo os valores dos parametros</param>
    ''' <param name="lstrNomeParametros">Um vetor contendo os parametros do RPT</param>
    ''' <param name="strRelatorio">Nome do relatorio com extensao para a abertura.</param>
    ''' <returns>
    ''' Entidade retorno contendo o erro ou a mensagem de sucesso na abertura
    ''' </returns>
    ''' <remarks></remarks>
    Public Shared Function CarregaRelatorio(ByRef CrystalReportViewer As Object, ByVal lobjValorParametros() As Object, ByVal lstrNomeParametros() As String, ByVal strRelatorio As String, ByVal lrptRelatorio As ReportDocument) As Retorno

        'Dim lrptRelatorio As New ReportDocument
        'Dim lconInfo As New CrystalDecisions.Shared.ConnectionInfo
        'Dim ltblInfo As New TableLogOnInfo
        'Dim lparParametro As ParameterFieldDefinition
        'Dim lfilCampo As ParameterFieldDefinitions
        'Dim ldisDiscrete As ParameterDiscreteValue
        'Dim lvalValor As ParameterValues
        'Dim lstrTexto As String
        Dim _retorno As New Retorno
        '        Dim strIsVerisure As String
        '        '_retorno.Sucesso = True

        '        'If IsNothing(lrptRelatorio) Then
        '        '    lrptRelatorio = New ReportDocument

        '        'End If

        '        '# Verifica o acesso ao arquivo .rpt
        '        'If (System.IO.File.Exists(My.Application.Info.DirectoryPath & "\" & strRelatorio) = False) Then
        '        '    _retorno.Sucesso = False
        '        '    _retorno.MsgErro = "+ CARREGA RELATORIO + Não foi possível acessar o relatório." & vbCrLf &
        '        '                    "Se o erro persistir, entre em contato com o Administrador."
        '        '    _retorno.NumErro = "Fcao"
        '        '    _retorno.TipoErro = DadosGenericos.TipoErro.Funcional
        '        '    _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
        '        '    Return _retorno
        '        '    Exit Function
        '        'End If


        '        Try

        '            ''# Passa os dados para conexão-------------------------------
        '            ''Propriedades exigidas para Conexao de banco diferente de windows autent.
        '            'lconInfo.DatabaseName = Conection.ParametroConexao(1)
        '            'lconInfo.ServerName = Conection.ParametroConexao(0)
        '            'lconInfo.UserID = Conection.ParametroConexao(2)
        '            'lconInfo.Password = Conection.ParametroConexao(3)

        '            'Para windows autent. IntegratedSecurity deve ser igual a  true
        '            'Busca Flag para identificar se o computador é Verisure ou Teleatlantic
        '            strIsVerisure = Funcoes.LerRegistroDoWindows("Configurações", "isVerisure", "0")
        '            'Versão Teste
        '            'strIsVerisure = "0"

        '            'If strIsVerisure.Equals("0") Then
        '            '    lconInfo.IntegratedSecurity = False
        '            '    lconInfo.ServerName = "SYSBASETEST"
        '            '    lconInfo.DatabaseName = "TelesystemHomologacao"
        '            '    lconInfo.UserID = "douglas"
        '            '    lconInfo.Password = "douglas"
        '            'End If

        '            '22/03/2019 - Lucas
        '            'Não é preciso mais verificar o registro, todas as máquinas fazem parte do domínio
        '            'Substituído também o server pelo listener do HA
        '            'If strIsVerisure.Equals("1") Then
        '            '    lconInfo.IntegratedSecurity = False
        '            '    lconInfo.ServerName = "SYSBASEPROD"
        '            '    lconInfo.DatabaseName = "Telesystem"
        '            '    lconInfo.UserID = "verisure"
        '            '    lconInfo.Password = "#VerTel@2001#"
        '            'Else
        '            '    lconInfo.IntegratedSecurity = True
        '            'End If
        '            lconInfo.ServerName = "LTPRODBR01"
        '            lconInfo.IntegratedSecurity = True
        '            '-----------------------------------------------------------

        '            '# Carrega o arquivo rpt
        '            lrptRelatorio.Load(My.Application.Info.DirectoryPath & "\" & strRelatorio)
        '            'lrptRelatorio.VerifyDatabase()
        '            '# Realiza a conexão das tabelas do relatório
        '            For Each ltblTable As Table In lrptRelatorio.Database.Tables
        '                ltblInfo.ConnectionInfo = lconInfo

        '                ltblTable.ApplyLogOnInfo(ltblInfo)
        '            Next

        '            '# Conexão das tabelas dos subrelatórios (CASO EXISTA)
        '            For lintContador = 1 To lrptRelatorio.Subreports.Count
        '                For Each ltblTable As Table In lrptRelatorio.Subreports(lintContador - 1).Database.Tables
        '                    ltblInfo.ConnectionInfo = lconInfo
        '                    ltblTable.ApplyLogOnInfo(ltblInfo)
        '                Next
        '            Next

        '            '# Recebe os parâmetros existentes no rpt
        '            lfilCampo = lrptRelatorio.DataDefinition.ParameterFields


        '            If lobjValorParametros(0) = "" And lstrNomeParametros(0) = "" Then GoTo CarregaRelatorio

        '            '# Loop que irá configurar cada parâmetro
        '            For lintContador As Integer = 0 To UBound(lobjValorParametros)

        '                lparParametro = lfilCampo.Item(lstrNomeParametros(lintContador))
        '                lvalValor = lparParametro.CurrentValues
        '                ldisDiscrete = New ParameterDiscreteValue

        '                '# Verifica o tipo de dado para passar o valor correto
        '                If IsDate(lobjValorParametros(lintContador)) Then
        '                    ldisDiscrete.Value = IIf(lobjValorParametros(lintContador) = "00:00:00" Or lobjValorParametros(lintContador) = "0001-01-01 00:00:00", "1899-12-30", lobjValorParametros(lintContador))

        '                ElseIf IsNumeric(lobjValorParametros(lintContador)) Then
        '                    ldisDiscrete.Value = lobjValorParametros(lintContador)

        '                Else
        '                    ldisDiscrete.Value = IIf(Trim(lobjValorParametros(lintContador)) = "NULL", DBNull.Value, Trim(lobjValorParametros(lintContador)))

        '                End If

        '                '# Adiciona e confirma os valores
        '                lvalValor.Add(ldisDiscrete)
        '                lparParametro.ApplyCurrentValues(lvalValor)
        '            Next

        'CarregaRelatorio:
        '            '# crvVisualizador é meu objeto CrystalReportViçewer
        '            CrystalReportViewer.ReportSource = lrptRelatorio
        '            CrystalReportViewer.Refresh()

        '        Catch ex As SqlException
        '            _retorno.Sucesso = False
        '            _retorno.MsgErro = "+ CARREGA RELATORIO + " & ex.Message
        '            _retorno.NumErro = "Fcao"
        '            _retorno.TipoErro = DadosGenericos.TipoErro.Funcional
        '            _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta

        '        Catch ex As Exception
        '            _retorno.Sucesso = False
        '            _retorno.MsgErro = "+ CARREGA RELATORIO + " & ex.Message
        '            _retorno.NumErro = "Fcao"
        '            _retorno.TipoErro = DadosGenericos.TipoErro.Funcional
        '            _retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
        '        End Try

        Return _retorno
    End Function

    ''' <summary>
    ''' Converta de Hex. para dec.
    ''' </summary>
    ''' <param name="strValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function hextodec(ByVal strValue As String) As Long
        If Left(strValue, 2) <> "&H" Then strValue = "&h" & strValue

        If InStr(1, strValue, ".") Then strValue = Left(strValue, (InStr(1, strValue, ".") - 1))

        hextodec = CLng(strValue)
        Exit Function

    End Function

    ''' <summary>
    ''' Incluir caracteres depois da palavra
    ''' </summary>
    ''' <param name="strPalavra"></param>
    ''' <param name="strCaracter"></param>
    ''' <param name="intTamanho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IncluiCaractere(ByVal strPalavra As String, ByVal strCaracter As String, ByVal intTamanho As Integer) As String
        Dim strPalavraNova As String = ""
        Dim intTotal As Integer = intTamanho - Len(strPalavra)

        strPalavraNova = strPalavra
        For i As Integer = 1 To intTotal
            strPalavraNova = strPalavraNova & strCaracter
        Next

        Return strPalavraNova
    End Function

    ''' <summary>
    ''' Incluir caracteres antes da palavra
    ''' </summary>
    ''' <param name="strPalavra"></param>
    ''' <param name="strCaracter"></param>
    ''' <param name="intTamanho"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IncluiCaractereAntes(ByVal strPalavra As String, ByVal strCaracter As String, ByVal intTamanho As Integer) As String
        Dim strPalavraNova As String = ""
        Dim intTotal As Integer = intTamanho - Len(strPalavra)

        strPalavraNova = strPalavra
        For i As Integer = 1 To intTotal
            strPalavraNova = strCaracter & strPalavraNova
        Next

        Return strPalavraNova
    End Function

    ''' <summary>
    ''' Retira ponto, virgula, hífem, parentes, barra, contrabarra e Underline
    ''' </summary>
    ''' <param name="sString">O string</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()>
    Public Shared Function FSepara(ByVal sString As String) As String
        Dim sVol As String, iConta As Integer

        sVol = ""

        For iConta = 1 To Len(sString)

            If Mid(sString, iConta, 1) <> "." And
               Mid(sString, iConta, 1) <> "," And
               Mid(sString, iConta, 1) <> "-" And
               Mid(sString, iConta, 1) <> "_" And
               Mid(sString, iConta, 1) <> "(" And
               Mid(sString, iConta, 1) <> ")" And
               Mid(sString, iConta, 1) <> "/" Then

                sVol = sVol & Mid(sString, iConta, 1)

            End If

        Next

        FSepara = sVol

    End Function

    Public Shared Function Padl(ByVal sTx As String, ByVal iTam As Integer) As String
        Dim iCon As Integer

        Padl = Left(Trim(sTx), iTam)

        For iCon = Len(Padl) + 1 To iTam
            Padl = Padl & " "
        Next

    End Function

    ''' <summary>
    ''' Remove os acentos.
    ''' </summary>
    ''' <param name="sTexto"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function RemoveAcento(ByVal sTexto As String) As String

        Dim X As Integer, sCarac As String, sNovoTexto As String

        For X = 1 To Len(sTexto)
            Select Case LCase(Mid(sTexto, X, 1))
                Case "á", "à", "â", "ã", "ä"
                    sCarac = "a"
                Case "é", "è", "ê", "ë"
                    sCarac = "e"
                Case "í", "ì", "î", "ï"
                    sCarac = "i"
                Case "ó", "ò", "ô", "õ", "ö"
                    sCarac = "o"
                Case "ú", "ù", "û", "ü"
                    sCarac = "u"
                Case "ç"
                    sCarac = "c"
                Case "ñ"
                    sCarac = "n"
                Case Else
                    If Asc(LCase(Mid(sTexto, X, 1))) = 186 Then 'Para tirar o ordinal
                        sCarac = "."
                    Else
                        sCarac = LCase(Mid(sTexto, X, 1))
                    End If
            End Select

            If UCase(Mid(sTexto, X, 1)) = Mid(sTexto, X, 1) Then sCarac = UCase(sCarac)

            sNovoTexto = sNovoTexto & sCarac
        Next X

        RemoveAcento = sNovoTexto
    End Function

    Public Shared Function fExtraiPalavra(ByVal sFrase As String, ByVal bPrimeira As Boolean) As String
        'retorna a primeira ou última palavra da frase
        Dim i As Integer
        fExtraiPalavra = ""

        If bPrimeira Then
            For i = 1 To Len(sFrase)
                If Mid(sFrase, i, 1) <> " " Then
                    fExtraiPalavra = fExtraiPalavra & Mid(sFrase, i, 1)
                Else
                    Exit For
                End If
            Next i
        Else
            For i = Len(sFrase) To 1 Step -1
                If Mid(sFrase, i, 1) <> " " Then
                    fExtraiPalavra = Mid(sFrase, i, 1) & fExtraiPalavra
                Else
                    Exit For
                End If
            Next i
        End If

    End Function

    ''' <summary>
    ''' Função utilizada para converter um double em uma string com o número por extenso
    ''' </summary>
    ''' <param name="dblValor">Valor à ser convertido</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fExtenso(ByVal dblValor As Double) As String
        Dim WValor, WExtenso, WMoeda, WMoedap As String
        Dim Unid1, Deze1, Deze2, Cente As String
        Dim Dez1, Dez0, Dez2, Dez3, Dez4 As String
        Dim Cent0 = 0, Cent1, Cent2, Cent3, Cent4 As String
        Dim Trilhao, Hist1, Hist2, Hist3, Hist4, Hist5, Hist6 As String
        Dim Flag As Integer

        If String.IsNullOrEmpty(dblValor) Then
            fExtenso = ""
            Exit Function
        End If
        WExtenso = ""
        WMoeda = "REAL"
        WMoedap = "REAIS"
        Unid1 = "UM    DOIS  TRÊS  QUATROCINCO SEIS  SETE  OITO  NOVE  "
        Deze1 = "ONZE     DOZE     TREZE    QUATORZE QUINZE   DEZESSEISDEZESSETEDEZOITO  DEZENOVE "
        Deze2 = "DEZ      VINTE    TRINTA   QUARENTA CINQUENTASESSENTA SETENTA  OITENTA  NOVENTA  "
        Cente = "CENTO       DUZENTOS    TREZENTOS   QUATROCENTOSQUINHENTOS  SEISCENTOS  SETECENTOS  OITOCENTOS  NOVECENTOS  "

        WValor = Format(dblValor, "0000000000000.00")
        Dez0 = Mid(WValor, 15, 2)
        Dez1 = Mid(WValor, 12, 2)
        Cent1 = Mid(WValor, 11, 1)
        Dez2 = Mid(WValor, 9, 2)
        Cent2 = Mid(WValor, 8, 1)
        Dez3 = Mid(WValor, 6, 2)
        Cent3 = Mid(WValor, 5, 1)
        Dez4 = Mid(WValor, 3, 2)
        Cent4 = Mid(WValor, 2, 1)
        Trilhao = Mid(WValor, 1, 1)

        If (Trilhao <> "0") Then
            Hist1 = Trim(Mid(Unid1, Val(Trilhao) * 6 - 5, 6))
            If (Val(Trilhao) > 1) Then
                Hist1 = Hist1 & " TRILHÕES"
            Else
                Hist1 = Hist1 + " TRILHÃO"
            End If
        Else
            Hist1 = ""
        End If

        Flag = False

        Select Case Val(Dez4)
            Case 1 To 9
                Hist2 = Trim(Mid(Unid1, Val(Dez4) * 6 - 5, 6))
            Case 11 To 19
                Hist2 = Trim(Mid(Deze1, Val(Mid(Dez4, 2, 1)) * 9 - 8, 9))
            Case 10
                Hist2 = " DEZ"
            Case 20 To 99
                Hist2 = Trim(Mid(Dez2, Val(Mid(Dez4, 1, 1)) * 9 - 8, 9))
                If (Val(Mid(Dez4, 2, 1))) <> 0 Then
                    Hist2 = Hist2 & " E " & Trim(Mid(Unid1, Val(Mid(Dez4, 2, 1)) * 6 - 5, 6))
                End If
            Case 0
                If (Val(Cent4) = 0) Then
                    Hist2 = ""
                End If

                If (Val(Cent4) = 1) Then
                    Hist2 = " CEM BILHÕES"
                End If

                If (Val(Cent4) > 1) Then
                    Hist2 = Trim(Mid(Cente, Val(Cent4) * 12 - 11, 12)) & " BILHÕES"
                End If

                Flag = True

            Case Else
                Hist2 = ""
                Flag = True

        End Select

        If (Flag = False) Then
            If (Cent4 <> "0") Then
                Hist2 = Trim(Mid(Cente, Val(Cent4) * 12 - 11, 12)) & " E " & Hist2 & "BILHÕE"
            Else
                If (Hist2 = "UM") Then
                    Hist2 = Hist2 & " BILHÃO"
                Else
                    Hist2 = Hist2 & " BILHÕES"
                End If
            End If
        End If

        Flag = False

        Select Case Val(Dez3)
            Case 1 To 9
                Hist3 = Trim(Mid(Unid1, Val(Dez3) * 6 - 5, 6))
            Case 11 To 19
                Hist3 = Trim(Mid(Deze1, Val(Mid(Dez3, 2, 1)) * 9 - 8, 9))
            Case 10
                Hist3 = " DEZ"
            Case 20 To 99
                Hist3 = Trim(Mid(Deze2, Val(Mid(Dez3, 1, 1)) * 9 - 8, 9))
                If (Val(Mid(Dez3, 2, 1)) <> 0) Then
                    Hist3 = Hist3 & " E " & Trim(Mid(Unid1, Val(Mid(Dez3, 2, 1)) * 6 - 5, 6))
                End If
            Case 0
                If (Val(Cent3) = 0) Then
                    Hist3 = ""
                End If

                If (Val(Cent3) = 1) Then
                    Hist3 = " CEM MILHÕES"
                End If

                If (Val(Cent3) > 1) Then
                    Hist3 = Trim(Mid(Cente, Val(Cent3) * 12 - 11, 12)) & " MILHÕES"
                End If

                Flag = True

            Case Else
                Hist3 = ""
                Flag = True

        End Select

        If (Flag = False) Then
            If (Val(Cent3) <> 0) Then
                Hist3 = Trim(Mid(Cente, Val(Cent3) * 12 - 11, 12)) & " E " & Hist3 & " MILHÕES"
            Else
                If (Hist3 = "UM") Then
                    Hist3 = Hist3 & " MILHÃO"
                Else
                    Hist3 = Hist3 & " MILHÕES"
                End If
            End If

            If (Val(Cent1) <> 0 And (Val(Dez2) = 0 Or Val(Cent2) = 0)) Then
                Hist3 = Hist3 & " E"
            End If

            If (Val(Cent2) = 0 And Val(Dez2) <> 0) Then
                Hist3 = Hist3 & " E"
            End If

            If (Val(Cent2) <> 0 And Val(Dez2) = 0) Then
                Hist3 = Hist3 & " E "
            End If

        End If

        Flag = False

        Select Case Val(Dez2)

            Case 1 To 9
                Hist4 = Trim(Mid(Unid1, Val(Dez2) * 6 - 5, 6))

            Case 11 To 19
                Hist4 = Trim(Mid(Deze1, Val(Mid(Dez2, 2, 1)) * 9 - 8, 9))

            Case 10
                Hist4 = " DEZ"

            Case 20 To 99
                Hist4 = Trim(Mid(Deze2, Val(Mid(Dez2, 1, 1)) * 9 - 8, 9))
                If (Val(Mid(Dez2, 2, 1)) <> 0) Then
                    Hist4 = Hist4 & " E " & Trim(Mid(Unid1, Val(Mid(Dez2, 2, 1)) * 6 - 5, 6))
                End If

            Case 0
                If (Val(Cent2) = 0) Then
                    Hist4 = ""
                End If

                If (Val(Cent2) = 1) Then
                    Hist4 = " CEM MIL"
                    If (Val(Dez1) <> 0) Then
                        Hist4 = Hist4 & " E"
                    End If
                End If

                If (Val(Cent2) > 1) Then
                    Hist4 = Trim(Mid(Cente, Val(Cent2) * 12 - 11, 12)) & " MIL"
                    If (Val(Dez1)) Then
                        Hist4 = Hist4 & " E"
                    End If
                End If

                Flag = True

            Case Else
                Hist4 = ""
                Flag = True

        End Select

        If (Flag = False) Then
            If (Val(Cent2) <> 0) Then
                If (Val(Cent1) <> 0 And Val(Dez1 = 0)) Then
                    Hist4 = Trim(Mid(Cente, Val(Cent2) * 12 - 11, 12)) & " E " & Hist4 & " MIL E"
                Else
                    Hist4 = Trim(Mid(Cente, Val(Cent2) * 12 - 11, 12)) & " E " & Hist4 & " MIL"
                End If
            Else
                If ((Val(Dez1) = 0 And Val(Cent1) <> 0) Or (Val(Dez1) <> 0 And Val(Cent1) = 0)) Then
                    Hist4 = Hist4 & " MIL E"
                Else
                    Hist4 = Hist4 & " MIL"
                End If
            End If
        End If

        Flag = False

        Select Case Val(Dez1)
            Case 1 To 9
                Hist5 = Trim(Mid(Unid1, Val(Dez1) * 6 - 5, 6))

            Case 11 To 19
                Hist5 = Trim(Mid(Deze1, Val(Mid(Dez1, 2, 1)) * 9 - 8, 9))

            Case 10
                Hist5 = " DEZ"

            Case 20 To 99
                Hist5 = Trim(Mid(Deze2, Val(Mid(Dez1, 1, 1)) * 9 - 8, 9))
                If (Val(Mid(Dez1, 2, 1)) <> 0) Then
                    Hist5 = Hist5 & " E " & Trim(Mid(Unid1, Val(Mid(Dez1, 2, 1)) * 6 - 5, 6))
                End If

            Case 0
                If (Val(Cent1) = 0) Then
                    Hist5 = " " & WMoedap
                End If

                If (Val(Cent1) = 1) Then
                    Hist5 = " CEM " & WMoedap
                End If

                If (Val(Cent1) > 1) Then
                    Hist5 = Trim(Mid(Cente, Val(Cent1) * 12 - 11, 12)) & " " & WMoedap
                End If

                Flag = True

            Case Else
                Hist5 = ""
                Flag = True

        End Select

        If (Flag = False) Then
            If (Cent1 <> "0") Then
                Hist5 = Trim(Mid(Cente, Val(Cent1) * 12 - 11, 12)) & " E " & Hist5 & " " & WMoedap
            Else
                If (Hist5 = "UM") Then
                    Hist5 = Hist5 & " " & WMoeda
                Else
                    Hist5 = Hist5 & " " & WMoedap
                End If
            End If
        End If

        Flag = False

        Select Case Val(Dez0)
            Case 1 To 9
                Hist6 = Trim(Mid(Unid1, Val(Dez0) * 6 - 5, 6))

            Case 11 To 19
                Hist6 = Trim(Mid(Deze1, Val(Mid(Dez0, 2, 1)) * 9 - 8, 9))

            Case 10
                Hist6 = " DEZ"

            Case 20 To 99
                Hist6 = Trim(Mid(Deze2, Val(Mid(Dez0, 1, 1)) * 9 - 8, 9))
                If (Val(Mid(Dez0, 2, 1)) <> 0) Then
                    Hist6 = Hist6 & " E " & Trim(Mid(Unid1, Val(Mid(Dez0, 2, 1)) * 6 - 5, 6))
                End If

            Case Else
                Hist6 = ""
                Flag = True

        End Select

        If (Flag = False) Then
            If (Hist6 = "UM") Then
                Hist6 = Hist6 & " CENTAVO"
            Else
                Hist6 = Hist6 & " CENTAVOS"
            End If
        End If

        WExtenso = ""

        If (Len(Trim(Hist1)) > 1) Then
            WExtenso = Trim(Hist1)
        End If

        If (Len(Trim(Hist2)) > 1) Then
            If (Len(WExtenso) > 1) Then
                WExtenso = WExtenso & " " & Trim(Hist2)
            Else
                WExtenso = Hist2
            End If
        End If

        If (Len(Trim(Hist3)) > 1) Then
            If (Len(WExtenso) > 1) Then
                WExtenso = WExtenso & " " & Trim(Hist3)
            Else
                WExtenso = Hist3
            End If
        End If

        If (Len(Trim(Hist4)) > 1) Then
            If (Len(WExtenso) > 1) Then
                WExtenso = WExtenso & " " & Trim(Hist4)
            Else
                WExtenso = Hist4
            End If
        End If

        If (Len(Trim(Hist5)) > 1) Then
            If (Len(WExtenso) > 1) Then
                If (Hist5 = " " & WMoedap And Val(Dez2) = 0 And Val(Cent2) = 0) Then
                    WExtenso = WExtenso & " DE " & Hist5
                Else
                    WExtenso = WExtenso & " " & Trim(Hist5)
                End If
            Else
                If (Hist5 <> " " & WMoedap) Then
                    WExtenso = Hist5
                End If
            End If
        End If

        If (Len(Trim(Hist6)) > 1) Then
            If (Len(WExtenso) > 1) Then
                WExtenso = WExtenso & " E " & Trim(Hist6)
            Else
                WExtenso = WExtenso & Trim(Hist6)
            End If
        End If

        '    If (Mid(WExtenso, 1, 2)) = "UM" Then
        '        WExtenso = "H" & WExtenso
        '    End If

        fExtenso = WExtenso

        Return WExtenso
    End Function

    Private Shared Function Valor() As Object
        Throw New NotImplementedException
    End Function

    ''' <summary>
    ''' Retorna uma string no Formato Projeto: "Projeto" - Classe: "Classe" - Método: "Método()"
    ''' </summary>
    ''' <param name="info">A Instância de System.Reflection.MethodoBase.GetCurrentMethod()</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '<DebuggerStepThrough()>
    Public Shared Function FormataLocalException(ByVal info As MethodBase) As String
        Dim lst = info.GetParameters()
        Dim strVariavel As String = ""
        For Each item In lst
            If strVariavel <> "" Then strVariavel += ","
            strVariavel += item.Name
        Next

        Return (String.Format("Projeto: {0} Classe: {1} Método: {2}({3})", info.ReflectedType.Namespace.Substring(17), info.ReflectedType.Name, info.Name, strVariavel))

    End Function

    <DebuggerStepThrough()>
    Public Shared Function FormataLocalException(ByVal info As MethodBase, ByVal ex As Exception) As String
        Dim lst = info.GetParameters()
        Dim strVariavel As String = ""
        For Each item In lst
            If strVariavel <> "" Then strVariavel += ","
            strVariavel += item.Name
        Next

        Dim st As StackTrace = New StackTrace(ex, True)
        Dim frame As StackFrame = st.GetFrames().FirstOrDefault()
        Dim strFrames As String = ""

        strFrames += String.Format("{0}:{1}({2},{3})", frame.GetFileName(), frame.GetMethod().Name, frame.GetFileLineNumber(), frame.GetFileColumnNumber()) + vbCrLf

        Return (String.Format("Projeto: {0} Classe: {1} Método: {2}({3})", info.ReflectedType.Namespace.Substring(17), info.ReflectedType.Name, info.Name, strVariavel)) + vbCrLf + strFrames
    End Function

    ''' <summary>
    ''' Formata a data no formato selecionado, padrão yyyy-MM-dd
    ''' </summary>
    ''' <param name="data"></param>
    ''' <param name="formato">
    ''' 0 - yyyy-MM-dd
    ''' 1 - ddMMyyyy
    ''' 2- dd/MM/yyyy
    ''' 3 - DD/MM/YYYY HH:MM:SS
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DebuggerStepThrough()>
    Public Shared Function FormatDate(ByVal data As Date, Optional ByVal formato As Integer = 0) As String
        Select Case formato
            Case 0
                Return data.Year.ToString().PadLeft(4, "0") + "-" + data.Month.ToString().PadLeft(2, "0") + "-" + data.Day.ToString().PadLeft(2, "0")
            Case 1
                Return data.Day.ToString().PadLeft(2, "0") + data.Month.ToString().PadLeft(2, "0") + data.Year.ToString()
            Case 2
                Return data.Day.ToString().PadLeft(2, "0") + "/" + data.Month.ToString().PadLeft(2, "0") + "/" + data.Year.ToString()
            Case 3
                Return data.Day.ToString().PadLeft(2, "0") + "/" + data.Month.ToString().PadLeft(2, "0") + "/" + data.Year.ToString() + Space(1) +
                    data.Hour.ToString().PadLeft(2, "0") + ":" + data.Minute.ToString().PadLeft(2, "0") + ":" + data.Second.ToString().PadLeft(2, "0")
            Case Else
                Return data.ToShortTimeString()

        End Select

    End Function

    ''' <summary>
    ''' Verifica se o dado passado é um DbNull e retorna o tipo especificado.
    ''' </summary>
    ''' <param name="value">Expressão que deseja comparar</param>
    ''' <param name="t">Tipo 0-Numérico, 1-Texto, 2-boolean, 3- DateTime, 4- Nothing </param>
    ''' <returns>Retorno o padrão (0, " ", false)
    ''' referente ao indice especificado,
    ''' em caso de índice inválido o retorno será 0.</returns>
    ''' <remarks></remarks>
    Public Shared Function VerificaDbNull(ByVal value As Object, Optional ByVal t As Integer = 1)

        If (IsDBNull(value)) Then

            Select Case t
                Case 0
                    Return 0
                Case 1
                    Return ""
                Case 2
                    Return False
                Case 3
                    Return DateTime.Today()
                Case 4
                    Return Nothing
                Case Else
                    Return 0
            End Select
        Else
            Select Case t
                Case 0
                    If (value.ToString().Contains(",") OrElse value.ToString().Contains(".")) Then
                        Convert.ToDouble(value.ToString().Replace(".", ""))
                    Else
                        Convert.ToInt64(value)
                    End If
                Case 3
                    Return Convert.ToDateTime(value, Funcoes.Cultura)
            End Select
        End If

        Return value
    End Function

    ''' <summary>
    ''' Obtem a quantidade de dias uteis entre uma data e outra
    ''' </summary>
    ''' <param name="dtInicio"></param>
    ''' <param name="dtFim"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function QuantDiasUteis(ByVal dtInicio As DateTime, ByVal dtFim As DateTime) _
        As Integer

        Dim intSemanas As Integer
        Dim varDataCont As Object
        Dim intFimDias As Integer

        intSemanas = DateDiff("w", dtInicio, dtFim)
        varDataCont = DateAdd("ww", intSemanas, dtInicio)
        intFimDias = 0

        Do While varDataCont < dtFim

            Using con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
                con.Open()

                Dim cmd As New SqlCommand("P_VerificaFeriado", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.Add(New SqlParameter("@data", varDataCont))

                If (DatePart("w", varDataCont) <> 1 _
                   AndAlso DatePart("w", varDataCont) <> 7 _
                   AndAlso cmd.ExecuteReader(CommandBehavior.CloseConnection).Read() = False) Then
                    intFimDias += 1
                End If
                varDataCont = DateAdd("d", 1, varDataCont)

                con.Close()
            End Using

        Loop


        Return intSemanas * 5 + intFimDias
    End Function

    ''' <summary>
    ''' Verifica se a tecla digita corresponde a alguma das teclas passadas no array. Caso sim retorna true. 
    ''' </summary>
    ''' <param name="teclas">Lista de código ASCII das teclas que não podem ser digitadas</param>
    ''' <param name="tecla">Tecla digitada</param>
    ''' <returns>True se houver sucesso nas comparações, false caso não.</returns>
    ''' <remarks></remarks>
    Public Shared Function VerificarTeclas(ByVal teclas() As Integer, ByVal tecla As String)
        Dim blnRetorno As Boolean = False
        For Each caractere As Integer In teclas
            If caractere = tecla Then
                blnRetorno = True
                Exit For
            End If
        Next
        Return blnRetorno
    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary> Edita uma error constante para obter a descrição personalizada. </summary>
    '''
    ''' <remarks> Elacerda, 16/05/2013. </remarks>
    '''
    ''' <param name="ec"> ErrorConstant </param>
    ''' <param name="strTag"> Tag a ser retirada </param>
    ''' <param name="novoValor"> Valor a ser substituido </param>
    '''
    ''' <returns> . </returns>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared Function EditaErrorConstant(ByVal ec As ErrorConstants, ByVal strTag As String, ByVal novoValor As String)

        ec.Descricao = ec.Descricao.Replace(strTag, novoValor)

        Return ec

    End Function

    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    ''' <summary> Retorno o próximo dia útil a partir do número de dias esperados. </summary>
    '''
    ''' <remarks> Elacerda, 27/05/2013. </remarks>
    '''
    ''' <param name="_DataInicio">      Data de início da contagem </param>
    ''' <param name="qtdeDiasCalcular"> Quantidade minima de dias para o próximo dia util </param>
    '''
    ''' <returns> . </returns>
    '''////////////////////////////////////////////////////////////////////////////////////////////////////
    Public Shared Function CalculaDataUtil(ByVal _DataInicio As DateTime, ByVal qtdeDiasCalcular As Integer)

        Dim con As New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim dataRetorno As New DateTime()

        Try
            con.Open()

            Dim cmd As New SqlCommand("P_CalcDtUtil", con)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add(New SqlParameter("@DtInicial", _DataInicio))
            cmd.Parameters.Add(New SqlParameter("@QtdeDiasUteis", qtdeDiasCalcular))

            Dim rdr As SqlDataReader = cmd.ExecuteReader()

            If (rdr.Read()) Then
                dataRetorno = Convert.ToDateTime(rdr(0))
            Else
                dataRetorno = DateAdd("DAY", qtdeDiasCalcular, _DataInicio)
            End If

        Catch ex As Exception

            Dim _Retorno As New Retorno()
            With _Retorno
                .Sucesso = False
                .NumErro = ErrorConstants.EXCEPTION_METODO_CALCULADATAUTIL.Id
                .MsgErro = ErrorConstants.EXCEPTION_METODO_CALCULADATAUTIL.Descricao + ex.Message
                .TipoErro = DadosGenericos.TipoErro.Arquitetura
                .ImagemErro = DadosGenericos.ImagemRetorno.Erro
            End With

            dataRetorno = DateAdd("DAY", qtdeDiasCalcular, _DataInicio)

            Throw New Exception(_Retorno.MsgErro)
        Finally
            con.Close()
        End Try


        Return dataRetorno
    End Function

    ''' <summary>
    ''' Cria um Format Provider de data, padrão pt-br.
    ''' *** Utilizar sempre ao converter uma data! ***
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CriarFormatProvider() As IFormatProvider
        Return New CultureInfo("pt-br")
    End Function

    ''' <summary>
    ''' Seta a cultura da aplicação
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub SetCultura()
        System.Threading.Thread.CurrentThread.CurrentCulture = Funcoes.Cultura
    End Sub

    Public Shared Function IsDDD9Digito(ByVal ddd As String) As Boolean

        Dim ddds As String() = {"11", "12", "13", "14", "15", "16", "17", "18", "19",
                                "21", "22", "24", "27", "28",
                                "31", "32", "33", "34", "35", "36", "37", "38", "41",
                                "42", "43", "44", "45", "46", "47", "48", "49",
                                "51", "53", "54", "55",
                                "61", "62", "63", "64", "65", "66", "67", "68", "69",
                                "71", "72", "73", "74", "75", "76", "77", "78", "79",
                                "81", "82", "83", "84", "85", "86", "87", "88", "89",
                                "91", "92", "93", "94", "95", "96", "97", "98", "99"}
        Return ddds.Contains(ddd.Trim())

    End Function

    Public Shared Function IsFormTodasTeclas(ByVal strFormName As String) As Boolean
        Dim arrForms As String() = {"TELESYSTEMFATNFSERVAVULSO1", "TELESYSTEMFATNFSERVEMISSAONFMONITORIA1", "TELESYSTEMFATNFVENDAAVULSO1", "TELESYSTEMFATNFVENDAEMISSAO1", "TELESYSTEMCTARECEBERLANCMANUTENCAOEDIT1",
         "TELESYSTEMHELPDESKHISTORICO1", "TELESYSTEMHELPDESKEXECUCAO1", "TELESYSTEMCONTABCCUSTOSCATEG1", "TELESYSTEMHELPDESKCADMOTIVOSALTERAR1", "TELESYSTEMHELPDESKCADATIVOSALTERAR1",
         "TELESYSTEMHELPDESKCADSTATUSATIVOSALTERAR1", "TELESYSTEMCADASTROSDEREPRESENTANTESINCLUSAOALTERACAO1", "TELESYSTEMHELPDESKADMIN1", "TELESYSTEMCADASTRODEPARTAMENTOINCLUIRALTERAR1", "TELESYSTEMCADASTRODEPTOSETORINCLUIRALTERAR1",
         "TELESYSTEMCOMERCIALCONFIGINDICACOES1", "TELESYSTEMHELPDESKPROJETOSALTERAR1", "TELESYSTEMCADASTROMOTIVOSINCLUIRALTERAR1", "TELESYSTEMCADASTROMOTIVOS1", "TELESYSTEMPESQMANUTENCAO1", "TELESYSTEMSOLICINTCOMPRAS1",
         "TELESYSTEMINSTGERENCAREARADIO", "TELESYSTEMINSTGERENCAREARADIOALTERAR", "TELESYSTEMINSTGERENCAREARADIOEXCECOES", "TELESYSTEMINSTGERENCAREARADIOEXCECOESALTERAR", "TELESYSTEMGERENCSOLICHEALTERAR1",
         "TELESYSTEMINSTTIPOLINKINSALT1", "TELESYSTEMINSTLINKMYTELEVIEWINSALT1", "TELESYSTEMADMMATRICULAINSALT1", "TELESYSTEMUSUARIOMANUTENCAO1", "TELESYSTEMTIVERSISTEMASINSALT1", "TELESYSTEMCOMERCIALCADMOTCLASSCONTATOEDI1",
         "TELESYSTEMCOMERCIALCADMOTCLASSCONTATO1", "TELESYSTEMINDICACAO1", "TELESYSTEMCONTATOSINDICACAO1", "TELESYSTEMCADASTROORIGEMALTERAR1", "TELESYSTEMCADASTROORIGEM1", "TELESYSTEMCOMCADCAMPANHAS1", "TELESYSTEMCOMCADCAMPANHASEDI1", "TELESYSTEMCLIENTE1",
         "TELESYSTEMCOMCADCATEGORIAORIGEM1", "TELESYSTEMCOMCADCATEGORIAORIGEMEDI1", "TELESYSTEMCOMCADTIPOORIGEM1", "TELESYSTEMCOMCADTIPOORIGEMEDI1", "TELESYSTEMPOSVENDARETEXTCADAREA1", "TELESYSTEMPOSVENDARETEXTCADAREAEDI1", "TELESYSTEMREVERSAO1", "TELESYSTEMFINESTQBALANCO1", "TELESYSTEMCOMCADMETAOPERADORESMKT1",
         "TELESYSTEMCOMCADCORRETORAEDI1", "TELESYSTEMCOMCADCORRETORA1", "TELESYSTEMCOMCADCORRETOR1", "TELESYSTEMCOMCADCORRETOREDI1", "TELESYSTEMCOBRANCAADMFILASMANUAIS1", "TELESYSTEMCOMCADTELEVENDASBLACKLIST1", "TELESYSTEMPOSVENDASCONFIG1", "TELESYSTEMALTEMAILSHISTCONTATO1", "TELESYSTEMALTCADCAMPANHASUPSELLING1", "TELESYSTEMCOMCADVDNEDIGENESYS1", "TELESYSTEMCOMCADVDNEDIGENESYS1",
         "TELESYSTEMMKTUPSELLINGAUDITORIAVISITASEDI1", "TELESYSTEMCTARECEBERDIVERGENTESEDI1", "TELESYSTEMCADASTROSCRIPTCALLCENTERALTERAR1", "TELESYSTEMPOSVENDASANALISECREDITODET1", "TELESYSTEMSACCADAVISOSINCLALT1", "TELESYSTEMLISTAGEMROBINSON1"}

        Return arrForms.Contains(Trim(strFormName).ToUpper)

    End Function

    Public Shared Function IsPlacaValida(ByVal value As String) As Boolean

        Dim regex As New Text.RegularExpressions.Regex("^[a-zA-Z]{3}\-\d{4}$")

        If (regex.IsMatch(value)) Then
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Verifica se arquivo existe, se não existe apaga
    ''' </summary>
    ''' <param name="strPath"></param>
    ''' <remarks></remarks>
    Public Shared Function ApagaArquivo(ByVal strPath As String) As Boolean

        Try
            If (File.Exists(strPath)) Then

                File.Delete(strPath)

            End If

        Catch ex As IOException
            Throw ex
        End Try

        Return True

    End Function

    Public Shared Function ValidaData(ByVal dataObject As Date?) As Boolean

        If (dataObject Is Nothing OrElse dataObject.Value.ToString() = "01/01/00001 00:00:00" OrElse dataObject.Value = Date.MinValue) Then
            Return False
        End If


        Return True

    End Function

    Public Shared Function IsComputadorSetorSistemas(ByVal machineName As String) As Boolean

        Dim machineNames As String() = {"D6JXB342", "D40PM1P1", "DCF0GXP1", "DDF0GXP1"}

        Return machineNames.Contains(machineName)

    End Function

    ''' <summary>
    ''' Retira valor truncado de um número com decimais (arredonda para baixo)
    ''' </summary>
    ''' <remarks> Renato - 21/05/2014 </remarks>
    Public Shared Function TiraVlrTruncado(ByVal vlr As Double, ByVal precisao As Integer)

        'Calcula o fator multiplicador
        Dim fator As Double = Double.Parse(Math.Pow(10D, precisao))

        'Obtém o valor junto com a parte inteira
        Dim vlrTruncado As Double = Math.Floor(vlr * fator)

        'Retorna o valor com as casas decimais
        Return (Math.Floor((Math.Round(vlrTruncado, precisao))) / fator)

    End Function

    ''' <summary>
    ''' Cria retorno padrão erro funcional.
    ''' </summary>
    ''' <remarks> Renato - 30/06/2014 </remarks>
    Public Shared Function RetornoFunc(ByVal msgErro As String)
        Dim retorno As New Retorno
        retorno.Sucesso = False
        retorno.TipoErro = DadosGenericos.TipoErro.Funcional
        retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
        retorno.MsgErro = msgErro
        Return retorno
    End Function

    ''' <summary>
    ''' Cria retorno padrão erro arquitetura.
    ''' </summary>
    ''' <remarks> Renato - 30/06/2014 </remarks>
    Public Shared Function RetornoArq(ByVal msgErro As String, ByVal numErro As String)
        Dim retorno As New Retorno
        retorno.Sucesso = False
        retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
        retorno.ImagemErro = DadosGenericos.ImagemRetorno.Erro
        retorno.MsgErro = msgErro
        retorno.NumErro = numErro
        Return retorno
    End Function

    Public Shared Function RetornoArq(ByVal msgErro As String, ByVal numErro As String, ByVal _ErrorImg As DadosGenericos.ImagemRetorno)
        Dim retorno As New Retorno
        retorno.Sucesso = False
        retorno.TipoErro = DadosGenericos.TipoErro.Funcional
        retorno.ImagemErro = _ErrorImg
        retorno.MsgErro = msgErro
        retorno.NumErro = numErro
        Return retorno
    End Function

    Public Shared Function UppercaseFirstLetter(ByVal strTexto As String) As String
        ' Test for nothing or empty.
        If String.IsNullOrEmpty(strTexto) Then
            Return strTexto
        End If

        ' Convert to character array.
        Dim array() As Char = strTexto.ToCharArray

        ' Uppercase first character.
        array(0) = Char.ToUpper(array(0))

        ' Return new string.
        Return New String(array)
    End Function

    Public Shared Function IsRedeVerisure() As Boolean
        Dim strIsVerisure As String = LerRegistroDoWindows("Configurações", "isVerisure", "0")
        Return ("1".Equals(strIsVerisure))
    End Function
#End Region

    Private Shared Function EnvialEmailWebServiceMobilePronto(ByVal strCredencial As String, ByVal strToken As String, ByVal strAssunto As String,
                                                       ByVal strDe As String, ByVal strBody As String, ByVal strTypeMessage As String,
                                                       ByVal strAutenticacao As String, ByVal strPara1 As String, Optional ByVal strPara2 As String = "",
                                                       Optional ByVal strBcc1 As String = "", Optional ByVal strBcc2 As String = "") As Boolean
        Try
            'Dim wsEnvioEmail As New WSEmailMobilePronto.MPGatewayMail
            Dim wsEnvioEmail As New WSMobilePronto10072015.MPGatewayEmail
            'Dim wsEnvioEmail As New WSEmailPitchWink.PWGatewayEmailSoapClient

            If strPara1 <> "" Then
                Dim strEmail() As String = strPara1.Split(";")
                For Each _Para In strEmail

                    'Dim strSucesso As String = wsEnvioEmail.MPGEmail_SendEmailViaSMTP(strCredencial, strToken, strAssunto, strDe, strBody, strTypeMessage, strAutenticacao, _Para, "", "", "")

                    'Select Case strSucesso
                    '    'Tipos de Retorno: 
                    '    '----------------------------------------------------------------------------------------------------------------------------------
                    '    '000 - Email enviado com Sucesso.
                    '    '----------------------------------------------------------------------------------------------------------------------------------
                    '    '001 - Credencial Inválida.
                    '    '002 - Token Inválido.
                    '    '003 - Message Vazia.
                    '    '004 - Type_Message inválida.
                    '    '005 - Autenticacao inválida.
                    '    '006 - To1 - obrigatório.
                    '    '-----------------------------------------------------------------------------------------------------------------------------------
                    '    '800 - Erro genérico de envio.
                    '    '900 - Erro interno.
                    '    '-----------------------------------------------------------------------------------------------------------------------------------
                    '    Case "002", "003", "004", "005", "006", "800", "900"
                    '        Throw New Exception(strSucesso)
                    'End Select

                    'Nova chamada para o Web Service atualizado da MobiPronto, para envio de e-mail no caso da alog não estar disponível
                    Dim strSucesso As String = wsEnvioEmail.MPG_Send_Email("TXT", "4.00|" & strCredencial & "|" & strToken & "||WEBAPI||" & strDe & "||" & _Para & "|||" & strTypeMessage & "|" & strBody)

                    Select Case strSucesso
                        Case "001", "005", "006", "008", "009", "010", "013", "015", "018", "019", "020", "021", "022", "023", "024", "025", "800", "900", "901"
                            Throw New Exception(strSucesso)
                    End Select

                Next
            End If

            Return True
        Catch ex As Exception
            Throw New Exception("Erro ao enviar e-mail! Código de erro: " & ex.Message & vbCrLf & " Caso o problema persista favor entrar em contato com o departamento de sistemas!")
        End Try
    End Function

    Public Shared Function CriaRelatorioXLSX_Extenso(ByVal dtData As DataTable, ByVal Path As SaveFileDialog, Optional ByVal dtDataTables() As DataTable = Nothing) As Retorno

        Dim retorno As New Retorno With {.Sucesso = False, .MsgErro = "Problemas ao gerar relatório!"}

        Try

            'Caso não seja pelo array, valida o DataTable individual
            If IsNothing(dtDataTables) Then
                'Valida se a DataTable está preenchida.
                If dtData.Rows.Count <= 0 Then
                    Return Funcoes.RetornoFunc("Nenhum registro encontrado para gerar o relatório!")
                End If
            End If

            Dim excelPackage As New ExcelPackage

            If IsNothing(dtDataTables) Then
                'Adiciona o data table ao worksheet. Vale lembrar que o datatable já deve vir com um NOME.
                Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets.Add(dtData.TableName)
                worksheet.Cells("A1").LoadFromDataTable(dtData, True)
            Else
                For Each _Dt As DataTable In dtDataTables
                    Dim worksheet As ExcelWorksheet = excelPackage.Workbook.Worksheets.Add(_Dt.TableName)
                    worksheet.Cells("A1").LoadFromDataTable(_Dt, True)
                Next
            End If

            Dim fi As New FileInfo(Path.FileName)

            excelPackage.SaveAs(fi)

            excelPackage.Dispose()

            retorno.Sucesso = True
            retorno.TipoErro = DadosGenericos.TipoErro.None

        Catch ex As Exception
            retorno.Sucesso = False
            retorno.TipoErro = DadosGenericos.TipoErro.Arquitetura
            retorno = Funcoes.RetornoArq(ex.Message(), "CriaRelatorioXLSX_Extenso")
        End Try

        Return retorno

    End Function

    Public Shared Function CriaRelatorioXLSX(ByVal dtData As DataTable,
                                             ByVal strCaminho As String,
                                             Optional ByVal dtDataTables() As DataTable = Nothing) As Retorno
        Dim _Retorno As New Retorno With {.Sucesso = False, .MsgErro = "Problemas ao gerar relatório!"}
        Try
            'Caso não seja pelo array, valida o DataTable individual
            If IsNothing(dtDataTables) Then
                'Valida se a DataTable está preenchida.
                If dtData.Rows.Count <= 0 Then Throw New Exception("Nenhum registro encontrado para gerar o relatório!")
            End If

            'Declara o objeto do tipo XLWorkbook
            Dim xlwPlanilha As New ClosedXML.Excel.XLWorkbook

            If IsNothing(dtDataTables) Then
                'Adiciona o data table ao worksheet. Vale lembrar que o datatable já deve vir com um NOME.
                xlwPlanilha.Worksheets.Add(dtData)
            Else
                For Each _Dt As DataTable In dtDataTables
                    xlwPlanilha.Worksheets.Add(_Dt)
                Next
            End If

            'Salva a planilha no caminho informado com a extensão informada.
            xlwPlanilha.SaveAs(strCaminho)


            'Dá um dispose no objeto
            xlwPlanilha.Dispose()

            'Define o retorno como true e mensagem de erro vazia
            _Retorno.Sucesso = True
            _Retorno.MsgErro = ""
        Catch ex As Exception
            _Retorno.MsgErro = ex.Message
            CriaLog(Application.StartupPath & "\RelatorioErros.txt", ex.Message)
        End Try
        Return _Retorno
    End Function

    'Public Shared Function GeraRelatExcel(ByVal dt As DataTable, ByVal form As RadForm) As Retorno
    '    Dim retorno As New Retorno
    '    Try
    '        Dim _Ret As New Retorno With {.Sucesso = True}

    '        ' Define o caminho para salvar o relatório
    '        Dim strPath As String = GetEnderecoArquivoXLSX()

    '        ' Se o caminho tiver vazio joga exception
    '        If Trim(strPath) = "" Then Throw New Exception(ErrorConstants.NENHUM_ARQUIVO_SELECIONADO.Descricao)

    '        ' Cria o relatório
    '        form.Cursor = Cursors.WaitCursor
    '        _Ret = Funcoes.CriaRelatorioXLSX(dt, strPath)
    '        form.Cursor = Cursors.Default

    '        ' Throw caso haja erro
    '        If Not _Ret.Sucesso Then Throw New Exception(_Ret.MsgErro)

    '        If RadMessageBox.Show("Relatório gerado com sucesso!" & vbCrLf & "Deseja visualizar agora?", "Relatório", MessageBoxButtons.YesNo, RadMessageIcon.Info) = vbYes Then
    '            System.Diagnostics.Process.Start(strPath)
    '        End If

    '        retorno.Sucesso = True

    '    Catch ex As Exception
    '        'If ex.Message.Contains("Nenhum arquivo selecionado!") Then
    '        '07/08/2017 - Fernando
    '        'Corrigido para evitar envio de e-mails de erro que deveriam ser apenas funcionais
    '        If IsErroFuncional(ex.Message) Then
    '            retorno = RetornoFunc(ex.Message)
    '        Else
    '            retorno = RetornoArq(ex.Message, "GeraRelExc")
    '        End If
    '    End Try

    '    Return retorno

    'End Function

    'Public Shared Function GeraRelatExcelOPENXML(ByVal dt As DataTable, ByVal form As RadForm) As Retorno
    '    Dim retorno As New Retorno
    '    Try
    '        Dim _Ret As New Retorno With {.Sucesso = True}

    '        ' Define o caminho para salvar o relatório
    '        Dim strPath As String = GetEnderecoArquivoXLSX()

    '        ' Se o caminho tiver vazio joga exception
    '        If Trim(strPath) = "" Then Throw New Exception(ErrorConstants.NENHUM_ARQUIVO_SELECIONADO.Descricao)

    '        ' Cria o relatório
    '        form.Cursor = Cursors.WaitCursor
    '        _Ret = CreateExcelFile.CreateExcelDocument(dt, strPath)
    '        form.Cursor = Cursors.Default

    '        ' Throw caso haja erro
    '        If Not _Ret.Sucesso Then Throw New Exception(_Ret.MsgErro)

    '        If RadMessageBox.Show("Relatório gerado com sucesso!" & vbCrLf & "Deseja visualizar agora?", "Relatório", MessageBoxButtons.YesNo, RadMessageIcon.Info) = vbYes Then
    '            System.Diagnostics.Process.Start(strPath)
    '        End If

    '        retorno.Sucesso = True

    '    Catch ex As Exception
    '        'If ex.Message.Contains("Nenhum arquivo selecionado!") Then
    '        '07/08/2017 - Fernando
    '        'Corrigido para evitar envio de e-mails de erro que deveriam ser apenas funcionais
    '        If IsErroFuncional(ex.Message) Then
    '            retorno = RetornoFunc(ex.Message)
    '        Else
    '            retorno = RetornoArq(ex.Message, "GeraRelExcOPXML")
    '        End If
    '    End Try

    '    Return retorno

    'End Function

    Public Shared Function getEnderecoArquivo_SaveFileDialog(ByVal strFilter As String, ByVal strInitialDirectory As String) As SaveFileDialog

        Dim ofdArquivo As New System.Windows.Forms.SaveFileDialog
        Dim strCaminho As String = ""
        ofdArquivo.Filter = strFilter
        ofdArquivo.InitialDirectory = strInitialDirectory
        ofdArquivo.ShowDialog()

        Return ofdArquivo

    End Function

    Public Shared Function GetEnderecoArquivoXLSX() As String
        Dim ofdArquivo As New System.Windows.Forms.SaveFileDialog
        Dim strCaminho As String = ""
        ofdArquivo.Filter = "xlsx files (*.xlsx)|*.xlsx"
        ofdArquivo.InitialDirectory = "C:\"
        ofdArquivo.ShowDialog()

        If ofdArquivo.FileName <> "" Then
            strCaminho = ofdArquivo.FileName
        End If

        Return strCaminho
    End Function

    ''' <summary>
    ''' Importa um arquivo excel do computador e retorna um data table.
    ''' </summary>
    ''' <param name="strFilePath">Caminho do arquivo</param>
    ''' <param name="strIsHDR">Informação se contém header ou não na planilha. "true" ou "false"</param>
    ''' <returns>Um data table contendo os dados da planilha</returns>
    ''' <remarks></remarks>
    Public Shared Function ImportarExcel(ByVal strFilePath As String, ByVal strIsHDR As String) As DataTable
        'Dim F As New Funcoes
        'Dim dtExcel As DataTable = F.ImportarExcel_Ver_02(strFilePath)
        'Return dtExcel
        'Define as variáveis de conexão
        Dim connExcel As OleDbConnection
        Dim cmdExcel As New OleDbCommand
        Dim odaAdapter As New OleDbDataAdapter
        Dim dtTable As New DataTable
        Dim strConnection As String = ""


        Try
            'Valida se existe um caminho
            If Trim(strFilePath) = "" Then Throw New FileNotFoundException("Caminho especificado não é válido!")

            'Determina a extensão do arquivo com base no caminho completo
            Dim strExtension As String = System.IO.Path.GetExtension(strFilePath)

            'Valida a extensão do arquivo
            If Trim(strExtension) <> ".xls" And Trim(strExtension) <> ".xlsx" Then Throw New FileNotFoundException("O tipo de arquivo selecionado não é válido, certifique-se de escolher uma planilha em excel.")

            'A conection string muda de acordo com a extensão(versão)
            Select Case strExtension
                Case ".xls"
                    'Excel 97-03
                    strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'"
                    Exit Select
                Case ".xlsx"
                    'Excel 07
                    strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'"
                    Exit Select
            End Select

            'Define a string de conexão na conexão
            strConnection = String.Format(strConnection, strFilePath, strIsHDR)
            connExcel = New OleDbConnection(strConnection)

            cmdExcel.Connection = connExcel

            'Pega o nome da primeira aba
            connExcel.Open()
            Dim dtExcelSchema As DataTable
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
            Dim SheetName As String = dtExcelSchema.Rows(0)("TABLE_NAME").ToString()
            connExcel.Close()

            'Lê as informações da primeira aba
            connExcel.Open()
            cmdExcel.CommandText = "SELECT * From [" & SheetName & "]"
            odaAdapter.SelectCommand = cmdExcel
            odaAdapter.Fill(dtTable)
            connExcel.Close()

        Catch ex As FileNotFoundException
            Throw New FileNotFoundException(ex.Message)
        Catch ex As Exception
            Throw ex
        End Try
        Return dtTable
    End Function

    'Private Function ImportarExcel_Ver_02(ByVal strFilePath As String) As DataTable

    '    If Trim(strFilePath) = "" Then Throw New FileNotFoundException("Caminho especificado não é válido!")

    '    Dim excelApp As ExcelApp.Application = New ExcelApp.Application()
    '    Dim myNewRow As DataRow
    '    Dim myTable As DataTable

    '    If (IsNothing(excelApp)) Then
    '        MessageBox.Show("Excel não instalado!")
    '        Return Nothing
    '    End If

    '    Dim excelBook As ExcelApp.Workbook = excelApp.Workbooks.Open(strFilePath)
    '    Dim excelSheet As ExcelApp._Worksheet = excelBook.Sheets(1)
    '    Dim excelRange As ExcelApp.Range = excelSheet.UsedRange

    '    Dim rows As Integer = excelRange.Rows.Count
    '    Dim cols As Integer = excelRange.Columns.Count

    '    myTable = New DataTable("MyDataTable")
    '    myTable.Columns.Add(CType("Unidade", String))
    '    myTable.Columns.Add(CType("Titulo", String))
    '    myTable.Columns.Add(CType("Banco Portador", String))
    '    myTable.Columns.Add(CType("Agência Portador", String))
    '    myTable.Columns.Add(CType("Conta Portador", String))
    '    myTable.Columns.Add(CType("Dt. Emissão", String))

    '    'Dim i As Integer = 2
    '    'For Each row As ExcelApp.Range In excelRange.Rows
    '    '    myNewRow = myTable.NewRow()
    '    '    myNewRow("Unidade") = excelRange.Cells(i, 1).Value2.ToString()
    '    '    myNewRow("Titulo") = excelRange.Cells(i, 2).Value2.ToString()
    '    '    myNewRow("Banco Portador") = excelRange.Cells(i, 3).Value2.ToString()
    '    '    myNewRow("Agência Portador") = excelRange.Cells(i, 4).Value2.ToString()
    '    '    myNewRow("Conta Portador") = excelRange.Cells(i, 5).Value2.ToString()
    '    '    myNewRow("Dt. Emissão") = excelRange.Cells(i, 6).Value2.ToString()
    '    '
    '    '    myTable.Rows.Add(myNewRow)
    '    '    i += 1
    '    'Next

    '    For i = 2 To rows
    '        myNewRow = myTable.NewRow()
    '        myNewRow("Unidade") = excelRange.Cells(i, 1).Value2.ToString()
    '        myNewRow("Titulo") = excelRange.Cells(i, 2).Value2.ToString()
    '        myNewRow("Banco Portador") = excelRange.Cells(i, 3).Value2.ToString()
    '        myNewRow("Agência Portador") = excelRange.Cells(i, 4).Value2.ToString()
    '        myNewRow("Conta Portador") = excelRange.Cells(i, 5).Value2.ToString()
    '        myNewRow("Dt. Emissão") = excelRange.Cells(i, 6).Value2.ToString()

    '        myTable.Rows.Add(myNewRow)
    '    Next

    '    excelApp.Quit()
    '    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)

    '    Return myTable
    'End Function

    ''' <summary>
    ''' Valida o header de um datatable baseado nas colunas informadas
    ''' </summary>
    ''' <param name="dtTable">Datatable</param>
    ''' <param name="strColunas">Nome das colunas separadas por ";"</param>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Shared Function ValidaHeaderDataTable(ByVal dtTable As DataTable, ByVal strColunas As String) As Boolean
        Dim arrColunas As String() = strColunas.Split(";")

        If arrColunas.Count < 1 Then Return False

        For Each _Coluna As String In arrColunas
            If IsNothing(dtTable.Columns(_Coluna)) Then Return False
        Next
        Return True
    End Function

    ''' <summary>
    ''' Valida se a linha passada está com todas as colunas em branco
    ''' </summary>
    ''' <param name="_Row">DataRow</param>
    ''' <param name="strColunas">Nome das colunas separadas por ";"</param>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Shared Function ValidaLinhaDataTableEmBranco(ByVal _Row As DataRow, ByVal strColunas As String) As Boolean
        Dim arrColunas As String() = strColunas.Split(";")

        If arrColunas.Count < 1 Then Return False

        For Each _Coluna As String In arrColunas
            If _Row.Item(_Coluna).ToString().Trim() <> "" Then Return True
        Next

        Return False
    End Function

    Public Shared Function GetExceptionLineNumber(ByVal stkTrace As StackTrace) As String
        Dim strLinhas As String = ""
        For Each _Frame As StackFrame In stkTrace.GetFrames
            If strLinhas = "" Then
                strLinhas = _Frame.GetFileLineNumber
            Else
                strLinhas += ", " & _Frame.GetFileLineNumber
            End If
        Next
        Return IIf(strLinhas = "", "N/I", strLinhas)
    End Function

    Public Shared Function getEnderecoAbrirArquivo(ByVal strFilter As String,
                                        ByVal strInitialDirectory As String)
        Dim sfdOpen As New OpenFileDialog
        Dim strCaminho As String = ""

        sfdOpen.Filter = strFilter
        sfdOpen.InitialDirectory = strInitialDirectory
        sfdOpen.ShowDialog()
        If sfdOpen.FileName <> "" Then
            strCaminho = sfdOpen.FileName
        End If

        Return strCaminho
    End Function

    Public Shared Function getEnderecoArquivo(ByVal strFilter As String,
                                        ByVal strInitialDirectory As String) As String
        Dim ofdArquivo As New System.Windows.Forms.SaveFileDialog
        Dim strCaminho As String = ""
        ofdArquivo.Filter = strFilter
        ofdArquivo.InitialDirectory = strInitialDirectory
        ofdArquivo.ShowDialog()

        If ofdArquivo.FileName <> "" Then
            strCaminho = ofdArquivo.FileName
        End If

        Return strCaminho
    End Function

    Public Shared Function data_Util_SabDomFer(ByVal dtDtInicial As DateTime,
                                               ByVal intSab As Integer,
                                               ByVal intDom As Integer,
                                               ByVal intFer As Integer) As DateTime
        Dim dtUtil As DateTime
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Dim command As SqlCommand = New SqlCommand("SELECT dbo.data_Util_SabDomFer ('" & Format(dtDtInicial, "yyyy-MM-dd HH:mm") &
                                                   "', " & intSab & ", " & intDom & ", " & intFer & ")", connection)
        command.CommandType = CommandType.Text
        Try
            connection.Open()
            Using rdr As SqlDataReader = command.ExecuteReader()
                If rdr.HasRows Then
                    rdr.Read()
                    dtUtil = Convert.ToDateTime(rdr(0))
                End If
                rdr.Close()
            End Using
        Catch ex As Exception
            Throw ex
        End Try


        Return dtUtil
    End Function

    Public Shared Function IsErroFuncional(ByVal strMsgErro As String) As Boolean
        If strMsgErro.ToUpper.Trim.Contains("NENHUM ARQUIVO SELECIONADO") Then Return True
        If strMsgErro.ToUpper.Trim.Contains("THE PROCESS CANNOT ACCESS THE FILE") Then Return True
        If strMsgErro.ToUpper.Trim.Contains("NENHUM REGISTRO ENCONTRADO") Then Return True
        If strMsgErro.ToUpper.Trim.Contains("O PROCESSO NÃO PODE ACESSAR O ARQUIVO") Then Return True

        Return False
    End Function

    'Public Shared Function ConverteRetornoVWS(ByVal _RetornoVWS As Verisure.Vws.Common.Retorno) As Retorno
    '    Dim _RetornoTELE As New Retorno
    '    With _RetornoTELE
    '        .Sucesso = _RetornoVWS.Sucesso
    '        .TipoErro = _RetornoVWS.TipoErro
    '        .NumErro = _RetornoVWS.NumErro
    '        .MsgErro = _RetornoVWS.MsgErro
    '    End With
    '    Return _RetornoTELE
    'End Function

    Public Shared Function IndicacaoisOrigemClienteIndicador(ByVal strRamo As String)
        Select Case strRamo.ToUpper()
            Case "SITE_PA"
                Return True
            Case "INDIQUE UM AMIGO"
                Return True
            Case "INDIQUE UM AMIGO II"
                Return True
        End Select
        Return False
    End Function

    Public Shared Function IndicacaoisOrigemClientedaBase(ByVal strRamo)
        Select Case strRamo.ToUpper()
            Case "MALA DIRETA VIDEO COMBO JULHO 2014"
                Return True
            Case "CLIENTE"
                Return True
            Case "EMAIL MKT TELE VIEW"
                Return True
            Case "E-MAIL MKT  AUTOMACAO"
                Return True
            Case "EMAIL MKT VIDEO COMBO NOVAS SOLUCOES"
                Return True
        End Select
        Return False
    End Function

    Public Shared Function IsFormAberto(ByVal Form As Form) As Boolean
        If Application.OpenForms.OfType(Of Form).Contains(Form) Then
            Return True
        End If
        Return False
    End Function

    Private Shared Function wrapValue(ByVal value As String, ByVal group As String, ByVal separator As String) As String
        If value.Contains(separator) Then
            If value.Contains(group) Then
                value = value.Replace(group, group + group)
            End If
            value = group & value & group
        End If
        Return value
    End Function

    Public Shared Function ExportToCSV(ByVal dtable As DataTable, ByVal fileName As String, ByVal exportHeader As Boolean) As Boolean
        Dim result As Boolean = True
        Try
            Dim sb As New System.Text.StringBuilder()
            Dim separator As String = ";"
            Dim group As String = """"
            Dim newLine As String = Environment.NewLine

            If exportHeader Then
                For Each column As DataColumn In dtable.Columns
                    'Se for a última coluna não coloca o ; no final.
                    If column.ColumnName.ToUpper() = dtable.Columns(dtable.Columns.Count - 1).ColumnName().ToUpper() Then
                        sb.Append(wrapValue(column.ColumnName, group, separator))
                    Else
                        sb.Append(wrapValue(column.ColumnName, group, separator) & separator)
                    End If
                Next
                sb.Append(newLine)
            End If
            For Each row As DataRow In dtable.Rows
                For Each col As DataColumn In dtable.Columns
                    'Se for a última coluna não coloca o ; no final.
                    If col.ColumnName.ToUpper() = dtable.Columns(dtable.Columns.Count - 1).ColumnName().ToUpper() Then
                        sb.Append(wrapValue(row(col).ToString(), group, separator))
                    Else
                        sb.Append(wrapValue(row(col).ToString(), group, separator) & separator)
                    End If
                Next

                sb.Append(newLine)
            Next
            Using fs As New StreamWriter(fileName)
                fs.Write(sb.ToString())
                fs.Close()
                fs.Dispose()
            End Using

        Catch ex As Exception
            'TrataException(ex, True, System.Reflection.MethodBase.GetCurrentMethod(), "FatGerArq")
            Throw ex
            result = False
        End Try
        Return result
    End Function
    Public Shared Function ExportToCSVRelatorios(ByVal dtable As DataTable, ByVal fileName As String, ByVal exportHeader As Boolean) As Boolean
        Dim result As Boolean = True
        Try
            Dim sb As New System.Text.StringBuilder()
            Dim separator As String = ";"
            Dim group As String = """"
            Dim newLine As String = Environment.NewLine


            If exportHeader Then
                For Each column As DataColumn In dtable.Columns
                    'Se for a última coluna não coloca o ; no final.
                    If column.ColumnName.ToUpper() = dtable.Columns(dtable.Columns.Count - 1).ColumnName().ToUpper() Then
                        sb.Append(wrapValue(column.ColumnName, group, separator))
                    Else
                        sb.Append(wrapValue(column.ColumnName, group, separator) & separator)
                    End If
                Next

                sb.Append(newLine)
            End If
            For Each row As DataRow In dtable.Rows
                For Each col As DataColumn In dtable.Columns
                    'Se for a última coluna não coloca o ; no final.
                    If col.ColumnName.ToUpper() = dtable.Columns(dtable.Columns.Count - 1).ColumnName().ToUpper() Then
                        sb.Append(wrapValue((row(col).ToString()), group, separator))
                    Else
                        sb.Append(wrapValue(row(col).ToString(), group, separator) & separator)
                    End If
                Next

                sb.Append(newLine)
            Next
            Using fs As New StreamWriter(fileName)
                fs.Write(sb.ToString())
            End Using
        Catch ex As Exception
            'TrataException(ex, True, System.Reflection.MethodBase.GetCurrentMethod(), "FatGerArq")
            Throw ex
            result = False
        End Try
        Return result
    End Function

    Public Shared Function GetHTML(ByVal link As String, ByVal useProxy As Boolean) As String

        Dim strResultado As String = ""

        Dim webResponse As System.Net.WebResponse

        Dim webRequest As System.Net.WebRequest = System.Net.HttpWebRequest.Create(link)

        webRequest.ContentType = "text/xml; charset=utf-8"

        If useProxy Then
            Dim Proxy As New WebProxy("192.168.1.10:3128")
            Proxy.Credentials = New NetworkCredential("teleatlantic-sp", "tele@tele12", "teleatlantic")
            webRequest.Proxy = Proxy
        End If
        webResponse = webRequest.GetResponse()


        Dim strReader As New System.IO.StreamReader(webResponse.GetResponseStream())
        strResultado = WebUtility.HtmlDecode(strReader.ReadToEnd())
        strReader.Close()

        Return strResultado
    End Function

    Public Shared Function getExceptionRentSoft(ByVal html As String) As String
        If html.ToUpper().Contains("BOLETO NAO ENCONTRADO") Then Return "Boleto não encontrado para envio do e-mail."
        If html.ToUpper().Contains("EXISTE MAIS DE UM BOLETO COM OS PARAMETROS INFORMADOS. MENSAGEM NAO ENVIADA.") Then Return "Existe mais de um boleto com estes dados para envio."
        If html.ToUpper().Contains("BOLETO CANCELADO. MENSAGEM NÃO ENVIADA.") Then Return "Boleto cancelado."
        If html.ToUpper().Contains("BOLETO VENCIDO. MENSAGEM NÃO ENVIADA.") Then Return "Boleto vencido."
        If html.ToUpper().Contains("EM PROCESSO DE ENVIO DE BOLETO. MENSAGEM NAO ENVIADA.") Then Return "Em processo de envio de boleto. Mensagem não enviada."
        If html.ToUpper().Contains("BOLETO EM PROCESSO DE ENVIO. TENTE NOVAMENTE DENTRO DE INSTANTES.") Then Return "Boleto em processo de envio. Tente novamente dentro de instantes."

        Return ""

    End Function

    Public Shared Sub UploadToFTPServer(ByVal FileName As String,
                                        ByVal UploadPath As String,
                                        ByVal FTPUser As String,
                                        ByVal FTPPass As String)

        Try

            'Traz os dados do arquivo
            Dim FileInfo As New System.IO.FileInfo(FileName)

            'Cria o FtpWebRequest do endereço recebido
            Dim FTPWebRequest As System.Net.FtpWebRequest = CType(System.Net.FtpWebRequest.Create(New Uri(UploadPath)), System.Net.FtpWebRequest)

            'Relaciona usuário e senha
            FTPWebRequest.Credentials = New System.Net.NetworkCredential(FTPUser, FTPPass)

            'Por default KeepAlive é true que mantém a conexão aberta
            'depois da execução de um comando
            FTPWebRequest.KeepAlive = False

            FTPWebRequest.UsePassive = False

            'Define o timeout para 30 segundos
            FTPWebRequest.Timeout = 30000

            'Específica o comando a ser executado
            FTPWebRequest.Method = System.Net.WebRequestMethods.Ftp.AppendFile

            'Específica o tipo de dado que será transferido
            FTPWebRequest.UseBinary = True

            'Notifica o servidor do tamanho do arquivo
            FTPWebRequest.ContentLength = FileInfo.Length

            'O tamanho do buffer é setado para 2kb
            Dim buffLength As Integer = 2048
            Dim buff(buffLength - 1) As Byte

            'Abre um file stream para ler o arquivo
            Dim FileStream As System.IO.FileStream = FileInfo.OpenRead()

            'Stream em qual o arquivo está escrito
            Dim Stream As System.IO.Stream = FTPWebRequest.GetRequestStream()

            'Lê do arquivo 2kb por vez
            Dim contentLen As Integer = FileStream.Read(buff, 0, buffLength)

            'Até o conteúdo acabar
            Do While contentLen <> 0
                'Escreve o conteúdo do arquivo no stream de upload
                Stream.Write(buff, 0, contentLen)
                contentLen = FileStream.Read(buff, 0, buffLength)
            Loop

            'Fecha o stream do arquivo e de requisição
            Stream.Close()
            Stream.Dispose()
            FileStream.Close()
            FileStream.Dispose()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub FileUploadSFTP(ByVal FileName As String,
                                        ByVal UploadPath As String,
                                        ByVal SFTPUser As String,
                                        ByVal SFTPPass As String,
                                        ByVal SFTPPort As String,
                                        ByVal ChangeDirectory As String)

        'Arthur 01/07/2021
        'Dim novoRegistro As String = ""
        'Using client As SftpClient = New SftpClient(UploadPath, SFTPPort, SFTPUser, SFTPPass)
        '    'Pega as informações do arquivo
        '    Dim FileInfo As New System.IO.FileInfo(FileName)
        '    'Conecta no SFTP
        '    client.Connect()
        '    'Muda de pasta após conectar
        '    If ChangeDirectory.ToString().ToUpper = "TRUE" Then
        '        client.ChangeDirectory("srentsoft")
        '    End If

        '    'Verifica se o arquivo já existe no SFTP
        '    If client.Exists(FileInfo.Name) Then
        '        'Le o conteudo do novo registro
        '        Dim objReader As New System.IO.StreamReader(FileName)
        '        Do While objReader.Peek() <> -1
        '            novoRegistro = novoRegistro + objReader.ReadLine() + Environment.NewLine
        '        Loop
        '        Dim txtNovoReg = novoRegistro
        '        'Escreve no arquivo existente no SFTP
        '        client.AppendAllText("/srentsoft/" + FileInfo.Name, txtNovoReg)
        '        'Garante fechar conexão com SFTP 
        '        client.Disconnect()
        '        client.Dispose()
        '        'Fecha e limpa o obj que leu arquivo 
        '        objReader.Close()
        '        objReader.Dispose()

        '    Else
        '        'Caso não exista, cria o arquivo;
        '        Using fs As FileStream = New FileStream(FileName, FileMode.Open)
        '            client.BufferSize = 4 * 1024
        '            client.UploadFile(fs, Path.GetFileName(FileName))
        '        End Using
        '        client.Disconnect()
        '        client.Dispose()

        '    End If
        'End Using
    End Sub

    Public Shared Function validarDataFutura(ByVal dtData As DateTime) As Boolean
        Dim dtHoje As DateTime = Funcoes.PegaData

        'Se o dia/mes/ano fore anterior a hoje já retorna false
        If dtData.Date < dtHoje.Date Then
            Return False
        End If

        'Se for pro mesmo dia, válida horários
        If dtData.Date = dtHoje.Date Then
            'Se for a hora passada, retorna false
            If dtData.Hour < dtHoje.Hour Then
                Return False
            End If

            'Se for a mesma hora, mas minuto passado, retorna false
            If dtData.Hour = dtHoje.Hour And dtData.Minute < dtHoje.Minute Then
                Return False
            End If
        End If

        Return True
    End Function

    Public Shared Function getSenhaRandomica() As String
        Dim rdm As New Random()
        Dim allowChrs() As Char = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLOMNOPQRSTUVWXYZ0123456789".ToCharArray()
        Dim sResult As String = ""

        For i As Integer = 0 To 6 - 1
            sResult += allowChrs(rdm.Next(0, allowChrs.Length))
        Next

        Return sResult
    End Function

    Public Shared Function getNomeArquivoRentsoft(ByVal dtEnvio As DateTime, ByVal hrLimiteEnvio As String) As String
        Dim tmsLimiteEnvio As New TimeSpan(Strings.Left(hrLimiteEnvio, 2), Strings.Right(hrLimiteEnvio, 2), 0)

        'Se for sábado joga o envio para dois dias depois (segunda)
        If dtEnvio.DayOfWeek = DayOfWeek.Saturday Then
            Return Format(DateAdd(DateInterval.Minute, 10, DateAdd(DateInterval.Day, 2, dtEnvio)), "yyyyMMdd0800")
        End If

        'Se for domingo joga o envio para um dia depois (segunda)
        If dtEnvio.DayOfWeek = DayOfWeek.Sunday Then
            Return Format(DateAdd(DateInterval.Minute, 10, DateAdd(DateInterval.Day, 1, dtEnvio)), "yyyyMMdd0800")
        End If

        'Se estiver dentro do horário, envia no próprio dia
        If dtEnvio.Hour < tmsLimiteEnvio.Hours Then
            Return Format(DateAdd(DateInterval.Minute, 10, dtEnvio), "yyyyMMddHHmm")
            'Se for o mesmo horário do limite, mas estiver antes do minuto final, envia no próprio dia.
        ElseIf dtEnvio.Hour = tmsLimiteEnvio.Hours Then
            If dtEnvio.Minute < tmsLimiteEnvio.Minutes Then
                Return Format(DateAdd(DateInterval.Minute, 10, dtEnvio), "yyyyMMddHHmm")
            Else
                'Se já tiver excedido o horário, envia no dia seguinte.
                Return Format(DateAdd(DateInterval.Minute, 10, DateAdd(DateInterval.Day, 1, dtEnvio)), "yyyyMMdd0800")
            End If
        Else
            Return Format(DateAdd(DateInterval.Minute, 10, DateAdd(DateInterval.Day, 1, dtEnvio)), "yyyyMMdd0800")
        End If

    End Function

    Public Shared Function validarCaminhoArquivo(caminho As String) As Boolean
        Dim X As Integer
        Dim Y As Integer = 1

        Do While Y > 0
            X = InStr(Y + 1, caminho, "\", vbTextCompare)
            If X = 0 Then Exit Do
            Y = X
        Loop

        If Len(Dir(Mid(caminho, 1, Y))) = 0 Then
            Return False
        End If
        Return True
    End Function

    'Public Shared Sub ValidaFormsAtivosRadRibbonForm(form As RadRibbonForm)
    '    If TypeOf form Is RadRibbonForm Then
    '        form.Activate()
    '    End If
    'End Sub

    'Public Shared Sub ValidaFormsAtivosRadForm(form As RadForm)
    '    If TypeOf form Is RadForm Then
    '        form.Activate()
    '    End If
    'End Sub

    Public Shared Function GetRetFun(ByVal num As String, ByVal msg As String) As Retorno
        Return New Retorno With {
            .Sucesso = False,
            .NumErro = "ErroFuncional",
            .MsgErro = msg.Replace("'", "").Replace("""", ""),
            .TipoErro = DadosGenericos.TipoErro.Funcional
        }
    End Function

    Public Shared Function getDescricaoTipoPgto(codTipoPgto As String) As String
        Select Case codTipoPgto.ToUpper()
            Case "CC", "CA"
                Return "Cartão de crédito"
            Case "TR"
                Return "Transferência"
            Case "HB"
                Return "Híbrido"
            Case "BO"
                Return "Boleto"
            Case "DC"
                Return "Débito em conta"
            Case "CD"
                Return "Cartão de débito"
            Case "CH"
                Return "Cheque"
            Case "DI"
                Return "Dinheiro"
            Case "DF"
                Return "Desconto em folha"
            Case "CN"
                Return "Carnê"
            Case "DB"
                Return "Depósito bancário"
            Case Else
                Return "Desconhecido"
        End Select
    End Function
    Public Shared Function ValidaCampoNumerico(ByVal charDigito As Char) As Boolean

        'Segundo a tabela ASCII números são os códigos 48 ao 57.

        If (Microsoft.VisualBasic.Asc(charDigito) > 47 And Microsoft.VisualBasic.Asc(charDigito) < 58) Or Microsoft.VisualBasic.Asc(charDigito) = 8 Then
            Return True
        End If

        Return False

    End Function
    Public Shared Function ConvertToDataTable(Of T)(ByVal list As IList(Of T), tableName As String) As DataTable
        Dim table As New DataTable(tableName)
        Dim fields() As PropertyInfo = GetType(T).GetProperties()
        For Each field As PropertyInfo In fields
            If Nullable.GetUnderlyingType(field.PropertyType) Is Nothing Then
                table.Columns.Add(field.Name, field.PropertyType)
            Else
                table.Columns.Add(field.Name, Nullable.GetUnderlyingType(field.PropertyType))
            End If

        Next
        For Each item As T In list
            Dim row As DataRow = table.NewRow()
            For Each field As PropertyInfo In fields
                row(field.Name) = field.GetValue(item, Nothing)
            Next
            table.Rows.Add(row)
        Next
        Return table
    End Function

    Public Shared Function gerarDVNossoNumeroSantander(nossoNumero As String) As String
        Dim DV As String = ""
        Dim soma As Integer = 0
        Dim multiplicador As Integer = 2

        'Para fazer o cálculo do DV é preciso reverter a string
        nossoNumero = StrReverse(nossoNumero)

        'Loop por cada caractere na string
        For Each c As Char In nossoNumero
            'Se tiver chego no máximo do multiplicador (9), volta para 2
            If multiplicador >= 10 Then multiplicador = 2

            'Efetua a multiplicação do número atual
            soma += Integer.Parse(c) * multiplicador

            'Acrescenta +1 no multiplicador
            multiplicador += 1
        Next

        'Resultado da soma é dividido por 11
        Dim divisao As Integer = Math.Floor(soma / 11)

        'Calcula o "resto" da divisão
        Dim resto As Integer = (divisao * 11) - soma

        'Se o "resto" for negativo, transforma para positivo
        If resto < 0 Then resto = resto * (-1)

        'Se o resto for igual a 10 o dígito deve ser 1
        If resto = 10 Then DV = 1

        'Se o resto for igual a 1 ou 0 o dígito é 0
        If resto = 0 Or resto = 1 Then DV = 0

        'Se o resto for diferente 1, 0 ou 10 o resto deve ser subtraído por 11
        If resto <> 0 And resto <> 1 And resto <> 10 Then DV = 11 - resto

        Return DV
    End Function

    Public Shared Function gerarDACNossoNumeroItau(agencia As String, conta As String, carteira As String, nossoNumero As String) As String
        Dim DAC As Integer = 0
        Dim Total As Integer = 0
        Dim seqCalculo As Integer = 1
        'Monta string com a junção dos campos para loop e geração do DAC
        Dim Numero As String = agencia & conta.Substring(0, 5) & carteira & nossoNumero

        Dim resultado As Integer = 0

        'Loop para gerar DAC com base em cada caractere da string montada acima
        For Each c As Char In Numero

            'Fal o cálculo para pegar resultado
            resultado = Integer.Parse(c) * seqCalculo

            'Se o resultado for maior do que dois dígitos, realiza novo calculo
            If resultado >= 10 Then resultado = Integer.Parse(resultado.ToString().Substring(0, 1)) + Integer.Parse(resultado.ToString().Substring(1, 1))

            'Calcula o total de acordo com lógica anterior
            Total = Total + resultado

            'Atualiza indice para sequencia de calculo
            If seqCalculo.Equals(1) Then seqCalculo = 2 Else seqCalculo = 1

        Next

        'Divide o total por 10 e coleta o restante da divisão
        DAC = Total Mod 10

        'Caso o resto da divsão seja 0, considera o DAC como 0
        If DAC.Equals(0) Then DAC = 0 Else DAC = 10 - DAC

        Return DAC
    End Function
    Public Sub TrataErro(ByVal exibir As Boolean,
                          ByRef lstRetorno As List(Of Retorno),
                          Optional ByVal conn As SqlConnection = Nothing,
                          Optional ByVal trans As SqlTransaction = Nothing)

        'Controle de transação
        If Not IsNothing(trans) Then trans.Rollback()
        If Not IsNothing(conn) Then conn.Close()

        'Exibe grid de avisos
        'If exibir Then ExibirErro(lstRetorno)

    End Sub

    Public Shared Function isErroLoginAD(msgErro As String) As String
        msgErro = msgErro.ToUpper()
        If msgErro.Contains("NOME DE USUÁRIO OU SENHA INCORRETOS.") Then Return "Usuário ou senha incorretos."

        Return String.Empty
    End Function
    Public Shared Function ValidaTelefoneInvalido(telefone As String) As Boolean
        If telefone.Count < 8 Then
            Return False
        End If

        If telefone.Count = 9 Then

            Select Case telefone
                Case "000000000"
                    Return False
                Case "111111111"
                    Return False
                Case "222222222"
                    Return False
                Case "333333333"
                    Return False
                Case "444444444"
                    Return False
                Case "555555555"
                    Return False
                Case "666666666"
                    Return False
                Case "777777777"
                    Return False
                Case "888888888"
                    Return False
                Case "999999999"
                    Return False
            End Select

            Return True
        ElseIf telefone.Count = 8 Then
            Select Case telefone
                Case "00000000"
                    Return False
                Case "11111111"
                    Return False
                Case "22222222"
                    Return False
                Case "33333333"
                    Return False
                Case "44444444"
                    Return False
                Case "55555555"
                    Return False
                Case "66666666"
                    Return False
                Case "77777777"
                    Return False
                Case "88888888"
                    Return False
                Case "99999999"
                    Return False
            End Select
        End If
        Return True
    End Function

    Public Shared Function validaCPF(ByVal strCPFCliente As String) As Boolean

        Dim strCPFOriginal As String = strCPFCliente.Replace(".", "").Replace("-", "")
        Dim strCPF As String = Mid(strCPFOriginal, 1, 9)
        Dim strCPFTemp As String
        Dim intSoma As Integer
        Dim intResto As Integer
        Dim strDigito As String
        Dim intMultiplicador As Integer = 10
        Const constIntMultiplicador As Integer = 11
        Dim i As Integer

        Select Case strCPFCliente
            Case "00000000000"
                Return False
            Case "11111111111"
                Return False
            Case "22222222222"
                Return False
            Case "33333333333"
                Return False
            Case "44444444444"
                Return False
            Case "55555555555"
                Return False
            Case "66666666666"
                Return False
            Case "77777777777"
                Return False
            Case "88888888888"
                Return False
            Case "99999999999"
                Return False
        End Select

        For i = 0 To strCPF.ToString.Length - 1
            intSoma += CInt(strCPF.ToString.Chars(i).ToString) * intMultiplicador
            intMultiplicador -= 1
        Next

        If (intSoma Mod constIntMultiplicador) < 2 Then
            intResto = 0
        Else
            intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
        End If

        strDigito = intResto
        intSoma = 0

        strCPFTemp = strCPF & strDigito
        intMultiplicador = 11

        For i = 0 To strCPFTemp.Length - 1
            intSoma += CInt(strCPFTemp.Chars(i).ToString) * intMultiplicador
            intMultiplicador -= 1
        Next

        If (intSoma Mod constIntMultiplicador) < 2 Then
            intResto = 0
        Else
            intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
        End If

        strDigito &= intResto

        If strDigito = Mid(strCPFOriginal, 10, strCPFOriginal.Length) Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Shared Function validaCNPJ(ByVal CNPJ As String) As Boolean

        Dim i As Integer
        Dim valida As Boolean
        Dim dadosArray() As String = {"11111111111111", "22222222222222", "33333333333333", "44444444444444",
                                              "55555555555555", "66666666666666", "77777777777777", "88888888888888", "99999999999999"}

        CNPJ = CNPJ.Trim

        For i = 0 To dadosArray.Length - 1
            If CNPJ.Length <> 14 Or dadosArray(i).Equals(CNPJ) Then
                Return False
            End If
        Next

        valida = efetivaValidacao(CNPJ)

        If valida Then
            validaCNPJ = True
        Else
            validaCNPJ = False
        End If

    End Function


    Public Shared Function efetivaValidacao(ByVal cnpj As String)

        Dim Numero(13) As Integer
        Dim soma As Integer
        Dim i As Integer
        Dim resultado1 As Integer
        Dim resultado2 As Integer

        For i = 0 To Numero.Length - 1
            Numero(i) = CInt(cnpj.Substring(i, 1))
        Next

        soma = Numero(0) * 5 + Numero(1) * 4 + Numero(2) * 3 + Numero(3) * 2 + Numero(4) * 9 + Numero(5) * 8 + Numero(6) * 7 +
                   Numero(7) * 6 + Numero(8) * 5 + Numero(9) * 4 + Numero(10) * 3 + Numero(11) * 2

        soma = soma - (11 * (Int(soma / 11)))

        If soma = 0 Or soma = 1 Then
            resultado1 = 0
        Else
            resultado1 = 11 - soma
        End If

        If resultado1 = Numero(12) Then

            soma = Numero(0) * 6 + Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 +
                         Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2

            soma = soma - (11 * (Int(soma / 11)))

            If soma = 0 Or soma = 1 Then
                resultado2 = 0
            Else
                resultado2 = 11 - soma
            End If

            If resultado2 = Numero(13) Then
                Return True
            Else
                Return False
            End If

        Else
            Return False
        End If

    End Function

    ''' <summary>
    ''' Retorna um datatable, cujas linhas em branco foram eliminadas
    ''' </summary>
    ''' <param name="DTable">DataTable a ser manipulado</param>
    ''' <returns>Retorna um DataTable modificado</returns>
    Public Function EliminarLinhasVaziasDT(ByVal DTable As DataTable) As DataTable
        Dim valuesarr As New StringBuilder
        For i As Integer = DTable.Rows.Count - 1 To 0 Step -1
            Dim lst As New List(Of Object)(DTable.Rows(i).ItemArray)
            For Each s As Object In lst
                valuesarr.Append(s.ToString)
            Next
            If String.IsNullOrEmpty(valuesarr.ToString) Then
                DTable.Rows.RemoveAt(i)
            End If
        Next
        Return DTable
    End Function

#Region "****Metodo para ocultar valores de uma string, transforma uma string em asteriscos"
    Public Shared Function TransformeString(StrEnter As String) As String
        Dim Result As String = ""

        If (IsNothing(StrEnter) Or StrEnter = "") Then
            Return Result
        End If

        StrEnter = StrEnter.Replace("-", "").Trim

        For i = 0 To StrEnter.Length - 1
            Result = Result + "*"
        Next

        Return Result
    End Function
#End Region

#Region "****Metodo para validação de numero de cartão de credito****"
    Public Shared Function ValidateCreditCard(ByVal CardNumber As String) As String
        Dim CheckSum As Integer = 0
        Dim CharPos As Integer
        Dim Digit As String

        Dim CardName As String = ""
        Dim CardValidate As String

        'Dim tChar As String

        For CharPos = Len(CardNumber) To 2 Step -2
            CheckSum = CheckSum + CInt(Mid(CardNumber, CharPos, 1))
            Digit = CStr((Mid(CardNumber, CharPos - 1, 1)) * 2)
            CheckSum = CheckSum + CInt(Left(Digit, 1))

            If Len(Digit) > 1 Then CheckSum = CheckSum + CInt(Right(Digit, 1))
        Next

        If Len(CardNumber) Mod 2 = 1 Then CheckSum = CheckSum + CInt(Left(CardNumber, 1))

        If (CheckSum Mod 10 = 0) Then
            CardValidate = Strings.Left(CardNumber, 6)
            ' Elo -- 636368,438935,504175,451416,636297 -- 15 length
            If (CardValidate = "636368" Or CardValidate = "438935" Or CardValidate = "504175" Or CardValidate = "451416" Or CardValidate = "636297") Then
                Return "Elo"
            End If

            CardValidate = Strings.Left(CardNumber, 5)

            ' Elo -- 5067,4576,4011  -- 15 length
            If (CardValidate = "36297") Then
                Return "Elo"
            End If

            CardValidate = Strings.Left(CardNumber, 4)

            ' enRoute -- 2014,2149 -- 15 length
            If (CardValidate = "2014" Or CardValidate = "2149") Then
                Return "Diners"
            End If

            '' JCB -- 2131, 1800 -- 15 length
            'If (CardValidate = "2131" Or CardValidate = "1800") Then
            '    Return "JCB"
            'End If

            '' Discover -- 6011 -- 16 length
            'If (CardValidate = "6011") Then
            '    Return "Discover"
            'End If

            ' Elo -- 5067,4576,4011  -- 15 length
            If (CardValidate = "5067" Or CardValidate = "4576" Or CardValidate = "4011") Then
                Return "Elo"
            End If

            CardValidate = Strings.Left(CardNumber, 3)

            ' Elo -- 509,650,651,655 15 length
            If (CardValidate = "509" Or CardValidate = "650" Or CardValidate = "651" Or CardValidate = "655") Then
                Return "Elo"
            End If

            ' Diners Club -- 36 or 38 -- 14 length
            If (CardValidate = "300" Or CardValidate = "301" Or CardValidate = "302" Or CardValidate = "303" Or CardValidate = "304" Or CardValidate = "305") Then
                Return "Diners"
            End If

            CardValidate = Strings.Left(CardNumber, 2)

            ' AMEX -- 34 or 37 -- 15 length
            If (CardValidate = "34" Or CardValidate = "37") Then
                Return "Amex"
            End If

            ' MasterCard -- 51 through 55 -- 16 length
            If (CardValidate = "51" Or CardValidate = "52" Or CardValidate = "53" Or CardValidate = "54" Or CardValidate = "55") Then
                Return "MasterCard"
            End If

            ' Diners Club -- 36 or 38 -- 14 length
            If (CardValidate = "36" Or CardValidate = "38") Then
                Return "Diners"
            End If

            CardValidate = Strings.Left(CardNumber, 1)

            ' VISA -- 4 -- 13 and 16 length
            If (CardValidate = "4") Then
                Return "Visa"
            End If

            '' JCB -- 3 -- 16 length
            'If (CardValidate = "4") Then
            '    Return "JCB"
            'End If
        End If

        Return CardName
    End Function

    'Public Shared Function ValidateCreditCard(ByVal CardNumber As String) As Boolean
    '    Dim CheckSum As Integer = 0
    '    Dim DoubleFlag As Boolean = (CardNumber.Length Mod 2 = 0)
    '    Dim Digit As Char
    '    Dim DigitValue As Integer

    '    For Each Digit In CardNumber
    '        DigitValue = Integer.Parse(Digit)

    '        If (DoubleFlag) Then
    '            DigitValue *= 2

    '            If (DigitValue > 9) Then
    '                DigitValue -= 9
    '            End If
    '        End If

    '        CheckSum += DigitValue
    '        DoubleFlag = Not DoubleFlag
    '    Next

    '    Return (CheckSum Mod 10 = 0)
    'End Function
#End Region

End Class

