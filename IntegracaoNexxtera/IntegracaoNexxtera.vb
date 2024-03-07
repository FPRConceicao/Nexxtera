Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports BoletoNet
Imports Telerik.WinControls


Module IntegracaoNexxtera

#Region "Declaração de Lista"
    Dim lstRetorno As New List(Of Retorno)
    Dim lstBanco As New List(Of Banco)
    Dim lstControleRetBco As New List(Of ControleRetBco)
    Dim lstRetornoBco As New List(Of RetornoBco)
    Dim lstBxAutoErros As New List(Of BxAutosErros)
#End Region

#Region "Variaveis Publicas"
    Public VG_bDescTit As Boolean = False
#End Region

#Region "Declaração de Variavel"
    Dim strTipoCta As String = ""
    Dim strNomeBanco As String = ""
    Dim strAgencia As String = ""
    Dim strConta As String = ""
    Dim blnRetorno As Boolean = False
    Dim strAgen As String = ""
    Dim strContaCorrente As String = ""
    Dim fsFile As System.IO.StreamReader
    Dim isErroExibido As Boolean = False
    Private strContaCorrenteItau As String = ""
    Private strAgenItau As String = ""
#End Region

#Region "Declaração de Metodos/Eventos"
    Dim clConsultaBancoDebtAutomatico As New BuscaBancoDebitoAutomatico
    Dim clBuscaBancoContaCorrente As New BuscaBancoContaCorrente
    Dim clConsultaControleRetBco As New ConsultaControleRetBco
    Dim clInserirControleRetBco As New InserirControleRetBco
    Dim clConsultaRetornoBco As New ConsultaRetornoBco
    'Dim clAlterarContaCorrente As New AlterarContaCorrente
    'Dim clConsultaParametro As New ConsultaParametroGeral
    'Dim clApagarTituloEnvBco As New ApagarTitulosEnvBco
    Dim clAlterarContaReceber As New ContasAReceberAlterar
    'Dim clConsultarContaCorrente As New ConsultarContaCorrente
    Dim clInserirRetornoBco As New InserirRetornoBco
    Dim clAlterarBaixaContasReceber As New AlterarBaixaContasReceber
    Dim clConsultaBxAutoErros As New ConsultaBxAutosErros
    Dim clConsultaCtasAReceber As New ConsultaContasReceber
    Dim consulta As New ConsultaContasReceber()
    Dim contaReceber As New ContaReceber()
    'Dim clInserirTitulosEnvBco As New InserirTitulosEnvBco
    Dim clConsultaContaReceber As New ConsultaContasReceber
    Dim clInserirTempCheckArqRetorno As New InserirTempCheckArqRetorno
    Dim clApagarTempCheckArqRetorno As New ApagarTempCheckArqRetorno
    Dim clConsultacontasReceber As New ConsultaContasReceber
    Dim clAlterarContasReceber As New ContasAReceberAlterar
    Dim clInserirHistoricoContato As New InserirHistoricoContato
    Dim lstContaReceber As New List(Of ContaReceber)

#End Region

    Sub Main()
        Try


            Dim diretorio As String
            Dim listaArquivos As New List(Of String)
            Dim mArquivo As String
            Dim sBuffer As String
            Dim tamanhoArquivo As Int32
            'Dim iCon As Integer
            'Dim fsFile As IO.StreamReader
            'diretorio = "C:\Documentos\Documentos\Analises\Nexxtera\Bancos\Itau\boleto\20240229\" 'Teste local
            'diretorio = "X:\inbox" 'Teste
            diretorio = "C:\Skyline\inbox" 'Produção

            For Each foundFile As String In My.Computer.FileSystem.GetFiles(diretorio)

                listaArquivos.Add(foundFile)

            Next
            For Each nomeArquivo As String In listaArquivos

                Dim arquivo As String = Path.GetFileName(nomeArquivo)
                Dim codigoBanco As String = arquivo.Substring(4, 3)
                Dim tipoArquivo As String = arquivo.Substring(0, 3)
                tamanhoArquivo = Convert.ToInt32(FileLen(nomeArquivo))
                If tipoArquivo = "COB" Or tipoArquivo = "DEB" Then
                    clApagarTempCheckArqRetorno.ApagaTudoTempCheckArqRetornoNexxtera()

                    'CreateFile(nomeArquivo, codigoBanco)
                    Select Case codigoBanco
                        Case "001" 'Banco do Brasil
                            If tamanhoArquivo > 500 Then
                                PBaixaTitulos001(nomeArquivo)
                                Check001(nomeArquivo)
                            End If


                        Case "237"
                            If tamanhoArquivo > 500 Then
                                If InStr(1, nomeArquivo, "\DD") > 0 Then
                                    PBaixaTitulos237DD(nomeArquivo)
                                    Check237(nomeArquivo)
                                Else
                                    PBaixaTitulos237(nomeArquivo)
                                    Check237(nomeArquivo)
                                End If
                            End If


                    '        Case "291"
                    '            PBaixaTitulos291()

                    '        Case "347"
                    '            PBaixaTitulos347()

                    '        Case "224"
                    '            PBaixaTitulos224()


                        Case "341" 'Banco Itau
                            If tamanhoArquivo > 500 Then
                                If InStr(1, nomeArquivo, "\CR") > 0 Then
                                    PBaixaTitulos341DD(nomeArquivo)
                                    Check341(nomeArquivo)
                                Else
                                    PBaixaTitulos341(nomeArquivo)
                                    Check341(nomeArquivo)
                                End If
                            End If
                    '        Case "409"
                    '            PBaixaTitulos409()
                        Case "033" 'Banco Santander
                            If tamanhoArquivo > 500 Then
                                PBaixaTitulos033(nomeArquivo)
                                Check033(nomeArquivo)
                            End If

                        Case "399" 'Banco HSBC
                        'PBaixaTitulos399(nomeArquivo)

                    '        Case "422"
                    '            PBaixaTitulos422()

                        Case "356" 'Banco Real
                        'PBaixaTitulos356(nomeArquivo)

                    '        Case "479"
                    '            PBaixaTitulos479()
                        Case "104" 'Caixa
                            If tamanhoArquivo > 500 Then
                                PBaixaTitulos104(nomeArquivo)
                                Check104(nomeArquivo)
                            End If
                        '05/04/2018 - Fernando
                        'Novo banco Caixa (TeleAlarme)

                        Case "041" 'Banrisul
                            If tamanhoArquivo > 500 Then
                                PBaixaTitulos041(nomeArquivo)
                                Check041(nomeArquivo)
                            End If
                            '19/12/2018
                            'Banco Banrisul

                    End Select

                    If tamanhoArquivo > 500 Then

                        CarregaRelatorioExcel(nomeArquivo)
                        If Not ArquivoEmUso(nomeArquivo) Then
                            My.Computer.FileSystem.MoveFile(nomeArquivo, "C:\Skyline\Processado" & arquivo)
                        End If
                        '
                        Console.Write(" Sucesso Banco: " & arquivo)
                    Else
                        If Not ArquivoEmUso(nomeArquivo) Then
                            My.Computer.FileSystem.MoveFile(nomeArquivo, "C:\Skyline\Processado\" & arquivo)
                        End If
                        Console.Write("Arquivo Vazio")
                        Funcoes.CriaLog(Application.StartupPath & "\RelatorioErros.txt", "Arquivo:" & arquivo & " vazio")
                    End If
                End If

            Next


            Console.Write("Arquivos executados com sucesso")
            Funcoes.CriaLog(Application.StartupPath & "\RelatorioErros.txt", "Arquivos executados com sucesso: " & Now)
        Catch ex As Exception
            Console.Write(ex.Message)
            Funcoes.CriaLog(Application.StartupPath & "\RelatorioErros.txt", ex.Message)
        End Try
    End Sub
    Private Function ArquivoEmUso(nomeArquivo As String) As Boolean
        Try
            Dim fs As System.IO.FileStream = System.IO.File.OpenWrite(nomeArquivo)
            fs.Close()
            Return False
        Catch ex As Exception
            Return True
        End Try
    End Function
    Private Sub BuscaBanco(ByVal strCodBanco As String)
        Dim dtpData As DateTime = Funcoes.PegaData

        lstBanco = clConsultaBancoDebtAutomatico.BuscaBancoContaTipoContaDoBanco(strCodBanco, "")
        'If Not lstBanco(0).Sucesso Then
        '    blnRetorno = True
        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBanco(0).MsgErro, lstBanco(0).NumErro, lstBanco(0).Sucesso, lstBanco(0).TipoErro, lstBanco(0).ImagemErro)
        'Else
        strCodBanco = lstBanco(0).CodBanco
        strNomeBanco = lstBanco(0).NomeBanco
        strTipoCta = lstBanco(0).TipoConta
        strAgencia = lstBanco(0).CodAgen
        strConta = lstBanco(0).Numcta

        '    sfdArquivo.FileName = ""
        'If strCodBanco = "224" Then
        '    txtEndArquivo.Enabled = False
        '    nomeArquivo = "C:\BCOFIBRA\COB" & Funcoes.StrZero(dtpData.Day, 2) & Funcoes.StrZero(dtpData.Month, 2) & ".REM"
        '    txtEndArquivo.Tag = "COB" & Funcoes.StrZero(dtpData.Day, 2) & Funcoes.StrZero(dtpData.Month, 2) & ".REM"
        '    btnLocalizaArq.Enabled = False
        'ElseIf strCodBanco = "104" Then
        '    btnLocalizaArq.Enabled = True
        '    nomeArquivo = RenomeiaArquivoRemessaCaixa201912()
        '    txtEndArquivo.Tag = RenomeiaArquivoRemessaCaixa201912()
        '    txtEndArquivo.Enabled = True
        'Else

        '    btnLocalizaArq.Enabled = True
        '        txtEndArquivo.Enabled = True
        '        nomeArquivo = ""
        '        txtEndArquivo.Select()
        '    End If

        '    'End If

        '    If blnRetorno Then
        '    ExibirErro()
        'End If
    End Sub
    'Private Sub FGeraRemessa237DESC(ByVal Connection As SqlConnection, ByVal Transaction As SqlTransaction, codigoBanco As String, nomeArquivo As String) 'BRADESCO
    '    Dim _Retorno As New Retorno
    '    Dim lstExportacao As New List(Of Exportacao)
    '    Dim _Parametro As New Parametros
    '    Dim _Banco As New Banco

    '    Dim iConREG As Integer
    '    Dim sAgen As String, sConta As String
    '    Dim strREG As String, iSeqRemes As Integer, iTotal As Integer
    '    Dim sDig As String, sCodPremes As String, dJurosDia As Double, dJurosMes As Double, strCodCarteira As String
    '    'Dim dblTaxaDia As Double, iFile As Integer
    '    Dim sRSEmpresa As String, sRSCliente As String, sEnd As String
    '    Dim sCNPJEmpresa As String
    '    Dim bErroSemValor As Boolean, bErroCNPJ As Boolean, bErroCEP As Boolean, sMsgErros As String

    '    If ddlTipo.Text.ToUpper = "DÉBITO AUTOMÁTICO - BOLETO UNIFICADO" Then
    '        _Retorno = clInserirTitulosEnvBco.InsereTitulos_Bu_Dc(codigoBanco, strAgencia, strConta, "DC", Connection, Transaction)
    '    Else
    '        _Retorno = clInserirTitulosEnvBco.InsereTitulos_Env_Bco(codigoBanco, strAgencia, strConta, Connection, Transaction)
    '    End If

    '    If Not _Retorno.Sucesso Then
    '        blnRetorno = True
    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '        Exit Sub
    '    End If

    '    sAgen = ""
    '    sConta = ""

    '    'PESQUISA DADOS PARA EXPORTACAO
    '    If ddlTipo.Text.ToUpper = "DÉBITO AUTOMÁTICO - BOLETO UNIFICADO" Then
    '        lstExportacao = clConsultaContaReceber.BuscaContaReceberGeraRemessaBU_DC(codigoBanco, strAgencia, strConta, IIf(cbHabilitaFiltros.Checked, Trim(mebCliente.Text), ""), GetDataInicio(), GetDataFim(), IIf(cbHabilitaFiltros.Checked, Trim(txtTitIni.Text), ""), IIf(cbHabilitaFiltros.Checked, Trim(txtTitFim.Text), ""))
    '    Else
    '        lstExportacao = clConsultaContaReceber.BuscaContaReceberGeraRemessa(codigoBanco, strAgencia, strConta, IIf(cbHabilitaFiltros.Checked, Trim(mebCliente.Text), ""), GetDataInicio(), GetDataFim(), IIf(cbHabilitaFiltros.Checked, Trim(txtTitIni.Text), ""), IIf(cbHabilitaFiltros.Checked, Trim(txtTitFim.Text), ""))
    '    End If
    '    If Not lstExportacao(0).Sucesso Then
    '        blnRetorno = True
    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstExportacao(0).MsgErro, lstExportacao(0).NumErro, lstExportacao(0).Sucesso, lstExportacao(0).TipoErro, lstExportacao(0).ImagemErro)
    '        Exit Sub
    '    End If

    '    'CONSULTA PARAMETRO
    '    _Parametro = clConsultaParametro.PesquisaParametro()
    '    If Not _Parametro.Sucesso Then
    '        blnRetorno = True
    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstExportacao(0).MsgErro, lstExportacao(0).NumErro, lstExportacao(0).Sucesso, lstExportacao(0).TipoErro, lstExportacao(0).ImagemErro)
    '        Exit Sub
    '    End If

    '    sAgen = Funcoes.StrZero(Strings.Left(strAgencia, 4), 5)
    '    sConta = Funcoes.StrZero(Strings.Left(strCtaVincDesc, Len(strCtaVincDesc) - 2), 7)
    '    sDig = Strings.Right(strCtaVincDesc, 1)

    '    iSeqRemes = 0

    '    'Aqui
    '    'Verificar alteracao de codigo e sequencia de remessa (para conta_corrente) ----------------------------------------------------------
    '    _Banco = clBuscaBancoContaCorrente.BuscaCodigoSeqRemessaBanco(codigoBanco, strAgencia, strConta)
    '    If Not _Banco.Sucesso Then
    '        blnRetorno = True
    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Banco.MsgErro, _Banco.NumErro, _Banco.Sucesso, _Banco.TipoErro, _Banco.ImagemErro)
    '        Exit Sub
    '    End If

    '    iSeqRemes = _Banco.ContaCorrente.SeqRemes
    '    sCodPremes = Funcoes.StrZero(_Banco.CodPRemes, 20)
    '    dJurosDia = _Banco.TaxaDia
    '    dJurosMes = _Banco.TaxaMes
    '    strCodCarteira = _Banco.ContaCorrente.CodCarteira

    '    'Alterar a alimentacao da sequencia de remessa, para a conta corrente
    '    _Retorno = clAlterarContaCorrente.AlterarSeqRemesContaCorrente(codigoBanco, strAgencia, IIf(_Banco.SeqRemesUnico = "S", "", strConta), Connection, Transaction)
    '    If Not _Retorno.Sucesso Then
    '        blnRetorno = True
    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '        Exit Sub
    '    End If
    '    '--------------------------------------------------------------------------------------------------------------------------------

    '    Dim fsFile As New System.IO.StreamWriter(nomeArquivo, False, System.Text.Encoding.Default)

    '    iTotal = 1

    '    sRSEmpresa = IIf(IsNothing(_Parametro.RazaoSocial), "", _Parametro.RazaoSocial)
    '    sCNPJEmpresa = IIf(IsNothing(_Parametro.CGC), "66526591000143", _Parametro.CGC)

    '    strREG = "0"
    '    strREG = strREG & "1"
    '    strREG = strREG & "REMESSA"
    '    strREG = strREG & "01"
    '    strREG = strREG & Funcoes.Padl("COBRANCA", 15)
    '    strREG = strREG & Funcoes.StrZero(sCodPremes, 20)
    '    strREG = strREG & Funcoes.RemoveAcento(Funcoes.Padl(sRSEmpresa, 30))
    '    strREG = strREG & "237"
    '    strREG = strREG & Funcoes.Padl("BRADESCO", 15)
    '    strREG = strREG & Format(Now, "ddMMyy")
    '    strREG = strREG & Space(8)
    '    strREG = strREG & "MX"
    '    strREG = strREG & Funcoes.StrZero(Str(iSeqRemes + 1), 7)
    '    strREG = strREG & Space(249)
    '    strREG = strREG & "DESC"
    '    strREG = strREG & Funcoes.StrZero(sAgen, 4) 'Número de autorização - Agencia
    '    strREG = strREG & sAgen
    '    strREG = strREG & Funcoes.StrZero(sConta, 7)
    '    strREG = strREG & sDig
    '    strREG = strREG & Space(7)
    '    strREG = strREG & "000001"

    '    iConREG = 2

    '    For i As Integer = 0 To lstExportacao.Count - 1

    '        bErroCEP = False
    '        bErroCNPJ = False
    '        bErroSemValor = False
    '        sMsgErros = ""

    '        If Len(lstExportacao(i).Cliente.Endereco.Cep) <> 8 Or IsNothing(lstExportacao(i).Cliente.Endereco.Cep) Then
    '            bErroCEP = True
    '            sMsgErros = "CEP inválido ou ausente !" & vbCr
    '        End If

    '        If Len(lstExportacao(i).ContaReceber.VlrInd) <= 0 Then
    '            bErroSemValor = True
    '            sMsgErros = sMsgErros & "Título com valor menor ou igual a zero !" & vbCr
    '        End If

    '        'As linhas abaixo deverão ser efetivadas tão logo o cadastro do cliente esteja normalizado
    '        'Fabricio - 09/04/2001
    '        '        If FValidaCgcCpf(rTrec.Fields("CodClie")) = False Then
    '        '            bErroCNPJ = True
    '        '            sMsgErros = sMsgErros & "CPF / CNPJ do cliente inválido !"
    '        '        End If

    '        If sMsgErros <> "" Then

    '            'RadMessageBox.Show(LoadMsgError(lstExportacao(i).ContaReceber.NumTit, lstExportacao(i).ContaReceber.SeqTit, vbCr, vbCr, sMsgErros, vbCr, vbCr), Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
    '            'MsgBox("Existem as seguintes inconsistências no título " & lstExportacao(i).ContaReceber.NumTit & "-" & lstExportacao(i).ContaReceber.SeqTit & ":" &
    '            'vbCr & vbCr & sMsgErros & vbCr & vbCr & "O título não será enviado ao banco.", vbExclamation, Me.Text)

    '            'Exclui o titulo da tabela de titulos enviados ao Banco
    '            _Retorno = clApagarTituloEnvBco.ApagarTitulosEnvBco(lstExportacao(i).ContaReceber.NumTit, lstExportacao(i).ContaReceber.SeqTit, lstExportacao(i).ContaReceber.DtEmissao, Connection, Transaction)
    '            If Not _Retorno.Sucesso Then
    '                blnRetorno = True
    '                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                Exit Sub
    '            End If

    '        Else

    '            iTotal = iTotal + 1

    '            sRSCliente = IIf(IsNothing(lstExportacao(i).Cliente.RazaoSocial), "", lstExportacao(i).Cliente.RazaoSocial)
    '            sEnd = IIf(IsNothing(lstExportacao(i).Cliente.Endereco.Endereco), "", lstExportacao(i).Cliente.Endereco.Endereco)

    '            strREG = "1"
    '            strREG = strREG & "02"
    '            strREG = strREG & sCNPJEmpresa
    '            strREG = strREG & "0000"
    '            strREG = strREG & Funcoes.StrZero(strCodCarteira, 3) & sAgen & Funcoes.StrZero(sConta, 7) & sDig 'carteira -Agencia - C.corrente - digito
    '            strREG = strREG & Space(25)
    '            strREG = strREG & Funcoes.StrZero("0", 8)
    '            strREG = strREG & Funcoes.StrZero("0", 12)
    '            strREG = strREG & Funcoes.StrZero("0", 10)
    '            strREG = strREG & Space(16)
    '            strREG = strREG & "01"     '109-110
    '            strREG = strREG & Space(10 - (Len(lstExportacao(i).ContaReceber.NumTit) + lstExportacao(i).ContaReceber.SeqTit)) & Trim(lstExportacao(i).ContaReceber.NumTit + lstExportacao(i).ContaReceber.SeqTit)  'Padl(rTrec.Fields("NumTit") + rTrec.Fields("SeqTit"), 10)
    '            strREG = strREG & Format(lstExportacao(i).ContaReceber.DtVcto, "ddMMyy")
    '            strREG = strREG & Funcoes.Padl(Funcoes.StrZero(Format(lstExportacao(i).ContaReceber.VlrInd * 100, "############0"), 13), 13)
    '            strREG = strREG & Space(8)
    '            strREG = strREG & "01"
    '            strREG = strREG & Space(1)
    '            strREG = strREG & Funcoes.StrZero("0", 68)
    '            strREG = strREG & IIf(Len(Trim(lstExportacao(i).ContaReceber.CodClie)) = 14, "02", "01")
    '            strREG = strREG & Funcoes.Padl(Funcoes.FSepara(lstExportacao(i).ContaReceber.CodClie), 14)
    '            strREG = strREG & Funcoes.RemoveAcento(Funcoes.Padl(sRSCliente.Trim(), 40))
    '            strREG = strREG & Funcoes.RemoveAcento(Funcoes.Padl(sEnd, 40))
    '            strREG = strREG & Space(12)
    '            strREG = strREG & Strings.Left(Funcoes.Padl(IIf(IsNothing(lstExportacao(i).Cliente.Endereco.Cep), "", lstExportacao(i).Cliente.Endereco.Cep), 8), 5)
    '            strREG = strREG & Strings.Right(Funcoes.Padl(IIf(IsNothing(lstExportacao(i).Cliente.Endereco.Cep), "", lstExportacao(i).Cliente.Endereco.Cep), 8), 3)
    '            strREG = strREG & Space(60)
    '            strREG = strREG & Funcoes.StrZero(Str(iConREG), 6)

    '            iConREG = iConREG + 1

    '            'ALTERA O STATUSBCO DE CONTAS A RECEBER
    '            _Retorno = clAlterarContaReceber.AlterarStatusBcoContaReceber(lstExportacao(i).ContaReceber.NumTit, lstExportacao(i).ContaReceber.SeqTit, lstExportacao(i).ContaReceber.DtEmissao, Connection, Transaction)
    '            If Not _Retorno.Sucesso Then
    '                blnRetorno = True
    '                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '            End If

    '        End If

    '    Next

    '    iTotal = iTotal + 1

    '    strREG = "9"
    '    strREG = strREG & Space(393)
    '    strREG = strREG & Funcoes.StrZero(Str(iConREG), 6)

    '    fsFile.WriteLine(strREG)
    '    fsFile.Close()

    '    _Retorno = New Retorno
    '    Dim clInsereLog As New InserirContasReceber
    '    _Retorno = clInsereLog.InsereLogEnvioArquivoRemessaBanco(False, codigoBanco, sConta, sAgen, ddlTipo.Text, rdbCartaoCredito.IsChecked, ddlEmpresa.Text, DateTime.Now, UsuarioGlobal.Usuario, iSeqRemes, rbtCadOptante.IsChecked, Connection, Transaction)
    '    If Not _Retorno.Sucesso Then
    '        blnRetorno = True
    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '    End If

    '    'RadMessageBox.Show(LoadMessage(Trim(nomeArquivo)), Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
    '    'MsgBox("Arquivo " & Trim(nomeArquivo) & " foi gerado com sucesso.", vbInformation, Me.Text)
    '    'lblTotalRegistro.Text = Idioma.RetornaMensagem("TotalRegistro") & " " & grdStatus.ChildRows.Count

    '    'codigoBanco = ""
    '    'strNomeBanco = ""
    '    'nomeArquivo = ""
    'End Sub
    'Private Sub CreateFile(nomeArquivo As String, codigoBanco As String)
    '    Dim _ContaCorrente As New ContaCorrente
    '    Dim crfile As IO.File
    '    BuscaBanco(codigoBanco)

    '    'blnContGera = True

    '    'Se for Cobranca Bancaria
    '    'If rdbBanco.IsChecked = True Or rbtCadOptante.IsChecked Then
    '    '    If codigoBanco = "" Then
    '    '        blnRetorno = True
    '    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.INFORME_CODIGO_BANCO.Descricao, ErrorConstants.INFORME_CODIGO_BANCO.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.CampoBranco)
    '    '    End If

    '    '    If codigoBanco <> "" And strNomeBanco = "" Then
    '    '        blnRetorno = True
    '    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.CAMPO_CODIGO_BANCO_INVALIDO.Descricao, ErrorConstants.CAMPO_CODIGO_BANCO_INVALIDO.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.CampoBranco)
    '    '    End If

    '    '    'Se for Cartão de Credito
    '    'Else
    '    '    If ddlEmpresa.SelectedIndex = -1 Then
    '    '        blnRetorno = True
    '    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.INFORME_TIPO_CARTAO_CREDITO_GERAR_ARQUIVO.Descricao, ErrorConstants.INFORME_TIPO_CARTAO_CREDITO_GERAR_ARQUIVO.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.CampoBranco)
    '    '    End If

    '    '    If ddlEmpresa.Text = "Adyen" Then
    '    '        If dtpDtAgendamento.Value <= Date.MinValue Or dtpDtAgendamento.Value.ToString("dd/MM/yyyy") < Date.Now.ToString("dd/MM/yyyy") Then
    '    '            blnRetorno = True
    '    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, "Data de agendamento inválida.", ErrorConstants.INFORME_TIPO_CARTAO_CREDITO_GERAR_ARQUIVO.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.CampoBranco)
    '    '        End If

    '    '        If rdpHrAgendamento.Value < DateTime.Now().ToString("HH:mm") And dtpDtAgendamento.Value.ToString("dd/MM/yyyy") = DateTime.Now.ToString("dd/MM/yyyy") Then
    '    '            blnRetorno = True
    '    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, "Data de agendamento inválida.", ErrorConstants.INFORME_TIPO_CARTAO_CREDITO_GERAR_ARQUIVO.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.CampoBranco)
    '    '        End If
    '    '    End If
    '    'End If

    '    'If Trim(nomeArquivo) = "" And ddlEmpresa.Text <> "Adyen" Then
    '    '    blnRetorno = True
    '    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.NOME_ARQUIVO_REMESSA_DEVE_SER_INFORMADA.Descricao, ErrorConstants.NOME_ARQUIVO_REMESSA_DEVE_SER_INFORMADA.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.CampoBranco)
    '    'End If

    '    'If blnEmpresaRJ Then
    '    '    'Validação das datas
    '    '    If dtpPeriodoInicial.Value <= Date.MinValue Then
    '    '        blnRetorno = True
    '    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.DATA_INICIAL_NAO_E_VALIDA.Descricao, ErrorConstants.DATA_INICIAL_NAO_E_VALIDA.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.CampoBranco)
    '    '        Beep()
    '    '    ElseIf dtpPeriodoFinal.Value <= Date.MinValue Then
    '    '        blnRetorno = True
    '    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.DATA_FINAL_NAO_E_VALIDA.Descricao, ErrorConstants.DATA_INICIAL_NAO_E_VALIDA.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.CampoBranco)
    '    '        Beep()

    '    '    End If

    '    '    If Not IsDate(dtpPeriodoInicial.Value) Or Not IsDate(dtpPeriodoFinal.Value) Then
    '    '        blnRetorno = True
    '    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.DATA_FINAL_NAO_E_VALIDA.Descricao, ErrorConstants.DATA_INICIAL_NAO_E_VALIDA.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '    '        Beep()
    '    '    End If

    '    '    If CDate(dtpPeriodoFinal.Value) < CDate(dtpPeriodoInicial.Value) Then
    '    '        blnRetorno = True
    '    '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.DATA_INICIAL_MAIOR_QUE_DATA_FINAL.Descricao, ErrorConstants.DATA_INICIAL_MAIOR_QUE_DATA_FINAL.Id, True, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '    '        Beep()
    '    '    End If
    '    'End If

    '    'EXIBIR ERRO
    '    If blnRetorno Then
    '        ExibirErro()
    '        Exit Sub
    '    End If


    '    '        If txtCodBanco.Text <> "409" And txtCodBanco.Text <> "237" And txtCodBanco.Text <> "347" Then
    '    '            MsgBox "Esse banco ainda não possui um driver para geração de arquivo remessa.", vbInformation, Me.Caption
    '    '            txtCodBanco.SetFocus
    '    '            Exit Sub
    '    '        End If

    '    'Trata o path informado

    '    '*********Retirado validação p/ o nome do arquivo na geração de arquivo de Cartão de Crédito**************
    '    '        Y = 1
    '    '        Do While Y > 0
    '    '            X = InStr(Y + 1, txtArquivo.Text, "\", vbTextCompare)
    '    '            If X = 0 Then Exit Do
    '    '            Y = X
    '    '        Loop

    '    '        If Len(Dir(Mid(txtArquivo.Text, 1, Y))) = 0 Then
    '    '            MsgBox "Caminho do arquivo para remessa inválido !", vbExclamation, Me.Caption
    '    '            If txtArquivo.Enabled Then txtArquivo.SetFocus
    '    '            Exit Sub
    '    '        End If
    '    '*********************************************************************************************************

    '    'MUDA O CURSOR NO MOUSE
    '    Me.Cursor = Cursors.WaitCursor
    '    '---------------------------------

    '    'CRIA ARQUIVO
    '    If ddlEmpresa.Text <> "Adyen" Then
    '        If Not IO.File.Exists(nomeArquivo) Then
    '            Dim file As System.IO.FileStream
    '            file = System.IO.File.Create(nomeArquivo)
    '            file.Close()
    '        End If
    '    End If

    '    'CONTROLE DE TRANSACAO
    '    Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '    connection.Open()
    '    Dim Transacao As SqlTransaction = connection.BeginTransaction()
    '    '-----------------------------------------------------------------------------------------------------------------------------------------------------------

    '    grdStatus.Rows.Clear()

    '    'Se for Cobranca Bancaria
    '    If rdbBanco.IsChecked = True Then
    '        If (RadMessageBox.Show("Confirma a geração do arquivo remessa para o banco " & vbLf & vbLf & Trim(lblNomeBanco.Text) & " ? ", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = Windows.Forms.DialogResult.Yes) Then
    '            'If MsgBox("Confirma a geração do arquivo remessa para o banco " & vbLf & vbLf & Trim(lblNomeBanco.Text) & " ? ", vbQuestion + vbYesNo, Me.Text) = vbYes Then
    '            If strTipoCta = "Desconto" Then

    '                _ContaCorrente = clConsultarContaCorrente.BuscaNumCtaVincDescContaCorrente(codigoBanco, strAgencia, strConta)
    '                If Not _ContaCorrente.Sucesso Then
    '                    blnRetorno = True
    '                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _ContaCorrente.MsgErro, _ContaCorrente.NumErro, _ContaCorrente.Sucesso, _ContaCorrente.TipoErro, _ContaCorrente.ImagemErro)

    '                    ExibirErro()
    '                End If

    '                Select Case codigoBanco
    '                    Case "237"
    '                        FGeraRemessa237DESC(connection, Transacao) ' foi foi
    '                    Case Else
    '                        RadMessageBox.Show("Não existe layout de remessa cadastrado para este banco no sistema !", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
    '                        'MsgBox("Não existe layout de remessa cadastrado para este banco no sistema !")
    '                End Select
    '            Else
    '                Select Case codigoBanco
    '                    Case "001"
    '                        FGeraRemessa001(connection, Transacao) 'foi foi
    '                        If blnRetorno Then
    '                            Transacao.Rollback()
    '                            connection.Close()

    '                            'ExibirErro()
    '                        Else
    '                            Transacao.Commit()
    '                            connection.Close()
    '                        End If
    '                        'Case "409"
    '                        '    FGeraRemessa409()
    '                        '    If blnRetorno Then
    '                        '        Transacao.Rollback()
    '                        '        connection.Close()

    '                        '        ExibirErro()
    '                        '    End If
    '                    Case "237"
    '                        FGeraRemessa237(connection, Transacao) 'foi foi foi
    '                        If blnRetorno Then
    '                            Transacao.Rollback()
    '                            connection.Close()

    '                            ExibirErro()
    '                        Else
    '                            Transacao.Commit()
    '                            connection.Close()
    '                        End If
    '                        'Case "347"
    '                        '    FGeraRemessa347()
    '                        '    If blnRetorno Then
    '                        '        Transacao.Rollback()
    '                        '        connection.Close()
    '                        '
    '                        '        ExibirErro()
    '                        '    End If
    '                        'Case "291"
    '                        '    'FGeraRemessa291()
    '                        '    If blnRetorno Then
    '                        '        Transacao.Rollback()
    '                        '        connection.Close()
    '                        '
    '                        '        ExibirErro()
    '                        '    End If
    '                    Case "341"
    '                        Select Case ddlTipo.Text.Trim().ToUpper()
    '                            Case "CARNÊ"
    '                                FGeraRemessa341Carne(connection, Transacao)
    '                            Case "BOLETO UNIFICADO"
    '                                FGeraRemessa341BoletoUnificado(connection, Transacao)
    '                            Case Else
    '                                FGeraRemessa341(connection, Transacao) 'foi foi foi ----
    '                        End Select

    '                        If blnRetorno Then
    '                            Transacao.Rollback()
    '                            connection.Close()

    '                            ExibirErro()
    '                        Else
    '                            Transacao.Commit()
    '                            connection.Close()
    '                        End If
    '                        'Case "224"
    '                        '    'FGeraRemessa224()
    '                        '    If blnRetorno Then
    '                        '        Transacao.Rollback()
    '                        '        connection.Close()
    '                        '
    '                        '        ExibirErro()
    '                        '    End If
    '                    Case "033"
    '                        Select Case ddlTipo.Text.Trim().ToUpper()
    '                            Case "BOLETO UNIFICADO"
    '                                FGeraRemessaBoletoUnificado033(connection, Transacao)
    '                            Case Else
    '                                FGeraRemessa033(connection, Transacao) 'foi foi foi
    '                        End Select

    '                        If blnRetorno Then
    '                            Transacao.Rollback()
    '                            connection.Close()

    '                            ExibirErro()
    '                        Else
    '                            Transacao.Commit()
    '                            connection.Close()
    '                        End If
    '                    Case "399"
    '                        FGeraRemessa399(connection, Transacao) 'foi foi foi
    '                        If blnRetorno Then
    '                            Transacao.Rollback()
    '                            connection.Close()

    '                            ExibirErro()
    '                        Else
    '                            Transacao.Commit()
    '                            connection.Close()
    '                        End If
    '                    Case "356"
    '                        FGeraRemessa356(connection, Transacao) 'foi foi foi
    '                        If blnRetorno Then
    '                            Transacao.Rollback()
    '                            connection.Close()

    '                            ExibirErro()
    '                        Else
    '                            Transacao.Commit()
    '                            connection.Close()
    '                        End If
    '                        'Case "422"
    '                        '    'FGeraRemessa422()
    '                        '    If blnRetorno Then
    '                        '        Transacao.Rollback()
    '                        '        connection.Close()
    '                        '
    '                        '        ExibirErro()
    '                        '    End If
    '                        'Case "479"
    '                        '    'FGeraRemessa479()
    '                        '    If blnRetorno Then
    '                        '        Transacao.Rollback()
    '                        '        connection.Close()
    '                        '
    '                        '        ExibirErro()
    '                        '    End If
    '                    Case "104"
    '                        '05/04/2018 - Fernando
    '                        'Novo banco Caixa (TeleAlarme)
    '                        FGeraRemessa104(connection, Transacao) 'foi foi foi
    '                        If blnRetorno Then
    '                            Transacao.Rollback()
    '                            connection.Close()

    '                            ExibirErro()
    '                        Else
    '                            Transacao.Commit()
    '                            connection.Close()
    '                        End If
    '                    Case "041"
    '                        '18/10/2018 - Douglas
    '                        'Novo banco Banrisul (Sul do país)
    '                        FGeraRemessa041(connection, Transacao) 'foi foi foi
    '                        If blnRetorno Then
    '                            Transacao.Rollback()
    '                            connection.Close()

    '                            ExibirErro()
    '                        Else
    '                            Transacao.Commit()
    '                            connection.Close()
    '                        End If

    '                    Case Else
    '                        RadMessageBox.Show("Não existe layout de remessa cadastrado para este banco no sistema!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
    '                        'MsgBox("Não existe layout de remessa cadastrado para este banco no sistema!")
    '                        ' Set clSig = New TOAcesso
    '                        ' With clSig
    '                        '     .OpenConnection gsAdo_StrConn
    '                        '     .Execute ("UPDATE CONTAS_A_RECEBER SET StatusBco='1' WHERE CodBanco = '" & txtCodBanco.Text & "'")
    '                        '     .CloseConnection
    '                        ' End With
    '                End Select
    '            End If
    '        End If
    '    ElseIf rbtCadOptante.IsChecked Then
    '        Select Case codigoBanco
    '            Case "104"
    '                FGeraRemessaCadOptante104(connection, Transacao) 'foi foi
    '                If blnRetorno Then
    '                    Transacao.Rollback()
    '                    connection.Close()

    '                    ExibirErro()
    '                Else
    '                    Transacao.Commit()
    '                    connection.Close()
    '                End If
    '        End Select
    '    Else
    '        If (RadMessageBox.Show(IIf(ddlEmpresa.Text <> "Adyen", "Confirma a geração do arquivo remessa para a Empresa " & vbLf & vbLf & Trim(ddlEmpresa.Text) & " ? ", "Confirma o agendamento para a Adyen no dia: " + dtpDtAgendamento.Value.ToString("dd/MM/yyyy") + " às " + rdpHrAgendamento.Value.ToString("HH:mm")), Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = Windows.Forms.DialogResult.Yes) Then
    '            'If MsgBox(IIf(ddlEmpresa.Text <> "Adyen", "Confirma a geração do arquivo remessa para a Empresa " & vbLf & vbLf & Trim(ddlEmpresa.Text) & " ? ", "Confirma o agendamento para a Adyen no dia: " + dtpDtAgendamento.Value.ToString("dd/MM/yyyy") + " às " + rdpHrAgendamento.Value.ToString("HH:mm")), vbQuestion + vbYesNo, Me.Text) = vbYes Then

    '            Select Case ddlEmpresa.Text
    '                Case "MasterCard/Diners"
    '                    FGeraRemessaMasterCard(connection, Transacao)
    '                    If blnRetorno Then
    '                        Transacao.Rollback()
    '                        connection.Close()

    '                        ExibirErro()
    '                    Else
    '                        Transacao.Commit()
    '                        connection.Close()
    '                    End If
    '                Case "Visa"
    '                    FGeraRemessaVisa(connection, Transacao)
    '                    If blnRetorno Then
    '                        Transacao.Rollback()
    '                        connection.Close()

    '                        ExibirErro()
    '                    Else
    '                        Transacao.Commit()
    '                        connection.Close()
    '                    End If
    '                Case "AMEX"
    '                    FGeraRemessaAMEX(connection, Transacao)
    '                    If blnRetorno Then
    '                        Transacao.Rollback()
    '                        connection.Close()

    '                        ExibirErro()
    '                    Else
    '                        Transacao.Commit()
    '                        connection.Close()
    '                    End If
    '                Case "Cielo"
    '                    FGeraRemessaCielo(connection, Transacao)
    '                    If blnRetorno Then
    '                        Transacao.Rollback()
    '                        connection.Close()

    '                        ExibirErro()
    '                    Else
    '                        Transacao.Commit()
    '                        connection.Close()
    '                    End If
    '                Case "Adyen"
    '                    AgendarTitulosAdyen(connection, Transacao)
    '                    If blnRetorno Then
    '                        Transacao.Rollback()
    '                        connection.Close()
    '                        ExibirErro()
    '                    Else
    '                        TrataSucesso("Títulos agendados com sucesso!", "Sucesso", connection, Transacao)
    '                    End If
    '            End Select

    '        End If

    '    End If


    '    'VOLTA O CURSOR NO MOUSE AO NORMAL
    '    'Me.Cursor = Cursors.Default
    '    '---------------------------------
    'End Sub
    Private Sub CarregaRelatorioExcel(nomeArquivo As String)

        Dim clBuscaChecagem As New BuscaTempCheckArqRetorno
        Dim retorno As New Retorno
        Dim dtArqRetorno As New DataTable("Chacagem Arq. Retorno")
        Dim nome = Path.GetFileNameWithoutExtension(nomeArquivo)

        retorno = clBuscaChecagem.BuscaRelExcelCheckArqRetornoMultasJuros(dtArqRetorno)

        ' Throw caso haja erro
        If Not retorno.Sucesso Then
            Funcoes.CriaLog(Application.StartupPath & "\RelatorioErros.txt", retorno.MsgErro)
            'Throw New Exception(retorno.MsgErro)
        Else



            ' Abre o diálogo SaveFileDialog
            Dim saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx"
            saveFileDialog.InitialDirectory = "C:\Users\flavio.conceicao\Documents\Analises\Nexxtera\Relatorios checagem e divergentes\TesteExportacao\"
            'If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Dim data As DateTime = Now
            Dim dataAtual As String = data.ToString("yyyy-MM-dd")
            'Dim strPath As String = "C:\Users\flavio.conceicao\Documents\Analises\Nexxtera\Relatorios checagem e divergentes\TesteExportacao\" + dataAtual + "\" + nome + ".xlsx"
            'Dim strPath As String = "L:\Contas a Receber\CHECAGENS\Relatório Checagem\" + dataAtual + "\" + nome + ".xlsx"
            Dim strPath As String = "\\br06filesrv02v\Faturacaobr\Contas a Receber\CHECAGENS\Relatório Checagem\" + dataAtual + "\" + nome + ".xlsx"



            ' Cria o relatório
            retorno = Funcoes.CriaRelatorioXLSX(dtArqRetorno, strPath)

            ' Throw caso haja erro
            If Not retorno.Sucesso Then Throw New Exception(retorno.MsgErro)

            'Me.Cursor = Cursors.Default
            'If RadMessageBox.Show("Relatório gerado com sucesso!" & vbCrLf & "Deseja visualizar agora?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Info) = vbYes Then
            'System.Diagnostics.Process.Start(strPath)
            'End If
            'End If
        End If
    End Sub
    Private Sub Check341(nomeArquivo As String) 'Check de Retorno do ITAU 
        Try


            Dim cr_tb As New ContaReceber
            Dim _TempCheckArqRetorno As New TempCheckArqRetorno
            Dim _Retorno As New Retorno

            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "ITAU ainda não liberado para faturamento. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------

            Dim iFile As Integer
            Dim sBuffer As String
            Dim sNomeBanco As String, sBanco As String, sNumAviso As String, sConta As String, sDataPagto As String, sNumtit As String, sNomeCliente As String
            Dim sAgencia As String, sCodIncons As String, sCodClie As String, sTamCodClie As Integer, sSeqTit As String
            Dim sTotalRegistros As String, sCodOcor As String, sBancoDeb As String, sAgenciaDeb As String, sNumCtaDeb As String, dblValor As Double, sEndereco As String, sUF As String, sCidade As String, sCEP As String, sEmpresa As String
            Dim iTamCodClie As Integer, sid As String
            Dim sDtVcto As String, sDtEmissao As String
            Dim sTituloCompleto As String, sTitDesc As String
            Dim sTitulosDesc As String, sTitulosJaDesc As String
            Dim sTitulosJaVolt As String = ""
            Dim sTitulosVolt As String = ""
            Dim sMsgTitDesc As String = ""
            Dim iCont As Integer, i As Integer
            Dim strNossoNumBco As String, sVcto As String, sMsgErros As String, sNumConta As String, sTaxaMulta As String
            Dim sDataArq As Date
            Dim aNumTit(iCont)
            Dim bPossuiInconsistencia As Boolean = False
            Dim MultaM As Double = 0
            Dim JurosM As Double = 0
            Dim vlrInd As Double = 0
            Dim _Banco As New Banco

            iFile = FreeFile()


            'VOLTA O CURSOR NO MOUSE
            'Me.Cursor = Cursors.WaitCursor
            '---------------------------------


            If Dir(nomeArquivo) = "" Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                Exit Sub
            End If

            fsFile = New System.IO.StreamReader(nomeArquivo)

            sBuffer = fsFile.ReadLine


            'IDENTIFICA SE LAYOUT EH DC OU BO
            'COBRANCA: LIQUIDACAO DE TITULOS NORMAIS
            'EMPRESTIMO: LIQUIDACAO DE TITULOS DESCONTADOS
            If Trim(Mid(sBuffer, 12, 15)) = "COBRANCA" Or Trim(Mid(sBuffer, 12, 15)) = "EMPRESTIMO" Then
                '*******************************
                '***** BOLETO ******************

                '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
                'MsgBox "ITAU ainda não liberado para checagem de retorno. Consultar analista de sistemas !"
                'Exit Sub
                '---------------------------------------------------------------
                ReDim aNumTit(0)
                iCont = 0

                sMsgErros = ""

                Do While (sBuffer) <> Nothing

                    'Checa se eh HEADER
                    If Mid(sBuffer, 1, 1) = "0" Then

                        'Checa se eh RETORNO 
                        If Mid(sBuffer, 2, 1) <> "2" Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                            Exit Sub
                        End If

                        sBanco = Mid(sBuffer, 77, 3)
                        sAgencia = Mid(sBuffer, 27, 4)
                        sConta = Mid(sBuffer, 33, 6)
                        sDataArq = CDate(Mid(sBuffer, 95, 2) & "/" & Mid(sBuffer, 97, 2) & "/20" & Mid(sBuffer, 99, 2))
                        sNumAviso = Mid(sBuffer, 109, 5)
                        sNomeBanco = Trim(Mid(sBuffer, 80, 15))
                        sEmpresa = Trim(Mid(sBuffer, 47, 28))

                        'If IIf(IsNothing(sBanco), "", sBanco) <> Trim(codigoBanco) Then
                        '    blnRetorno = True
                        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        '    Exit Sub
                        'End If

                    End If

                    'Checa se eh DETALHE
                    If Trim(Mid(sBuffer, 1, 1)) = "1" Then

                        sCodOcor = Mid(sBuffer, 109, 2)
                        sCodIncons = Trim(Mid(sBuffer, 378, 2))
                        If sCodOcor = "25" Or sCodOcor = "24" Or sCodOcor = "57" Then
                            sCodIncons = Trim(Mid(sBuffer, 302, 4))
                        End If
                        sNomeCliente = Trim(Mid(sBuffer, 325, 30))
                        sBancoDeb = ""
                        sAgenciaDeb = ""
                        sNumCtaDeb = ""

                        If Mid(sBuffer, 109, 2) = "06" Or Mid(sBuffer, 109, 2) = "10" Then
                            'Se for liquidacao pega valor pago
                            dblValor = (Val(Mid(sBuffer, 254, 13)) / 100) ' + (Val(Mid(sBuffer, 267, 13)) / 100) - (Val(Mid(sBuffer, 241, 13)) / 100) - (Val(Mid(sBuffer, 228, 13)) / 100)
                        Else
                            'Se for confirmacao de recebimento pega valor do titulo
                            dblValor = (Val(Mid(sBuffer, 153, 13)) / 100)
                        End If

                        If Mid(sBuffer, 109, 2) = "06" Or Mid(sBuffer, 109, 2) = "10" Or Mid(sBuffer, 109, 2) = "16" Or Mid(sBuffer, 109, 2) = "30" Then
                            'Se for liquidacao pega dt pgto
                            sDataPagto = CDate("20" & Mid(sBuffer, 115, 2) & "/" & Mid(sBuffer, 113, 2) & "/" & Mid(sBuffer, 111, 2))
                        Else
                            'Se for confirmacao de recebimento pega dt vcto

                            sDataPagto = CDate("20" & Mid(sBuffer, 151, 2) & "/" & Mid(sBuffer, 149, 2) & "/" & Mid(sBuffer, 147, 2))
                        End If

                        sEndereco = ""
                        sUF = ""
                        sCidade = ""
                        sCEP = ""
                        If Mid(sBuffer, 109, 2) = "30" Then
                            sNumtit = ""
                            sSeqTit = ""
                        Else
                            If sEmpresa = "TELEATLANTIC RIO MON AL LTDA" Then
                                'Carteira sem cobranca nao traz dados do titulo
                                If Trim(Mid(sBuffer, 83, 3)) <> "175" Then
                                    sNumtit = Trim(Mid(sBuffer, 117, 10))
                                    sSeqTit = Strings.Right(sNumtit, 2)
                                    sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 3)
                                Else
                                    sNumtit = "**CNR**"
                                    sSeqTit = ""
                                End If
                            Else
                                sNumtit = Trim(Mid(sBuffer, 38, 25))
                                sSeqTit = Strings.Right(sNumtit, 2)
                                sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
                            End If
                        End If

                        ''TESTT

                        'If Mid(sBuffer, 109, 2) = "30" Then
                        '    sNumtit = ""
                        '    sSeqTit = ""
                        'End If
                        'If sEmpresa = "TELEATLANTIC RIO MON AL LTDA" Then
                        '    'Carteira sem cobranca nao traz dados do titulo
                        '    If Trim(Mid(sBuffer, 83, 3)) <> "175" Then
                        '        sNumtit = Trim(Mid(sBuffer, 117, 10))
                        '        sSeqTit = Strings.Right(sNumtit, 2)
                        '        sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 3)
                        '    Else
                        '        sNumtit = "**CNR**"
                        '        sSeqTit = ""
                        '    End If
                        'Else
                        '    sNumtit = Trim(Mid(sBuffer, 38, 25))
                        '    sSeqTit = Strings.Right(sNumtit, 2)
                        '    sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
                        'End If

                        'Vencimento
                        'Tratativa para títulos já baixados/quitado (Código ocorrencia 16)
                        If Mid(sBuffer, 109, 2) = "16" Or Mid(sBuffer, 109, 2) = "30" Then
                            sVcto = CDate("20" & Mid(sBuffer, 115, 2) & "/" & Mid(sBuffer, 113, 2) & "/" & Mid(sBuffer, 111, 2))
                        Else
                            sVcto = CDate(Trim(Mid(sBuffer, 147, 2)) & "/" & Trim(Mid(sBuffer, 149, 2)) & "/20" & Trim(Mid(sBuffer, 151, 2)))
                        End If


                        'nosso numero no banco
                        strNossoNumBco = Trim(Mid(sBuffer, 86, 9))


                        sNumConta = Regex.Replace(sConta, "[^0-9]", "")

                        If Not String.IsNullOrEmpty(sNumConta) Then
                            _Banco = clBuscaBancoContaCorrente.BuscaTaxaJurosParametrizada("341", sAgencia, sNumConta)
                        End If

                        'Consulta para pegar valor original do titulo.
                        contaReceber.VlrInd = consulta.ConsultarValorTituloContaReceber(sNumtit, sSeqTit)

                        If contaReceber.VlrInd > 0 Then
                            vlrInd = contaReceber.VlrInd
                        End If

                        sTaxaMulta = _Banco.TaxaMulta / 100

                        If Val(Mid(sBuffer, 267, 13)) = 0 Then 'Caso cliente pague na data correta juros e multa e atribuido 0 
                            JurosM = 0
                            MultaM = 0
                        Else
                            MultaM = Val(vlrInd) * sTaxaMulta 'Calcula o valor da multa com base no valor real do titulo.
                            MultaM = FormatNumber(MultaM, 2)  'Formata número para duas casas decimais
                            If MultaM > Val(Mid(sBuffer, 267, 13)) / 100 Then
                                MultaM = 0
                                JurosM = Val(Mid(sBuffer, 267, 13)) / 100
                            Else
                                JurosM = ((Val(Mid(sBuffer, 267, 13)) / 100) - MultaM)
                            End If
                        End If


                        'Entrada Rejeitada/Baixa Rejeitada/Instrução Rejeitada/Alteração de Dados Rejeitada/Cobrança Contratual Bloqueada
                        If (sCodOcor = "03" Or sCodOcor = "15" Or sCodOcor = "16" Or sCodOcor = "17" Or sCodOcor = "18") Then
                            'As linhas abaixo deverão ser analisadas no final do processo
                            'sMsgErros = sMsgErros & IIf(Trim(sMsgErros) = "", "", vbCr) & "Nosso Número : " & Trim(strNossoNumBco) & "  Cód.Ocorrência : " & sCodOcor & " Cód.Inconsistência : " & sCodIncons
                            bPossuiInconsistencia = True
                            Funcoes.CriaLog("LogChecagem.log", Format(Funcoes.PegaData, "dd/MM/yyyy HH:mm:ss") & " - " & System.IO.Path.GetFileName(nomeArquivo) & " - " & IIf(Trim(sMsgErros) = "", "", vbCr) & "Nosso Número : " & Trim(strNossoNumBco) & "  Cód.Ocorrência : " & sCodOcor & " Cód.Inconsistência : " & sCodIncons)
                        Else

                            _TempCheckArqRetorno.CodBanco = sBanco
                            _TempCheckArqRetorno.NomeBanco = sNomeBanco
                            _TempCheckArqRetorno.NumAviso = sNumAviso
                            _TempCheckArqRetorno.CodAgen = sAgencia
                            _TempCheckArqRetorno.NumCta = sConta
                            _TempCheckArqRetorno.DtArq = sDataArq
                            _TempCheckArqRetorno.NumTit = sNumtit
                            _TempCheckArqRetorno.DtPagto = sDataPagto
                            _TempCheckArqRetorno.NomeCliente = sNomeCliente
                            _TempCheckArqRetorno.CodOcor = sCodOcor
                            _TempCheckArqRetorno.CodIncons = sCodIncons
                            _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                            _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                            _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                            _TempCheckArqRetorno.Valor = dblValor
                            _TempCheckArqRetorno.Endereco = sEndereco
                            _TempCheckArqRetorno.Cidade = sCidade
                            _TempCheckArqRetorno.UF = sUF
                            _TempCheckArqRetorno.Cep = sCEP
                            _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                            _TempCheckArqRetorno.SeqTit = sSeqTit
                            _TempCheckArqRetorno.DtVcto = sVcto
                            _TempCheckArqRetorno.NossoNumeroBco = strNossoNumBco
                            _TempCheckArqRetorno.SeqTit = Trim(sSeqTit)
                            _TempCheckArqRetorno.VlrJuros = JurosM 'validar caso seja nulo por 0
                            _TempCheckArqRetorno.VlrMulta = MultaM 'validar caso seja nulo por 0

                            cr_tb.CodBanco = sBanco
                            cr_tb.CodAgen = sAgencia
                            cr_tb.NumCta = sConta
                            cr_tb.NumTit = sNumtit + sSeqTit
                            cr_tb.DtPgto = sDataPagto
                            cr_tb.VlrInd = dblValor
                            cr_tb.SeqTit = sSeqTit
                            cr_tb.DtVcto = sVcto
                            cr_tb.NossoNumeroBco = strNossoNumBco
                            cr_tb.SeqTit = Trim(sSeqTit)
                            cr_tb.VlrJuros = JurosM 'validar caso seja nulo por 0
                            cr_tb.Carteira = _Banco.ContaCorrente.CodCarteira

                            'Gera o codigo de barras e grava na table
                            Dim bb_tb As New BoletoNet.BoletoBancario
                            bb_tb = ObterBoletoBancario(cr_tb)
                            If Not IsNothing(bb_tb) Then
                                _Retorno = clInserirTempCheckArqRetorno.InsertCodigoBarra(LoadCodigoBarra(cr_tb, bb_tb.Boleto.CodigoBarra))
                            End If
                            If Not _Retorno.Sucesso Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                                Exit Sub
                            End If

                            _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                            If Not _Retorno.Sucesso Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                                Exit Sub
                            End If
                        End If

                        'Verifica marcação dos títulos (Desconto)
                        If sCodOcor = "04" Or sCodOcor = "47" Then
                            sDtVcto = CDate("20" & Mid(sBuffer, 151, 2) & "-" & Mid(sBuffer, 149, 2) & "-" & Mid(sBuffer, 147, 2))
                            lstContaReceber = clConsultacontasReceber.BuscaDtEmissaoTitDescDtVctoSeqTitNumTitContasReceber(sNumtit, sSeqTit, sDtVcto)
                            If Not lstContaReceber(0).Sucesso And lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstContaReceber(0).MsgErro, lstContaReceber(0).NumErro, lstContaReceber(0).Sucesso, lstContaReceber(0).TipoErro, lstContaReceber(0).ImagemErro)

                                Exit Sub
                            ElseIf lstContaReceber(0).TipoErro = DadosGenericos.TipoErro.None Then

                                sTitDesc = lstContaReceber(0).TitDesc
                                sDtEmissao = lstContaReceber(0).DtEmissao
                                sTituloCompleto = "Título: " & sNumtit & " Seq.: " & sSeqTit & " Dt. Emissão: " & sDtEmissao
                                If sCodOcor = "47" Then
                                    If sTitDesc = "S" Then
                                        sTitulosJaDesc = sTitulosJaDesc & sTituloCompleto & vbNewLine
                                    Else
                                        ReDim aNumTit(0)
                                        sTitulosDesc = sTitulosDesc & sTituloCompleto & vbNewLine
                                        aNumTit(iCont).NumTit = sNumtit
                                        aNumTit(iCont).SeqTit = sSeqTit
                                        aNumTit(iCont).dtEmissao = Format(sDtEmissao, "yyyy-MM-dd")
                                        aNumTit(iCont).TipoMarcacao = "S"
                                        iCont = iCont + 1
                                    End If
                                ElseIf sCodOcor = "04" Then
                                    If sTitDesc = "S" Then
                                        ReDim aNumTit(0)
                                        sTitulosVolt = sTitulosVolt & sTituloCompleto & vbNewLine
                                        aNumTit(iCont).NumTit = sNumtit
                                        aNumTit(iCont).SeqTit = sSeqTit
                                        aNumTit(iCont).dtEmissao = Format(sDtEmissao, "yyyy-MM-dd")
                                        aNumTit(iCont).TipoMarcacao = "N"
                                        iCont = iCont + 1
                                    Else
                                        sTitulosJaVolt = sTitulosJaVolt & sTituloCompleto & vbNewLine
                                    End If
                                End If
                            End If

                        End If
                    End If

                    sBuffer = fsFile.ReadLine
                Loop

                If sTitulosJaDesc <> "" Then
                    sTitulosJaDesc = "TÍTULOS QUE VIERAM PARA MARCAR COMO DESCONTO E JÁ ESTAVAM MARCADOS EM NOSSA BASE: " & vbNewLine & sTitulosJaDesc
                    sMsgTitDesc = sTitulosJaDesc & vbNewLine
                End If
                If sTitulosDesc <> "" Then
                    sTitulosDesc = "TÍTULOS QUE VIERAM PARA MARCAR COMO DESCONTO: " & vbNewLine & sTitulosDesc
                    sMsgTitDesc = sMsgTitDesc & sTitulosDesc & vbNewLine
                End If
                If sTitulosJaVolt <> "" Then
                    sTitulosJaVolt = "TÍTULOS QUE VIERAM PARA MARCAR COMO NÃO DESCONTADOS E JÁ ESTAVAM MARCADOS NA NOSSA BASE: " & vbNewLine & sTitulosJaVolt
                    sMsgTitDesc = sMsgTitDesc & sTitulosJaVolt & vbNewLine
                End If
                If sTitulosVolt <> "" Then
                    sTitulosVolt = "TÍTULOS QUE VIERAM PARA MARCAR COMO NÃO DESCONTADOS: " & vbNewLine & sTitulosVolt
                    sMsgTitDesc = sMsgTitDesc & sTitulosVolt & vbNewLine
                End If

                If sMsgTitDesc <> "" Then
                    VG_bDescTit = False
                    'TelesystemCtaReceberTitDescInfo.txtInfoTit.Text = sMsgTitDesc
                    'TelesystemCtaReceberTitDescInfo.ShowDialog()

                    If VG_bDescTit = True Then
                        'CONTROLE DE TRANSACAO
                        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
                        connection.Open()
                        Dim Transacao As SqlTransaction = connection.BeginTransaction()

                        For i = 0 To iCont - 1
                            _Retorno = clAlterarContasReceber.AlterarDtTitDescTitDesContasReceber(aNumTit(i).NumTit, aNumTit(i).SeqTit, aNumTit(i).dtEmissao, aNumTit(i).TipoMarcacao, connection, Transacao)
                            If Not _Retorno.Sucesso Then
                                Transacao.Rollback()
                                connection.Close()
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                Exit Sub
                            End If
                        Next

                        Transacao.Commit()
                        connection.Close()
                    End If
                End If

                'IsCobEmp = True
                ''Atualiza a tabela contas_a_receber com o nosso numero no banco
                'If MsgBox("Deseja gravar o Nosso Número ?", vbYesNo, Me.Text) = vbYes Then
                '    _Retorno = clAlterarContasReceber.AlteraNossoNumeroBcoContasAReceber(False, connection, Transacao)
                '    If Not _Retorno.Sucesso Then
                '        blnRetorno = True
                '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                '        Exit Sub
                '    End If
                'End If

                ''Atualiza a tabela contas_a_receber se pagamento será por DDA
                'If MsgBox("Deseja gravar o DDA ?", vbYesNo, Me.Caption) = vbYes Then
                '    .Execute "UPDATE Contas_A_Receber " & _
                '             "SET ObsTit = (CASE WHEN LEFT(tempCheckArqRetorno.CodIncons,4)='1842' THEN '(DDA) ' ELSE '' END) + CONVERT(CHAR(100),ISNULL(obstit,'')) " & _
                '             "FROM tempCheckArqRetorno " & _
                '             "INNER JOIN Contas_A_Receber ON tempCheckArqRetorno.NumTit = Contas_A_Receber.NumTit AND " & _
                '                                            "tempCheckArqRetorno.SeqTit = Contas_A_Receber.SeqTit AND " & _
                '                                            "CONVERT(VARCHAR(8), tempCheckArqRetorno.DtVcto, 112) = CONVERT(VARCHAR(8), Contas_A_Receber.DtVcto, 112) " & _
                '            "WHERE tempCheckArqRetorno.CodOcor = '25' "
                'End If

                'If sMsgErros <> "" Then
                If bPossuiInconsistencia = True Then
                    'MsgBox("Existe(m) a(s) seguinte(s) inconsistência(s) : " & _
                    'vbCr & vbCr & sMsgErros & vbCr, vbExclamation, "Problemas para checar Arquivo Retorno")
                    'MsgBox("Existe(m) inconsistência(s) no arquivo, por favor checar o arquivo de log ""LogChecagem.log"" no caminho: " & Application.StartupPath, vbExclamation, "Problemas para checar Arquivo Retorno")
                    'RadMessageBox.Show("Existe(m) inconsistência(s) no arquivo, por favor checar o arquivo de log ""LogChecagem.log"" no caminho: " & Application.StartupPath & vbCr & "Problemas para checar Arquivo Retorno", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
                End If

            Else
                '*******************************
                '***** DEB AUTO ****************

                Do While (sBuffer) <> Nothing

                    'Checa se eh HEADER
                    If Mid(sBuffer, 8, 1) = "0" Then

                        'Checa se eh RETORNO
                        If Mid(sBuffer, 143, 1) <> "2" Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                            Exit Sub
                        End If

                        sBanco = Mid(sBuffer, 1, 3)

                        sAgencia = Mid(sBuffer, 54, 4)

                        If (sAgencia) = "    " Then
                            sAgencia = strAgenItau 'lstBanco(0).CodAgen
                        End If

                        sConta = Mid(sBuffer, 66, 5)

                        If (sConta) = "     " Then
                            sConta = strContaCorrenteItau 'lstBanco(0).Numcta
                        End If

                        sDataArq = CDate(Mid(sBuffer, 148, 4) & "-" & Mid(sBuffer, 146, 2) & "-" & Mid(sBuffer, 144, 2))
                        sNumAviso = Strings.Right(Mid(sBuffer, 158, 6), 5)
                        sNomeBanco = Trim(Mid(sBuffer, 103, 30))

                        'If IIf(IsNothing(sBanco), "", sBanco) <> Trim(codigoBanco) Then
                        '    blnRetorno = True
                        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        '    Exit Sub
                        'End If

                    End If


                    'Checa se eh DETALHE
                    If Trim(Mid(sBuffer, 8, 1)) = "3" Then

                        sCodOcor = "00"
                        sCodIncons = Trim(Mid(sBuffer, 231, 2))
                        sNomeCliente = Trim(Mid(sBuffer, 44, 30))
                        sBancoDeb = Trim(Mid(sBuffer, 1, 3))
                        sAgenciaDeb = Trim(Mid(sBuffer, 25, 4))
                        sNumCtaDeb = Mid(sBuffer, 37, 5) & "-" & Mid(sBuffer, 43, 1)
                        dblValor = Val(Mid(sBuffer, 120, 15)) / 100

                        If ("0000".Equals(Trim(Mid(sBuffer, 159, 4)))) Then
                            sDataPagto = CDate(Trim(Mid(sBuffer, 98, 4)) & "-" & Trim(Mid(sBuffer, 96, 2)) & "-" & Trim(Mid(sBuffer, 94, 2)))
                        Else
                            sDataPagto = CDate(Trim(Mid(sBuffer, 159, 4)) & "-" & Trim(Mid(sBuffer, 157, 2)) & "-" & Trim(Mid(sBuffer, 155, 2)))
                        End If

                        sEndereco = ""
                        sUF = ""
                        sCidade = ""
                        sCEP = ""

                        sNumtit = Strings.Left(Trim(Mid(sBuffer, 74, 15)), Len(Trim(Mid(sBuffer, 74, 15))) - 2)
                        sSeqTit = Strings.Right(Trim(Mid(sBuffer, 74, 15)), 2)
                        Dim TipoDupRetorno As List(Of ContaReceber) = clConsultacontasReceber.BuscaTipoDup(sNumtit, sSeqTit)
                        Dim tipoDup As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).TipoDup, " ")
                        Dim Situacao As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).Situacao, " ")

                        _TempCheckArqRetorno.CodBanco = sBanco
                        _TempCheckArqRetorno.NomeBanco = sNomeBanco
                        _TempCheckArqRetorno.NumAviso = sNumAviso
                        _TempCheckArqRetorno.CodAgen = sAgencia
                        _TempCheckArqRetorno.NumCta = sConta
                        _TempCheckArqRetorno.DtArq = sDataArq
                        _TempCheckArqRetorno.NumTit = sNumtit
                        _TempCheckArqRetorno.DtPagto = sDataPagto
                        _TempCheckArqRetorno.NomeCliente = sNomeCliente
                        _TempCheckArqRetorno.CodOcor = sCodOcor
                        _TempCheckArqRetorno.CodIncons = sCodIncons
                        _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                        _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                        _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                        _TempCheckArqRetorno.Valor = dblValor
                        _TempCheckArqRetorno.Endereco = sEndereco
                        _TempCheckArqRetorno.Cidade = sCidade
                        _TempCheckArqRetorno.UF = sUF
                        _TempCheckArqRetorno.Cep = sCEP
                        _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                        _TempCheckArqRetorno.SeqTit = Trim(sSeqTit)
                        _TempCheckArqRetorno.TipoDup = tipoDup
                        _TempCheckArqRetorno.Situacao = Situacao

                        _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                        If Not _Retorno.Sucesso Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                            Exit Sub
                        End If

                    End If

                    sBuffer = fsFile.ReadLine
                Loop
            End If
            Dim Con As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            Con.Open()
            Dim Tran As SqlTransaction = Con.BeginTransaction()

            'If MsgBox("Deseja gravar o Nosso Número e o aviso de DDA ?", vbYesNo, Me.Text) = vbYes Then
            _Retorno = clAlterarContasReceber.AlteraNossoNumeroBcoContasAReceber(True, Con, Tran)
            '    If Not _Retorno.Sucesso Then
            '        blnRetorno = True
            '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            '        Exit Sub
            '    End If
            'End If

            Dim Arquivo As String() = Split(nomeArquivo, "\")
            Dim FileName = Arquivo(Arquivo.Length - 1)


            'VERIFICA SE O ARQUIVO JA FOI CHECADO PRA GRAVAR HISTORICO DO CLIENTE

            'INSERE HISTORICO DE CONTATO
            _Retorno = clInserirHistoricoContato.IncluiHistoricoContatoChecagem(sBanco, sNumAviso, sConta, sDataArq, "Verisure", FileName, Con, Tran)
            If Not _Retorno.Sucesso Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            End If

            'Douglas - 15/05/2019
            'Insere os mesmos dados na tabela nova
            _Retorno = clInserirTempCheckArqRetorno.InserirChecagemClientesBancariosLog(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "Verisure", Con, Tran)
            If Not _Retorno.Sucesso Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            End If
            'End If

            Tran.Commit()
            Con.Close()

            'VOLTA O CURSOR NO MOUSE
            'Me.Cursor = Cursors.Default
            '---------------------------------


            'Chama relatorio
            'MsgBox("Checagem concluída com sucesso (ver relatório)!")

            'Exibe Relatório
            'CarregaRelatorio()
            fsFile.Close()
            ValidateBankFiles(sBanco, sNumAviso, sConta, sDataArq, nomeArquivo)
        Catch ex As Exception
            Funcoes.CriaLog(Application.StartupPath & "\RelatorioErros.txt", ex.Message)
        End Try
    End Sub
    Public Function ObterBoletoBancario(ByVal cr_tb As ContaReceber) As BoletoNet.BoletoBancario
        Dim b_tb = New Boleto(cr_tb.CodBanco)
        Dim Teste = ""

        Select Case CType(cr_tb.CodBanco, BoletoNet.Enums.Bancos)
            'Case Enums.Bancos.BancoBrasil
            '    Return b_tb.BancodoBrasil(cr_tb)
            'Case Enums.Bancos.Banrisul
            '    Return b_tb.Banrisul(cr_tb)
            'Case Enums.Bancos.Bradesco
            '    Return b_tb.Bradesco(cr_tb)
            'Case Enums.Bancos.Caixa
            '    Return b_tb.Caixa(cr_tb)
            'Case Enums.Bancos.HSBC
            '    Return b_tb.HSBC(cr_tb)
            Case BoletoNet.Enums.Bancos.Itau
                Return b_tb.CodigoBarraItau(cr_tb)
                'Case Enums.Bancos.Santander
                '    Return b_tb.Santander(cr_tb)
                'Case Enums.Bancos.Semear
                '    Return b_tb.Semear(cr_tb)
            Case Else
                Throw New ArgumentException("Banco não implementado")
        End Select
    End Function
    Private Function LoadCodigoBarra(cr_tb As ContaReceber, codigoBarra As BoletoNet.CodigoBarra) As ContaReceber
        cr_tb.CodigoBarra = codigoBarra.LinhaDigitavel
        Return cr_tb
    End Function
    Private Sub Check041(nomeArquivo As String)
        Dim _TempCheckArqRetorno As New TempCheckArqRetorno
        Dim _Retorno As New Retorno

        Dim iFile As Integer
        Dim sBuffer As String
        Dim sNomeBanco As String, sBanco As String, sNumAviso As String, sConta As String, sDataPagto As Date, sNumtit As String, sNomeCliente As String
        Dim sAgencia As String, sCodIncons As String, sCodClie As String = "", sTamCodClie As Integer = 0, sSeqTit As String
        Dim sTotalRegistros As String = "", sCodOcor As String, sBancoDeb As String, sAgenciaDeb As String, sNumCtaDeb As String, dblValor As Double, sEndereco As String, sUF As String, sCidade As String, sCEP As String
        Dim iTamCodClie As Integer, sid As String = "", strNossoNumBco As String, sVcto As Date, sMsgErros As String = ""

        Dim sDataArq As String = ""


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------


        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            Exit Sub
        End If

        fsFile = New System.IO.StreamReader(nomeArquivo)

        sBuffer = fsFile.ReadLine

        If Mid(sBuffer, 82, 17) = "DEBITO AUTOMATICO" Then
            '****************DEBITO
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "SANTANDER ainda não liberado Checagem para Débito. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------

            Do While (sBuffer) <> Nothing

                'Checa se eh HEADER
                If Mid(sBuffer, 1, 1) = "A" Then

                    'Checa se eh RETORNO
                    If Mid(sBuffer, 2, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        Exit Sub
                    End If

                    sBanco = Mid(sBuffer, 43, 3)
                    sAgencia = strAgen
                    sConta = strContaCorrente

                    sDataArq = CDate(Mid(sBuffer, 66, 4) & "-" & Mid(sBuffer, 70, 2) & "-" & Mid(sBuffer, 72, 2))
                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)
                    sNomeBanco = Trim(Mid(sBuffer, 46, 20))

                    'If sBanco <> Trim(codigoBanco) Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                    '    Exit Sub
                    'End If

                    sBuffer = fsFile.ReadLine
                End If


                'Checa se eh DETALHE
                'Lucas - 21/06/2018
                'Adicionado tratativa para ignorar cadastro de optante por débito automático no mesmo arquivo de retorno
                If Trim(Mid(sBuffer, 1, 1)) = "F" AndAlso Trim(Mid(sBuffer, 150, 1)) = "0" Then

                    sCodOcor = "00"
                    sCodIncons = Mid(sBuffer, 68, 2)
                    sNomeCliente = ""
                    sBancoDeb = sBanco
                    sAgenciaDeb = Mid(sBuffer, 27, 4)
                    sNumCtaDeb = Mid(sBuffer, 34, 7) & "-" & Mid(sBuffer, 42, 1)
                    dblValor = Val(Mid(sBuffer, 53, 15)) / 100
                    sDataPagto = CDate(Trim(Mid(sBuffer, 45, 4)) & "-" & Trim(Mid(sBuffer, 49, 2)) & "-" & Trim(Mid(sBuffer, 51, 2)))
                    sEndereco = ""
                    sUF = ""
                    sCidade = ""
                    sCEP = ""

                    'sid = Mid(sBuffer, 86, 4) & "-" & Mid(sBuffer, 90, 2) & "-" & Mid(sBuffer, 92, 2)

                    sNumtit = Trim(Mid(sBuffer, 70, 8))
                    sSeqTit = Trim(Mid(sBuffer, 78, 2))

                    Dim TipoDupRetorno As List(Of ContaReceber) = clConsultacontasReceber.BuscaTipoDup(sNumtit, sSeqTit)
                    Dim tipoDup As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).TipoDup, " ")
                    Dim Situacao As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).Situacao, " ")

                    _TempCheckArqRetorno.CodBanco = sBanco
                    _TempCheckArqRetorno.NomeBanco = sNomeBanco
                    _TempCheckArqRetorno.NumAviso = sNumAviso
                    _TempCheckArqRetorno.CodAgen = sAgencia
                    _TempCheckArqRetorno.NumCta = sConta
                    _TempCheckArqRetorno.DtArq = sDataArq
                    _TempCheckArqRetorno.NumTit = sNumtit
                    _TempCheckArqRetorno.DtPagto = sDataPagto
                    _TempCheckArqRetorno.NomeCliente = sNomeCliente
                    _TempCheckArqRetorno.CodOcor = sCodOcor
                    _TempCheckArqRetorno.CodIncons = sCodIncons
                    _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                    _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                    _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                    _TempCheckArqRetorno.Valor = dblValor
                    _TempCheckArqRetorno.Endereco = sEndereco
                    _TempCheckArqRetorno.Cidade = sCidade
                    _TempCheckArqRetorno.UF = sUF
                    _TempCheckArqRetorno.Cep = sCEP
                    _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                    _TempCheckArqRetorno.SeqTit = sSeqTit
                    _TempCheckArqRetorno.TipoDup = tipoDup
                    _TempCheckArqRetorno.Situacao = Situacao

                    _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                        Exit Sub
                    End If

                End If

                'Checa se eh DETALHE DE PROCESSAMENTO
                'If Trim(Mid(sBuffer, 1, 1)) = "J" Then
                '    MsgBox("Registro de Confirmação de Processamento " & vbLf & vbLf &
                '            "NSA: " & Mid(sBuffer, 2, 6) & vbLf &
                '            "Data de geração do arquivo: " & Format(Mid(sBuffer, 8, 4) & "-" & Mid(sBuffer, 12, 2) & "-" & Mid(sBuffer, 14, 2), "dd/MM/yyyy") & vbLf &
                '            "Total de registros no arquivo: " & Mid(sBuffer, 16, 6) & vbLf &
                '            "Valor total do arquivo: " & CStr(CLng(Mid(sBuffer, 22, 17)) / 100) & vbLf &
                '            "Data de processamento do arquivo: " & Format(Mid(sBuffer, 39, 4) & "-" & Mid(sBuffer, 43, 2) & "-" & Mid(sBuffer, 45, 2), "dd/MM/yyyy"))
                'End If

                sBuffer = fsFile.ReadLine

            Loop
        Else

            'RadMessageBox.Show(Me.Text, "Caixa (boleto) ainda não liberado para faturamento. Consultar a área de sistemas!", MessageBoxButtons.OK, RadMessageIcon.Exclamation)
#Region "codigo comentado"
            'Exit Sub

            'sMsgErros = ""

            'Do While (sBuffer) <> Nothing

            '    'Checa se eh HEADER
            '    If Mid(sBuffer, 8, 1) = "0" Then

            '        'Checa se eh RETORNO
            '        If Mid(sBuffer, 143, 1) <> "2" Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            '            Exit Sub
            '        End If

            '        sBanco = Mid(sBuffer, 1, 3)
            '        sAgencia = Mid(sBuffer, 33, 4)
            '        sConta = Mid(sBuffer, 38, 9)
            '        sDataArq = CDate(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4))
            '        sNumAviso = Mid(sBuffer, 158, 6)
            '        sNomeBanco = Trim(Mid(sBuffer, 103, 30))


            '        If IIf(String.IsNullOrEmpty(sBanco), "", sBanco) <> Trim(codigoBanco) Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            '            Exit Sub
            '        End If


            '    End If


            '    'Checa se eh DETALHE
            '    If Trim(Mid(sBuffer, 8, 1)) = "3" And Trim(Mid(sBuffer, 14, 1)) = "T" Then

            '        sCodOcor = Trim(Mid(sBuffer, 16, 2))
            '        'Instrução Rejeitada ou Título Não Registrado
            '        If (sCodOcor = "26" Or sCodOcor = "17") And Strings.Left(Trim(Mid(sBuffer, 55, 15)), 1) = "" Then
            '            'As linhas abaixo deverão ser analisadas no final do processo
            '            sMsgErros = sMsgErros & IIf(Trim(sMsgErros) = "", "", vbCr) & "Nosso Número : " & Trim(Mid(sBuffer, 41, 13)) & "  Cód.Ocorrência : " & sCodOcor & " Valor : " & Format(dblValor, "C")
            '        Else

            '            sCodIncons = Trim(Mid(sBuffer, 209, 2))
            '            sNomeCliente = Trim(Mid(sBuffer, 144, 30))
            '            sBancoDeb = ""
            '            sAgenciaDeb = ""
            '            sNumCtaDeb = ""
            '            sEndereco = ""
            '            sUF = ""
            '            sCidade = ""
            '            sCEP = ""
            '            sNumtit = IIf(Trim(Mid(sBuffer, 55, 15)) = "", "000000000000000", Trim(Mid(sBuffer, 55, 15)))
            '            sSeqTit = Strings.Right(sNumtit, 2)
            '            sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
            '            sVcto = CDate(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4)))
            '            strNossoNumBco = Trim(Mid(sBuffer, 41, 13))

            '            If sCodOcor = "06" Or sCodOcor = "09" Or sCodOcor = "17" Then
            '                'Se for liquidacao pega valor pago e dt. pgto no registro seguinte
            '                sBuffer = fsFile.ReadLine
            '                dblValor = (Val(Mid(sBuffer, 78, 15)) / 100)
            '                sDataPagto = CDate(Mid(sBuffer, 142, 4) & "-" & Mid(sBuffer, 140, 2) & "-" & Mid(sBuffer, 138, 2))
            '            Else
            '                'Se for confirmacao de recebimento pega valor do titulo e dt de vcto
            '                dblValor = (Val(Mid(sBuffer, 78, 15)) / 100)
            '                sDataPagto = CDate(Mid(sBuffer, 74, 4) & "-" & Mid(sBuffer, 72, 2) & "-" & Mid(sBuffer, 70, 2))
            '            End If

            '            _TempCheckArqRetorno.CodBanco = sBanco
            '            _TempCheckArqRetorno.NomeBanco = sNomeBanco
            '            _TempCheckArqRetorno.NumAviso = Strings.Right(sNumAviso, 5)
            '            _TempCheckArqRetorno.CodAgen = sAgencia
            '            _TempCheckArqRetorno.NumCta = sConta
            '            _TempCheckArqRetorno.DtArq = sDataArq
            '            _TempCheckArqRetorno.NumTit = Strings.Right(sNumtit, 8)
            '            _TempCheckArqRetorno.SeqTit = sSeqTit
            '            _TempCheckArqRetorno.DtPagto = sDataPagto
            '            _TempCheckArqRetorno.NomeCliente = sNomeCliente
            '            _TempCheckArqRetorno.CodOcor = sCodOcor
            '            _TempCheckArqRetorno.CodIncons = sCodIncons
            '            _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
            '            _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
            '            _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
            '            _TempCheckArqRetorno.Valor = dblValor
            '            _TempCheckArqRetorno.Endereco = sEndereco
            '            _TempCheckArqRetorno.Cidade = sCidade
            '            _TempCheckArqRetorno.UF = sUF
            '            _TempCheckArqRetorno.Cep = sCEP
            '            _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
            '            _TempCheckArqRetorno.DtVcto = sVcto
            '            _TempCheckArqRetorno.NossoNumeroBco = strNossoNumBco

            '            _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno, connection, Transacao)
            '            If Not _Retorno.Sucesso Then
            '                blnRetorno = True
            '                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            '                Exit Sub
            '            End If

            '        End If
            '    End If

            '    sBuffer = fsFile.ReadLine
            'Loop
#End Region
        End If


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.Default
        '---------------------------------
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Connection.Open()
        Dim Transacao As SqlTransaction = Connection.BeginTransaction()

        'Atualiza  a tabela contas_a_receber com o nosso numero no banco
        'If MsgBox("Deseja gravar o Nosso Número ?", vbYesNo, Me.Text) = vbYes Then
        _Retorno = clAlterarContasReceber.AlteraNossoNumeroBcoContasAReceber(False, Connection, Transacao)
        '    If Not _Retorno.Sucesso Then
        '        blnRetorno = True
        '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

        '        Exit Sub
        '    End If
        'End If

        Dim Arquivo As String() = Split(nomeArquivo, "\")
        Dim FileName = Arquivo(Arquivo.Length - 1)

        'INSERE HISTORICO DE CONTATO
        _Retorno = clInserirHistoricoContato.IncluiHistoricoContatoChecagem(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "Verisure", FileName, Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        End If

        'Douglas - 15/05/2019
        'Insere os mesmos dados na tabela nova
        _Retorno = clInserirTempCheckArqRetorno.InserirChecagemClientesBancariosLog(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "Verisure", Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        End If

        'If sMsgErros <> "" Then
        '    RadMessageBox.Show(Me.Text, "Existe(m) a(s) seguinte(s) inconsistência(s) : " & vbCr & vbCr & sMsgErros & vbCr & "Problemas para checar Arquivo Retorno", MessageBoxButtons.OK, RadMessageIcon.Info)
        'End If

        Transacao.Commit()
        Connection.Close()

        'Chama relatorio
        'MsgBox("Checagem concluída com sucesso (ver relatório)!")

        'Exibe Relatório
        'CarregaRelatorio()
        fsFile.Close()
        ValidateBankFiles(sBanco, sNumAviso, sConta, sDataArq, nomeArquivo)
    End Sub

    Private Sub Check104(nomeArquivo As String) 'Check de Retorno do SANTANDER

        Dim _TempCheckArqRetorno As New TempCheckArqRetorno
        Dim _Retorno As New Retorno

        Dim iFile As Integer
        Dim sBuffer As String
        Dim sNomeBanco As String, sBanco As String, sNumAviso As String, sConta As String, sDataPagto As Date, sNumtit As String, sNomeCliente As String
        Dim sAgencia As String, sCodIncons As String, sCodClie As String = "", sTamCodClie As Integer = 0, sSeqTit As String
        Dim sTotalRegistros As String = "", sCodOcor As String, sBancoDeb As String, sAgenciaDeb As String, sNumCtaDeb As String, dblValor As Double, sEndereco As String, sUF As String, sCidade As String, sCEP As String
        Dim iTamCodClie As Integer, sid As String = "", strNossoNumBco As String, sVcto As Date, sMsgErros As String = ""

        Dim sDataArq As String = ""


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------


        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            Exit Sub
        End If

        fsFile = New System.IO.StreamReader(nomeArquivo)

        sBuffer = fsFile.ReadLine

        If Mid(sBuffer, 82, 17) = "DEBITO AUTOMATICO" Then
            '****************DEBITO
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "SANTANDER ainda não liberado Checagem para Débito. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------

            Do While (sBuffer) <> Nothing

                'Checa se eh HEADER
                If Mid(sBuffer, 1, 1) = "A" Then

                    'Checa se eh RETORNO
                    If Mid(sBuffer, 2, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        Exit Sub
                    End If

                    sBanco = Mid(sBuffer, 43, 3)
                    sAgencia = strAgen
                    sConta = strContaCorrente

                    sDataArq = CDate(Mid(sBuffer, 66, 4) & "-" & Mid(sBuffer, 70, 2) & "-" & Mid(sBuffer, 72, 2))
                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)
                    sNomeBanco = Trim(Mid(sBuffer, 46, 20))

                    'If sBanco <> Trim(codigoBanco) Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                    '    Exit Sub
                    'End If

                End If


                'Checa se eh DETALHE
                'Lucas - 21/06/2018
                'Adicionado tratativa para ignorar cadastro de optante por débito automático no mesmo arquivo de retorno
                If Trim(Mid(sBuffer, 1, 1)) = "F" AndAlso Trim(Mid(sBuffer, 150, 1)) = "0" Then

                    sCodOcor = "00"
                    sCodIncons = Mid(sBuffer, 68, 2)
                    sNomeCliente = ""
                    sBancoDeb = sBanco
                    sAgenciaDeb = Mid(sBuffer, 27, 4)
                    sNumCtaDeb = Mid(sBuffer, 34, 7) & "-" & Mid(sBuffer, 42, 1)
                    dblValor = Val(Mid(sBuffer, 53, 15)) / 100
                    sDataPagto = CDate(Trim(Mid(sBuffer, 45, 4)) & "-" & Trim(Mid(sBuffer, 49, 2)) & "-" & Trim(Mid(sBuffer, 51, 2)))
                    sEndereco = ""
                    sUF = ""
                    sCidade = ""
                    sCEP = ""

                    'sid = Mid(sBuffer, 86, 4) & "-" & Mid(sBuffer, 90, 2) & "-" & Mid(sBuffer, 92, 2)
                    Dim split As String()
                    split = sBuffer.Split("#")

                    Dim numtitSeqTit As String

                    numtitSeqTit = Trim(Mid(split(2), 1, 10))
                    sNumtit = Mid(numtitSeqTit, 1, 8)
                    sSeqTit = Mid(numtitSeqTit, 9, 2)

                    Dim TipoDupRetorno As List(Of ContaReceber) = clConsultacontasReceber.BuscaTipoDup(sNumtit, sSeqTit)
                    Dim tipoDup As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).TipoDup, " ")
                    Dim Situacao As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).Situacao, " ")

                    _TempCheckArqRetorno.CodBanco = sBanco
                    _TempCheckArqRetorno.NomeBanco = sNomeBanco
                    _TempCheckArqRetorno.NumAviso = sNumAviso
                    _TempCheckArqRetorno.CodAgen = sAgencia
                    _TempCheckArqRetorno.NumCta = sConta
                    _TempCheckArqRetorno.DtArq = sDataArq
                    _TempCheckArqRetorno.NumTit = sNumtit
                    _TempCheckArqRetorno.DtPagto = sDataPagto
                    _TempCheckArqRetorno.NomeCliente = sNomeCliente
                    _TempCheckArqRetorno.CodOcor = sCodOcor
                    _TempCheckArqRetorno.CodIncons = sCodIncons
                    _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                    _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                    _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                    _TempCheckArqRetorno.Valor = dblValor
                    _TempCheckArqRetorno.Endereco = sEndereco
                    _TempCheckArqRetorno.Cidade = sCidade
                    _TempCheckArqRetorno.UF = sUF
                    _TempCheckArqRetorno.Cep = sCEP
                    _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                    _TempCheckArqRetorno.SeqTit = sSeqTit
                    _TempCheckArqRetorno.TipoDup = tipoDup
                    _TempCheckArqRetorno.Situacao = Situacao

                    _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                        Exit Sub
                    End If

                End If

                'Checa se eh DETALHE DE PROCESSAMENTO
                'If Trim(Mid(sBuffer, 1, 1)) = "J" Then
                '    'MsgBox("Registro de Confirmação de Processamento " & vbLf & vbLf &
                '    '        "NSA: " & Mid(sBuffer, 2, 6) & vbLf &
                '    '        "Data de geração do arquivo: " & Format(Mid(sBuffer, 8, 4) & "-" & Mid(sBuffer, 12, 2) & "-" & Mid(sBuffer, 14, 2), "dd/MM/yyyy") & vbLf &
                '    '        "Total de registros no arquivo: " & Mid(sBuffer, 16, 6) & vbLf &
                '    '        "Valor total do arquivo: " & CStr(CLng(Mid(sBuffer, 22, 17)) / 100) & vbLf &
                '    '        "Data de processamento do arquivo: " & Format(Mid(sBuffer, 39, 4) & "-" & Mid(sBuffer, 43, 2) & "-" & Mid(sBuffer, 45, 2), "dd/MM/yyyy"))
                '    RadMessageBox.Show("Registro de Confirmação de Processamento " & vbLf & vbLf &
                '        "NSA: " & Mid(sBuffer, 2, 6) & vbLf &
                '        "Data de geração do arquivo: " & Format(Date.Parse(Mid(sBuffer, 8, 4) & "-" & Mid(sBuffer, 12, 2) & "-" & Mid(sBuffer, 14, 2)), "dd/MM/yyyy") & vbLf &
                '        "Total de registros no arquivo: " & Mid(sBuffer, 16, 6) & vbLf &
                '        "Valor total do arquivo: " & CStr(CLng(Mid(sBuffer, 22, 17)) / 100) & vbLf &
                '        "Data de processamento do arquivo: " & Format(Date.Parse(Mid(sBuffer, 39, 4) & "-" & Mid(sBuffer, 43, 2) & "-" & Mid(sBuffer, 45, 2)), "dd/MM/yyyy"), Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
                'End If

                sBuffer = fsFile.ReadLine

            Loop
        Else

            'RadMessageBox.Show(Me.Text, "Caixa (boleto) ainda não liberado para faturamento. Consultar a área de sistemas!", MessageBoxButtons.OK, RadMessageIcon.Exclamation)
#Region "codigo comentado"
            'Exit Sub

            'sMsgErros = ""

            'Do While (sBuffer) <> Nothing

            '    'Checa se eh HEADER
            '    If Mid(sBuffer, 8, 1) = "0" Then

            '        'Checa se eh RETORNO
            '        If Mid(sBuffer, 143, 1) <> "2" Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            '            Exit Sub
            '        End If

            '        sBanco = Mid(sBuffer, 1, 3)
            '        sAgencia = Mid(sBuffer, 33, 4)
            '        sConta = Mid(sBuffer, 38, 9)
            '        sDataArq = CDate(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4))
            '        sNumAviso = Mid(sBuffer, 158, 6)
            '        sNomeBanco = Trim(Mid(sBuffer, 103, 30))


            '        If IIf(String.IsNullOrEmpty(sBanco), "", sBanco) <> Trim(codigoBanco) Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            '            Exit Sub
            '        End If


            '    End If


            '    'Checa se eh DETALHE
            '    If Trim(Mid(sBuffer, 8, 1)) = "3" And Trim(Mid(sBuffer, 14, 1)) = "T" Then

            '        sCodOcor = Trim(Mid(sBuffer, 16, 2))
            '        'Instrução Rejeitada ou Título Não Registrado
            '        If (sCodOcor = "26" Or sCodOcor = "17") And Strings.Left(Trim(Mid(sBuffer, 55, 15)), 1) = "" Then
            '            'As linhas abaixo deverão ser analisadas no final do processo
            '            sMsgErros = sMsgErros & IIf(Trim(sMsgErros) = "", "", vbCr) & "Nosso Número : " & Trim(Mid(sBuffer, 41, 13)) & "  Cód.Ocorrência : " & sCodOcor & " Valor : " & Format(dblValor, "C")
            '        Else

            '            sCodIncons = Trim(Mid(sBuffer, 209, 2))
            '            sNomeCliente = Trim(Mid(sBuffer, 144, 30))
            '            sBancoDeb = ""
            '            sAgenciaDeb = ""
            '            sNumCtaDeb = ""
            '            sEndereco = ""
            '            sUF = ""
            '            sCidade = ""
            '            sCEP = ""
            '            sNumtit = IIf(Trim(Mid(sBuffer, 55, 15)) = "", "000000000000000", Trim(Mid(sBuffer, 55, 15)))
            '            sSeqTit = Strings.Right(sNumtit, 2)
            '            sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
            '            sVcto = CDate(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4)))
            '            strNossoNumBco = Trim(Mid(sBuffer, 41, 13))

            '            If sCodOcor = "06" Or sCodOcor = "09" Or sCodOcor = "17" Then
            '                'Se for liquidacao pega valor pago e dt. pgto no registro seguinte
            '                sBuffer = fsFile.ReadLine
            '                dblValor = (Val(Mid(sBuffer, 78, 15)) / 100)
            '                sDataPagto = CDate(Mid(sBuffer, 142, 4) & "-" & Mid(sBuffer, 140, 2) & "-" & Mid(sBuffer, 138, 2))
            '            Else
            '                'Se for confirmacao de recebimento pega valor do titulo e dt de vcto
            '                dblValor = (Val(Mid(sBuffer, 78, 15)) / 100)
            '                sDataPagto = CDate(Mid(sBuffer, 74, 4) & "-" & Mid(sBuffer, 72, 2) & "-" & Mid(sBuffer, 70, 2))
            '            End If

            '            _TempCheckArqRetorno.CodBanco = sBanco
            '            _TempCheckArqRetorno.NomeBanco = sNomeBanco
            '            _TempCheckArqRetorno.NumAviso = Strings.Right(sNumAviso, 5)
            '            _TempCheckArqRetorno.CodAgen = sAgencia
            '            _TempCheckArqRetorno.NumCta = sConta
            '            _TempCheckArqRetorno.DtArq = sDataArq
            '            _TempCheckArqRetorno.NumTit = Strings.Right(sNumtit, 8)
            '            _TempCheckArqRetorno.SeqTit = sSeqTit
            '            _TempCheckArqRetorno.DtPagto = sDataPagto
            '            _TempCheckArqRetorno.NomeCliente = sNomeCliente
            '            _TempCheckArqRetorno.CodOcor = sCodOcor
            '            _TempCheckArqRetorno.CodIncons = sCodIncons
            '            _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
            '            _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
            '            _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
            '            _TempCheckArqRetorno.Valor = dblValor
            '            _TempCheckArqRetorno.Endereco = sEndereco
            '            _TempCheckArqRetorno.Cidade = sCidade
            '            _TempCheckArqRetorno.UF = sUF
            '            _TempCheckArqRetorno.Cep = sCEP
            '            _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
            '            _TempCheckArqRetorno.DtVcto = sVcto
            '            _TempCheckArqRetorno.NossoNumeroBco = strNossoNumBco

            '            _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno, connection, Transacao)
            '            If Not _Retorno.Sucesso Then
            '                blnRetorno = True
            '                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            '                Exit Sub
            '            End If

            '        End If
            '    End If

            '    sBuffer = fsFile.ReadLine
            'Loop
#End Region
        End If


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.Default
        '---------------------------------
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Connection.Open()
        Dim Transacao As SqlTransaction = Connection.BeginTransaction()

        'Atualiza  a tabela contas_a_receber com o nosso numero no banco
        'If MsgBox("Deseja gravar o Nosso Número ?", vbYesNo, Me.Text) = vbYes Then
        _Retorno = clAlterarContasReceber.AlteraNossoNumeroBcoContasAReceber(False, Connection, Transacao)
        '    If Not _Retorno.Sucesso Then
        '        blnRetorno = True
        '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

        '        Exit Sub
        '    End If
        'End If

        Dim Arquivo As String() = Split(nomeArquivo, "\")
        Dim FileName = Arquivo(Arquivo.Length - 1)


        'INSERE HISTORICO DE CONTATO
        _Retorno = clInserirHistoricoContato.IncluiHistoricoContatoChecagem(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "Verisure", FileName, Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        End If
        'Douglas - 15/05/2019
        'Insere os mesmos dados na tabela nova
        _Retorno = clInserirTempCheckArqRetorno.InserirChecagemClientesBancariosLog(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "Verisure", Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        End If

        'If sMsgErros <> "" Then
        '    'MsgBox("Existe(m) a(s) seguinte(s) inconsistência(s) : " & vbCr & vbCr & sMsgErros & vbCr, vbExclamation, "Problemas para checar Arquivo Retorno")
        '    RadMessageBox.Show("Existe(m) a(s) seguinte(s) inconsistência(s) : " & vbCr & vbCr & sMsgErros & vbCr & "Problemas para checar Arquivo Retorno", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
        'End If

        Transacao.Commit()
        Connection.Close()

        'Chama relatorio
        'MsgBox("Checagem concluída com sucesso (ver relatório)!")

        'Exibe Relatório
        'CarregaRelatorio()
        ValidateBankFiles(sBanco, sNumAviso, sConta, sDataArq, nomeArquivo)
    End Sub

    Private Sub Check237(nomeArquivo As String) 'Check de Retorno do BRADESCO

        'SOMENTE PARA GRAVAÇÃO DO NOSSO NÚMERO
        'A CHECAGEM COMPLETA EH FEITA PELO SITE DO BRADESCO

        Dim _TempCheckArqRetorno As New TempCheckArqRetorno
        Dim _Retorno As New Retorno

        Dim iFile As Integer
        Dim sBuffer As String
        Dim sNomeBanco As String, sBanco As String, sNumAviso As String, sConta As String, sDataPagto As String, sNumtit As String, sNomeCliente As String
        Dim sAgencia As String, sCodIncons As String, sCodClie As String, sTamCodClie As Integer, sSeqTit As String
        Dim sTotalRegistros As String, sCodOcor As String, sBancoDeb As String, sAgenciaDeb As String, sNumCtaDeb As String, dblValor As Double, sEndereco As String, sUF As String, sCidade As String, sCEP As String
        Dim iTamCodClie As Integer, sid As String, strNossoNumBco As String, sVcto As Date
        Dim sDataArq As Date


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------


        iFile = FreeFile()

        'If Dir(nomeArquivo) = "" Then
        '    blnRetorno = True
        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
        '    Exit Sub
        'End If
        fsFile = New System.IO.StreamReader(nomeArquivo)

        sBuffer = fsFile.ReadLine

        Do While (sBuffer) <> Nothing

            'Checa se eh HEADER
            If Mid(sBuffer, 1, 1) = "0" Then
                'Checa se eh RETORNO
                If Mid(sBuffer, 1, 1) = "0" Then
                    If Mid(sBuffer, 2, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        Exit Sub
                    End If
                End If

                sBanco = Trim(Mid(sBuffer, 77, 3))
                sDataArq = Convert.ToDateTime(Mid(sBuffer, 95, 2) & "/" & Mid(sBuffer, 97, 2) & "/" & Mid(sBuffer, 99, 2), Funcoes.Cultura)
                sNumAviso = Mid(sBuffer, 109, 5)
                sNomeBanco = Trim(Mid(sBuffer, 80, 15))

                'If sBanco <> Trim(codigoBanco) Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                '    Exit Sub
                'End If

            End If

            'Checa se eh DETALHE
            If Trim(Mid(sBuffer, 1, 1)) = "1" Then

                sCodOcor = Trim(Mid(sBuffer, 109, 2))
                sCodIncons = Trim(Mid(sBuffer, 319, 10))
                sAgencia = Trim(Mid(sBuffer, 25, 5))
                sConta = Trim(Str(Val(Mid(sBuffer, 30, 7)))) & "-" & Trim(Mid(sBuffer, 37, 1))

                sNomeCliente = ""
                sBancoDeb = Trim(Mid(sBuffer, 166, 3))
                sAgenciaDeb = Trim(Mid(sBuffer, 169, 5))
                sNumCtaDeb = ""
                sEndereco = ""
                sUF = ""
                sCidade = ""
                sCEP = ""
                dblValor = getValor(sCodOcor, sBuffer)
                'If sCodOcor = "02" Then
                If CStr(Trim(Mid(sBuffer, 117, 10))) <> "" Then
                    sNumtit = Strings.Left(Trim(Mid(sBuffer, 117, 10)), IIf(Len(Trim(Mid(sBuffer, 117, 10))) - 2 <= 0, 0, Len(Trim(Mid(sBuffer, 117, 10))) - 2))
                End If
                sSeqTit = Strings.Right(Trim(Mid(sBuffer, 117, 10)), 2)
                If Trim(Mid(sBuffer, 147, 2)) <> "00" Then
                    sVcto = CDate(Trim(Mid(sBuffer, 147, 2)) & "/" & Trim(Mid(sBuffer, 149, 2)) & "/20" & Trim(Mid(sBuffer, 151, 2)))
                    'sVcto = CDate(Trim(Mid(sBuffer, 149, 2)) & "/" & Trim(Mid(sBuffer, 147, 2)) & "/20" & Trim(Mid(sBuffer, 151, 2)))

                    'medida provisória
                ElseIf Trim(Mid(sBuffer, 147, 2)) = "00" Then
                    sVcto = Funcoes.PegaData()
                    sNumtit = ""
                End If

                strNossoNumBco = Trim(Mid(sBuffer, 71, 12))


                If (Not IsDate(sDataArq)) Then
                    Throw New Exception("Check237 - Data inválida: " + sDataArq)
                End If

                _TempCheckArqRetorno.CodBanco = sBanco
                _TempCheckArqRetorno.NomeBanco = sNomeBanco
                _TempCheckArqRetorno.NumAviso = sNumAviso
                _TempCheckArqRetorno.CodAgen = sAgencia
                _TempCheckArqRetorno.NumCta = sConta
                _TempCheckArqRetorno.DtArq = Convert.ToDateTime(sDataArq, Funcoes.Cultura)
                _TempCheckArqRetorno.NumTit = sNumtit
                _TempCheckArqRetorno.SeqTit = sSeqTit
                _TempCheckArqRetorno.DtPagto = Convert.ToDateTime(sDataArq, Funcoes.Cultura)
                _TempCheckArqRetorno.NomeCliente = sNomeCliente
                _TempCheckArqRetorno.CodOcor = Mid(sCodOcor, 1, 2)
                _TempCheckArqRetorno.CodIncons = Mid(sCodIncons, 1, 2)
                _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                _TempCheckArqRetorno.Valor = dblValor
                _TempCheckArqRetorno.Endereco = sEndereco
                _TempCheckArqRetorno.Cidade = sCidade
                _TempCheckArqRetorno.UF = sUF
                _TempCheckArqRetorno.Cep = sCEP
                _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                _TempCheckArqRetorno.DtVcto = sVcto
                _TempCheckArqRetorno.NossoNumeroBco = strNossoNumBco
                Dim TipoDupRetorno As List(Of ContaReceber) = clConsultacontasReceber.BuscaTipoDup(sNumtit, sSeqTit)
                Dim tipoDup As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).TipoDup, " ")
                Dim Situacao As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).Situacao, " ")
                _TempCheckArqRetorno.TipoDup = tipoDup
                _TempCheckArqRetorno.Situacao = Situacao

                _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                If Not _Retorno.Sucesso Then
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                    Exit Sub
                End If

                'and If

                Debug.Print(sNumtit & "/" & sCodOcor & "/" & strNossoNumBco & "/" & sVcto)

            End If
            sBuffer = fsFile.ReadLine
        Loop


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.Default
        '---------------------------------
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Connection.Open()
        Dim Transacao As SqlTransaction = Connection.BeginTransaction()

        'Atualiza  a tabela contas_a_receber com o nosso numero no banco e se pagamento será por DDA
        'If MsgBox("Deseja gravar o Nosso Número e o aviso de DDA ?", vbYesNo, Me.Text) = vbYes Then
        _Retorno = clAlterarContasReceber.AlteraNossoNumeroBcoContasAReceber(True, Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            Exit Sub
        End If
        'End If

        Dim Arquivo As String() = Split(nomeArquivo, "\")
        Dim FileName = Arquivo(Arquivo.Length - 1)

        'VERIFICA SE O ARQUIVO JA FOI CHECADO PRA GRAVAR HISTORICO DO CLIENTE
        'lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoChecagemPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sConta, Connection, Transacao)
        ' If lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
        'INSERE HISTORICO DE CONTATO
        _Retorno = clInserirHistoricoContato.IncluiHistoricoContatoChecagem(sBanco, sNumAviso, sConta, sDataArq, "VERISURE", FileName, Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        End If
        'Douglas - 15/05/2019
        'Insere os mesmos dados na tabela nova
        _Retorno = clInserirTempCheckArqRetorno.InserirChecagemClientesBancariosLog(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "VERISURE", Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        End If
        'End If

        Transacao.Commit()
        Connection.Close()

        'MsgBox("Checagem concluída com sucesso !")

        'Exibe Relatório
        'CarregaRelatorio()

        'Chama relatorio
        'MsgBox "Checagem concluída com sucesso (ver relatório)!"
        'ShowReport
        fsFile.Close()
        ValidateBankFiles(sBanco, sNumAviso, sConta, sDataArq, nomeArquivo)
    End Sub
    Private Function getValor(ByVal CodOcor As String, ByVal Buffer As String) As Double

        'Select Case CodOcor
        '    Case "02"
        '        Return Val(Mid(Buffer, 153, 13)) / 100
        '    Case Else
        '        Return Val(Mid(Buffer, 254, 13)) / 100
        'End Select

        If Val(Mid(Buffer, 254, 13)) / 100 > 0 Then
            Return Val(Mid(Buffer, 254, 13)) / 100
        Else
            Return Val(Mid(Buffer, 153, 13)) / 100
        End If


    End Function
    Private Sub Check033(nomeArquivo As String) 'Check de Retorno do SANTANDER

        Dim _TempCheckArqRetorno As New TempCheckArqRetorno
        Dim _Retorno As New Retorno

        Dim iFile As Integer
        Dim sBuffer As String
        Dim sNomeBanco As String, sBanco As String, sNumAviso As String, sConta As String, sDataPagto As Date, sNumtit As String, sNomeCliente As String
        Dim sAgencia As String, sCodIncons As String, sCodClie As String = "", sTamCodClie As Integer = 0, sSeqTit As String
        Dim sTotalRegistros As String = "", sCodOcor As String, sBancoDeb As String, sAgenciaDeb As String, sNumCtaDeb As String, dblValor As Double, sEndereco As String, sUF As String, sCidade As String, sCEP As String
        Dim iTamCodClie As Integer, sid As String = "", strNossoNumBco As String, sVcto As Date, sMsgErros As String = ""

        Dim sDataArq As String = ""


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------


        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            Exit Sub
        End If

        fsFile = New System.IO.StreamReader(nomeArquivo)

        sBuffer = fsFile.ReadLine

        If Mid(sBuffer, 82, 17) = "DEBITO AUTOMATICO" Then
            '****************DEBITO
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "SANTANDER ainda não liberado Checagem para Débito. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------

            Do While (sBuffer) <> Nothing

                'Checa se eh HEADER
                If Mid(sBuffer, 1, 1) = "A" Then

                    'Checa se eh RETORNO
                    If Mid(sBuffer, 2, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        Exit Sub
                    End If

                    sBanco = Mid(sBuffer, 43, 3)
                    'sAgencia = "0319"
                    'sConta = "13007469-5"
                    sAgencia = strAgen
                    sConta = strContaCorrente

                    sDataArq = CDate(Mid(sBuffer, 66, 4) & "-" & Mid(sBuffer, 70, 2) & "-" & Mid(sBuffer, 72, 2))
                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)
                    sNomeBanco = Trim(Mid(sBuffer, 46, 20))

                    'If sBanco <> Trim(codigoBanco) Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                    '    Exit Sub
                    'End If

                End If

                'Lucas - 13/11/2018
                'Verifica se o registro está de acordo com novo layout
                If String.IsNullOrEmpty(Trim(Mid(sBuffer, 120, 8)) + Trim(Mid(sBuffer, 128, 2))) Then
                    sBuffer = fsFile.ReadLine
                    Continue Do
                End If

                'Checa se eh DETALHE
                If Trim(Mid(sBuffer, 1, 1)) = "F" Then

                    sCodOcor = "00"
                    sCodIncons = Mid(sBuffer, 68, 2)
                    sNomeCliente = ""
                    sBancoDeb = sBanco
                    sAgenciaDeb = Mid(sBuffer, 27, 4)
                    sNumCtaDeb = Mid(sBuffer, 31, 8) & "-" & Mid(sBuffer, 39, 1)
                    dblValor = Val(Mid(sBuffer, 53, 15)) / 100
                    sDataPagto = CDate(Trim(Mid(sBuffer, 45, 4)) & "-" & Trim(Mid(sBuffer, 49, 2)) & "-" & Trim(Mid(sBuffer, 51, 2)))
                    sEndereco = ""
                    sUF = ""
                    sCidade = ""
                    sCEP = ""

                    'sid = Mid(sBuffer, 86, 4) & "-" & Mid(sBuffer, 90, 2) & "-" & Mid(sBuffer, 92, 2)

                    'Lucas - 17/10/2018
                    'Alterado posições para ler o numtit e seqtit do cliente de acordo com novo layout
                    'sNumtit = Strings.Left(Trim(Mid(sBuffer, 2, 25)), Len(Trim(Mid(sBuffer, 2, 25))) - 2)
                    'sSeqTit = Strings.Right(Trim(Mid(sBuffer, 2, 25)), 2)
                    sNumtit = Trim(Mid(sBuffer, 120, 8))
                    sSeqTit = Trim(Mid(sBuffer, 128, 2))
                    Dim TipoDupRetorno As List(Of ContaReceber) = clConsultacontasReceber.BuscaTipoDup(sNumtit, sSeqTit)
                    Dim tipoDup As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).TipoDup, " ")
                    Dim Situacao As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).Situacao, " ")

                    _TempCheckArqRetorno.CodBanco = sBanco
                    _TempCheckArqRetorno.NomeBanco = sNomeBanco
                    _TempCheckArqRetorno.NumAviso = sNumAviso
                    _TempCheckArqRetorno.CodAgen = sAgencia
                    _TempCheckArqRetorno.NumCta = sConta
                    _TempCheckArqRetorno.DtArq = sDataArq
                    _TempCheckArqRetorno.NumTit = sNumtit
                    _TempCheckArqRetorno.DtPagto = sDataPagto
                    _TempCheckArqRetorno.NomeCliente = sNomeCliente
                    _TempCheckArqRetorno.CodOcor = sCodOcor
                    _TempCheckArqRetorno.CodIncons = sCodIncons
                    _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                    _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                    _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                    _TempCheckArqRetorno.Valor = dblValor
                    _TempCheckArqRetorno.Endereco = sEndereco
                    _TempCheckArqRetorno.Cidade = sCidade
                    _TempCheckArqRetorno.UF = sUF
                    _TempCheckArqRetorno.Cep = sCEP
                    _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                    _TempCheckArqRetorno.SeqTit = sSeqTit
                    _TempCheckArqRetorno.TipoDup = tipoDup
                    _TempCheckArqRetorno.Situacao = Situacao

                    _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                        Exit Sub
                    End If

                End If

                'Checa se eh DETALHE DE PROCESSAMENTO
                'If Trim(Mid(sBuffer, 1, 1)) = "J" Then
                '    RadMessageBox.Show("Registro de Confirmação de Processamento " & vbLf & vbLf &
                '                        "NSA: " & Mid(sBuffer, 2, 6) & vbLf &
                '                        "Data de geração do arquivo: " & Format(Date.Parse(Mid(sBuffer, 8, 4) & "-" & Mid(sBuffer, 12, 2) & "-" & Mid(sBuffer, 14, 2)), "dd/MM/yyyy") & vbLf &
                '                        "Total de registros no arquivo: " & Mid(sBuffer, 16, 6) & vbLf &
                '                        "Valor total do arquivo: " & CStr(CLng(Mid(sBuffer, 22, 17)) / 100) & vbLf &
                '                        "Data de processamento do arquivo: " & Format(Date.Parse(Mid(sBuffer, 39, 4) & "-" & Mid(sBuffer, 43, 2) & "-" & Mid(sBuffer, 45, 2)), "dd/MM/yyyy"), Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
                'End If

                sBuffer = fsFile.ReadLine

            Loop
        Else

            '*************BOLETO
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "SANTANDER ainda não liberado Checagem de Boleto. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------

            sMsgErros = ""

            Do While (sBuffer) <> Nothing

                'Checa se eh HEADER
                If Mid(sBuffer, 8, 1) = "0" Then

                    'Checa se eh RETORNO
                    If Mid(sBuffer, 143, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        Exit Sub
                    End If

                    sBanco = Mid(sBuffer, 1, 3)
                    sAgencia = Mid(sBuffer, 33, 4)
                    sConta = Mid(sBuffer, 38, 9)
                    sDataArq = CDate(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4))
                    sNumAviso = Mid(sBuffer, 158, 6)
                    sNomeBanco = Trim(Mid(sBuffer, 103, 30))


                    'If IIf(String.IsNullOrEmpty(sBanco), "", sBanco) <> Trim(codigoBanco) Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                    '    Exit Sub
                    'End If


                End If


                'Checa se eh DETALHE
                If Trim(Mid(sBuffer, 8, 1)) = "3" And Trim(Mid(sBuffer, 14, 1)) = "T" Then

                    sCodOcor = Trim(Mid(sBuffer, 16, 2))
                    'Instrução Rejeitada ou Título Não Registrado
                    If (sCodOcor = "26" Or sCodOcor = "17") And Strings.Left(Trim(Mid(sBuffer, 55, 15)), 1) = "" Then
                        'As linhas abaixo deverão ser analisadas no final do processo
                        sMsgErros = sMsgErros & IIf(Trim(sMsgErros) = "", "", vbCr) & "Nosso Número : " & Trim(Mid(sBuffer, 41, 13)) & "  Cód.Ocorrência : " & sCodOcor & " Valor : " & Format(dblValor, "C")
                    Else

                        sCodIncons = Trim(Mid(sBuffer, 209, 2))
                        sNomeCliente = Trim(Mid(sBuffer, 144, 30))
                        sBancoDeb = ""
                        sAgenciaDeb = ""
                        sNumCtaDeb = ""
                        sEndereco = ""
                        sUF = ""
                        sCidade = ""
                        sCEP = ""
                        sNumtit = IIf(Trim(Mid(sBuffer, 55, 15)) = "", "000000000000000", Trim(Mid(sBuffer, 55, 15)))
                        sSeqTit = Strings.Right(sNumtit, 2)
                        sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
                        sVcto = CDate(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4)))
                        strNossoNumBco = Trim(Mid(sBuffer, 41, 13))

                        If sCodOcor = "06" Or sCodOcor = "09" Or sCodOcor = "17" Then
                            'Se for liquidacao pega valor pago e dt. pgto no registro seguinte
                            sBuffer = fsFile.ReadLine
                            dblValor = (Val(Mid(sBuffer, 78, 15)) / 100)
                            sDataPagto = CDate(Mid(sBuffer, 142, 4) & "-" & Mid(sBuffer, 140, 2) & "-" & Mid(sBuffer, 138, 2))
                        Else
                            'Se for confirmacao de recebimento pega valor do titulo e dt de vcto
                            dblValor = (Val(Mid(sBuffer, 78, 15)) / 100)
                            sDataPagto = CDate(Mid(sBuffer, 74, 4) & "-" & Mid(sBuffer, 72, 2) & "-" & Mid(sBuffer, 70, 2))
                        End If

                        _TempCheckArqRetorno.CodBanco = sBanco
                        _TempCheckArqRetorno.NomeBanco = sNomeBanco
                        _TempCheckArqRetorno.NumAviso = Strings.Right(sNumAviso, 5)
                        _TempCheckArqRetorno.CodAgen = sAgencia
                        _TempCheckArqRetorno.NumCta = sConta
                        _TempCheckArqRetorno.DtArq = sDataArq
                        _TempCheckArqRetorno.NumTit = Strings.Right(sNumtit, 8)
                        _TempCheckArqRetorno.SeqTit = sSeqTit
                        _TempCheckArqRetorno.DtPagto = sDataPagto
                        _TempCheckArqRetorno.NomeCliente = sNomeCliente
                        _TempCheckArqRetorno.CodOcor = sCodOcor
                        _TempCheckArqRetorno.CodIncons = sCodIncons
                        _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                        _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                        _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                        _TempCheckArqRetorno.Valor = dblValor
                        _TempCheckArqRetorno.Endereco = sEndereco
                        _TempCheckArqRetorno.Cidade = sCidade
                        _TempCheckArqRetorno.UF = sUF
                        _TempCheckArqRetorno.Cep = sCEP
                        _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                        _TempCheckArqRetorno.DtVcto = sVcto
                        _TempCheckArqRetorno.NossoNumeroBco = strNossoNumBco

                        _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                        If Not _Retorno.Sucesso Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                            Exit Sub
                        End If

                    End If
                End If

                sBuffer = fsFile.ReadLine
            Loop

        End If


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.Default
        '---------------------------------
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Connection.Open()
        Dim Transacao As SqlTransaction = Connection.BeginTransaction()

        'Atualiza  a tabela contas_a_receber com o nosso numero no banco
        'If MsgBox("Deseja gravar o Nosso Número ?", vbYesNo, Me.Text) = vbYes Then
        _Retorno = clAlterarContasReceber.AlteraNossoNumeroBcoContasAReceber(False, Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            Exit Sub
        End If
        'End If

        Dim Arquivo As String() = Split(nomeArquivo, "\")
        Dim FileName = Arquivo(Arquivo.Length - 1)

        'VERIFICA SE O ARQUIVO JA FOI CHECADO PRA GRAVAR HISTORICO DO CLIENTE
        'INSERE HISTORICO DE CONTATO
        _Retorno = clInserirHistoricoContato.IncluiHistoricoContatoChecagem(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "VERISURE", FileName, Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        End If
        'Douglas - 15/05/2019
        'Insere os mesmos dados na tabela nova
        _Retorno = clInserirTempCheckArqRetorno.InserirChecagemClientesBancariosLog(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "VERISURE", Connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        End If
        'End If

        Transacao.Commit()
        Connection.Close()

        'Chama relatorio
        'MsgBox("Checagem concluída com sucesso (ver relatório)!")

        'Exibe Relatório
        'CarregaRelatorio()
        fsFile.Close()
        ValidateBankFiles(sBanco, sNumAviso, sConta, sDataArq, nomeArquivo)
    End Sub

    Private Sub Check001(nomeArquivo As String) 'Check de Retorno do BANCO DO BRASIL

        Dim _TempCheckArqRetorno As New TempCheckArqRetorno
        Dim _Retorno As New Retorno
        '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
        'MsgBox "BANCO DO BRASIL ainda não liberado para checagem de retono. Consultar analista de sistemas !"
        'Exit Sub
        '---------------------------------------------------------------

        Dim iFile As Integer
        Dim sBuffer As String

        Dim sNomeBanco As String, sBanco As String, sNumAviso As String, sConta As String, sDataPagto As String, sNumtit As String, sNomeCliente As String
        Dim sAgencia As String, sCodIncons As String, sCodClie As String = "", sTamCodClie As Integer = 0, sSeqTit As String
        Dim sTotalRegistros As String = "", sCodOcor As String, sBancoDeb As String, sAgenciaDeb As String, sNumCtaDeb As String, dblValor As Double, sEndereco As String, sUF As String, sCidade As String, sCEP As String
        Dim iTamCodClie As Integer = 0, sid As String = ""
        Dim strNossoNumBco As String, sVcto As String
        Dim sDataArq As Date

        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            Exit Sub
        End If

        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------


        fsFile = New System.IO.StreamReader(nomeArquivo)

        sBuffer = fsFile.ReadLine

        If Mid(sBuffer, 82, 17) = "DEBITO AUTOMATICO" Then
            '****************DEBITO
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "BANCO DO BRASIL ainda não liberado Checagem para Débito. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------


            Do While (sBuffer) <> Nothing

                'Checa se eh HEADER
                If Mid(sBuffer, 1, 1) = "A" Then

                    'Checa se eh RETORNO
                    If Mid(sBuffer, 2, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        Exit Sub
                    End If

                    sBanco = Mid(sBuffer, 43, 3)
                    'sAgencia = "2962-9"
                    'sConta = "691933-2"                    
                    sAgencia = strAgen
                    sConta = strContaCorrente

                    sDataArq = CDate(Mid(sBuffer, 66, 4) & "-" & Mid(sBuffer, 70, 2) & "-" & Mid(sBuffer, 72, 2))
                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)
                    sNomeBanco = Trim(Mid(sBuffer, 46, 20))

                    'If sBanco <> Trim(codigoBanco) Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                    '    Exit Sub
                    'End If

                End If


                'Checa se eh DETALHE DE DEB AUTOMATICO
                If Trim(Mid(sBuffer, 1, 1)) = "F" Then

                    sCodOcor = "00"
                    sCodIncons = Mid(sBuffer, 68, 2)
                    sNomeCliente = ""
                    sNumCtaDeb = Mid(sBuffer, 31, 14)
                    dblValor = Val(Mid(sBuffer, 53, 15)) / 100
                    sDataPagto = Trim(Mid(sBuffer, 45, 4)) & "-" & Trim(Mid(sBuffer, 49, 2)) & "-" & Trim(Mid(sBuffer, 51, 2))
                    sEndereco = ""
                    sUF = ""
                    sCidade = ""
                    sCEP = ""

                    'sid = Mid(sBuffer, 86, 4) & "-" & Mid(sBuffer, 90, 2) & "-" & Mid(sBuffer, 92, 2)

                    'sNumtit = Strings.Left(Trim(Mid(sBuffer, 2, 25)), Len(Trim(Mid(sBuffer, 2, 25))) - 2)
                    'sSeqTit = Strings.Right(Trim(Mid(sBuffer, 2, 25)), 2)

                    ''Liberar quenado Banco do Brasil for testado
                    sNumtit = Strings.Left(Trim(Mid(sBuffer, 70, 25)), Len(Trim(Mid(sBuffer, 70, 25))) - 2)
                    sSeqTit = Strings.Right(Trim(Mid(sBuffer, 70, 25)), 2)
                    Dim TipoDupRetorno As List(Of ContaReceber) = clConsultacontasReceber.BuscaTipoDup(sNumtit, sSeqTit)
                    Dim tipoDup As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).TipoDup, " ")
                    Dim Situacao As String = If(TipoDupRetorno.Count > 0 AndAlso TipoDupRetorno(0).Sucesso, TipoDupRetorno(0).Situacao, " ")


                    _TempCheckArqRetorno.CodBanco = sBanco
                    _TempCheckArqRetorno.NomeBanco = sNomeBanco
                    _TempCheckArqRetorno.NumAviso = sNumAviso
                    _TempCheckArqRetorno.CodAgen = sAgencia
                    _TempCheckArqRetorno.NumCta = sConta
                    _TempCheckArqRetorno.DtArq = sDataArq
                    _TempCheckArqRetorno.NumTit = sNumtit
                    _TempCheckArqRetorno.DtPagto = sDataPagto
                    _TempCheckArqRetorno.NomeCliente = sNomeCliente
                    _TempCheckArqRetorno.CodOcor = sCodOcor
                    _TempCheckArqRetorno.CodIncons = sCodIncons
                    _TempCheckArqRetorno.Mensagem = IIf(sCodOcor = "09", "Baixa", IIf(sCodOcor = "50", "Título pago com cheque, pendente de compensação", "")) 'Informar msg, conforme Mary 27/08/2010
                    _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                    _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                    _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                    _TempCheckArqRetorno.Valor = dblValor
                    _TempCheckArqRetorno.Endereco = sEndereco
                    _TempCheckArqRetorno.Cidade = sCidade
                    _TempCheckArqRetorno.UF = sUF
                    _TempCheckArqRetorno.Cep = sCEP
                    _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                    _TempCheckArqRetorno.SeqTit = sSeqTit
                    _TempCheckArqRetorno.TipoDup = tipoDup
                    _TempCheckArqRetorno.Situacao = Situacao


                    _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                        Exit Sub
                    End If

                End If

                ''Checa se eh DETALHE DE PROCESSAMENTO
                'If Trim(Mid(sBuffer, 1, 1)) = "J" Then
                '    RadMessageBox.Show("Registro de Confirmação de Processamento " & vbLf & vbLf &
                '                        "NSA: " & Mid(sBuffer, 2, 6) & vbLf &
                '                        "Data de geração do arquivo: " & Format(Date.Parse(Mid(sBuffer, 8, 4) & "-" & Mid(sBuffer, 12, 2) & "-" & Mid(sBuffer, 14, 2)), "dd/MM/yyyy") & vbLf &
                '                        "Total de registros no arquivo: " & Mid(sBuffer, 16, 6) & vbLf &
                '                        "Valor total do arquivo: " & CStr(CLng(Mid(sBuffer, 22, 17)) / 100) & vbLf &
                '                        "Data de processamento do arquivo: " & Format(Date.Parse(Mid(sBuffer, 39, 4) & "-" & Mid(sBuffer, 43, 2) & "-" & Mid(sBuffer, 45, 2)), "dd/MM/yyyy"), Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
                'End If

                sBuffer = fsFile.ReadLine
            Loop

        Else

            '*************BOLETO
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "BANCO DO BRASIL ainda não liberado Checagem de Boleto. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------
            Do While (sBuffer) <> Nothing

                'Checa se eh HEADER
                If Mid(sBuffer, 8, 1) = "0" Then

                    'Checa se eh RETORNO
                    If Mid(sBuffer, 143, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                        Exit Sub
                    End If

                    sBanco = Mid(sBuffer, 1, 3)
                    sAgencia = Mid(sBuffer, 53, 6)
                    sConta = Mid(sBuffer, 59, 13)
                    sDataArq = Format(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4), "yyyy-mm-dd")
                    sNumAviso = Mid(sBuffer, 158, 6)
                    sNomeBanco = Trim(Mid(sBuffer, 103, 30))


                    'If IIf(String.IsNullOrEmpty(sBanco), "", sBanco) <> Trim(codigoBanco) Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                    '    Exit Sub
                    'End If

                End If

                'Checa se eh DETALHE
                If Trim(Mid(sBuffer, 8, 1)) = "3" And Trim(Mid(sBuffer, 14, 1)) = "T" Then

                    sCodOcor = Trim(Mid(sBuffer, 16, 2))
                    sCodIncons = Trim(Mid(sBuffer, 214, 10))
                    sNomeCliente = Trim(Mid(sBuffer, 106, 25))
                    sBancoDeb = ""
                    sAgenciaDeb = ""
                    sNumCtaDeb = ""
                    sEndereco = ""
                    sUF = ""
                    sCidade = ""
                    sCEP = ""
                    sNumtit = Trim(Mid(sBuffer, 59, 15))
                    sSeqTit = Strings.Right(sNumtit, 2)
                    sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
                    sVcto = Format(Format(Trim(Mid(sBuffer, 74, 2)) & "/" & Trim(Mid(sBuffer, 76, 2)) & "/" & Trim(Mid(sBuffer, 78, 4)), "dd/MM/yyyy"), "yyyy-MM-dd")
                    strNossoNumBco = Trim(Mid(sBuffer, 38, 20))

                    If sCodOcor = "06" Or sCodOcor = "09" Or sCodOcor = "17" Then
                        'Se for liquidacao pega valor pago e dt. pgto no registro seguinte
                        sBuffer = fsFile.ReadLine
                        dblValor = (Val(Mid(sBuffer, 78, 15)) / 100)
                        sDataPagto = Mid(sBuffer, 142, 4) & "-" & Mid(sBuffer, 140, 2) & "-" & Mid(sBuffer, 138, 2)
                    Else
                        'Se for confirmacao de recebimento pega valor do titulo e dt de vcto
                        dblValor = (Val(Mid(sBuffer, 82, 15)) / 100)
                        sDataPagto = sVcto
                    End If



                    _TempCheckArqRetorno.CodBanco = sBanco
                    _TempCheckArqRetorno.NomeBanco = sNomeBanco
                    _TempCheckArqRetorno.NumAviso = Strings.Right(sNumAviso, 5)
                    _TempCheckArqRetorno.CodAgen = sAgencia
                    _TempCheckArqRetorno.NumCta = sConta
                    _TempCheckArqRetorno.DtArq = sDataArq
                    _TempCheckArqRetorno.NumTit = Strings.Right(sNumtit, 8)
                    _TempCheckArqRetorno.SeqTit = sSeqTit
                    _TempCheckArqRetorno.DtPagto = sDataPagto
                    _TempCheckArqRetorno.NomeCliente = sNomeCliente
                    _TempCheckArqRetorno.CodOcor = sCodOcor
                    _TempCheckArqRetorno.CodIncons = sCodIncons
                    _TempCheckArqRetorno.Mensagem = IIf(sCodOcor = "09", "Baixa", IIf(sCodOcor = "50", "Título pago com cheque, pendente de compensação", "")) 'Informar msg, conforme Mary 27/08/2010
                    _TempCheckArqRetorno.CodBancoDeb = sBancoDeb
                    _TempCheckArqRetorno.CodAgenDeb = sAgenciaDeb
                    _TempCheckArqRetorno.CodNumCtaDeb = sNumCtaDeb
                    _TempCheckArqRetorno.Valor = dblValor
                    _TempCheckArqRetorno.Endereco = sEndereco
                    _TempCheckArqRetorno.Cidade = sCidade
                    _TempCheckArqRetorno.UF = sUF
                    _TempCheckArqRetorno.Cep = sCEP
                    _TempCheckArqRetorno.Arquivo = Trim(nomeArquivo)
                    _TempCheckArqRetorno.DtVcto = sVcto
                    _TempCheckArqRetorno.NossoNumeroBco = strNossoNumBco
                    _TempCheckArqRetorno.SeqTit = sSeqTit

                    _Retorno = clInserirTempCheckArqRetorno.InserirTempCheckArqRetorno(_TempCheckArqRetorno)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                        Exit Sub
                    End If

                End If

                sBuffer = fsFile.ReadLine

            Loop

        End If

        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.Default
        '---------------------------------
        Dim Connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        Connection.Open()
        Dim Transacao As SqlTransaction = Connection.BeginTransaction()

        'Atualiza  a tabela contas_a_receber com o nosso numero no banco
        'If MsgBox("Deseja gravar o Nosso Número ?", vbYesNo, Me.Text) = vbYes Then
        _Retorno = clAlterarContasReceber.AlteraNossoNumeroBcoContasAReceber(False, Connection, Transacao)
        'If Not _Retorno.Sucesso Then
        '    blnRetorno = True
        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

        '    Exit Sub
        'End If
        'End If

        Dim Arquivo As String() = Split(nomeArquivo, "\")
        Dim FileName = Arquivo(Arquivo.Length - 1)

        'VERIFICA SE O ARQUIVO JA FOI CHECADO PRA GRAVAR HISTORICO DO CLIENTE
        lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoChecagemPorCodBancoNumAvisoNumCta(sBanco, Strings.Right(sNumAviso, 5), sConta, Connection, Transacao)
        If lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
            'INSERE HISTORICO DE CONTATO
            _Retorno = clInserirHistoricoContato.IncluiHistoricoContatoChecagem(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "VERISURE", FileName, Connection, Transacao)
            If Not _Retorno.Sucesso Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            End If
            'Douglas - 15/05/2019
            'Insere os mesmos dados na tabela nova
            _Retorno = clInserirTempCheckArqRetorno.InserirChecagemClientesBancariosLog(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "VERISURE", Connection, Transacao)
            If Not _Retorno.Sucesso Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            End If
        End If
        Transacao.Commit()
        Connection.Close()

        'Chama relatorio
        'MsgBox("Checagem concluída com sucesso (ver relatório)!")

        'Exibe Relatório
        'CarregaRelatorio()
        fsFile.Close()
        ValidateBankFiles(sBanco, sNumAviso, sConta, sDataArq, nomeArquivo)
    End Sub
    Private Sub ValidateBankFiles(sBanco As String, sNumAviso As String, sConta As String, sDataArq As String, nomeArquivo As String)
        'If RadMessageBox.Show("Deseja gravar o Nosso Número?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Info) = vbYes Then
        'CONTROLE DE TRANSACAO
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        connection.Open()
        Dim Transacao As SqlTransaction = connection.BeginTransaction()
        Dim _Retorno As New Retorno

        'Atualiza  a tabela contas_a_receber com o nosso numero no banco

        _Retorno = clAlterarContasReceber.AlteraNossoNumeroBcoContasAReceber(False, connection, Transacao)

        If Not _Retorno.Sucesso Then
            Transacao.Rollback()
            connection.Close()
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            Exit Sub
        End If

        Dim Arquivo As String() = Split(nomeArquivo, "\")
        Dim FileName = Arquivo(Arquivo.Length - 1)

        'VERIFICA SE O ARQUIVO JA FOI CHECADO PRA GRAVAR HISTORICO DO CLIENTE nomeArquivo
        'lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoChecagemPorCodBancoNumAvisoNumCta(sBanco, Strings.Right(sNumAviso, 5), sConta, FileName, connection, Transacao)
        lstControleRetBco = clConsultaControleRetBco.BuscarArquivoJaProcessado(FileName, connection, Transacao)
        'If lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
        If Not (lstControleRetBco(0).Sucesso) Then
            'INSERE HISTORICO DE CONTATO
            _Retorno = clInserirHistoricoContato.IncluiHistoricoContatoChecagem(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "VERISURE", FileName, connection, Transacao)
            If Not _Retorno.Sucesso Then
                Transacao.Rollback()
                connection.Close()
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                Exit Sub
            End If

            'Douglas - 15/05/2019
            'Insere os mesmos dados na tabela nova
            _Retorno = clInserirTempCheckArqRetorno.InserirChecagemClientesBancariosLog(sBanco, Strings.Right(sNumAviso, 5), sConta, sDataArq, "VERISURE", connection, Transacao)
            If Not _Retorno.Sucesso Then
                Transacao.Rollback()
                connection.Close()
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                Exit Sub
            End If
        End If

        Transacao.Commit()
        connection.Close()
        'End If
    End Sub

    Private Sub CarregaRelatorio()
        Dim _Retorno As New Retorno
        Dim strRelatorio As String = ""
        Dim strParametro As String = ""
        Dim strValor As String = ""
        'Dim frmCrystal As New TelesystemRelatorio


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------

        'Abre relatorio
        'If rdbBanco.IsChecked = True Then
        '    strRelatorio = "RelatChkRetorno.Rpt"
        'ElseIf rbCadOptante.IsChecked Then
        '    strRelatorio = "RelatChkRetornoOptante.Rpt"
        'Else
        strRelatorio = "RelatChkRetCartao.Rpt"
        'End If

        Dim lrptRelatorio As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        '_Retorno = Funcoes.CarregaRelatorio(frmCrystal.CrystalReportViewer1, strParametro.Split("|"), strValor.Split("|"), strRelatorio, lrptRelatorio)
        'If Not _Retorno.Sucesso Then
        '    blnRetorno = True
        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
        'End If

        'VOLTA O CURSOR NO MOUSE AO NORMAL
        'Me.Cursor = Cursors.Default
        '---------------------------------

        'EXIBE ERRO
        'If blnRetorno Then
        '    ExibirErro()
        '    Exit Sub
        'End If

        'FuncoesUI.CriaRelatorio(strValor, strParametro, strRelatorio, System.Reflection.MethodBase.GetCurrentMethod(), lstRetorno)

        'On Error GoTo Erro
        '
        'CrystalRelatCheckArqRet.EnableParameterPrompting = False
        'CrystalRelatCheckArqRet.DiscardSavedData
        'frmCrystalRelatCheckArqRet.Show

    End Sub

    Private Sub PBaixaTitulos041(nomeArquivo As String)
        Dim iFile As Integer
        Dim sBuffer As String, sEvento As String
        Dim iCon As Integer
        Dim sBanco As String = ""
        Dim sNumAviso As String = ""
        Dim sNumtit As String = ""
        Dim sSeqTit As String = ""
        Dim sVcto As String = ""
        Dim sDtPgto As String = ""
        Dim X As Integer, iNumLcto As Integer
        Dim dDataServ As Date, iUltNum As Integer
        Dim arrContas() As String
        Dim arrValores() As Double
        Dim sAgencia As String
        Dim sConvenio As String = ""
        Dim sConta As String = ""
        Dim sCCorrente As String, sDataArq As String
        Dim sErros As String = ""
        Dim _ControleRetBco As New ControleRetBco
        Dim _Retorno As New Retorno
        Dim _RetornoBco As New RetornoBco


        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            Exit Sub
        End If


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine



        'CONTROLE DE TRANSACAO
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        connection.Open()
        Dim Transacao As SqlTransaction = connection.BeginTransaction()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------

        'Verifica se BO ou DC
        If Trim(Mid(sBuffer, 82, 17)) = "DEBITO AUTOMATICO" Then

            '*****DEBITO AUTOMATICO*************
            'VERIFICA SE EH DA CAIXA
            sBanco = Mid(sBuffer, 43, 3)
            'If sBanco <> codigoBanco Then
            '    blnRetorno = True
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '    fsFile.Close()

            '    GoTo ErrBaixa
            'End If

            Do While (sBuffer) <> Nothing

                'CHECA SE EH HEADER E RETORNO
                If Mid(sBuffer, 1, 1) = "A" And Mid(sBuffer, 2, 1) = "2" Then

                    'Define agência e conta
                    sAgencia = Trim("0413")
                    sConta = Trim("0610678002")

                    'Define o convênio
                    sConvenio = Trim(Mid(sBuffer, 3, 20))

                    'Define numaviso (número sequencial do arquivo NSA)
                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)

                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sConta, connection, Transacao)
                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If


                    'CARREGA A ENTIDADE
                    _ControleRetBco.CodBanco = sBanco
                    _ControleRetBco.NumAviso = sNumAviso
                    _ControleRetBco.NumCta = sConta 'SANTANDER nao possui num. conta no header
                    sDataArq = Format(CDate(Mid(sBuffer, 66, 4) & "/" & Mid(sBuffer, 70, 2) & "/" & Mid(sBuffer, 72, 2)), "yyyy-MM-dd")
                    _ControleRetBco.DtArq = sDataArq

                    'INSERE O ARQUIVO DE RETORNO
                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If
                    sBuffer = fsFile.ReadLine
                End If


                'CHECA SE EH DETALHE DE RETORNO
                'A posição 150 = 5 é retorno de cadastro de optante, não deve inserir nada.
                'Na checagem de arquivo de retorno existe a regra que verifica se o optante foi cadastrado com sucesso ou não
                If Trim(Mid(sBuffer, 1, 1)) = "F" And Trim(Mid(sBuffer, 150, 1)) = "0" Then

                    If Trim(Mid(sBuffer, 68, 2)) = "00" Then

                        'Dados do título
                        sNumtit = Trim(Mid(sBuffer, 70, 8))
                        sSeqTit = Trim(Mid(sBuffer, 78, 2))

                        'Dt vcto do título
                        'sVcto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 70, 4)) & "/" & Trim(Mid(sBuffer, 74, 2)) & "/" & Trim(Mid(sBuffer, 76, 2))), "yyyy-MM-dd")
                        sDtPgto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 45, 4)) & "/" & Trim(Mid(sBuffer, 49, 2)) & "/" & Trim(Mid(sBuffer, 51, 2))), "yyyy-MM-dd")
                        'sVcto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 122, 4)) & "/" & Trim(Mid(sBuffer, 126, 2)) & "/" & Trim(Mid(sBuffer, 128, 2))), "yyyy-MM-dd")
                        sVcto = sDtPgto


                        lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                        If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
                            'CARREGA ENTIDADE
                            _RetornoBco.CodBanco = sBanco
                            _RetornoBco.NumAviso = sNumAviso
                            _RetornoBco.NumTit = sNumtit
                            _RetornoBco.SeqTit = sSeqTit
                            _RetornoBco.CodAgen = sAgencia
                            _RetornoBco.NumCta = sConta
                            _RetornoBco.VlrPago = Val(Mid(sBuffer, 53, 15)) / 100
                            _RetornoBco.VlrJuros = 0
                            _RetornoBco.VlrDesc = 0
                            _RetornoBco.VlrIOF = 0
                            _RetornoBco.VlrAbat = 0
                            _RetornoBco.Processado = "N"
                            _RetornoBco.DtVcto = sVcto
                            _RetornoBco.DtPagto = sDtPgto
                            _RetornoBco.DtArq = sDataArq

                            'INSERE RETORNO BANCO
                            _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                            If Not _Retorno.Sucesso Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                fsFile.Close()

                                GoTo ErrBaixa
                            End If
                        End If

                    ElseIf isErroCaixa(Trim(Mid(sBuffer, 68, 2))) Then

                        'Grava histórico de contato do cliente com o motivo da não baixa do título
                        Dim Inconsistencia As New CodInconsistencias
                        Dim clBuscaCodInconsistencia As New ConsultaCodInconsistencias
                        Dim clInsereErroDadosPgto As New InsereCodInconsistencia
                        Dim clInsereHistContato As New InserirHistoricoContato
                        Dim clBuscaContasAReceber As New ConsultaContasReceber
                        Dim ContasAReceber As New ContaReceber
                        Dim blnIsErro As Boolean
                        Dim Retorno As New Retorno

                        sNumtit = Trim(Mid(sBuffer, 70, 8))
                        sSeqTit = Trim(Mid(sBuffer, 78, 2))

                        'Busca mensagem de inconsistencia
                        Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 68, 2)), "104", blnIsErro, connection, Transacao)

                        If Not Inconsistencia.Sucesso Then
                            Retorno = Funcoes.RetornoFunc(Inconsistencia.MsgErro)
                            lstRetorno.Add(Retorno)
                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        ContasAReceber = clBuscaContasAReceber.ConsultaContaReceberNumTitSeqTit(sNumtit, sSeqTit, connection, Transacao)

                        If Not ContasAReceber.Sucesso Then
                            Retorno = Funcoes.RetornoFunc(ContasAReceber.MsgErro)
                            lstRetorno.Add(Retorno)
                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        'Insere no histórico de contato do cliente motivo de não baixa do título
                        Retorno = clInsereHistContato.IncluiHistoricoContato(ContasAReceber.CodIntClie, "VERISURE", "Título " & sNumtit & "-" & sSeqTit & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

                        If Not Retorno.Sucesso Then
                            lstRetorno.Add(Retorno)
                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        'Caso o código de incosistencia for um erro para assinalar o cliente, realizar update no cliente - Lucas 15/05/2017
                        If blnIsErro Then

                            Retorno = clInsereErroDadosPgto.InsereErroDadosPgtoCliente(ContasAReceber.CodIntClie, connection, Transacao)

                            If Not Retorno.Sucesso Then
                                lstRetorno.Add(Retorno)
                                'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                                blnRetorno = True
                                Exit Sub
                            End If

                        End If

                        'sBuffer = fsFile.ReadLine

                    End If

                End If
                sBuffer = fsFile.ReadLine
            Loop
        Else
            '**********BOLETO******************
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "SANTANDER (boleto) ainda não liberado para faturamento. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------
            'RadMessageBox.Show(Me.Text, "Caixa (boleto) ainda não liberado para faturamento. Consultar a área de sistemas!", MessageBoxButtons.OK, RadMessageIcon.Exclamation)
            'Exit Sub
            ''VERIFICA SE EH DO SANTANDER
            'sBanco = Mid(sBuffer, 1, 3)
            'If sBanco <> codigoBanco Then
            '    blnRetorno = True
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '    fsFile.Close()

            '    GoTo ErrBaixa
            'End If

            'Do While (sBuffer) <> Nothing
            '    'CHECA SE EH HEADER
            '    If Mid(sBuffer, 8, 1) = "0" Then
            '        'CHECA SE EH RETORNO
            '        If Mid(sBuffer, 143, 1) <> "2" Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If

            '        sBanco = Mid(sBuffer, 1, 3)

            '        'VERIFICA SE EH DO SANTANDER
            '        If sBanco <> codigoBanco Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If
            '        sAgencia = Mid(sBuffer, 33, 4)
            '        sCCorrente = Mid(sBuffer, 39, 9)

            '        'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
            '        If codigoBanco <> sBanco Or Trim(lblAgen.Text) <> Trim(sAgencia) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(sCCorrente)) Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", sBanco).Replace("<%AGENCIA%>", sAgencia).Replace("<%CONTA%>", sCCorrente), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If

            '        sDataArq = Format(CDate(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4)), "yyyy-MM-dd")
            '        sNumAviso = Strings.Right(Mid(sBuffer, 158, 6), 5)

            '        'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
            '        lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sCCorrente, connection, Transacao)
            '        If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If


            '        'CARREGA A ENTIDADE
            '        _ControleRetBco.CodBanco = sBanco
            '        _ControleRetBco.NumAviso = sNumAviso
            '        _ControleRetBco.NumCta = sCCorrente 'SANTANDER nao possui num. conta no header
            '        _ControleRetBco.DtArq = sDataArq

            '        'INSERE O ARQUIVO DE RETORNO
            '        _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
            '        If Not _Retorno.Sucesso Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If

            '    End If


            '    'CHECA SE EH DETALHE
            '    '*********************************************
            '    'DETALHE SEGMENTO T
            '    '*********************************************
            '    If Trim(Mid(sBuffer, 14, 1)) = "T" Then
            '        'sNumtit = Trim(Mid(sBuffer, 55, 15))
            '        'sSeqTit = Right(sNumtit, 2)
            '        'sNumtit = Left(sNumtit, Len(sNumtit) - 2)
            '        'sVcto = Format(Format(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4)), "dd/mm/yyyy"), "yyyy-mm-dd")

            '        'CHECA SE EH LIQUIDACAO
            '        If Mid(sBuffer, 16, 2) = "06" Then 'Or Mid(sBuffer, 16, 2) = "09" Or Mid(sBuffer, 16, 2) = "17" Then

            '            sNumtit = Trim(Mid(sBuffer, 55, 15))
            '            sSeqTit = Strings.Right(sNumtit, 2)
            '            sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
            '            sVcto = Format(CDate(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4))), "yyyy-MM-dd")

            '            lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
            '            If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            '                blnRetorno = True
            '                lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
            '                fsFile.Close()

            '                GoTo ErrBaixa
            '            Else
            '                'CARREGA ENTIDADE
            '                _RetornoBco.CodBanco = sBanco
            '                _RetornoBco.NumAviso = sNumAviso
            '                _RetornoBco.NumTit = sNumtit
            '                _RetornoBco.SeqTit = sSeqTit
            '                _RetornoBco.CodAgen = sAgencia
            '                _RetornoBco.NumCta = sCCorrente

            '                sBuffer = fsFile.ReadLine

            '                If Trim(Mid(sBuffer, 14, 1)) = "U" Then
            '                    _RetornoBco.VlrPago = Val(Mid(sBuffer, 78, 15)) / 100
            '                    _RetornoBco.VlrJuros = (Val(Mid(sBuffer, 18, 15)) / 100) + (Val(Mid(sBuffer, 123, 15)) / 100)
            '                    _RetornoBco.VlrDesc = Val(Mid(sBuffer, 33, 15)) / 100
            '                    _RetornoBco.VlrIOF = 0
            '                    _RetornoBco.VlrAbat = Val(Mid(sBuffer, 48, 15)) / 100
            '                    _RetornoBco.Processado = "N"
            '                    _RetornoBco.DtVcto = sVcto
            '                    _RetornoBco.DtPagto = Format(CDate(Trim(Mid(sBuffer, 138, 2)) & "/" & Trim(Mid(sBuffer, 140, 2)) & "/" & Trim(Mid(sBuffer, 142, 4))), "dd/MM/yyyy")
            '                    _RetornoBco.DtArq = sDataArq

            '                    'INSERE RETORNO BANCO
            '                    _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
            '                    If Not _Retorno.Sucesso Then
            '                        blnRetorno = True
            '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            '                        fsFile.Close()

            '                        GoTo ErrBaixa
            '                    End If
            '                Else
            '                    blnRetorno = True
            '                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Descricao, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '                    fsFile.Close()

            '                    GoTo ErrBaixa
            '                End If

            '            End If

            '        End If
            '    End If
            '    sBuffer = fsFile.ReadLine
            'Loop
        End If
        '**************************************************************************************************************
        fsFile.Close()


        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(sDataArq), Trim("0610678002"), sAgencia, connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            GoTo ErrBaixa
        End If


        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

            GoTo ErrBaixa
        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            'For i As Integer = 0 To lstBxAutoErros.Count - 1
            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            'Next
            'grdTitJaBaixados.DataSource = lstBxAutoErros
            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

            'If (grdTitJaBaixados.Rows.Count > 0) Then
            '    gbxTitJaBaixados.Visible = True
            'Else
            '    gbxTitJaBaixados.Visible = False
            'End If
        End If

        'EXIBE MESSAGEM
        If blnRetorno Then
            'ExibirErro()
        End If

        Transacao.Commit()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

        'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
        Exit Sub

ErrBaixa:

        Transacao.Rollback()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------
    End Sub
    Private Function isErroCaixa(ByVal CodInconsistencia As String) As Boolean

        Select Case CodInconsistencia
            Case "01"
                Return True
            Case "02"
                Return True
            Case "04"
                Return True
            Case "10"
                Return True
            Case "12"
                Return True
            Case "13"
                Return True
            Case "14"
                Return True
            Case "15"
                Return True
            Case "18"
                Return True
            Case "30"
                Return True
            Case "96"
                Return True
            Case "97"
                Return True
            Case "98"
                Return True
            Case "99"
                Return True
            Case Else
                Return False
        End Select

    End Function
    Private Sub PBaixaTitulos104(nomeArquivo As String) 'CAIXA

        '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
        'MsgBox "SANTANDER ainda não liberado para faturamento. Consultar analista de sistemas !"
        'Exit Sub
        '---------------------------------------------------------------

        Dim iFile As Integer
        Dim sBuffer As String, sEvento As String
        Dim iCon As Integer
        Dim sBanco As String = ""
        Dim sNumAviso As String = ""
        Dim sNumtit As String = ""
        Dim sSeqTit As String = ""
        Dim sVcto As String = ""
        Dim sDtPgto As String = ""
        Dim X As Integer, iNumLcto As Integer
        Dim dDataServ As Date, iUltNum As Integer
        Dim arrContas() As String
        Dim arrValores() As Double
        Dim sAgencia As String
        Dim sConvenio As String = ""
        Dim sConta As String = ""
        Dim sCCorrente As String, sDataArq As String
        Dim sErros As String = ""
        Dim _ControleRetBco As New ControleRetBco
        Dim _Retorno As New Retorno
        Dim _RetornoBco As New RetornoBco


        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            Exit Sub
        End If


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine


        'CONTROLE DE TRANSACAO
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        connection.Open()
        Dim Transacao As SqlTransaction = connection.BeginTransaction()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------

        'Verifica se BO ou DC
        If Trim(Mid(sBuffer, 82, 17)) = "DEBITO AUTOMATICO" Then

            '*****DEBITO AUTOMATICO*************
            'VERIFICA SE EH DA CAIXA
            sBanco = Mid(sBuffer, 43, 3)
            'If sBanco <> codigoBanco Then
            '    blnRetorno = True
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '    fsFile.Close()

            '    GoTo ErrBaixa
            'End If

            Do While (sBuffer) <> Nothing

                'CHECA SE EH HEADER E RETORNO
                If Mid(sBuffer, 1, 1) = "A" And Mid(sBuffer, 2, 1) = "2" Then

                    'Define agência e conta
                    sAgencia = Trim("42714")
                    sConta = Trim("00906923-2")

                    'Define o convênio
                    sConvenio = Trim(Mid(sBuffer, 3, 20))

                    'Define numaviso (número sequencial do arquivo NSA)
                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)

                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sConta, connection, Transacao)
                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If


                    'CARREGA A ENTIDADE
                    _ControleRetBco.CodBanco = sBanco
                    _ControleRetBco.NumAviso = sNumAviso
                    _ControleRetBco.NumCta = sConta 'SANTANDER nao possui num. conta no header
                    sDataArq = Format(CDate(Mid(sBuffer, 66, 4) & "/" & Mid(sBuffer, 70, 2) & "/" & Mid(sBuffer, 72, 2)), "yyyy-MM-dd")
                    _ControleRetBco.DtArq = sDataArq

                    'INSERE O ARQUIVO DE RETORNO
                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                End If


                'CHECA SE EH DETALHE DE RETORNO
                'A posição 150 = 5 é retorno de cadastro de optante, não deve inserir nada.
                'Na checagem de arquivo de retorno existe a regra que verifica se o optante foi cadastrado com sucesso ou não
                If Trim(Mid(sBuffer, 1, 1)) = "F" And Trim(Mid(sBuffer, 150, 1)) = "0" Then

                    If Trim(Mid(sBuffer, 68, 2)) = "00" Then

                        Dim split As String()
                        split = sBuffer.Split("#")

                        Dim codintclie As String = split(1)
                        sNumtit = Trim(Mid(split(2), 1, 8))
                        sSeqTit = Trim(Mid(split(2), 9, 2))


                        'Dados do título
                        'sNumtit = Trim(Mid(sBuffer, 70, 8))
                        'sSeqTit = Trim(Mid(sBuffer, 78, 2))

                        'Dt vcto do título
                        'sVcto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 70, 4)) & "/" & Trim(Mid(sBuffer, 74, 2)) & "/" & Trim(Mid(sBuffer, 76, 2))), "yyyy-MM-dd")
                        sDtPgto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 45, 4)) & "/" & Trim(Mid(sBuffer, 49, 2)) & "/" & Trim(Mid(sBuffer, 51, 2))), "yyyy-MM-dd")
                        sVcto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 122, 4)) & "/" & Trim(Mid(sBuffer, 126, 2)) & "/" & Trim(Mid(sBuffer, 128, 2))), "yyyy-MM-dd")
                        'sVcto = sDtPgto


                        lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                        If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
                            'CARREGA ENTIDADE
                            _RetornoBco.CodBanco = sBanco
                            _RetornoBco.NumAviso = sNumAviso
                            _RetornoBco.NumTit = sNumtit
                            _RetornoBco.SeqTit = sSeqTit
                            _RetornoBco.CodAgen = strAgen
                            _RetornoBco.NumCta = strContaCorrente
                            _RetornoBco.VlrPago = Val(Mid(sBuffer, 53, 15)) / 100
                            _RetornoBco.VlrJuros = 0
                            _RetornoBco.VlrDesc = 0
                            _RetornoBco.VlrIOF = 0
                            _RetornoBco.VlrAbat = 0
                            _RetornoBco.Processado = "N"
                            _RetornoBco.DtVcto = sVcto
                            _RetornoBco.DtPagto = sDtPgto
                            _RetornoBco.DtArq = sDataArq

                            'INSERE RETORNO BANCO
                            _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                            If Not _Retorno.Sucesso Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                fsFile.Close()

                                GoTo ErrBaixa
                            End If
                        End If

                    ElseIf isErroCaixa(Trim(Mid(sBuffer, 68, 2))) Then

                        'Grava histórico de contato do cliente com o motivo da não baixa do título
                        Dim Inconsistencia As New CodInconsistencias
                        Dim clBuscaCodInconsistencia As New ConsultaCodInconsistencias
                        Dim clInsereErroDadosPgto As New InsereCodInconsistencia
                        Dim clInsereHistContato As New InserirHistoricoContato
                        Dim clBuscaContasAReceber As New ConsultaContasReceber
                        Dim ContasAReceber As New ContaReceber
                        Dim blnIsErro As Boolean
                        Dim Retorno As New Retorno

                        Dim split As String()
                        split = sBuffer.Split("#")

                        Dim codintclie As String = split(1)
                        sNumtit = Trim(Mid(split(2), 1, 8))
                        sSeqTit = Trim(Mid(split(2), 9, 2))

                        'sNumtit = Trim(Mid(sBuffer, 70, 8))
                        'sSeqTit = Trim(Mid(sBuffer, 78, 2))

                        'Busca mensagem de inconsistencia
                        Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 68, 2)), "104", blnIsErro, connection, Transacao)

                        If Not Inconsistencia.Sucesso Then
                            Retorno = Funcoes.RetornoFunc(Inconsistencia.MsgErro)
                            lstRetorno.Add(Retorno)
                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        ContasAReceber = clBuscaContasAReceber.ConsultaContaReceberNumTitSeqTit(sNumtit, sSeqTit, connection, Transacao)

                        If Not ContasAReceber.Sucesso Then
                            Retorno = Funcoes.RetornoFunc(ContasAReceber.MsgErro)
                            lstRetorno.Add(Retorno)
                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        'Insere no histórico de contato do cliente motivo de não baixa do título
                        Retorno = clInsereHistContato.IncluiHistoricoContato(ContasAReceber.CodIntClie, "VERISURE", "Título " & sNumtit & "-" & sSeqTit & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

                        If Not Retorno.Sucesso Then
                            lstRetorno.Add(Retorno)
                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        'Caso o código de incosistencia for um erro para assinalar o cliente, realizar update no cliente - Lucas 15/05/2017
                        If blnIsErro Then

                            Retorno = clInsereErroDadosPgto.InsereErroDadosPgtoCliente(ContasAReceber.CodIntClie, connection, Transacao)

                            If Not Retorno.Sucesso Then
                                lstRetorno.Add(Retorno)
                                'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                                blnRetorno = True
                                Exit Sub
                            End If

                        End If

                        'sBuffer = fsFile.ReadLine

                    End If

                End If
                sBuffer = fsFile.ReadLine
            Loop
        Else
            '**********BOLETO******************
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "SANTANDER (boleto) ainda não liberado para faturamento. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------
            'RadMessageBox.Show(Me.Text, "Caixa (boleto) ainda não liberado para faturamento. Consultar a área de sistemas!", MessageBoxButtons.OK, RadMessageIcon.Exclamation)
            'Exit Sub
            ''VERIFICA SE EH DO SANTANDER
            'sBanco = Mid(sBuffer, 1, 3)
            'If sBanco <> codigoBanco Then
            '    blnRetorno = True
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '    fsFile.Close()

            '    GoTo ErrBaixa
            'End If

            'Do While (sBuffer) <> Nothing
            '    'CHECA SE EH HEADER
            '    If Mid(sBuffer, 8, 1) = "0" Then
            '        'CHECA SE EH RETORNO
            '        If Mid(sBuffer, 143, 1) <> "2" Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If

            '        sBanco = Mid(sBuffer, 1, 3)

            '        'VERIFICA SE EH DO SANTANDER
            '        If sBanco <> codigoBanco Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If
            '        sAgencia = Mid(sBuffer, 33, 4)
            '        sCCorrente = Mid(sBuffer, 39, 9)

            '        'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
            '        If codigoBanco <> sBanco Or Trim(lblAgen.Text) <> Trim(sAgencia) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(sCCorrente)) Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", sBanco).Replace("<%AGENCIA%>", sAgencia).Replace("<%CONTA%>", sCCorrente), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If

            '        sDataArq = Format(CDate(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4)), "yyyy-MM-dd")
            '        sNumAviso = Strings.Right(Mid(sBuffer, 158, 6), 5)

            '        'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
            '        lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sCCorrente, connection, Transacao)
            '        If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If


            '        'CARREGA A ENTIDADE
            '        _ControleRetBco.CodBanco = sBanco
            '        _ControleRetBco.NumAviso = sNumAviso
            '        _ControleRetBco.NumCta = sCCorrente 'SANTANDER nao possui num. conta no header
            '        _ControleRetBco.DtArq = sDataArq

            '        'INSERE O ARQUIVO DE RETORNO
            '        _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
            '        If Not _Retorno.Sucesso Then
            '            blnRetorno = True
            '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            '            fsFile.Close()

            '            GoTo ErrBaixa
            '        End If

            '    End If


            '    'CHECA SE EH DETALHE
            '    '*********************************************
            '    'DETALHE SEGMENTO T
            '    '*********************************************
            '    If Trim(Mid(sBuffer, 14, 1)) = "T" Then
            '        'sNumtit = Trim(Mid(sBuffer, 55, 15))
            '        'sSeqTit = Right(sNumtit, 2)
            '        'sNumtit = Left(sNumtit, Len(sNumtit) - 2)
            '        'sVcto = Format(Format(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4)), "dd/mm/yyyy"), "yyyy-mm-dd")

            '        'CHECA SE EH LIQUIDACAO
            '        If Mid(sBuffer, 16, 2) = "06" Then 'Or Mid(sBuffer, 16, 2) = "09" Or Mid(sBuffer, 16, 2) = "17" Then

            '            sNumtit = Trim(Mid(sBuffer, 55, 15))
            '            sSeqTit = Strings.Right(sNumtit, 2)
            '            sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
            '            sVcto = Format(CDate(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4))), "yyyy-MM-dd")

            '            lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
            '            If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            '                blnRetorno = True
            '                lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
            '                fsFile.Close()

            '                GoTo ErrBaixa
            '            Else
            '                'CARREGA ENTIDADE
            '                _RetornoBco.CodBanco = sBanco
            '                _RetornoBco.NumAviso = sNumAviso
            '                _RetornoBco.NumTit = sNumtit
            '                _RetornoBco.SeqTit = sSeqTit
            '                _RetornoBco.CodAgen = sAgencia
            '                _RetornoBco.NumCta = sCCorrente

            '                sBuffer = fsFile.ReadLine

            '                If Trim(Mid(sBuffer, 14, 1)) = "U" Then
            '                    _RetornoBco.VlrPago = Val(Mid(sBuffer, 78, 15)) / 100
            '                    _RetornoBco.VlrJuros = (Val(Mid(sBuffer, 18, 15)) / 100) + (Val(Mid(sBuffer, 123, 15)) / 100)
            '                    _RetornoBco.VlrDesc = Val(Mid(sBuffer, 33, 15)) / 100
            '                    _RetornoBco.VlrIOF = 0
            '                    _RetornoBco.VlrAbat = Val(Mid(sBuffer, 48, 15)) / 100
            '                    _RetornoBco.Processado = "N"
            '                    _RetornoBco.DtVcto = sVcto
            '                    _RetornoBco.DtPagto = Format(CDate(Trim(Mid(sBuffer, 138, 2)) & "/" & Trim(Mid(sBuffer, 140, 2)) & "/" & Trim(Mid(sBuffer, 142, 4))), "dd/MM/yyyy")
            '                    _RetornoBco.DtArq = sDataArq

            '                    'INSERE RETORNO BANCO
            '                    _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
            '                    If Not _Retorno.Sucesso Then
            '                        blnRetorno = True
            '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
            '                        fsFile.Close()

            '                        GoTo ErrBaixa
            '                    End If
            '                Else
            '                    blnRetorno = True
            '                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Descricao, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '                    fsFile.Close()

            '                    GoTo ErrBaixa
            '                End If

            '            End If

            '        End If
            '    End If
            '    sBuffer = fsFile.ReadLine
            'Loop
        End If
        '**************************************************************************************************************
        fsFile.Close()


        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(sDataArq), Trim("00906923-2"), sAgencia, connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            GoTo ErrBaixa
        End If


        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

            GoTo ErrBaixa
        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            'For i As Integer = 0 To lstBxAutoErros.Count - 1
            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            'Next
            'grdTitJaBaixados.DataSource = lstBxAutoErros
            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

            'If (grdTitJaBaixados.Rows.Count > 0) Then
            '    gbxTitJaBaixados.Visible = True
            'Else
            '    gbxTitJaBaixados.Visible = False
            'End If
        End If

        'EXIBE MESSAGEM
        If blnRetorno Then
            'ExibirErro()
        End If

        Transacao.Commit()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

        'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
        Exit Sub

ErrBaixa:

        Transacao.Rollback()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------
    End Sub
    '    Private Sub PBaixaTitulos356(nomeArquivo As String) 'Retorno do Banco REAL

    '        '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
    '        'MsgBox "REAL ainda não liberado para faturamento. Consultar analista de sistemas !"
    '        'Exit Sub
    '        '---------------------------------------------------------------

    '        Dim iFile As Integer
    '        Dim sBuffer As String, sEvento As String
    '        Dim iCon As Integer
    '        Dim sBanco As String, sNumAviso As String, sNumtit As String, sSeqTit As String, sVcto As String
    '        Dim X As Integer, iNumLcto As Integer
    '        Dim dDataServ As Date, iUltNum As Integer
    '        Dim arrContas() As String
    '        Dim arrValores() As Double
    '        Dim sAgencia As String, sCCorrente As String
    '        Dim sErros As String = ""
    '        Dim sConta As String
    '        Dim sDataArq As String = ""

    '        Dim _ControleRetBco As New ControleRetBco
    '        Dim _Retorno As New Retorno
    '        Dim _RetornoBco As New RetornoBco

    '        iFile = FreeFile()

    '        If Dir(nomeArquivo) = "" Then
    '            blnRetorno = True
    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

    '            Exit Sub
    '        End If

    '        'VOLTA O CURSOR NO MOUSE
    '        'Me.Cursor = Cursors.WaitCursor
    '        '---------------------------------

    '        fsFile = New System.IO.StreamReader(nomeArquivo)
    '        sBuffer = fsFile.ReadLine


    '        'CONTROLE DE TRANSACAO
    '        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '        connection.Open()
    '        Dim Transacao As SqlTransaction = connection.BeginTransaction()
    '        '-----------------------------------------------------------------------------------------------------------------------------------------------------------

    '        iCon = 1

    '        'Verifica se é DC ou BO
    '        If Mid(sBuffer, 82, 17) = "DEBITO AUTOMATICO" Then


    '            '******************************************
    '            '********** DEBITO AUTOMATICO *************

    '            ''***Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado*****************
    '            'Do While Not EOF(iFile)
    '            '
    '            '    If Mid(sBuffer, 1, 1) = "A" And Mid(sBuffer, 2, 1) = "2" Then
    '            '
    '            '        If Mid(sBuffer, 2, 1) <> "2" Then
    '            '            MsgBox "Este arquivo não é de retorno !", vbExclamation, Me.Caption
    '            '            Close #iFile
    '            '            Set clSig = Nothing
    '            '            Exit Sub
    '            '        End If
    '            '
    '            '        sBanco = Mid(sBuffer, 43, 3)
    '            '        sCCorrente = strContas
    '            '
    '            '        'VERIFICA SE EH DO REAL
    '            '        If sBanco <> txtCodBanco.Text Then
    '            '            MsgBox "Este arquivo não é do do Banco Real !", vbExclamation, Me.Caption
    '            '            Close #iFile
    '            '            Set clSig = Nothing
    '            '            Exit Sub
    '            '        End If
    '            '
    '            '        sNumAviso = Right(Mid(sBuffer, 74, 6), 5)
    '            '        .OpenTable "SELECT * FROM CONTROLE_RET_BCO WHERE CodBanco = '" & sBanco & "' and NumAviso = '" & sNumAviso & "' and NumCta = '" & sCCorrente & "'"
    '            '
    '            '        If Not .Cursor.EOF Then
    '            '            MsgBox "Este arquivo de retorno já foi atualizado no sistema !", vbExclamation, Me.Caption
    '            '            Close #iFile
    '            '            Set clSig = Nothing
    '            '            Exit Sub
    '            '        End If
    '            '    End If
    '            '
    '            '    If Trim(Mid(sBuffer, 1, 1)) = "F" And Trim(Mid(sBuffer, 68, 2)) = "00" Then
    '            '        sAgencia = Trim(Mid(sBuffer, 27, 4))
    '            '        sConta = Trim(Mid(sBuffer, 31, 14))
    '            '
    '            '        If txtCodBanco.Text <> sBanco Or CInt(Trim(strAgencia)) <> CInt(sAgencia) Or CLng(FSepara(strContas)) <> CLng(FSepara(sConta)) Then
    '            '           MsgBox "Arquivo de Retorno Banco/Agência/Conta Corrente : [" & sBanco & "/" & sAgencia & "/" & sConta & "] não corresponde ao selecionado !!", vbExclamation, Me.Caption
    '            '           Close #iFile
    '            '           Set clSig = Nothing
    '            '           Exit Sub
    '            '        End If
    '            '        Close #iFile
    '            '        Exit Do
    '            '    End If
    '            'Loop
    '            ''******************************************************************************************

    '            'Open Trim(txtArquivo.Text) For Input As #iFile




    '            'VERIFICA SE EH DO REAL
    '            sBanco = Mid(sBuffer, 43, 3)
    '            'If sBanco <> codigoBanco Then
    '            '    blnRetorno = True
    '            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '            '    fsFile.Close()

    '            '    GoTo ErrBaixa
    '            'End If

    '            Do While (sBuffer) <> Nothing

    '                'CHECA SE EH HEADER E RETORNO
    '                If Mid(sBuffer, 1, 1) = "A" And Mid(sBuffer, 2, 1) = "2" Then

    '                    If Mid(sBuffer, 2, 1) <> "2" Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                    sBanco = Mid(sBuffer, 43, 3)
    '                    sCCorrente = lblCCorrente.Text

    '                    'VERIFICA SE EH DO REAL
    '                    'If sBanco <> codigoBanco Then
    '                    '    blnRetorno = True
    '                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                    '    fsFile.Close()

    '                    '    GoTo ErrBaixa
    '                    'End If


    '                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
    '                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)
    '                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sCCorrente, connection, Transacao)
    '                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                    'CARREGA A ENTIDADE
    '                    _ControleRetBco.CodBanco = sBanco
    '                    _ControleRetBco.NumAviso = sNumAviso
    '                    _ControleRetBco.NumCta = sCCorrente
    '                    sDataArq = Convert.ToDateTime(Format(CDate(Mid(sBuffer, 66, 4) & "/" & Mid(sBuffer, 70, 2) & "/" & Mid(sBuffer, 72, 2)), "yyyy-MM-dd")).Date
    '                    _ControleRetBco.DtArq = sDataArq

    '                    'INSERE O ARQUIVO DE RETORNO
    '                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
    '                    If Not _Retorno.Sucesso Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If
    '                End If


    '                'CHECA SE EH DETALHE DE RETORNO
    '                If Trim(Mid(sBuffer, 1, 1)) = "F" And Trim(Mid(sBuffer, 68, 2)) = "00" Then
    '                    sNumtit = Trim(Mid(sBuffer, 70, 52))
    '                    sSeqTit = Strings.Right(sNumtit, 2)
    '                    sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
    '                    sVcto = Mid(sBuffer, 122, 4) & "-" & Mid(sBuffer, 126, 2) & "-" & Mid(sBuffer, 128, 2)


    '                    lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
    '                    If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
    '                        'CARREGA ENTIDADE
    '                        _RetornoBco.CodBanco = sBanco
    '                        _RetornoBco.NumAviso = sNumAviso
    '                        _RetornoBco.NumTit = sNumtit
    '                        _RetornoBco.SeqTit = sSeqTit
    '                        _RetornoBco.CodAgen = strAgen 'Trim(Mid(sBuffer, 27, 4))
    '                        _RetornoBco.NumCta = strContaCorrente 'Trim(Str(Val(Mid(sBuffer, 38, 5)))) & Trim(Mid(sBuffer, 43, 2))
    '                        _RetornoBco.VlrPago = Val(Mid(sBuffer, 53, 15)) / 100
    '                        _RetornoBco.VlrJuros = 0
    '                        _RetornoBco.VlrDesc = 0
    '                        _RetornoBco.VlrIOF = 0
    '                        _RetornoBco.VlrAbat = 0
    '                        _RetornoBco.Processado = "N"
    '                        _RetornoBco.DtVcto = sVcto
    '                        _RetornoBco.DtPagto = Mid(sBuffer, 45, 4) & " - " & Mid(sBuffer, 49, 2) & " - " & Mid(sBuffer, 51, 2)
    '                        _RetornoBco.DtArq = sDataArq

    '                        'INSERE RETORNO BANCO
    '                        _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
    '                        If Not _Retorno.Sucesso Then
    '                            blnRetorno = True
    '                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                            fsFile.Close()

    '                            GoTo ErrBaixa
    '                        End If

    '                    End If
    '                    'Else

    '                    '    'Grava histórico de contato do cliente com o motivo da não baixa do título
    '                    '    Dim Inconsistencia As New CodInconsistencias
    '                    '    Dim clBuscaCodInconsistencia As New Teleatlantic.TLS.CodInconsistenciasBC.ConsultaCodInconsistencias
    '                    '    Dim clInsereHistContato As New Teleatlantic.TLS.HistoricoContatoBC.InserirHistoricoContato
    '                    '    Dim ContasAReceber As New ContaReceber
    '                    '    Dim Retorno As New Retorno

    '                    '    'Busca mensagem de inconsistencia
    '                    '    Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 68, 2)), "356", connection, Transacao)

    '                    '    If Not Inconsistencia.Sucesso Then
    '                    '        Retorno = Funcoes.RetornoFunc(Inconsistencia.MsgErro)
    '                    '        lstRetorno.Add(Retorno)
    '                    '        FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
    '                    '        blnRetorno = True
    '                    '        Exit Sub
    '                    '    End If

    '                    '    'Insere no histórico de contato do cliente motivo de não baixa do título
    '                    '    Retorno = clInsereHistContato.IncluiHistoricoContato(ContasAReceber.CodIntClie, "VERISURE"(), "Título " & Trim(Mid(sBuffer, 70, 52)) & "-" & Strings.Right(Trim(Mid(sBuffer, 70, 52)), 2) & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

    '                    '    If Not Retorno.Sucesso Then
    '                    '        lstRetorno.Add(Retorno)
    '                    '        FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
    '                    '        blnRetorno = True
    '                    '        Exit Sub
    '                    '    End If

    '                    '    sBuffer = fsFile.ReadLine

    '                End If
    '                sBuffer = fsFile.ReadLine
    '            Loop
    '        Else
    '            '*******************************
    '            '***********BOLETO *************

    '            'VERIFICA SE EH DO REAL
    '            sBanco = Mid(sBuffer, 77, 3)
    '            'If sBanco <> codigoBanco Then
    '            '    blnRetorno = True
    '            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '            '    fsFile.Close()

    '            '    GoTo ErrBaixa
    '            'End If

    '            Do While (sBuffer) <> Nothing

    '                'CHECA SE EH HEADER E RETORNO
    '                If Mid(sBuffer, 1, 1) = "0" Then

    '                    If Trim(Mid(sBuffer, 2, 25)) <> "2RETORNO01COBRANCA" Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                    sBanco = Mid(sBuffer, 77, 3)
    '                    'sCCorrente = Mid(sBuffer, 33, 7)
    '                    sCCorrente = "B" & Mid(sBuffer, 33, 7) 'Para diferenciar baixa com mesmo seq da mesma conta DC de BO
    '                    sAgencia = Mid(sBuffer, 28, 4)

    '                    'VERIFICA SE EH DO REAL
    '                    'If sBanco <> codigoBanco Then
    '                    '    blnRetorno = True
    '                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                    '    fsFile.Close()

    '                    '    GoTo ErrBaixa
    '                    'End If

    '                    'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
    '                    'If codigoBanco <> sBanco Or Trim(lblAgen.Text) <> Trim(sAgencia) Or Trim(Funcoes.FSepara(Strings.Left(lblCCorrente.Text, 7))) <> Trim(Funcoes.FSepara(Mid(sBuffer, 33, 7))) Then
    '                    '    blnRetorno = True
    '                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", sBanco).Replace("<%AGENCIA%>", sAgencia).Replace("<%CONTA%>", Funcoes.FSepara(Mid(sBuffer, 33, 7))), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                    '    fsFile.Close()

    '                    '    GoTo ErrBaixa
    '                    'End If

    '                    sNumAviso = Strings.Right(Mid(sBuffer, 109, 8), 5)

    '                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
    '                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sCCorrente, connection, Transacao)
    '                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                    sDataArq = Format(CDate(Mid(sBuffer, 95, 2) & "/" & Mid(sBuffer, 97, 2) & "/20" & Mid(sBuffer, 99, 2)), "yyyy-MM-dd")

    '                    'CARREGA A ENTIDADE
    '                    _ControleRetBco.CodBanco = sBanco
    '                    _ControleRetBco.NumAviso = sNumAviso
    '                    _ControleRetBco.NumCta = sCCorrente
    '                    _ControleRetBco.DtArq = sDataArq

    '                    'INSERE O ARQUIVO DE RETORNO
    '                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
    '                    If Not _Retorno.Sucesso Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                End If

    '                'CHECA SE EH DETALHE
    '                If Trim(Mid(sBuffer, 1, 1)) = "1" Then

    '                    'CHECA SE EH LIQUIDACAO
    '                    If Mid(sBuffer, 109, 2) = "06" Or Mid(sBuffer, 109, 2) = "10" Or Mid(sBuffer, 109, 2) = "98" Then
    '                        sNumtit = Trim(Mid(sBuffer, 117, 10))
    '                        sSeqTit = Strings.Right(sNumtit, 2)
    '                        sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
    '                        'Debug.Print "'" & sNumTit & "',"
    '                        sVcto = Format(CDate(Trim(Mid(sBuffer, 147, 2)) & "/" & Trim(Mid(sBuffer, 149, 2)) & "/20" & Trim(Mid(sBuffer, 151, 2))), "yyyy-MM-dd")


    '                        lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
    '                        If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
    '                            blnRetorno = True
    '                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
    '                            fsFile.Close()

    '                            GoTo ErrBaixa
    '                        ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
    '                            'CARREGA ENTIDADE
    '                            _RetornoBco.CodBanco = sBanco
    '                            _RetornoBco.NumAviso = sNumAviso
    '                            _RetornoBco.NumTit = sNumtit
    '                            _RetornoBco.SeqTit = sSeqTit
    '                            _RetornoBco.CodAgen = strAgen
    '                            _RetornoBco.NumCta = strContaCorrente
    '                            _RetornoBco.VlrPago = Val(Mid(sBuffer, 254, 13)) / 100
    '                            _RetornoBco.VlrJuros = (Val(Mid(sBuffer, 267, 13)) / 100) + (Val(Mid(sBuffer, 215, 13)) / 100) + (Val(Mid(sBuffer, 280, 13)) / 100) 'Multa + Juros + Outros
    '                            _RetornoBco.VlrDesc = Val(Mid(sBuffer, 241, 13)) / 100
    '                            _RetornoBco.VlrIOF = 0
    '                            _RetornoBco.VlrAbat = Val((Val(Mid(sBuffer, 228, 13))) + (Val(Mid(sBuffer, 189, 13)))) / 100
    '                            _RetornoBco.Processado = "N"
    '                            _RetornoBco.DtVcto = sVcto
    '                            _RetornoBco.DtPagto = "20" & Mid(sBuffer, 115, 2) & "-" & Mid(sBuffer, 113, 2) & "-" & Mid(sBuffer, 111, 2)
    '                            _RetornoBco.DtArq = sDataArq

    '                            'INSERE RETORNO BANCO
    '                            _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
    '                            If Not _Retorno.Sucesso Then
    '                                blnRetorno = True
    '                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                                fsFile.Close()

    '                                GoTo ErrBaixa
    '                            End If

    '                        End If
    '                    End If
    '                End If
    '                sBuffer = fsFile.ReadLine
    '            Loop
    '        End If

    '        fsFile.Close()

    '        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", CDate(Format(CDate(sDataArq), "yyyy-MM-dd")), lblCCorrente.Text, "", connection, Transacao)
    '        If Not _Retorno.Sucesso Then
    '            blnRetorno = True
    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

    '            GoTo ErrBaixa
    '        End If


    '        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
    '        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
    '            blnRetorno = True
    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

    '            GoTo ErrBaixa
    '        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
    '            blnRetorno = True
    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

    '            'For i As Integer = 0 To lstBxAutoErros.Count - 1
    '            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
    '            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '            'Next
    '            'grdTitJaBaixados.DataSource = lstBxAutoErros
    '            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

    '            'If (grdTitJaBaixados.Rows.Count > 0) Then
    '            '    gbxTitJaBaixados.Visible = True
    '            'Else
    '            '    gbxTitJaBaixados.Visible = False
    '            'End If
    '        End If

    '        'EXIBE MESSAGEM
    '        If blnRetorno Then
    '            'ExibirErro()
    '        End If

    '        Transacao.Commit()
    '        connection.Close()

    '        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
    '        'Me.Cursor = Cursors.Default
    '        '------------------------------------

    '        'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
    '        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
    '        Exit Sub

    'ErrBaixa:

    '        Transacao.Rollback()
    '        connection.Close()

    '        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
    '        'Me.Cursor = Cursors.Default
    '        '------------------------------------
    '    End Sub

    '    Private Sub PBaixaTitulos399(nomeArquivo As String) 'Retorno do HSBC

    '        '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
    '        'MsgBox "HSBC ainda não liberado para faturamento. Consultar analista de sistemas !"
    '        'Exit Sub
    '        '---------------------------------------------------------------

    '        Dim iFile As Integer
    '        Dim sBuffer As String, sEvento As String
    '        Dim iCon As Integer
    '        Dim sBanco As String, sNumAviso As String, sNumtit As String, sSeqTit As String, sVcto As String
    '        Dim X As Integer, iNumLcto As Integer
    '        Dim dDataServ As Date, iUltNum As Integer
    '        Dim arrContas() As String
    '        Dim arrValores() As Double
    '        Dim sAgencia As String, sConvenio As String, sConta As String
    '        Dim sCCorrente As String
    '        Dim sDataArq As String = ""
    '        Dim sErros As String = ""

    '        Dim _Retorno As New Retorno
    '        Dim _ControleRetBco As New ControleRetBco
    '        Dim _RetornoBco As New RetornoBco

    '        iFile = FreeFile()

    '        If Dir(nomeArquivo) = "" Then
    '            blnRetorno = True
    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

    '            Exit Sub
    '        End If

    '        'VOLTA O CURSOR NO MOUSE
    '        'Me.Cursor = Cursors.WaitCursor
    '        '---------------------------------

    '        fsFile = New System.IO.StreamReader(nomeArquivo)
    '        sBuffer = fsFile.ReadLine


    '        'CONTROLE DE TRANSACAO
    '        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
    '        connection.Open()
    '        Dim Transacao As SqlTransaction = connection.BeginTransaction()
    '        '-----------------------------------------------------------------------------------------------------------------------------------------------------------

    '        iCon = 1

    '        'Verifica se BO ou DC
    '        If Trim(Mid(sBuffer, 82, 17)) = "DEBITO AUTOMATICO" Then
    '            '*****DEBITO AUTOMATICO*************


    '            'VERIFICA SE EH DO HSBC
    '            sBanco = Mid(sBuffer, 43, 3)
    '            'If sBanco <> codigoBanco Then
    '            '    sBanco = Mid(sBuffer, 43, 3)
    '            '    blnRetorno = True
    '            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '            '    fsFile.Close()

    '            '    GoTo ErrBaixa
    '            'End If

    '            Do While (sBuffer) <> Nothing

    '                'CHECA SE EH HEADER E RETORNO
    '                If Mid(sBuffer, 1, 1) = "A" And Mid(sBuffer, 2, 1) = "2" Then

    '                    If Mid(sBuffer, 2, 1) <> "2" Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                    sBanco = Mid(sBuffer, 43, 3)

    '                    'VERIFICA SE EH DO HSBC
    '                    'If sBanco <> codigoBanco Then
    '                    '    blnRetorno = True
    '                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                    '    fsFile.Close()

    '                    '    GoTo ErrBaixa
    '                    'End If
    '                    sAgencia = CStr(CInt(Mid(sBuffer, 53, 5)))
    '                    sCCorrente = "00" & Strings.Left(CStr(CInt(Mid(sBuffer, 59, 12))), Len(CStr(CInt(Mid(sBuffer, 59, 12)))) - 2) & "-" & Strings.Right(Mid(sBuffer, 59, 12), 2)


    '                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
    '                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)
    '                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sCCorrente, connection, Transacao)
    '                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If


    '                    'CARREGA A ENTIDADE
    '                    _ControleRetBco.CodBanco = sBanco
    '                    _ControleRetBco.NumAviso = sNumAviso
    '                    _ControleRetBco.NumCta = strContaCorrente '"99999"
    '                    sDataArq = Convert.ToDateTime(Mid(sBuffer, 66, 4) & "/" & Mid(sBuffer, 70, 2) & "/" & Mid(sBuffer, 72, 2)).Date
    '                    _ControleRetBco.DtArq = sDataArq

    '                    'INSERE O ARQUIVO DE RETORNO
    '                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
    '                    If Not _Retorno.Sucesso Then
    '                        If _Retorno.MsgErro.ToUpper().Contains("VIOLATION OF PRIMARY KEY") Then
    '                            _Retorno.MsgErro = "Arquivo já baixado no sistema!"
    '                        End If
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                End If


    '                'CHECA SE EH DETALHE DE RETORNO
    '                If Trim(Mid(sBuffer, 1, 1)) = "F" And Trim(Mid(sBuffer, 68, 2)) = "00" Then
    '                    sNumtit = Trim(Mid(sBuffer, 2, 18))
    '                    sSeqTit = Trim(Mid(sBuffer, 20, 2))
    '                    sVcto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 45, 4)) & "/" & Trim(Mid(sBuffer, 49, 2)) & "/" & Trim(Mid(sBuffer, 51, 2))), "yyyy-MM-dd")

    '                    lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
    '                    If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
    '                        'CARREGA ENTIDADE
    '                        _RetornoBco.CodBanco = sBanco
    '                        _RetornoBco.NumAviso = sNumAviso
    '                        _RetornoBco.NumTit = sNumtit
    '                        _RetornoBco.SeqTit = sSeqTit
    '                        _RetornoBco.CodAgen = strAgen 'Trim(Mid(sBuffer, 27, 4))
    '                        _RetornoBco.NumCta = strContaCorrente 'Trim(Str(Val(Mid(sBuffer, 38, 5)))) & Trim(Mid(sBuffer, 43, 2))
    '                        _RetornoBco.VlrPago = Val(Mid(sBuffer, 53, 15)) / 100
    '                        _RetornoBco.VlrJuros = 0
    '                        _RetornoBco.VlrDesc = 0
    '                        _RetornoBco.VlrIOF = 0
    '                        _RetornoBco.VlrAbat = 0
    '                        _RetornoBco.Processado = "N"
    '                        _RetornoBco.DtVcto = sVcto
    '                        _RetornoBco.DtPagto = Mid(sBuffer, 45, 4) & " - " & Mid(sBuffer, 49, 2) & " - " & Mid(sBuffer, 51, 2)
    '                        _RetornoBco.DtArq = sDataArq

    '                        'INSERE RETORNO BANCO
    '                        _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
    '                        If Not _Retorno.Sucesso Then
    '                            blnRetorno = True
    '                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                            fsFile.Close()

    '                            GoTo ErrBaixa
    '                        End If

    '                    End If

    '                ElseIf Trim(Mid(sBuffer, 109, 2)) = "03" Then
    '                    'erro
    '                    'Grava histórico de contato do cliente com o motivo da não baixa do título
    '                    Dim Inconsistencia As New CodInconsistencias
    '                    Dim clBuscaCodInconsistencia As New ConsultaCodInconsistencias
    '                    Dim clInsereErroDadosPgto As New InsereCodInconsistencia
    '                    Dim clInsereHistContato As New InserirHistoricoContato
    '                    Dim clBuscaContasAReceber As New ConsultaContasReceber
    '                    Dim ContasAReceber As New ContaReceber
    '                    Dim blnIsErro As Boolean
    '                    Dim Retorno As New Retorno

    '                    'Busca mensagem de inconsistencia
    '                    Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 319, 2)), "237", blnIsErro, connection, Transacao)

    '                    If Not Inconsistencia.Sucesso Then
    '                        Retorno = Funcoes.RetornoFunc(Inconsistencia.MsgErro)
    '                        lstRetorno.Add(Retorno)
    '                        'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
    '                        blnRetorno = True
    '                        Exit Sub
    '                    End If

    '                    ContasAReceber = clBuscaContasAReceber.ConsultaContaReceberNumTitSeqTit(Strings.Left(Trim(Mid(sBuffer, 117, 10)), Len(Trim(Mid(sBuffer, 117, 10))) - 2), Strings.Right(Trim(Mid(sBuffer, 117, 10)), 2), connection, Transacao)

    '                    If Not ContasAReceber.Sucesso Then
    '                        Retorno = Funcoes.RetornoFunc(ContasAReceber.MsgErro)
    '                        lstRetorno.Add(Retorno)
    '                        'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
    '                        blnRetorno = True
    '                        Exit Sub
    '                    End If

    '                    'Insere no histórico de contato do cliente motivo de não baixa do título
    '                    Retorno = clInsereHistContato.IncluiHistoricoContato(ContasAReceber.CodIntClie, "VERISURE", "Título " & Strings.Left(Trim(Mid(sBuffer, 117, 10)), Len(Trim(Mid(sBuffer, 117, 10))) - 2) & "-" & Strings.Right(Trim(Mid(sBuffer, 117, 10)), 2) & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

    '                    If Not Retorno.Sucesso Then
    '                        lstRetorno.Add(Retorno)
    '                        'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
    '                        blnRetorno = True
    '                        Exit Sub
    '                    End If

    '                    'Caso o código de incosistencia for um erro para assinalar o cliente, realizar update no cliente - Lucas 15/05/2017
    '                    If blnIsErro Then

    '                        Retorno = clInsereErroDadosPgto.InsereErroDadosPgtoCliente(ContasAReceber.CodIntClie, connection, Transacao)

    '                        If Not Retorno.Sucesso Then
    '                            lstRetorno.Add(Retorno)
    '                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
    '                            blnRetorno = True
    '                            Exit Sub
    '                        End If

    '                    End If


    '                    'sBuffer = fsFile.ReadLine

    '                End If

    '                sBuffer = fsFile.ReadLine
    '            Loop
    '        Else
    '            '**********BOLETO******************
    '            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
    '            'MsgBox "SANTANDER (boleto) ainda não liberado para faturamento. Consultar analista de sistemas !"
    '            'Exit Sub
    '            '---------------------------------------------------------------

    '            'VERIFICA SE E DO HSBC
    '            sBanco = Mid(sBuffer, 1, 3)
    '            'If sBanco <> codigoBanco Then
    '            '    blnRetorno = True
    '            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '            '    fsFile.Close()

    '            '    GoTo ErrBaixa
    '            'End If

    '            Do While (sBuffer) <> Nothing
    '                'CHECA SE EH HEADER
    '                If Mid(sBuffer, 8, 1) = "0" Then
    '                    'CHECA SE EH RETORNO
    '                    If Mid(sBuffer, 143, 1) <> "2" Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                    sBanco = Mid(sBuffer, 1, 3)

    '                    'VERIFICA SE E DO HSBC
    '                    'If sBanco <> codigoBanco Then
    '                    '    blnRetorno = True
    '                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                    '    fsFile.Close()

    '                    '    GoTo ErrBaixa
    '                    'End If
    '                    sAgencia = CStr(CInt(Mid(sBuffer, 53, 5)))
    '                    sCCorrente = "00" & Strings.Left(CStr(CInt(Mid(sBuffer, 59, 12))), Len(CStr(CInt(Mid(sBuffer, 59, 12)))) - 2) & "-" & Strings.Right(Mid(sBuffer, 59, 12), 2)

    '                    'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
    '                    'If codigoBanco <> sBanco Or Trim(lblAgen.Text) <> Trim(sAgencia) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(sCCorrente)) Then
    '                    '    blnRetorno = True
    '                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", sBanco).Replace("<%AGENCIA%>", sAgencia).Replace("<%CONTA%>", sCCorrente), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                    '    fsFile.Close()

    '                    '    GoTo ErrBaixa
    '                    'End If


    '                    sDataArq = Format(Convert.ToDateTime(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4)), "yyyy-MM-dd")
    '                    sNumAviso = Strings.Right(Mid(sBuffer, 158, 6), 5)


    '                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
    '                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sCCorrente, connection, Transacao)
    '                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If


    '                    'CARREGA A ENTIDADE
    '                    _ControleRetBco.CodBanco = sBanco
    '                    _ControleRetBco.NumAviso = sNumAviso
    '                    _ControleRetBco.NumCta = sCCorrente
    '                    _ControleRetBco.DtArq = sDataArq

    '                    'INSERE O ARQUIVO DE RETORNO
    '                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
    '                    If Not _Retorno.Sucesso Then
    '                        blnRetorno = True
    '                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                        fsFile.Close()

    '                        GoTo ErrBaixa
    '                    End If

    '                End If


    '                'CHECA SE EH DETALHE
    '                '*********************************************
    '                'DETALHE SEGMENTO T
    '                '*********************************************
    '                If Trim(Mid(sBuffer, 14, 1)) = "T" Then
    '                    'sNumtit = Trim(Mid(sBuffer, 55, 15))
    '                    'sSeqTit = Right(sNumtit, 2)
    '                    'sNumtit = Left(sNumtit, Len(sNumtit) - 2)
    '                    'sVcto = Format(Format(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4)), "dd/mm/yyyy"), "yyyy-mm-dd")

    '                    'CHECA SE EH LIQUIDACAO
    '                    If Mid(sBuffer, 16, 2) = "06" Then 'Or Mid(sBuffer, 16, 2) = "09" Or Mid(sBuffer, 16, 2) = "17" Then

    '                        sNumtit = Trim(Mid(sBuffer, 59, 15))
    '                        sSeqTit = Strings.Right(sNumtit, 2)
    '                        sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
    '                        sVcto = Format(CDate(Trim(Mid(sBuffer, 74, 2)) & "/" & Trim(Mid(sBuffer, 76, 2)) & "/" & Trim(Mid(sBuffer, 78, 4))), "yyyy-MM-dd")


    '                        lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
    '                        If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
    '                            blnRetorno = True
    '                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
    '                            fsFile.Close()

    '                            GoTo ErrBaixa
    '                        Else
    '                            'CARREGA ENTIDADE
    '                            _RetornoBco.CodBanco = sBanco
    '                            _RetornoBco.NumAviso = sNumAviso
    '                            _RetornoBco.NumTit = sNumtit
    '                            _RetornoBco.SeqTit = sSeqTit
    '                            _RetornoBco.CodAgen = strAgen
    '                            _RetornoBco.NumCta = strContaCorrente

    '                            '*********************************************
    '                            'DETALHE SEGMENTO U
    '                            '*********************************************
    '                            sBuffer = fsFile.ReadLine

    '                            If Trim(Mid(sBuffer, 14, 1)) = "U" Then
    '                                _RetornoBco.VlrPago = Val(Mid(sBuffer, 78, 15)) / 100
    '                                _RetornoBco.VlrJuros = (Val(Mid(sBuffer, 18, 15)) / 100) + (Val(Mid(sBuffer, 123, 15)) / 100)
    '                                _RetornoBco.VlrDesc = Val(Mid(sBuffer, 33, 15)) / 100
    '                                _RetornoBco.VlrIOF = 0
    '                                _RetornoBco.VlrAbat = Val(Mid(sBuffer, 48, 15)) / 100
    '                                _RetornoBco.Processado = "N"
    '                                _RetornoBco.DtVcto = sVcto
    '                                _RetornoBco.DtPagto = Format(CDate(Trim(Mid(sBuffer, 138, 2)) & "/" & Trim(Mid(sBuffer, 140, 2)) & "/" & Trim(Mid(sBuffer, 142, 4))), "dd/MM/yyyy")
    '                                _RetornoBco.DtArq = sDataArq

    '                                'INSERE RETORNO BANCO
    '                                _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
    '                                If Not _Retorno.Sucesso Then
    '                                    blnRetorno = True
    '                                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
    '                                    fsFile.Close()

    '                                    GoTo ErrBaixa
    '                                End If
    '                            Else
    '                                blnRetorno = True
    '                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Descricao, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '                                fsFile.Close()

    '                                GoTo ErrBaixa
    '                            End If

    '                        End If

    '                    End If
    '                End If
    '                sBuffer = fsFile.ReadLine
    '            Loop
    '        End If
    '        '**************************************************************************************************************

    '        fsFile.Close()


    '        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(sDataArq), sCCorrente, "", connection, Transacao)
    '        If Not _Retorno.Sucesso Then
    '            blnRetorno = True
    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

    '            GoTo ErrBaixa
    '        End If


    '        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
    '        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
    '            blnRetorno = True
    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

    '            GoTo ErrBaixa
    '        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
    '            blnRetorno = True
    '            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

    '            'For i As Integer = 0 To lstBxAutoErros.Count - 1
    '            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
    '            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
    '            'Next
    '            'grdTitJaBaixados.DataSource = lstBxAutoErros
    '            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

    '            'If (grdTitJaBaixados.Rows.Count > 0) Then
    '            '    gbxTitJaBaixados.Visible = True
    '            'Else
    '            '    gbxTitJaBaixados.Visible = False
    '            'End If
    '        End If

    '        'EXIBE MESSAGEM
    '        If blnRetorno Then
    '            'ExibirErro()
    '        End If

    '        Transacao.Commit()
    '        connection.Close()

    '        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
    '        'Me.Cursor = Cursors.Default
    '        '------------------------------------

    '        'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
    '        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
    '        Exit Sub

    'ErrBaixa:

    '        Transacao.Rollback()
    '        connection.Close()

    '        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
    '        'Me.Cursor = Cursors.Default
    '        '------------------------------------

    '    End Sub

    Private Sub PBaixaTitulos237(nomeArquivo As String) 'Retorno do Banco Bradesco
        Dim iFile As Integer
        Dim sBuffer As String, sEvento As String
        Dim iCon As Integer
        Dim sBanco As String, sNumAviso As String
        Dim X As Integer, iNumLcto As Integer
        Dim dDataServ As Date, iUltNum As Integer
        Dim arrContas() As String
        Dim arrValores() As Double
        Dim sAgencia As String, sErros As String
        Dim sConta As String
        Dim sDataArq As String = ""
        Dim lstControleRetBco As New List(Of ControleRetBco)

        Dim _ControleRetBco As New ControleRetBco
        Dim _Retorno As New Retorno
        Dim _RetornoBco As New RetornoBco


        iFile = FreeFile()

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine


        'CONTROLE DE TRANSACAO
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        connection.Open()
        Dim Transacao As SqlTransaction = connection.BeginTransaction()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------


        sBanco = Trim(Mid(sBuffer, 77, 3))

        Do While (sBuffer) <> Nothing

            If Mid(sBuffer, 1, 1) = "0" Then

                sBanco = Trim(Mid(sBuffer, 77, 3))

            End If

            If Trim(Mid(sBuffer, 109, 2)) = "06" Or Trim(Mid(sBuffer, 109, 2)) = "15" Then
                sAgencia = Trim(Mid(sBuffer, 26, 5))
                sConta = Trim(Str(Val(Mid(sBuffer, 30, 7)))) & "-" & Trim(Mid(sBuffer, 37, 1))

                fsFile.Close()
                Exit Do
            End If

            sBuffer = fsFile.ReadLine
        Loop

        '******************************************************************************************
        fsFile.Close()

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine

        iCon = 1


        Do While (sBuffer) <> Nothing


            If Mid(sBuffer, 1, 1) = "0" Then

                sDataArq = Format(CDate(Mid(sBuffer, 95, 2) & "/" & Mid(sBuffer, 97, 2) & "/" & Mid(sBuffer, 99, 2)), "dd/MM/yyyy")

                lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBcoNumAvisoNumCtaNumCta(Mid(sBuffer, 77, 3), Mid(sBuffer, 109, 5), "22161302", sDataArq, connection, Transacao)

                sBanco = Mid(sBuffer, 77, 3)
                sNumAviso = Mid(sBuffer, 109, 5)


                'CARREGA A ENTIDADE
                _ControleRetBco.CodBanco = sBanco
                _ControleRetBco.NumAviso = sNumAviso
                _ControleRetBco.NumCta = "221613-2"
                _ControleRetBco.DtArq = sDataArq

                'INSERE O ARQUIVO DE RETORNO
                _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)


            End If


            If Trim(Mid(sBuffer, 109, 2)) = "06" Or Trim(Mid(sBuffer, 109, 2)) = "15" Then
                'Se não vier o título, aborta o registro
                'If Trim(Mid(sBuffer, 117, 10)) = "" Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.EXISTE_UM_REGISTRO_SEM_NUMERO_TITULO.Descricao, ErrorConstants.EXISTE_UM_REGISTRO_SEM_NUMERO_TITULO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                'Else

                lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, Strings.Left(Trim(Mid(sBuffer, 117, 10)), Len(Trim(Mid(sBuffer, 117, 10))) - 2), Trim(Mid(sBuffer, 125, 2)), CDate(Trim(Mid(sBuffer, 147, 2)) & "/" & Trim(Mid(sBuffer, 149, 2)) & "/20" & Trim(Mid(sBuffer, 151, 2))), connection, Transacao)
                'If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then

                'If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
                'CARREGA ENTIDADE
                _RetornoBco.CodBanco = sBanco
                _RetornoBco.NumAviso = sNumAviso
                _RetornoBco.NumTit = Strings.Left(Trim(Mid(sBuffer, 117, 10)), Len(Trim(Mid(sBuffer, 117, 10))) - 2)
                _RetornoBco.SeqTit = Strings.Right(Trim(Mid(sBuffer, 117, 10)), 2)
                _RetornoBco.CodAgen = strAgen 'Trim(Mid(sBuffer, 25, 5))
                _RetornoBco.NumCta = strContaCorrente 'Trim(Str(Val(Mid(sBuffer, 30, 7)))) & "-" & Trim(Mid(sBuffer, 37, 1))
                _RetornoBco.VlrPago = (Val(Mid(sBuffer, 254, 13)) / 100)
                _RetornoBco.VlrJuros = Val(Mid(sBuffer, 267, 13)) / 100
                _RetornoBco.VlrDesc = Val(Mid(sBuffer, 241, 13)) / 100
                _RetornoBco.VlrIOF = Val(Mid(sBuffer, 215, 13)) / 100
                _RetornoBco.VlrAbat = Val(Mid(sBuffer, 228, 13)) / 100
                _RetornoBco.Processado = "N"
                _RetornoBco.DtVcto = CDate(Trim(Mid(sBuffer, 147, 2)) & "/" & Trim(Mid(sBuffer, 149, 2)) & "/20" & Trim(Mid(sBuffer, 151, 2)))
                '_RetornoBco.DtPagto = "20" & Mid(sBuffer, 115, 2) & "-" & Mid(sBuffer, 113, 2) & "-" & Mid(sBuffer, 111, 2)
                '_RetornoBco.DtPagto = IIf(Trim(Mid(sBuffer, 319, 2)) = "15", Format(sDataArq, "yyyy-MM-dd"), "20" & Mid(sBuffer, 115, 2) & "-" & Mid(sBuffer, 113, 2) & "-" & Mid(sBuffer, 111, 2))
                _RetornoBco.DtPagto = IIf(Trim(Mid(sBuffer, 319, 2)) = "15", CDate(sDataArq), CDate(Trim(Mid(sBuffer, 111, 2) & "/" & Mid(sBuffer, 113, 2) & "/20" & Mid(sBuffer, 115, 2))))
                _RetornoBco.DtArq = sDataArq

                'INSERE RETORNO BANCO
                _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                'If Not _Retorno.Sucesso Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'End If



                'End If
            ElseIf Trim(Mid(sBuffer, 109, 2)) = "03" Then

                'Grava histórico de contato do cliente com o motivo da não baixa do título
                Dim Inconsistencia As New CodInconsistencias
                Dim clBuscaCodInconsistencia As New ConsultaCodInconsistencias
                Dim clInsereErroDadosPgto As New InsereCodInconsistencia
                Dim clInsereHistContato As New InserirHistoricoContato
                Dim clBuscaContasAReceber As New ConsultaContasReceber
                Dim ContasAReceber As New ContaReceber
                Dim blnIsErro As Boolean
                Dim Retorno As New Retorno

                'Busca mensagem de inconsistencia
                Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 319, 2)), "237", blnIsErro, connection, Transacao)

                ContasAReceber = clBuscaContasAReceber.ConsultaContaReceberNumTitSeqTit(Strings.Left(Trim(Mid(sBuffer, 117, 10)), Len(Trim(Mid(sBuffer, 117, 10))) - 2), Strings.Right(Trim(Mid(sBuffer, 117, 10)), 2), connection, Transacao)

                'Insere no histórico de contato do cliente motivo de não baixa do título
                Retorno = clInsereHistContato.IncluiHistoricoContato(ContasAReceber.CodIntClie, "VERISURE", "Título " & Strings.Left(Trim(Mid(sBuffer, 117, 10)), Len(Trim(Mid(sBuffer, 117, 10))) - 2) & "-" & Strings.Right(Trim(Mid(sBuffer, 117, 10)), 2) & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

                'Caso o código de incosistencia for um erro para assinalar o cliente, realizar update no cliente - Lucas 15/05/2017
                If blnIsErro Then

                    Retorno = clInsereErroDadosPgto.InsereErroDadosPgtoCliente(ContasAReceber.CodIntClie, connection, Transacao)


                End If


                'sBuffer = fsFile.ReadLine

            End If

            sBuffer = fsFile.ReadLine
        Loop

        fsFile.Close()


        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(sDataArq), "221613-2", sAgencia, connection, Transacao)


        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
        'If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
        '    blnRetorno = True
        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)


        'ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
        '    blnRetorno = True
        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
        '    For i As Integer = 0 To lstBxAutoErros.Count - 1
        '        sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
        '        lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
        '    Next
        'End If

        Transacao.Commit()
        connection.Close()


    End Sub

    Private Sub PBaixaTitulos237DD(nomeArquivo As String) 'Retorno do Bradesco por Deposito Identificado

        '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
        'MsgBox "Bradesco DI ainda não liberado para faturamento. Consultar analista de sistemas !"
        'Exit Sub
        '---------------------------------------------------------------

        Dim iFile As Integer
        Dim sBuffer As String, sEvento As String
        Dim iCon As Integer
        Dim sBanco As String, sNumAviso As String, sNumtit As String, sSeqTit As String, sVcto As String
        Dim X As Integer, iNumLcto As Integer
        Dim dDataServ As Date, iUltNum As Integer
        Dim arrContas() As String
        Dim arrValores() As Double
        Dim sAgencia As String, sErros As String, sConta As String, sDataArq As String
        Dim strCodDI As String, sNumCta As String, dValorTitulo As Double, dValorDesc As Double, dValorPago As Double, dValorJuros As Double

        Dim lstCtaAReceber As New List(Of ContaReceber)
        Dim _Retorno As New Retorno
        Dim _RetornoBco As New RetornoBco




        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            Exit Sub
        End If

        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine


        'CONTROLE DE TRANSACAO
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        connection.Open()
        Dim Transacao As SqlTransaction = connection.BeginTransaction()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------

        iCon = 1

        Do While (sBuffer) <> Nothing


            'CHECA SE EH HEADER E RETORNO
            If Mid(sBuffer, 1, 1) = "0" Then
                'If Trim(Mid(sBuffer, 2, 45)) <> "DEPOSITO COM IDENTIFICACAO NUMERICA - DP06" Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_RETORNO_DI.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_RETORNO_DI.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'End If

                sBanco = Mid(sBuffer, 104, 3)

                'VERIFICA SE EH DO BRADESCO
                'If sBanco <> codigoBanco Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'End If

                sNumAviso = Trim(Strings.Right(Mid(sBuffer, 162, 8), 5))
                sDataArq = Format(CDate(Mid(sBuffer, 150, 2) & "/" & Mid(sBuffer, 152, 4) & "/" & Mid(sBuffer, 148, 2)), "yyyy-MM-dd")
                sAgencia = Trim(Mid(sBuffer, 26, 5))
                sNumCta = "221613-2"

                '.OpenTable "Select * FROM CONTROLE_RET_BCO WHERE CodBanco = '" & sBanco & "' and NumAviso = '" & sNumAviso & "'"
                'If Not .Cursor.EOF Then
                '    MsgBox "Este arquivo de retorno já foi atualizado no sistema !", vbExclamation, Me.Caption
                '    Close #iFile
                '    Set clSig = Nothing
                '    Exit Sub
                'End If
                '.Cursor.AddNew
                '.Cursor.Fields("CodBanco") = sBanco
                '.Cursor.Fields("NumAviso") = sNumAviso
                '.Cursor.Fields("NumCta") = "99999"
                '.Cursor.Fields("DtArq") = sDataArq
                '.Cursor.Update
            End If

            'CHECA SE EH DETALHE DE RETORNO -------------------------------------------------------------------------------------------------------------------------------
            If Trim(Mid(sBuffer, 1, 1)) = "1" Then
                strCodDI = CStr(CDbl(Mid(sBuffer, 43, 17))) & "-" & CStr(CDbl(Mid(sBuffer, 60, 2)))

                'PROCURA O DI NA TABELA CONTAS A RECEBER
                lstCtaAReceber = clConsultaCtasAReceber.BuscaNumTitSeqTitDtVctoCodAgenNumCtaVlrIndContasAReceberPorCodDI(strCodDI, connection, Transacao) 'TRATA EXCESSAO
                If Not lstCtaAReceber(0).Sucesso And lstCtaAReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstCtaAReceber(0).MsgErro, lstCtaAReceber(0).NumErro, lstCtaAReceber(0).Sucesso, lstCtaAReceber(0).TipoErro, lstCtaAReceber(0).ImagemErro)
                    fsFile.Close()

                    GoTo ErrBaixa
                ElseIf lstCtaAReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional Then ' TRATA DI INEXISTENTE
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.NAO_ENCONTRADO_TITULO_REFERENTE_CODIGO_DI.Descricao.Replace("<%DI%>", strCodDI), ErrorConstants.NAO_ENCONTRADO_TITULO_REFERENTE_CODIGO_DI.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    'fsFile.Close()
                    'GoTo ErrBaixa
                Else

                    If lstCtaAReceber.Count > 1 Then 'TRATA DI IGUAIS
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.EXISTE_MAIS_TITULOS_REFERENTE_AO_CODIGO_DI.Descricao.Replace("<%DI%>", strCodDI), ErrorConstants.EXISTE_MAIS_TITULOS_REFERENTE_AO_CODIGO_DI.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()
                        GoTo ErrBaixa
                    Else
                        sNumtit = lstCtaAReceber(0).NumTit
                        sSeqTit = lstCtaAReceber(0).SeqTit
                        sVcto = Format(lstCtaAReceber(0).DtVcto, "yyyy-MM-dd")
                        dValorTitulo = lstCtaAReceber(0).VlrInd
                        dValorPago = Val(Mid(sBuffer, 92, 15)) / 100
                        If dValorTitulo - dValorPago > 0 Then
                            dValorDesc = dValorTitulo - dValorPago
                            dValorJuros = 0
                        Else
                            dValorJuros = dValorPago - dValorTitulo
                            dValorDesc = 0
                        End If
                    End If


                    lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                    If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
                        'CARREGA ENTIDADE
                        _RetornoBco.CodBanco = sBanco
                        _RetornoBco.NumAviso = sNumAviso
                        _RetornoBco.NumTit = sNumtit
                        _RetornoBco.SeqTit = sSeqTit
                        _RetornoBco.CodAgen = sAgencia
                        _RetornoBco.NumCta = sNumCta
                        _RetornoBco.VlrPago = dValorPago
                        _RetornoBco.VlrJuros = dValorJuros
                        _RetornoBco.VlrDesc = dValorDesc
                        _RetornoBco.VlrIOF = 0
                        _RetornoBco.VlrAbat = 0
                        _RetornoBco.Processado = "N"
                        _RetornoBco.DtVcto = sVcto
                        _RetornoBco.DtPagto = Format(CDate(Mid(sBuffer, 6, 4) & "-" & Mid(sBuffer, 4, 2) & "-" & Mid(sBuffer, 2, 2)), "yyyy-MM-dd")
                        _RetornoBco.DtArq = sDataArq

                        'INSERE RETORNO BANCO
                        _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                        If Not _Retorno.Sucesso Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        End If
                    End If


                End If
            End If
            '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            sBuffer = fsFile.ReadLine
        Loop

        fsFile.Close()

        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(sDataArq), Trim(sNumCta), sAgencia, connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            GoTo ErrBaixa
        End If


        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

            GoTo ErrBaixa
        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            'For i As Integer = 0 To lstBxAutoErros.Count - 1
            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            'Next
            'grdTitJaBaixados.DataSource = lstBxAutoErros
            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

            'If (grdTitJaBaixados.Rows.Count > 0) Then
            '    gbxTitJaBaixados.Visible = True
            'Else
            '    gbxTitJaBaixados.Visible = False
            'End If
        End If

        'EXIBE MESSAGEM
        If blnRetorno Then
            'ExibirErro()
        End If

        Transacao.Commit()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
        Exit Sub

ErrBaixa:

        Transacao.Rollback()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------
    End Sub
    Private Sub PBaixaTitulos341(nomeArquivo As String) 'Retorno do Banco Itau

        '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
        'MsgBox "ITAU ainda não liberado para faturamento. Consultar analista de sistemas !"
        'Exit Sub
        '---------------------------------------------------------------
        Dim arquivo As String = Path.GetFileName(nomeArquivo)
        Dim CodigoAgenciaTele As String = arquivo.Substring(8, 4)

        Dim iFile As Integer
        Dim sBuffer As String, sEvento As String
        Dim iCon As Integer
        Dim sBanco As String, sNumAviso As String, sNumtit As String, sSeqTit As String, sVcto As String
        Dim X As Integer, iNumLcto As Integer
        Dim dDataServ As Date, iUltNum As Integer, sEmpresa As String, sTaxaMulta As String
        Dim arrContas() As String
        Dim arrValores() As Double, dOutrosCreditos As Double, dDesconto As String
        Dim sAgencia As String, sCCorrente As String, sErros As String, sNumConta As String
        Dim sDataArq As String = ""
        Dim NumCCorrente As String


        Dim MultaM As Double = 0
        Dim JurosM As Double = 0
        Dim vlrInd As Double = 0

        Dim _ControleRetBco As New ControleRetBco
        Dim _Retorno As New Retorno
        Dim _RetornoBco As New RetornoBco
        Dim _Banco As New Banco


        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            Exit Sub
        End If

        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine


        'CONTROLE DE TRANSACAO
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        connection.Open()
        Dim Transacao As SqlTransaction = connection.BeginTransaction()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------

        iCon = 1

        'IDENTIFICA SE LAYOUT EH DC OU BO
        If Trim(Mid(sBuffer, 12, 15)) = "COBRANCA" Or Trim(Mid(sBuffer, 12, 15)) = "EMPRESTIMO" Then
            '*******************************
            '***** BOLETO ******************

            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "ITAU (boleto) ainda não liberado para faturamento. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------


            'VERIFICA SE EH DO ITAU
            sBanco = Mid(sBuffer, 77, 3)
            'If sBanco <> codigoBanco Then
            '    blnRetorno = True
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '    fsFile.Close()

            '    GoTo ErrBaixa
            'End If

            Do While (sBuffer) <> Nothing

                'CHECA SE EH HEADER E RETORNO
                If Mid(sBuffer, 1, 1) = "0" Then

                    '''''CHECA SE EH BX DO RJ
                    sEmpresa = Trim(Mid(sBuffer, 47, 28))
                    If sEmpresa = "TELEATLANTIC RIO MON AL LTDA" Then
                        PBaixaTitulos341RJ(connection, Transacao, nomeArquivo)
                        Exit Sub
                    End If

                    If Mid(sBuffer, 2, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                    sBanco = Mid(sBuffer, 77, 3)
                    'Para diferenciar baixa com mesmo seq da mesma conta DC de BO
                    If Trim(Mid(sBuffer, 12, 15)) = "COBRANCA" Then
                        sCCorrente = "B" & Mid(sBuffer, 33, 6) 'COBRANCA NORMAL
                    Else
                        sCCorrente = "E" & Mid(sBuffer, 33, 6) 'EMPRESTIMO

                    End If
                    sAgencia = Mid(sBuffer, 27, 4)
                    NumCCorrente = Mid(sBuffer, 33, 6)


                    'VERIFICA SE EH DO ITAU
                    'If sBanco <> "341" Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    '    fsFile.Close()

                    '    GoTo ErrBaixa
                    'End If

                    'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
                    'If "341" <> sBanco Or Trim(lblAgen.Text) <> Trim(sAgencia) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(Mid(sBuffer, 33, 6))) Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", sBanco).Replace("<%AGENCIA%>", sAgencia).Replace("<%CONTA%>", Funcoes.FSepara(Trim(Mid(sBuffer, 33, 6)))), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    '    fsFile.Close()

                    '    GoTo ErrBaixa
                    'End If


                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
                    sNumAviso = Mid(sBuffer, 109, 5)
                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, Mid(sBuffer, 33, 6), connection, Transacao)
                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                    sDataArq = Format(CDate(Mid(sBuffer, 95, 2) & "/" & Mid(sBuffer, 97, 2) & "/20" & Mid(sBuffer, 99, 2)), "yyyy-MM-dd")


                    'CARREGA A ENTIDADE
                    _ControleRetBco.CodBanco = sBanco
                    _ControleRetBco.NumAviso = sNumAviso
                    _ControleRetBco.NumCta = sCCorrente
                    _ControleRetBco.DtArq = sDataArq

                    'INSERE O ARQUIVO DE RETORNO
                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                End If

                'CHECA SE EH DETALHE
                If Trim(Mid(sBuffer, 1, 1)) = "1" Then

                    'CHECA SE EH LIQUIDACAO
                    If Mid(sBuffer, 109, 2) = "06" Or Mid(sBuffer, 109, 2) = "10" Then
                        Dim numeroCarteira = Mid(sBuffer, 83, 3)
                        sNumtit = Trim(Mid(sBuffer, 38, 25))
                        sSeqTit = Strings.Right(sNumtit, 2)
                        sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
                        sVcto = Format(CDate((Trim(Mid(sBuffer, 147, 2)) & "/" & Trim(Mid(sBuffer, 149, 2)) & "/20" & Trim(Mid(sBuffer, 151, 2)))), "yyyy-MM-dd")
                        sNumConta = Regex.Replace(sCCorrente, "[^0-9]", "")

                        If numeroCarteira = "109" Then
                            sNumConta = "403109"
                        End If

                        If numeroCarteira = "112" Then
                            sNumConta = "40310_9"
                        End If

                        lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                        If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then

                            'Consulta a taxa parametrizada
                            If Not String.IsNullOrEmpty(sNumConta) Then
                                _Banco = clBuscaBancoContaCorrente.BuscaTaxaJurosParametrizada("341", sAgencia, sNumConta)
                            End If

                            'Consulta para pegar valor original do titulo.
                            contaReceber.VlrInd = consulta.ConsultarValorTituloContaReceber(sNumtit, sSeqTit)

                            If contaReceber.VlrInd > 0 Then
                                vlrInd = contaReceber.VlrInd
                            End If

                            'Verificacao para deduzir credito de origem de desconto sobre titulo descontado
                            dOutrosCreditos = Val(Mid(sBuffer, 280, 13)) / 100
                            dDesconto = Val(Mid(sBuffer, 241, 13)) / 100
                            dOutrosCreditos = IIf(dOutrosCreditos = dDesconto, dOutrosCreditos, 0)
                            sTaxaMulta = _Banco.TaxaMulta / 100

                            If Val(Mid(sBuffer, 267, 13)) = 0 Then 'Caso cliente pague na data correta juros e multa e atribuido 0 
                                JurosM = 0
                                MultaM = 0
                            Else
                                MultaM = Val(vlrInd) * sTaxaMulta 'Calcula o valor da multa com base no valor real do titulo.
                                MultaM = FormatNumber(MultaM, 2)  'Formata número para duas casas decimais
                                If MultaM > Val(Mid(sBuffer, 267, 13)) / 100 Then
                                    MultaM = 0
                                    JurosM = Val(Mid(sBuffer, 267, 13)) / 100
                                Else
                                    JurosM = ((Val(Mid(sBuffer, 267, 13)) / 100) - MultaM)
                                End If
                            End If

                            'CARREGA ENTIDADE
                            _RetornoBco.CodBanco = sBanco
                            _RetornoBco.NumAviso = sNumAviso
                            _RetornoBco.NumTit = sNumtit
                            _RetornoBco.SeqTit = sSeqTit
                            _RetornoBco.CodAgen = strAgen
                            _RetornoBco.NumCta = sNumConta
                            _RetornoBco.VlrPago = (Val(Mid(sBuffer, 254, 13)) / 100) - dOutrosCreditos
                            _RetornoBco.VlrJuros = JurosM
                            _RetornoBco.VlrMulta = MultaM

                            _RetornoBco.VlrDesc = Val(Mid(sBuffer, 241, 13)) / 100
                            _RetornoBco.VlrIOF = Val(Mid(sBuffer, 215, 13)) / 100
                            _RetornoBco.VlrAbat = Val(Mid(sBuffer, 228, 13)) / 100
                            _RetornoBco.Processado = "N"
                            _RetornoBco.DtVcto = sVcto
                            _RetornoBco.DtPagto = "20" & Mid(sBuffer, 115, 2) & "-" & Mid(sBuffer, 113, 2) & "-" & Mid(sBuffer, 111, 2)
                            _RetornoBco.DtArq = sDataArq

                            'INSERE RETORNO BANCO
                            _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                            If Not _Retorno.Sucesso Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                fsFile.Close()

                                GoTo ErrBaixa
                            End If

                        End If

                    ElseIf Mid(sBuffer, 109, 2) = "03" Or Mid(sBuffer, 109, 2) = "15" Or Mid(sBuffer, 109, 2) = "16" Or Mid(sBuffer, 109, 2) = "17" Or Mid(sBuffer, 109, 2) = "18" Then 'Somente entrada rejeitada

                        'Grava histórico de contato do cliente com o motivo da não baixa do título
                        Dim Inconsistencia As New CodInconsistencias
                        Dim clBuscaCodInconsistencia As New ConsultaCodInconsistencias
                        Dim clInsereErroDadosPgto As New InsereCodInconsistencia
                        Dim clInsereHistContato As New InserirHistoricoContato
                        Dim clBuscaContasAReceber As New ConsultaContasReceber
                        Dim ContasAReceber As New ContaReceber
                        Dim blnIsErro As Boolean
                        Dim Retorno As New Retorno

                        sNumtit = Trim(Mid(sBuffer, 38, 25))
                        sSeqTit = Strings.Right(sNumtit, 2)
                        sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)

                        'Busca mensagem de inconsistencia
                        Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 378, 2)), "341", blnIsErro, connection, Transacao)

                        'If Not Inconsistencia.Sucesso Then
                        '    Retorno = Funcoes.RetornoFunc(Inconsistencia.MsgErro)
                        '    lstRetorno.Add(Retorno)
                        '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                        '    blnRetorno = True
                        '    Exit Sub
                        'End If

                        ContasAReceber = clBuscaContasAReceber.ConsultaContaReceberNumTitSeqTit(sNumtit, sSeqTit, connection, Transacao)

                        'If Not ContasAReceber.Sucesso Then
                        '    Retorno = Funcoes.RetornoFunc(ContasAReceber.MsgErro)
                        '    lstRetorno.Add(Retorno)
                        '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                        '    blnRetorno = True
                        '    Exit Sub
                        'End If

                        'Insere no histórico de contato do cliente motivo de não baixa do título
                        Retorno = clInsereHistContato.IncluiHistoricoContato(ContasAReceber.CodIntClie, "VERISURE", "Título " & sNumtit & "-" & sSeqTit & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

                        'If Not Retorno.Sucesso Then
                        '    lstRetorno.Add(Retorno)
                        '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                        '    blnRetorno = True
                        '    Exit Sub
                        'End If

                        'Caso o código de incosistencia for um erro para assinalar o cliente, realizar update no cliente - Lucas 15/05/2017
                        If blnIsErro Then

                            Retorno = clInsereErroDadosPgto.InsereErroDadosPgtoCliente(ContasAReceber.CodIntClie, connection, Transacao)

                            'If Not Retorno.Sucesso Then
                            '    lstRetorno.Add(Retorno)
                            '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            '    blnRetorno = True
                            '    Exit Sub
                            'End If

                        End If

                        'sBuffer = fsFile.ReadLine

                    End If
                End If
                sBuffer = fsFile.ReadLine
            Loop
        Else
            '*******************************
            '***** DEB AUTO ****************
            sBanco = Mid(sBuffer, 1, 3)
            'If sBanco <> "341" Then
            '    blnRetorno = True
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '    fsFile.Close()

            '    GoTo ErrBaixa
            'End If

            Do While (sBuffer) <> Nothing

                'CHECA SE EH HEADER E RETORNO
                If Mid(sBuffer, 8, 1) = "0" And Mid(sBuffer, 143, 1) = "2" Then

                    If Mid(sBuffer, 143, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                    sBanco = Mid(sBuffer, 1, 3)
                    sCCorrente = "40310-9" 'Mid(sBuffer, 33, 10)
                    NumCCorrente = "40310-9"
                    'VERIFICA SE EH DO ITAU
                    'If sBanco <> "341" Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    '    fsFile.Close()

                    '    GoTo ErrBaixa
                    'End If

                    sNumAviso = Strings.Right(Mid(sBuffer, 158, 6), 5)
                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sCCorrente, connection, Transacao)
                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If


                    'CARREGA A ENTIDADE
                    sDataArq = Format(CDate(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4)), "yyyy-MM-dd")
                    _ControleRetBco.CodBanco = sBanco
                    _ControleRetBco.NumAviso = sNumAviso
                    _ControleRetBco.NumCta = Mid(sBuffer, 33, 10)
                    _ControleRetBco.DtArq = sDataArq

                    'INSERE O ARQUIVO DE RETORNO
                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                End If


                'CHECA SE EH DETALHE de liquidacao
                If Trim(Mid(sBuffer, 8, 1)) = "3" And (Trim(Mid(sBuffer, 231, 2)) = "00" Or Trim(Mid(sBuffer, 231, 2)) = "03") Then
                    sNumtit = Trim(Mid(sBuffer, 74, 13))
                    sSeqTit = Trim(Mid(sBuffer, 87, 2))
                    sVcto = Format(CDate(Trim(Mid(sBuffer, 94, 2)) & "/" & Trim(Mid(sBuffer, 96, 2)) & "/" & Trim(Mid(sBuffer, 98, 4))), "yyyy-MM-dd")
                    sCCorrente = Trim(Str(Val(Mid(sBuffer, 37, 5)))) & Trim(Mid(sBuffer, 43, 1))

                    lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                    If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then

                        'CARREGA ENTIDADE
                        _RetornoBco.CodBanco = sBanco
                        _RetornoBco.NumAviso = sNumAviso
                        _RetornoBco.NumTit = sNumtit
                        _RetornoBco.SeqTit = sSeqTit
                        _RetornoBco.CodAgen = strAgen 'Trim(Mid(sBuffer, 25, 4))
                        _RetornoBco.NumCta = strContaCorrente
                        _RetornoBco.VlrPago = Val(Mid(sBuffer, 163, 15)) / 100
                        _RetornoBco.VlrJuros = 0
                        _RetornoBco.VlrDesc = 0
                        _RetornoBco.VlrIOF = 0
                        _RetornoBco.VlrAbat = 0
                        _RetornoBco.Processado = "N"
                        _RetornoBco.DtVcto = sVcto
                        _RetornoBco.DtPagto = Mid(sBuffer, 159, 4) & "-" & Mid(sBuffer, 157, 2) & "-" & Mid(sBuffer, 155, 2)
                        _RetornoBco.DtArq = sDataArq

                        'INSERE RETORNO BANCO
                        _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                        If Not _Retorno.Sucesso Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        End If

                    End If

                ElseIf Not Trim(Mid(sBuffer, 380, 6)) = "" Then

                    'Grava histórico de contato do cliente com o motivo da não baixa do título
                    Dim Inconsistencia As New CodInconsistencias
                    Dim clBuscaCodInconsistencia As New ConsultaCodInconsistencias
                    Dim clInsereErroDadosPgto As New InsereCodInconsistencia
                    Dim clInsereHistContato As New InserirHistoricoContato
                    Dim clBuscaContasAReceber As New ConsultaContasReceber
                    Dim ContasAReceber As New ContaReceber
                    Dim blnIsErro As Boolean
                    Dim Retorno As New Retorno

                    sNumtit = Trim(Mid(sBuffer, 38, 25))
                    sSeqTit = Strings.Right(sNumtit, 2)
                    sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)

                    'Busca mensagem de inconsistencia
                    Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 380, 6)), "341", blnIsErro, connection, Transacao)

                    'If Not Inconsistencia.Sucesso Then
                    '    Retorno = Funcoes.RetornoFunc(Inconsistencia.MsgErro)
                    '    lstRetorno.Add(Retorno)
                    '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                    '    blnRetorno = True
                    '    Exit Sub
                    'End If

                    ContasAReceber = clBuscaContasAReceber.ConsultaContaReceberNumTitSeqTit(sNumtit, sSeqTit, connection, Transacao)

                    'If Not ContasAReceber.Sucesso Then
                    '    Retorno = Funcoes.RetornoFunc(ContasAReceber.MsgErro)
                    '    lstRetorno.Add(Retorno)
                    '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                    '    blnRetorno = True
                    '    Exit Sub
                    'End If

                    'Insere no histórico de contato do cliente motivo de não baixa do título
                    Retorno = clInsereHistContato.IncluiHistoricoContato(ContasAReceber.CodIntClie, "VERISURE", "Título " & sNumtit & "-" & sSeqTit & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

                    'If Not Retorno.Sucesso Then
                    '    lstRetorno.Add(Retorno)
                    '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                    '    blnRetorno = True
                    '    Exit Sub
                    'End If

                    'Caso o código de incosistencia for um erro para assinalar o cliente, realizar update no cliente - Lucas 15/05/2017
                    If blnIsErro Then

                        Retorno = clInsereErroDadosPgto.InsereErroDadosPgtoCliente(ContasAReceber.CodIntClie, connection, Transacao)

                        'If Not Retorno.Sucesso Then
                        '    lstRetorno.Add(Retorno)
                        '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                        '    blnRetorno = True
                        '    Exit Sub
                        'End If

                    End If

                    'sBuffer = fsFile.ReadLine

                End If
                sBuffer = fsFile.ReadLine

            Loop
        End If

        fsFile.Close()

        If String.IsNullOrEmpty(sCCorrente) Then
            sCCorrente = "40310-9"
        End If

        If String.IsNullOrEmpty(sAgencia) Then
            sAgencia = "1608"
        End If

        If CodigoAgenciaTele = "0196" Then
            sAgencia = "0196"
            NumCCorrente = "63932-2"
        End If



        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(sDataArq), NumCCorrente, sAgencia, connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            GoTo ErrBaixa
        End If


        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

            GoTo ErrBaixa
        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            'For i As Integer = 0 To lstBxAutoErros.Count - 1
            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            'Next
            'grdTitJaBaixados.DataSource = lstBxAutoErros
            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

            'If (grdTitJaBaixados.Rows.Count > 0) Then
            '    gbxTitJaBaixados.Visible = True
            'Else
            '    gbxTitJaBaixados.Visible = False
            'End If
        End If

        'EXIBE MESSAGEM
        If blnRetorno Then
            'ExibirErro()
        End If


        Transacao.Commit()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

        'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
        Exit Sub

ErrBaixa:

        Transacao.Rollback()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

    End Sub

    Private Sub PBaixaTitulos341RJ(ByVal connection As SqlConnection, ByVal Transacao As SqlTransaction, nomeArquivo As String) 'Retorno do Banco Itau do Rio de Janeiro

        Dim iFile As Integer
        Dim sBuffer As String, sEvento As String
        Dim iCon As Integer
        Dim sBanco As String, sNumAviso As String
        Dim X As Integer, iNumLcto As Integer
        Dim dDataServ As Date, iUltNum As Integer
        Dim arrContas() As String
        Dim arrValores() As Double
        Dim sAgencia As String, sErros As String
        Dim strContas As String = ""
        Dim sDataArq As String = ""
        Dim sNumtit As String = ""
        Dim sSeqTit As String = ""
        Dim sVcto As String = ""


        Dim _ControleRetBco As New ControleRetBco
        Dim _Retorno As New Retorno
        Dim _RetornoBco As New RetornoBco



        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            Exit Sub
        End If

        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine


        'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
        'If codigoBanco <> Mid(sBuffer, 77, 3) Or Trim(lblAgen.Text) <> Trim(Mid(sBuffer, 27, 4)) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(Mid(sBuffer, 33, 6))) Then
        '    blnRetorno = True
        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", Mid(sBuffer, 77, 3)).Replace("<%AGENCIA%>", Trim(Mid(sBuffer, 27, 4))).Replace("<%CONTA%>", Trim(Mid(sBuffer, 33, 5)) & "-" & Trim(Mid(sBuffer, 38, 1))), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
        '    fsFile.Close()

        '    GoTo ErrBaixa
        'End If

        iCon = 1

        Do While (sBuffer) <> Nothing

            If Mid(sBuffer, 1, 1) = "0" Then

                If Mid(sBuffer, 2, 1) <> "2" Then
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    fsFile.Close()

                    GoTo ErrBaixa
                End If


                'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
                'If "341" <> Mid(sBuffer, 77, 3) Or Trim(lblAgen.Text) <> Trim(Mid(sBuffer, 27, 4)) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(Mid(sBuffer, 33, 6))) Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", Mid(sBuffer, 77, 3)).Replace("<%AGENCIA%>", Trim(Mid(sBuffer, 27, 4))).Replace("<%CONTA%>", Trim(Mid(sBuffer, 33, 5)) & "-" & Trim(Mid(sBuffer, 38, 1))), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'End If


                lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAviso(Mid(sBuffer, 77, 3), Mid(sBuffer, 109, 5), strContaCorrente, connection, Transacao)
                If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                    fsFile.Close()

                    GoTo ErrBaixa
                ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    fsFile.Close()

                    GoTo ErrBaixa
                End If


                sBanco = Mid(sBuffer, 77, 3)
                sNumAviso = Mid(sBuffer, 109, 5)
                sDataArq = Format(CDate(Mid(sBuffer, 95, 2) & "/" & Mid(sBuffer, 97, 2) & "/" & Mid(sBuffer, 99, 2)), "dd/MM/yyyy")

                'CARREGA A ENTIDADE
                _ControleRetBco.CodBanco = sBanco
                _ControleRetBco.NumAviso = sNumAviso
                _ControleRetBco.NumCta = strContaCorrente
                _ControleRetBco.DtArq = sDataArq

                'INSERE O ARQUIVO DE RETORNO
                _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                If Not _Retorno.Sucesso Then
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                    fsFile.Close()

                    GoTo ErrBaixa
                End If

            End If


            If Trim(Mid(sBuffer, 109, 2)) = "06" And Trim(Mid(sBuffer, 83, 3)) = "112" Then

                sNumtit = Mid(Mid(sBuffer, 117, 10), 1, InStr(1, Mid(sBuffer, 117, 10), "/", vbTextCompare) - 1)
                sSeqTit = Mid(Mid(sBuffer, 117, 10), InStr(1, Mid(sBuffer, 117, 10), "/", vbTextCompare) + 1, 2)
                sVcto = Format(CDate(Trim(Mid(sBuffer, 147, 2)) & "/" & Trim(Mid(sBuffer, 149, 2)) & "/20" & Trim(Mid(sBuffer, 151, 2))), "yyyy/MM/dd")


                lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                    fsFile.Close()

                    GoTo ErrBaixa
                ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then

                    'CARREGA ENTIDADE
                    _RetornoBco.CodBanco = sBanco
                    _RetornoBco.NumAviso = Strings.Right(sNumAviso, 5)
                    _RetornoBco.NumTit = sNumtit
                    _RetornoBco.SeqTit = sSeqTit
                    _RetornoBco.CodAgen = strAgen 'Trim(Mid(sBuffer, 18, 4))
                    _RetornoBco.NumCta = strContaCorrente 'Trim(Mid(sBuffer, 24, 5)) & "-" & Trim(Mid(sBuffer, 29, 1))
                    _RetornoBco.VlrPago = (Val(Mid(sBuffer, 254, 13)) / 100)
                    _RetornoBco.VlrJuros = Val(Mid(sBuffer, 267, 13)) / 100
                    _RetornoBco.VlrDesc = Val(Mid(sBuffer, 241, 13)) / 100
                    _RetornoBco.VlrIOF = Val(Mid(sBuffer, 215, 13)) / 100
                    _RetornoBco.VlrAbat = Val(Mid(sBuffer, 228, 13)) / 100
                    _RetornoBco.Processado = "N"
                    _RetornoBco.DtVcto = sVcto
                    _RetornoBco.DtPagto = sDataArq
                    _RetornoBco.DtArq = sDataArq

                    'INSERE RETORNO BANCO
                    _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If
                End If
            End If
            sBuffer = fsFile.ReadLine
        Loop

        fsFile.Close()


        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(Format(sDataArq, "yyyyMMdd")), "E" & Mid(sBuffer, 33, 6), "", connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            GoTo ErrBaixa
        End If


        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

            GoTo ErrBaixa
        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            'For i As Integer = 0 To lstBxAutoErros.Count - 1
            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            'Next
            'grdTitJaBaixados.DataSource = lstBxAutoErros
            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

            'If (grdTitJaBaixados.Rows.Count > 0) Then
            '    gbxTitJaBaixados.Visible = True
            'Else
            '    gbxTitJaBaixados.Visible = False
            'End If
        End If

        'EXIBE MESSAGEM
        'If blnRetorno Then
        '    ExibirErro()
        'End If

        Transacao.Commit()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

        'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
        Exit Sub

ErrBaixa:

        Transacao.Rollback()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

    End Sub

    Private Sub PBaixaTitulos033(nomeArquivo As String) 'SANTANDER

        '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
        'MsgBox "SANTANDER ainda não liberado para faturamento. Consultar analista de sistemas !"
        'Exit Sub
        '---------------------------------------------------------------

        Dim iFile As Integer
        Dim sBuffer As String, sEvento As String
        Dim iCon As Integer
        Dim sBanco As String = ""
        Dim sNumAviso As String = ""
        Dim sNumtit As String = ""
        Dim sSeqTit As String = ""
        Dim sVcto As String = ""
        Dim X As Integer, iNumLcto As Integer
        Dim dDataServ As Date, iUltNum As Integer
        Dim arrContas() As String
        Dim arrValores() As Double
        Dim sAgencia As String
        Dim sConvenio As String = ""
        Dim sConta As String = ""
        Dim sCCorrente As String, sDataArq As String
        Dim sErros As String = ""
        Dim _ControleRetBco As New ControleRetBco
        Dim _Retorno As New Retorno
        Dim _RetornoBco As New RetornoBco
        Dim sDtPgto As String = ""

        Try

            iFile = FreeFile()

            If Dir(nomeArquivo) = "" Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                Exit Sub
            End If

            '***Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado*****************
            'Open Trim(txtArquivo.Text) For Input As #iFile
            'Set clSig = New TOAcesso
            'With clSig
            '    .OpenConnection gsAdo_StrConn
            '        Do While Not EOF(iFile)
            '            Line Input #iFile, sBuffer
            '
            '            If Mid(sBuffer, 1, 1) = "A" And Mid(sBuffer, 2, 1) = "2" Then
            '
            '                If Mid(sBuffer, 2, 1) <> "2" Then
            '                   MsgBox "Este arquivo não é de retorno !", vbExclamation, Me.Caption
            '                   Close #iFile
            '                    Set clSig = Nothing
            '                    Exit Sub
            '                End If
            '
            '                sBanco = Mid(sBuffer, 43, 3)
            '
            '                'VERIFICA SE EH DO SANTANDER
            '                If sBanco <> txtCodBanco.Text Then
            '                    MsgBox "Este arquivo não é do SANTANDER !", vbExclamation, Me.Caption
            '                    Close #iFile
            '                    Set clSig = Nothing
            '                    Exit Sub
            '                End If
            '
            '                sNumAviso = Right(Mid(sBuffer, 74, 6), 5)
            '                .OpenTable "SELECT * FROM CONTROLE_RET_BCO WHERE CodBanco = '" & sBanco & "' and NumAviso = '" & sNumAviso & "'"
            '
            '                If Not .Cursor.EOF Then
            '                   MsgBox "Este arquivo de retorno já foi atualizado no sistema !", vbExclamation, Me.Caption
            '                   Close #iFile
            '                   Set clSig = Nothing
            '                   Exit Sub
            '                End If
            '
            '            End If
            '
            '            If Trim(Mid(sBuffer, 1, 1)) = "F" And Trim(Mid(sBuffer, 68, 2)) = "00" Then
            '                sAgencia = "0319"
            '                sConta = "13007469-5"
            '
            '                If txtCodBanco.Text <> sBanco Or CInt(Trim(strAgencia)) <> CInt(sAgencia) Or Trim(FSepara(strContas)) <> Trim(FSepara(sConta)) Then
            '                   MsgBox "Arquivo de Retorno Banco/Agência/Conta Corrente : [" & sBanco & "/" & sAgencia & "/" & sConta & "] não corresponde ao selecionado !!", vbExclamation, Me.Caption
            '                   Close #iFile
            '                   Set clSig = Nothing
            '                   Exit Sub
            '                End If
            '                Close #iFile
            '                Exit Do
            '            End If
            '        Loop
            'End With
            'Close #iFile
            '******************************************************************************************


            'VOLTA O CURSOR NO MOUSE
            'Me.Cursor = Cursors.WaitCursor
            '---------------------------------

            fsFile = New System.IO.StreamReader(nomeArquivo)
            sBuffer = fsFile.ReadLine


            'CONTROLE DE TRANSACAO
            Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
            connection.Open()
            Dim Transacao As SqlTransaction = connection.BeginTransaction()
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------

            'Verifica se BO ou DC
            If Trim(Mid(sBuffer, 82, 17)) = "DEBITO AUTOMATICO" Then

                '*****DEBITO AUTOMATICO*************
                'VERIFICA SE EH DO SANTANDER
                sBanco = Mid(sBuffer, 43, 3)
                'If sBanco <> codigoBanco Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'End If

                Do While (sBuffer) <> Nothing

                    'CHECA SE EH HEADER E RETORNO
                    If Mid(sBuffer, 1, 1) = "A" And Mid(sBuffer, 2, 1) = "2" Then

                        If Mid(sBuffer, 2, 1) <> "2" Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                            fsFile.Close()

                            GoTo ErrBaixa
                        End If


                        sBanco = Mid(sBuffer, 43, 3)

                        'VERIFICA SE EH DO SANTANDER
                        'If sBanco <> codigoBanco Then
                        '    blnRetorno = True
                        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        '    fsFile.Close()

                        '    GoTo ErrBaixa
                        'End If

                        'O ARQUIVO DE DEBITO DE AUTOMATICO DO SANTANDER NAO POSSUI INFORMAÇÂO DA CONTA E AGENCIA, RETIRADO A VALIDACAO ABAIXO (Hercules 06/01/2015)
                        'sAgencia = "0319"
                        'sConta = "13007469-5"

                        ''Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
                        'If codigoBanco <> sBanco Or CInt(Trim(lblAgen.Text)) <> CInt(sAgencia) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(sConta)) Then
                        '    blnRetorno = True
                        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", sBanco).Replace("<%AGENCIA%>", sAgencia).Replace("<%CONTA%>", sConta), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        '    fsFile.Close()

                        '    GoTo ErrBaixa
                        'End If

                        sAgencia = Trim("0900")
                        sConta = Trim("13000168-2")

                        sConvenio = Trim(Mid(sBuffer, 3, 20))
                        sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)


                        'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
                        lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sConta, connection, Transacao)
                        If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                            fsFile.Close()

                            GoTo ErrBaixa
                        End If


                        'CARREGA A ENTIDADE
                        _ControleRetBco.CodBanco = sBanco
                        _ControleRetBco.NumAviso = sNumAviso
                        _ControleRetBco.NumCta = sConta 'SANTANDER nao possui num. conta no header
                        sDataArq = Format(CDate(Mid(sBuffer, 66, 4) & "/" & Mid(sBuffer, 70, 2) & "/" & Mid(sBuffer, 72, 2)), "yyyy-MM-dd")
                        _ControleRetBco.DtArq = sDataArq

                        'INSERE O ARQUIVO DE RETORNO
                        _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                        If Not _Retorno.Sucesso Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        End If

                    End If


                    'CHECA SE EH DETALHE DE RETORNO
                    If Trim(Mid(sBuffer, 1, 1)) = "F" Then

                        'Lucas - 13/11/2018
                        'Verifica se o registro está de acordo com novo layout
                        If String.IsNullOrEmpty(Trim(Mid(sBuffer, 120, 8)) + Trim(Mid(sBuffer, 128, 2))) Then
                            sBuffer = fsFile.ReadLine
                            Continue Do
                        End If

                        If Trim(Mid(sBuffer, 68, 2)) = "00" Then

                            'Lucas - 17/10/2018
                            'Alterado posições para ler o numtit e seqtit do cliente de acordo com novo layout
                            'sNumtit = Trim(Mid(sBuffer, 2, 23))
                            'sSeqTit = Trim(Mid(sBuffer, 25, 2))
                            sNumtit = Trim(Mid(sBuffer, 120, 8))
                            sSeqTit = Trim(Mid(sBuffer, 128, 2))

                            sDtPgto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 45, 4)) & "/" & Trim(Mid(sBuffer, 49, 2)) & "/" & Trim(Mid(sBuffer, 51, 2))), "yyyy-MM-dd")
                            sVcto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 70, 4)) & "/" & Trim(Mid(sBuffer, 74, 2)) & "/" & Trim(Mid(sBuffer, 76, 2))), "yyyy-MM-dd")

                            lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                            If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                                fsFile.Close()

                                GoTo ErrBaixa
                            ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
                                'CARREGA ENTIDADE
                                _RetornoBco.CodBanco = sBanco
                                _RetornoBco.NumAviso = sNumAviso
                                _RetornoBco.NumTit = sNumtit
                                _RetornoBco.SeqTit = sSeqTit
                                _RetornoBco.CodAgen = strAgen  'SANTANDER nao possui num. conta no header
                                _RetornoBco.NumCta = strContaCorrente 'SANTANDER nao possui num. conta no header
                                _RetornoBco.VlrPago = Val(Mid(sBuffer, 53, 15)) / 100
                                _RetornoBco.VlrJuros = 0
                                _RetornoBco.VlrDesc = 0
                                _RetornoBco.VlrIOF = 0
                                _RetornoBco.VlrAbat = 0
                                _RetornoBco.Processado = "N"
                                _RetornoBco.DtVcto = sVcto
                                _RetornoBco.DtPagto = sDtPgto
                                _RetornoBco.DtArq = sDataArq

                                'INSERE RETORNO BANCO
                                _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                                If Not _Retorno.Sucesso Then
                                    blnRetorno = True
                                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                    fsFile.Close()

                                    GoTo ErrBaixa
                                End If
                            End If

                        ElseIf isErroSantander(Trim(Mid(sBuffer, 68, 2))) Then

                            'Grava histórico de contato do cliente com o motivo da não baixa do título
                            Dim Inconsistencia As New CodInconsistencias
                            Dim clBuscaCodInconsistencia As New ConsultaCodInconsistencias
                            Dim clInsereErroDadosPgto As New InsereCodInconsistencia
                            Dim clInsereHistContato As New InserirHistoricoContato
                            'Dim clBuscaContasAReceber As New Teleatlantic.TLS.ContasAReceberBC.ConsultaContasReceber
                            'Dim ContasAReceber As New ContaReceber
                            Dim blnIsErro As Boolean
                            Dim Retorno As New Retorno

                            'Lucas - 17/10/2018
                            'Alterado posições para ler o numtit e seqtit do cliente de acordo com novo layout
                            'sNumtit = Trim(Mid(sBuffer, 2, 23))
                            'sSeqTit = Trim(Mid(sBuffer, 25, 2))
                            sNumtit = Trim(Mid(sBuffer, 120, 8))
                            sSeqTit = Trim(Mid(sBuffer, 128, 2))

                            'Lucas - 17/10/2018
                            'Busca mensagem de inconsistencia referente ao código 04 de acordo com tabela (mais de uma opção para este tipo de erro 04)
                            If Trim(Mid(sBuffer, 68, 2)).Equals("04") Then
                                'ou Trim(Mid(sBuffer, 130, 2)
                                Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistencias04(Trim(Mid(sBuffer, 147, 2)), "033", blnIsErro, connection, Transacao)
                            Else
                                'Busca mensagem de inconsistencia
                                Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 68, 2)), "033", blnIsErro, connection, Transacao)
                            End If


                            If Not Inconsistencia.Sucesso Then

                                Retorno = Funcoes.RetornoFunc(Inconsistencia.MsgErro)
                                lstRetorno.Add(Retorno)
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                fsFile.Close()

                                GoTo ErrBaixa

                            End If

                            Dim CodIntClie As String = Trim(Mid(sBuffer, 20, 7))

                            'ContasAReceber = clBuscaContasAReceber.ConsultaContaReceberNumTitSeqTit(sNumtit, sSeqTit, connection, Transacao)

                            'If Not ContasAReceber.Sucesso Then
                            '    Retorno = Funcoes.RetornoFunc(ContasAReceber.MsgErro)
                            '    lstRetorno.Add(Retorno)
                            '    FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            '    blnRetorno = True
                            '    Exit Sub
                            'End If

                            'Insere no histórico de contato do cliente motivo de não baixa do título
                            Retorno = clInsereHistContato.IncluiHistoricoContato(CodIntClie, "VERISURE", "Título " & sNumtit & "-" & sSeqTit & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

                            If Not Retorno.Sucesso Then

                                lstRetorno.Add(Retorno)
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                fsFile.Close()

                                GoTo ErrBaixa

                            End If

                            'Caso o código de incosistencia for um erro para assinalar o cliente, realizar update no cliente - Lucas 15/05/2017
                            If blnIsErro Then

                                Retorno = clInsereErroDadosPgto.InsereErroDadosPgtoCliente(CodIntClie, connection, Transacao)

                                If Not Retorno.Sucesso Then

                                    lstRetorno.Add(Retorno)
                                    blnRetorno = True
                                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                    fsFile.Close()

                                    GoTo ErrBaixa

                                End If

                            End If

                            'sBuffer = fsFile.ReadLine

                        End If

                    End If
                    sBuffer = fsFile.ReadLine
                Loop
            Else
                '**********BOLETO******************
                '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
                'MsgBox "SANTANDER (boleto) ainda não liberado para faturamento. Consultar analista de sistemas !"
                'Exit Sub
                '---------------------------------------------------------------

                'VERIFICA SE EH DO SANTANDER
                sBanco = Mid(sBuffer, 1, 3)
                'If sBanco <> codigoBanco Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'End If

                Do While (sBuffer) <> Nothing
                    'CHECA SE EH HEADER
                    If Mid(sBuffer, 8, 1) = "0" Then
                        'CHECA SE EH RETORNO
                        If Mid(sBuffer, 143, 1) <> "2" Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                            fsFile.Close()

                            GoTo ErrBaixa
                        End If

                        sBanco = Mid(sBuffer, 1, 3)

                        'VERIFICA SE EH DO SANTANDER
                        'If sBanco <> codigoBanco Then
                        '    blnRetorno = True
                        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        '    fsFile.Close()

                        '    GoTo ErrBaixa
                        'End If
                        sAgencia = Mid(sBuffer, 33, 4)
                        sCCorrente = Mid(sBuffer, 39, 9)

                        'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
                        'If codigoBanco <> sBanco Or Trim(lblAgen.Text) <> Trim(sAgencia) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(sCCorrente)) Then
                        '    blnRetorno = True
                        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", sBanco).Replace("<%AGENCIA%>", sAgencia).Replace("<%CONTA%>", sCCorrente), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        '    fsFile.Close()

                        '    GoTo ErrBaixa
                        'End If

                        sDataArq = Format(CDate(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4)), "yyyy-MM-dd")
                        sNumAviso = Strings.Right(Mid(sBuffer, 158, 6), 5)

                        'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
                        lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sCCorrente, connection, Transacao)
                        If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                            fsFile.Close()

                            GoTo ErrBaixa
                        End If


                        'CARREGA A ENTIDADE
                        _ControleRetBco.CodBanco = sBanco
                        _ControleRetBco.NumAviso = sNumAviso
                        _ControleRetBco.NumCta = sCCorrente 'SANTANDER nao possui num. conta no header
                        _ControleRetBco.DtArq = sDataArq

                        'INSERE O ARQUIVO DE RETORNO
                        _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                        If Not _Retorno.Sucesso Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        End If

                    End If


                    'CHECA SE EH DETALHE
                    '*********************************************
                    'DETALHE SEGMENTO T
                    '*********************************************
                    If Trim(Mid(sBuffer, 14, 1)) = "T" Then
                        'sNumtit = Trim(Mid(sBuffer, 55, 15))
                        'sSeqTit = Right(sNumtit, 2)
                        'sNumtit = Left(sNumtit, Len(sNumtit) - 2)
                        'sVcto = Format(Format(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4)), "dd/mm/yyyy"), "yyyy-mm-dd")

                        'CHECA SE EH LIQUIDACAO
                        If Mid(sBuffer, 16, 2) = "06" Then 'Or Mid(sBuffer, 16, 2) = "09" Or Mid(sBuffer, 16, 2) = "17" Then

                            sNumtit = Trim(Mid(sBuffer, 55, 15))
                            sSeqTit = Strings.Right(sNumtit, 2)
                            sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
                            sVcto = Format(CDate(Trim(Mid(sBuffer, 70, 2)) & "/" & Trim(Mid(sBuffer, 72, 2)) & "/" & Trim(Mid(sBuffer, 74, 4))), "yyyy-MM-dd")

                            lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                            If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                                fsFile.Close()

                                GoTo ErrBaixa
                            Else
                                'CARREGA ENTIDADE
                                _RetornoBco.CodBanco = sBanco
                                _RetornoBco.NumAviso = sNumAviso
                                _RetornoBco.NumTit = sNumtit
                                _RetornoBco.SeqTit = sSeqTit
                                _RetornoBco.CodAgen = strAgen
                                _RetornoBco.NumCta = strContaCorrente

                                sBuffer = fsFile.ReadLine

                                If Trim(Mid(sBuffer, 14, 1)) = "U" Then
                                    _RetornoBco.VlrPago = Val(Mid(sBuffer, 78, 15)) / 100
                                    _RetornoBco.VlrJuros = (Val(Mid(sBuffer, 18, 15)) / 100) + (Val(Mid(sBuffer, 123, 15)) / 100)
                                    _RetornoBco.VlrDesc = Val(Mid(sBuffer, 33, 15)) / 100
                                    _RetornoBco.VlrIOF = 0
                                    _RetornoBco.VlrAbat = Val(Mid(sBuffer, 48, 15)) / 100
                                    _RetornoBco.Processado = "N"
                                    _RetornoBco.DtVcto = sVcto
                                    _RetornoBco.DtPagto = Format(CDate(Trim(Mid(sBuffer, 138, 2)) & "/" & Trim(Mid(sBuffer, 140, 2)) & "/" & Trim(Mid(sBuffer, 142, 4))), "dd/MM/yyyy")
                                    _RetornoBco.DtArq = sDataArq

                                    'INSERE RETORNO BANCO
                                    _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                                    If Not _Retorno.Sucesso Then
                                        blnRetorno = True
                                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                        fsFile.Close()

                                        GoTo ErrBaixa
                                    End If
                                Else
                                    blnRetorno = True
                                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Descricao, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                                    fsFile.Close()

                                    GoTo ErrBaixa
                                End If

                            End If

                        End If
                    End If
                    sBuffer = fsFile.ReadLine
                Loop
            End If
            '**************************************************************************************************************
            fsFile.Close()


            _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(sDataArq), Trim("13000168-2"), sAgencia, connection, Transacao)
            If Not _Retorno.Sucesso Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

                GoTo ErrBaixa
            End If


            lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
            If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

                GoTo ErrBaixa
            ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
                blnRetorno = True
                lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                'For i As Integer = 0 To lstBxAutoErros.Count - 1
                '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                'Next
                'grdTitJaBaixados.DataSource = lstBxAutoErros
                'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

                'If (grdTitJaBaixados.Rows.Count > 0) Then
                '    gbxTitJaBaixados.Visible = True
                'Else
                '    gbxTitJaBaixados.Visible = False
                'End If
            End If

            'EXIBE MESSAGEM
            'If blnRetorno Then
            '    ExibirErro()
            'End If

            Transacao.Commit()
            connection.Close()

            'MUDA O CURSOR DO MOUSE PARA AMPULHETA
            'Me.Cursor = Cursors.Default
            '------------------------------------

            'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
            'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
            Exit Sub

ErrBaixa:

            Transacao.Rollback()
            connection.Close()

            'MUDA O CURSOR DO MOUSE PARA AMPULHETA
            'Me.Cursor = Cursors.Default
            '------------------------------------
        Catch ex As Exception
            'FuncoesUI.TrataException(ex, True, System.Reflection.MethodBase.GetCurrentMethod(), "PBaixaTitulos033")
        End Try
    End Sub

    Private Function isErroSantander(ByVal CodInconsistencia As String) As Boolean

        Select Case CodInconsistencia
            Case "01"
                Return True
            Case "02"
                Return True
            Case "04"
                Return True
            Case "10"
                Return True
            Case "12"
                Return True
            Case "13"
                Return True
            Case "14"
                Return True
            Case "15"
                Return True
            Case "18"
                Return True
            Case "30"
                Return True
            Case "96"
                Return True
            Case "97"
                Return True
            Case "98"
                Return True
            Case "99"
                Return True
            Case Else
                Return False
        End Select

    End Function
    Private Sub PBaixaTitulos341DD(nomeArquivo As String) 'Retorno do Itaú por Depósito Identificado

        '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
        'MsgBox "Itaú DI ainda não liberado para faturamento. Consultar analista de sistemas !"
        'Exit Sub
        '---------------------------------------------------------------

        Dim iFile As Integer, iCon As Integer, iCountDi As Integer = 0, iCountRegistros As Integer = 0
        Dim sBanco As String, sAgencia As String, sNumCta As String
        Dim sNumtit As String, sSeqTit As String, sVcto As String
        Dim sNumAviso As String, sErros As String, strCodDI As String, sBuffer As String
        Dim dValorTitulo As Double, dValorDesc As Double, dValorPago As Double, dValorJuros As Double
        Dim sDataArq As String = ""

        Dim lstCtaAReceber As New List(Of ContaReceber)
        Dim _RetornoBco As New RetornoBco
        Dim _Retorno As New Retorno


        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            fsFile.Close()

            GoTo ErrBaixa
        End If


        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine


        'CONTROLE DE TRANSACAO
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        connection.Open()
        Dim Transacao As SqlTransaction = connection.BeginTransaction()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------

        iCon = 1


        sBanco = Mid(sBuffer, 91, 3)

        'VERIFICA SE EH DO ITAÚ
        'If sBanco <> "341" Then
        '    blnRetorno = True
        '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
        '    fsFile.Close()

        '    GoTo ErrBaixa
        'End If


        Do While (sBuffer) <> Nothing


            'CHECA SE EH HEADER E RETORNO
            If Mid(sBuffer, 14, 2) = "00" Then
                If Trim(Mid(sBuffer, 94, 2)) <> "06" Then
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO_DE_DI.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO_DE_DI.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    fsFile.Close()

                    GoTo ErrBaixa
                End If

                sBanco = Mid(sBuffer, 91, 3)

                'VERIFICA SE EH DO ITAÚ
                'If sBanco <> "341" Then
                '    blnRetorno = True
                '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                '    fsFile.Close()

                '    GoTo ErrBaixa
                'End If
            End If

            '' Checa se é o trailer do arquivo e quantidade de DIs processados é coerente com o que foi escrito no arquivo. -- Eduardo 
            If Trim(Mid(sBuffer, 14, 2)) = "99" Then
                iCountRegistros = Convert.ToInt32(Mid(sBuffer, 79, 6))
            End If

            'CHECA SE EH DETALHE DE RETORNO
            If Trim(Mid(sBuffer, 14, 2)) = "10" Then
                strCodDI = CStr(CDbl(Mid(sBuffer, 35, 15))) & "-" & CStr(CDbl(Mid(sBuffer, 50, 1)))

                'PROCURA O DI NA TABELA CONTAS A RECEBER
                lstCtaAReceber = clConsultaCtasAReceber.BuscaNumTitSeqTitDtVctoCodAgenNumCtaVlrIndContasAReceberPorCodDI(strCodDI, connection, Transacao) 'TRATA EXCEÇÃO
                If Not lstCtaAReceber(0).Sucesso And lstCtaAReceber(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstCtaAReceber(0).MsgErro, lstCtaAReceber(0).NumErro, lstCtaAReceber(0).Sucesso, lstCtaAReceber(0).TipoErro, lstCtaAReceber(0).ImagemErro)
                    fsFile.Close()

                    GoTo ErrBaixa
                ElseIf lstCtaAReceber(0).TipoErro = DadosGenericos.TipoErro.Funcional Then ' TRATA DI INEXISTENTE
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.NAO_ENCONTRADO_TITULO_REFERENTE_CODIGO_DI.Descricao.Replace("<%DI%>", strCodDI), ErrorConstants.NAO_ENCONTRADO_TITULO_REFERENTE_CODIGO_DI.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    GoTo NextLine
                Else
                    If lstCtaAReceber.Count > 1 Then 'TRATA DI IGUAIS
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.EXISTE_MAIS_TITULOS_REFERENTE_AO_CODIGO_DI.Descricao.Replace("<%DI%>", strCodDI), ErrorConstants.EXISTE_MAIS_TITULOS_REFERENTE_AO_CODIGO_DI.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

                    Else
                        sNumAviso = Mid(sBuffer, 26, 3)
                        sDataArq = Format(CDate("20" & Mid(sBuffer, 29, 2) & "-" & Mid(sBuffer, 31, 2) & "-" & Mid(sBuffer, 33, 2)), "yyyy-MM-dd")
                        sAgencia = sAgencia = Mid(sBuffer, 27, 4)
                        sNumCta = "221613-2"

                        sNumtit = lstCtaAReceber(0).NumTit
                        sSeqTit = lstCtaAReceber(0).SeqTit
                        sVcto = Format(lstCtaAReceber(0).DtVcto, "yyyy-MM-dd")
                        dValorTitulo = lstCtaAReceber(0).VlrInd
                        dValorPago = Val(Mid(sBuffer, 55, 17)) / 100
                        If dValorTitulo - dValorPago > 0 Then
                            dValorDesc = dValorTitulo - dValorPago
                            dValorJuros = 0
                        Else
                            dValorJuros = dValorPago - dValorTitulo
                            dValorDesc = 0
                        End If

                        '' Soma 1 no contador de DIs processados.
                        iCountDi += 1
                    End If

                End If

                lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                    blnRetorno = True
                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                    fsFile.Close()

                    GoTo ErrBaixa
                ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then

                    'CARREGA ENTIDADE
                    _RetornoBco.CodBanco = sBanco
                    _RetornoBco.NumAviso = sNumAviso
                    _RetornoBco.NumTit = sNumtit
                    _RetornoBco.SeqTit = sSeqTit
                    _RetornoBco.CodAgen = sAgencia
                    _RetornoBco.NumCta = sNumCta
                    _RetornoBco.VlrPago = dValorPago
                    _RetornoBco.VlrJuros = dValorJuros
                    _RetornoBco.VlrDesc = dValorDesc
                    _RetornoBco.VlrIOF = 0
                    _RetornoBco.VlrAbat = 0
                    _RetornoBco.Processado = "N"
                    _RetornoBco.DtVcto = sVcto
                    _RetornoBco.DtPagto = Format(CDate("20" & Mid(sBuffer, 16, 2) & "-" & Mid(sBuffer, 18, 2) & "-" & Mid(sBuffer, 20, 2)), "yyyy-MM-dd")
                    _RetornoBco.DtArq = sDataArq

                    'INSERE RETORNO BANCO
                    _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                End If

            End If
NextLine:
            sBuffer = fsFile.ReadLine
        Loop

        fsFile.Close()

        If (iCountDi <> iCountRegistros) Then
            _Retorno = New Retorno()
            _Retorno.Sucesso = False
            _Retorno.TipoErro = DadosGenericos.TipoErro.Funcional
            _Retorno.NumErro = ErrorConstants.NAO_FORAM_PROCESSADOS_X_DI.Id
            _Retorno.MsgErro = ErrorConstants.NAO_FORAM_PROCESSADOS_X_DI.Descricao.Replace("<%X%>", (iCountRegistros - iCountDi).ToString())
            _Retorno.ImagemErro = DadosGenericos.ImagemRetorno.Alerta
            lstRetorno.Add(_Retorno)
        End If

        '' Significa que não achou nenhum DI válido e não registrou nenhuma data.
        If (String.IsNullOrWhiteSpace(sDataArq)) OrElse iCountRegistros <= 0 Then GoTo ErrBaixa

        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", Convert.ToDateTime(sDataArq), sNumCta, sAgencia, connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            GoTo ErrBaixa
        End If


        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

            GoTo ErrBaixa
        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            'For i As Integer = 0 To lstBxAutoErros.Count - 1
            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            'Next
            'grdTitJaBaixados.DataSource = lstBxAutoErros
            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

            'If (grdTitJaBaixados.Rows.Count > 0) Then
            '    gbxTitJaBaixados.Visible = True
            'Else
            '    gbxTitJaBaixados.Visible = False
            'End If
        End If

        'EXIBE MESSAGEM
        If blnRetorno Then
            'ExibirErro()
        End If

        Transacao.Commit()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

        'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
        Exit Sub

ErrBaixa:

        Transacao.Rollback()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

    End Sub

    Private Sub PBaixaTitulos001(nomeArquivo As String) 'Retorno do BANCO DO BRASIL
        Dim iFile As Integer
        Dim sBuffer As String, sEvento As String
        Dim iCon As Integer
        Dim sBanco As String, sNumAviso As String, sNumtit As String, sSeqTit As String, sVcto As String
        Dim X As Integer, iNumLcto As Integer
        Dim dDataServ As Date, iUltNum As Integer
        Dim arrContas() As String
        Dim arrValores() As Double
        Dim sAgencia As String, sErros As String, sConta As String
        Dim sDataArq As DateTime = Date.MinValue
        Dim _ControleRetBco As New ControleRetBco
        Dim _Retorno As New Retorno
        Dim _RetornoBco As New RetornoBco


        iFile = FreeFile()

        If Dir(nomeArquivo) = "" Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Descricao, ErrorConstants.O_ARQUIVO_ESCOLHIDO_NAO_EXISTE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            Exit Sub
        End If

        '***Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado*****************

        'VOLTA O CURSOR NO MOUSE
        'Me.Cursor = Cursors.WaitCursor
        '---------------------------------

        fsFile = New System.IO.StreamReader(nomeArquivo)
        sBuffer = fsFile.ReadLine


        'CONTROLE DE TRANSACAO
        Dim connection As SqlConnection = New SqlConnection(Conection.STRING_CONEXAO.StringConexao)
        connection.Open()
        Dim Transacao As SqlTransaction = connection.BeginTransaction()
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------


        iCon = 1
        'Verifica se BO ou DC
        If Trim(Mid(sBuffer, 82, 17)) = "DEBITO AUTOMATICO" Then
            '*****DEBITO AUTOMATICO*************
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "BANCO DO BRASIL (Débito Automático) ainda não liberado para baixa automática. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------

            'VERIFICA SE EH DO BANCO DO BRASIL
            sBanco = Mid(sBuffer, 43, 3)
            'If sBanco <> codigoBanco Then
            '    blnRetorno = True
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '    fsFile.Close()

            '    GoTo ErrBaixa
            'End If

            Do While (sBuffer) <> Nothing

                If Mid(sBuffer, 1, 1) = "A" And Mid(sBuffer, 2, 1) = "2" Then

                    If Mid(sBuffer, 2, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                    sBanco = Mid(sBuffer, 43, 3)

                    'VERIFICA SE EH DO BANCO DO BRASIL
                    'If sBanco <> codigoBanco Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    '    fsFile.Close()

                    '    GoTo ErrBaixa
                    'End If

                    sNumAviso = Strings.Right(Mid(sBuffer, 74, 6), 5)


                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, "22267-4", connection, Transacao) 'TRATAMENTO DE ERRO
                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                    'CARREGA A ENTIDADE
                    _ControleRetBco.CodBanco = sBanco
                    _ControleRetBco.NumAviso = sNumAviso
                    _ControleRetBco.NumCta = strContaCorrente
                    sDataArq = Format(Convert.ToDateTime(Mid(sBuffer, 66, 4) & "/" & Mid(sBuffer, 70, 2) & "/" & Mid(sBuffer, 72, 2)), "yyyy-MM-dd")
                    _ControleRetBco.DtArq = sDataArq

                    'INSERE O ARQUIVO DE RETORNO
                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                End If

                'CHECA SE EH DETALHE DE RETORNO
                If Trim(Mid(sBuffer, 1, 1)) = "F" Then

                    If Trim(Mid(sBuffer, 68, 2)) = "00" Or Trim(Mid(sBuffer, 68, 2)) = "31" Then

                        'sNumtit = Strings.Left(Trim(Mid(sBuffer, 2, 25)), Len(Trim(Mid(sBuffer, 2, 25))) - 2)
                        'sSeqTit = Strings.Right(Trim(Mid(sBuffer, 2, 25)), 2)

                        ''Liberar quenado Banco do Brasil for testado
                        sNumtit = Strings.Left(Trim(Mid(sBuffer, 70, 25)), Len(Trim(Mid(sBuffer, 70, 25))) - 2)
                        sSeqTit = Strings.Right(Trim(Mid(sBuffer, 70, 25)), 2)

                        sVcto = Format(Convert.ToDateTime(Trim(Mid(sBuffer, 122, 4)) & "/" & Trim(Mid(sBuffer, 126, 2)) & "/" & Trim(Mid(sBuffer, 128, 2))), "yyyy-MM-dd")

                        lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                        If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
                            'CARREGA ENTIDADE
                            _RetornoBco.CodBanco = sBanco
                            _RetornoBco.NumAviso = sNumAviso
                            _RetornoBco.NumTit = sNumtit
                            _RetornoBco.SeqTit = sSeqTit
                            _RetornoBco.CodAgen = Trim(Mid(sBuffer, 27, 4))
                            _RetornoBco.NumCta = "22267-4" 'Trim(Str(Val(Mid(sBuffer, 31, 14)))) '& Trim(Mid(sBuffer, 43, 2))")
                            _RetornoBco.VlrPago = Val(Mid(sBuffer, 53, 15)) / 100
                            _RetornoBco.VlrJuros = 0
                            _RetornoBco.VlrDesc = 0
                            _RetornoBco.VlrIOF = 0
                            _RetornoBco.VlrAbat = 0
                            _RetornoBco.Processado = "N"
                            _RetornoBco.DtVcto = sVcto
                            _RetornoBco.DtPagto = Mid(sBuffer, 45, 4) & " - " & Mid(sBuffer, 49, 2) & " - " & Mid(sBuffer, 51, 2)
                            _RetornoBco.DtArq = sDataArq

                            'INSERE RETORNO BANCO
                            _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                            If Not _Retorno.Sucesso Then
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                fsFile.Close()

                                GoTo ErrBaixa
                            End If

                        End If
                    ElseIf isErroBancoDoBrasil(Trim(Mid(sBuffer, 68, 2))) Then

                        'Grava histórico de contato do cliente com o motivo da não baixa do título
                        Dim Inconsistencia As New CodInconsistencias
                        Dim clBuscaCodInconsistencia As New ConsultaCodInconsistencias
                        Dim clInsereErroDadosPgto As New InsereCodInconsistencia
                        Dim clInsereHistContato As New InserirHistoricoContato
                        Dim clBuscaContasAReceber As New ConsultaContasReceber
                        Dim ContasAReceber As New ContaReceber
                        Dim blnIsErro As Boolean
                        Dim Retorno As New Retorno

                        'Busca mensagem de inconsistencia
                        Inconsistencia = clBuscaCodInconsistencia.BuscaCodInconsistenciasDA(Trim(Mid(sBuffer, 68, 2)), "001", blnIsErro, connection, Transacao)

                        If Not Inconsistencia.Sucesso Then
                            Retorno = Funcoes.RetornoFunc(Inconsistencia.MsgErro)
                            lstRetorno.Add(Retorno)
                            'Funcoes.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        ContasAReceber = clBuscaContasAReceber.ConsultaContaReceberNumTitSeqTit(Strings.Left(Trim(Mid(sBuffer, 70, 25)), Len(Trim(Mid(sBuffer, 70, 25))) - 2), Strings.Right(Trim(Mid(sBuffer, 70, 25)), 2), connection, Transacao)

                        If Not ContasAReceber.Sucesso Then
                            Retorno = Funcoes.RetornoFunc(ContasAReceber.MsgErro)
                            lstRetorno.Add(Retorno)
                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        'Insere no histórico de contato do cliente motivo de não baixa do título
                        Retorno = clInsereHistContato.IncluiHistoricoContato(ContasAReceber.CodIntClie, "VERISURE", "Título " & Strings.Left(Trim(Mid(sBuffer, 70, 25)), Len(Trim(Mid(sBuffer, 70, 25))) - 2) & "-" & Strings.Right(Trim(Mid(sBuffer, 70, 25)), 2) & " não baixado automaticamente pelo seguinte motivo: " & Inconsistencia.Mensagem, connection, Transacao)

                        If Not Retorno.Sucesso Then
                            lstRetorno.Add(Retorno)
                            'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                            blnRetorno = True
                            Exit Sub
                        End If

                        'Caso o código de incosistencia for um erro para assinalar o cliente, realizar update no cliente - Lucas 15/05/2017
                        If blnIsErro Then

                            Retorno = clInsereErroDadosPgto.InsereErroDadosPgtoCliente(ContasAReceber.CodIntClie, connection, Transacao)

                            If Not Retorno.Sucesso Then
                                lstRetorno.Add(Retorno)
                                'FuncoesUI.TrataErro(True, lstRetorno, connection, Transacao)
                                blnRetorno = True
                                Exit Sub
                            End If

                        End If

                        'sBuffer = fsFile.ReadLine

                    End If
                End If
                sBuffer = fsFile.ReadLine
            Loop

        Else

            'VERIFICA SE EH DO BANCO DO BRASIL
            sBanco = Mid(sBuffer, 1, 3)
            'If sBanco <> codigoBanco Then
            '    blnRetorno = True
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            '    fsFile.Close()

            '    GoTo ErrBaixa
            'End If

            '**********BOLETO******************
            '-RETIRAR QDO EM PRODUCAO---------------------------------------------------------
            'MsgBox "BANCO DO BRASIL (boleto) ainda não liberado para faturamento. Consultar analista de sistemas !"
            'Exit Sub
            '---------------------------------------------------------------
            Do While (sBuffer) <> Nothing
                'CHECA SE EH HEADER
                If Mid(sBuffer, 8, 1) = "0" Then
                    'CHECA SE EH RETORNO
                    If Mid(sBuffer, 143, 1) <> "2" Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Descricao, ErrorConstants.ESTE_ARQUIVO_NAO_E_DE_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                    sBanco = Mid(sBuffer, 1, 3)

                    'VERIFICA SE EH DO BANCO DO BRASIL
                    'If sBanco <> codigoBanco Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Descricao & " - " & strNomeBanco, ErrorConstants.ESTE_ARQUIVO_NAO_E_DO_BANCO_SELECIONADO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    '    fsFile.Close()

                    '    GoTo ErrBaixa
                    'End If
                    sAgencia = CStr(CLng(Mid(sBuffer, 53, 6)))
                    sConta = CStr(CLng(Mid(sBuffer, 59, 13)))

                    'Verifica se o arquivo é corresponte ao Banco/Agencia/Conta selecionado
                    'If codigoBanco <> sBanco Or Trim(lblAgen.Text) <> Trim(sAgencia) Or Trim(Funcoes.FSepara(lblCCorrente.Text)) <> Trim(Funcoes.FSepara(sConta)) Then
                    '    blnRetorno = True
                    '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Descricao.Replace("<%BANCO%>", sBanco).Replace("<%AGENCIA%>", sAgencia).Replace("<%CONTA%>", sConta), ErrorConstants.ARQUIVO_DE_RETORNO_BANCOAGENCIACONTA_CORRENTE_NAO_CORRESPONDE.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                    '    fsFile.Close()

                    '    GoTo ErrBaixa
                    'End If


                    sDataArq = Format(Mid(sBuffer, 144, 2) & "/" & Mid(sBuffer, 146, 2) & "/" & Mid(sBuffer, 148, 4), "yyyy-MM-dd")
                    sNumAviso = Strings.Right(Mid(sBuffer, 158, 6), 5)


                    'VERIFICA SE O ARQUIVO JA FOI ATUALIZADO
                    lstControleRetBco = clConsultaControleRetBco.BuscaControleRetBcoPorCodBancoNumAvisoNumCta(sBanco, sNumAviso, sConta, connection, Transacao)
                    If Not lstControleRetBco(0).Sucesso And lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then ''TRATA ERRO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstControleRetBco(0).MsgErro, lstControleRetBco(0).NumErro, lstControleRetBco(0).Sucesso, lstControleRetBco(0).TipoErro, lstControleRetBco(0).ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    ElseIf lstControleRetBco(0).TipoErro = DadosGenericos.TipoErro.None Then 'ARQUIVO JA ATUALIZADO
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Descricao, ErrorConstants.ESTE_ARQUIVO_RETORNO_JA_FOI_ATUALIZADO_NO_SISTEMA.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If


                    'CARREGA A ENTIDADE
                    _ControleRetBco.CodBanco = sBanco
                    _ControleRetBco.NumAviso = sNumAviso
                    _ControleRetBco.NumCta = sConta
                    _ControleRetBco.DtArq = sDataArq

                    'INSERE O ARQUIVO DE RETORNO
                    _Retorno = clInserirControleRetBco.InserirControleRetBco(_ControleRetBco, connection, Transacao)
                    If Not _Retorno.Sucesso Then
                        blnRetorno = True
                        lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                        fsFile.Close()

                        GoTo ErrBaixa
                    End If

                End If

                'CHECA SE EH DETALHE
                '*********************************************
                'DETALHE SEGMENTO T
                '*********************************************
                If Trim(Mid(sBuffer, 8, 1)) = "3" And Trim(Mid(sBuffer, 14, 1)) = "T" Then
                    sNumtit = Trim(Mid(sBuffer, 59, 15))
                    sSeqTit = Strings.Right(sNumtit, 2)
                    sNumtit = Strings.Left(sNumtit, Len(sNumtit) - 2)
                    sVcto = Format(Format(Trim(Mid(sBuffer, 74, 2)) & "/" & Trim(Mid(sBuffer, 76, 2)) & "/" & Trim(Mid(sBuffer, 78, 4)), "dd/MM/yyyy"), "yyyy-MM-dd")

                    'CHECA SE EH LIQUIDACAO
                    If Mid(sBuffer, 16, 2) = "06" Then 'Or Mid(sBuffer, 16, 2) = "09" Or Mid(sBuffer, 16, 2) = "17" Then

                        lstRetornoBco = clConsultaRetornoBco.BuscaControleRetBcoPorCodBancoNumAviso(sBanco, sNumAviso, sNumtit, sSeqTit, Convert.ToDateTime(sVcto), connection, Transacao)
                        If Not lstRetornoBco(0).Sucesso And lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
                            blnRetorno = True
                            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstRetornoBco(0).MsgErro, lstRetornoBco(0).NumErro, lstRetornoBco(0).Sucesso, lstRetornoBco(0).TipoErro, lstRetornoBco(0).ImagemErro)
                            fsFile.Close()

                            GoTo ErrBaixa
                        ElseIf lstRetornoBco(0).TipoErro = DadosGenericos.TipoErro.Funcional Then
                            'CARREGA ENTIDADE
                            _RetornoBco.CodBanco = sBanco
                            _RetornoBco.NumAviso = sNumAviso
                            _RetornoBco.NumTit = sNumtit
                            _RetornoBco.SeqTit = sSeqTit
                            _RetornoBco.CodAgen = strAgen
                            _RetornoBco.NumCta = strContaCorrente '& Trim(Mid(sBuffer, 43, 2))")
                            If Trim(Mid(sBuffer, 14, 1)) = "U" Then
                                _RetornoBco.VlrPago = Val(Mid(sBuffer, 78, 15)) / 100
                                _RetornoBco.VlrJuros = (Val(Mid(sBuffer, 18, 15)) / 100) + (Val(Mid(sBuffer, 123, 15)) / 100)
                                _RetornoBco.VlrDesc = Val(Mid(sBuffer, 33, 15)) / 100
                                _RetornoBco.VlrIOF = 0
                                _RetornoBco.VlrAbat = Val(Mid(sBuffer, 48, 15)) / 100
                                _RetornoBco.Processado = "N"
                                _RetornoBco.DtVcto = sVcto
                                _RetornoBco.DtPagto = Format(Trim(Mid(sBuffer, 138, 2)) & "/" & Trim(Mid(sBuffer, 140, 2)) & "/" & Trim(Mid(sBuffer, 142, 4)), "dd/MM/yyyy")
                                _RetornoBco.DtArq = sDataArq

                                'INSERE RETORNO BANCO
                                _Retorno = clInserirRetornoBco.InserirRetornoBco(_RetornoBco, connection, Transacao)
                                If Not _Retorno.Sucesso Then
                                    blnRetorno = True
                                    lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)
                                    fsFile.Close()

                                    GoTo ErrBaixa
                                End If
                            Else
                                blnRetorno = True
                                lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Descricao, ErrorConstants.PROBLEMA_NA_SEQUENCIA_DADOS_ARQUIVO_RETORNO.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
                                fsFile.Close()

                                GoTo ErrBaixa
                            End If

                        End If

                    End If
                End If
                sBuffer = fsFile.ReadLine
            Loop
        End If
        '**************************************************************************************************************
        fsFile.Close()

        _Retorno = clAlterarBaixaContasReceber.BaixaCRec(sNumAviso, sBanco, "VERISURE", sDataArq, "22267-4", "29629", connection, Transacao)
        If Not _Retorno.Sucesso Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, _Retorno.MsgErro, _Retorno.NumErro, _Retorno.Sucesso, _Retorno.TipoErro, _Retorno.ImagemErro)

            GoTo ErrBaixa
        End If


        lstBxAutoErros = clConsultaBxAutoErros.BuscaBxAutoErros(connection, Transacao)
        If Not lstBxAutoErros(0).Sucesso And lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.Arquitetura Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, lstBxAutoErros(0).MsgErro, lstBxAutoErros(0).NumErro, lstBxAutoErros(0).Sucesso, lstBxAutoErros(0).TipoErro, lstBxAutoErros(0).ImagemErro)

            GoTo ErrBaixa
        ElseIf lstBxAutoErros(0).TipoErro = DadosGenericos.TipoErro.None Then
            blnRetorno = True
            lstRetorno = Funcoes.CriaRetorno(lstRetorno, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Descricao, ErrorConstants.O_SISTEMA_DETECTOU_SEGUINTES_TITULOS_RETORNADO_PELO_BANCO_BAIXA_MANUAL.Id, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)

            'For i As Integer = 0 To lstBxAutoErros.Count - 1
            '    sErros = sErros & vbCr & "Título " & lstBxAutoErros(i).NumTit & ", Sequência " & lstBxAutoErros(i).SeqTit & ", emitido em " & lstBxAutoErros(i).DtEmissao
            '    lstRetorno = Funcoes.CriaRetorno(lstRetorno, "", sErros, False, DadosGenericos.TipoErro.Funcional, DadosGenericos.ImagemRetorno.Alerta)
            'Next
            'grdTitJaBaixados.DataSource = lstBxAutoErros
            'UpdateTotal(lblTotalTitJaBaixados, grdTitJaBaixados.ChildRows.Count)

            'If (grdTitJaBaixados.Rows.Count > 0) Then
            '    gbxTitJaBaixados.Visible = True
            'Else
            '    gbxTitJaBaixados.Visible = False
            'End If
        End If

        'EXIBE MESSAGEM
        'If blnRetorno Then
        '    ExibirErro()
        'End If

        Transacao.Commit()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

        'MsgBox("Baixa dos títulos realizada com sucesso !", vbInformation, Me.Text)
        'RadMessageBox.Show("Baixa dos títulos realizada com sucesso!", Me.Text, MessageBoxButtons.OK, RadMessageIcon.Info)
        Exit Sub

ErrBaixa:

        Transacao.Rollback()
        connection.Close()

        'MUDA O CURSOR DO MOUSE PARA AMPULHETA
        'Me.Cursor = Cursors.Default
        '------------------------------------

    End Sub

    Private Function isErroBancoDoBrasil(ByVal CodInconsistencia As String) As Boolean

        Select Case CodInconsistencia
            Case "01"
                Return True
            Case "02"
                Return True
            Case "04"
                Return True
            Case "10"
                Return True
            Case "12"
                Return True
            Case "13"
                Return True
            Case "14"
                Return True
            Case "15"
                Return True
            Case "18"
                Return True
            Case "30"
                Return True
            Case "96"
                Return True
            Case "97"
                Return True
            Case "98"
                Return True
            Case "99"
                Return True
            Case Else
                Return False
        End Select


    End Function


End Module
