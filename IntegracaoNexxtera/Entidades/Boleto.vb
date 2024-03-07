Imports BoletoNet
Public Class Boleto
    Public Enum Bancos
        BancodoBrasil = 1
        Banrisul = 41
        Basa = 3
        Bradesco = 237
        BRB = 70
        Caixa = 104
        HSBC = 399
        Itau = 341
        Real = 356
        Safra = 422
        Santander = 33
        Sicoob = 756
        Sicred = 748
        Sudameris = 347
        Unibanco = 409
        Semear = 743
    End Enum

    Public Sub New(ByVal Banco As Integer)
        boletoBancario = New BoletoBancario()
        boletoBancario.CodigoBanco = CShort(Boleto.Bancos.Itau)
    End Sub

    Public Property boletoBancario As BoletoBancario

    Public Function BancodoBrasil(cr_tb As ContaReceber) As BoletoBancario
        Dim NossoNumero As String = ""
        Dim NumCta As String = ""
        Dim DigConta As String = ""
        Dim Conta As Integer
        Dim CodAgen As String = ""

        NumCta = cr_tb.NumCta.PadLeft(7, "0")
        CodAgen = cr_tb.CodAgen.PadLeft(4, "0")
        NossoNumero = Trim(Right(cr_tb.NossoNumeroBco, 13))

        NumCta = Replace(NumCta, "-", "")
        NumCta = Replace(NumCta, "_", "")
        Conta = Left(NumCta, 6).PadLeft(8, "0")
        DigConta = NumCta.Substring(NumCta.Length - 1, 1)

        Dim c As Cedente = New Cedente("11.660.106/0001-38", "VERISURE BRASIL MONITORAMENTO DE ALARMES SA", CodAgen, Conta, DigConta)
        c.Codigo = Left(NossoNumero, 7)
        Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(cr_tb.DtVcto, cr_tb.VlrInd, cr_tb.Carteira, NossoNumero, c, New EspecieDocumento(1, "1"))

        'b.NumeroDocumento = TitlesDatas.ContasAReceber.NumTit
        'b.Sacado = New Sacado(TitlesDatas.Cliente.CPF_CNPJ, TitlesDatas.Cliente.RazaoSocial)
        'b.Sacado.Endereco.[End] = TitlesDatas.Cliente.endereco.Endereco & ", " & TitlesDatas.Cliente.endereco.NumEndereco
        'b.Sacado.Endereco.Bairro = TitlesDatas.Cliente.endereco.Bairro
        'b.Sacado.Endereco.Cidade = TitlesDatas.Cliente.endereco.Cidade
        'b.Sacado.Endereco.CEP = TitlesDatas.Cliente.endereco.CEP
        'b.Sacado.Endereco.UF = TitlesDatas.Cliente.endereco.UF
        'b.Sacado.InformacoesSacado.Add(New InfoSacado(String.Format("TÍTULO: {0}{1}", TitlesDatas.ContasAReceber.NumTit, TitlesDatas.ContasAReceber.SeqTit)))
        b.DataDocumento = cr_tb.DtEmissao
        b.DataVencimento = cr_tb.DtVcto
        b.DataProcessamento = cr_tb.DtEmissao
        b.PercMulta = cr_tb.TaxaDia
        b.PercJurosMora = cr_tb.TaxaMes

        b.LocalPagamento = String.Format("Até o vencimento, preferencialmente no {0}. <br>Após o vencimento, somente no {0}.", b.EspecieDocumento.Banco.Nome)

        'b.Instrucoes.Add(New Instrucao_BancoBrasil() With {
        '    .Descricao = String.Format("<br><br>PARA SEU CONFORTO, SOLICITE O PAGAMENTO ATRAVÉS DE DÉBITO AUTOMÁTICO.<br>LIGUE PARA O SAT - (11) 3811-1000 DEMAIS LOCALIDADES 4002-7222", TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes)
        '})

        boletoBancario.Boleto = b
        boletoBancario.Boleto.Valida()

        Return boletoBancario
    End Function

    'Public Function Banrisul(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim NossoNumero As String = ""
    '    Dim NumCta As String = ""
    '    Dim Carteira As String = TitlesDatas.ContasAReceber.Carteira
    '    Dim DigConta As String = ""
    '    Dim Conta As Integer
    '    Dim CodAgen As String = ""

    '    NumCta = TitlesDatas.ContasAReceber.NumCta.PadLeft(11, "0")
    '    CodAgen = TitlesDatas.ContasAReceber.CodAgen.PadLeft(4, "0")
    '    NossoNumero = Left(TitlesDatas.ContasAReceber.NossoNumero, 12)

    '    NumCta = Replace(NumCta, "-", "")
    '    NumCta = Replace(NumCta, "_", "")
    '    Conta = NumCta
    '    DigConta = NumCta.Substring(NumCta.Length - 1, 1)

    '    Dim c As Cedente = New Cedente("11.660.106/0001-38", "VERISURE BRASIL MONITORAMENTO DE ALARMES SA", CodAgen, Conta)
    '    'c.Codigo = "00000000504"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(TitlesDatas.ContasAReceber.DtVcto, TitlesDatas.ContasAReceber.vlrInd, Carteira, NossoNumero, c, New EspecieDocumento(33, "1"))

    '    b.NumeroDocumento = TitlesDatas.ContasAReceber.NumTit
    '    b.Sacado = New Sacado(TitlesDatas.Cliente.CPF_CNPJ, TitlesDatas.Cliente.RazaoSocial)
    '    b.Sacado.Endereco.[End] = TitlesDatas.Cliente.endereco.Endereco & ", " & TitlesDatas.Cliente.endereco.NumEndereco
    '    b.Sacado.Endereco.Bairro = TitlesDatas.Cliente.endereco.Bairro
    '    b.Sacado.Endereco.Cidade = TitlesDatas.Cliente.endereco.Cidade
    '    b.Sacado.Endereco.CEP = TitlesDatas.Cliente.endereco.CEP
    '    b.Sacado.Endereco.UF = TitlesDatas.Cliente.endereco.UF
    '    'b.Sacado.InformacoesSacado.Add(New InfoSacado(String.Format("TÍTULO: {0}{1}", TitlesDatas.ContasAReceber.NumTit, TitlesDatas.ContasAReceber.SeqTit)))
    '    b.DataDocumento = TitlesDatas.ContasAReceber.DtEmissao
    '    b.DataVencimento = TitlesDatas.ContasAReceber.DtVcto
    '    b.DataProcessamento = TitlesDatas.ContasAReceber.DtEmissao
    '    b.PercMulta = TitlesDatas.ContasAReceber.TaxaDia
    '    b.PercJurosMora = TitlesDatas.ContasAReceber.TaxaMes

    '    b.LocalPagamento = String.Format("Até o vencimento, preferencialmente no {0}. <br>Após o vencimento, somente no {0}.", b.EspecieDocumento.Banco.Nome)

    '    b.Instrucoes.Add(New Instrucao_Banrisul() With {
    '        .Descricao = String.Format("<br><br>PARA SEU CONFORTO, SOLICITE O PAGAMENTO ATRAVÉS DE DÉBITO AUTOMÁTICO.<br>LIGUE PARA O SAT - (11) 3811-1000 DEMAIS LOCALIDADES 4002-7222", TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes)
    '    })

    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    boletoBancario.Boleto.Carteira = TitlesDatas.ContasAReceber.Carteira

    '    Return boletoBancario
    'End Function

    'Public Function Basa(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim vencimento As DateTime = DateTime.Now.AddDays(10)
    '    Dim c As Cedente = New Cedente("00.000.000/0000-00", "Empresa de Atacado", "1234", "5", "12345678", "9")
    '    c.Codigo = "12548"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 45.5D, "CNR", "125478", c)
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    b.NumeroDocumento = "12345678901"
    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    boletoBancario.Cedente.Endereco = New Endereco() With {
    '        .[End] = "Endereço do Cedente",
    '        .Bairro = "Bairro",
    '        .Cidade = "Cidade",
    '        .Cep = "70000000",
    '        .UF = "DF"
    '    }
    '    boletoBancario.RemoveSimboloMoedaValorDocumento = False

    '    Return boletoBancario
    'End Function

    'Public Function Bradesco(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim NossoNumero As String = ""
    '    Dim NumCta As String = ""
    '    Dim Carteira As String = Right(TitlesDatas.ContasAReceber.Carteira, 2)
    '    Dim DigConta As String = ""
    '    Dim Conta As Integer
    '    Dim CodAgen As String = ""

    '    NumCta = TitlesDatas.ContasAReceber.NumCta.PadLeft(9, "0")
    '    CodAgen = Left(TitlesDatas.ContasAReceber.CodAgen, 4).PadLeft(4, "0")
    '    NossoNumero = Left(TitlesDatas.ContasAReceber.NossoNumero, 11)

    '    NumCta = Replace(NumCta, "-", "")
    '    NumCta = Replace(NumCta, "_", "")
    '    Conta = NumCta
    '    DigConta = NumCta.Substring(NumCta.Length - 1, 1)

    '    Dim c As Cedente = New Cedente("11.660.106/0001-38", "VERISURE BRASIL MONITORAMENTO DE ALARMES SA", CodAgen, Conta)
    '    'c.Codigo = "00000000504"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(TitlesDatas.ContasAReceber.DtVcto, TitlesDatas.ContasAReceber.vlrInd, Carteira, NossoNumero, c, New EspecieDocumento(33, "1"))

    '    b.NumeroDocumento = TitlesDatas.ContasAReceber.NumTit
    '    b.Sacado = New Sacado(TitlesDatas.Cliente.CPF_CNPJ, TitlesDatas.Cliente.RazaoSocial)
    '    b.Sacado.Endereco.[End] = TitlesDatas.Cliente.endereco.Endereco & ", " & TitlesDatas.Cliente.endereco.NumEndereco
    '    b.Sacado.Endereco.Bairro = TitlesDatas.Cliente.endereco.Bairro
    '    b.Sacado.Endereco.Cidade = TitlesDatas.Cliente.endereco.Cidade
    '    b.Sacado.Endereco.CEP = TitlesDatas.Cliente.endereco.CEP
    '    b.Sacado.Endereco.UF = TitlesDatas.Cliente.endereco.UF
    '    'b.Sacado.InformacoesSacado.Add(New InfoSacado(String.Format("TÍTULO: {0}{1}", TitlesDatas.ContasAReceber.NumTit, TitlesDatas.ContasAReceber.SeqTit)))
    '    b.DataDocumento = TitlesDatas.ContasAReceber.DtEmissao
    '    b.DataVencimento = TitlesDatas.ContasAReceber.DtVcto
    '    b.DataProcessamento = TitlesDatas.ContasAReceber.DtEmissao
    '    b.PercMulta = TitlesDatas.ContasAReceber.TaxaDia
    '    b.PercJurosMora = TitlesDatas.ContasAReceber.TaxaMes

    '    b.LocalPagamento = String.Format("Até o vencimento, preferencialmente no {0}. <br>Após o vencimento, somente no {0}.", b.EspecieDocumento.Banco.Nome)

    '    b.Instrucoes.Add(New Instrucao_Bradesco() With {
    '        .Descricao = String.Format("<br><br>PARA SEU CONFORTO, SOLICITE O PAGAMENTO ATRAVÉS DE DÉBITO AUTOMÁTICO.<br>LIGUE PARA O SAT - (11) 3811-1000 DEMAIS LOCALIDADES 4002-7222", TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes)
    '    })

    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    boletoBancario.Boleto.Carteira = TitlesDatas.ContasAReceber.Carteira

    '    Return boletoBancario
    'End Function

    'Public Function BRB(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim vencimento As DateTime = DateTime.Now.AddDays(10)
    '    Dim c As Cedente = New Cedente("00.000.000/0000-00", "Empresa de Atacado", "208", "", "010357", "6")
    '    c.Codigo = "13000"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 0.01D, "COB", "119964", c)
    '    b.NumeroDocumento = "119964"
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    Return boletoBancario
    'End Function

    'Public Function Caixa(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim NossoNumero As String = ""
    '    Dim NumCta As String = ""
    '    Dim Carteira As String = TitlesDatas.ContasAReceber.Carteira
    '    Dim DigConta As String = ""
    '    Dim Conta As Integer
    '    Dim CodAgen As String = ""

    '    NumCta = TitlesDatas.ContasAReceber.NumCta.PadLeft(9, "0")
    '    CodAgen = Left(TitlesDatas.ContasAReceber.CodAgen, 4).PadLeft(4, "0")
    '    NossoNumero = TitlesDatas.ContasAReceber.NossoNumero

    '    NumCta = Replace(NumCta, "-", "")
    '    NumCta = Replace(NumCta, "_", "")
    '    Conta = NumCta
    '    DigConta = NumCta.Substring(NumCta.Length - 1, 1)

    '    Dim c As Cedente = New Cedente("11.660.106/0001-38", "VERISURE BRASIL MONITORAMENTO DE ALARMES SA", CodAgen, Conta)
    '    c.Codigo = "112233"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(TitlesDatas.ContasAReceber.DtVcto, TitlesDatas.ContasAReceber.vlrInd, Carteira, NossoNumero, c, New EspecieDocumento(33, "1"))

    '    b.NumeroDocumento = TitlesDatas.ContasAReceber.NumTit
    '    b.Sacado = New Sacado(TitlesDatas.Cliente.CPF_CNPJ, TitlesDatas.Cliente.RazaoSocial)
    '    b.Sacado.Endereco.[End] = TitlesDatas.Cliente.endereco.Endereco & ", " & TitlesDatas.Cliente.endereco.NumEndereco
    '    b.Sacado.Endereco.Bairro = TitlesDatas.Cliente.endereco.Bairro
    '    b.Sacado.Endereco.Cidade = TitlesDatas.Cliente.endereco.Cidade
    '    b.Sacado.Endereco.CEP = TitlesDatas.Cliente.endereco.CEP
    '    b.Sacado.Endereco.UF = TitlesDatas.Cliente.endereco.UF
    '    'b.Sacado.InformacoesSacado.Add(New InfoSacado(String.Format("TÍTULO: {0}{1}", TitlesDatas.ContasAReceber.NumTit, TitlesDatas.ContasAReceber.SeqTit)))
    '    b.DataDocumento = TitlesDatas.ContasAReceber.DtEmissao
    '    b.DataVencimento = TitlesDatas.ContasAReceber.DtVcto
    '    b.DataProcessamento = TitlesDatas.ContasAReceber.DtEmissao
    '    b.PercMulta = TitlesDatas.ContasAReceber.TaxaDia
    '    b.PercJurosMora = TitlesDatas.ContasAReceber.TaxaMes

    '    b.LocalPagamento = String.Format("Até o vencimento, preferencialmente no {0}. <br>Após o vencimento, somente no {0}.", b.EspecieDocumento.Banco.Nome)

    '    b.Instrucoes.Add(New Instrucao_Caixa() With {
    '        .Descricao = String.Format("<br><br>PARA SEU CONFORTO, SOLICITE O PAGAMENTO ATRAVÉS DE DÉBITO AUTOMÁTICO.<br>LIGUE PARA O SAT - (11) 3811-1000 DEMAIS LOCALIDADES 4002-7222", TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes)
    '    })

    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    boletoBancario.Boleto.Carteira = TitlesDatas.ContasAReceber.Carteira

    '    Return boletoBancario
    'End Function

    'Private Function SearchClientsDatas(codIntClie As String) As Cliente
    '    Dim c_tb As New Cliente

    '    Try
    '        Dim consulta As ICliente
    '        Dim _Ret As New Retorno With {.sucesso = True}
    '        consulta = New ClienteDAO
    '        'Preenche o datatable
    '        Dim dt As New DataTable("dt")

    '        _Ret = consulta.SearchClientsDatas(codIntClie, dt)

    '        'Throw caso haja erro
    '        If Not _Ret.sucesso Then Throw New Exception(_Ret.msgErro)

    '        If (dt.Rows.Count > 0) Then
    '            c_tb.CodIntClie = dt.Rows(0).Item("CodIntClie")
    '            c_tb.RazaoSocial = dt.Rows(0).Item("RazaoSocial")
    '            c_tb.Endereco.Endereco = dt.Rows(0).Item("Endereco")
    '            c_tb.Endereco.NumEndereco = dt.Rows(0).Item("NumEndereco")
    '            c_tb.Endereco.Complemento = dt.Rows(0).Item("Complemento")
    '            c_tb.Endereco.UF = dt.Rows(0).Item("UF")
    '            c_tb.Endereco.Cep = dt.Rows(0).Item("CEP")
    '            c_tb.Endereco.Bairro = dt.Rows(0).Item("Bairro")
    '            c_tb.Endereco.Cidade = dt.Rows(0).Item("Cidade")
    '            'c_tb.CodAgen = dt.Rows(0).Item("CodAgen")
    '            'c_tb.NumCta = dt.Rows(0).Item("NumCta")
    '            'c_tb.CPF_CNPJ = dt.Rows(0).Item("CPF_CNPJ")
    '            'c_tb.Carteira = dt.Rows(0).Item("CodCarteira")
    '        End If

    '    Catch ex As Exception
    '        Common.EnviarEmail("BR.DG.Sistemas@verisure.com.br", "Erro faturas verisure - Contrato " & codIntClie & " - " & System.Reflection.MethodBase.GetCurrentMethod().Name, ex.Message)
    '    End Try

    '    Return c_tb
    'End Function

    'Public Function HSBC(TitlesDatas As TitlesDatas) As BoletoBancario
    '    'Carteira = "1"
    '    'conta = Left(conta, 5) & "-" & Right(conta, 2)
    '    'numerodoc = numerodoc & Session("SeqTit")

    '    ''DESCARTO DIGITO VERIFICAR DO NOSSO NUMERO
    '    'Dim nossonumero As String = Left(Trim(Session("NossoNumero")), 10)

    '    'Cedente = Left(nossonumero, 5)
    '    'Session("NossoNumero") = Right(nossonumero, 5)

    '    Dim vencimento As DateTime = DateTime.Now.AddDays(10)
    '    Dim c As Cedente = New Cedente("00.000.000/0000-00", "Minha empresa", "0000", "", "00000", "00")
    '    c.Codigo = "0000000"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 2, "CNR", "1330001490684", c)
    '    b.NumeroDocumento = "9999999999999"
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    Return boletoBancario
    'End Function

    Public Function CodigoBarraItau(cr_tb As ContaReceber) As BoletoBancario
        Dim NossoNumero As String = ""
        Dim NumCta As String = ""
        Dim DigConta As String = ""
        Dim Conta As Integer
        Dim CodAgen As String = ""

        NumCta = cr_tb.NumCta.PadLeft(7, "0")
        CodAgen = cr_tb.CodAgen.PadLeft(4, "0")
        NossoNumero = Left(cr_tb.NossoNumeroBco, 8)

        NumCta = Replace(NumCta, "-", "")
        NumCta = Replace(NumCta, "_", "")
        Conta = Left(NumCta, NumCta.Length - 1)
        DigConta = NumCta.Substring(NumCta.Length - 1, 1)

        Dim c As Cedente = New Cedente("11.660.106/0001-38", "VERISURE BRASIL MONITORAMENTO DE ALARMES SA", CodAgen, Conta, DigConta)

        Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(cr_tb.DtVcto, cr_tb.VlrInd, cr_tb.Carteira, NossoNumero, c, New EspecieDocumento(341, "5"))

        'b.NumeroDocumento = TitlesDatas.ContasAReceber.NumTit
        'b.Sacado = New Sacado(TitlesDatas.Cliente.CPF_CNPJ, TitlesDatas.Cliente.RazaoSocial)
        'b.Sacado.Endereco.[End] = TitlesDatas.Cliente.endereco.Endereco & ", " & TitlesDatas.Cliente.endereco.NumEndereco
        'b.Sacado.Endereco.Bairro = TitlesDatas.Cliente.endereco.Bairro
        'b.Sacado.Endereco.Cidade = TitlesDatas.Cliente.endereco.Cidade
        'b.Sacado.Endereco.CEP = TitlesDatas.Cliente.endereco.CEP
        'b.Sacado.Endereco.UF = TitlesDatas.Cliente.endereco.UF
        'b.Sacado.InformacoesSacado.Add(New InfoSacado(String.Format("TÍTULO: {0}{1}", TitlesDatas.ContasAReceber.NumTit, TitlesDatas.ContasAReceber.SeqTit)))
        'b.DataDocumento = cr_tb.DtEmissao
        b.DataVencimento = cr_tb.DtVcto
        'b.DataProcessamento = cr_tb.DtEmissao
        b.PercMulta = cr_tb.TaxaDia
        b.PercJurosMora = cr_tb.TaxaMes

        'b.LocalPagamento = String.Format("Até o vencimento, preferencialmente no {0} e Após o vencimento, somente no {0}.", b.EspecieDocumento.Banco.Nome)

        'b.Instrucoes.Add(New Instrucao_Itau() With {
        '    .Descricao = String.Format("<br> NF-E: {0} / COD.: {1} JUROS AO DIA: {2}% MULTA MENSAL: {3}%.", TitlesDatas.ContasAReceber.NumNFe, TitlesDatas.ContasAReceber.CodVerNfe, TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes * 100)
        '})

        'b.Instrucoes.Add(New Instrucao_Itau() With {
        '    .Descricao = String.Format("<br><br>PARA SEU CONFORTO, SOLICITE O PAGAMENTO ATRAVÉS DE DÉBITO AUTOMÁTICO.<br>LIGUE PARA O SAT - (11) 3811-1000 DEMAIS LOCALIDADES 4002-7222", TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes)
        '})

        boletoBancario.Boleto = b
        boletoBancario.Boleto.Valida()
        boletoBancario.Boleto.Carteira = cr_tb.Carteira

        Return boletoBancario
    End Function

    'Public Function Real(TitlesDatas As TitlesDatas) As BoletoBancario
    '    'agencia = agencia.PadLeft(4, "0")
    '    'conta = conta.PadLeft(7, "0")
    '    'Cedente = conta.PadLeft(7, "0")

    '    Dim vencimento As DateTime = DateTime.Now.AddDays(10)
    '    Dim c As Cedente = New Cedente("00.000.000/0000-00", "Coloque a Razão Social da sua empresa aqui", "1234", "12345")
    '    c.Codigo = "12345"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 0.1D, "57", "123456", c, New EspecieDocumento(356, "9"))
    '    b.NumeroDocumento = "1234567"
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    boletoBancario.Boleto = b
    '    Dim ed As EspeciesDocumento = EspecieDocumento_Real.CarregaTodas()
    '    boletoBancario.Boleto.Valida()
    '    Return boletoBancario
    'End Function

    'Public Function Safra(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim vencimento As DateTime = DateTime.Now.AddDays(10)
    '    Dim c As Cedente = New Cedente("00.000.000/0000-00", "Empresa de Atacado", "0542", "5413000")
    '    c.Codigo = "13000"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 1642, "198", "02592082835", c)
    '    b.NumeroDocumento = "1008073"
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    Dim instrucao As Instrucao_Safra = New Instrucao_Safra()
    '    instrucao.Descricao = "Instrução 1"
    '    b.Instrucoes.Add(instrucao)
    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    Return boletoBancario
    'End Function

    'Public Function Santander(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim NossoNumero As String = ""
    '    Dim NumCta As String = ""
    '    Dim Carteira As String = IIf(TitlesDatas.ContasAReceber.Carteira = "ECR", "104", TitlesDatas.ContasAReceber.Carteira)
    '    Dim DigConta As String = ""
    '    Dim Conta As Integer
    '    Dim CodAgen As String = ""

    '    NumCta = TitlesDatas.ContasAReceber.NumCta.PadLeft(11, "0")
    '    CodAgen = TitlesDatas.ContasAReceber.CodAgen.PadLeft(4, "0")
    '    NossoNumero = Left(TitlesDatas.ContasAReceber.NossoNumero, 12)

    '    NumCta = Replace(NumCta, "-", "")
    '    NumCta = Replace(NumCta, "_", "")
    '    Conta = NumCta
    '    DigConta = NumCta.Substring(NumCta.Length - 1, 1)

    '    Dim c As Cedente = New Cedente("11.660.106/0001-38", "VERISURE BRASIL MONITORAMENTO DE ALARMES SA", CodAgen, Conta)
    '    c.Codigo = "5043433"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(TitlesDatas.ContasAReceber.DtVcto, TitlesDatas.ContasAReceber.vlrInd, Carteira, NossoNumero, c, New EspecieDocumento(33, "1"))

    '    b.NumeroDocumento = TitlesDatas.ContasAReceber.NumTit
    '    b.Sacado = New Sacado(TitlesDatas.Cliente.CPF_CNPJ, TitlesDatas.Cliente.RazaoSocial)
    '    b.Sacado.Endereco.[End] = TitlesDatas.Cliente.endereco.Endereco & ", " & TitlesDatas.Cliente.endereco.NumEndereco
    '    b.Sacado.Endereco.Bairro = TitlesDatas.Cliente.endereco.Bairro
    '    b.Sacado.Endereco.Cidade = TitlesDatas.Cliente.endereco.Cidade
    '    b.Sacado.Endereco.CEP = TitlesDatas.Cliente.endereco.CEP
    '    b.Sacado.Endereco.UF = TitlesDatas.Cliente.endereco.UF
    '    'b.Sacado.InformacoesSacado.Add(New InfoSacado(String.Format("TÍTULO: {0}{1}", TitlesDatas.ContasAReceber.NumTit, TitlesDatas.ContasAReceber.SeqTit)))
    '    b.DataDocumento = TitlesDatas.ContasAReceber.DtEmissao
    '    b.DataVencimento = TitlesDatas.ContasAReceber.DtVcto
    '    b.DataProcessamento = TitlesDatas.ContasAReceber.DtEmissao
    '    b.PercMulta = TitlesDatas.ContasAReceber.TaxaDia
    '    b.PercJurosMora = TitlesDatas.ContasAReceber.TaxaMes

    '    'b.LocalPagamento = String.Format("Até o vencimento, preferencialmente no {0} e Após o vencimento, somente no {0}.", b.EspecieDocumento.Banco.Nome)
    '    b.LocalPagamento = String.Format("Pagável em qualquer banco até o vencimento")

    '    b.Instrucoes.Add(New Instrucao_Santander() With {
    '        .Descricao = String.Format("<br> JUROS AO DIA: {2}% MULTA MENSAL: {3}%.", TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes * 100)
    '    })

    '    'b.Instrucoes.Add(New Instrucao_Santander() With {
    '    '    .Descricao = String.Format("<br> NF-E: {0} / COD.: {1} JUROS AO DIA: {2}% MULTA MENSAL: {3}%.", TitlesDatas.ContasAReceber.NumNFe, TitlesDatas.ContasAReceber.CodVerNfe, TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes * 100)
    '    '})

    '    'b.Instrucoes.Add(New Instrucao_Santander() With {
    '    '    .Descricao = String.Format("<br><br>PARA SEU CONFORTO, SOLICITE O PAGAMENTO ATRAVÉS DE DÉBITO AUTOMÁTICO.<br>LIGUE PARA O SAT - (11) 3811-1000 DEMAIS LOCALIDADES 4002-7222", TitlesDatas.ContasAReceber.TaxaDia, TitlesDatas.ContasAReceber.TaxaMes)
    '    '})

    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    boletoBancario.Boleto.Carteira = TitlesDatas.ContasAReceber.Carteira
    '    Return boletoBancario
    'End Function

    'Public Function Sicoob(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim vencimento As DateTime = DateTime.Now.AddDays(10)
    '    Dim c As Cedente = New Cedente("00.000.000/0000-00", "Empresa de Atacado", "4444", "", "", "")
    '    c.Codigo = "123456"
    '    c.DigitoCedente = 7
    '    c.Carteira = "1"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 10, "1", "897654321", c)
    '    b.NumeroDocumento = "119964"
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    Return boletoBancario
    'End Function

    'Public Function Sicred(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim vencimento As DateTime = DateTime.Now.AddDays(1)
    '    Dim item1 As Instrucao_Sicredi = New Instrucao_Sicredi(9, 5)
    '    Dim item2 As Instrucao_Sicredi = New Instrucao_Sicredi()
    '    Dim c As Cedente = New Cedente("10.823.650/0001-90", "SAFIRALIFE", "0811", "81111")
    '    c.Codigo = "08111081111"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 0.1D, "1", "00000001", c)
    '    b.NumeroDocumento = "00000001"
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    b.Sacado.InformacoesSacado.Add(New InfoSacado("TÍTULO: " & "2541245"))
    '    item2.Descricao += " " & item1.QuantidadeDias.ToString() & " dias corridos do vencimento."
    '    b.Instrucoes.Add(item1)
    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    Return boletoBancario
    'End Function

    'Public Function Sudameris(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim vencimento As DateTime = DateTime.Now.AddDays(10)
    '    Dim c As Cedente = New Cedente("00.000.000/0000-00", "Empresa de Atacado", "0501", "6703255")
    '    c.Codigo = "13000"
    '    Dim nn As String = "0003020"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 1642, "198", nn, c)
    '    b.NumeroDocumento = "1008073"
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    Return boletoBancario
    'End Function

    'Public Function Unibanco(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim vencimento As DateTime = DateTime.Now.AddDays(10)
    '    Dim c As Cedente = New Cedente("00.000.000/0000-00", "Next Consultoria Ltda.", "0123", "100618", "9")
    '    c.Codigo = "203167"
    '    Dim b As BoletoNet.Boleto = New BoletoNet.Boleto(vencimento, 2952.95D, "20", "1803029901", c)
    '    b.NumeroDocumento = b.NossoNumero
    '    b.Sacado = New Sacado("000.000.000-00", "Nome do seu Cliente ")
    '    b.Sacado.Endereco.[End] = "Endereço do seu Cliente "
    '    b.Sacado.Endereco.Bairro = "Bairro"
    '    b.Sacado.Endereco.Cidade = "Cidade"
    '    b.Sacado.Endereco.CEP = "00000000"
    '    b.Sacado.Endereco.UF = "UF"
    '    boletoBancario.Boleto = b
    '    boletoBancario.Boleto.Valida()
    '    Return boletoBancario
    'End Function

    'Public Function Semear(TitlesDatas As TitlesDatas) As BoletoBancario
    '    Dim boleto = New BoletoNet.Boleto()
    '    boleto.Cedente = New Cedente() With {
    '        .Codigo = "743",
    '        .MostrarCNPJnoBoleto = True,
    '        .Nome = "BANCO SEMEAR",
    '        .CPFCNPJ = "65825481000110",
    '        .Carteira = "2",
    '        .DigCedente = "9",
    '        .ContaBancaria = New ContaBancaria() With {
    '            .Agencia = "001",
    '            .DigitoAgencia = "0",
    '            .Conta = "65456",
    '            .DigitoConta = "5"
    '        },
    '        .Endereco = New Endereco()
    '    }
    '    boleto.LocalPagamento = "Este boleto poderá ser pago em toda a Rede Bancária até 5 dias após o vencimento. "
    '    boleto.Instrucoes.Add(New Instrucao_Bradesco() With {
    '        .Descricao = "Não receber após o vencimento",
    '        .QuantidadeDias = 3
    '    })
    '    boleto.ValorBoleto = 251.51D
    '    boleto.ValorCobrado = 251.51D
    '    boleto.NossoNumero = "35148373401"
    '    boleto.NumeroDocumento = "051483734"
    '    boleto.DataVencimento = New DateTime(2017, 12, 4)
    '    boleto.DataProcessamento = DateTime.Now
    '    boleto.DataDocumento = DateTime.Now
    '    boleto.Carteira = "03"
    '    boleto.Sacado = New Sacado() With {
    '        .CPFCNPJ = "05461883893",
    '        .Nome = "Joãozinho Testador",
    '        .Endereco = New Endereco() With {
    '            .Complemento = "Bla bla",
    '            .Numero = "80",
    '            .Bairro = "",
    '            .Cep = "32310535",
    '            .Cidade = "Contagem",
    '            .UF = "Minas Gerais"
    '        }
    '    }
    '    boleto.CodigoBarra.CodigoBanco = "743"
    '    boleto.CodigoBarra.Moeda = 9
    '    boleto.CodigoBarra.CampoLivre = "0001023514837340110996818"
    '    boleto.CodigoBarra.ValorDocumento = "0000025151"
    '    boleto.CodigoBarra.FatorVencimento = 7363
    '    Dim linhaDigitavel = boleto.CodigoBarra.LinhaDigitavelFormatada
    '    boleto.CodigoBarra.Codigo = boleto.CodigoBarra.LinhaDigitavelFormatada.Trim().Replace(".", String.Empty).Replace(" ", String.Empty)
    '    Dim boletoBancario = New BoletoBancario With {
    '        .CodigoBanco = 743,
    '        .boleto = boleto,
    '        .MostrarEnderecoCedente = True,
    '        .MostrarContraApresentacaoNaDataVencimento = False,
    '        .GerarArquivoRemessa = True
    '    }
    '    Return boletoBancario
    'End Function
End Class
