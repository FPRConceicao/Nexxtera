
''' <summary>
''' Rotinas genéricas
''' </summary>
''' <remarks>
''' 
''' Data Criação:     08/04/2011
''' Auttor:           Edson Ferreira
''' 
''' Modificações: 
''' 08/04/2011
''' EDF - TL200001 - Rotinas genéricas
''' Autor da Modificação: Edson Ferreira 
''' 
''' </remarks>
Public Class DadosGenericos : Implements IDisposable

    ' booleano para controlar se
    ' o método Dispose já foi chamado
    Dim disposed As Boolean = False


    Public Const EMAIL_COMPRAS As String = "compras@teleatlantic.com.br"
    Public Const EMAIL_MALA_DIRETA_FINANCEIRO As String = "hercules.dealmeida@verisure.com.br"

    Protected Overridable Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    ' método privado para controle
    ' da liberação dos recursos
    Private Sub Dispose(ByVal disposing As Boolean)
        ' Verifique se Dispose já foi chamado.
        If Not Me.disposed Then

            If disposing Then
                ' Liberando recursos gerenciados
            End If
            ' Seta a variável booleana para true,
            ' indicando que os recursos já foram liberados
            disposed = True
        End If
    End Sub
    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub

    Public Enum TipoErro As Byte
        None = 0
        Arquitetura = 1
        Funcional = 2
    End Enum


    Public Enum TipoExportacao As Byte
        PDF = 5
        EXCEL = 4
        WORD = 12
        HTML = 7
    End Enum

    Public Enum PontoEletronicoStatus As Byte
        SaidaIntervalo1 = 1
        EntradaIntervalo1 = 2
        SaidaIntervalo2 = 3
        EntradaIntervalo2 = 4
        Saida = 5
        SaidaIntervalo3 = 6
        EntradaIntervalo3 = 7
    End Enum

    Public Enum PegaIdioma As Byte
        Portugues = 0
        Ingles = 1
        espanhol = 2
    End Enum

    Public Enum ImagemRetorno As Byte
        CampoBranco = 0
        Alerta = 1
        Erro = 2
    End Enum

    Public Enum Servicos As Byte
        Supervisao_Motorizada = 0
        Radio = 1
    End Enum

    Public Enum Timeout As Integer
		Faturamento = 2000
		Query = 120
        FaturamentoCartaoCredito = 900
        Boletagem = 1600
		ContImprodutivos = 1800
		ListaClientes = 1000
	End Enum

    Public Enum Chamado As Integer

        APonto = 1
        AgendaInstalacao = 2
        InclusaoServico = 3
        SolicitacaoCancelamento = 4
        AltCadastro = 5
        Adesao = 6
        Manutecao = 7
        Reclamacao = 8
        MudancaRazaoSocial = 9
        Reativacao = 10
        BaixaTitulo = 11
        TeleVideo = 12
        DescAcrecMonitoria = 13
        ExclusaoServico = 14
        CancelamentoNF = 15

    End Enum

    Public Enum TipoServicoParcela
        Compra = 0
        Venda = 1
    End Enum

    Public Property Descricao() As [String]
        Get
            Return m_Descricao
        End Get
        Private Set(ByVal value As [String])
            m_Descricao = value
        End Set
    End Property
    Private m_Descricao As [String]

    Public Property Id() As [String]
        Get
            Return m_Id
        End Get
        Private Set(ByVal value As [String])
            m_Id = value
        End Set
    End Property
    Private m_Id As [String]


    Private Sub New(ByVal descricao As String, Optional ByVal Id As String = "")
        m_Descricao = descricao
        m_Id = Id
    End Sub

    Public Enum TipoEqpto As Integer
        RADIO = 1
        BUTTON = 2
    End Enum

    Public Shared ENVIO_EMAIL_ADITIVO_CONTRATUAL As New DadosGenericos("Aditivo Contratual" & vbLf)
    Public Shared ENVIO_EMAIL_FORMALIZACAO_DE_ACORDO As New DadosGenericos("Formalização de Acordo" & vbLf)
    Public Shared ENVIO_EMAIL_ORDEMINSTALACAO As New DadosGenericos("Ordem de Instalação" & vbLf)

    Public Shared ModoCadastroINSERE As New DadosGenericos("Insere", "0")
    Public Shared ModoCadastroALTERAR As New DadosGenericos("Alterar", "01")
    Public Shared ModoCadastroPRECADASTRO As New DadosGenericos("Pré-Cadastro", "03")
    Public Shared ModoCadastroENCAMINHAR As New DadosGenericos("Encaminhar", "04")
    Public Shared ModoCadastroCONSULTAR As New DadosGenericos("Consultar", "05")
    Public Shared ModoCadastroFINALIZAR As New DadosGenericos("Finalizar", "06")
    Public Shared ModoCadastroCONFERIR As New DadosGenericos("Conferir", "07")

    Public Shared STATUS_CLIENTE_ATIVO_BONIFICADO As New DadosGenericos("ATIVO - BONIFICADO", "01")
    Public Shared STATUS_CLIENTE_CANCELAMENTO_EM_NEGOCIACAO As New DadosGenericos("CANCELAMENTO EM NEGOCIACAO", "02")
    Public Shared STATUS_CLIENTE_CANCELADO As New DadosGenericos("CANCELADO", "03")
    Public Shared STATUS_CLIENTE_ATIVO_COMODATO As New DadosGenericos("ATIVO - COMODATO", "04")
    Public Shared STATUS_CLIENTE_ATIVO As New DadosGenericos("ATIVO", "05")
    Public Shared STATUS_CLIENTE_COBRANCA_CENTRALIZADA As New DadosGenericos("COBRANÇA CENTRALIZADA", "06")
    Public Shared STATUS_CLIENTE_CANCELADO_MUDANÇA_RSOCIAL As New DadosGenericos("CANCELADO - MUDANÇA  R. SOCIAL", "07")
    Public Shared STATUS_CLIENTE_CANCELADO_INADIMPLENCIA As New DadosGenericos("CANCELADO - INADIMPLENCIA", "08")
    Public Shared STATUS_CLIENTE_ATIVO_INADIMPLENTE As New DadosGenericos("ATIVO - INADIMPLENTE", "09")
    Public Shared STATUS_CLIENTE_BLOQUEADO_SAT As New DadosGenericos("BLOQUEADO SAT", "10")
    Public Shared STATUS_CLIENTE_ATIVO_PEDIDO_NAO_FINALIZADO As New DadosGenericos("ATIVO - PEDIDO NÃO FINALIZADO", "11")
    Public Shared STATUS_CLIENTE_CANCELADO_COBRANCA_EXTERNA As New DadosGenericos("CANCELADO COBRANÇA EXTERNA", "12")
    Public Shared STATUS_CLIENTE_ATIVO_CONTRATO_MANUTENCAO As New DadosGenericos("ATIVO - CONTRATO MANUTENÇÃO", "13")
    Public Shared STATUS_CLIENTE_ATIVO_RETENCAO_EXTERNA As New DadosGenericos("ATIVO - RETENCAO EXTERNA", "14")
    Public Shared STATUS_CLIENTE_ATIVO_COBRANCA_EXTERNA As New DadosGenericos("ATIVO COBRANÇA EXTERNA", "15")
    Public Shared STATUS_CLIENTE_ATIVO_VERISURE As New DadosGenericos("ATIVO - VERISURE", "16")
    Public Shared STATUS_CLIENTE_ATIVO_VERISURE_MONITORING As New DadosGenericos("ATIVO - VERISURE (MONITORING)", "17")

    Public Shared LOG_FICHACONFIDENCIAL_PRECADASTRO As New DadosGenericos("PRE CADASTRO DA FICHA CONF. EFETUADO", "01")
    Public Shared LOG_FICHACONFIDENCIAL_CADASTROLIBERADOPARACENTRAL As New DadosGenericos("CAD. LIB. CENTRAL", "02")
    Public Shared LOG_FICHACONFIDENCIAL_FAX_RECEBIDO As New DadosGenericos("FAX DA FICHA CONF. RECEBIDO", "03")

    Public Shared IdiomaBr As New DadosGenericos("Pt")
    Public Shared IdiomaEn As New DadosGenericos("En")
    Public Shared IdiomaEs As New DadosGenericos("Es")

    Public Shared TipoManutencaoSM1 As New DadosGenericos("SM1")
    Public Shared TipoManutencaoSM2 As New DadosGenericos("SM2")
    Public Shared TipoManutencaoSM3 As New DadosGenericos("SM3")
    Public Shared TipoManutencaoSM4 As New DadosGenericos("SM4")

    Public Shared Pessoa_Fisica As New DadosGenericos("Fisica", "F")
    Public Shared Pessoa_Juridica As New DadosGenericos("Juridica", "J")
    Public Shared Pessoa_Outros As New DadosGenericos("Outros", "O")

    Public Shared Ativo As New DadosGenericos("A")
    Public Shared Inativo As New DadosGenericos("I")
    Public Shared Cancelado As New DadosGenericos("C")
    Public Shared Quitado As New DadosGenericos("Q")

    Public Shared SIM As New DadosGenericos("S")
    Public Shared NAO As New DadosGenericos("N")
    Public Shared TODOS As New DadosGenericos("Todos")

    'TeleEmergencia
    Public Shared Residencial As New DadosGenericos("Residencial", "0")
    Public Shared Comercio0_50 As New DadosGenericos("Comércio 1 à 50 Funcionários", "1")
    Public Shared Comercio51_200_ As New DadosGenericos("Comércio 51 à 200 Funcionários", "2")
    Public Shared Comercio201 As New DadosGenericos("Comércio 201 Funcionários", "3")

    Public Shared TipoUsuarioMASTER As New DadosGenericos("M")
    Public Shared TipoUsuarioUSUARIO As New DadosGenericos("U")
    Public Shared TipoUsuarioSUPERVISOR As New DadosGenericos("S")

    ''Filiais Teleatlantic
    Public Shared Filial_MATRIZ As New DadosGenericos("Matriz", "1")
    Public Shared Filial_CAMPINAS As New DadosGenericos("Campinas", "2")
    Public Shared Filial_RIO_DE_JANEIRO As New DadosGenericos("Rio de Janeiro", "3")
    Public Shared Filial_SANTOS As New DadosGenericos("Santos", "4")

    Public Shared STATUSBLOQUEADA As New DadosGenericos("B")
    Public Shared STATUSLIBERADA As New DadosGenericos("L")
    Public Shared STATUSNEGADA As New DadosGenericos("N")
    Public Shared STATUSEXCLUIDA As New DadosGenericos("X")
    Public Shared STATUSFATURADA As New DadosGenericos("F")


    ''Departamento - a descrição e o Id estão iguais ao da tabela Departamento 
    ''Obs: somente altere a descrição e o id caso houver uma alteração na tabela Departamento
    Public Shared DepartamentoINSTALACAO As New DadosGenericos("Instalação", "01")
    Public Shared DepartamentoADMFINANCEIRO As New DadosGenericos("Adm / Financeiro", "02")
    Public Shared DepartamentoCENTRALOPERACOES As New DadosGenericos("Central de Operações", "04")
    Public Shared DepartamentoSAT As New DadosGenericos("SAT", "06")
    Public Shared DepartamentoMANUTENCAO As New DadosGenericos("Manutenção", "07")
    Public Shared DepartamentoSISTEMAS As New DadosGenericos("Sistemas", "08")
    Public Shared DepartamentoPRESIDENCIA As New DadosGenericos("Presidência", "09")
    Public Shared DepartamentoRECURSOSHUMANOS As New DadosGenericos("Recursos Humanos", "11")
    Public Shared DepartamentoDIRETORIAGERAL As New DadosGenericos("Diretoria Geral", "12")
    Public Shared DepartamentoTI As New DadosGenericos("TI", "13")
    Public Shared DepartamentoPOSVENDA As New DadosGenericos("Pós Venda", "14")
    Public Shared DepartamentoCOMERCIAL As New DadosGenericos("Comercial", "145")
    Public Shared DepartamentoSUPERVISAOMOTORIZADA As New DadosGenericos("Supervisão Motorizada", "150")
    Public Shared DepartamentoJURIDICO As New DadosGenericos("Jurídico / Compliance", "18")
    Public Shared DepartamentoRETENCAO As New DadosGenericos("Retenção", "19")
    Public Shared DepartamentoCADASTRO As New DadosGenericos("Cadastro", "20")
    Public Shared DepartamentoCOBRANCA As New DadosGenericos("Cobrança", "21")
    Public Shared DepartamentoOUVIDORIA As New DadosGenericos("Ouvidoria", "23")
    Public Shared DepartamentoQUALIDADEEPERFORMANCE As New DadosGenericos("Qualidade e Performance", "22")
    Public Shared DepartamentoMARKETING As New DadosGenericos("Marketing", "10")

    ''Setor - a descrição e o Id estão iguais ao da tabela Departamento_Setor 
    ''Obs: somente altere a descrição e o id caso houver uma alteração na tabela Departamento_Setor
    Public Shared SetorFINANCEIRO As New DadosGenericos("Financeiro", "02")
    Public Shared SetorCENTRAL As New DadosGenericos("Central", "04")
    Public Shared SetorPRESIDENCIA As New DadosGenericos("Presidência", "09")
    Public Shared SetorSAT As New DadosGenericos("SAT", "06")
    Public Shared SetorMANUTENCAO As New DadosGenericos("Manutenção", "07")
    Public Shared SetorCOMERCIAL As New DadosGenericos("Comercial", "145")
    Public Shared SetorCONEXAO As New DadosGenericos("Conexão", "163")
    Public Shared SetorDIRETORIAGERAL As New DadosGenericos("Diretoria Geral", "12")
    Public Shared SetorVISITA As New DadosGenericos("Visita", "156")
    Public Shared SetorGARANTIA As New DadosGenericos("Garantia", "157")
    Public Shared SetorADMINISTRATIVO As New DadosGenericos("Administrativo", "158")
    Public Shared SetorAUMENTODEPONTO As New DadosGenericos("Aumento de Ponto", "159")
    Public Shared SetorTI As New DadosGenericos("TI", "13")
    Public Shared SetorSUPERVISORDEVENDAS As New DadosGenericos("Supervisor de Vendas", "92")
    Public Shared SetorSUPERVISAO As New DadosGenericos("Supervisão", "150")
    Public Shared SetorQUALIDADE As New DadosGenericos("Qualidade", "17")
    Public Shared SetorRETENCAO As New DadosGenericos("Retenção", "112")
    Public Shared SetorSISTEMAS As New DadosGenericos("Sistemas", "08")
    Public Shared SetorTELEVIDEO As New DadosGenericos("Tele Video", "160")
    Public Shared SetorSERVICOSGERAIS As New DadosGenericos("Serviços Gerais", "10")
    Public Shared SetorRECURSOSHUMANOS As New DadosGenericos("Recursos Humanos", "11")
    Public Shared SetorPOSVENDA As New DadosGenericos("Pós Venda", "14")
    Public Shared SetorCOBRANCA As New DadosGenericos("Cobrança", "114")
    Public Shared SetorOUVIDORIA As New DadosGenericos("Ouvidoria", "116")
    Public Shared SetorGPRS As New DadosGenericos("GPRS", "161")
    'Public Shared SetorRETENCAOEXTERNA As New DadosGenericos("Retenção Externa", "115")
    Public Shared SetorRETENCAOEXTERNA As New DadosGenericos("Retenção Externa", "121")
    Public Shared SetorCONTASARECEBER As New DadosGenericos("Contas a Receber", "108")
    Public Shared SetorPROJETOS As New DadosGenericos("Projetos", "152")
    Public Shared SetorMARKETING As New DadosGenericos("Marketing", "10")
    Public Shared SetorAuditoriaEventos As New DadosGenericos("Auditoria de Eventos", "115")
    Public Shared SetorCadastro As New DadosGenericos("Cadastro", "113")
    Public Shared SetorONBOARDING As New DadosGenericos("Onboarding", "119")
    Public Shared SetorCALLCENTERMARKETING As New DadosGenericos("Call Center Marketing", "111")
    Public Shared SetorBACKOFFICE As New DadosGenericos("BACK OFFICE", "103")

    'DIA DA SEMANA
    Public Shared Semana_Seg_a_Sex As New DadosGenericos("Seg à Sex")
    Public Shared Semana_Segunda As New DadosGenericos("Segunda - Feira")
    Public Shared Semana_Terca As New DadosGenericos("Terça - Feira")
    Public Shared Semana_Quarta As New DadosGenericos("Quarta - Feira")
    Public Shared Semana_Quinta As New DadosGenericos("Quinta - Feira")
    Public Shared Semana_Sexta As New DadosGenericos("Sexta - Feira")
    Public Shared Semana_Sabado As New DadosGenericos("Sábado")
    Public Shared Semana_Domingo As New DadosGenericos("Domingo")

    'TIPO ORIGEM
    Public Shared TIPO_ORIGEM_CONCILIACAOBANCARIA As New DadosGenericos("CONCILIACAO", "C")
    Public Shared TIPO_ORIGEM_MOVCCORRENTE As New DadosGenericos("MOVCCORRENTE", "M")
    Public Shared TIPO_ORIGEM_AUMENTODEPONTO As New DadosGenericos("Aum. Pto.")
    Public Shared TIPO_ORIGEM_COBRANCA As New DadosGenericos("Cobranca", "COBR")
    Public Shared TIPO_ORIGEM_MUNDANCARAZAOSOCIAL As New DadosGenericos("MRS")
    Public Shared TIPO_ORIGEM_CLIENTE As New DadosGenericos("CLIENTE")
    Public Shared TIPO_ORIGEM_PEDIDO As New DadosGenericos("PEDIDO")
    Public Shared TIPO_ORIGEM_PORDEM_INSTALACAO As New DadosGenericos("ORDEMINSTALACAO")
    Public Shared TIPO_ORIGEM_EXECUCAOPROCEDIMENTO As New DadosGenericos("EXECUCAOPROCEDIMENTO")
    Public Shared TIPO_ORIGEM_MANUTENCAO As New DadosGenericos("MANUT", "MANUT")
    Public Shared TIPO_ORIGEM_DEBTOCREDITO As New DadosGenericos("DÉBITO/CRÉDITO")
    Public Shared TIPO_ORIGEM_INDICACAO As New DadosGenericos("INDICACAO", "I")
    Public Shared TIPO_ORIGEM_PROSPECCAO As New DadosGenericos("PROSPECCAO", "P")
    Public Shared TIPO_ORIGEM_ARRASTAO As New DadosGenericos("ARRASTAO", "A")
    Public Shared TIPO_ORIGEM_GERENCIAL As New DadosGenericos("GERENCIAL", "G")
    Public Shared TIPO_ORIGEM_ESTOQUE As New DadosGenericos("ESTOQUE", "E")
    Public Shared TIPO_ORIGEM_NFE As New DadosGenericos("NFE", "N")
    Public Shared TIPO_ORIGEM_CC As New DadosGenericos("CC", "C")
    Public Shared TIPO_ORIGEM_IMPBORDERO As New DadosGenericos("IMPBORDERO", "B")
    Public Shared TIPO_ORIGEM_IMPVERSOBORDERO As New DadosGenericos("IMPVERSOBORDERO", "VB")
    Public Shared TIPO_ORIGEM_REQUISICAO As New DadosGenericos("Requis", "R")
    Public Shared TIPO_ORIGEM_DEVOLUCAO As New DadosGenericos("DEV", "D")
    Public Shared TIPO_ORIGEM_AGENDAMENTO As New DadosGenericos("AGD", "AG")
    Public Shared TIPO_ORIGEM_CONTROLEINSTALACAO As New DadosGenericos("CONTINST", "CI")
    Public Shared TIPO_ORIGEM_HELPDESKPRINCIPAL As New DadosGenericos("HELPDESKPRINCIPAL", "HDP")
    Public Shared TIPO_ORIGEM_HELPDESKEXECUCAO As New DadosGenericos("HELPDESKEXECUAO", "HDE")
    Public Shared TIPO_ORIGEM_NOTAFISCAL As New DadosGenericos("NOTAFISCAL", "NF")
    Public Shared TIPO_ORIGEM_COMISSAOREPRESENTANTE As New DadosGenericos("COMISSAOREPRESENTANTE", "CR")
    Public Shared TIPO_ORIGEM_DEALERPROGRAM As New DadosGenericos("DEALERPROGRAM", "DP")
    Public Shared TIPO_ORIGEM_NOTASFORNECEDORES As New DadosGenericos("NOTASFORNECEDORES", "NFF")
    Public Shared TIPO_ORIGEM_TELE_VENDAS As New DadosGenericos("TELEVENDAS", "TV")
    Public Shared TIPO_ORIGEM_RETENCAO As New DadosGenericos("RETENCAO", "RET")
    Public Shared TIPO_ORIGEM_BLOQUEIO As New DadosGenericos("BLOQUEIO", "BLQ")
    'TELAS PARA LOG - TELESYSTEM

    ''' <summary>
    ''' Esta constante foi criada para que na busca de um log na 'tLogChangeData' para efeito
    ''' de histórico de alteração na tela do cliente. para inserção o frmClieEdi foi substituido pela
    ''' constante LOG__MANUTENCAO_CLIENTE_TELESYSTEM2 que contem o seguite texto 'ManutCliente'
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared LOG__MANUTENCAO_CLIENTE_TELESYSTEM As New DadosGenericos("frmClieEdi")



    'TELAS PARA LOG - TELESYSTEM 2
    ''' <summary>
    ''' Esta constante está sendo gravada na 'tLogChangeData' no campo FORM  no lugar da 'frmClieEdi' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared LOG__MANUTENCAO_CLIENTE_TELESYSTEM2 As New DadosGenericos("ManutCliente")


    'TELAS PARA LOG - TELESYSTEM 2
    ''' <summary>
    ''' Esta constante está sendo gravada na 'tLogChangeData' no campo FORM  no lugar da 'frmOrdInst' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared LOG__ORDEM_INST_TELESYSTEM2 As New DadosGenericos("frmOrdInst")

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared LOG__VENDADEBITOAUTOMATICO_TELESYSTEM2 As New DadosGenericos("VendaDebitoAutomatico")

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared LOG__REVERSAO_TELESYSTEM2 As New DadosGenericos("Reversao")

    Public Shared LOG__EXECUCAOPROCEDIMENO_TELESYSTEM2 As New DadosGenericos("ExecucaoProcedimento")

    'CAMPOS PARA LOG - TELESYSTEM 
    Public Shared LOG__MANUTENCAO_CLIENTE_OPTLIBFAT_TELESYSTEM As New DadosGenericos("optLibFat")
    Public Shared LOG__BOTON_TELESYSTEM As New DadosGenericos("BUTTON")
    Public Shared LOG__RADIO_TELESYSTEM As New DadosGenericos("RADIO")
    Public Shared LOG__GPRS_TELESYSTEM As New DadosGenericos("GPRS")

    'CAMPO PARA LOG - TELESYSTEM 2
    Public Shared LOG__MANUTENCAO_CLIENTE_TXTSTATUS_TELESYSTEM2 As New DadosGenericos("txtStatus")
    Public Shared LOG__MANUTENCAO_CLIENTE_CHKLIBERADO_TELESYSTEM2 As New DadosGenericos("optLibFat")
    Public Shared LOG__ALTERACAO_INCLUSAO_OS_TXTPROTOCOLO_TELESYSTEM2 As New DadosGenericos("txtProtocolo")
    Public Shared LOG__BOTON_TELESYSTEM2 As New DadosGenericos("BUTTON")
    Public Shared LOG__RADIO_TELESYSTEM2 As New DadosGenericos("RADIO")
    Public Shared LOG__GPRS_TELESYSTEM2 As New DadosGenericos("GPRS")
    Public Shared LOG__REVERSAO_SUCESSO_TELESYSTEM2 As New DadosGenericos("REVERSAO")
    'Motivo ção
    Public Shared MotivoReclaRECLAMACAO_DO_CLIENTE As New DadosGenericos("Reclamação do cliente", "000002")
    Public Shared MotivoReclaMANUTENCAO As New DadosGenericos("Manutenção", "000003")
    Public Shared MotivoReclaCARTA_DE_REAJUSTE As New DadosGenericos("Carta de Reajuste", "000005")
    Public Shared MotivoReclaCARTA_DE_CORRECAO As New DadosGenericos("Carta de Correção", "000008")
    Public Shared MotivoReclaINSATISFACAO As New DadosGenericos("Insatisfação", "000013")
    Public Shared MotivoReclaINSATISFACAO_COM_SERVICOS_AGREGADOS As New DadosGenericos("Insatisfação com Serviços Agregados", "000014")
    Public Shared MotivoReclaCONCORRENCIA As New DadosGenericos("Concorrência", "000017")
    Public Shared MotivoReclaVISITA_RETIRADA_EQUIPAMENTO As New DadosGenericos("Visita - Retirada equipamentos", "000018")
    Public Shared MotivoReclaVISITA_AUMENTO_PONTO As New DadosGenericos("Visita - Aumento de ponto", "000019")
    Public Shared MotivoReclaDISPAROS_FALSOS As New DadosGenericos("Disparos falsos", "000021")
    Public Shared MotivoReclaENVIO_DE_ORCAMENTO As New DadosGenericos("Envio de Orçamento", "000022")
    Public Shared MotivoReclaRETORNO_TROCA_EQUIPAMENTO As New DadosGenericos("Retorno troca de equipamento", "000023")
    Public Shared MotivoRecla2VIA_DE_BOLETO_BANCARIO As New DadosGenericos("2ª Via de Boleto Bancário", "000025")
    Public Shared MotivoReclaEMISSAO_DE_NOTA_FISCAL As New DadosGenericos("Emissão de Nota Fiscal", "000028")
    Public Shared MotivoReclaALTERACAO_DE_CADASTRO As New DadosGenericos("Alteração de cadastro", "000029")
    Public Shared MotivoReclaDEMORA_DA_SUPERVISAO As New DadosGenericos("Demora da supervisão", "000031")
    Public Shared MotivoReclaTROCAR_PLACA_DE_PROTEGIDO As New DadosGenericos("Trocar placa de protegido", "000043")
    Public Shared MotivoReclaREATIVACAO_DO_ALARME As New DadosGenericos("Reativação do alarme", "000045")
    Public Shared MotivoReclaVISITA_AO_CLIENTE As New DadosGenericos("Visita ao cliente", "000050")
    Public Shared MotivoReclaREATIVACAO As New DadosGenericos("Reativação", "000051")
    Public Shared MotivoReclaLIGACAO_DA_GERENCIA As New DadosGenericos("Ligação da gerência", "000052")
    Public Shared MotivoReclaEMISSAO_DE_RELATORIO As New DadosGenericos("Emissão de relatório", "000053")
    Public Shared MotivoReclaRETENCAO_DE_IMPOSTOS As New DadosGenericos("Retenção de Impostos", "000055")
    Public Shared MotivoReclaDISPAROS_FREQUENTES As New DadosGenericos("Disparos frequentes", "000056")
    Public Shared MotivoReclaBATERIA_FRACA As New DadosGenericos("Bateria Fraca", "000059")
    Public Shared MotivoReclaPROBLEMA_COM_A_CERCA As New DadosGenericos("Problema com a Cerca", "000060")
    Public Shared MotivoReclaRETIRADA_EQUIPAMENTO_COMODATO As New DadosGenericos("Retirada Equipamento em Comodato", "000064")
    Public Shared MotivoReclaVISITA_REINSTALACAO_DE_PONTO As New DadosGenericos("Visita - Reinstalação de ponto", "000067")
    Public Shared MotivoReclaCOMUNICACAO_INTERNA As New DadosGenericos("Comunicação interna", "000073")
    Public Shared MotivoReclaSOLICITACAO_CANC_MONITORAMENTO_SIMS As New DadosGenericos("Solicitacao Canc. Monitoramento no SIMS", "000075")
    Public Shared MotivoReclaMUDANCA_LINHA_TELEFONICA As New DadosGenericos("Mudança de linha telefônica", "000076")
    Public Shared MotivoReclaVISITA_MUDANCA_ENDERECO As New DadosGenericos("Visita - Mudança de endereço", "000079")
    Public Shared MotivoReclaCANCELOU_ANTES_DA_INSTALACAO As New DadosGenericos("Cancelou antes da instalação", "000081")
    Public Shared MotivoReclaMUDANCA_RAZAO_SOCIAL As New DadosGenericos("Mudança de razão social", "000082")
    Public Shared MotivoReclaMUDANCA_RAZAO_SOCIAL_SAT As New DadosGenericos("Mudança de Razão Social-SAT", "000084")
    Public Shared MotivoReclaLINHA_TELEFONICA_COM_PROBLEMAS As New DadosGenericos("Linha telefônica c/ problemas", "000089")
    Public Shared MotivoReclaZONA_ABERTA_FIACAO_ROMPIDA As New DadosGenericos("Zona aberta/Fiação rompida", "000090")
    Public Shared MotivoReclaDUVIDAS_DO_CLIENTE_REENTREGA As New DadosGenericos("Duvidas do cliente / Reentrega", "000091")
    Public Shared MotivoReclaRETORNO_TERMINO_INSTALACAO As New DadosGenericos("Retorno -Término da instalação", "000092")
    Public Shared MotivoReclaSEM_SINAIS_DE_MONITORAMENTO As New DadosGenericos("Sem sinais no monitoramento", "000093")
    Public Shared MotivoReclaBLOQUEIO_TEMPORARIO_REFORMA As New DadosGenericos("Bloqueio Temporário - Reforma", "000095")
    Public Shared MotivoReclaDAD_CADASTRAIS_RAZAO_SOCIAL_ABREVIADA As New DadosGenericos("Dad. Cadastrais - Razão Social Abreviada", "000096")
    Public Shared MotivoReclaDAD_CADASTRAIS_RAZAO_SOCIAL_DIVERGENTE As New DadosGenericos("Dad. Cadastrais - Razão Social Divergente CNPJ", "000097")
    Public Shared MotivoReclaDAD_CADASTRAIS_NOME_DIVERGENTE_CPF As New DadosGenericos("Dad. Cadastrais - Nome Divergente CPF", "000098")
    Public Shared MotivoReclaDAD_CADASTRAIS_NOME_ABREVIADO As New DadosGenericos("Dad. Cadastrais - Nome Abreviado", "000099")
    Public Shared MotivoReclaDAD_CADASTRAIS_ENDERECO_INCOMPLETO As New DadosGenericos("Dad. Cadastrais - Endereço Incompleto", "000100")
    Public Shared MotivoReclaDAD_CADASTRAIS_ENDERECO_COBRANCA_DIVERGENTE As New DadosGenericos("Dad. Cadastrais - Endereço de Cobrança Divergente", "000101")
    Public Shared MotivoReclaDAD_CADASTRAIS_ENDERECO_INSTALACAO_DIVERGENTE As New DadosGenericos("Dad. Cadastrais- Endereço de Instalação Divergente", "000102")
    Public Shared MotivoReclaDAD_CADASTRAIS_ENDERECO_ERRADO As New DadosGenericos("Dad. Cadastrais - Endereço Errado", "000103")
    Public Shared MotivoReclaDAD_CADASTRAIS_NR_TELEFONE_NAO_ATENDE As New DadosGenericos("Dad. Cadastrais - Nº Telefone não Atende", "000104")
    Public Shared MotivoReclaDAD_CADASTRAIS_NR_TELEFONE_NAO_EXISTE As New DadosGenericos("Dad. Cadastrais - Nº Telefone não Existe", "000105")
    Public Shared MotivoReclaDAD_CADASTRAIS_NR_TELEFONE_ERRADO As New DadosGenericos("Dad. Cadastrais - Nº Telefone Errado", "000106")
    Public Shared MotivoReclaDAD_CADASTRAIS_NOME_DO_CONTATO_INCOMPLETO As New DadosGenericos("Dad. Cadastrais - Nome do Contato Incompleto", "000107")
    Public Shared MotivoReclaDAD_CADASTRAIS_NOME_DO_CONTATOERRADO As New DadosGenericos("Dad. Cadastrais - Nome do Contato Errado", "000108")
    Public Shared MotivoReclaDAD_CADASTRAIS_INSCRICAO_ESTADUAL_ERRADA As New DadosGenericos("Dad. Cadastrais - Inscrição Estadual Errada", "000109")
    Public Shared MotivoReclaDAD_CADASTRAIS_FALTA_INSCRICAO_MUNICIPAL As New DadosGenericos("Dad. Cadastrais - Falta Inscrição Municipal", "000110")
    Public Shared MotivoReclaDAD_CADASTRAIS_INSCRICAO_MUNICIPAL_ERRAD As New DadosGenericos("Dad. Cadastrais - Inscrição Municipal Errada", "000111")
    Public Shared MotivoReclaDAD_CADASTRAIS_RG_ERRADO As New DadosGenericos("Dad. Cadastrais - RG Errado", "000112")
    Public Shared MotivoReclaRECEPTIVO_COBRANCA As New DadosGenericos("Receptivo Cobrança", "001010")
    Public Shared MotivoReclaCANCELAMENTO_COMPULSORIO As New DadosGenericos("Cancelamento Compulsório", "000186")
    Public Shared MotivoReclaREVERSAO_MOTIVO_MUDANCA_LOCAL As New DadosGenericos("Reversão Motivo Mudança de Local", "001061")
    Public Shared MotivoReclaVISITA_VENDEDOR_MUDANCA_DE_LOCAL As New DadosGenericos("Visita Vendedor - MUDANÇA DE LOCAL", "001060")
    Public Shared MotivoReclaCONTATO_OUVIDORIA As New DadosGenericos("Contato Ouvidoria", "000816")
    Public Shared MotivoReclaVISITA_ENVIO_ORCAMENTO As New DadosGenericos("Visita - Envio de orçamento", "001719")
    Public Shared MotivoReclaDUVIDAS_CLIENTE As New DadosGenericos("Dúvidas do Cliente", "001395")
    Public Shared MotivoReclaRETIRADA_EQUIPAMENTO_CONEXAO As New DadosGenericos("Retirada Equipamento - Conexão", "002267")



    Public Shared Cargo_OPERADORI As New DadosGenericos("Operador I", "1")
    Public Shared Cargo_OPERADORII As New DadosGenericos("Operador II", "2")
    Public Shared Cargo_SECRETARIA As New DadosGenericos("Secretária", "3")
    Public Shared Cargo_DIRETOR As New DadosGenericos("Diretor", "4")
    Public Shared Cargo_ASSISTENTEADMINISTRATIVOJUNIOR As New DadosGenericos("Assistente Administrativo Junior", "5")
    Public Shared Cargo_AUXILIARDERETENCAO As New DadosGenericos("Auxiliar de Retenção", "6")
    Public Shared Cargo_ASSISTENTEADMINISTRATIVO As New DadosGenericos("Assistente Administrativo", "7")
    Public Shared Cargo_ARQUIVISTA As New DadosGenericos("Arquivista", "8")
    Public Shared Cargo_COORDENADOR As New DadosGenericos("Coordenador", "9")
    Public Shared Cargo_SUPERVISOR As New DadosGenericos("Supervisor", "10")
    Public Shared Cargo_TECNICOONLINE As New DadosGenericos("Técnico On-Line", "11")
    Public Shared Cargo_REPRESENTANTE As New DadosGenericos("Representante", "12")
    Public Shared Cargo_COACHING_PERFORMANCE As New DadosGenericos("Coaching Performance", "33")
    Public Shared Cargo_GERENTE As New DadosGenericos("Gerente", "13")
    Public Shared Cargo_SUPERVISORCENTRAL As New DadosGenericos("Supervisor Central", "14")
    Public Shared Cargo_OPERADORCENTRAL As New DadosGenericos("Operador Central", "15")
    Public Shared Cargo_CONSULTORCENTRAL As New DadosGenericos("Consultor Central", "16")
    Public Shared Cargo_LIDERSUPERVISOROPERACIONAL As New DadosGenericos("Lider Supervisor Operacional", "17")
    Public Shared Cargo_OPERADORCENTRAL_TRAFEGO As New DadosGenericos(" Operador Central Tráfego ", "18")
    Public Shared Cargo_OPERADOR_CALL_CENTER As New DadosGenericos("Operador de Call Center", "76")
    Public Shared Cargo_SUPERVISORA_CALL_CENTER As New DadosGenericos("Supervisora Call Center", "73")
    Public Shared Cargo_GERENTE_MARKETING As New DadosGenericos("Gerente de Marketing", "74")
    Public Shared Cargo_ANALISTA_MARKETING As New DadosGenericos("Analista de Marketing", "75")

    'UTILIZADOS NA TELA DE ORÇAMENTO, ENVIO DE ORÇAMENTO
    Public Shared Email_Orcamentos As New DadosGenericos("orcamentos@teleatlantic.com.br", "EO")
    Public Shared Fax_Orcamentos As New DadosGenericos("(11) 3811-1022", "FO")

    ' Status Pedido
    Public Structure StatusPedido
        Public Const INSTALANDO As String = "I"
        Public Const APROVADO As String = "A"
        Public Const DEVOLVIDO As String = "D"
        Public Const CANCELADO As String = "C"
        Public Const FINALIZADO As String = "F"
        Public Const REPROVADO As String = "R"
        Public Const APROVACAOGERENCIAL As String = "G"
    End Structure

    'Forma Pagamento
    Public Structure FormaPagamento
        Public Const CARTAO_CREDITO As String = "CC"
        Public Const BOLETO As String = "BO"
        Public Const DEBITO_CONTA As String = "DC"
        Public Const TRANSFERENCIA As String = "TR"
        Public Const DEPOSITO_BANCARIO As String = "DB"
        Public Const CARTAO_DEBITO As String = "CD"
        Public Const HIBRIDA As String = "HB"
        Public Const CHEQUE As String = "CH"
        Public Const DINHEIRO As String = "DI"
    End Structure

    Public Structure StatusInstalacao
        Public Const AGUARDANDO_CONFERENCIA_DADOS As String = "029"
        Public Const AGUARDANDO_PRE_VISTORIA As String = "043"
        Public Const AGUARDANDO_ANALISE_CARTAO_CREDITO As String = "049"
	End Structure

	Public Enum OrigemCondicaoPagamento
		COMPRA
		VENDA
	End Enum



End Class
