Imports System.Xml.Serialization

''' <summary>
''' Classe responsável pela String de Conexão.
''' </summary>
''' <remarks>
''' 
''' Data Criação:     08/04/2011
''' Auttor:           Edson Ferreira
''' 
''' Modificações: 
''' 08/04/2011
''' EDF - TL200001 - Classe responsável pela String de Conexão.
''' Autor da Modificação: Edson Ferreira
''' 
''' </remarks>
Public Class Conection

    Public ReadOnly Property StringConexao() As String
        Get
            Return m_StringConexao
        End Get

    End Property


    Private m_StringConexao As String

    Private Sub New(ByVal stringconexao As String)
        m_StringConexao = stringconexao

    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="codID">
    ''' 0-Servidor
    ''' 1-Base de dados
    ''' 2-User
    ''' 3-Senha
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ParametroConexao(ByVal codID As Integer) As String
        Dim ConectionString = STRING_CONEXAO.StringConexao.Split(";")
        Dim strDados As String = ""

        Select Case codID
            Case 0
                strDados = ConectionString(0).Replace("Data Source=", "")
            Case 1
                strDados = ConectionString(1).Replace("Initial Catalog=", "")
            Case 2
                strDados = ConectionString(2).Replace("User Id=", "")
            Case 3
                strDados = ConectionString(3).Replace("Password=", "")
        End Select
        Return strDados
    End Function

    Public Shared Function getBD() As String
        Dim strCampoBD As String

        'Servidor OFICIAL
        strCampoBD = "Data Source=LTPRODBR01; Initial Catalog=Telesystem;Integrated Security=SSPI;"

        'Servidor HOMOLOGACAO (Não utilizado mais)
        'strCampoBD = "Data Source=SYSBASETEST; Initial Catalog=TelesystemHomologacao;User Id=verisure;Password=BR25v3riSure#4-87;"

        'Servidor HOMOLOGACAO (Não utilizado mais)
        'strCampoBD = "Data Source=LTPRODBR01; Initial Catalog=Telesystem;User Id=NexxteraAPPUsr;Password=G45#@&kf48&!882L;"

        'Servidor HOMOLOGACAO
        'strCampoBD = "Data Source=BR08SQLPRE02V; Initial Catalog=TelesystemHomologacao;Integrated Security=SSPI;"

#Region "******* Codigo comentado *******"
        'Dim strIsVerisure As String
        'Busca Flag para identificar se o computador é Verisure ou Teleatlantic
        'strIsVerisure = Funcoes.LerRegistroDoWindows("Configurações", "isVerisure", "0")

        'If strIsVerisure.Equals("1") Then
        '09/05/2016 - Fernando 
        'Com a inativação do sharedgateway o IP que a rede Veri enxerga deixa de ser o 10.8.151.6 e passa a ser SYSBASEPROD (192.168.1.232).
        'strCampoBD = "Data Source=SYSBASEPROD; Initial Catalog=Telesystem;User Id=verisure;Password=#VerTel@2001#;"
        'strCampoBD = "Data Source=LTPRODBR01; Initial Catalog=Telesystem;Integrated Security=SSPI;"
        'strCampoBD = "Data Source=10.8.150.10; Initial Catalog=Telesystem; Integrated Security=SSPI;"
        'strCampoBD = "Data Source=SYSBASETEST; Initial Catalog=Telesystem;User Id=verisure;Password=#VerTel@2001#;"
        'strCampoBD = "Data Source=SYSBASETEST; Initial Catalog=Telesystem;User Id=verisure;Password=BR25v3riSure#4-87;"
        'Else
        'strCampoBD = "Data Source=SYSBASEPROD; Initial Catalog=Telesystem; Integrated Security=SSPI;"
        'strCampoBD = "Data Source=SYSBASETEST; Initial Catalog=Telesystem; Integrated Security=SSPI;"
        'End If
#End Region

        Return strCampoBD
    End Function


    ''' <summary>
    ''' String de conexao do Banco de Dados.
    ''' </summary>
    ''' <remarks>
    ''' 
    ''' Data Criação:     08/04/2011
    ''' Autor:           Edson Ferreira
    ''' 
    ''' Modificações: 
    ''' 08/04/2011
    ''' EDF - TL200001 - String de conexao do Banco de Dados.
    ''' Autor da Modificação: Edson Ferreira
    ''' 
    ''' 
    '''  STRINGS UTILIZADAS
    '''  Public Shared STRING_CONEXAO As New Conection("Data Source=AQIJ8IFASEF; Initial Catalog=Telesystem; Integrated Security=SSPI;")
    '''  Public Shared STRING_CONEXAO As New Conection("Data Source=7D7FSV1; Initial Catalog=Telesystem; Integrated Security=SSPI;")
    '''  Public Shared STRING_CONEXAO As New Conection("Data Source=D82SX861; Initial Catalog=Telesystem; Integrated Security=SSPI;")
    '''  Public Shared STRING_CONEXAO As New Conection("Data Source=AQIJ8IFASEF;Initial Catalog=Telesystem;User Id=sa;Password=Ftadm_00;")
    '''  Public Shared STRING_CONEXAO As New Conection("Data Source=D82SX861\D82SX861;Initial Catalog=Telesystem;User Id=sa;Password=Ftadm_00;")
    '''  Public Shared STRING_CONEXAO As New Conection("Data Source=POG-D82;Initial Catalog=Telesystem;User Id=sa;Password=tele4tl4nt!c;")
    '''  Public Shared STRING_CONEXAO As New Conection("Data Source=POG-D82; Initial Catalog=Telesystem; Integrated Security=SSPI;")
    ''' </remarks>


    Public Shared STRING_CONEXAO As New Conection(getBD())


End Class