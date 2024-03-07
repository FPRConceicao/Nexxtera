Imports Teleatlantic.TLS.Common

Public Class LogEnvioArquivoRemessa : Inherits Retorno
    Public Property Id As Integer
    Public Property EnviaBanco As Boolean
    Public Property CodBanco As String
    Public Property NumCta As String
    Public Property CodAgen As String
    Public Property Tipo As String
    Public Property Cartao As Boolean
    Public Property Empresa As String
    Public Property DtHoraGeracao As DateTime
    Public Property UsrGeracao As String
    Public Property NumRemessa As Integer
    Public Property IsOptante As Boolean
    Public Property DtHoraEnvioRemessa As DateTime
    Public Property UsrEnvioRemessa As String
    Public Property EnviaBancoDesc As String
    Public Property NomeBanco As String
    Public Property NomeUsr As String
    Public Property NomeUsrEnvioRemessa As String
    Public Property CartaoDesc As String


End Class
