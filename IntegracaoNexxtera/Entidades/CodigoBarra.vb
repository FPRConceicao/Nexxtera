
Namespace CodigoBarra
    Public Class CodigoBarra
        Public Sub New()

        Public Property CampoLivre As String
        Public Property Chave As String
        Public Property Codigo As String
        Public Property CodigoBanco As String
        Public ReadOnly Property DigitoVerificador As String
        Public Property FatorVencimento As Long
        Public ReadOnly Property Imagem As Image
        Public Property LinhaDigitavel As String
        Public ReadOnly Property LinhaDigitavelFormatada As String
        Public Property Moeda As Integer
        Public Property ValorDocumento As String

        Public Sub PreencheValores(codigoBanco As Integer, moeda As Integer, fatorVencimento As Long, valorDocumento As String, campoLivre As String)
    End Class
End Namespace
