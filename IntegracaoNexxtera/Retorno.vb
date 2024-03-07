Imports Teleatlantic.TLS.Common.DadosGenericos

''' <summary>
''' Classe de tratamento de erros.
''' </summary>
''' <remarks>
''' 
''' Data Criação:     08/04/2011
''' Auttor:           Edson Ferreira
''' 
''' Modificações: 
''' 08/04/2011
''' EDF - TL200001 - Classe de tratamento de erros.
''' Autor da Modificação: Edson Ferreira  
''' 
''' </remarks>
<DebuggerStepThrough()>
Public Class Retorno : Implements IDisposable

    ' booleano para controlar se
    ' o método Dispose já foi chamado
    Dim disposed As Boolean = False

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

    Public Sub New()

    End Sub
    Public Sub New(sucesso As Boolean, msgErro As String)
        Me.Sucesso = sucesso
        Me.MsgErro = msgErro
    End Sub

    Public Sub New(ByVal erro As ErrorConstants, ByVal sucesso As Boolean, ByVal tipoErro As TipoErro, _
                    ByVal imagemErro As ImagemRetorno)

        Me.NumErro = erro.Id
        Me.MsgErro = erro.Descricao
        Me.Sucesso = sucesso
        Me.TipoErro = tipoErro
        Me.ImagemErro = imagemErro

    End Sub

    Private m_NumErro As String
    Private m_MsgErro As String
    Private m_Sucesso As Boolean
    Private m_TipoErro As TipoErro
    Private m_ImagemErro As String

    Public Property NumErro() As String
        Get
            Return m_NumErro
        End Get
        Set(ByVal value As String)
            m_NumErro = value
        End Set
    End Property

    Public Property MsgErro() As String
        Get
            Return m_MsgErro
        End Get
        Set(ByVal value As String)
            m_MsgErro = value
        End Set
    End Property

    Public Property Sucesso() As Boolean
        Get
            Return m_Sucesso
        End Get
        Set(ByVal value As Boolean)
            m_Sucesso = value
        End Set
    End Property

    Public Property TipoErro() As TipoErro
        Get
            Return m_TipoErro
        End Get
        Set(ByVal value As TipoErro)
            m_TipoErro = value
        End Set
    End Property

    Public Property ImagemErro() As ImagemRetorno
        Get
            Return m_ImagemErro
        End Get
        Set(ByVal value As ImagemRetorno)
            m_ImagemErro = value
        End Set
    End Property


    Public Shared Function CriaInstanciaParaException(ByVal ec As ErrorConstants, ex As Exception) As Retorno

        Return New Retorno() With {
            .Sucesso = False,
            .NumErro = ec.Id,
            .MsgErro = ec.Descricao + ex.Message,
            .TipoErro = DadosGenericos.TipoErro.Arquitetura,
            .ImagemErro = DadosGenericos.ImagemRetorno.Erro
        }


    End Function

    Public Shared Function SelecionaImagemRetorno(ByVal tipoErro As DadosGenericos.TipoErro)
        Dim imagemErro As DadosGenericos.ImagemRetorno
        Select Case tipoErro
            Case DadosGenericos.TipoErro.Arquitetura
                imagemErro = ImagemRetorno.Erro
            Case DadosGenericos.TipoErro.Funcional
                imagemErro = ImagemRetorno.Alerta
            Case DadosGenericos.TipoErro.None
                imagemErro = ImagemRetorno.CampoBranco
        End Select


        Return imagemErro
	End Function

	Public Shared Function CriaInstanciaParaAlerta(ByVal ec As ErrorConstants, Optional msgErro As String = "") As Retorno

		Return New Retorno() With {
		 .Sucesso = False,
		 .NumErro = ec.Id,
		 .MsgErro = ec.Descricao + msgErro,
		 .TipoErro = DadosGenericos.TipoErro.Funcional,
		 .ImagemErro = DadosGenericos.ImagemRetorno.Alerta
		}


	End Function


	Public Shared Function CriaRetornoSucesso(Optional msg As String = "") As Retorno

		Return New Retorno() With {
		 .Sucesso = True,
		 .MsgErro = msg,
		 .TipoErro = DadosGenericos.TipoErro.None,
		 .ImagemErro = DadosGenericos.ImagemRetorno.CampoBranco
		}

	End Function


End Class
