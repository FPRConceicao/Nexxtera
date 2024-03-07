Imports Teleatlantic.TLS.Common
Imports System.Reflection

Public Class ContaCorrente : Inherits Retorno

	Public Property Descricao() As String
		Get
			Return m_Descricao
		End Get
		Set(ByVal value As String)
			m_Descricao = value
		End Set
	End Property
	Private m_Descricao As String

	Public Property IsDebAuto() As String
		Get
			Return m_IsDebAuto
		End Get
		Set(ByVal value As String)
			m_IsDebAuto = value
		End Set
	End Property
	Private m_IsDebAuto As String


	Public Property CodBanco() As String
		Get
			Return m_CodBanco
		End Get
		Set(ByVal value As String)
			m_CodBanco = value
		End Set
	End Property
	Private m_CodBanco As String


	Public Property CodAgen() As String
		Get
			Return m_CodAgen
		End Get
		Set(ByVal value As String)
			m_CodAgen = value
		End Set
	End Property
	Private m_CodAgen As String


	Public Property NumCta() As String
		Get
			Return m_NumCta
		End Get
		Set(ByVal value As String)
			m_NumCta = value
		End Set
	End Property
	Private m_NumCta As String


	Public Property SeqRemes() As Integer
		Get
			Return m_SeqRemes
		End Get
		Set(ByVal value As Integer)
			m_SeqRemes = value
		End Set
	End Property
	Private m_SeqRemes As Integer


	Public Property Convenio() As String
		Get
			Return m_Convenio
		End Get
		Set(ByVal value As String)
			m_Convenio = value
		End Set
	End Property
	Private m_Convenio As String


	Public Property CodCarteira() As String
		Get
			Return m_CodCarteira
		End Get
		Set(ByVal value As String)
			m_CodCarteira = value
		End Set
	End Property
	Private m_CodCarteira As String


	Public Property NossoNumero() As String
		Get
			Return m_NossoNumero
		End Get
		Set(ByVal value As String)
			m_NossoNumero = value
		End Set
	End Property
	Private m_NossoNumero As String


	Public Property NumCtaVincDesc() As String
		Get
			Return m_NumCtaVincDesc
		End Get
		Set(ByVal value As String)
			m_NumCtaVincDesc = value
		End Set
	End Property
	Private m_NumCtaVincDesc As String
	Public Property NomeBanco As String
		Get
			Return m_NomeBanco
		End Get
		Set(ByVal value As String)
			m_NomeBanco = value
		End Set
	End Property
	Private m_NomeBanco As String
	Public Property NomeAgen As String
		Get
			Return m_NomeAgen
		End Get
		Set(ByVal value As String)
			m_NomeAgen = value
		End Set
	End Property
	Private m_NomeAgen As String
	Public Property ContaFinan As String
		Get
			Return m_ContaFinan
		End Get
		Set(ByVal value As String)
			m_ContaFinan = value
		End Set
	End Property
	Private m_ContaFinan As String
	Public Property LimiteCredito As Double
		Get
			Return m_LimiteCredito
		End Get
		Set(ByVal value As Double)
			m_LimiteCredito = value
		End Set
	End Property
	Private m_LimiteCredito As Double
	Public Property Fluxo As String
		Get
			Return m_Fluxo
		End Get
		Set(ByVal value As String)
			m_Fluxo = value
		End Set
	End Property
	Private m_Fluxo As String
	Public Property Tipo As String
		Get
			Return m_Tipo
		End Get
		Set(ByVal value As String)
			m_Tipo = value
		End Set
	End Property
	Private m_Tipo As String
	Public Property LayoutCheque As String
		Get
			Return m_LayoutCheque
		End Get
		Set(ByVal value As String)
			m_LayoutCheque = value
		End Set
	End Property
	Private m_LayoutCheque As String
	Public Property TextoBordero As String
		Get
			Return m_TextoBordero
		End Get
		Set(ByVal value As String)
			m_TextoBordero = value
		End Set
	End Property
	Private m_TextoBordero As String
	Public Property CodCCusto As String
		Get
			Return m_CodCCusto
		End Get
		Set(ByVal value As String)
			m_CodCCusto = value
		End Set
	End Property
	Private m_CodCCusto As String
	Public Property ContaCtbl As String
		Get
			Return m_ContaCtbl
		End Get
		Set(ByVal value As String)
			m_ContaCtbl = value
		End Set
	End Property
	Private m_ContaCtbl As String
	Public Property FloatCredito As String
		Get
			Return m_FloatCredito
		End Get
		Set(ByVal value As String)
			m_FloatCredito = value
		End Set
	End Property
	Private m_FloatCredito As String
	Public Property TipoConta As String
		Get
			Return m_TipoConta
		End Get
		Set(ByVal value As String)
			m_TipoConta = value
		End Set
	End Property
	Private m_TipoConta As String
	Public Property NumCtaVinc As String
		Get
			Return m_NumCtaVinc
		End Get
		Set(ByVal value As String)
			m_NumCtaVinc = value
		End Set
	End Property
	Private m_NumCtaVinc As String
	Public Property Status As String
		Get
			Return m_Status
		End Get
		Set(ByVal value As String)
			m_Status = value
		End Set
	End Property
	Private m_Status As String
	Public Property PermiteDI As String
		Get
			Return m_PermiteDI
		End Get
		Set(ByVal value As String)
			m_PermiteDI = value
		End Set
	End Property
	Private m_PermiteDI As String
	Public Property PagamentoEletronico As String
		Get
			Return m_PagamentoEletronico
		End Get
		Set(ByVal value As String)
			m_PagamentoEletronico = value
		End Set
	End Property
	Private m_PagamentoEletronico As String
	Public Property BancoAgen As String
		Get
			Return m_BancoAgen
		End Get
		Set(ByVal value As String)
			m_BancoAgen = value
		End Set
	End Property
	Private m_BancoAgen As String
	Public Property SaldoTotal As Double
		Get
			Return m_SaldoTotal
		End Get
		Set(ByVal value As Double)
			m_SaldoTotal = value
		End Set
	End Property
	Private m_SaldoTotal As Double
	Public Property SaldoLiq As Double
		Get
			Return m_SaldoLiq
		End Get
		Set(ByVal value As Double)
			m_SaldoLiq = value
		End Set
	End Property
	Private m_SaldoLiq As Double
	Public Property SaldoRes As Double
		Get
			Return m_SaldoRes
		End Get
		Set(ByVal value As Double)
			m_SaldoRes = value
		End Set
	End Property
	Private m_SaldoRes As Double
	Public Property LayoutBordero As String
		Get
			Return m_LayoutBordero
		End Get
		Set(ByVal value As String)
			m_LayoutBordero = value
		End Set
	End Property
	Private m_LayoutBordero As String

	Public Shared Function ComparaContas(_CC1 As ContaCorrente, _CC2 As ContaCorrente) As String()

		Dim oldValue As String = ""
		Dim newValue As String = ""

		For Each prop As PropertyInfo In GetType(ContaCorrente).GetProperties()

			Dim v1 As Object = prop.GetValue(_CC1, Nothing)
			Dim v2 As Object = prop.GetValue(_CC2, Nothing)

			If (prop.Name = "CodAgen" Or prop.Name = "ContaCtbl") Then
				v1 = v1.ToString().Replace("-", "")
			ElseIf (prop.Name = "IsDebAuto") Then
				v1 = IIf(v1.ToString() = "Sim", "1", "0")
			ElseIf ({"Fluxo", "Tipo", "TipoConta", "Status", "PermiteDI", "PagamentoEletronico"}.Contains(prop.Name)) Then
				v1 = v1.ToString().Substring(0, 1)
			End If

			If (prop.Name <> "Sucesso" Or prop.Name <> "NomeBanco" Or prop.Name <> "NomeAgen") Then
				If (v1 Is Nothing) Then
					If (v2 IsNot Nothing) Then
						oldValue += prop.Name + ": NULL;"
						newValue += prop.Name + ": " + v2.ToString()
					End If
				ElseIf (v2 Is Nothing) Then
					If (v1 IsNot Nothing) Then
						oldValue += prop.Name + ": " + v1.ToString() + ";"
						newValue += prop.Name + ": NULL;"
					End If
				ElseIf (Not v1.Equals(v2)) Then
					oldValue += prop.Name + ":" + v1.ToString() + ";"
					newValue += prop.Name + ":" + v2.ToString() + ";"
				End If
			End If
		Next

		Return {oldValue, newValue}
    End Function

    Public Property CodCCustoDNI As String

    Public Property CodTransmissao() As String
        Get
            Return m_CodTransmissao
        End Get
        Set(ByVal value As String)
            m_CodTransmissao = value
        End Set
    End Property
    Private m_CodTransmissao As String

    Public Property CodFlash As String
    Public Property ContaPadraoDebAuto As Integer
    Public Property ContaPadraoDebAutoDesc As String
    Public Property IsGerarNossoNumero As Integer
End Class
