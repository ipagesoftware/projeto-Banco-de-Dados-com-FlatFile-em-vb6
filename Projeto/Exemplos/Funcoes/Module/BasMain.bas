Attribute VB_Name = "BasMain"
Option Explicit

' Função que calcula o valor total de uma venda e retorna o valor (As Currency)
Public Function CalcularTotalVenda(ByVal Quantidade As Long, _
                                   ByVal PrecoUnitario As Currency) As Currency
  Dim valorTotal As Currency
    
  ' Realiza o cálculo matemático
  valorTotal = Quantidade * PrecoUnitario
    
  ' O valor é retornado atribuindo-o ao nome da função
  CalcularTotalVenda = valorTotal
End Function

Public Function ValidarSenhaAdmin(ByVal SenhaDigitada As String) As Boolean
  If SenhaDigitada = "ipage" Then
    ValidarSenhaAdmin = True   ' Senha correta
  Else
    ValidarSenhaAdmin = False  ' Senha incorreta
  End If
End Function

