Attribute VB_Name = "BasMain"
Option Explicit

' Sub-rotina que recebe uma mensagem personalizada e exibe na tela
Public Sub ExibirAlerta(ByVal mensagem As String, ByVal tipoTitulo As String)
  MsgBox mensagem, vbExclamation, tipoTitulo
End Sub

