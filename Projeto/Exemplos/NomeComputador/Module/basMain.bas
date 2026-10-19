Attribute VB_Name = "basMain"
Option Explicit

' 1. Declaração da API do Windows no topo de um Módulo (.bas)
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
   
Public Sub Main()
  ExibirNomeDoComputador
End Sub

' 2. Uso prático da API dentro de uma rotina do sistema
Public Sub ExibirNomeDoComputador()
  Dim bufferNome    As String
  Dim tamanhoBuffer As Long
  Dim resultado     As Long
  ' Prepara o espaço de memória para receber o texto do Windows
  bufferNome = Space$(255)
  tamanhoBuffer = 255
  ' Chama a API nativa do Windows
  resultado = GetComputerName(bufferNome, tamanhoBuffer)
    
  If resultado <> 0 Then
    ' Remove os caracteres vazios extras gerados pelo buffer
    bufferNome = Left$(bufferNome, tamanhoBuffer)
    MsgBox "O nome desta máquina na rede é: " & bufferNome, vbInformation, "API do Windows"
  Else
    MsgBox "Não foi possível recuperar o nome do computador.", vbCritical
  End If
End Sub

