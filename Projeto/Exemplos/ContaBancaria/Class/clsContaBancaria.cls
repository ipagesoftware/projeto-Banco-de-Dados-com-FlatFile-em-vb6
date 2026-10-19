VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =======================================================
' 1. VARIÁVEIS PRIVADAS (Estado Interno)
' =======================================================
Private m_Titular As String
Private m_Saldo As Double
' =======================================================
' 2. CICLO DE VIDA DO OBJETO
' =======================================================
' Executado automaticamente quando o objeto é instanciado (Set ... = New)
Private Sub Class_Initialize()
    m_Titular = "Não informado"
    m_Saldo = 0
End Sub
' Executado quando o objeto é destruído da memória (Set ... = Nothing)
Private Sub Class_Terminate()
    ' Limpeza de memória ou registro de logs, se necessário
End Sub


' =======================================================
' 3. PROPRIEDADES (Encapsulamento com Validação)
' =======================================================
' Leitura do Titular
Public Property Get Titular() As String
    Titular = m_Titular
End Property

' Escrita do Titular
Public Property Let Titular(ByVal Value As String)
    m_Titular = Trim(Value)
End Property

' Leitura do Saldo (Apenas leitura externamente, sem Property Let para o Saldo!)
Public Property Get Saldo() As Double
    Saldo = m_Saldo
End Property

' =======================================================
' 4. MÉTODOS (Comportamentos / Ações)
' =======================================================
' Método para depositar dinheiro
Public Sub Depositar(ByVal Valor As Double)
  If Valor > 0 Then
    m_Saldo = m_Saldo + Valor
  Else
    MsgBox "O valor do depósito deve ser positivo!", vbExclamation, "Atenção"
  End If
End Sub

' Método para sacar dinheiro
Public Function Sacar(ByVal Valor As Double) As Boolean
    If Valor > 0 And Valor <= m_Saldo Then
        m_Saldo = m_Saldo - Valor
        Sacar = True ' Saque realizado com sucesso
    Else
        MsgBox "Saldo insuficiente ou valor inválido!", vbCritical, "Erro"
        Sacar = False ' Falha no saque
    End If
End Function

