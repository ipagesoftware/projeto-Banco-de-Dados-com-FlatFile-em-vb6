VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exemplo Class"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTestarConta 
      Caption         =   "Testar Conta"
      Height          =   465
      Left            =   1050
      TabIndex        =   0
      Top             =   1770
      Width           =   2820
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnTestarConta_Click()
    ' 1. Declaração da variável do tipo da Classe
    Dim minhaConta As clsContaBancaria
    
    ' 2. Instanciação (Criação do objeto na memória -> dispara Class_Initialize)
    Set minhaConta = New clsContaBancaria
    
    ' 3. Uso das Propriedades e Métodos
    minhaConta.Titular = "Guilherme Santos"
    minhaConta.Depositar 500#
    
    MsgBox "Titular: " & minhaConta.Titular & vbCrLf & _
           "Saldo Inicial: R$ " & minhaConta.Saldo, vbInformation, "Conta Criada"
           
    ' Tentando realizar um saque
    If minhaConta.Sacar(150#) Then
        MsgBox "Saque de R$ 150,00 realizado! Saldo atual: R$ " & minhaConta.Saldo
    End If
    
    ' 4. Destruição do objeto (Dispara Class_Terminate e libera a RAM)
    Set minhaConta = Nothing
End Sub


