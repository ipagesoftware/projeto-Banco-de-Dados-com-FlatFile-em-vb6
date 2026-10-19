VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exemplo Função"
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
   Begin VB.Frame FrameCampos 
      Height          =   2325
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5115
      Begin VB.TextBox txtSenha 
         Height          =   375
         Left            =   180
         MaxLength       =   20
         TabIndex        =   0
         Top             =   660
         Width           =   4680
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SENHA:"
         Height          =   225
         Left            =   540
         TabIndex        =   5
         ToolTipText     =   "Chave Primária"
         Top             =   360
         Width           =   720
      End
      Begin VB.Image imgSenha 
         Height          =   240
         Left            =   210
         Picture         =   "Form1.frx":0000
         ToolTipText     =   "A senha e a palavra: ipage"
         Top             =   330
         Width           =   240
      End
   End
   Begin VB.Frame FrameRodape 
      Height          =   915
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   5145
      Begin VB.CommandButton btnCalcular 
         Caption         =   "Calcular"
         Height          =   465
         Left            =   2640
         TabIndex        =   3
         Top             =   270
         Width           =   1350
      End
      Begin VB.CommandButton btnLogin 
         Caption         =   "Login"
         Height          =   465
         Left            =   1140
         TabIndex        =   1
         Top             =   270
         Width           =   1350
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCalcular_Click()
  Dim qtd            As Long
  Dim preco          As Currency
  Dim resultadoFinal As Currency
    
  qtd = 5
  preco = 45.5
    
  ' Chamando a função e guardando o valor retornado na variável
  resultadoFinal = CalcularTotalVenda(qtd, preco)
    
  MsgBox "O valor total da compra é: R$ " & Format(resultadoFinal, "Standard"), vbInformation, "Cálculo Realizado"

End Sub

Private Sub btnLogin_Click()
  If ValidarSenhaAdmin(txtSenha.Text) = True Then
    MsgBox "Acesso autorizado!", vbInformation
  Else
    MsgBox "Senha inválida!", vbCritical
  End If
End Sub

