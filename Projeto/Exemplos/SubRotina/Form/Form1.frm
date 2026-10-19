VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exemplo Sub-Rotina"
   ClientHeight    =   4590
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
   ScaleHeight     =   4590
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameCampos 
      Height          =   3435
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5115
      Begin VB.TextBox txtCodigo 
         Height          =   375
         Left            =   180
         MaxLength       =   5
         TabIndex        =   0
         Top             =   630
         Width           =   1425
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   180
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1470
         Width           =   4680
      End
      Begin VB.TextBox txtPrecoCompra 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "0,00"
         Top             =   2310
         Width           =   2055
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓDIGO:"
         Height          =   225
         Left            =   500
         TabIndex        =   9
         ToolTipText     =   "Chave Primária"
         Top             =   330
         Width           =   840
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOME:"
         Height          =   225
         Left            =   540
         TabIndex        =   8
         ToolTipText     =   "Chave Primária"
         Top             =   1170
         Width           =   600
      End
      Begin VB.Label lblPrecoCompra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO COMPRA R$:"
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   1980
         Width           =   1920
      End
      Begin VB.Image imgNome 
         Height          =   240
         Left            =   210
         Picture         =   "Form1.frx":0000
         ToolTipText     =   "Chave Primária"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Left            =   180
         Picture         =   "Form1.frx":058A
         ToolTipText     =   "Chave Primária"
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Frame FrameRodape 
      Height          =   915
      Left            =   0
      TabIndex        =   5
      Top             =   3630
      Width           =   5145
      Begin VB.CommandButton btnNovo 
         Caption         =   "Novo"
         Height          =   465
         Left            =   1050
         TabIndex        =   3
         Top             =   240
         Width           =   1350
      End
      Begin VB.CommandButton cmdAlerta 
         Caption         =   "Alerta"
         Height          =   465
         Left            =   2550
         TabIndex        =   4
         Top             =   240
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

Private Sub btnNovo_Click()
  ' Chama a sub-rotina que criamos acima
  Call LimparCamposFormulario
  MsgBox "Formulário limpo e pronto para novo cadastro!", vbInformation
End Sub

Private Sub cmdAlerta_Click()
  ExibirAlerta "O estoque deste produto está crítico!", "Atenção de Estoque"
End Sub

' Declaração de uma Sub-rotina privada que limpa os campos de cadastro
Private Sub LimparCamposFormulario()
  txtCodigo.Text = ""
  txtNome.Text = ""
  txtPrecoCompra.Text = "0,00"
  txtCodigo.SetFocus ' Retorna o cursor para a caixa de texto do nome
End Sub


