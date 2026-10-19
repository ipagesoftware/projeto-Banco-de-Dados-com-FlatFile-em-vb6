VERSION 5.00
Begin VB.Form frmInsert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CADASTRO PRODUTO"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6645
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameBotton 
      Height          =   1065
      Left            =   40
      TabIndex        =   20
      Top             =   4710
      Width           =   6555
      Begin VB.CommandButton btnSalvar 
         Caption         =   "Salvar"
         Height          =   780
         Left            =   1620
         Picture         =   "frmInsert.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1620
      End
      Begin VB.CommandButton btnSair 
         Caption         =   "Sair"
         Height          =   780
         Left            =   3300
         Picture         =   "frmInsert.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1620
      End
   End
   Begin VB.Frame frameCampos 
      Height          =   3465
      Left            =   40
      TabIndex        =   11
      Top             =   1230
      Width           =   6555
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   180
         MaxLength       =   20
         TabIndex        =   0
         Top             =   600
         Width           =   4680
      End
      Begin VB.TextBox txtPrecoCompra 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "0,0000"
         Top             =   1410
         Width           =   2055
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "1,00"
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtTotalPago 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2340
         MaxLength       =   20
         TabIndex        =   4
         Text            =   "0,0000"
         Top             =   2160
         Width           =   2025
      End
      Begin VB.TextBox txtPrecoVenda 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   2
         Text            =   "0,0000"
         Top             =   1410
         Width           =   2025
      End
      Begin VB.TextBox txtPrecoTabela 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "0,0000"
         Top             =   2910
         Width           =   2025
      End
      Begin VB.TextBox txtUltimoPreco 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "0,0000"
         Top             =   2910
         Width           =   2025
      End
      Begin VB.Label lblNomeProduto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOME:"
         Height          =   225
         Left            =   540
         TabIndex        =   18
         ToolTipText     =   "Chave Primária"
         Top             =   330
         Width           =   600
      End
      Begin VB.Label lblPrecoCompra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO COMPRA R$:"
         Height          =   225
         Left            =   210
         TabIndex        =   17
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label lblQuantidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTIDADE:"
         Height          =   225
         Left            =   210
         TabIndex        =   16
         Top             =   1860
         Width           =   1320
      End
      Begin VB.Label lblTotalPago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAGO R$:"
         Height          =   225
         Left            =   2370
         TabIndex        =   15
         Top             =   1860
         Width           =   1920
      End
      Begin VB.Label lblPrecoVenda 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO VENDA R$:"
         Height          =   225
         Left            =   2370
         TabIndex        =   14
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label lblPrecoTabela 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO TABELA R$:"
         Height          =   225
         Left            =   210
         TabIndex        =   13
         Top             =   2610
         Width           =   1920
      End
      Begin VB.Label lblUltimoPreco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÚLTIMO PREÇO R$:"
         Height          =   225
         Left            =   2370
         TabIndex        =   12
         Top             =   2610
         Width           =   1920
      End
      Begin VB.Image imgNome 
         Height          =   240
         Left            =   210
         Picture         =   "frmInsert.frx":1956
         ToolTipText     =   "Chave Primária"
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   6645
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   6645
      Begin VB.Image imgBanner 
         Height          =   480
         Left            =   180
         Picture         =   "frmInsert.frx":1EE0
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "Informe os dados abaixo referentes as informações do cadastro de produtos."
         Height          =   525
         Left            =   960
         TabIndex        =   10
         Top             =   105
         Width           =   5460
      End
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*** MODO ADIÇÃO ***"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   810
      Width           =   6735
   End
   Begin VB.Menu mnuPrinicpal 
      Caption         =   "&Principal"
      Begin VB.Menu mnuSalvar 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'                    ########\ ########\ ########\
'                    ##  _____|\__##  __|##  _____|
'                    ## |         ## |   ## |
'                    #####\       ## |   #####\
'                    ##  __|      ## |   ##  __|
'                    ## |         ## |   ## |
'                    ########\    ## |   ########\
'                    \________|   \__|   \________|
'
'
' ____  ____   ___  __   __     __     ____  ____  ___  __ _  __  ___   __
'(  __)/ ___) / __)/  \ (  )   / _\   (_  _)(  __)/ __)(  ( \(  )/ __) / _\
' ) _) \___ \( (__(  O )/ (_/\/    \    )(   ) _)( (__ /    / )(( (__ /    \
'(____)(____/ \___)\__/ \____/\_/\_/   (__) (____)\___)\_)__)(__)\___)\_/\_/
'               ____  ____  ____  __   ____  _  _   __   __
'              (  __)/ ___)(_  _)/ _\ (    \/ )( \ / _\ (  )
'               ) _) \___ \  )( /    \ ) D () \/ (/    \/ (_/\
'              (____)(____/ (__)\_/\_/(____/\____/\_/\_/\____/
'        ___  __  ___  ____  ____   __     ____  __   __   ____
'       / __)(  )/ __)(  __)(  _ \ /  \   (    \(  ) / _\ / ___)
'      ( (__  )(( (__  ) _)  )   /(  O )   ) D ( )( /    \\___ \
'       \___)(__)\___)(____)(__\_) \__/   (____/(__)\_/\_/(____/
'
' © Copyright Ipage Software - Diógenes Dias - 1995/2026
' Visite meu github: https://github.com/ipagesoftware
'
'
Option Explicit 'FORÇA DECLARAÇÃO DO TIPO DA VARIÁVEL
Private cMask As New clsMaskEdit

Private Sub btnSalvar_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Result      As Long
  Dim Produto     As TProduto
  Dim nomeProduto As String
  '
  nomeProduto = Me.TxtNome
  '
  Screen.MousePointer = vbDefault
  '
  If MsgBox("Confirma os dados digitados?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Atenção") = vbNo Then
    TxtNome.SetFocus
    Exit Sub
  End If
  '
  Screen.MousePointer = vbHourglass
  Result = basProduto.Pesquisar(nomeProduto)
  ' SE O VALOR RETORNADO FOR FALSE NÃO INSERE OS DADOS
  If Result = False Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  ' OK TÁ TUDO BEM HORA DE INSERIR OS DADOS
  'PASSO OS DADOS DAS CAIXAS DE TEXTO A VARIÁVEL GLOBAL DO TIPO TYPE
  Produto.Nome = TxtNome.Text
  Produto.PrecoCompra = TxtPrecoCompra.Text
  Produto.PrecoVenda = TxtPrecoVenda.Text
  Produto.Quantidade = TxtQuantidade.Text
  Produto.TotalPago = TxtTotalPago.Text
  Produto.Tabela = TxtPrecoTabela.Text
  Produto.Ultimo = TxtUltimoPreco.Text
  '
  Result = basProduto.Incluir(Produto)
  Screen.MousePointer = vbDefault
  '
  If Result > 0 Then
    If MsgBox("Operação realizada com sucesso, deseja incluir mais algum registro?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Atenção") = vbNo Then
      Unload Me
      Exit Sub
    End If
  End If
  ' REINICIA O PROCESSO DE GRAVAÇÃO DE DADOS
  TxtNome.Text = ""
  TxtPrecoCompra.Text = "0,0000"
  TxtPrecoVenda.Text = "0,0000"
  TxtQuantidade.Text = "0,00"
  TxtTotalPago.Text = "0,0000"
  TxtPrecoTabela.Text = "0,0000"
  TxtUltimoPreco.Text = "0,0000"
  TxtNome.SetFocus
End Sub

Private Sub btnSair_Click()
  Unload Me
End Sub

Private Sub Form_Initialize()
  Set cMask = New clsMaskEdit
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Set cMask = Nothing
End Sub

Private Sub mnuSair_Click()
  Call btnSair_Click
End Sub

Private Sub mnuSalvar_Click()
  Call btnSalvar_Click
End Sub

Private Sub TxtNome_GotFocus()
  cMask.SelText TxtNome, 0, Len(TxtNome.Text)
End Sub

Private Sub TxtNome_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TxtPrecoCompra.SetFocus
  End If
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
  cMask.Mask TxtNome, KeyAscii, "U"
End Sub

Private Sub TxtNome_LostFocus()
  If Len(TxtNome.Text) = 0 Then
    TxtNome.Text = "INDEFINIDO"
  End If
End Sub

Private Sub TxtPrecoCompra_GotFocus()
  cMask.SelText TxtPrecoCompra, 0, Len(TxtPrecoCompra.Text)
End Sub

Private Sub TxtPrecoCompra_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TxtPrecoVenda.SetFocus
  End If
End Sub

Private Sub TxtPrecoCompra_KeyPress(KeyAscii As Integer)
  cMask.Mask TxtPrecoCompra, KeyAscii, "C", "0,0000"
End Sub

Private Sub TxtPrecoCompra_LostFocus()
  If Len(TxtPrecoCompra.Text) = 0 Then
    TxtPrecoCompra.Text = 0
  End If
  '
  TxtPrecoCompra.Text = Format(TxtPrecoCompra.Text, "#,##0.0000")
End Sub

Private Sub TxtPrecoVenda_GotFocus()
  cMask.SelText TxtPrecoVenda, 0, Len(TxtPrecoVenda.Text)
End Sub

Private Sub TxtPrecoVenda_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TxtQuantidade.SetFocus
  End If
End Sub

Private Sub TxtPrecoVenda_KeyPress(KeyAscii As Integer)
  cMask.Mask TxtPrecoVenda, KeyAscii, "C"
End Sub

Private Sub TxtPrecoVenda_LostFocus()
  If Len(TxtPrecoVenda.Text) = 0 Then
    TxtPrecoVenda.Text = 0
  End If
  '
  TxtPrecoVenda.Text = Format(TxtPrecoVenda.Text, "#,##0.0000")
  '
  If CCur(TxtUltimoPreco.Text) = 0 Then
    TxtUltimoPreco.Text = TxtPrecoCompra.Text
  End If
End Sub

Private Sub TxtQuantidade_GotFocus()
  cMask.SelText TxtQuantidade, 0, Len(TxtQuantidade.Text)
End Sub

Private Sub TxtQuantidade_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TxtTotalPago.SetFocus
  End If
End Sub

Private Sub TxtQuantidade_KeyPress(KeyAscii As Integer)
  cMask.Mask TxtQuantidade, KeyAscii, "C"
End Sub

Private Sub TxtQuantidade_LostFocus()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim TotalPagar As Currency
  '
  If Len(TxtQuantidade.Text) = 0 Then
    TxtQuantidade.Text = 1
  End If
  '
  TxtQuantidade.Text = Format(TxtQuantidade.Text, "#,##0.00")
  TotalPagar = CCur(TxtPrecoCompra.Text) * CCur(TxtQuantidade.Text)
  TxtTotalPago.Text = Format(TotalPagar, "#,##0.0000")
End Sub

Private Sub TxtTotalPago_GotFocus()
  cMask.SelText TxtTotalPago, 0, Len(TxtTotalPago.Text)
End Sub

Private Sub TxtTotalPago_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TxtPrecoTabela.SetFocus
  End If
End Sub

Private Sub TxtTotalPago_KeyPress(KeyAscii As Integer)
  cMask.Mask TxtTotalPago, KeyAscii, "C"
End Sub

Private Sub TxtTotalPago_LostFocus()
  If Len(TxtTotalPago.Text) = 0 Then
    TxtTotalPago.Text = 0
  End If
  '
  TxtTotalPago.Text = Format(TxtTotalPago.Text, "#,##0.0000")
End Sub

Private Sub TxtPrecoTabela_GotFocus()
  cMask.SelText TxtPrecoTabela, 0, Len(TxtPrecoTabela.Text)
End Sub

Private Sub TxtPrecoTabela_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TxtUltimoPreco.SetFocus
  End If
End Sub

Private Sub TxtPrecoTabela_KeyPress(KeyAscii As Integer)
  cMask.Mask TxtPrecoTabela, KeyAscii, "C"
End Sub

Private Sub TxtPrecoTabela_LostFocus()
  If Len(TxtPrecoTabela.Text) = 0 Then
    TxtPrecoTabela.Text = 0
  End If
  '
  TxtPrecoTabela.Text = Format(TxtPrecoTabela.Text, "#,##0.0000")
End Sub

Private Sub TxtUltimoPreco_GotFocus()
  cMask.SelText TxtUltimoPreco, 0, Len(TxtUltimoPreco.Text)
End Sub

Private Sub TxtUltimoPreco_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    btnSalvar.SetFocus
  End If
End Sub

Private Sub TxtUltimoPreco_KeyPress(KeyAscii As Integer)
  cMask.Mask TxtUltimoPreco, KeyAscii, "C"
End Sub

Private Sub TxtUltimoPreco_LostFocus()
  If Len(TxtUltimoPreco.Text) = 0 Then
    TxtUltimoPreco.Text = 0
  End If
  '
  TxtUltimoPreco.Text = Format(TxtUltimoPreco.Text, "#,##0.0000")
End Sub
