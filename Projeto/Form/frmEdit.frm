VERSION 5.00
Begin VB.Form frmEdit 
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
   Icon            =   "frmEdit.frx":0000
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
      TabIndex        =   21
      Top             =   4650
      Width           =   6555
      Begin VB.CommandButton btnSair 
         Caption         =   "Sair"
         Height          =   780
         Left            =   3300
         Picture         =   "frmEdit.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   1620
      End
      Begin VB.CommandButton btnSalvar 
         Caption         =   "Salvar"
         Height          =   780
         Left            =   1620
         Picture         =   "frmEdit.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1620
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   6645
      Begin VB.Image imgBanner 
         Height          =   480
         Left            =   180
         Picture         =   "frmEdit.frx":1956
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "Edite os dados abaixo referentes ao cadastro de produtos."
         Height          =   465
         Left            =   960
         TabIndex        =   20
         Top             =   105
         Width           =   5400
      End
   End
   Begin VB.Frame frameCampos 
      Height          =   3435
      Left            =   40
      TabIndex        =   9
      Top             =   1170
      Width           =   6555
      Begin VB.TextBox txtUltimoPreco 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "0,0000"
         Text            =   "0,0000"
         Top             =   2940
         Width           =   2025
      End
      Begin VB.TextBox txtPrecoTabela 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "0,0000"
         Text            =   "0,0000"
         Top             =   2940
         Width           =   2025
      End
      Begin VB.TextBox txtPrecoVenda 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "0,0000"
         Text            =   "0,0000"
         Top             =   1440
         Width           =   2025
      End
      Begin VB.TextBox txtTotalPago 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2340
         MaxLength       =   20
         TabIndex        =   5
         Tag             =   "0,0000"
         Text            =   "0,0000"
         Top             =   2190
         Width           =   2025
      End
      Begin VB.TextBox txtQuantidade 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "1,00"
         Text            =   "1,00"
         Top             =   2190
         Width           =   2055
      End
      Begin VB.TextBox txtPrecoCompra 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   180
         MaxLength       =   15
         TabIndex        =   2
         Tag             =   "0,0000"
         Text            =   "0,0000"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   1710
         MaxLength       =   20
         TabIndex        =   1
         Top             =   630
         Width           =   4680
      End
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   180
         MaxLength       =   5
         TabIndex        =   0
         Tag             =   "00000"
         Top             =   630
         Width           =   1425
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Left            =   180
         Picture         =   "frmEdit.frx":2220
         ToolTipText     =   "Chave Primária"
         Top             =   300
         Width           =   240
      End
      Begin VB.Image imgNome 
         Height          =   240
         Left            =   1740
         Picture         =   "frmEdit.frx":27AA
         ToolTipText     =   "Chave Primária"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label lblUltimoPreco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÚLTIMO PREÇO R$:"
         Height          =   225
         Left            =   2370
         TabIndex        =   18
         Top             =   2640
         Width           =   1920
      End
      Begin VB.Label lblPrecoTabela 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO TABELA R$:"
         Height          =   225
         Left            =   210
         TabIndex        =   17
         Top             =   2640
         Width           =   1920
      End
      Begin VB.Label lblPrecoVenda 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO VENDA R$:"
         Height          =   225
         Left            =   2370
         TabIndex        =   16
         Top             =   1110
         Width           =   1920
      End
      Begin VB.Label lblTotalPago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAGO R$:"
         Height          =   225
         Left            =   2370
         TabIndex        =   15
         Top             =   1890
         Width           =   1920
      End
      Begin VB.Label lblQuantidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTIDADE:"
         Height          =   225
         Left            =   210
         TabIndex        =   14
         Top             =   1890
         Width           =   1320
      End
      Begin VB.Label lblPrecoCompra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO COMPRA R$:"
         Height          =   225
         Left            =   210
         TabIndex        =   13
         Top             =   1110
         Width           =   1920
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOME:"
         Height          =   225
         Left            =   2070
         TabIndex        =   12
         ToolTipText     =   "Chave Primária"
         Top             =   330
         Width           =   600
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓDIGO:"
         Height          =   225
         Left            =   500
         TabIndex        =   11
         ToolTipText     =   "Chave Primária"
         Top             =   330
         Width           =   840
      End
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "*** MODO EDIÇÃO ***"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   810
      Width           =   6735
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Principal"
      Begin VB.Menu mnuSalvar 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmEdit"
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
Private mDados     As String
Private mRecordSet As Long

Private Sub btnSalvar_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Result  As Long
  Dim Produto As TProduto
  '
  If MsgBox("Confirma os dados digitados?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Atenção") = vbNo Then
    txtNome.SetFocus
    Exit Sub
  End If
  '-------------------------------------------------------------------
  'PASSA OS DADOS DAS CAIXAS DE TEXTO A VARIÁVEL GLOBAL DO TIPO TYPE
  '-------------------------------------------------------------------
  Produto.Nome = txtNome.Text
  Produto.Codigo = txtCodigo.Text
  Produto.PrecoCompra = txtPrecoCompra.Text
  Produto.PrecoVenda = txtPrecoVenda.Text
  Produto.Quantidade = txtQuantidade.Text
  Produto.TotalPago = txtTotalPago.Text
  Produto.Tabela = txtPrecoTabela.Text
  Produto.Ultimo = txtUltimoPreco.Text
  '
  Result = basProduto.Duplicidade(Produto.Codigo, Produto.Nome)
  '
  If Result = True Then
    Result = basProduto.Editar(RecordSet, Produto)
    Screen.MousePointer = vbDefault
    '
    If Result = True Then
      MsgBox "Operação realizada com sucesso!", vbInformation Or vbOKOnly, "Atenção"
      Unload Me
      Exit Sub
    Else
      MsgBox "Ocorreu uma exceção ao editar os dados atuais, por gentiliza verifique!", vbCritical, "Atenção"
      txtNome.SetFocus
    End If
  End If
End Sub

Private Sub btnSair_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  If Len(Trim$(Dados)) Then
    '---------------------
    ' DEFINIÇÃO VARIÁVEIS
    '---------------------
    Dim Valores() As String
    '
    Valores = Split(Dados, "|")
    '
    txtNome.Text = Valores(0)
    txtCodigo.Text = Valores(1)
    txtPrecoCompra.Text = Valores(2)
    txtPrecoVenda.Text = Valores(3)
    txtQuantidade.Text = Valores(4)
    txtPrecoTabela.Text = Valores(5)
    txtTotalPago.Text = Valores(6)
    txtUltimoPreco.Text = Valores(7)
  End If
  '
  Set MaskedEdit = New clsMaskedEdit
  MaskedEdit.Format txtQuantidade, "MOEDA", 2
  MaskedEdit.Format txtPrecoCompra, "MOEDA", 4
  MaskedEdit.Format txtPrecoVenda, "MOEDA", 4
  MaskedEdit.Format txtPrecoTabela, "MOEDA", 4
  MaskedEdit.Format txtUltimoPreco, "MOEDA", 4
  MaskedEdit.Format txtTotalPago, "MOEDA", 4
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  MaskedEdit.ReleaseForm Me
End Sub

Private Sub mnuSair_Click()
  Call btnSair_Click
End Sub

Private Sub mnuSalvar_Click()
  Call btnSalvar_Click
End Sub

Private Sub TxtCodigo_GotFocus()
  MaskedEdit.SelText txtCodigo, 0, Len(txtCodigo.Text)
End Sub

Private Sub TxtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtNome.SetFocus
  End If
End Sub

Private Sub TxtCodigo_LostFocus()
  If Len(txtCodigo.Text) = 0 Then
    txtCodigo.Text = txtCodigo.Tag
  End If
  '
  txtCodigo.Text = Format$(txtCodigo.Text, "00000")
End Sub

Private Sub TxtNome_GotFocus()
  MaskedEdit.SelText txtNome, 0, Len(txtNome.Text)
End Sub

Private Sub TxtNome_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtPrecoCompra.SetFocus
  End If
End Sub

Private Sub TxtNome_LostFocus()
  If Len(txtNome.Text) = 0 Then
    txtNome.Text = "INDEFINIDO"
  Else
    txtNome.Text = UCase$(txtNome.Text)
  End If
End Sub

Private Sub TxtPrecoCompra_GotFocus()
  MaskedEdit.SelText txtPrecoCompra, 0, Len(txtPrecoCompra.Text)
End Sub

Private Sub TxtPrecoCompra_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtPrecoVenda.SetFocus
  End If
End Sub

Private Sub TxtPrecoCompra_LostFocus()
  If Len(txtPrecoCompra.Text) = 0 Then
    txtPrecoCompra.Text = txtPrecoCompra.Tag
  End If
  '
  txtPrecoCompra.Text = Format$(txtPrecoCompra.Text, "#,##0.0000")
End Sub

Private Sub TxtPrecoTabela_GotFocus()
  MaskedEdit.SelText txtPrecoTabela, 0, Len(txtPrecoTabela.Text)
End Sub

Private Sub TxtPrecoTabela_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtUltimoPreco.SetFocus
  End If
End Sub

Private Sub TxtPrecoTabela_LostFocus()
  If Len(txtPrecoTabela.Text) = 0 Then
    txtPrecoTabela.Text = txtPrecoTabela.Tag
  End If
  '
  txtPrecoTabela.Text = Format$(txtPrecoTabela.Text, "#,##0.0000")
End Sub

Private Sub TxtPrecoVenda_GotFocus()
  MaskedEdit.SelText txtPrecoVenda, 0, Len(txtPrecoVenda.Text)
End Sub

Private Sub TxtPrecoVenda_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtQuantidade.SetFocus
  End If
End Sub

Private Sub TxtPrecoVenda_LostFocus()
  If Len(txtPrecoVenda.Text) = 0 Then
    txtPrecoVenda.Text = txtPrecoVenda.Tag
  End If
  '
  txtPrecoVenda.Text = Format$(txtPrecoVenda.Text, "#,##0.0000")
  '
  If CCur(txtUltimoPreco.Text) = 0 Then
    txtUltimoPreco.Text = txtPrecoCompra.Text
  End If
End Sub

Private Sub TxtQuantidade_GotFocus()
  MaskedEdit.SelText txtQuantidade, 0, Len(txtQuantidade.Text)
End Sub

Private Sub TxtQuantidade_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtTotalPago.SetFocus
  End If
End Sub

Private Sub TxtQuantidade_LostFocus()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim TotalPagar As Currency
  '
  If Len(txtQuantidade.Text) = 0 Then
    txtQuantidade.Text = txtQuantidade.Tag
  End If
  '
  txtQuantidade.Text = Format$(txtQuantidade.Text, "#,##0.00")
  TotalPagar = CCur(txtPrecoCompra.Text) * CCur(txtQuantidade.Text)
  txtTotalPago.Text = Format$(TotalPagar, "#,##0.0000")
End Sub

Private Sub TxtTotalPago_GotFocus()
  MaskedEdit.SelText txtTotalPago, 0, Len(txtTotalPago.Text)
End Sub

Private Sub TxtTotalPago_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtPrecoTabela.SetFocus
  End If
End Sub

Private Sub TxtTotalPago_LostFocus()
  If Len(txtTotalPago.Text) = 0 Then
    txtTotalPago.Text = txtTotalPago.Tag
  End If
  '
  txtTotalPago.Text = Format$(txtTotalPago.Text, "#,##0.0000")
End Sub

Private Sub TxtUltimoPreco_GotFocus()
  MaskedEdit.SelText txtUltimoPreco, 0, Len(txtUltimoPreco.Text)
End Sub

Private Sub TxtUltimoPreco_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    btnSalvar.SetFocus
  End If
End Sub

Private Sub TxtUltimoPreco_LostFocus()
  If Len(txtUltimoPreco.Text) = 0 Then
    txtUltimoPreco.Text = txtUltimoPreco.Tag
  End If
  '
  txtUltimoPreco.Text = Format$(txtUltimoPreco.Text, "#,##0.0000")
End Sub

Public Property Get RecordSet() As String
  RecordSet = mRecordSet
End Property

Public Property Let RecordSet(ByVal vNewValue As String)
  mRecordSet = vNewValue
End Property

Public Property Get Dados() As String
  Dados = mDados
End Property

Public Property Let Dados(ByVal vNewValue As String)
  mDados = vNewValue
End Property

