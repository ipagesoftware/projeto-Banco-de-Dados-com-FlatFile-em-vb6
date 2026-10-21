VERSION 5.00
Begin VB.Form frmProduto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CADASTRO PRODUTOS"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProduto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameDados 
      Caption         =   "DADOS"
      Height          =   4455
      Left            =   4980
      TabIndex        =   26
      Top             =   1620
      Width           =   4890
      Begin VB.TextBox TxtNome 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1440
         Width           =   4680
      End
      Begin VB.TextBox TxtPrecoVenda 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "0,0000"
         Top             =   2220
         Width           =   2025
      End
      Begin VB.TextBox TxtPrecoCompra 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "0,0000"
         Top             =   2220
         Width           =   2055
      End
      Begin VB.TextBox TxtTotalPago 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   20
         TabIndex        =   17
         Tag             =   "0,0000"
         Top             =   3000
         Width           =   2025
      End
      Begin VB.TextBox TxtQuantidade 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "1,00"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox TxtUltimoPreco 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "0,0000"
         Top             =   3780
         Width           =   2025
      End
      Begin VB.TextBox TxtPrecoTabela 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "0,0000"
         Top             =   3780
         Width           =   2055
      End
      Begin VB.TextBox TxtCodigo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaxLength       =   5
         TabIndex        =   11
         Top             =   630
         Width           =   1425
      End
      Begin VB.Image imgNome 
         Height          =   240
         Left            =   150
         Picture         =   "frmProduto.frx":0E42
         ToolTipText     =   "Chave Primária"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOME:"
         Height          =   225
         Left            =   480
         TabIndex        =   34
         ToolTipText     =   "Chave Primária"
         Top             =   1170
         Width           =   600
      End
      Begin VB.Label lblPrecoVenda 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO VENDA R$:"
         Height          =   225
         Left            =   2310
         TabIndex        =   33
         Top             =   1890
         Width           =   1920
      End
      Begin VB.Label lblPrecoCompra 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO COMPRA R$:"
         Height          =   225
         Left            =   150
         TabIndex        =   32
         Top             =   1890
         Width           =   1920
      End
      Begin VB.Label lblTotalPago 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAGO R$:"
         Height          =   225
         Left            =   2310
         TabIndex        =   31
         Top             =   2700
         Width           =   1920
      End
      Begin VB.Label lblQuantidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTIDADE:"
         Height          =   225
         Left            =   150
         TabIndex        =   30
         Top             =   2700
         Width           =   1320
      End
      Begin VB.Label lblUltimoPreco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÚLTIMO PREÇO R$"
         Height          =   225
         Left            =   2310
         TabIndex        =   29
         Top             =   3480
         Width           =   1320
      End
      Begin VB.Label lblPrecoTabela 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREÇO TABELA R$"
         Height          =   225
         Left            =   150
         TabIndex        =   28
         Top             =   3480
         Width           =   1320
      End
      Begin VB.Image imgCodigo 
         Height          =   240
         Left            =   180
         Picture         =   "frmProduto.frx":13CC
         ToolTipText     =   "Chave Primária"
         Top             =   330
         Width           =   240
      End
      Begin VB.Label lblCodigo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CÓDIGO:"
         Height          =   225
         Left            =   510
         TabIndex        =   27
         ToolTipText     =   "Chave Primária"
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame FramePesquisar 
      Caption         =   "PESQUISAR"
      Height          =   4455
      Left            =   30
      TabIndex        =   24
      Top             =   1620
      Width           =   4875
      Begin VB.ListBox ListPesquisar 
         Height          =   3210
         Left            =   90
         TabIndex        =   12
         Top             =   1140
         Width           =   4680
      End
      Begin VB.TextBox TxtPesquisar 
         Height          =   345
         Left            =   90
         MaxLength       =   20
         TabIndex        =   10
         Top             =   630
         Width           =   4680
      End
      Begin VB.Image imgPesquisar 
         Height          =   240
         Left            =   120
         Picture         =   "frmProduto.frx":1956
         ToolTipText     =   "Chave Primária"
         Top             =   330
         Width           =   240
      End
      Begin VB.Label lblPesquisar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOME PRODUTO:"
         Height          =   225
         Left            =   450
         TabIndex        =   25
         ToolTipText     =   "Chave Primária"
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Frame FrameToolbar 
      Height          =   735
      Left            =   30
      TabIndex        =   22
      Top             =   810
      Width           =   9855
      Begin VB.CommandButton btnImprimir 
         Height          =   495
         Left            =   5370
         Picture         =   "frmProduto.frx":1EE0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir os dados da tabela produto"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnAtualizar 
         Height          =   495
         Left            =   2850
         Picture         =   "frmProduto.frx":246A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Faz uma releitura dos dados"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnExcluir 
         Height          =   500
         Left            =   4740
         Picture         =   "frmProduto.frx":29F4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exclui o registro atualmente selecionado"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnInserir 
         Height          =   500
         Left            =   3480
         Picture         =   "frmProduto.frx":2F7E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Adiciona um novo registro"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnEditar 
         Height          =   500
         Left            =   4110
         Picture         =   "frmProduto.frx":3508
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Edita o registro atualmente selecionado"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnUltimo 
         Height          =   500
         Left            =   2010
         Picture         =   "frmProduto.frx":3A92
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Move os dados para o final do registro"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnProximo 
         Height          =   500
         Left            =   1380
         Picture         =   "frmProduto.frx":3BDC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Move os dados para o próximo registro"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnAnterior 
         Height          =   500
         Left            =   735
         Picture         =   "frmProduto.frx":3D26
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Move os dados para o registro anterior"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnInicio 
         Height          =   500
         Left            =   90
         Picture         =   "frmProduto.frx":3E70
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Move os dados para o início do registro"
         Top             =   160
         Width           =   600
      End
      Begin VB.CommandButton btnSair 
         Height          =   500
         Left            =   6000
         Picture         =   "frmProduto.frx":3FBA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Encerra as atividades neste cadatro"
         Top             =   160
         Width           =   600
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
      ScaleWidth      =   9915
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   9915
      Begin VB.Image imgBanner 
         Height          =   480
         Left            =   180
         Picture         =   "frmProduto.frx":4544
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmProduto.frx":4E0E
         Height          =   495
         Left            =   960
         TabIndex        =   21
         Top             =   90
         Width           =   8640
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image imgDataBase 
      Height          =   240
      Left            =   120
      Picture         =   "frmProduto.frx":4E9C
      ToolTipText     =   "Banco de dados local"
      Top             =   6630
      Width           =   240
   End
   Begin VB.Image imgTotalReg 
      Height          =   240
      Left            =   120
      Picture         =   "frmProduto.frx":5426
      Top             =   6180
      Width           =   240
   End
   Begin VB.Label lblTotalReg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Registro(s): 0"
      Height          =   225
      Left            =   420
      TabIndex        =   35
      Top             =   6180
      Width           =   2400
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   225
      Left            =   420
      TabIndex        =   23
      Top             =   6630
      Width           =   120
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Principal"
      Begin VB.Menu mnuAtualizar 
         Caption         =   "Atualizar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuInserir 
         Caption         =   "Novo Registro"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuExcluir 
         Caption         =   "Excluir"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditar 
         Caption         =   "Editar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuNavegacao 
      Caption         =   "&Navegação"
      Begin VB.Menu mnuInicio 
         Caption         =   "&Início"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuAnterior 
         Caption         =   "&Anterior"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuProximo 
         Caption         =   "P&róximo"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuUltimo 
         Caption         =   "&Último"
         Shortcut        =   ^U
      End
   End
End
Attribute VB_Name = "frmProduto"
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

'Private mValue As String

Private Sub btnAnterior_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim RecordSet As Long
  '
  RecordSet = ListPesquisar.ListIndex - 1
  If RecordSet < 0 Then RecordSet = 0
  ListPesquisar.ListIndex = RecordSet
End Sub

Private Sub btnEditar_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim RecordSet As Long
  Dim Dados     As String
  '
  RecordSet = ListPesquisar.ListIndex + 1
  '
  If RecordSet <= 0 Then
    btnEditar.Enabled = False
    Exit Sub
  End If
  '
  Dados = RetornarRegistro(RecordSet)
  '
  If Len(Trim$(Dados)) Then
    With frmEdit
      .Dados = Dados
      .RecordSet = RecordSet
      .Show vbModal
      Call CarregarDados
      ListPesquisar.ListIndex = RecordSet - 1
    End With
  End If

End Sub

Private Function ValorPadrao() As String
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Tmp As String
  ' MONTA O VALOR PADRÃO DE RETORNO
  Tmp = txtNome.Tag & "|" 'Nome
  Tmp = Tmp & txtCodigo.Tag & "|" 'Codigo
  Tmp = Tmp & txtPrecoCompra.Tag & "|" 'PrecoCompra
  Tmp = Tmp & txtPrecoVenda.Tag & "|" 'PrecoVenda
  Tmp = Tmp & txtQuantidade.Tag & "|" 'Quantidade
  Tmp = Tmp & txtPrecoTabela.Tag & "|" 'Tabela
  Tmp = Tmp & txtTotalPago.Tag & "|" 'TotalPago
  Tmp = Tmp & Me.txtUltimoPreco.Tag & "|" 'Ultimo
  ValorPadrao = Tmp
End Function

Private Function RetornarRegistro(ByVal RecordSet As Long) As String
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Buffer         As Integer
  Dim Produto        As TProduto
  Dim Tmp            As String
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Buffer = FreeFile
  Tmp = ValorPadrao()
  '
  RetornarRegistro = Tmp 'VALOR PADRÃO
  '
  Open dataBaseProduto For Random As #Buffer Len = Len(Produto)
  '
  Get #Buffer, RecordSet, Produto
  Tmp = Trim$(Produto.Nome) & "|"
  Tmp = Tmp & Trim$(Produto.Codigo) & "|"
  Tmp = Tmp & Trim$(Produto.PrecoCompra) & "|"
  Tmp = Tmp & Trim$(Produto.PrecoVenda) & "|"
  Tmp = Tmp & Trim$(Produto.Quantidade) & "|"
  Tmp = Tmp & Trim$(Produto.Tabela) & "|"
  Tmp = Tmp & Trim$(Produto.TotalPago) & "|"
  Tmp = Tmp & Trim$(Produto.Ultimo) & "|"
  '
  RetornarRegistro = Tmp
  '
  Close #Buffer
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Close 'FECHA TODOS OS ARQUIVOS ABERTOS
  Screen.MousePointer = vbDefault
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox Err.Number & ", " & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function

Private Sub btnImprimir_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Result As Long
  '
  Result = MsgBox("Deseja Imprimir o relatório do cadastro produtos?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Atenção")
  '
  If Result = vbYes Then
    Call basProdutoReport.GerarRelatorio
    Call CarregarDados
  End If
  '
  TxtPesquisar.SetFocus
  
End Sub

Private Sub btnInicio_Click()
  ListPesquisar.ListIndex = 0
End Sub

Private Sub btnInserir_Click()
  frmInsert.Show vbModal
  Call CarregarDados
  Call btnUltimo_Click
End Sub

Private Sub btnProximo_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim RecordSet As Long
  Dim Total     As Long
  '
  Total = ListPesquisar.ListCount - 1
  '
  RecordSet = ListPesquisar.ListIndex + 1
  If RecordSet > Total Then RecordSet = Total
  ListPesquisar.ListIndex = RecordSet
End Sub

Private Sub btnSair_Click()
  Unload Me
End Sub

Private Sub btnAtualizar_Click()
  Call CarregarDados
  TxtPesquisar.SetFocus
End Sub

Private Sub btnExcluir_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Result        As Long
  Dim RegistroAtual As Long
  '
  RegistroAtual = ListPesquisar.ListIndex + 1
  '
  If RegistroAtual <= 0 Then
    btnExcluir.Enabled = False
    Exit Sub
  End If
  '
  Result = MsgBox("Deseja excluir o registro corrente?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Atenção")
  '
  If Result = vbYes Then
    Call basProduto.Excluir(RegistroAtual)
    Call CarregarDados
  End If
  '
  TxtPesquisar.SetFocus
End Sub

Private Sub btnUltimo_Click()
  ListPesquisar.ListIndex = ListPesquisar.ListCount - 1
End Sub

Private Sub Form_Activate()
  TxtPesquisar.SetFocus
End Sub

Private Sub Form_Load()
  'Define o nome do banco de dados
  basProduto.dataBaseProduto = App.Path & "\database\produtos.dat"
  lblDataBase.Caption = basProduto.dataBaseProduto
  Call CarregarDados
End Sub

Private Sub ListPesquisar_Click()
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  If ListPesquisar.ListCount > 0 Then
    '---------------------
    ' DEFINIÇÃO VARIÁVEIS
    '---------------------
    Dim RegistroAtual As Long
    btnEditar.Enabled = True
    '
    RegistroAtual = ListPesquisar.ListIndex + 1
    '
    If (ListPesquisar.ListIndex >= 0) Then
      '---------------------
      ' DEFINIÇÃO VARIÁVEIS
      '---------------------
      Dim Dados() As String
      Dados = Split(RetornarRegistro(RegistroAtual), "|")
      txtNome.Text = Dados(0)
      txtCodigo.Text = Dados(1)
      txtPrecoCompra.Text = Dados(2)
      txtPrecoVenda.Text = Dados(3)
      txtQuantidade.Text = Dados(4)
      txtPrecoTabela.Text = Dados(5)
      txtTotalPago.Text = Dados(6)
      txtUltimoPreco.Text = Dados(7)
    Else
      ' REINICIA O PROCESSO DE GRAVAÇÃO DE DADOS
      Call basMain.valorPadraoTextBox(Me)
    End If
  Else
    btnEditar.Enabled = False
  End If
  '
  Call InfoRecordset
  Exit Sub
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Screen.MousePointer = vbDefault
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox Err.Number & ", " & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Sub

Private Sub ListPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 38    '^
      If ListPesquisar.ListIndex = 0 Then
        KeyCode = 0
        TxtPesquisar.SetFocus
        ListPesquisar.ListIndex = -1
      End If
  End Select
End Sub

Private Sub ListPesquisar_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    Call btnEditar_Click
  End If
End Sub

Private Sub mnuAnterior_Click()
  Call btnAnterior_Click
End Sub

Private Sub mnuAtualizar_Click()
  Call btnAtualizar_Click
End Sub

Private Sub mnuEditar_Click()
  Call btnEditar_Click
End Sub

Private Sub mnuExcluir_Click()
  Call btnExcluir_Click
End Sub

Private Sub mnuImprimir_Click()
  Call btnImprimir_Click
End Sub

Private Sub mnuInicio_Click()
  Call btnInicio_Click
End Sub

Private Sub mnuInserir_Click()
  Call btnInserir_Click
End Sub

Private Sub mnuProximo_Click()
  Call btnProximo_Click
End Sub

Private Sub mnuSair_Click()
  Call btnSair_Click
End Sub

Private Sub mnuUltimo_Click()
  Call btnUltimo_Click
End Sub

Private Sub TxtPesquisar_Change()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Result As Long
  '
  Result = SendMessage(ListPesquisar.hwnd, LB_FINDSTRING, -1, ByVal CStr(TxtPesquisar.Text))
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  ListPesquisar.ListIndex = Result
End Sub

Private Sub CarregarDados()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Buffer         As Integer
  Dim Produto        As TProduto
  Dim Registro       As Long
  Dim TotalRegistros As Long
  ' LIMPA A CAIXA DE SELEÇÃO
  ListPesquisar.Clear
  Call basMain.valorPadraoTextBox(Me)
  '
  Buffer = FreeFile
  '
  Open dataBaseProduto For Random As #Buffer Len = Len(Produto)
  TotalRegistros = LOF(Buffer) \ Len(Produto) 'ENCONTRO O NÚMERO TOTAL DE LINHAS

  ' INICIO O LOOP FOR
  For Registro = 1 To TotalRegistros
    Get #Buffer, Registro, Produto
    ListPesquisar.AddItem Trim$(Produto.Nome)
  Next Registro

  Close #Buffer

  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '> VERIFICO SE O LISTBOX TEM DADOS <
  '
  If ListPesquisar.ListCount > 0 Then
    ListPesquisar.ListIndex = 0
    Call StatusMenuToolbar(True)
  Else
    Call StatusMenuToolbar(False)
  End If
  '
  Call InfoRecordset
  Screen.MousePointer = vbDefault
End Sub

Private Sub TxtPesquisar_KeyDown(KeyCode As Integer, Shift As Integer)
  '
  Select Case KeyCode
    Case 38    '
      KeyCode = 0
      TxtPesquisar.SetFocus
      'v
      KeyCode = 0
      '
      If ListPesquisar.ListCount > 0 Then
        ListPesquisar.SetFocus
      End If
      '
  End Select
End Sub

Private Sub TxtPesquisar_GotFocus()
  TxtPesquisar.SelStart = 0
  TxtPesquisar.SelLength = Len(TxtPesquisar.Text)
End Sub

Private Sub StatusMenuToolbar(ByVal Status As Boolean)
  ' MENUS
  mnuEditar.Enabled = Status
  mnuImprimir.Enabled = Status
  mnuExcluir.Enabled = Status
  mnuUltimo.Enabled = Status
  mnuProximo.Enabled = Status
  mnuAnterior.Enabled = Status
  mnuInicio.Enabled = Status
  'TOOLBAR
  btnEditar.Enabled = Status
  btnImprimir.Enabled = Status
  btnExcluir.Enabled = Status
  btnUltimo.Enabled = Status
  btnProximo.Enabled = Status
  btnAnterior.Enabled = Status
  btnInicio.Enabled = Status
End Sub

Private Sub InfoRecordset()
  Dim RecordSet As Long
  '
  RecordSet = ListPesquisar.ListIndex + 1
  '
  Select Case ListPesquisar.ListCount
    Case Is <= 0
      lblTotalReg.Caption = "Nenhum registro encontrado!"
    Case 1
      lblTotalReg.Caption = "Total Registro: " & RecordSet & " de " & ListPesquisar.ListCount
    Case Else
      lblTotalReg.Caption = "Total Registro(s): " & RecordSet & " de " & ListPesquisar.ListCount
  End Select
End Sub

