VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Exemplo de CRUD com arquivo texto"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgBackGround 
      Height          =   5655
      Left            =   360
      Picture         =   "frmMain.frx":1782
      Stretch         =   -1  'True
      Top             =   300
      Width           =   8355
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "&Principal"
      Begin VB.Menu mnuCadProdutos 
         Caption         =   "Ca&dastro Produtos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuBarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Sair"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuLicenca 
         Caption         =   "&Licença de Uso"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjudaBar6 
         Caption         =   "* * SUPORTE * *"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpDesk 
         Caption         =   "Help Desk"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuWhatsApp 
         Caption         =   "WhatsApp"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuTelegram 
         Caption         =   "Telegram"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuGitHub 
         Caption         =   "GitHub"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "&Sobre..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> 1 Then
    If MsgBox("Deseja sair desta aplicação?", vbQuestion Or vbYesNo, "Atenção") = vbYes Then
      Unload Me
    Else
      Cancel = True
    End If
    Exit Sub
  End If
End Sub

Private Sub Form_Resize()
  'AJUSTA O TAMANHO DA IMAGEM DE FUNDO AUTOMATICAMENTE PARA FICAR
  'COM O MESMO TAMANHO DA JANELA
  imgBackGround.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub imgBackGround_MouseDown(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    y As Single)
  If Button = 2 Then
    'AO CLICAR COM O BOTÃO DIREITO DO MOUSE O
    'MENU PRINCIPAL SERÁ EXIBIDO
    PopupMenu mnuPrincipal
  End If
End Sub

Private Sub mnuCadProdutos_Click()
  frmProduto.Show vbModal
End Sub

Private Sub mnuGitHub_Click()
  Const URL As String = "https://github.com/ipagesoftware"
  Call ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuHelpDesk_Click()
  Const URL As String = "https://www.ipagesoftware.com.br/helpdesk/"
  Call ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuLicenca_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim FileName As String
  '
  FileName = App.Path & "\Doc\LICENSE.pdf"
  Call ShellExecute(0, "open", FileName, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuSair_Click()
  If MsgBox("Deseja sair desta aplicação?", vbQuestion Or vbYesNo, "Atenção") = vbYes Then
    Unload Me
  End If
End Sub

Private Sub mnuSobre_Click()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Tmp As String
  '
  Tmp = App.CompanyName & " - Ver. " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine
  Tmp = Tmp & App.EXEName
  '
  ShellAbout Me.hwnd, App.LegalCopyright, Tmp, Me.Icon
End Sub

Private Sub mnuTelegram_Click()
  Call ShellExecute(0, "open", "https://t.me/ipagesoftware", vbNullString, vbNullString, 1)
End Sub

Private Sub mnuWhatsApp_Click()
  Dim URL      As String
  '"https://api.whatsapp.com/send?phone=5581986352819&text=Ol%C3%A1%20tudo%20bem%3F%0D%0A%0D%0AGostaria%20que%20voc%C3%AA%20me%20ajudasse%20a%20fazer%20um%20or%C3%A7amento.%0D%0A%0D%0ADesde%20j%C3%A1%20agrade%C3%A7o%20a%20aten%C3%A7%C3%A3o%20e%20fico%20no%20aguardo%20do%20seu%20pronunciamento.%0D%0A%0D%0ANOTA%3A%0D%0ASolicita%C3%A7%C3%A3o%20realizada%20atrav%C3%A9s%20do%20nosso%20WEBSITE%20em%3A%2007%2F04%2F2025%2023%3A13%3A47" target="_blank"><i class="fa fa-whatsapp"></i> <?= FONES[1] ?>
                                            
  URL = "https://api.whatsapp.com/send"
  URL = URL & "?phone=+5581986152352"
  URL = URL & "&text=Olá, gostaria de esclarecer algumas dúvidas sobre o "
  URL = URL & "aplicativo FlatFile em Visual Basic 6.0."
  '
  Call ShellExecute(0, "open", URL, vbNullString, vbNullString, 1)
End Sub

