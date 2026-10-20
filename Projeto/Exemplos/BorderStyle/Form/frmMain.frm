VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exemplo de Como Definir o Estilo de Uma Janela"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
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
   ScaleHeight     =   4200
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnSair 
      Caption         =   "Sair"
      Height          =   570
      Left            =   2730
      TabIndex        =   2
      Top             =   3450
      Width           =   1230
   End
   Begin VB.CommandButton btnExibir 
      Caption         =   "Exibir"
      Height          =   570
      Left            =   1290
      TabIndex        =   1
      Top             =   3450
      Width           =   1230
   End
   Begin VB.ComboBox cboBorderStyle 
      Height          =   345
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   810
      Width           =   4605
   End
   Begin VB.Label lblBorderStyle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione o Tipo de Borda da Janela:"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   4320
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExibir_Click()
  Select Case cboBorderStyle.ListIndex
    Case 0
      frmNone.Show vbModal
    Case 1
      frmFixedSingle.Show vbModal
    Case 2
      frmSizable.Show vbModal
    Case 3
      frmFixedDialog.Show vbModal
    Case 4
      frmFixedToolWindow.Show vbModal
    Case 5
      frmSizableToolWindow.Show vbModal
    Case Else
      MsgBox "Valor selecionado para a borda do formulário é inválido, tente outra vez!", vbExclamation
  End Select
  '
  cboBorderStyle.SetFocus
End Sub

Private Sub btnSair_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Call LoadCombo
End Sub

Private Function LoadCombo()
  With cboBorderStyle
    .Clear
    .AddItem "None"
    .AddItem "Fixed Single"
    .AddItem "Sizable"
    .AddItem "Fixed Dialog"
    .AddItem "Fixed Tool Window"
    .AddItem "Sizable Tool Window"
  End With
  '
End Function
