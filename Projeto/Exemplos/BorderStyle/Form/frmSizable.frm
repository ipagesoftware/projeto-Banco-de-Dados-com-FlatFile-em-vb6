VERSION 5.00
Begin VB.Form frmSizable 
   BackColor       =   &H00C0FFFF&
   Caption         =   "#"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5595
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
   ScaleHeight     =   3135
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSair 
      Caption         =   "Sair"
      Height          =   570
      Left            =   2250
      TabIndex        =   0
      Top             =   2430
      Width           =   1230
   End
   Begin VB.Label lblPropriedadeBorderStyle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   225
      Left            =   80
      TabIndex        =   1
      Top             =   150
      Width           =   120
   End
End
Attribute VB_Name = "frmSizable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSair_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Title As String
  '
  Title = "Propriedade da Janela BorderStyle = " & Me.BorderStyle
  '
  lblPropriedadeBorderStyle.Caption = Title
  Me.Caption = Title
End Sub
