VERSION 5.00
Begin VB.Form frmNone 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
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
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
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
      Caption         =   "Propriedade BorderStyle = 0"
      Height          =   225
      Left            =   80
      TabIndex        =   1
      Top             =   390
      Width           =   3240
   End
End
Attribute VB_Name = "frmNone"
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
End Sub
