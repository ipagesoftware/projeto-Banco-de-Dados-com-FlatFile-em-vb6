VERSION 5.00
Begin VB.Form frmSizableToolWindow 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   5625
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
   ScaleHeight     =   3135
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
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
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   80
      TabIndex        =   1
      Top             =   150
      Width           =   120
   End
End
Attribute VB_Name = "frmSizableToolWindow"
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

