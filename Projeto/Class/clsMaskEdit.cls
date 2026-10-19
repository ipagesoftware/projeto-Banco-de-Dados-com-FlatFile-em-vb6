VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaskEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
'
Public Enum eFormatText
  'U = Ucase, L = Lcase, P = ProperCase, D = Data, C = Currency, E = Email
  emUCase = 0
  emLCase = 1
  emProperCase = 2
  emData = 3
  emCurrency = 4
  emEmail = 5
  emCheck = 6
  emInput = 7
  emOBS = 8
  emNone = -1
End Enum
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : Mask                                           <
'> Data/Hora             : 29/10/2006 11:35                               <
'> Autor                 : Diógenes Dias                                  <
'> Descrição             : <
'> Parâmetros Passados   : <
'> Parâmetros Retornados : <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 29/10/2006 11:35                               <
'>                                                                        <
'> © Copyright DIÓGENES DIAS
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Sub Mask(ByRef p_Ctrl As Object, _
                ByRef p_KeyAscii As Integer, _
                Optional ByVal M_FormatText As String, _
                Optional ByVal p_Mask As String = "")

  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  'U = Ucase, L = Lcase, P = ProperCase, D = Data, C = Currency
  '
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Const m_F = "UuLuPpDdCc"
  Const ONLY_WORD = "0123456789"
  Dim Valid_CHAR     As String
  Dim Flag_Format    As Integer
  Dim CHAR1          As String
  Dim CHAR2          As String
  Dim L              As Integer
  Dim Tmp            As String
  Dim mFormatData    As String
  Dim FormatCurrency As String
  Dim Slash          As Variant
  Dim Slash2         As Variant
  Dim m_Decimal      As Integer
  '
  Valid_CHAR = "abcdefghijklmnopqrstuwxyz_.@0123456789-" & Chr$(8) & Chr$(3) & Chr$(22) & Chr$(24) & Chr$(26) & Chr$(13) & Chr$(10)
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '> Habilita as teclas de atalho <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '
  Select Case p_KeyAscii
    Case 3, 22, 24, 26
      'Permite Ctrl+C, Ctrl+V e Ctrl+X
      Exit Sub
    Case vbKeyReturn
      'SendKeys "{TAB}"
      Exit Sub
  End Select
  '
Begin:
  '
  L = Len(p_Mask)
  '
  Flag_Format = InStr(1, m_F, M_FormatText, vbTextCompare)
  '
  If Flag_Format > 0 Or Len(Flag_Format) < Len(m_F) Then
    '
    Select Case UCase(Mid(M_FormatText, 1, 1))
      Case "C"
        If Len(Trim(p_Mask)) = False Then
          FormatCurrency = ""
          Slash = ","
          Slash2 = "."
        Else
          If InStr(1, p_Mask, ".", vbTextCompare) <> 0 Then
            m_Decimal = InStr(1, p_Mask, ".", vbTextCompare) - 1
            FormatCurrency = Mid(p_Mask, 1, m_Decimal) + Mid(p_Mask, m_Decimal + 2)
            Slash = ","
            Slash2 = "."
          ElseIf InStr(1, p_Mask, ".", vbTextCompare) <> 0 Then
            m_Decimal = InStr(1, p_Mask, ".", vbTextCompare) - 1
            FormatCurrency = Mid(p_Mask, 1, m_Decimal) + Mid(p_Mask, m_Decimal + 2)
            Slash = ","
            Slash2 = "."
          Else
            FormatCurrency = ""
            Slash = ","
            Slash2 = "."
          End If
          '
          Tmp = p_Ctrl.Text
          '
          If InStr(1, Tmp, FormatCurrency, vbTextCompare) = False Then
            Tmp = FormatCurrency & " " & Tmp
            p_Ctrl.Text = Tmp
            p_Ctrl.SelStart = Len(Tmp)
          End If
        End If
        '
        'Currency
        '*** Permite só números com máscara para Moeda ***
        '
        Select Case p_KeyAscii
          Case 3, 22, 24, 26
            'Permite Ctrl+C, Ctrl+V e Ctrl+X
            Exit Sub
          Case 8
            Exit Sub
          Case Asc(Slash2)
            '*** Evita a repetição da vírgula ***
            If InStr(p_Ctrl, Slash) <> 0 Then
              p_KeyAscii = 0
              Exit Sub
            End If
            '
            p_KeyAscii = Asc(Slash)
          Case Asc(Slash)
            '*** Evita a repetição da vírgula ***
            If InStr(p_Ctrl, Slash) <> 0 Then
              p_KeyAscii = 0
            End If
          Case 48 To 57, Asc(Slash)
          Case Asc("-")
            '*** Evita a repetição do sinal de menos ***
            If InStr(p_Ctrl, "-") <> 0 Then
              p_KeyAscii = 0
              Exit Sub
            End If
            '
          Case Else
            p_KeyAscii = 0
        End Select
        '
        Exit Sub
        '
      Case "D"
        If Len(Trim(p_Mask)) = False Then
          mFormatData = "DD/MM/YYYY"
          p_Mask = "99/99/9999"
          Slash = Split(p_Mask, "/", , vbTextCompare)
        Else
          '
          If InStr(1, p_Mask, "-", vbTextCompare) <> 0 Then
            Slash = Split(p_Mask, "-", , vbTextCompare)
            mFormatData = "YYYY-MM-DD"    'Formato do MYSQL
          ElseIf InStr(1, p_Mask, "/", vbTextCompare) <> 0 Then
            Slash = Split(p_Mask, "/", , vbTextCompare)
            mFormatData = "MM/DD/YYYY"    'Formato americano
          End If
          '
          If UBound(Slash) < 2 Then
            Slash = Split("99/99/9999", "/", , vbTextCompare)
          End If
          '
        End If
        '
        '*** Permite só números com máscara para data ***
        '
        Select Case p_KeyAscii
          Case vbKeyReturn
            p_KeyAscii = 0
            p_Ctrl.Text = Format(p_Ctrl.Text, mFormatData)
          Case 8
            Exit Sub
          Case Asc("/")
            Exit Sub
          Case 3, 22, 24, 26
            'Permite Ctrl+C, Ctrl+V e Ctrl+X
            Exit Sub
          Case 48 To 57
            If Len(p_Ctrl.Text) = Len(Slash(0)) Then
              p_Ctrl.Text = p_Ctrl.Text & "/"
              p_Ctrl.SelStart = Len(p_Ctrl.Text)
            ElseIf Len(p_Ctrl.Text) = (Len(Slash(0)) + Len(Slash(1)) + 1) Then
              p_Ctrl.Text = p_Ctrl.Text & "/"
              p_Ctrl.SelStart = Len(p_Ctrl.Text)
            End If
            '
          Case Else
            p_KeyAscii = 0
        End Select
        '
        Exit Sub
        '
      Case "U"    ' maiúscula
        p_KeyAscii = Asc(UCase(Chr(p_KeyAscii)))
      Case "L"    ' Minúscula
        p_KeyAscii = Asc(LCase(Chr(p_KeyAscii)))
      Case "E"
        p_KeyAscii = Asc(LCase(Chr(p_KeyAscii)))
        '
        If InStr(1, Valid_CHAR, Chr$(p_KeyAscii), vbTextCompare) = 0 Then
          p_KeyAscii = 0
        End If
        '
      Case "P"    'Nome próprio
        Tmp = p_Ctrl.Text + Chr$(p_KeyAscii)
        '
        '> Não avalia os espaços <
        '
        If p_KeyAscii <> 32 Then
          Tmp = ProperCase(Tmp)
        End If
        '
        p_Ctrl.Text = Left(Tmp, Len(Tmp) - 1)
        p_Ctrl.SelStart = Len(p_Ctrl)
        p_KeyAscii = Asc(Right(Tmp, 1))
        '
    End Select
    '
  End If
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '> Caractere pode gerar erros na consulta <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '
  If p_KeyAscii = Asc("'") Then
    p_KeyAscii = 0
    Exit Sub
  ElseIf p_KeyAscii = 8 Then    'BackSpace
    Exit Sub
  End If
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '> Não possui nenhuma formatação <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '
  If L = 0 Then
    Exit Sub
  End If
  '
  Tmp = Left(p_Ctrl.Text + Chr$(p_KeyAscii), Len(p_Mask))
  '
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '> Não permite que o texto seja maior do que a máscara <
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  '
  If Len(Tmp) >= Len(p_Mask) Then
    p_Ctrl.Text = Tmp
    p_Ctrl.SelStart = Len(Tmp)
    p_KeyAscii = 0
    Exit Sub
  End If
  '
  CHAR1 = Mid(p_Mask, Len(Tmp), 1)
  CHAR2 = Mid(Tmp, Len(Tmp), 1)
  '
  Select Case UCase(CHAR1)
    Case "9"
      'Apenas número
      If IsNumeric(CHAR2) = False Then
        p_KeyAscii = 0
      End If
    Case "@"
      'Apenas letras
      If InStr(1, ONLY_WORD, UCase(Chr(p_KeyAscii)), vbTextCompare) <> 0 Then
        p_KeyAscii = 0
      End If
      '
    Case "-", "/", "(", ")", ".", "+", "*", " ", ":"
      Tmp = p_Ctrl.Text & CHAR1
      p_Ctrl.Text = Tmp
      p_Ctrl.SelStart = Len(Tmp)
      GoTo Begin
    Case Else
      If InStr(1, p_Mask, Chr$(p_KeyAscii), vbTextCompare) = False Then
        p_KeyAscii = 0
      End If
  End Select
  '
  Exit Sub
  'TRATAMENTO DE ERROS
LogErrorHandler:
  p_KeyAscii = 0
  Err.Clear
End Sub
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : ProperCase                                     <
'> Data/Hora             : 29/10/2006 21:25                               <
'> Autor                 : Diógenes Dias                                  <
'> Descrição             : Converte as primeiras letra dos nomes em       <
'>                         maiúsculas menos: de, da, do, das, dos, a, e e.<
'>                         Retira os espaços em branco excedentes entre   <
'>                         as p_Words.                                    <
'> Parâmetros Passados   : Uma String.                                    <
'> Parâmetros Retornados : Uma string no formato nome próprio.            <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 29/10/2006 21:25                               <
'>                                                                        <
'> © Copyright DIÓGENES DIAS
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Function ProperCase(ByVal p_Word As String) As String
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Resto As String
  Dim Parte As String
  Dim Posic As Byte
  '
  Resto = Trim$(LCase$(p_Word)) & " "
  '
  Do Until Resto = " "
    Posic = InStr(1, Resto, " ")
    Parte = Mid$(Resto, 1, Posic)
    '
    If Trim$(Parte) <> "e" And Trim$(Parte) <> "a" And Trim$(Parte) <> "do" And Trim$(Parte) <> "dos" And Trim$(Parte) <> "da" And Trim$(Parte) <> "das" And Trim$(Parte) <> "de" Then
      '
      Parte = UCase$(Left$(Parte, 1)) & Mid$(Parte, 2)
      '
    End If
    '
    ProperCase = ProperCase & Parte
    Resto = Trim$(Mid$(Resto, Posic)) & " "
  Loop
  '
  ProperCase = Trim$(ProperCase)
  '
End Function
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : SelText                                        <
'> Data/Hora             : 29/10/2006 21:27                               <
'> Autor                 : Diógenes Dias                                  <
'> Descrição             : Seleciona o p_Ctrl quando entra em foco em uma <
'>                         caixa de controle.                             <
'> Parâmetros Passados   : O objeto p_Ctrl, a posição inicial de seleção  <
'>                         a posição final de seleção.                    <
'> Parâmetros Retornados : Nenhum                                         <
'> Dependências          :                                                <
'> Categoria             :                                                <
'> Modificada em         : 29/10/2006 21:27                               <
'>                                                                        <
'> © Copyright DIÓGENES DIAS
'>                                                                        <
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'
Public Sub SelText(ByVal p_Ctrl As Object, _
                   Optional ByVal p_Start As Integer = False, _
                   Optional ByVal p_MaxLength As Integer = 255)

  '
  p_Ctrl.SelStart = p_Start
  p_Ctrl.SelLength = p_MaxLength
  '
End Sub

