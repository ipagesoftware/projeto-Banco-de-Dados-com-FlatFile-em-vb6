Attribute VB_Name = "basMain"
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

Public Sub Main()
  frmProduto.Show vbModal
End Sub

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : valorPadraoTextBox
' Description: Sub que faz uma varredura em todos os objetos de um form em
'                    busca de uma caixa de texto, ela lê a propriedade Tag desta
'                    caixa de texto e atribui este valor a mesma
' Created by : Diógens Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 10/3/2026 09:52:27
'
' Parameters : Frm é o objeto Form que deseja realizar a varredura
'--------------------------------------------------------------------------------
Public Sub valorPadraoTextBox(ByRef frm As Form)
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Objeto As Variant
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  For Each Objeto In frm
    Select Case TypeName(Objeto)
      Case "TextBox"
        Objeto.Text = Objeto.Tag
    End Select
  Next
  '
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

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : DataPorExtenso
' Description: [type_description_here]
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 10/4/2026 23:31:15
'
' Parameters : Data (Date)
'--------------------------------------------------------------------------------
Public Function DataPorExtenso(ByVal Data As Date) As String
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Result As String
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Result = WeekdayName(Weekday(Data), False, vbUseSystemDayOfWeek) & ", " & Format$(Data, "dd") & " de " & MonthName(Month(Data), False) & " de " & Format$(Data, "yyyy") & " " & Format$(Now(), "hh:mm:ss")
  DataPorExtenso = HtmlEncode(Result)
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Screen.MousePointer = vbDefault
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox Err.Number & ", " & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : HtmlEncode
' Description: Converte os caracteres acentuados para acentuação HTML
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 10/04/2026 23:31:40
'
' Parameters : pTexto (String)
'
' Example    :
' Dim Menssagem As String
' Dim Result   As String
' Menssagem = "Diógens Dias"
' Result = HtmlEncode(Menssagem)
' Debug.Print Result
' Resultado:
' Di&oacute;gens Dias
'--------------------------------------------------------------------------------
Public Function HtmlEncode(ByVal pTexto As String) As String
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  pTexto = Replace(pTexto, "&", "&amp;")
  pTexto = Replace(pTexto, "á", "&aacute;")
  pTexto = Replace(pTexto, "à", "&agrave;")
  pTexto = Replace(pTexto, "â", "&acirc;")
  pTexto = Replace(pTexto, "ã", "&atilde;")
  pTexto = Replace(pTexto, "ä", "&auml;")
  pTexto = Replace(pTexto, "Á", "&Aacute;")
  pTexto = Replace(pTexto, "À", "&Agrave;")
  pTexto = Replace(pTexto, "Â", "&Acirc;")
  pTexto = Replace(pTexto, "Ã", "&Atilde;")
  pTexto = Replace(pTexto, "Ä", "&Auml;")
  pTexto = Replace(pTexto, "é", "&eacute;")
  pTexto = Replace(pTexto, "è", "&egrave;")
  pTexto = Replace(pTexto, "ê", "&ecirc;")
  pTexto = Replace(pTexto, "ë", "&euml;")
  pTexto = Replace(pTexto, "É", "&Eacute;")
  pTexto = Replace(pTexto, "È", "&Egrave;")
  pTexto = Replace(pTexto, "Ê", "&Ecirc;")
  pTexto = Replace(pTexto, "Ë", "&Euml;")
  pTexto = Replace(pTexto, "í", "&iacute;")
  pTexto = Replace(pTexto, "ì", "&igrave;")
  pTexto = Replace(pTexto, "î", "&icirc;")
  pTexto = Replace(pTexto, "ï", "&iuml;")
  pTexto = Replace(pTexto, "Í", "&Iacute;")
  pTexto = Replace(pTexto, "Ì", "&Igrave;")
  pTexto = Replace(pTexto, "Î", "&Icirc;")
  pTexto = Replace(pTexto, "Ï", "&Iuml;")
  pTexto = Replace(pTexto, "ó", "&oacute;")
  pTexto = Replace(pTexto, "ò", "&ograve;")
  pTexto = Replace(pTexto, "ô", "&ocirc;")
  pTexto = Replace(pTexto, "õ", "&otilde;")
  pTexto = Replace(pTexto, "ö", "&ouml;")
  pTexto = Replace(pTexto, "Ó", "&Oacute;")
  pTexto = Replace(pTexto, "Ò", "&Ograve;")
  pTexto = Replace(pTexto, "Ô", "&Ocirc;")
  pTexto = Replace(pTexto, "Õ", "&Otilde;")
  pTexto = Replace(pTexto, "Ö", "&Ouml;")
  pTexto = Replace(pTexto, "ú", "&uacute;")
  pTexto = Replace(pTexto, "ù", "&ugrave;")
  pTexto = Replace(pTexto, "û", "&ucirc;")
  pTexto = Replace(pTexto, "ü", "&uuml;")
  pTexto = Replace(pTexto, "Ú", "&Uacute;")
  pTexto = Replace(pTexto, "Ù", "&Ugrave;")
  pTexto = Replace(pTexto, "Û", "&Ucirc;")
  pTexto = Replace(pTexto, "Ü", "&Uuml;")
  pTexto = Replace(pTexto, "ç", "&ccedil;")
  pTexto = Replace(pTexto, "Ç", "&Ccedil;")
  pTexto = Replace(pTexto, "ñ", "&ntilde;")
  pTexto = Replace(pTexto, "Ñ", "&Ntilde;")
  pTexto = Replace(pTexto, "Ñ", "&copy;")
  '
  HtmlEncode = pTexto 'RETORNA O TEXTO CONVERTIDO PARA ACENTUAÇÃO HTML
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Screen.MousePointer = vbDefault
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox Err.Number & ", " & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function
