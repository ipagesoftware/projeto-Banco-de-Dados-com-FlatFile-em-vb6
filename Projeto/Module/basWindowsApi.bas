Attribute VB_Name = "basWindowsApi"
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
Public Declare Function ShellAbout _
               Lib "Shell32.dll" _
               Alias "ShellAboutA" (ByVal hwnd As Long, _
                                    ByVal szApp As String, _
                                    ByVal szOtherStuff As String, _
                                    ByVal hIcon As Long) As Long
'*** início: UTILIZADA NO ROTINA DE BUSCA ****
'
Public Const LB_FINDSTRING = &H18F
Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long

'*** fim: UTILIZADA NO ROTINA DE BUSCA ****

Private Declare Function GetTempPath _
                Lib "kernel32" _
                Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                      ByVal lpBuffer As String) As Long

Private Declare Function GetTempFileName _
                Lib "kernel32" _
                Alias "GetTempFileNameA" (ByVal lpPathName As String, _
                                          ByVal lpPrefixString As String, _
                                          ByVal uUnique As Long, _
                                          ByVal lpTempFileName As String) As Long

' UTILIZADO PELA ROTINA DE IMPRESSÃO
Public Declare Function ShellExecute _
               Lib "Shell32.dll" _
               Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long
    
'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : PastaTemp
' Description: Encontra a pasta temporário do windows
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 01:52:15
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function PastaTemp() As String
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Buffer  As String
  Dim Tamanho As Long
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Buffer = String$(260, vbNullChar)
  Tamanho = GetTempPath(260, Buffer)
  '
  If Tamanho > 0 Then
    PastaTemp = Left$(Buffer, Tamanho)
  End If
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
' Procedure  : CriarArquivoTemp
' Description: Cria um nome de arquivo com caractres randômicos e extenção .TMP
'
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 01:52:39
'
' Parameters : Prefixo (String)
'--------------------------------------------------------------------------------
Public Function CriarArquivoTemp(ByVal Prefixo As String) As String
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Buffer  As String
  Dim Retorno As Long
  Dim Pasta   As String
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Pasta = PastaTemp()
  Buffer = String$(260, vbNullChar)
  Retorno = GetTempFileName(Pasta, Prefixo, 0, Buffer)
  '
  If Retorno <> 0 Then
    CriarArquivoTemp = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
  End If
  '
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

