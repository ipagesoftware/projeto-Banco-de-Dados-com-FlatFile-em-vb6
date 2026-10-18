Attribute VB_Name = "basProdutoReport"
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

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : CarregarTemplate
' Description: [type_description_here]
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 01:14:38
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function CarregarTemplate() As String
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Buffer   As Integer
  Dim Texto    As String
  Dim Template As String
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Template = App.Path & "\report\produto.html"
  Buffer = FreeFile
  '
  Open Template For Binary Access Read As #Buffer
  '
  Texto = Space$(LOF(Buffer))
  Get #Buffer, , Texto
  Close #Buffer
  '
  CarregarTemplate = Texto
  Exit Function

  'TRATAMENTO DE ERROS
LogErrorHandler:              On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Close 'FECHA TODOS OS ARQUIVOS ABERTOS
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox "Erro ao abrir o arquivo de tempalte" & vbCrLf & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
  CarregarTemplate = ""
End Function

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : GerarRelatorio
' Description: [type_description_here]
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 01:14:56
'
' Parameters :
'--------------------------------------------------------------------------------
Public Sub GerarRelatorio()
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim ArqOrigem                 As Integer
  Dim Produto                   As TProduto
  Dim TotalRegistros            As Long
  Dim Contador                  As Long
  Dim ArquivoTemporario         As String
  Dim Template                  As String
  Dim Template2                 As String
  Dim Body                      As String
  Dim TBody                     As String
  Dim Footer                    As String
  Dim Linha                     As Long
  Dim Row                       As String
  Dim NumPagina                 As Integer
  Const TOTAL_LINHAS_POR_PAGINA As Integer = 27 '27 É O NÚMERO MÁXIMO DE LINHAS POR PÁGINA TIPO A4
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  Screen.MousePointer = vbHourglass
  DoEvents
  ' VERIFICA SE O NOME DO ARQUIVO DE TEXTO PARA O BANCO DE DADOS FOI DEFINIDO
  If Len(dataBaseProduto) = 0 Then
    dataBaseProduto = App.Path & "\database\produtos.dat"
  End If
  
  '----------------------------------------
  ' Cria o nome do arquivo temporário e
  ' Modifica a extenção do arquivo, de TMP
  ' para HTML
  '----------------------------------------
  ArquivoTemporario = CriarArquivoTemp("ib2")
  ArquivoTemporario = Replace(ArquivoTemporario, ".tmp", ".html")
  '----------------------------------------
  ' Abre o arquivo do banco de dados
  '----------------------------------------
  ArqOrigem = FreeFile
  Open dataBaseProduto For Random As #ArqOrigem Len = Len(Produto)
  '----------------------------------------
  ' Descobre quantidade de registros
  '----------------------------------------
  TotalRegistros = LOF(ArqOrigem) \ Len(Produto)
  '----------------------------------------
  ' Validação
  '----------------------------------------
  If TotalRegistros <= 0 Then
    Close #ArqOrigem
    MsgBox "Registro inexistente.", vbExclamation, "Atenção"
    Exit Sub
  End If
  
  ' INICIO O LOOP FOR
  TBody = ""
  Template = ""
  '
  Linha = 0
  NumPagina = 0
  Template = CarregarTemplate()
  '
  If TotalRegistros < TOTAL_LINHAS_POR_PAGINA Then
    For Contador = 1 To TotalRegistros
      Get #ArqOrigem, Contador, Produto
      Linha = Linha + 1
      GoSub GETROW
    Next Contador
    '
    GoSub HEADER
    '
    TBody = ""
    Template2 = Template2 & basMain.HtmlEncode(Body)
    GoSub RODAPE
    Template2 = Template2 & Footer
    Template = Replace(Template, "[@PAGE]", Template2)
  Else
    For Contador = 1 To TotalRegistros
      Get #ArqOrigem, Contador, Produto
      Linha = Linha + 1
      '
      GoSub GETROW
      '
      If Linha = TOTAL_LINHAS_POR_PAGINA Then
        GoSub HEADER
        If Contador = TotalRegistros Then
          Body = Body & "<div class=" & Chr$(34) & "total" & Chr$(34) & ">Total Geral de registro(s) Listado(s): " & TotalRegistros & "</div>"
        End If
        Linha = 0
        TBody = ""
        Template2 = Template2 & basMain.HtmlEncode(Body)
        GoSub RODAPE
        Template2 = Template2 & Footer
      End If
    Next Contador
    '
    If Len(TBody) > 0 Then
      GoSub HEADER
      Body = Body & "<div class=" & Chr$(34) & "total" & Chr$(34) & ">Total Geral de registro(s) Listado(s): " & TotalRegistros & "</div>"
      Linha = 0
      TBody = ""
      Template2 = Template2 & basMain.HtmlEncode(Body)
      GoSub RODAPE
      Template2 = Template2 & Footer
    End If
    Template = Replace(Template, "[@PAGE]", Template2)
  End If
  '
  Call SalvarReport(ArquivoTemporario, Template)
  '----------------------------------------
  ' Fecha os arquivos
  '----------------------------------------
  Close #ArqOrigem
  Screen.MousePointer = vbDefault
  '
  Exit Sub
HEADER:
  NumPagina = NumPagina + 1
  Body = GetBody()
  Body = Replace(Body, "[@DATAEXTENSO]", basMain.DataPorExtenso(Date))
  Body = Replace(Body, "[@TITULO]", "RELATÓRIO CADASTRO PRODUTOS ORDENADO POR CÓDIGO")
  Body = Replace(Body, "[@REGINICIAL]", NumPagina)
  Body = Replace(Body, "[@SUBTOTALREG]", Linha)
  Body = Replace(Body, "[@TBODY]", TBody)
  Return
GETROW:
  Row = "<tr>"
  Row = Row & "<td>" & Trim$(Produto.Codigo) & "</td>"
  Row = Row & "<td>" & Trim$(Produto.Nome) & "</td>"
  Row = Row & "<td style=" & Chr$(34) & "text-align: right;" & Chr$(34) & ">" & Trim$(Produto.PrecoCompra) & "</td>"
  Row = Row & "<td style=" & Chr$(34) & "text-align: right;" & Chr$(34) & ">" & Trim$(Produto.PrecoVenda) & "</td>"
  Row = Row & "<td style=" & Chr$(34) & "text-align: right;" & Chr$(34) & ">" & Trim$(Produto.Quantidade) & "</td>"
  Row = Row & "<td style=" & Chr$(34) & "text-align: right;" & Chr$(34) & ">" & Trim$(Produto.Tabela) & "</td>"
  Row = Row & "<td style=" & Chr$(34) & "text-align: right;" & Chr$(34) & ">" & Trim$(Produto.TotalPago) & "</td>"
  Row = Row & "<td style=" & Chr$(34) & "text-align: right;" & Chr$(34) & ">" & Trim$(Produto.Ultimo) & "</td>"
  Row = Row & "</tr>"
  TBody = TBody & Row
  Return
RODAPE:
  Footer = "<div class=" & Chr$(34) & "footer" & Chr$(34) & ">"
  Footer = Footer & "<div>Programa Exemplo</div>"
  Footer = Footer & "<div class=" & Chr$(34) & "copyright" & Chr$(34) & ">"
  Footer = Footer & "&copy; Copyright Ipage Software - <a href=" & Chr$(34) & "https://www.ipage.com.br" & Chr$(34) & " target=" & Chr$(34) & "_blank" & Chr$(34) & ">www.ipage.com.br</a>"
  Footer = Footer & "</div>"
  Footer = Footer & "</div>"
  Footer = Footer & "</div>"
  Footer = Footer & "</div>"
  Return
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Close 'FECHA TODOS OS ARQUIVOS ABERTOS
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox "Erro ao excluir registro:" & vbCrLf & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Sub

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : GetBody
' Description: [type_description_here]
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 01:15:04
'
' Parameters :
'--------------------------------------------------------------------------------
Private Function GetBody() As String
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Html As String
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Html = "<div class=" & Chr$(34) & "page-break" & Chr$(34) & ">"
  Html = Html & "<div class=" & Chr$(34) & "page" & Chr$(34) & ">"
  Html = Html & "<div class=" & Chr$(34) & "top" & Chr$(34) & ">"
  Html = Html & "<div>"
  Html = Html & "<h1>CADASTRO PRODUTOS</h1>"
  Html = Html & "</div>"
  Html = Html & "<div class=" & Chr$(34) & "page-info" & Chr$(34) & ">"
  Html = Html & "<div>Página [@REGINICIAL]</div>"
  Html = Html & "<div>[@DATAEXTENSO]</div>"
  Html = Html & "</div>"
  Html = Html & "</div>"
  Html = Html & "<div class=" & Chr$(34) & "subtitle" & Chr$(34) & ">[@TITULO]</div>"
  Html = Html & "<table>"
  Html = Html & "<thead>"
  Html = Html & "<tr>"
  Html = Html & "<th style=" & Chr$(34) & "background-color: #e0e0e0;color: black;" & Chr$(34) & ">CÓDIGO</th>"
  Html = Html & "<th>NOME DO PRODUTO</th>"
  Html = Html & "<th>PREÇO COMPRA</th>"
  Html = Html & "<th>PREÇO VENDA</th>"
  Html = Html & "<th>QUANTIDADE</th>"
  Html = Html & "<th>TOTAL PAGO R$</th>"
  Html = Html & "<th>PREÇO TABELA</th>"
  Html = Html & "<th>ÚLTIMO PREÇO</th>"
  Html = Html & "</tr>"
  Html = Html & "</thead>"
  Html = Html & "<tbody>"
  Html = Html & "[@TBODY]"
  Html = Html & "</tbody>"
  Html = Html & "</table>"
  Html = Html & "<div class=" & Chr$(34) & "total" & Chr$(34) & ">Total de registro(s) Listado(s): [@SUBTOTALREG]</div>"
  GetBody = Html
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox "Ocorreu um erro ao executar o método GetBody." & vbCrLf & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : SalvarReport
' Description: Salva o relatório gerado na pasta temporário do Windows
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 01:15:09
'
' Parameters : ArquivoTemporario (String)
'              Html (String)
'--------------------------------------------------------------------------------
Public Function SalvarReport(ByVal ArquivoTemporario As String, _
                             ByVal Html As String) As Long
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Stream  As Object
  Dim Retorno As Long
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Set Stream = CreateObject("ADODB.Stream")
  '
  With Stream
    .Type = 2                 ' adTypeText
    .Charset = "windows-1252" 'utf-8
    .Open
    .WriteText Html
    .SaveToFile ArquivoTemporario, 2     ' 2 = adSaveCreateOverWrite
    .Close
  End With
  '
  Set Stream = Nothing
  '
  Retorno = ShellExecute(0, "open", ArquivoTemporario, vbNullString, vbNullString, 1)

  SalvarReport = (Retorno > 32)
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '-----------------------------------------------
  MsgBox "Ocorreu um erro ao executar o método SalvarReport." & vbCrLf & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function

