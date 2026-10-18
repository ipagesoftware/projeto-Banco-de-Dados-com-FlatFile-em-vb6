Attribute VB_Name = "basProduto"
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
' DEFINIÇÃO DA ESTRUTURA DOS CAMPOS DO ARQUIVO TEXTO
' QUE REPRESENTARÁ A TABELA PRODUTOS
'
Public Type TProduto
  Nome        As String * 20
  Codigo      As String * 5
  PrecoCompra As String * 15
  PrecoVenda  As String * 10
  Quantidade  As String * 10
  TotalPago   As String * 20
  Tabela      As String * 10
  Ultimo      As String * 10
End Type

' Declaração da estrutura (tipo de dado customizado)
Public Type Cliente
  Codigo As Long
  Nome As String * 50       ' String de tamanho fixo
  CPF As String * 11
  ValorLimite As Double
  Ativo As Boolean
End Type

'
Public dataBaseProduto As String

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : Duplicidade
' Description: Verifica se o nome do produto está cadastrado para o
'                    respectivo código.
'                    Retorna True se não houver duplicatas no banco de dados.
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 00:44:50
'
' Parameters : Codigo (String)
'              nomeProduto (String)
'--------------------------------------------------------------------------------
Public Function Duplicidade(ByVal Codigo As String, _
                            ByVal nomeProduto As String) As Boolean
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Buffer         As Integer
  Dim Produto        As TProduto
  Dim RecordSet      As Long
  Dim TotalRegistros As Long
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Duplicidade = True 'VALOR PADRÃO DoEven RETORNO DA FUNCÇÃO
  '
  Buffer = FreeFile
  nomeProduto = UCase(Trim(nomeProduto))
  ' VERIFICA SE O NOME DO ARQUIVO DE TEXTO PARA O BANCO DE DADOS FOI DEFINIDO
  If Len(dataBaseProduto) = 0 Then
    dataBaseProduto = App.Path & "\database\produtos.dat"
  End If
  '
  Open dataBaseProduto For Random As #Buffer Len = Len(Produto)

  TotalRegistros = LOF(Buffer) \ Len(Produto) 'ENCONTRO O NÚMERO TOTAL DE LINHAS
  ' INICIO O LOOP FOR
  For RecordSet = 1 To TotalRegistros
    Get #Buffer, RecordSet, Produto
    '
    If Val(Produto.Codigo) <> Val(Codigo) And UCase(Trim(Produto.Nome)) = nomeProduto Then
      Screen.MousePointer = vbDefault
      MsgBox "O Nome do Produto: " & nomeProduto & ",  já encontra-se cadastrado para o código: " & Produto.Codigo & "." & vbCrLf & vbCrLf & "Tente outro valor!", vbExclamation Or vbOKOnly, "Atenção"
      frmEdit.TxtNome.SetFocus
      Duplicidade = False
      Exit For
    End If
  Next RecordSet
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

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : Editar
' Description: Edita o valor do registro passado
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 01:02:57
'
' Parameters : RecordSet (Long)
'              Produto (TProduto)
'--------------------------------------------------------------------------------
Public Function Editar(ByVal RecordSet As Long, ByRef Produto As TProduto) As Boolean
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Arquivo        As Integer
  Dim TotalRegistros As Long
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Editar = False 'VALOR PADRÃO DO RETORNO DA FUNÇÃO
  ' VERIFICA SE O NOME DO ARQUIVO DE TEXTO PARA O BANCO DE DADOS FOI DEFINIDO
  If Len(dataBaseProduto) = 0 Then
    dataBaseProduto = App.Path & "\database\produtos.dat"
  End If
  '
  Arquivo = FreeFile
  '
  Open dataBaseProduto For Random As #Arquivo Len = Len(Produto)
  ' ENCONTRA A ÚLTIMA LINHA DO ARQUIVO TEXTO
  TotalRegistros = LOF(Arquivo) \ Len(Produto)
  '
  If RecordSet < 1 Or RecordSet > TotalRegistros Then
    Close #Arquivo
    MsgBox "Registro não encontrado!", vbCritical, "Atenção"
    Exit Function
  End If
  '
  Put #Arquivo, RecordSet, Produto
  '
  Close #Arquivo
  'RETORNA A ÚLTIMA LINHA ADICIONADA
  Editar = True
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Close 'FECHA TODOS OS ARQUIVOS ABERTOS
  Editar = False
  Screen.MousePointer = vbDefault
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox Err.Number & ", " & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : Excluir
' Description: Exclui o registro passado como parâmetro
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 01:04:13
'
' Parameters : RecordSet (Long)
'--------------------------------------------------------------------------------
Public Function Excluir(ByVal RecordSet As Long) As Boolean
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim ArqOrigem         As Integer
  Dim ArqTemp           As Integer
  Dim Produto           As TProduto
  Dim TotalRegistros    As Long
  Dim Contador          As Long
  Dim NovoRegistro      As Long
  Dim ArquivoTemporario As String
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  ' VERIFICA SE O NOME DO ARQUIVO DE TEXTO PARA O BANCO DE DADOS FOI DEFINIDO
  If Len(dataBaseProduto) = 0 Then
    dataBaseProduto = App.Path & "\database\produtos.dat"
  End If
  '----------------------------------------
  ' Nome do arquivo temporário
  '----------------------------------------
  ArquivoTemporario = CriarArquivoTemp("ib2")
  '----------------------------------------
  ' Abre o arquivo original
  '----------------------------------------
  ArqOrigem = FreeFile
  '
  Open dataBaseProduto For Random As #ArqOrigem Len = Len(Produto)
  '----------------------------------------
  ' Descobre quantidade de registros
  '----------------------------------------
  TotalRegistros = LOF(ArqOrigem) \ Len(Produto)
  '----------------------------------------
  ' Validação
  '----------------------------------------
  If RecordSet < 1 Or RecordSet > TotalRegistros Then
    Close #ArqOrigem
    MsgBox "Registro inexistente.", vbExclamation, "Atenção"
    Exit Function
  End If
  '----------------------------------------
  ' Cria arquivo temporário
  '----------------------------------------
  ArqTemp = FreeFile
  '
  Open ArquivoTemporario For Random As #ArqTemp Len = Len(Produto)
  '----------------------------------------
  ' Copia todos os registros,
  ' exceto o registro a excluir
  '----------------------------------------
  NovoRegistro = 1
  '
  For Contador = 1 To TotalRegistros
    If Contador <> RecordSet Then
      Get #ArqOrigem, Contador, Produto
      Put #ArqTemp, NovoRegistro, Produto
      NovoRegistro = NovoRegistro + 1
    End If
  Next Contador
  '----------------------------------------
  ' Fecha os arquivos
  '----------------------------------------
  Close #ArqOrigem
  Close #ArqTemp
  '----------------------------------------
  ' Exclui o arquivo original
  '----------------------------------------
  Kill dataBaseProduto
    
  '----------------------------------------
  ' Renomeia o temporário
  '----------------------------------------
  Name ArquivoTemporario As dataBaseProduto
  Excluir = True
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Close 'FECHA TODOS OS ARQUIVOS ABERTOS
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '------------------------------------
  MsgBox "Erro ao excluir registro:" & vbCrLf & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
  Excluir = False
End Function

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : Incluir
' Description: Adiciona um novo registro na base de dados
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 00:57:49
'
' Parameters : Produto (TProduto)
'--------------------------------------------------------------------------------
Public Function Incluir(ByRef Produto As TProduto) As Long
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Arquivo      As Integer
  Dim NovoRegistro As Long
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Produto.Codigo = Format(basProduto.UltimoRegistro, "00000")
  '
  Incluir = 0 'VALOR PADRÃO DoEvent RETORNO DA FUNÇÃO
  ' VERIFICA SE O NOME DO ARQUIVO DE TEXTO PARA O BANCO DE DADOS FOI DEFINIDO
  If Len(dataBaseProduto) = 0 Then
    dataBaseProduto = App.Path & "\database\produtos.dat"
  End If
  '
  Arquivo = FreeFile
  '
  Open dataBaseProduto For Random As #Arquivo Len = Len(Produto)
  ' ENCONTRA A ÚLTIMA LINHA DO ARQUIVO TEXTO
  NovoRegistro = (LOF(Arquivo) \ Len(Produto)) + 1
  '
  Put #Arquivo, NovoRegistro, Produto
  '
  Close #Arquivo
  'RETORNA A ÚLTIMA LINHA ADICIONADA + 1
  Incluir = NovoRegistro + 1
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Close 'FECHA TODOS OS ARQUIVOS ABERTOS
  Screen.MousePointer = vbDefault
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '-----------------------------------------------
  MsgBox Err.Number & ", " & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : Pesquisar
' Description: Verifica se o nome do produto já está cadstrado.
'                    Retorna True se o nome do produto ainda não foi cadastrado
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 00:40:57
'
' Parameters : nomeProduto (String)
'--------------------------------------------------------------------------------
Public Function Pesquisar(ByVal nomeProduto As String) As Boolean
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Buffer         As Integer
  Dim Produto        As TProduto
  Dim RecordSet      As Long
  Dim TotalRegistros As Long
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  Pesquisar = True 'VALOR PADRÃO DE RETORNO DA FUNCÇÃO
  ' VERIFICA SE O NOME DO ARQUIVO DE TEXTO PARA O BANCO DE DADOS FOI DEFINIDO
  If Len(dataBaseProduto) = 0 Then
    dataBaseProduto = App.Path & "\database\produtos.dat"
  End If
  '
  Buffer = FreeFile
  '
  Open dataBaseProduto For Random As #Buffer Len = Len(Produto)

  TotalRegistros = LOF(Buffer) \ Len(Produto) 'ENCONTRO O NÚMERO TOTAL DE LINHAS
  ' INICIO O LOOP FOR
  For RecordSet = 1 To TotalRegistros
    Get #Buffer, RecordSet, Produto
    '
    If UCase(Trim(Produto.Nome)) = UCase(Trim(nomeProduto)) Then
      Screen.MousePointer = vbDefault
      MsgBox "O Nome do Produto: " & nomeProduto & ",  já encontra-se cadastrado para o código: " & Produto.Codigo & "." & vbCrLf & vbCrLf & "Tente outro valor!", vbExclamation Or vbOKOnly, "Atenção"
      frmInsert.TxtNome.SetFocus
      Pesquisar = False
      Exit Function
    End If
        
  Next RecordSet

  Close #Buffer
  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Close 'FECHA TODOS OS ARQUIVOS ABERTOS
  Screen.MousePointer = vbDefault
  '------------------------------------
  ' EXIBE MENSAGEM AMIGÁVEL AO USUÁRIO
  '-----------------------------------------------
  MsgBox Err.Number & ", " & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function

'--------------------------------------------------------------------------------
' Project    : FlatFile
' Procedure  : UltimoRegistro
' Description: Faz uma varredura no banco de dados e retorna o maior
'              código do produto + 1
' Created by : Diógenes Dias
' Machine    : DESKTOP-3M0KOU8
' Date-Time  : 05/10/2026 00:53:28
'
' Parameters :
'--------------------------------------------------------------------------------
Public Function UltimoRegistro() As Long
  '---------------------
  ' DEFINIÇÃO VARIÁVEIS
  '---------------------
  Dim Buffer         As Integer
  Dim RecordSet      As Long
  Dim TotalRegistros As Long
  Dim Maior          As Long
  Dim Produto        As TProduto
  '
  On Error GoTo LogErrorHandler 'ATIVA O TRATAMENTO DE ERROS
  '
  UltimoRegistro = 0 'VALOR PADRÃO
  Maior = 0
  ' VERIFICA SE O NOME DO ARQUIVO DE TEXTO PARA O BANCO DE DADOS FOI DEFINIDO
  If Len(dataBaseProduto) = 0 Then
    dataBaseProduto = App.Path & "\database\produtos.dat"
  End If
  '
  Buffer = FreeFile
  '
  Open dataBaseProduto For Random As #Buffer Len = Len(Produto)
  TotalRegistros = LOF(Buffer) \ Len(Produto) 'ENCONTRO O NÚMERO TOTAL DE LINHAS

  ' INICIO O LOOP FOR
  For RecordSet = 1 To TotalRegistros
    Get #Buffer, RecordSet, Produto

    ' VAL CONVERTE A STRING PARA NÚMERO
    If Val(Produto.Codigo) > Maior Then
      Maior = Produto.Codigo
    End If

  Next RecordSet
  '
  Close #Buffer
  '
  UltimoRegistro = Maior + 1

  Exit Function
  'TRATAMENTO DE ERROS
LogErrorHandler:
  On Error Resume Next 'IGNORA O TRATAMENTO DE ERROS DO VB.
  Close 'FECHA TODOS OS ARQUIVOS ABERTOS
  Screen.MousePointer = vbDefault
  MsgBox Err.Number & ", " & Err.Description, vbExclamation Or vbOKOnly, "Atenção"
End Function

