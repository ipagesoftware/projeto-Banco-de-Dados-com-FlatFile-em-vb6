VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaskedEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'Parênteses: chame Format sem parênteses, como acima. Com parênteses e dois
'argumentos o VB6 dá erro de sintaxe, a não ser que use Call.
'Máscara: 9 ou # é dígito, A é letra (vira maiúscula), X é letra ou dígito,
'* é qualquer caractere.
'
'O resto é literal e entra automaticamente. As máscaras alternativas separadas
'por | devem estar em ordem crescente de tamanho.
'
'Validação: IsValid confere se a máscara está completa e valida data e hora
'(ano mínimo 1900), CPF e CNPJ pelos dígitos verificadores. Campo vazio conta
'como válido, então trate obrigatoriedade no seu código.
'Se quiser que o campo seja apagado ao perder o foco quando inválido, use
'MaskedEdit.ClearOnInvalid = True antes dos Format.
'
'Edição no meio do texto: o cursor é preservado. Passar por cima de um literal
'exige dois Backspaces, de propósito, para o literal não “travar”.
'Valores: MaskedEdit.SetValue txtValor, 1234.5 formata o campo, e Value devolve
'Double para MOEDA.
'Limitações: o MOEDA usa vírgula e ponto fixos (padrão brasileiro) e não aceita
'valor negativo.
'Não testei com controles em array (txt(0), txt(1)...).
'

'MaskedEdit.Format TxtPrecoCompra, "MOEDA", 4    ' 1.234,5678
'MaskedEdit.Format TxtPrecoVenda, "MOEDA", 2     ' 1.234,56
'MaskedEdit.Format txtTotal, "MOEDA"             ' omitido = 2 casas
'MaskedEdit.Format txtQtde, "MOEDA", 0           ' sem decimais: 1.234


'==============================================================================
' clsMaskedEdit - gerenciador de mascaras para TextBox (sem usar o MaskEdBox).
' Use a instancia global declarada em modMask.bas:  MaskedEdit
'
'   MaskedEdit.Format txtCPF, "CPF"
'   MaskedEdit.Format txtX, "999.999.999-99"
'
' Formatos prontos: CPF, CNPJ, CPFCNPJ, CEP, TELEFONE, DATA, DATAHORA, HORA,
'                   NUMERO, MOEDA, MOEDA:3 (casas decimais)
'==============================================================================

Private mItems As Collection

' True = ao sair do campo (LostFocus) apaga o conteudo se estiver invalido/incompleto.
' Defina ANTES de chamar Format.
Public ClearOnInvalid As Boolean

Private Sub Class_Initialize()
    Set mItems = New Collection
    ClearOnInvalid = False
End Sub

Private Sub Class_Terminate()
    Set mItems = Nothing
End Sub

' Aplica a mascara/formato ao TextBox (chame no Form_Load)
' Decimals (so para "MOEDA"): numero de casas decimais; omitido = 2
Public Sub Format$(ByVal txt As Object, ByVal Spec As String, Optional ByVal Decimals As Long = -1)
    Dim it As clsMaskItem
    Release txt                                   ' se ja existia, substitui
    Set it = New clsMaskItem
    it.Init txt, Spec, ClearOnInvalid, Decimals
    mItems.Add it, KeyOf(txt)
End Sub

' Remove o tratamento de um TextBox
Public Sub Release(ByVal txt As Object)
    On Error Resume Next
    mItems.Remove KeyOf(txt)
End Sub

' Remove o tratamento de todos os TextBox de um form (chame no Form_Unload)
Public Sub ReleaseForm(ByVal frm As Object)
    Dim c As Object
    For Each c In frm.Controls
        If TypeOf c Is VB.TextBox Then Release c
    Next
End Sub

' True se o conteudo e valido (campo vazio = True)
Public Function IsValid(ByVal txt As Object) As Boolean
    Dim it As clsMaskItem
    Set it = GetItem(txt)
    If it Is Nothing Then IsValid = True Else IsValid = it.IsValid
End Function

' Conteudo sem mascara (CPF -> "12345678901"; MOEDA -> Double)
Public Function Value(ByVal txt As Object) As Variant
    Dim it As clsMaskItem
    Set it = GetItem(txt)
    If it Is Nothing Then Value = txt.Text Else Value = it.Value
End Function

' Atribui valor ja formatado (MOEDA: passe um numero, ex. 1234.5)
Public Sub SetValue(ByVal txt As Object, ByVal v As Variant)
    Dim it As clsMaskItem
    Set it = GetItem(txt)
    If it Is Nothing Then txt.Text = CStr(v) Else it.SetValue v
End Sub

Private Function GetItem(ByVal txt As Object) As clsMaskItem
    On Error Resume Next
    Set GetItem = mItems(KeyOf(txt))
End Function

Private Function KeyOf(ByVal txt As Object) As String
    KeyOf = "k" & Hex$(txt.hwnd)
End Function '
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'> Nome                  : SelText                                        <
'> Data/Hora             : 29/10/2006 21:27                               <
'> Autor                 : Diógenes Dias                                  <
'> Descrição             : Seleciona o txt quando entra em foco em uma <
'>                         caixa de controle.                             <
'> Parâmetros Passados   : O objeto txt, a posição inicial de seleção  <
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
Public Sub SelText(ByVal txt As Object, _
                   Optional ByVal p_Start As Integer = 0, _
                   Optional ByVal p_MaxLength As Integer = 255)

  '
  txt.SelStart = p_Start
  txt.SelLength = p_MaxLength
  '
End Sub

