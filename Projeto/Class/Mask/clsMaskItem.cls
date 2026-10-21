VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaskItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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
'
'
'==============================================================================
' clsMaskItem - uma instancia por TextBox (criada pela clsMaskedEdit).
'
' Caracteres da mascara:
'   9 ou #  = digito (0-9)
'   A       = letra (A-Z, convertida para maiuscula)
'   X       = letra ou digito (letra convertida para maiuscula)
'   *       = qualquer caractere
'   outros  = literais (inseridos automaticamente)
'
' Mascaras alternativas: separe com "|" em ordem CRESCENTE de tamanho.
'   Ex.: "(99) 9999-9999|(99) 99999-9999"
'==============================================================================

Private Const MAXDIG As Long = 15          ' maximo de digitos no formato MOEDA

Private WithEvents mTxt As VB.TextBox
Attribute mTxt.VB_VarHelpID = -1

Private mMasks() As String                 ' mascaras alternativas
Private mCount As Long
Private mKind As Long                      ' 0 = mascara, 1 = numero, 2 = moeda
Private mDecimals As Long
Private mValidation As String              ' "", DATA, DATAHORA, HORA, CPF, CNPJ, CPFCNPJ
Private mClear As Boolean
Private mBusy As Boolean

'------------------------------------------------------------------------------
' Inicializacao
'------------------------------------------------------------------------------
Public Sub Init(ByVal txt As Object, ByVal Spec As String, ByVal ClearOnInvalid As Boolean, _
                Optional ByVal Decimals As Long = -1)
    Dim s As String, mask As String, p As Long, i As Long

    Set mTxt = txt
    mClear = ClearOnInvalid
    mKind = 0
    mValidation = ""
    s = UCase$(Trim$(Spec))

    Select Case True
        Case s = "CPF":      mask = "999.999.999-99": mValidation = "CPF"
        Case s = "CNPJ":     mask = "99.999.999/9999-99": mValidation = "CNPJ"
        Case s = "CPFCNPJ":  mask = "999.999.999-99|99.999.999/9999-99": mValidation = "CPFCNPJ"
        Case s = "CEP":      mask = "99999-999"
        Case s = "TELEFONE": mask = "(99) 9999-9999|(99) 99999-9999"
        Case s = "DATA":     mask = "99/99/9999": mValidation = "DATA"
        Case s = "DATAHORA": mask = "99/99/9999 99:99": mValidation = "DATAHORA"
        Case s = "HORA":     mask = "99:99": mValidation = "HORA"
        Case s = "NUMERO":   mKind = 1
        Case Left$(s, 5) = "MOEDA"
            mKind = 2
            mDecimals = 2
            p = InStr(s, ":")
            If p > 0 Then mDecimals = Val(Mid$(s, p + 1))
            If Decimals >= 0 Then mDecimals = Decimals     ' 3o parametro tem prioridade
            If mDecimals < 0 Then mDecimals = 0
            If mDecimals > 6 Then mDecimals = 6
        Case Else:           mask = Spec          ' mascara livre
    End Select

    mMasks = Split(mask, "|")
    mCount = UBound(mMasks) + 1

    If mKind = 0 Then
        ' MaxLength = tamanho da maior mascara
        p = 0
        For i = 0 To mCount - 1
            If Len(mMasks(i)) > p Then p = Len(mMasks(i))
        Next
        mTxt.MaxLength = p
    End If

    mTxt_Change                                   ' formata o texto ja existente
End Sub

'------------------------------------------------------------------------------
' Eventos do TextBox
'------------------------------------------------------------------------------
Private Sub mTxt_KeyPress(KeyAscii As Integer)
    Dim ch As String, txt As String, st As Long
    Dim mask As String, restante As String, t As String

    If KeyAscii < 32 Then Exit Sub                ' Backspace, Enter, Tab, Ctrl+C/V/X...

    ch = Chr$(KeyAscii)
    txt = mTxt.Text
    st = mTxt.SelStart

    Select Case mKind
        Case 1                                    ' NUMERO
            If Not IsDigit(ch) Then KeyAscii = 0

        Case 2                                    ' MOEDA
            If Not IsDigit(ch) Then
                KeyAscii = 0
            ElseIf mTxt.SelLength = 0 And Len(DigitsOnly(txt)) >= MAXDIG Then
                KeyAscii = 0
            End If

        Case Else                                 ' MASCARA
            mask = LastMask()
            ' o que sobra depois que a selecao (se houver) for substituida
            restante = Strip(Left$(txt, st) & Mid$(txt, st + mTxt.SelLength + 1), mask)
            If Len(restante) >= SlotCount(mask) Then
                KeyAscii = 0
            Else
                t = SlotTypeAt(mask, Len(Strip(Left$(txt, st), mask)) + 1)
                If Accepts(t, ch) Then
                    If t = "A" Or t = "X" Then KeyAscii = Asc(UCase$(ch))
                Else
                    KeyAscii = 0
                End If
            End If
    End Select
End Sub

Private Sub mTxt_Change()
    Dim txt As String, novo As String, Raw As String, mk As String
    Dim st As Long, antes As Long

    If mBusy Then Exit Sub
    mBusy = True
    On Error GoTo Fim

    txt = mTxt.Text
    st = mTxt.SelStart

    Select Case mKind
        Case 0
            Raw = Strip(txt, LastMask())
            antes = Len(Strip(Left$(txt, st), LastMask()))
            mk = PickMask(Len(Raw))
            novo = Apply(Raw, mk)
            If novo <> txt Then
                mTxt.Text = novo
                mTxt.SelStart = CursorPos(mk, Len(novo), antes, Len(Raw))
            End If

        Case 1
            novo = DigitsOnly(txt)
            If novo <> txt Then
                antes = Len(DigitsOnly(Left$(txt, st)))
                mTxt.Text = novo
                mTxt.SelStart = antes
            End If

        Case 2
            novo = FormatMoney(txt)
            If novo <> txt Then mTxt.Text = novo
            mTxt.SelStart = Len(mTxt.Text)
    End Select
Fim:
    mBusy = False
End Sub

Private Sub mTxt_LostFocus()
    If mClear Then
        If mTxt.Text <> "" Then
            If Not IsValid() Then mTxt.Text = ""
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' Publicos
'------------------------------------------------------------------------------
' Valor sem mascara: texto "cru" (CPF -> so digitos), numero (MOEDA -> Double)
Public Function Value() As Variant
    Select Case mKind
        Case 0: Value = Strip(mTxt.Text, LastMask())
        Case 1: Value = mTxt.Text
        Case 2: Value = Val(DigitsOnly(mTxt.Text)) / 10 ^ mDecimals
    End Select
End Function

' Atribui um valor ja formatando (para MOEDA passe um numero: 1234.5)
Public Sub SetValue(ByVal v As Variant)
    Dim d As String
    If mKind = 2 Then
        d = CStr(Int(CDbl(v) * 10 ^ mDecimals + 0.5))
        If Len(d) < mDecimals + 1 Then d = String$(mDecimals + 1 - Len(d), "0") & d
        mTxt.Text = d
    Else
        mTxt.Text = CStr(v)
    End If
End Sub

' Vazio e considerado valido (trate campo obrigatorio no seu codigo)
Public Function IsValid() As Boolean
    Dim t As String, Raw As String

    t = mTxt.Text
    If t = "" Or mKind <> 0 Then IsValid = True: Exit Function

    Raw = Strip(t, LastMask())
    If Len(Raw) <> SlotCount(PickMask(Len(Raw))) Then Exit Function    ' incompleto

    Select Case mValidation
        Case "DATA"
            IsValid = ValidDate(CLng(Mid$(Raw, 1, 2)), CLng(Mid$(Raw, 3, 2)), CLng(Mid$(Raw, 5, 4)))
        Case "DATAHORA"
            IsValid = ValidDate(CLng(Mid$(Raw, 1, 2)), CLng(Mid$(Raw, 3, 2)), CLng(Mid$(Raw, 5, 4))) _
                  And ValidTime(CLng(Mid$(Raw, 9, 2)), CLng(Mid$(Raw, 11, 2)))
        Case "HORA"
            IsValid = ValidTime(CLng(Mid$(Raw, 1, 2)), CLng(Mid$(Raw, 3, 2)))
        Case "CPF"
            IsValid = ValidCPF(Raw)
        Case "CNPJ"
            IsValid = ValidCNPJ(Raw)
        Case "CPFCNPJ"
            If Len(Raw) = 11 Then IsValid = ValidCPF(Raw) Else IsValid = ValidCNPJ(Raw)
        Case Else
            IsValid = True
    End Select
End Function

'------------------------------------------------------------------------------
' Rotinas de mascara
'------------------------------------------------------------------------------
Private Function LastMask() As String
    LastMask = mMasks(mCount - 1)
End Function

Private Function IsSlot(ByVal c As String) As Boolean
    IsSlot = (InStr("9#AX*", c) > 0)
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    IsDigit = (ch >= "0" And ch <= "9")
End Function

Private Function SlotCount(ByVal mask As String) As Long
    Dim i As Long, n As Long
    For i = 1 To Len(mask)
        If IsSlot(Mid$(mask, i, 1)) Then n = n + 1
    Next
    SlotCount = n
End Function

' Tipo do n-esimo "slot" (posicao editavel) da mascara
Private Function SlotTypeAt(ByVal mask As String, ByVal n As Long) As String
    Dim i As Long, k As Long, c As String
    For i = 1 To Len(mask)
        c = Mid$(mask, i, 1)
        If IsSlot(c) Then
            k = k + 1
            If k = n Then SlotTypeAt = c: Exit Function
        End If
    Next
End Function

Private Function Accepts(ByVal SlotType As String, ByVal ch As String) As Boolean
    Dim u As String
    u = UCase$(ch)
    Select Case SlotType
        Case "9", "#": Accepts = (ch >= "0" And ch <= "9")
        Case "A":      Accepts = (u >= "A" And u <= "Z")
        Case "X":      Accepts = (ch >= "0" And ch <= "9") Or (u >= "A" And u <= "Z")
        Case "*":      Accepts = (Asc(ch) >= 32)
    End Select
End Function

' Remove os literais: devolve so os caracteres que "cabem" nos slots, em ordem
Private Function Strip(ByVal Text As String, ByVal mask As String) As String
    Dim i As Long, n As Long, ch As String, t As String, r As String
    n = 1
    For i = 1 To Len(Text)
        ch = Mid$(Text, i, 1)
        t = SlotTypeAt(mask, n)
        If t = "" Then Exit For
        If Accepts(t, ch) Then
            If t = "A" Or t = "X" Then ch = UCase$(ch)
            r = r & ch
            n = n + 1
        End If
    Next
    Strip = r
End Function

' Aplica a mascara; literais finais so aparecem quando ha o proximo caractere
' (assim o Backspace nunca "trava" em um literal)
Private Function Apply(ByVal Raw As String, ByVal mask As String) As String
    Dim i As Long, k As Long, c As String, r As String
    k = 1
    For i = 1 To Len(mask)
        If k > Len(Raw) Then Exit For
        c = Mid$(mask, i, 1)
        If IsSlot(c) Then
            r = r & Mid$(Raw, k, 1)
            k = k + 1
        Else
            r = r & c
        End If
    Next
    Apply = r
End Function

' Escolhe a primeira mascara alternativa que comporta a quantidade digitada
Private Function PickMask(ByVal RawLen As Long) As String
    Dim i As Long
    For i = 0 To mCount - 1
        If SlotCount(mMasks(i)) >= RawLen Then PickMask = mMasks(i): Exit Function
    Next
    PickMask = LastMask()
End Function

' Posicao do cursor (SelStart) logo apos o RawBefore-esimo caractere digitado
Private Function CursorPos(ByVal mask As String, ByVal NewLen As Long, _
                           ByVal RawBefore As Long, ByVal RawTotal As Long) As Long
    Dim i As Long, k As Long
    If RawBefore >= RawTotal Then CursorPos = NewLen: Exit Function
    For i = 1 To Len(mask)
        If IsSlot(Mid$(mask, i, 1)) Then
            k = k + 1
            If k = RawBefore Then CursorPos = i: Exit Function
        End If
    Next
    CursorPos = 0
End Function

'------------------------------------------------------------------------------
' Numero / Moeda
'------------------------------------------------------------------------------
Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(s)
        If IsDigit(Mid$(s, i, 1)) Then r = r & Mid$(s, i, 1)
    Next
    DigitsOnly = r
End Function

' Digitacao "da direita para a esquerda" (como calculadora): 5 -> 0,05 ; 1234567 -> 12.345,67
Private Function FormatMoney(ByVal Text As String) As String
    Dim d As String, ip As String, r As String

    d = DigitsOnly(Text)
    If Len(d) > MAXDIG Then d = Left$(d, MAXDIG)

    Do While Len(d) > mDecimals + 1 And Left$(d, 1) = "0"
        d = Mid$(d, 2)
    Loop

    If Len(d) < mDecimals + 1 Then
        If Len(Replace(d, "0", "")) = 0 Then Exit Function      ' vazio ou so zeros -> limpa
        d = String$(mDecimals + 1 - Len(d), "0") & d
    End If

    ip = Left$(d, Len(d) - mDecimals)
    Do While Len(ip) > 3
        r = "." & Right$(ip, 3) & r
        ip = Left$(ip, Len(ip) - 3)
    Loop
    r = ip & r
    If mDecimals > 0 Then r = r & "," & Right$(d, mDecimals)
    FormatMoney = r
End Function

'------------------------------------------------------------------------------
' Validacoes
'------------------------------------------------------------------------------
Private Function ValidDate(ByVal d As Long, ByVal m As Long, ByVal y As Long) As Boolean
    If y < 1900 Or m < 1 Or m > 12 Or d < 1 Then Exit Function
    ValidDate = (Day(DateSerial(y, m, d)) = d)    ' dia inexistente "rola" p/ o mes seguinte
End Function

Private Function ValidTime(ByVal h As Long, ByVal mi As Long) As Boolean
    ValidTime = (h <= 23 And mi <= 59)
End Function

Private Function ValidCPF(ByVal s As String) As Boolean
    Dim i As Long, soma As Long, dv As Long
    If Len(s) <> 11 Then Exit Function
    If s = String$(11, Mid$(s, 1, 1)) Then Exit Function

    For i = 1 To 9: soma = soma + CLng(Mid$(s, i, 1)) * (11 - i): Next
    dv = (soma * 10) Mod 11: If dv = 10 Then dv = 0
    If dv <> CLng(Mid$(s, 10, 1)) Then Exit Function

    soma = 0
    For i = 1 To 10: soma = soma + CLng(Mid$(s, i, 1)) * (12 - i): Next
    dv = (soma * 10) Mod 11: If dv = 10 Then dv = 0
    ValidCPF = (dv = CLng(Mid$(s, 11, 1)))
End Function

Private Function ValidCNPJ(ByVal s As String) As Boolean
    Dim p1 As Variant, p2 As Variant, i As Long, soma As Long, dv As Long
    If Len(s) <> 14 Then Exit Function
    If s = String$(14, Mid$(s, 1, 1)) Then Exit Function

    p1 = Array(5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    p2 = Array(6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)

    For i = 0 To 11: soma = soma + CLng(Mid$(s, i + 1, 1)) * p1(i): Next
    dv = soma Mod 11: If dv < 2 Then dv = 0 Else dv = 11 - dv
    If dv <> CLng(Mid$(s, 13, 1)) Then Exit Function

    soma = 0
    For i = 0 To 12: soma = soma + CLng(Mid$(s, i + 1, 1)) * p2(i): Next
    dv = soma Mod 11: If dv < 2 Then dv = 0 Else dv = 11 - dv
    ValidCNPJ = (dv = CLng(Mid$(s, 14, 1)))
End Function

