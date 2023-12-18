Public Function verificaCNPJ(sCNPJ As String) As Boolean
    Dim d1 As Integer
    Dim d2 As Integer
    Dim d3 As Integer
    Dim d4 As Integer
    Dim d5 As Integer
    Dim d6 As Integer
    Dim d7 As Integer
    Dim d8 As Integer
    Dim d9 As Integer
    Dim d10 As Integer
    Dim d11 As Integer
    Dim d12 As Integer
    Dim d13 As Integer
    Dim d14 As Integer
    Dim DV1 As Integer
    Dim DV2 As Integer
    Dim UltDig As Integer
    'Completa com zeros à esquerda caso não esteja com os 14 digitos
    If Len(sCNPJ) < 14 Then
        sCNPJ = String(14 - Len(sCNPJ), "0") & sCNPJ
    End If
    'Pega a posição do último dígito
    UltDig = Len(sCNPJ)
    'Sai da função caso a célula esteja vazia
    If sCNPJ = "00000000000000" Then
        verificaCNPJ = ""
        Exit Function
    End If
    'Pega cada dígito do CNPJ informado e
    'coloca nas variáveis específicas
    d1 = CInt(Mid(sCNPJ, UltDig - 13, 1))
    d2 = CInt(Mid(sCNPJ, UltDig - 12, 1))
    d3 = CInt(Mid(sCNPJ, UltDig - 11, 1))
    d4 = CInt(Mid(sCNPJ, UltDig - 10, 1))
    d5 = CInt(Mid(sCNPJ, UltDig - 9, 1))
    d6 = CInt(Mid(sCNPJ, UltDig - 8, 1))
    d7 = CInt(Mid(sCNPJ, UltDig - 7, 1))
    d8 = CInt(Mid(sCNPJ, UltDig - 6, 1))
    d9 = CInt(Mid(sCNPJ, UltDig - 5, 1))
    d10 = CInt(Mid(sCNPJ, UltDig - 4, 1))
    d11 = CInt(Mid(sCNPJ, UltDig - 3, 1))
    d12 = CInt(Mid(sCNPJ, UltDig - 2, 1))
    d13 = CInt(Mid(sCNPJ, UltDig - 1, 1))    '<----- Aqui são os DVs informados
    d14 = CInt(Mid(sCNPJ, UltDig, 1))    '<----- no CNPJ analizado
    '----------- Aqui é executado o calculo para obter os digitos verificadores corretos
    DV1 = (d1 * 6) + (d2 * 7) + (d3 * 8) + (d4 * 9) + (d5 * 2) + (d6 * 3) + (d7 * 4) + (d8 * 5) + (d9 * 6) + (d10 * 7) + (d11 * 8) + (d12 * 9)
    DV1 = DV1 Mod 11    'Obtem o resto
    'se o resto for igual a 10 altera pra 0
    If DV1 = 10 Then
        DV1 = 0
    End If
    DV2 = (d1 * 5) + (d2 * 6) + (d3 * 7) + (d4 * 8) + (d5 * 9) + (d6 * 2) + (d7 * 3) + (d8 * 4) + (d9 * 5) + (d10 * 6) + (d11 * 7) + (d12 * 8) + (DV1 * 9)
    DV2 = DV2 Mod 11    'Obtem o resto
    'se o resto for igual a 10 altera pra 0
    If DV2 = 10 Then
        DV2 = 0
    End If
    '---------- Fazendo a comparação dos dvs informados -------
    If d13 = DV1 And d14 = DV2 Then
        verificaCNPJ = True
    Else
        verificaCNPJ = False
    End If
End Function
