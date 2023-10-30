Public Function lfValidaCPF(ByVal lNumCPF As String) As Boolean
    Application.Volatile
    
    Dim lMultiplicador  As Integer
    Dim lDv1            As Integer
    Dim lDv2            As Integer
    
    lMultiplicador = 2
    
    'Realiza o preenchimento dos zeros á esquerda
    lNumCPF = String(11 - Len(lNumCPF), "0") & lNumCPF
    
    'Realiza o cálculo do dividendo para o dv1 e o dv2
    For i = 9 To 1 Step -1
        lDv1 = (Mid(lNumCPF, i, 1) * lMultiplicador) + lDv1
        
        lDv2 = (Mid(lNumCPF, i, 1) * (lMultiplicador + 1)) + lDv2
        
        lMultiplicador = lMultiplicador + 1
    Next
    
    'Realiza o cálculo para chegar no primeiro dígio
    lDv1 = lDv1 Mod 11
    
    If lDv1 >= 2 Then
        lDv1 = 11 - lDv1
    Else
        lDv1 = 0
    End If
    
    'Realiza o cálculo para chegar no segundo dígido
    lDv2 = lDv2 + (lDv1 * 2)
    
    lDv2 = lDv2 Mod 11
    
    If lDv2 >= 2 Then
        lDv2 = 11 - lDv2
    Else
        lDv2 = 0
    End If
    
    'Realiza a validação e retorna na função
    If Right(lNumCPF, 2) = CStr(lDv1) & CStr(lDv2) Then
        lfValidaCPF = True
    Else
        lfValidaCPF = False
    End If
End Function