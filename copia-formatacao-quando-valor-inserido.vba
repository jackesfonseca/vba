Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Nome da Sua Planilha") ' Substitua pelo nome da sua planilha

    ' Verifique se a mudança ocorreu nas colunas E a U e se apenas uma célula foi alterada
    If Not Intersect(Target, ws.Range("E:U")) Is Nothing And Target.Cells.Count = 1 Then
        Dim rowAbove As Range
        Set rowAbove = ws.Rows(Target.Row - 1) ' Linha acima da célula alterada

        ' Copie a formatação da linha acima para a nova linha
        rowAbove.Copy
        Target.PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If
End Sub