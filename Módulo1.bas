Attribute VB_Name = "M�dulo1"
Sub apagar_linhas_colunas()
    Dim linha As Long
    Dim coluna As Long
    linha = 4

    ' Apaga toda linha com algum dado a partir da linha 4 (checa a linha
    ' inteira, nao so a coluna A, para nao deixar linha orfa para tras)
    Do While Application.WorksheetFunction.CountA(Rows(linha)) > 0
        Rows(linha).Delete
    Loop

    coluna = 1

    ' Apaga toda coluna com algum dado a partir da coluna A (checa a coluna
    ' inteira, nao so a linha 4, pelo mesmo motivo)
    Do While Application.WorksheetFunction.CountA(Columns(coluna)) > 0
        Columns(coluna).Delete
    Loop
End Sub



