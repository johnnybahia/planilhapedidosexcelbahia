Attribute VB_Name = "Módulo1"
Sub apagar_linhas_colunas()
    Dim linha As Long
    Dim coluna As Long
    linha = 4
    
    ' Verifica cada linha na planilha a partir da linha 4
    Do While Cells(linha, 1) <> ""
        ' Verifica se a linha contém dados
        If Application.WorksheetFunction.CountA(Rows(linha)) > 0 Then
            ' Apaga a linha
            Rows(linha).Delete
            ' Como a linha foi excluída, precisamos continuar verificando a mesma linha novamente
            linha = linha - 1
        End If
        linha = linha + 1
    Loop
    
    coluna = 1
    
    ' Verifica cada coluna na planilha a partir da coluna A
    Do While Cells(4, coluna) <> ""
        ' Verifica se a coluna contém dados
        If Application.WorksheetFunction.CountA(Columns(coluna)) > 0 Then
            ' Apaga a coluna
            Columns(coluna).Delete
            ' Como a coluna foi excluída, precisamos continuar verificando a mesma coluna novamente
            coluna = coluna - 1
        End If
        coluna = coluna + 1
    Loop
End Sub



