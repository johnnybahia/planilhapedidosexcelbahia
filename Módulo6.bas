Attribute VB_Name = "Módulo6"
Sub VerificarItens()
    Dim ws As Worksheet
    Dim lastRowA As Long, lastRowB As Long, lastRowC As Long
    Dim rngA As Range, rngB As Range, rngC As Range
    Dim cellA As Range, cellB As Range, foundRange As Range
    Dim uniqueItems As Collection
    Dim item As Variant
    
    ' Definir a planilha de trabalho
    Set ws = ThisWorkbook.Sheets("VERIFICAÇÃO DE PEDIDOS") ' Substitua "NomeDaPlanilha" pelo nome da sua planilha
    
    ' Obter a última linha em cada coluna
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    lastRowC = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Definir os intervalos de dados nas colunas A, B e C
    Set rngA = ws.Range("A2:A" & lastRowA)
    Set rngB = ws.Range("B2:B" & lastRowB)
    Set rngC = ws.Range("C2:C" & lastRowC)
    
    ' Inicializar a coleção de itens únicos
    Set uniqueItems = New Collection
    
    ' Verificar cada item na coluna B
    For Each cellB In rngB
        ' Verificar se o item da coluna B não está presente na coluna A
        Set foundRange = rngA.Find(What:=cellB.Value, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Se não encontrar correspondência, adicionar o item à coleção de itens únicos
        If foundRange Is Nothing Then
            On Error Resume Next
            uniqueItems.Add cellB.Value, CStr(cellB.Value)
            On Error GoTo 0
        End If
    Next cellB
    
    ' Escrever os itens únicos na coluna C
    For Each item In uniqueItems
        ws.Cells(lastRowC + 1, "C").Value = item
        lastRowC = lastRowC + 1
    Next item
    
    ' Limpar objetos da memória
    Set ws = Nothing
    Set rngA = Nothing
    Set rngB = Nothing
    Set rngC = Nothing
    Set uniqueItems = Nothing
    
    MsgBox "Concluído! Itens únicos foram listados na coluna C.", vbInformation
End Sub

