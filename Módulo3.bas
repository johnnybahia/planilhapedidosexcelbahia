Attribute VB_Name = "Módulo3"
Option Explicit

Private limpado As Boolean

Sub GerarRelatorio()
    Dim xml_ws As Worksheet
    Dim relatorio_ws As Worksheet
    Dim copy_columns As Variant
    Dim paste_columns As Variant
    Dim i As Integer
    Dim last_row As Long
    
    ' Define as planilhas a serem usadas
    Set xml_ws = ThisWorkbook.Sheets("XML")
    Set relatorio_ws = ThisWorkbook.Sheets("RELATORIO")
    
    copy_columns = Array("O", "P", "A", "E", "B", "C", "D", "F", "I", "R", "K", "L", "M", "N", "S", "T")
    paste_columns = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "X")
    
    ' Define o valor inicial de last_row
    last_row = 4
    
    ' Copia as colunas da aba "XML" para a aba "Relatorio"
    For i = LBound(copy_columns) To UBound(copy_columns)
        If xml_ws.Cells(xml_ws.Rows.Count, copy_columns(i)).End(xlUp).Row >= 4 Then
            xml_ws.Range(copy_columns(i) & "4", xml_ws.Cells(xml_ws.Rows.Count, copy_columns(i)).End(xlUp)).Copy _
            Destination:=relatorio_ws.Range(paste_columns(i) & "4")
        End If
    Next i

    ' Atualiza a última linha usada na planilha "Relatorio" baseada na coluna A
    last_row = relatorio_ws.Cells(relatorio_ws.Rows.Count, 1).End(xlUp).Row
    If last_row < 4 Then last_row = 4

    ' ======================================================================
    ' --- INÍCIO DO TRATAMENTO DE CLIENTES (DASS, DILLY, DAKOTA, ANIGER) ---
    ' ======================================================================
    Dim r As Long
    Dim textoOC As String
    Dim partes() As String
    Dim valorExtraido As String
    Dim nomeCliente As String
    Dim dictCodigosDilly As Object
    Dim dictPendentesDilly As Object
    Dim descricaoOriginalDilly As String
    Dim descricaoChaveDilly As String
    Dim tamanhoDilly As String

    ' Cadastro perene de códigos Dilly (Módulo8) - carregado uma única vez
    Set dictCodigosDilly = CarregarDictCodigosDilly()
    Set dictPendentesDilly = CreateObject("Scripting.Dictionary")

    ' Loop pelas linhas copiadas no Relatório
    For r = 4 To last_row
        ' Pega o nome do cliente na Coluna B (2) e remove espaços extras
        nomeCliente = UCase(Trim(relatorio_ws.Cells(r, 2).Value))
        
        ' Pega o texto da Coluna J (10)
        textoOC = relatorio_ws.Cells(r, 10).Value
        
        ' --- REGRA CLIENTE: DASS ---
        If InStr(nomeCliente, "DASS") > 0 Then
            If InStr(textoOC, "-") > 0 Then
                partes = Split(textoOC, "-")
                valorExtraido = Trim(partes(UBound(partes)))
                
                relatorio_ws.Cells(r, 4).NumberFormat = "@"
                relatorio_ws.Cells(r, 4).Value = valorExtraido
            End If
            
        ' --- REGRA CLIENTE: DILLY ---
        ' O XML da Dilly deixou de trazer "OC:" na observação (Coluna J).
        ' A coluna D passa a ser deduzida do cadastro perene (Módulo8),
        ' usando a DESCRIÇÃO (Coluna F) como chave e o TAMANHO (Coluna G)
        ' como sufixo.
        ElseIf InStr(nomeCliente, "DILLY") > 0 Then
            descricaoOriginalDilly = Trim(CStr(relatorio_ws.Cells(r, 6).Value))
            descricaoChaveDilly = NormalizarDescricao(descricaoOriginalDilly)

            If descricaoChaveDilly <> "" Then
                If dictCodigosDilly.Exists(descricaoChaveDilly) Then
                    tamanhoDilly = ExtrairTamanhoSemUnidade(relatorio_ws.Cells(r, 7).Value)
                    relatorio_ws.Cells(r, 4).NumberFormat = "@"
                    relatorio_ws.Cells(r, 4).Value = dictCodigosDilly(descricaoChaveDilly) & "-" & tamanhoDilly
                ElseIf Not dictPendentesDilly.Exists(descricaoChaveDilly) Then
                    dictPendentesDilly.Add descricaoChaveDilly, descricaoOriginalDilly
                End If
            End If
            
        ' --- REGRA CLIENTES: DAKOTA E ANIGER ---
        ElseIf InStr(nomeCliente, "DAKOTA") > 0 Or InStr(nomeCliente, "ANIGER") > 0 Then
            If InStr(1, textoOC, "OC", vbTextCompare) > 0 Then
                ' Divide o texto usando "OC" como corte
                partes = Split(UCase(textoOC), "OC")
                
                ' Pega a PRIMEIRA parte do array (tudo antes do OC) e remove espaços
                valorExtraido = Trim(partes(0))
                
                ' Remove todos os zeros à esquerda da string
                Do While Left(valorExtraido, 1) = "0" And Len(valorExtraido) > 1
                    valorExtraido = Mid(valorExtraido, 2)
                Loop
                
                ' Força formato de texto e insere na Coluna 4 (D)
                relatorio_ws.Cells(r, 4).NumberFormat = "@"
                relatorio_ws.Cells(r, 4).Value = valorExtraido
            End If
        End If
    Next r

    ' Resolve (via InputBox, uma vez por descrição nova) as linhas DILLY
    ' cuja descrição ainda não tinha código base cadastrado.
    If dictPendentesDilly.Count > 0 Then
        Call ResolverPendenciasDilly(relatorio_ws, last_row, dictCodigosDilly, dictPendentesDilly)
    End If
    ' ======================================================================
    ' --- FIM DO TRATAMENTO DE CLIENTES ---
    ' ======================================================================

    ' Configurações de Fonte
    Dim font_size As Integer
    font_size = 14

    ' Loop para formatar (Número, Fonte e Alinhamento)
    For i = LBound(paste_columns) To UBound(paste_columns)
        With relatorio_ws.Range(paste_columns(i) & "4:" & paste_columns(i) & last_row)
            ' Converte texto para número (exceto colunas D, E, L e M que são datas)
            If paste_columns(i) <> "D" And paste_columns(i) <> "E" And _
               paste_columns(i) <> "L" And paste_columns(i) <> "M" Then
                .NumberFormat = "General"
                .Value = .Value
            End If
            
            ' Aplica Fonte e Tamanho
            .Font.Size = font_size
            .Font.Name = "Arial"
        End With
    Next i

    ' --- FORMATAÇÃO GERAL E BORDAS ---
    With relatorio_ws.Range("A4:X" & last_row)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Adiciona o filtro na linha 3
    If relatorio_ws.AutoFilterMode Then relatorio_ws.AutoFilterMode = False
    relatorio_ws.Range("A3:X" & last_row).AutoFilter

    ' Atualiza a variável de controle de estado
    limpado = False

    ' Verifica se há dados na aba "XML" e ativa a planilha Relatorio
    If xml_ws.Cells(4, "B").Value <> "" Then
        relatorio_ws.Activate
    End If

    ' --- TEXT TO COLUMNS ---
    Application.ScreenUpdating = False
    
    Call AplicarTextToColumns(relatorio_ws.Columns("A:A"))
    Call AplicarTextToColumns(relatorio_ws.Columns("C:C"))
    Call AplicarTextToColumns(relatorio_ws.Columns("D:D"))
    Call AplicarTextToColumns(relatorio_ws.Columns("E:E"))
    Call AplicarTextToColumns(relatorio_ws.Columns("G:G"))
    Call AplicarTextToColumns(relatorio_ws.Columns("H:H"))
    Call AplicarTextToColumns(relatorio_ws.Columns("I:I"))
    Call AplicarTextToColumns(relatorio_ws.Columns("J:J"))
    Call AplicarTextToColumns(relatorio_ws.Columns("K:K"))
    
    ' Colunas L e M são datas já em formato brasileiro - apenas garante exibição correta
    relatorio_ws.Range("L4:L" & last_row).NumberFormat = "DD/MM/YYYY"
    relatorio_ws.Range("M4:M" & last_row).NumberFormat = "DD/MM/YYYY"
    
    Call AplicarTextToColumns(relatorio_ws.Columns("N:N"))

    Application.ScreenUpdating = True
    
    ' Seleciona A4 no final
    relatorio_ws.Range("A4").Select

    MsgBox "Relatório gerado com sucesso!"

End Sub

' Sub rotina auxiliar para TextToColumns padrão
Sub AplicarTextToColumns(rng As Range)
    rng.TextToColumns Destination:=rng.Cells(1, 1), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
End Sub

