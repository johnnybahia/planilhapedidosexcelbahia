Attribute VB_Name = "M�dulo7"
Option Explicit

Private Const WEBAPP_URL       As String = "https://script.google.com/macros/s/AKfycbxLmPvyBTXC5VK67Uo6Hi9ubiaBRiBUaC595q0elIALW7LVaz-kJ2Fk88GAuyu7_ypj-Q/exec"
Private Const TOKEN            As String = "mrf_rel_9K7xQp2Zv4Ld8Tn3"
Private Const ABA_ORIGEM       As String = "RELATORIO"
Private Const COL_FIM          As Long = 24
Private Const LIN_FIM          As Long = 5000
Private Const LINHAS_POR_BLOCO As Long = 500
Private Const MIN_LINHAS       As Long = 2
Private Const MAX_TENT_BLOCO   As Long = 3
Private Const ESPERA_TENT_SEG  As Long = 5

Public Sub EnviarRelatorio()
    Dim ws As Worksheet, arr As Variant
    Dim ultLin As Long, r As Long, c As Long
    Dim totalBlocos As Long, bloco As Long, linIni As Long, linFimB As Long
    Dim sid As String, resp As String, payload As String, tentativaBloco As Long

    On Error GoTo Falha
    Set ws = LocalizarAba(ABA_ORIGEM)
    If ws Is Nothing Then Err.Raise vbObjectError + 1, , "Aba '" & ABA_ORIGEM & "' nao encontrada."

    Application.ScreenUpdating = False
    Application.StatusBar = "Lendo dados..."
    arr = ws.Range(ws.Cells(1, 1), ws.Cells(LIN_FIM, COL_FIM)).Value

    For r = LIN_FIM To 1 Step -1
        For c = 1 To COL_FIM
            If EstaPreenchida(arr(r, c)) Then ultLin = r: Exit For
        Next c
        If ultLin > 0 Then Exit For
    Next r

    If ultLin < MIN_LINHAS Then
        Err.Raise vbObjectError + 2, , "Somente " & ultLin & " linha(s) preenchida(s). Envio abortado."
    End If

    Randomize
    sid = Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int(Rnd() * 100000), "00000")
    totalBlocos = -Int(-ultLin / LINHAS_POR_BLOCO)

    For bloco = 1 To totalBlocos
        linIni = (bloco - 1) * LINHAS_POR_BLOCO + 1
        linFimB = linIni + LINHAS_POR_BLOCO - 1
        If linFimB > ultLin Then linFimB = ultLin
        payload = MontarPayload(arr, sid, bloco, totalBlocos, ultLin, linIni, linFimB)

        For tentativaBloco = 1 To MAX_TENT_BLOCO
            Application.StatusBar = "Enviando bloco " & bloco & " de " & totalBlocos & _
                IIf(tentativaBloco > 1, " (tentativa " & tentativaBloco & ")", "") & "..."
            resp = PostarComRetentativa(payload)
            If InStr(1, resp, """ok"":true", vbTextCompare) > 0 Then Exit For
            If Not EhFalhaTransitoria(resp) Then
                Err.Raise vbObjectError + 3, , "Falha no bloco " & bloco & " de " & totalBlocos & ": " & resp
            End If
            If tentativaBloco = MAX_TENT_BLOCO Then
                Err.Raise vbObjectError + 3, , "Falha no bloco " & bloco & " de " & totalBlocos & _
                    " apos " & MAX_TENT_BLOCO & " tentativas: " & resp
            End If
            Application.Wait Now + TimeSerial(0, 0, ESPERA_TENT_SEG)
        Next tentativaBloco
    Next bloco

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox ultLin & " linhas enviadas.", vbInformation
    Exit Sub

Falha:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbCritical
End Sub

Private Function MontarPayload(ByRef arr As Variant, ByVal sid As String, ByVal bloco As Long, _
                               ByVal totalBlocos As Long, ByVal totalLinhas As Long, _
                               ByVal linIni As Long, ByVal linFimB As Long) As String
    Dim buf() As String, cel() As String, r As Long, c As Long, i As Long
    ReDim buf(0 To linFimB - linIni)
    ReDim cel(1 To COL_FIM)
    i = 0
    For r = linIni To linFimB
        For c = 1 To COL_FIM
            cel(c) = CelulaJson(arr(r, c))
        Next c
        buf(i) = "[" & Join(cel, ",") & "]"
        i = i + 1
    Next r
    MontarPayload = "{""token"":""" & TOKEN & """,""sessionId"":""" & sid & _
                    """,""bloco"":" & bloco & ",""totalBlocos"":" & totalBlocos & _
                    ",""totalLinhas"":" & totalLinhas & ",""linhaInicial"":" & linIni & _
                    ",""dados"":[" & Join(buf, ",") & "]}"
End Function

Private Function CelulaJson(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        CelulaJson = """"""
    ElseIf VarType(v) = vbDate Then
        CelulaJson = """" & Format$(v, "dd/mm/yyyy") & """"
    ElseIf VarType(v) = vbBoolean Then
        CelulaJson = IIf(v, "true", "false")
    ElseIf VarType(v) = vbString Then
        CelulaJson = """" & EscJson(CStr(v)) & """"
    ElseIf IsNumeric(v) Then
        CelulaJson = Trim$(Str$(v))
    Else
        CelulaJson = """" & EscJson(CStr(v)) & """"
    End If
End Function

Private Function EstaPreenchida(ByVal v As Variant) As Boolean
    If IsError(v) Then EstaPreenchida = True: Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbString Then
        EstaPreenchida = (Len(Trim$(v)) > 0)
    Else
        EstaPreenchida = True
    End If
End Function

Private Function EscJson(ByVal s As String) As String
    Dim i As Long
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    For i = 0 To 31
        s = Replace(s, Chr$(i), "")
    Next i
    EscJson = s
End Function

Private Function Utf8Bytes(ByVal s As String) As Variant
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1
    st.Position = 3
    Utf8Bytes = st.Read
    st.Close
End Function

Private Function EhFalhaTransitoria(ByVal resp As String) As Boolean
    ' Timeout reportado pelo Apps Script ao acessar a planilha, ou erro de
    ' rede/HTTP jah esgotado pelo PostarComRetentativa: vale tentar de novo.
    EhFalhaTransitoria = (InStr(1, resp, "tempo limite", vbTextCompare) > 0) _
                       Or (Left$(resp, 5) = "Erro ") _
                       Or (Left$(resp, 5) = "HTTP ") _
                       Or (Left$(Trim$(resp), 1) <> "{")
End Function

Private Function PostarComRetentativa(ByVal payload As String) As String
    ' Uma unica tentativa aqui: a retentativa (com espera de
    ' ESPERA_TENT_SEG entre elas) ja e feita pelo chamador, por bloco.
    ' Duplicar a retentativa aqui so multiplicava requisicoes contra o
    ' Google em cima de um cenario de rate-limit/instabilidade.
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Option(6) = True
    http.SetTimeouts 30000, 30000, 60000, 360000
    http.Open "POST", WEBAPP_URL, False
    http.SetRequestHeader "Content-Type", "application/json;charset=UTF-8"
    http.Send Utf8Bytes(payload)
    If Err.Number = 0 And http.Status = 200 Then
        PostarComRetentativa = http.ResponseText
    ElseIf Err.Number <> 0 Then
        PostarComRetentativa = "Erro " & Err.Number & ": " & Err.Description
    Else
        PostarComRetentativa = "HTTP " & http.Status
    End If
    On Error GoTo 0
End Function

Private Function LocalizarAba(ByVal nome As String) As Worksheet
    Dim ws As Worksheet, alvo As String
    alvo = Normalizar(nome)
    For Each ws In ThisWorkbook.Worksheets
        If Normalizar(ws.Name) = alvo Then
            Set LocalizarAba = ws
            Exit Function
        End If
    Next ws
End Function

Private Function Normalizar(ByVal s As String) As String
    Dim de As String, pa As String, i As Long
    de = "����������������������������������������������"
    pa = "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC"
    For i = 1 To Len(de)
        s = Replace(s, Mid$(de, i, 1), Mid$(pa, i, 1))
    Next i
    Normalizar = LCase$(Trim$(s))
End Function

