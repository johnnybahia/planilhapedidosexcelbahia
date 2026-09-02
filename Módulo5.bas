Attribute VB_Name = "Módulo5"
Sub PUXAR_OC()
Attribute PUXAR_OC.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PUXAR_OC Macro
'

'
    Sheets("RELATORIO").Select
    Range("H4:H10000").Select
    Selection.Copy
    Sheets("VERIFICAÇÃO DE PEDIDOS").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C1").Select
End Sub
Sub APAGAR_OC()
Attribute APAGAR_OC.VB_ProcData.VB_Invoke_Func = " \n14"
'
' APAGAR_OC Macro
'

'
    Range("A2:c10020").Select
    Selection.ClearContents
    Range("C1").Select
End Sub
