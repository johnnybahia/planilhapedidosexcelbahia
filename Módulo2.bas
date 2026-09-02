Attribute VB_Name = "Módulo2"
Sub Macro15()
Attribute Macro15.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro15 Macro
'

'
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=7, Criteria1:="<>*MM*" _
        , Operator:=xlAnd
    Range("J5").Select
End Sub
Sub Macro16()
Attribute Macro16.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro16 Macro
'

'
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=7, Criteria1:="=*MM*", _
        Operator:=xlAnd
    ActiveWindow.SmallScroll Down:=-12
    Range("J164").Select
End Sub
Sub Macro17()
Attribute Macro17.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro17 Macro
'

'
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=1
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=2
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=3
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=4
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=4
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=4
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=4
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=5
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=5
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=6
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=7
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=8
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=9
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=10
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=11
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=12
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=13
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveSheet.Range("$A$3:$W$3964").AutoFilter Field:=14
    Range("J4").Select
End Sub
