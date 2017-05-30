Attribute VB_Name = "Module3"
Sub LogFileCleanUp()
Attribute LogFileCleanUp.VB_Description = "Macro recorded 2/21/2008 by Osmetech CCD"
Attribute LogFileCleanUp.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Dim LoopNo As Integer
    Dim R1 As Integer
    Dim R2 As Integer
    Dim y As Integer
    
    LoopNo = Range("F42")
    R2 = 1
    
    Sheets("Paste LogFile Here").Select
    
    Cells.Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("Q1") = "=SUM(P:P)"
    
    For y = 1 To LoopNo
    
        Range("A" & R2).Select
    
        Cells.Find(What:="  endpoint.c 771", After:=ActiveCell, LookIn:= _
                xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(0, 15) = "=ROW()"
        R1 = Range("Q1")
        Columns("P:P").ClearContents
    
        Rows(R1).Select
        Selection.Delete Shift:=xlUp
        Rows(R1).Select
        Selection.Delete Shift:=xlUp
        
        R2 = R1 - 2

    Next
    
    Cells.Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Range("Q1").ClearContents
    Range("A1").Select

End Sub
