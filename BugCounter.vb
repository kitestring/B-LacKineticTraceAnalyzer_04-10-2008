Attribute VB_Name = "Module2"
Sub BugCounter()
Attribute BugCounter.VB_Description = "Macro recorded 2/14/2008 by Osmetech CCD"
Attribute BugCounter.VB_ProcData.VB_Invoke_Func = " \n14"
    
'Title #1*************************************************************************
'---------------------------------------------------------------------------------
'Define & Set Variables, Open/Copy Template, & Copy/Paste LogFile Into Template
'---------------------------------------------------------------------------------
'*********************************************************************************
    
    Application.DisplayAlerts = False
    
    Dim TDryCal As Byte
    Dim TSam As Byte
    Dim TGasCal As Byte

    'Dim DryCalRow As Long
    'Dim SamRow As Long
    'Dim GasCalRow As Long
    
    Dim LoopNo As Byte
    Dim RunCounter As Byte
    Dim x As Byte
    Dim bugs As Byte
    
    TDryCal = Range("F39")
    TSam = Range("F40")
    TGasCal = Range("F41")
    RunCounter = 1
    bugs = 0
    
    Select Case TDryCal
        Case Is = TSam
            LoopNo = TSam
        Case Is > TSam
            LoopNo = TSam
        Case Is < TSam
            LoopNo = TDryCal
    End Select
    
    Workbooks.Open ("N:\RD\K_Kite\Macro\Lactate_LogFile_BugScan_rev A.xls")
    Windows("Lactate_LogFile_BugScan_rev A.xls").Activate
    ActiveWorkbook.SaveAs Filename:="N:\RD\2008 Experiments\Lactate\08-022 Macro Destination Folder\BugScan_LogFile Analysis.xls", FileFormat:=xlWorkbookNormal, AddToMru:=True

    Windows("Lactate_Logfile_Macro_rev B.xls").Activate
    Sheets("Paste LogFile Here").Select
    Cells.Select
    Selection.Copy
    
    Windows("BugScan_LogFile Analysis.xls").Activate
    Sheets("LogFile").Select
    Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
    Range("A1").Select
    
    Windows("Lactate_Logfile_Macro_rev B.xls").Activate
    Sheets("Macro Controls").Select
    Windows("BugScan_LogFile Analysis.xls").Activate
    Sheets("LogFile").Select
    ActiveWindow.Zoom = 70

'Title #1*************************************************************************
'---------------------------------------------------------------------------------
'Software Bug Scan & Label
'---------------------------------------------------------------------------------
'*********************************************************************************

    Range("A1").Select
    
    For x = 1 To TSam
    
        Cells.Find(What:="MEASURE 240.0  individual endpoint", After:=ActiveCell, _
                LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Replace What:="MEASURE 240.0  individual endpoint", Replacement _
                :="Sam Meas " & RunCounter, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False _
                , SearchFormat:=False, ReplaceFormat:=False
                Selection.Font.ColorIndex = 3
                Selection.Font.Bold = True
    
        ActiveCell.Offset(51, 0).Activate
        If ActiveCell = "Setting CO2 ETH to 0.00150" Then
            Selection.EntireRow.Font.Bold = True
            Selection.EntireRow.Font.ColorIndex = 5
            ActiveCell.Offset(0, 15) = "Software Bug"
            bugs = bugs + 1
            Sheets("Software Bug Results").Select
            Range("D" & 8 + bugs) = RunCounter
            Sheets("LogFile").Select
        End If
    
        Range("A1").Select
        RunCounter = RunCounter + 1

    Next
    
    If bugs = 0 Then
        Sheets("Software Bug Results").Select
        Range("D9") = "No Software bugs detected"
    End If
    
    Sheets("Software Bug Results").Select
End Sub
