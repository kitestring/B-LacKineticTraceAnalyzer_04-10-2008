Attribute VB_Name = "Module1"
Sub LogFileAnalysis()


'Title #1*************************************************************************
'---------------------------------------------------------------------------------
'Open/Copy Template, Copy/Paste LogFile Into Template, & Define & Set Variables
'---------------------------------------------------------------------------------
'*********************************************************************************

    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim ExpData As Date
    Dim ExpTitle As String
    Dim NoL1 As Byte
    Dim NoL2 As Byte
    Dim NoL3 As Byte
    Dim TotalRuns As Byte
    Dim RunCounter As Byte
    Dim MF1 As Byte
    Dim MF2 As Byte
    Dim MF13 As Byte
    Dim MF14 As String
    Dim RowRef1 As Long
    Dim RowRef2 As Long
    Dim RowDif As Integer
    Dim RowsHolder As Integer
    
    Dim CO2_DryCal As ChartObject
    Dim CO2_GasCal As ChartObject
    Dim CO2_ETH As ChartObject
    Dim O2_DryCal As ChartObject
    Dim O2_GasCal As ChartObject
    Dim O2_ETH As ChartObject
    Dim Lac_DryCal As ChartObject
    Dim Lac_GasCal As ChartObject
    Dim Lac_ETH As ChartObject
    Dim pH_DryCal As ChartObject
    Dim pH_GasCal As ChartObject
    Dim pH_ETH As ChartObject
    
    Dim x As Double
    
    Dim CO2_Min_DryCal As Double
    Dim CO2_Min_GasCal As Double
    Dim CO2_Min_ETH As Double
    Dim O2_Min_DryCal As Double
    Dim O2_Min_GasCal As Double
    Dim O2_Min_ETH As Double
    Dim Lac_Min_DryCal As Double
    Dim Lac_Min_GasCal As Double
    Dim Lac_Min_ETH As Double
    Dim pH_Min_DryCal As Double
    Dim pH_Min_GasCal As Double
    Dim pH_Min_ETH As Double
    
    Dim CO2_Max_DryCal As Double
    Dim CO2_Max_GasCal As Double
    Dim CO2_Max_ETH As Double
    Dim O2_Max_DryCal As Double
    Dim O2_Max_GasCal As Double
    Dim O2_Max_ETH As Double
    Dim Lac_Max_DryCal As Double
    Dim Lac_Max_GasCal As Double
    Dim Lac_Max_ETH As Double
    Dim pH_Max_DryCal As Double
    Dim pH_Max_GasCal As Double
    Dim pH_Max_ETH As Double
    
    Dim GMin As Double
    Dim GMax As Double
    Dim xMax As Double
    
    Dim LoopNo As Integer
    Dim R1 As Integer
    Dim R2 As Integer
    Dim y As Integer
    
    LoopNo = Range("F42")
    ExpData = Range("B11")
    ExpTitle = Range("F12")
    NoL1 = Range("G12")
    NoL2 = Range("H12")
    NoL3 = Range("I12")
    TotalRuns = NoL1 + NoL2 + NoL3
    RunCounter = 1
    R2 = 1
    


    Workbooks.Open ("M:\My Documents\Macros\Lactate Macros\Lactate_LogFile_Template_rev B.xls")
    Windows("Lactate_LogFile_Template_rev B.xls").Activate
    ActiveWorkbook.SaveAs Filename:="M:\My Documents\Macros\Lactate Macros\CassetteNo_CCA SN_()_LogFile Analysis.xls", FileFormat:=xlWorkbookNormal, AddToMru:=True
    Sheets("Graphs").Unprotect Password:="water18AMU"
    Sheets("Analyzed LogFile").Unprotect Password:="water18AMU"
    Sheets("Cleaned Logfile").Unprotect Password:="water18AMU"
    Sheets("Original Logfile").Unprotect Password:="water18AMU"
    Sheets("Data").Unprotect Password:="water18AMU"
    ActiveWorkbook.Unprotect Password:="water18AMU"

    Windows("Lactate_Logfile_Macro_rev B.xls").Activate
    Sheets("Paste LogFile Here").Select
    Cells.Select
    Selection.Copy
    
    Windows("CassetteNo_CCA SN_()_LogFile Analysis.xls").Activate
    Sheets("Original LogFile").Select
    Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        'Application.CutCopyMode = False
    Range("A1").Select
    
    Sheets("Cleaned LogFile").Select
    Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
    Range("A1").Select
    ActiveWindow.Zoom = 60
    
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
    
    Range("Q1").ClearContents
    Range("A1").Select
    Cells.Copy
    Sheets("Analyzed LogFile").Select
    Cells.Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
    Range("A1").Select
    ActiveWindow.Zoom = 60
    
    
    Windows("Lactate_Logfile_Macro_rev B.xls").Activate
    Sheets("Macro Controls").Select
    Windows("CassetteNo_CCA SN_()_LogFile Analysis.xls").Activate
    Sheets("Analyzed LogFile").Select
    
'Title #2*************************************************************************
'---------------------------------------------------------------------------------
'Label Data Clusters Within LogFile
'---------------------------------------------------------------------------------
'*********************************************************************************
    
    
    For MF1 = 1 To TotalRuns
    
        Cells.Find(What:="Setting hot_flag=FALSE", After:=ActiveCell, LookIn:= _
            xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Replace What:="Setting hot_flag=FALSE", Replacement:= _
            "Dry Cal " & RunCounter, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
            SearchFormat:=False, ReplaceFormat:=False
            Selection.Font.ColorIndex = 3
            Selection.Font.Bold = True
        Range("A1").Select
    
        Cells.Find(What:="MEASURE 240.0  individual endpoint", After:=ActiveCell, _
            LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Replace What:="MEASURE 240.0  individual endpoint", Replacement _
            :="Sam Meas " & RunCounter, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False _
            , SearchFormat:=False, ReplaceFormat:=False
            Selection.Font.ColorIndex = 3
            Selection.Font.Bold = True
        Range("A1").Select
    
        Cells.Find(What:="CALGAS2 90.0  individual endpoint", After:=ActiveCell, _
            LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Replace What:="CALGAS2 90.0  individual endpoint", Replacement:= _
            "Gas Cal " & RunCounter, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
            SearchFormat:=False, ReplaceFormat:=False
            Selection.Font.ColorIndex = 3
            Selection.Font.Bold = True
        Range("A1").Select
        
        RunCounter = RunCounter + 1
    
    Next
    
    RunCounter = 1
    
    GoTo CopyPasteWkSht

'Title #3*************************************************************************
'---------------------------------------------------------------------------------
'Copy/Paste Run # Worksheet
'---------------------------------------------------------------------------------
'*********************************************************************************

CopyPasteWkSht:

    Windows("CassetteNo_CCA SN_()_LogFile Analysis.xls").Activate
    Sheets("Data").Select
    Sheets("Data").Copy After:=Sheets("Data")
    ActiveSheet.Name = "Run " & RunCounter
    
    GoTo DetDryCalRange
Exit Sub
    
'Title #4*************************************************************************
'---------------------------------------------------------------------------------
'Determine Dry Calibration Data Range Locations & Set Live Links In Run Worksheet
'---------------------------------------------------------------------------------
'*********************************************************************************
    
DetDryCalRange:
    
    Sheets("Analyzed LogFile").Select
    Range("A1").Select
    
    Cells.Find(What:="Dry Cal " & RunCounter, After:=ActiveCell, LookIn:= _
            xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Activate
        Selection.End(xlToRight).Select
        Selection.End(xlDown).Select
        Selection.End(xlToLeft).Select
        ActiveCell.Offset(-2, 0).Activate
        ActiveCell.Range(Cells(1, 1), Cells(3, 14)).Select
            Selection.Font.ColorIndex = 5
            Selection.Font.Bold = True
        ActiveCell.Offset(0, 15).Activate
        ActiveCell = "=ROW()"
        RowRef1 = ActiveCell
        ActiveCell.ClearContents
        
    Sheets("Run " & RunCounter).Select
        Range("AH5") = "='Analyzed LogFile'!A" & RowRef1
        Range("AH5").Copy
        Range("AH5:AU7").Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    GoTo DetGasCalRange
Exit Sub

'Title #5*************************************************************************
'---------------------------------------------------------------------------------
'Determine Gas Calibration Data Range Locations & Set Live Links In Run Worksheet
'---------------------------------------------------------------------------------
'*********************************************************************************
    
DetGasCalRange:

    Sheets("Analyzed LogFile").Select
    Range("A1").Select
    
    Cells.Find(What:="Gas Cal " & RunCounter, After:=ActiveCell, LookIn:= _
            xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
            xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Activate
        Selection.End(xlToRight).Select
        Selection.End(xlDown).Select
        Selection.End(xlToLeft).Select
        ActiveCell.Offset(-2, 0).Activate
        ActiveCell.Range(Cells(1, 1), Cells(3, 14)).Select
            Selection.Font.ColorIndex = 5
            Selection.Font.Bold = True
        ActiveCell.Offset(0, 15).Activate
        ActiveCell = "=ROW()"
        RowRef1 = ActiveCell
        ActiveCell.ClearContents
        
    Sheets("Run " & RunCounter).Select
        Range("AW5") = "='Analyzed LogFile'!A" & RowRef1
        Range("AW5").Copy
        Range("AW5:BJ7").Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
      GoTo DetSamMeasRange
Exit Sub

'Title #6*************************************************************************
'---------------------------------------------------------------------------------
'Determine Sample Measurement Data Range Locations & Set Live Links In Run Worksheet
'---------------------------------------------------------------------------------
'*********************************************************************************

DetSamMeasRange:

    Sheets("Analyzed LogFile").Select
    Range("A1").Select
    
    Cells.Find(What:="Sam Meas " & RunCounter, After:=ActiveCell, _
            LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(4, 0).Activate
        Selection.End(xlToRight).Select
        Selection.End(xlDown).Select
        Selection.End(xlToLeft).Select
        ActiveCell.Range(Cells(2, 1), Cells(4, 1)).Select
        Selection.EntireRow.Delete
        
    Range("A1").Select
    
    Cells.Find(What:="Sam Meas " & RunCounter, After:=ActiveCell, _
            LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(48, 0).Activate
        If ActiveCell = "Setting CO2 ETH to 0.00150" Then
            Selection.EntireRow.Delete
        End If
        
    Range("A1").Select
    
    Cells.Find(What:="Sam Meas " & RunCounter, After:=ActiveCell, _
            LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(4, 0).Activate
        ActiveCell.Offset(0, 15).Activate
        ActiveCell = "=ROW()"
        RowRef1 = ActiveCell
        ActiveCell.ClearContents
        
        ActiveCell.Offset(0, -15).Activate
        Selection.End(xlToRight).Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(0, 2).Activate
        ActiveCell = "=ROW()"
        RowRef2 = ActiveCell
        ActiveCell.ClearContents
        
        RowDif = RowRef2 - RowRef1 + 5
        RowsHolder = RowRef2 - RowRef1 + 1
        
        Range("A" & RowRef1 & ":N" & RowRef2).Select
            Selection.Font.ColorIndex = 5
            Selection.Font.Bold = True
            
    Sheets("Run " & RunCounter).Select
        Range("T2") = RowsHolder
        Range("S5") = "='Analyzed LogFile'!A" & RowRef1
        Range("S5").Copy
        Range("S5:AF" & RowDif).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        
    Range("Z5:AE7").ClearContents
    
    Range("B103") = "=MAX(B5:B" & RowsHolder + 4 & ")"
    
    Range("C102") = "=MIN(C5:C" & RowsHolder + 4 & ")"
    Range("C103") = "=MAX(C5:C" & RowsHolder + 4 & ")"
    
    Range("D102") = "=MIN(D5:D" & RowsHolder + 4 & ")"
    Range("D103") = "=MAX(D5:D" & RowsHolder + 4 & ")"
    
    Range("E102") = "=MIN(E5:E" & RowsHolder + 4 & ")"
    Range("E103") = "=MAX(E5:E" & RowsHolder + 4 & ")"
    
    Range("H102") = "=MIN(H5:H" & RowsHolder + 4 & ")"
    Range("H103") = "=MAX(H5:H" & RowsHolder + 4 & ")"
    
    Range("K102") = "=MIN(K5:K" & RowsHolder + 4 & ")"
    Range("K103") = "=MAX(K5:K" & RowsHolder + 4 & ")"
    
    Range("L102") = "=MIN(L5:L" & RowsHolder + 4 & ")"
    Range("L103") = "=MAX(L5:L" & RowsHolder + 4 & ")"
    
    Range("M102") = "=MIN(M5:M" & RowsHolder + 4 & ")"
    Range("M103") = "=MAX(M5:M" & RowsHolder + 4 & ")"
    
    Range("P102") = "=MIN(P5:P" & RowsHolder + 4 & ")"
    Range("P103") = "=MAX(P5:P" & RowsHolder + 4 & ")"
    
    Range("Z102") = "=MIN(Z8:Z" & RowsHolder + 4 & ")"
    Range("Z103") = "=MAX(Z8:Z" & RowsHolder + 4 & ")"
    
    Range("AA102") = "=MIN(AA8:AA" & RowsHolder + 4 & ")"
    Range("AA103") = "=MAX(AA8:AA" & RowsHolder + 4 & ")"
    
    Range("AB102") = "=MIN(AB8:AB" & RowsHolder + 4 & ")"
    Range("AB103") = "=MAX(AB8:AB" & RowsHolder + 4 & ")"
    
    Range("AE102") = "=MIN(AE8:AE" & RowsHolder + 4 & ")"
    Range("AE103") = "=MAX(AE8:AE" & RowsHolder + 4 & ")"
        
    Range("A1").Select
    
    If RunCounter = TotalRuns Then
        GoTo SetScales
    End If
        
    RunCounter = RunCounter + 1
        
    GoTo CopyPasteWkSht
Exit Sub

'Title #7*************************************************************************
'---------------------------------------------------------------------------------
'SetScales
'---------------------------------------------------------------------------------
'*********************************************************************************

SetScales:

    Sheets("Run 1").Select
    
        x = Range("B103")
    
        CO2_Min_DryCal = Range("C102")
        CO2_Min_GasCal = Range("K102")
        CO2_Min_ETH = Range("Z102")
        O2_Min_DryCal = Range("D102")
        O2_Min_GasCal = Range("L102")
        O2_Min_ETH = Range("AA102")
        Lac_Min_DryCal = Range("E102")
        Lac_Min_GasCal = Range("M102")
        Lac_Min_ETH = Range("AB102")
        pH_Min_DryCal = Range("H102")
        pH_Min_GasCal = Range("P102")
        pH_Min_ETH = Range("AE102")
    
        CO2_Max_DryCal = Range("C103")
        CO2_Max_GasCal = Range("K103")
        CO2_Max_ETH = Range("Z103")
        O2_Max_DryCal = Range("D103")
        O2_Max_GasCal = Range("L103")
        O2_Max_ETH = Range("AA103")
        Lac_Max_DryCal = Range("E103")
        Lac_Max_GasCal = Range("M103")
        Lac_Max_ETH = Range("AB103")
        pH_Max_DryCal = Range("H103")
        pH_Max_GasCal = Range("P103")
        pH_Max_ETH = Range("AE103")
        
        MF13 = 2
        
        Do While MF13 < (TotalRuns + 1)
            
            Sheets("Run " & MF13).Select
            
                If x < Range("B103") Then
                    x = Range("B103")
                End If
            
                If CO2_Min_DryCal > Range("C102") Then
                    CO2_Min_DryCal = Range("C102")
                End If
                If CO2_Min_GasCal > Range("K102") Then
                    CO2_Min_GasCal = Range("K102")
                End If
                If CO2_Min_ETH > Range("Z102") Then
                    CO2_Min_ETH = Range("Z102")
                End If
                If O2_Min_DryCal > Range("D102") Then
                    O2_Min_DryCal = Range("D102")
                End If
                If O2_Min_GasCal > Range("L102") Then
                    O2_Min_GasCal = Range("L102")
                End If
                If O2_Min_ETH > Range("AA102") Then
                    O2_Min_ETH = Range("AA102")
                End If
                If Lac_Min_DryCal > Range("E102") Then
                    Lac_Min_DryCal = Range("E102")
                End If
                If Lac_Min_GasCal > Range("M102") Then
                    Lac_Min_GasCal = Range("M102")
                End If
                If Lac_Min_ETH > Range("AB102") Then
                    Lac_Min_ETH = Range("AB102")
                End If
                If pH_Min_DryCal > Range("H102") Then
                    pH_Min_DryCal = Range("H102")
                End If
                If pH_Min_GasCal > Range("P102") Then
                    pH_Min_GasCal = Range("P102")
                End If
                If pH_Min_ETH > Range("AE102") Then
                    pH_Min_ETH = Range("AE102")
                End If
    
    
    
                If CO2_Max_DryCal < Range("C103") Then
                    CO2_Max_DryCal = Range("C103")
                End If
                If CO2_Max_GasCal < Range("K103") Then
                    CO2_Max_GasCal = Range("K103")
                End If
                If CO2_Max_ETH < Range("Z103") Then
                    CO2_Max_ETH = Range("Z103")
                End If
                If O2_Max_DryCal < Range("D103") Then
                    O2_Max_DryCal = Range("D103")
                End If
                If O2_Max_GasCal < Range("L103") Then
                    O2_Max_GasCal = Range("L103")
                End If
                If O2_Max_ETH < Range("AA103") Then
                    O2_Max_ETH = Range("AA103")
                End If
                If Lac_Max_DryCal < Range("E103") Then
                    Lac_Max_DryCal = Range("E103")
                End If
                If Lac_Max_GasCal < Range("M103") Then
                    Lac_Max_GasCal = Range("M103")
                End If
                If Lac_Max_ETH < Range("AB103") Then
                    Lac_Max_ETH = Range("AB103")
                End If
                If pH_Max_DryCal < Range("H103") Then
                    pH_Max_DryCal = Range("H103")
                End If
                If pH_Max_GasCal < Range("P103") Then
                    pH_Max_GasCal = Range("P103")
                End If
                If pH_Max_ETH < Range("AE103") Then
                    pH_Max_ETH = Range("AE103")
                End If
                
                MF13 = MF13 + 1
                
            Loop
    
    Sheets("Graphs").Select
    
            Range("A1") = x
    
            Range("C1") = CO2_Min_DryCal
            Range("O1") = CO2_Min_GasCal
            Range("AB1") = CO2_Min_ETH
            Range("C36") = O2_Min_DryCal
            Range("O36") = O2_Min_GasCal
            Range("AB36") = O2_Min_ETH
            Range("C71") = Lac_Min_DryCal
            Range("O71") = Lac_Min_GasCal
            Range("AB71") = Lac_Min_ETH
            Range("C105") = pH_Min_DryCal
            Range("O105") = pH_Min_GasCal
            Range("AB105") = pH_Min_ETH
    
            Range("C2") = CO2_Max_DryCal
            Range("O2") = CO2_Max_GasCal
            Range("AB2") = CO2_Max_ETH
            Range("C37") = O2_Max_DryCal
            Range("O37") = O2_Max_GasCal
            Range("AB37") = O2_Max_ETH
            Range("C72") = Lac_Max_DryCal
            Range("O72") = Lac_Max_GasCal
            Range("AB72") = Lac_Max_ETH
            Range("C106") = pH_Max_DryCal
            Range("O106") = pH_Max_GasCal
            Range("AB106") = pH_Max_ETH
        
    GoTo Graphing
Exit Sub

'Title #8*************************************************************************
'---------------------------------------------------------------------------------
'Generate and Populate Graphs
'---------------------------------------------------------------------------------
'*********************************************************************************

Graphing:

     RunCounter = 1

'***************
'---------------
'CO2 Graphs
'---------------
'***************

    Set CO2_DryCal = Sheets("Graphs").ChartObjects.Add _
    (Left:=5, Top:=40, Width:=600, Height:=400)
    CO2_DryCal.Chart.ChartType = xlXYScatter

    Set CO2_GasCal = Sheets("Graphs").ChartObjects.Add _
    (Left:=610, Top:=40, Width:=600, Height:=400)
    CO2_GasCal.Chart.ChartType = xlXYScatter
    
    Set CO2_ETH = Sheets("Graphs").ChartObjects.Add _
    (Left:=1215, Top:=40, Width:=600, Height:=400)
    CO2_ETH.Chart.ChartType = xlXYScatter
    
    For MF2 = 1 To TotalRuns

        CO2_DryCal.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("C5:C101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With
    
        CO2_GasCal.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("K5:K101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With

        CO2_ETH.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("Z5:Z101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With
    
        RunCounter = RunCounter + 1
    Next

'------------------------------
'CO2 Dry Cal Graph Formatting
'------------------------------

    xMax = Range("A2")
    GMin = FormatNumber(Range("K1"), 2, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("K2"), 2, vbUseDefault, vbUseDefault, vbTrue)
    
    CO2_DryCal.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Sam/Dry Cal PCO2" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "Normalized Intensity"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
    
'------------------------------
'CO2 Gas Cal Graph Formatting
'------------------------------
    
    GMin = FormatNumber(Range("W1"), 2, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("W2"), 2, vbUseDefault, vbUseDefault, vbTrue)
    
    Range("U3").Select
    CO2_GasCal.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Sam/Gas Cal PCO2" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "Normalized Intensity"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        
'------------------------------
'CO2 ETH Graph Formatting
'------------------------------
        
    GMin = FormatNumber(Range("AJ1"), 1, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("AJ2"), 1, vbUseDefault, vbUseDefault, vbTrue)
        
    Range("AG3").Select
    CO2_ETH.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "ETH PCO2" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "ETH Value"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            '.MinimumScaleIsAuto = True
            '.MaximumScaleIsAuto = True
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlCustom
            .CrossesAt = -5000
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
               
    RunCounter = 1
    Range("A51").Select

'***************
'---------------
'O2 Graphs
'---------------
'***************

    Set O2_DryCal = Sheets("Graphs").ChartObjects.Add _
    (Left:=5, Top:=480, Width:=600, Height:=400)
    O2_DryCal.Chart.ChartType = xlXYScatter

    Set O2_GasCal = Sheets("Graphs").ChartObjects.Add _
    (Left:=610, Top:=480, Width:=600, Height:=400)
    O2_GasCal.Chart.ChartType = xlXYScatter
    
    Set O2_ETH = Sheets("Graphs").ChartObjects.Add _
    (Left:=1215, Top:=480, Width:=600, Height:=400)
    O2_ETH.Chart.ChartType = xlXYScatter
    
    For MF2 = 1 To TotalRuns

        O2_DryCal.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("D5:D101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With
    
        O2_GasCal.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("L5:L101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With

        O2_ETH.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("AA5:AA101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With
    
        RunCounter = RunCounter + 1
    Next

'------------------------------
'O2 Dry Cal Graph Formatting
'------------------------------

    GMin = FormatNumber(Range("K36"), 2, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("K37"), 2, vbUseDefault, vbUseDefault, vbTrue)
    
    O2_DryCal.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Sam/Dry Cal PO2" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "Normalized Intensity"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
    
'------------------------------
'O2 Gas Cal Graph Formatting
'------------------------------
    
    GMin = FormatNumber(Range("W36"), 2, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("W37"), 2, vbUseDefault, vbUseDefault, vbTrue)
    
    Range("U51").Select
    O2_GasCal.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Sam/Gas Cal PO2" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "Normalized Intensity"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        
'------------------------------
'O2 ETH Graph Formatting
'------------------------------
        
    GMin = FormatNumber(Range("AJ36"), 1, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("AJ37"), 1, vbUseDefault, vbUseDefault, vbTrue)
    
    Range("AG51").Select
    O2_ETH.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "ETH PO2" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "ETH Value"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            '.MinimumScaleIsAuto = True
            '.MaximumScaleIsAuto = True
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlCustom
            .CrossesAt = -2000
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        
    RunCounter = 1
    Range("A85").Select

'***************
'---------------
'Lac Graphs
'---------------
'***************

    Set Lac_DryCal = Sheets("Graphs").ChartObjects.Add _
    (Left:=5, Top:=920, Width:=600, Height:=400)
    Lac_DryCal.Chart.ChartType = xlXYScatter

    Set Lac_GasCal = Sheets("Graphs").ChartObjects.Add _
    (Left:=610, Top:=920, Width:=600, Height:=400)
    Lac_GasCal.Chart.ChartType = xlXYScatter
    
    Set Lac_ETH = Sheets("Graphs").ChartObjects.Add _
    (Left:=1215, Top:=920, Width:=600, Height:=400)
    Lac_ETH.Chart.ChartType = xlXYScatter
    
    For MF2 = 1 To TotalRuns

        Lac_DryCal.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("E5:E101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With
    
        Lac_GasCal.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("M5:M101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With

        Lac_ETH.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("AB5:AB101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With
    
        RunCounter = RunCounter + 1
    Next

'------------------------------
'Lac Dry Cal Graph Formatting
'------------------------------

    GMin = FormatNumber(Range("K71"), 2, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("K72"), 2, vbUseDefault, vbUseDefault, vbTrue)
    
    Lac_DryCal.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Sam/Dry Cal Lac" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "Normalized Intensity"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
    
'------------------------------
'Lac Gas Cal Graph Formatting
'------------------------------
    
    GMin = FormatNumber(Range("W71"), 2, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("W72"), 2, vbUseDefault, vbUseDefault, vbTrue)
    
    Range("U85").Select
    Lac_GasCal.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Sam/Gas Cal Lac" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "Normalized Intensity"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        
'------------------------------
'Lac ETH Graph Formatting
'------------------------------
        
    GMin = FormatNumber(Range("AJ71"), 1, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("AJ72"), 1, vbUseDefault, vbUseDefault, vbTrue)
    
    Range("AG85").Select
    Lac_ETH.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "ETH Lac" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "ETH Value"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            '.MinimumScaleIsAuto = True
            '.MaximumScaleIsAuto = True
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlCustom
            .CrossesAt = 0
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
    
    RunCounter = 1
    Range("A121").Select

'***************
'---------------
'pH Graphs
'---------------
'***************

    Set pH_DryCal = Sheets("Graphs").ChartObjects.Add _
    (Left:=5, Top:=1360, Width:=600, Height:=400)
    pH_DryCal.Chart.ChartType = xlXYScatter

    Set pH_GasCal = Sheets("Graphs").ChartObjects.Add _
    (Left:=610, Top:=1360, Width:=600, Height:=400)
    pH_GasCal.Chart.ChartType = xlXYScatter
    
    Set pH_ETH = Sheets("Graphs").ChartObjects.Add _
    (Left:=1215, Top:=1360, Width:=600, Height:=400)
    pH_ETH.Chart.ChartType = xlXYScatter
    
    For MF2 = 1 To TotalRuns

        pH_DryCal.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("H5:H101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With
    
        pH_GasCal.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("P5:P101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With

        pH_ETH.Activate
        With ActiveChart.SeriesCollection.NewSeries
            .Name = "Run " & RunCounter
            .Values = Worksheets("Run " & RunCounter).Range("AE5:AE101")
            .XValues = Worksheets("Run " & RunCounter).Range("B5:B101")
        End With
    
        RunCounter = RunCounter + 1
    Next

'------------------------------
'pH Dry Cal Graph Formatting
'------------------------------

    GMin = FormatNumber(Range("K105"), 2, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("K106"), 2, vbUseDefault, vbUseDefault, vbTrue)

    pH_DryCal.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Sam/Dry Cal pH" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "Normalized Intensity"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
    
'------------------------------
'pH Gas Cal Graph Formatting
'------------------------------
    
    GMin = FormatNumber(Range("W105"), 2, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("W106"), 2, vbUseDefault, vbUseDefault, vbTrue)
    
    Range("U121").Select
    pH_GasCal.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "Sam/Gas Cal pH" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "Normalized Intensity"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        
'------------------------------
'pH ETH Graph Formatting
'------------------------------
        
    GMin = FormatNumber(Range("AJ105"), 1, vbUseDefault, vbUseDefault, vbTrue)
    GMax = FormatNumber(Range("AJ106"), 1, vbUseDefault, vbUseDefault, vbTrue)
    
    Range("AG121").Select
    pH_ETH.Activate
        
        ActiveChart.PlotArea.Select
        Selection.ClearFormats
        ActiveChart.PlotArea.Select
        With ActiveChart.Axes(xlCategory)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        With ActiveChart.Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
        End With
        ActiveChart.ChartArea.Select
        With Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .Background = xlAutomatic
        End With
        ActiveChart.PlotArea.Select
        With ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = "ETH pH" & Chr(10) & ExpTitle
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Time (sec)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = _
            "ETH Value"
        End With
        ActiveChart.PlotArea.Select
        Selection.Top = 1
        Selection.Height = 370
        ActiveChart.Axes(xlValue).Select
        With ActiveChart.Axes(xlValue)
            '.MinimumScaleIsAuto = True
            '.MaximumScaleIsAuto = True
            .MinimumScale = GMin
            .MaximumScale = GMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlCustom
            .CrossesAt = -5000
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With
        ActiveChart.Axes(xlCategory).Select
        With ActiveChart.Axes(xlCategory)
            .MinimumScale = 0
            .MaximumScale = xMax
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = xlLinear
            .DisplayUnit = xlNone
        End With

    Range("A1").Select
    ActiveWindow.Zoom = 70
    
    Sheets("Data").Visible = False
    Range("A1:A2,C1:C2,K1:K2,O1:O2,W1:W2,AB1:AB2,AJ1,AJ2").ClearContents
    Range("A36:A37,C36:C37,K36:K37,O36:O37,W36:W37,AB36:AB37,AJ36,AJ37").ClearContents
    Range("A71:A72,C71:C72,K71:K72,O71:O72,W71:W72,AB71:AB72,AJ71,AJ72").ClearContents
    Range("A105:A106,C105:C106,K105:K106,O105:O106,W105:W106,AB105:AB106,AJ105,AJ106").ClearContents
    
    Range("L2:N2") = ExpData

Exit Sub

End Sub
