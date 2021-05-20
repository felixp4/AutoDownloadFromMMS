Attribute VB_Name = "Module1"
Sub UploadCSV()
    myPath = ThisWorkbook.Path
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & myPath & "\generate.csv", Destination:= _
        Range("$A$1"))
'        .CommandType = 0
        .Name = "generate"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & myPath & "\exchange.csv", Destination:= _
        Range("$N$1"))
'        .CommandType = 0
        .Name = "exchange"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & myPath & "\supply.csv", Destination:= _
        Range("$AJ$1"))
'        .CommandType = 0
        .Name = "exchange"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & myPath & "\rainbow.csv", _
        Destination:=Range("$AV$1"))
'        .CommandType = 0
        .Name = "rainbow"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    Cells.Replace What:=",0000", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    Cells.Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
    ' --- AZ заповнення колонки DTEK_out -----------------------------------------------------------------
    With ActiveSheet
        For i = 2 To 746
            .Cells(i, 52).Value = .Cells(i, 21).Value + .Cells(i, 50).Value - .Cells(i, 49).Value
        Next i
    End With
    
End Sub

Sub UploadEAMD()
    
    Application.ScreenUpdating = False
    
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H13").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H14").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
'    Windows("UPLOAD_month.xlsm").Activate
'    Application.CutCopyMode = False
'

    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H38").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H39").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
 ' --- ZAP_GEN -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H21").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- UZHUK_GEN -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H26").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("K2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H27").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    
' --- RIVN_00000 -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("O2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H15").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("P2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H16").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- RIVN_01500 -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("Q2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H19").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("R2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H20").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- RIVN_00600 -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("S2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H17").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H18").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- ZAP_00000 -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H22").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("Z2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H23").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- ZAP_TPP -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("V2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H24").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AZ2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- UZHUK_02500 -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H28").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AD2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H29").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- UZHUK_00000 -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AB2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H30").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AA2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H31").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- UZHUK_01700 -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AF2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H32").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AE2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H33").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- UZHUK_TASHL -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H34").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AT2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H35").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

' --- UZHUK_OLEKS -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H36").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("AP2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H37").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
' --- KYM_00000 -------------------------------------------------------------------------------------
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("W2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H40").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
    Windows("UPLOAD_month.xlsm").Activate
    Application.CutCopyMode = False
    Range("X2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("EAMD_NAEK_month.xlsm").Activate
    Range("H41").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
               
    Application.ScreenUpdating = True
    
End Sub
