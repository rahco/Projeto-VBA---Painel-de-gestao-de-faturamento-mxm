Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False

    Call Bv_get
    Call Bv_adq
    Call Base_tt
    Call Base_tt_tratada
    Call Base_de_resultados
    Call Quadro_resumo_cidade

    Sheets("MACROS").Select
    Range("B7").Select

    Application.ScreenUpdating = True

End Sub

Sub Bv_get()
Attribute Bv_get.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BV - G&T").Range("C2").Value)
    final = Abs(Worksheets("BV - G&T").Range("B2").Value)
 
    Do While atual > final
        Sheets("BV - G&T").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BV - G&T").Range("C2").Value)
        final = Abs(Worksheets("BV - G&T").Range("B2").Value)
    Loop

    Sheets("BV - G&T").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("BV - G&T - M-1 - TT").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BV - G&T").Select
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B4").Select
    Selection.Copy
    Range("C5").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BV - G&T - PARCIAL").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BV - G&T").Select
    Range("C4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "PARCIAIS"
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B4").Select
    Range("J4").Select
    Selection.Copy
    Range("I4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("J5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B4").Select
    
    Application.ScreenUpdating = True

End Sub

Sub Bv_adq()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BV - ADQ").Range("C2").Value)
    final = Abs(Worksheets("BV - ADQ").Range("B2").Value)
 
    Do While atual > final
        Sheets("BV - ADQ").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BV - ADQ").Range("C2").Value)
        final = Abs(Worksheets("BV - ADQ").Range("B2").Value)
    Loop

    Sheets("BV - ADQ").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B4").Select
    Sheets("BV - ADQ - M-1 - TT").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BV - ADQ").Select
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B4").Select
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B4").Select
    Sheets("BV - ADQ - PARCIAL").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BV - ADQ").Select
    Range("C4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B3").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "PARCIAIS"
    Selection.Copy
    Range("C3").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, -1).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B4").Select
    Range("G4").Select
    Selection.Copy
    Range("F4").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("G5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B4").Select
    
    Application.ScreenUpdating = True

End Sub

Sub Base_tt()
Attribute Base_tt.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False

    Sheets("BASE TT").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    Range("B3").Select
    Sheets("BV - G&T").Select
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B4").Select
    Sheets("BASE TT").Select
    Range("B3").Select
    ActiveSheet.Paste
    Range("B3").Select
    Application.CutCopyMode = False
    Sheets("BV - ADQ").Select
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B4").Select
    Sheets("BASE TT").Select
    Range("B3").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("B3").Select
    ActiveSheet.Range("$B$2:$B$100000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Range("B3").Select
    ActiveWorkbook.RefreshAll
    
    Application.ScreenUpdating = True
    
End Sub

Sub Base_tt_tratada()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE TT - TRATADA").Range("C5").Value)
    final = Abs(Worksheets("BASE TT - TRATADA").Range("B5").Value)
 
    Do While atual > final
        Sheets("BASE TT - TRATADA").Select
        Range("B7").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B7").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE TT - TRATADA").Range("C5").Value)
        final = Abs(Worksheets("BASE TT - TRATADA").Range("B5").Value)
    Loop

    Sheets("BASE TT - TRATADA").Select
    Range("B7").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C5").Value > 0 Then
        linhaf = linhai - Range("C5").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C5").Value < 0 Then
        linhaf = linhai + Range("C5").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B7").Select

    Sheets("BASE TT").Select
    Range("B3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B3").Select
    Sheets("BASE TT - TRATADA").Select
    Range("B7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("C8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B7").Select

    Application.ScreenUpdating = True

End Sub

Sub Base_de_resultados()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE DE RESULTADOS").Range("C1").Value)
    final = Abs(Worksheets("BASE DE RESULTADOS").Range("B1").Value)
 
    Do While atual > final
        Sheets("BASE DE RESULTADOS").Select
        Range("B6").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B6").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE DE RESULTADOS").Range("C1").Value)
        final = Abs(Worksheets("BASE DE RESULTADOS").Range("B1").Value)
    Loop

    Sheets("BASE DE RESULTADOS").Select
    Range("B6").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B6").Select
    
    Sheets("BASE TT - TRATADA").Select
    Range("B7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B7").Select
    Sheets("BASE DE RESULTADOS").Select
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("P5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.Columns.AutoFit
    Range("E5").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("E5:E5132"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D5").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("D5:D5132"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("F5").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("F5:F5132"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G5").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("G5:G5132"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("J5").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("J5:J5132"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("K5").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("K5:K5132"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B6").Select
    ActiveWorkbook.RefreshAll
    
    Application.ScreenUpdating = True

End Sub

Sub Quadro_resumo_cidade()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("QUADRO RESUMO | CIDADE").Range("C2").Value)
    final = Abs(Worksheets("QUADRO RESUMO | CIDADE").Range("B2").Value)
 
    Do While atual > final
        Sheets("QUADRO RESUMO | CIDADE").Select
        Range("B7").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B7").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("QUADRO RESUMO | CIDADE").Range("C2").Value)
        final = Abs(Worksheets("QUADRO RESUMO | CIDADE").Range("B2").Value)
    Loop

    Sheets("QUADRO RESUMO | CIDADE").Select
    Range("B7").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B7").Select
    
    Sheets("TD - CIDADE").Select
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B4").Select
    Sheets("QUADRO RESUMO | CIDADE").Select
    Range("B7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C6").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("QUADRO RESUMO | CIDADE").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("QUADRO RESUMO | CIDADE").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("C6"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("QUADRO RESUMO | CIDADE").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B7").Select
    
    Application.ScreenUpdating = True

End Sub

Sub Arquivo_de_envio()

    Application.ScreenUpdating = False
    
    ActiveWorkbook.Save
    
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("MACROS").Range("C16").Value & " - Gestão de Faturamento Bruto MxM - Dados até dia " & Worksheets("MACROS").Range("C17").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    
    Sheets("TD - RESUMO").Select
    ActiveWindow.SelectedSheets.Visible = False
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B7").Select
    Application.CutCopyMode = False
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    ActiveWindow.DisplayHeadings = False
    Sheets("QUADRO RESUMO | HC - CAP").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B7").Select
    Application.CutCopyMode = False
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    ActiveWindow.DisplayHeadings = False
    Sheets("QUADRO RESUMO | CIDADE").Select
    Cells.Select
    Selection.Copy
    Range("B7").Select
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("QUADRO RESUMO | CIDADE").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("QUADRO RESUMO | CIDADE").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("C6"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("QUADRO RESUMO | CIDADE").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B7").Select
    Application.CutCopyMode = False
    Range("B2:C2").Select
    Selection.Clear
    Range("C7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B7").Select
    Selection.AutoFilter
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE DE RESULTADOS").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False
    Range("B1:C1").Select
    Selection.Clear
    Range("B6").Select
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets(Array("TD - HC", "TD - CIDADE", "GRÁFICOS")).Select
    Sheets("TD - HC").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=-10
    Sheets(Array("BV - G&T - M-1 - TT", "BV - ADQ - M-1 - TT", "BV - G&T", "BV - ADQ", _
        "BASE TT", "TD - G&T", "TD - ADQ", "BASE TT - TRATADA", "TD - HC", "TD - CIDADE", _
        "GRÁFICOS")).Select
    Sheets("TD - HC").Activate
    ActiveWindow.ScrollWorkbookTabs Sheets:=-4
    Sheets(Array("MACROS", "BD - CLI", "BV - G&T - PARCIAL", "BV - ADQ - PARCIAL", _
        "ÁREA-HC", "BV - G&T - M-1 - TT", "BV - ADQ - M-1 - TT", "BV - G&T", "BV - ADQ", _
        "BASE TT", "TD - G&T", "TD - ADQ", "BASE TT - TRATADA", "TD - HC", "TD - CIDADE", _
        "GRÁFICOS")).Select
    Sheets("MACROS").Activate
    ActiveWindow.SelectedSheets.Delete
    Sheets("QUADRO RESUMO").Select
    ActiveWorkbook.Save

    
    Application.ScreenUpdating = True

End Sub
