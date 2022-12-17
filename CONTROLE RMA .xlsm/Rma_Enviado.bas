Attribute VB_Name = "Rma_Enviado"
Sub Envia_Rma_Para_Enviados()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    line_number = Selection.Row
    full_line = "A" & line_number & ":Z" & line_number
    Range(full_line).Copy
    
    worksheet_name = ActiveSheet.name
    file_adress = ("\\nefile\Controle RMA\Eviados\" & worksheet_name & ".xlsx")
    
    Workbooks.Open (file_adress), False, ReadOnly:=False
    Windows(worksheet_name).Activate

    ActiveSheet.Unprotect
    ActiveSheet.Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.PasteSpecial
        
    line_second_worksheet = ActiveCell.Row
         
    'Coloca bordas finas
    Range("A" & line_second_worksheet & ":Z" & line_second_worksheet).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
    End With
        
    'Classificar data mais antiga
    Range("A7").Select
    ActiveSheet.Unprotect
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets(worksheet_name).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(worksheet_name).Sort.SortFields.Add2 Key:=Range("E8:E500"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="dd/mm/yyyy", _
        DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets(worksheet_name).Sort
        .SetRange Range("A7:XFC500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    Windows("CONTROLE RMA").Activate
    Sheets(worksheet_name).Activate
    ActiveSheet.Unprotect
    Rows(line_number).Delete Shift:=xlUp
    
    'protect planilha
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
    AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, AllowInsertingColumns:=True
    
    Workbooks(worksheet_name).Close (True)
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub




