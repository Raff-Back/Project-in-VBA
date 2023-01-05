Attribute VB_Name = "Deleta_rma"
Sub delete_line()
    
    delete_confirm = MsgBox("Este cadastro de RMA será apagado permanentemente, deseja continuar?", vbYesNo, "ATENÇÃO!!!")
    
    If delete_confirm = vbYes Then
    
        ActiveSheet.Unprotect
        
        line_number = Selection.Row
        Rows(line_number).Delete Shift:=xlUp
        
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True
    End If
    
End Sub
