VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoverFornecedor 
   Caption         =   "REMOVER FORNECEDOR"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4665
   OleObjectBlob   =   "RemoverFornecedor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoverFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Botao_Remover_Click()

    fornecedor = Caixa_SelecionarFornecedor
    Dim exclusion_confirm As VbMsgBoxResult
    Dim concluded As VbMsgBoxResult
    
    exclusion_confirm = MsgBox("Todos os dados contidos na planilha serão apagados, Deseja continuar?", vbYesNo, "ATENÇÃO!!!")
    If exclusion_confirm = vbYes Then
    
        Application.ScreenUpdating = False
        
        Dim provider_name As String
        
        provider_name = RemoverFornecedor!Caixa_SelecionarFornecedor.Value
        
        Sheets("DADOS").Unprotect
        Range("DADOS!B:B").Cells.Find(provider_name).Delete
        Sheets("DADOS").Protect
        
        Application.DisplayAlerts = False
        Sheets(provider_name).Delete
        Application.DisplayAlerts = True
        
        concluded = MsgBox("Concluído!", vbInformation, "Exclusão de fornecedor.")
 
        Application.ScreenUpdating = True
    End If
    
End Sub


Private Sub UserForm_Initialize()

    end_list = Sheets("DADOS").Range("B2").End(xlDown).Row
    Caixa_SelecionarFornecedor.RowSource = "DADOS!B2:B" & end_list
    
End Sub
