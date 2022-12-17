VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cadastro_Fornecedor 
   Caption         =   "CADASTRO DE FORNECEDOR"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Cadastro_Fornecedor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cadastro_Fornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

   'Mostra nome do usuario ativo
Function UsuarioRede() As String
    Dim GetUserN
    Dim ObjNetwork
    Set ObjNetwork = CreateObject("WScript.Network")
    GetUserN = ObjNetwork.UserName
    UsuarioRede = GetUserN
End Function

Private Sub Botao_Cadastrar_Click()

    
    Application.ScreenUpdating = False  'Desativa atualização de tela
    
    ' Verifica se o nome do fornecedor ja consta na planilha DADOS
    Dim fornecedor As Double
    fornecedor = WorksheetFunction.CountIf(Planilha4.Range("B:B"), Caixa_NomeFornecedor.Text)
    
    If fornecedor > 0 Then
        MsgBox "Fonecedor ja foi cadastrado!", vbCritical, "ATTENTION!"
        Exit Sub
    End If
    
    Dim next_line As Double

    Sheets("DADOS").Unprotect  'Desprotege a planilha "DADOS"
    
    next_line = WorksheetFunction.CountA(Planilha4.Range("B:B")) + 1 'selecionar proxima celula em branco de "DADOS"
    Planilha4.Cells(next_line, 2).Value = Caixa_NomeFornecedor  'Cola o nome do novo fornecedor

    'Colocar os DADOS em ordem alfabética
    ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela3").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela3").Sort.SortFields.Add2 _
        Key:=Columns("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DADOS").ListObjects("Tabela3").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Sheets("DADOS").Protect  'Protege a planilha "DADOS"
    
    'Reexibe e desprotege a planilha
    Sheets("ESTRUTURA").Visible = True
    Sheets("ESTRUTURA").Unprotect
    
    'Criar uma cópia da planilha "ESTRUTURA"
    Sheets("ESTRUTURA").Copy After:=Sheets(ThisWorkbook.Worksheets.Count)
    
    'Pega a ultima planilha, que no caso foi a criada e renomeia com o nome do fornecedor
    Dim newSheet As Worksheet
    Set newSheet = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    newSheet.name = Caixa_NomeFornecedor
    
    'Escreve o nome do fornecedor no título
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 6")).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Caixa_NomeFornecedor
    Selection.Font.Bold = True
    Range("A8").Select
    
    'Oculta planilha e desprotege a planilha "ESTRUTURA"
    Sheets("ESTRUTURA").Visible = False
    Sheets("ESTRUTURA").Protect
    
    'Verifica se ja existe um arquivo com o nome do fornecedor na pasta de enviados, caso não, cria um novo arquivo
    strPath = "C:\Users\" & UsuarioRede & "\Desktop\" & Caixa_NomeFornecedor & ".xlsx"
    If Dir(strPath) = vbNullString Then
        strCheck = False
    Else
        strCheck = True
    End If
    
    If strCheck = False Then
        Dim nome As String
        nome = Caixa_NomeFornecedor
        ActiveSheet.Copy
        With ActiveWorkbook
        .SaveAs "C:\Users\" & UsuarioRede & "\Desktop" & "\" & nome & " enviados.xlsx"
        ActiveWorkbook.Close
        End With
    End If

    'Proteger a nova planilha
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
    AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    AllowFormattingRows:=True, AllowInsertingColumns:=True
    
    'Ativa a atualização de tela e fecha o formualrio de cadastro de fornecedor
    Application.ScreenUpdating = True
    Unload Cadastro_Fornecedor
    
    Dim Pasta_Criada As VbMsgBoxStyle
    
    Pasta_Criada = MsgBox("Novo arquivo """ & Caixa_NomeFornecedor & ".xlsx"", criado em sua área de trabalho!", vbInformation, "Concluído!")

End Sub

Private Sub Caixa_NomeFornecedor_Change()

Caixa_NomeFornecedor.Value = VBA.UCase(Caixa_NomeFornecedor.Text)
'Todas as entradas da caixa de texto ficarão em maiúsculo

End Sub


