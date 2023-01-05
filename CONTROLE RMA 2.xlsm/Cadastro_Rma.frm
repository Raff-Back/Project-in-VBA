VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cadastro_Rma 
   Caption         =   "CADASTRO RMA"
   ClientHeight    =   10800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6930
   OleObjectBlob   =   "Cadastro_Rma.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cadastro_Rma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Caixa_NotaDeCompra_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    Select Case KeyAscii
        Case 8  '''Backspace
        Case 48 To 57  '''Numeros de 0 a 9
        Case Else: KeyAscii = 0 '''Ignora os outros caracteres
    End Select
    
End Sub

Private Sub Caixa_NotaDeVenda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    Select Case KeyAscii
        Case 8  '''Backspace
        Case 48 To 57  '''Numeros de 0 a 9
        Case Else: KeyAscii = 0 '''Ignora os outros caracteres
    End Select
    
End Sub

Private Sub Caixa_PrazoDeGarantia_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Caixa_PrazoDeGarantia.MaxLength = 10 '''Permite digitar no máximo 10 caracteres

    Select Case KeyAscii
        Case 8  '''Backspace
        Case 13: SendKeys "{TAB}"  '''Emula o TAB
        Case 48 To 57  '''Numeros de 0 a 9
        If Caixa_PrazoDeGarantia.SelStart = 2 Then Caixa_PrazoDeGarantia.SelText = "/" '''insere barra ao digitar dia
        If Caixa_PrazoDeGarantia.SelStart = 5 Then Caixa_PrazoDeGarantia.SelText = "/" '''Insere barra ao digitar mes
        Case Else: KeyAscii = 0 '''Ignora os outros caracteres
    End Select

End Sub

Private Sub Caixa_PrazoDeGarantia_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsDate(Caixa_PrazoDeGarantia) And Caixa_PrazoDeGarantia <> "" Then '''valida Data
        
        Dim invalid_date As VbMsgBoxResult
        
        invalid_date = MsgBox("Data inválida!", vbInformation, "Atenção!")
        Caixa_PrazoDeGarantia = ""
        Cancel = True
    End If

End Sub


Private Sub Caixa_Quantidade_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8  '''Backspace
        Case 48 To 57  '''numeros de 0 a 9
        Case Else: KeyAscii = 0 '''Ignora os outros caracteres
    End Select
    
End Sub

Private Sub UserForm_Initialize()
    
    '''Coleta a ultima celula contendo dados da coluna B a partir de B2
    end_list = Sheets("DADOS").Range("B2").End(xlDown).Row
    Caixa_NomeFornecedor.RowSource = "DADOS!B2:B" & end_list

End Sub

Private Sub Botao_Cadastrar_Click()
    
    Dim next_line As Integer
    Dim not_selected As VbMsgBoxResult

    Application.ScreenUpdating = False  '''Dessabilita atualização de tela
    
    fornecedor = Caixa_NomeFornecedor
    codigo_produto = Caixa_CodigoDoProduto
    descricao_produto = Caixa_DescricaoDoProduto
    numero_serie = Caixa_NumeroDeSerie
    quantidade = Caixa_Quantidade
    prazo_garantia = Caixa_PrazoDeGarantia
    nota_compra = Caixa_NotaDeCompra
    nota_venda = Caixa_NotaDeVenda
    chave_acesso = Caixa_ChaveDeAcesso
    
    If Caixa_NumeroDeSerie <> "" Then
        forn = Caixa_NomeFornecedor
        
        Dim ns As Double
        ns = WorksheetFunction.CountIf(Sheets(forn).Range("C:C"), Caixa_NumeroDeSerie.Text)
                     
        '''verifica duplicidade no numero de serie
        If ns > 0 Then
            MsgBox "Numero de serie ja foi cadastrado!", vbCritical, "ATENÇÃO!"
            Exit Sub
        End If
    End If
   
    '''Verifica se a caixa fornecedor foi preenchida
    If fornecedor = "" Then
        not_selected = MsgBox("Inserir fornecedor!", vbInformation, "Atenção!")
    Else
        next_line = Sheets(fornecedor).Range("H1048576").End(xlUp).Row + 1
        
        Sheets(fornecedor).Activate
        ActiveSheet.Unprotect
        
        '''Formata celulas para texto e alinha para o centro
        Rows(next_line).Select
        Selection.NumberFormat = "@"
        With Selection
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        '''Formata para data
        Range("E" & next_line).Select
        Selection.NumberFormat = "mm/dd/yyyy"
        
        '''Adicionar codigo do produto
        Sheets(fornecedor).Cells(next_line, 1).Value = codigo_produto
        '''Adicionar descrição do produto
        Sheets(fornecedor).Cells(next_line, 2).Value = descricao_produto
        '''Adicionar numero de série
        Sheets(fornecedor).Cells(next_line, 3).Value = numero_serie
        '''Adcionar Quantidade
        Sheets(fornecedor).Cells(next_line, 4).Value = quantidade
        '''Adcionar Prazo de garantia
        Sheets(fornecedor).Cells(next_line, 5).Value = prazo_garantia
        '''Adcionar Nota de compra
        Sheets(fornecedor).Cells(next_line, 6).Value = nota_compra
        ''Adicionar Nota de venda
        Sheets(fornecedor).Cells(next_line, 7).Value = nota_venda
        '''Adicionar fornecedor
        Sheets(fornecedor).Cells(next_line, 8).Value = fornecedor
        '''Adicionar chave de acesso
        Sheets(fornecedor).Cells(next_line, 9).Value = chave_acesso
    
        '''Adciona bordas finas
        Rows(next_line).Select
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
        
        '''Classifica por data mais antiga
        Range("A7").Select
        ActiveSheet.Unprotect
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Range(Selection, Selection.End(xlDown)).Select
        ActiveWorkbook.Worksheets(fornecedor).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(fornecedor).Sort.SortFields.Add2 Key:=Range("E8:E500"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="dd/mm/yyyy", _
            DataOption:=xlSortTextAsNumbers
        With ActiveWorkbook.Worksheets(fornecedor).Sort
            .SetRange Range("A7:XFC500")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        '''protege planilha
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingColumns:=True
        
        '''Limpa os campos
        Dim c As Control
        Dim Objeto As String, nome As String
        
        With Cadastro_Rma
            For Each c In .Controls
                nome = c.name
                Objeto = VBA.TypeName(c)
                
                If Objeto = "TextBox" Then
                    .Controls(nome).Value = Empty
                End If
            Next c
        End With
    End If

    Application.ScreenUpdating = True
    
 End Sub
