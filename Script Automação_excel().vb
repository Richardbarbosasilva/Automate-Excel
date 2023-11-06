Sub AutomaçãoFoda()
    
    Dim FolderPathOld As String
    Dim Filename As String
    Dim FolderPathNew As String
    Dim wb As Workbook
    Dim FileNameFormatation As String
    
    FolderPathOld = "C:\Users\Richard.silva\Downloads\notacsv\"
    FolderPathNew = "C:\Users\Richard.silva\Downloads\notaxlsx\"
    
    ' Identificando o arquivo
    Filename = Dir(FolderPathOld & "*.csv")
    
    Application.ScreenUpdating = False
    
    Do While Filename <> ""
        ' Usando Workbooks.OpenText para importar o CSV diretamente como uma nova pasta de trabalho
        Workbooks.OpenText Filename:=FolderPathOld & Filename, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, Comma:=False, Space:=False, Other:=False, Local:=True
        
        ' Atualizando o nome da pasta de trabalho para remover a extensão .csv
        ActiveWorkbook.SaveAs Filename:=FolderPathNew & Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & ".xlsx", FileFormat:=51 ' 51 é o formato para .xlsx
        ActiveWorkbook.Close SaveChanges:=False
        
        ' Obter o próximo arquivo .csv na pasta
        Filename = Dir
    Loop
    
    FileNameFormatation = Dir(FolderPathNew & "*.xlsx")
    
    Do While FileNameFormatation <> ""
        ' Abrir cada arquivo .xlsx
        Set wb = Workbooks.Open(FolderPathNew & FileNameFormatation)
        
        ' Ativar a atualização da tela para a execução da macro
        Application.ScreenUpdating = False
        
        ' Executar a macro na pasta de trabalho
        Columns("A:A").EntireColumn.AutoFit
        Columns("B:B").EntireColumn.AutoFit
        Columns("C:C").EntireColumn.AutoFit
        Columns("D:D").EntireColumn.AutoFit
        Columns("E:E").EntireColumn.AutoFit
        Columns("F:F").EntireColumn.AutoFit
        Columns("G:G").EntireColumn.AutoFit
        Range("A1:G1").Select
        Range("G1").Activate
        Selection.Font.Bold = True
        Columns("A:G").Select
        Range("G18").Activate
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
        Columns("G:G").Select
        Selection.Style = "Currency"
        Selection.ColumnWidth = 15.5
        Columns("C:C").Select
        Selection.Style = "Currency"
        Columns("D:D").Select
        Selection.Style = "Currency"
        Selection.ColumnWidth = 80
        Columns("C:C").Select
        Selection.Style = "Currency"
        
        ' Salvar as alterações na pasta de trabalho
        wb.Save
        
        ' Fechar a pasta de trabalho sem perguntar para salvar novamente
        wb.Close SaveChanges:=False
        
        ' Obter o próximo arquivo .xlsx na pasta
        FileNameFormatation = Dir
    Loop
    
    ' Ativar as atualizações de tela novamente
    Application.ScreenUpdating = True
End Sub

