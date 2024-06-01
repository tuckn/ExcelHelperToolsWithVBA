Attribute VB_Name = "Main"
Option Explicit

Sub UpdateConnectionsTable()
    Dim connectorsCsvOutput As String
    connectorsCsvOutput = MakeConnectorShapesCsvString(SHEETNAME_CANVAS)
    Call InsertCsvAsTable(SHEETNAME_CONNECTORS, TABLENAME_CONNECTOR, connectorsCsvOutput)
End Sub

Sub UpdateShapesTable()
    Dim shapesCsvOutput As String
    shapesCsvOutput = MakeShapesCsvString(SHEETNAME_CANVAS)
    Call InsertCsvAsTable(SHEETNAME_SHAPES, TABLENAME_SHAPES, shapesCsvOutput)
    Call BackupTableData(SHEETNAME_SHAPES, TABLENAME_SHAPES, SHEETNAME_BACKUP_SHAPES)
End Sub

Private Sub BackupShapesTable()
    Call BackupTableData(SHEETNAME_SHAPES, TABLENAME_SHAPES, SHEETNAME_BACKUP_SHAPES)
End Sub

Sub UpdateTables()
    ' Performance optimization settings
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Call UpdateConnectionsTable
    Call UpdateShapesTable
    
    ' Clear the status bar
    Application.StatusBar = False
    ' Restore performance settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub UpdateCanvasShapes()
    ' Performance optimization settings
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Call UpdateCanvasShapesFromConnectorsTable @TODO
    Call UpdateCanvasShapesFromShapesTable
    Call SelectCanvasShapesFromTables
    
    ' Clear the status bar
    Application.StatusBar = False
    ' Restore performance settings
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub UpdateCanvasShapesFromShapesTable()
    Dim wsCanvas As Worksheet
    Dim wsShapes As Worksheet
    Dim tblShapes As ListObject
    Dim shp As Shape
    Dim shapeRow As ListRow
    Dim shapeID As Long
    Dim totalShapes As Long
    Dim processedShapes As Long
    Dim shapeFound As Boolean
    Dim i As Integer

    ' Set sheets
    Set wsCanvas = ThisWorkbook.Sheets(SHEETNAME_CANVAS)
    Set wsShapes = ThisWorkbook.Sheets(SHEETNAME_SHAPES)
    ' Set tables
    Set tblShapes = wsShapes.ListObjects(TABLENAME_SHAPES)

    ' Check all shapes in the Canvas sheet and delete shapes not present in the Shapes table (excluding connectors)
    totalShapes = wsCanvas.Shapes.Count ' Count the total number of shapes
    processedShapes = 0

    For Each shp In wsCanvas.Shapes
        If Not shp.Connector Then
            shapeID = shp.ID
            shapeFound = False

            For Each shapeRow In tblShapes.ListRows
                If shapeRow.Range.Cells(1, colShpID).Value = shapeID Then
                    shapeFound = True
                    Exit For
                End If
            Next shapeRow

            If Not shapeFound Then
                shp.Delete
            End If

            processedShapes = processedShapes + 1
            If processedShapes Mod 10 = 0 Then
                Application.StatusBar = "Processing Canvas shapes: " & processedShapes & " of " & totalShapes
                DoEvents
            End If
        End If
    Next shp

    totalShapes = wsCanvas.Shapes.Count
    processedShapes = 0
    i = 1

    ' Reflect changes from the Shapes table to the Canvas sheet
    For Each shapeRow In tblShapes.ListRows
        shapeID = shapeRow.Range.Cells(1, colShpID).Value
        Set shp = Nothing

        ' Find the shape in the Canvas sheet by ID
        For Each shp In wsCanvas.Shapes
            If shp.ID = shapeID Then
                Exit For
            End If
        Next shp

        If shp Is Nothing Then
            MsgBox "Shape with ID " & shapeID & " does not exist in Canvas sheet.", vbCritical
            Exit Sub
        End If

        ' Update only elements that have changed
        If shapeRow.Range.Cells(1, colShpName).Value <> shp.Name Then shp.Name = shapeRow.Range.Cells(1, colShpName).Value
        If shapeRow.Range.Cells(1, colShpForeColorRGB).Value <> ConvertForeColorRgbToHexString(shp.Fill.foreColor.RGB) Then shp.Fill.foreColor.RGB = ConvertRGBFromHex(shapeRow.Range.Cells(1, colShpForeColorRGB).Value)

        ' Update properties only if the values are numeric
        If IsNumeric(shapeRow.Range.Cells(1, colShpTop).Value) Then
            If shapeRow.Range.Cells(1, colShpTop).Value <> shp.Top Then shp.Top = shapeRow.Range.Cells(1, colShpTop).Value
        End If
        If IsNumeric(shapeRow.Range.Cells(1, colShpLeft).Value) Then
            If shapeRow.Range.Cells(1, colShpLeft).Value <> shp.Left Then shp.Left = shapeRow.Range.Cells(1, colShpLeft).Value
        End If
        If IsNumeric(shapeRow.Range.Cells(1, colShpHeight).Value) Then
            If shapeRow.Range.Cells(1, colShpHeight).Value <> shp.Height Then shp.Height = shapeRow.Range.Cells(1, colShpHeight).Value
        End If
        If IsNumeric(shapeRow.Range.Cells(1, colShpWidth).Value) Then
            If shapeRow.Range.Cells(1, colShpWidth).Value <> shp.Width Then shp.Width = shapeRow.Range.Cells(1, colShpWidth).Value
        End If

        ' Update text if shape can contain text
        If shp.TextFrame2.HasText Then
            If shapeRow.Range.Cells(1, colShpText).Value <> shp.TextFrame2.TextRange.Text Then
                shp.TextFrame2.TextRange.Text = shapeRow.Range.Cells(1, colShpText).Value
            End If
        End If

        ' ZOrder changes are commented out because ZOrderPosition is read-only
        ' If shapeRow.Range.Cells(1, colShpZOrderPosition).Value <> shp.ZOrderPosition Then
        '     Select Case shapeRow.Range.Cells(1, colShpZOrderPosition).Value
        '         Case Is < shp.ZOrderPosition
        '             Do While shp.ZOrderPosition > shapeRow.Range.Cells(1, colShpZOrderPosition).Value
        '                 shp.ZOrder msoSendBackward
        '             Loop
        '         Case Is > shp.ZOrderPosition
        '             Do While shp.ZOrderPosition < shapeRow.Range.Cells(1, colShpZOrderPosition).Value
        '                 shp.ZOrder msoBringForward
        '             Loop
        '     End Select
        ' End If

        processedShapes = processedShapes + 1
        If processedShapes Mod 10 = 0 Then
            Application.StatusBar = "Processing Shapes table: " & processedShapes & " of " & totalShapes
            DoEvents
        End If
    Next shapeRow

    MsgBox "Canvas shapes have been updated from Shapes table.", vbInformation
End Sub


Sub SelectCanvasShapesFromTables()
    Dim wsCanvas As Worksheet
    Dim wsConnectors As Worksheet
    Dim wsShapes As Worksheet
    Dim tblConns As ListObject
    Dim tblShapes As ListObject
    Dim shp As Shape
    Dim connRow As ListRow
    Dim shapeRow As ListRow
    Dim shapeID As Long
    Dim selectedRowCount As Long
    Dim selectedShapeNames() As String
    Dim i As Integer

    ' Set sheets
    Set wsCanvas = ThisWorkbook.Sheets(SHEETNAME_CANVAS)
    Set wsConnectors = ThisWorkbook.Sheets(SHEETNAME_CONNECTORS)
    Set wsShapes = ThisWorkbook.Sheets(SHEETNAME_SHAPES)
    ' Set tables
    Set tblConns = wsConnectors.ListObjects(TABLENAME_CONNECTOR)
    Set tblShapes = wsShapes.ListObjects(TABLENAME_SHAPES)

    ' Count the number of rows with True in the Selected column
    selectedRowCount = 0

    For Each shapeRow In tblConns.ListRows
        If shapeRow.Range.Cells(1, colConnSelected).Value = True _
                Or shapeRow.Range.Cells(1, colConnSelected).Value = "True" Then
            selectedRowCount = selectedRowCount + 1
        End If
    Next shapeRow

    For Each shapeRow In tblShapes.ListRows
        If shapeRow.Range.Cells(1, colShpSelected).Value = True _
                Or shapeRow.Range.Cells(1, colShpSelected).Value = "True" Then
            selectedRowCount = selectedRowCount + 1
        End If
    Next shapeRow

    ' Resize the array based on the number of selected shapes
    If Not selectedRowCount = 0 Then
        ReDim selectedShapeNames(1 To selectedRowCount)
    End If

    i = 1

    ' Find the shape in the Canvas sheet by ID on the Connector table
    For Each connRow In tblConns.ListRows
        shapeID = connRow.Range.Cells(1, colConnID).Value
        Set shp = Nothing

        For Each shp In wsCanvas.Shapes
            If shp.ID = shapeID Then
                Exit For
            End If
        Next shp

        If shp Is Nothing Then
            MsgBox "Shape with ID " & shapeID & " does not exist in Canvas sheet.", vbCritical
            Exit Sub
        End If

        ' Collect selected shape names
        If connRow.Range.Cells(1, colConnSelected).Value Then
            selectedShapeNames(i) = shp.Name
            i = i + 1
        End If
    Next connRow

    ' Find the shape in the Canvas sheet by ID on the Shapes table
    For Each shapeRow In tblShapes.ListRows
        shapeID = shapeRow.Range.Cells(1, colShpID).Value
        Set shp = Nothing

        For Each shp In wsCanvas.Shapes
            If shp.ID = shapeID Then
                Exit For
            End If
        Next shp

        If shp Is Nothing Then
            MsgBox "Shape with ID " & shapeID & " does not exist in Canvas sheet.", vbCritical
            Exit Sub
        End If

        ' Collect selected shape names
        If shapeRow.Range.Cells(1, colShpSelected).Value Then
            selectedShapeNames(i) = shp.Name
            i = i + 1
        End If
    Next shapeRow

    ' Deselect any selected shapes
    wsCanvas.Activate
    wsCanvas.Cells(1, 1).Select

    ' Select multiple shapes
    If Not selectedRowCount = 0 Then
        wsCanvas.Shapes.Range(selectedShapeNames).Select
    End If

    MsgBox "Shapes have been selected from tables.", vbInformation
End Sub



