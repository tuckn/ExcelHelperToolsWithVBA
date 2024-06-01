Attribute VB_Name = "Scratchs"
Option Explicit

Private Sub DebugPrintShapeNamesAndTexts()
    Dim shp As Shape
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEETNAME_CANVAS)

    For Each shp In ws.Shapes
        If Not shp.TextFrame2.HasText = msoFalse Then
            Debug.Print shp.Name & ": " & shp.TextFrame2.TextRange.Text
        End If
    Next shp
End Sub

' 図形に書かれたテキストを取得し、配置位置を考慮して並べて表示
Private Sub DebugPrintShapeTextsInOrder()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEETNAME_CANVAS)
    
    Dim shp As Shape
    Dim i As Long
    Dim j As Long
    Dim tempY As Single
    Dim tempText As String
    Dim shapeCount As Long
    shapeCount = ws.Shapes.Count
    
    ' 配列をShapeの数だけ用意
    Dim shapeInfo() As Variant
    ReDim shapeInfo(1 To shapeCount, 1 To 2)
    
    ' ShapeのテキストとY座標を配列に格納
    i = 1
    For Each shp In ws.Shapes
        If Not shp.TextFrame2.HasText = msoFalse Then
            shapeInfo(i, 1) = shp.Top ' Y軸座標
            shapeInfo(i, 2) = shp.TextFrame2.TextRange.Text ' テキスト
            i = i + 1
        End If
    Next shp
    
    ' Y軸座標で配列をソート（単純なバブルソート）
    For i = 1 To shapeCount - 1
        For j = i + 1 To shapeCount
            If shapeInfo(i, 1) > shapeInfo(j, 1) Then
                ' Y軸座標で交換
                tempY = shapeInfo(i, 1)
                shapeInfo(i, 1) = shapeInfo(j, 1)
                shapeInfo(j, 1) = tempY
                ' テキストで交換
                tempText = shapeInfo(i, 2)
                shapeInfo(i, 2) = shapeInfo(j, 2)
                shapeInfo(j, 2) = tempText
            End If
        Next j
    Next i
    
    ' ソートされたテキストを出力
    For i = 1 To shapeCount
        If shapeInfo(i, 2) <> "" Then
            Debug.Print shapeInfo(i, 2)
        End If
    Next i
End Sub

' すべての図形の情報を出力
Private Sub DebugPrintShapesInfo()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim shapeTypeName As String

    Set ws = ThisWorkbook.Sheets(SHEETNAME_CANVAS)
    
    For Each shp In ws.Shapes
        Debug.Print vbCrLf & "Name: " & shp.Name
        Debug.Print "Type: " & shp.Type
        
        ' AutoShape or not
        If shp.Type = msoAutoShape Then ' AutoShape
            Debug.Print "TypeName: AutoShape"
            shapeTypeName = ConvertAutoShapeTypeNumberToName(shp.autoShapeType)
            Debug.Print "AutoShapeType: " & shp.autoShapeType & " (" & shapeTypeName & ")"
        Else
            shapeTypeName = ConvertShapeTypeNumberToName(shp.Type)
            Debug.Print "TypeName: " & shapeTypeName
        End If
        
        Debug.Print "Connector: " & shp.Connector
        
        If shp.Connector Then
            Select Case shp.connectorFormat.Type
                Case msoConnectorStraight
                    shapeTypeName = "Straight Connector"
                Case msoConnectorElbow
                    shapeTypeName = "Elbow Connector"
                Case msoConnectorCurve
                    shapeTypeName = "Curved Connector"
                Case Else
                    shapeTypeName = "Other Connector"
            End Select
            
            Debug.Print "ConnectorFormat: " & shp.connectorFormat.Type & " (" & shapeTypeName & ")"
        End If
    
    Next shp
End Sub

Private Sub DebugPrintConnectorShapes()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim connBegin As Shape
    Dim connEnd As Shape

    Set ws = ThisWorkbook.Sheets(SHEETNAME_CANVAS)

    ' コネクタの情報を収集
    For Each shp In ws.Shapes
        If shp.Connector Then
            ' コネクタの始点と終点の図形を取得
            On Error Resume Next
            Set connBegin = shp.connectorFormat.BeginConnectedShape
            Set connEnd = shp.connectorFormat.EndConnectedShape
            On Error GoTo 0
            
            ' コネクタに接続されている図形の情報を出力
            Debug.Print "Connector Name: " & shp.Name
            If Not connBegin Is Nothing Then
                Debug.Print "  Begin Connected Shape: " & connBegin.Name
            Else
                Debug.Print "  Begin Connected Shape: None"
            End If
            If Not connEnd Is Nothing Then
                Debug.Print "  End Connected Shape: " & connEnd.Name
            Else
                Debug.Print "  End Connected Shape: None"
            End If
        End If
    Next shp
End Sub

Private Sub DebugPrintConnectorsCsvString()
    Dim csvOutput As String
    csvOutput = MakeConnectorShapesCsvString(SHEETNAME_CANVAS)
    Debug.Print csvOutput
End Sub

Private Sub DebugPrintShapesCsvString()
    Dim csvOutput As String
    csvOutput = MakeShapesCsvString(SHEETNAME_CANVAS)
    Debug.Print csvOutput
End Sub

