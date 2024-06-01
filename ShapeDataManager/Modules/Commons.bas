Attribute VB_Name = "Commons"
'Functions that can be reused in other Excel workbooks
Option Explicit

Public Function ConvertForeColorRgbToRgbString(ByVal rgbValue As Long) As String
    Dim r As Long
    Dim g As Long
    Dim b As Long

    r = rgbValue Mod 256
    g = (rgbValue \ 256) Mod 256
    b = (rgbValue \ 65536) Mod 256

    ConvertForeColorRgbToRgbString = "RGB(" & r & "," & g & "," & b & ")"
End Function

Public Function ConvertForeColorRgbToHexString(ByVal rgbValue As Long) As String
    Dim r As Long
    Dim g As Long
    Dim b As Long

    r = rgbValue Mod 256
    g = (rgbValue \ 256) Mod 256
    b = (rgbValue \ 65536) Mod 256

    ConvertForeColorRgbToHexString = "#" & Right("0" & hex(r), 2) & Right("0" & hex(g), 2) & Right("0" & hex(b), 2)
End Function

Public Function ConvertArrowheadStyle(ByVal arrowheadStyle As Long) As String
    Dim arrowheadName As String

    Select Case arrowheadStyle
        Case msoArrowheadDiamond ' Diamond-shaped
            arrowheadName = "ArrowheadDiamond"
        Case msoArrowheadNone ' No arrowhead
            arrowheadName = "ArrowheadNone"
        Case msoArrowheadOpen ' Open
            arrowheadName = "ArrowheadOpen"
        Case msoArrowheadOval ' Oval-shaped
            arrowheadName = "ArrowheadOval"
        Case msoArrowheadStealth ' Stealth-shaped
            arrowheadName = "ArrowheadStealth"
        Case msoArrowheadStyleMixed ' Return value only; indicates a combination of the other states.
            arrowheadName = "ArrowheadStyleMixed"
        Case msoArrowheadTriangle ' Triangular
            arrowheadName = "ArrowheadTriangle"
        Case Else
            arrowheadName = "Unknown"
    End Select
    
    ConvertArrowheadStyle = arrowheadName
End Function

Public Function ConvertShapeTypeNumberToName(ByVal shapeType As Long) As String
    Dim shapeTypeName As String

    Select Case shapeType
        Case mso3DModel ' 3D model
            shapeTypeName = "3DModel"
        Case msoAutoShape ' AutoShape
            shapeTypeName = "AutoShape"
        Case msoCallout ' Callout
            shapeTypeName = "Callout"
        Case msoCanvas ' Canvas
            shapeTypeName = "Canvas"
        Case msoChart ' Chart
            shapeTypeName = "Chart"
        Case msoComment ' Comment
            shapeTypeName = "Comment"
        Case msoContentApp ' Content Office Add-in
            shapeTypeName = "ContentApp"
        Case msoDiagram ' Diagram
            shapeTypeName = "Diagram"
        Case msoEmbeddedOLEObject ' Embedded OLE object
            shapeTypeName = "EmbeddedOLEObject"
        Case msoFormControl ' Form control
            shapeTypeName = "FormControl"
        Case msoFreeform ' Freeform
            shapeTypeName = "Freeform"
        Case msoGraphic ' Graphic
            shapeTypeName = "Graphic"
        Case msoGroup ' Group
            shapeTypeName = "Group"
'        Case msoIgxGraphic ' SmartArt graphic
'            shapeTypeName = "IgxGraphic"
        Case msoInk ' Ink
            shapeTypeName = "Ink"
        Case msoInkComment ' Ink comment
            shapeTypeName = "InkComment"
        Case msoLine ' Line
            shapeTypeName = "Line"
        Case msoLinked3DModel ' Linked 3D model
            shapeTypeName = "Linked3DModel"
        Case msoLinkedGraphic ' Linked graphic
            shapeTypeName = "LinkedGraphic"
        Case msoLinkedOLEObject ' Linked OLE object
            shapeTypeName = "LinkedOLEObject"
        Case msoLinkedPicture ' Linked picture
            shapeTypeName = "LinkedPicture"
        Case msoMedia ' Media
            shapeTypeName = "Media"
        Case msoOLEControlObject ' OLE control object
            shapeTypeName = "OLEControlObject"
        Case msoPicture ' Picture
            shapeTypeName = "Picture"
        Case msoPlaceholder ' Placeholder
            shapeTypeName = "Placeholder"
        Case msoScriptAnchor ' Script anchor
            shapeTypeName = "ScriptAnchor"
        Case msoShapeTypeMixed ' Mixed shape?type
            shapeTypeName = "msoShapeTypeMixed"
        Case msoSlicer ' Slicer
            shapeTypeName = "Slicer"
        Case msoTable ' Table
            shapeTypeName = "Table"
        Case msoTextBox ' Text box
            shapeTypeName = "TextBox"
        Case msoTextEffect ' Text effect
            shapeTypeName = "TextEffect"
        Case msoWebVideo ' Web video
            shapeTypeName = "WebVideo"
        Case Else
            shapeTypeName = "Unknown"
    End Select
    
    ConvertShapeTypeNumberToName = shapeTypeName
End Function

Public Function ConvertAutoShapeTypeNumberToName(ByVal autoShapeType As Long) As String
    Dim autoShapeTypeName As String
    
    ' https://learn.microsoft.com/en-us/office/vba/api/office.msoautoshapetype
    Select Case autoShapeType
        Case msoShape10pointStar ' 10-point star
            autoShapeTypeName = "Shape10pointStar"
        Case msoShape12pointStar ' 12-point star
            autoShapeTypeName = "Shape12pointStar"
        Case msoShape16pointStar ' 16-point star
            autoShapeTypeName = "Shape16pointStar"
        Case msoShape24pointStar ' 24-point star
            autoShapeTypeName = "Shape24pointStar"
        Case msoShape32pointStar ' 32-point star
            autoShapeTypeName = "Shape32pointStar"
        Case msoShape4pointStar ' 4-point star
            autoShapeTypeName = "Shape4pointStar"
        Case msoShape5pointStar ' 5-point star
            autoShapeTypeName = "Shape5pointStar"
        Case msoShape6pointStar ' 6-point star
            autoShapeTypeName = "Shape6pointStar"
        Case msoShape7pointStar ' 7-point star
            autoShapeTypeName = "Shape7pointStar"
        Case msoShape8pointStar ' 8-point star
            autoShapeTypeName = "Shape8pointStar"
        Case msoShapeActionButtonBackorPrevious ' Back?or?Previous?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonBackorPrevious"
        Case msoShapeActionButtonBeginning ' Beginning?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonBeginning"
        Case msoShapeActionButtonCustom ' Button with no default picture or text. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonCustom"
        Case msoShapeActionButtonDocument ' Document?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonDocument"
        Case msoShapeActionButtonEnd ' End?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonEnd"
        Case msoShapeActionButtonForwardorNext ' Forward?or?Next?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonForwardorNext"
        Case msoShapeActionButtonHelp ' Help?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonHelp"
        Case msoShapeActionButtonHome ' Home?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonHome"
        Case msoShapeActionButtonInformation ' Information?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonInformation"
        Case msoShapeActionButtonMovie ' Movie?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonMovie"
        Case msoShapeActionButtonReturn ' Return?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonReturn"
        Case msoShapeActionButtonSound ' Sound?button. Supports mouse-click and mouse-over actions.
            autoShapeTypeName = "ShapeActionButtonSound"
        Case msoShapeArc ' Arc
            autoShapeTypeName = "ShapeArc"
        Case msoShapeBalloon ' Balloon
            autoShapeTypeName = "ShapeBalloon"
        Case msoShapeBentArrow ' Block arrow that follows a curved 90-degree angle.
            autoShapeTypeName = "ShapeBentArrow"
        Case msoShapeBentUpArrow ' Block arrow that follows a sharp 90-degree angle. Points up by default.
            autoShapeTypeName = "ShapeBentUpArrow"
        Case msoShapeBevel ' Bevel
            autoShapeTypeName = "ShapeBevel"
        Case msoShapeBlockArc ' Block arc
            autoShapeTypeName = "ShapeBlockArc"
        Case msoShapeCan ' Can
            autoShapeTypeName = "ShapeCan"
        Case msoShapeChartPlus ' Square divided vertically and horizontally into four quarters
            autoShapeTypeName = "ShapeChartPlus"
        Case msoShapeChartStar ' Square divided into six parts along vertical and diagonal lines
            autoShapeTypeName = "ShapeChartStar"
        Case msoShapeChartX ' Square divided into four parts along diagonal lines
            autoShapeTypeName = "ShapeChartX"
        Case msoShapeChevron ' Chevron
            autoShapeTypeName = "ShapeChevron"
        Case msoShapeChord ' Circle with a line connecting two points on the perimeter through the interior of the circle; a circle with a chord
            autoShapeTypeName = "ShapeChord"
        Case msoShapeCircularArrow ' Block arrow that follows a curved 180-degree angle
            autoShapeTypeName = "ShapeCircularArrow"
        Case msoShapeCloud ' Cloud shape
            autoShapeTypeName = "ShapeCloud"
        Case msoShapeCloudCallout ' Cloud callout
            autoShapeTypeName = "ShapeCloudCallout"
        Case msoShapeCorner ' Rectangle with rectangular-shaped hole.
            autoShapeTypeName = "ShapeCorner"
        Case msoShapeCornerTabs ' Four right triangles aligning along a rectangular path; four 'snipped' corners.
            autoShapeTypeName = "ShapeCornerTabs"
        Case msoShapeCross ' Cross
            autoShapeTypeName = "ShapeCross"
        Case msoShapeCube ' Cube
            autoShapeTypeName = "ShapeCube"
        Case msoShapeCurvedDownArrow ' Block arrow that curves down
            autoShapeTypeName = "ShapeCurvedDownArrow"
        Case msoShapeCurvedDownRibbon ' Ribbon banner that curves down
            autoShapeTypeName = "ShapeCurvedDownRibbon"
        Case msoShapeCurvedLeftArrow ' Block arrow that curves left
            autoShapeTypeName = "ShapeCurvedLeftArrow"
        Case msoShapeCurvedRightArrow ' Block arrow that curves right
            autoShapeTypeName = "ShapeCurvedRightArrow"
        Case msoShapeCurvedUpArrow ' Block arrow that curves up
            autoShapeTypeName = "ShapeCurvedUpArrow"
        Case msoShapeCurvedUpRibbon ' Ribbon banner that curves up
            autoShapeTypeName = "ShapeCurvedUpRibbon"
        Case msoShapeDecagon ' Decagon
            autoShapeTypeName = "ShapeDecagon"
        Case msoShapeDiagonalStripe ' Rectangle with two triangles-shapes removed; a diagonal stripe
            autoShapeTypeName = "ShapeDiagonalStripe"
        Case msoShapeDiamond ' Diamond
            autoShapeTypeName = "ShapeDiamond"
        Case msoShapeDodecagon ' Dodecagon
            autoShapeTypeName = "ShapeDodecagon"
        Case msoShapeDonut ' Donut
            autoShapeTypeName = "ShapeDonut"
        Case msoShapeDoubleBrace ' Double brace
            autoShapeTypeName = "ShapeDoubleBrace"
        Case msoShapeDoubleBracket ' Double bracket
            autoShapeTypeName = "ShapeDoubleBracket"
        Case msoShapeDoubleWave ' Double wave
            autoShapeTypeName = "ShapeDoubleWave"
        Case msoShapeDownArrow ' Block arrow that points down
            autoShapeTypeName = "ShapeDownArrow"
        Case msoShapeDownArrowCallout ' Callout with arrow that points down
            autoShapeTypeName = "ShapeDownArrowCallout"
        Case msoShapeDownRibbon ' Ribbon banner with center area below ribbon ends
            autoShapeTypeName = "ShapeDownRibbon"
        Case msoShapeExplosion1 ' Explosion
            autoShapeTypeName = "ShapeExplosion1"
        Case msoShapeExplosion2 ' Explosion
            autoShapeTypeName = "ShapeExplosion2"
        Case msoShapeFlowchartAlternateProcess ' Alternate process flowchart symbol
            autoShapeTypeName = "ShapeFlowchartAlternateProcess"
        Case msoShapeFlowchartCard ' Card flowchart symbol
            autoShapeTypeName = "ShapeFlowchartCard"
        Case msoShapeFlowchartCollate ' Collate flowchart symbol
            autoShapeTypeName = "ShapeFlowchartCollate"
        Case msoShapeFlowchartConnector ' Connector flowchart symbol
            autoShapeTypeName = "ShapeFlowchartConnector"
        Case msoShapeFlowchartData ' Data flowchart symbol
            autoShapeTypeName = "ShapeFlowchartData"
        Case msoShapeFlowchartDecision ' Decision flowchart symbol
            autoShapeTypeName = "ShapeFlowchartDecision"
        Case msoShapeFlowchartDelay ' Delay flowchart symbol
            autoShapeTypeName = "ShapeFlowchartDelay"
        Case msoShapeFlowchartDirectAccessStorage ' Direct access storage flowchart symbol
            autoShapeTypeName = "ShapeFlowchartDirectAccessStorage"
        Case msoShapeFlowchartDisplay ' Display flowchart symbol
            autoShapeTypeName = "ShapeFlowchartDisplay"
        Case msoShapeFlowchartDocument ' Document flowchart symbol
            autoShapeTypeName = "ShapeFlowchartDocument"
        Case msoShapeFlowchartExtract ' Extract flowchart symbol
            autoShapeTypeName = "ShapeFlowchartExtract"
        Case msoShapeFlowchartInternalStorage ' Internal storage flowchart symbol
            autoShapeTypeName = "ShapeFlowchartInternalStorage"
        Case msoShapeFlowchartMagneticDisk ' Magnetic disk flowchart symbol
            autoShapeTypeName = "ShapeFlowchartMagneticDisk"
        Case msoShapeFlowchartManualInput ' Manual input flowchart symbol
            autoShapeTypeName = "ShapeFlowchartManualInput"
        Case msoShapeFlowchartManualOperation ' Manual operation flowchart symbol
            autoShapeTypeName = "ShapeFlowchartManualOperation"
        Case msoShapeFlowchartMerge ' Merge flowchart symbol
            autoShapeTypeName = "ShapeFlowchartMerge"
        Case msoShapeFlowchartMultidocument ' Multi-document flowchart symbol
            autoShapeTypeName = "ShapeFlowchartMultidocument"
        Case msoShapeFlowchartOfflineStorage ' Offline storage flowchart symbol
            autoShapeTypeName = "ShapeFlowchartOfflineStorage"
        Case msoShapeFlowchartOffpageConnector ' Off-page connector flowchart symbol
            autoShapeTypeName = "ShapeFlowchartOffpageConnector"
        Case msoShapeFlowchartOr ' "Or" flowchart symbol
            autoShapeTypeName = "ShapeFlowchartOr"
        Case msoShapeFlowchartPredefinedProcess ' Predefined process flowchart symbol
            autoShapeTypeName = "ShapeFlowchartPredefinedProcess"
        Case msoShapeFlowchartPreparation ' Preparation flowchart symbol
            autoShapeTypeName = "ShapeFlowchartPreparation"
        Case msoShapeFlowchartProcess ' Process flowchart symbol
            autoShapeTypeName = "ShapeFlowchartProcess"
        Case msoShapeFlowchartPunchedTape ' Punched tape flowchart symbol
            autoShapeTypeName = "ShapeFlowchartPunchedTape"
        Case msoShapeFlowchartSequentialAccessStorage ' Sequential access storage flowchart symbol
            autoShapeTypeName = "ShapeFlowchartSequentialAccessStorage"
        Case msoShapeFlowchartSort ' Sort flowchart symbol
            autoShapeTypeName = "ShapeFlowchartSort"
        Case msoShapeFlowchartStoredData ' Stored data flowchart symbol
            autoShapeTypeName = "ShapeFlowchartStoredData"
        Case msoShapeFlowchartSummingJunction ' Summing junction flowchart symbol
            autoShapeTypeName = "ShapeFlowchartSummingJunction"
        Case msoShapeFlowchartTerminator ' Terminator flowchart symbol
            autoShapeTypeName = "ShapeFlowchartTerminator"
        Case msoShapeFoldedCorner ' Folded corner
            autoShapeTypeName = "ShapeFoldedCorner"
        Case msoShapeFrame ' Rectangular picture frame
            autoShapeTypeName = "ShapeFrame"
        Case msoShapeFunnel ' Funnel
            autoShapeTypeName = "ShapeFunnel"
        Case msoShapeGear6 ' Gear with six teeth
            autoShapeTypeName = "ShapeGear6"
        Case msoShapeGear9 ' Gear with nine teeth
            autoShapeTypeName = "ShapeGear9"
        Case msoShapeHalfFrame ' Half of a rectangular picture frame
            autoShapeTypeName = "ShapeHalfFrame"
        Case msoShapeHeart ' Heart
            autoShapeTypeName = "ShapeHeart"
        Case msoShapeHeptagon ' Heptagon
            autoShapeTypeName = "ShapeHeptagon"
        Case msoShapeHexagon ' Hexagon
            autoShapeTypeName = "ShapeHexagon"
        Case msoShapeHorizontalScroll ' Horizontal scroll
            autoShapeTypeName = "ShapeHorizontalScroll"
        Case msoShapeIsoscelesTriangle ' Isosceles triangle
            autoShapeTypeName = "ShapeIsoscelesTriangle"
        Case msoShapeLeftArrow ' Block arrow that points left
            autoShapeTypeName = "ShapeLeftArrow"
        Case msoShapeLeftArrowCallout ' Callout with arrow that points left
            autoShapeTypeName = "ShapeLeftArrowCallout"
        Case msoShapeLeftBrace ' Left brace
            autoShapeTypeName = "ShapeLeftBrace"
        Case msoShapeLeftBracket ' Left bracket
            autoShapeTypeName = "ShapeLeftBracket"
        Case msoShapeLeftCircularArrow ' Circular arrow pointing counter-clockwise
            autoShapeTypeName = "ShapeLeftCircularArrow"
        Case msoShapeLeftRightArrow ' Block arrow with arrowheads that point both left and right
            autoShapeTypeName = "ShapeLeftRightArrow"
        Case msoShapeLeftRightArrowCallout ' Callout with arrowheads that point both left and right
            autoShapeTypeName = "ShapeLeftRightArrowCallout"
        Case msoShapeLeftRightCircularArrow ' Circular arrow pointing clockwise and counter-clockwise; a curved arrow with points at both ends
            autoShapeTypeName = "ShapeLeftRightCircularArrow"
        Case msoShapeLeftRightRibbon ' Ribbon with an arrow at both ends
            autoShapeTypeName = "ShapeLeftRightRibbon"
        Case msoShapeLeftRightUpArrow ' Block arrow with arrowheads that point left, right, and up
            autoShapeTypeName = "ShapeLeftRightUpArrow"
        Case msoShapeLeftUpArrow ' Block arrow with arrowheads that point left and up
            autoShapeTypeName = "ShapeLeftUpArrow"
        Case msoShapeLightningBolt ' Lightning bolt
            autoShapeTypeName = "ShapeLightningBolt"
        Case msoShapeLineCallout1 ' Callout with border and horizontal callout line
            autoShapeTypeName = "ShapeLineCallout1"
        Case msoShapeLineCallout1AccentBar ' Callout with horizontal accent bar
            autoShapeTypeName = "ShapeLineCallout1AccentBar"
        Case msoShapeLineCallout1BorderandAccentBar ' Callout with border and horizontal accent bar
            autoShapeTypeName = "ShapeLineCallout1BorderandAccentBar"
        Case msoShapeLineCallout1NoBorder ' Callout with horizontal line
            autoShapeTypeName = "ShapeLineCallout1NoBorder"
        Case msoShapeLineCallout2 ' Callout with diagonal straight line
            autoShapeTypeName = "ShapeLineCallout2"
        Case msoShapeLineCallout2AccentBar ' Callout with diagonal callout line and accent bar
            autoShapeTypeName = "ShapeLineCallout2AccentBar"
        Case msoShapeLineCallout2BorderandAccentBar ' Callout with border, diagonal straight line, and accent bar
            autoShapeTypeName = "ShapeLineCallout2BorderandAccentBar"
        Case msoShapeLineCallout2NoBorder ' Callout with no border and diagonal callout line
            autoShapeTypeName = "ShapeLineCallout2NoBorder"
        Case msoShapeLineCallout3 ' Callout with angled line
            autoShapeTypeName = "ShapeLineCallout3"
        Case msoShapeLineCallout3AccentBar ' Callout with angled callout line and accent bar
            autoShapeTypeName = "ShapeLineCallout3AccentBar"
        Case msoShapeLineCallout3BorderandAccentBar ' Callout with border, angled callout line, and accent bar
            autoShapeTypeName = "ShapeLineCallout3BorderandAccentBar"
        Case msoShapeLineCallout3NoBorder ' Callout with no border and angled callout line
            autoShapeTypeName = "ShapeLineCallout3NoBorder"
        Case msoShapeLineCallout4 ' Callout with callout line segments forming a U-shape
            autoShapeTypeName = "ShapeLineCallout4"
        Case msoShapeLineCallout4AccentBar ' Callout with accent bar and callout line segments forming a U-shape
            autoShapeTypeName = "ShapeLineCallout4AccentBar"
        Case msoShapeLineCallout4BorderandAccentBar ' Callout with border, accent bar, and callout line segments forming a U-shape
            autoShapeTypeName = "ShapeLineCallout4BorderandAccentBar"
        Case msoShapeLineCallout4NoBorder ' Callout with no border and callout line segments forming a U-shape
            autoShapeTypeName = "ShapeLineCallout4NoBorder"
        Case msoShapeLineInverse ' Line inverse
            autoShapeTypeName = "ShapeLineInverse"
        Case msoShapeMathDivide ' Division symbol?÷
            autoShapeTypeName = "ShapeMathDivide"
        Case msoShapeMathEqual ' Equivalence symbol?=
            autoShapeTypeName = "ShapeMathEqual"
        Case msoShapeMathMinus ' Subtraction symbol?-
            autoShapeTypeName = "ShapeMathMinus"
        Case msoShapeMathMultiply ' Multiplication symbol?x
            autoShapeTypeName = "ShapeMathMultiply"
        Case msoShapeMathNotEqual ' Non-equivalence symbol?≠
            autoShapeTypeName = "ShapeMathNotEqual"
        Case msoShapeMathPlus ' Addition symbol?+
            autoShapeTypeName = "ShapeMathPlus"
        Case msoShapeMixed ' Return value only; indicates a combination of the other states.
            autoShapeTypeName = "ShapeMixed"
        Case msoShapeMoon ' Moon
            autoShapeTypeName = "ShapeMoon"
        Case msoShapeNonIsoscelesTrapezoid ' Trapezoid with asymmetrical non-parallel sides
            autoShapeTypeName = "ShapeNonIsoscelesTrapezoid"
        Case msoShapeNoSymbol ' "No" symbol
            autoShapeTypeName = "ShapeNoSymbol"
        Case msoShapeNotchedRightArrow ' Notched block arrow that points right
            autoShapeTypeName = "ShapeNotchedRightArrow"
        Case msoShapeNotPrimitive ' Not supported
            autoShapeTypeName = "ShapeNotPrimitive"
        Case msoShapeOctagon ' Octagon
            autoShapeTypeName = "ShapeOctagon"
        Case msoShapeOval ' Oval
            autoShapeTypeName = "ShapeOval"
        Case msoShapeOvalCallout ' Oval-shaped callout
            autoShapeTypeName = "ShapeOvalCallout"
        Case msoShapeParallelogram ' Parallelogram
            autoShapeTypeName = "ShapeParallelogram"
        Case msoShapePentagon ' Pentagon
            autoShapeTypeName = "ShapePentagon"
        Case msoShapePie ' Circle ('pie') with a portion missing
            autoShapeTypeName = "ShapePie"
        Case msoShapePieWedge ' Quarter of a circular shape
            autoShapeTypeName = "ShapePieWedge"
        Case msoShapePlaque ' Plaque
            autoShapeTypeName = "ShapePlaque"
        Case msoShapePlaqueTabs ' Four quarter-circles defining a rectangular shape
            autoShapeTypeName = "ShapePlaqueTabs"
        Case msoShapeQuadArrow ' Block arrows that point up, down, left, and right
            autoShapeTypeName = "ShapeQuadArrow"
        Case msoShapeQuadArrowCallout ' Callout with arrows that point up, down, left, and right
            autoShapeTypeName = "ShapeQuadArrowCallout"
        Case msoShapeRectangle ' Rectangle
            autoShapeTypeName = "ShapeRectangle"
        Case msoShapeRectangularCallout ' Rectangular callout
            autoShapeTypeName = "ShapeRectangularCallout"
        Case msoShapeRegularPentagon ' Pentagon
            autoShapeTypeName = "ShapeRegularPentagon"
        Case msoShapeRightArrow ' Block arrow that points right
            autoShapeTypeName = "ShapeRightArrow"
        Case msoShapeRightArrowCallout ' Callout with arrow that points right
            autoShapeTypeName = "ShapeRightArrowCallout"
        Case msoShapeRightBrace ' Right brace
            autoShapeTypeName = "ShapeRightBrace"
        Case msoShapeRightBracket ' Right bracket
            autoShapeTypeName = "ShapeRightBracket"
        Case msoShapeRightTriangle ' Right triangle
            autoShapeTypeName = "ShapeRightTriangle"
        Case msoShapeRound1Rectangle ' Rectangle with one rounded corner
            autoShapeTypeName = "ShapeRound1Rectangle"
        Case msoShapeRound2DiagRectangle ' Rectangle with two rounded corners, diagonally-opposed
            autoShapeTypeName = "ShapeRound2DiagRectangle"
        Case msoShapeRound2SameRectangle ' Rectangle with two-rounded corners that share a side
            autoShapeTypeName = "ShapeRound2SameRectangle"
        Case msoShapeRoundedRectangle ' Rounded rectangle
            autoShapeTypeName = "ShapeRoundedRectangle"
        Case msoShapeRoundedRectangularCallout ' Rounded rectangle-shaped callout
            autoShapeTypeName = "ShapeRoundedRectangularCallout"
        Case msoShapeSmileyFace ' Smiley face
            autoShapeTypeName = "ShapeSmileyFace"
        Case msoShapeSnip1Rectangle ' Rectangle with one snipped corner
            autoShapeTypeName = "ShapeSnip1Rectangle"
        Case msoShapeSnip2DiagRectangle ' Rectangle with two snipped corners, diagonally-opposed
            autoShapeTypeName = "ShapeSnip2DiagRectangle"
        Case msoShapeSnip2SameRectangle ' Rectangle with two snipped corners that share a side
            autoShapeTypeName = "ShapeSnip2SameRectangle"
        Case msoShapeSnipRoundRectangle ' Rectangle with one snipped corner and one rounded corner
            autoShapeTypeName = "ShapeSnipRoundRectangle"
        Case msoShapeSquareTabs ' Four small squares that define a rectangular shape
            autoShapeTypeName = "ShapeSquareTabs"
        Case msoShapeStripedRightArrow ' Block arrow that points right with stripes at the tail
            autoShapeTypeName = "ShapeStripedRightArrow"
        Case msoShapeSun ' Sun
            autoShapeTypeName = "ShapeSun"
        Case msoShapeSwooshArrow ' Curved arrow
            autoShapeTypeName = "ShapeSwooshArrow"
        Case msoShapeTear ' Water droplet
            autoShapeTypeName = "ShapeTear"
        Case msoShapeTrapezoid ' Trapezoid
            autoShapeTypeName = "ShapeTrapezoid"
        Case msoShapeUpArrow ' Block arrow that points up
            autoShapeTypeName = "ShapeUpArrow"
        Case msoShapeUpArrowCallout ' Callout with arrow that points up
            autoShapeTypeName = "ShapeUpArrowCallout"
        Case msoShapeUpDownArrow ' Block arrow that points up and down
            autoShapeTypeName = "ShapeUpDownArrow"
        Case msoShapeUpDownArrowCallout ' Callout with arrows that point up and down
            autoShapeTypeName = "ShapeUpDownArrowCallout"
        Case msoShapeUpRibbon ' Ribbon banner with center area above ribbon ends
            autoShapeTypeName = "ShapeUpRibbon"
        Case msoShapeUTurnArrow ' Block arrow forming a U shape
            autoShapeTypeName = "ShapeUTurnArrow"
        Case msoShapeVerticalScroll ' Vertical scroll
            autoShapeTypeName = "ShapeVerticalScroll"
        Case msoShapeWave ' Wave
            autoShapeTypeName = "ShapeWave"
        Case Else
            autoShapeTypeName = "Other AutoShape"
    End Select
    
    ConvertAutoShapeTypeNumberToName = autoShapeTypeName
End Function

Private Function ParseCsvLine(ByVal line As String) As Variant
    Dim result As Collection
    Dim inQuotes As Boolean
    Dim currentField As String
    Dim i As Long
    Dim char As String
    
    Set result = New Collection
    inQuotes = False
    currentField = ""
    
    For i = 1 To Len(line)
        char = Mid(line, i, 1)
        
        If char = """" Then
            If inQuotes And Mid(line, i + 1, 1) = """" Then
                ' Escaped quote inside quoted field
                currentField = currentField & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf char = "," And Not inQuotes Then
            result.Add currentField
            currentField = ""
        Else
            currentField = currentField & char
        End If
    Next i
    result.Add currentField ' Add the last field

    ParseCsvLine = CollectionToArray(result)
End Function

Private Function CollectionToArray(coll As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    
    ReDim arr(0 To coll.Count - 1)
    For i = 1 To coll.Count
        arr(i - 1) = coll(i)
    Next i

    CollectionToArray = arr
End Function

Public Function MakeConnectorShapesCsvString(sheetName As String) As String
    Dim ws As Worksheet
    Dim shp As Shape
    Dim connBegin As Shape
    Dim connEnd As Shape
    Dim csvOutput As String
    Dim connectorFormatType As String
    Dim connectorFormatName As String
    Dim lineColor As String
    Dim lineColorRGB As String
    Dim dashType As String
    Dim lineWidth As Double
    Dim beginArrowType As String
    Dim beginArrowTypeName As String
    Dim endArrowType As String
    Dim endArrowTypeName As String
    
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' CSVヘッダ行を設定
    csvOutput = "ID,Name,Format,FormatName,Color,ColorRGB,DashType,Width,BeginArrowType,BeginArrowTypeName,BeginConnectedShapeID,BeginConnectedShapeInnerText,BeginConnectionSite,EndArrowType,EndArrowTypeName,EndConnectedShapeID,EndConnectedShapeInnerText,EndConnectionSite,Selected" & vbCrLf

    ' コネクタの情報を収集
    For Each shp In ws.Shapes
        If shp.Connector Then
            ' コネクタの始点と終点の図形を取得
            On Error Resume Next
            Set connBegin = Nothing
            Set connEnd = Nothing
            Set connBegin = shp.connectorFormat.BeginConnectedShape
            Set connEnd = shp.connectorFormat.EndConnectedShape
            On Error GoTo 0

            ' コネクタの属性を取得
            lineColor = shp.line.foreColor.RGB
            lineColorRGB = ConvertForeColorRgbToHexString(lineColor)
            dashType = shp.line.DashStyle
            lineWidth = shp.line.Weight
            connectorFormatType = shp.connectorFormat.Type
            beginArrowType = shp.line.BeginArrowheadStyle
            beginArrowTypeName = ConvertArrowheadStyle(beginArrowType)
            endArrowType = shp.line.EndArrowheadStyle
            endArrowTypeName = ConvertArrowheadStyle(endArrowType)

            Select Case connectorFormatType
                Case msoConnectorStraight
                    connectorFormatName = "Straight Connector"
                Case msoConnectorElbow
                    connectorFormatName = "Elbow Connector"
                Case msoConnectorCurve
                    connectorFormatName = "Curved Connector"
                Case Else
                    connectorFormatName = "Other Connector"
            End Select
            
            ' コネクタに接続されている図形の情報を収集
            csvOutput = csvOutput & shp.ID & ","
            csvOutput = csvOutput & shp.Name & ","
            csvOutput = csvOutput & connectorFormatType & ","
            csvOutput = csvOutput & connectorFormatName & ","
            csvOutput = csvOutput & lineColor & ","
            csvOutput = csvOutput & lineColorRGB & ","
            csvOutput = csvOutput & dashType & ","
            csvOutput = csvOutput & lineWidth & ","
            
            csvOutput = csvOutput & beginArrowType & ","
            csvOutput = csvOutput & beginArrowTypeName & ","
            
            If Not connBegin Is Nothing Then
                csvOutput = csvOutput & connBegin.ID & ","

                If Not connBegin.TextFrame2.HasText = msoFalse Then
                    csvOutput = csvOutput & """" & connBegin.TextFrame2.TextRange.Text & """" & ","
                Else
                    csvOutput = csvOutput & ","
                End If
        
                csvOutput = csvOutput & shp.connectorFormat.BeginConnectionSite & ","
            Else
                csvOutput = csvOutput & ",,,"
            End If
            
            csvOutput = csvOutput & endArrowType & ","
            csvOutput = csvOutput & endArrowTypeName & ","
            
            If Not connEnd Is Nothing Then
                csvOutput = csvOutput & connEnd.ID & ","
                
                If Not connEnd.TextFrame2.HasText = msoFalse Then
                    csvOutput = csvOutput & """" & connEnd.TextFrame2.TextRange.Text & """" & ","
                Else
                    csvOutput = csvOutput & ","
                End If
                
                csvOutput = csvOutput & shp.connectorFormat.EndConnectionSite & ","
            Else
                csvOutput = csvOutput & ",,,"
            End If
            
            csvOutput = csvOutput & "False" & vbCrLf
        End If
    Next shp
    
    ' Remove the last newline character
    If Right(csvOutput, 2) = vbCrLf Then
        csvOutput = Left(csvOutput, Len(csvOutput) - 2)
    End If
    
    ' Debug.Print csvOutput
    MakeConnectorShapesCsvString = csvOutput
End Function


Public Function MakeShapesCsvString(sheetName As String) As String
    Dim ws As Worksheet
    Dim shp As Shape
    Dim csvOutput As String
    Dim shapeType As String
    Dim shapeTypeName As String
    Dim foreColor As String
    Dim foreColorRGB As String
    Dim posTop As Single
    Dim posLeft As Single
    Dim shapeHeight As Single
    Dim shapeWidth As Single
    Dim autoShapeType As String
    Dim autoShapeTypeName As String
    Dim innerText As String
    
    Set ws = ThisWorkbook.Sheets(sheetName)

    csvOutput = "ID,Name,AlternativeText,Type,TypeName,AutoShapeType,AutoShapeName,ForeColor,ForeColorRGB,Top,Left,Height,Width,ZOrderPosition,Text,Selected" & vbCrLf

    For Each shp In ws.Shapes
        If Not shp.Connector Then
            shapeType = shp.Type
            shapeTypeName = ConvertShapeTypeNumberToName(shapeType)
            foreColor = shp.Fill.foreColor.RGB
            foreColorRGB = ConvertForeColorRgbToHexString(foreColor)
            posTop = shp.Top
            posLeft = shp.Left
            shapeHeight = shp.Height
            shapeWidth = shp.Width
                
            If shp.Type = msoAutoShape Then
                autoShapeType = shp.autoShapeType
                autoShapeTypeName = ConvertAutoShapeTypeNumberToName(autoShapeType)
            Else
                autoShapeType = ""
                autoShapeTypeName = ""
            End If
        
            If Not shp.TextFrame2.HasText = msoFalse Then
                innerText = """" & shp.TextFrame2.TextRange.Text & """"
            Else
                innerText = ""
            End If
        
            csvOutput = csvOutput & shp.ID & ","
            csvOutput = csvOutput & shp.Name & ","
            csvOutput = csvOutput & shp.AlternativeText & ","
            csvOutput = csvOutput & shapeType & ","
            csvOutput = csvOutput & shapeTypeName & ","
            csvOutput = csvOutput & autoShapeType & ","
            csvOutput = csvOutput & autoShapeTypeName & ","
            csvOutput = csvOutput & foreColor & ","
            csvOutput = csvOutput & foreColorRGB & ","
            csvOutput = csvOutput & posTop & ","
            csvOutput = csvOutput & posLeft & ","
            csvOutput = csvOutput & shapeHeight & ","
            csvOutput = csvOutput & shapeWidth & ","
            csvOutput = csvOutput & shp.ZOrderPosition & ","
            ' csvOutput = csvOutput & shp.Nodes.Count & ","
            csvOutput = csvOutput & innerText & ","
            csvOutput = csvOutput & "False" & vbCrLf
        End If
    Next shp
    
    ' Remove the last newline character
    If Right(csvOutput, 2) = vbCrLf Then
        csvOutput = Left(csvOutput, Len(csvOutput) - 2)
    End If
    
    ' Debug.Print csvOutput
    MakeShapesCsvString = csvOutput
End Function


Public Function InsertCsvAsTable(sheetName As String, tableName As String, csvData As String)
    Dim ws As Worksheet
    Dim csvLines As Variant
    Dim csvFields As Variant
    Dim i As Long, j As Long
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Sheets(sheetName)

    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    If Not tbl Is Nothing Then
        tbl.Delete
    End If
    On Error GoTo 0

    ' CSVデータを行ごとに分割
    csvLines = Split(csvData, vbCrLf)
    
    For i = LBound(csvLines) To UBound(csvLines)
        If csvLines(i) <> "" Then
            csvFields = ParseCsvLine(csvLines(i))
            For j = LBound(csvFields) To UBound(csvFields)
                ws.Cells(i + 1, j + 1).Value = csvFields(j)
            Next j
        End If
    Next i

    ' テーブル範囲を設定（データの範囲を適宜調整）
    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(1 + UBound(csvLines), 1 + UBound(csvFields)))

    ' テーブルとしてフォーマット
    Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tbl.Name = tableName
End Function

Public Function BackupTableData(sheetNameSrc As String, tableName As String, sheetNameDist As String)
    Dim wsSrc As Worksheet
    Dim wsDist As Worksheet
    Dim tbl As ListObject
    Dim backupTbl As ListObject
    Dim backupTblName As String: backupTblName = "_Backup_" & tableName

    ' バックアップ用シートを検索／作成
    On Error Resume Next
    Set wsDist = ThisWorkbook.Sheets(sheetNameDist)
    If wsDist Is Nothing Then
        Set wsDist = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDist.Name = sheetNameDist
    End If
    wsDist.Visible = xlSheetHidden
    On Error GoTo 0

    ' 既存のバックアップテーブルを削除
    On Error Resume Next
    Set backupTbl = wsDist.ListObjects(backupTblName)
    If Not backupTbl Is Nothing Then
        backupTbl.Delete
    End If
    On Error GoTo 0

    ' テーブルをバックアップシートにコピー
    Set wsSrc = ThisWorkbook.Sheets(sheetNameSrc)
    Set tbl = wsSrc.ListObjects(tableName)
    tbl.Range.Copy
    wsDist.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False ' コピーの選択をクリア
    
    ' 貼り付けられたテーブルを取得して名前を変更
    On Error Resume Next
    Set backupTbl = wsDist.ListObjects(1)
    On Error GoTo 0

    If Not backupTbl Is Nothing Then
        backupTbl.Name = backupTblName
    End If
End Function

' Function to convert RGB to Hex (e.g., #FF0000)
Function ConvertRGBFromHex(hex As String) As Long
    Dim r As Long, g As Long, b As Long
    hex = Replace(hex, "#", "")
    r = CLng("&H" & Mid(hex, 1, 2))
    g = CLng("&H" & Mid(hex, 3, 2))
    b = CLng("&H" & Mid(hex, 5, 2))
    ConvertRGBFromHex = RGB(r, g, b)
End Function

Function ExportAllModules()
    Dim vbComp As Object
    Dim destFolder As String
    Dim wbPath As String
    Dim fso As Object

    ' Get the path of the current workbook
    wbPath = ThisWorkbook.Path

    ' Create the Modules folder path
    destFolder = wbPath & "\Modules"
    
    ' Create the FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the Modules folder exists, if not, create it
    If Not fso.FolderExists(destFolder) Then
        fso.CreateFolder destFolder
    End If
    
    ' Loop through all VBA components (modules)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Export the module to the Modules folder
        vbComp.Export destFolder & "\" & vbComp.Name & ".bas"
    Next vbComp
    
    ' Clean up
    Set fso = Nothing
    
    ' Notify the user
    MsgBox "All modules have been exported to the 'Modules' folder.", vbInformation
End Function


