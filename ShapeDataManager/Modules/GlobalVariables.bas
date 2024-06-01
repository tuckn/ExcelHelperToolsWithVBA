Attribute VB_Name = "GlobalVariables"
Option Explicit

' Sheet name
Public Const SHEETNAME_CANVAS As String = "Canvas"
Public Const SHEETNAME_CONNECTORS As String = "Connectors"
Public Const SHEETNAME_SHAPES As String = "Shapes"
Public Const SHEETNAME_BACKUP_SHAPES As String = "_Backup_Shapes"

' Table name
Public Const TABLENAME_CONNECTOR As String = "Connectors"
Public Const TABLENAME_SHAPES As String = "Shapes"
Public Const TABLENAME_BKUPSHAPES As String = "_Backup_Shapes"

' Column index
Public Const colConnID As Long = 1
Public Const colConnName As Long = 2
Public Const colConnFormat As Long = 3
Public Const colConnFormatName As Long = 4
Public Const colConnColor As Long = 5
Public Const colConnColorRGB As Long = 6
Public Const colConnDashType As Long = 7
Public Const colConnWidth As Long = 8
Public Const colConnBeginArrowType As Long = 9
Public Const colConnBeginArrowTypeName As Long = 10
Public Const colConnBeginConnectedShapeID As Long = 11
Public Const colConnBeginConnectedShapeInnerText As Long = 12
Public Const colConnBeginConnectionSite As Long = 13
Public Const colConnEndArrowType As Long = 14
Public Const colConnEndArrowTypeName As Long = 15
Public Const colConnEndConnectedShapeID As Long = 16
Public Const colConnEndConnectedShapeInnerText As Long = 17
Public Const colConnEndConnectionSite As Long = 18
Public Const colConnSelected As Long = 19

Public Const colShpID As Long = 1
Public Const colShpName As Long = 2
Public Const colShpAlternativeText As Long = 3
Public Const colShpType As Long = 4
Public Const colShpTypeName As Long = 5
Public Const colShpAutoShapeType As Long = 6
Public Const colShpAutoShapeName As Long = 7
Public Const colShpForeColor As Long = 8
Public Const colShpForeColorRGB As Long = 9
Public Const colShpTop As Long = 10
Public Const colShpLeft As Long = 11
Public Const colShpHeight As Long = 12
Public Const colShpWidth As Long = 13
Public Const colShpZOrderPosition As Long = 14
Public Const colShpText As Long = 15
Public Const colShpSelected As Long = 16

