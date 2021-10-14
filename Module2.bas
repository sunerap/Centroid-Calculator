Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub ConnectPoints()
'
' ConnectPoints Macro
'

'
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 405.7618897638, _
        158.7428346457, 405.7618897638, 297.7142519685).Select
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes( _
        "Oval 1952"), 1
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes("Oval 1954" _
        ), 5
    
End Sub

