Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub Plot()

Dim p As Shape 'point
Dim Axis As Shape
Dim Xrange As Range
Dim Yrange As Range
Dim FullRnge As Range
Dim Xval As Double
Dim Yval As Double
Dim XvalOld As Double
Dim YvalOld As Double
Dim Xinit As Double
Dim Yinit As Double
Dim CenX As Double
Dim CenY As Double
Dim Acheck As Double
Dim MaxXY As Double
Dim S As Long 'number of stringers
Dim Label As String

Dim scle As Double 'scale
'independent x,y translate
Dim OtransX As Double
Dim OtransY As Double

Set FullRnge = Range("B2:C51") 'range for all x,y values
Set Xrange = Range("B2:B51") 'range for x values
Set Yrange = Range("C2:C51") 'range for y values

    'Centroid range
    CenX = Cells(3, 11).Value
    CenY = Cells(3, 12).Value


    'scale - find max value, divide by max wanted resolution size
    
    MaxXY = MaxAbsR(FullRnge)
    scle = (MaxXY / 300)
    'shift origin
    OtransX = 600 'pixels from left
    OtransY = 400 'pixels from top
    
    'delete existing points
    For Each p In ActiveSheet.Shapes
        If p.AlternativeText <> "Run" Then 'delete everything with out the text "Run"
              p.Delete
        End If
    Next p
    
    For Each Axis In ActiveSheet.Shapes 'delete existing Axis
        If Axis.AlternativeText <> "Run" Then
              Axis.Delete
        End If
    Next Axis


    'draw axis
    
    Set Axis = ActiveSheet.Shapes.AddShape(1, trans_xcoords(0, scle, OtransX) - 320, trans_ycoords(0, scle, OtransY) - 300, 650, 610)

    Axis.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Axis.Line.Weight = 1.75
    
    'Vertical Axis
    ActiveSheet.Shapes.AddLine(0 + OtransX + 5, 0 + OtransY + 5, 0 + OtransX + 5, trans_ycoords(200, 1, OtransY) - 95).Select
    With Selection.ShapeRange.Line
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadOpen
        .Weight = 2.75
        .Transparency = 0.75
        .DashStyle = msoLineLongDashDot
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    'Horizontal Axis
    ActiveSheet.Shapes.AddLine(0 + OtransX + 5, 0 + OtransY + 5, trans_xcoords(200, 1, OtransX) + 95, 0 + OtransY + 5).Select
    With Selection.ShapeRange.Line
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadOpen
        .Weight = 2.75
        .Transparency = 0.75
        .DashStyle = msoLineLongDashDot
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
    ActiveSheet.Range("A1").Select

    
    For i = 2 To 51 'set number of stringers
        Xval = Cells(i, 2).Value
        Yval = Cells(i, 3).Value
        
        Label = Cells(i, 1).Value
        Acheck = Cells(i, 4).Value
    'create outer border
'        Range("C8:L35").BorderAround _
'                    ColorIndex:=3, Weight:=xlThick
                    
      
        If Acheck <> 0 Then
            'draw points
            Set p = ActiveSheet.Shapes.AddShape(9, trans_xcoords(Xval, scle, OtransX), trans_ycoords(Yval, scle, OtransY), 10, 10)
            p.Fill.ForeColor.RGB = RGB(0, 0, 0)
            'add labels
            Set p = ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, trans_xcoords(Xval, scle, OtransX) + 5, trans_ycoords(Yval, scle, OtransY), 100, 100)
            p.TextFrame.Characters.Text = Label
                 
            If i = 2 Then
                Xinit = trans_xcoords(Xval, scle, OtransX)
                Yinit = trans_ycoords(Yval, scle, OtransY)
            End If
            
            If i > 2 Then
                Set p = ActiveSheet.Shapes.AddLine(XvalOld + 5, YvalOld + 5, trans_xcoords(Xval, scle, OtransX) + 5, trans_ycoords(Yval, scle, OtransY) + 5)
                p.Line.ForeColor.RGB = RGB(0, 0, 0)
                p.Line.Weight = 1.75
                p.Line.Transparency = 0.25
            End If
            
            XvalOld = trans_xcoords(Xval, scle, OtransX)
            YvalOld = trans_ycoords(Yval, scle, OtransY)
        End If

    Next
    
    Set p = ActiveSheet.Shapes.AddShape(9, trans_xcoords(CenX, scle, OtransX), trans_ycoords(CenY, scle, OtransY), 10, 10)
    p.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Set p = ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, trans_xcoords(CenX, scle, OtransX) + 5, trans_ycoords(CenY, scle, OtransY), 100, 100)
    p.TextFrame.Characters.Text = "C"
    
        'comment in/out the 3 lines below to switch from open/closed section
        Set p = ActiveSheet.Shapes.AddLine(XvalOld + 5, YvalOld + 5, Xinit + 5, Yinit + 5)
        p.Line.ForeColor.RGB = RGB(0, 0, 0)
        p.Line.Weight = 1.75
        p.Line.Transparency = 0.25

End Sub
'X Coord translate function
Function trans_xcoords(x As Double, aScl As Double, aOtrans As Double)

    trans_xcoords = (x / aScl) + aOtrans
    
End Function
'Y Coord translate function
Function trans_ycoords(y As Double, bScl As Double, bOtrans As Double)

    trans_ycoords = (-y / bScl) + bOtrans

End Function

Function MaxAbsR(Dataa As Range)
Dim MaxVal1 As Double, MinVall As Double
    MaxVal1 = WorksheetFunction.Max(Dataa)
    MinVall = -WorksheetFunction.Min(Dataa)
 
    If MaxVal1 > MinVall Then MaxAbsR = MaxVal1 Else MaxAbsR = MinVall
 
End Function




