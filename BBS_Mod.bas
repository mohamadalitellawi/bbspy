Attribute VB_Name = "BBS_Mod"
Option Base 1
Option Explicit
Const ShapeWeight = 3
Const ShapeColor = 56
Const RrR = 10
Const GgG = 40
Const BbB = 110
Const ColBMark = 2
Const ColUnitLength = 7
Const ColShape = 11
Const ColShapeCode = 14
Const ColA = 15
Const ColB = 16
Const ColC = 17
Const ColD = 18
Const ColE = 19
Const ColF = 20
Const ColG = 21
Const ColBendDiameter = 22
Const ColShapeName = 23
Const RowHeight = 82
Const Pi = 3.14159265358979
'
Dim ShapeArray(200) As Shape
Dim ShapeArrayNames(200) As String

Sub DeleteAllShapes()
Dim ALLSHAPES As Shape
Dim i As Integer

If MsgBox("Are you shure", vbOKCancel, "MaT") = vbOK Then
For Each ALLSHAPES In ActiveSheet.Shapes
    If Left(ALLSHAPES.Name, 5) = "Group" Then ALLSHAPES.Delete
Next
For i = 1 To 200
Cells(i + 5, ColShapeName).Value = ""
Next
End If
End Sub
Sub ReDrawAllShapes()
'ShapeColor = Int(Rnd * 10)
DRAW_SAHPES
End Sub
Sub DRAW_SAHPES()
Dim ShapeIndex As Long
Dim i As Long: i = 1

'CLEAR ALL SHAPES
Dim ALLSHAPES As Shape
'For Each ALLSHAPES In ActiveSheet.Shapes
'    ALLSHAPES.Delete
'Next
If MsgBox("Are you shure", vbOKCancel, "MaT") = vbCancel Then Exit Sub
For Each ALLSHAPES In ActiveSheet.Shapes
    If Left(ALLSHAPES.Name, 5) = "Group" Then ALLSHAPES.Delete
Next

For i = 1 To 200
    Cells(i + 5, ColShapeName).Value = ""
    If IsNumeric(Cells(i, ColShapeCode).Value) Then
        ShapeIndex = Cells(i, ColShapeCode).Value
        Select Case ShapeIndex
            Case 1
                shape_1 (i)
            Case 2
                shape_2 (i)
            Case 3
                shape_3 (i)
            Case 4
                shape_4 (i)
            Case 5
                shape_5 (i)
            Case 6
                shape_6 (i)
             Case 7
                shape_7 (i)
            Case 8
                shape_8 (i)
            Case 9
                shape_9 (i)
            Case 10
                shape_10 (i)
            Case 11
                shape_11 (i)
            Case 12
                shape_12 (i)
            Case 13
                shape_13 (i)
            Case 14
                shape_14 (i)
            Case 15
                shape_15 (i)
            Case 16
                shape_16 (i)
            Case 17
                shape_17 (i)
            Case 18
                shape_18 (i)
            Case 19
                Shape_19 (i)
            Case 20
                Shape_20 (i)
            Case 21
                Shape_21 (i)
            Case 22
                Shape_22 (i)
            Case 23
                Shape_23 (i)
            Case 24
                Shape_24 (i)
            Case 25
                Shape_25 (i)
            Case 26
                Shape_26 (i)
            Case 27
                Shape_27 (i)
            Case 28
                Shape_28 (i)
            Case 29
                Shape_29 (i)
            Case 30
                Shape_30 (i)
            Case 31
                Shape_31 (i)
            Case 32
                Shape_32 (i)
                
            Case 33
                shape_33 (i)
                
            Case 34
                shape_34 (i)
            
            Case 35
                shape_35 (i)
                
            Case 36
                shape_36 (i)
                
            Case 37
                Shape_37 (i)
                
            Case 38
                Shape_38 (i)
            'Case 39
            '    shape_39 (i)
            'Case 40
            '    shape_40 (i)
                
            Case 99
                Shape_99 (i)
            End Select
    End If
Next i

End Sub
Sub Del_SHape(ShapeRowIndex As Long)
    On Error Resume Next
    If Len(Cells(ShapeRowIndex, ColShapeName).Value) > 0 Then ActiveSheet.Shapes.Item(Cells(ShapeRowIndex, ColShapeName).Value).Delete
    Cells(ShapeRowIndex, ColShapeName).Value = ""
End Sub
Sub DRAW_SAHPES2(ShapeRowIndex As Long)
Dim ShapeIndex As Long
Dim i As Long: i = 1

'CLEAR ALL SHAPES
'Dim ALLSHAPES As Shape
'For Each ALLSHAPES In ActiveSheet.Shapes
'    ALLSHAPES.Delete
'Next
'Stop
'On Error Resume Next
''On Error GoTo ERRR
'If Len(Cells(ShapeRowIndex, ColShapeName).Value) > 0 Then ActiveSheet.Shapes.Item(Cells(ShapeRowIndex, ColShapeName).Value).Delete
''Cells(ShapeRowIndex, ColShapeName).Value = ""
'On Error GoTo 0

'For i = 1 To 200
i = ShapeRowIndex
    If IsNumeric(Cells(i, ColShapeCode).Value) Then
        ShapeIndex = Cells(i, ColShapeCode).Value
        Select Case ShapeIndex
            Case 1
                shape_1 (i)
            Case 2
                shape_2 (i)
            Case 3
                shape_3 (i)
            Case 4
                shape_4 (i)
            Case 5
                shape_5 (i)
            Case 6
                shape_6 (i)
             Case 7
                shape_7 (i)
            Case 8
                shape_8 (i)
            Case 9
                shape_9 (i)
            Case 10
                shape_10 (i)
            Case 11
                shape_11 (i)
            Case 12
                shape_12 (i)
            Case 13
                shape_13 (i)
            Case 14
                shape_14 (i)
            Case 15
                shape_15 (i)
            Case 16
                shape_16 (i)
            Case 17
                shape_17 (i)
            Case 18
                shape_18 (i)
            Case 19
                Shape_19 (i)
            Case 20
                Shape_20 (i)
            Case 21
                Shape_21 (i)
            Case 22
                Shape_22 (i)
            Case 23
                Shape_23 (i)
            Case 24
                Shape_24 (i)
            Case 25
                Shape_25 (i)
            Case 26
                Shape_26 (i)
            Case 27
                Shape_27 (i)
            Case 28
                Shape_28 (i)
            Case 29
                Shape_29 (i)
            Case 30
                Shape_30 (i)
            Case 31
                Shape_31 (i)
            Case 32
                Shape_32 (i)
            Case 33
                shape_33 (i)
                
            Case 34
                shape_34 (i)
            
            Case 35
                shape_35 (i)
                
            Case 36
                shape_36 (i)
                
            Case 37
                Shape_37 (i)
                
            Case 38
                Shape_38 (i)
            'Case 39
            '    shape_39 (i)
            'Case 40
            '    shape_40 (i)
                
            Case 99
                Shape_99 (i)
            End Select
    End If
'Next i
Exit Sub
'ERRR:
'MsgBox Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source & vbCrLf & Len(Cells(ShapeRowIndex, ColShapeName).Value) & vbCrLf & Cells(ShapeRowIndex, ColShapeName).Value
End Sub
Sub tryDim()
Dim Shape1 As Shape

Set Shape1 = DrawLeader2(250, 250, 350, 150, "hello")

End Sub
Sub testt()
'MsgBox Cos(45 / 180 * 3.14159265358979)

' Draw Dimensions
Dim DimShape As Shape


Set DimShape = DrawAlDimensionsR(100, 300, 300, 500, "hhh", 50)
'
ActiveSheet.Shapes.Range(Array(DimShape.Name)).Select
'Set Gr1 = Selection.ShapeRange.Group
End Sub
Function DrawArcDimensions(StartX As Double, StartY As Double, EndX As Double, EndY As Double, DimText As String, Optional Extension = 18) As Shape

End Function
Function DrawAlDimensionsR(StartX As Double, StartY As Double, EndX As Double, EndY As Double, DimText As String, Optional Extension = 18) As Shape

' Draw Dimensions
Dim P1(2) As Double
Dim P2(2) As Double
Dim P3(2) As Double
Dim P4(2) As Double
Dim P5(2) As Double
Dim P6(2) As Double

Dim DimLine1 As Shape
Dim DimLine2 As Shape
Dim DimLine3 As Shape
'
P1(1) = StartX
P1(2) = StartY
P2(1) = StartX + Extension * Cos(45 / 180 * 3.14159265358979)
P2(2) = StartY - Extension * Cos(45 / 180 * 3.14159265358979)
'
P3(1) = EndX
P3(2) = EndY
P4(1) = EndX + Extension * Cos(45 / 180 * 3.14159265358979)
P4(2) = EndY - Extension * Cos(45 / 180 * 3.14159265358979)
'
P5(1) = P2(1) - 5 * Cos(45 / 180 * 3.14159265358979)
P5(2) = P2(2) + 5 * Cos(45 / 180 * 3.14159265358979)
P6(1) = P4(1) - 5 * Cos(45 / 180 * 3.14159265358979)
P6(2) = P4(2) + 5 * Cos(45 / 180 * 3.14159265358979)
'
Set DimLine1 = ActiveSheet.Shapes.AddLine(P1(1), P1(2), P2(1), P2(2))
Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
'Set DimLine3 = ActiveSheet.Shapes.AddLine((P1(1) + P2(1)) / 2, (P1(2) + P2(2)) / 2, (P3(1) + P4(1)) / 2, (P3(2) + P4(2)) / 2)
Set DimLine3 = ActiveSheet.Shapes.AddLine(P5(1), P5(2), P6(1), P6(2))
DimLine3.Line.BeginArrowheadStyle = msoArrowheadTriangle
DimLine3.Line.EndArrowheadStyle = msoArrowheadTriangle
'
'DIMENSIONS TEXT
Dim txtDimText1 As String
txtDimText1 = DimText
Dim txtP1(2) As Double
Dim txtWidth1 As Long
Dim txtHeight1 As Long
Dim DimText1 As Shape
'draw dimension text 1

txtWidth1 = 0
txtHeight1 = 12
txtP1(1) = (P5(1) + P6(1)) / 2 + txtHeight1 * 0.8
txtP1(2) = (P5(2) + P6(2)) / 2 - txtHeight1 * 0.8
'
Set DimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtWidth1, txtHeight1)
DimText1.TextFrame.Characters.Text = txtDimText1
DimText1.TextFrame.MarginBottom = 0
DimText1.TextFrame.MarginLeft = 0
DimText1.TextFrame.MarginRight = 0
DimText1.TextFrame.MarginTop = 0
DimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
DimText1.TextFrame.VerticalAlignment = xlVAlignCenter
DimText1.Rotation = 45
DimText1.TextFrame.AutoSize = True
DimText1.Line.Visible = msoFalse
'
ActiveSheet.Shapes.Range(Array(DimLine1.Name, DimLine2.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawAlDimensionsR = Selection.ShapeRange.Group
End Function

Function DrawAlDimensionsL(StartX As Double, StartY As Double, EndX As Double, EndY As Double, DimText As String, Optional Extension = 18) As Shape

' Draw Dimensions
Dim P1(2) As Double
Dim P2(2) As Double
Dim P3(2) As Double
Dim P4(2) As Double
Dim P5(2) As Double
Dim P6(2) As Double

Dim DimLine1 As Shape
Dim DimLine2 As Shape
Dim DimLine3 As Shape
'
P1(1) = StartX
P1(2) = StartY
P2(1) = StartX - Extension * Cos(45 / 180 * 3.14159265358979)
P2(2) = StartY - Extension * Cos(45 / 180 * 3.14159265358979)
'
P3(1) = EndX
P3(2) = EndY
P4(1) = EndX - Extension * Cos(45 / 180 * 3.14159265358979)
P4(2) = EndY - Extension * Cos(45 / 180 * 3.14159265358979)
'
P5(1) = P2(1) + 5 * Cos(45 / 180 * 3.14159265358979)
P5(2) = P2(2) + 5 * Cos(45 / 180 * 3.14159265358979)
P6(1) = P4(1) + 5 * Cos(45 / 180 * 3.14159265358979)
P6(2) = P4(2) + 5 * Cos(45 / 180 * 3.14159265358979)
'
Set DimLine1 = ActiveSheet.Shapes.AddLine(P1(1), P1(2), P2(1), P2(2))
Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
'Set DimLine3 = ActiveSheet.Shapes.AddLine((P1(1) + P2(1)) / 2, (P1(2) + P2(2)) / 2, (P3(1) + P4(1)) / 2, (P3(2) + P4(2)) / 2)
Set DimLine3 = ActiveSheet.Shapes.AddLine(P5(1), P5(2), P6(1), P6(2))
DimLine3.Line.BeginArrowheadStyle = msoArrowheadTriangle
DimLine3.Line.EndArrowheadStyle = msoArrowheadTriangle
'
'DIMENSIONS TEXT
Dim txtDimText1 As String
txtDimText1 = DimText
Dim txtP1(2) As Double
Dim txtWidth1 As Long
Dim txtHeight1 As Long
Dim DimText1 As Shape
'draw dimension text 1

txtWidth1 = 0
txtHeight1 = 12
txtP1(1) = (P5(1) + P6(1)) / 2 - txtHeight1 * 0.8
txtP1(2) = (P5(2) + P6(2)) / 2 - txtHeight1 * 0.8
'
Set DimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtWidth1, txtHeight1)
DimText1.TextFrame.Characters.Text = txtDimText1
DimText1.TextFrame.MarginBottom = 0
DimText1.TextFrame.MarginLeft = 0
DimText1.TextFrame.MarginRight = 0
DimText1.TextFrame.MarginTop = 0
DimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
DimText1.TextFrame.VerticalAlignment = xlVAlignCenter
DimText1.Rotation = -45
DimText1.TextFrame.AutoSize = True
DimText1.Line.Visible = msoFalse
'
ActiveSheet.Shapes.Range(Array(DimLine1.Name, DimLine2.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawAlDimensionsL = Selection.ShapeRange.Group
End Function

Function DrawHorDimensions(StartX As Double, StartY As Double, EndX As Double, DimText As String, Optional Extension = 18) As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
txtDimText1 = DimText

' Draw Dimensions
Dim P1(2) As Double
Dim P2(2) As Double
Dim P3(2) As Double
Dim P4(2) As Double
Dim DimLine1 As Shape
Dim DimLine2 As Shape
Dim DimLine3 As Shape

Dim txtP1(2) As Double


Dim txtWidth1 As Long
Dim txtHeight1 As Long

Dim DimText1 As Shape

P1(1) = StartX
P1(2) = StartY
P2(1) = StartX
P2(2) = StartY + Extension

P3(1) = EndX
P3(2) = StartY
P4(1) = EndX
P4(2) = StartY + Extension

Set DimLine1 = ActiveSheet.Shapes.AddLine(P1(1), P1(2), P2(1), P2(2))
Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
Set DimLine3 = ActiveSheet.Shapes.AddLine(P1(1), P2(2) - 5, P3(1), P2(2) - 5)
DimLine3.Line.BeginArrowheadStyle = msoArrowheadTriangle
DimLine3.Line.EndArrowheadStyle = msoArrowheadTriangle
'
'draw dimension text 1
txtP1(1) = StartX
txtP1(2) = StartY + 7
txtWidth1 = EndX - StartX
txtHeight1 = 12

Set DimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtWidth1, txtHeight1)
DimText1.TextFrame.Characters.Text = txtDimText1
DimText1.TextFrame.MarginBottom = 0
DimText1.TextFrame.MarginLeft = 0
DimText1.TextFrame.MarginRight = 0
DimText1.TextFrame.MarginTop = 0
DimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
DimText1.TextFrame.VerticalAlignment = xlVAlignCenter
DimText1.TextFrame.AutoSize = True
DimText1.Line.Visible = msoFalse
'
ActiveSheet.Shapes.Range(Array(DimLine1.Name, DimLine2.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawHorDimensions = Selection.ShapeRange.Group

End Function
'
Function DrawHorDimensionsT(StartX As Double, StartY As Double, EndX As Double, DimText As String, Optional Extension = 17) As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
txtDimText1 = DimText

' Draw Dimensions
Dim P1(2) As Double
Dim P2(2) As Double
Dim P3(2) As Double
Dim P4(2) As Double
Dim DimLine1 As Shape
Dim DimLine2 As Shape
Dim DimLine3 As Shape

Dim txtP1(2) As Double


Dim txtWidth1 As Long
Dim txtHeight1 As Long

Dim DimText1 As Shape

P1(1) = StartX
P1(2) = StartY
P2(1) = StartX
P2(2) = StartY - Extension

P3(1) = EndX
P3(2) = StartY
P4(1) = EndX
P4(2) = StartY - Extension

Set DimLine1 = ActiveSheet.Shapes.AddLine(P1(1), P1(2), P2(1), P2(2))
Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
Set DimLine3 = ActiveSheet.Shapes.AddLine(P1(1), P2(2) + 5, P3(1), P2(2) + 5)
DimLine3.Line.BeginArrowheadStyle = msoArrowheadTriangle
DimLine3.Line.EndArrowheadStyle = msoArrowheadTriangle
'
'draw dimension text 1
txtP1(1) = StartX
txtP1(2) = StartY - 18
txtWidth1 = EndX - StartX
txtHeight1 = 12

Set DimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtWidth1, txtHeight1)
DimText1.TextFrame.Characters.Text = txtDimText1
DimText1.TextFrame.MarginBottom = 0
DimText1.TextFrame.MarginLeft = 0
DimText1.TextFrame.MarginRight = 0
DimText1.TextFrame.MarginTop = 0
DimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
DimText1.TextFrame.VerticalAlignment = xlVAlignCenter
DimText1.TextFrame.AutoSize = True
DimText1.Line.Visible = msoFalse
'
ActiveSheet.Shapes.Range(Array(DimLine1.Name, DimLine2.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawHorDimensionsT = Selection.ShapeRange.Group

End Function
Function DrawVerDimensions(StartX As Double, StartY As Double, EndY As Double, DimText As String, Optional Extension = 28) As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
txtDimText1 = DimText

' Draw Dimensions
Dim P1(2) As Double
Dim P2(2) As Double
Dim P3(2) As Double
Dim P4(2) As Double
Dim DimLine1 As Shape
Dim DimLine2 As Shape
Dim DimLine3 As Shape

Dim txtP1(2) As Double


Dim txtWidth1 As Long
Dim txtHeight1 As Long

Dim DimText1 As Shape

P1(1) = StartX
P1(2) = StartY
P2(1) = StartX + Extension
P2(2) = StartY

P3(1) = StartX
P3(2) = EndY
P4(1) = StartX + Extension
P4(2) = EndY

Set DimLine1 = ActiveSheet.Shapes.AddLine(P1(1), P1(2), P2(1), P2(2))
Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
Set DimLine3 = ActiveSheet.Shapes.AddLine(P2(1) - 5, P2(2), P4(1) - 5, P4(2))
DimLine3.Line.BeginArrowheadStyle = msoArrowheadTriangle
DimLine3.Line.EndArrowheadStyle = msoArrowheadTriangle
'
'draw dimension text 1
txtP1(1) = StartX + (Extension - 20)
txtP1(2) = StartY
txtWidth1 = EndY - StartY
txtHeight1 = 12

Set DimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtHeight1, txtWidth1)
DimText1.TextFrame.Characters.Text = txtDimText1
DimText1.TextFrame.Orientation = msoTextOrientationUpward
DimText1.TextFrame.MarginBottom = 0
DimText1.TextFrame.MarginLeft = 0
DimText1.TextFrame.MarginRight = 0
DimText1.TextFrame.MarginTop = 0
DimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
DimText1.TextFrame.VerticalAlignment = xlVAlignCenter
DimText1.TextFrame.AutoSize = True
DimText1.Line.Visible = msoFalse
'
ActiveSheet.Shapes.Range(Array(DimLine1.Name, DimLine2.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawVerDimensions = Selection.ShapeRange.Group

End Function
'
Function DrawVerDimensionsL(StartX As Double, StartY As Double, EndY As Double, DimText As String, Optional Extension = 22) As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
txtDimText1 = DimText

' Draw Dimensions
Dim P1(2) As Double
Dim P2(2) As Double
Dim P3(2) As Double
Dim P4(2) As Double
Dim DimLine1 As Shape
Dim DimLine2 As Shape
Dim DimLine3 As Shape

Dim txtP1(2) As Double


Dim txtWidth1 As Long
Dim txtHeight1 As Long

Dim DimText1 As Shape

P1(1) = StartX
P1(2) = StartY
P2(1) = StartX - Extension
P2(2) = StartY

P3(1) = StartX
P3(2) = EndY
P4(1) = StartX - Extension
P4(2) = EndY

Set DimLine1 = ActiveSheet.Shapes.AddLine(P1(1), P1(2), P2(1), P2(2))
Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
Set DimLine3 = ActiveSheet.Shapes.AddLine(P2(1) + 5, P2(2), P4(1) + 5, P4(2))
DimLine3.Line.BeginArrowheadStyle = msoArrowheadTriangle
DimLine3.Line.EndArrowheadStyle = msoArrowheadTriangle
'
'draw dimension text 1
txtP1(1) = StartX - (Extension + 8)
txtP1(2) = StartY
txtWidth1 = EndY - StartY
txtHeight1 = 12

Set DimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtHeight1, txtWidth1)
DimText1.TextFrame.Characters.Text = txtDimText1
DimText1.TextFrame.Orientation = msoTextOrientationUpward
DimText1.TextFrame.MarginBottom = 0
DimText1.TextFrame.MarginLeft = 0
DimText1.TextFrame.MarginRight = 0
DimText1.TextFrame.MarginTop = 0
DimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
DimText1.TextFrame.VerticalAlignment = xlVAlignCenter
DimText1.TextFrame.AutoSize = True
DimText1.Line.Visible = msoFalse
'
ActiveSheet.Shapes.Range(Array(DimLine1.Name, DimLine2.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawVerDimensionsL = Selection.ShapeRange.Group

End Function

Function DrawLeader1(StartX As Double, StartY As Double, EndX As Double, EndY As Double, LeaderText As String) As Shape
Dim P1(2) As Double
Dim P2(2) As Double
Dim P3(2) As Double
Dim P4(2) As Double
'
Dim txtP1(2) As Double
Dim txtWidth1 As Long
Dim txtHeight1 As Long
'
Dim Line1 As Shape
Dim Line2 As Shape
Dim Text1 As Shape
'
Set Line1 = ActiveSheet.Shapes.AddLine(StartX, StartY, EndX, EndY)
Line1.Line.BeginArrowheadStyle = msoArrowheadTriangle

Set Line2 = ActiveSheet.Shapes.AddLine(EndX, EndY, EndX - 30, EndY)

'draw dimension text 1
txtP1(1) = EndX - 30
txtP1(2) = EndY - 15
txtWidth1 = 30
txtHeight1 = 12

Set Text1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtWidth1, txtHeight1)
Text1.TextFrame.Characters.Text = LeaderText
Text1.TextFrame.MarginBottom = 0
Text1.TextFrame.MarginLeft = 0
Text1.TextFrame.MarginRight = 0
Text1.TextFrame.MarginTop = 0
Text1.TextFrame.HorizontalAlignment = xlHAlignCenter
Text1.TextFrame.VerticalAlignment = xlVAlignCenter
Text1.TextFrame.AutoSize = True
Text1.Line.Visible = msoFalse

ActiveSheet.Shapes.Range(Array(Line1.Name, Line2.Name, Text1.Name)).Select
Set DrawLeader1 = Selection.ShapeRange.Group
End Function
Function DrawLeader2(StartX As Double, StartY As Double, EndX As Double, EndY As Double, LeaderText As String) As Shape
Dim P1(2) As Double
Dim P2(2) As Double
Dim P3(2) As Double
Dim P4(2) As Double
'
Dim txtP1(2) As Double
Dim txtWidth1 As Long
Dim txtHeight1 As Long
'
Dim Line1 As Shape
Dim Line2 As Shape
Dim Text1 As Shape
'
Set Line1 = ActiveSheet.Shapes.AddLine(StartX, StartY, EndX, EndY)
Line1.Line.BeginArrowheadStyle = msoArrowheadTriangle

Set Line2 = ActiveSheet.Shapes.AddLine(EndX, EndY, EndX + 30, EndY)

'draw dimension text 1
txtP1(1) = EndX
txtP1(2) = EndY - 15
txtWidth1 = 30
txtHeight1 = 12

Set Text1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtWidth1, txtHeight1)
Text1.TextFrame.Characters.Text = LeaderText
Text1.TextFrame.MarginBottom = 0
Text1.TextFrame.MarginLeft = 0
Text1.TextFrame.MarginRight = 0
Text1.TextFrame.MarginTop = 0
Text1.TextFrame.HorizontalAlignment = xlHAlignCenter
Text1.TextFrame.VerticalAlignment = xlVAlignCenter
Text1.TextFrame.AutoSize = True
Text1.Line.Visible = msoFalse

ActiveSheet.Shapes.Range(Array(Line1.Name, Line2.Name, Text1.Name)).Select
Set DrawLeader2 = Selection.ShapeRange.Group
End Function
Sub shape_1(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight

'Stop
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"


txtDimText1 = Cells(RowIndex, ColA).Value
' Calculate The length
If IsNumeric(txtDimText1) Then
Cells(RowIndex, ColUnitLength).Value = txtDimText1
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 20


Point1(1) = X0 + 35: Point1(2) = Y0 + 40
Point2(1) = X0 + 150: Point2(2) = Y0 + 40

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'=======================================
'=======================================
' Draw Dimensions
Set DimShape = DrawHorDimensions(Point1(1), Point1(2), Point2(1), txtDimText1)
'
ActiveSheet.Shapes.Range(Array(shLine1.Name, DimShape.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group

'Set ShapeArray(RowIndex) = Gr1
'ShapeArrayNames(RowIndex) = Gr1.Name
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub

Sub shape_2(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight



' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value


txtLeader = Cells(RowIndex, ColBendDiameter).Value
' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 35: Point1(2) = Y0 + 60
Point2(1) = X0 + 138: Point2(2) = Y0 + 60
Point3(1) = X0 + 150: Point3(2) = Y0 + 48
Point4(1) = X0 + 150: Point4(2) = Y0 + 20


' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc As Shape
Dim CenterX: CenterX = X0 + 138
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 12
Set Arc = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc.Adjustments.Item(1) = 0
Arc.Adjustments.Item(2) = 90
Arc.Line.Weight = ShapeWeight: Arc.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point1(1), Point1(2), Point3(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensions(Point4(1), Point4(2), Point2(2) + ShapeWeight / 4, txtDimText2)

'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point2(1) + Point3(1)) / 2 + 1
LeaderStartPoint(2) = (Point2(2) + Point3(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, Arc.Name, DimShape1.Name, DimShape2.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub

'
Sub shape_3(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight



' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 35: Point1(2) = Y0 + 20
Point2(1) = X0 + 35: Point2(2) = Y0 + 48
Point3(1) = X0 + 47: Point3(2) = Y0 + 60
Point4(1) = X0 + 138: Point4(2) = Y0 + 60
Point5(1) = X0 + 150: Point5(2) = Y0 + 48
Point6(1) = X0 + 150: Point6(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 90
Arc1.Adjustments.Item(2) = 180
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 0
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point2(1) - ShapeWeight / 4, Point3(2), Point5(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2), Point3(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point6(1), Point6(2), Point4(2) + ShapeWeight / 4, txtDimText3)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
'
Sub shape_4(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight



' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) And IsNumeric(txtDimText5) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4) + CDbl(txtDimText5)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 25: Point1(2) = Y0 + 20
Point2(1) = X0 + 53: Point2(2) = Y0 + 20
Point3(1) = X0 + 65: Point3(2) = Y0 + 32
Point4(1) = X0 + 65: Point4(2) = Y0 + 48
Point5(1) = X0 + 77: Point5(2) = Y0 + 60
Point6(1) = X0 + 108: Point6(2) = Y0 + 60
Point7(1) = X0 + 120: Point7(2) = Y0 + 48
Point8(1) = X0 + 120: Point8(2) = Y0 + 32
Point9(1) = X0 + 132: Point9(2) = Y0 + 20
Point10(1) = X0 + 160: Point10(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 53
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 270
Arc1.Adjustments.Item(2) = 0
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 77
CenterY = Y0 + 48
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 108
CenterY = Y0 + 48
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 4
CenterX = X0 + 132
CenterY = Y0 + 32
Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 180
Arc4.Adjustments.Item(2) = 270
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2, 12)
Set DimShape3 = DrawVerDimensions(Point10(1), Point10(2) - ShapeWeight / 4, Point6(2) + ShapeWeight / 4, txtDimText3, 18)
Set DimShape4 = DrawHorDimensionsT(Point1(1), Point1(2), Point3(1) + ShapeWeight / 4, txtDimText4)
Set DimShape5 = DrawHorDimensionsT(Point8(1) - ShapeWeight / 4, Point10(2), Point10(1), txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, DimShape5.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
'
Sub shape_5(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight



' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) And IsNumeric(txtDimText5) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4) + CDbl(txtDimText5)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 75: Point1(2) = Y0 + 20
Point2(1) = X0 + 47: Point2(2) = Y0 + 20
Point3(1) = X0 + 35: Point3(2) = Y0 + 32
Point4(1) = X0 + 35: Point4(2) = Y0 + 48
Point5(1) = X0 + 47: Point5(2) = Y0 + 60
Point6(1) = X0 + 138: Point6(2) = Y0 + 60
Point7(1) = X0 + 150: Point7(2) = Y0 + 48
Point8(1) = X0 + 150: Point8(2) = Y0 + 32
Point9(1) = X0 + 138: Point9(2) = Y0 + 20
Point10(1) = X0 + 110: Point10(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 180
Arc1.Adjustments.Item(2) = 270
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 47
CenterY = Y0 + 48
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 138
CenterY = Y0 + 48
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 4
CenterX = X0 + 138
CenterY = Y0 + 32
Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 270
Arc4.Adjustments.Item(2) = 0
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point3(1), Point2(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point8(1), Point9(2) - ShapeWeight / 4, Point6(2) + ShapeWeight / 4, txtDimText3)
Set DimShape4 = DrawHorDimensionsT(Point3(1) - ShapeWeight / 4, Point2(2), Point1(1) + ShapeWeight / 4, txtDimText4)
Set DimShape5 = DrawHorDimensionsT(Point10(1), Point10(2), Point8(1) + ShapeWeight / 4, txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, DimShape5.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub

Sub shape_6(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"


txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 60: Point1(2) = Y0 + 20
Point2(1) = X0 + 40: Point2(2) = Y0 + 40
Point3(1) = X0 + 49: Point3(2) = Y0 + 60
Point4(1) = X0 + 138: Point4(2) = Y0 + 60
Point5(1) = X0 + 150: Point5(2) = Y0 + 48
Point6(1) = X0 + 150: Point6(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 49
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 90
Arc1.Adjustments.Item(2) = 225
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 0
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 35, Point3(2), X0 + 150, txtDimText1)
Set DimShape2 = DrawAlDimensionsL(X0 + 31, Y0 + 49, Point1(1), Point1(2), txtDimText2)
Set DimShape3 = DrawVerDimensions(Point6(1), Point6(2), Point4(2) + ShapeWeight / 4, txtDimText3)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub

Sub shape_7(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"


txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If


X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 60: Point1(2) = Y0 + 20
Point2(1) = X0 + 40: Point2(2) = Y0 + 40
Point3(1) = X0 + 49: Point3(2) = Y0 + 60
Point4(1) = X0 + 138: Point4(2) = Y0 + 60
Point5(1) = X0 + 145: Point5(2) = Y0 + 40
Point6(1) = X0 + 125: Point6(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 49
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 90
Arc1.Adjustments.Item(2) = 225
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 136
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -45
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 35, Point3(2), X0 + 150, txtDimText1)
Set DimShape2 = DrawAlDimensionsL(X0 + 31, Y0 + 49, Point1(1), Point1(2), txtDimText2)
Set DimShape3 = DrawAlDimensionsR(X0 + 154, Y0 + 49, Point6(1), Point6(2), txtDimText3)
'Set DimShape3 = DrawAlDimensionsR(X0 + 154, Y0 + 49, Point4(1), Point4(2), txtDimText3)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 6, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_8(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
'
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) And IsNumeric(txtDimText5) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4) + CDbl(txtDimText5)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12
Point1(1) = X0 + 55: Point1(2) = Y0 + 20
Point2(1) = X0 + 38.51: Point2(2) = Y0 + 36.49
Point3(1) = X0 + 35: Point3(2) = Y0 + 44.97
Point4(1) = X0 + 35: Point4(2) = Y0 + 64
Point5(1) = X0 + 47: Point5(2) = Y0 + 76
Point6(1) = X0 + 138: Point6(2) = Y0 + 76
Point7(1) = X0 + 150: Point7(2) = Y0 + 64
Point8(1) = X0 + 150: Point8(2) = Y0 + 44.97
Point9(1) = X0 + 146.49: Point9(2) = Y0 + 36.49
Point10(1) = X0 + 130: Point10(2) = Y0 + 20
'
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 64
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 90
Arc1.Adjustments.Item(2) = 180
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 64

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 0
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc 3
CenterX = X0 + 47
CenterY = Y0 + 44.97

Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 180
Arc3.Adjustments.Item(2) = 225
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'Draw Arc 4
CenterX = X0 + 138
CenterY = Y0 + 44.97

Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = -45
Arc4.Adjustments.Item(2) = 0
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 35, Point5(2), X0 + 150, txtDimText1)
Set DimShape2 = DrawAlDimensionsL(X0 + 35, Y0 + 40, Point1(1), Point1(2), txtDimText2)
Set DimShape3 = DrawAlDimensionsR(X0 + 150, Y0 + 40, Point10(1), Point10(2), txtDimText3)
'
'Set DimShape4 = DrawVerDimensions(X0 + 150, Y0 + 76, X0 + 40, txtDimText4)
Set DimShape4 = DrawVerDimensions(X0 + 150, Y0 + 40, Y0 + 76, txtDimText4)
Set DimShape5 = DrawVerDimensionsL(X0 + 35, Y0 + 40, Y0 + 76, txtDimText5)

'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, DimShape5.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub shape_9(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String

'
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"


txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value


txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12
Point1(1) = X0 + 55: Point1(2) = Y0 + 20
Point2(1) = X0 + 38.51: Point2(2) = Y0 + 36.49
Point3(1) = X0 + 35: Point3(2) = Y0 + 44.97
Point4(1) = X0 + 35: Point4(2) = Y0 + 64
Point5(1) = X0 + 47: Point5(2) = Y0 + 76
Point6(1) = X0 + 138: Point6(2) = Y0 + 76
Point7(1) = X0 + 150: Point7(2) = Y0 + 64
Point8(1) = X0 + 150: Point8(2) = Y0 + 30

'
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)




'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape


Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 64
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 90
Arc1.Adjustments.Item(2) = 180
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 64

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 0
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc 3
CenterX = X0 + 47
CenterY = Y0 + 44.97

Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 180
Arc3.Adjustments.Item(2) = 225
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 35, Point5(2), X0 + 150, txtDimText1)
Set DimShape2 = DrawAlDimensionsL(X0 + 35, Y0 + 40, Point1(1), Point1(2), txtDimText2)

'
'Set DimShape4 = DrawVerDimensions(X0 + 150, Y0 + 76, X0 + 40, txtDimText4)
Set DimShape3 = DrawVerDimensions(X0 + 150, Y0 + 30, Y0 + 76, txtDimText3)
Set DimShape4 = DrawVerDimensionsL(X0 + 35, Y0 + 40, Y0 + 76, txtDimText4)

'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, Arc1.Name, Arc2.Name, Arc3.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_10(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'
Dim DimShape6 As Shape
Dim DimShape7 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
'
Dim txtDimText6 As String
Dim txtDimText7 As String
'
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"
'
If Cells(RowIndex, ColF).Value = "" Then Cells(RowIndex, ColF).Value = "F"
If Cells(RowIndex, ColG).Value = "" Then Cells(RowIndex, ColG).Value = "G"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
txtDimText5 = Cells(RowIndex, ColE).Value
'
txtDimText6 = Cells(RowIndex, ColF).Value
txtDimText7 = Cells(RowIndex, ColG).Value
'
txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) And IsNumeric(txtDimText5) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4) + CDbl(txtDimText5)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12
'Point1(1) = X0 + 35: Point1(2) = Y0 + 30
'Point2(1) = X0 + 55: Point2(2) = Y0 + 30
'Point3(1) = X0 + 63.5: Point3(2) = Y0 + 33.5
'Point4(1) = X0 + 76.5: Point4(2) = Y0 + 46.5
'Point5(1) = X0 + 85: Point5(2) = Y0 + 50
'Point6(1) = X0 + 100: Point6(2) = Y0 + 50
'Point7(1) = X0 + 108.5: Point7(2) = Y0 + 46.5
'Point8(1) = X0 + 121.5: Point8(2) = Y0 + 33.5
'Point9(1) = X0 + 130: Point9(2) = Y0 + 30
'Point10(1) = X0 + 150: Point10(2) = Y0 + 30

Point1(1) = X0 + 10: Point1(2) = Y0 + 30
Point2(1) = X0 + 40: Point2(2) = Y0 + 30
Point3(1) = X0 + 48.5: Point3(2) = Y0 + 33.5
Point4(1) = X0 + 61.5: Point4(2) = Y0 + 46.5
Point5(1) = X0 + 70: Point5(2) = Y0 + 50
Point6(1) = X0 + 115: Point6(2) = Y0 + 50
Point7(1) = X0 + 123.5: Point7(2) = Y0 + 46.5
Point8(1) = X0 + 136.5: Point8(2) = Y0 + 33.5
Point9(1) = X0 + 145: Point9(2) = Y0 + 30
Point10(1) = X0 + 175: Point10(2) = Y0 + 30
'
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 40
Dim CenterY: CenterY = Y0 + 42
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 270
Arc1.Adjustments.Item(2) = -45
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 70
CenterY = Y0 + 38

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -270
Arc2.Adjustments.Item(2) = -224
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc 3
CenterX = X0 + 115
CenterY = Y0 + 38

Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 45
Arc3.Adjustments.Item(2) = -270
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'Draw Arc 4
CenterX = X0 + 145
CenterY = Y0 + 42

Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 225
Arc4.Adjustments.Item(2) = -90
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 65, Point5(2), X0 + 120, txtDimText1)
Set DimShape2 = DrawHorDimensionsT(Point1(1), Point1(2), X0 + 45, txtDimText2)
Set DimShape3 = DrawHorDimensionsT(X0 + 140, Point10(2), Point10(1), txtDimText3)
Set DimShape4 = DrawAlDimensionsR(X0 + 65, Y0 + 49, X0 + 47, Y0 + 31, txtDimText4)
Set DimShape5 = DrawAlDimensionsL(X0 + 120, Y0 + 49, X0 + 140, Y0 + 31, txtDimText5)
'
'Set DimShape6 = DrawVerDimensionsL(X0 + 25, Y0 + 30, Y0 + 50, txtDimText6, 20)
Set DimShape6 = MOD2.DrawVerDimensions2(X0 + 15, Y0 + 30, Y0 + 50, txtDimText6, 8)
Set DimShape7 = MOD2.DrawVerDimensions2(X0 + 155, Y0 + 30, Y0 + 50, txtDimText7)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader2(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) + 9, LeaderStartPoint(2) + 24, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, DimShape5.Name, SLeader.Name, DimShape6.Name, DimShape7.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_11(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'
Dim DimShape4 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
'
Dim txtDimText4 As String
'
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'
txtDimText4 = Cells(RowIndex, ColD).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12
'Point1(1) = X0 + 35: Point1(2) = Y0 + 30
'Point2(1) = X0 + 55: Point2(2) = Y0 + 30
'Point3(1) = X0 + 63.5: Point3(2) = Y0 + 33.5
'Point4(1) = X0 + 76.5: Point4(2) = Y0 + 46.5
'Point5(1) = X0 + 85: Point5(2) = Y0 + 50
'Point6(1) = X0 + 100: Point6(2) = Y0 + 50
'Point7(1) = X0 + 108.5: Point7(2) = Y0 + 46.5
'Point8(1) = X0 + 121.5: Point8(2) = Y0 + 33.5
'Point9(1) = X0 + 130: Point9(2) = Y0 + 30
'Point10(1) = X0 + 150: Point10(2) = Y0 + 30

Point1(1) = X0 + 10: Point1(2) = Y0 + 20
Point2(1) = X0 + 45: Point2(2) = Y0 + 20
Point3(1) = X0 + 53.5: Point3(2) = Y0 + 23.5
Point4(1) = X0 + 86.5: Point4(2) = Y0 + 56.5
Point5(1) = X0 + 95: Point5(2) = Y0 + 60
Point6(1) = X0 + 175: Point6(2) = Y0 + 60

'
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)




'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape


Dim CenterX: CenterX = X0 + 45
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 270
Arc1.Adjustments.Item(2) = -45
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 95
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -270
Arc2.Adjustments.Item(2) = -224
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 90, Point5(2), Point6(1), txtDimText1)
Set DimShape2 = DrawHorDimensionsT(Point1(1), Point1(2), X0 + 49, txtDimText2)

Set DimShape3 = DrawAlDimensionsR(X0 + 90, Y0 + 59, X0 + 49, Y0 + 21, txtDimText3)
'
Set DimShape4 = MOD2.DrawVerDimensions3(X0 + 135, Y0 + 20, Y0 + 60, txtDimText4)
'

'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point2(1) + Point3(1)) / 2 + 1
LeaderStartPoint(2) = (Point2(2) + Point3(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) + 18, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name, DimShape4.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_12(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.05
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"


txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If


X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
Y0 = Y0 - 15


Point1(1) = X0 + 35: Point1(2) = Y0 + 20
Point2(1) = X0 + 35: Point2(2) = Y0 + 36
Point3(1) = X0 + 47: Point3(2) = Y0 + 48
Point4(1) = X0 + 138: Point4(2) = Y0 + 48
Point5(1) = X0 + 150: Point5(2) = Y0 + 60
Point6(1) = X0 + 150: Point6(2) = Y0 + 76

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 36
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 90
Arc1.Adjustments.Item(2) = 180
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 60

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 270
Arc2.Adjustments.Item(2) = 0
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 35, Point6(2), X0 + 150, txtDimText1, 17)
Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2), Point3(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point6(1), Point4(2) - ShapeWeight / 4, Point6(2), txtDimText3)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 - 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) + 17, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_13(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) And IsNumeric(txtDimText5) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4) + CDbl(txtDimText5)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 73.9: Point1(2) = Y0 + 41.9
Point2(1) = X0 + 55.5: Point2(2) = Y0 + 23.5
Point3(1) = X0 + 35: Point3(2) = Y0 + 32
Point4(1) = X0 + 35: Point4(2) = Y0 + 64
Point5(1) = X0 + 47: Point5(2) = Y0 + 76
Point6(1) = X0 + 138: Point6(2) = Y0 + 76
Point7(1) = X0 + 150: Point7(2) = Y0 + 64
Point8(1) = X0 + 150: Point8(2) = Y0 + 32
Point9(1) = X0 + 129.5: Point9(2) = Y0 + 23.5
Point10(1) = X0 + 111.1: Point10(2) = Y0 + 41.9

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 180
Arc1.Adjustments.Item(2) = -45
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 47
CenterY = Y0 + 64
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 138
CenterY = Y0 + 64
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 4
CenterX = X0 + 138
CenterY = Y0 + 32
Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 225
Arc4.Adjustments.Item(2) = 0
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point3(1), Y0 + 17, Point5(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point8(1), Y0 + 17, Point6(2) + ShapeWeight / 4, txtDimText3)
Set DimShape4 = DrawAlDimensionsR(Point1(1), Point1(2), X0 + 47, Y0 + 16, txtDimText4, 16)
Set DimShape5 = DrawAlDimensionsL(Point10(1), Point10(2), X0 + 140, Y0 + 16, txtDimText5, 16)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 8, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, DimShape5.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_14(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
'
Dim Point9(2) As Double
Dim Point10(2) As Double
Dim Point11(2) As Double
Dim Point12(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'Dim DimShape4 As Shape
'Dim DimShape5 As Shape
'Dim DimShape6 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
'Dim txtDimText4 As String
'Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"


txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) * 2 + CDbl(txtDimText2) * 2 + CDbl(txtDimText3) * 2
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 138: Point1(2) = Y0 + 20
Point2(1) = X0 + 47: Point2(2) = Y0 + 20
Point3(1) = X0 + 35: Point3(2) = Y0 + 32
Point4(1) = X0 + 35: Point4(2) = Y0 + 64
Point5(1) = X0 + 47: Point5(2) = Y0 + 76
Point6(1) = X0 + 138: Point6(2) = Y0 + 76
Point7(1) = X0 + 150: Point7(2) = Y0 + 64
Point8(1) = X0 + 150: Point8(2) = Y0 + 32
'
Point9(1) = X0 + 119: Point9(2) = Y0 + 20
Point10(1) = X0 + 110: Point10(2) = Y0 + 30
Point11(1) = X0 + 150: Point10(2) = Y0 + 51
Point12(1) = X0 + 141: Point10(2) = Y0 + 61
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'Stop
' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(X0 + 119, Y0 + 20, X0 + 110, Y0 + 30)
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
' Draw Line 6
Dim ShLine6 As Shape
Set ShLine6 = ActiveSheet.Shapes.AddLine(X0 + 150, Y0 + 51, X0 + 141, Y0 + 61)
ShLine6.Line.Weight = ShapeWeight: ShLine6.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 180
Arc1.Adjustments.Item(2) = 270
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 47
CenterY = Y0 + 64
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 138
CenterY = Y0 + 64
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 4
CenterX = X0 + 138
CenterY = Y0 + 32
Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 270
Arc4.Adjustments.Item(2) = 0
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point3(1), Point2(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point8(1), Point9(2) - ShapeWeight / 4, Y0 + 51, txtDimText3)
'Set DimShape4 = DrawHorDimensionsT(Point3(1) - ShapeWeight / 4, Point2(2), Point1(1) + ShapeWeight / 4, txtDimText4)
'Set DimShape5 = DrawHorDimensionsT(Point10(1), Point10(2), Point8(1) + ShapeWeight / 4, txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 - 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 + 1

Set SLeader = DrawLeader2(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) + 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, ShLine6.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_15(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
'
Dim Point9(2) As Double
Dim Point10(2) As Double
Dim Point11(2) As Double
Dim Point12(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'Dim DimShape4 As Shape
'Dim DimShape5 As Shape
'Dim DimShape6 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
'Dim txtDimText4 As String
'Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"


txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) * 2 + CDbl(txtDimText2) * 2 + CDbl(txtDimText3) * 2
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 138: Point1(2) = Y0 + 20
Point2(1) = X0 + 47: Point2(2) = Y0 + 20
Point3(1) = X0 + 35: Point3(2) = Y0 + 32
Point4(1) = X0 + 35: Point4(2) = Y0 + 64
Point5(1) = X0 + 47: Point5(2) = Y0 + 76
Point6(1) = X0 + 138: Point6(2) = Y0 + 76
Point7(1) = X0 + 150: Point7(2) = Y0 + 64
Point8(1) = X0 + 150: Point8(2) = Y0 + 32
'
Point9(1) = X0 + 119: Point9(2) = Y0 + 20
Point10(1) = X0 + 110: Point10(2) = Y0 + 30
Point11(1) = X0 + 150: Point10(2) = Y0 + 51
Point12(1) = X0 + 141: Point10(2) = Y0 + 61
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'Stop
' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), X0 + 117, Y0 + 41)
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
' Draw Line 6
Dim ShLine6 As Shape
Set ShLine6 = ActiveSheet.Shapes.AddLine(Point8(1), Point8(2), X0 + 129, Y0 + 53)
ShLine6.Line.Weight = ShapeWeight: ShLine6.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 180
Arc1.Adjustments.Item(2) = 270
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 47
CenterY = Y0 + 64
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 138
CenterY = Y0 + 64
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 4
CenterX = X0 + 138
CenterY = Y0 + 32
Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 270
Arc4.Adjustments.Item(2) = 0
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point3(1), Point2(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2)
'Set DimShape3 = DrawVerDimensions(Point8(1), Point9(2) - ShapeWeight / 4, Y0 + 51, txtDimText3)
Set DimShape3 = DrawAlDimensionsL(X0 + 117, Y0 + 41, X0 + 142, Y0 + 16, txtDimText3)

'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 - 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 + 1

Set SLeader = DrawLeader2(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) + 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, ShLine6.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_16(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"


txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) Then
Cells(RowIndex, ColUnitLength).Value = Pi * CDbl(txtDimText1) + CDbl(txtDimText2)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12
Point1(1) = X0 + 71.8: Point1(2) = Y0 + 29.2
Point2(1) = X0 + 78.4: Point2(2) = Y0 + 38.8
Point3(1) = X0 + 113.2: Point3(2) = Y0 + 29.2
Point4(1) = X0 + 106.6: Point4(2) = Y0 + 38.8

' Draw Line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Dim Arc1 As Shape
Dim CenterX: CenterX = X0 + 92.5
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 28
Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 0
Arc1.Adjustments.Item(2) = 0
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 64.5, Y0 + 76, X0 + 120.5, txtDimText1)
'
'************************************
'**** draw arc dimensions************
'************************************
Dim L1 As Shape
Dim L2 As Shape
Dim A1 As Shape
Dim R1 As Double
Dim ADimText1 As Shape

Set L1 = ActiveSheet.Shapes.AddLine(X0 + 71.8, Y0 + 29.2, X0 + 58.6, Y0 + 17.2)
Set L2 = ActiveSheet.Shapes.AddLine(X0 + 113.2, Y0 + 29.2, X0 + 126.4, Y0 + 17.2)
'L1.Line.Weight = ShapeWeight: L2.Line.Weight = ShapeWeight
R1 = 38
Set A1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - R1, R1, R1)
A1.Adjustments.Item(1) = 222
A1.Adjustments.Item(2) = -43
'A1.Line.Weight = ShapeWeight
A1.Line.BeginArrowheadStyle = msoArrowheadTriangle
A1.Line.EndArrowheadStyle = msoArrowheadTriangle
'
'draw dimension text 1
Dim tp1(2) As Double
Dim tWidth1 As Double
Dim tHeight1 As Double

tp1(1) = CenterX
tp1(2) = Y0 + 4
tWidth1 = 0
tHeight1 = 12

Set ADimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, tp1(1), tp1(2), tWidth1, tHeight1)
ADimText1.TextFrame.Characters.Text = txtDimText2
ADimText1.TextFrame.MarginBottom = 0
ADimText1.TextFrame.MarginLeft = 0
ADimText1.TextFrame.MarginRight = 0
ADimText1.TextFrame.MarginTop = 0
ADimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
ADimText1.TextFrame.VerticalAlignment = xlVAlignCenter
ADimText1.TextFrame.AutoSize = True
ADimText1.Line.Visible = msoFalse
'_________________________

ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, Arc1.Name, DimShape1.Name, L1.Name, L2.Name, A1.Name, ADimText1.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub

Sub shape_17(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String

Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
Y0 = Y0 - 7


Point1(1) = X0 + 35: Point1(2) = Y0 + 20
Point2(1) = X0 + 35: Point2(2) = Y0 + 48
Point3(1) = X0 + 47: Point3(2) = Y0 + 60
Point4(1) = X0 + 138: Point4(2) = Y0 + 60
Point5(1) = X0 + 138: Point5(2) = Y0 + 36
Point6(1) = X0 + 100: Point6(2) = Y0 + 36

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 90
Arc1.Adjustments.Item(2) = 180
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -90
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point2(1) - ShapeWeight / 4, Point3(2), Point5(1) + ShapeWeight / 4 + 12, txtDimText1)
Set DimShape4 = DrawHorDimensionsT(Point6(1), Point6(2), Point5(1) + 13, txtDimText4)
Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2), Point3(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point5(1) + 12, Point5(2) - 1.5, Point4(2) + ShapeWeight / 4, txtDimText3)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 12
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 65, LeaderStartPoint(2), "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub shape_18(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
'
Dim DimShape3 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String

Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
Y0 = Y0 + 10


Point1(1) = X0 + 35: Point1(2) = Y0 + 60
Point2(1) = X0 + 71.5: Point2(2) = Y0 + 23.5
Point3(1) = X0 + 80: Point3(2) = Y0 + 20
Point4(1) = X0 + 150: Point4(2) = Y0 + 20


' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)



'Draw Arc
Dim Arc1 As Shape

Dim CenterX: CenterX = X0 + 80
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 225
Arc1.Adjustments.Item(2) = -90
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawAlDimensionsL(Point1(1), Point1(2), X0 + 74.6, Y0 + 20.4, txtDimText1)
Set DimShape2 = DrawHorDimensionsT(X0 + 74.6, Point3(2), Point4(1), txtDimText2)
Set DimShape3 = DrawVerDimensions(Point4(1), Point4(2), Point1(2), txtDimText3)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point2(1) + Point3(1)) / 2
LeaderStartPoint(2) = (Point2(2) + Point3(2)) / 2

Set SLeader = DrawLeader2(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) + 15, LeaderStartPoint(2) + 25, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, Arc1.Name, DimShape1.Name, DimShape2.Name, SLeader.Name, DimShape3.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub
Sub Shape_19(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'
Dim DimShape4 As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
'
Dim txtDimText4 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'
txtDimText4 = Cells(RowIndex, ColD).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 35: Point1(2) = Y0 + 76
Point2(1) = X0 + 59.5: Point2(2) = Y0 + 51.5
Point3(1) = X0 + 68: Point3(2) = Y0 + 48
Point4(1) = X0 + 138: Point4(2) = Y0 + 48
Point5(1) = X0 + 150: Point5(2) = Y0 + 36
Point6(1) = X0 + 150: Point6(2) = Y0 + 12

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 68
Dim CenterY: CenterY = Y0 + 60
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 225
Arc1.Adjustments.Item(2) = -90
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 36

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 0
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawAlDimensionsL(Point1(1), Point1(2), X0 + 63.3, Y0 + 47.7, txtDimText1)
Set DimShape2 = DrawHorDimensions(X0 + 63.3, Point3(2), Point5(1), txtDimText2)
Set DimShape3 = DrawVerDimensions(Point6(1), Point6(2), Point4(2) + ShapeWeight / 4, txtDimText3)
'
Set DimShape4 = DrawVerDimensions(Point6(1), Point4(2) + ShapeWeight / 4, Point1(2), txtDimText4)

'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name, DimShape4.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name


End Sub
Sub Shape_20(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'
Dim DimShape4 As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
'
Dim txtDimText4 As String

Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'
txtDimText4 = Cells(RowIndex, ColD).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
X0 = X0 + 5
Y0 = Y0 + 8


Point1(1) = X0 + 35: Point1(2) = Y0 + 60
Point2(1) = X0 + 71.5: Point2(2) = Y0 + 23.5
Point3(1) = X0 + 80: Point3(2) = Y0 + 20
Point4(1) = X0 + 138: Point4(2) = Y0 + 20
Point5(1) = X0 + 150: Point5(2) = Y0 + 32
Point6(1) = X0 + 150: Point6(2) = Y0 + 52

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 80
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 225
Arc1.Adjustments.Item(2) = -90
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 32

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -90
Arc2.Adjustments.Item(2) = 0
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawAlDimensionsL(Point1(1), Point1(2), X0 + 74.6, Y0 + 20.4, txtDimText1)
Set DimShape2 = DrawHorDimensionsT(X0 + 74.6, Point3(2), Point5(1), txtDimText2)
Set DimShape3 = DrawVerDimensions(Point6(1), Point4(2) + ShapeWeight / 4, Point6(2), txtDimText3)
'
Set DimShape4 = DrawVerDimensionsL(Point1(1) - 5, Point3(2), Point1(2), txtDimText4)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 - 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) + 20, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name, DimShape4.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name


End Sub
Sub Shape_21(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
'Dim Point1(2) As Double
'Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String

Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'txtDimText4 = Cells(RowIndex, ColD).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
Y0 = Y0 - 7


'Point1(1) = X0 + 35: Point1(2) = Y0 + 20
'Point2(1) = X0 + 35: Point2(2) = Y0 + 48
Point3(1) = X0 + 35: Point3(2) = Y0 + 60
Point4(1) = X0 + 138: Point4(2) = Y0 + 60
Point5(1) = X0 + 138: Point5(2) = Y0 + 36
Point6(1) = X0 + 100: Point6(2) = Y0 + 36

' Draw line 1
'Dim shLine1 As Shape
'Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
'shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
'Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 12

'Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc1.Adjustments.Item(1) = 90
'Arc1.Adjustments.Item(2) = 180
'Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -90
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point3(1), Point3(2), Point5(1) + ShapeWeight / 4 + 12, txtDimText1)
Set DimShape2 = DrawHorDimensionsT(Point6(1), Point6(2), Point5(1) + 13, txtDimText3)
'Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2), Point3(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point5(1) + 12, Point5(2) - 1.5, Point4(2) + ShapeWeight / 4, txtDimText2)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 12
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 65, LeaderStartPoint(2), "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine2.Name, shLine3.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub

Sub Shape_22(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String

Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) * 2 + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
Y0 = Y0 - 7


Point1(1) = X0 + 70: Point1(2) = Y0 + 36
Point2(1) = X0 + 47: Point2(2) = Y0 + 36
Point3(1) = X0 + 47: Point3(2) = Y0 + 60
Point4(1) = X0 + 138: Point4(2) = Y0 + 60
Point5(1) = X0 + 138: Point5(2) = Y0 + 36
Point6(1) = X0 + 115: Point6(2) = Y0 + 36

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 90
Arc1.Adjustments.Item(2) = -90
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 138
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -90
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point2(1) - ShapeWeight / 4 - 12, Point3(2), Point5(1) + ShapeWeight / 4 + 12, txtDimText1)
Set DimShape4 = DrawHorDimensionsT(Point6(1), Point6(2), Point5(1) + 13, txtDimText4)
Set DimShape3 = DrawHorDimensionsT(Point2(1) - 13, Point2(2), Point1(1), txtDimText3)
'Set DimShape2 = DrawVerDimensionsL(Point2(1) - 12, Point1(2), Point3(2) + ShapeWeight / 4, txtDimText2)
Set DimShape2 = DrawVerDimensions(Point5(1) + 12, Point5(2) - 1.5, Point4(2) + ShapeWeight / 4, txtDimText2)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 12
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 44, LeaderStartPoint(2), "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub Shape_23(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
'If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


'Point1(1) = X0 + 25: Point1(2) = Y0 + 20
'Point2(1) = X0 + 53: Point2(2) = Y0 + 20
'Point3(1) = X0 + 65: Point3(2) = Y0 + 32
'Point4(1) = X0 + 65: Point4(2) = Y0 + 48
Point5(1) = X0 + 25: Point5(2) = Y0 + 60
Point6(1) = X0 + 108: Point6(2) = Y0 + 60
Point7(1) = X0 + 120: Point7(2) = Y0 + 48
Point8(1) = X0 + 120: Point8(2) = Y0 + 32
Point9(1) = X0 + 132: Point9(2) = Y0 + 20
Point10(1) = X0 + 160: Point10(2) = Y0 + 20

' Draw line 1
'Dim shLine1 As Shape
'Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
'shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
'Dim shLine2 As Shape
'Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
'shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
'Dim Arc1 As Shape
'Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 53
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

'Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc1.Adjustments.Item(1) = 270
'Arc1.Adjustments.Item(2) = 0
'Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
'CenterX = X0 + 77
'CenterY = Y0 + 48
'Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc2.Adjustments.Item(1) = 90
'Arc2.Adjustments.Item(2) = 180
'Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 108
CenterY = Y0 + 48
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 4
CenterX = X0 + 132
CenterY = Y0 + 32
Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 180
Arc4.Adjustments.Item(2) = 270
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point5(1), Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
'Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2, 12)
Set DimShape2 = DrawVerDimensions(Point10(1), Point10(2) - ShapeWeight / 4, Point6(2) + ShapeWeight / 4, txtDimText2, 18)
'Set DimShape4 = DrawHorDimensionsT(Point1(1), Point1(2), Point3(1) + ShapeWeight / 4, txtDimText4)
Set DimShape3 = DrawHorDimensionsT(Point8(1) - ShapeWeight / 4, Point10(2), Point10(1), txtDimText3)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine3.Name, shLine4.Name, ShLine5.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub Shape_24(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
'If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"


txtDimText1 = Cells(RowIndex, ColA).Value
'txtDimText2 = Cells(RowIndex, ColB).Value

If IsNumeric(txtDimText1) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
Y0 = Y0 + 5
'Point1(1) = X0 + 71.8: Point1(2) = Y0 + 29.2
'Point2(1) = X0 + 78.4: Point2(2) = Y0 + 38.8
'Point3(1) = X0 + 113.2: Point3(2) = Y0 + 29.2
'Point4(1) = X0 + 106.6: Point4(2) = Y0 + 38.8
'
'' Draw Line 1
'Dim shLine1 As Shape
'Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
'shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'' Draw Line 2
'Dim shLine2 As Shape
'Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
'shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Dim Arc1 As Shape
Dim CenterX: CenterX = X0 + 92.5
Dim CenterY: CenterY = Y0 + 116.9
'Dim Radius: Radius = 88.3
'Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc1.Adjustments.Item(1) = 0
'Arc1.Adjustments.Item(2) = 0
'Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'' Draw Dimensions
'Set DimShape1 = DrawHorDimensions(X0 + 64.5, Y0 + 76, X0 + 120.5, txtDimText1)
'
'************************************
'**** draw arc dimensions************
'************************************
Dim L1 As Shape
Dim L2 As Shape
Dim A1 As Shape
Dim A2 As Shape
Dim R1 As Double
Dim R2 As Double
Dim ADimText1 As Shape

Set L1 = ActiveSheet.Shapes.AddLine(X0 + 25, Y0 + 60, X0 + 8.8, Y0 + 46.3)
Set L2 = ActiveSheet.Shapes.AddLine(X0 + 160, Y0 + 60, X0 + 176.2, Y0 + 46.3)
'L1.Line.Weight = ShapeWeight: L2.Line.Weight = ShapeWeight
R1 = 105
Set A1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - R1, R1, R1)
A1.Adjustments.Item(1) = 220
A1.Adjustments.Item(2) = -40
'A1.Line.Weight = ShapeWeight
A1.Line.BeginArrowheadStyle = msoArrowheadTriangle
A1.Line.EndArrowheadStyle = msoArrowheadTriangle
'
R2 = 88.3
Set A2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - R2, R2, R2)
A2.Adjustments.Item(1) = 220
A2.Adjustments.Item(2) = -40
A2.Line.Weight = ShapeWeight: A2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'draw dimension text 1
Dim tp1(2) As Double
Dim tWidth1 As Double
Dim tHeight1 As Double

tp1(1) = CenterX
tp1(2) = Y0 + 6
tWidth1 = 0
tHeight1 = 12

Set ADimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, tp1(1), tp1(2), tWidth1, tHeight1)
ADimText1.TextFrame.Characters.Text = txtDimText1
ADimText1.TextFrame.MarginBottom = 0
ADimText1.TextFrame.MarginLeft = 0
ADimText1.TextFrame.MarginRight = 0
ADimText1.TextFrame.MarginTop = 0
ADimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
ADimText1.TextFrame.VerticalAlignment = xlVAlignCenter
ADimText1.TextFrame.AutoSize = True
ADimText1.Line.Visible = msoFalse
'_________________________

Dim txtLeader As String
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = X0 + 119
LeaderStartPoint(2) = Y0 + 33
txtLeader = Cells(RowIndex, ColBendDiameter).Value
Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) + 22, "D " & txtLeader)


ActiveSheet.Shapes.Range(Array(L1.Name, L2.Name, A1.Name, A2.Name, ADimText1.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub Shape_25(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
'
Dim DimShape5 As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
'
'
Dim txtDimText5 As String
'
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
'
txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12
'Point1(1) = X0 + 35: Point1(2) = Y0 + 30
'Point2(1) = X0 + 55: Point2(2) = Y0 + 30
'Point3(1) = X0 + 63.5: Point3(2) = Y0 + 33.5
'Point4(1) = X0 + 76.5: Point4(2) = Y0 + 46.5
'Point5(1) = X0 + 85: Point5(2) = Y0 + 50
'Point6(1) = X0 + 100: Point6(2) = Y0 + 50
'Point7(1) = X0 + 108.5: Point7(2) = Y0 + 46.5
'Point8(1) = X0 + 121.5: Point8(2) = Y0 + 33.5
'Point9(1) = X0 + 130: Point9(2) = Y0 + 30
'Point10(1) = X0 + 150: Point10(2) = Y0 + 30

Point1(1) = X0 + 10: Point1(2) = Y0 + 20
Point2(1) = X0 + 45: Point2(2) = Y0 + 20
Point3(1) = X0 + 53.5: Point3(2) = Y0 + 23.5
Point4(1) = X0 + 86.5: Point4(2) = Y0 + 56.5
Point5(1) = X0 + 95: Point5(2) = Y0 + 60
Point6(1) = X0 + 138: Point6(2) = Y0 + 60
Point7(1) = X0 + 150: Point7(2) = Y0 + 48
Point8(1) = X0 + 150: Point8(2) = Y0 + 20
'
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)




'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape

Dim CenterX: CenterX = X0 + 45
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 270
Arc1.Adjustments.Item(2) = -45
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 95
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -270
Arc2.Adjustments.Item(2) = -224
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc 3
CenterX = X0 + 138
CenterY = Y0 + 48
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 90, Point5(2), Point7(1) + 1, txtDimText1)
Set DimShape2 = DrawHorDimensionsT(Point1(1), Point1(2), X0 + 49, txtDimText2)

Set DimShape3 = DrawAlDimensionsR(X0 + 90, Y0 + 59, X0 + 49, Y0 + 21, txtDimText3)
Set DimShape4 = DrawVerDimensions(Point8(1), Point8(2), Point6(2) + 1, txtDimText4)
'
Set DimShape5 = MOD2.DrawVerDimensions4(Point2(1), Point2(2), Point5(2), txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 8, LeaderStartPoint(2) - 19, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, Arc1.Name, Arc2.Name, Arc3.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, SLeader.Name, DimShape5.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub Shape_26(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
'If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 85: Point1(2) = Y0 + 20
Point2(1) = X0 + 47: Point2(2) = Y0 + 20
Point3(1) = X0 + 35: Point3(2) = Y0 + 32
Point4(1) = X0 + 35: Point4(2) = Y0 + 48
Point5(1) = X0 + 47: Point5(2) = Y0 + 60
Point6(1) = X0 + 150: Point6(2) = Y0 + 60
'Point7(1) = X0 + 150: Point7(2) = Y0 + 48
'Point8(1) = X0 + 150: Point8(2) = Y0 + 32
'Point9(1) = X0 + 138: Point9(2) = Y0 + 20
'Point10(1) = X0 + 110: Point10(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'' Draw Line 4
'Dim shLine4 As Shape
'Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
'shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'' Draw Line 5
'Dim ShLine5 As Shape
'Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
'ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
'Dim Arc3 As Shape
'Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 180
Arc1.Adjustments.Item(2) = 270
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 47
CenterY = Y0 + 48
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

''Draw Arc 3
'CenterX = X0 + 138
'CenterY = Y0 + 48
'Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc3.Adjustments.Item(1) = 0
'Arc3.Adjustments.Item(2) = 90
'Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
''Draw Arc 4
'CenterX = X0 + 138
'CenterY = Y0 + 32
'Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc4.Adjustments.Item(1) = 270
'Arc4.Adjustments.Item(2) = 0
'Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point6(1), txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point3(1), Point2(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2)
'Set DimShape3 = DrawVerDimensions(Point8(1), Point9(2) - ShapeWeight / 4, Point6(2) + ShapeWeight / 4, txtDimText3)
Set DimShape3 = DrawHorDimensionsT(Point3(1) - ShapeWeight / 4, Point2(2), Point1(1) + ShapeWeight / 4, txtDimText3)
'Set DimShape5 = DrawHorDimensionsT(Point10(1), Point10(2), Point8(1) + ShapeWeight / 4, txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 - 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 + 1

Set SLeader = DrawLeader2(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) + 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub Shape_27(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
'Dim Point9(2) As Double
'Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
'If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 75: Point1(2) = Y0 + 20
Point2(1) = X0 + 47: Point2(2) = Y0 + 20
Point3(1) = X0 + 35: Point3(2) = Y0 + 32
Point4(1) = X0 + 35: Point4(2) = Y0 + 48
Point5(1) = X0 + 47: Point5(2) = Y0 + 60
Point6(1) = X0 + 138: Point6(2) = Y0 + 60
Point7(1) = X0 + 150: Point7(2) = Y0 + 48
Point8(1) = X0 + 150: Point8(2) = Y0 + 20
'Point9(1) = X0 + 138: Point9(2) = Y0 + 20
'Point10(1) = X0 + 110: Point10(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'' Draw Line 5
'Dim ShLine5 As Shape
'Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
'ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 180
Arc1.Adjustments.Item(2) = 270
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 47
CenterY = Y0 + 48
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 138
CenterY = Y0 + 48
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

''Draw Arc 4
'CenterX = X0 + 138
'CenterY = Y0 + 32
'Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc4.Adjustments.Item(1) = 270
'Arc4.Adjustments.Item(2) = 0
'Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point3(1), Point2(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point8(1), Point8(2), Point6(2) + ShapeWeight / 4, txtDimText3)
Set DimShape4 = DrawHorDimensionsT(Point3(1) - ShapeWeight / 4, Point2(2), Point1(1) + ShapeWeight / 4, txtDimText4)
'Set DimShape5 = DrawHorDimensionsT(Point10(1), Point10(2), Point8(1) + ShapeWeight / 4, txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, Arc1.Name, Arc2.Name, Arc3.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub Shape_28(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
'Dim Point9(2) As Double
'Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
'If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 25: Point1(2) = Y0 + 20
Point2(1) = X0 + 53: Point2(2) = Y0 + 20
Point3(1) = X0 + 65: Point3(2) = Y0 + 32
Point4(1) = X0 + 65: Point4(2) = Y0 + 48
Point5(1) = X0 + 77: Point5(2) = Y0 + 60
Point6(1) = X0 + 138: Point6(2) = Y0 + 60
Point7(1) = X0 + 150: Point7(2) = Y0 + 48
Point8(1) = X0 + 150: Point8(2) = Y0 + 20
'Point9(1) = X0 + 132: Point9(2) = Y0 + 20
'Point10(1) = X0 + 160: Point10(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
'Dim ShLine5 As Shape
'Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
'ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 53
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 270
Arc1.Adjustments.Item(2) = 0
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 77
CenterY = Y0 + 48
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 138
CenterY = Y0 + 48
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

''Draw Arc 4
'CenterX = X0 + 132
'CenterY = Y0 + 32
'Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc4.Adjustments.Item(1) = 180
'Arc4.Adjustments.Item(2) = 270
'Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2, 12)
Set DimShape3 = DrawVerDimensions(Point8(1), Point8(2), Point6(2) + ShapeWeight / 4, txtDimText3)
Set DimShape4 = DrawHorDimensionsT(Point1(1), Point1(2), Point3(1) + ShapeWeight / 4, txtDimText4)
'Set DimShape5 = DrawHorDimensionsT(Point8(1) - ShapeWeight / 4, Point10(2), Point10(1), txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, Arc1.Name, Arc2.Name, Arc3.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub Shape_29(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.25
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'
Dim Point01(2) As Double
Dim Point02(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape


'
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText5 As String
'
Dim txtDimText4 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText5 = Cells(RowIndex, ColD).Value
'
txtDimText4 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText5) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText5)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
X0 = X0 + 16
'Y0 = Y0 - 12


Point1(1) = X0 + 38.5: Point1(2) = Y0 + 72.5
Point2(1) = X0 + 59.5: Point2(2) = Y0 + 51.5
Point3(1) = X0 + 68: Point3(2) = Y0 + 48
Point4(1) = X0 + 125: Point4(2) = Y0 + 48
Point5(1) = X0 + 137: Point5(2) = Y0 + 36
Point6(1) = X0 + 137: Point6(2) = Y0 + 12
'
Point01(1) = X0 - 5: Point01(2) = Y0 + 76
Point02(1) = X0 + 30: Point02(2) = Y0 + 76

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point01(1), Point01(2), Point02(1), Point02(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim CenterX: CenterX = X0 + 68
Dim CenterY: CenterY = Y0 + 60
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 225
Arc1.Adjustments.Item(2) = -90
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 125
CenterY = Y0 + 36

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 0
Arc2.Adjustments.Item(2) = 90
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
CenterX = X0 + 30
CenterY = Y0 + 64

Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 45
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawAlDimensionsL(X0 + 35.7, Y0 + 75.3, X0 + 63.3, Y0 + 47.7, txtDimText1)
Set DimShape2 = DrawHorDimensions(X0 + 63.3, Point3(2), Point5(1), txtDimText2)
Set DimShape3 = DrawVerDimensions(Point6(1), Point6(2), Point4(2) + ShapeWeight / 4, txtDimText3)
'
Set DimShape4 = DrawVerDimensions(Point6(1), Point4(2) + ShapeWeight / 4, Point1(2), txtDimText4)
'
Set DimShape5 = DrawHorDimensions(Point01(1), Point01(2), X0 + 35.7, txtDimText5)

'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point4(1) + Point5(1)) / 2 + 1
LeaderStartPoint(2) = (Point4(2) + Point5(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, Arc1.Name, Arc2.Name, Arc3.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape5.Name, SLeader.Name, DimShape4.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name


End Sub
Sub Shape_30(RowIndex As Long)

Cells(RowIndex, ColA).RowHeight = RowHeight

' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) And IsNumeric(txtDimText5) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4) + CDbl(txtDimText5)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12

Point1(1) = X0 + 25: Point1(2) = Y0 + 20
Point2(1) = X0 + 53: Point2(2) = Y0 + 20
Point3(1) = X0 + 65: Point3(2) = Y0 + 32
Point4(1) = X0 + 65: Point4(2) = Y0 + 48
Point5(1) = X0 + 77: Point5(2) = Y0 + 60
Point6(1) = X0 + 140: Point6(2) = Y0 + 60
Point7(1) = X0 + 152: Point7(2) = Y0 + 48
Point8(1) = X0 + 152: Point8(2) = Y0 + 32
Point9(1) = X0 + 140: Point9(2) = Y0 + 20
Point10(1) = X0 + 100: Point10(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 53
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 270
Arc1.Adjustments.Item(2) = 0
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 77
CenterY = Y0 + 48
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 140
CenterY = Y0 + 48
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 0
Arc3.Adjustments.Item(2) = 90
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 4
CenterX = X0 + 140
CenterY = Y0 + 32
Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 270
Arc4.Adjustments.Item(2) = 0
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2, 12)
Set DimShape3 = DrawVerDimensions(Point9(1) + 20, Point10(2) - ShapeWeight / 4, Point6(2) + ShapeWeight / 4, txtDimText3, 18)
Set DimShape4 = DrawHorDimensionsT(Point1(1), Point1(2), Point3(1) + ShapeWeight / 4, txtDimText4)
Set DimShape5 = DrawHorDimensionsT(Point10(1), Point10(2), Point8(1) + ShapeWeight / 4, txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 + 1
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, DimShape5.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
Sub Shape_31(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.3
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
'Dim Point9(2) As Double
'Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
'If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 75 + 25: Point1(2) = Y0 + 20
Point2(1) = X0 + 47: Point2(2) = Y0 + 20
Point3(1) = X0 + 35: Point3(2) = Y0 + 32
Point4(1) = X0 + 35: Point4(2) = Y0 + 48
Point5(1) = X0 + 47: Point5(2) = Y0 + 60
Point6(1) = X0 + 138: Point6(2) = Y0 + 60
Point7(1) = X0 + 150: Point7(2) = Y0 + 72
Point8(1) = X0 + 150: Point8(2) = Y0 + 100
'Point9(1) = X0 + 138: Point9(2) = Y0 + 20
'Point10(1) = X0 + 110: Point10(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'' Draw Line 5
'Dim ShLine5 As Shape
'Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
'ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 180
Arc1.Adjustments.Item(2) = 270
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 47
CenterY = Y0 + 48
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 138
CenterY = Y0 + 72
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = -90
Arc3.Adjustments.Item(2) = 0
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

''Draw Arc 4
'CenterX = X0 + 138
'CenterY = Y0 + 32
'Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc4.Adjustments.Item(1) = 270
'Arc4.Adjustments.Item(2) = 0
'Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point3(1), Point2(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point8(1), Point6(2) + ShapeWeight / 4, Point8(2), txtDimText3)
Set DimShape4 = DrawHorDimensionsT(Point3(1) - ShapeWeight / 4, Point2(2), Point1(1) + ShapeWeight / 4, txtDimText4)
'Set DimShape5 = DrawHorDimensionsT(Point10(1), Point10(2), Point8(1) + ShapeWeight / 4, txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point6(1) + Point7(1)) / 2 - 2
LeaderStartPoint(2) = (Point6(2) + Point7(2)) / 2 - 2

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) + 34, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, Arc1.Name, Arc2.Name, Arc3.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub

Sub Shape_32(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.3
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
'Dim Point9(2) As Double
'Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
'If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 90 + 25: Point1(2) = Y0 + 100
Point2(1) = X0 + 90 + 25: Point2(2) = Y0 + 32
Point3(1) = X0 + 78 + 25: Point3(2) = Y0 + 20
Point4(1) = X0 + 47: Point4(2) = Y0 + 20
Point5(1) = X0 + 35: Point5(2) = Y0 + 32
Point6(1) = X0 + 35: Point6(2) = Y0 + 48
Point7(1) = X0 + 47: Point7(2) = Y0 + 60
Point8(1) = X0 + 150 + 5: Point8(2) = Y0 + 60
'Point9(1) = X0 + 138: Point9(2) = Y0 + 20
'Point10(1) = X0 + 110: Point10(2) = Y0 + 20

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'' Draw Line 5
'Dim ShLine5 As Shape
'Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
'ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 47
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 180
Arc1.Adjustments.Item(2) = 270
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 47
CenterY = Y0 + 48
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 90
Arc2.Adjustments.Item(2) = 180
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 78 + 25
CenterY = Y0 + 32
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = -90
Arc3.Adjustments.Item(2) = 0
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

''Draw Arc 4
'CenterX = X0 + 138
'CenterY = Y0 + 32
'Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
'Arc4.Adjustments.Item(1) = 270
'Arc4.Adjustments.Item(2) = 0
'Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point6(1) - ShapeWeight / 4, Point7(2), Point8(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensionsL(Point5(1), Point4(2) - ShapeWeight / 4, Point7(2) + ShapeWeight / 4, txtDimText2)
Set DimShape3 = DrawVerDimensions(Point1(1) + 42, Point3(2) + ShapeWeight / 4, Point1(2), txtDimText3)
Set DimShape4 = DrawHorDimensionsT(Point5(1) - ShapeWeight / 4, Point3(2), Point2(1) + ShapeWeight / 4, txtDimText4)
'Set DimShape5 = DrawHorDimensionsT(Point10(1), Point10(2), Point8(1) + ShapeWeight / 4, txtDimText5)
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point2(1) + Point3(1)) / 2 + 1
LeaderStartPoint(2) = (Point2(2) + Point3(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 10, LeaderStartPoint(2) + 17, "D " & txtLeader)
' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, Arc1.Name, Arc2.Name, Arc3.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, SLeader.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub

Sub Shape_99(RowIndex)

Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
Dim Point7(2) As Double
Dim Point8(2) As Double
Dim Point9(2) As Double
Dim Point10(2) As Double
'Define the group
Dim Gr1 As Shape
'Dim DimShape1 As Shape
'Dim DimShape2 As Shape
'Dim DimShape3 As Shape
'Dim DimShape4 As Shape
'Dim DimShape5 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
'Dim txtDimText4 As String
'Dim txtDimText5 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
'If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'txtDimText4 = Cells(RowIndex, ColD).Value
'txtDimText5 = Cells(RowIndex, ColE).Value

txtLeader = Cells(RowIndex, ColBendDiameter).Value

' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) Then
Cells(RowIndex, ColUnitLength).Value = (CDbl(txtDimText1) + CDbl(txtDimText2)) * 2 + CDbl(txtDimText3)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
Y0 = Y0 + 5


Point1(1) = X0 + 35: Point1(2) = Y0 + 76
Point2(1) = X0 + 60: Point2(2) = Y0 + 60.4
Point3(1) = X0 + 64.7: Point3(2) = Y0 + 51.9
Point4(1) = X0 + 64.7: Point4(2) = Y0 + 13.5
Point5(1) = X0 + 70.8: Point5(2) = Y0 + 10.1
Point6(1) = X0 + 103: Point6(2) = Y0 + 29.4
Point7(1) = X0 + 106.9: Point7(2) = Y0 + 36.2
Point8(1) = X0 + 106.9: Point8(2) = Y0 + 75.5
Point9(1) = X0 + 113: Point9(2) = Y0 + 78.9
Point10(1) = X0 + 136.6: Point10(2) = Y0 + 64.2

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point7(1), Point7(2), Point8(1), Point8(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 5
Dim ShLine5 As Shape
Set ShLine5 = ActiveSheet.Shapes.AddLine(Point9(1), Point9(2), Point10(1), Point10(2))
ShLine5.Line.Weight = ShapeWeight: ShLine5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'
'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape
Dim Arc3 As Shape
Dim Arc4 As Shape

Dim CenterX: CenterX = X0 + 54.7
Dim CenterY: CenterY = Y0 + 51.9
Dim Radius: Radius = 10

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 0
Arc1.Adjustments.Item(2) = 58
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 68.7
CenterY = Y0 + 13.5
Radius = 4
Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = 180
Arc2.Adjustments.Item(2) = -52
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 3
CenterX = X0 + 98.9
CenterY = Y0 + 36.2
Radius = 8
Set Arc3 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc3.Adjustments.Item(1) = 295
Arc3.Adjustments.Item(2) = 0
Arc3.Line.Weight = ShapeWeight: Arc3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 4
CenterX = X0 + 110.9
CenterY = Y0 + 75.5
Radius = 4
Set Arc4 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc4.Adjustments.Item(1) = 45
Arc4.Adjustments.Item(2) = 180
Arc4.Line.Weight = ShapeWeight: Arc4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
'Set DimShape1 = DrawHorDimensions(Point4(1) - ShapeWeight / 4, Point5(2), Point7(1) + ShapeWeight / 4, txtDimText1)
'Set DimShape2 = DrawVerDimensionsL(Point1(1), Point1(2) - ShapeWeight / 4, Point5(2) + ShapeWeight / 4, txtDimText2, 12)
'Set DimShape3 = DrawVerDimensions(Point8(1), Point8(2), Point6(2) + ShapeWeight / 4, txtDimText3)
'Set DimShape4 = DrawHorDimensionsT(Point1(1), Point1(2), Point3(1) + ShapeWeight / 4, txtDimText4)
'Set DimShape5 = DrawHorDimensionsT(Point8(1) - ShapeWeight / 4, Point10(2), Point10(1), txtDimText5)
'Draw Leader
Dim DL1 As Shape
Dim DL2 As Shape
Dim DL3 As Shape
Dim DL4 As Shape
Dim DL5 As Shape
Dim DL6 As Shape
Dim DL7 As Shape
'
Set DL1 = ActiveSheet.Shapes.AddLine(X0 + 35.9, Y0 + 76.5, X0 + 49.6, Y0 + 84.7)
Set DL2 = ActiveSheet.Shapes.AddLine(X0 + 66.4, Y0 + 57.5, X0 + 80.2, Y0 + 65.7)
Set DL3 = ActiveSheet.Shapes.AddLine(X0 + 68.1, Y0 + 7.2, X0 + 81.8, Y0 - 1.3)
Set DL4 = ActiveSheet.Shapes.AddLine(X0 + 107.6, Y0 + 30.8, X0 + 121.3, Y0 + 22.3)

Set DL5 = ActiveSheet.Shapes.AddLine(X0 + 45.8, Y0 + 82.4, X0 + 76.3, Y0 + 63.4)
DL5.Line.BeginArrowheadStyle = msoArrowheadTriangle
DL5.Line.EndArrowheadStyle = msoArrowheadTriangle

Set DL6 = ActiveSheet.Shapes.AddLine(X0 + 78, Y0 + 1.1, X0 + 117.5, Y0 + 24.6)
DL6.Line.BeginArrowheadStyle = msoArrowheadTriangle
DL6.Line.EndArrowheadStyle = msoArrowheadTriangle

Set DL7 = ActiveSheet.Shapes.AddLine(X0 + 117.5, Y0 + 74.6, X0 + 117.5, Y0 + 24.6)
DL7.Line.BeginArrowheadStyle = msoArrowheadTriangle
DL7.Line.EndArrowheadStyle = msoArrowheadTriangle


Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point2(1) + Point3(1)) / 2 - 1
LeaderStartPoint(2) = (Point2(2) + Point3(2)) / 2 - 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)

'''''''''' draw the texts ''''''''''''''
Dim DimText1 As Shape
Dim DimText2 As Shape
Dim DimText3 As Shape

Set DimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, X0 + 63, Y0 + 73, 12, 0)
DimText1.TextFrame.Characters.Text = txtDimText1
DimText1.TextFrame.MarginBottom = 0
DimText1.TextFrame.MarginLeft = 0
DimText1.TextFrame.MarginRight = 0
DimText1.TextFrame.MarginTop = 0
DimText1.TextFrame.AutoSize = True
DimText1.Line.Visible = msoFalse
'
Set DimText2 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, X0 + 114, Y0 + 40, 12, 0)
DimText2.TextFrame.Characters.Text = txtDimText2
DimText2.TextFrame.MarginBottom = 0
DimText2.TextFrame.MarginLeft = 0
DimText2.TextFrame.MarginRight = 0
DimText2.TextFrame.MarginTop = 0
DimText2.TextFrame.AutoSize = True
DimText2.Line.Visible = msoFalse
'
Set DimText3 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, X0 + 87, Y0 + 2, 12, 0)
DimText3.TextFrame.Characters.Text = txtDimText3
DimText3.TextFrame.MarginBottom = 0
DimText3.TextFrame.MarginLeft = 0
DimText3.TextFrame.MarginRight = 0
DimText3.TextFrame.MarginTop = 0
DimText3.TextFrame.AutoSize = True
DimText3.Line.Visible = msoFalse

' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, ShLine5.Name, Arc1.Name, Arc2.Name, Arc3.Name, Arc4.Name, SLeader.Name, DimText1.Name, DimText2.Name, DimText3.Name, DL1.Name, DL2.Name, DL3.Name, DL4.Name, DL5.Name, DL6.Name, DL7.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name


End Sub



Sub shape_33(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight

'Stop
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"


txtDimText1 = Cells(RowIndex, ColA).Value
' Calculate The length
If IsNumeric(txtDimText1) Then
Cells(RowIndex, ColUnitLength).Value = txtDimText1
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 20


Point1(1) = X0 + 35: Point1(2) = Y0 + 40
Point2(1) = X0 + 150: Point2(2) = Y0 + 40

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
' Draw Dimensions
Set DimShape = DrawHorDimensions(Point1(1), Point1(2), Point2(1), txtDimText1)
'
' add thread to straight

Dim thrdL1 As Shape
Dim thrdL2 As Shape
Dim thrdL3 As Shape
Dim thrdL4 As Shape
Dim thrdL5 As Shape

Set thrdL1 = ActiveSheet.Shapes.AddLine(Point1(1) + 4, Point1(2) - 6, Point1(1) + 4, Point1(2) + 6)
thrdL1.Line.Weight = ShapeWeight / 2: thrdL1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL2 = ActiveSheet.Shapes.AddLine(Point1(1) + 8, Point1(2) - 6, Point1(1) + 8, Point1(2) + 6)
thrdL2.Line.Weight = ShapeWeight / 2: thrdL2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL3 = ActiveSheet.Shapes.AddLine(Point1(1) + 12, Point1(2) - 6, Point1(1) + 12, Point1(2) + 6)
thrdL3.Line.Weight = ShapeWeight / 2: thrdL3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL4 = ActiveSheet.Shapes.AddLine(Point1(1) + 16, Point1(2) - 6, Point1(1) + 16, Point1(2) + 6)
thrdL4.Line.Weight = ShapeWeight / 2: thrdL4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL5 = ActiveSheet.Shapes.AddLine(Point1(1) + 20, Point1(2) - 6, Point1(1) + 20, Point1(2) + 6)
thrdL5.Line.Weight = ShapeWeight / 2: thrdL5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


ActiveSheet.Shapes.Range(Array(shLine1.Name, DimShape.Name, thrdL1.Name, thrdL2.Name, thrdL3.Name, thrdL4.Name, thrdL5.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group


Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub




Sub shape_34(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight



' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtLeader As String

If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value


txtLeader = Cells(RowIndex, ColBendDiameter).Value
' Calculate The length
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12


Point1(1) = X0 + 35: Point1(2) = Y0 + 60
Point2(1) = X0 + 138: Point2(2) = Y0 + 60
Point3(1) = X0 + 150: Point3(2) = Y0 + 48
Point4(1) = X0 + 150: Point4(2) = Y0 + 20


' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc
Dim Arc As Shape
Dim CenterX: CenterX = X0 + 138
Dim CenterY: CenterY = Y0 + 48
Dim Radius: Radius = 12
Set Arc = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc.Adjustments.Item(1) = 0
Arc.Adjustments.Item(2) = 90
Arc.Line.Weight = ShapeWeight: Arc.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(Point1(1), Point1(2), Point3(1) + ShapeWeight / 4, txtDimText1)
Set DimShape2 = DrawVerDimensions(Point4(1), Point4(2), Point2(2) + ShapeWeight / 4, txtDimText2)

'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point2(1) + Point3(1)) / 2 + 1
LeaderStartPoint(2) = (Point2(2) + Point3(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) - 12, "D " & txtLeader)



' add thread to L Bar

Dim thrdL1 As Shape
Dim thrdL2 As Shape
Dim thrdL3 As Shape
Dim thrdL4 As Shape
Dim thrdL5 As Shape

Set thrdL1 = ActiveSheet.Shapes.AddLine(Point1(1) + 4, Point1(2) - 6, Point1(1) + 4, Point1(2) + 6)
thrdL1.Line.Weight = ShapeWeight / 2: thrdL1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL2 = ActiveSheet.Shapes.AddLine(Point1(1) + 8, Point1(2) - 6, Point1(1) + 8, Point1(2) + 6)
thrdL2.Line.Weight = ShapeWeight / 2: thrdL2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL3 = ActiveSheet.Shapes.AddLine(Point1(1) + 12, Point1(2) - 6, Point1(1) + 12, Point1(2) + 6)
thrdL3.Line.Weight = ShapeWeight / 2: thrdL3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL4 = ActiveSheet.Shapes.AddLine(Point1(1) + 16, Point1(2) - 6, Point1(1) + 16, Point1(2) + 6)
thrdL4.Line.Weight = ShapeWeight / 2: thrdL4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL5 = ActiveSheet.Shapes.AddLine(Point1(1) + 20, Point1(2) - 6, Point1(1) + 20, Point1(2) + 6)
thrdL5.Line.Weight = ShapeWeight / 2: thrdL5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)



' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, Arc.Name, DimShape1.Name, DimShape2.Name, SLeader.Name, _
    thrdL1.Name, thrdL2.Name, thrdL3.Name, thrdL4.Name, thrdL5.Name)).Select
    
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
'
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub







Sub shape_35(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight

'Stop
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double

'Define the group
Dim Gr1 As Shape
Dim DimShape As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"


txtDimText1 = Cells(RowIndex, ColA).Value
' Calculate The length
If IsNumeric(txtDimText1) Then
Cells(RowIndex, ColUnitLength).Value = txtDimText1
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 20


Point1(1) = X0 + 35: Point1(2) = Y0 + 40
Point2(1) = X0 + 150: Point2(2) = Y0 + 40

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
' Draw Dimensions
Set DimShape = DrawHorDimensions(Point1(1), Point1(2), Point2(1), txtDimText1)
'
' add thread to straight

Dim thrdL1 As Shape
Dim thrdL2 As Shape
Dim thrdL3 As Shape
Dim thrdL4 As Shape
Dim thrdL5 As Shape

Set thrdL1 = ActiveSheet.Shapes.AddLine(Point1(1) + 4, Point1(2) - 6, Point1(1) + 4, Point1(2) + 6)
thrdL1.Line.Weight = ShapeWeight / 2: thrdL1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL2 = ActiveSheet.Shapes.AddLine(Point1(1) + 8, Point1(2) - 6, Point1(1) + 8, Point1(2) + 6)
thrdL2.Line.Weight = ShapeWeight / 2: thrdL2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL3 = ActiveSheet.Shapes.AddLine(Point1(1) + 12, Point1(2) - 6, Point1(1) + 12, Point1(2) + 6)
thrdL3.Line.Weight = ShapeWeight / 2: thrdL3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL4 = ActiveSheet.Shapes.AddLine(Point1(1) + 16, Point1(2) - 6, Point1(1) + 16, Point1(2) + 6)
thrdL4.Line.Weight = ShapeWeight / 2: thrdL4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL5 = ActiveSheet.Shapes.AddLine(Point1(1) + 20, Point1(2) - 6, Point1(1) + 20, Point1(2) + 6)
thrdL5.Line.Weight = ShapeWeight / 2: thrdL5.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)




Dim thrdL11 As Shape
Dim thrdL12 As Shape
Dim thrdL13 As Shape
Dim thrdL14 As Shape
Dim thrdL15 As Shape

Set thrdL11 = ActiveSheet.Shapes.AddLine(Point2(1) - 4, Point1(2) - 6, Point2(1) - 4, Point1(2) + 6)
thrdL11.Line.Weight = ShapeWeight / 2: thrdL11.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL12 = ActiveSheet.Shapes.AddLine(Point2(1) - 8, Point1(2) - 6, Point2(1) - 8, Point1(2) + 6)
thrdL12.Line.Weight = ShapeWeight / 2: thrdL12.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL13 = ActiveSheet.Shapes.AddLine(Point2(1) - 12, Point1(2) - 6, Point2(1) - 12, Point1(2) + 6)
thrdL13.Line.Weight = ShapeWeight / 2: thrdL13.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL14 = ActiveSheet.Shapes.AddLine(Point2(1) - 16, Point1(2) - 6, Point2(1) - 16, Point1(2) + 6)
thrdL14.Line.Weight = ShapeWeight / 2: thrdL14.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

Set thrdL15 = ActiveSheet.Shapes.AddLine(Point2(1) - 20, Point1(2) - 6, Point2(1) - 20, Point1(2) + 6)
thrdL15.Line.Weight = ShapeWeight / 2: thrdL15.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)




ActiveSheet.Shapes.Range(Array(shLine1.Name, DimShape.Name, thrdL1.Name, thrdL2.Name, thrdL3.Name, thrdL4.Name, thrdL5.Name, _
       thrdL11.Name, thrdL12.Name, thrdL13.Name, thrdL14.Name, thrdL15.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group


Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub












Sub shape_36(RowIndex As Long)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.8
' Origin point coordinate
Dim dSpace As Double: dSpace = 75
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double
Dim Point6(2) As Double
'
Dim Point11(2) As Double
Dim Point12(2) As Double
Dim Point13(2) As Double
Dim Point14(2) As Double
'
'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
'
Dim DimShape11 As Shape
Dim DimShape12 As Shape

'
Dim DimShape4 As Shape
'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
'
Dim txtDimText4 As String
'
Dim txtLeader As String

'
Dim txtDimText11 As String
Dim txtDimText12 As String
Dim txtLeader10 As String
'
If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
'
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"
'
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
'If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
'
txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
'
txtDimText4 = Cells(RowIndex, ColE).Value
'
txtDimText11 = Cells(RowIndex, ColD).Value
'
txtLeader = Cells(RowIndex, ColBendDiameter).Value
'
txtLeader10 = Cells(RowIndex, ColBendDiameter).Value
'
If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And _
 IsNumeric(txtDimText3) And IsNumeric(txtDimText11) Then
    Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText11)
Else
    Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top
'ALLIGN SHAPE
'X0 = X0 - 10
'Y0 = Y0 - 12
'Point1(1) = X0 + 35: Point1(2) = Y0 + 30
'Point2(1) = X0 + 55: Point2(2) = Y0 + 30
'Point3(1) = X0 + 63.5: Point3(2) = Y0 + 33.5
'Point4(1) = X0 + 76.5: Point4(2) = Y0 + 46.5
'Point5(1) = X0 + 85: Point5(2) = Y0 + 50
'Point6(1) = X0 + 100: Point6(2) = Y0 + 50
'Point7(1) = X0 + 108.5: Point7(2) = Y0 + 46.5
'Point8(1) = X0 + 121.5: Point8(2) = Y0 + 33.5
'Point9(1) = X0 + 130: Point9(2) = Y0 + 30
'Point10(1) = X0 + 150: Point10(2) = Y0 + 30

Point1(1) = X0 + 10: Point1(2) = Y0 + 20
Point2(1) = X0 + 45: Point2(2) = Y0 + 20
Point3(1) = X0 + 53.5: Point3(2) = Y0 + 23.5
Point4(1) = X0 + 86.5: Point4(2) = Y0 + 56.5
Point5(1) = X0 + 95: Point5(2) = Y0 + 60
Point6(1) = X0 + 150: Point6(2) = Y0 + 60
'
Point11(1) = X0 + 10: Point11(2) = Y0 + 60 + dSpace
Point12(1) = X0 + 138: Point12(2) = Y0 + 60 + dSpace
Point13(1) = X0 + 150: Point13(2) = Y0 + 48 + dSpace
Point14(1) = X0 + 150: Point14(2) = Y0 + 20 + dSpace
'

' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point5(1), Point5(2), Point6(1), Point6(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'
' Draw line 11
Dim shLine11 As Shape
Set shLine11 = ActiveSheet.Shapes.AddLine(Point11(1), Point11(2), Point12(1), Point12(2))
shLine11.Line.Weight = ShapeWeight: shLine11.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 12
Dim shLine12 As Shape
Set shLine12 = ActiveSheet.Shapes.AddLine(Point13(1), Point13(2), Point14(1), Point14(2))
shLine12.Line.Weight = ShapeWeight: shLine12.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)




'Draw Arc
Dim Arc1 As Shape
Dim Arc2 As Shape


Dim CenterX: CenterX = X0 + 45
Dim CenterY: CenterY = Y0 + 32
Dim Radius: Radius = 12

Set Arc1 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc1.Adjustments.Item(1) = 270
Arc1.Adjustments.Item(2) = -45
Arc1.Line.Weight = ShapeWeight: Arc1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

'Draw Arc 2
CenterX = X0 + 95
CenterY = Y0 + 48

Set Arc2 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc2.Adjustments.Item(1) = -270
Arc2.Adjustments.Item(2) = -224
Arc2.Line.Weight = ShapeWeight: Arc2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)
'

'Draw Arc 10
Dim Arc10 As Shape
CenterX = X0 + 138 + 0
CenterY = Y0 + 48 + dSpace
Radius = 12
Set Arc10 = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
Arc10.Adjustments.Item(1) = 0
Arc10.Adjustments.Item(2) = 90
Arc10.Line.Weight = ShapeWeight: Arc10.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)


'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawHorDimensions(X0 + 90, Point5(2), Point6(1), txtDimText1)
Set DimShape2 = DrawHorDimensionsT(Point1(1), Point1(2), X0 + 49, txtDimText2)

Set DimShape3 = DrawAlDimensionsR(X0 + 90, Y0 + 59, X0 + 49, Y0 + 21, txtDimText3)
'
Set DimShape4 = MOD2.DrawVerDimensions3(X0 + 110, Y0 + 20, Y0 + 60, txtDimText4)
'
Set DimShape11 = DrawVerDimensions(Point14(1), Point14(2), Point12(2) + ShapeWeight / 4, txtDimText11)
'
'Draw Leader
Dim SLeader As Shape
Dim LeaderStartPoint(2) As Double
LeaderStartPoint(1) = (Point2(1) + Point3(1)) / 2 + 1
LeaderStartPoint(2) = (Point2(2) + Point3(2)) / 2 + 1

Set SLeader = DrawLeader1(LeaderStartPoint(1), LeaderStartPoint(2), LeaderStartPoint(1) - 12, LeaderStartPoint(2) + 18, "D " & txtLeader)

'
'Draw Leader 10
Dim SLeader10 As Shape
Dim LeaderStartPoint10(2) As Double
LeaderStartPoint10(1) = (Point12(1) + Point13(1)) / 2 + 1
LeaderStartPoint10(2) = (Point12(2) + Point13(2)) / 2 + 1

Set SLeader10 = DrawLeader1(LeaderStartPoint10(1), LeaderStartPoint10(2), LeaderStartPoint10(1) - 12, LeaderStartPoint10(2) - 12, "D " & txtLeader10)
'

Dim DimText1 As Shape
Dim DimText2 As Shape
Dim txtWidth1 As Long
Dim txtHeight1 As Long
Dim txtP1(2) As Double

txtHeight1 = 12
txtWidth1 = 20
txtP1(1) = Point1(1) + 10
txtP1(2) = Point6(2)
Set DimText1 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtWidth1, txtHeight1)
DimText1.TextFrame.Characters.Text = "_PLAN_"
DimText1.TextFrame.MarginBottom = 0
DimText1.TextFrame.MarginLeft = 0
DimText1.TextFrame.MarginRight = 0
DimText1.TextFrame.MarginTop = 0
DimText1.TextFrame.HorizontalAlignment = xlHAlignCenter
DimText1.TextFrame.VerticalAlignment = xlVAlignCenter
DimText1.TextFrame.AutoSize = True
DimText1.Line.Visible = msoTrue
'
txtP1(1) = Point1(1) + 25
txtP1(2) = Point11(2) - 25
Set DimText2 = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, txtP1(1), txtP1(2), txtWidth1, txtHeight1)
DimText2.TextFrame.Characters.Text = "_ELEVATION_"
DimText2.TextFrame.MarginBottom = 0
DimText2.TextFrame.MarginLeft = 0
DimText2.TextFrame.MarginRight = 0
DimText2.TextFrame.MarginTop = 0
DimText2.TextFrame.HorizontalAlignment = xlHAlignCenter
DimText2.TextFrame.VerticalAlignment = xlVAlignCenter
DimText2.TextFrame.AutoSize = True
DimText2.Line.Visible = msoTrue
'
'expression. AddShape( _Type_ , _Left_ , _Top_ , _Width_ , _Height_ )
'add a circle
Dim sCircle As Shape
Dim dCirLeft As Double
Dim dCirTop As Double
Dim dCirWidth As Double
Dim dCirHeight As Double

'dCirLeft = Point6(1) - 12
'dCirTop = Point6(2) - 12
'dCirWidth = Point6(1) + 12
'dCirHeight = Point6(2) + 12
dCirLeft = Point6(1) - 6
dCirTop = Point6(2) - 6
dCirWidth = 12
dCirHeight = 12
Set sCircle = ActiveSheet.Shapes.AddShape(9, dCirLeft, dCirTop, dCirWidth, dCirHeight)

'format this circle
sCircle.Line.Visible = msoFalse
With sCircle.Fill
'give the shape a colour
.ForeColor.RGB = RGB(RrR, GgG, BbB)
End With

' Set the group
ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, _
Arc1.Name, Arc2.Name, DimShape1.Name, DimShape2.Name, DimShape3.Name, _
SLeader.Name, DimShape4.Name, _
shLine11.Name, shLine12.Name, Arc10.Name, DimShape11.Name, SLeader10.Name, _
DimText1.Name, DimText2.Name, sCircle.Name)).Select

Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name
End Sub






Sub Shape_37(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double


'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
'
Dim DimShape5 As Shape
Dim DimShape6 As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
'
'
Dim txtDimText5 As String
Dim txtDimText6 As String
'


If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"
If Cells(RowIndex, ColF).Value = "" Then Cells(RowIndex, ColF).Value = "F"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
'
txtDimText5 = Cells(RowIndex, ColE).Value
txtDimText6 = Cells(RowIndex, ColF).Value



If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top


Point1(1) = X0 + 60: Point1(2) = Y0 + 90
Point2(1) = X0 + 60: Point2(2) = Y0 + 55
Point3(1) = X0 + 95: Point3(2) = Y0 + 20
Point4(1) = X0 + 130: Point4(2) = Y0 + 55
Point5(1) = X0 + 130: Point5(2) = Y0 + 90

'
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point2(1), Point2(2), Point3(1), Point3(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point4(1), Point4(2), Point5(1), Point5(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)



'=======================================
'=======================================
' Draw Dimensions
Set DimShape1 = DrawVerDimensionsL(Point2(1), Point2(2), Point1(2), txtDimText1)
Set DimShape2 = DrawAlDimensionsL(Point2(1), Point2(2), Point3(1), Point3(2), txtDimText2)
Set DimShape3 = DrawAlDimensionsR(Point3(1), Point3(2), Point4(1), Point4(2), txtDimText3)
Set DimShape4 = DrawVerDimensions(Point4(1), Point4(2), Point5(2), txtDimText4)
'
Set DimShape5 = DrawVerDimensionsL(Point2(1) - 20, Point3(2), Point2(2), txtDimText5)
Set DimShape6 = DrawVerDimensions(Point4(1) + 20, Point3(2), Point4(2), txtDimText6)



' Set the group

ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, _
DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, DimShape5.Name, DimShape6.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub





Sub Shape_38(RowIndex)
Cells(RowIndex, ColA).RowHeight = RowHeight * 1.2
' Origin point coordinate
Dim X0 As Double
Dim Y0 As Double
' Define Shape points
Dim Point1(2) As Double
Dim Point2(2) As Double
Dim Point3(2) As Double
Dim Point4(2) As Double
Dim Point5(2) As Double


'Define the group
Dim Gr1 As Shape
Dim DimShape1 As Shape
Dim DimShape2 As Shape
Dim DimShape3 As Shape
Dim DimShape4 As Shape
'
Dim DimShape5 As Shape
Dim DimShape6 As Shape

'DIMENSIONS TEXT
Dim txtDimText1 As String
Dim txtDimText2 As String
Dim txtDimText3 As String
Dim txtDimText4 As String
'
'
Dim txtDimText5 As String
Dim txtDimText6 As String
'


If Cells(RowIndex, ColA).Value = "" Then Cells(RowIndex, ColA).Value = "A"
If Cells(RowIndex, ColB).Value = "" Then Cells(RowIndex, ColB).Value = "B"
If Cells(RowIndex, ColC).Value = "" Then Cells(RowIndex, ColC).Value = "C"
If Cells(RowIndex, ColD).Value = "" Then Cells(RowIndex, ColD).Value = "D"
If Cells(RowIndex, ColE).Value = "" Then Cells(RowIndex, ColE).Value = "E"
If Cells(RowIndex, ColF).Value = "" Then Cells(RowIndex, ColF).Value = "F"

txtDimText1 = Cells(RowIndex, ColA).Value
txtDimText2 = Cells(RowIndex, ColB).Value
txtDimText3 = Cells(RowIndex, ColC).Value
txtDimText4 = Cells(RowIndex, ColD).Value
'
txtDimText5 = Cells(RowIndex, ColE).Value
txtDimText6 = Cells(RowIndex, ColF).Value



If IsNumeric(txtDimText1) And IsNumeric(txtDimText2) And IsNumeric(txtDimText3) And IsNumeric(txtDimText4) Then
Cells(RowIndex, ColUnitLength).Value = CDbl(txtDimText1) + CDbl(txtDimText2) + CDbl(txtDimText3) + CDbl(txtDimText4)
Else
Cells(RowIndex, ColUnitLength).Value = ""
End If

X0 = Cells(RowIndex, ColShape).Left
Y0 = Cells(RowIndex, ColShape).Top


Point1(1) = X0 + 10: Point1(2) = Y0 + 55
Point2(1) = X0 + 60: Point2(2) = Y0 + 55
Point3(1) = X0 + 95: Point3(2) = Y0 + 20
Point4(1) = X0 + 130: Point4(2) = Y0 + 55
Point5(1) = X0 + 130: Point5(2) = Y0 + 90

'
' Draw line 1
Dim shLine1 As Shape
Set shLine1 = ActiveSheet.Shapes.AddLine(Point1(1), Point1(2), Point2(1), Point2(2))
shLine1.Line.Weight = ShapeWeight: shLine1.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 2
Dim shLine2 As Shape
Set shLine2 = ActiveSheet.Shapes.AddLine(Point2(1), Point2(2), Point3(1), Point3(2))
shLine2.Line.Weight = ShapeWeight: shLine2.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 3
Dim shLine3 As Shape
Set shLine3 = ActiveSheet.Shapes.AddLine(Point3(1), Point3(2), Point4(1), Point4(2))
shLine3.Line.Weight = ShapeWeight: shLine3.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)

' Draw Line 4
Dim shLine4 As Shape
Set shLine4 = ActiveSheet.Shapes.AddLine(Point4(1), Point4(2), Point5(1), Point5(2))
shLine4.Line.Weight = ShapeWeight: shLine4.Line.ForeColor.RGB = RGB(RrR, GgG, BbB)



'=======================================
'=======================================
' Draw Dimensions
'Set DimShape1 = DrawVerDimensionsL(Point2(1), Point2(2), Point1(2), txtDimText1)
Set DimShape1 = DrawHorDimensions(Point1(1), Point1(2), Point2(1), txtDimText1, 20)
Set DimShape2 = DrawAlDimensionsL(Point2(1), Point2(2), Point3(1), Point3(2), txtDimText2)
Set DimShape3 = DrawAlDimensionsR(Point3(1), Point3(2), Point4(1), Point4(2), txtDimText3)
Set DimShape4 = DrawVerDimensions(Point4(1), Point4(2), Point5(2), txtDimText4)
'
'Set DimShape5 = DrawVerDimensionsL(Point2(1) - 20, Point3(2), Point2(2), txtDimText5)
Set DimShape5 = MOD2.DrawVerDimensions5(Point2(1) - 20, Point3(2), Point2(2), txtDimText5)
Set DimShape6 = DrawVerDimensions(Point4(1) + 20, Point3(2), Point4(2), txtDimText6)



' Set the group

ActiveSheet.Shapes.Range(Array(shLine1.Name, shLine2.Name, shLine3.Name, shLine4.Name, _
DimShape1.Name, DimShape2.Name, DimShape3.Name, DimShape4.Name, DimShape5.Name, DimShape6.Name)).Select
Selection.ShapeRange.Shadow.Type = msoShadow21
Set Gr1 = Selection.ShapeRange.Group
Cells(RowIndex, ColShapeName).Value = Gr1.Name

End Sub
























'=========================================
'=========================================
'====== Support Modules ==================
'=========================================
'=========================================
Sub addsahpe()
Dim top1 As Double
Dim left1 As Double
Dim top2 As Double
Dim left2 As Double
Dim v
left2 = Cells(3, 5).Left
top2 = Cells(3, 4).RowHeight
top1 = Cells(3, 4).Top
left1 = Cells(3, 4).Left

Dim Sh1 As Shape
Dim Sh2 As Shape
Set Sh1 = ActiveSheet.Shapes.AddLine(left1, top1, left2, top1 + top2)
Sh1.Line.DashStyle = msoLineSolid
Sh1.Line.BeginArrowheadStyle = msoArrowheadOpen
Sh1.Line.EndArrowheadStyle = msoArrowheadOpen

Sh1.Line.Weight = 1

Set Sh2 = ActiveSheet.Shapes.AddLine(left1, top1 + top2, left2, top1)
Sh2.Line.DashStyle = msoLineSolid
Sh2.Line.BeginArrowheadStyle = msoArrowheadOpen
Sh2.Line.EndArrowheadStyle = msoArrowheadOpen

Sh2.Line.Weight = 4

Cells(3, 4).Select
ActiveSheet.Shapes.Range(Array(Sh1.Name, Sh2.Name)).Select
Set v = Selection.ShapeRange.Group
Cells(4, 4).Value = top1 & "; " & (top1 + top2)
End Sub

Sub DrawArrowArc1()
   Dim Arc As Shape
   Const CenterX = 200
   Const CenterY = 200
   Const Radius = 100
   'Draw 90-degree arc with 100 point radius and center at (200,200)
   Set Arc = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
   'Add arrowhead on clockwise end
'   Arc.Line.EndArrowheadStyle = msoArrowheadTriangle
'   'adjust arrow to start at -30 degrees and end at 63 degrees
'   '(measured clockwise positive from vertical)
   Arc.Adjustments.Item(1) = 0
   Arc.Adjustments.Item(2) = 90
End Sub
Sub DrawArrowArc2()
   Dim Arc As Shape
   Const CenterX = 200
   Const CenterY = 200
   Const Radius = 100
   'Draw 90-degree arc with 100 point radius and center at (200,200)
   Set Arc = ActiveSheet.Shapes.AddShape(msoShapeArc, CenterX, CenterY - Radius, Radius, Radius)
   'Add arrowhead on clockwise end
'   Arc.Line.EndArrowheadStyle = msoArrowheadTriangle
'   'adjust arrow to start at -30 degrees and end at 63 degrees
'   '(measured clockwise positive from vertical)
   Arc.Adjustments.Item(1) = 90
   Arc.Adjustments.Item(2) = 180
End Sub


