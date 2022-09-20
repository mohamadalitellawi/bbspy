Attribute VB_Name = "MOD2"
Function DrawVerDimensions2(StartX As Double, StartY As Double, EndY As Double, DimText As String, Optional Extension = 12) As Shape
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

P3(1) = StartX - 8
P3(2) = EndY
P4(1) = StartX + Extension
P4(2) = EndY

'Set DimLine1 = ActiveSheet.Shapes.AddLine(P1(1), P1(2), P2(1), P2(2))
Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
Set DimLine3 = ActiveSheet.Shapes.AddLine(P2(1) - 5, P2(2), P4(1) - 5, P4(2))
DimLine3.Line.BeginArrowheadStyle = msoArrowheadTriangle
DimLine3.Line.EndArrowheadStyle = msoArrowheadTriangle
'
'draw dimension text 1
txtP1(1) = StartX + (Extension + 3)
txtP1(2) = StartY + 9
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
ActiveSheet.Shapes.Range(Array(DimLine2.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawVerDimensions2 = Selection.ShapeRange.Group

End Function
Function DrawVerDimensions3(StartX As Double, StartY As Double, EndY As Double, DimText As String, Optional Extension = 28) As Shape
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
'Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
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
ActiveSheet.Shapes.Range(Array(DimLine1.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawVerDimensions3 = Selection.ShapeRange.Group

End Function

Function DrawVerDimensions4(StartX As Double, StartY As Double, EndY As Double, DimText As String, Optional Extension = 22) As Shape
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

'Set DimLine1 = ActiveSheet.Shapes.AddLine(P1(1), P1(2), P2(1), P2(2))
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
ActiveSheet.Shapes.Range(Array(DimLine2.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawVerDimensions4 = Selection.ShapeRange.Group

End Function


Function DrawVerDimensions5(StartX As Double, StartY As Double, EndY As Double, DimText As String, Optional Extension = 22) As Shape
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
'Set DimLine2 = ActiveSheet.Shapes.AddLine(P3(1), P3(2), P4(1), P4(2))
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
ActiveSheet.Shapes.Range(Array(DimLine1.Name, DimLine3.Name, DimText1.Name)).Select
Set DrawVerDimensions5 = Selection.ShapeRange.Group

End Function

