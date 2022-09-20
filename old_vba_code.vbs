

Sub LinkAttrToReinfBar()
On Error GoTo ErrorCenter

    Dim returnObj As AcadObject
    Dim basePnt As Variant
    Dim MarkBlock As AcadBlockReference
    Dim BarBlock As AcadBlockReference
    Dim objID As String
    Dim varAttributes As Variant
    
    Dim Field As String
    Dim aFields(10) As String
    
    Dim blkName As String
   
    
    ThisDrawing.Utility.GetEntity returnObj, basePnt, vbLf & "Select an Attribute Block (MARK):"
    Set MarkBlock = returnObj
    
    ThisDrawing.Utility.GetEntity returnObj, basePnt, vbLf & "Select a Rebar Block:"
    Set BarBlock = returnObj
    
    ThisDrawing.StartUndoMark
    
    If MarkBlock.Name = "MARK" Then
        objID = "_ObjId " & BarBlock.ObjectID
        blkName = BarBlock.EffectiveName
        
        If (InStr(1, blkName, "XX", vbTextCompare) > 0) And BarBlock.IsDynamicBlock Then
            blkName = Split(blkName, "XX", , vbTextCompare)(0)
        End If
        
        Select Case blkName
            Case "S-BAR", "S-BAR-COUPLER-32&40", "S-BAR-COUPLER-32"
                Field = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(23).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                
                DrawSBar CDbl(basePnt(0)), CDbl(basePnt(1)) - 600
                AddText Field, 160, CDbl(basePnt(0)), CDbl(basePnt(1)) - 500
                
            Case "BAR-1"
                aFields(0) = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(17).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                aFields(1) = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(9).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                
                Field = "%<\AcExpr ( " & aFields(0) & "+" & aFields(1) & ") \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                
                
                DrawLBar CDbl(basePnt(0)), CDbl(basePnt(1)) - 600
                
                AddText aFields(0), 160, CDbl(basePnt(0)), CDbl(basePnt(1)) - 500
                AddText aFields(1), 160, CDbl(basePnt(0)) + 1150, CDbl(basePnt(1)) - 600, 90
                
            Case "U-BAR"
                
                aFields(0) = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(1).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                aFields(1) = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(15).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                aFields(2) = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(8).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                
                Field = "%<\AcExpr ( " & aFields(0) & "+" & aFields(1) & "+" & aFields(2) & ") \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                
                
                DrawUBar1 CDbl(basePnt(0)), CDbl(basePnt(1)) - 250
                
                AddText aFields(0), 160, CDbl(basePnt(0)) + 150, CDbl(basePnt(1)) - 500
                AddText aFields(1), 160, CDbl(basePnt(0)) + 1150, CDbl(basePnt(1)) - 600, 90
                AddText aFields(2), 160, CDbl(basePnt(0)) - 100, CDbl(basePnt(1)) - 600, 90
                
            Case "U-BAR-1"
                
                
                aFields(0) = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(9).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                aFields(1) = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(16).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                aFields(2) = "%<\AcObjProp Object(%<\" & objID & ">%).Parameter(2).UpdatedDistance \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                
                Field = "%<\AcExpr ( " & aFields(0) & "+" & aFields(1) & "+" & aFields(2) & ") \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
                
                
                DrawUBar2 CDbl(basePnt(0)), CDbl(basePnt(1)) - 500
                
                AddText aFields(0), 160, CDbl(basePnt(0)) - 1000, CDbl(basePnt(1)) - 900, 90
                AddText aFields(1), 160, CDbl(basePnt(0)) - 700, CDbl(basePnt(1)) - 830
                AddText aFields(2), 160, CDbl(basePnt(0)) - 700, CDbl(basePnt(1)) - 370
                
                
            'Case "Z-BAR-INCLIN.", "Z-BAR"
                
            'Case "Z-BAR-2-SIDE"
                
            'Case "Z-BAR-INC.HOOK"
                
        Case Else
            ThisDrawing.Utility.Prompt vbCr & "Error!: Select a Correct Block for REBAR."
        End Select
        
        If Len(Field) > 1 Then
            varAttributes = MarkBlock.GetAttributes
            varAttributes(2).textString = Split(varAttributes(2).textString, "X", , vbTextCompare)(0) & "X" & Field
            
        End If
    Else
        ThisDrawing.Utility.Prompt vbCr & "Error!: Select an Attribute Block First (Name = MARK)"
    End If
    
    
    
'    'Update Bar Count
'    Dim dimObj As AcadDimension
'    Dim currBarCount As String
'    Dim UpdatedCount As String
'
'    ThisDrawing.Utility.GetEntity returnObj, basePnt, vbLf & "Select a Dimension to update Bar Count (MARK):"
'    Set dimObj = returnObj
'
'    If dimObj.StyleName = "D-DIM-MILL-MILL-0100-DIS." Then
'
'        objID = "_ObjId " & dimObj.ObjectID
'
'        aFields(0) = "%<\AcObjProp Object(%<\" & objID & ">%).Measurement \f " & Chr(34) & "%lu2" & Chr(34) & ">%"
'        aFields(1) = "%<\AcObjProp Object(%<\" & objID & ">%).AltUnitsScale \f " & Chr(34) & "%lu2" & Chr(34) & ">%"
'
'        Field = "%<\AcExpr ( " & aFields(0) & "*" & aFields(1) & "+1) \f " & Chr(34) & "%lu2%pr0" & Chr(34) & ">%"
'
'        currBarCount = Split(varAttributes(2).textString, "-T", , vbTextCompare)(0)
'
'        If Len(currBarCount) > 2 Then
'            varAttributes(2).textString = Replace(varAttributes(2).textString, currBarCount & "-T", Field & "-T", , , vbTextCompare)
'        End If
'     End If
    
ExitSub:

    ThisDrawing.EndUndoMark
    
    Exit Sub
    
ErrorCenter:
    MsgBox Err.Description, vbCritical, "ERROR!!"
    Resume ExitSub
End Sub