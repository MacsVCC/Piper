Imports Inventor
Imports System.Runtime.InteropServices

Module Piper_Module


    Public Sub Execute()

        Console.WriteLine("Button was clicked.")

        Dim partDoc As Inventor.PartDocument
        If InvApp.ActiveDocumentType = DocumentTypeEnum.kPartDocumentObject Then
            partDoc = InvApp.ActiveDocument
        Else
            MsgBox("Current document is not a Part document")
            Exit Sub
        End If

        Console.WriteLine(partDoc.SelectSet.Count & " selected object(s)")


        If partDoc.SelectSet.Count = 1 Then
            'Continue
        ElseIf partDoc.SelectSet.Count = 0 Then
            MsgBox("No Sketchpoints selected")
            Exit Sub
        Else

            For Each occ In partDoc.SelectSet
                If partDoc.SelectSet(1).Type = occ.Type Then
                    'Continue
                Else
                    'Mismatching selections found!
                End If
            Next

        End If

        Dim StartPoints As New ArrayList

        If partDoc.SelectSet.Count = 1 Then
            If partDoc.SelectSet(1).Type = ObjectTypeEnum.kSketchPointObject Or partDoc.SelectSet(1).Type = ObjectTypeEnum.kSketchPoint3DObject Then
                StartPoints.Add(partDoc.SelectSet(1))
            Else
                MsgBox("Sketchpoint not found in selection.")
                Exit Sub
            End If
        Else
            If MsgBox("Multiple start points selected. Perform batch sweep?", vbYesNo) = MsgBoxResult.Yes Then

                For Each occ In partDoc.SelectSet
                    If occ.Type <> ObjectTypeEnum.kSketchPointObject And occ.Type <> ObjectTypeEnum.kSketchPoint3DObject Then
                        MsgBox("Not all selected objects are " & [Enum].GetName(GetType(ObjectTypeEnum), occ.Type))
                        Exit Sub
                    Else
                        StartPoints.Add(occ)
                    End If
                Next

            Else
                Exit Sub
            End If

        End If


        For Each occ In StartPoints
            Dim Line As Object
            Dim Point As Object = occ

            If Point.Type = ObjectTypeEnum.kSketchPointObject Then
                Line = getAttachedLine2D(Point)
                If IsNothing(Line) = False Then
                    createSweep2D(partDoc, Point, Line)
                End If
            ElseIf Point.Type = ObjectTypeEnum.kSketchPoint3dObject Then
                Line = getAttachedLine3D(Point)
                If IsNothing(Line) = False Then
                    createSweep3D(partDoc, Point, Line)
                End If
            Else
                    MsgBox("Somehow a non-point has sneaked throught to the createSweep stage...")
            End If
        Next

    End Sub



    Private Function getAttachedLine2D(ByRef targetPoint As Object) As SketchLine
        Dim AttachedLines As New ArrayList

        For Each L In targetPoint.AttachedEntities
            If L.Type = ObjectTypeEnum.kSketchLineObject Then
                AttachedLines.Add(L)
            End If
        Next

        If AttachedLines.Count = 0 Then
            MsgBox("No attached lines found.")
            Return Nothing
        End If

        Return AttachedLines(0)
    End Function

    Private Function getAttachedLine3D(ByRef targetPoint As Object) As SketchLine3D

        Dim AttachedLines3D As New ArrayList
        For Each L In targetPoint.AttachedEntities
            If L.Type = ObjectTypeEnum.kSketchLine3DObject Then
                AttachedLines3D.Add(L)
            End If
        Next
        If AttachedLines3D.Count = 0 Then
            MsgBox("No attached lines found.")
        End If

        Return AttachedLines3D(0)

    End Function

    Private Sub createSweep3D(dPart As PartDocument, targetPoint As Object, targetLine As Object)

        Dim WP As WorkPlane
        Dim SK As Sketch
        Dim cp As SketchPoint
        Dim C As SketchCircle
        Dim cons1 As DiameterDimConstraint
        Dim Crad As Double
        Dim parentSketch As Sketch3D = targetLine.Parent


        'Get intended circle diameter
        Dim allParams As Parameters
        allParams = dPart.ComponentDefinition.Parameters
        For Each p As Parameter In allParams
            If p.Name = "D" Then
                Crad = p.Value / 2
            End If
        Next
        If Crad = 0 Then
            MsgBox("No diameter value found. Exiting.")
            Exit Sub
        End If


        WP = dPart.ComponentDefinition.WorkPlanes.AddByNormalToCurve(targetLine, targetPoint)
        WP.Visible = False
        SK = dPart.ComponentDefinition.Sketches.Add(WP)
        cp = SK.AddByProjectingEntity(targetPoint)
        C = SK.SketchCircles.AddByCenterRadius(cp, Crad)
        cons1 = SK.DimensionConstraints.AddDiameter(C, cp.Geometry)
        C.CenterSketchPoint.Merge(cp)


        Dim path As Path = get3DPath(dPart, targetPoint)
        Dim PlanSK As PlanarSketch = SK
        Dim profile As Profile = PlanSK.Profiles.AddForSolid

        Dim SW As SweepDefinition
        SW = dPart.ComponentDefinition.Features.SweepFeatures.CreateSweepDefinition(SweepTypeEnum.kPathSweepType, profile, path, PartFeatureOperationEnum.kJoinOperation)
        dPart.ComponentDefinition.Features.SweepFeatures.Add(SW)

    End Sub
    Function get3DPath(ByRef dPart As PartDocument, targetPoint As SketchPoint3D) As Path

        'Dim CurveArray As ArrayList = getAttachedLines(0, targetPoint, Nothing, False)
        Dim CurveArray As New ArrayList

        For Each Sc In targetPoint.AttachedEntities
            If Sc.Type = ObjectTypeEnum.kSketchLine3DObject Then
                CurveArray.Add(Sc)
            End If
        Next

        If CurveArray.Count > 0 Then
            Return dPart.ComponentDefinition.Features.CreatePath(CurveArray(0))
        Else
            MsgBox("No lines attached to selected point")
            Return Nothing
        End If

    End Function

    Function getAttachedLines(recursiveCounter As Integer, beginningPoint As SketchPoint3D, Optional prevLine As SketchLine3D = Nothing, Optional Verbose As Boolean = True) As ArrayList

        'Check for recursion limit
        If recursiveCounter = 10 Or recursiveCounter >= 10 Then
            MsgBox("Recursion limit reached!")
            Return New ArrayList
        End If

        Dim Coll As New ArrayList

        'Count attached lines
        Dim counter% = 0
        For Each Sl As SketchLine3D In beginningPoint.AttachedEntities
            counter += 1
            If Verbose Then
                MsgBox("Recursion level: " & (recursiveCounter + 1) & vbNewLine & "Line: " & (counter + 1) & vbNewLine & Sl.Length)
            End If
        Next

        If Verbose Then
            MsgBox(counter & " attached lines found")
        End If

        'Check if end of line is reached
        If IsNothing(prevLine) = False Then
            If counter = 1 & beginningPoint.AttachedEntities(1).AssociativeID = prevLine.AssociativeID Then
                'END OF LINE
                'return empty arraylist
                Return Coll
            End If
        Else
            'Continue
        End If


        For Each Sl As SketchLine3D In beginningPoint.AttachedEntities

            If IsNothing(prevLine) = False Then
                If Sl.AssociativeID = prevLine.AssociativeID Then
                    'Ignore line behind pilot
                    Continue For
                End If
            End If

            'add next line to coll
            Coll.Add(Sl)
            If Verbose Then
                MsgBox("Added " & Sl.Length & " to the collection")
            End If
            'get end ppoint of line
            Dim n As SketchPoint3D
            If Sl.StartSketchPoint.AssociativeID = beginningPoint.AssociativeID Then
                n = Sl.EndSketchPoint
            Else
                n = Sl.StartSketchPoint
            End If

            If Verbose Then
                MsgBox("Following: " & Sl.Length)
            End If
            Coll.AddRange(getAttachedLines(recursiveCounter + 1, n, Sl, Verbose))
            If Verbose Then
                MsgBox("Returned to: " & Sl.Length)
            End If

        Next

        Return Coll

    End Function

    Private Sub createSweep2D(dPart As PartDocument, targetPoint As Object, targetLine As Object)

        Dim WP As WorkPlane
        Dim SK As Sketch
        Dim cp As SketchPoint
        Dim C As SketchCircle
        Dim cons1 As DiameterDimConstraint
        Dim Crad As Double
        Dim parentSketch As Sketch = targetLine.Parent


        'Get intended circle diameter
        Dim allParams As Parameters
        allParams = dPart.ComponentDefinition.Parameters
        For Each p As Parameter In allParams
            If p.Name = "D" Then
                Crad = p.Value / 2
            End If
        Next
        If Crad = 0 Then
            MsgBox("No diameter value found. Exiting.")
            Exit Sub
        End If

        WP = dPart.ComponentDefinition.WorkPlanes.AddByNormalToCurve(targetLine, targetPoint)
        WP.Visible = False
        SK = dPart.ComponentDefinition.Sketches.Add(WP)
        cp = SK.AddByProjectingEntity(targetPoint)
        C = SK.SketchCircles.AddByCenterRadius(cp, Crad)
        cons1 = SK.DimensionConstraints.AddDiameter(C, cp.Geometry)
        C.CenterSketchPoint.Merge(cp)

        Dim path As Path = get2DPath(dPart, targetPoint)
        Dim PlanSK As PlanarSketch = SK
        Dim profile As Profile = PlanSK.Profiles.AddForSolid

        Dim SW As SweepDefinition
        SW = dPart.ComponentDefinition.Features.SweepFeatures.CreateSweepDefinition(SweepTypeEnum.kPathSweepType, profile, path, PartFeatureOperationEnum.kJoinOperation)
        dPart.ComponentDefinition.Features.SweepFeatures.Add(SW)

    End Sub

    Function get2DPath(ByRef dPart As PartDocument, targetPoint As SketchPoint) As Path

        Dim CurveArray As New ArrayList

        For Each Sc In targetPoint.AttachedEntities
            If Sc.Type = ObjectTypeEnum.kSketchLineObject Then
                CurveArray.Add(Sc)
            End If
        Next

        If CurveArray.Count > 0 Then
            Return dPart.ComponentDefinition.Features.CreatePath(CurveArray(0))
        Else
            MsgBox("No lines attached to selected point")
            Return Nothing
        End If

    End Function

End Module
