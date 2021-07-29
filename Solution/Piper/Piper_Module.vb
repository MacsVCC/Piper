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
        Else

            For Each occ In partDoc.SelectSet
                If partDoc.SelectSet(1).Type = occ.Type Then
                    'Continue
                Else
                    'Mismatching selections found!
                End If
            Next

        End If

        Dim TPoint As Object = partDoc.SelectSet(1)

        If TPoint.Type = ObjectTypeEnum.kSketchPointObject Then
            Console.WriteLine("2D Sketchpoint found")

            Dim StartPoints As New ArrayList

            If partDoc.SelectSet.Count = 1 Then
                StartPoints.Add(partDoc.SelectSet(1))
            Else
                If MsgBox("Multiple start points selected. Perform batch sweep?", vbYesNo) = MsgBoxResult.Yes Then

                    For Each occ In partDoc.SelectSet
                        If occ.Type <> ObjectTypeEnum.kSketchPointObject Then
                            MsgBox("Not all selected objects are " & [Enum].GetName(GetType(ObjectTypeEnum), TPoint.Type))
                            Exit Sub
                        Else
                            StartPoints.Add(occ)
                        End If
                    Next

                Else
                    Exit Sub
                End If

            End If

            For Each occ As SketchPoint In StartPoints
                Dim Line As Object = Nothing
                Dim Point As Object = occ
                get2DObjects(Point, Line)
                createSweep2D(partDoc, TPoint, Line)
            Next


        ElseIf TPoint.Type = ObjectTypeEnum.kSketchPoint3DObject Then
            Console.WriteLine("3D Sketchpoint found")

            Dim StartPoints As New ArrayList

            If partDoc.SelectSet.Count = 1 Then
                StartPoints.Add(partDoc.SelectSet(1))
            Else
                If MsgBox("Multiple start points selected. Continue?", vbYesNo) = MsgBoxResult.Yes Then

                    For Each occ In partDoc.SelectSet
                        If occ.Type <> ObjectTypeEnum.kSketchPoint3DObject Then
                            MsgBox("Not all selected objects are " & [Enum].GetName(GetType(ObjectTypeEnum), TPoint.Type))
                            Exit Sub
                        Else
                            StartPoints.Add(occ)
                        End If
                    Next

                End If

            End If

            For Each occ As SketchPoint3D In StartPoints
                Dim Line As Object = Nothing
                Dim Point As Object = occ
                get3DObjects(Point, Line)
                createSweep3D(partDoc, Point, Line)
            Next

        Else
            MsgBox("Unhandled object selected:" & vbNewLine & [Enum].GetName(GetType(ObjectTypeEnum), TPoint.Type) & vbNewLine & TPoint.Type)
        End If

    End Sub



    Private Sub get2DObjects(ByRef targetPoint As Object, ByRef targetLine As Object)
        If targetPoint.AttachedEntities.Count = 1 Then
            If targetPoint.AttachedEntities(1).Type = ObjectTypeEnum.kSketchLineObject Or targetPoint.AttachedEntities(1).Type = ObjectTypeEnum.kSketchArcObject Then
                targetLine = targetPoint.AttachedEntities(1)
                Console.WriteLine("2D Sketchline found")
            Else
                MsgBox("Attached object not recognised")
                Exit Sub
            End If
        Else
            Console.WriteLine("More than one attached object found. Exiting.")
            Exit Sub
        End If
    End Sub

    Private Sub get3DObjects(ByRef targetPoint As Object, ByRef targetLine As Object)
        If targetPoint.AttachedEntities.Count = 1 Then
            If targetPoint.AttachedEntities(1).Type = ObjectTypeEnum.kSketchLine3DObject Or targetPoint.AttachedEntities(1).Type = ObjectTypeEnum.kSketchArc3DObject Then
                targetLine = targetPoint.AttachedEntities(1)
                Console.WriteLine("3D Sketchline found")
            Else
                MsgBox("Attached object not recognised")
                Exit Sub
            End If
        Else
            Console.WriteLine("More than one attached object found. Exiting.")
            Exit Sub
        End If
    End Sub

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

        For Each Sc As SketchLine3D In targetPoint.AttachedEntities
            CurveArray.Add(Sc)
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

        'Dim CurveArray As ArrayList = getAttachedLines(0, targetPoint, Nothing, False)
        Dim CurveArray As New ArrayList

        For Each Sc As SketchLine In targetPoint.AttachedEntities
            CurveArray.Add(Sc)
        Next

        If CurveArray.Count > 0 Then
            Return dPart.ComponentDefinition.Features.CreatePath(CurveArray(0))
        Else
            MsgBox("No lines attached to selected point")
            Return Nothing
        End If

    End Function

End Module
