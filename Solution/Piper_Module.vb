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

        If partDoc.SelectSet.Count <> 1 Then
            MsgBox("Select one point to begin")
            Exit Sub
        End If

        Dim TPoint As Object = partDoc.SelectSet(1)
        Dim TLine As Object = Nothing

        If TPoint.Type = ObjectTypeEnum.kSketchPointObject Then
            Console.WriteLine("2D Sketchpoint found")

            get2DObjects(TPoint, TLine)
            createSweep2D(partDoc, TPoint, TLine)

        ElseIf TPoint.Type = ObjectTypeEnum.kSketchPoint3DObject Then
            Console.WriteLine("3D Sketchpoint found")

            get3DObjects(TPoint, TLine)
            createSweep3D(partDoc, TPoint, TLine)

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
        Dim cons2 As GeometricConstraint
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
        cons2 = SK.GeometricConstraints.AddCoincident(cp, C.CenterSketchPoint)

        'Get path to sweep along
        Dim CurvColl As ObjectCollection = InvApp.TransientObjects.CreateObjectCollection
        For Each e As SketchEntity3D In parentSketch.SketchEntities3D
            If e.Type = ObjectTypeEnum.kSketchLine3DObject Or e.Type = ObjectTypeEnum.kSketchArc3DObject Then
                CurvColl.Add(e)
            End If
        Next

        Dim path As Path = dPart.ComponentDefinition.Features.CreateSpecifiedPath(CurvColl)
        Dim PlanSK As PlanarSketch = SK
        Dim profile As Profile = PlanSK.Profiles.AddForSolid

        Dim SW As SweepDefinition
        SW = dPart.ComponentDefinition.Features.SweepFeatures.CreateSweepDefinition(SweepTypeEnum.kPathSweepType, profile, path, PartFeatureOperationEnum.kJoinOperation)
        dPart.ComponentDefinition.Features.SweepFeatures.Add(SW)

    End Sub

    Private Sub createSweep2D(dPart As PartDocument, targetPoint As Object, targetLine As Object)

        Dim WP As WorkPlane
        Dim SK As Sketch
        Dim cp As SketchPoint
        Dim C As SketchCircle
        Dim cons1 As DiameterDimConstraint
        Dim cons2 As GeometricConstraint
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
        cons2 = SK.GeometricConstraints.AddCoincident(cp, C.CenterSketchPoint)

        'Get path to sweep along
        Dim CurvColl As ObjectCollection = InvApp.TransientObjects.CreateObjectCollection
        For Each e As SketchEntity In parentSketch.SketchEntities
            If e.Type = ObjectTypeEnum.kSketchLineObject Or e.Type = ObjectTypeEnum.kSketchArcObject Then
                CurvColl.Add(e)
            End If
        Next

        Dim path As Path = dPart.ComponentDefinition.Features.CreateSpecifiedPath(CurvColl)
        Dim PlanSK As PlanarSketch = SK
        Dim profile As Profile = PlanSK.Profiles.AddForSolid

        Dim SW As SweepDefinition
        SW = dPart.ComponentDefinition.Features.SweepFeatures.CreateSweepDefinition(SweepTypeEnum.kPathSweepType, profile, path, PartFeatureOperationEnum.kJoinOperation)
        dPart.ComponentDefinition.Features.SweepFeatures.Add(SW)

    End Sub


End Module
