Imports Inventor
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Namespace Piper
    <ProgIdAttribute("Piper.StandardAddInServer"), _
    GuidAttribute("56761c4d-8dad-41ef-ad69-b59cd902ac1c")> _
    Public Class StandardAddInServer
        Implements Inventor.ApplicationAddInServer

        Private WithEvents m_uiEvents As UserInterfaceEvents
        Private WithEvents m_sampleButton As ButtonDefinition

#Region "ApplicationAddInServer Members"

        ' This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
        ' to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
        ' the first time. However, with the introduction of the ribbon this argument is always true.
        Public Sub Activate(ByVal addInSiteObject As Inventor.ApplicationAddInSite, ByVal firstTime As Boolean) Implements Inventor.ApplicationAddInServer.Activate
            ' Initialize AddIn members.
            g_inventorApplication = addInSiteObject.Application

            ' Connect to the user-interface events to handle a ribbon reset.
            m_uiEvents = g_inventorApplication.UserInterfaceManager.UserInterfaceEvents

            ' TODO: Add button definitions.

            ' Sample to illustrate creating a button definition.
            Dim largeIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.PipeIcon)
            'Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.YourSmallImage)
            Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions
            m_sampleButton = controlDefs.AddButtonDefinition("Auto Sweep", "Act_ivate", CommandTypesEnum.kShapeEditCmdType, AddInClientID, , , largeIcon, largeIcon)

            ' Add to the user interface, if it's the first time.
            If firstTime Then
                AddToUserInterface()
            End If
        End Sub

        ' This method is called by Inventor when the AddIn is unloaded. The AddIn will be
        ' unloaded either manually by the user or when the Inventor session is terminated.
        Public Sub Deactivate() Implements Inventor.ApplicationAddInServer.Deactivate

            ' TODO:  Add ApplicationAddInServer.Deactivate implementation

            ' Release objects.
            m_uiEvents = Nothing
            g_inventorApplication = Nothing

            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Sub

        ' This property is provided to allow the AddIn to expose an API of its own to other 
        ' programs. Typically, this  would be done by implementing the AddIn's API
        ' interface in a class and returning that class object through this property.
        Public ReadOnly Property Automation() As Object Implements Inventor.ApplicationAddInServer.Automation
            Get
                Return Nothing
            End Get
        End Property

        ' Note:this method is now obsolete, you should use the 
        ' ControlDefinition functionality for implementing commands.
        Public Sub ExecuteCommand(ByVal commandID As Integer) Implements Inventor.ApplicationAddInServer.ExecuteCommand
        End Sub

#End Region

#Region "User interface definition"
        ' Sub where the user-interface creation is done.  This is called when
        ' the add-in loaded and also if the user interface is reset.
        Private Sub AddToUserInterface()
            ' This is where you'll add code to add buttons to the ribbon.

            '** Sample to illustrate creating a button on a new panel of the Tools tab of the Part ribbon.

            '' Get the part ribbon.
            Dim partRibbon As Ribbon = g_inventorApplication.UserInterfaceManager.Ribbons.Item("Part")

            '' Get the "Tools" tab.
            Dim toolsTab As RibbonTab = partRibbon.RibbonTabs.Item("id_TabTools")

            '' Create a new panel.
            Dim customPanel As RibbonPanel = toolsTab.RibbonPanels.Add("MH Tools", "MHCustom", AddInClientID)

            '' Add a button.
            customPanel.CommandControls.AddButton(m_sampleButton, True)
        End Sub

        Private Sub m_uiEvents_OnResetRibbonInterface(Context As NameValueMap) Handles m_uiEvents.OnResetRibbonInterface
            ' The ribbon was reset, so add back the add-ins user-interface.
            AddToUserInterface()
        End Sub

        ' Sample handler for the button.
        Private Sub m_sampleButton_OnExecute(Context As NameValueMap) Handles m_sampleButton.OnExecute
            Console.WriteLine("Button was clicked.")

            Dim partDoc As Inventor.PartDocument
            If g_inventorApplication.ActiveDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                partDoc = g_inventorApplication.ActiveDocument
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
            Dim CurvColl As ObjectCollection = g_inventorApplication.TransientObjects.CreateObjectCollection
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
            Dim CurvColl As ObjectCollection = g_inventorApplication.TransientObjects.CreateObjectCollection
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
#End Region

    End Class
End Namespace


Public Module Globals
    ' Inventor application object.
    Public g_inventorApplication As Inventor.Application

#Region "Function to get the add-in client ID."
    ' This function uses reflection to get the GuidAttribute associated with the add-in.
    Public Function AddInClientID() As String
        Dim guid As String = ""
        Try
            Dim t As Type = GetType(Piper.StandardAddInServer)
            Dim customAttributes() As Object = t.GetCustomAttributes(GetType(GuidAttribute), False)
            Dim guidAttribute As GuidAttribute = CType(customAttributes(0), GuidAttribute)
            guid = "{" + guidAttribute.Value.ToString() + "}"
        Catch
        End Try

        Return guid
    End Function
#End Region

#Region "hWnd Wrapper Class"
    ' This class is used to wrap a Win32 hWnd as a .Net IWind32Window class.
    ' This is primarily used for parenting a dialog to the Inventor window.
    '
    ' For example:
    ' myForm.Show(New WindowWrapper(g_inventorApplication.MainFrameHWND))
    '
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As IntPtr _
          Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

        Private _hwnd As IntPtr
    End Class
#End Region

#Region "Image Converter"
    ' Class used to convert bitmaps and icons from their .Net native types into
    ' an IPictureDisp object which is what the Inventor API requires. A typical
    ' usage is shown below where MyIcon is a bitmap or icon that's available
    ' as a resource of the project.
    '
    ' Dim smallIcon As stdole.IPictureDisp = PictureDispConverter.ToIPictureDisp(My.Resources.MyIcon)

    Public NotInheritable Class PictureDispConverter
        <DllImport("OleAut32.dll", EntryPoint:="OleCreatePictureIndirect", ExactSpelling:=True, PreserveSig:=False)> _
        Private Shared Function OleCreatePictureIndirect( _
            <MarshalAs(UnmanagedType.AsAny)> ByVal picdesc As Object, _
            ByRef iid As Guid, _
            <MarshalAs(UnmanagedType.Bool)> ByVal fOwn As Boolean) As stdole.IPictureDisp
        End Function

        Shared iPictureDispGuid As Guid = GetType(stdole.IPictureDisp).GUID

        Private NotInheritable Class PICTDESC
            Private Sub New()
            End Sub

            'Picture Types
            Public Const PICTYPE_BITMAP As Short = 1
            Public Const PICTYPE_ICON As Short = 3

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Icon
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Icon))
                Friend picType As Integer = PICTDESC.PICTYPE_ICON
                Friend hicon As IntPtr = IntPtr.Zero
                Friend unused1 As Integer
                Friend unused2 As Integer

                Friend Sub New(ByVal icon As System.Drawing.Icon)
                    Me.hicon = icon.ToBitmap().GetHicon()
                End Sub
            End Class

            <StructLayout(LayoutKind.Sequential)> _
            Public Class Bitmap
                Friend cbSizeOfStruct As Integer = Marshal.SizeOf(GetType(PICTDESC.Bitmap))
                Friend picType As Integer = PICTDESC.PICTYPE_BITMAP
                Friend hbitmap As IntPtr = IntPtr.Zero
                Friend hpal As IntPtr = IntPtr.Zero
                Friend unused As Integer

                Friend Sub New(ByVal bitmap As System.Drawing.Bitmap)
                    Me.hbitmap = bitmap.GetHbitmap()
                End Sub
            End Class
        End Class

        Public Shared Function ToIPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
            Dim pictIcon As New PICTDESC.Icon(icon)
            Return OleCreatePictureIndirect(pictIcon, iPictureDispGuid, True)
        End Function

        Public Shared Function ToIPictureDisp(ByVal bmp As System.Drawing.Bitmap) As stdole.IPictureDisp
            Dim pictBmp As New PICTDESC.Bitmap(bmp)
            Return OleCreatePictureIndirect(pictBmp, iPictureDispGuid, True)
        End Function
    End Class
#End Region

End Module
