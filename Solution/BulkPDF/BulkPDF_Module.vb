Imports Inventor
Imports System.Runtime.InteropServices

Module BulkPDF_Module
    Public Sub Execute(Context As NameValueMap)

        Dim oPDF As TranslatorAddIn
        oPDF = InvApp.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}")

        Dim Drgs As New ArrayList
        Dim msg As String = ""
        For Each i As Document In InvApp.Documents
            If i.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                Drgs.Add(i)
                msg = msg & vbNewLine & i.DisplayName
            End If
        Next


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' TESTING
        Dim F As New BulkPDF_Form
        F.DocumentList = Drgs
        F.ShowDialog()
        Exit Sub
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' TESTING


        If MsgBox("Drawings found: " & msg & vbNewLine & "Continue?", vbYesNo) = MsgBoxResult.No Then
            Exit Sub
        End If

        Dim Overwrite As Boolean = False
        If MsgBox("Overwrite all present drawings?", vbYesNo) = MsgBoxResult.Yes Then
            Overwrite = True
        End If

        For Each D As DrawingDocument In Drgs
            Export(D, F.OverwriteAll)
        Next

    End Sub

    Private Sub Export(Doc As DrawingDocument, overwriteAll As Boolean, Optional PDFName As String = "")

        If Doc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            Dim drawDoc As Inventor.DrawingDocument = Doc
            'warn user if errors (show error list, in log and on screen), choose to continue to plot or skip file and quit this routine
            Dim errorMng As ErrorManager = InvApp.ErrorManager
            If errorMng.HasErrors = True Then
                Dim errors As String = errorMng.AllMessages
                MsgBox("This drawing has errors: " & drawDoc.FullFileName & vbCr & errors, MsgBoxStyle.SystemModal)
            End If

            'export file to PDF using 'all sheets', 'default sizes', 'file name - Rev name' (keep color)
            Dim PDFAddin As TranslatorAddIn
            PDFAddin = InvApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
            Dim tContext As TranslationContext
            tContext = InvApp.TransientObjects.CreateTranslationContext
            tContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
            Dim nvmOptions As NameValueMap
            nvmOptions = InvApp.TransientObjects.CreateNameValueMap
            Dim dm As DataMedium
            dm = InvApp.TransientObjects.CreateDataMedium

            If PDFName = "" Then
                dm.FileName = Doc.FullFileName.Replace(".idw", "") & " R" & Doc.Sheets(1).Revision & ".pdf"
            Else
                dm.FileName = PDFName
            End If


            If overwriteAll = False Then
                If My.Computer.FileSystem.FileExists(dm.FileName) Then
                    Dim Filename = My.Computer.FileSystem.GetFileInfo(dm.FileName).Name
                    If (MsgBox("File " & Filename & " exists. Overwrite?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
                        My.Computer.FileSystem.DeleteFile(dm.FileName)
                    Else
                        Exit Sub
                    End If
                End If
            Else
                If My.Computer.FileSystem.FileExists(dm.FileName) Then
                    My.Computer.FileSystem.DeleteFile(dm.FileName)
                End If
            End If

            If PDFAddin.HasSaveCopyAsOptions(drawDoc, tContext, nvmOptions) Then
                With nvmOptions
                    .Value("All_Color_AS_Black") = 0
                    .Value("Remove_Line_Weights") = 0
                    .Value("Vector_Resolution") = 400
                    .Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
                End With

                Try
                    PDFAddin.SaveCopyAs(drawDoc, tContext, nvmOptions, dm)
                Catch ex As Exception
                    MsgBox("Error removing or writing file; perhaps existing file can't be overwritten.  Skipping file.")
                End Try
            End If

        End If

    End Sub

End Module
