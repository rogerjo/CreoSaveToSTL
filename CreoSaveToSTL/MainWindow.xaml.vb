Imports pfcls

Class MainWindow

    Dim asyncConnection As IpfcAsyncConnection = Nothing
    Dim model As IpfcModel
    Dim activeserver As IpfcServer
    Dim paramval As IpfcParamValue
    Dim session As IpfcBaseSession
    Dim Moditem As CMpfcModelItem
    Dim State As String = ""
    Dim FileEnd As String = ""
    Dim ConvertType As Boolean
    Dim FileNameComplete As String

    Sub Creo_Connect()

        Dim asyncConnection As IpfcAsyncConnection = Nothing

        Try
            myInfo.Text = "Connecting..."

            asyncConnection = (New CCpfcAsyncConnection).Connect(Nothing, Nothing, Nothing, Nothing)
            session = asyncConnection.Session
            activeserver = session.GetActiveServer
            model = session.CurrentModel
            myInfo.Text = "Connection established"

        Catch ex As Exception
            MsgBox(ex.Message.ToString + Chr(13) + ex.StackTrace.ToString)
            If Not asyncConnection Is Nothing AndAlso asyncConnection.IsRunning Then
                asyncConnection.Disconnect(1)
            End If
            myInfo.Text = "Error occurred while connecting"

        End Try
    End Sub
    Private Sub MyWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles myWindow.Loaded
        myInfo.Text = "Working..."

        Call Creo_Connect()

        Call DecideFileType()
    End Sub

    Private Sub DecideFileType()

        Try
            If model Is Nothing Then
                MsgBox("Model is not present",, "Script message")
                asyncConnection.Disconnect(1)
                Environment.Exit(0)
            End If

            If activeserver.IsObjectCheckedOut(activeserver.ActiveWorkspace, model.FileName) Then
                MsgBox("Please check in model first...",, "Script Message")
                asyncConnection.Disconnect(1)
                Environment.Exit(0)
            End If

            Select Case model.ReleaseLevel
                Case "Concept"
                    State = "C"
                Case "Design", "Redesign"
                    State = "D"
                Case "PreReleased", "PreReleased RD"
                    State = "P"
                Case "Released"
                    State = "R"
                Case Else
            End Select

            Select Case model.Type
                Case 0
                    FileEnd = ".stp"

                Case 1
                    FileEnd = ".stp"

                Case 2
                    FileEnd = ".pdf"

                Case Else
                    MsgBox("Model not supported. Only Drawings, Models or Assemblies", "Script Message")
                    asyncConnection.Disconnect(1)
                    Environment.Exit(0)
            End Select

            FileNameComplete = model.FullName + "_" + model.Revision + "_" + model.Version + "_" + State + FileEnd

            Call ExportFileToDisc(FileNameComplete, model.Type)
        Catch ex As Exception

        End Try

    End Sub

    Sub TestForDir(workdir As String)
        Try
            If Dir(workdir, vbDirectory) = "" Then
                MkDir(workdir)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ExportFileToDisc(FileNameComplete As String, ConvertType As Integer)
        Dim Workdir As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory).ToString() & "\Fileshuffler Files\"
        Dim Destination As String = Workdir & FileNameComplete

        TestForDir(Workdir)

        myInfo.Text = FileNameComplete.ToString()

        Try
            If (ConvertType = 0) Then 'Export assy to STEP
                Dim cDesExStep As CCpfcSTEP3DExportInstructions
                Dim DesFlags As IpfcGeometryFlags
                Dim Des3DEx As IpfcExport3DInstructions
                Dim DesEx As IpfcExportInstructions
                Dim DesExStep As IpfcSTEP3DExportInstructions

                cDesExStep = New CCpfcSTEP3DExportInstructions
                DesFlags = (New CCpfcGeometryFlags).Create()
                DesFlags.AsSolids = True
                DesExStep = cDesExStep.Create(EpfcAssemblyConfiguration.EpfcEXPORT_ASM_SINGLE_FILE, DesFlags)
                Des3DEx = DesExStep
                DesEx = Des3DEx

                session.CurrentModel.Export(Destination, Des3DEx)

            ElseIf (ConvertType = 1) Then 'Export model to STEP
                Dim cDesExStep As CCpfcSTEP3DExportInstructions
                Dim DesFlags As IpfcGeometryFlags
                Dim Des3DEx As IpfcExport3DInstructions
                Dim DesEx As IpfcExportInstructions
                Dim DesExStep As IpfcSTEP3DExportInstructions

                cDesExStep = New CCpfcSTEP3DExportInstructions
                DesFlags = (New CCpfcGeometryFlags).Create()
                DesFlags.AsSolids = True
                DesExStep = cDesExStep.Create(EpfcAssemblyConfiguration.EpfcEXPORT_ASM_FLAT_FILE, DesFlags)
                Des3DEx = DesExStep
                DesEx = Des3DEx

                session.CurrentModel.Export(Destination, Des3DEx)
            ElseIf (ConvertType = 2) Then 'Export drawing to PDF

                Dim Drawing As IpfcModel2D
                Dim Sheet As IpfcSheetOwner
                Dim numSheets As Integer

                Dim PDFExportInstrCreate As New CCpfcPDFExportInstructions
                Dim PDFExportInstr As IpfcPDFExportInstructions
                PDFExportInstr = PDFExportInstrCreate.Create
                Dim PDF_Options As New CpfcPDFOptions

                ' Set Stroke All Fonts PDF Option
                Dim PDFOptionCreate_SAF As New CCpfcPDFOption
                Dim PDFOption_SAF As IpfcPDFOption
                PDFOption_SAF = PDFOptionCreate_SAF.Create
                PDFOption_SAF.OptionType = EpfcPDFOptionType.EpfcPDFOPT_FONT_STROKE
                Dim newArg_SAF As New CMpfcArgument
                PDFOption_SAF.OptionValue = newArg_SAF.CreateIntArgValue(EpfcPDFFontStrokeMode.EpfcPDF_USE_TRUE_TYPE_FONTS)
                Call PDF_Options.Append(PDFOption_SAF)

                ' Set COLOR_DEPTH value (Set EpfcPDF_CD_MONO to have Black & White output)
                Dim PDFOptionCreate_CD As New CCpfcPDFOption
                Dim PDFOption_CD As IpfcPDFOption
                PDFOption_CD = PDFOptionCreate_CD.Create
                PDFOption_CD.OptionType = EpfcPDFOptionType.EpfcPDFOPT_COLOR_DEPTH
                Dim newArg_CD As New CMpfcArgument
                PDFOption_CD.OptionValue = newArg_CD.CreateIntArgValue(EpfcPDFColorDepth.EpfcPDF_CD_MONO)
                Call PDF_Options.Append(PDFOption_CD)

                ' Set PDF EpfcPDFOPT_LAUNCH_VIEWER(Set TRUE to Launch Adobe reader)
                Dim PDFOptionCreate_LV As New CCpfcPDFOption
                Dim PDFOption_LV As IpfcPDFOption
                PDFOption_LV = PDFOptionCreate_LV.Create
                PDFOption_LV.OptionType = EpfcPDFOptionType.EpfcPDFOPT_LAUNCH_VIEWER
                Dim newArg_LV As New CMpfcArgument
                PDFOption_LV.OptionValue = newArg_LV.CreateBoolArgValue(True)
                Call PDF_Options.Append(PDFOption_LV)

                PDFExportInstr.Options = PDF_Options

                Sheet = CType(session.CurrentModel, IpfcSheetOwner)
                numSheets = Sheet.NumberOfSheets

                'Loop through every sheet and regenerate before creating PDF 
                For index = 1 To numSheets
                    Sheet.CurrentSheetNumber = index
                    Sheet.RegenerateSheet(index)
                Next


                session.CurrentModel.Export(Destination, PDFExportInstr)


            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub MyButton_Click(sender As Object, e As RoutedEventArgs) Handles myButton.Click
        Try
            asyncConnection.Disconnect(1)
        Catch ex As Exception

        End Try
        Close()
    End Sub

    Private Sub myWindow_Closing(sender As Object, e As ComponentModel.CancelEventArgs) Handles myWindow.Closing
        Try
            asyncConnection.Disconnect(1)
        Catch ex As Exception

        End Try

    End Sub
End Class

