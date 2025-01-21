Option Strict Off
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports NXOpen
Imports NXOpen.UF

Module NXJournal

    Sub Main()

        Dim theSession As Session = Session.GetSession
        
		Dim theUI As UI = UI.GetUI()
		If IsNothing(theSession.Parts.Display) Then
            MessageBox.Show("Active Part Required", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        Dim myPdfExporter As New NXJ_PdfExporter

        '$ prompt user for output folder
        myPdfExporter.PickExportFolder()

        '$ preliminary print?
        myPdfExporter.PromptPreliminaryPrint()

        '$ desired watermark text (if preliminary print)
        myPdfExporter.PromptWatermarkText()
        myPdfExporter.WatermarkAddDatestamp = True

        '$ show confirmation dialog box on completion
        myPdfExporter.ShowConfirmationDialog = False

		'additional code added by Noah Curfman to loop through all open parts
		
        Dim partName As String = Nothing
		
		Dim relatedDrawings() As String
		
		Dim basePart1 As NXOpen.BasePart = Nothing
		Dim partLoadStatus1 As PartLoadStatus = Nothing	
			
		'open drawing files for all loaded parts and assemblies
		
		For Each part As Part In theSession.Parts
			partName = part.Leaf
			'theUI.NXMessageBox.Show("Part Name", NXMessageBox.DialogType.Information, "The part name is: " & partName)
			If Not partName.Contains("DWG")
				theSession.Parts.SetDisplay(part, False, False, Nothing)
				'theUI.NXMessageBox.Show("Oi ", NXMessageBox.DialogType.Error, "In the open drawing files loop" )        
				Try
					relatedDrawings = part.PDMPart.GetRelatedDrawings()
				Catch ex As Exception
					theUI.NXMessageBox.Show("Error", NXMessageBox.DialogType.Error, "Unable to retrieve related drawings: " & ex.Message)
					Return
				End Try
	
				' Open all the related drawings for part files
				For Each dwg As String In relatedDrawings
					Try
						'theUI.NXMessageBox.Show("Part Name", NXMessageBox.DialogType.Information, "The part name is: " & dwg)
						basePart1 = theSession.Parts.OpenActiveDisplay(dwg, NXOpen.DisplayPartOption.AllowAdditional, partLoadStatus1)
					Catch ex As Exception
						theUI.NXMessageBox.Show("Error", NXMessageBox.DialogType.Error, "Failed to open related drawing: " & ex.Message)
					End Try
				Next
			
			End If
		Next

		'export PDFS for all open drawing files
		
        For Each part As Part In theSession.Parts
		    partName = part.Leaf
			If partName.Contains("DWG")
			
					theSession.Parts.SetDisplay(part, False, False, Nothing)

					myPdfExporter.Part = part
					Try
						myPdfExporter.Commit()
					Catch ex As Exception
						MessageBox.Show("Error:" & ControlChars.CrLf & ex.GetType.ToString & " : " & ex.Message, "PDF export error", MessageBoxButtons.OK, MessageBoxIcon.Error)

					'Finally

					End Try
			End If

		Next
	myPdfExporter = Nothing
    End Sub
	
End Module


'*******************************************************************************



Class NXJ_PdfExporter


#Region "information"

    'NXJournaling.com
    'Jeff Gesiakowski
    'December 9, 2013
    '
    'NX 8.5
    'class to export drawing sheets to pdf file
    '
    'Please send any bug reports and/or feature requests to: info@nxjournaling.com
    '
    'version 0.4 {beta}, initial public release
    '  special thanks to Mike H. and Ryan G.
    '
    'November 3, 2014
    'update to version 0.6
    '  added a public Sort(yourSortFunction) method to sort the drawing sheet collection according to a custom supplied function.
    '
    'November 7, 2014
    'Added public property: ExportSheetsIndividually and related code changes default value: False
    'Changing this property to True will cause each sheet to be exported to an individual pdf file in the specified export folder.
    '
    'Added new public method: New(byval thePart as Part)
    '  allows you to specify the part to use at the time of the NXJ_PdfExporter object creation
    '
    '
    'December 1, 2014
    'update to version 1.0
    'Added public property: SkipBlankSheets [Boolean] {read/write} default value: True
    '   If the drawing sheet contains no visible objects, it is not output to the pdf file.
    '   Checking the sheet for visible objects requires the sheet to be opened;
    '   display updating is suppressed while the check takes place.
    'Bugfix:
    '   If the PickExportFolder method was used and the user pressed cancel, a later call to Commit would still execute and send the output to a temp folder.
    '   Now, if cancel is pressed on the folder browser dialog a boolean flag is set (_cancelOutput), which the other methods will check before executing.
    '
    '
    'December 4, 2014
    'update to version 1.0.1
    'Made changes to .OutputPdfFileName property Set method: you can pass in the full path to the file (e.g. C:\temp\pdf-output\12345.pdf);
    'or you can simply pass in the base file name for the pdf file (e.g. 12345 or 12345.pdf). The full path will be built based on the
    '.OutputFolder property (default value - same folder as the part file, if an invalid folder path is specified it will default to the user's
    '"Documents" folder).
    '
    '
    'Public Properties:
    '  ExportSheetsIndividually [Boolean] {read/write} - flag indicating that the drawing sheets should be output to individual pdf files.
    '           cannot be used if ExportToTc = True
    '           default value: False
    '  ExportToTc [Boolean] {read/write} - flag indicating that the pdf should be output to the TC dataset, False value = output to filesystem
    '           default value: False
    '  IsTcRunning [Boolean] {read only} - True if NX is running under TC, false if native NX
    '  OpenPdf [Boolean] {read/write} - flag to indicate whether the journal should attempt to open the pdf after creation
    '           default value: False
    '  OutputFolder [String] {read/write} - path to the output folder
    '           default value (native): folder of current part
    '           default value (TC): user's Documents folder
    '  OutputPdfFileName [String] {read/write} - full file name of outut pdf file (if exporting to filesystem)
    '           default value (native): <folder of current part>\<part name>_<part revision>{_preliminary}.pdf
    '           default value (TC): <current user's Documents folder>\<DB_PART_NO>_<DB_PART_REV>{_preliminary}.pdf
    '  OverwritePdf [Boolean] {read/write} - flag indicating that the pdf file should be overwritten if it already exists
    '                                           currently only applies when exporting to the filesystem
    '           default value: True
    '  Part [NXOpen.Part] {read/write} - part that contains the drawing sheets of interest
    '           default value: none, must be set by user
    '  PartFilePath [String] {read only} - for native NX part files, the path to the part file
    '  PartNumber [String] {read only} - for native NX part files: part file name
    '                                    for TC files: value of DB_PART_NO attribute
    '  PartRevision [String] {read only} - for native NX part files: value of part "Revision" attribute, if present
    '                                      for TC files: value of DB_PART_REV
    '  PreliminaryPrint [Boolean] {read/write} - flag indicating that the pdf should be marked as an "preliminary"
    '                                       when set to True, the output file will be named <filename>_preliminary.pdf
    '           default value: False
    '  SheetCount [Integer] {read only} - integer indicating the total number of drawing sheets found in the file
    '  ShowConfirmationDialog [Boolean] {read/write} - flag indicating whether to show the user a confirmation dialog after pdf is created
    '                                                   if set to True and ExportToTc = False, user will be asked if they want to open the pdf file
    '                                                   if user chooses "Yes", the code will attempt to open the pdf with the default viewer
    '           default value: False
    '  SkipBlankSheets [Boolean] {read/write} - flag indicating if the user wants to skip drawing sheets with no visible objects.
    '           default value: True
    '  SortSheetsByName [Boolean] {read/write} - flag indicating that the sheets should be sorted by name before output to pdf
    '           default value: True
    '  TextAsPolylines [Boolean] {read/write} - flag indicating that text objects should be output as polylines instead of text objects
    '           default value: False        set this to True if you are using an NX font and the output changes the 'look' of the text
    '  UseWatermark [Boolean] {read/write} - flag indicating that watermark text should be applied to the face of the drawing
    '           default value: False
    '  WatermarkAddDatestamp [Boolean] {read/write} - flag indicating that today's date should be added to the end of the
    '                                                   watermark text
    '           default value: True
    '  WatermarkText [String] {read/write} - watermark text to use
    '           default value: "PRELIMINARY PRINT NOT TO BE USED FOR PRODUCTION"
    '
    'Public Methods:
    '  New() - initializes a new instance of the class
    '  New(byval thePart as Part) - initializes a new instance of the class and specifies the NXOpen.Part to use
    '  PickExportFolder() - displays a FolderPicker dialog box, the user's choice will be set as the output folder
    '  PromptPreliminaryPrint() - displays a yes/no dialog box asking the user if the print should be marked as preliminary
    '                               if user chooses "Yes", PreliminaryPrint and UseWatermark are set to True
    '  PromptWatermarkText() - displays an input box prompting the user to enter text to use for the watermark
    '                           if cancel is pressed, the default value is used
    '                           if Me.UseWatermark = True, an input box will appear prompting the user for the desired watermark text. Initial text = Me.WatermarkText
    '                           if Me.UseWatermark = False, calling this method will have no effect
    '  Commit() - using the specified options, export the given part's sheets to pdf
    '  SortDrawingSheets() - sorts the drawing sheets in alphabetic order
    '  SortDrawingSheets(ByVal customSortFunction As System.Comparison(Of NXOpen.Drawings.DrawingSheet)) - sorts the drawing sheets by the custom supplied function
    '    signature of the sort function must be: {function name}(byval x as Drawings.Drawingsheet, byval y as Drawings.DrawingSheet) as Integer
    '    a return value < 0 means x comes before y
    '    a return value > 0 means x comes after y
    '    a return value = 0 means they are equal (it doesn't matter which is first in the resulting list)
    '    after writing your custom sort function in the module, pass it in like this: myPdfExporter.Sort(AddressOf {function name})


#End Region


#Region "properties and private variables"

    Private Const Version As String = "1.0.1"

    Private _theSession As Session = Session.GetSession
    Private _theUfSession As UFSession = UFSession.GetUFSession
    Private lg As LogFile = _theSession.LogFile

    Private _cancelOutput As Boolean = False
    Private _drawingSheets As New List(Of Drawings.DrawingSheet)

    Private _exportFile As String = ""
    Private _partUnits As Integer
    Private _watermarkTextFinal As String = ""
    Private _outputPdfFiles As New List(Of String)

    Private _exportSheetsIndividually As Boolean = False
    Public Property ExportSheetsIndividually() As Boolean
        Get
            Return _exportSheetsIndividually
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property ExportSheetsIndividually")
            _exportSheetsIndividually = value
            lg.WriteLine("  ExportSheetsIndividually: " & value.ToString)
            lg.WriteLine("exiting Set Property ExportSheetsIndividually")
            lg.WriteLine("")
        End Set
    End Property

    Private _exportToTC As Boolean = False
    Public Property ExportToTc() As Boolean
        Get
            Return _exportToTC
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property ExportToTc")
            _exportToTC = value
            lg.WriteLine("  exportToTc: " & _exportToTC.ToString)
            Me.GetOutputName()
            lg.WriteLine("exiting Set Property ExportToTc")
            lg.WriteLine("")
        End Set
    End Property

    Private _isTCRunning As Boolean
    Public ReadOnly Property IsTCRunning() As Boolean
        Get
            Return _isTCRunning
        End Get
    End Property

    Private _openPdf As Boolean = False
    Public Property OpenPdf() As Boolean
        Get
            Return _openPdf
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property OpenPdf")
            _openPdf = value
            lg.WriteLine("  openPdf: " & _openPdf.ToString)
            lg.WriteLine("exiting Set Property OpenPdf")
            lg.WriteLine("")
        End Set
    End Property

    Private _outputFolder As String = ""
    Public Property OutputFolder() As String
        Get
            Return _outputFolder
        End Get
        Set(ByVal value As String)
            lg.WriteLine("Set Property OutputFolder")
            If _cancelOutput Then
                lg.WriteLine("  export pdf canceled")
                Exit Property
            End If
            If Not Directory.Exists(value) Then
                Try
                    lg.WriteLine("  specified directory does not exist, trying to create it...")
                    Directory.CreateDirectory(value)
                    lg.WriteLine("  directory created: " & value)
                Catch ex As Exception
                    lg.WriteLine("  ** error while creating directory: " & value)
                    lg.WriteLine("  " & ex.GetType.ToString & " : " & ex.Message)
                    lg.WriteLine("  defaulting to: " & My.Computer.FileSystem.SpecialDirectories.MyDocuments)
                    value = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                End Try
            End If
            _outputFolder = value
            _outputPdfFile = IO.Path.Combine(_outputFolder, _exportFile & ".pdf")
            lg.WriteLine("  outputFolder: " & _outputFolder)
            lg.WriteLine("  outputPdfFile: " & _outputPdfFile)
            lg.WriteLine("exiting Set Property OutputFolder")
            lg.WriteLine("")
        End Set
    End Property

    Private _outputPdfFile As String = ""
    Public Property OutputPdfFileName() As String
        Get
            Return _outputPdfFile
        End Get
        Set(ByVal value As String)
            lg.WriteLine("Set Property OutputPdfFileName")
            lg.WriteLine("  value passed to property: " & value)
            _exportFile = IO.Path.GetFileName(value)
            If _exportFile.Substring(_exportFile.Length - 4, 4).ToLower = ".pdf" Then
                'strip off ".pdf" extension
                _exportFile = _exportFile.Substring(_exportFile.Length - 4, 4)
            End If
            lg.WriteLine("  _exportFile: " & _exportFile)
            If Not value.Contains("\") Then
                lg.WriteLine("  does not appear to contain path information")
                'file name only, need to add output path
                _outputPdfFile = IO.Path.Combine(Me.OutputFolder, _exportFile & ".pdf")
            Else
                'value contains path, update _outputFolder
                lg.WriteLine("  value contains path, updating the output folder...")
                lg.WriteLine("  parent path: " & Me.GetParentPath(value))
                Me.OutputFolder = Me.GetParentPath(value)
                _outputPdfFile = IO.Path.Combine(Me.OutputFolder, _exportFile & ".pdf")
            End If
            '_outputPdfFile = value
            lg.WriteLine("  outputPdfFile: " & _outputPdfFile)
            lg.WriteLine("  outputFolder: " & Me.OutputFolder)
            lg.WriteLine("exiting Set Property OutputPdfFileName")
            lg.WriteLine("")
        End Set
    End Property

    Private _overwritePdf As Boolean = True
    Public Property OverwritePdf() As Boolean
        Get
            Return _overwritePdf
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property OverwritePdf")
            _overwritePdf = value
            lg.WriteLine("  overwritePdf: " & _overwritePdf.ToString)
            lg.WriteLine("exiting Set Property OverwritePdf")
            lg.WriteLine("")
        End Set
    End Property

    Private _thePart As Part = Nothing
    Public Property Part() As Part
        Get
            Return _thePart
        End Get
        Set(ByVal value As Part)
            lg.WriteLine("Set Property Part")
            _thePart = value
            _partUnits = _thePart.PartUnits
            Me.GetPartInfo()
            Me.GetDrawingSheets()
            If Me.SortSheetsByName Then
                Me.SortDrawingSheets()
            End If
            lg.WriteLine("exiting Set Property Part")
            lg.WriteLine("")
        End Set
    End Property

    Private _partFilePath As String
    Public ReadOnly Property PartFilePath() As String
        Get
            Return _partFilePath
        End Get
    End Property

    Private _partNumber As String
    Public ReadOnly Property PartNumber() As String
        Get
            Return _partNumber
        End Get
    End Property

    Private _partRevision As String = ""
    Public ReadOnly Property PartRevision() As String
        Get
            Return _partRevision
        End Get
    End Property

    Private _preliminaryPrint As Boolean = False
    Public Property PreliminaryPrint() As Boolean
        Get
            Return _preliminaryPrint
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property PreliminaryPrint")
            _preliminaryPrint = value
            If String.IsNullOrEmpty(_exportFile) Then
                'do nothing
            Else
                Me.GetOutputName()
            End If
            lg.WriteLine("  preliminaryPrint: " & _preliminaryPrint.ToString)
            lg.WriteLine("exiting Set Property PreliminaryPrint")
            lg.WriteLine("")
        End Set
    End Property

    Public ReadOnly Property SheetCount() As Integer
        Get
            Return _drawingSheets.Count
        End Get
    End Property

    Private _showConfirmationDialog As Boolean = False
    Public Property ShowConfirmationDialog() As Boolean
        Get
            Return _showConfirmationDialog
        End Get
        Set(ByVal value As Boolean)
            _showConfirmationDialog = value
        End Set
    End Property

    Private _skipBlankSheets As Boolean = True
    Public Property SkipBlankSheets() As Boolean
        Get
            Return _skipBlankSheets
        End Get
        Set(ByVal value As Boolean)
            _skipBlankSheets = value
        End Set
    End Property

    Private _sortSheetsByName As Boolean
    Public Property SortSheetsByName() As Boolean
        Get
            Return _sortSheetsByName
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property SortSheetsByName")
            _sortSheetsByName = value
            If _sortSheetsByName = True Then
                'sort alphabetically by sheet name
                Me.SortDrawingSheets()
            Else
                'get original collection order of sheets
                Me.GetDrawingSheets()
            End If
            lg.WriteLine("  sortSheetsByName: " & _sortSheetsByName.ToString)
            lg.WriteLine("exiting Set Property SortSheetsByName")
            lg.WriteLine("")
        End Set
    End Property

    Private _textAsPolylines As Boolean = False
    Public Property TextAsPolylines() As Boolean
        Get
            Return _textAsPolylines
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property TextAsPolylines")
            _textAsPolylines = value
            lg.WriteLine("  textAsPolylines: " & _textAsPolylines.ToString)
            lg.WriteLine("exiting Set Property TextAsPolylines")
            lg.WriteLine("")
        End Set
    End Property

    Private _useWatermark As Boolean = False
    Public Property UseWatermark() As Boolean
        Get
            Return _useWatermark
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property UseWatermark")
            _useWatermark = value
            lg.WriteLine("  useWatermark: " & _useWatermark.ToString)
            lg.WriteLine("exiting Set Property UseWatermark")
            lg.WriteLine("")
        End Set
    End Property

    Private _watermarkAddDatestamp As Boolean = True
    Public Property WatermarkAddDatestamp() As Boolean
        Get
            Return _watermarkAddDatestamp
        End Get
        Set(ByVal value As Boolean)
            lg.WriteLine("Set Property WatermarkAddDatestamp")
            _watermarkAddDatestamp = value
            lg.WriteLine("  watermarkAddDatestamp: " & _watermarkAddDatestamp.ToString)
            If _watermarkAddDatestamp Then
                'to do: internationalization for dates
                _watermarkTextFinal = _watermarkText & " " & Today
            Else
                _watermarkTextFinal = _watermarkText
            End If
            lg.WriteLine("  watermarkTextFinal: " & _watermarkTextFinal)
            lg.WriteLine("exiting Set Property WatermarkAddDatestamp")
            lg.WriteLine("")
        End Set
    End Property

    Private _watermarkText As String = "PRELIMINARY PRINT NOT TO BE USED FOR PRODUCTION"
    Public Property WatermarkText() As String
        Get
            Return _watermarkText
        End Get
        Set(ByVal value As String)
            lg.WriteLine("Set Property WatermarkText")
            _watermarkText = value
            lg.WriteLine("  watermarkText: " & _watermarkText)
            lg.WriteLine("exiting Set Property WatermarkText")
            lg.WriteLine("")
        End Set
    End Property


#End Region





#Region "public methods"

    Public Sub New()

        Me.StartLog()

    End Sub

    Public Sub New(ByVal thePart As NXOpen.Part)

        Me.StartLog()
        Me.Part = thePart

    End Sub

    Public Sub PickExportFolder()

        'Requires:
        '    Imports System.IO
        '    Imports System.Windows.Forms
        'if the user presses OK on the dialog box, the chosen path is returned
        'if the user presses cancel on the dialog box, 0 is returned
        lg.WriteLine("Sub PickExportFolder")

        If Me.ExportToTc Then
            lg.WriteLine("  N/A when ExportToTc = True")
            lg.WriteLine("  exiting Sub PickExportFolder")
            lg.WriteLine("")
            Return
        End If

        Dim strLastPath As String

        'Key will show up in HKEY_CURRENT_USER\Software\VB and VBA Program Settings
        Try
            'Get the last path used from the registry
            lg.WriteLine("  attempting to retrieve last export path from registry...")
            strLastPath = GetSetting("NX journal", "Export pdf", "ExportPath")
            'msgbox("Last Path: " & strLastPath)
        Catch e As ArgumentException
            lg.WriteLine("  ** Argument Exception: " & e.Message)
        Catch e As Exception
            lg.WriteLine("  ** Exception type: " & e.GetType.ToString)
            lg.WriteLine("  ** Exception message: " & e.Message)
            'MsgBox(e.GetType.ToString)
        Finally
        End Try

        Dim FolderBrowserDialog1 As New FolderBrowserDialog

        ' Then use the following code to create the Dialog window
        ' Change the .SelectedPath property to the default location
        With FolderBrowserDialog1
            ' Desktop is the root folder in the dialog.
            .RootFolder = Environment.SpecialFolder.Desktop
            ' Select the D:\home directory on entry.
            If Directory.Exists(strLastPath) Then
                .SelectedPath = strLastPath
            Else
                .SelectedPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            End If
            ' Prompt the user with a custom message.
            .Description = "Select the directory to export .pdf file"
            If .ShowDialog = DialogResult.OK Then
                ' Display the selected folder if the user clicked on the OK button.
                Me.OutputFolder = .SelectedPath
                lg.WriteLine("  selected output path: " & .SelectedPath)
                ' save the output folder path in the registry for use on next run
                SaveSetting("NX journal", "Export pdf", "ExportPath", .SelectedPath)
            Else
                'user pressed 'cancel', keep original value of output folder
                _cancelOutput = True
                Me.OutputFolder = Nothing
                lg.WriteLine("  folder browser dialog cancel button pressed")
                lg.WriteLine("  current output path: {nothing}")
            End If
        End With

        lg.WriteLine("exiting Sub PickExportFolder")
        lg.WriteLine("")

    End Sub

    Public Sub PromptPreliminaryPrint()

        lg.WriteLine("Sub PromptPreliminaryPrint")

        If _cancelOutput Then
            lg.WriteLine("  output canceled")
            Return
        End If

        Dim rspPreliminaryPrint As DialogResult
        rspPreliminaryPrint = MessageBox.Show("Add preliminary print watermark?", "Add Watermark?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If rspPreliminaryPrint = DialogResult.Yes Then
            Me.PreliminaryPrint = True
            Me.UseWatermark = True
            lg.WriteLine("  this is a preliminary print")
        Else
            Me.PreliminaryPrint = False
            lg.WriteLine("  this is not a preliminary print")
        End If

        lg.WriteLine("exiting Sub PromptPreliminaryPrint")
        lg.WriteLine("")

    End Sub

    Public Sub PromptWatermarkText()

        lg.WriteLine("Sub PromptWatermarkText")
        lg.WriteLine("  useWatermark: " & Me.UseWatermark.ToString)

        Dim theWatermarkText As String = Me.WatermarkText

        If Me.UseWatermark Then
            theWatermarkText = InputBox("Enter watermark text", "Watermark", theWatermarkText)
            Me.WatermarkText = theWatermarkText
        Else
            lg.WriteLine("  suppressing watermark prompt")
        End If

        lg.WriteLine("exiting Sub PromptWatermarkText")
        lg.WriteLine("")

    End Sub

    Public Sub SortDrawingSheets()

        If _cancelOutput Then
            Return
        End If

        If Not IsNothing(_thePart) Then
            Me.GetDrawingSheets()
            _drawingSheets.Sort(AddressOf Me.CompareSheetNames)
        End If

    End Sub

    Public Sub SortDrawingSheets(ByVal customSortFunction As System.Comparison(Of NXOpen.Drawings.DrawingSheet))

        If _cancelOutput Then
            Return
        End If

        If Not IsNothing(_thePart) Then
            Me.GetDrawingSheets()
            _drawingSheets.Sort(customSortFunction)
        End If

    End Sub

    Public Sub Commit()

        If _cancelOutput Then
            Return
        End If

        lg.WriteLine("Sub Commit")
        lg.WriteLine("  number of drawing sheets in part file: " & _drawingSheets.Count.ToString)

        _outputPdfFiles.Clear()
        For Each tempSheet As Drawings.DrawingSheet In _drawingSheets
            If Me.PreliminaryPrint Then
                _outputPdfFiles.Add(IO.Path.Combine(Me.OutputFolder, tempSheet.Name & "_preliminary.pdf"))
            Else
                _outputPdfFiles.Add(IO.Path.Combine(Me.OutputFolder, tempSheet.Name & ".pdf"))
            End If
        Next

        'make sure we can output to the specified file(s)
        If Me.ExportSheetsIndividually Then
            'check each sheet
            For Each newPdf As String In _outputPdfFiles

                If Not Me.DeleteExistingPdfFile(newPdf) Then
                    If _overwritePdf Then
                        'file could not be deleted
                        MessageBox.Show("The pdf file: " & newPdf & " exists and could not be overwritten." & ControlChars.NewLine & _
                                        "PDF export exiting", "PDF export error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        'file already exists and will not be overwritten
                        MessageBox.Show("The pdf file: " & newPdf & " exists and the overwrite option is set to False." & ControlChars.NewLine & _
                                        "PDF export exiting", "PDF file exists", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                    Return

                End If

            Next

        Else
            'check _outputPdfFile
            If Not Me.DeleteExistingPdfFile(_outputPdfFile) Then
                If _overwritePdf Then
                    'file could not be deleted
                    MessageBox.Show("The pdf file: " & _outputPdfFile & " exists and could not be overwritten." & ControlChars.NewLine & _
                                    "PDF export exiting", "PDF export error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    'file already exists and will not be overwritten
                    MessageBox.Show("The pdf file: " & _outputPdfFile & " exists and the overwrite option is set to False." & ControlChars.NewLine & _
                                    "PDF export exiting", "PDF file exists", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return
            End If

        End If

        Dim sheetCount As Integer = 0
        Dim sheetsExported As Integer = 0

        Dim numPlists As Integer = 0
        Dim myPlists() As Tag

        _theUfSession.Plist.AskTags(myPlists, numPlists)
        For i As Integer = 0 To numPlists - 1
            _theUfSession.Plist.Update(myPlists(i))
        Next

        For Each tempSheet As Drawings.DrawingSheet In _drawingSheets

            sheetCount += 1

            lg.WriteLine("  working on sheet: " & tempSheet.Name)
            lg.WriteLine("  sheetCount: " & sheetCount.ToString)


			'Removed this to prevent view update problems - Noah Curfman
            'update any views that are out of date
            'lg.WriteLine("  updating OutOfDate views on sheet: " & tempSheet.Name)
            'Me.Part.DraftingViews.UpdateViews(Drawings.DraftingViewCollection.ViewUpdateOption.OutOfDate, tempSheet)

        Next

        If Me._drawingSheets.Count > 0 Then

            lg.WriteLine("  done updating views on all sheets")

            Try
                If Me.ExportSheetsIndividually Then
                    For Each tempSheet As Drawings.DrawingSheet In _drawingSheets
                        lg.WriteLine("  calling Sub ExportPdf")
                        lg.WriteLine("")
                        If Me.PreliminaryPrint Then
                            Me.ExportPdf(tempSheet, IO.Path.Combine(Me.OutputFolder, tempSheet.Name & "_preliminary.pdf"))
                        Else
                            Me.ExportPdf(tempSheet, IO.Path.Combine(Me.OutputFolder, tempSheet.Name & ".pdf"))
                        End If
                    Next
                Else
                    lg.WriteLine("  calling Sub ExportPdf")
                    lg.WriteLine("")
                    Me.ExportPdf()
                End If
            Catch ex As Exception
                lg.WriteLine("  ** error exporting PDF")
                lg.WriteLine("  " & ex.GetType.ToString & " : " & ex.Message)
                'MessageBox.Show("Error occurred in PDF export" & vbCrLf & ex.Message, "Error exporting PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Throw ex
            End Try

        Else
            'no sheets in file
            lg.WriteLine("  ** no drawing sheets in file: " & Me._partNumber)

        End If

        If Me.ShowConfirmationDialog Then
            Me.DisplayConfirmationDialog()
        End If

        If (Not Me.ExportToTc) AndAlso (Me.OpenPdf) AndAlso (Me._drawingSheets.Count > 0) Then
            'open new pdf print
            lg.WriteLine("  trying to open newly created pdf file")
            Try
                If Me.ExportSheetsIndividually Then
                    For Each newPdf As String In _outputPdfFiles
                        System.Diagnostics.Process.Start(newPdf)
                    Next
                Else
                    System.Diagnostics.Process.Start(Me.OutputPdfFileName)
                End If
                lg.WriteLine("  pdf open process successful")
            Catch ex As Exception
                lg.WriteLine("  ** error opening pdf **")
                lg.WriteLine("  " & ex.GetType.ToString & " : " & ex.Message)
            End Try
        End If

        lg.WriteLine("  exiting Sub ExportSheetsToPdf")
        lg.WriteLine("")

    End Sub

#End Region





#Region "private methods"

    Private Sub GetPartInfo()

        lg.WriteLine("Sub GetPartInfo")

        If Me.IsTCRunning Then
            _partNumber = _thePart.GetStringAttribute("DB_PART_NO")
            _partRevision = _thePart.GetStringAttribute("DB_PART_REV")

            lg.WriteLine("  TC running")
            lg.WriteLine("  partNumber: " & _partNumber)
            lg.WriteLine("  partRevision: " & _partRevision)

        Else 'running in native mode

            _partNumber = IO.Path.GetFileNameWithoutExtension(_thePart.FullPath)
            _partFilePath = IO.Directory.GetParent(_thePart.FullPath).ToString

            lg.WriteLine("  Native NX")
            lg.WriteLine("  partNumber: " & _partNumber)
            lg.WriteLine("  partFilePath: " & _partFilePath)

            Try
                _partRevision = _thePart.GetStringAttribute("REVISION")
                _partRevision = _partRevision.Trim
            Catch ex As Exception
                _partRevision = ""
            End Try

            lg.WriteLine("  partRevision: " & _partRevision)

        End If

        If String.IsNullOrEmpty(_partRevision) Then
            _exportFile = _partNumber
        Else
            _exportFile = _partNumber & "_" & _partRevision
        End If

        lg.WriteLine("")
        Me.GetOutputName()

        lg.WriteLine("  exportFile: " & _exportFile)
        lg.WriteLine("  outputPdfFile: " & _outputPdfFile)
        lg.WriteLine("  exiting Sub GetPartInfo")
        lg.WriteLine("")

    End Sub

    Private Sub GetOutputName()

        lg.WriteLine("Sub GetOutputName")

        _exportFile.Replace("_preliminary", "")
        _exportFile.Replace("_PDF_1", "")

        If IsNothing(Me.Part) Then
            lg.WriteLine("  Me.Part is Nothing")
            lg.WriteLine("  exiting Sub GetOutputName")
            lg.WriteLine("")
            Return
        End If

        If Not IsTCRunning And _preliminaryPrint Then
            _exportFile &= "_preliminary"
        End If

        If Me.ExportToTc Then      'export to Teamcenter dataset
            lg.WriteLine("  export to TC option chosen")
            If Me.IsTCRunning Then
                lg.WriteLine("  TC is running")
                _exportFile &= "_PDF_1"
            Else
                'error, cannot export to a dataset if TC is not running
                lg.WriteLine("  ** error: export to TC option chosen, but TC is not running")
                'todo: throw error
            End If
        Else                    'export to file system
            lg.WriteLine("  export to filesystem option chosen")
            If Me.IsTCRunning Then
                lg.WriteLine("  TC is running")
                'exporting from TC to filesystem, no part folder to default to
                'default to "MyDocuments" folder
                'Commented this line out and added next line to allow select folder to work in TC - Noah Curfman
				'_outputPdfFile = IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyDocuments, _exportFile & ".pdf")
				_outputPdfFile = IO.Path.Combine(_outputFolder, _exportFile & ".pdf")
            Else
                lg.WriteLine("  native NX")
                'exporting from native to file system
                'use part folder as default output folder
                If _outputFolder = "" Then
                    _outputFolder = _partFilePath
                End If
                _outputPdfFile = IO.Path.Combine(_outputFolder, _exportFile & ".pdf")
            End If

        End If

        lg.WriteLine("  exiting Sub GetOutputName")
        lg.WriteLine("")

    End Sub

    Private Sub GetDrawingSheets()

        _drawingSheets.Clear()

        For Each tempSheet As Drawings.DrawingSheet In _thePart.DrawingSheets
            If _skipBlankSheets Then
                _theUfSession.Disp.SetDisplay(UFConstants.UF_DISP_SUPPRESS_DISPLAY)
                Dim currentSheet As Drawings.DrawingSheet = _thePart.DrawingSheets.CurrentDrawingSheet
                tempSheet.Open()
                If Not IsSheetEmpty(tempSheet) Then
                    _drawingSheets.Add(tempSheet)
                End If
                Try
                    currentSheet.Open()
                Catch ex As NXException
                    lg.WriteLine("  NX current sheet error: " & ex.Message)
                Catch ex As Exception
                    lg.WriteLine("  current sheet error: " & ex.Message)
                End Try
                _theUfSession.Disp.SetDisplay(UFConstants.UF_DISP_UNSUPPRESS_DISPLAY)
                _theUfSession.Disp.RegenerateDisplay()
            Else
                _drawingSheets.Add(tempSheet)
            End If
        Next

    End Sub

    Private Function CompareSheetNames(ByVal x As Drawings.DrawingSheet, ByVal y As Drawings.DrawingSheet) As Integer

        'case-insensitive sort
        Dim myStringComp As StringComparer = StringComparer.CurrentCultureIgnoreCase

        'for a case-sensitive sort (A-Z then a-z), change the above option to:
        'Dim myStringComp As StringComparer = StringComparer.CurrentCulture

        Return myStringComp.Compare(x.Name, y.Name)

    End Function

    Private Function GetParentPath(ByVal thePath As String) As String

        lg.WriteLine("Function GetParentPath(" & thePath & ")")

        Try
            Dim directoryInfo As System.IO.DirectoryInfo
            directoryInfo = System.IO.Directory.GetParent(thePath)
            lg.WriteLine("  returning: " & directoryInfo.FullName)
            lg.WriteLine("exiting Function GetParentPath")
            lg.WriteLine("")

            Return directoryInfo.FullName
        Catch ex As ArgumentNullException
            lg.WriteLine("  Path is a null reference.")
            Throw ex
        Catch ex As ArgumentException
            lg.WriteLine("  Path is an empty string, contains only white space, or contains invalid characters")
            Throw ex
        End Try

        lg.WriteLine("exiting Function GetParentPath")
        lg.WriteLine("")

    End Function

    Private Sub ExportPdf()

        lg.WriteLine("Sub ExportPdf")

        Dim printPDFBuilder1 As PrintPDFBuilder

        printPDFBuilder1 = _thePart.PlotManager.CreatePrintPdfbuilder()
        printPDFBuilder1.Scale = 1.0
        printPDFBuilder1.Colors = PrintPDFBuilder.Color.BlackOnWhite
        printPDFBuilder1.Size = PrintPDFBuilder.SizeOption.ScaleFactor
        printPDFBuilder1.RasterImages = True
        printPDFBuilder1.ImageResolution = PrintPDFBuilder.ImageResolutionOption.Medium

        If _thePart.PartUnits = BasePart.Units.Inches Then
            lg.WriteLine("  part units: English")
            printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.English
        Else
            lg.WriteLine("  part units: Metric")
            printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.Metric
        End If

        If _textAsPolylines Then
            lg.WriteLine("  output text as polylines")
            printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Polylines
        Else
            lg.WriteLine("  output text as text")
            printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Text
        End If

        lg.WriteLine("  useWatermark: " & _useWatermark.ToString)
        If _useWatermark Then
            printPDFBuilder1.AddWatermark = True
            printPDFBuilder1.Watermark = _watermarkTextFinal
        Else
            printPDFBuilder1.AddWatermark = False
            printPDFBuilder1.Watermark = ""
        End If

        lg.WriteLine("  export to TC? " & _exportToTC.ToString)
        If _exportToTC Then
            'output to dataset
            printPDFBuilder1.Relation = PrintPDFBuilder.RelationOption.Manifestation
            printPDFBuilder1.DatasetType = "PDF"
            printPDFBuilder1.NamedReferenceType = "PDF_Reference"
            'printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Overwrite
            printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
            printPDFBuilder1.DatasetName = _exportFile & "_PDF_1"
            lg.WriteLine("  dataset name: " & _exportFile)

            Try
                lg.WriteLine("  printPDFBuilder1.Assign")
                printPDFBuilder1.Assign()
            Catch ex As NXException
                lg.WriteLine("  ** error with printPDFBuilder1.Assign")
                lg.WriteLine("  " & ex.ErrorCode & " : " & ex.Message)
            End Try

        Else
            'output to filesystem
            lg.WriteLine("  pdf file: " & _outputPdfFile)
            printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Native
            printPDFBuilder1.Append = False
            printPDFBuilder1.Filename = _outputPdfFile

        End If

        printPDFBuilder1.SourceBuilder.SetSheets(_drawingSheets.ToArray)

        Dim nXObject1 As NXObject
        Try
            lg.WriteLine("  printPDFBuilder1.Commit")
            nXObject1 = printPDFBuilder1.Commit()

        Catch ex As NXException
            lg.WriteLine("  ** error with printPDFBuilder1.Commit")
            lg.WriteLine("  " & ex.ErrorCode & " : " & ex.Message)

            'If Me.ExportToTc Then

            '    Try
            '        lg.WriteLine("  trying new dataset option")
            '        printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
            '        printPDFBuilder1.Commit()
            '    Catch ex2 As NXException
            '        lg.WriteLine("  ** error with printPDFBuilder1.Commit")
            '        lg.WriteLine("  " & ex2.ErrorCode & " : " & ex2.Message)

            '    End Try

            'End If

        Finally
            printPDFBuilder1.Destroy()
        End Try

        lg.WriteLine("  exiting Sub ExportPdf")
        lg.WriteLine("")

    End Sub

    Private Sub ExportPdf(ByVal theSheet As Drawings.DrawingSheet, ByVal pdfFile As String)

        lg.WriteLine("Sub ExportPdf(" & theSheet.Name & ", " & pdfFile & ")")

        Dim printPDFBuilder1 As PrintPDFBuilder

        printPDFBuilder1 = _thePart.PlotManager.CreatePrintPdfbuilder()
        printPDFBuilder1.Scale = 1.0
        printPDFBuilder1.Colors = PrintPDFBuilder.Color.BlackOnWhite
        printPDFBuilder1.Size = PrintPDFBuilder.SizeOption.ScaleFactor
        printPDFBuilder1.RasterImages = True
        printPDFBuilder1.ImageResolution = PrintPDFBuilder.ImageResolutionOption.Medium

        If _thePart.PartUnits = BasePart.Units.Inches Then
            lg.WriteLine("  part units: English")
            printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.English
        Else
            lg.WriteLine("  part units: Metric")
            printPDFBuilder1.Units = PrintPDFBuilder.UnitsOption.Metric
        End If

        If _textAsPolylines Then
            lg.WriteLine("  output text as polylines")
            printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Polylines
        Else
            lg.WriteLine("  output text as text")
            printPDFBuilder1.OutputText = PrintPDFBuilder.OutputTextOption.Text
        End If

        lg.WriteLine("  useWatermark: " & _useWatermark.ToString)
        If _useWatermark Then
            printPDFBuilder1.AddWatermark = True
            printPDFBuilder1.Watermark = _watermarkTextFinal
        Else
            printPDFBuilder1.AddWatermark = False
            printPDFBuilder1.Watermark = ""
        End If

        lg.WriteLine("  export to TC? " & _exportToTC.ToString)
        'If _exportToTC Then
        '    'output to dataset
        '    printPDFBuilder1.Relation = PrintPDFBuilder.RelationOption.Manifestation
        '    printPDFBuilder1.DatasetType = "PDF"
        '    printPDFBuilder1.NamedReferenceType = "PDF_Reference"
        '    'printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Overwrite
        '    printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
        '    printPDFBuilder1.DatasetName = _exportFile & "_PDF_1"
        '    lg.WriteLine("  dataset name: " & _exportFile)

        '    Try
        '        lg.WriteLine("  printPDFBuilder1.Assign")
        '        printPDFBuilder1.Assign()
        '    Catch ex As NXException
        '        lg.WriteLine("  ** error with printPDFBuilder1.Assign")
        '        lg.WriteLine("  " & ex.ErrorCode & " : " & ex.Message)
        '    End Try

        'Else
        'output to filesystem
        lg.WriteLine("  pdf file: " & pdfFile)
        printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.Native
        printPDFBuilder1.Append = False
        printPDFBuilder1.Filename = pdfFile

        'End If

        Dim outputSheets(0) As Drawings.DrawingSheet
        outputSheets(0) = theSheet
        printPDFBuilder1.SourceBuilder.SetSheets(outputSheets)

        Dim nXObject1 As NXObject
        Try
            lg.WriteLine("  printPDFBuilder1.Commit")
            nXObject1 = printPDFBuilder1.Commit()

        Catch ex As NXException
            lg.WriteLine("  ** error with printPDFBuilder1.Commit")
            lg.WriteLine("  " & ex.ErrorCode & " : " & ex.Message)

            'If Me.ExportToTc Then

            '    Try
            '        lg.WriteLine("  trying new dataset option")
            '        printPDFBuilder1.Action = PrintPDFBuilder.ActionOption.New
            '        printPDFBuilder1.Commit()
            '    Catch ex2 As NXException
            '        lg.WriteLine("  ** error with printPDFBuilder1.Commit")
            '        lg.WriteLine("  " & ex2.ErrorCode & " : " & ex2.Message)

            '    End Try

            'End If

        Finally
            printPDFBuilder1.Destroy()
        End Try

        lg.WriteLine("  exiting Sub ExportPdf")
        lg.WriteLine("")

    End Sub

    Private Sub DisplayConfirmationDialog()

        Dim sb As New Text.StringBuilder

        If Me._drawingSheets.Count = 0 Then
            MessageBox.Show("No drawing sheets found in file.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        sb.Append("The following sheets were output to PDF:")
        sb.AppendLine()
        For Each tempSheet As Drawings.DrawingSheet In _drawingSheets
            sb.AppendLine("   " & tempSheet.Name)
        Next
        sb.AppendLine()

        If Not Me.ExportToTc Then
            If Me.ExportSheetsIndividually Then
                sb.AppendLine("Open pdf files now?")
            Else
                sb.AppendLine("Open pdf file now?")
            End If
        End If

        Dim prompt As String = sb.ToString

        Dim response As DialogResult
        If Me.ExportToTc Then
            response = MessageBox.Show(prompt, Me.OutputPdfFileName, MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            response = MessageBox.Show(prompt, Me.OutputPdfFileName, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)
        End If

        If response = DialogResult.Yes Then
            Me.OpenPdf = True
        Else
            Me.OpenPdf = False
        End If

    End Sub

    Private Sub StartLog()

        lg.WriteLine("")
        lg.WriteLine("~ NXJournaling.com: Start of drawing to PDF journal ~")
        lg.WriteLine("  ~~ Version: " & Version & " ~~")
        lg.WriteLine("  ~~ Timestamp of run: " & DateTime.Now.ToString & " ~~")
        lg.WriteLine("PdfExporter Sub StartLog()")

        'determine if we are running under TC or native
        _theUfSession.UF.IsUgmanagerActive(_isTCRunning)
        lg.WriteLine("IsTcRunning: " & _isTCRunning.ToString)

        lg.WriteLine("exiting Sub StartLog")
        lg.WriteLine("")


    End Sub

    Private Function DeleteExistingPdfFile(ByVal thePdfFile As String) As Boolean

        lg.WriteLine("Function DeleteExistingPdfFile(" & thePdfFile & ")")

        If File.Exists(thePdfFile) Then
            lg.WriteLine("  specified PDF file already exists")
            If Me.OverwritePdf Then
                Try
                    lg.WriteLine("  user chose to overwrite existing PDF file")
                    File.Delete(thePdfFile)
                    lg.WriteLine("  file deleted")
                    lg.WriteLine("  returning: True")
                    lg.WriteLine("  exiting Function DeleteExistingPdfFile")
                    lg.WriteLine("")
                    Return True
                Catch ex As Exception
                    'rethrow error?
                    lg.WriteLine("  ** error while attempting to delete existing pdf file")
                    lg.WriteLine("  " & ex.GetType.ToString & " : " & ex.Message)
                    lg.WriteLine("  returning: False")
                    lg.WriteLine("  exiting Function DeleteExistingPdfFile")
                    lg.WriteLine("")
                    Return False
                End Try
            Else
                'file exists, overwrite option is set to false - do nothing
                lg.WriteLine("  specified pdf file exists, user chose not to overwrite")
                lg.WriteLine("  returning: False")
                lg.WriteLine("  exiting Function DeleteExistingPdfFile")
                lg.WriteLine("")
                Return False
            End If
        Else
            'file does not exist
            Return True
        End If

    End Function

    Private Function IsSheetEmpty(ByVal theSheet As Drawings.DrawingSheet) As Boolean

        theSheet.Open()
        Dim sheetTag As NXOpen.Tag = theSheet.View.Tag
        Dim sheetObj As NXOpen.Tag = NXOpen.Tag.Null
        _theUfSession.View.CycleObjects(sheetTag, UFView.CycleObjectsEnum.VisibleObjects, sheetObj)
        If (sheetObj = NXOpen.Tag.Null) And (theSheet.GetDraftingViews.Length = 0) Then
            Return True
        End If

        Return False

    End Function

#End Region

End Class
