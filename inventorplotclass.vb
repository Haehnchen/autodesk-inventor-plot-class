'$Id$

Imports Inventor

Public Class myPlotter
    Private _UpdatePlotstyles As Boolean = False
    Private _AllColorsAsBlack As Boolean = False
    Private _UpdateIProperties As Boolean = False
    Private _SetSystemPrinter As String = ""
    Private _FixedPaperSize As PaperSizeEnum = Nothing
    Public RotatePlot As New sRotatePlot
    Private _NumberOfCopies As Integer = 1
    Private _PrintScaleMode As PrintScaleModeEnum = PrintScaleModeEnum.kPrintBestFitScale
    Private _ActiveDocument As Inventor.Document
#Region "sRotatePlot"
    Public Class sRotatePlot
        Private _A0 As Boolean = False
        Private _A1 As Boolean = False
        Private _A2 As Boolean = False
        Private _A3 As Boolean = False
        Private _A4 As Boolean = False
        Public Property A0() As Boolean
            Get
                Return _A0
            End Get
            Set(ByVal value As Boolean)
                _A0 = value
            End Set
        End Property
        Public Property A1() As Boolean
            Get
                Return _A1
            End Get
            Set(ByVal value As Boolean)
                _A1 = value
            End Set
        End Property
        Public Property A2() As Boolean
            Get
                Return _A2
            End Get
            Set(ByVal value As Boolean)
                _A2 = value
            End Set
        End Property
        Public Property A3() As Boolean
            Get
                Return _A3
            End Get
            Set(ByVal value As Boolean)
                _A3 = value
            End Set
        End Property
        Public Property A4() As Boolean
            Get
                Return _A4
            End Get
            Set(ByVal value As Boolean)
                _A4 = value
            End Set
        End Property
    End Class
#End Region
#Region "Properties"
    ''' <summary>
    ''' Check if (Plot)Styles different from StylesManager and Update (default=false)
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property UpdatePlotstyles() As Boolean
        Set(ByVal value As Boolean)
            _UpdatePlotstyles = value
        End Set
    End Property
    ''' <summary>
    ''' Plot in back and white (default=false)
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property AllColorsAsBlack() As Boolean
        Set(ByVal value As Boolean)
            _AllColorsAsBlack = value
        End Set
    End Property
    ''' <summary>
    ''' Set plotuser and plotdate (default=false)
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property UpdateIProperties() As Boolean
        Set(ByVal value As Boolean)
            _UpdateIProperties = value
        End Set
    End Property
    ''' <summary>
    ''' Force plotting on definied Papersize (default=auto)
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property FixedPaperSize() As PaperSizeEnum
        Set(ByVal value As PaperSizeEnum)
            _FixedPaperSize = value
        End Set
    End Property
    ''' <summary>
    ''' Set Printer/Plotter to print on
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property SetSystemPrinter() As String
        Set(ByVal value As String)
            _SetSystemPrinter = value
        End Set
    End Property
    ''' <summary>
    ''' Numbers of Copies to plot
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property NumberOfCopies() As Integer
        Set(ByVal value As Integer)
            _NumberOfCopies = value
        End Set
    End Property
    ''' <summary>
    ''' Scalemode for plotting (default=bestfit)
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property PrintScaleMode() As PrintScaleModeEnum
        Set(ByVal value As PrintScaleModeEnum)
            _PrintScaleMode = value
        End Set
    End Property
#End Region
    Structure strings
        Dim [empty] As String
        Shared PrinterNotFound As String = "Unknown Printer"
        Shared NoIDW As String = "No IDW-File"
        Shared ErrorPageOrientation As String = "Page orientation unknown"
        Shared ErrorPaperSize As String = "Papersize unknown"
    End Structure
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ActiveDocument">should be m_inventorApplication.ActiveDocument</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal ActiveDocument As Inventor.Document)
        _ActiveDocument = ActiveDocument
    End Sub
    ''' <summary>
    ''' Force an update on StylesManger->Drawing
    ''' </summary>
    ''' <remarks></remarks>
    Sub UpdateStyles()
        Dim inv_document As ApprenticeServerDocument = _ActiveDocument

        Dim oDrgDoc As DrawingDocument = inv_document
        Dim oIDWStyles As Inventor.DrawingStylesManager = oDrgDoc.StylesManager

        For i As Integer = 1 To oIDWStyles.DimensionStyles.Count
            If oIDWStyles.DimensionStyles.Item(i).UpToDate = False Then
                oIDWStyles.DimensionStyles.Item(i).UpdateFromGlobal()
            End If
        Next

    End Sub
    ''' <summary>
    ''' Force update/rebuild of Drawing; need after some changes
    ''' </summary>
    ''' <remarks></remarks>
    Sub UpdateDocument()
        _ActiveDocument.Update()
    End Sub
    ''' <summary>
    ''' Force generating or updating the IProperties: plotuser, plotdate
    ''' </summary>
    ''' <remarks></remarks>
    Sub GenerateIProperties()
        Dim inv_document As ApprenticeServerDocument = _ActiveDocument
        Dim oDrgDoc As DrawingDocument = inv_document

        For Each opropset As PropertySet In oDrgDoc.PropertySets
            If opropset.Name = "Inventor User Defined Properties" Then
                Dim oUserPropertySet As PropertySet = opropset
                ChangeOrAddProperty("plotuser", System.Environment.UserName, oUserPropertySet)
                ChangeOrAddProperty("plotdate", Now, oUserPropertySet)
                Exit For
            End If
        Next opropset
    End Sub
    ''' <summary>
    ''' OK all done, plotting...
    ''' </summary>
    ''' <remarks></remarks>
    Sub plot()
        Try
            If _SetSystemPrinter.ToLower = "defaultprinter" Then _SetSystemPrinter = DefaultPrinterName()

            If SystemPrinterContains(_SetSystemPrinter) = False Then
                MsgBox(myPlotter.strings.PrinterNotFound & _SetSystemPrinter, MsgBoxStyle.Critical)
                Exit Sub
            End If

            If _ActiveDocument.DocumentType = Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                Dim inv_document As ApprenticeServerDocument = _ActiveDocument

                Dim oDrgDoc As DrawingDocument = inv_document
                If _UpdateIProperties = True Then Me.GenerateIProperties()
                If _UpdatePlotstyles = True Then Me.UpdateStyles()
                If _UpdateIProperties = True Or _UpdatePlotstyles = True Then Me.UpdateDocument()

                Try
                    If Me._FixedPaperSize > 0 Then
                        Me.PlotDin()
                    Else
                        Me.plotBestFit()
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

            Else
                MsgBox(myPlotter.strings.NoIDW)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub setOrientation(ByVal tt As Sheet, ByVal oDrgPrintMgr As PrintManager)
        Select Case tt.Orientation
            Case PageOrientationTypeEnum.kLandscapePageOrientation
                oDrgPrintMgr.Orientation = PrintOrientationEnum.kLandscapeOrientation
                'MsgBox("Querformat")
            Case PageOrientationTypeEnum.kPortraitPageOrientation
                oDrgPrintMgr.Orientation = PrintOrientationEnum.kPortraitOrientation
                'MsgBox("Hochformat")
            Case Else    ' other values.
                MsgBox(myPlotter.strings.ErrorPageOrientation, MsgBoxStyle.Critical)
                Exit Sub
        End Select
    End Sub
    Private Sub plotBestFit()

        Dim inv_document As ApprenticeServerDocument = _ActiveDocument

        Dim oDrgDoc As DrawingDocument = inv_document
        Dim oDrgPrintMgr As PrintManager = oDrgDoc.PrintManager

        oDrgPrintMgr.Printer = Me._SetSystemPrinter

        Dim tt As Sheet = oDrgDoc.ActiveSheet
        setOrientation(tt, oDrgPrintMgr)

        Select Case tt.Size
            Case DrawingSheetSizeEnum.kA4DrawingSheetSize
                'MsgBox("A4")
                oDrgPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeA4
                If Me.RotatePlot.A4 = True Then oDrgPrintMgr.Rotate90Degrees = True
            Case DrawingSheetSizeEnum.kA3DrawingSheetSize
                'MsgBox("A3")
                oDrgPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeA3
                If Me.RotatePlot.A3 = True Then oDrgPrintMgr.Rotate90Degrees = True
            Case DrawingSheetSizeEnum.kA2DrawingSheetSize
                'MsgBox("A2")
                oDrgPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeA2
                If Me.RotatePlot.A2 = True Then oDrgPrintMgr.Rotate90Degrees = True
            Case DrawingSheetSizeEnum.kA1DrawingSheetSize
                'MsgBox("A1")
                oDrgPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeA1
                If Me.RotatePlot.A1 = True Then oDrgPrintMgr.Rotate90Degrees = True
            Case DrawingSheetSizeEnum.kA0DrawingSheetSize
                ' MsgBox("A0")
                oDrgPrintMgr.PaperSize = PaperSizeEnum.kPaperSizeA0
                If Me.RotatePlot.A0 = True Then oDrgPrintMgr.Rotate90Degrees = True
            Case Else    ' Andere Werte.
                MsgBox(myPlotter.strings.ErrorPaperSize, MsgBoxStyle.Critical)
                Exit Sub
        End Select

        Plotting(oDrgPrintMgr)

    End Sub
    Private Sub PlotDin()
        Dim oDrgDoc As DrawingDocument = _ActiveDocument
        Dim tt As Sheet = oDrgDoc.ActiveSheet
        Dim oDrgPrintMgr As PrintManager = oDrgDoc.PrintManager

        oDrgPrintMgr.Printer = Me._SetSystemPrinter

        setOrientation(tt, oDrgPrintMgr)

        oDrgPrintMgr.PaperSize = Me._FixedPaperSize

        Plotting(oDrgPrintMgr)

    End Sub
    Private Sub Plotting(ByVal oDrgPrintMgr As PrintManager)
        oDrgPrintMgr.AllColorsAsBlack = Me._AllColorsAsBlack
        If Me._NumberOfCopies > 1 Then oDrgPrintMgr.NumberOfCopies = 1
        oDrgPrintMgr.ScaleMode = _PrintScaleMode
        oDrgPrintMgr.SubmitPrint()
    End Sub
    Private Function SystemPrinterContains(ByVal drucker As String) As Boolean
        For i As Integer = 0 To System.Drawing.Printing.PrinterSettings.InstalledPrinters.Count - 1
            If System.Drawing.Printing.PrinterSettings.InstalledPrinters.Item(i).ToLower = drucker.ToLower Then Return True
        Next
        Return False
    End Function
    Private Function DefaultPrinterName() As String
        Dim oPS As New System.Drawing.Printing.PrinterSettings

        Try
            DefaultPrinterName = oPS.PrinterName
        Catch ex As System.Exception
            DefaultPrinterName = ""
        Finally
            oPS = Nothing
        End Try

    End Function
#Region "TranslatorFunctions"


    ''' <summary>
    ''' Use the Inventor translator Add-In to generate a PDF without any further tools
    ''' 
    ''' http://modthemachine.typepad.com/my_weblog/2009/01/translating-files-with-the-api.html
    ''' </summary>
    ''' <param name="ThisApplication">Inventor Application object; a Drawing must be the current Document!</param>
    ''' <param name="OutFile">FullName with complete Path for PDF-File</param>
    ''' <remarks></remarks>
    Shared Sub SaveAsPDF(ByVal ThisApplication As Inventor.Application, ByVal OutFile As String, Optional ByVal AllColorASBlack As Boolean = True)
        ' Get the PDF translator Add-In. 
        Dim oPDFTrans As TranslatorAddIn
        oPDFTrans = ThisApplication.ApplicationAddIns.ItemById( _
                                "{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
        If oPDFTrans Is Nothing Then
            MsgBox("Could not access PDF translator.")
            Exit Sub
        End If

        ' Create some objects that are used to pass information to  the translator Add-In.  
        Dim oContext As TranslationContext
        oContext = ThisApplication.TransientObjects.CreateTranslationContext
        Dim oOptions As NameValueMap
        oOptions = ThisApplication.TransientObjects.CreateNameValueMap
        If oPDFTrans.HasSaveCopyAsOptions(ThisApplication.ActiveDocument, _
                                                  oContext, oOptions) Then
            ' Set to print all sheets.  This can also have the value 
            ' kPrintCurrentSheet or kPrintSheetRange. If kPrintSheetRange 
            ' is used then you must also use the CustomBeginSheet and 
            ' Custom_End_Sheet to define the sheet range. 
            oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets

            ' Other possible options... 
            'oOptions.Value("Custom_Begin_Sheet") = 1 
            'oOptions.Value("Custom_End_Sheet") = 5 
            oOptions.Value("All_Color_AS_Black") = True
            'oOptions.Value("Remove_Line_Weights") = True 
            'oOptions.Value("Vector_Resolution") = 200

            ' Define various settings and input to provide the translator. 
            oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism
            Dim oData As DataMedium
            oData = ThisApplication.TransientObjects.CreateDataMedium
            oData.FileName = OutFile

            ' Call the translator. 
            Call oPDFTrans.SaveCopyAs(ThisApplication.ActiveDocument, _
                                              oContext, oOptions, oData)
        End If
    End Sub
    ''' <summary>
    ''' Save File as DWF
    ''' 
    ''' Autodesk: Discussion Groups - Publish DWF without partlists and/or revision table
    ''' 
    ''' http://discussion.autodesk.com/forums/message.jspa?messageID=6183760
    ''' </summary>
    ''' <param name="ThisApplication"></param>
    ''' <param name="OutFile">FileName</param>
    ''' <param name="Launch_Viewer">Start DWF View after saving</param>
    ''' <param name="Publish_3D_Models"></param>
    ''' <param name="Publish_All_Sheets"></param>
    ''' <param name="Publish_Mode"></param>
    ''' <param name="Enable_Printing"></param>
    ''' <remarks></remarks>
    Shared Sub SaveAsDWF(ByVal ThisApplication As Inventor.Application, ByVal OutFile As String, Optional ByVal Launch_Viewer As Boolean = False, Optional ByVal Publish_3D_Models As Boolean = False, Optional ByVal Publish_All_Sheets As Boolean = True, Optional ByVal Publish_Mode As Inventor.DWFPublishModeEnum = Inventor.DWFPublishModeEnum.kCustomDWFPublish, Optional ByVal Enable_Printing As Boolean = True)

        ' Get the DWF translator Add-In.
        Dim DWFAddIn As TranslatorAddIn
        DWFAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}")

        ' Set a reference to the active document (the document to be published).
        Dim oDocument As Document
        oDocument = ThisApplication.ActiveDocument

        Dim oContext As TranslationContext
        oContext = ThisApplication.TransientObjects.CreateTranslationContext
        oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = ThisApplication.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = ThisApplication.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If DWFAddIn.HasSaveCopyAsOptions(oDataMedium, oContext, oOptions) Then

            oOptions.Value("Launch_Viewer") = Launch_Viewer
            oOptions.Value("Enable_Printing") = Enable_Printing

            If TypeOf oDocument Is DrawingDocument Then

                ' Drawing options
                oOptions.Value("Publish_Mode") = Publish_Mode
                oOptions.Value("Publish_All_Sheets") = Publish_All_Sheets
                oOptions.Value("Publish_3D_Models") = Publish_3D_Models

            End If

        End If

        oDataMedium.FileName = OutFile

        'Publish document.
        Call DWFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)

    End Sub
#End Region
#Region "IPropertiesHelperFunctions"
    Private Function chkPropInt(ByVal PropName As String, ByVal oUserPropertySet As PropertySet) As Integer
        For i As Integer = 1 To oUserPropertySet.Count
            If oUserPropertySet.Item(i).Name = PropName Then Return i
        Next i
        Return -1
    End Function
    Private Sub ChangeOrAddProperty(ByVal PropName As String, ByVal value As Object, ByVal oUserPropertySet As PropertySet)
        If Me.chkProp(PropName, oUserPropertySet) = True Then
            oUserPropertySet(Me.chkPropInt(PropName, oUserPropertySet)).Value = value
        Else
            oUserPropertySet.Add(value, PropName)
        End If
    End Sub
    Private Function chkProp(ByVal PropName As String, ByVal oUserPropertySet As PropertySet) As Boolean
        For i As Integer = 1 To oUserPropertySet.Count
            If oUserPropertySet.Item(i).Name = PropName Then Return True
        Next i
        Return False
    End Function
#End Region

End Class
