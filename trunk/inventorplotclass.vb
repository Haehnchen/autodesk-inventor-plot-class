'$Id$

Imports Inventor

Public Class myPlotter
    Private _UpdatePlotstyles As Boolean = False
    Private _AllColorsAsBlack As Boolean = False
    Private _UpdateIFeatures As Boolean = False
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
    Public WriteOnly Property UpdatePlotstyles() As Boolean
        Set(ByVal value As Boolean)
            _UpdatePlotstyles = value
        End Set
    End Property
    Public WriteOnly Property AllColorsAsBlack() As Boolean
        Set(ByVal value As Boolean)
            _AllColorsAsBlack = value
        End Set
    End Property
    Public WriteOnly Property UpdateIFeatures() As Boolean
        Set(ByVal value As Boolean)
            _UpdateIFeatures = value
        End Set
    End Property
    Public WriteOnly Property FixedPaperSize() As PaperSizeEnum
        Set(ByVal value As PaperSizeEnum)
            _FixedPaperSize = value
        End Set
    End Property
    Public WriteOnly Property SetSystemPrinter() As String
        Set(ByVal value As String)
            _SetSystemPrinter = value
        End Set
    End Property
    Public WriteOnly Property NumberOfCopies() As Integer
        Set(ByVal value As Integer)
            _NumberOfCopies = value
        End Set
    End Property
    Public WriteOnly Property PrintScaleMode() As PrintScaleModeEnum
        Set(ByVal value As PrintScaleModeEnum)
            _PrintScaleMode = value
        End Set
    End Property
    'Public Property SetRotatePlot1() As RotatePlot
    '    Get
    '        Return _SetRotatePlot
    '    End Get
    '    Set(ByVal value As RotatePlot)
    '        _SetRotatePlot = value
    '    End Set
    'End Property
#End Region
    Public Sub New(ByVal ActiveDocument As Inventor.Document)
        _ActiveDocument = ActiveDocument
    End Sub
    Sub UpdateStile()
        'Dim oDrgPrintMgr As PrintManager = oDrgDoc.PrintManager
        Dim inv_document As ApprenticeServerDocument = _ActiveDocument

        Dim oDrgDoc As DrawingDocument = inv_document
        Dim oIDWStyles As Inventor.DrawingStylesManager = oDrgDoc.StylesManager

        For i As Integer = 1 To oIDWStyles.DimensionStyles.Count
            If oIDWStyles.DimensionStyles.Item(i).UpToDate = False Then
                oIDWStyles.DimensionStyles.Item(i).UpdateFromGlobal()
            End If
        Next

    End Sub
    Sub UpdateDocument()
        _ActiveDocument.Update()
    End Sub
    Sub GenerateIProperties()
        Dim inv_document As ApprenticeServerDocument = _ActiveDocument
        Dim oDrgDoc As DrawingDocument = inv_document
        ' Dim oPropSets As PropertySets = oDrgDoc.PropertySets

        For Each opropset As PropertySet In oDrgDoc.PropertySets
            If opropset.Name = "Inventor User Defined Properties" Then
                Dim oUserPropertySet As PropertySet = opropset
                ChangeOrAddProperty("plotuser", System.Environment.UserName, oUserPropertySet)
                ChangeOrAddProperty("plotdate", Now, oUserPropertySet)
                Exit For
            End If
        Next opropset
    End Sub
    Sub plot()
        Try

            If _SetSystemPrinter.ToLower = "defaultprinter" Then _SetSystemPrinter = DefaultPrinterName()

            If SystemPrinterContains(_SetSystemPrinter) = False Then
                MsgBox("Der Drucker '" & _SetSystemPrinter & "' wurde nicht gefunden", MsgBoxStyle.Critical)
                Exit Sub
            End If

            If _ActiveDocument.DocumentType = Inventor.DocumentTypeEnum.kDrawingDocumentObject Then
                Dim inv_document As ApprenticeServerDocument = _ActiveDocument

                Dim oDrgDoc As DrawingDocument = inv_document
                If _UpdateIFeatures = True Then Me.GenerateIProperties()
                If _UpdatePlotstyles = True Then Me.UpdateStile()
                If _UpdateIFeatures = True Or _UpdatePlotstyles = True Then Me.UpdateDocument()

                Try


                    If Me._FixedPaperSize > 0 Then
                        Me.PlotDin()
                    Else
                        Me.plotBestFit()
                    End If
                Catch ex As Exception
                    'lager.DebugLog(ex.Message)
                    MsgBox(ex.Message)
                End Try

            Else
                MsgBox("keine IDW-Zeichnung")
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
            Case Else    ' Andere Werte.
                MsgBox("ungültige Seiten-Orientierung", MsgBoxStyle.Critical)
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
                MsgBox("ungültiges Papierformat", MsgBoxStyle.Critical)
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
    Public Shared Function DefaultPrinterName() As String
        Dim oPS As New System.Drawing.Printing.PrinterSettings

        Try
            DefaultPrinterName = oPS.PrinterName
        Catch ex As System.Exception
            DefaultPrinterName = ""
        Finally
            oPS = Nothing
        End Try
    End Function
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
