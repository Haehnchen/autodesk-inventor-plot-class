'$Id$

Dim beispiel As New InventorPlotClass(Inventor.Document.ActiveDocument)
beispiel.RotatePlot.A0 = True
beispiel.AllColorsAsBlack = True
beispiel.NumberOfCopies = 2
beispiel.SetSystemPrinter = "PDFCreator"
beispiel.UpdateIFeatures = True
beispiel.UpdatePlotstyles = True
beispiel.plot()