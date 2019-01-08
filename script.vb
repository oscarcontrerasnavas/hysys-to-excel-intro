Private Sub BtnRun_Click()
    ' HYSYS main objects
    Dim hyApp As HYSYS.Application
    Dim hyCase As HYSYS.SimulationCase
    Set hyApp = CreateObject("HYSYS.Application")
    Set hyCase = hyApp.ActiveDocument
    
    ' Check if the hyCase is open or it is neccesary to find it in the path
    ' specified in the Sheets("SetUp").Range("B4")
    If hyCase Is Nothing Then
        Dim hyPath As String
        hyPath = Sheets("SetUp").Range("B4").Value2
        If hyPath = "FALSE" Or hyPath = "" Then
            MsgBox ("The Cell B4 is empty.")
        Else
            Set hyCase = GetObject(hyPath)
        End If
    End If
    
    ' HYSYS child objects
    Dim hyStreams As HYSYS.Streams
    Dim stream As HYSYS.ProcessStream
    Dim hyComponents As HYSYS.Components
    Dim component As HYSYS.component
    
    'Read and write properties
    Set hyStreams = hyCase.Flowsheet.MaterialStreams
    Dim i As Integer
    Dim j As Integer
    Dim startRow As Integer
    Dim value As Variant
    i = 0
    For Each stream In hyStreams
        
        ' Some properties, you can add more but be careful with the cell numbers
        Cells(4, 2 + i) = stream.Name
        Cells(5, 2 + i) = stream.VapourFractionValue
        Cells(6, 2 + i) = stream.TemperatureValue
        Cells(7, 2 + i) = stream.PressureValue
        Cells(8, 2 + i) = stream.MolarFlowValue
        Cells(9, 2 + i) = stream.MassFlowValue
        Cells(10, 2 + i) = stream.IdealLiquidVolumeFlowValue
        Cells(11, 2 + i) = stream.MolarEnthalpyValue
        Cells(12, 2 + i) = stream.MolarEntropyValue
        Cells(13, 2 + i) = stream.HeatFlowValue
        Cells(14, 2 + i) = stream.StdLiqVolFlowValue
        
        ' Extract MolarComposition of Components
        ' The startRow is one after the last Row where the properties has been written into
        startRow = 15
        j = 0
        For Each value In stream.ComponentMolarFractionValue
            Cells(startRow + j, 2 + i) = value
            j = j + 1
        Next value
        
        i = i + 1
    Next stream
    
    ' Read and Write components names
    Set hyComponents = hyCase.BasisManager.ComponentLists.Item(0).Components
    j = 0
    For Each component In hyComponents
        Cells(startRow + j, 1) = "Molar Fraction " & component.Name
        j = j + 1
    Next component
        
        
    ' Sort
    Dim header As Range
    Dim group As Range
    Set header = Range(Cells(4, 2), Cells(4, 2 + i - 1))
    Set group = Range(Cells(4, 2), Cells(14 + j, 2 + i - 1))
    Call Module1.sortStreams(header, group)
End Sub