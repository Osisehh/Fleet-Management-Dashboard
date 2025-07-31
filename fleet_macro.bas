Attribute VB_Name = "Module1"
Sub ToggleView()
    Dim ws As Worksheet
    Set ws = Worksheets("2nd Dashboard")

    With ws
        If .ChartObjects("MapChart").Visible = True Then
            ' Hide map, show city chart
            .ChartObjects("MapChart").Visible = False
            .ChartObjects("CityChart").Visible = True
            .Shapes("HeadingState").Visible = msoFalse
            .Shapes("HeadingCity").Visible = msoTrue
            .Shapes("ToggleButton").TextFrame2.TextRange.Text = "View Revenue by State"
        Else
            ' Show map, hide city chart
            .ChartObjects("MapChart").Visible = True
            .ChartObjects("CityChart").Visible = False
            .Shapes("HeadingState").Visible = msoTrue
            .Shapes("HeadingCity").Visible = msoFalse
            .Shapes("ToggleButton").TextFrame2.TextRange.Text = "View Revenue by City"
        End If
    End With
End Sub


