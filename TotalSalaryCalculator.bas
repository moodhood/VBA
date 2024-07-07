Attribute VB_Name = "Module1"
Sub Color(RngF As range, RngG As range, C2 As range)
    Dim Loods As Long
    Dim TTAssen As Long
    Dim BarMedewerker As Long
    Dim BarRunner As Long
    Dim Barhoofd As Long

    Loods = RGB(172, 185, 202)        ' Blue-Gray, Text 2, Lighter 60%
    TTAssen = RGB(187, 190, 169)      ' Brown, Accent 5, Lighter 60%
    BarMedewerker = RGB(200, 201, 190) ' Olive Green, Accent 4, Lighter 60%
    BarRunner = RGB(202, 198, 149)    ' Tan, Accent 1, 40% Lighter
    Barhoofd = RGB(174, 170, 170)     ' Light Gray, Background 2, Darker 25%

    If RngG.Interior.Color = Loods Then
        C2.Value = C2.Value + RngF.Value * 20
    End If

    If RngG.Interior.Color = TTAssen Then
        C2.Value = C2.Value + RngF.Value * 22.5
    End If

    If RngG.Interior.Color = BarMedewerker Then
        C2.Value = C2.Value + RngF.Value * 16.5
    End If

    If RngG.Interior.Color = BarRunner Then
        C2.Value = C2.Value + RngF.Value * 17.5
    End If

    If RngG.Interior.Color = Barhoofd Then
        C2.Value = C2.Value + RngF.Value * 18.5
    End If
End Sub

Sub TestColor()
    Dim RngG As range
    Dim RngF As range
    Dim CellF As range
    Dim CellG As range
    Dim C2 As range
    
    Set RngF = ThisWorkbook.Worksheets("Time sheet").range("F7:F68")
    Set RngG = ThisWorkbook.Worksheets("Time sheet").range("G7:G68")
    Set C2 = ThisWorkbook.Worksheets("Time sheet").range("C2")  ' Target cell where the result will be accumulated
    
    ' Initialize the target cell
    C2.Value = 0

    ' Iterate through each cell in RngG and RngF simultaneously
    Dim i As Long
    For i = 1 To RngG.Rows.Count
        Set CellF = RngF.Cells(i, 1) '.Cells(row, column) RngF.Cells(1,1) = F7
        Set CellG = RngG.Cells(i, 1)
        Call Color(CellF, CellG, C2)
    Next i
End Sub
