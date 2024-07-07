Attribute VB_Name = "Module1"
Sub Color(RngF As range, RngG As range, Salary As range)
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
        Salary.Value = Salary.Value + RngF.Value * 20
    End If

    If RngG.Interior.Color = TTAssen Then
        Salary.Value = Salary.Value + RngF.Value * 22.5
    End If

    If RngG.Interior.Color = BarMedewerker Then
        Salary.Value = Salary.Value + RngF.Value * 16.5
    End If

    If RngG.Interior.Color = BarRunner Then
        Salary.Value = Salary.Value + RngF.Value * 17.5
    End If

    If RngG.Interior.Color = Barhoofd Then
        Salary.Value = Salary.Value + RngF.Value * 18.5
    End If
End Sub


Sub TotalSalary()
    Dim RngG As range
    Dim RngF As range
    Dim CellF As range
    Dim CellG As range
    Dim B2 As range ' Total Salary
    
    Set RngF = ThisWorkbook.Worksheets("Time sheet").range("F7:F68")
    Set RngG = ThisWorkbook.Worksheets("Time sheet").range("G7:G68")
    Set B2 = ThisWorkbook.Worksheets("Time sheet").range("B2")  ' Target cell where the result will be accumulated
    
    ' Initialize the target cell
    B2.Value = 0
    
    ' Iterate through each cell in RngG and RngF simultaneously
    Dim i As Long
    For i = 1 To RngG.Rows.Count
        Set CellF = RngF.Cells(i, 1)
        Set CellG = RngG.Cells(i, 1)
        Call Color(CellF, CellG, B2)
    Next i
End Sub

