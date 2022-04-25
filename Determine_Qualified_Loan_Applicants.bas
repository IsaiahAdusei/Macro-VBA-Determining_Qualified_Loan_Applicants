Attribute VB_Name = "Module1"
Option Explicit

Sub Determine_Qualified_Loan_Applicants()
Attribute Determine_Qualified_Loan_Applicants.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Determine_Qualified_Loan_Applicants Macro
'

' Bolden the heading of the data and apply centre alignment to the data

    Columns("A:A").EntireColumn.AutoFit
    Cells.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Rows("1:1").Select
    Selection.Font.Bold = True
    
    
 ' Removing all rows with empty cells
 
    Dim Rng, CI As Range
    Set Rng = ActiveSheet.Range("A2:L3000")
    For Each CI In Rng
    If CI.Value = "" Then
    CI.EntireRow.Delete
    End If
    Next CI
    
    
'
' Applying a border to the cells
'

'
    Cells.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
'

' Summing the applicant and Co-Applicant income
'

'
    ActiveCell.FormulaR1C1 = "Total_Applicant&CoIncome"
    Range("M2").Select
    Columns("M:M").EntireColumn.AutoFit
    ActiveCell.FormulaR1C1 = "=SUM(RC[-6],RC[-5])"
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M497"), Type:=xlFillDefault
    Range("M2:M497").Select
    Range("N485").Select

'
' Calculating the per dependent income(Total Income/Dependents) with a condition
'

'
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "Per_Dependent_Income(Total Income/Dependents)"
    Range("N2").Select
    Columns("N:N").EntireColumn.AutoFit
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=IF(RC[-10]=0,RC[-1],RC[-1]/RC[-10])"
    Selection.AutoFill Destination:=Range("N2:N497"), Type:=xlFillDefault
    Range("N2:N497").Select
    Range("O495").Select
    ActiveWindow.ScrollRow = 466
    ActiveWindow.ScrollRow = 446
    ActiveWindow.ScrollRow = 403
    ActiveWindow.ScrollRow = 384
    ActiveWindow.ScrollRow = 380
    ActiveWindow.ScrollRow = 369
    ActiveWindow.ScrollRow = 334
    ActiveWindow.ScrollRow = 329
    ActiveWindow.ScrollRow = 327
    ActiveWindow.ScrollRow = 325
    ActiveWindow.ScrollRow = 322
    ActiveWindow.ScrollRow = 315
    ActiveWindow.ScrollRow = 254
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 230
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 211
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1

'
' Calculating Per Dependent Income Over Loan amount
'

'
    ActiveCell.FormulaR1C1 = "Per_Dependent_Income Over Loan_Amount"
    Range("O2").Select
    Columns("O:O").EntireColumn.AutoFit
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/RC[-6]"
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O497"), Type:=xlFillDefault
    Range("O2:O497").Select
    Range("P492").Select
    ActiveWindow.ScrollRow = 451
    ActiveWindow.ScrollRow = 407
    ActiveWindow.ScrollRow = 382
    ActiveWindow.ScrollRow = 353
    ActiveWindow.ScrollRow = 347
    ActiveWindow.ScrollRow = 289
    ActiveWindow.ScrollRow = 244
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 219
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 171
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 129
    ActiveWindow.ScrollRow = 127
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 95
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1

'
' Determine the qualification status
'

'
    ActiveCell.FormulaR1C1 = "Qualification_Status"
    Range("P2").Select
    Columns("P:P").EntireColumn.AutoFit
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-1]>2,RC[-5]>0),""Qualified"",""Not Qualified"")"
    Selection.AutoFill Destination:=Range("P2:P497"), Type:=xlFillDefault
    Range("P2:P497").Select
    ActiveWindow.ScrollRow = 465
    ActiveWindow.ScrollRow = 438
    ActiveWindow.ScrollRow = 390
    ActiveWindow.ScrollRow = 348
    ActiveWindow.ScrollRow = 346
    ActiveWindow.ScrollRow = 341
    ActiveWindow.ScrollRow = 317
    ActiveWindow.ScrollRow = 298
    ActiveWindow.ScrollRow = 277
    ActiveWindow.ScrollRow = 275
    ActiveWindow.ScrollRow = 273
    ActiveWindow.ScrollRow = 272
    ActiveWindow.ScrollRow = 270
    ActiveWindow.ScrollRow = 263
    ActiveWindow.ScrollRow = 236
    ActiveWindow.ScrollRow = 235
    ActiveWindow.ScrollRow = 234
    ActiveWindow.ScrollRow = 231
    ActiveWindow.ScrollRow = 214
    ActiveWindow.ScrollRow = 210
    ActiveWindow.ScrollRow = 207
    ActiveWindow.ScrollRow = 200
    ActiveWindow.ScrollRow = 199
    ActiveWindow.ScrollRow = 198
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 165
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 161
    ActiveWindow.ScrollRow = 159
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 138
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 120
    ActiveWindow.ScrollRow = 119
    ActiveWindow.ScrollRow = 116
    ActiveWindow.ScrollRow = 113
    ActiveWindow.ScrollRow = 112
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 109
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 107
    ActiveWindow.ScrollRow = 105
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 80
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 75
    Range("O91").Select
    ActiveWindow.SmallScroll ToRight:=-3

'
' Colour Coding Qualification Status
'

'
    Columns("P:P").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Not Qualified"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Qualified"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("Q6").Select
End Sub
