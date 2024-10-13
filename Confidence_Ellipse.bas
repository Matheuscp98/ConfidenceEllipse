Attribute VB_Name = "ConfidenceEllipse"
Private Sub UserForm_Initialize()
    ' Set the range of the scroll bar
    ScrollBar1.Min = -99
    ScrollBar1.Max = 99
    ' Set the increment size of the scroll bar
    ScrollBar1.SmallChange = 1
    ScrollBar1.LargeChange = 10
    ' Set the initial position of the scroll bar
    ScrollBar1.Value = 0
    ' Set the label for displaying data
    lblValor.Caption = "Value: 0"
End Sub

Private Sub ScrollBar1_Change()
    ' Convert the scroll bar value to a decimal value
    Dim valor As Double
    valor = ScrollBar1.Value / 100
    ' Set the value for the data label
    lblValor.Caption = "Value: " & Format(valor, "0.00")
    ' Set the value of cell D200 with 2 decimal places
    Range("D200").Value = Format(valor, "0.00")
End Sub

Private Sub cmdatualizar1_Click()
    Dim valor As Double
    valor = ScrollBar1.Value / 100

    Range("D200").Value = valor

    lblValor.Caption = "Value: " & Format(valor, "0.00")
    MsgBox "The correlation has been updated to " & Format(valor, "0.00")
    
End Sub

Private Sub cmdcancelar1_Click()

    Unload Me

End Sub

Private Sub cmdgravar1_Click()
    
    ' Check if the value entered by the user is between -1 and 1
    If Not IsNumeric(txtcorr.Value) Then
        MsgBox "Please enter a numeric value.", vbExclamation
        Exit Sub
    ElseIf txtcorr.Value < -1 Or txtcorr.Value > 1 Then
        MsgBox "Please enter a value between -1 and 1.", vbExclamation
        Exit Sub
    End If
    
    Range("D200").Value = txtcorr.Value
    corr01 = txtcorr.Value
    
    Unload Correlacao
    MsgBox ("The correlation has been changed to " & corr01 & "!")
    
End Sub

Private Sub cmdgravar_Click()
    
    ' Check if the values entered by the user are between 0 and 100
    If Not IsNumeric(txtelipse1.Value) Or Not IsNumeric(txtelipse2.Value) Or Not IsNumeric(txtelipse3.Value) Then
        MsgBox "Please enter numeric values.", vbExclamation
        Exit Sub
    ElseIf txtelipse1.Value < 0 Or txtelipse1.Value > 100 Or txtelipse2.Value < 0 Or txtelipse2.Value > 100 Or txtelipse3.Value < 0 Or txtelipse3.Value > 100 Then
        MsgBox "Please enter values between 0 and 100.", vbExclamation
        Exit Sub
    End If
    
    Range("A200").Value = txtelipse1.Value
    Range("A201").Value = txtelipse2.Value
    Range("A202").Value = txtelipse3.Value
    
    elipse01 = txtelipse1.Value
    elipse02 = txtelipse2.Value
    elipse03 = txtelipse3.Value
    
    Unload Elipses
    MsgBox ("The confidence degrees for the ellipses have been changed to " & elipse01 & "%, " & elipse02 & "%, and " & elipse03 & "%!")
    
End Sub

Private Sub cmdgravar1_Click()
If IsNumeric(txtmedia1.Value) And IsNumeric(txtmedia2.Value) Then
    Range("B200").Value = txtmedia1.Value
    Range("B201").Value = txtmedia2.Value

    Unload Medias
    MsgBox ("The chart axes have been updated according to the mean vector!")
Else
    MsgBox ("Please enter numeric values only!")
End If

ActiveSheet.Unprotect

Dim meu_grafico As ChartObject
Set meu_grafico = Worksheets("Dashboard").ChartObjects("Chart 3")
    'Set x-axis
    meu_grafico.Chart.Axes(xlCategory).MinimumScale = txtmedia1.Value - 10
    meu_grafico.Chart.Axes(xlCategory).MaximumScale = txtmedia1.Value + 10
    
    'Set y-axis
    meu_grafico.Chart.Axes(xlValue).MinimumScale = txtmedia2.Value - 20
    meu_grafico.Chart.Axes(xlValue).MaximumScale = txtmedia2.Value + 20
    
    Set meu_grafico = Worksheets("Dashboard").ChartObjects("Chart 6")
    
    'Set x-axis
    meu_grafico.Chart.Axes(xlCategory).MinimumScale = txtmedia1.Value - 20
    meu_grafico.Chart.Axes(xlCategory).MaximumScale = txtmedia1.Value + 20
    
    'Set y-axis
    meu_grafico.Chart.Axes(xlValue).MinimumScale = txtmedia2.Value - 20
    meu_grafico.Chart.Axes(xlValue).MaximumScale = txtmedia2.Value + 20
    
ActiveSheet.Protect
    
End Sub

Private Sub cmdgravar2_Click()
    If IsNumeric(txtvar1.Value) And IsNumeric(txtvar2.Value) Then
        Range("C200").Value = txtvar1.Value
        Range("C201").Value = txtvar2.Value

        var01 = txtvar1.Value
        var02 = txtvar2.Value

        Unload Variancias
        MsgBox ("The variances of sigma1 (a11) and sigma2 (a22) have been changed to " & var01 & " and " & var02 & "!")
    Else
        MsgBox ("Please enter numeric values only!")
    End If
End Sub

Sub TabPoints()

    Application.ScreenUpdating = False
    
    Sheets("Points").Select
    ActiveWindow.SmallScroll Down:=-27
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub TabData()

    Application.ScreenUpdating = False

    Sheets("Data").Select
    ActiveWindow.SmallScroll Down:=-12
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub TabDashboard()

    Application.ScreenUpdating = False

    Sheets("Dashboard").Select
    ActiveWindow.SmallScroll Down:=-27
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub ToggleBetweenGraphs()
    If ActiveSheet.ChartObjects("Graph 3").Visible = True Then
        ActiveSheet.ChartObjects("Graph 3").Visible = False
        ActiveSheet.ChartObjects("Graph 6").Visible = True
    Else
        ActiveSheet.ChartObjects("Graph 3").Visible = True
        ActiveSheet.ChartObjects("Graph 6").Visible = False
    End If
End Sub

Sub BarCorrelation()

    Application.ScreenUpdating = False
    
    BarCorrelation.Show
    
    Application.ScreenUpdating = True
    
End Sub


Sub Correlation()

    Application.ScreenUpdating = False
    
    Correlation.Show
    
    Application.ScreenUpdating = True
    
End Sub


Sub ChangeAxisValues()

Dim min_x As Double, max_x As Double, min_y As Double, max_y As Double

' Checks if a chart is selected
If TypeName(Selection) <> "ChartArea" Then
    MsgBox "Please select a chart before changing its axes.", vbExclamation
    Exit Sub
End If

' Get the minimum and maximum x-axis values from the user
On Error Resume Next
min_x = CDbl(InputBox("Enter the minimum value for the x-axis:", "X-Axis"))
max_x = CDbl(InputBox("Enter the maximum value for the x-axis:", "X-Axis"))
If Err.Number <> 0 Then
    MsgBox "Please enter only numeric values.", vbExclamation
    Err.Clear
    Exit Sub
End If

' Change the x-axis
ActiveChart.Axes(xlCategory).MinimumScale = min_x ' Set the minimum x-axis value to the user input
ActiveChart.Axes(xlCategory).MaximumScale = max_x ' Set the maximum x-axis value to the user input

' Get the minimum and maximum y-axis values from the user
On Error Resume Next
min_y = CDbl(InputBox("Enter the minimum value for the y-axis:", "Y-Axis"))
max_y = CDbl(InputBox("Enter the maximum value for the y-axis:", "Y-Axis"))
If Err.Number <> 0 Then
    MsgBox "Please enter only numeric values.", vbExclamation
    Err.Clear
    Exit Sub
End If

' Change the y-axis
ActiveChart.Axes(xlValue).MinimumScale = min_y ' Set the minimum y-axis value to the user input
ActiveChart.Axes(xlValue).MaximumScale = max_y ' Set the maximum y-axis value to the user input

End Sub


Sub DisableFullScreen()

    Application.DisplayFullScreen = False
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    
End Sub

Sub UnlockEllipse()
'
' Unblock Ellipse Macro
'

    Application.ScreenUpdating = False
    
    Sheets("Cholesky 90%").Select
    Range("E11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("E510:F510").Select
    Range("F510").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("E11:F510").Select
    ActiveWindow.SmallScroll Down:=-9
    Selection.ClearContents
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "=NORM.S.INV(RAND())"
    Range("E11").Select
    Selection.AutoFill Destination:=Range("E11:F11"), Type:=xlFillDefault
    Range("E11:F11").Select
    Selection.AutoFill Destination:=Range("E11:F510")
    Range("E11:F510").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
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
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 315
    ActiveWindow.ScrollRow = 367
    ActiveWindow.ScrollRow = 492
    ActiveWindow.ScrollRow = 458
    ActiveWindow.ScrollRow = 408
    ActiveWindow.ScrollRow = 265
    ActiveWindow.ScrollRow = 221
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("E13").Select
    Sheets("Dashboard").Select
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub


Sub UnlockPoint()
'
' Unblock Points Macro
'

    Application.ScreenUpdating = False
    
    Sheets("Cholesky 90%").Select
    Range("E11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("E510:F510").Select
    Range("F510").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("E11:F510").Select
    ActiveWindow.SmallScroll Down:=-9
    Selection.ClearContents
    Range("E11").Select
    ActiveCell.FormulaR1C1 = "=NORM.S.INV(RAND())"
    Range("E11").Select
    Selection.AutoFill Destination:=Range("E11:F11"), Type:=xlFillDefault
    Range("E11:F11").Select
    Selection.AutoFill Destination:=Range("E11:F510")
    Range("E11:F510").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
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
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 315
    ActiveWindow.ScrollRow = 367
    ActiveWindow.ScrollRow = 492
    ActiveWindow.ScrollRow = 458
    ActiveWindow.ScrollRow = 408
    ActiveWindow.ScrollRow = 265
    ActiveWindow.ScrollRow = 221
    ActiveWindow.ScrollRow = 70
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("E13").Select
    Sheets("Points").Select
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Sub Ellipse()
    
    Application.ScreenUpdating = False
    
    Ellipse.Show
    
    Application.ScreenUpdating = True
    
End Sub


Sub EnableFullScreen()

    Application.DisplayFullScreen = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    
End Sub

Sub Means()

    Application.ScreenUpdating = False
    
    Means.Show
    
    Application.ScreenUpdating = True
    
End Sub

Sub SaveChanges()

    Application.ScreenUpdating = False

    ThisWorkbook.Save

    Application.ScreenUpdating = True

End Sub

Sub Variances()

    Application.ScreenUpdating = False
    
    Variances.Show
    
    Application.ScreenUpdating = True
    
End Sub

Sub LockEllipse()
'
' Lock Macro
'

    Application.ScreenUpdating = False
    
    Sheets("Cholesky 90%").Select
    Range("E11:F11").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.ScrollRow = 490
    ActiveWindow.ScrollRow = 489
    ActiveWindow.ScrollRow = 485
    ActiveWindow.ScrollRow = 473
    ActiveWindow.ScrollRow = 425
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 230
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 1
    Range("E11:F11").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveWindow.ScrollRow = 471
    ActiveWindow.ScrollRow = 446
    ActiveWindow.ScrollRow = 394
    ActiveWindow.ScrollRow = 364
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("Q11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("E11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Q11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("E11").Select
    Sheets("Dashboard").Select
    Range("A1").Select
    
    Application.ScreenUpdating = True
        
End Sub

Sub LockPoints()
'
' Lock Macro
'

    Application.ScreenUpdating = False
    
    Sheets("Cholesky 90%").Select
    Range("E11:F11").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.ScrollRow = 490
    ActiveWindow.ScrollRow = 489
    ActiveWindow.ScrollRow = 485
    ActiveWindow.ScrollRow = 473
    ActiveWindow.ScrollRow = 425
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 230
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 1
    Range("E11:F11").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveWindow.ScrollRow = 471
    ActiveWindow.ScrollRow = 446
    ActiveWindow.ScrollRow = 394
    ActiveWindow.ScrollRow = 364
    ActiveWindow.ScrollRow = 227
    ActiveWindow.ScrollRow = 163
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("Q11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("E11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Q11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("E11").Select
    Sheets("Points").Select
    Range("A1").Select
    
    Application.ScreenUpdating = True
        
End Sub

