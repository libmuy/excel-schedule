Sub GetLastRowCol(ByRef lastRow As Long, ByRef lastCol As Long, ws As Worksheet)
    lastRow = ws.Cells(ws.Rows.Count, COL_NO).End(xlUp).row
    lastCol = ws.Cells(ROW_START_DATE, ws.Columns.Count).End(xlToLeft).Column
End Sub


Function GetDateColumn(ws As Worksheet, endCol As Long, targetDate As Date) As Long
    Dim col As Long
    Dim prevCol As Long
    Dim prevDate As Date
    prevCol = -1
    prevDate = 0

    For col = COL_START_DATE To endCol
        If ws.Cells(ROW_START_DATE, col).Value = targetDate Then
            GetDateColumn = col
            Exit Function
        ElseIf ws.Cells(ROW_START_DATE, col).Value > targetDate Then
            If prevCol <> -1 Then
                GetDateColumn = prevCol
            Else
                GetDateColumn = -1
            End If
            Exit Function
        End If
        prevCol = col
        prevDate = ws.Cells(ROW_START_DATE, col).Value
    Next col
    ' If targetDate is after the last week, return the last column
    GetDateColumn = endCol
End Function


Function MaxValue(a As Double, b As Double) As Double
    MaxValue = a
    If a < b Then MaxValue = b
End Function

Function MinValue(a As Double, b As Double) As Double
    MinValue = a
    If a > b Then MinValue = b
End Function


Private Sub Swap(ByRef task1 As task, ByRef task2 As task)
    Dim tempTask As task
    Set tempTask = task1
    Set task1 = task2
    Set task2 = tempTask
End Sub


Sub sortTasks(ByRef tasks() As task, byNo As Boolean)
    Dim doSwap As Boolean
    Dim i As Long, j As Long

    ' �^�X�N�z����^�X�NNo���Ƀ\�[�g
    For i = LBound(tasks) To UBound(tasks) - 1
        For j = i + 1 To UBound(tasks)
            If byNo Then
                doSwap = CInt(tasks(i).TaskNo) > CInt(tasks(j).TaskNo)
            Else
                doSwap = CInt(tasks(i).Priority) > CInt(tasks(j).Priority)
            End If
            
            If doSwap Then Swap tasks(i), tasks(j)
        Next j
    Next i
    
End Sub

Sub ClearGanttChart(ws as Worksheet)
    Dim r As Long, c As Long
    
    GetLastRowCol r, c, ws
    ws.Range(ws.Cells(ROW_TSK_START, COL_START_DATE), ws.Cells(r, c)).Interior.ColorIndex = xlNone
    
End Sub

Sub ClearProgressBar(ws as Worksheet)
    Dim shp As Shape
    Dim i As Long

    If ws.Shapes.Count > 0 Then
        For i = ws.Shapes.Count To 1 Step -1
            Set shp = ws.Shapes(i)
            If Left(shp.Name, Len(PROGRESS_BAR_PREFIX)) = PROGRESS_BAR_PREFIX Then
                shp.Delete
            End If
        Next i
    End If
End Sub

Sub ClearContent()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ClearGanttChart ws
    ClearProgressBar ws
End Sub

Sub test1()
    Debug.Print GetRedmineIssueProgress("42671", 1)
End Sub



