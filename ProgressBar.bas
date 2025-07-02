' すべてのタスクの進捗バーを描画するサブルーチン
Public Sub DrawAllProgressBars()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim taskList() As task
    Dim taskRow As Long
    Dim TaskNo As String
    Dim task As task
    Dim i As Long, j As Long
    Dim workerNum As Long
    Dim startX As Double
    Dim endX As Double
    Dim midY As Double
    Dim r As Range
    Dim sh As Shape

    Set ws = ActiveSheet
    GetLastRowCol lastRow, lastCol, ws
    
    UpdateProgressFromRedmine lastRow, ws

    taskList = GetTaskList(ws, lastRow, True)

    ClearProgressBar ws
    
    For taskRow = ROW_TSK_START To lastRow
        TaskNo = ws.Cells(taskRow, COL_NO).Value
        Set task = Nothing
        For i = LBound(taskList) To UBound(taskList)
            If taskList(i).TaskNo = TaskNo Then
                Set task = taskList(i)
                Exit For
            End If
        Next i
        If Not task Is Nothing Then
            Call DrawProgressBar(ws, task, taskRow, lastCol)
        End If
    Next taskRow


    ' タスク配列をタスクNo順にソート
    sortTasks taskList, True
    
    
    For i = LBound(taskList) To UBound(taskList)
        startX = taskList(i).startX
        endX = taskList(i).endX
        Debug.Print "Child :" & taskList(i).TaskNo & ", start:" & startX & ", end:" & endX
    Next i
    
    i = LBound(taskList)
    Do While i <= UBound(taskList)
        If taskList(i).IsParent Then
            Call GetTaskDrawRange(taskList, i, j)
            i = j + 1
        Else
            i = i + 1
        End If
    Loop
    
    
    For i = LBound(taskList) To UBound(taskList)
        startX = taskList(i).startX
        endX = taskList(i).endX
        Debug.Print "Task :" & taskList(i).TaskNo & ", start:" & startX & ", end:" & endX
        
        
        
        ' 親進捗バーを描画
        If taskList(i).IsParent Then
            startX = taskList(i).startX
            endX = taskList(i).endX
            Set r = ws.Cells(ROW_TSK_START + i - 1, COL_NAME)
            midY = r.Top + (r.Height / 2)
            Set sh = ws.Shapes.AddLine(startX, midY, endX, midY)
            sh.Line.ForeColor.RGB = RGB(110, 110, 110)
            sh.Line.BeginArrowheadStyle = msoArrowheadDiamond
            sh.Line.EndArrowheadStyle = msoArrowheadDiamond
            sh.Line.DashStyle = msoLineDash
            sh.Line.Weight = 2
            sh.Name = PROGRESS_BAR_PREFIX & "parent_" & task.TaskNo & "_" & taskRow
            
        End If
        
    Next i

    ' Draw vertical line for current week
    DrawCurrentWeekLine ws, lastCol, lastRow
End Sub

' --- Draw vertical line for current week ---
Sub DrawCurrentWeekLine(ws As Worksheet, lastCol As Long, lastRow As Long)
    Dim currentDate As Date
    Dim col As Long
    Dim cell As Range
    Dim x As Double
    Dim y1 As Double, y2 As Double
    Dim sh As Shape

    currentDate = Date
    ' Use GetDateColumn to find the column for the current week
    col = GetDateColumn(ws, lastCol, currentDate)
    If col = -1 Then Exit Sub

    Set cell = ws.Cells(ROW_TSK_START, col)
    x = cell.Left + cell.Width / 2
    y1 = ws.Cells(ROW_TSK_START, 1).Top
    y2 = ws.Cells(lastRow, 1).Top + ws.Cells(lastRow, 1).Height

    Set sh = ws.Shapes.AddLine(x, y1, x, y2)
    sh.Line.ForeColor.RGB = RGB(255, 0, 0)
    sh.Line.Weight = 2
    sh.Name = PROGRESS_BAR_PREFIX & "current_week"
End Sub

Sub DrawProgressBar(ws As Worksheet, task As task, taskRow As Long, lastCol As Long)
    Dim progressStartCol As Long, progressEndCol As Long
    Dim doneEndCol As Long, notDoneEndCol As Long
    Dim doneShape As Shape, notDoneShape As Shape
    Dim realStartDate As Date
    Dim scheduledStartDate As Date
    Dim period As Long
    Dim progress As Double
    Dim startCell As Range, endCell As Range
    Dim startX As Double, midY As Double
    Dim radio As Double
    Dim weekStartDate As Date
    Dim donePeriod As Double
    Dim notDonePeriod As Double
    Dim doneEndX As Double
    Dim endX As Double

    If task.IsParent Then Exit Sub

    realStartDate = task.StartDate
    scheduledStartDate = task.scheduledStartDate
    period = task.period
    progress = task.progress

    ' --- Progress bar should start at the later of scheduledStartDate and realStartDate ---
    Dim barStartDate As Date
    If realStartDate > scheduledStartDate Then
        barStartDate = realStartDate
    Else
        barStartDate = scheduledStartDate
    End If

    ' 進捗バーの開始列（ガントチャートの週の開始列）
    progressStartCol = GetDateColumn(ws, lastCol, barStartDate)
    If progressStartCol = -1 Then Exit Sub

    ' 進捗バーの終了列（タスク終了日）
    progressEndCol = progressStartCol + period - 1
    If progressEndCol > lastCol Then progressEndCol = lastCol

    ' 週の開始日
    weekStartDate = ws.Cells(ROW_START_DATE, progressStartCol).Value

    Set startCell = ws.Cells(taskRow, progressStartCol)
    Set endCell = ws.Cells(taskRow, progressEndCol)

    radio = (barStartDate - weekStartDate) / 5
    startX = startCell.Left + (startCell.Width * radio)

    ' 進捗バーのY座標（セル中央）
    midY = startCell.Top + (startCell.Height / 2)

    ' 完了期間と未完了期間を計算
    donePeriod = period * progress
    notDonePeriod = period - donePeriod
    
    Dim doneEndCell As Range, notDoneEndCell As Range
    Dim doneRadio As Double
    doneRadio = donePeriod + radio
    Set doneEndCell = Cells(taskRow, progressStartCol + Int(doneRadio))
    doneEndX = doneEndCell.Left + (doneEndCell.Width * (doneRadio - Int(doneRadio)))
    
    Set notDoneEndCell = Cells(taskRow, progressStartCol + period)
    endX = notDoneEndCell.Left + notDoneEndCell.Width * radio


    ' 未完了部分のバーを描画
    Set notDoneShape = ws.Shapes.AddLine(doneEndX, midY, endX, midY)
    notDoneShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    notDoneShape.Line.Weight = 2
    notDoneShape.Name = PROGRESS_BAR_PREFIX & "notdone_" & task.TaskNo & "_" & taskRow
    
    ' 完了部分のバーを描画
    If donePeriod > 0 And progress > 0 Then
        Set doneShape = ws.Shapes.AddLine(startX, midY, doneEndX, midY)
        doneShape.Line.ForeColor.RGB = RGB(0, 0, 255)
        doneShape.Line.Weight = 2
        doneShape.Name = PROGRESS_BAR_PREFIX & "done_" & task.TaskNo & "_" & taskRow
        'doneShape.Line.BeginArrowheadStyle = msoArrowheadOval
        doneShape.Line.EndArrowheadStyle = msoArrowheadOval
        
    End If
    
    task.startX = startX
    task.endX = endX
End Sub

Sub GetTaskDrawRange(list() As task, idx As Long, ByRef endIdx As Long)
    Dim i As Long
    Dim maxIdx As Long
    Dim startX As Double, endX As Double
    
    maxIdx = UBound(list)
    endIdx = idx
    
    If idx >= maxIdx Then Exit Sub
    
    i = idx + 1
    startX = 9999999
    endX = -1
    Do While i <= maxIdx
    
        If list(i).level > list(idx).level Then
            If list(i).IsParent Then
                Call GetTaskDrawRange(list, i, endIdx)
                If endIdx = i Then Exit Sub
                startX = MinValue(startX, list(i).startX)
                endX = MaxValue(endX, list(i).endX)
                list(idx).startX = startX
                list(idx).endX = endX
            Else
                endIdx = i
                startX = MinValue(startX, list(i).startX)
                endX = MaxValue(endX, list(i).endX)
            End If
        Else
            list(idx).startX = startX
            list(idx).endX = endX
            Exit Sub
        End If
        
        i = endIdx + 1
    Loop
End Sub



Public Sub UpdateProgressFromRedmine(lastRow As Long, ws As Worksheet)
    Dim taskRow As Long
    Dim redmineId As String
    Dim idParts() As String
    Dim repoId As Integer
    Dim ticketId As String
    Dim progress As Double

    For taskRow = ROW_TSK_START To lastRow
        redmineId = ws.Cells(taskRow, COL_REDMINE_ID).Value
        
        If redmineId <> "" Then
            idParts = Split(redmineId, ":")
            
            If UBound(idParts) = 1 Then
                repoId = CInt(idParts(0))
                ticketId = idParts(1)
                
                progress = GetRedmineIssueProgress(ticketId, repoId)
                
                If progress >= 0 Then
                    ws.Cells(taskRow, COL_PROGRESS).Value = progress
                End If
            End If
        End If
    Next taskRow
End Sub




