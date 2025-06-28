Attribute VB_Name = "Module1"
Option Explicit

Sub GenerateGanttChart()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim taskRow As Long, dateCol As Long
    Dim StartDate As Date, endDate As Date
    Dim taskStartCol As Long, taskEndCol As Long
    Dim TaskName As String, TaskNo As String
    Dim taskPeriod As Long, taskPriority As Long
    Dim PrevTasks As String
    Dim Progress As Double
    Dim startDateRange As Range, endDateRange As Range
    Dim progressStartCol As Long, progressEndCol As Long
    Dim progressShape As Shape
    Dim workerNum As Long
    Dim taskList As Collection
    Dim task As task
    Dim i As Long, j As Long
    
    ' シートの初期設定
    Set ws = ThisWorkbook.Sheets("Sheet1") ' シート名を変更してください
    'ws.Cells.ClearFormats
    ws.Shapes.Delete
    
    ' 定数の初期設定
    Const TSK_WORKER_NUM_ROW As Long = 1 ' TSK_WORKER_NUMの行番号
    Const TSK_DATE_START_ROW As Long = 3 ' TSK_DATE_STARTの行番号
    Const TSK_DATE_START_COL As Long = 10 ' TSK_DATE_STARTの列番号（J列）
    Const TSK_NAME_COL As Long = 5 ' TSK_NAMEの列番号（E列）
    Const TSK_NO_COL As Long = 1 ' TSK_NOの列番号（A列）
    Const TSK_PERIOD_COL As Long = 4 ' TSK_PERIODの列番号（D列）
    Const TSK_PRIORITY_COL As Long = 2 ' TSK_PRIORITYの列番号（B列）
    Const TSK_PREV_TSK_COL As Long = 3 ' TSK_PREV_TSKの列番号（C列）
    Const TSK_START_DATE_COL As Long = 16 ' TSK_START_DATEの列番号（P列）
    Const TSK_PROGRESS_COL As Long = 17 ' TSK_PROGRESSの列番号（Q列）
    
    ' 最終行と最終列の取得
    lastRow = ws.Cells(ws.Rows.Count, TSK_NAME_COL).End(xlUp).Row
    lastCol = ws.Cells(TSK_DATE_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    
    ' 実行可能平行タスク数の取得
    workerNum = ws.Cells(TSK_WORKER_NUM_ROW, TSK_DATE_START_COL).Value
    
    ' タスクリストの初期化
    Set taskList = New Collection
    
    ' タスクデータの読み込み
    For taskRow = 4 To lastRow
        If ws.Cells(taskRow, TSK_NAME_COL).Value <> "" Then
            TaskName = ws.Cells(taskRow, TSK_NAME_COL).Value
            TaskNo = ws.Cells(taskRow, TSK_NO_COL).Value
            taskPeriod = ws.Cells(taskRow, TSK_PERIOD_COL).Value
            taskPriority = ws.Cells(taskRow, TSK_PRIORITY_COL).Value
            PrevTasks = ws.Cells(taskRow, TSK_PREV_TSK_COL).Value
            StartDate = ws.Cells(taskRow, TSK_START_DATE_COL).Value
            Progress = ws.Cells(taskRow, TSK_PROGRESS_COL).Value / 100
            
            ' タスクオブジェクトの作成
            Set task = New task
            task.TaskNo = TaskNo
            task.TaskName = TaskName
            task.Period = taskPeriod
            task.Priority = taskPriority
            task.PrevTasks = PrevTasks
            task.StartDate = StartDate
            task.Progress = Progress
            
            ' タスクリストに追加
            taskList.Add task, TaskNo
        End If
    Next taskRow
    
    ' スケジューリング計算
    ScheduleTasks taskList, workerNum
    
    ' ガントチャートの描画
    For taskRow = 4 To lastRow
        TaskNo = ws.Cells(taskRow, TSK_NO_COL).Value
        On Error Resume Next
        Set task = taskList(TaskNo)
        On Error GoTo 0
        
        If Not task Is Nothing Then
            ' 予定の描画
            taskStartCol = GetDateColumn(ws, TSK_DATE_START_ROW, TSK_DATE_START_COL, lastCol, task.ScheduledStartDate)
            taskEndCol = taskStartCol + task.Period - 1
            If taskStartCol >= TSK_DATE_START_COL And taskEndCol <= lastCol Then
                ws.Range(ws.Cells(taskRow, taskStartCol), ws.Cells(taskRow, taskEndCol)).Interior.Color = RGB(200, 200, 200)
            End If
            
            ' 実績の描画
            If task.StartDate <> 0 Then
                progressStartCol = GetDateColumn(ws, TSK_DATE_START_ROW, TSK_DATE_START_COL, lastCol, task.StartDate)
                progressEndCol = progressStartCol + task.Period * task.Progress - 1
                If progressStartCol >= TSK_DATE_START_COL And progressEndCol <= lastCol Then
                    Set progressShape = ws.Shapes.AddLine( _
                        ws.Cells(taskRow, progressStartCol).Left + 5, _
                        ws.Cells(taskRow, progressStartCol).Top + 10, _
                        ws.Cells(taskRow, progressEndCol).Left + 5, _
                        ws.Cells(taskRow, progressEndCol).Top + 10 _
                    )
                    progressShape.Line.ForeColor.RGB = RGB(0, 0, 255)
                    progressShape.Line.Weight = 2
                End If
            End If
        End If
    Next taskRow
    
    MsgBox "ガントチャートの生成が完了しました！", vbInformation
End Sub

Function GetDateColumn(ws As Worksheet, dateRow As Long, startCol As Long, endCol As Long, targetDate As Date) As Long
    Dim col As Long
    For col = startCol To endCol
        If ws.Cells(dateRow, col).Value = targetDate Then
            GetDateColumn = col
            Exit Function
        End If
    Next col
    GetDateColumn = -1
End Function

Sub ScheduleTasks(taskList As Collection, workerNum As Long)
    Dim task As task
    Dim scheduledTasks As Collection
    Dim availableWorkers As Long
    Dim currentWeek As Date
    Dim taskArray() As task
    Dim i As Long, j As Long
    
    ' タスク配列の初期化
    ReDim taskArray(1 To taskList.Count)
    i = 1
    For Each task In taskList
        taskArray(i) = task
        i = i + 1
    Next task
    
    ' タスク配列の優先度順にソート
    For i = LBound(taskArray) To UBound(taskArray) - 1
        For j = i + 1 To UBound(taskArray)
            If taskArray(i).Priority > taskArray(j).Priority Then
                Swap taskArray(i), taskArray(j)
            End If
        Next j
    Next i
    
    ' スケジューリング
    Set scheduledTasks = New Collection
    availableWorkers = workerNum
    currentWeek = Now
    
    For i = LBound(taskArray) To UBound(taskArray)
        If Not IsTaskScheduled(taskArray(i), scheduledTasks) Then
            If CanStartTask(taskArray(i), scheduledTasks) Then
                If availableWorkers > 0 Then
                    taskArray(i).ScheduledStartDate = currentWeek
                    scheduledTasks.Add taskArray(i), taskArray(i).TaskNo
                    availableWorkers = availableWorkers - 1
                Else
                    currentWeek = currentWeek + 7
                    availableWorkers = workerNum - 1
                    taskArray(i).ScheduledStartDate = currentWeek
                    scheduledTasks.Add taskArray(i), taskArray(i).TaskNo
                End If
            End If
        End If
    Next i
End Sub

Function IsTaskScheduled(task As task, scheduledTasks As Collection) As Boolean
    Dim scheduledTask As task
    For Each scheduledTask In scheduledTasks
        If scheduledTask.TaskNo = task.TaskNo Then
            IsTaskScheduled = True
            Exit Function
        End If
    Next scheduledTask
    IsTaskScheduled = False
End Function

Function CanStartTask(task As task, scheduledTasks As Collection) As Boolean
    Dim prevTask As Variant
    Dim prevTaskScheduled As Boolean
    prevTaskScheduled = True
    
    If task.PrevTasks <> "" Then
        For Each prevTask In Split(task.PrevTasks, ",")
            prevTaskScheduled = prevTaskScheduled And IsTaskScheduled(taskList(prevTask), scheduledTasks)
        Next prevTask
    End If
    
    CanStartTask = prevTaskScheduled
End Function

Sub Swap(ByRef task1 As task, ByRef task2 As task)
    Dim tempTask As task
    Set tempTask = task1
    Set task1 = task2
    Set task2 = tempTask
End Sub


Sub RemoveArrows()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Long
    Dim prefix As String
    
    prefix = "ZZZ"
    Set ws = ActiveSheet
    
    If ws.Shapes.Count > 0 Then
        For i = ws.Shapes.Count To 1 Step -1
            Set shp = ws.Shapes(i)
            If Left(shp.Name, Len(prefix)) = prefix Then
                shp.Delete
            End If
        Next i
        MsgBox "所有名称以'" & prefix & "'??的形状已被?除。", vbInformation
    End If
End Sub

