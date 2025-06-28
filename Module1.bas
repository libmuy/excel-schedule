Attribute VB_Name = "Module1"
Option Explicit

'================================================================================
' 定数
'================================================================================
' 実行可能平行タスク数
Public Const TSK_WORKER_NUM As String = "P2"
' タスクのID
Public Const COL_NO As Long = 1
' タスク優先度、5段階（１～５）、1は最高優先度
Public Const COL_PRIORITY As Long = 2
' 先行タスクを定義する
Public Const COL_PREV_TSK As Long = 3
' タスクを実施する期間（週単位）
Public Const COL_PERIOD As Long = 4
' タスクの名前
Public Const COL_NAME As Long = 5
' 実際にタスクを開始した日付
Public Const COL_REAL_START As Long = 16
' 進捗率（％）
Public Const COL_PROGRESS As Long = 17
' 全体スケジュールの開始日付
Public Const COL_START_DATE As Long = 18
Public Const ROW_START_DATE As Long = 5
' タスクの開始行
Public Const ROW_TSK_START As Long = 6

'================================================================================
' メインロジック (ここに処理を記述します)
'================================================================================
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
    
    ' シートの設定
    Set ws = ActiveSheet
    RemoveArrows
    
    ' 最終行と最終列の取得
    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row
    lastCol = ws.Cells(ROW_START_DATE, ws.Columns.Count).End(xlToLeft).Column
    
    ' 作業者数の取得
    workerNum = ws.Cells(ROW_START_DATE, COL_REAL_START).Value
    
    ' タスクリストの作成
    Set taskList = New Collection
    
    ' タスクデータの読み込み
    For taskRow = ROW_TSK_START To lastRow
        If ws.Cells(taskRow, COL_NAME).Value <> "" Then
            TaskName = ws.Cells(taskRow, COL_NAME).Value
            TaskNo = ws.Cells(taskRow, COL_NO).Value
            taskPeriod = ws.Cells(taskRow, COL_PERIOD).Value
            taskPriority = ws.Cells(taskRow, COL_PRIORITY).Value
            PrevTasks = ws.Cells(taskRow, COL_PREV_TSK).Value
            StartDate = ws.Cells(taskRow, COL_REAL_START).Value
            Progress = ws.Cells(taskRow, COL_PROGRESS).Value / 100
            
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
    
    ' スケジューリング処理
    ScheduleTasks taskList, workerNum
    
    ' ガントチャートの描画
    For taskRow = 4 To lastRow
        TaskNo = ws.Cells(taskRow, COL_NO).Value
        On Error Resume Next
        Set task = taskList(TaskNo)
        On Error GoTo 0
        
        If Not task Is Nothing Then
            ' タスクの描画
            taskStartCol = GetDateColumn(ws, ROW_START_DATE, COL_REAL_START, lastCol, task.ScheduledStartDate)
            taskEndCol = taskStartCol + task.Period - 1
            If taskStartCol >= COL_REAL_START And taskEndCol <= lastCol Then
                ws.Range(ws.Cells(taskRow, taskStartCol), ws.Cells(taskRow, taskEndCol)).Interior.Color = RGB(200, 200, 200)
            End If
            
            ' 進捗の描画
            If task.StartDate <> 0 Then
                progressStartCol = GetDateColumn(ws, ROW_START_DATE, COL_REAL_START, lastCol, task.StartDate)
                progressEndCol = progressStartCol + task.Period * task.Progress - 1
                If progressStartCol >= COL_REAL_START And progressEndCol <= lastCol Then
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
    
    ' タスク配列の作成
    ReDim taskArray(1 To taskList.Count)
    i = 1
    For Each task In taskList
        taskArray(i) = task
        i = i + 1
    Next task
    
    ' タスク配列を優先度順にソート
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
        MsgBox "プレフィックス'" & prefix & "'の矢印を削除しました。", vbInformation
    End If
End Sub

