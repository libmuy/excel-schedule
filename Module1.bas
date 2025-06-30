Option Explicit

'================================================================================
' 定数
'================================================================================
' 実行可能平行タスク数
Public Const TSK_WORKER_NUM As String = "P2"

' Cell A1's ROW=1,COLUMN=1
' タスクのID
Public Const COL_NO As Long = 2
' タスク優先度、5段階（１～５）、1は最高優先度
Public Const COL_PRIORITY As Long = 3
' 先行タスクを定義する
Public Const COL_PREV_TSK As Long = 4
' タスクを実施する期間（週単位）
Public Const COL_PERIOD As Long = 5
' タスクの名前
Public Const COL_NAME As Long = 6
' 実際にタスクを開始した日付
Public Const COL_REAL_START As Long = 17
' 進捗率（％）
Public Const COL_PROGRESS As Long = 18
' 全体スケジュールの開始日付
Public Const COL_START_DATE As Long = 19
Public Const ROW_START_DATE As Long = 5
' タスクの開始行
Public Const ROW_TSK_START As Long = 6

' Progress bar shape prefix
Public Const PROGRESS_BAR_PREFIX As String = "ProgressBar_"

' 予定塗りつぶし色 RGB(200,200,200)
Public Const SCHEDULE_COLOR As Long = 13158600


'================================================================================
' メインロジック (ここに処理を記述します)
'================================================================================
Sub GenerateGanttChart()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim taskList() As task
    Dim workerNum As Long
    
    ' シートの設定
    Set ws = ActiveSheet
    
    ' 最終行と最終列の取得
    lastRow = ws.Cells(ws.Rows.Count, COL_NO).End(xlUp).Row
    lastCol = ws.Cells(ROW_START_DATE, ws.Columns.Count).End(xlToLeft).Column
    
    ' 作業者数の取得
    workerNum = Range(TSK_WORKER_NUM).Value
    
    ' タスクリストの作成
    taskList = GetTaskList(ws, lastRow, False)
    
    ' スケジューリング処理
    ScheduleTasks taskList, workerNum
    
    ' ガントチャートの描画
    DrawGanttChart ws, taskList, lastRow, lastCol
    
    'MsgBox "ガントチャートの生成が完了しました！", vbInformation
End Sub

Function GetTaskList(ws As Worksheet, lastRow As Long, getScheduledDate As Boolean) As task()
    Dim taskList() As task
    Dim taskRow As Integer
    Dim taskName As String, TaskNo As String
    Dim taskPeriod As Long, taskPriority As Long
    Dim PrevTasks As String
    Dim progress As Double
    Dim StartDate As Date
    Dim tsk As task
    Dim i As Long, j As Long
    Dim currentLevel As Integer, previousLevel As Integer
    Dim taskCount As Long
    Dim lastCol As Long
    
    
    lastCol = ws.Cells(ROW_START_DATE, ws.Columns.Count).End(xlToLeft).Column
    
    ' タスクリストの作成
    taskCount = lastRow - ROW_TSK_START + 1
    ReDim taskList(1 To taskCount)
    previousLevel = 0
    i = 1
    
    ' タスクデータの読み込み
    For taskRow = ROW_TSK_START To lastRow
        Call GetName(ws, taskRow, taskName, currentLevel)
        If taskName = "" Then Exit For
        
        TaskNo = ws.Cells(taskRow, COL_NO).Value
        taskPeriod = ws.Cells(taskRow, COL_PERIOD).Value
        taskPriority = ws.Cells(taskRow, COL_PRIORITY).Value
        PrevTasks = ws.Cells(taskRow, COL_PREV_TSK).Value
        StartDate = ws.Cells(taskRow, COL_REAL_START).Value
        Dim rawProgress As Variant
        rawProgress = ws.Cells(taskRow, COL_PROGRESS).Value
        If IsNumeric(rawProgress) Then
            If rawProgress > 1 Then
                progress = rawProgress / 100
            Else
                progress = rawProgress
            End If
        Else
            progress = 0
        End If
        
        ' タスクオブジェクトの作成
        Set tsk = New task
        tsk.TaskNo = TaskNo
        tsk.taskName = taskName
        tsk.period = taskPeriod
        tsk.Priority = taskPriority
        tsk.PrevTasks = PrevTasks
        tsk.StartDate = StartDate
        tsk.progress = progress
        tsk.IsParent = False
        tsk.level = currentLevel
        
        ' 階層情報の更新
        If currentLevel > previousLevel Then
            taskList(i - 1).IsParent = True
        End If
        
        '塗りつぶしから予定日を探す
        If getScheduledDate Then
            For j = COL_START_DATE To lastCol
                If ws.Cells(taskRow, j).Interior.Color = SCHEDULE_COLOR Then
                    tsk.scheduledStartDate = ws.Cells(ROW_START_DATE, j).Value
                    Exit For
                End If
            Next j
        End If
        
        ' タスクリストに追加
        Set taskList(i) = tsk
        i = i + 1
        previousLevel = currentLevel
    Next taskRow
    
    GetTaskList = taskList
End Function

Sub DrawGanttChart(ws As Worksheet, taskList() As task, lastRow As Long, lastCol As Long)
    Dim taskRow As Long
    Dim TaskNo As String
    Dim task As task
    Dim i As Long
    Dim taskStartCol As Long, taskEndCol As Long

    ' --- Clear Gantt chart area color ---
    ws.Range(ws.Cells(ROW_TSK_START, COL_START_DATE), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone

    ' ガントチャートの描画
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
            ' タスクの描画
            taskStartCol = GetDateColumn(ws, lastCol, task.scheduledStartDate)
            taskEndCol = taskStartCol + task.period - 1
            If taskStartCol >= COL_START_DATE And taskEndCol <= lastCol Then
                ws.Range(ws.Cells(taskRow, taskStartCol), ws.Cells(taskRow, taskEndCol)).Interior.Color = SCHEDULE_COLOR
            End If
        End If
    Next taskRow
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

    ' 完了部分のバーを描画
    If donePeriod > 0 And progress > 0 Then
        Set doneShape = ws.Shapes.AddLine(startX, midY, doneEndX, midY)
        doneShape.Line.ForeColor.RGB = RGB(0, 0, 255)
        doneShape.Line.Weight = 2
        doneShape.Name = PROGRESS_BAR_PREFIX & "done_" & task.TaskNo & "_" & taskRow
        'doneShape.Line.BeginArrowheadStyle = msoArrowheadOval
        doneShape.Line.EndArrowheadStyle = msoArrowheadOval
        
    End If

    ' 未完了部分のバーを描画
    Set notDoneShape = ws.Shapes.AddLine(doneEndX, midY, endX, midY)
    notDoneShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    notDoneShape.Line.Weight = 2
    notDoneShape.Name = PROGRESS_BAR_PREFIX & "notdone_" & task.TaskNo & "_" & taskRow
    
    task.startX = startX
    task.endX = endX
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

Sub ScheduleTasks(taskList() As task, workerNum As Long)
    Dim scheduledTasks() As task
    Dim currentWeek As Date
    Dim i As Long, j As Long
    Dim scheduledCount As Long
    Dim childStart As Date, childEnd As Date
    Dim k As Long, child As task
    Dim parentIdx As Long
    Dim usedWorkers As Long
    Dim earliestStartDate As Date
    Dim candidateWeek As Date

    ' Validate workerNum
    If workerNum <= 0 Then
        MsgBox "作業者数が無効です。1以上の値を設定してください。", vbCritical
        Exit Sub
    End If

    currentWeek = Cells(ROW_START_DATE, COL_START_DATE).Value

    ' タスク配列を優先度順にソート
    For i = LBound(taskList) To UBound(taskList) - 1
        For j = i + 1 To UBound(taskList)
            If taskList(i).Priority > taskList(j).Priority Then
                Swap taskList(i), taskList(j)
            End If
        Next j
    Next i

    ' スケジューリング
    ReDim scheduledTasks(1 To UBound(taskList))
    scheduledCount = 0

    ' まず子タスクのみスケジューリング
    For i = LBound(taskList) To UBound(taskList)
        If Not taskList(i).IsParent Then
            If Not IsTaskScheduled(taskList(i), scheduledTasks, scheduledCount) Then
                If CanStartTask(taskList(i), scheduledTasks, scheduledCount, earliestStartDate) Then
                    ' candidateWeek is the later of currentWeek and earliestStartDate
                    If earliestStartDate > currentWeek Then
                        candidateWeek = earliestStartDate
                    Else
                        candidateWeek = currentWeek
                    End If
                    Do
                        usedWorkers = CountWorkingTasksInWeek(scheduledTasks, scheduledCount, candidateWeek)
                        If usedWorkers < workerNum Then
                            taskList(i).scheduledStartDate = candidateWeek
                            scheduledCount = scheduledCount + 1
                            Set scheduledTasks(scheduledCount) = taskList(i)
                            Exit Do
                        Else
                            candidateWeek = candidateWeek + 7
                        End If
                    Loop
                End If
            End If
        End If
    Next i

    ' 親タスクの期間と開始日を子タスクから決定
    For parentIdx = LBound(taskList) To UBound(taskList)
        If taskList(parentIdx).IsParent Then
            childStart = 0
            childEnd = 0
            ' 子タスクは親タスクの直後に並んでいると仮定
            For k = parentIdx + 1 To UBound(taskList)
                Set child = taskList(k)
                If child.IsParent Then Exit For
                If child.scheduledStartDate <> 0 Then
                    If childStart = 0 Or child.scheduledStartDate < childStart Then
                        childStart = child.scheduledStartDate
                    End If
                    If child.scheduledStartDate + child.period - 1 > childEnd Then
                        childEnd = child.scheduledStartDate + child.period - 1
                    End If
                End If
            Next k
            If childStart <> 0 And childEnd <> 0 Then
                taskList(parentIdx).scheduledStartDate = childStart
                taskList(parentIdx).period = childEnd - childStart + 1
            End If
        End If
    Next parentIdx
End Sub

' 現在の週に作業中のタスク数をカウント
Function CountWorkingTasksInWeek(scheduledTasks() As task, scheduledCount As Long, weekStart As Date) As Long
    Dim i As Long
    Dim cnt As Long
    For i = 1 To scheduledCount
        If scheduledTasks(i).scheduledStartDate <> 0 Then
            If weekStart >= scheduledTasks(i).scheduledStartDate And weekStart < scheduledTasks(i).scheduledStartDate + (scheduledTasks(i).period * 7) Then
                cnt = cnt + 1
            End If
        End If
    Next i
    CountWorkingTasksInWeek = cnt
End Function

Function IsTaskScheduled(task As task, scheduledTasks() As task, scheduledCount As Long) As Boolean
    Dim i As Long
    For i = 1 To scheduledCount
        If scheduledTasks(i).TaskNo = task.TaskNo Then
            IsTaskScheduled = True
            Exit Function
        End If
    Next i
    IsTaskScheduled = False
End Function

' Returns: CanStartTask = True/False, and sets earliestStartDate to the earliest possible start date
Function CanStartTask(task As task, scheduledTasks() As task, scheduledCount As Long, ByRef earliestStartDate As Date) As Boolean
    Dim prevTask As Variant
    Dim prevTaskScheduled As Boolean
    Dim i As Long, j As Long
    Dim maxEndDate As Date
    Dim found As Boolean

    prevTaskScheduled = True
    maxEndDate = 0

    If task.PrevTasks <> "" Then
        For Each prevTask In Split(task.PrevTasks, ",")
            found = False
            For j = 1 To scheduledCount
                If scheduledTasks(j).TaskNo = Trim(prevTask) Then
                    found = True
                    ' Calculate end date of the dependency
                    Dim depStart As Date, depPeriod As Long, depEnd As Date
                    depStart = scheduledTasks(j).scheduledStartDate
                    depPeriod = scheduledTasks(j).period
                    depEnd = depStart + (depPeriod * 7) ' period is in weeks, so *7 for days
                    If depEnd > maxEndDate Then maxEndDate = depEnd
                    Exit For
                End If
            Next j
            If Not found Then
                prevTaskScheduled = False
                Exit For
            End If
        Next prevTask
    End If

    earliestStartDate = maxEndDate
    CanStartTask = prevTaskScheduled
End Function

Sub Swap(ByRef task1 As task, ByRef task2 As task)
    Dim tempTask As task
    Set tempTask = task1
    Set task1 = task2
    Set task2 = tempTask
End Sub

Private Sub ClearProgressBar()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim i As Long
    Dim prefix As String

    prefix = PROGRESS_BAR_PREFIX
    Set ws = ActiveSheet

    If ws.Shapes.Count > 0 Then
        For i = ws.Shapes.Count To 1 Step -1
            Set shp = ws.Shapes(i)
            If Left(shp.Name, Len(prefix)) = prefix Then
                shp.Delete
            End If
        Next i
        'MsgBox "プレフィックス'" & prefix & "'の矢印を削除しました。", vbInformation
    End If
End Sub

Sub GetName(ws As Worksheet, taskRow As Integer, ByRef taskName As String, ByRef taskLevel As Integer)
    Dim i As Long
    Dim n As String
    
    taskName = ""
    
    For i = COL_NAME To COL_NAME + 5
        n = ws.Cells(taskRow, i).Value
        If n <> "" Then
            taskName = n
            taskLevel = i - COL_NAME
        End If
    Next i
End Sub

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
    lastRow = ws.Cells(ws.Rows.Count, COL_NO).End(xlUp).Row
    lastCol = ws.Cells(ROW_START_DATE, ws.Columns.Count).End(xlToLeft).Column
    taskList = GetTaskList(ws, lastRow, True)

    ClearProgressBar
    
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
    For i = LBound(taskList) To UBound(taskList) - 1
        For j = i + 1 To UBound(taskList)
            If CInt(taskList(i).TaskNo) > CInt(taskList(j).TaskNo) Then
                Swap taskList(i), taskList(j)
            End If
        Next j
    Next i
    
    
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

Function MaxValue(a As Double, b As Double) As Double
    MaxValue = a
    If a < b Then MaxValue = b
End Function

Function MinValue(a As Double, b As Double) As Double
    MinValue = a
    If a > b Then MinValue = b
End Function

