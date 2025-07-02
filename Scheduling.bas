
Sub ScheduleTasks(taskList() As task, workerNum As Long)
    Dim scheduledTasks() As task
    Dim unscheduledTasks() As task
    Dim currentWeek As Date
    Dim i As Long, j As Long, k As Long
    Dim scheduledCount As Long
    Dim unscheduledCount As Long
    Dim childStart As Date, childEnd As Date
    Dim parentIdx As Long
    Dim usedWorkers As Long
    Dim earliestStartDate As Date
    Dim tasksScheduledThisWeek As Boolean
    Dim allTasksScheduled As Boolean
    Dim tempUnscheduled() As task
    Dim tempCount As Long

    ' Validate workerNum
    If workerNum <= 0 Then
        MsgBox "作業者数が無効です。1以上の値を設定してください。", vbCritical
        Exit Sub
    End If

    currentWeek = Cells(ROW_START_DATE, COL_START_DATE).Value

    ' タスク配列を優先度順にソート
    sortTasks taskList, False

    ' スケジューリング
    ReDim scheduledTasks(1 To UBound(taskList))
    scheduledCount = 0
    ReDim unscheduledTasks(1 To UBound(taskList))
    unscheduledCount = 0

    ' 親タスクと子タスクを分離
    For i = LBound(taskList) To UBound(taskList)
        If Not taskList(i).IsParent Then
            unscheduledCount = unscheduledCount + 1
            Set unscheduledTasks(unscheduledCount) = taskList(i)
        End If
    Next i
    
    ' 週ごとにスケジューリング
    Do While unscheduledCount > 0
        usedWorkers = CountWorkingTasksInWeek(scheduledTasks, scheduledCount, currentWeek)
        
        tempCount = 0
        ReDim tempUnscheduled(1 To unscheduledCount)

        For i = 1 To unscheduledCount
            Dim taskToSchedule As task
            Set taskToSchedule = unscheduledTasks(i)

            If usedWorkers < workerNum Then
                If CanStartTask(taskToSchedule, scheduledTasks, scheduledCount, earliestStartDate) Then
                    If earliestStartDate <= currentWeek Then
                        taskToSchedule.scheduledStartDate = currentWeek
                        scheduledCount = scheduledCount + 1
                        Set scheduledTasks(scheduledCount) = taskToSchedule
                        usedWorkers = usedWorkers + 1
                    Else
                        ' 依存関係がまだ満たされていないので、次の週に持ち越す
                        tempCount = tempCount + 1
                        Set tempUnscheduled(tempCount) = taskToSchedule
                    End If
                Else
                    ' 依存関係がまだ満たされていないので、次の週に持ち越す
                    tempCount = tempCount + 1
                    Set tempUnscheduled(tempCount) = taskToSchedule
                End If
            Else
                ' 今週はワーカーがいないので、次の週に持ち越す
                tempCount = tempCount + 1
                Set tempUnscheduled(tempCount) = taskToSchedule
            End If
        Next i
        
        unscheduledCount = tempCount
        If unscheduledCount > 0 Then
            ReDim Preserve tempUnscheduled(1 To unscheduledCount)
            unscheduledTasks = tempUnscheduled
        End If

        currentWeek = currentWeek + 7
    Loop

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
            If weekStart >= scheduledTasks(i).scheduledStartDate And weekStart < (scheduledTasks(i).scheduledStartDate + (scheduledTasks(i).period * 7)) Then
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



