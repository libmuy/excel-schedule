Attribute VB_Name = "Scheduling"

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
    sortTasks taskList, False

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


