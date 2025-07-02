Sub GenerateGanttChart()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim taskList() As task
    Dim workerNum As Long
    
    ' シートの設定
    Set ws = ActiveSheet
    
    ' 最終行と最終列の取得
    GetLastRowCol lastRow, lastCol, ws
    
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

' Draw Gantt chart using Redmine start/end dates if ticket id is filled
Sub DrawRedmineGanttChart(ws As Worksheet, lastRow As Long, lastCol As Long)
    Dim taskRow As Long
    Dim redmineId As String
    Dim idParts() As String
    Dim repoId As Integer
    Dim ticketId As String
    Dim startDate As Date, endDate As Date
    Dim startCol As Long, endCol As Long

    ClearGanttChart ws

    For taskRow = ROW_TSK_START To lastRow
        redmineId = ws.Cells(taskRow, COL_REDMINE_ID).Value
        If redmineId <> "" Then
            idParts = Split(redmineId, ":")
            If UBound(idParts) = 1 Then
                repoId = CInt(idParts(0))
                ticketId = idParts(1)
                If GetRedmineIssueStartEndDate(ticketId, repoId, startDate, endDate) Then
                    startCol = GetDateColumn(ws, lastCol, startDate)
                    endCol = GetDateColumn(ws, lastCol, endDate)
                    If startCol >= COL_START_DATE And endCol <= lastCol And startCol <= endCol Then
                        ws.Range(ws.Cells(taskRow, startCol), ws.Cells(taskRow, endCol)).Interior.Color = SCHEDULE_COLOR
                    End If
                End If
            End If
        End If
    Next taskRow
End Sub
