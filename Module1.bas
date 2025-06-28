Attribute VB_Name = "Module1"
Option Explicit

'================================================================================
' 定数
'================================================================================

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
Public Const COL_REAL_START As Long = 6

' 進捗率（％）
Public Const COL_PROGRESS As Long = 7

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
    
    ' �V�[�g�̏����ݒ�
    Set ws = ThisWorkbook.Sheets("Sheet1") ' �V�[�g����ύX���Ă�������
    'ws.Cells.ClearFormats
    ws.Shapes.Delete
    
    ' �萔�̏����ݒ�
    Const TSK_WORKER_NUM_ROW As Long = 1 ' TSK_WORKER_NUM�̍s�ԍ�
    Const TSK_DATE_START_ROW As Long = 3 ' TSK_DATE_START�̍s�ԍ�
    Const TSK_DATE_START_COL As Long = 10 ' TSK_DATE_START�̗�ԍ��iJ��j
    Const TSK_NAME_COL As Long = 5 ' TSK_NAME�̗�ԍ��iE��j
    Const TSK_NO_COL As Long = 1 ' TSK_NO�̗�ԍ��iA��j
    Const TSK_PERIOD_COL As Long = 4 ' TSK_PERIOD�̗�ԍ��iD��j
    Const TSK_PRIORITY_COL As Long = 2 ' TSK_PRIORITY�̗�ԍ��iB��j
    Const TSK_PREV_TSK_COL As Long = 3 ' TSK_PREV_TSK�̗�ԍ��iC��j
    Const TSK_START_DATE_COL As Long = 16 ' TSK_START_DATE�̗�ԍ��iP��j
    Const TSK_PROGRESS_COL As Long = 17 ' TSK_PROGRESS�̗�ԍ��iQ��j
    
    ' �ŏI�s�ƍŏI��̎擾
    lastRow = ws.Cells(ws.Rows.Count, TSK_NAME_COL).End(xlUp).Row
    lastCol = ws.Cells(TSK_DATE_START_ROW, ws.Columns.Count).End(xlToLeft).Column
    
    ' ���s�\���s�^�X�N���̎擾
    workerNum = ws.Cells(TSK_WORKER_NUM_ROW, TSK_DATE_START_COL).Value
    
    ' �^�X�N���X�g�̏�����
    Set taskList = New Collection
    
    ' �^�X�N�f�[�^�̓ǂݍ���
    For taskRow = 4 To lastRow
        If ws.Cells(taskRow, TSK_NAME_COL).Value <> "" Then
            TaskName = ws.Cells(taskRow, TSK_NAME_COL).Value
            TaskNo = ws.Cells(taskRow, TSK_NO_COL).Value
            taskPeriod = ws.Cells(taskRow, TSK_PERIOD_COL).Value
            taskPriority = ws.Cells(taskRow, TSK_PRIORITY_COL).Value
            PrevTasks = ws.Cells(taskRow, TSK_PREV_TSK_COL).Value
            StartDate = ws.Cells(taskRow, TSK_START_DATE_COL).Value
            Progress = ws.Cells(taskRow, TSK_PROGRESS_COL).Value / 100
            
            ' �^�X�N�I�u�W�F�N�g�̍쐬
            Set task = New task
            task.TaskNo = TaskNo
            task.TaskName = TaskName
            task.Period = taskPeriod
            task.Priority = taskPriority
            task.PrevTasks = PrevTasks
            task.StartDate = StartDate
            task.Progress = Progress
            
            ' �^�X�N���X�g�ɒǉ�
            taskList.Add task, TaskNo
        End If
    Next taskRow
    
    ' �X�P�W���[�����O�v�Z
    ScheduleTasks taskList, workerNum
    
    ' �K���g�`���[�g�̕`��
    For taskRow = 4 To lastRow
        TaskNo = ws.Cells(taskRow, TSK_NO_COL).Value
        On Error Resume Next
        Set task = taskList(TaskNo)
        On Error GoTo 0
        
        If Not task Is Nothing Then
            ' �\��̕`��
            taskStartCol = GetDateColumn(ws, TSK_DATE_START_ROW, TSK_DATE_START_COL, lastCol, task.ScheduledStartDate)
            taskEndCol = taskStartCol + task.Period - 1
            If taskStartCol >= TSK_DATE_START_COL And taskEndCol <= lastCol Then
                ws.Range(ws.Cells(taskRow, taskStartCol), ws.Cells(taskRow, taskEndCol)).Interior.Color = RGB(200, 200, 200)
            End If
            
            ' ���т̕`��
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
    
    MsgBox "�K���g�`���[�g�̐������������܂����I", vbInformation
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
    
    ' �^�X�N�z��̏�����
    ReDim taskArray(1 To taskList.Count)
    i = 1
    For Each task In taskList
        taskArray(i) = task
        i = i + 1
    Next task
    
    ' �^�X�N�z��̗D��x���Ƀ\�[�g
    For i = LBound(taskArray) To UBound(taskArray) - 1
        For j = i + 1 To UBound(taskArray)
            If taskArray(i).Priority > taskArray(j).Priority Then
                Swap taskArray(i), taskArray(j)
            End If
        Next j
    Next i
    
    ' �X�P�W���[�����O
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
        MsgBox "���L���̈�'" & prefix & "'??�I�`��ߔ�?���B", vbInformation
    End If
End Sub

