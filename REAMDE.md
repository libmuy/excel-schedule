# Excel タスク計画表 要件定義書

## 1. 概要

- **目的**：  
  Excelでタスクの優先度・依存関係・並行実行数を考慮したスケジューリングを行い、VBAを使ってガントチャートを自動生成する。

- **利用者**：  
  日本人（表記はすべて日本語）

- **出力**：  
  ガントチャート（VBAにより自動描画）
  予定は、セルの色塗りで示す
  実績は、矢印（Line Shape）で示す

---

## 2. スケジューリング要件

### **名前の定義**：  

- TSK_WORKER_NUM：実行可能平行タスク数
- TSK_DATE_START:
  全体スケジュールの開始日付。
  このセルの右は、週単位で、未来日付を記載している
  一セルは、セルに記載している日から1週間の期間を示す

### **列番号定数の定義**：  

- COL_NO：タスクのID
- COL_PRIORITY：タスク優先度、5段階（１～５）、1は最高優先度
- COL_PREV_TSK：先行タスクを定義する
- COL_PERIOD：タスクを実施する期間（週単位）
- COL_NAME：タスクの名前（階層の定義は、これで決める。詳細は、「タスク階層の定義」を参照）
- COL_REAL_START：実際にタスクを開始した日付、実績を示す矢印を描画時に使用
- COL_PROGRESS：進捗率（％）、実績を示す矢印を描画時に使用


### タスク階層の定義

- タスクの名前のインデントで、タスクの階層を定義する
- 親タスク名前の**右下のセル**に子タスクの名前を記入
- **右下セルに内容あり**：子タスクが存在
- **右下セルが空**：最下層タスク


### 表示仕様

- 予定
  セルの色塗りで、タスクの開始と終了を示す
  開始と終了は、優先度・依存関係・並行実行数で自動計算
- 実績
  矢印で、現在の進捗を示す
  矢印の開始と終了は、TSK_START_DATE 列に記載した日付（開始実績日）とTSK_PROGRESS列に記載した％（進捗率）で計算する



abtout the progress bar drawing
	draw 2 lines
		one line represent already done part (blue)
		another line represent not done part (black)
	
	already done part line
		the start date: 
			user inputed real task start date which is in column COL_REAL_START
			this date may be not the start of a week, in case of that the start of the line should be calculate like:
				radio = <real start date> / (<start of week> + 5)
				start_x = cell.left + (cell.width * radio)
		the end date
			start date + (period * progress)
			
		height
			cell.top + (cell.height / 2)
			
	
	not done part line
		the start date: 
			equal to the end of already done part line
		the end date
			real start date + period
			
		height
			cell.top + (cell.height / 2)
			



I want to correct this scheduling algorithm with below algorithm
do you think this algorithm is OK? 
1. sort tasks with priority
2. loop with week from start to end of all range of weeks
	1. schedule task within current available worker numbers
	2. skip task if its previous task is not done
	3. if there is no available worker shift to next week



update progress from redmine
* the column of task's redmine id is COL_REDMINE_ID
* the format of redmine id is: <RepoID>:<TicketId>
* use GetRedmineIssueProgress() to get progress of task in redmine
* create a sub get all task's redmine progress and update to task's progress column(COL_PROGRESS)
  1. loop from first task to last task
  2. if redmine id is not empty get redmine progress
  3. if get progress succeeded update the progress column